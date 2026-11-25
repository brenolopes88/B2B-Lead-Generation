#!/usr/bin/env python3
"""
Enriquecimento de leads (etapa 3): acha o site de cada laboratório e extrai
canais de contato.

Entrada : CSV/XLSX de EMPRESAS gerado pelo consolida_leads.py
          (leads_pr_consolidado_empresas.csv ou aba "Empresas" do .xlsx)
Saída   : planilha com as colunas originais + colunas novas:
          site, whatsapp, email_site_sugerido, emails_site, telefones_site,
          instagram, facebook, linkedin, responsavel_tecnico_possivel,
          marcas_citadas, qualidade_contato (+ dados do Google, se usado)

Como o site é encontrado
  1. Domínio do e-mail da Receita, quando é domínio PRÓPRIO do laboratório
     (ou compartilhado, mas com nome parecido com o da empresa).
  2. Opcional (--places-key): Google Places (API oficial) só para quem
     ficou sem site. Veja o aviso sobre os termos do Google abaixo.

Uso
  pip install pandas openpyxl requests beautifulsoup4
  python enriquece_leads.py --entrada leads_pr_consolidado_empresas.csv
  python enriquece_leads.py --entrada ... --prioridades A --limite 30   # teste
  python enriquece_leads.py --entrada ... --places-key SUA_CHAVE

Retomada: cada empresa processada é gravada em <saida>_cache.jsonl; se o
script for interrompido, rode de novo e ele continua de onde parou.

Boas práticas embutidas: respeita robots.txt, 1 requisição por segundo por
servidor, identifica-se no User-Agent (troque o e-mail de contato abaixo).

AVISO (Google): os termos do Google Maps Platform restringem armazenar e
redistribuir conteúdo do Places. Se a lista será ENTREGUE a terceiros, confira
os termos vigentes antes de usar --places-key (ou use só para descobrir o site).
"""
import argparse
import json
import os
import re
import threading
import time
import unicodedata
import urllib.robotparser
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
from difflib import SequenceMatcher
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ----------------------------- CONFIGURAÇÃO -----------------------------
HEADERS = {"User-Agent": "lead-research-bot/1.0 (contato: seu@email.com)"}  # TROQUE
TIMEOUT = 12
PAUSA = 1.0                 # segundos entre requisições ao mesmo servidor
MAX_PAGINAS_EXTRAS = 4      # páginas de contato/sobre além da home
CHAVES_PAGINA = ("contato", "fale-conosco", "faleconosco", "atendimento", "sobre",
                 "quem-somos", "institucional", "unidades", "onde-estamos", "equipe")
ADMIN = (r"contab|contador|escrit|assessor|fiscal|societ|juridic|financ|folha|tribut|"
         r"cont[_.\d]|[_.]cont$|^rh@|^dp@")
PRIORIDADE_EMAIL = ["compras", "comercial", "vendas", "diretoria", "gerencia",
                    "contato", "atendimento", "laborat", "lab"]
LIXO_EMAIL = ("sentry", "wixpress", "example", "seuemail", "email@", "@dominio",
              "@site", "@seusite", "u003e", "noreply", "no-reply")
EXT_IMG = (".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".css", ".js")
MARCAS = ["Roche", "Cobas", "Abbott", "Architect", "Alinity", "Siemens", "Advia",
          "Dimension", "Sysmex", "Mindray", "Beckman", "Bio-Rad", "Wiener", "Labtest",
          "Bioclin", "Doles", "Biotecnica", "Biotécnica", "Kovalent", "Horiba",
          "Werfen", "Stago", "Vitros"]
DDDS = set("""11 12 13 14 15 16 17 18 19 21 22 24 27 28 31 32 33 34 35 37 38 41 42 43 44 45
46 47 48 49 51 53 54 55 61 62 63 64 65 66 67 68 69 71 73 74 75 77 79 81 82 83 84 85 86 87
88 89 91 92 93 94 95 96 97 98 99""".split())
STOP_NOME = {"laboratorio", "laboratorios", "analises", "clinicas", "clinica", "clinico",
             "diagnostico", "diagnosticos", "medicina", "centro", "instituto", "ltda"}
# -------------------------------------------------------------------------

COLUNAS_NOVAS = ["site", "fonte_site", "status_crawl", "whatsapp", "email_site_sugerido",
                 "emails_site", "telefones_site", "instagram", "facebook", "linkedin",
                 "responsavel_tecnico_possivel", "marcas_citadas", "google_nota",
                 "google_avaliacoes", "google_status", "google_maps_url",
                 "data_enriquecimento"]

_locks = defaultdict(threading.Lock)
_locks_guard = threading.Lock()
_robots = {}
_cache_lock = threading.Lock()


def sem_acento(t: str) -> str:
    return unicodedata.normalize("NFKD", t or "").encode("ascii", "ignore").decode().lower()


def lock_do_host(host):
    with _locks_guard:
        return _locks[host]


def pode_acessar(url: str) -> bool:
    p = urlparse(url)
    raiz = f"{p.scheme}://{p.netloc}"
    if raiz not in _robots:
        rp = urllib.robotparser.RobotFileParser()
        try:
            r = requests.get(raiz + "/robots.txt", headers=HEADERS, timeout=8)
            if r.status_code == 200:
                rp.parse(r.text.splitlines())
            elif r.status_code in (401, 403):
                rp.disallow_all = True
            else:
                rp.allow_all = True
        except requests.RequestException:
            rp.allow_all = True
        _robots[raiz] = rp
    return _robots[raiz].can_fetch(HEADERS["User-Agent"], url)


def baixar(url: str):
    host = urlparse(url).netloc
    with lock_do_host(host):
        try:
            if not pode_acessar(url):
                return None
            r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        except requests.RequestException:
            time.sleep(PAUSA)
            return None
        time.sleep(PAUSA)
    ctype = r.headers.get("content-type", "html").lower()
    if r.status_code != 200 or "html" not in ctype:
        return None
    if "charset" not in ctype:      # sem charset o requests assume latin-1 e quebra acentos
        r.encoding = r.apparent_encoding or "utf-8"
    return r


def varrer_site(base: str):
    """Home + até MAX_PAGINAS_EXTRAS páginas de contato/sobre. Retorna lista de HTML."""
    r = baixar(base)
    if r is None:
        return None
    paginas = {r.url: r.text}
    soup = BeautifulSoup(r.text, "html.parser")
    host = urlparse(r.url).netloc.replace("www.", "")
    extras = []
    for a in soup.find_all("a", href=True):
        href = urljoin(r.url, a["href"]).split("#")[0]
        p = urlparse(href)
        if p.scheme not in ("http", "https") or p.netloc.replace("www.", "") != host:
            continue
        alvo = (p.path + " " + a.get_text(" ")).lower()
        if any(k in alvo for k in CHAVES_PAGINA) and href not in paginas and href not in extras:
            extras.append(href)
    for u in extras[:MAX_PAGINAS_EXTRAS]:
        rr = baixar(u)
        if rr is not None:
            paginas[rr.url] = rr.text
    return list(paginas.values()), host


# ------------------------------ EXTRAÇÃO --------------------------------
EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
WA_RE = re.compile(r"(?:wa\.me/|whatsapp\.com/send/?\?(?:[^\"'\s]*&)?phone=)\+?(\d{10,13})")
FONE_RE = re.compile(r"(?<!\d)\(?(\d{2})\)?[\s.\-]?(9?\d{4})[\s.\-]?(\d{4})(?!\d)")
IG_RE = re.compile(r"https?://(?:www\.)?instagram\.com/([A-Za-z0-9_.]+)")
FB_RE = re.compile(r"https?://(?:www\.|pt-br\.)?facebook\.com/([A-Za-z0-9_.\-]+)")
LI_RE = re.compile(r"https?://(?:[a-z]{2,3}\.)?linkedin\.com/company/([A-Za-z0-9_\-%]+)")
RT_RE = re.compile(
    r"Respons[áa]vel\s+T[ée]cnic[oa]\s*[:\-–]?\s*((?:Dr[a]?\.?\s*|Prof[a]?\.?\s*)?"
    r"[A-ZÁÉÍÓÚÂÊÔÃÕÇ][a-záéíóúâêôãõç]+(?:\s+(?:de|da|do|dos|das|e|"
    r"[A-ZÁÉÍÓÚÂÊÔÃÕÇ][a-záéíóúâêôãõç\.]+)){1,5})")
IG_IGNORAR = {"p", "reel", "reels", "explore", "accounts", "tv", "stories", "share"}
FB_IGNORAR = {"sharer", "sharer.php", "share.php", "tr", "plugins", "dialog", "login", "policies"}


def escolher_email(emails, host):
    bons = [e for e in emails if not re.search(ADMIN, e.split("@")[0], re.I)]
    if not bons:
        return ""

    def chave(e):
        local, dom = e.split("@")
        rank = next((i for i, k in enumerate(PRIORIDADE_EMAIL) if k in local), 99)
        return (0 if dom.endswith(host) else 1, rank)

    return sorted(bons, key=chave)[0]


def extrair(htmls, host):
    texto, brutos = [], "\n".join(htmls)
    emails, fones = set(), set()
    for h in htmls:
        soup = BeautifulSoup(h, "html.parser")
        for a in soup.select('a[href^="mailto:"]'):
            emails.add(a["href"][7:].split("?")[0].strip().lower())
        texto.append(soup.get_text(" "))
    t = " ".join(texto)
    emails.update(e.lower() for e in EMAIL_RE.findall(t))
    emails = {e for e in emails if not e.endswith(EXT_IMG)
              and not any(x in e for x in LIXO_EMAIL) and "@" in e}
    for d, a, b in FONE_RE.findall(t):
        if d in DDDS:
            fones.add(f"({d}) {a}-{b}")
    whats = sorted(set(WA_RE.findall(brutos)))
    ig = next((f"https://instagram.com/{n}" for n in IG_RE.findall(brutos)
               if n.lower() not in IG_IGNORAR), "")
    fb = next((f"https://facebook.com/{n}" for n in FB_RE.findall(brutos)
               if n.lower() not in FB_IGNORAR), "")
    li = next((f"https://linkedin.com/company/{n}" for n in LI_RE.findall(brutos)), "")
    rt = RT_RE.search(t)
    marcas = sorted({m for m in MARCAS if re.search(rf"\b{re.escape(m)}\b", t, re.I)})
    return {
        "whatsapp": "; ".join("+" + w for w in whats[:2]),
        "email_site_sugerido": escolher_email(sorted(emails), host),
        "emails_site": "; ".join(sorted(emails)[:6]),
        "telefones_site": "; ".join(sorted(fones)[:5]),
        "instagram": ig, "facebook": fb, "linkedin": li,
        "responsavel_tecnico_possivel": rt.group(1).strip() if rt else "",
        "marcas_citadas": ", ".join(marcas),
    }


# --------------------------- DESCOBERTA DE SITE -------------------------
def dominio_combina_nome(dominio: str, texto: str) -> bool:
    label = sem_acento(dominio.replace("http://", "").replace("https://", "")).split(".")[0]
    palavras = [p for p in re.findall(r"[a-z]{5,}", sem_acento(texto)) if p not in STOP_NOME]
    return len(label) >= 4 and any(p in label or label in p for p in palavras)


def site_do_email(row):
    dom = (row.get("dominio_para_pesquisar") or "").strip()
    tipo = row.get("email_tipo", "")
    if not dom:
        return ""
    ok = tipo == "próprio" or (
        tipo == "compartilhado"
        and dominio_combina_nome(dom, f"{row.get('razao_social', '')} {row.get('nome_fantasia', '')}"))
    if not ok:
        return ""
    return dom if dom.startswith("http") else "https://" + dom


def buscar_places(row, chave):
    """Google Places (API nova). Retorna dict ou None se sem correspondência confiável."""
    nome = row.get("nome_fantasia") or row.get("razao_social")
    cidade = row.get("cidade_ref", "")
    corpo = {"textQuery": f"{nome} {cidade} PR", "languageCode": "pt-BR",
             "regionCode": "BR", "pageSize": 3}
    cab = {"X-Goog-Api-Key": chave, "Content-Type": "application/json",
           "X-Goog-FieldMask": ("places.displayName,places.formattedAddress,places.websiteUri,"
                                "places.rating,places.userRatingCount,places.googleMapsUri,"
                                "places.businessStatus")}
    try:
        r = requests.post("https://places.googleapis.com/v1/places:searchText",
                          json=corpo, headers=cab, timeout=TIMEOUT)
        lugares = r.json().get("places", []) if r.status_code == 200 else []
    except (requests.RequestException, ValueError):
        return None
    return escolher_lugar(lugares, row)


def escolher_lugar(lugares, row):
    nomes = [sem_acento(row.get("nome_fantasia", "")), sem_acento(row.get("razao_social", ""))]
    cidade = sem_acento(row.get("cidade_ref", ""))
    melhor, nota = None, 0.0
    for p in lugares:
        nome_p = sem_acento(p.get("displayName", {}).get("text", ""))
        if cidade and cidade not in sem_acento(p.get("formattedAddress", "")):
            continue
        s = max(SequenceMatcher(None, nome_p, n).ratio() for n in nomes if n) if any(nomes) else 0
        if s > nota:
            melhor, nota = p, s
    return melhor if melhor and nota >= 0.5 else None


# ------------------------------- PIPELINE -------------------------------
def enriquecer(row, chave_places):
    res = {"key": row["cnpj_matriz_ou_ref"], "site": "", "fonte_site": "", "status_crawl": "",
           "google_nota": "", "google_avaliacoes": "", "google_status": "",
           "google_maps_url": "", "data_enriquecimento": date.today().isoformat()}
    site = site_do_email(row)
    if site:
        res["fonte_site"] = "dominio_do_email"
    elif chave_places:
        lugar = buscar_places(row, chave_places)
        if lugar:
            res.update({
                "google_nota": lugar.get("rating", ""),
                "google_avaliacoes": lugar.get("userRatingCount", ""),
                "google_status": lugar.get("businessStatus", ""),
                "google_maps_url": lugar.get("googleMapsUri", ""),
            })
            site = lugar.get("websiteUri", "")
            res["fonte_site"] = "google_places" if site else ""
    if not site:
        res["status_crawl"] = "sem_site"
        return res
    res["site"] = site
    varrido = varrer_site(site)
    if varrido is None:
        res["status_crawl"] = "site_inacessivel_ou_bloqueado"
        return res
    htmls, host = varrido
    res.update(extrair(htmls, host))
    res["status_crawl"] = "ok"
    return res


def qualidade(r):
    direto = bool(r.get("whatsapp")) or bool(r.get("email_site_sugerido")) \
        or r.get("email_tipo") == "próprio"
    fone = bool(r.get("telefones_site")) or bool(r.get("telefone_principal"))
    if direto and fone:
        return "Pronto"
    if direto or fone:
        return "Parcial"
    return "Sem contato"


def formatar(caminho):
    wb = load_workbook(caminho)
    ws = wb.active
    for row in ws.iter_rows():
        for c in row:
            c.font = Font(name="Arial", size=10)
    for c in ws[1]:
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color="1F3864")
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions
    for i, col in enumerate(ws.columns, 1):
        larg = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[get_column_letter(i)].width = min(max(10, larg + 2), 48)
    wb.save(caminho)


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--entrada", required=True)
    ap.add_argument("--saida", default="leads_enriquecidos.xlsx")
    ap.add_argument("--prioridades", nargs="+", default=["A", "B"])
    ap.add_argument("--threads", type=int, default=8)
    ap.add_argument("--limite", type=int, help="processa só N empresas (teste)")
    ap.add_argument("--places-key", default=os.environ.get("GOOGLE_PLACES_API_KEY"))
    a = ap.parse_args()

    if a.entrada.lower().endswith(".xlsx"):
        df = pd.read_excel(a.entrada, sheet_name="Empresas", dtype=str)
    else:
        df = pd.read_csv(a.entrada, dtype=str, encoding="utf-8-sig")
    df = df.fillna("")
    alvo = df[df["prioridade"].isin(a.prioridades)]
    if a.limite:
        alvo = alvo.head(a.limite)

    cache_path = a.saida.rsplit(".", 1)[0] + "_cache.jsonl"
    cache = {}
    if os.path.exists(cache_path):
        for linha in open(cache_path, encoding="utf-8"):
            d = json.loads(linha)
            cache[d["key"]] = d
    pendentes = [r for r in alvo.to_dict("records") if r["cnpj_matriz_ou_ref"] not in cache]
    print(f"{len(alvo)} empresas no filtro | {len(cache)} já em cache | "
          f"{len(pendentes)} a processar | Places: {'sim' if a.places_key else 'não'}")

    feitos = 0
    with ThreadPoolExecutor(max_workers=a.threads) as ex, \
            open(cache_path, "a", encoding="utf-8") as fcache:
        futs = {ex.submit(enriquecer, r, a.places_key): r for r in pendentes}
        for f in as_completed(futs):
            try:
                d = f.result()
            except Exception as e:  # nunca derrubar a rodada por 1 site problemático
                r = futs[f]
                d = {"key": r["cnpj_matriz_ou_ref"], "status_crawl": f"erro: {type(e).__name__}"}
            cache[d["key"]] = d
            with _cache_lock:
                fcache.write(json.dumps(d, ensure_ascii=False) + "\n")
                fcache.flush()
            feitos += 1
            if feitos % 25 == 0 or feitos == len(pendentes):
                print(f"  {feitos}/{len(pendentes)}", flush=True)

    enr = pd.DataFrame(cache.values()).rename(columns={"key": "cnpj_matriz_ou_ref"})
    out = df.merge(enr, on="cnpj_matriz_ou_ref", how="left")
    for c in COLUNAS_NOVAS:          # garante as colunas mesmo se nenhum site respondeu
        if c not in out.columns:
            out[c] = ""
    out = out.fillna("")
    out["status_crawl"] = out["status_crawl"].replace("", "nao_processado")
    out["qualidade_contato"] = out.apply(qualidade, axis=1)
    out.loc[out["status_crawl"] == "nao_processado", "qualidade_contato"] = ""
    out.to_excel(a.saida, index=False)
    formatar(a.saida)
    out.to_csv(a.saida.rsplit(".", 1)[0] + ".csv", index=False, encoding="utf-8-sig")

    proc = out[out["status_crawl"] != "nao_processado"]
    print(f"\nProcessadas: {len(proc)}")
    print(proc["status_crawl"].value_counts().to_string())
    print("Com WhatsApp:", (proc["whatsapp"] != "").sum(),
          "| com e-mail do site:", (proc["email_site_sugerido"] != "").sum(),
          "| responsável técnico:", (proc["responsavel_tecnico_possivel"] != "").sum())
    print(proc["qualidade_contato"].value_counts().to_string())
    print(f"Salvo em {a.saida}")


if __name__ == "__main__":
    main()