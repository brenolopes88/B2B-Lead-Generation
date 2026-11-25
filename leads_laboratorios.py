#!/usr/bin/env python3
"""
Gerador de leads B2B: laboratórios de análises clínicas.

Fontes (todas públicas/gratuitas):
  1. OpenStreetMap (Overpass API): laboratórios por cidade/estado
  2. Site de cada laboratório: e-mails e WhatsApp publicados
  3. (opcional) BrasilAPI: dados do CNPJ, se você já tiver o CNPJ

Uso:
  pip install requests beautifulsoup4
  python leads_laboratorios.py --area "Paraná" --saida leads_pr.csv
  python leads_laboratorios.py --area "Curitiba" --nivel city --saida leads_cwb.csv

Boas práticas: só dados B2B públicos, respeite robots.txt, ofereça opt-out
no primeiro contato (LGPD, legítimo interesse).
"""
import argparse
import csv
import re
import time
import urllib.robotparser
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

OVERPASS = "https://overpass-api.de/api/interpreter"
HEADERS = {"User-Agent": "lead-research-bot/1.0 (contato: seu@email.com)"}
EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
WHATS_RE = re.compile(r"(?:wa\.me/|api\.whatsapp\.com/send\?phone=)(\d{10,13})")
PHONE_RE = re.compile(r"\(?\d{2}\)?\s?9?\d{4}[-\s]?\d{4}")
IGNORAR = ("png", "jpg", "jpeg", "gif", "webp", "svg", "sentry", "example")
PAGINAS_CONTATO = ("", "contato", "contatos", "fale-conosco", "sobre", "quem-somos")


def buscar_osm(area: str, nivel: str):
    """Laboratórios/clínicas de análises no OSM."""
    admin = "4" if nivel == "state" else "8"
    query = f"""
    [out:json][timeout:120];
    area["name"="{area}"]["boundary"="administrative"]["admin_level"="{admin}"]->.a;
    (
      nwr["healthcare"="laboratory"](area.a);
      nwr["amenity"="clinic"]["healthcare:speciality"~"laboratory"](area.a);
      nwr["name"~"[Ll]aborat[óo]rio|[Aa]n[áa]lises [Cc]l[íi]nicas"](area.a);
    );
    out center tags;
    """
    r = requests.post(OVERPASS, data={"data": query}, headers=HEADERS, timeout=180)
    r.raise_for_status()
    leads = []
    for el in r.json().get("elements", []):
        t = el.get("tags", {})
        if not t.get("name"):
            continue
        leads.append({
            "nome": t["name"],
            "telefone": t.get("phone") or t.get("contact:phone", ""),
            "site": t.get("website") or t.get("contact:website", ""),
            "email": t.get("email") or t.get("contact:email", ""),
            "cidade": t.get("addr:city", ""),
            "endereco": " ".join(filter(None, [t.get("addr:street"), t.get("addr:housenumber")])),
            "whatsapp": "",
        })
    return leads


def pode_acessar(url: str) -> bool:
    try:
        rp = urllib.robotparser.RobotFileParser()
        rp.set_url(urljoin(url, "/robots.txt"))
        rp.read()
        return rp.can_fetch(HEADERS["User-Agent"], url)
    except Exception:
        return True


def varrer_site(site: str):
    """Extrai e-mails, WhatsApp e telefones das páginas de contato."""
    if not site.startswith("http"):
        site = "http://" + site
    emails, whats, fones = set(), set(), set()
    for pagina in PAGINAS_CONTATO:
        url = urljoin(site if site.endswith("/") else site + "/", pagina)
        if not pode_acessar(url):
            continue
        try:
            r = requests.get(url, headers=HEADERS, timeout=10)
            if r.status_code != 200:
                continue
        except requests.RequestException:
            continue
        html = r.text
        soup = BeautifulSoup(html, "html.parser")
        for a in soup.select('a[href^="mailto:"]'):
            emails.add(a["href"][7:].split("?")[0].strip().lower())
        emails.update(e.lower() for e in EMAIL_RE.findall(soup.get_text(" ")))
        whats.update(WHATS_RE.findall(html))
        fones.update(PHONE_RE.findall(soup.get_text(" ")))
        time.sleep(1)  # educado com o servidor
    emails = {e for e in emails if not e.endswith(IGNORAR)}
    return sorted(emails), sorted(whats), sorted(fones)


def enriquecer_cnpj(cnpj: str):
    """Opcional: BrasilAPI (porte, CNAE, situação, e-mail cadastral)."""
    cnpj = re.sub(r"\D", "", cnpj)
    r = requests.get(f"https://brasilapi.com.br/api/cnpj/v1/{cnpj}", timeout=15)
    return r.json() if r.status_code == 200 else {}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--area", required=True, help="Nome do estado ou cidade")
    ap.add_argument("--nivel", choices=["state", "city"], default="state")
    ap.add_argument("--saida", default="leads.csv")
    args = ap.parse_args()

    print(f"Buscando laboratórios em {args.area}...")
    leads = buscar_osm(args.area, args.nivel)
    print(f"{len(leads)} encontrados. Varrendo sites...")

    vistos = set()
    finais = []
    for i, l in enumerate(leads, 1):
        chave = l["nome"].lower()
        if chave in vistos:
            continue
        vistos.add(chave)
        if l["site"]:
            emails, whats, fones = varrer_site(l["site"])
            if emails and not l["email"]:
                l["email"] = "; ".join(emails[:3])
            l["whatsapp"] = "; ".join(whats[:2])
            if fones and not l["telefone"]:
                l["telefone"] = fones[0]
        # pontuação simples: quanto mais canais de contato, melhor o lead
        l["score"] = sum(bool(l[c]) for c in ("email", "whatsapp", "telefone", "site"))
        finais.append(l)
        print(f"[{i}/{len(leads)}] {l['nome']} (score {l['score']})")

    finais.sort(key=lambda x: x["score"], reverse=True)
    campos = ["nome", "cidade", "endereco", "telefone", "whatsapp", "email", "site", "score"]
    with open(args.saida, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=campos)
        w.writeheader()
        w.writerows(finais)
    print(f"Pronto: {args.saida} ({len(finais)} leads)")


if __name__ == "__main__":
    main()