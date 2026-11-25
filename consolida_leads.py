#!/usr/bin/env python3
"""
Consolida o CSV de estabelecimentos (saída do leads_cnpj_receita.py) em uma
lista de EMPRESAS qualificada.

Uso:
  pip install pandas openpyxl
  python consolida_leads.py --entrada leads_pr_cnpj.csv --saida leads_pr_consolidado.xlsx

O que faz
  1. Agrupa por CNPJ básico (8 primeiros dígitos) -> 1 linha por empresa
  2. Marca postos de coleta (não compram reagentes) e redes nacionais
  3. Classifica o e-mail: próprio / gratuito / administrativo (contábil, jurídico...)
  4. Classifica o segmento e marca possíveis falsos positivos para revisão
  5. Recalcula a pontuação (0-100) no nível da empresa e define prioridade A/B/C

Os pesos estão no bloco PESOS abaixo; ajuste conforme o perfil de cliente
que a empresa de reagentes quer (independentes x redes).
"""
import argparse
import re

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ----------------------------- CONFIGURAÇÃO -----------------------------
# Pesos da pontuação (máximo teórico = 100). Premissas minhas, ajuste ao negócio.
PESOS = {
    "tipo": {"Independente": 25, "Pequeno grupo": 20, "Rede regional": 15, "Rede nacional": 0},
    "email": {"próprio": 25, "compartilhado": 10, "gratuito": 15, "administrativo": 5, "sem e-mail": 0},
    "telefone": 10,
    "porte": {"Demais": 15, "Pequeno": 10, "Micro": 5, "": 0},
    "maturidade_3_anos": 5,
    "tem_matriz_na_lista": 10,
    "segmento": {"Clínico": 10, "Veterinário": 5},  # demais segmentos = 0
    "so_postos_coleta": -30,
}
CORTE_A, CORTE_B = 90, 75   # calibrado p/ ~15% A, ~50% B no PR; ajuste se mudar os pesos
REDE_MIN_UNIDADES = 5          # 5+ estabelecimentos = rede
PEQUENO_GRUPO_MIN = 2
DOMINIO_COMPARTILHADO_MIN = 3   # nº de empresas distintas no mesmo domínio

# Redes nacionais: compram centralizado fora do estado (CNPJ básico ou nome)
REDES_NACIONAIS_BASE = {"61486650": "DASA", "60840055": "Fleury"}
REDES_NACIONAIS_NOME = r"DIAGNOSTICOS DA AMERICA|FLEURY|SABIN|HERMES PARDINI|DB DIAGNOSTICOS"

FREE = r"@(gmail|hotmail|yahoo|outlook|uol|terra|bol|ig|live|msn|icloud)\."
ADMIN = (r"contab|contador|escrit|assessor|fiscal|societ|juridic|financ|folha|"
         r"tribut|cont[_.\d]|[_.]cont$|^rh@|^dp@")
POSTO = r"\bCOLETA\b|\bPSC\b"
TERMO_LAB = r"LABORAT|ANALISE|ANALISES|PATOLOG|CITOLOG|DIAGNOSTIC|EXAMES|\bLAB\b|TOXICOL"
SEGMENTOS = [  # (nome, regex, só se NÃO tiver termo de laboratório?)
    ("Veterinário", r"VETERIN", False),
    ("Plano de saúde/Cooperativa", r"UNIMED|COOPERATIVA", False),
    ("Ocupacional (revisar)", r"MEDICINA DO TRABALHO|SEGURANCA DO TRABALHO|OCUPACIONAL|SEGURANCA E SAUDE|SAUDE E SEGURANCA|MEDICINA E SEGURANCA", True),
    ("Genética/Nutrição/Estética (revisar)", r"GENETIC|NUTRI|ESTETIC|ODONTO", True),
    ("Consultoria (revisar)", r"CONSULTORIA|ASSESSORIA|ENGENHARIA|PROMOCOES", True),
    ("Hospital/Instituição", r"HOSPITAL|SANTA CASA|IRMANDADE|ASSOCIACAO|FUNDACAO|UNIVERSIDADE|INSTITUTO", False),
]
# -------------------------------------------------------------------------


def classificar_email(email: str) -> str:
    if not email:
        return "sem e-mail"
    local = email.split("@")[0]
    if re.search(ADMIN, local, re.I):
        return "administrativo"
    if re.search(FREE, email, re.I):
        return "gratuito"
    return "próprio"


def classificar_segmento(texto: str) -> str:
    t = texto.upper()
    tem_lab = bool(re.search(TERMO_LAB, t))
    for nome, rx, so_sem_lab in SEGMENTOS:
        if re.search(rx, t) and not (so_sem_lab and tem_lab):
            return nome
    return "Clínico"


def moda(serie):
    s = serie[serie != ""]
    return s.value_counts().index[0] if len(s) else ""


def consolidar(df: pd.DataFrame):
    df = df.fillna("").astype(str)
    df["base"] = df["cnpj"].str.replace(r"\D", "", regex=True).str[:8]
    texto = df["razao_social"] + " " + df["nome_fantasia"]
    df["posto_coleta"] = texto.str.upper().str.contains(POSTO, regex=True)
    df["email"] = df["email"].str.lower().str.strip()
    df["email_tipo"] = df["email"].map(classificar_email)
    # domínio próprio usado por 3+ empresas distintas = grupo econômico ou escritório
    # contábil/holding; o contato não é, necessariamente, do laboratório
    df["dominio"] = df["email"].str.split("@").str[-1]
    proprio = df["email_tipo"] == "próprio"
    n_emp_dom = df[proprio].groupby("dominio")["base"].nunique()
    df["n_empresas_dominio"] = df["dominio"].map(n_emp_dom).fillna(0).astype(int)
    df.loc[proprio & (df["n_empresas_dominio"] >= DOMINIO_COMPARTILHADO_MIN), "email_tipo"] = "compartilhado"
    df["anos_atividade"] = pd.to_numeric(df["anos_atividade"], errors="coerce").fillna(0)
    df["capital_social"] = pd.to_numeric(
        df["capital_social"].str.replace(",", ".", regex=False), errors="coerce").fillna(0)

    ordem_email = {"próprio": 0, "gratuito": 1, "compartilhado": 2, "administrativo": 3, "sem e-mail": 4}
    linhas = []
    for base, g in df.groupby("base", sort=False):
        g = g.assign(
            _ord_tipo=(g["tipo"] != "Matriz").astype(int),
            _ord_email=g["email_tipo"].map(ordem_email),
        )
        ref = g.sort_values(["_ord_tipo", "posto_coleta"]).iloc[0]   # matriz preferida
        n = len(g)
        n_postos = int(g["posto_coleta"].sum())
        nome_ref = ref["razao_social"] + " " + moda(g["nome_fantasia"])
        nacional = (base in REDES_NACIONAIS_BASE
                    or bool(re.search(REDES_NACIONAIS_NOME, nome_ref.upper())))
        if nacional:
            tipo_emp = "Rede nacional"
        elif n >= REDE_MIN_UNIDADES:
            tipo_emp = "Rede regional"
        elif n >= PEQUENO_GRUPO_MIN:
            tipo_emp = "Pequeno grupo"
        else:
            tipo_emp = "Independente"

        # melhor e-mail: próprio > gratuito > administrativo; empate = mais frequente
        com_email = g[g["email"] != ""].sort_values(["_ord_email", "_ord_tipo"])
        if len(com_email):
            melhor_tipo = com_email.iloc[0]["email_tipo"]
            cand = com_email[com_email["email_tipo"] == melhor_tipo]["email"]
            email_p = cand.value_counts().index[0]
            outros = [e for e in com_email["email"].unique() if e != email_p][:3]
        else:
            melhor_tipo, email_p, outros = "sem e-mail", "", []
        dominio = (email_p.split("@")[-1] if melhor_tipo in ("próprio", "compartilhado") else "")

        tem_matriz = bool((g["tipo"] == "Matriz").any())
        so_postos = n_postos == n
        cidade_ref = ref["cidade"]
        n_cid = g["cidade"].nunique()
        seg = classificar_segmento(nome_ref)

        # pontuação
        sc = PESOS["tipo"][tipo_emp] + PESOS["email"][melhor_tipo]
        sc += PESOS["telefone"] if ref["telefone1"] or moda(g["telefone1"]) else 0
        sc += PESOS["porte"].get(ref["porte_desc"], 0)
        sc += PESOS["maturidade_3_anos"] if g["anos_atividade"].max() >= 3 else 0
        sc += PESOS["tem_matriz_na_lista"] if tem_matriz else 0
        sc += PESOS["segmento"].get(seg, 0)
        sc += PESOS["so_postos_coleta"] if so_postos else 0
        sc = max(0, min(100, sc))

        linhas.append({
            "prioridade": "A" if sc >= CORTE_A else "B" if sc >= CORTE_B else "C",
            "score": sc,
            "razao_social": ref["razao_social"],
            "nome_fantasia": ref["nome_fantasia"] or moda(g["nome_fantasia"]),
            "cnpj_matriz_ou_ref": ref["cnpj"],
            "tipo_empresa": tipo_emp,
            "segmento": seg,
            "n_unidades": n,
            "n_postos_coleta": n_postos,
            "n_cidades": n_cid,
            "cidade_ref": cidade_ref,
            "porte": ref["porte_desc"],
            "anos_atividade": int(g["anos_atividade"].max()),
            "email_principal": email_p,
            "email_tipo": melhor_tipo,
            "dominio_para_pesquisar": dominio,
            "empresas_no_dominio": int(g["n_empresas_dominio"].max()),
            "outros_emails": "; ".join(outros),
            "telefone_principal": ref["telefone1"] or moda(g["telefone1"]),
            "telefone_2": ref["telefone2"] or moda(g["telefone2"]),
            "endereco_ref": ref["endereco"],
            "matriz_no_arquivo": "sim" if tem_matriz else "não",
            "so_postos_coleta": "sim" if so_postos else "não",
            "revisar": "sim" if "revisar" in seg else "",
        })
    emp = pd.DataFrame(linhas).sort_values(
        ["score", "n_unidades"], ascending=[False, True]).reset_index(drop=True)

    chave = emp.assign(base=emp["cnpj_matriz_ou_ref"].str.replace(r"\D", "", regex=True).str[:8])
    estab = df.drop(columns=["score"], errors="ignore").merge(
        chave[["base", "prioridade", "tipo_empresa"]], on="base", how="left")
    estab = estab[["cnpj", "razao_social", "nome_fantasia", "tipo", "posto_coleta", "cidade",
                   "endereco", "telefone1", "telefone2", "email", "email_tipo",
                   "prioridade", "tipo_empresa"]]
    return emp, estab


def resumo(emp: pd.DataFrame) -> pd.DataFrame:
    rows = [("Total de empresas", len(emp)), ("", "")]
    rows += [(f"Prioridade {p}", int((emp["prioridade"] == p).sum())) for p in "ABC"]
    rows += [("", ""), ("Por tipo de empresa", "")]
    rows += [(f"  {k}", int(v)) for k, v in emp["tipo_empresa"].value_counts().items()]
    rows += [("", ""), ("Por tipo de e-mail", "")]
    rows += [(f"  {k}", int(v)) for k, v in emp["email_tipo"].value_counts().items()]
    rows += [("", ""), ("Por segmento", "")]
    rows += [(f"  {k}", int(v)) for k, v in emp["segmento"].value_counts().items()]
    rows += [("", ""), ("Pesos da pontuação (premissas, editáveis no script)", "")]
    for k, v in PESOS.items():
        rows.append((f"  {k}", str(v)))
    rows += [(f"  Corte A / B", f">= {CORTE_A} / >= {CORTE_B}"),
             ("  Rede = nº de unidades >=", REDE_MIN_UNIDADES)]
    return pd.DataFrame(rows, columns=["Item", "Valor"])


def formatar(caminho: str):
    wb = load_workbook(caminho)
    azul = PatternFill("solid", start_color="1F3864")
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                c.font = Font(name="Arial", size=10)
        if ws.title == "Resumo":
            ws.column_dimensions["A"].width = 58
            ws.column_dimensions["B"].width = 24
            ws["A1"].font = ws["B1"].font = Font(name="Arial", size=10, bold=True)
            continue
        for c in ws[1]:
            c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            c.fill = azul
            c.alignment = Alignment(vertical="center", wrap_text=True)
        ws.freeze_panes = "C2" if ws.title == "Empresas" else "B2"
        ws.auto_filter.ref = ws.dimensions
        for i, col in enumerate(ws.columns, 1):
            largura = max(len(str(c.value)) if c.value is not None else 0 for c in col)
            ws.column_dimensions[get_column_letter(i)].width = min(max(10, largura + 2), 48)
    wb.save(caminho)


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--entrada", required=True)
    ap.add_argument("--saida", default="leads_consolidado.xlsx")
    a = ap.parse_args()

    df = pd.read_csv(a.entrada, dtype=str, encoding="utf-8-sig")
    emp, estab = consolidar(df)
    with pd.ExcelWriter(a.saida, engine="openpyxl") as xw:
        resumo(emp).to_excel(xw, sheet_name="Resumo", index=False)
        emp.to_excel(xw, sheet_name="Empresas", index=False)
        emp[emp["revisar"] == "sim"].to_excel(xw, sheet_name="Revisar", index=False)
        estab.to_excel(xw, sheet_name="Estabelecimentos", index=False)
    formatar(a.saida)
    emp.to_csv(a.saida.rsplit(".", 1)[0] + "_empresas.csv", index=False, encoding="utf-8-sig")
    print(f"{len(df)} estabelecimentos -> {len(emp)} empresas")
    print(emp["prioridade"].value_counts().sort_index().to_string())
    print(f"Salvo em {a.saida}")


if __name__ == "__main__":
    main()