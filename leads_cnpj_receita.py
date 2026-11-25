#!/usr/bin/env python3
"""
Leads de laboratórios a partir da base aberta de CNPJ da Receita Federal.

Etapa 1 (este script): filtra TODOS os estabelecimentos ATIVOS por CNAE + UF
e junta razão social, porte e município.
Etapa 2 (depois): qualificar (site, WhatsApp, tamanho, volume de exames).

Como usar
---------
1. Baixe os .zip da Receita (Empresas0-9, Estabelecimentos0-9, Municipios,
   Cnaes) em uma pasta, por exemplo ./cnpj_zips
   Fonte: https://dados.gov.br/dados/conjuntos-dados/cadastro-nacional-da-pessoa-juridica---cnpj
   (a Receita mudou o layout/caminhos em jan/2026; baixe pelo portal acima).
   Só os Estabelecimentos, Empresas e Municipios são necessários.
2. pip install pandas
3. python leads_cnpj_receita.py --pasta ./cnpj_zips --uf PR --saida leads_pr_cnpj.csv

Opções úteis
  --uf PR SC          vários estados
  --municipio CURITIBA  filtra por nome de município (após o filtro de UF)
  --perfil ampliado   inclui hemoterapia, análises técnicas, veterinária, hospitais
  --cnaes 8640202 ... define os CNAEs manualmente (7 dígitos, sem pontuação)
  --incluir-secundario  também pega quem tem o CNAE núcleo só como secundário

Nota: o e-mail e o telefone da Receita muitas vezes são do CONTADOR da
empresa, não do laboratório. Trate como canal de partida, não como decisor.
"""
import argparse
import glob
import os
import re
import sys
import zipfile
from datetime import date

import pandas as pd

# 8640-2/02 laboratórios clínicos | 8640-2/01 anatomia patológica e citológica
CNAES_NUCLEO = {"8640202", "8640201"}
# + 8640-2/12 hemoterapia | 7120-1/00 testes e análises técnicas
# + 7500-1/00 veterinária | 8610-1/01 hospitais
CNAES_AMPLIADO = CNAES_NUCLEO | {"8640212", "7120100", "7500100", "8610101"}

CNAES_MERCADOS_PARALELOS = {
    "8630502", "8630503",  # Medicina do Trabalho / Exames
    "8630501",             # Clínicas de Fertilidade / Reprodução
    "1051100", "1012101",  # Indústria Alimentícia (Laticínios / Frigoríficos)
    "9602502",             # Clínicas de Estética Avançada (Compradores de Centrífuga/Tubos)
    "8532500", "8542200"   # Universidades e Escolas Técnicas de Saúde
}


COLS_ESTAB = [
    "cnpj_basico", "cnpj_ordem", "cnpj_dv", "matriz_filial", "nome_fantasia",
    "situacao", "data_situacao", "motivo_situacao", "cidade_exterior", "pais",
    "data_inicio", "cnae_principal", "cnae_secundaria", "tipo_logradouro",
    "logradouro", "numero", "complemento", "bairro", "cep", "uf", "municipio",
    "ddd1", "tel1", "ddd2", "tel2", "ddd_fax", "fax", "email",
    "situacao_especial", "data_situacao_especial",
]
USAR_ESTAB = [
    "cnpj_basico", "cnpj_ordem", "cnpj_dv", "matriz_filial", "nome_fantasia",
    "situacao", "data_inicio", "cnae_principal", "cnae_secundaria",
    "tipo_logradouro", "logradouro", "numero", "complemento", "bairro", "cep",
    "uf", "municipio", "ddd1", "tel1", "ddd2", "tel2", "email",
]
COLS_EMPRESA = [
    "cnpj_basico", "razao_social", "natureza_juridica", "qualif_resp",
    "capital_social", "porte", "ente_federativo",
]
PORTE = {"00": "Não informado", "01": "Micro", "03": "Pequeno", "05": "Demais"}
CHUNK = 500_000


def achar_zips(pasta, prefixo):
    achados = [
        p for p in glob.glob(os.path.join(pasta, "**", "*.zip"), recursive=True)
        if os.path.basename(p).lower().startswith(prefixo.lower())
    ]
    return sorted(achados)


def ler_zip(caminho, **kw):
    """Itera sobre o CSV (sem cabeçalho, ';', latin-1) dentro do zip."""
    z = zipfile.ZipFile(caminho)
    membro = z.namelist()[0]
    with z.open(membro) as f:
        yield from pd.read_csv(
            f, sep=";", header=None, encoding="latin-1", dtype=str,
            chunksize=CHUNK, **kw,
        )


def etapa_estabelecimentos(zips, ufs, cnaes, cnaes_sec):
    partes = []
    total = 0
    padrao_sec = "|".join(sorted(cnaes_sec)) if cnaes_sec else None
    for z in zips:
        print(f"  lendo {os.path.basename(z)} ...", flush=True)
        for df in ler_zip(z, names=COLS_ESTAB, usecols=USAR_ESTAB):
            total += len(df)
            base = (df["situacao"] == "02") & df["uf"].isin(ufs)
            alvo = df["cnae_principal"].isin(cnaes)
            if padrao_sec:
                alvo |= df["cnae_secundaria"].fillna("").str.contains(padrao_sec)
            achou = df[base & alvo]
            if len(achou):
                partes.append(achou)
    print(f"  {total:,} estabelecimentos lidos")
    if not partes:
        return pd.DataFrame(columns=USAR_ESTAB)
    return pd.concat(partes, ignore_index=True)


def etapa_empresas(zips, basicos):
    partes = []
    for z in zips:
        print(f"  lendo {os.path.basename(z)} ...", flush=True)
        for df in ler_zip(z, names=COLS_EMPRESA):
            achou = df[df["cnpj_basico"].isin(basicos)]
            if len(achou):
                partes.append(achou)
    if not partes:
        return pd.DataFrame(columns=COLS_EMPRESA)
    return pd.concat(partes, ignore_index=True).drop_duplicates("cnpj_basico")


def ler_municipios(pasta):
    zs = achar_zips(pasta, "Municipios")
    if not zs:
        print("  aviso: Municipios.zip não encontrado; saída terá só o código.")
        return {}
    mapa = {}
    for df in ler_zip(zs[0], names=["codigo", "nome"]):
        mapa.update(dict(zip(df["codigo"], df["nome"])))
    return mapa


def fmt_cnpj(r):
    c = f"{r['cnpj_basico']}{r['cnpj_ordem']}{r['cnpj_dv']}"
    if len(c) == 14:
        return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"
    return c


def fmt_fone(ddd, tel):
    ddd = "" if pd.isna(ddd) else str(ddd).strip()
    tel = "" if pd.isna(tel) else str(tel).strip()
    return f"({ddd}) {tel}" if ddd and tel else tel


def anos_atividade(s):
    try:
        ano = int(s[:4])
        return max(0, date.today().year - ano)
    except (TypeError, ValueError):
        return 0


def montar_saida(estab, empresas, municipios):
    df = estab.merge(empresas, on="cnpj_basico", how="left")
    df["cnpj"] = df.apply(fmt_cnpj, axis=1)
    df["tipo"] = df["matriz_filial"].map({"1": "Matriz", "2": "Filial"})
    df["porte_desc"] = df["porte"].map(PORTE).fillna("")
    df["cidade"] = df["municipio"].map(municipios).fillna(df["municipio"])
    df["endereco"] = (
        df["tipo_logradouro"].fillna("") + " " + df["logradouro"].fillna("")
        + ", " + df["numero"].fillna("") + " " + df["complemento"].fillna("")
        + " - " + df["bairro"].fillna("")
    ).str.replace(r"\s+", " ", regex=True).str.strip()
    df["telefone1"] = [fmt_fone(a, b) for a, b in zip(df["ddd1"], df["tel1"])]
    df["telefone2"] = [fmt_fone(a, b) for a, b in zip(df["ddd2"], df["tel2"])]
    df["email"] = df["email"].fillna("").str.lower().str.strip()
    df["anos_atividade"] = df["data_inicio"].map(anos_atividade)

    # Pontuação inicial simples (refinar na etapa de qualificação)
    df["score"] = (
        (df["email"] != "").astype(int) * 2
        + (df["telefone1"] != "").astype(int)
        + (df["tipo"] == "Matriz").astype(int)
        + df["porte"].map({"05": 2, "03": 1}).fillna(0).astype(int)
        + (df["anos_atividade"] >= 3).astype(int)
    )
    cols = [
        "cnpj", "razao_social", "nome_fantasia", "tipo", "cnae_principal",
        "porte_desc", "capital_social", "cidade", "uf", "cep", "endereco",
        "telefone1", "telefone2", "email", "anos_atividade", "score",
    ]
    return df[cols].sort_values(["score", "razao_social"], ascending=[False, True])


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--pasta", required=True, help="Pasta com os .zip da Receita")
    ap.add_argument("--uf", nargs="+", default=["PR"])
    ap.add_argument("--municipio", help="Filtrar por nome do município")
    ap.add_argument("--perfil", choices=["nucleo", "ampliado"], default="nucleo")
    ap.add_argument("--cnaes", nargs="+", help="CNAEs (7 dígitos) manuais")
    ap.add_argument("--incluir-secundario", action="store_true")
    ap.add_argument("--saida", default="leads_cnpj.csv")
    a = ap.parse_args()

    ufs = {u.upper() for u in a.uf}
    if a.cnaes:
        cnaes = {re.sub(r"\D", "", c) for c in a.cnaes}
    else:
        cnaes = CNAES_AMPLIADO if a.perfil == "ampliado" else CNAES_NUCLEO
    cnaes_sec = CNAES_NUCLEO if a.incluir_secundario else None

    z_estab = achar_zips(a.pasta, "Estabelecimentos")
    z_emp = achar_zips(a.pasta, "Empresas")
    if not z_estab or not z_emp:
        sys.exit("Não achei Estabelecimentos*.zip e/ou Empresas*.zip em " + a.pasta)

    print(f"UF: {sorted(ufs)} | CNAEs: {sorted(cnaes)}")
    print("1/3 Estabelecimentos ativos...")
    estab = etapa_estabelecimentos(z_estab, ufs, cnaes, cnaes_sec)
    print(f"  {len(estab):,} estabelecimentos ativos no filtro")
    if estab.empty:
        sys.exit("Nenhum resultado. Confira UF e CNAEs.")

    print("2/3 Empresas (razão social, porte)...")
    empresas = etapa_empresas(z_emp, set(estab["cnpj_basico"]))

    print("3/3 Municípios e saída...")
    municipios = ler_municipios(a.pasta)
    out = montar_saida(estab, empresas, municipios)
    if a.municipio:
        out = out[out["cidade"].str.upper() == a.municipio.upper()]

    out.to_csv(a.saida, index=False, encoding="utf-8-sig")
    print(f"Pronto: {a.saida} ({len(out):,} leads)")
    print(f"  com e-mail: {(out['email'] != '').sum():,} | "
          f"com telefone: {(out['telefone1'] != '').sum():,}")


if __name__ == "__main__":
    main()