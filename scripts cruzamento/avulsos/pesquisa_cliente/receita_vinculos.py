"""Pesquisa de um cliente na copia local da Receita: o CNPJ dele e as empresas ligadas por razao, telefone, e-mail ou endereco.

Uso:
  python receita_vinculos.py SAIDA.json --raiz 12345678 [--razao REGEX] [--tel 8 a 9 digitos ...] [--email TRECHO ...]
                             [--mun COD_MUN_RECEITA --logr TRECHO --num NUMERO] [--fantasia REGEX --uf UF] [--pasta P]
Vinculos procurados (uma passada em cada zip):
  - Empresas*.zip: razao social que casa com --razao (ex.: sobrenome da familia; MEI traz o nome do dono na razao);
  - Estabelecimentos*.zip: raiz do cliente ou das razoes achadas; telefone (DDD1/tel1, DDD2/tel2, fax) terminado em --tel;
    e-mail com --email; mesmo endereco (municipio + logradouro + numero); nome fantasia com --fantasia na --uf;
  - depois, Empresas*.zip de novo, para trazer a razao social de toda raiz achada.
Estabelecimentos (latin-1, ';'): 0 raiz, 1 ordem, 2 dv, 3 matriz/filial, 4 fantasia, 5 situacao, 6 data situacao,
  7 motivo, 10 inicio, 11 CNAE, 12 CNAEs secundarios, 13-17 endereco, 18 CEP, 19 UF, 20 municipio, 21-26 telefones,
  27 e-mail.  Empresas: 0 raiz, 1 razao, 2 natureza, 3 qualificacao do responsavel, 4 capital, 5 porte.
Pasta da Receita: --pasta, AUTOPEL_RECEITA ou a pasta mais nova de C:\\dados_receita.
"""
import argparse
import csv
import io
import json
import os
import re
import zipfile
from multiprocessing import Pool


def pasta_receita(arg=None):
    if arg:
        return arg
    if os.environ.get("AUTOPEL_RECEITA"):
        return os.environ["AUTOPEL_RECEITA"]
    base = r"C:\dados_receita"
    return os.path.join(base, sorted(os.listdir(base))[-1])


def linhas(caminho_zip):
    with zipfile.ZipFile(caminho_zip) as z:
        for nome in z.namelist():
            with z.open(nome) as f:
                yield from csv.reader(io.TextIOWrapper(f, encoding="latin-1", newline=""), delimiter=";", quotechar='"')


def empresas_por_razao(args):
    caminho, padrao = args
    rx = re.compile(padrao)
    return [r for r in linhas(caminho) if rx.search(r[1])]


def empresas_por_raiz(args):
    caminho, raizes = args
    return [r for r in linhas(caminho) if r[0] in raizes]


def estab(args):
    caminho, p = args
    raizes = set(p["raizes"])
    tels = p["tel"]
    emails = [e.lower() for e in p["email"]]
    rx_fant = re.compile(p["fantasia"]) if p["fantasia"] else None
    achados = []
    for r in linhas(caminho):
        if len(r) < 28:
            continue
        motivos = []
        if r[0] in raizes:
            motivos.append("raiz")
        fones = [r[21] + r[22], r[23] + r[24], r[25] + r[26]]
        if tels and any(f and any(f.endswith(t) for t in tels) for f in fones):
            motivos.append("telefone")
        if emails and any(e in r[27].lower() for e in emails):
            motivos.append("e-mail")
        if p["mun"] and r[20] == p["mun"] and p["logr"] in r[14].upper() and r[15].strip() == p["num"]:
            motivos.append("endereco")
        if rx_fant and r[19] == p["uf"] and rx_fant.search(r[4]):
            motivos.append("fantasia")
        if motivos:
            achados.append(r + ["|".join(motivos)])
    return achados


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("saida")
    ap.add_argument("--raiz", required=True)
    ap.add_argument("--razao", default="")
    ap.add_argument("--tel", nargs="*", default=[])
    ap.add_argument("--email", nargs="*", default=[])
    ap.add_argument("--mun", default="")
    ap.add_argument("--logr", default="")
    ap.add_argument("--num", default="")
    ap.add_argument("--fantasia", default="")
    ap.add_argument("--uf", default="")
    ap.add_argument("--pasta")
    a = ap.parse_args()
    pasta = pasta_receita(a.pasta)
    emp_zips = [os.path.join(pasta, f"Empresas{i}.zip") for i in range(10)]
    est_zips = [os.path.join(pasta, f"Estabelecimentos{i}.zip") for i in range(10)]
    with Pool(4) as pool:
        por_razao = []
        if a.razao:
            por_razao = [r for parte in pool.map(empresas_por_razao, [(z, a.razao) for z in emp_zips]) for r in parte]
        print(len(por_razao), "empresas pela razao", flush=True)
        p = {"raizes": sorted({a.raiz} | {r[0] for r in por_razao}), "tel": a.tel, "email": a.email,
             "mun": a.mun, "logr": a.logr.upper(), "num": a.num, "fantasia": a.fantasia, "uf": a.uf}
        estabs = [r for parte in pool.map(estab, [(z, p) for z in est_zips]) for r in parte]
        print(len(estabs), "estabelecimentos", flush=True)
        raizes = {r[0] for r in estabs} | {a.raiz}
        empresas = [r for parte in pool.map(empresas_por_raiz, [(z, raizes) for z in emp_zips]) for r in parte]
    mun = {r[0]: r[1] for r in linhas(os.path.join(pasta, "Municipios.zip"))}
    for r in estabs:
        r.append(mun.get(r[20], ""))
    json.dump({"pasta": pasta, "empresas_razao": por_razao, "empresas": empresas, "estabelecimentos": estabs},
              open(a.saida, "w", encoding="utf-8"), ensure_ascii=False)
    print(len(empresas), "empresas no total")
