"""Acha na copia local da Receita as empresas e os estabelecimentos de um grupo (ex.: Kraft Heinz).

Uso: python -I receita_grupo.py SAIDA.json "REGEX DA RAZAO" [pasta_receita]
- passa 1: Empresas*.zip -> raizes cuja razao casa com o regex (0 raiz, 1 razao, 2 natureza, 4 capital, 5 porte)
- passa 2: Estabelecimentos*.zip -> todos os estabelecimentos dessas raizes
Pasta da Receita: argumento, AUTOPEL_RECEITA ou a pasta mais nova de C:\\dados_receita.
"""
import csv
import io
import json
import os
import re
import sys
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
                texto = io.TextIOWrapper(f, encoding="latin-1", newline="")
                yield from csv.reader(texto, delimiter=";", quotechar='"')


def busca_empresas(args):
    caminho, padrao = args
    regex = re.compile(padrao)
    return [r for r in linhas(caminho) if regex.search(r[1])]


def busca_estab(args):
    caminho, raizes = args
    return [r for r in linhas(caminho) if r[0] in raizes]


def municipios(pasta):
    return {r[0]: r[1] for r in linhas(os.path.join(pasta, "Municipios.zip"))}


if __name__ == "__main__":
    saida, padrao = sys.argv[1], sys.argv[2]
    pasta = pasta_receita(sys.argv[3] if len(sys.argv) > 3 else None)
    emp_zips = [os.path.join(pasta, f"Empresas{i}.zip") for i in range(10)]
    est_zips = [os.path.join(pasta, f"Estabelecimentos{i}.zip") for i in range(10)]
    with Pool(4) as pool:
        empresas = [r for parte in pool.map(busca_empresas, [(z, padrao) for z in emp_zips]) for r in parte]
        raizes = {r[0] for r in empresas}
        print(len(empresas), "empresas", flush=True)
        estab = [r for parte in pool.map(busca_estab, [(z, raizes) for z in est_zips]) for r in parte]
    mun = municipios(pasta)
    for r in estab:
        r.append(mun.get(r[20], ""))
    json.dump({"pasta": pasta, "empresas": empresas, "estabelecimentos": estab},
              open(saida, "w", encoding="utf-8"), ensure_ascii=False)
    print(len(estab), "estabelecimentos")
