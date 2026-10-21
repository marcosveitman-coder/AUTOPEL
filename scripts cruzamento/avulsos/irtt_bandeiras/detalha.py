"""Detalha raizes de CNPJ na Receita: razao, porte, matriz (CNPJ, situacao, CNAE, UF, municipio), ativos e nomes fantasia.
Uso: python detalha.py raizes.pkl saida.pkl   (base em AUTOPEL_RECEITA, padrao C:/dados_receita/<ultimo mes>)"""
import os, sys, glob, zipfile, io, pickle, collections
from multiprocessing import Pool
BASE = os.environ.get("AUTOPEL_RECEITA") or sorted(glob.glob("C:/dados_receita/20*"))[-1]
F = lambda s: s.strip().strip('"')
ROOTS = set()
def init(r):
    global ROOTS; ROOTS = r
def emp(path):
    out = {}
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            r = raw[1:9]
            if r in ROOTS:
                p = raw.split('";"'); out[r] = (F(p[1]), F(p[2]), F(p[5]))
    return out
def est(path):
    out = {}
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            r = raw[1:9]
            if r not in ROOTS: continue
            p = raw.split('";"')
            d = out.setdefault(r, {"ativos": 0, "fant": collections.Counter(), "uf": collections.Counter(), "matriz": None})
            if F(p[3]) == "1":
                d["matriz"] = dict(cnpj=r + F(p[1]) + F(p[2]), fant=F(p[4]), sit=F(p[5]), cnae=F(p[11]), uf=F(p[19]), mun=F(p[20]))
            if F(p[5]) == "02":
                d["ativos"] += 1; d["fant"][F(p[4])] += 1; d["uf"][F(p[19])] += 1
    return out
if __name__ == "__main__":
    roots = pickle.load(open(sys.argv[1], "rb"))
    E, I = {}, {}
    with Pool(4, initializer=init, initargs=(roots,)) as pool:
        for o in pool.imap_unordered(emp, sorted(glob.glob(os.path.join(BASE, "Empresas*.zip")))): E.update(o)
        for o in pool.imap_unordered(est, sorted(glob.glob(os.path.join(BASE, "Estabelecimentos*.zip")))):
            for r, d in o.items():
                if r in I:
                    x = I[r]; x["ativos"] += d["ativos"]; x["fant"].update(d["fant"]); x["uf"].update(d["uf"]); x["matriz"] = x["matriz"] or d["matriz"]
                else: I[r] = d
    pickle.dump({"emp": E, "info": I}, open(sys.argv[2], "wb"))
    print("empresas:", len(E), "| com estabelecimentos:", len(I))
