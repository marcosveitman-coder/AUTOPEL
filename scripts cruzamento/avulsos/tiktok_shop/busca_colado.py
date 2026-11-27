"""2a passada: nome da loja sem espacos (6+ letras) = nome fantasia (ativo) ou razao (sem LTDA etc.) sem espacos.
Uso: python busca_colado.py <tt_q.json> <saida.pkl>
Saida: {"est": [linhas como no busca_receita], "emp": {raiz: razao}, "hits": [(termo_colado, raiz, como)]}"""
import os, sys, glob, zipfile, io, pickle, json, re
from multiprocessing import Pool
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from prep_termos import TBL
DIR = os.environ.get("AUTOPEL_RECEITA") or sorted(glob.glob(r"C:\dados_receita\*"))[-1]
Q = json.load(open(sys.argv[1], encoding="utf8"))
COL = {t.replace(" ", "") for q in Q for t in q["terms"] if len(t.replace(" ", "")) >= 6}
TIPO = {"LTDA", "ME", "EPP", "EIRELI", "SA", "LIMITADA", "MEI", "SLU", "UNIPESSOAL"}
def col(s):
    return "".join(s.translate(TBL).split())
def emp(path):
    out = {}
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            p = raw.split('";"', 2)
            if len(p) < 2: continue
            w = re.sub(r"\d{11}\s*$", "", p[1]).translate(TBL).split()
            while w and w[-1] in TIPO: w = w[:-1]
            c = "".join(w)
            if c in COL: out[p[0].strip('"')] = (p[1], c)
    return out
def est(path):
    rows = []
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            r = raw[1:9]
            if r not in ALVO:
                p = raw.split('";"', 6)
                if len(p) < 6 or p[5] != "02" or not p[4] or col(p[4]) not in COL: continue
            p = [x.strip('"') for x in raw.rstrip("\n").split('";"')]
            if len(p) < 21: continue
            rows.append((p[0], p[1], p[2], p[3], p[4], p[5], p[6], p[10], p[11], p[12], p[19], p[20]))
    return rows
def init(a):
    global ALVO
    ALVO = a
ALVO = set()
if __name__ == "__main__":
    E = sorted(glob.glob(os.path.join(DIR, "Empresas*.zip"))); S = sorted(glob.glob(os.path.join(DIR, "Estabelecimentos*.zip")))
    with Pool(4) as pool: res = pool.map(emp, E)
    em = {}
    for a in res: em.update(a)
    print("empresas", len(em), flush=True)
    with Pool(4, initializer=init, initargs=(set(em),)) as pool: res = pool.map(est, S)
    rows = [x for a in res for x in a]
    print("estab", len(rows), flush=True)
    novas = {x[0] for x in rows} - set(em)
    def _n(): pass
    COL_ALL = COL
    raz = {k: v[0] for k, v in em.items()}
    hits = [(v[1], k, "razao") for k, v in em.items()]
    hits += [(col(x[4]), x[0], "fantasia") for x in rows if x[5] == "02" and x[4] and col(x[4]) in COL]
    pickle.dump({"est": rows, "emp": raz, "novas": sorted(novas), "hits": hits}, open(sys.argv[2], "wb"))
    print("fim", len(hits))
