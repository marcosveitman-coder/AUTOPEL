"""Para as linhas sem nº de lojas: (1) acha a raiz pela razao social exata (Empresas*.zip) das linhas sem CNPJ;
(2) lista todos os estabelecimentos das raizes (Estabelecimentos*.zip): ordem, matriz/filial, situacao, CNAE, UF, municipio, fantasia."""
import os, re, glob, zipfile, io, pickle, json, time, unicodedata
from multiprocessing import Pool
DIR = os.environ.get("AUTOPEL_RECEITA") or sorted(glob.glob(r"C:\dados_receita\*"))[-1]
HERE = os.path.dirname(os.path.abspath(__file__))
SUF = {"LTDA", "SA", "S", "A", "EIRELI", "ME", "EPP", "LIMITADA", "CIA", "COMPANHIA"}
def norm(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().upper().replace("&", " E ")
    w = re.findall(r"[A-Z0-9]+", s)
    while w and w[-1] in SUF: w.pop()
    return " ".join(w)
L0 = json.load(open(os.path.join(HERE, "..", "l0.json"), encoding="utf8"))
NOMES = {norm(o["raz"]) for o in L0 if not o["cnpj"] and o["raz"]} | {norm(o["fant"]) for o in L0 if not o["cnpj"] and o["fant"]}
NOMES.discard("")
ROOTS = {o["cnpj"][:8] for o in L0 if o["cnpj"]}
def emp(path):
    out = []
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            p = raw.split(";", 3)
            n = norm(p[1].strip('"'))
            if n in NOMES: out.append((p[0].strip('"'), n, p[1].strip('"')))
    return out
def init(a):
    global ALVO
    ALVO = a
def est(path):
    out = []
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            r = raw[:9].strip('"')
            if r not in ALVO: continue
            p = [x.strip('"') for x in raw.rstrip("\n").split('";"')]
            out.append((p[0].strip('"'), p[1], p[3], p[4], p[5], p[11], p[19], p[20]))
    return out
if __name__ == "__main__":
    t = time.time()
    with Pool(4) as pool:
        cand = [x for lst in pool.map(emp, sorted(glob.glob(os.path.join(DIR, "Empresas*.zip")))) for x in lst]
    print("empresas com razao igual:", len(cand), f"{time.time()-t:.0f}s", flush=True)
    pickle.dump(cand, open(os.path.join(HERE, "cand_razao.pkl"), "wb"))
    ALVO = ROOTS | {c[0] for c in cand}
    pickle.dump(ALVO, open(os.path.join(HERE, "alvo.pkl"), "wb"))
    with Pool(4, initializer=init, initargs=(ALVO,)) as pool:
        ests = [x for lst in pool.map(est, sorted(glob.glob(os.path.join(DIR, "Estabelecimentos*.zip")))) for x in lst]
    pickle.dump(ests, open(os.path.join(HERE, "estab_l0.pkl"), "wb"))
    print("estabelecimentos:", len(ests), "raizes alvo", len(ALVO), f"{time.time()-t:.0f}s")
