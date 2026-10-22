"""3a rodada: razao social com iniciais preservadas e sem o CPF/numero do fim (empresario individual)."""
import os, re, glob, zipfile, io, pickle, json, time, unicodedata
from multiprocessing import Pool
DIR = sorted(glob.glob(r"C:\dados_receita\*"))[-1]
HERE = os.path.dirname(os.path.abspath(__file__))
GEN = {'SUPERMERCADO','SUPERMERCADOS','SUPER','MERCADO','MERCADOS','MERCADINHO','MINIMERCADO','MERCEARIA','ATACADO','ATACADISTA',
       'ATACAREJO','COMERCIO','COMERCIAL','COM','DE','DO','DA','DOS','DAS','E','LTDA','ME','EPP','EIRELI','SA','REDE','GRUPO',
       'CIA','ALIMENTOS','ALIM','ALIMENTICIOS','PRODUTOS','PROD','GENEROS','GEN','VAREJISTA','VAREJO','DISTRIBUIDORA','MICROEMPRESA',
       'LIMITADA','AUTOSSERVICO','LOJA','PAD','PANIFICADORA','DELICATESSEN'}
def miolo3(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().upper().replace("&", " E ")
    w = re.findall(r"[A-Z0-9]+", s)
    while w and w[-1].isdigit(): w.pop()
    return " ".join(t for t in w if t not in GEN)
ALVOS = json.load(open(os.path.join(HERE, "pend3.json"), encoding="utf8"))
M = {miolo3(p["raz"]): p["lin"] for p in ALVOS} | {miolo3(p["fant"]): p["lin"] for p in ALVOS}
M.pop("", None)
def emp(path):
    out = []
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            p = raw.split('";"', 2)
            if len(p) < 2: continue
            m = miolo3(p[1])
            if m in M: out.append((p[0].strip('"'), p[1], m))
    return out
def init(a):
    global ALVO
    ALVO = a
def est_raiz(path):
    out = []
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            r = raw[1:9]
            if r not in ALVO: continue
            p = raw.split('";"')
            if len(p) < 21: continue
            out.append((r, p[1], p[4], p[5], p[11], p[19], p[20]))
    return out
if __name__ == "__main__":
    t = time.time()
    with Pool(4) as pool:
        e2 = [x for l in pool.map(emp, sorted(glob.glob(os.path.join(DIR, "Empresas*.zip")))) for x in l]
    print("empresas:", len(e2), f"{time.time()-t:.0f}s", flush=True)
    alvo = {x[0] for x in e2}
    with Pool(4, initializer=init, initargs=(alvo,)) as pool:
        e3 = [x for l in pool.map(est_raiz, sorted(glob.glob(os.path.join(DIR, "Estabelecimentos*.zip")))) for x in l]
    pickle.dump(dict(raz=e2, est=e3, M=M), open(os.path.join(HERE, "pass3.pkl"), "wb"))
    print("estab:", len(e3), "raizes", len(alvo), f"{time.time()-t:.0f}s")
