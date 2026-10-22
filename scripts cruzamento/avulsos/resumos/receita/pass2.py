"""2a rodada: miolo do nome (fantasia ou razao) dentro do nome fantasia (Estabelecimentos) ou da razao social
(Empresas), so na UF da linha e so em estabelecimento com CNAE de loja."""
import os, re, glob, zipfile, io, pickle, json, time, unicodedata, sys
from multiprocessing import Pool
DIR = sorted(glob.glob(r"C:\dados_receita\*"))[-1]
HERE = os.path.dirname(os.path.abspath(__file__))
STORE = {'4711301','4711302','4712100','4729699','4691500','4639701','4639702','4693100'}
GEN = {'SUPERMERCADO','SUPERMERCADOS','SUPER','MERCADO','MERCADOS','MERCADINHO','MINIMERCADO','MERCEARIA','ATACADO','ATACADISTA',
       'ATACAREJO','COMERCIO','COMERCIAL','COM','DE','DO','DA','DOS','DAS','E','LTDA','ME','EPP','EIRELI','SA','S','A','REDE','GRUPO',
       'CIA','ALIMENTOS','ALIM','ALIMENTICIOS','PRODUTOS','GENEROS','GEN','VAREJISTA','VAREJO','HIPERMERCADO','DISTRIBUIDORA',
       'MICROEMPRESA','LIMITADA','COOP','COOPERATIVA','AUTOSSERVICO','LOJA','LOJAS','MEGASTORE','DELICATESSEN','PAD','EXPRESS'}
def toks(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().upper().replace("&", " E ")
    s = re.sub(r"\(.*?\)", " ", s)
    return re.findall(r"[A-Z0-9]+", s)
def miolo(s):
    w = [t for t in toks(s) if t not in GEN and len(t) >= 2]
    return " ".join(w)
PEND = json.load(open(os.path.join(HERE, "pend.json"), encoding="utf8"))
Q = {}   # miolo -> set(UF)
for p in PEND:
    for m in p["miolos"]:
        Q.setdefault(m, set()).add(p["uf"])
QS = sorted(Q, key=len, reverse=True)
IDX = {}
for m in QS:
    IDX.setdefault(m.split()[0], []).append(m)
def acha(nome):
    tk = toks(nome)
    if not tk: return []
    st = set(tk)
    cand = [m for w in st for m in IDX.get(w, ())]
    if not cand: return []
    t = " " + " ".join(tk) + " "
    return [m for m in cand if " " + m + " " in t]
def est(path):
    out = []
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            p = raw.split('";"')
            if len(p) < 21 or p[11] not in STORE: continue
            ms = acha(p[4])
            if not ms: continue
            uf = p[19]
            ms = [m for m in ms if uf in Q[m]]
            if ms: out.append((p[0].strip('"'), p[1], p[4], p[5], p[11], uf, p[20], ms))
    return out
MRAZ = {miolo(p["raz"]) for p in PEND if p.get("raz")} - {""}
def emp(path):
    out = []
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            p = raw.split('";"', 2)
            if len(p) < 2: continue
            m = miolo(p[1])
            if m in MRAZ: out.append((p[0].strip('"'), p[1], m))
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
        e1 = [x for l in pool.map(est, sorted(glob.glob(os.path.join(DIR, "Estabelecimentos*.zip")))) for x in l]
    print("estab por fantasia:", len(e1), f"{time.time()-t:.0f}s", flush=True)
    pickle.dump(e1, open(os.path.join(HERE, "p2_fant.pkl"), "wb"))
    with Pool(4) as pool:
        e2 = [x for l in pool.map(emp, sorted(glob.glob(os.path.join(DIR, "Empresas*.zip")))) for x in l]
    print("empresas por razao:", len(e2), f"{time.time()-t:.0f}s", flush=True)
    pickle.dump(e2, open(os.path.join(HERE, "p2_raz.pkl"), "wb"))
    alvo = {x[0] for x in e1} | {x[0] for x in e2} | set(json.load(open(os.path.join(HERE, "raizes_extra.json"))))
    with Pool(4, initializer=init, initargs=(alvo,)) as pool:
        e3 = [x for l in pool.map(est_raiz, sorted(glob.glob(os.path.join(DIR, "Estabelecimentos*.zip")))) for x in l]
    # razao social das raizes com estabelecimento
    raz = {x[0]: x[1] for x in e2}
    pickle.dump(dict(fant=e1, raz=e2, est=e3), open(os.path.join(HERE, "pass2.pkl"), "wb"))
    print("estab das raizes:", len(e3), "raizes", len(alvo), f"{time.time()-t:.0f}s")
