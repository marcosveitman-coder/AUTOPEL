"""Acha redes pela propria Receita, por regiao: conta estabelecimentos ATIVOS por raiz de CNPJ e por nome fantasia,
so com CNAE principal de pet (PET 4789-0/04 pet shop; VET 7500-1/00 clinica/hospital veterinario; BANHO 9609-2/08
banho e tosa; HOTEL 9609-2/07 hotel/creche de animais; VETMED 4771-7/04 farmacia veterinaria) e de estacionamento
(ESTAC 5223-1/00). Base das abas PET SHOP  FINAL e ESTACIONAMENTO FINAL. Uso: python redes_pet_estac.py r11.pkl
Saida: raizes com muitas lojas proprias e nomes fantasia repetidos em muitas raizes (franquias)."""
import os as _os
_AQUI = _os.path.dirname(_os.path.abspath(__file__))
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_AQUI, 'dados'))
import os, sys, glob, zipfile, io, pickle, time, unicodedata, collections
from multiprocessing import Pool
DIR = os.environ.get("AUTOPEL_RECEITA") or sorted(glob.glob(r"C:\dados_receita\*"))[-1]
OUT = sys.argv[1]
SEG = {"4789004": "PET", "7500100": "VET", "9609208": "BANHO", "9609207": "HOTEL", "4771704": "VETMED", "5223100": "ESTAC"}
def mk():
    t = {}
    for i in range(256):
        c = chr(i); a = unicodedata.normalize("NFKD", c).encode("ascii", "ignore").decode().upper()
        t[i] = " E " if c == "&" else (a if a and a.isalnum() else " ")
    return t
TBL = mk()
def linhas(path):
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            p = raw.split('";"', 22)
            if len(p) < 21 or p[5] != "02": continue
            s = SEG.get(p[11])
            if not s: continue
            yield p[0].strip('"'), s, p[11], p[19], " ".join(p[4].translate(TBL).split())
def conta(path):
    cr = collections.Counter(); cf = collections.Counter()
    for r, s, cn, uf, fn in linhas(path):
        cr[(r, s)] += 1
        if fn: cf[(fn, s)] += 1
    return cr, cf
def init(a, b):
    global RS, FS
    RS, FS = a, b
def detalhe(path):
    R = {}; F = {}
    for r, s, cn, uf, fn in linhas(path):
        k = (r, s)
        if k in RS:
            x = R.setdefault(k, [collections.Counter(), collections.Counter(), collections.Counter()])
            x[0][cn] += 1; x[1][uf] += 1
            if fn: x[2][fn] += 1
        if fn and (fn, s) in FS:
            y = F.setdefault((fn, s), [set(), collections.Counter(), collections.Counter()])
            y[0].add(r); y[1][uf] += 1; y[2][cn] += 1
    return R, F
if __name__ == "__main__":
    t0 = time.time()
    S = sorted(glob.glob(os.path.join(DIR, "Estabelecimentos*.zip")))
    with Pool(4) as pool:
        res = pool.map(conta, S)
    CR = collections.Counter(); CF = collections.Counter()
    for a, b in res: CR.update(a); CF.update(b)
    TH = {"PET": 3, "VET": 3, "BANHO": 3, "HOTEL": 3, "VETMED": 3, "ESTAC": 3}
    RS = {k for k, n in CR.items() if n >= TH[k[1]]}
    FS = {k for k, n in CF.items() if n >= 5}
    print("raizes", len(RS), "fantasias", len(FS), f"{time.time()-t0:.0f}s", flush=True)
    with Pool(4, initializer=init, initargs=(RS, FS)) as pool:
        res = pool.map(detalhe, S)
    R = {}; F = {}
    for a, b in res:
        for k, v in a.items():
            o = R.setdefault(k, [collections.Counter(), collections.Counter(), collections.Counter()])
            for i in range(3): o[i].update(v[i])
        for k, v in b.items():
            o = F.setdefault(k, [set(), collections.Counter(), collections.Counter()])
            o[0] |= v[0]; o[1].update(v[1]); o[2].update(v[2])
    R = {k: (dict(v[0]), dict(v[1]), v[2].most_common(8), sum(v[0].values())) for k, v in R.items()}
    F = {k: (CF[k], len(v[0]), dict(v[1]), dict(v[2]), sorted(v[0])) for k, v in F.items()}
    pickle.dump(dict(raiz=R, fant=F), open(OUT, "wb"))
    print("fim", len(R), len(F), f"{time.time()-t0:.0f}s")
