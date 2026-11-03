"""Busca redes na copia local da Receita (C:\dados_receita\AAAA-MM).
Entrada: JSON com [{id, terms:[nomes normalizados], cnpjs:[14 digitos]}].
1) Empresas: razao social que contem um termo (sequencia de palavras inteiras) + razao das raizes dos CNPJs dados.
2) Estabelecimentos: todas as linhas das raizes achadas (ou dadas) + linhas ATIVAS cujo nome fantasia contem um termo.
3) Empresas: razao das raizes que so apareceram pelo nome fantasia.
Saida: pickle {emp: {raiz: (razao, natureza, capital, porte)}, hits_raz: [(raiz, termo)], est: [linha], hits_fant: [(idx_linha, termo)]}"""
import os, sys, glob, zipfile, io, pickle, json, time, unicodedata
import os as _os
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'dados'))
from multiprocessing import Pool
DIR = os.environ.get("AUTOPEL_RECEITA") or sorted(glob.glob(r"C:\dados_receita\*"))[-1]
SRC, OUT = sys.argv[1], sys.argv[2]
Q = json.load(open(SRC, encoding="utf8"))
def mk_table():
    t = {}
    for i in range(256):
        c = chr(i)
        a = unicodedata.normalize("NFKD", c).encode("ascii", "ignore").decode().upper()
        if c == "&": a = " E "
        elif not a or not a.isalnum(): a = " " if not (a and a.isalnum()) else a
        t[i] = a
    return t
TBL = mk_table()
TERMS = {}
for q in Q:
    for t in q["terms"]:
        TERMS.setdefault(t, set()).add(q["id"])
IDX = {}
for t in TERMS:
    w = tuple(t.split())
    IDX.setdefault(w[0], []).append(w)
# termos de 1 palavra so contam na razao social se tiverem 6+ letras
RAZ_OK = {t for t in TERMS if len(t.split()) > 1 or len(t) >= 6}
ROOTS = {c[:8] for q in Q for c in q.get("cnpjs", [])}
def acha(txt):
    tk = txt.translate(TBL).split()
    if not tk: return ()
    hits = set()
    for i, w in enumerate(tk):
        for seq in IDX.get(w, ()):
            n = len(seq)
            if tuple(tk[i:i+n]) == seq: hits.add(" ".join(seq))
    return hits
def emp(path):
    raz, hits = {}, []
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            p = raw.rstrip("\n").split('";"')
            if len(p) < 6: continue
            r = p[0].strip('"')
            h = [t for t in acha(p[1]) if t in RAZ_OK]
            if h or r in ROOTS or r in ALVO:
                raz[r] = (p[1], p[2], p[4], p[5].strip('"'))
                for t in h: hits.append((r, t))
    return raz, hits
def init(a):
    global ALVO
    ALVO = a
ALVO = set()
def est(path):
    rows, hits = [], []
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            r = raw[1:9]
            p = None
            h = ()
            if r not in ALVO:
                # so olha o fantasia de estabelecimento ativo
                p = raw.split('";"', 6)
                if len(p) < 6 or p[5] != "02" or not p[4]: continue
                h = acha(p[4])
                if not h: continue
            p = [x.strip('"') for x in raw.rstrip("\n").split('";"')]
            if len(p) < 21: continue
            rows.append((p[0], p[1], p[2], p[3], p[4], p[5], p[6], p[10], p[11], p[12], p[19], p[20], p[27] if len(p) > 27 else ""))
            if r not in ALVO or not h:
                h = acha(p[4]) if p[4] else ()
            for t in h: hits.append((len(rows) - 1, t))
    return rows, hits
if __name__ == "__main__":
    t0 = time.time()
    E = sorted(glob.glob(os.path.join(DIR, "Empresas*.zip")))
    S = sorted(glob.glob(os.path.join(DIR, "Estabelecimentos*.zip")))
    with Pool(4) as pool:
        res = pool.map(emp, E)
    raz = {}; hits_raz = []
    for a, b in res: raz.update(a); hits_raz += b
    cnt = {}
    for r, t in hits_raz: cnt[t] = cnt.get(t, 0) + 1
    # termo de razao muito generico (300+ empresas) nao puxa os estabelecimentos dessas raizes
    alvo = set(ROOTS) | {r for r, t in hits_raz if cnt[t] <= 300}
    print("empresas", len(raz), "hits", len(hits_raz), "alvo", len(alvo), f"{time.time()-t0:.0f}s", flush=True)
    with Pool(4, initializer=init, initargs=(alvo,)) as pool:
        res = pool.map(est, S)
    est_rows = []; hits_f = []
    for rows, hits in res:
        off = len(est_rows); est_rows += rows; hits_f += [(i + off, t) for i, t in hits]
    print("estab", len(est_rows), "hits fantasia", len(hits_f), f"{time.time()-t0:.0f}s", flush=True)
    novas = {x[0] for x in est_rows} - set(raz)
    with Pool(4, initializer=init, initargs=(novas,)) as pool:
        res = pool.map(emp, E)
    for a, b in res:
        for k, v in a.items(): raz.setdefault(k, v)
    pickle.dump(dict(emp=raz, hits_raz=hits_raz, est=est_rows, hits_fant=hits_f, terms={t: sorted(v) for t, v in TERMS.items()}), open(OUT, "wb"))
    print("fim", len(raz), f"{time.time()-t0:.0f}s")
