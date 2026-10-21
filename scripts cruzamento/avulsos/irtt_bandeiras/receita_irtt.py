"""Busca na Receita (dados abertos do CNPJ) as bandeiras do IRTT da aba "I.R.T.T.__Por bandeira (FINAL)".
1) Empresas*.zip: razao social de toda raiz que tenha a marca da bandeira/grupo no nome + razao das raizes indicadas (col U).
2) Estabelecimentos*.zip (so ATIVOS): por (marca, raiz) quantos estabelecimentos tem a marca no nome fantasia e em quais
   CNAEs; para as raizes de interesse, total de ativos, nomes fantasia mais comuns, UF e CNAE da matriz.
Saida: receita_irtt.pkl.  Rodar na pasta do mapa_irtt.pkl (gerado a partir da aba)."""
import os, re, sys, glob, zipfile, io, pickle, time, collections
from multiprocessing import Pool
REPO = os.environ.get("AUTOPEL_SCRIPTS", os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", ".."))
sys.path.insert(0, REPO); sys.path.insert(0, os.path.join(REPO, "avulsos"))
from common import norm_legal

DIR = os.environ.get("AUTOPEL_RECEITA") or sorted(glob.glob("C:/dados_receita/20*"))[-1]  # copia local da Receita
F = lambda s: s.strip().strip('"')
AP = lambda s: re.sub(r"['’`´]", "", str(s or ""))
KEYS = PREF = ROOTS = None
MANUAIS = {"C&A": {"C A MODAS", "CEA MODAS", "C E A MODAS"}, "Dia %": {"DIA BRASIL"}}  # nomes que viram so simbolos


def chaves():
    from irtt import brand_keys
    M = pickle.load(open("mapa_irtt.pkl", "rb"))
    por_band, keys = [], set()
    for _, r in M["data"]:
        ks = set()
        for txt in (r[1], re.sub(r"(?i)\bgrupo\b", "", str(r[0] or ""))):
            cores, bare, glued = brand_keys(txt)
            ks |= {c for c in cores | glued}
            n = norm_legal(AP(txt))
            if len(n.replace(" ", "")) >= 4:  # nome inteiro (bandeiras curtas: Dia, Big, Oba...)
                ks.add(n)
        ks |= MANUAIS.get(str(r[1]).strip(), set())
        por_band.append(ks); keys |= ks
    roots = {re.sub(r"\D", "", str(r[20] or ""))[:8] for _, r in M["data"]}
    return por_band, keys, {x for x in roots if len(x) == 8}


def init(k, p, r):
    global KEYS, PREF, ROOTS
    KEYS, PREF, ROOTS = k, p, r


def hits(text):
    toks = norm_legal(AP(text)).split()
    out = set()
    for a, t in enumerate(toks):
        if t not in PREF:
            continue
        sp, gl = t, t
        if sp in KEYS: out.add(sp)
        for b in range(a + 1, min(a + 5, len(toks))):
            sp += " " + toks[b]; gl += toks[b]
            if sp in KEYS: out.add(sp)
            if gl in KEYS: out.add(gl)
    return out


def emp(path):
    out = {}
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            p = raw.split('";"')
            if len(p) < 6:
                continue
            r = F(p[0])
            if r in ROOTS or hits(p[1]):
                out[r] = (F(p[1]), F(p[2]), F(p[5]))
    return out


def est(path):
    kr, kc, info = collections.Counter(), collections.Counter(), {}
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            p = raw.split('";"')
            if len(p) < 28 or F(p[5]) != "02":
                continue
            r, fant, cnae = F(p[0]), F(p[4]), F(p[11])
            for k in hits(fant):
                kr[(k, r)] += 1; kc[(k, r, cnae[:4])] += 1
            if r in ROOTS:
                d = info.setdefault(r, {"ativos": 0, "fant": collections.Counter(), "uf": collections.Counter(), "matriz": None})
                d["ativos"] += 1; d["fant"][fant] += 1; d["uf"][F(p[19])] += 1
                if F(p[3]) == "1":
                    d["matriz"] = (r + F(p[1]) + F(p[2]), fant, cnae, F(p[19]))
    for d in info.values():
        d["fant"] = d["fant"].most_common(15)
    return kr, kc, info


if __name__ == "__main__":
    t = time.time()
    por_band, keys, roots = chaves()
    pref = {k.split()[0] for k in keys} | {k[:n] for k in keys if " " not in k for n in range(2, len(k) + 1)}
    print("bandeiras:", len(por_band), "| chaves:", len(keys), "| raizes indicadas:", len(roots), flush=True)
    E = {}
    with Pool(4, initializer=init, initargs=(keys, pref, roots)) as pool:
        for o in pool.imap_unordered(emp, sorted(glob.glob(os.path.join(DIR, "Empresas*.zip")))):
            E.update(o)
    print("empresas:", len(E), f"{time.time()-t:.0f}s", flush=True)
    # raizes de interesse: indicadas + as que tem a marca na razao social
    roots2 = roots | set(E)
    KR, KC, INFO = collections.Counter(), collections.Counter(), {}
    with Pool(4, initializer=init, initargs=(keys, pref, roots2)) as pool:
        for kr, kc, info in pool.imap_unordered(est, sorted(glob.glob(os.path.join(DIR, "Estabelecimentos*.zip")))):
            KR.update(kr); KC.update(kc)
            for r, d in info.items():
                if r in INFO:
                    x = INFO[r]; x["ativos"] += d["ativos"]; x["uf"].update(d["uf"])
                    x["fant"] = (collections.Counter(dict(x["fant"])) + collections.Counter(dict(d["fant"]))).most_common(15)
                    x["matriz"] = x["matriz"] or d["matriz"]
                else:
                    INFO[r] = d
            print("estab ok", f"{time.time()-t:.0f}s", flush=True)
    pickle.dump({"por_band": por_band, "emp": E, "kr": dict(KR), "kc": dict(KC), "info": INFO},
                open("receita_irtt.pkl", "wb"))
    print("fim", f"{time.time()-t:.0f}s")
