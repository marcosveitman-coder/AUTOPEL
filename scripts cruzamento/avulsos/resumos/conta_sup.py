"""Lojas ativas por raiz na Receita: total e so supermercado/hiper/minimercado (CNAE 4711-3/01, 4711-3/02, 4712-1/00)."""
import os, sys, glob, zipfile, io, pickle, time
from multiprocessing import Pool
DIR = os.environ.get("AUTOPEL_RECEITA") or sorted(glob.glob(r"C:\dados_receita\*"))[-1]
HERE = os.path.dirname(os.path.abspath(__file__))
ALVO = pickle.load(open(os.path.join(HERE, "raizes_sup.pkl"), "rb"))
SUP = {"4711301", "4711302", "4712100"}
def work(path):
    t = time.time(); cnt = {}
    with zipfile.ZipFile(path) as z:
        with z.open(z.namelist()[0]) as f:
            for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
                p = raw.split(";", 12)
                raiz = p[0].strip('"')
                if raiz not in ALVO: continue
                if p[5].strip('"').lstrip("0") != "2": continue
                c = cnt.setdefault(raiz, [0, 0])
                c[0] += 1
                if p[11].strip('"') in SUP: c[1] += 1
    return os.path.basename(path), cnt, time.time() - t
if __name__ == "__main__":
    total = {}
    with Pool(4) as p:
        for name, cnt, t in p.imap_unordered(work, sorted(glob.glob(os.path.join(DIR, "Estabelecimentos*.zip")))):
            print(name, len(cnt), f"{t:.0f}s", flush=True)
            for k, v in cnt.items():
                tv = total.setdefault(k, [0, 0]); tv[0] += v[0]; tv[1] += v[1]
    pickle.dump(total, open(os.path.join(HERE, "lojas_sup.pkl"), "wb"))
    print("raizes:", len(total), "de", len(ALVO))
