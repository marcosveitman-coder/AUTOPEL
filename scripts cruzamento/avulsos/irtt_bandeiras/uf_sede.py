"""UF da matriz (e do 1o estabelecimento ativo) de cada raiz, na copia local da Receita."""
import os, glob, zipfile, io, pickle, json
from multiprocessing import Pool
DIR = sorted(glob.glob(r"C:\dados_receita\*"))[-1]
HERE = os.path.dirname(os.path.abspath(__file__))
ALVO = set(json.load(open(os.path.join(HERE, "raizes_irtt.json"))))
def work(path):
    out = []
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            r = raw[1:9]
            if r not in ALVO: continue
            p = raw.split('";"')
            if len(p) < 21: continue
            out.append((r, p[1], p[3], p[5], p[19]))
    return out
if __name__ == "__main__":
    with Pool(4) as pool:
        e = [x for l in pool.map(work, sorted(glob.glob(os.path.join(DIR, "Estabelecimentos*.zip")))) for x in l]
    pickle.dump(e, open(os.path.join(HERE, "uf_irtt.pkl"), "wb")); print(len(e))
