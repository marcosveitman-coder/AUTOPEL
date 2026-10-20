"""Conta estabelecimentos por raiz de CNPJ nos arquivos Estabelecimentos*.zip da Receita (dados abertos).
Saida: filiais.pkl {raiz: [ativos_total, ativas_filiais, todos, ativos_matriz]} so para as raizes alvo."""
import os, sys, re, glob, zipfile, io, pickle, time
from multiprocessing import Pool

DIR = os.path.join(os.environ["TEMP"], "receita_cnpj")
HERE = os.path.dirname(os.path.abspath(__file__))
ALVO = pickle.load(open(os.path.join(HERE, "raizes_alvo.pkl"), "rb"))


def work(path):
    t = time.time()
    cnt = {}
    n = 0
    with zipfile.ZipFile(path) as z:
        name = z.namelist()[0]
        with z.open(name) as f:
            for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
                n += 1
                p = raw.split(";", 6)
                raiz = p[0].strip('"')
                if raiz not in ALVO:
                    continue
                ident = p[3].strip('"')
                sit = p[5].strip('"').lstrip("0")
                c = cnt.setdefault(raiz, [0, 0, 0, 0])
                c[2] += 1
                if sit == "2":
                    c[0] += 1
                    if ident == "2":
                        c[1] += 1
                    else:
                        c[3] += 1
    return os.path.basename(path), n, cnt, time.time() - t


if __name__ == "__main__":
    files = sorted(glob.glob(os.path.join(DIR, "Estabelecimentos*.zip")))
    total = {}
    with Pool(4) as p:
        for name, n, cnt, t in p.imap_unordered(work, files):
            print(f"{name}: {n} linhas, {len(cnt)} raizes alvo, {t:.0f}s", flush=True)
            for k, v in cnt.items():
                tv = total.setdefault(k, [0, 0, 0, 0])
                for j in range(4):
                    tv[j] += v[j]
    pickle.dump(total, open(os.path.join(HERE, "filiais.pkl"), "wb"))
    print("raizes encontradas:", len(total), "de", len(ALVO))
