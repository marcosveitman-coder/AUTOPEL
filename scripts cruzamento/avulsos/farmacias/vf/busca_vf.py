"""Busca na Receita (razao social em Empresas e nome fantasia da matriz em Estabelecimentos) os nomes das redes novas.
Saida: busca_vf.pkl {raiz: dict(razao, natureza, porte, fantasia, sit, cnae, uf, mun, cnpj)}"""
import os, glob, zipfile, io, pickle, re, time
from multiprocessing import Pool
DIR = os.path.join(os.environ["TEMP"], "receita_cnpj")
F = lambda s: s.strip().strip('"')
TERMS = ["MASTERFARMA", "MASTER FARMA", "AMR SAUDE", "FARMA JUSTA", "MELHOR COMPRA", "MULTIDROGAS", "MULTI DROGAS", "LIDER FARMA",
         "LIDERFARMA", "ECONOMIZE", "NATUS FARMA", "FARMACIA ARTESANAL", "DROGARIAS MAX", "COMPRE BEM", "CB DROGARIAS", "DROGA LIDER",
         "MINAS MASTER", "MASTER BRASIL", "CONCEITO FARMA", "FARMACIA NACIONAL", "REDE SANAR", "DROGA CLARA", "STYLOFARMA",
         "DROGASHOP", "REDEFARMA", "DROGARIA EXATA", "FARMAGENTE", "FARMAVIP", "FAZFARMA", "FARMAXIMA", "CONVIVA", "SANTA BRANCA",
         "MEDFACIL", "UNIDASFARMA", "DROGA REDE", "DROGARIA ATUAL", "OFERTAO", "VIVA MAIS", "FRANKFARMA", "REDEMED", "IN FARMA",
         "INOVFARMA", "LUCREMAIS", "MAIS VALE", "MAISFARMA", "PIX FARMA", "PULSAFARMA", "MG FARMA", "SUPER FARMA", "TCHE FARMACIAS",
         "ULTRA DESCONTAO", "WCL SAUDE", "ADMPHARMA", "HIPER SAUDE", "HIPER POPULAR", "GRUPO TOTAL", "TOTAL POPULAR", "FARMELHOR",
         "UNIFARMA", "ULTRA POPULAR", "MULTMAIS", "DROGARIAS ULTRA POPULAR", "GRUPO NR", "GRUPO MF", "GRUPO DC", "MSV"]
RX = re.compile(r"\b(" + "|".join(re.escape(t) for t in TERMS) + r")\b")


def emp(path):
    out = {}
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            p = raw.split('";"')
            if len(p) > 5 and RX.search(p[1]):
                out[F(p[0])] = (F(p[1]), F(p[2]), F(p[5]))
    return out


ROOTS = set()


def init(r):
    global ROOTS; ROOTS = r


def est(path):
    out = []
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            p = raw.split('";"')
            if len(p) < 28 or F(p[3]) != "1":
                continue
            r = F(p[0])
            if r in ROOTS or RX.search(p[4]):
                out.append(dict(root=r, cnpj=r + F(p[1]) + F(p[2]), fantasia=F(p[4]), sit=F(p[5]).lstrip("0"), cnae=F(p[11]),
                                uf=F(p[19]), mun=F(p[20])))
    return out


if __name__ == "__main__":
    t = time.time(); E = {}
    with Pool(4) as pool:
        for o in pool.imap_unordered(emp, sorted(glob.glob(os.path.join(DIR, "Empresas*.zip")))): E.update(o)
    print("razoes:", len(E), f"{time.time()-t:.0f}s", flush=True)
    M = []
    with Pool(4, initializer=init, initargs=(set(E),)) as pool:
        for o in pool.imap_unordered(est, sorted(glob.glob(os.path.join(DIR, "Estabelecimentos*.zip")))): M += o
    print("matrizes:", len(M), f"{time.time()-t:.0f}s", flush=True)
    pickle.dump({"emp": E, "mat": M}, open("busca_vf.pkl", "wb"))
