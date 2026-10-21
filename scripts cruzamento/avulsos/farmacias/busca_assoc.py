"""Busca ampla (qualquer CNAE) de razao social/nome fantasia com nomes de redes sem CNPJ de associacao."""
import os, glob, zipfile, io, pickle, re, time, sys
from multiprocessing import Pool
DIR = os.path.join(os.environ["TEMP"], "receita_cnpj")
F = lambda s: s.strip().strip('"')
TERMS = ["FARMARCAS", "FARMACIAS ASSOCIADAS", "PROPRIETARIOS E OFICIAIS", "AUGEFARMA", "AUGE FARMA", "HIPERFARMA", "REDEFARMES",
         "COOPFARMES", "REDE FARMES", "DESCONTO FACIL", "ASFAR", "UNIFARMA", "ULTRA POPULAR", "ULTRAPOPULAR", "SUPER POPULAR",
         "MAXI POPULAR", "MEGA POPULAR", "BIGFORT", "MEGA PHARMA", "MEGAPHARMA", "INOVA DROGARIAS", "REDE INOVA", "RM FARMA",
         "DSG FARMA", "FARMELHOR", "ULTRAFARMA", "DROGARIA TOTAL", "TOTAL POPULAR", "FARMA E FARMA", "CAPILE", "GRUPOFARMA",
         "CITYFARMA", "LEGITIMA", "PRECO MAIS POPULAR", "COMPRE CERTO", "DESCONTAO", "REDE SOMA", "SUPERMAIS", "UNIVERSAL FARMA",
         "MOVA MARCAS", "MULTIFARMA", "USIFARMA", "SMALL FARMA", "SMALLFARMA", "REDE 2000", "POUPAQUI", "ACOFARMA", "PRECO JUSTO",
         "OTIMA FARMA", "HIPERFARMAC", "FARMACIA DO TRABALHADOR", "BIFARMA", "PEDPHARMA", "REDEBELLA", "MAESTRA", "FARMAVALE",
         "CONFARMA", "BOA FARMA", "BOAFARMA", "REDE BEM", "FORTE FARMA", "FORTEFARMA", "COPERFARMA", "THE FARMA"]
RX = re.compile(r"\b(" + "|".join(re.escape(t) for t in TERMS) + r")\b")

def emp(path):
    out = {}
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            p = raw.split('";"')
            if len(p) > 1 and RX.search(p[1]):
                out[F(p[0])] = (F(p[1]), F(p[2]))
    return out

def est(path):
    out = []
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            p = raw.split('";"')
            if len(p) >= 28 and F(p[3]) == "1" and (RX.search(p[4]) or F(p[0]) in ROOTS):
                out.append((F(p[0]), F(p[0]) + F(p[1]) + F(p[2]), F(p[4]), F(p[5]).lstrip("0"), F(p[11]), F(p[19]), F(p[20]), F(p[27]).lower()))
    return out

ROOTS = set()
if __name__ == "__main__":
    t = time.time(); E = {}
    with Pool(4) as pool:
        for o in pool.imap_unordered(emp, sorted(glob.glob(os.path.join(DIR, "Empresas*.zip")))): E.update(o)
    print("empresas por razao:", len(E), f"{time.time()-t:.0f}s", flush=True)
    pickle.dump(E, open("busca_emp.pkl", "wb"))
