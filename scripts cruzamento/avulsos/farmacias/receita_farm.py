"""Le Estabelecimentos*.zip e Empresas*.zip (Receita, dados abertos).
Estabelecimentos: guarda farmacias (CNAE principal 47717xx, qualquer situacao) + qualquer estabelecimento de raiz da aba FARMACIA.
Empresas: razao social / porte / capital das raizes guardadas.   Saida: rf_estab.pkl, rf_emp.pkl"""
import os, glob, zipfile, io, pickle, time, sys
from multiprocessing import Pool
DIR = os.path.join(os.environ["TEMP"], "receita_cnpj")
HERE = os.path.dirname(os.path.abspath(__file__))
ALVO = pickle.load(open(os.path.join(HERE, "raizes_farmacia_aba.pkl"), "rb"))
F = lambda s: s.strip().strip('"')

def estab(path):
    out = []
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            p = raw.split('";"')
            if len(p) < 28:
                continue
            root = F(p[0]); cnae = F(p[11])
            if not (cnae.startswith("47717") or root in ALVO):
                continue
            out.append((root, root + F(p[1]) + F(p[2]), F(p[3]), F(p[4]), F(p[5]).lstrip("0"), F(p[6]), cnae, F(p[19]),
                        F(p[20]), F(p[27]).lower(), F(p[12])))
    return os.path.basename(path), out

import re
ASSOC = re.compile(r"(ASSOC|COOPERATIV|REDE|FEDERA|FRANQUIA|ADMINISTRADORA|FARMARCAS|FEBRAFAR|CENTRAL DE (COMPRAS|NEGOCIOS)).*(FARMA|DROG|MEDICAMENT)|"
                   r"(FARMA|DROG|MEDICAMENT).*(ASSOC|COOPERATIV|FEDERA|FRANQUIA|ADMINISTRADORA|CENTRAL DE (COMPRAS|NEGOCIOS))|FARMARCAS|FEBRAFAR")

def emp(args):
    path, roots = args
    out, assoc = {}, {}
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            p = raw.split('";"')
            r = F(p[0])
            if r in roots:
                out[r] = (F(p[1]), F(p[2]), F(p[4]), F(p[5]) if len(p) > 5 else "")
            elif ASSOC.search(p[1] if len(p) > 1 else ""):
                assoc[r] = (F(p[1]), F(p[2]), F(p[4]), F(p[5]) if len(p) > 5 else "")
    return os.path.basename(path), out, assoc

def estab_assoc(args):
    path, roots = args
    out = []
    with zipfile.ZipFile(path) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding="latin-1", newline=""):
            p = raw.split('";"')
            if len(p) < 28 or F(p[0]) not in roots:
                continue
            out.append((F(p[0]), F(p[0]) + F(p[1]) + F(p[2]), F(p[3]), F(p[4]), F(p[5]).lstrip("0"), F(p[6]), F(p[11]),
                        F(p[19]), F(p[20]), F(p[27]).lower(), F(p[12])))
    return os.path.basename(path), out

if __name__ == "__main__":
    t = time.time()
    rows = []
    with Pool(4) as pool:
        for name, out in pool.imap_unordered(estab, sorted(glob.glob(os.path.join(DIR, "Estabelecimentos*.zip")))):
            rows += out; print(name, len(out), f"{time.time()-t:.0f}s", flush=True)
    pickle.dump(rows, open(os.path.join(HERE, "rf_estab.pkl"), "wb"))
    roots = {r[0] for r in rows}
    print("estabelecimentos guardados:", len(rows), "| raizes:", len(roots), flush=True)
    emp_, assoc = {}, {}
    with Pool(4) as pool:
        for name, out, ass in pool.imap_unordered(emp, [(p, roots) for p in sorted(glob.glob(os.path.join(DIR, "Empresas*.zip")))]):
            emp_.update(out); assoc.update(ass); print(name, len(out), len(ass), f"{time.time()-t:.0f}s", flush=True)
    pickle.dump(emp_, open(os.path.join(HERE, "rf_emp.pkl"), "wb"))
    print("empresas:", len(emp_), "| associacoes/redes:", len(assoc), flush=True)
    est_ass = []
    with Pool(4) as pool:
        for name, out in pool.imap_unordered(estab_assoc, [(p, set(assoc)) for p in sorted(glob.glob(os.path.join(DIR, "Estabelecimentos*.zip")))]):
            est_ass += out
    pickle.dump({"emp": assoc, "estab": est_ass}, open(os.path.join(HERE, "rf_assoc.pkl"), "wb"))
    print("estab. de associacoes:", len(est_ass), f"TOTAL {time.time()-t:.0f}s")
