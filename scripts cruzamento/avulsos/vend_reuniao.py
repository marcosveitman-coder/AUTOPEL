import os, sys, re, pickle, collections, datetime
import pandas as pd, openpyxl
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from build import choose, to_date
from common import pick_vendor
D = lambda s: re.sub(r"\D", "", str(s or ""))
CARTEIRA = {"INATIVOS", "CLIENTE SEM COMPRA", "RD", "REPRESENTANTE", ""}

# codigo -> nome do vendedor
uf = pd.read_csv(os.path.join(common.BASE, "Ultimo faturamento - SQL.csv"), sep=";", skiprows=1, dtype=str, encoding="utf-8-sig")
uf.columns = [c.strip() for c in uf.columns]
vname = {}
for c, n in zip(uf["Codigo"].fillna("").str.strip(), uf.iloc[:, 15].fillna("").str.strip()):
    if c and n: vname.setdefault(c, n)
# cadastro: CNPJ -> codigo vendedor ; raiz -> contagem
cs = pd.read_csv(os.environ.get("AUTOPEL_CLIENTES_CSV", "Clientes - SQL.csv"), sep=";", skiprows=1, dtype=str, encoding="utf-8-sig")
cs.columns = [c.strip() for c in cs.columns]
by_cnpj, by_root = {}, collections.defaultdict(collections.Counter)
for cnpj, v in zip(cs["CNPJ/CPF"].map(D), cs["Vendedor"].fillna("").str.strip()):
    if len(cnpj) != 14 or not v: continue
    by_cnpj.setdefault(cnpj, v)
    by_root[cnpj[:8]][v] += 1
real = lambda code: vname.get(code, "") not in CARTEIRA and code != "000000"

def cad_vendor(cnpj14):
    """vendedor do cadastro: CNPJ exato; senao o vendedor mais comum da raiz (sem carteiras)"""
    if cnpj14 in by_cnpj and real(by_cnpj[cnpj14]):
        return vname.get(by_cnpj[cnpj14]), "cadastro (CNPJ)"
    if len(cnpj14) >= 8:
        for code, _ in by_root.get(cnpj14[:8], collections.Counter()).most_common():
            if real(code):
                return vname.get(code), "cadastro (raiz)"
    return None, None

cand = pickle.load(open("cand_reuniao.pkl", "rb"))
rows = pickle.load(open("reuniao_rows.pkl", "rb"))
ws = openpyxl.load_workbook("backup_mapa_mercado.xlsx", read_only=True)["SUPERMERCADOS (reuniao)"]
sheet = {r: row for r, row in enumerate(ws.iter_rows(values_only=True), start=1)}
out = {}
for k, c in enumerate(cand["cand"]):
    r = rows[k]; row = sheet[r]
    ch = choose(c)
    fat_v = pick_vendor(ch[1]["vends"]) if ch else None
    fat_cnpj = ch[1]["cnpj"] if ch else ""
    cad_v, cad_how = cad_vendor(fat_cnpj) if fat_cnpj else (None, None)
    if not cad_v:
        cad_v, cad_how = cad_vendor(D(row[5]))
    out[r] = dict(fat=fat_v, fat_dt=ch[1]["dt"] if ch else None, fat_cli=ch[1]["cliente"] if ch else None,
                  cad=cad_v, cad_how=cad_how, ag=row[32], crm=row[30])
pickle.dump(out, open("vend_reuniao.pkl", "wb"))
