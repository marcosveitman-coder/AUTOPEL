import os, sys, re, io, pickle, runpy, contextlib, collections
import pandas as pd, openpyxl
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from common import pick_vendor
from choose_irtt import to_date
D = lambda s: re.sub(r"\D", "", str(s or ""))
CARTEIRA = {"INATIVOS", "CLIENTE SEM COMPRA", "RD", "REPRESENTANTE", ""}

with contextlib.redirect_stdout(io.StringIO()):
    g = runpy.run_path("write_irtt.py")          # escolhe a ultima compra de cada bandeira (sem gravar)
res = g["res"]                                    # linha do IRTT (2..571) -> (metodo, registro)

uf = pd.read_csv(os.path.join(common.BASE, "Ultimo faturamento - SQL.csv"), sep=";", skiprows=1, dtype=str, encoding="utf-8-sig")
uf.columns = [c.strip() for c in uf.columns]
vname = {}
for c, n in zip(uf["Codigo"].fillna("").str.strip(), uf.iloc[:, 15].fillna("").str.strip()):
    if c and n: vname.setdefault(c, n)
cs = pd.read_csv(os.environ.get("AUTOPEL_CLIENTES_CSV", "Clientes - SQL.csv"), sep=";", skiprows=1, dtype=str, encoding="utf-8-sig")
cs.columns = [c.strip() for c in cs.columns]
by_cnpj, by_root, by_cod = {}, collections.defaultdict(collections.Counter), collections.defaultdict(collections.Counter)
for cod, cnpj, v in zip(cs["Codigo"].fillna("").str.strip(), cs["CNPJ/CPF"].map(D), cs["Vendedor"].fillna("").str.strip()):
    if not v: continue
    if cod: by_cod[cod.zfill(6)][v] += 1
    if len(cnpj) == 14:
        by_cnpj.setdefault(cnpj, v); by_root[cnpj[:8]][v] += 1
real = lambda code: code != "000000" and vname.get(code, "") not in CARTEIRA
top = lambda cnt: next((vname[c] for c, _ in cnt.most_common() if real(c)), None)

ws = openpyxl.load_workbook("backup_mapa_mercado_v2.xlsx", read_only=True)["I.R.T.T.__Por bandeira (FINAL)"]
sheet = {r: row for r, row in enumerate(ws.iter_rows(values_only=True), start=1)}
irtt = {r: row for r, row in enumerate(openpyxl.load_workbook(os.environ.get("AUTOPEL_IRTT", "IRTT (claude).xlsx"), read_only=True).active.iter_rows(values_only=True), start=1)}
out, check = {}, collections.Counter()
for r in range(5, 575):
    row = sheet.get(r)
    if not row or not row[1]: continue
    ir = r - 3
    check["bandeira igual ao IRTT" if irtt[ir][1] == row[1] else "BANDEIRA DIFERENTE"] += 1
    ch = res.get(ir)
    if ch and row[6] is not None:
        check["data G igual" if to_date(ch[1]["dt"]) == row[6] else "DATA G DIFERENTE"] += 1
    v, how = None, None
    if ch and len(ch[1]["cnpj"]) == 14:
        c14 = ch[1]["cnpj"]
        if c14 in by_cnpj and real(by_cnpj[c14]): v, how = vname[by_cnpj[c14]], "cliente da última compra (CNPJ)"
        elif top(by_root.get(c14[:8], collections.Counter())): v, how = top(by_root[c14[:8]]), "cliente da última compra (raiz)"
    if not v:
        cnt = collections.Counter()
        for c in re.findall(r"\d+", str(row[18] or "")): cnt.update(by_cod.get(c.zfill(6), {}))
        if top(cnt): v, how = top(cnt), "códigos de cliente (col S)"
    if not v:
        m = re.match(r"^\s*(\d{2})\.?(\d{3})\.?(\d{3})", str(row[20] or ""))
        if m and top(by_root.get("".join(m.groups()), collections.Counter())): v, how = top(by_root["".join(m.groups())]), "CNPJ indicado (col U)"
    if not v and ch:
        pv = pick_vendor(ch[1]["vends"])
        if pv: v, how = pv, "vendedor da última venda (sem cadastro)"
    out[r] = (v, how, row[1], row[12])
pickle.dump(out, open("vend_irtt.pkl", "wb"))
print(dict(check))
print(collections.Counter(h for v, h, b, m in out.values()))
print("igual ao 'Vendedor principal' (col M):", sum(1 for v, h, b, m in out.values() if v and m and v[:30] == str(m)[:30]),
      "| diferente:", sum(1 for v, h, b, m in out.values() if v and m and v[:30] != str(m)[:30]))
for r, (v, h, b, m) in list(out.items())[:12]: print("  ", r, b[:28], "|", v, "|", h, "| col M:", m)
