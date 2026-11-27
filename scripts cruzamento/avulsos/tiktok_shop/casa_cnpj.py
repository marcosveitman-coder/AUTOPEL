"""Escolhe o CNPJ de cada loja do TikTok Shop a partir da saida do busca_receita.py.
Uso: python casa_cnpj.py <tt_q.json> <tt_rec.pkl> <saida.json> <CRM.xlsx> [tt_col.pkl]
CNAE (principal ou secundario) tem de combinar com a Categoria da loja (ramo.py); senao fica "RAMO DIFERENTE".
Vale so nome IGUAL (normalizado): nome fantasia de estabelecimento ativo, ou razao social sem LTDA/ME/EIRELI/SA
(e sem o CPF do fim, no MEI). Varias raizes: fica a de CNAE de comercio/industria (46, 47, 10-33) e, se ainda
empatar, nao escolhe (AMBIGUO)."""
import sys, json, pickle, re
sys.path.insert(0, __import__("os").path.dirname(__file__))
from prep_termos import norm
from ramo import compativel
import openpyxl
wb = openpyxl.load_workbook(sys.argv[4], read_only=True); ws = wb["Leads"]
hdr = list(next(ws.iter_rows(max_row=1, values_only=True)))
CAT = {r[hdr.index("ID CRM")]: norm(r[hdr.index("Categoria")] or "") for r in ws.iter_rows(min_row=2, values_only=True) if r[0]}
wb.close()
Q = json.load(open(sys.argv[1], encoding="utf8"))
D = pickle.load(open(sys.argv[2], "rb"))
EMP, EST = D["emp"], D["est"]
TIPO = {"LTDA", "ME", "EPP", "EIRELI", "S", "A", "SA", "S A", "LIMITADA", "MEI", "SLU", "SOCIEDADE", "UNIPESSOAL"}
def core_raz(r):
    w = norm(re.sub(r"\d{11}\s*$", "", r.strip())).split()
    while w and w[-1] in TIPO: w = w[:-1]
    return " ".join(w)
by_raiz = {}
for i, row in enumerate(EST):
    by_raiz.setdefault(row[0], []).append(row)
cand = {}  # termo -> {raiz: como}
for i, t in D["hits_fant"]:
    row = EST[i]
    if row[5] == "02" and norm(row[4]) == t:
        cand.setdefault(t, {}).setdefault(row[0], set()).add("fantasia")
for r, t in D["hits_raz"]:
    if core_raz(EMP[r][0]) == t:
        cand.setdefault(t, {}).setdefault(r, set()).add("razao")
# 2a passada (busca_colado.py): nome sem espacos
if len(sys.argv) > 5:
    C = pickle.load(open(sys.argv[5], "rb"))
    for row in C["est"]:
        if row[0] not in by_raiz or row not in by_raiz[row[0]]:
            by_raiz.setdefault(row[0], []).append(tuple(row) + ("",))
    for k, v in C["emp"].items(): EMP.setdefault(k, (v, "", "", ""))
    for c, r, como in C["hits"]:
        cand.setdefault("#" + c, {}).setdefault(r, set()).add(como + " (colado)")
def _bom_cnae(c):
    c = c or ""
    return c[:2] in ("46", "47") or (c[:2].isdigit() and 10 <= int(c[:2]) <= 33)
def info(r):
    rows = by_raiz.get(r, [])
    at = [x for x in rows if x[5] == "02"]
    m = next((x for x in at if x[3] == "1"), None) or (at[0] if at else (rows[0] if rows else None))
    return m, len(at)
out = {}
for q in Q:
    c = {}
    for t in list(q["terms"]) + ["#" + t.replace(" ", "") for t in q["terms"]]:
        for r, como in cand.get(t, {}).items():
            c.setdefault(r, set()).update(como)
    ok = {}
    for r, como in c.items():
        m, nat = info(r)
        if not m or nat == 0: continue
        ok[r] = (m, como)
    if not ok:
        out[q["id"]] = {"st": "NAO ACHADO"}; continue
    cat = CAT.get(q["id"], "")
    comp = {r: v for r, v in ok.items() if compativel(cat, v[0][8], v[0][9])[0]}
    if not comp:
        out[q["id"]] = {"st": "RAMO DIFERENTE", "n": len(ok), "cands": [v[0][0]+v[0][1]+v[0][2]+" "+v[0][8] for v in ok.values()][:5]}; continue
    ok = comp
    if len(ok) > 1:
        pri = {r: v for r, v in ok.items() if compativel(cat, v[0][8], "")[1]}
        if pri: ok = pri
    if len(ok) > 1:
        out[q["id"]] = {"st": "AMBIGUO", "n": len(ok), "cands": [v[0][0] + v[0][1] + v[0][2] for v in ok.values()][:10]}
        continue
    r, (m, como) = next(iter(ok.items()))
    out[q["id"]] = {"st": "OK", "cnpj": m[0] + m[1] + m[2], "razao": EMP.get(r, ("",))[0], "fantasia": m[4],
                    "cnae": m[8], "uf": m[10], "inicio": m[7], "como": sorted(como)}
json.dump(out, open(sys.argv[3], "w", encoding="utf8"), ensure_ascii=False, indent=0)
from collections import Counter
print(Counter(v["st"] for v in out.values()), "sem termo:", 1606 - len(Q))
