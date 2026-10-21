"""Passo 1b: junta os codigos de cliente (col. S), escolhe a data final e compara com a coluna G atual."""
import os, sys, re, glob, pickle, datetime, collections
A = os.environ["AUTOPEL_SCRIPTS"]; sys.path.insert(0, A); sys.path.insert(0, os.path.join(A, "avulsos"))
import common, match
from common import norm_legal
from irtt import SEG
AP = lambda s: re.sub(r"['’`´]", "", str(s or ""))
cod = {}
for p in glob.glob(os.path.join(os.environ["AUTOPEL_COD_DIR"], "*.pkl")):
    common.merge_into(cod, pickle.load(open(p, "rb")))
R = pickle.load(open("cand_v.pkl", "rb")); M = pickle.load(open("mapa_irtt.pkl", "rb"))
ORDER = ["codigo", "cnpj", "razao", "email", "fantasia", "desc"]
LBL = {"codigo": "Código de cliente (col. S)", "cnpj": "CNPJ validado", "razao": "Razão social (Receita)",
       "email": "Domínio de e-mail", "fantasia": "Nome da bandeira no cliente", "desc": "Produto personalizado"}
def sig(text):
    toks = [t for t in norm_legal(AP(text)).split() if len(t) >= 4 and t not in match.WEAK]
    return set(toks), "".join(norm_legal(AP(text)).split())
def related(rec, row):
    if len(rec[3]) == 14 and re.sub(r"\D", "", str(row[19] or ""))[:8] == rec[3][:8]: return True
    ct, cns = sig(rec[2])
    for ref in (row[0], row[1], row[20]):
        rt, rns = sig(ref)
        if ct & rt or any(len(t) >= 5 and t in cns for t in rt): return True
    return False
for k, row in enumerate(R["data"]):
    best = None
    for c in re.findall(r"\d+", str(row[17] or "")):
        rec = cod.get(c.zfill(6))
        if rec and row[2] != "Outros Segmentos" and not related(rec, row) and rec[4] and rec[4] not in SEG.get(row[2], set()):
            continue
        if rec and (best is None or rec[0] > best["dt"]):
            best = {"dt": rec[0], "vends": dict(rec[1]), "cliente": rec[2], "cnpj": rec[3], "seg": rec[4], "how": "codigo"}
    if best: R["cand"][k]["codigo"] = best
def choose(c):
    strong = [(m, c[m]) for m in ("codigo", "cnpj", "razao") if m in c]
    weak = [(m, c[m]) for m in ("email", "fantasia", "desc") if m in c]
    pick = lambda l: max(l, key=lambda md: (md[1]["dt"], -ORDER.index(md[0])))
    if not strong: return pick(weak) if weak else None
    best = pick(strong)
    over = [md for md in weak if md[1].get("buyer_has_brand") and md[1]["dt"] > best[1]["dt"]]
    return pick(over) if over else best
toD = lambda dt: datetime.date(dt // 10000, dt // 100 % 100, dt % 100)
res, st, dif = {}, collections.Counter(), []
for k, c in enumerate(R["cand"]):
    lin, row = M["data"][k]
    ch = choose(c); res[lin] = ch
    g = row[6].date() if hasattr(row[6], "date") else None
    if ch: st[LBL[ch[0]]] += 1
    n = toD(ch[1]["dt"]) if ch else None
    tag = ("igual" if n == g else "G vazio, achamos" if n and not g else "não achamos, G tem" if g and not n else
           "nenhum" if not g and not n else "nossa mais recente" if n > g else "nossa mais antiga")
    st["cmp: " + tag] += 1
    if tag not in ("igual", "nenhum"):
        dif.append((lin, row[1][:28], str(g), str(n), f'{ch[0]}: {ch[1]["cliente"][:34]}' if ch else ""))
pickle.dump({"res": res, "cand": R["cand"]}, open("passo1.pkl", "wb"))
for k, v in sorted(st.items()): print(f"{v:4}  {k}")
print("\nDIFERENÇAS:")
for d in dif: print("  ", d)
