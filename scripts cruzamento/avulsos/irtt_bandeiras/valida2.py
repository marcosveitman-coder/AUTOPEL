"""Passo 2 (v2): classifica o CNPJ da col. U e escolhe o candidato da Receita quando U nao confere ou esta vazio."""
import pickle, re, collections, sys, os
sys.path.insert(0, os.environ["AUTOPEL_SCRIPTS"])
from common import norm_legal
import match
exec(open("valida.py", encoding="utf-8").read().split("out = []")[0])  # reaproveita cargas e funcoes
def sig(t):
    toks = norm_legal(AP(t)).split()
    return {x for x in toks if len(x) >= 4 and x not in match.WEAK and not x.isdigit()}, "".join(toks)
GRP_OK = {  # operadoras conhecidas do grupo com razao social sem o nome da bandeira
    "COMPANHIA BRASILEIRA DE DISTRIBUICAO", "RAIZEN", "DIA BRASIL", "LUXOTTICA", "NS2 COM", "TREND FOODS", "SPT FRANCHISING",
    "FRM FRANQUIA", "UNI CO", "VENBO", "OUTBACK", "WMS SUPERMERCADOS", "VF DO BRASIL", "GUARANA BRASIL", "NIBS PARTICIPACOES",
    "E VINO", "HNT COMERCIO", "PORTOBELLO", "SBF COMERCIO", "ADICAO DISTRIBUICAO"}
def razao_ok(rz, band, grupo):
    rs, rg = sig(rz)
    for t in (band, grupo):
        ts, tg = sig(t)
        if rs & ts or any(len(x) >= 5 and x in rg for x in ts) or (len(tg) >= 5 and tg in rg):
            return True
    return any(g.replace(" ", "") in rg for g in GRP_OK)
out = []
for (lin, row), keys in zip(M["data"], pb):
    seg, band, grupo = row[2], str(row[1] or ""), str(row[0] or "")
    bn = norm_legal(AP(band)).replace(" ", "")
    bkeys = {k for k in keys if k.replace(" ", "") in bn} or keys
    u = str(row[20] or ""); ru = D(u)[:8] if len(D(u)) >= 8 else ""
    x = dict(lin=lin, band=band, grupo=grupo, seg=seg, u=u, ru=ru)
    cand = collections.Counter()
    for k in bkeys:
        for n, r in BYKEY.get(k, ()):
            m = n_band([k], r, seg)
            if m and (m >= 2 or nroots[k] <= 3): cand[r] = max(cand[r], m)
    x["cand"] = cand.most_common(6)
    if ru:
        nb = n_band(bkeys, ru, seg); rz = emp.get(ru, ("",))[0]; x["rz_u"] = rz
        if nb: x["status"], x["obs"] = "confere", f"{nb} lojas ativas da raiz com o nome da bandeira"
        elif rz and razao_ok(rz, band, grupo): x["status"], x["obs"] = "confere", "empresa do grupo (razão social)"
        elif not rz: x["status"], x["obs"] = "nao_confere", "raiz não existe na Receita"
        else: x["status"], x["obs"] = "nao_confere", f"raiz é de outra empresa ({rz[:40]})"
    else:
        x["status"] = "sem_cnpj"
    out.append(x)
pickle.dump(out, open("valida2.pkl", "wb"))
print(collections.Counter(x["status"] for x in out))
need = {r for x in out for r, n in x["cand"]} | {x["ru"] for x in out if x["ru"]}
pickle.dump(need, open("need_roots.pkl", "wb")); print("raizes para detalhar:", len(need))
