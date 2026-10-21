"""Passos 3 e 4: faturamento 2025 da bandeira (mercado e Autopel) e numero de lojas."""
import os, sys, re, pickle, collections
A = os.environ["AUTOPEL_SCRIPTS"]; sys.path.insert(0, A); sys.path.insert(0, os.path.join(A, "avulsos"))
from common import norm_legal
from web2025 import WEB
from decisoes import ABRAS_MAP
exec(open("valida.py", encoding="utf-8").read().split("out = []")[0])  # kr, kc, n_band, seg_ok, pb...
M = pickle.load(open("mapa_irtt.pkl", "rb")); F = pickle.load(open("final_cnpj.pkl", "rb"))
Dt = pickle.load(open("detalhe_final.pkl", "rb")); E, I = Dt["emp"], Dt["info"]
G = pickle.load(open("irtt_grp.pkl", "rb"))["grp"]; grp_of = pickle.load(open("grp_of.pkl", "rb"))
RK = pickle.load(open("rankings.pkl", "rb")); P1 = pickle.load(open("passo1.pkl", "rb"))
SOMA = pickle.load(open(os.environ.get("AUTOPEL_SOMA2025", "soma2025.pkl"), "rb"))  # de ../farmacias/vf/soma2025.py
cod = {}
import glob, common
for p in glob.glob(os.path.join(os.environ["AUTOPEL_COD_DIR"], "*.pkl")): common.merge_into(cod, pickle.load(open(p, "rb")))
NL = lambda s: norm_legal(AP(s)).replace(" ", "")
abras = collections.defaultdict(list)
for a in RK["abras"]: abras[NL(a["razao"])].append(a)
farm = {r: f for f in RK["farm"] for r in f["roots"]}
ECOM = re.compile(r"\.com|\bonline\b|e-?commerce|marketplace|web\b", re.I)

# Carrefour: divisao das vendas brutas 1T25 (Cash & Carry 20.680, Varejo 5.476, Clube 1.676 - R$ mi)
_t = 20680 + 5476 + 1676
SPLIT = {"GRUPO CARREFOUR BRASIL": {"Atacadão": 20680 / _t, "SAM's Club": 1676 / _t, "Carrefour": 5476 / _t}}
SPLIT_MAIN = {"GRUPO CARREFOUR BRASIL": "Carrefour (varejo: hiper, bairro, express, Nacional e Bompreço)"}
rows = [(lin, row, pb[k]) for k, (lin, row) in enumerate(M["data"])]
bygrp = collections.defaultdict(list)
for lin, row, keys in rows: bygrp[grp_of[lin]].append(lin)
info = {}
for lin, row, keys in rows:
    band = str(row[1] or ""); bn = NL(band)
    bkeys = {k for k in keys if k.replace(" ", "") in bn} or keys
    info[lin] = dict(band=band, seg=row[2], bkeys=bkeys, root=F[lin]["root"], grp=grp_of[lin], ecom=bool(ECOM.search(band)))
# ---- lojas (passo 4)
out = {}
for g, lins in bygrp.items():
    gi = G[g]; roots = {info[l]["root"] for l in lins if info[l]["root"]}
    web = WEB.get(g); fr = next((farm[r] for r in roots if r in farm), None)
    tot, fonte_tot = (web[3], "2025 – " + web[2]) if web and web[3] else ((fr["lojas"], "2025 – ranking de farmácias") if fr and isinstance(fr["lojas"], (int, float)) else (gi["lojas24"], "2024 – IRTT"))
    for l in lins:
        x = info[l]; rec = n_receita = 0
        for r in roots: n_receita += n_band(x["bkeys"], r, x["seg"])
        if len(lins) == 1:
            lojas, fonte = tot, fonte_tot
            obs = f"Receita: {n_receita} estabelecimentos ativos com o nome da bandeira" if n_receita else ""
        elif x["ecom"]:
            lojas, fonte, obs = None, "loja on-line", ""
        else:
            lojas, fonte = (n_receita or None), "Receita – estabelecimentos ativos do grupo com o nome da bandeira"
            obs = f"grupo: {tot} lojas ({fonte_tot})" if tot else ""
        out[l] = dict(lojas=lojas, fonte_lojas=fonte, obs_lojas=obs)
for g, lins in bygrp.items():
    if len(lins) < 2: continue
    porn = collections.defaultdict(list)
    for l in lins:
        if out[l]["lojas"] and out[l]["fonte_lojas"].startswith("Receita"): porn[out[l]["lojas"]].append(l)
    for n, ls in porn.items():
        if len(ls) < 2: continue
        main = min(ls, key=lambda l: len(info[l]["band"]))
        if not all(NL(info[main]["band"]) in NL(info[l]["band"]) for l in ls):
            main = None  # nome em comum (ex.: MATEUS) nao separa as bandeiras
        for l in ls:
            if l != main and main is None:
                out[l].update(lojas=None, obs_lojas="Receita não separa esta bandeira das outras do grupo; " + out[l]["obs_lojas"])
            elif l != main:
                out[l].update(lojas=None, obs_lojas=f"lojas contadas na linha de {info[main]['band']} (mesmo nome na Receita); " + out[l]["obs_lojas"])
# ---- faturamento de mercado 2025 (passo 3)
for g, lins in bygrp.items():
    gi = G[g]; roots = [info[l]["root"] for l in lins if info[l]["root"]]
    val = fonte = None
    seg = info[lins[0]]["seg"]
    if WEB.get(g):
        w = WEB[g]; val, fonte = w[0], f"2025 – {w[1]} ({w[2]})"
    if val is None:
        fr = next((farm[r] for r in roots if r in farm), None)
        if fr and fr.get("fat_est"): val, fonte = fr["fat_est"] * 1e6, "2025 – ranking de farmácias (faturamento estimado)"
    if val is None and g in ABRAS_MAP:
        a_ = next((a for a in RK["abras"] if a["razao"].startswith(ABRAS_MAP[g])), None)
        if a_: val, fonte = a_["fat"], "2025 – Ranking ABRAS 2026"
    if val is None:
        ab = {}
        for r in set(roots):
            for a in abras.get(NL(E.get(r, ("",))[0]), []): ab[a["pos"]] = a
        if ab: val, fonte = sum(a["fat"] for a in ab.values()), "2025 – Ranking ABRAS 2026"
    if val is None and gi["venda24"] and gi["venda24"] >= 50:  # valores < R$ 50 mi no IRTT = dado nao divulgado
        val, fonte = gi["venda24"] * 1e6, "2024 – IRTT (2025 não publicado)"
    # rateio pelas lojas entre as bandeiras do grupo
    fis = [l for l in lins if not info[l]["ecom"]]
    pesos = {l: (out[l]["lojas"] or 0) for l in fis}
    tot = sum(pesos.values())
    gtot = G[g]["lojas24"] or 0
    if g in SPLIT and val is not None:
        for l in lins:
            k = SPLIT[g].get(info[l]["band"])
            out[l].update(fat=round(val * k) if k else None,
                          fonte_fat=f"{fonte}; parte da bandeira pela divisão do release 1T25" if k else f"incluído na linha de {SPLIT_MAIN[g]}")
        continue
    if len(lins) > 1 and val is not None and (not tot or (gtot and tot < 0.6 * gtot) or any(not pesos[l] for l in fis)):
        first = max(lins, key=lambda l: (out[l]["lojas"] or 0, -lins.index(l)))
        for l in lins:
            out[l].update(fat=val if l == first else None,
                          fonte_fat=(f"{fonte}; valor do GRUPO (todas as bandeiras) – sem número por bandeira" if l == first
                                     else f"incluído no valor do grupo, na linha de {info[first]['band']}"))
        continue
    for l in lins:
        if val is None:
            out[l].update(fat=None, fonte_fat="não encontrado")
        elif len(lins) == 1:
            out[l].update(fat=val, fonte_fat=fonte)
        elif info[l]["ecom"] or not tot:
            out[l].update(fat=None, fonte_fat=f"grupo: R$ {val/1e9:.2f} bi ({fonte}); sem base para separar a bandeira")
        else:
            out[l].update(fat=round(val * pesos[l] / tot) if pesos[l] else None,
                          fonte_fat=f"{fonte}; rateio do grupo (R$ {val/1e9:.2f} bi) pelas lojas da bandeira")
# ---- valor faturado Autopel 2025
dono = {}
for lin, row, keys in rows:
    ch = P1["res"].get(lin)
    rs = set()
    if info[lin]["root"]: rs.add(info[lin]["root"])
    if ch and len(ch[1]["cnpj"]) == 14: rs.add(ch[1]["cnpj"][:8])
    for c in re.findall(r"\d+", str(row[18] or "")):
        rec = cod.get(c.zfill(6))
        if rec and len(rec[3]) == 14: rs.add(rec[3][:8])
    v, ja = 0.0, []
    for r in rs:
        if r in dono: ja.append(dono[r]); continue
        dono[r] = info[lin]["band"]; v += SOMA["root"].get(r, 0)
    out[lin].update(aut=round(v, 2) if v else None,
                    obs_aut=("compras do CNPJ já somadas na linha de " + ", ".join(sorted(set(ja)))) if ja and not v else "")
pickle.dump(out, open("passo34.pkl", "wb"))
c = collections.Counter(o["fonte_fat"].split(";")[0][:30] for o in out.values()); print(c.most_common(12))
print("lojas:", sum(1 for o in out.values() if o["lojas"]), "| fat mercado:", sum(1 for o in out.values() if o["fat"]),
      "| autopel 2025:", sum(1 for o in out.values() if o["aut"]), "total R$", round(sum(o["aut"] or 0 for o in out.values())))
for lin in (5, 6, 8, 9, 204, 205, 210, 211, 212, 213, 247, 249, 163, 412, 239):
    o = out[lin]; print(lin, info[lin]["band"][:22], "| lojas", o["lojas"], o["fonte_lojas"][:30], "|", o["obs_lojas"][:40], "| fat", o["fat"], o["fonte_fat"][:70], "| aut", o["aut"], o["obs_aut"][:40])
