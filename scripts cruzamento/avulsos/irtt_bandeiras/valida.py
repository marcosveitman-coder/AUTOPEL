"""Passo 2: valida o CNPJ da coluna U (raiz) contra a Receita e sugere o CNPJ da bandeira quando nao confere."""
import pickle, re, collections, sys, os
sys.path.insert(0, os.environ["AUTOPEL_SCRIPTS"])
from common import norm_legal
AP = lambda s: re.sub(r"['’`´]", "", str(s or ""))
D = lambda s: re.sub(r"\D", "", str(s or ""))
R = pickle.load(open("receita_irtt.pkl", "rb")); M = pickle.load(open("mapa_irtt.pkl", "rb"))
G = pickle.load(open("irtt_grp.pkl", "rb"))
kr, kc, emp, info, pb = R["kr"], R["kc"], R["emp"], R["info"], R["por_band"]
nroots = collections.Counter(k for (k, r) in kr)
SEGC = {
    "Super, Hiper, Atacarejo e Conveniência": ("4711", "4712", "4721", "4722", "4723", "4724", "4729", "4639", "4691", "4637", "4731"),
    "Foodservice": ("5611", "5612", "5620", "1091", "4721", "4729", "1053", "4722"),
    "Moda, Calçados e Artigos Esportivos": ("4781", "4782", "4783", "4763", "1412", "1531", "1539", "4642", "4643", "4755", "4789", "1411", "1413"),
    "Drogarias e Perfumarias": ("4771", "4772", "2063", "4644", "4646"),
    "Eletromóveis": ("4753", "4754", "4752", "4751", "4759", "4713", "4757", "3104", "3101"),
    "Lojas de Departamento, Artigos do Lar e Mercadorias em Geral": ("4713", "4759", "4753", "4789", "4755", "4754", "4781", "4711", "4712", "4761", "4763"),
    "Material de Construção": ("4744", "4742", "4741", "4743", "4679", "4754", "4759"),
    "Óticas, Jóias, Bijoux, Bolsas e Acessórios": ("4774", "4783", "3211", "4782", "3212"),
    "Outros Segmentos": ("47", "56", "1093", "9609", "7500"),
}
def seg_ok(seg, cnae4):
    return any(cnae4.startswith(p) for p in SEGC.get(seg, ("47",)))
def n_band(keys, r, seg=None):
    """estabelecimentos ativos da raiz com a bandeira no nome fantasia (maior contagem entre as chaves)"""
    best = 0
    for k in keys:
        if seg:
            best = max(best, sum(n for c4, n in ((c, kc.get((k, r, c), 0)) for c in {c for (kk, rr, c) in KCIDX.get((k, r), ())}) if seg_ok(seg, c4)))
        else:
            best = max(best, kr.get((k, r), 0))
    return best
KCIDX = collections.defaultdict(set)
for (k, r, c) in kc: KCIDX[(k, r)].add((k, r, c))
BYKEY = collections.defaultdict(list)
for (k, r), n in kr.items(): BYKEY[k].append((n, r))

out = []
for (lin, row), keys in zip(M["data"], pb):
    seg, band, grupo = row[2], row[1], row[0]
    gkeys = {k for k in keys if k in pb_g} if False else set()
    bkeys = set()
    for k in keys:  # chaves da bandeira (nao do grupo)
        if k.replace(" ", "") in norm_legal(AP(band)).replace(" ", ""): bkeys.add(k)
    bkeys = bkeys or keys
    u = str(row[20] or ""); ru = D(u)[:8] if len(D(u)) >= 8 else ""
    res = dict(lin=lin, band=band, grupo=grupo, seg=seg, u=u, ru=ru)
    # candidatos: raizes com a bandeira no nome fantasia, CNAE do segmento
    cand = collections.Counter()
    for k in bkeys:
        for n, r in BYKEY.get(k, ()):
            m = n_band([k], r, seg)
            if m: cand[r] = max(cand[r], m)
    res["cand"] = cand.most_common(5)
    if ru:
        nb = n_band(bkeys, ru, seg); nall = n_band(keys, ru)
        rz = emp.get(ru, ("",))[0]
        rzn = norm_legal(AP(rz)).replace(" ", "")
        if nb:
            res["status"] = "confere"; res["obs"] = f"{nb} lojas ativas da raiz com a bandeira no nome fantasia"
        elif any(k.replace(" ", "") in rzn for k in keys):
            res["status"] = "confere"; res["obs"] = "razão social da raiz traz o nome da bandeira/grupo"
        elif nall:
            res["status"] = "confere"; res["obs"] = f"{nall} estabelecimentos da raiz com o nome da bandeira/grupo"
        else:
            res["status"] = "nao_confere"; res["obs"] = f"raiz {'não encontrada' if ru not in emp else 'sem relação de nome'} na Receita"
        res["rz_u"] = rz; res["info_u"] = info.get(ru)
    else:
        res["status"] = "sem_cnpj"
    out.append(res)
pickle.dump(out, open("valida.pkl", "wb"))
c = collections.Counter(x["status"] for x in out); print(c)
for x in out[:0]: pass
