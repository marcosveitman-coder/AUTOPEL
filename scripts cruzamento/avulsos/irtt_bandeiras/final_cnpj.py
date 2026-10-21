"""Junta a validacao automatica (valida2.pkl) com a revisao manual (decisoes.py) e o fallback de grupo -> raiz final por bandeira."""
import pickle, re, collections
from decisoes import DEC, SEM_FALLBACK
V = pickle.load(open("valida2.pkl", "rb"))
res = {}
for x in V:
    lin = x["lin"]; d = DEC.get(lin)
    if d:
        r, nota = d
        if r == "U":
            res[lin] = dict(root=x["ru"], tipo="confere" if x["status"] == "confere" or x["ru"] else "x", nota=nota)
            res[lin]["tipo"] = "confere_manual"
        elif r is None:
            res[lin] = dict(root=None, tipo="nao_encontrado", nota=nota)
        else:
            res[lin] = dict(root=r, tipo=("corrigido" if x["status"] == "nao_confere" and x["ru"] and x["ru"] != r
                                          else "completado" if x["ru"] == r else "encontrado"), nota=nota)
    elif x["status"] == "confere":
        res[lin] = dict(root=x["ru"], tipo="confere", nota=x["obs"])
    elif x["status"] == "nao_confere":
        res[lin] = dict(root=None, tipo="nao_confere", nota=x["obs"])
    else:
        res[lin] = dict(root=None, tipo="nao_encontrado", nota="")
# fallback: bandeira sem raiz -> raiz de outra bandeira do mesmo grupo
porg = collections.defaultdict(collections.Counter)
for x in V:
    r = res[x["lin"]]
    if r["root"]: porg[x["grupo"]][r["root"]] += 1
for x in V:
    r = res[x["lin"]]
    if not r["root"] and porg[x["grupo"]] and x["grupo"] not in SEM_FALLBACK:
        r.update(root=porg[x["grupo"]].most_common(1)[0][0], tipo="grupo",
                 nota=("CNPJ indicado era de outra empresa; " if r["tipo"] == "nao_confere" else "") + "bandeira sem CNPJ próprio identificado: CNPJ do grupo")
pickle.dump(res, open("final_cnpj.pkl", "wb"))
pickle.dump({r["root"] for r in res.values() if r["root"]}, open("roots_final.pkl", "wb"))
print(collections.Counter(r["tipo"] for r in res.values()))
