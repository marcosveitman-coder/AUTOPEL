import os
import sys, re, pickle, collections
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
from common import norm
from redes import REDES
AP = lambda s: re.sub(r"['’`´]", "", str(s or ""))
E = pickle.load(open("rf_estab.pkl", "rb")); EMP = pickle.load(open("rf_emp.pkl", "rb"))
act = [e for e in E if e[4] == "2" and e[6][:7] in ("4771701", "4771702", "4771703")]
fant = [(" " + norm(AP(e[3])) + " ", " " + norm(AP(EMP.get(e[0], ("",))[0])) + " ", e) for e in act]
res = {}
for nome, kws, uf, l1, l2, ufr, mae in REDES:
    if not kws: continue
    keys = set(kws) | {k.replace(" ", "") for k in kws}
    hit = [e for f, rz, e in fant if (ufr is None or e[7] in ufr) and any(f" {k} " in f or f" {k} " in rz for k in keys)]
    cnt = collections.Counter(e[0] for e in hit)
    tot = len(hit)
    chains = [r for r, c in cnt.items() if c >= 10 and c >= 0.15 * tot]
    tipo = "REDE" if chains and sum(cnt[r] for r in chains) >= 0.5 * tot else "ASSOC"
    res[nome] = dict(tipo=tipo, chains=chains, cnt=cnt, tot=tot, members=set(cnt), uf=uf, l1=l1, l2=l2, mae=mae, kws=kws, ufr=ufr)
    top = ", ".join(f"{r}:{c}({EMP.get(r,('?',))[0][:22]})" for r, c in cnt.most_common(3))
    print(f"{nome[:38]:38} {tipo:5} lojas_rf={tot:5} empresas={len(cnt):5} inf={l1}/{l2} | {top}")
pickle.dump(res, open("redes_res.pkl", "wb"))
