"""Monta CNPJ + pessoas por loja. Uso: python pessoas.py <tt_cnpj.json> <tt_qsa.json> <saida.json>
CEO = Presidente > Socio-Administrador > Administrador > Socio (unico) > titular do Empresario Individual (razao sem o CPF/CNPJ).
Diretor = Diretor; senao o 2o socio-administrador/administrador. Gerente: o QSA nao tem.
CNPJ BAIXADO/INAPTO na BrasilAPI nao entra."""
import sys, json, re
d = json.load(open(sys.argv[1], encoding="utf8")); Q = json.load(open(sys.argv[2], encoding="utf8"))
ORD = {"Presidente": 0, "Sócio-Administrador": 1, "Administrador": 2, "Sócio": 3}
out = {}
for k, v in d.items():
    if v["st"] != "OK": out[k] = {"st": v["st"]}; continue
    q = Q.get(v["cnpj"])
    if not q or q["situacao"] not in ("ATIVA",):
        out[k] = {"st": "CNPJ " + (q or {}).get("situacao", "SEM CONSULTA")}; continue
    soc = [(n.title(), ql) for n, ql in q["qsa"] if n]
    ceo = dire = None; ceo_q = dire_q = ""
    cand = sorted([s for s in soc if s[1] in ORD], key=lambda s: ORD[s[1]])
    if cand and (cand[0][1] != "Sócio" or len(soc) == 1):
        ceo, ceo_q = cand[0]
    if not soc and "Individual" in (q["natureza"] or ""):
        n = re.sub(r"^[\d.\s]+", "", q["razao"] or ""); n = re.sub(r"\s*\d{11}\s*$", "", n).strip()
        if n: ceo, ceo_q = n.title(), "Titular (Empresário Individual)"
    dd = [s for s in soc if s[1] == "Diretor" and s[0] != ceo]
    if not dd: dd = [s for s in soc if s[1] in ("Sócio-Administrador", "Administrador") and s[0] != ceo]
    if dd: dire, dire_q = dd[0]
    out[k] = {"st": "OK", "cnpj": v["cnpj"], "razao": q["razao"], "fantasia": v["fantasia"], "uf": v["uf"],
              "cnae": v["cnae"], "como": v["como"], "nivel": v.get("nivel", "CONFIRMADO (nome único)"), "fonte": v.get("fonte", ""), "ceo": ceo, "ceo_q": ceo_q, "dir": dire, "dir_q": dire_q}
json.dump(out, open(sys.argv[3], "w", encoding="utf8"), ensure_ascii=False, indent=0)
from collections import Counter
print(Counter(v["st"] for v in out.values()), "ceo", sum(1 for v in out.values() if v.get("ceo")), "dir", sum(1 for v in out.values() if v.get("dir")))
