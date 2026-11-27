"""Desempata as lojas AMBIGUO do casa_cnpj.py.
Uso: python desempata.py <tt_cnpj.json> <dec_ambiguos.json> <saida.json>
1) decisao manual (dec_ambiguos.json, local): {ID: {"cnpj": "14 digitos" ou "mun": "CIDADE", "fonte": "..."}}
   - "mun": o unico candidato daquela cidade (se houver mais de um, vale a regra 2 dentro da cidade);
2) todos os candidatos (ou os da cidade) na mesma cidade: a empresa cuja razao contem o nome; senao a mais nova.
Grava st = "OK" com "nivel" = "PROVAVEL (web)" ou "PROVAVEL (mesma cidade)"."""
import sys, json
sys.path.insert(0, __import__("os").path.dirname(__import__("os").path.abspath(__file__)))
from prep_termos import norm
d = json.load(open(sys.argv[1], encoding="utf8")); dec = json.load(open(sys.argv[2], encoding="utf8"))
def escolhe(cs, nomes):
    if len({c["mun"] for c in cs}) != 1: return None
    com = [c for c in cs if any(n.replace(" ", "") in norm(c["razao"]).replace(" ", "") for n in nomes)]
    return sorted(com or cs, key=lambda c: c["inicio"])[-1]
n = {"web": 0, "cidade": 0}
for k, v in d.items():
    if v["st"] != "AMBIGUO": continue
    cs = v["det"]; nomes = [norm(c["fantasia"]) for c in cs if c["fantasia"]]
    m = dec.get(k); esc = None; nivel = ""
    if m and m.get("cnpj"):
        esc = next((c for c in cs if c["cnpj"] == m["cnpj"]), None); nivel = "PROVAVEL (web)"
    elif m and m.get("mun"):
        sub = [c for c in cs if c["munome"] == m["mun"]]
        esc = sub[0] if len(sub) == 1 else escolhe(sub, nomes) if sub else None; nivel = "PROVAVEL (web)"
    if not esc and not m:
        esc = escolhe(cs, nomes); nivel = "PROVAVEL (mesma cidade)"
    if not esc: continue
    n["web" if "web" in nivel else "cidade"] += 1
    d[k] = {"st": "OK", "cnpj": esc["cnpj"], "razao": esc["razao"], "fantasia": esc["fantasia"], "cnae": esc["cnae"],
            "uf": esc["uf"], "inicio": esc["inicio"], "como": esc["como"], "nivel": nivel,
            "fonte": (m or {}).get("fonte", f"{len(cs)} candidatos, todos em {esc['munome'].title()}"), "n_cands": len(cs)}
json.dump(d, open(sys.argv[3], "w", encoding="utf8"), ensure_ascii=False)
print(n, "ainda ambiguos:", sum(1 for v in d.values() if v["st"] == "AMBIGUO"))
