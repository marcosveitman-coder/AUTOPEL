"""Ajuda do lote do LinkedIn: grava os achados (JSON no argv[1], pode ser {}), marca o lote anterior como feito e imprime o proximo.
Uso: python li_passo.py '<json achados>' <ini_feito> <fim_feito> <ini_prox> <fim_prox>"""
import sys, json
r = json.load(open("dados/li_result.json", encoding="utf8")); r.update(json.loads(sys.argv[1]))
json.dump(r, open("dados/li_result.json", "w", encoding="utf8"), ensure_ascii=False, indent=0)
f = json.load(open("dados/li_feitos.json")); L = json.load(open("dados/li_lista.json", encoding="utf8"))
a, b, c, d = map(int, sys.argv[2:6])
f += [x["id"] for x in L[a:b] if x["id"] not in f]; json.dump(f, open("dados/li_feitos.json", "w"))
B = [{k: x[k] for k in ("id", "q", "lojas", "ceo", "dir", "estado")} for x in L[c:d]]
print(len(f), "feitos;", len(r), "lojas com perfil; prox", L[c]["id"] if B else "-", "pri", L[d - 1]["pri"] if B else "-")
print(json.dumps(B, ensure_ascii=False, separators=(",", ":")))
