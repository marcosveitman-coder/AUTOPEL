"""Monta as acoes do browser_batch (Claude in Chrome) para os itens a..b (1-based) de dados/li_lista.json.
Uso: python linkedin_lote_tt.py <li_lista.json> <a> <b> <tabId> [--fn]  (--fn grava a funcao no localStorage; 1a vez)   -> imprime o JSON das acoes
Pausa de 8-10 s entre buscas (24 seguidas derrubaram a sessao em 25/11). Funcao em linkedin_funcoes.js."""
import sys, json, os, random, urllib.parse as u
L = json.load(open(sys.argv[1], encoding="utf8")); a, b, tab = int(sys.argv[2]), int(sys.argv[3]), int(sys.argv[4])
B = [{k: x[k] for k in ("id", "q", "lojas", "ceo", "dir", "estado")} for x in L[a - 1:b]]
fn = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "linkedin_funcoes.js"), encoding="utf8").read()
fn = fn[fn.index("(async"):]
setup = ("localStorage.__tm=%s;" % json.dumps(fn) if "--fn" in sys.argv else "") + "localStorage.__tb=%s;localStorage.__tr='{}';'ok'" % json.dumps(json.dumps(B, ensure_ascii=False, separators=(",", ":")), ensure_ascii=False)
acts = [{"name": "javascript_tool", "input": {"action": "javascript_exec", "tabId": tab, "text": setup}},
        {"name": "navigate", "input": {"tabId": tab, "url": "https://www.linkedin.com/search/results/people/?keywords=" + u.quote('"' + B[0]["q"] + '"')}}]
for i in range(len(B)):
    acts.append({"name": "computer", "input": {"action": "wait", "duration": random.choice([4, 5, 6]), "tabId": tab}})
    acts.append({"name": "javascript_tool", "input": {"action": "javascript_exec", "tabId": tab, "text": "await eval(localStorage.__tm)(%d)" % i}})
    if i + 1 < len(B): acts.append({"name": "computer", "input": {"action": "wait", "duration": random.choice([4, 5]), "tabId": tab}})
acts.append({"name": "javascript_tool", "input": {"action": "javascript_exec", "tabId": tab, "text": "JSON.stringify(Object.values(JSON.parse(localStorage.__tr||'{}')).filter(r=>r.u||r.erro))"}})
print(json.dumps(acts, ensure_ascii=False))
