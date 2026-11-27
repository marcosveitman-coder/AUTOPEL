"""QSA dos CNPJs achados, pela BrasilAPI (via curl; o urllib leva 403). Guarda tudo em cache json.
Uso: python qsa.py <tt_cnpj.json> <cache_qsa.json>"""
import sys, json, subprocess, time, os
d = json.load(open(sys.argv[1], encoding="utf8"))
cache = json.load(open(sys.argv[2], encoding="utf8")) if os.path.exists(sys.argv[2]) else {}
todo = sorted({v["cnpj"] for v in d.values() if v["st"] == "OK"} - set(cache))
print(len(todo), "a consultar", flush=True)
for i, c in enumerate(todo):
    for tent in range(4):
        r = subprocess.run(["curl", "-s", "-m", "30", f"https://brasilapi.com.br/api/cnpj/v1/{c}"], capture_output=True)
        try:
            j = json.loads(r.stdout.decode("utf8"))
        except Exception:
            time.sleep(5 * (tent + 1)); continue
        if "razao_social" in j or "não encontrado" in str(j.get("message", "")).lower():
            cache[c] = {"razao": j.get("razao_social"), "situacao": j.get("descricao_situacao_cadastral"),
                        "porte": j.get("porte"), "natureza": j.get("natureza_juridica"),
                        "qsa": [(q.get("nome_socio"), q.get("qualificacao_socio")) for q in j.get("qsa") or []]}
            break
        time.sleep(5 * (tent + 1))
    if i % 25 == 0:
        json.dump(cache, open(sys.argv[2], "w", encoding="utf8"), ensure_ascii=False); print(i, flush=True)
    time.sleep(0.7)
json.dump(cache, open(sys.argv[2], "w", encoding="utf8"), ensure_ascii=False)
print("fim", len(cache))
