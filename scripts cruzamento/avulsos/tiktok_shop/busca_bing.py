"""Busca no Bing (por script, sem credito) as lojas AMBIGUO sem decisao e cruza os trechos com a cidade dos candidatos.
Uso: python busca_bing.py <tt_cnpj2.json> <dec_ambiguos.json> <CRM.xlsx> <cache_bing.json> <sugestoes.json>
Sugestao so quando os trechos que citam a loja trazem a cidade de UM candidato (revisar a mao antes de aceitar)."""
import sys, json, os, re, html, time, subprocess, urllib.parse, openpyxl
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from prep_termos import norm
d = json.load(open(sys.argv[1], encoding="utf8")); dec = json.load(open(sys.argv[2], encoding="utf8"))
wb = openpyxl.load_workbook(sys.argv[3], read_only=True); ws = wb["Leads"]
L = {r[0]: (r[1], r[30]) for r in ws.iter_rows(min_row=2, values_only=True) if r[0]}; wb.close()
cache = json.load(open(sys.argv[4], encoding="utf8")) if os.path.exists(sys.argv[4]) else {}
def bing(q):
    u = "https://www.bing.com/search?setlang=pt-br&cc=BR&q=" + urllib.parse.quote(q)
    t = subprocess.run(["curl", "-s", "-m", "30", "-A", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)", u], capture_output=True).stdout.decode("utf8", "ignore")
    out = []
    for m in re.finditer(r'<h2[^>]*><a[^>]*href="([^"]+)"[^>]*>(.*?)</a></h2>(.*?)</li>', t, re.S):
        tx = html.unescape(re.sub("<.*?>", " ", m.group(3)))
        out.append([m.group(1), html.unescape(re.sub("<.*?>", "", m.group(2))), " ".join(tx.split())[:400]])
    return out
sug = {}
amb = [k for k, v in d.items() if v["st"] == "AMBIGUO" and k not in dec]
for i, k in enumerate(amb):
    loja = L[k][0]
    for q in (f'"{loja}" instagram', f'"{loja}" loja cidade'):
        if q not in cache:
            cache[q] = bing(q); time.sleep(4)
    json.dump(cache, open(sys.argv[4], "w", encoding="utf8"), ensure_ascii=False)
    nl = norm(loja).replace(" ", "")
    hits = {}
    for q in (f'"{loja}" instagram', f'"{loja}" loja cidade'):
        for url, tit, tx in cache[q]:
            alltx = norm(tit + " " + tx + " " + url.replace(".", " ").replace("_", " "))
            if nl[:8] not in alltx.replace(" ", ""): continue
            if re.search(r"cnpj\.biz|serasa|econodata|jusbrasil|cnpja|casadosdados|cnpj\.info|consultas", url): continue
            for c in v_det if (v_det := d[k]["det"]) else []:
                cid = norm(c["munome"])
                if cid and re.search(r"\b" + re.escape(cid) + r"\b", alltx):
                    hits.setdefault(c["munome"], []).append(url)
    if len(hits) == 1:
        mun = next(iter(hits)); sug[k] = {"loja": loja, "mun": mun, "urls": sorted(set(hits[mun]))[:3],
                                         "n_cand_cidade": sum(1 for c in d[k]["det"] if c["munome"] == mun)}
    print(i, k, loja, "->", list(hits)[:4], flush=True)
json.dump(sug, open(sys.argv[5], "w", encoding="utf8"), ensure_ascii=False, indent=1)
print("sugestoes:", len(sug), "de", len(amb))
