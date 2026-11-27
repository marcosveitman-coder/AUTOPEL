"""Lista de LOJAS do CRM TikTok Shop para buscar pessoas no LinkedIn pelo nome da loja (prioridade A > B > C).
Uso: python linkedin_lista.py <tt_pes3.json> <CRM.xlsx> <saida.json>
Cada item: id, q (nome da loja na busca), lojas (nomes para conferir o cartao), ceo/dir (1o nome..sobrenomes do QSA), estado."""
import sys, json, re, openpyxl
from prep_termos import norm
UF = {"AC":"ACRE","AL":"ALAGOAS","AP":"AMAPA","AM":"AMAZONAS","BA":"BAHIA","CE":"CEARA","DF":"DISTRITO FEDERAL","ES":"ESPIRITO SANTO",
      "GO":"GOIAS","MA":"MARANHAO","MT":"MATO GROSSO","MS":"MATO GROSSO DO SUL","MG":"MINAS GERAIS","PA":"PARA","PB":"PARAIBA","PR":"PARANA",
      "PE":"PERNAMBUCO","PI":"PIAUI","RJ":"RIO DE JANEIRO","RN":"RIO GRANDE DO NORTE","RS":"RIO GRANDE DO SUL","RO":"RONDONIA","RR":"RORAIMA",
      "SC":"SANTA CATARINA","SP":"SAO PAULO","SE":"SERGIPE","TO":"TOCANTINS"}
TIPO = {"LTDA","ME","EPP","EIRELI","SA","S","A","LIMITADA","MEI","COMERCIO","DE","DA","DO","DAS","DOS","E","PRODUTOS","EM","GERAL","INDUSTRIA","VAREJISTA","ATACADISTA","ARTIGOS","BR"}
P = json.load(open(sys.argv[1], encoding="utf8"))
wb = openpyxl.load_workbook(sys.argv[2], read_only=True); ws = wb["Leads"]
L = {r[0]: (r[1], r[2]) for r in ws.iter_rows(min_row=2, values_only=True) if r[0]}; wb.close()
CORTE = {"MODA","MODAS","PRAIA","VARIEDADES","COMERCIO","CONFECCOES","STORE","IMPORTS","LTDA","ME","EPP"}
def tks(n):
    out = []
    for w in (norm(re.split(r"\s-\s|\d", n)[0]).split() if n else []):
        if w in CORTE: break
        if w not in {"DA","DE","DO","DAS","DOS","E"} and len(w) > 1: out.append(w)
    return out
out = []
for k, v in P.items():
    if v["st"] != "OK": continue
    loja, pri = L[k]
    pessoas = set(tks(v.get("ceo"))) | set(tks(v.get("dir")))
    raz = " ".join(w for w in norm(re.sub(r"\d{6,}", "", v["razao"] or "")).split() if w not in TIPO)
    nomes = {norm(loja), norm(v["fantasia"] or ""), raz}
    lojas = sorted({x for x in nomes if len(x.replace(" ", "")) >= 5 and not set(x.split()) <= pessoas})
    fan = norm(v["fantasia"] or "")
    q = fan if fan in lojas else norm(loja) if norm(loja) in lojas else (lojas[0] if lojas else "")
    if not q: continue
    out.append({"id": k, "pri": pri or "C", "q": q.title(), "lojas": lojas, "ceo": tks(v.get("ceo")) or None,
                "dir": tks(v.get("dir")) or None, "estado": UF.get(v["uf"], "")})
out.sort(key=lambda x: (x["pri"], x["id"]))
json.dump(out, open(sys.argv[3], "w", encoding="utf8"), ensure_ascii=False)
from collections import Counter
print(len(out), Counter(x["pri"] for x in out))
