import datetime, pickle
ORDER = ["cnpj", "razao", "email", "fantasia", "desc"]
LBL = {"cnpj": "CNPJ indicado", "razao": "Razão social indicada", "email": "Domínio de e-mail",
       "fantasia": "Nome da bandeira no cliente", "desc": "Produto personalizado"}
def to_date(dt): return datetime.datetime(dt // 10000, dt // 100 % 100, dt % 100)
def choose(c):
    strong = [(m, c[m]) for m in ("cnpj", "razao") if m in c]
    weak = [(m, c[m]) for m in ("email", "fantasia", "desc") if m in c]
    pick = lambda l: max(l, key=lambda md: (md[1]["dt"], -ORDER.index(md[0])))
    if not strong: return pick(weak) if weak else None
    best = pick(strong)
    over = [md for md in weak if md[1].get("buyer_has_brand") and md[1]["dt"] > best[1]["dt"]]
    return pick(over) if over else best
