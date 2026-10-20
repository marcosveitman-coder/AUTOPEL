"""Conferencia opcional dos candidatos: python conferir.py [resumo | cnpj | razao | fantasia | desc [so] | conflito]"""
import os, pickle, sys, collections
from common import CAND, fmt_cnpj, pick_vendor

with open(CAND, "rb") as f:
    D = pickle.load(f)
cand, UF, FANT, CNPJ = D["cand"], D["UF"], D["FANT"], D["CNPJ"]
what = sys.argv[1] if len(sys.argv) > 1 else "resumo"


def line(i, m):
    d = cand[i][m]
    extra = f' nucleo="{d["core"]}" comprador_tem_marca={d["buyer_has_brand"]}' if "core" in d else ""
    if m == "desc":
        extra += f' prod="{d["produto"]}"'
    return f'[{i+2}] {UF[i]} "{FANT[i]}" -> {d["dt"]} {pick_vendor(d["vends"])} | {d["cliente"]} {fmt_cnpj(d["cnpj"])} {d["seg"]}{extra}'


if what == "resumo":
    combos = collections.Counter(tuple(m for m in ("cnpj", "razao", "fantasia", "desc") if m in c) for c in cand)
    for k, v in combos.most_common():
        print(v, k)
    strong = lambda c: max([c[m]["dt"] for m in ("cnpj", "razao") if m in c], default=None)
    n = 0
    for i, c in enumerate(cand):
        s = strong(c)
        for m in ("fantasia", "desc"):
            if m in c and s and c[m]["dt"] > s:
                n += 1
    print("casos em que nome/descricao e mais recente que CNPJ/razao:", n)
elif what in ("cnpj", "razao", "fantasia", "desc"):
    only = len(sys.argv) > 2 and sys.argv[2] == "so"
    for i, c in enumerate(cand):
        if what in c and (not only or not any(m in c for m in ("cnpj", "razao") if m != what)):
            print(line(i, what))
elif what == "conflito":
    for i, c in enumerate(cand):
        s = [c[m] for m in ("cnpj", "razao") if m in c]
        for m in ("fantasia", "desc"):
            if m in c and s and c[m]["dt"] > max(x["dt"] for x in s):
                print(line(i, m))
                for mm in ("cnpj", "razao"):
                    if mm in c:
                        print("     ", line(i, mm))
