"""Aba "SUPERMERCADOS (reuniao)" (22/10/2026, 2a rodada): nº de lojas das linhas com L = 0 e linhas repetidas.

1. Raiz de CNPJ de cada linha com L = 0: a da aba; senao a da razao social igual + UF (Receita, 1a rodada); senao o
   "miolo" do nome igual na razao social ou dentro do nome fantasia, na UF da linha e em estabelecimento com CNAE de loja
   (2a rodada).  A raiz so vale se tiver loja ativa (CNAE de supermercado, minimercado, varejo alimenticio ou
   atacarejo, sem CD/deposito/escritorio).
2. Nº de lojas = esses estabelecimentos ativos da raiz (mesma regra da coluna L, que bate com a Receita).  Grupos com
   varias bandeiras/CNPJs ou holding sem lojas: numero publicado na internet (web_lojas.json, com fonte).
3. Linhas repetidas (a mesma empresa): mesma raiz de CNPJ, mesma posicao no ranking (regra da propria aba) ou a mesma
   rede confirmada (grupos com bandeiras em outras linhas).  Fica a linha de maior porte (ranking); ela recebe o nº de
   lojas e a compra mais recente do grupo (data AK, status AL, vendedor AN); as outras sao excluidas.
Saida: ajuste_spec.json (para ajusta_reuniao.ps1) e consolida.pkl (conferencia).
"""
import json, pickle, re, collections as C, datetime as dt, unicodedata, os, sys
from openpyxl.utils import column_index_from_string as ci

D = lambda x: re.sub(r"\D", "", str(x or ""))
STORE = {"4711301", "4711302", "4712100", "4729699", "4691500", "4639701", "4639702", "4693100"}
CDW = re.compile(r"\bCD\b|DEPOSIT|CENTRO DE DISTRIB|ESCRIT|ADMINISTR|LOGISTIC|ARMAZEM GERAL|GALPAO|MANUTENCAO")
AI = {"A": 20, "B": 15, "C": 10, "D": 8}
# a mesma rede em linhas sem CNPJ em comum (confirmado na internet): linha que fica -> linhas da mesma rede
MESMA_REDE = {21: [20, 861, 867, "Plurix: Superpão e Paraná Supermercados são bandeiras da Plurix"],
              15: [14, 863, "Dom Atacadista (MG e RJ) é bandeira do Mart Minas"],
              67: [72, 865, "Costazul Alimentos opera o Super Bom Supermercados, do Grupo Barcelos"]}


def sa(x):
    return unicodedata.normalize("NFKD", str(x or "")).encode("ascii", "ignore").decode().upper()


def conta(estabs):
    """estabs: [(raiz, fantasia, situacao, cnae, uf)] -> nº de lojas ativas"""
    return sum(1 for r, f, s, c, u in estabs if s.lstrip("0") == "2" and c in STORE and not CDW.search(sa(f)))


def main():
    a = pickle.load(open("abas.pkl", "rb"))["SUPERMERCADOS (reuniao)"]
    g = lambda r, c: r[ci(c) - 1]
    rows = {i + 1: r for i, r in enumerate(a) if i >= 4 and any(v is not None for v in r)}
    p1 = {r["lin"]: r for r in json.load(open("rf_l0.json", encoding="utf8"))}
    web = {int(k): v for k, v in json.load(open("web_lojas.json", encoding="utf8")).items()}
    e1 = pickle.load(open("receita/estab_l0.pkl", "rb"))
    p2 = pickle.load(open("receita/pass2.pkl", "rb"))
    sys.path.insert(0, "receita")
    from pass2 import miolo
    est = C.defaultdict(dict)  # raiz -> ordem do estabelecimento -> (raiz, fantasia, situacao, cnae, uf)
    for x in e1:
        est[x[0]][x[1]] = (x[0], x[3], x[4], x[5], x[6])
    for x in p2["est"]:
        est[x[0]][x[1]] = (x[0], x[2], x[3], x[4], x[5])
    est = {k: list(v.values()) for k, v in est.items()}
    loja_uf = lambda root, uf: any(s.lstrip("0") == "2" and c in STORE and u == uf and not CDW.search(sa(f)) for r, f, s, c, u in est.get(root, []))
    pend = {p["lin"]: p for p in json.load(open("receita/pend.json", encoding="utf8"))}
    fant_hits = C.defaultdict(C.Counter)  # (miolo, uf) -> raiz -> nº de estabelecimentos
    for x in p2["fant"]:
        for m in x[7]:
            fant_hits[(m, x[5])][x[0]] += 1
    raz_hits = C.defaultdict(set)
    for x in p2["raz"]:
        raz_hits[x[2]].add(x[0])
    # 3a rodada: razao com iniciais e sem o CPF do fim; raizes conferidas a mao (raiz_manual.json)
    p3 = pickle.load(open("receita/pass3.pkl", "rb")) if os.path.exists("receita/pass3.pkl") else dict(raz=[], est=[])
    for x in p3["est"]:
        est.setdefault(x[0], []).append((x[0], x[2], x[3], x[4], x[5]))
    raz3 = C.defaultdict(set)
    for x in p3["raz"]:
        raz3[x[2]].add(x[0])
    from pass3 import miolo3
    manual = {int(k): v for k, v in json.load(open("raiz_manual.json", encoding="utf8")).items()}

    # ---- 1/2. raiz e nº de lojas das linhas com L = 0
    lojas, fonte, raiz = {}, {}, {}
    for n, r in rows.items():
        c = D(g(r, "F"))
        if len(c) == 14:
            raiz[n] = c[:8]
    for n, o in p1.items():
        uf = o["uf"]
        if n in web:
            lojas[n], fonte[n] = web[n]
            continue
        rt, como = None, ""
        if o.get("raiz") and (o["cnpj"] or o["grupo"] == "nome: com loja"):
            rt, como = o["raiz"], ("CNPJ da aba" if o["cnpj"] else "CNPJ achado pela razão social + UF")
        if not o["cnpj"] and (rt is None or not loja_uf(rt, uf)) and n in pend:
            rt = None
            cand = {x for x in raz_hits.get(miolo(o["raz"]), set()) if loja_uf(x, uf)}
            cand3 = {x for x in raz3.get(miolo3(o["raz"]), set()) | raz3.get(miolo3(o["fant"]), set()) if loja_uf(x, uf)}
            if n in manual:
                rt, como = manual[n][0], "CNPJ conferido à mão: " + manual[n][1]
            elif len(cand) == 1:
                rt, como = cand.pop(), "CNPJ achado pelo miolo da razão social + UF"
            elif len(cand3) == 1:
                rt, como = cand3.pop(), "CNPJ achado pela razão social (com iniciais) + UF"
            else:
                cf = C.Counter()
                for m in pend[n]["miolos"]:
                    for x, k in fant_hits.get((m, uf), {}).items():
                        if loja_uf(x, uf):
                            cf[x] += k
                if len(cf) == 1:
                    rt, como = next(iter(cf)), "CNPJ achado pelo nome fantasia + UF"
                elif len(cf) > 1:
                    como = f"ambíguo: {len(cf)} empresas com esse nome na UF"
        if rt:
            raiz.setdefault(n, rt)
            k = conta(est.get(rt, []))
            sit = C.Counter(s.lstrip("0") for _, _, s, _, _ in est.get(rt, []))
            if k:
                lojas[n], fonte[n] = k, f"Receita (out/2026): {k} loja(s) ativa(s) de supermercado/atacarejo – {como} (raiz {rt})"
            elif sit and not sit.get("2"):
                lojas[n], fonte[n] = 0, f"0 – CNPJ {'inapto' if sit.get('4') else 'baixado/suspenso'} na Receita: nenhuma loja ativa (raiz {rt})"
            else:
                cn = C.Counter(c for _, _, s, c, _ in est.get(rt, []) if s.lstrip("0") == "2").most_common(1)
                lojas[n], fonte[n] = 0, f"0 – CNPJ ativo, mas sem estabelecimento de supermercado (CNAE {cn[0][0] if cn else '?'}) (raiz {rt})"
        elif (g(rows[n], "K") or 0) > 0:  # declarou faturamento ao Ranking: tem ao menos 1 loja
            lojas[n], fonte[n] = 1, ("Mínimo de 1 loja: a empresa declarou faturamento ao Ranking ABRAS, mas o CNPJ não foi "
                                     "localizado na Receita nem na internet (" + (como or "sem CNPJ e sem nome único") + ")")
        else:
            lojas[n], fonte[n] = 0, "0 – não localizado na Receita nem na internet (" + (como or "sem CNPJ e sem nome único") + ")"

    # ---- 3. linhas repetidas
    par = {n: n for n in rows}
    def f(x):
        while par[x] != x:
            par[x] = par[par[x]]; x = par[x]
        return x
    motivo = C.defaultdict(set)
    def une(x, y, why):
        rx, ry = f(x), f(y)
        if rx != ry:
            par[ry] = rx
        motivo[x].add(why); motivo[y].add(why)
    for chave, why in ((lambda n: raiz.get(n), "mesmo CNPJ"), (lambda n: ("pos", g(rows[n], "A")), "mesma posição no ranking")):
        b = C.defaultdict(list)
        for n in rows:
            k = chave(n)
            if k:
                b[k].append(n)
        for v in b.values():
            for x in v[1:]:
                une(v[0], x, why)
    for k, v in MESMA_REDE.items():
        for x in v[:-1]:
            une(k, x, v[-1])
    grupos = C.defaultdict(list)
    for n in rows:
        grupos[f(n)].append(n)
    grupos = [sorted(v) for v in grupos.values() if len(v) > 1]
    porte = lambda n: g(rows[n], "J") or "D"
    akd = lambda n: g(rows[n], "AK") if isinstance(g(rows[n], "AK"), dt.datetime) else dt.datetime(1900, 1, 1)
    excluir, cel, uni = [], {}, []
    for v in grupos:
        # fica: a rede confirmada; senao a linha com fonte da internet; senao a de maior porte/faturamento, a mais recente
        com_L = [n for n in v if (g(rows[n], "L") or 0) > 0]
        fica = (next((k for k in MESMA_REDE if k in v), None)
                or (None if com_L else next((n for n in v if n in web and web[n][0] > 0), None))
                or min(v, key=lambda n: (porte(n), not (g(rows[n], "L") or 0), -(g(rows[n], "K") or 0), -akd(n).toordinal(), n)))
        sai = [n for n in v if n != fica]
        Lfica = lojas.get(fica, g(rows[fica], "L") or 0)
        if not Lfica:  # sem loja na linha que fica: a maior L do grupo (mesma empresa)
            Lfica = max((lojas.get(n, g(rows[n], "L") or 0) for n in v), default=0)
        recente = max(v, key=akd)
        nomes = "; ".join(f'"{g(rows[n], "D")}" ({g(rows[n], "C")})' for n in sai)
        porque = "; ".join(sorted(set().union(*[motivo[n] for n in v])))
        nota = f"Linha única da empresa: também estava como {nomes} – {porque}"
        if fica in fonte:
            nota = fonte[fica] + " | " + nota
        elif Lfica != (g(rows[fica], "L") or 0):
            nota = f"nº de lojas da linha unificada ({Lfica}) | " + nota
        cel[fica] = dict(L=Lfica, AO=nota)
        if recente != fica and akd(recente) > akd(fica):
            cel[fica].update(AK=g(rows[recente], "AK"), AL=g(rows[recente], "AL"), AN=g(rows[recente], "AN"))
        excluir += sai
        uni.append(dict(fica=fica, sai=sai, porque=porque))
    for n in lojas:
        if n in excluir or n in cel:
            continue
        cel[n] = dict(L=lojas[n], AO=fonte[n])

    # ---- saida
    serial = lambda d: (d - dt.datetime(1899, 12, 30)).days
    cells = []
    for n, c in sorted(cel.items()):
        cells.append(dict(cel=f"L{n}", tipo="num", valor=int(c["L"])))
        cells.append(dict(cel=f"AO{n}", tipo="txt", valor=c["AO"]))
        if "AK" in c:
            cells.append(dict(cel=f"AK{n}", tipo="data", valor=serial(c["AK"])))
            cells.append(dict(cel=f"AL{n}", tipo="txt", valor=c["AL"]))
            cells.append(dict(cel=f"AN{n}", tipo="txt" if c["AN"] else "vazio", valor=c["AN"] or ""))
    spec = dict(aba="SUPERMERCADOS (reuniao)", celulas=cells, excluir=sorted(excluir),
                cabecalho=dict(cel="AO4", valor="Fonte do nº de lojas (L) / linha unificada", modelo="AN4", col="AO"),
                atualizar_abas=["TAB DIN. SUPERMERCADOS final"])
    json.dump(spec, open("ajuste_spec.json", "w", encoding="utf8"), ensure_ascii=True, indent=0)
    pickle.dump(dict(lojas=lojas, fonte=fonte, raiz=raiz, grupos=uni, cel=cel, excluir=excluir), open("consolida.pkl", "wb"))

    # ---- conferencia
    antes = sum(g(r, "AJ") for r in rows.values())
    depois = 0
    for n, r in rows.items():
        if n in excluir:
            continue
        L = cel[n]["L"] if n in cel else (g(r, "L") or 0)
        depois += AI.get(g(r, "J"), 0) * L
    print("linhas", len(rows), "excluídas", len(excluir), "ficam", len(rows) - len(excluir), "grupos", len(uni))
    print("caixas antes", antes, "depois", depois)
    print("fontes L=0:", C.Counter(re.sub(r"\(raiz.*|\d+ loja\(s\)|: \d+ .*", "", v)[:60] for v in fonte.values()).most_common(14))
    l0 = [n for n in lojas if n not in excluir]
    print("das", len(lojas), "linhas sem loja:", len(l0), "ficam;", sum(1 for n in l0 if lojas[n] > 0), "ganharam nº de lojas,",
          sum(lojas[n] for n in l0), "lojas")


if __name__ == "__main__":
    main()
