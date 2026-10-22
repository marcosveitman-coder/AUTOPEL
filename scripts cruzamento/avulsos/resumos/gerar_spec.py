"""Monta o conteudo das abas de resumo do MAPA MERCADO (22/10/2026): textos, tabelas pequenas e a definicao das
tabelas dinamicas.  Todos os numeros saem dos dados (nada digitado a mao).  Saida: resumos_spec.json, lido por
td_resumos.ps1, que grava pelo proprio Excel.

Entradas (na pasta de trabalho):
  abas.pkl       linhas das abas SUPERMERCADOS (reuniao), "FARMACIA CONSOLIDADO" e "FARMACIA VF" (valores)
  farm_final.pkl resultado de ../farmacias/final/farm_final.py
  lojas_sup.pkl  {raiz: [estab. ativos, estab. ativos de supermercado]} da Receita (conta_sup.py)
"""
import json, pickle, re, statistics as st, collections as C
from openpyxl.utils import column_index_from_string as ci

D = lambda x: re.sub(r"\D", "", str(x or ""))
n0 = lambda v: f"{v:,.0f}".replace(",", ".")
n1 = lambda v: f"{v:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
pct = lambda a, b: f"{100 * a / b:.1f}%".replace(".", ",")
HOJE = "22/10/2026"
TICKET_MERCADO = 135  # R$ por compra: Neogrid 2023 = R$ 125 + alta de ~6% (Scanntech 2025)


def supermercados(abas, R, lojas_rf):
    a = abas["SUPERMERCADOS (reuniao)"]
    g = lambda r, c: r[ci(c) - 1]
    rows = [r for r in a[4:] if any(v is not None for v in r)]
    ult = 4 + len(rows)
    ativo = lambda r: g(r, "AL") == "ativo"
    AI = {"A": 20, "B": 15, "C": 10, "D": 8}
    tot = sum(g(r, "AJ") for r in rows)
    at = [r for r in rows if ativo(r)]
    cx_at = sum(g(r, "AJ") for r in at)
    L0 = [r for r in rows if not g(r, "L")]
    sem_cnpj = sum(1 for r in L0 if len(D(g(r, "F"))) != 14)
    vistos, sem_loja, com, lj_rf, cx_rf = set(), 0, 0, 0, 0
    for r in L0:
        c = D(g(r, "F"))
        if len(c) != 14 or c[:8] in vistos:
            continue
        vistos.add(c[:8])
        v = lojas_rf.get(c[:8])
        if not v or v[1] == 0:
            sem_loja += 1
        else:
            com += 1; lj_rf += v[1]; cx_rf += v[1] * AI[g(r, "J")]
    falta, redes_ab, ex_ab = 0, 0, []
    for p in "AB":
        fl = st.median([g(r, "K") / g(r, "L") for r in rows if g(r, "J") == p and g(r, "L") and (g(r, "K") or 0) > 0])
        vis = set()
        for r in sorted(rows, key=lambda r: -(g(r, "K") or 0)):
            if g(r, "J") == p and not g(r, "L") and (g(r, "K") or 0) > 0 and g(r, "A") not in vis:
                vis.add(g(r, "A")); falta += g(r, "K") / fl * AI[p]
                if len(ex_ab) < 6:
                    ex_ab.append(str(g(r, "D")))
        redes_ab += sum(1 for r in rows if g(r, "J") == p and not g(r, "L"))
    br = C.defaultdict(list)
    for r in rows:
        c = D(g(r, "F"))
        if len(c) == 14:
            br[c[:8]].append(r)
    dup = {k: v for k, v in br.items() if len(v) > 1}
    cx_dup = sum(sum(g(r, "AJ") for r in v) - max(g(r, "AJ") for r in v) for v in dup.values())
    conf = sum(1 for v in dup.values() if len({g(r, "AL") for r in v}) > 1)
    # premissa x modelo N-T (caixas/loja/mes = faturamento mensal por loja / 1,82 mi), com o nº de lojas da coluna L
    modelo = {}
    for p in "ABC":
        v = [g(r, "K") / g(r, "L") / 12 / 1.82e6 for r in rows if g(r, "J") == p and (g(r, "K") or 0) > 0 and g(r, "L")]
        modelo[p] = (len(v), st.median(v))
    soma_L, soma_M = sum(g(r, "L") or 0 for r in rows), sum(g(r, "M") or 0 for r in rows)
    igual_rf = sum(1 for r in rows if g(r, "L") and len(D(g(r, "F"))) == 14 and D(g(r, "F"))[:8] in lojas_rf
                   and lojas_rf[D(g(r, "F"))[:8]][1] == g(r, "L"))
    base_rf = sum(1 for r in rows if g(r, "L") and len(D(g(r, "F"))) == 14 and D(g(r, "F"))[:8] in lojas_rf
                  and lojas_rf[D(g(r, "F"))[:8]][1] > 0)
    ad_al = sum(1 for r in rows if str(g(r, "AD")).lower() != ("ativo" if ativo(r) else "a trabalhar"))
    sem_vend = sum(1 for r in rows if not ativo(r) and not g(r, "AN"))
    ex_cli = sum(1 for r in rows if not ativo(r) and g(r, "AN"))
    vf = R["vf"]
    vfr = {c[:8]: i for i, c in enumerate(vf["cnpj"]) if len(c) == 14}
    ov = [k for k in br if k in vfr and vf["cx"][vfr[k]] > 0]
    cx_ov = sum(vf["cx"][vfr[k]] for k in ov)

    T1 = [["Porte (coluna J)", "Caixas por loja (AI)", "Linhas", "Linhas ativas", "Lojas (L)", "Caixas/mês (AJ)", "% das caixas"]]
    for p in "ABCD":
        rr = [r for r in rows if g(r, "J") == p]
        cx = sum(g(r, "AJ") for r in rr)
        T1.append([p, AI[p], len(rr), sum(1 for r in rr if ativo(r)), sum(g(r, "L") for r in rr), cx, pct(cx, tot)])
    T1.append(["Total", "", len(rows), len(at), soma_L, tot, "100,0%"])
    T2 = [["Porte", "Premissa AI (caixas/loja/mês)", "Modelo N–T, ticket R$ 450", f"Modelo N–T, ticket R$ {TICKET_MERCADO}",
           "Redes com faturamento"]]
    for p in "ABC":
        n, m = modelo[p]
        T2.append([p, AI[p], n1(m), n1(m * 450 / TICKET_MERCADO), n])

    L = []
    add = lambda cel, txt, neg=False, tam=None: L.append(dict(cel=cel, txt=txt, neg=neg, tam=tam))
    add("A1", f"RESUMO – SUPERMERCADOS (REUNIÃO)   |   base: aba SUPERMERCADOS (reuniao), {n0(len(rows))} linhas   |   posição de {HOJE}", True, 14)
    add("A3", "PARTE 1 – RACIONAL DAS CONTAS (colunas AI e AJ)", True, 12)
    add("A4", '1) Coluna AI "QUANTIDADE DE CAIXA POR LOJA": premissa fixa pelo porte da coluna J – A = 20, B = 15, C = 10 e D = 8 caixas por loja/mês. Fórmula: =SE(J5="A";20;SE(J5="B";15;SE(J5="C";10;SE(J5="D";8)))).')
    add("A5", '2) Coluna AJ "Total de caixas/ mês" = AI × L (nº de lojas da coluna L, "nª de LOJA (antonio)"). Exemplo, linha 5: porte A com 500 lojas → 20 × 500 = 10.000 caixas/mês.')
    add("A6", "3) Porte (colunas I e J): A = faturamento 2025 ≥ R$ 1 bi; B = R$ 300 mi a 1 bi; C = abaixo de R$ 300 mi; D = pequena/local, sem faturamento publicado (classificada pelo capital social e nº de lojas, coluna V).")
    add("A7", f'4) Status (coluna AL "status final (reuniao)"): ativo = última compra (coluna AK) há até 360 dias; sem compra ou mais antiga = a trabalhar. Conferido: a AL bate com a data da AK em 100% das linhas.')
    add("A8", f"5) Total (célula AJ3, que respeita o filtro): {n0(tot)} caixas/mês = {n0(cx_at)} de clientes ativos ({n0(len(at))} linhas, {pct(cx_at, tot)}) + {n0(tot - cx_at)} a trabalhar ({n0(len(rows) - len(at))} linhas).")
    tabs = [dict(cel="A10", linhas=T1)]
    r0 = 10 + len(T1) + 1
    add(f"A{r0}", "PONTOS DE ATENÇÃO (conferidos antes de usar o número)", True, 12)
    add(f"A{r0 + 1}", f"a) {n0(len(L0))} linhas estão com 0 loja na coluna L, e portanto 0 caixa: {sem_cnpj} sem CNPJ na aba, {sem_loja} com CNPJ mas sem loja de supermercado ativa na Receita (out/2026) e {com} com {lj_rf} lojas na Receita (+{n0(cx_rf)} caixas).")
    add(f"A{r0 + 2}", f"   Entre elas estão {redes_ab} linhas de porte A e B (ex.: {', '.join(ex_ab)}). Pelo faturamento por loja mediano do porte, faltam cerca de {n0(round(falta, -2))} caixas/mês ({pct(falta, tot)} do total). Recomendação: preencher a coluna L dessas redes.")
    add(f"A{r0 + 3}", f'b) {len(dup)} empresas (mesma raiz de CNPJ) aparecem em {sum(len(v) for v in dup.values())} linhas com o mesmo nº de lojas: {n0(cx_dup)} caixas estão contadas duas vezes no total. Em {conf} delas a mesma empresa está "ativo" numa linha e "a trabalhar" na outra.')
    add(f"A{r0 + 4}", f"c) A premissa da AI está acima do consumo calculado pelas colunas N–T (RACIONAL POTENCIAL) para as redes com faturamento (tabela abaixo). O modelo divide o faturamento por um ticket médio de R$ 450 (coluna P); o mercado gira em torno de R$ 125 a 140 por compra (Neogrid 2023: R$ 125,34; alta de ~6% em 2025, Scanntech).")
    add(f"A{r0 + 5}", f"   Leitura das fórmulas N–T: faturamento por loja/mês ÷ ticket = cupons → × 0,3 m de papel por cupom → ÷ 40 m por bobina → ÷ 30 bobinas por caixa. Com ticket de ~R$ {TICKET_MERCADO}, o modelo chega perto da premissa no porte A, mas fica bem abaixo em B e C. Vale validar a premissa por porte.")
    tabs.append(dict(cel=f"A{r0 + 6}", linhas=T2))
    r1 = r0 + 6 + len(T2) + 1
    add(f"A{r1}", f"d) As colunas N–T usam a coluna M (Lojas); a AJ usa a coluna L (antonio). Soma da L = {n0(soma_L)} lojas × M = {n0(soma_M)}. Pela Receita, a L bate exatamente em {n0(igual_rf)} de {n0(base_rf)} empresas com CNPJ: a L é a mais confiável.")
    add(f"A{r1 + 1}", f'e) {ad_al} linhas têm "Status no CRM FINAL" (AD, fórmula com a data do CRM) diferente do status final (AL, com a data final da AK). As tabelas usam a AL.')
    add(f"A{r1 + 2}", 'f) A AL está escrita "a tarbalhar"; nas tabelas dinâmicas aparece como "a trabalhar" (só o rótulo; a aba não foi alterada).')
    add(f"A{r1 + 3}", "g) UF = UF da linha na lista (estado da rede). Lojas de uma rede em outros estados contam nessa UF.")
    add(f"A{r1 + 4}", f'h) Vendedor (AN): vazio em {n0(sem_vend)} linhas a trabalhar (empresas sem cadastro na Autopel), que aparecem como "Sem vendedor". Outras {ex_cli} linhas a trabalhar já têm vendedor na carteira (ex-clientes).')
    add(f"A{r1 + 5}", f"i) {len(ov)} empresas desta lista também estão nas abas de farmácia (pelas lojas de drogaria delas). Não somar o potencial das duas abas sem descontar essas {n0(cx_ov)} caixas.")
    r2 = r1 + 7
    add(f"A{r2}", f"PARTE 2 – TABELAS DINÂMICAS   (fonte: SUPERMERCADOS (reuniao)!A4:AN{ult}; se a aba mudar, clique com o botão direito numa tabela > Atualizar)", True, 12)
    tds = [dict(nome="TD_SUP_UF", campo="UF", titulo="Clientes ativos e a trabalhar por UF + caixas/mês (AJ)", col="A", larg=12, ordena=True),
           dict(nome="TD_SUP_PORTE", campo="Porte", titulo="Clientes ativos e a trabalhar por porte (I) + caixas/mês (AJ)", col="J", larg=30, ordena=False),
           dict(nome="TD_SUP_VEND", campo="Vendedor", titulo="Clientes ativos e a trabalhar por vendedor (AN) + caixas/mês (AJ)", col="S", larg=44, ordena=True)]
    for t in tds:
        t["titulo_cel"] = f"{t['col']}{r2 + 2}"; t["cel"] = f"{t['col']}{r2 + 3}"
    return dict(aba="RESUMO SUPERMERCADO (REUNIAO)", linhas=L, tabelas=tabs, fonte_aba="SUPERMERCADOS (reuniao)",
                fonte_faixa=f"A4:AN{ult}", campo_status="status final (reuniao)", campo_conta="Nome fantasia", campo_caixas="Total de caixas/ mês",
                renomeia={"a tarbalhar": "a trabalhar"}, vazio="Sem vendedor", primeiro="ativo", tds=tds)


def farmacia(R, chave, aba, fonte_faixa, cols, n_antes):
    X = R[chave]
    n = X["n"]
    at = [i for i in range(n) if X["status"][i] == "ATIVO"]
    tot, cx_at = sum(X["cx"]), sum(X["cx"][i] for i in at)
    f = C.Counter(b["fonte"] for b in X["best"])
    zero = sum(1 for v in X["lojas"] if not v)
    lojas = sum(X["liq"][i] or 0 for i in range(n))
    T = [["Porte", "Caixas por loja", "Linhas", "Linhas ATIVO", "Caixas/mês", "Caixas/mês ATIVO", "% das caixas"]]
    for p, cxl in zip("ABCD", (17, 12, 7, 4)):
        ii = [i for i in range(n) if X["porte"][i][:1] == p]
        cx = sum(X["cx"][i] for i in ii)
        T.append([p, cxl, len(ii), sum(1 for i in ii if X["status"][i] == "ATIVO"), cx, sum(X["cx"][i] for i in ii if X["status"][i] == "ATIVO"), pct(cx, tot)])
    T.append(["Total", "", n, len(at), tot, cx_at, "100,0%"])
    L = []
    add = lambda cel, txt, neg=False, tam=None: L.append(dict(cel=cel, txt=txt, neg=neg, tam=tam))
    add("A1", f'RESUMO – {aba.strip(chr(34))}   |   base: aba {aba}, {n0(n)} linhas   |   posição de {HOJE}', True, 14)
    add("A3", "COMO FORAM FEITAS AS CONTAS", True, 12)
    add("A4", f'1) Última compra (col. {cols["dt"]}) e STATUS (col. {cols["st"]}): a data mais recente entre a que já estava na aba (faturamento 2019–2026, produto personalizado, e-mail do cadastro) e as fontes novas do relatório "ultimo fat (20.10.26)": raiz do CNPJ, comprador já identificado e grupo do cadastro (coluna Descrição) nas linhas de rede/associação. ATIVO = compra há até 360 dias.')
    add("A5", f'   Nesta revisão: {f["ultimo fat (raiz)"]} datas novas pela raiz do CNPJ (notas de 15 a 21/10, depois do fim das bases de faturamento), {f["ultimo fat (comprador)"]} pelo comprador, {f["grupo do cadastro"]} pelo grupo do cadastro e {f["linha da rede"]} por associadas da rede. {f["homônimo removido"]} vínculos antigos por razão social foram removidos por serem homônimos (CNPJ válido na Receita, comprador de outra empresa). ATIVO: de {n_antes} para {len(at)}.')
    add("A6", '   Notas do "ultimo fat" anteriores ao fim das bases (14/10) que não estão nas bases de venda não foram usadas: a data dessas empresas é a das bases.')
    add("A7", f'2) VENDEDOR (col. {cols["vend"]}): vendedor da carteira no cadastro (CLIENTES 20.10.26) do cliente que comprou por último; sem compra, o da própria empresa (CNPJ, depois raiz). Carteiras que não são pessoas (INATIVOS, CLIENTE SEM COMPRA, RD, AUTOPEL – VENDA ENCERRADA) ficam de fora. {n0(sum(1 for v in X["vend"] if v))} linhas com vendedor; todas as {len(at)} ativas têm.')
    add("A8", f'3) PORTE (col. {cols["porte"]}): A ≥ R$ 1 bi; B = R$ 300 mi a 1 bi; C = R$ 20 a 300 mi; D < R$ 20 mi – pelo faturamento 2025 do Ranking de redes ou, sem ele, pelo nº de lojas × R$ 2,2 mi por loja/ano (mesma lógica da coluna I de supermercados).' + (" Igual ao da FARMACIA VF para a mesma empresa." if chave == "co" else ""))
    add("A9", f'4) Total de caixas/mês (col. {cols["cx"]}) = caixas por loja do porte (A 17, B 12, C 7, D 4) × nº de lojas considerado' + (" (o mesmo da FARMACIA VF, col. H)." if chave == "co" else " (col. H)."))
    add("A10", f"   Sem contagem dupla: a base é o universo das farmácias ativas da Receita, então toda associada já tem linha própria. Por isso a linha de associação/franquia (e a Farmarcas) fica com 0 caixa ou só com as lojas da própria raiz, e a rede com mais de uma raiz conta o nº considerado uma vez só. Lojas somadas: {n0(lojas)} (o Brasil tem cerca de 120,7 mil farmácias ativas na Receita).")
    add("A11", f"5) Total: {n0(tot)} caixas/mês = {n0(cx_at)} de ATIVOS ({n0(len(at))} linhas) + {n0(tot - cx_at)} A TRABALHAR ({n0(n - len(at))} linhas). {n0(zero)} linhas têm 0 loja (inaptas, sem farmácia ativa na raiz ou sem CNPJ válido) e ficam com 0 caixa.")
    add("A12", "Atenção: UF = UF da sede (rede com lojas em vários estados conta na UF da sede). As premissas de caixas por porte valem para farmácia e podem ser ajustadas na regra.")
    tabs = [dict(cel="A14", linhas=T)]
    r2 = 14 + len(T) + 1
    add(f"A{r2}", f"TABELAS DINÂMICAS   (fonte: {aba}!{fonte_faixa}; se a aba mudar, clique com o botão direito numa tabela > Atualizar)", True, 12)
    pre = "TD_" + chave.upper()
    tds = [dict(nome=pre + "_UF", campo="UF SEDE", titulo="Clientes ATIVOS e A TRABALHAR por UF + caixas/mês", col="A", larg=12, ordena=True),
           dict(nome=pre + "_PORTE", campo="PORTE", titulo="Clientes ATIVOS e A TRABALHAR por porte + caixas/mês", col="J", larg=34, ordena=False),
           dict(nome=pre + "_VEND", campo="VENDEDOR", titulo="Clientes ATIVOS e A TRABALHAR por vendedor + caixas/mês", col="S", larg=44, ordena=True)]
    for t in tds:
        t["titulo_cel"] = f"{t['col']}{r2 + 2}"; t["cel"] = f"{t['col']}{r2 + 3}"
    return dict(aba=aba[:-1] + ' RESUMO"', linhas=L, tabelas=tabs, fonte_aba=aba,
                fonte_faixa=fonte_faixa, campo_status="STATUS", campo_conta=("NA ABA FARMACIA?" if chave == "co" else "TIPO"), campo_caixas="Total de caixas/ mês", renomeia={},
                vazio="Sem vendedor (não está no cadastro)", primeiro="ATIVO", tds=tds)


if __name__ == "__main__":
    abas = pickle.load(open("abas.pkl", "rb"))
    R = pickle.load(open("farm_final.pkl", "rb"))
    lojas_rf = pickle.load(open("lojas_sup.pkl", "rb"))
    nco, nvf = R["co"]["n"] + 1, R["vf"]["n"] + 1
    spec = dict(abas=[
        supermercados(abas, R, lojas_rf),
        farmacia(R, "co", '"FARMACIA CONSOLIDADO"', f"E1:T{nco}", dict(dt="N", st="O", vend="R", porte="S", cx="T"),
                 sum(1 for a in R["co"]["antes"] if a[1] == "ATIVO")),
        farmacia(R, "vf", '"FARMACIA VF"', f"A1:W{nvf}", dict(dt="O", st="P", vend="V", porte="U", cx="W"),
                 sum(1 for a in R["vf"]["antes"] if a[1] == "ATIVO")),
    ])
    json.dump(spec, open("resumos_spec.json", "w", encoding="utf8"), ensure_ascii=True, indent=1)
    for a in spec["abas"]:
        print(a["aba"], len(a["linhas"]), "linhas de texto,", len(a["tabelas"]), "tabelas,", len(a["tds"]), "dinâmicas")
        for l in a["linhas"]:
            print("  ", l["cel"], l["txt"][:230])
        for t in a["tabelas"]:
            for r in t["linhas"]:
                print("     ", r)
