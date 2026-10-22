"""Monta o conteudo das abas de resumo do MAPA MERCADO (22/10/2026): textos, tabelas pequenas e a definicao das
tabelas dinamicas.  Todos os numeros saem dos dados (nada digitado a mao).  Saida: resumos_spec.json, lido por
td_resumos.ps1, que grava pelo proprio Excel.

Entradas (na pasta de trabalho):
  abas.pkl       linhas das abas SUPERMERCADOS (reuniao), "FARMACIA CONSOLIDADO" e "FARMACIA VF" (valores)
  farm_final.pkl resultado de ../farmacias/final/farm_final.py
  consolida.pkl  resultado de consolida_reuniao.py (nº de lojas preenchido e linhas unificadas)
  abas_antes.pkl a aba da reuniao antes desses ajustes (numeros de antes/depois)
"""
import json, pickle, re, collections as C
from openpyxl.utils import column_index_from_string as ci

D = lambda x: re.sub(r"\D", "", str(x or ""))
n0 = lambda v: f"{v:,.0f}".replace(",", ".")
n1 = lambda v: f"{v:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
pct = lambda a, b: f"{100 * a / b:.1f}%".replace(".", ",")
HOJE = "22/10/2026"


def supermercados(abas, R, K, antes):
    """K: consolida.pkl (lojas preenchidas e linhas unificadas); antes: linhas da aba antes dos ajustes de 22/10."""
    a = abas["SUPERMERCADOS (reuniao)"]
    g = lambda r, c: r[ci(c) - 1]
    rows = [r for r in a[4:] if any(v is not None for v in r)]
    ult = 4 + len(rows)
    ativo = lambda r: g(r, "AL") == "ativo"
    AI = {"A": 20, "B": 15, "C": 10, "D": 8}
    tot = sum(g(r, "AJ") for r in rows)
    at = [r for r in rows if ativo(r)]
    cx_at = sum(g(r, "AJ") for r in at)
    old = {i + 1: r for i, r in enumerate(antes) if i >= 4 and any(v is not None for v in r)}
    tot_antes = sum(g(r, "AJ") for r in old.values())
    ex = set(K["excluir"])
    # nº de lojas preenchido (linhas que estavam com L = 0 e ficaram)
    l0 = {n: v for n, v in K["fonte"].items() if n not in ex}
    ganhou = [n for n in l0 if K["cel"].get(n, {}).get("L", 0) > 0]
    lojas_novas = sum(K["cel"][n]["L"] for n in ganhou)
    cx_novas = sum(AI[g(old[n], "J")] * K["cel"][n]["L"] for n in ganhou)
    por = C.Counter()
    for n in ganhou:
        f = K["fonte"][n]
        por["Receita" if f.startswith("Receita") else "internet" if f.startswith("Internet") else "mínimo"] += 1
    zero = C.Counter()
    for n, f in l0.items():
        if K["cel"].get(n, {}).get("L", 0) == 0:
            zero["inapto" if "inapto" in f else "outro ramo" if "sem estabelecimento" in f else "outros"] += 1
    # linhas repetidas
    cx_saiu = sum(g(old[n], "AJ") for n in ex)
    T1 = [["Porte (coluna J)", "Caixas por loja (AI)", "Linhas", "Linhas ativas", "Lojas (L)", "Caixas/mês (AJ)", "% das caixas"]]
    for p in "ABCD":
        rr = [r for r in rows if g(r, "J") == p]
        cx = sum(g(r, "AJ") for r in rr)
        T1.append([p, AI[p], len(rr), sum(1 for r in rr if ativo(r)), sum(g(r, "L") or 0 for r in rr), cx, pct(cx, tot)])
    T1.append(["Total", "", len(rows), len(at), sum(g(r, "L") or 0 for r in rows), tot, "100,0%"])
    vf = R["vf"]
    vfr = {c[:8]: i for i, c in enumerate(vf["cnpj"]) if len(c) == 14}
    roots = {D(g(r, "F"))[:8] for r in rows if len(D(g(r, "F"))) == 14}
    ov = [k for k in roots if k in vfr and vf["cx"][vfr[k]] > 0]
    sem_vend = sum(1 for r in rows if not ativo(r) and not g(r, "AN"))
    ex_cli = sum(1 for r in rows if not ativo(r) and g(r, "AN"))
    rede_sup = sum(g(r, "L") or 0 for r in rows if str(g(r, "D")).startswith("Rede Supermarket"))
    L = []
    add = lambda cel, txt, neg=False, tam=None: L.append(dict(cel=cel, txt=txt, neg=neg, tam=tam))
    add("A1", f"RESUMO – SUPERMERCADOS (REUNIÃO)   |   base: aba SUPERMERCADOS (reuniao), {n0(len(rows))} linhas   |   posição de {HOJE}", True, 14)
    add("A3", "PARTE 1 – RACIONAL DAS CONTAS (colunas AI e AJ)", True, 12)
    add("A4", '1) Coluna AI "QUANTIDADE DE CAIXA POR LOJA": premissa fixa pelo porte da coluna J – A = 20, B = 15, C = 10 e D = 8 caixas por loja/mês. Fórmula: =SE(J5="A";20;SE(J5="B";15;SE(J5="C";10;SE(J5="D";8)))).')
    add("A5", '2) Coluna AJ "Total de caixas/ mês" = AI × L (nº de lojas da coluna L, "nª de LOJA (antonio)"). Exemplo, linha 5: porte A com 500 lojas → 20 × 500 = 10.000 caixas/mês.')
    add("A6", "3) Porte (colunas I e J): A = faturamento 2025 ≥ R$ 1 bi; B = R$ 300 mi a 1 bi; C = abaixo de R$ 300 mi; D = pequena/local, sem faturamento publicado (classificada pelo capital social e nº de lojas, coluna V).")
    add("A7", '4) Status (coluna AL "status final (reuniao)"): ativo = última compra (coluna AK) há até 360 dias; sem compra ou mais antiga = a trabalhar.')
    add("A8", f"5) Total (célula AJ3, que respeita o filtro): {n0(tot)} caixas/mês = {n0(cx_at)} de clientes ativos ({n0(len(at))} linhas, {pct(cx_at, tot)}) + {n0(tot - cx_at)} a trabalhar ({n0(len(rows) - len(at))} linhas).")
    tabs = [dict(cel="A10", linhas=T1)]
    r0 = 10 + len(T1) + 1
    add(f"A{r0}", f"AJUSTES FEITOS NA ABA EM {HOJE}   (antes: {n0(len(old))} linhas e {n0(tot_antes)} caixas/mês; agora: {n0(len(rows))} linhas e {n0(tot)} caixas/mês)", True, 12)
    add(f"A{r0 + 1}", f"a) Nº de lojas (coluna L) das linhas que estavam com 0: {n0(len(ganhou))} linhas ganharam {n0(lojas_novas)} lojas (+{n0(cx_novas)} caixas) – {por['Receita']} pela Receita Federal (estabelecimentos ativos de supermercado/atacarejo do CNPJ, out/2026), {por['internet']} pela internet e {por['mínimo']} com o mínimo de 1 loja (empresa com faturamento no Ranking ABRAS, mas CNPJ não localizado).")
    add(f"A{r0 + 2}", f"   Continuam com 0 ({n0(len(l0) - len(ganhou))} linhas): {zero['inapto']} com CNPJ inapto na Receita (nenhuma loja ativa), {zero['outro ramo']} com CNPJ ativo de outro ramo e {zero['outros']} outras (associação, minimercado autônomo, CNPJ baixado). A fonte de cada número está na nova coluna AO.")
    add(f"A{r0 + 3}", f"b) Linhas repetidas: {n0(len(ex))} linhas excluídas ({len(K['grupos'])} empresas apareciam mais de uma vez; repetição confirmada pelo mesmo CNPJ, pela mesma posição no ranking ou por serem bandeiras da mesma rede, ex.: Plurix, Mart Minas/Dom Atacadista, Super Bom/Costazul). Elas somavam {n0(cx_saiu)} caixas. Conta: {n0(tot_antes)} + {n0(cx_novas)} − {n0(cx_saiu)} = {n0(tot)}.")
    add(f"A{r0 + 4}", "   Ficou uma linha por empresa (a do ranking, de maior porte), com o nº de lojas da empresa e a compra mais recente do grupo (data, status e vendedor). A coluna AO diz com quais nomes ela também aparecia.")
    add(f"A{r0 + 5}", "PONTOS DE ATENÇÃO", True, 12)
    add(f"A{r0 + 6}", f"c) Rede Supermarket (linha da associação, {n0(rede_sup)} lojas): é uma rede de associados, e algumas associadas também têm linha própria (ex.: Alvorada, com 43 lojas). Pode haver lojas contadas nas duas.")
    add(f"A{r0 + 7}", "d) Market4u: 2.554 minimercados autônomos em condomínios (sem caixa nem bobina) – deixado com 0 loja. Se quiser contar, L = 2.554 (+38.310 caixas).")
    add(f"A{r0 + 8}", 'e) A AL está escrita "a tarbalhar"; nas tabelas dinâmicas aparece como "a trabalhar" (só o rótulo). UF = UF da linha na lista: lojas de uma rede em outros estados contam nessa UF.')
    add(f"A{r0 + 9}", f'f) Vendedor (AN): vazio em {n0(sem_vend)} linhas a trabalhar (empresas sem cadastro na Autopel), que aparecem como "Sem vendedor"; outras {ex_cli} linhas a trabalhar já têm vendedor (ex-clientes). {len(ov)} empresas desta lista também estão nas abas de farmácia (pelas lojas de drogaria): não somar os dois potenciais sem descontar.')
    r2 = r0 + 11
    add(f"A{r2}", f"PARTE 2 – TABELAS DINÂMICAS   (fonte: SUPERMERCADOS (reuniao)!A4:AN{ult}; se a aba mudar, clique com o botão direito numa tabela > Atualizar)", True, 12)
    tds = [dict(nome="TD_SUP_UF", campo="UF", titulo="Clientes ativos e a trabalhar por UF + caixas/mês (AJ)", col="A", larg=12, ordena=True),
           dict(nome="TD_SUP_PORTE", campo="Porte", titulo="Clientes ativos e a trabalhar por porte (I) + caixas/mês (AJ)", col="J", larg=30, ordena=False),
           dict(nome="TD_SUP_VEND", campo="Vendedor", titulo="Clientes ativos e a trabalhar por vendedor (AN) + caixas/mês (AJ)", col="S", larg=44, ordena=True)]
    for t in tds:
        t["titulo_cel"] = f"{t['col']}{r2 + 2}"; t["cel"] = f"{t['col']}{r2 + 3}"
    return dict(aba="RESUMO SUPERMERCADO (REUNIAO)", linhas=L, tabelas=tabs, fonte_aba="SUPERMERCADOS (reuniao)",
                fonte_faixa=f"A4:AN{ult}", campo_status="status final (reuniao)", campo_conta="Nome fantasia",
                campo_caixas="Total de caixas/ mês", renomeia={"a tarbalhar": "a trabalhar"}, vazio="Sem vendedor",
                primeiro="ativo", tds=tds)


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
    K = pickle.load(open("consolida.pkl", "rb"))
    antes = pickle.load(open("abas_antes.pkl", "rb"))["SUPERMERCADOS (reuniao)"]
    nco, nvf = R["co"]["n"] + 1, R["vf"]["n"] + 1
    spec = dict(abas=[
        supermercados(abas, R, K, antes),
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
