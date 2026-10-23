"""Conteudo das abas "I.R.T. T. resumo" e '"FARMACIA VF RESUMO" port A e B' (23/10/2026), lido por td_resumos.ps1.
Numeros tirados da planilha ja gravada (argumento 1) e de confere_g.pkl / vendedor_ag.pkl.  Saida: resumos_23_spec.json"""
import json, pickle, sys, datetime as dt, collections as C
import openpyxl

n0 = lambda v: f"{v:,.0f}".replace(",", ".")
HOJE = "23/10/2026"
ABA_IRTT = "I.R.T.T.__Por bandeira (FINAL)"
AB = ["A – Grande (≥ R$ 1 bi)", "B – Média (R$ 300 mi a 1 bi)"]


def irtt(rows, O):
    dados = [r for r in rows[4:] if r and r[1]]
    ult = 4 + len(dados)
    st = C.Counter(r[5] for r in dados)
    mud = sum(1 for o in O.values() if o["forte"] and o["g0"] and o["forte"][0] > o["g0"]) + sum(1 for o in O.values() if o["forte"] and not o["g0"])
    L = []
    add = lambda cel, txt, neg=False, tam=None: L.append(dict(cel=cel, txt=txt, neg=neg, tam=tam))
    add("A1", f"RESUMO – I.R.T.T. POR BANDEIRA   |   base: aba {ABA_IRTT}, {n0(len(dados))} bandeiras   |   posição de {HOJE}", True, 14)
    add("A3", "COMO FORAM FEITAS AS CONTAS", True, 12)
    add("A4", f'1) Última compra (col. G): conferida com as bases de faturamento 2019–2026 e o relatório "ultimo fat" (notas até 21/10) – CNPJ validado (col. Y), códigos de cliente (col. S), grupo do cadastro (col. R = coluna Descrição do relatório) e os vínculos por nome, produto personalizado e e-mail já validados. {mud} datas ficaram mais recentes.')
    add("A5", "   Vínculos por código ou grupo que eram de outra empresa foram descartados (distribuidor, courier, financeira, outra rede da mesma central de compras, homônimo de outra UF). Em franquias (ex.: O Boticário, Cacau Show) vale a compra do franqueado.")
    add("A6", f"2) Status (col. F, fórmula): ativo = compra há até 360 dias; sem compra ou mais antiga = a trabalhar. Hoje: {st['ativo']} ativas e {st['a trabalhar']} a trabalhar.")
    add("A7", f"3) Vendedor (col. AG): carteira no cadastro CLIENTES (20.10.26) do cliente da última compra; sem ele, os códigos de cliente (S) e o CNPJ validado (Y). {sum(1 for r in dados if r[32])} bandeiras com vendedor; carteiras que não são pessoas (INATIVOS, CLIENTE SEM COMPRA, VENDA ENCERRADA…) ficam de fora.")
    add("A8", f"4) UF (col. AH): UF da sede na Receita, pelo CNPJ validado (Y); {sum(1 for r in dados if r[33] == '(sem CNPJ)')} bandeiras sem CNPJ ficam como \"(sem CNPJ)\". A coluna Q lista várias UFs por bandeira e por isso não serve para a tabela.")
    add("A9", "Atenção: cada linha é uma bandeira – um grupo com várias bandeiras (ex.: GPA, Magalu, Grupo Trigo) conta uma vez por bandeira.")
    add("A11", f"TABELAS DINÂMICAS   (fonte: {ABA_IRTT}!A4:AH{ult}; se a aba mudar, clique com o botão direito numa tabela > Atualizar)", True, 12)
    tds = [dict(nome="TD_IRTT_STATUS", campo="Status no CRM FINAL", titulo="Bandeiras ativas e a trabalhar", col="A", larg=22, ordena=False),
           dict(nome="TD_IRTT_UF", campo="UF SEDE (Receita, CNPJ validado)", titulo="Bandeiras ativas e a trabalhar por UF (sede)", col="E", larg=14, ordena=True),
           dict(nome="TD_IRTT_SEG", campo="Segmento (IRTT)", titulo="Bandeiras ativas e a trabalhar por segmento", col="K", larg=58, ordena=True)]
    for t in tds:
        t["titulo_cel"] = f"{t['col']}13"; t["cel"] = f"{t['col']}14"
    return dict(aba="I.R.T. T. resumo", linhas=L, tabelas=[], fonte_aba=ABA_IRTT, fonte_faixa=f"A4:AH{ult}",
                campo_status="Status no CRM FINAL", campo_conta="Bandeira", campo_caixas="", renomeia={},
                vazio="(sem CNPJ)", primeiro="ativo", tds=tds)


def vf_ab(rows, cab):
    hdr = rows[cab - 1]
    ix = {h: i for i, h in enumerate(hdr)}
    dados = [r for r in rows[cab:] if r and r[0] is not None]
    ult = cab + len(dados)
    ab = [r for r in dados if r[ix["PORTE"]] in AB]
    at = [r for r in ab if r[ix["STATUS AUTOPEL"]] == "ATIVO"]
    cx = lambda rs: sum(r[ix["Total de caixas/ mês"]] or 0 for r in rs)
    zero = sum(1 for r in ab if not (r[ix["Total de caixas/ mês"]] or 0))
    T = [["Porte", "Caixas por loja", "Linhas", "Linhas ATIVO", "Caixas/mês", "Caixas/mês ATIVO"]]
    for p, k in zip(AB, (17, 12)):
        rs = [r for r in ab if r[ix["PORTE"]] == p]
        T.append([p, k, len(rs), sum(1 for r in rs if r[ix["STATUS AUTOPEL"]] == "ATIVO"), cx(rs), cx([r for r in rs if r[ix["STATUS AUTOPEL"]] == "ATIVO"])])
    T.append(["Total", "", len(ab), len(at), cx(ab), cx(at)])
    L = []
    add = lambda cel, txt, neg=False, tam=None: L.append(dict(cel=cel, txt=txt, neg=neg, tam=tam))
    add("A1", f'RESUMO – FARMACIA VF, SÓ PORTES A E B   |   base: aba "FARMACIA VF" (coluna U = A ou B), {n0(len(ab))} de {n0(len(dados))} linhas   |   posição de {HOJE}', True, 14)
    add("A3", "COMO LER", True, 12)
    add("A4", '1) Mesmas contas da aba "FARMACIA VF RESUMO", só para as linhas de porte A – Grande (≥ R$ 1 bi) e B – Média (R$ 300 mi a 1 bi). As tabelas dinâmicas têm o filtro de porte já aplicado.')
    add("A5", '2) Total de caixas/mês (col. X) = caixas por loja do porte (parâmetros X1:X4 da aba: A 17, B 12) × nº de lojas considerado (col. H), pela fórmula SE da coluna X; nas linhas de associação/rede a fórmula usa o nº de lojas sem contagem dupla.')
    add("A6", f"3) {zero} dessas linhas ficam com 0 caixa: são associações/franquias (as lojas estão nas linhas das associadas, de porte menor) ou redes sem loja ativa. Por isso o porte A/B soma {n0(cx(ab))} caixas/mês, e não o total das redes inteiras.")
    add("A7", f"4) Total: {n0(cx(ab))} caixas/mês = {n0(cx(at))} de ATIVOS ({len(at)} linhas) + {n0(cx(ab) - cx(at))} A TRABALHAR ({len(ab) - len(at)} linhas).")
    r2 = 9 + len(T) + 1
    add(f"A{r2}", f'TABELAS DINÂMICAS   (fonte: "FARMACIA VF"!A{cab}:X{ult}, filtro PORTE = A e B; se a aba mudar, clique com o botão direito numa tabela > Atualizar)', True, 12)
    tds = [dict(nome="TD_VFAB_UF", campo="UF SEDE", titulo="Porte A e B: ATIVOS e A TRABALHAR por UF + caixas/mês", col="A", larg=12, ordena=True),
           dict(nome="TD_VFAB_PORTE", campo="PORTE", titulo="Porte A e B: ATIVOS e A TRABALHAR por porte + caixas/mês", col="J", larg=34, ordena=False),
           dict(nome="TD_VFAB_VEND", campo="VENDEDOR", titulo="Porte A e B: ATIVOS e A TRABALHAR por vendedor + caixas/mês", col="S", larg=44, ordena=True)]
    for t in tds:
        t["titulo_cel"] = f"{t['col']}{r2 + 2}"; t["cel"] = f"{t['col']}{r2 + 5}"   # espaco para o filtro de porte
    return dict(aba='"FARMACIA VF RESUMO" port A e B', linhas=L, tabelas=[dict(cel="A9", linhas=T)], fonte_aba='"FARMACIA VF"',
                fonte_faixa=f"A{cab}:X{ult}", campo_status="STATUS AUTOPEL", campo_conta="TIPO", campo_caixas="Total de caixas/ mês",
                renomeia={}, vazio="Sem vendedor (não está no cadastro)", primeiro="ATIVO", tds=tds,
                filtro=dict(campo="PORTE", manter=AB))


if __name__ == "__main__":
    wb = openpyxl.load_workbook(sys.argv[1], read_only=True, data_only=True)
    ri = list(wb[ABA_IRTT].iter_rows(values_only=True)); rv = list(wb['"FARMACIA VF"'].iter_rows(values_only=True)); wb.close()
    cab = next(i + 1 for i, r in enumerate(rv) if r and r[0] == "POSIÇÃO")
    O = pickle.load(open("confere_g.pkl", "rb"))["out"]
    spec = dict(abas=[irtt(ri, O), vf_ab(rv, cab)])
    json.dump(spec, open("resumos_23_spec.json", "w", encoding="utf8"), ensure_ascii=True, indent=1)
    for a in spec["abas"]:
        print(a["aba"], a["fonte_faixa"])
        for l in a["linhas"]:
            print("  ", l["cel"], l["txt"][:170])
        for t in a["tabelas"]:
            for r in t["linhas"]:
                print("     ", r)
