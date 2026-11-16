"""Refaz os textos fixos do topo de "FARMACIA VF RESUMO", "... port A e B" e "I.R.T. T. resumo" com os números do
próprio arquivo (16/11/2026). As dinâmicas se atualizam sozinhas; o texto não. Rodar depois de toda mudança de status.
Gera um tsv para o grava_status.ps1 (aba, coluna, linha, tipo s/n, valor).
Uso: python textos_resumos.py "<MAPA.xlsm já gravado e salvo pelo Excel>" "<saida.tsv>" DD/MM/AAAA"""
import sys, collections as C
import openpyxl

ARQ, SAIDA, DIA = sys.argv[1:4]
br = lambda n: f'{n:,.0f}'.replace(',', '.')
wb = openpyxl.load_workbook(ARQ, read_only=True, data_only=True)
# ---------------- FARMACIA VF
ws = wb['"FARMACIA VF"']
it = ws.iter_rows(min_row=6, values_only=True)
cab = [str(x or '').strip() for x in next(it)]
ix = lambda c: cab.index(c)
iS, iX, iP, iV, iU = ix('STATUS AUTOPEL'), ix('Total de caixas/ mês'), ix('P'), ix('VENDEDOR'), ix('PORTE')
par = {'A': 17, 'B': 12, 'C': 7, 'D': 4}
lin = C.Counter(); cx = C.Counter(); por = C.defaultdict(C.Counter); ab = C.defaultdict(C.Counter)
com_v = ativ_v = zero = zero_ab = 0; lojas = 0.0; n = 0
for r in it:
    if not r or all(x is None for x in r[:8]): continue
    n += 1; st = r[iS]; x = r[iX] or 0; p = r[iP]
    lin[st] += 1; cx[st] += x
    por[p]['l'] += 1; por[p]['x'] += x
    if st == 'ATIVO': por[p]['la'] += 1; por[p]['xa'] += x
    if p in ('A', 'B'): ab[st]['l'] += 1; ab[st]['x'] += x
    if r[iV]: com_v += 1
    if st == 'ATIVO' and r[iV]: ativ_v += 1
    if not x:
        zero += 1
        if p in ('A', 'B'): zero_ab += 1
    if p in par: lojas += x / par[p]
tot = sum(cx.values())
A, I, N = 'ATIVO', 'INATIVO', 'NÃO É CLIENTE'
out = []
def s(a, c, r, v): out.append((a, c, r, 's', v))
def nn(a, c, r, v): out.append((a, c, r, 'n', str(int(v))))
VF = '"FARMACIA VF RESUMO"'; AB = '"FARMACIA VF RESUMO" port A e B'; IR = 'I.R.T. T. resumo'
s(VF, 'A', 1, f'RESUMO – FARMACIA VF   |   base: aba "FARMACIA VF", {br(n)} linhas   |   posição de {DIA}')
s(VF, 'A', 7, f'2) VENDEDOR (col. T): vendedor da carteira no cadastro de clientes do cliente da última compra; quando a carteira não '
              f'é uma pessoa (INATIVOS, CLIENTE SEM COMPRA, RD, AUTOPEL – VENDA ENCERRADA), o vendedor da nota. {br(com_v)} linhas com '
              f'vendedor; {"todas as " + br(lin[A]) + " ativas têm" if ativ_v == lin[A] else br(ativ_v) + " das " + br(lin[A]) + " ativas têm"}.')
s(VF, 'A', 10, '   Sem contagem dupla: a base é o universo das farmácias ativas da Receita, então toda associada já tem linha própria. Por isso '
               'a linha de associação/franquia (e a Farmarcas) fica com 0 caixa ou só com as lojas da própria raiz, e a rede com mais de uma '
               f'raiz conta o nº considerado uma vez só. Lojas somadas: {br(lojas)} (o Brasil tem cerca de 120,7 mil farmácias ativas na Receita).')
s(VF, 'A', 11, f'5) Total: {br(tot)} caixas/mês = {br(cx[A])} de ATIVOS ({br(lin[A])} linhas) + {br(cx[I])} de INATIVOS ({br(lin[I])} linhas) + '
               f'{br(cx[N])} de NÃO É CLIENTE ({br(lin[N])} linhas). {br(zero)} linhas ficam com 0 caixa (associações sem contagem dupla, '
               'inaptas, sem farmácia ativa na raiz ou sem CNPJ válido).')
for r, p in zip(range(15, 19), 'ABCD'):
    nn(VF, 'C', r, por[p]['l']); nn(VF, 'D', r, por[p]['la']); nn(VF, 'E', r, por[p]['x']); nn(VF, 'F', r, por[p]['xa'])
nn(VF, 'C', 19, n); nn(VF, 'D', 19, lin[A]); nn(VF, 'E', 19, tot); nn(VF, 'F', 19, cx[A])
s(AB, 'A', 6, f'3) {br(zero_ab)} dessas linhas ficam com 0 caixa: são associações/franquias (as lojas estão nas linhas das associadas, de '
              f'porte menor) ou redes sem loja ativa. Por isso o porte A/B soma {br(sum(v["x"] for v in ab.values()))} caixas/mês, e não o '
              'total das redes inteiras.')
s(AB, 'A', 1, f'RESUMO – FARMACIA VF, SÓ PORTES A E B   |   base: aba "FARMACIA VF" (coluna S = A ou B), '
              f'{br(por["A"]["l"] + por["B"]["l"])} de {br(n)} linhas   |   posição de {DIA}')
s(AB, 'A', 7, f'4) Total: {br(sum(v["x"] for v in ab.values()))} caixas/mês = {br(ab[A]["x"])} de ATIVOS ({br(ab[A]["l"])} linhas) + '
              f'{br(ab[I]["x"])} de INATIVOS ({br(ab[I]["l"])} linhas) + {br(ab[N]["x"])} de NÃO É CLIENTE ({br(ab[N]["l"])} linhas).')
for r, p in zip((10, 11), 'AB'):
    nn(AB, 'C', r, por[p]['l']); nn(AB, 'D', r, por[p]['la']); nn(AB, 'E', r, por[p]['x']); nn(AB, 'F', r, por[p]['xa'])
nn(AB, 'C', 12, por['A']['l'] + por['B']['l']); nn(AB, 'D', 12, por['A']['la'] + por['B']['la'])
nn(AB, 'E', 12, por['A']['x'] + por['B']['x']); nn(AB, 'F', 12, por['A']['xa'] + por['B']['xa'])
# ---------------- IRTT
ws = wb['I.R.T.T.__Por bandeira (FINAL)']
rows = list(ws.iter_rows(min_row=4, values_only=True))
cab = [str(x or '').strip() for x in rows[0]]
dat = [r for r in rows[1:] if r and r[1]]
iE, iF = cab.index('Status'), cab.index('Status no CRM FINAL')
iAG = next(i for i, c in enumerate(cab) if c.startswith('VENDEDOR (CLIENTES'))
e = C.Counter(r[iE] for r in dat); f = C.Counter(r[iF] for r in dat)
s(IR, 'A', 1, f'RESUMO – I.R.T.T. POR BANDEIRA   |   base: aba I.R.T.T.__Por bandeira (FINAL), {len(dat)} bandeiras   |   posição de {DIA}')
s(IR, 'A', 6, f'2) Status (col. F, fórmula): ativo = compra há até 360 dias; sem compra ou mais antiga = a trabalhar. Hoje: {f["ativo"]} ativas '
              f'e {f["a trabalhar"]} a trabalhar (col. E: {e["Ativo"]} Ativo, {e["Inativo"]} Inativo, {e["Cadastrado sem compra"]} '
              f'Cadastrado sem compra e {e["Não atendido"]} Não atendido).')
s(IR, 'A', 7, f'3) Vendedor (col. AG): carteira no cadastro de clientes do cliente da última compra; quando a carteira não é uma pessoa '
              f'(INATIVOS, CLIENTE SEM COMPRA, VENDA ENCERRADA…), o vendedor da nota. {sum(1 for r in dat if r[iAG])} bandeiras com '
              'vendedor. O título da coluna AG ainda diz 20.10.26 porque uma tabela dinâmica usa o campo; os dados são de 16/11.')
wb.close()
assert f['ativo'] == e['Ativo'], ('coluna F (fórmula) diferente da E', f, e)
with open(SAIDA, 'w', encoding='utf-8-sig', newline='') as fo:
    for x in out: fo.write('\t'.join(map(str, x)) + '\r\n')
print(len(out), 'células | VF', dict(lin), 'caixas', br(tot), '| A/B', {k: dict(v) for k, v in ab.items()}, '| IRTT F', dict(f))
