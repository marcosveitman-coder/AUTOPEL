"""Conferência final do MAPA (16/11/2026), só leitura, tudo a partir do próprio arquivo:
  1. dinâmicas da FARMACIA VF RESUMO, da "port A e B", da I.R.T. T. resumo e da Alessandra Dayan VF x os dados das abas;
  2. tabela fixa e texto do total dos resumos x os dados;
  3. potencial do ORGANOGRAMA impressao (redes, PDV, caixas por segmento) x as abas;
  4. nenhuma célula de erro (#REF!, #N/D...) nas abas do pedido;
  5. exemplos: redes que precisam estar ATIVO (argumentos, nomes de linha das abas FINAL).
Uso: python confere_final.py "<MAPA.xlsm>" ["Rede 1" "Rede 2" ...]"""
import sys, re, collections as C
import openpyxl

ARQ, EXEMPLOS = sys.argv[1], sys.argv[2:]
wb = openpyxl.load_workbook(ARQ, read_only=True, data_only=True)
br = lambda n: f'{n:,.0f}'.replace(',', '.')
falhas = []
def ok(cond, txt):
    print(('OK    ' if cond else 'FALHA ') + txt)
    if not cond: falhas.append(txt)

# ---- dados da FARMACIA VF
it = wb['"FARMACIA VF"'].iter_rows(min_row=6, values_only=True)
cab = [str(x or '').strip() for x in next(it)]
iS, iX, iP, iUF = cab.index('STATUS AUTOPEL'), cab.index('Total de caixas/ mês'), cab.index('P'), cab.index('UF SEDE')
lin, cx = C.Counter(), C.Counter(); ab_l, ab_x = C.Counter(), C.Counter(); por = C.defaultdict(C.Counter)
for r in it:
    if not r or all(x is None for x in r[:8]): continue
    st, x, p = r[iS], r[iX] or 0, r[iP]
    lin[st] += 1; cx[st] += x; por[p]['la' if st == 'ATIVO' else 'lo'] += 1; por[p]['xa' if st == 'ATIVO' else 'xo'] += x
    if p in ('A', 'B'): ab_l[st] += 1; ab_x[st] += x
A, I, N = 'ATIVO', 'INATIVO', 'NÃO É CLIENTE'
def linha(aba, rot, col=0, maxc=16):
    for r in wb[aba].iter_rows(values_only=True):
        if r and len(r) > col and r[col] == rot: return r[:maxc]
    return None
# 1) dinâmica por UF da FARMACIA VF RESUMO (Total Geral: NÃO clientes, caixas, INATIVO, caixas, ATIVO, caixas, total, total caixas)
t = linha('"FARMACIA VF RESUMO"', 'Total Geral')
ok(t is not None and list(t[1:9]) == [lin[N], cx[N], lin[I], cx[I], lin[A], cx[A], sum(lin.values()), sum(cx.values())],
   f'FARMACIA VF RESUMO dinâmica = dados ({lin[A]} ativos, {br(cx[A])} caixas)')
t = linha('"FARMACIA VF RESUMO" port A e B', 'Total Geral')
ok(t is not None and list(t[1:9]) == [ab_l[N], ab_x[N], ab_l[I], ab_x[I], ab_l[A], ab_x[A], sum(ab_l.values()), sum(ab_x.values())],
   f'port A e B dinâmica = dados ({ab_l[A]} ativos, {br(ab_x[A])} caixas)')
# 2) tabela fixa e texto
t = linha('"FARMACIA VF RESUMO"', 'Total')
ok(t is not None and t[3] == lin[A] and t[5] == cx[A], 'FARMACIA VF RESUMO tabela fixa (Linhas ATIVO, Caixas ATIVO)')
txt = ' '.join(str(x) for r in wb['"FARMACIA VF RESUMO"'].iter_rows(max_row=12, max_col=1, values_only=True) for x in r if x)
ok(f'{br(cx[A])} de ATIVOS ({br(lin[A])} linhas)' in txt, 'FARMACIA VF RESUMO texto do total')
t = linha('"FARMACIA VF RESUMO" port A e B', 'Total')
ok(t is not None and t[3] == ab_l[A] and t[5] == ab_x[A], 'port A e B tabela fixa')
# IRTT
rows = list(wb['I.R.T.T.__Por bandeira (FINAL)'].iter_rows(min_row=4, values_only=True))
cab = [str(x or '').strip() for x in rows[0]]
f = C.Counter(r[cab.index('Status no CRM FINAL')] for r in rows[1:] if r and r[1])
e = C.Counter(r[cab.index('Status')] for r in rows[1:] if r and r[1])
ok(f['ativo'] == e['Ativo'], f'IRTT coluna F (fórmula) = coluna E ({f["ativo"]} ativas)')
t = linha('I.R.T. T. resumo', 'ativo')
ok(t is not None and t[1] == f['ativo'], 'I.R.T. T. resumo dinâmica = dados')
txt = ' '.join(str(x) for r in wb['I.R.T. T. resumo'].iter_rows(max_row=9, max_col=1, values_only=True) for x in r if x)
ok(f'Hoje: {f["ativo"]} ativas e {f["a trabalhar"]} a trabalhar' in txt, 'I.R.T. T. resumo texto')
# Alessandra Dayan VF (filtro STATUS RECEITA = ATIVA)
rows = list(wb['Alessandra Dayan VF'].iter_rows(values_only=True))
cnt = C.defaultdict(C.Counter)
for r in rows[1:]:
    if r and r[0] and r[4] == 'ATIVA': cnt[r[7]][r[6]] += 1
dif = [r[11] for r in rows[4:40] if r and r[11] and r[11] in cnt and tuple(x or 0 for x in r[12:15]) != (cnt[r[11]][A], cnt[r[11]][I], cnt[r[11]][N])]
ok(not dif and any(r and r[11] in cnt for r in rows[4:40]), f'Alessandra Dayan VF dinâmica = dados ({len(cnt)} segmentos) {dif}')
# 3) organograma x abas
def pot(aba):
    it = wb[aba].iter_rows(min_row=6, values_only=True); cab = [str(x or '').strip() for x in next(it)]
    iS = cab.index('STATUS AUTOPEL'); iL = next(i for i, c in enumerate(cab) if c.upper().startswith('Nº DE LOJAS')); iX = next(i for i, c in enumerate(cab) if c.upper().startswith('TOTAL DE CAIXAS'))
    t = red = pdv = cxs = 0
    for r in it:
        if not r or not (r[1] or r[4]): continue
        t += 1
        if r[iS] != 'ATIVO':
            red += 1; pdv += r[iL] if isinstance(r[iL], (int, float)) else 0; cxs += r[iX] if isinstance(r[iX], (int, float)) else 0
    return t, red, round(pdv), round(cxs)
for r in wb['ORGANOGRAMA impressao'].iter_rows(values_only=True):
    v = [x for x in r if x is not None]
    if len(v) >= 9 and isinstance(v[1], str) and v[1].strip().endswith('FINAL (W)'):
        aba = v[1].strip()[:-4]
        t, red, pdv, cxs = pot(aba)
        ok((v[2], v[6], v[7], v[8]) == (t, red, pdv, cxs), f'organograma {aba}: {red} redes, {br(pdv)} PDV, {br(cxs)} caixas')
    if len(v) >= 9 and isinstance(v[1], str) and 'FARMACIA VF' in v[1]:
        ok(v[2] == sum(lin.values()) and (v[3], v[4], v[5]) == (lin[A], lin[I], lin[N]) and v[6] == lin[I] + lin[N] and v[8] == cx[I] + cx[N],
           f'organograma Drogarias: {lin[I] + lin[N]} redes, {br(cx[I] + cx[N])} caixas')
# 4) células de erro
ERR = {'#REF!', '#N/D', '#N/A', '#VALOR!', '#VALUE!', '#DIV/0!', '#NOME?', '#NAME?', '#NULO!', '#NUM!'}
ABAS = [s for s in wb.sheetnames if s.endswith('FINAL') or s.endswith('(TOP 50)') or s.startswith('Alessandra') or 'FARMACIA VF' in s
        or s.startswith('I.R.T') or s.startswith('ORGANOGRAMA') or s == 'TAB DIN. I.R.T.T.']
erros = C.Counter()
for s in ABAS:
    for r in wb[s].iter_rows(values_only=True):
        for x in r:
            if isinstance(x, str) and x.strip() in ERR: erros[s] += 1
ok(not erros, f'nenhuma célula de erro em {len(ABAS)} abas {dict(erros)}')
# 5) exemplos
if EXEMPLOS:
    achou = {}
    for s in [x for x in wb.sheetnames if x.endswith('FINAL')]:
        it = wb[s].iter_rows(min_row=6, values_only=True); cab = [str(x or '').strip() for x in next(it)]
        iS = cab.index('STATUS AUTOPEL')
        for r in it:
            if r and r[1] in EXEMPLOS: achou[r[1]] = r[iS]
    for x in EXEMPLOS: ok(achou.get(x) == 'ATIVO', f'exemplo {x}: {achou.get(x)}')
wb.close()
print('\nRESULTADO:', 'TUDO CERTO' if not falhas else f'{len(falhas)} FALHA(S)')
