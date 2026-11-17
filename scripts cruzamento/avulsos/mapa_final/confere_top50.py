"""Confere a gravação do top50_cnpj_caixas.ps1 (17/11/2026), só leitura, arquivo de partida x arquivo gravado:
  1. nas 11 "(TOP 50)" e na "Alessandra Dayan VF": cada célula antiga (A:H e, na VF, a dinâmica L:P) está igual, só
     descida (3 ou 5 linhas); cabeçalho na linha 6; tabela a/b/c/d nas linhas 1-4;
  2. coluna I = CNPJ do tsv (e vazia fora dele); coluna J = valor do porte x lojas, recalculado aqui;
  3. nenhuma célula de erro nessas abas;
  4. todas as outras abas iguais célula a célula (menos as listadas em --exceto).
Uso: python confere_top50.py "<partida.xlsm>" "<gravado.xlsm>" "<cnpj.tsv>" [--exceto "ABA1" "ABA2"]"""
import sys, collections as C
import openpyxl
A0, A1, TSV = sys.argv[1:4]
EXC = set(sys.argv[sys.argv.index('--exceto') + 1:]) if '--exceto' in sys.argv else set()
DEM, SUP = [17, 12, 7, 4], [20, 15, 10, 8]
ERR = {'#REF!', '#N/D', '#N/A', '#VALOR!', '#VALUE!', '#DIV/0!', '#NOME?', '#NAME?', '#NULO!', '#NUM!'}
falhas = []
def ok(c, t):
    print(('OK    ' if c else 'FALHA ') + t)
    if not c: falhas.append(t)
cn = C.defaultdict(dict)
for l in open(TSV, encoding='utf-8'):
    p = l.rstrip('\r\n').split('\t')
    if len(p) >= 4: cn[p[0]][int(p[1])] = p[3]
w0 = openpyxl.load_workbook(A0, read_only=True, data_only=True)
w1 = openpyxl.load_workbook(A1, read_only=True, data_only=True)
ok(w0.sheetnames == w1.sheetnames, 'mesmas abas, mesma ordem')
ABAS = [a for a in w0.sheetnames if a.endswith('(TOP 50)') or a == 'Alessandra Dayan VF']
for aba in ABAS:
    vf = aba == 'Alessandra Dayan VF'; ins = 5 if vf else 3
    r0 = [list(r) for r in w0[aba].iter_rows(values_only=True)]
    r1 = [list(r) for r in w1[aba].iter_rows(values_only=True)]
    g = lambda rows, i, j: rows[i][j] if i < len(rows) and j < len(rows[i]) else None
    cols = list(range(8)) + (list(range(11, 16)) if vf else [])
    dif = [(i + 1, j) for i in range(len(r0)) for j in cols if g(r0, i, j) != g(r1, i + ins, j)]
    ok(not dif, f'{aba}: células antigas iguais, {ins} linhas abaixo {dif[:5]}')
    ok(g(r1, 5, 0) == 'EMPRESA' and g(r1, 5, 8) == 'CNPJ' and g(r1, 5, 9) == 'Total de caixas/ mês', f'{aba}: cabeçalho na linha 6')
    par = SUP if aba.startswith('SUPERMERCADO') else DEM
    ok([g(r1, i, 8) for i in range(4)] == list('abcd') and [g(r1, i, 9) for i in range(4)] == par
       and (not vf or [g(r1, i, 10) for i in range(4)] == SUP), f'{aba}: tabela a/b/c/d {par}')
    errados, n_cnpj, soma = [], 0, 0
    for i in range(6, len(r1)):
        if not g(r1, i, 0): continue
        lin = i + 1 - ins
        porte, lojas, seg = str(g(r1, i, 5) or ''), g(r1, i, 2), g(r1, i, 7)
        p = (SUP if (vf and seg == 'SUPERMERCADO / ATACAREJO') else par)
        esp = p['ABCD'.index(porte[:1])] * (lojas if isinstance(lojas, (int, float)) else 0) if porte[:1] in 'ABCD' and porte else 0
        if g(r1, i, 9) != esp: errados.append((i + 1, g(r1, i, 9), esp))
        soma += esp
        if g(r1, i, 8) != cn[aba].get(lin): errados.append((i + 1, 'CNPJ', g(r1, i, 8), cn[aba].get(lin)))
        n_cnpj += bool(g(r1, i, 8))
    ok(not errados, f'{aba}: caixas e CNPJ conferidos ({n_cnpj} CNPJs, {soma:,.0f} caixas/mês) {errados[:4]}'.replace(',', '.'))
    e = sum(1 for r in r1 for x in r if isinstance(x, str) and x.strip() in ERR)
    ok(e == 0, f'{aba}: sem célula de erro ({e})')
outras = [a for a in w0.sheetnames if a not in ABAS and a not in EXC]
for aba in outras:
    it0, it1 = w0[aba].iter_rows(values_only=True), w1[aba].iter_rows(values_only=True)
    d = 0
    for a, b in zip(it0, it1):
        if tuple(a) != tuple(b):
            a, b = list(a), list(b); n = max(len(a), len(b)); a += [None] * (n - len(a)); b += [None] * (n - len(b))
            if a != b: d += 1
    resto = sum(1 for _ in it0) + sum(1 for _ in it1)
    if d or resto: ok(False, f'outra aba mudou: {aba} ({d} linhas diferentes, {resto} a mais)')
print(f'outras abas conferidas: {len(outras)} (fora: {sorted(EXC)})')
w0.close(); w1.close()
print('\nRESULTADO:', 'TUDO CERTO' if not falhas else f'{len(falhas)} FALHA(S)')
