"""Confere a cópia gravada pelo grava_status.ps1 (16/11/2026):
  1. cada célula do gravar.tsv está no arquivo (FARMACIA VF: colunas depois da H andaram 2 com a exclusão de I e J);
  2. as outras abas estão iguais célula a célula ao arquivo de partida (fora as gravadas e as das dinâmicas atualizadas);
  3. nas abas gravadas, as células que não estavam no tsv continuam iguais (FARMACIA VF: comparando com o deslocamento);
  4. dinâmicas e caches: mesma quantidade; nomes definidos.
Uso: python confere_gravacao.py "<partida.xlsm>" "<copia.xlsm>" "<gravar.tsv>" """
import sys, re, zipfile, datetime as dt, collections as C
import openpyxl
from openpyxl.utils import column_index_from_string as ci, get_column_letter as L

A, B, TSV = sys.argv[1:4]
FARM = '"FARMACIA VF"'
DIN = {'Alessandra Dayan VF', 'TAB DIN. I.R.T.T.', 'I.R.T. T. resumo', '"FARMACIA VF RESUMO"', '"FARMACIA VF RESUMO" port A e B'}
esp = C.defaultdict(dict)
for l in open(TSV, encoding='utf-8-sig'):
    a, c, r, t, v = l.rstrip('\r\n').split('\t'); esp[a][(int(r), ci(c))] = (t, v)
wa = openpyxl.load_workbook(A, read_only=True); wb = openpyxl.load_workbook(B, read_only=True)
assert wa.sheetnames == wb.sheetnames
desloca = lambda c: c - 2 if c > 10 else c      # FARMACIA VF: I e J (9, 10) apagadas
tot = C.Counter()
for n in wa.sheetnames:
    ra = {(i, j): v for i, row in enumerate(wa[n].iter_rows(values_only=True), 1) for j, v in enumerate(row, 1) if v is not None}
    rb = {(i, j): v for i, row in enumerate(wb[n].iter_rows(values_only=True), 1) for j, v in enumerate(row, 1) if v is not None}
    e = esp.get(n, {})
    dif = gravadas_erradas = 0
    if n == FARM:
        ra = {(i, desloca(j)): v for (i, j), v in ra.items() if j not in (9, 10)}
        e = {(i, desloca(j)): tv for (i, j), tv in e.items()}
    for k in set(ra) | set(rb):
        if k in e:
            t, v = e[k]; x = rb.get(k)
            if t == 'd': ok = (x is None and not v) or (isinstance(x, dt.datetime) and x.date().isoformat() == v)
            else: ok = (x is None and not v) or str(x) == v
            if not ok:
                gravadas_erradas += 1
                if gravadas_erradas <= 3: print('  GRAVADA ERRADA', n, L(k[1]) + str(k[0]), repr(x), repr(v))
        elif n in DIN:
            continue
        elif ra.get(k) != rb.get(k):
            dif += 1
            if dif <= 3: print('  DIF', n, L(k[1]) + str(k[0]), repr(ra.get(k))[:60], '->', repr(rb.get(k))[:60])
    if dif or gravadas_erradas or e: print(f'{n[:40]:40} gravadas={len(e):7} erradas={gravadas_erradas} outras_diferentes={dif}')
    tot['dif'] += dif; tot['erradas'] += gravadas_erradas
za, zb = zipfile.ZipFile(A), zipfile.ZipFile(B)
for pat in [r'xl/pivotTables/pivotTable\d+\.xml$', r'xl/pivotCache/pivotCacheDefinition\d+\.xml$', r'xl/charts/.*\.xml$']:
    print(pat, len([x for x in za.namelist() if re.match(pat, x)]), len([x for x in zb.namelist() if re.match(pat, x)]))
dn = lambda z: sorted(re.findall(r'<definedName [^>]*>[^<]*</definedName>', z.read('xl/workbook.xml').decode('utf8')))
na, nb = dn(za), dn(zb)
print('nomes só na partida:', [x for x in na if x not in nb]); print('nomes só na cópia:', [x for x in nb if x not in na])
print('vbaProject:', 'xl/vbaProject.bin' in zb.namelist())
print('TOTAL', dict(tot))
