"""Confere a copia gravada contra o arquivo de partida: todas as abas (menos a gravada) celula a celula (valor ou
formula), dinamicas, caches e nomes definidos. Uso: python org_confere.py <partida.xlsx> <copia.xlsx> "<aba gravada>" """
import sys, zipfile, re, openpyxl
A, B, ABA = sys.argv[1:4]
wa = openpyxl.load_workbook(A, read_only=True); wb = openpyxl.load_workbook(B, read_only=True)
assert wa.sheetnames == wb.sheetnames, 'ordem/nome das abas mudou'
dif = 0
for n in wa.sheetnames:
    if n == ABA: continue
    for i, (ra, rb) in enumerate(zip(wa[n].iter_rows(values_only=True), wb[n].iter_rows(values_only=True))):
        if tuple(x for x in ra) != tuple(x for x in rb):
            ra2 = list(ra) + [None] * max(0, len(rb) - len(ra)); rb2 = list(rb) + [None] * max(0, len(ra) - len(rb))
            if ra2 != rb2: dif += 1; print('DIF', n, i + 1) if dif < 20 else None
    print('ok' if not dif else 'dif', n, flush=True)
print('diferencas', dif)
za, zb = zipfile.ZipFile(A), zipfile.ZipFile(B)
for pat in [r'xl/pivotTables/.*\.xml$', r'xl/pivotCache/pivotCacheDefinition.*\.xml$', r'xl/charts/.*\.xml$', r'xl/slicers?/.*\.xml$']:
    print(pat, len([x for x in za.namelist() if re.match(pat, x)]), len([x for x in zb.namelist() if re.match(pat, x)]))
dn = lambda z: sorted(re.findall(r'<definedName [^>]*>[^<]*</definedName>', z.read('xl/workbook.xml').decode('utf8')))
na, nb = dn(za), dn(zb)
print('nomes so na partida:', [x for x in na if x not in nb]); print('nomes so na copia:', [x for x in nb if x not in na])
