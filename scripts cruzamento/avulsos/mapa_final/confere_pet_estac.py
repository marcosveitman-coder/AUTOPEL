"""Confere a gravacao das abas PET SHOP  FINAL e ESTACIONAMENTO FINAL contra o arquivo de partida:
- todas as outras abas iguais celula a celula; MAIORES POR REGIAO e R_E_S_U_M_O iguais ate a ultima linha antiga;
- nas 2 abas novas: digito verificador de todo CNPJ, CNPJ ativo na Receita local (r12/r13/r14), nome sem repeticao;
- dinamicas, caches e nomes definidos iguais.
Uso: python confere_pet_estac.py <partida.xlsx> <copia.xlsx>"""
import os as _os
import sys as _sys
_sys.argv[1:] = [_os.path.abspath(a) for a in _sys.argv[1:]]
_AQUI = _os.path.dirname(_os.path.abspath(__file__))
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_AQUI, 'dados'))
import sys, re, zipfile, pickle, collections, unicodedata, openpyxl
A, B = sys.argv[1:3]
NOVAS = ['PET SHOP  FINAL', 'ESTACIONAMENTO FINAL']; ACRESC = ['MAIORES POR REGIÃO', 'R_E_S_U_M_O']
sit = {}
for pk in ['r12.pkl', 'r13.pkl', 'r14.pkl', 'r15.pkl', 'r16.pkl']:
    if _os.path.exists(pk):
        for x in pickle.load(open(pk, 'rb'))['est']: sit[x[0] + x[1] + x[2]] = x[5]
def dv(c):
    def k(b, w):
        s = sum(int(x) * y for x, y in zip(b, w)); r = s % 11; return '0' if r < 2 else str(11 - r)
    w1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    return len(c) == 14 and k(c[:12], w1) == c[12] and k(c[:13], [6] + w1) == c[13]
def n(s): return ' '.join(unicodedata.normalize('NFKD', s or '').encode('ascii', 'ignore').decode().upper().split())
wa = openpyxl.load_workbook(A, read_only=True); wb = openpyxl.load_workbook(B, read_only=True)
assert wa.sheetnames == wb.sheetnames
dif = 0
for nome in wa.sheetnames:
    if nome in NOVAS: continue
    ra = list(wa[nome].iter_rows(values_only=True)); rb = list(wb[nome].iter_rows(values_only=True))
    lim = len(ra)
    if nome not in ACRESC and len(ra) != len(rb): print('nº de linhas mudou', nome, len(ra), len(rb)); dif += 1
    for i in range(lim):
        x, y = list(ra[i]), list(rb[i]); m = max(len(x), len(y)); x += [None] * (m - len(x)); y += [None] * (m - len(y))
        if x != y: dif += 1; print('DIF', nome, i + 1) if dif < 15 else None
print('outras abas: diferenças =', dif)
for nome in NOVAS:
    vals = list(wb[nome].iter_rows(values_only=True)); H = {h: i for i, h in enumerate(vals[0])}
    erros = []; nomes = collections.Counter(); conf = collections.Counter()
    for r in vals[1:]:
        nomes[n(r[H['NOME FANTASIA']])] += 1; conf[r[H['CONFERÊNCIA DO CNPJ']]] += 1
        for col in ['CNPJ', 'CNPJ DE REFERÊNCIA (UNIDADE DA REDE)']:
            c = re.sub(r'\D', '', r[H[col]] or '')
            if not c: continue
            if not dv(c): erros.append(('DV', r[1], c))
            elif sit.get(c) != '02': erros.append(('não ativo/sem dado', r[1], c, sit.get(c)))
        if r[H['CONFERÊNCIA DO CNPJ']] in ('CNPJ CONFIRMADO NA RECEITA', 'CNPJ PROVÁVEL (CONFERIR)') and not r[H['CNPJ']]: erros.append(('sem CNPJ', r[1]))
    print(nome, len(vals) - 1, 'linhas |', dict(conf), '| nomes repetidos:', [k for k, v in nomes.items() if v > 1], '| erros:', erros[:10])
za, zb = zipfile.ZipFile(A), zipfile.ZipFile(B)
for pat in [r'xl/pivotTables/.*\.xml$', r'xl/pivotCache/pivotCacheDefinition.*\.xml$']:
    print(pat, len([x for x in za.namelist() if re.match(pat, x)]), len([x for x in zb.namelist() if re.match(pat, x)]))
dn = lambda z: re.findall(r'<definedName [^>]*>[^<]*</definedName>', z.read('xl/workbook.xml').decode('utf8'))
print('nomes definidos iguais:', dn(za) == dn(zb), '| partes diferentes:', [x for x in zb.namelist() if za.read(x) != zb.read(x)])
