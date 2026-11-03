"""Confere as 3 abas gravadas: DV de todo CNPJ, CNPJ ativo na Receita local, duplicidade de nome e de CNPJ+nome por aba."""
import openpyxl, sys, re, pickle, collections, unicodedata
import os as _os
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'dados'))
p = sys.argv[1]
rows_by = {}
for pk in ['r1.pkl', 'r3.pkl']:
    for r in pickle.load(open(pk, 'rb'))['est']: rows_by[r[0] + r[1] + r[2]] = r
def dv(c):
    def k(b, w):
        s = sum(int(x) * y for x, y in zip(b, w)); r = s % 11; return '0' if r < 2 else str(11 - r)
    w1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    return len(c) == 14 and k(c[:12], w1) == c[12] and k(c[:13], [6] + w1) == c[13]
def n(s): return ' '.join(unicodedata.normalize('NFKD', s or '').encode('ascii', 'ignore').decode().upper().split())
wb = openpyxl.load_workbook(p, read_only=True)
for aba in [x for x in wb.sheetnames if x.endswith('FINAL') and x.startswith(('ALIMENTACAO', 'REDE DE LOJAS', 'CONSTRU'))]:
    ws = wb[aba]; vals = list(ws.iter_rows(values_only=True))
    head, data = vals[0], vals[1:]
    erros = []; nomes = collections.Counter(); st = collections.Counter(); ncnpj = 0
    for r in data:
        nome, cnpj, conf = r[1], re.sub(r'\D', '', r[5] or ''), r[6]
        nomes[n(nome)] += 1; st[conf] += 1
        if cnpj:
            ncnpj += 1
            if not dv(cnpj): erros.append(('DV', nome, cnpj))
            e = rows_by.get(cnpj)
            if not e: erros.append(('sem dado Receita', nome, cnpj))
            elif e[5] != '02': erros.append(('não ativo', nome, cnpj, e[5]))
        elif conf == 'CNPJ CONFIRMADO NA RECEITA': erros.append(('confirmado sem CNPJ', nome))
    dup = [k for k, v in nomes.items() if v > 1]
    print(f"{aba}: {len(data)} linhas | {ncnpj} com CNPJ | {dict(st)} | nomes repetidos: {dup} | erros: {erros[:5]} ({len(erros)})")
wb.close()
