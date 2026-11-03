"""Confere as 3 abas FINAL (2a rodada) e a aba MAIORES POR REGIAO: DV de todo CNPJ (e do CNPJ de referencia), CNPJ ativo na
Receita local (ou na consulta BrasilAPI), duplicidade de nome, e se o ranking por regiao bate com as abas."""
import os as _os
import sys as _sys
_sys.argv[1:] = [_os.path.abspath(a) for a in _sys.argv[1:]]  # caminhos da planilha, antes do chdir
_AQUI = _os.path.dirname(_os.path.abspath(__file__))
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_AQUI, 'dados'))
import openpyxl, sys, re, pickle, json, collections, unicodedata
p = sys.argv[1]
rows_by = {}
for pk in ['r1.pkl', 'r3.pkl', 'r4.pkl', 'r5.pkl']:
    for r in pickle.load(open(pk, 'rb'))['est']: rows_by[r[0] + r[1] + r[2]] = r[5]
for c, v in json.load(open('brasilapi.json', encoding='utf8')).items(): rows_by.setdefault(c, '02' if 'ATIVA' in str(v[3]).upper() else v[3])
def dv(c):
    def k(b, w):
        s = sum(int(x) * y for x, y in zip(b, w)); r = s % 11; return '0' if r < 2 else str(11 - r)
    w1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    return len(c) == 14 and k(c[:12], w1) == c[12] and k(c[:13], [6] + w1) == c[13]
def n(s): return ' '.join(unicodedata.normalize('NFKD', s or '').encode('ascii', 'ignore').decode().upper().split())
wb = openpyxl.load_workbook(p, read_only=True)
lojas = {}
for aba in ['ALIMENTACAO FINAL', 'REDE DE LOJAS FINAL', 'CONSTRUÇÃO FINAL']:
    vals = list(wb[aba].iter_rows(values_only=True)); head, data = vals[0], vals[1:]
    H = {h: i for i, h in enumerate(head)}
    erros = []; nomes = collections.Counter(); st = collections.Counter(); ncnpj = nref = 0
    for r in data:
        nome, cnpj, conf, ref = r[H['NOME FANTASIA']], re.sub(r'\D', '', r[H['CNPJ']] or ''), r[H['CONFERÊNCIA DO CNPJ']], re.sub(r'\D', '', r[H['CNPJ DE REFERÊNCIA (UNIDADE DA REDE)']] or '')
        nomes[n(nome)] += 1; st[conf] += 1; lojas[(aba, nome)] = r[H['Nº DE LOJAS']]
        for tipo, c in (('cnpj', cnpj), ('ref', ref)):
            if not c: continue
            if tipo == 'cnpj': ncnpj += 1
            else: nref += 1
            if not dv(c): erros.append(('DV', tipo, nome, c))
            s = rows_by.get(c)
            if s is None: erros.append(('sem dado Receita', tipo, nome, c))
            elif s != '02': erros.append(('não ativo', tipo, nome, c, s))
        if not cnpj and conf in ('CNPJ CONFIRMADO NA RECEITA', 'CNPJ PROVÁVEL (CONFERIR)'): erros.append(('sem CNPJ', nome, conf))
        if cnpj and ref: erros.append(('CNPJ e ref juntos', nome))
    dup = [k for k, v in nomes.items() if v > 1]
    print(f"{aba}: {len(data)} linhas | {ncnpj} CNPJ + {nref} ref | {dict(st)} | repetidos: {dup} | erros {len(erros)}: {erros[:6]}")
vals = list(wb['MAIORES POR REGIÃO'].iter_rows(values_only=True))
blocos = [v[0] for v in vals if v and isinstance(v[0], str) and ' — ' in v[0]]
print('MAIORES POR REGIÃO:', len(vals), 'linhas,', len(blocos), 'blocos')
wb.close()
