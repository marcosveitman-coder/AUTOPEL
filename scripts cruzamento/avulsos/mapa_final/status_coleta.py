"""STATUS RECEITA / PORTE / STATUS AUTOPEL (05/11/2026) - passo 1: coleta as linhas que recebem as 3 colunas.
Abas: as 7 abas FINAL da semana (ALIMENTACAO, REDE DE LOJAS, CONSTRUCAO, POSTOS, PET SHOP, ESTACIONAMENTO, OUTROS) e as
550 empresas da aba Alessandra Dayan (11 tabelas lado a lado, ligadas as linhas das outras abas por
dados/alessandra_cruzamento.json e pela coluna ORIGEM da OUTROS FINAL).
Para as empresas da Alessandra que estao na FARMACIA VF, no IRTT por bandeira ou na SUPERMERCADOS (reuniao), le o CNPJ, a
situacao, o porte, a data da ultima compra e o faturamento dessas abas.
Saidas (pasta dados): status_itens.pkl e q19.json (CNPJs para busca_receita.py).
Uso: python status_coleta.py "<MAPA MERCADO.xlsx>" """
import os, sys, re, json, pickle, datetime as dt
import openpyxl
from openpyxl.utils import get_column_letter as L, column_index_from_string as ci
DADOS = os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dados')
D = lambda x: re.sub(r'\D', '', str(x or ''))
s = lambda x: str(x).strip() if x is not None else ''
ABAS = ['ALIMENTACAO FINAL', 'REDE DE LOJAS FINAL', 'CONSTRUÇÃO FINAL', 'POSTOS GAS & CONVENIENCIA FINAL', 'PET SHOP  FINAL',
        'ESTACIONAMENTO FINAL', 'OUTROS FINAL']
CAMPOS = {'nome': 'NOME FANTASIA', 'raz': 'RAZÃO SOCIAL', 'grupo': 'GRUPO ADMINISTRADOR', 'lojas': 'Nº DE LOJAS', 'cnpj': 'CNPJ',
          'conf': 'CONFERÊNCIA DO CNPJ', 'ref': 'CNPJ DE REFERÊNCIA (UNIDADE DA REDE)', 'uf': 'UF SEDE', 'sit': 'SITUAÇÃO NA RECEITA',
          'segm': 'SEGMENTO', 'origem': 'ORIGEM', 'obs': 'OBSERVAÇÃO', 'U': 'STATUS RECEITA', 'V': 'PORTE', 'W': 'STATUS AUTOPEL',
          'sub': 'SUBSEGMENTO', 'segp': 'SEGMENTO (PESQUISA)', 'rede': 'LOJAS ATIVAS COM O NOME NA RECEITA (TODOS OS CNPJs)'}


def dig(c):
    def k(b, w):
        r = sum(int(x) * y for x, y in zip(b, w)) % 11
        return '0' if r < 2 else str(11 - r)
    w1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    return len(c) == 14 and k(c[:12], w1) == c[12] and k(c[:13], [6] + w1) == c[13]


def main(mapa):
    wb = openpyxl.load_workbook(mapa, read_only=True, data_only=True)
    itens, cnpjs = [], set()
    for aba in ABAS:
        rows = list(wb[aba].iter_rows(values_only=True))
        h = [s(x) for x in rows[0]]
        col = {}
        for k, nome in CAMPOS.items():
            j = [i for i, x in enumerate(h) if x == nome] or ([i for i, x in enumerate(h) if x.startswith(nome)] if k == 'lojas' else [])
            if j: col[k] = j[0]
        print(aba, len(rows) - 1, 'linhas | sem coluna:', [k for k in CAMPOS if k not in col], '| colunas:', len(h))
        for n, r in enumerate(rows[1:], start=2):
            if not any(v not in (None, '') for v in r[:6]): continue
            it = dict(aba=aba, lin=n, **{k: r[j] if j < len(r) else None for k, j in col.items()})
            it['cnpj'] = D(it.get('cnpj')); it['ref'] = D(it.get('ref'))
            for c in (it['cnpj'], it['ref']):
                if len(c) == 14: cnpjs.add(c)
            itens.append(it)
    # aba Alessandra Dayan: 11 tabelas (EMPRESA, UF, LOJAS, BASE) com 3 colunas vazias depois de cada uma
    ws = wb['Alessandra Dayan']
    A = list(ws.iter_rows(values_only=True))
    tabelas = [(j, s(v)) for j, v in enumerate(A[0]) if v not in (None, '')]
    gaps = {}
    for j, seg in tabelas:
        assert s(A[1][j]) == 'EMPRESA' and s(A[1][j + 3]) == 'BASE', (seg, A[1][j:j + 4])
        livres = 0  # colunas vazias depois da tabela (a ultima tabela tem espaco livre a direita)
        while livres < 3 and all(A[i][j + 4 + livres] in (None, '') if j + 4 + livres < len(A[i]) else True for i in range(len(A))):
            livres += 1
        gaps[seg] = (j + 1, livres)
    ALE = {o['ref']: o for o in json.load(open(os.path.join(DADOS, 'alessandra_cruzamento.json'), encoding='utf8'))}
    por_origem = {}
    for it in itens:
        for m in re.findall(r'Alessandra Dayan: ([A-Z]+\d+)', s(it.get('origem'))):
            por_origem.setdefault(m, (it['aba'], it['lin']))
    # abas de apoio (FARMACIA VF, IRTT por bandeira, SUPERMERCADOS (reuniao))
    VF = {n: r for n, r in enumerate(wb['"FARMACIA VF"'].iter_rows(values_only=True), start=1) if n >= 7}
    IR = {n: r for n, r in enumerate(wb['I.R.T.T.__Por bandeira (FINAL)'].iter_rows(values_only=True), start=1) if n >= 5}
    RE = {n: r for n, r in enumerate(wb['SUPERMERCADOS (reuniao)'].iter_rows(values_only=True), start=1) if n >= 5}
    g = lambda r, c: r[ci(c) - 1] if r and len(r) >= ci(c) else None
    ale = []
    for j, seg in tabelas:
        for i in range(2, len(A)):
            nome = A[i][j] if j < len(A[i]) else None
            if nome in (None, ''): continue
            ref = f'{L(j + 1)}{i + 1}'
            o = ALE.get(ref, {})
            a = dict(ref=ref, lin=i + 1, col=j + 1, seg=seg, nome=s(nome), uf=s(A[i][j + 1]), lojas=A[i][j + 2], liga=None, apoio=[])
            if ref in por_origem:
                a['liga'] = por_origem[ref]
            for w in o.get('onde', []):
                m = re.match(r'^(.*?) linha (\d+) ', w)
                if not m: continue
                aba, n = m.group(1), int(m.group(2))
                if aba in ABAS and not a['liga']:
                    a['liga'] = (aba, n)
                elif aba == '"FARMACIA VF"' and n in VF:
                    r = VF[n]
                    a['apoio'].append(dict(fonte='FARMACIA VF', lin=n, cnpj=D(g(r, 'G')), sit=s(g(r, 'M')), data=g(r, 'O'), status=s(g(r, 'P')),
                                           porte=s(g(r, 'U')), fat=g(r, 'S'), nome=s(g(r, 'B'))))
                elif aba == 'I.R.T.T.__Por bandeira (FINAL)' and n in IR:
                    r = IR[n]
                    a['apoio'].append(dict(fonte='IRTT', lin=n, cnpj=D(g(r, 'Y')), sit='', data=g(r, 'G'), status=s(g(r, 'E')), porte='',
                                           fat=g(r, 'AA'), fat_fonte=s(g(r, 'AB')), nome=s(g(r, 'B')), uf=s(g(r, 'AH'))))
                elif aba == 'SUPERMERCADOS (reuniao)' and n in RE:
                    r = RE[n]
                    a['apoio'].append(dict(fonte='SUPERMERCADOS (reuniao)', lin=n, cnpj=D(g(r, 'F')), sit=s(g(r, 'X')), data=g(r, 'AK'),
                                           status=s(g(r, 'AL')), porte=s(g(r, 'I')), fat=g(r, 'K'), nome=s(g(r, 'D'))))
            for x in a['apoio']:
                if len(x['cnpj']) == 14: cnpjs.add(x['cnpj'])
            ale.append(a)
    wb.close()
    # apoio do IRTT para as linhas das abas FINAL: faturamento da bandeira (AA) por raiz do CNPJ validado (Y)
    irtt_raiz = {}
    for n, r in IR.items():
        y = D(g(r, 'Y'))
        if len(y) == 14 and isinstance(g(r, 'AA'), (int, float)) and g(r, 'AA') > 0:
            irtt_raiz.setdefault(y[:8], []).append(dict(lin=n, nome=s(g(r, 'B')), grupo=s(g(r, 'A')), fat=g(r, 'AA'), fonte=s(g(r, 'AB')),
                                                       lojas=g(r, 'AD'), seg=s(g(r, 'C'))))
    irtt_nome = [dict(lin=n, nome=s(g(r, 'B')), grupo=s(g(r, 'A')), fat=g(r, 'AA'), fonte=s(g(r, 'AB')), lojas=g(r, 'AD'), seg=s(g(r, 'C')),
                      cnpj=D(g(r, 'Y'))) for n, r in IR.items() if g(r, 'B')]
    sem = [a['ref'] for a in ale if not a['liga'] and not a['apoio']]
    print('Alessandra:', len(ale), 'empresas |', sum(1 for a in ale if a['liga']), 'ligadas a uma aba FINAL |',
          sum(1 for a in ale if not a['liga'] and a['apoio']), 'só nas abas de apoio |', len(sem), 'sem ligação:', sem)
    bad = sorted(c for c in cnpjs if not dig(c))
    print('CNPJs:', len(cnpjs), '| dígito verificador errado:', bad)
    print('Alessandra: colunas livres depois de cada tabela', gaps)
    pickle.dump(dict(itens=itens, ale=ale, gaps=gaps, irtt_raiz=irtt_raiz, irtt_nome=irtt_nome, gerado=dt.datetime.now()),
                open(os.path.join(DADOS, 'status_itens.pkl'), 'wb'))
    json.dump([dict(id='CNPJS', terms=[], cnpjs=sorted(cnpjs))], open(os.path.join(DADOS, 'q19.json'), 'w', encoding='utf8'))
    print(len(itens), 'linhas das abas FINAL')


if __name__ == '__main__':
    main(sys.argv[1])
