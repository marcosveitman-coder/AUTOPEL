"""Atualiza o clientes.pkl (base_clientes.py, 16/11/2026) com os CSVs do dia da pasta de relatórios TOTVS (19/11/2026):
  Clientes - SQL.csv            cadastro: clientes novos, nome, fantasia, CNPJ, vendedor, UF, e-mail, grupo
  Ultimo faturamento - SQL.csv  grupo + descrição, segmento, vendedor e DT_FAT (última nota de venda do cliente+loja)
  FAT - SQL.csv                 notas do mês: produto (DES_PROD) e data, só nota com valor > 0
  Pedidos abertos - SQL.csv     pedidos em aberto: c['ped'] = {desc_prod: data do pedido} (informação; não conta como compra)
Sem isso o status ficava com a foto de 16/11 (notas dos dias seguintes e clientes novos de fora).
Uso: python atualiza_clientes.py <clientes.pkl de partida> <pasta CSV do TOTVS> <clientes.pkl de saída>"""
import sys, os, re, csv, pickle, datetime as dt
from base_clientes import dominio, D

ENT, PASTA, SAI = sys.argv[1:4]
B = pickle.load(open(ENT, 'rb')); cli, vend = B['cli'], B['vend']
s = lambda x: (x or '').strip()


def linhas(nome):
    with open(os.path.join(PASTA, nome), encoding='utf-8-sig', newline='') as f:
        r = csv.reader(f, delimiter=';'); next(r); cab = [s(x) for x in next(r)]
        for x in r:
            if x and any(v.strip() for v in x): yield cab, [s(v) for v in x]


def data(t):
    t = s(t)[:10]
    try: return dt.datetime.strptime(t, '%d/%m/%Y').date()
    except ValueError: return None


def num(t):
    t = s(t).replace('.', '').replace(',', '.')
    try: return float(t)
    except ValueError: return 0.0


def novo():
    return dict(cnpj='', raiz='', nome='', fant='', email='', dom=set(), uf='', seg='', grp='', grp_desc='', vend_cad='',
                dt_ultfat=None, dt_fat=None, vend_fat='', prods={}, no_cadastro=False)


n_novo = n = 0
for cab, x in linhas('Clientes - SQL.csv'):
    g = lambda c: x[cab.index(c)] if cab.index(c) < len(x) else ''
    k = (g('Codigo'), g('Loja'))
    if k not in cli: cli[k] = novo(); n_novo += 1
    c = cli[k]; n += 1
    cnpj = D(g('CNPJ/CPF'))
    if cnpj: c['cnpj'] = cnpj
    for campo, col in (('nome', 'Nome'), ('fant', 'N Fantasia'), ('vend_cad', 'Vendedor'), ('uf', 'Estado'), ('grp', 'Grp.Vendas')):
        if g(col): c[campo] = g(col)
    if g('E-Mail NF-e'): c['email'] = g('E-Mail NF-e'); c['dom'] |= dominio(g('E-Mail NF-e'))
    c['no_cadastro'] = True
print('cadastro:', n, '| clientes novos:', n_novo)
n = n_dt = 0
for cab, x in linhas('Ultimo faturamento - SQL.csv'):
    x += [''] * (29 - len(x))
    k = (x[0], x[1])
    if k not in cli: cli[k] = novo()
    c = cli[k]; n += 1
    if not c['cnpj']: c['cnpj'] = D(x[2])
    if x[3]: c['grp'] = x[3]
    if x[4]: c['grp_desc'] = x[4]
    if not c['nome']: c['nome'] = x[5]
    if not c['fant']: c['fant'] = x[6]
    if x[8]: c['dom'] |= dominio(x[8]); c['email'] = c['email'] or x[8]
    if x[13]: c['seg'] = x[13]
    if not c['uf']: c['uf'] = x[28]
    if x[14]: vend[x[14]] = x[15] or x[16] or vend.get(x[14], '')
    d = data(x[24])
    if d and num(x[25]) > 0 and (c['dt_ultfat'] is None or d > c['dt_ultfat']): c['dt_ultfat'] = d; n_dt += 1
print('ultimo fat:', n, '| datas novas:', n_dt)
n = 0
for cab, x in linhas('FAT - SQL.csv'):
    g = lambda col: x[cab.index(col)] if cab.index(col) < len(x) else ''
    if num(g('VLR_TOTAL')) <= 0: continue
    k = (g('COD_CLI'), g('COD_LOJA')); d = data(g('EMISSAO'))
    if not d: continue
    if k not in cli: cli[k] = novo()
    c = cli[k]; n += 1
    if not c['cnpj']: c['cnpj'] = D(g('CNPJ'))
    if not c['nome']: c['nome'] = g('CLIENTE')
    if not c['seg']: c['seg'] = g('SEGMENTO')
    if c['dt_fat'] is None or d > c['dt_fat']: c['dt_fat'] = d; c['vend_fat'] = g('NOME_VENDEDOR')
    des = g('DES_PROD'); a = c['prods'].get(des)
    if a is None or d > a[0]: c['prods'][des] = (d, g('NOME_VENDEDOR'))
print('FAT do mês: notas com valor', n)
n = 0
for cab, x in linhas('Pedidos abertos - SQL.csv'):
    g = lambda col: x[cab.index(col)] if cab.index(col) < len(x) else ''
    k = (g('COD_CLI'), g('LOJA_CLI'))
    if k not in cli: cli[k] = novo(); cli[k].update(cnpj=D(g('CNPJ')), nome=g('CLIENTE'), fant=g('FANTASIA'), uf=g('ESTADO'))
    c = cli[k]; n += 1
    c.setdefault('ped', {})[g('DESC_PROD')] = data(g('DATA_PED'))
print('pedidos em aberto (itens):', n)
for c in cli.values():
    c['raiz'] = c['cnpj'][:8] if len(c['cnpj']) == 14 else ''
    c['ult'] = max([d for d in (c['dt_ultfat'], c['dt_fat']) if d], default=None)
    c.setdefault('ped', {})
gd = {}
for c in cli.values():
    if c['grp'] and c['grp_desc']: gd.setdefault(c['grp'], c['grp_desc'])
for c in cli.values():
    if c['grp'] and not c['grp_desc']: c['grp_desc'] = gd.get(c['grp'], '')
pickle.dump(dict(cli=cli, vend=vend), open(SAI, 'wb'))
print('clientes+loja:', len(cli), '| compra mais recente:', max(c['ult'] for c in cli.values() if c['ult']))
