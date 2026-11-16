"""Índice de clientes (16/11/2026): junta, por cliente+loja do TOTVS,
  - CADASTRO DE CLIENTES.RLT.xlsx: nome, fantasia, CNPJ, vendedor do cadastro (código), UF, e-mail, grupo e segmento;
  - ultimo fat.xlsx: grupo do cliente com a descrição, e-mail, segmento, vendedor (código, nome e nome reduzido) e DT_FAT
    (última nota fiscal do cliente+loja);
  - os pickles do extrai_fat.py: última nota de cada cliente+loja+produto nas bases de faturamento 2020-2026.
Saída: <pasta>/clientes.pkl com
  cli[(cod, loja)] = dict(cnpj, raiz, nome, fant, email, dom, uf, seg, grp, grp_desc, vend_cad, dt_ultfat, dt_fat,
                          vend_fat, prods={des_prod: (data, vendedor)})
  vend[codigo] = nome do vendedor (para traduzir o vendedor do cadastro)
Uso: python base_clientes.py "<pasta 12.11.2026>" "<pasta dos pickles>" """
import sys, os, re, glob, pickle, datetime as dt, collections as C
import openpyxl

D = lambda x: re.sub(r'\D', '', str(x or ''))
s = lambda x: str(x).strip() if x is not None else ''


def ler(arq, linha_cab):
    wb = openpyxl.load_workbook(arq, read_only=True)
    it = wb.worksheets[0].iter_rows(values_only=True)
    for _ in range(linha_cab - 1): next(it)
    cab = [s(x) for x in next(it)]
    for r in it:
        if r and any(v is not None for v in r): yield cab, r
    wb.close()


def data(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    t = s(v)
    for f in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y'):
        try: return dt.datetime.strptime(t, f).date()
        except ValueError: pass
    return None


def dominio(e):
    out = set()
    for m in re.findall(r'@([a-z0-9.\-]+)', s(e).lower()):
        out.add(m)
    return out


def main(pasta, pk):
    cli = C.defaultdict(lambda: dict(cnpj='', raiz='', nome='', fant='', email='', dom=set(), uf='', seg='', grp='', grp_desc='',
                                     vend_cad='', dt_ultfat=None, dt_fat=None, vend_fat='', prods={}, no_cadastro=False))
    vend = {}
    # 1) cadastro (linha 2 = cabeçalho)
    n = 0
    for cab, r in ler(os.path.join(pasta, 'CADASTRO DE CLIENTES.RLT.xlsx'), 2):
        g = lambda c: s(r[cab.index(c)])
        k = (g('Codigo'), g('Loja'))
        c = cli[k]; n += 1
        c.update(cnpj=D(g('CNPJ/CPF')), nome=g('Nome'), fant=g('N Fantasia'), vend_cad=g('Vendedor'), uf=g('Estado'),
                 email=g('E-Mail NF-e'), grp=g('Grp.Vendas'), no_cadastro=True)
        c['dom'] |= dominio(g('E-Mail NF-e'))
    print('cadastro:', n)
    # 2) ultimo fat (cabeçalho na linha 1; colunas repetidas: Descricao do grupo e do segmento, Nome do cliente e do vendedor)
    n = 0
    for cab, r in ler(os.path.join(pasta, 'ultimo fat.xlsx'), 1):
        x = [s(v) for v in r]
        k = (x[0], x[1]); c = cli[k]; n += 1
        if not c['cnpj']: c['cnpj'] = D(x[2])
        c['grp'] = c['grp'] or x[3]; c['grp_desc'] = x[4]
        if not c['nome']: c['nome'] = x[5]
        if not c['fant']: c['fant'] = x[6]
        if not c['email']: c['email'] = x[8]
        c['dom'] |= dominio(x[8])
        c['seg'] = x[13]
        if not c['uf']: c['uf'] = x[28] if len(x) > 28 else ''
        if not c['vend_cad']: c['vend_cad'] = x[14]
        if x[14]: vend.setdefault(x[14], x[15] or x[16])
        c['dt_ultfat'] = data(r[24])
    print('ultimo fat:', n)
    # 3) bases de faturamento
    for f in sorted(glob.glob(os.path.join(pk, '*.pkl'))):
        P = pickle.load(open(f, 'rb'))
        for (cod, loja, prod), v in P['dados'].items():
            d, cv, nv, val, cnpj, grp, nome, uf, mun, des, seg = v
            c = cli[(cod, loja)]
            if not c['cnpj']: c['cnpj'] = D(cnpj)
            if not c['nome']: c['nome'] = nome
            if not c['uf']: c['uf'] = uf
            if not c['seg']: c['seg'] = seg
            if not c['grp']: c['grp'] = grp
            if cv: vend.setdefault(cv, nv)
            if c['dt_fat'] is None or d > c['dt_fat']: c['dt_fat'] = d; c['vend_fat'] = nv
            a = c['prods'].get(des)
            if a is None or d > a[0]: c['prods'][des] = (d, nv)
        print(os.path.basename(f), len(P['dados']), flush=True)
    for c in cli.values():
        c['raiz'] = c['cnpj'][:8] if len(c['cnpj']) == 14 else ''
        c['ult'] = max([d for d in (c['dt_ultfat'], c['dt_fat']) if d], default=None)
    # descrição de grupo para quem só tem o código
    gd = {}
    for c in cli.values():
        if c['grp'] and c['grp_desc']: gd.setdefault(c['grp'], c['grp_desc'])
    for c in cli.values():
        if c['grp'] and not c['grp_desc']: c['grp_desc'] = gd.get(c['grp'], '')
    out = dict(cli={k: dict(v) for k, v in cli.items()}, vend=vend)
    pickle.dump(out, open(os.path.join(pk, '..', 'clientes.pkl'), 'wb'))
    ult = [c['ult'] for c in cli.values() if c['ult']]
    print('clientes+loja:', len(cli), '| com compra:', len(ult), '| compra mais recente:', max(ult), '| vendedores:', len(vend))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
