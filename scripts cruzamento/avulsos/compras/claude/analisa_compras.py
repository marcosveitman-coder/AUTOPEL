"""Historico de entradas (TOTvS_HistoricodeCompras) e ordens de compra (TOTvS_OrdemCompra): lote, frequencia e prazo.

    python analisa_compras.py HISTORICO.csv ORDEMCOMPRA.csv SAIDA.pkl

- por produto (armazem 40, ultimos 12 meses ate a data mais recente do arquivo): entradas, quantidade total,
  quantidade tipica (mediana), maior divisor comum das quantidades (multiplo de compra), fornecedor principal,
  data da ultima entrada;
- por fornecedor (ordens de compra): prazo prometido = Entrega - Emissao (mediana), pedidos em aberto, atrasados.
"""
import collections
import csv
import datetime as dt
import math
import pickle
import statistics as st
import sys
from functools import reduce


def num(s):
    s = str(s or '').strip()
    if s in ('', '-'):
        return 0.0
    return float(s.replace('.', '').replace(',', '.'))


def data(s):
    try:
        return dt.datetime.strptime(str(s).strip(), '%d/%m/%Y').date()
    except ValueError:
        return None


def le(arq):
    with open(arq, encoding='latin-1', newline='') as f:
        r = csv.reader(f, delimiter=';')
        cab = [c.strip() for c in next(r)]
        for linha in r:
            yield {k: v.strip() for k, v in zip(cab, linha)}


def main(hist, oc, saida):
    ent = list(le(hist))
    fim = max(d for d in (data(e['DT Digitacao']) for e in ent) if d)
    ini = fim - dt.timedelta(days=365)
    prod = collections.defaultdict(list)
    for e in ent:
        d = data(e['DT Digitacao'])
        if not d or e.get('Armazem') != '40' or e.get('Tipo') != 'RV':
            continue
        prod[e['Produto']].append((d, num(e['Quantidade']), num(e['Vlr.Unitario']), e['Razao Social'], e['Forn/Cliente'], e['Unidade']))
    por_prod = {}
    for cod, l in prod.items():
        l.sort()
        ult = [x for x in l if x[0] >= ini]
        qs = [x[1] for x in ult if x[1] > 0]
        forn = collections.Counter()
        for x in ult or l[-5:]:
            forn[x[3]] += x[1] * x[2]
        mdc = reduce(math.gcd, [int(q) for q in qs if float(q).is_integer()], 0) if qs else 0
        datas = sorted({x[0] for x in ult})
        intervalos = [(b - a).days for a, b in zip(datas, datas[1:])]
        por_prod[cod] = {'entradas12': len(datas), 'qtd12': sum(qs), 'qtd_tipica': st.median(qs) if qs else None,
                         'mdc': mdc, 'forn_principal': forn.most_common(1)[0][0] if forn else '',
                         'n_forn12': len({x[3] for x in ult}), 'ultima': l[-1][0], 'ult_preco': l[-1][2],
                         'intervalo_mediano': st.median(intervalos) if intervalos else None, 'unidade': l[-1][5]}
    # ordens de compra
    ocs = list(le(oc))
    forn = collections.defaultdict(list)
    for o in ocs:
        e, p = data(o['Emissao']), data(o['Entrega'])
        if not e or not p or not o['Produto'].startswith('V'):
            continue
        forn[o['Razao Social']].append({'pc': o['Num.PC'].lstrip('0'), 'prod': o['Produto'], 'emissao': e, 'entrega': p,
                                         'prazo': (p - e).days, 'saldo': num(o['Saldo Receber']), 'qtd': num(o['Quantidade'])})
    por_forn = {}
    for f, l in forn.items():
        pz = [x['prazo'] for x in l if 0 <= x['prazo'] <= 180]
        por_forn[f] = {'itens': len(l), 'prazo_prometido_mediano': st.median(pz) if pz else None, 'pcs': len({x['pc'] for x in l})}
    pc_forn = {x['pc']: (f, x['emissao'], x['entrega']) for f, l in forn.items() for x in l}
    pickle.dump({'prod': por_prod, 'forn': por_forn, 'pc_forn': pc_forn, 'fim_hist': fim,
                 'oc_data_max': max((x['emissao'] for l in forn.values() for x in l), default=None)}, open(saida, 'wb'))
    print('historico ate', fim, '| produtos RV/40:', len(por_prod), '| com entrada em 12 meses:',
          sum(1 for v in por_prod.values() if v['entradas12']))
    e12 = [v['entradas12'] for v in por_prod.values() if v['entradas12']]
    print('  entradas por produto em 12 meses: mediana', st.median(e12), '| 12+ (mensal ou mais):', sum(e >= 12 for e in e12))
    iv = [v['intervalo_mediano'] for v in por_prod.values() if v['intervalo_mediano']]
    print('  intervalo mediano entre entradas (dias): mediana', st.median(iv))
    print('  produtos com multiplo de compra > 1:', sum(1 for v in por_prod.values() if v['mdc'] > 1))
    print('ordens de compra (RV):', sum(len(l) for l in forn.values()), 'itens,', len(forn), 'fornecedores')
    pz = [v['prazo_prometido_mediano'] for v in por_forn.values() if v['prazo_prometido_mediano'] is not None]
    print('  prazo prometido por fornecedor (mediana das medianas):', st.median(pz))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], sys.argv[3])
