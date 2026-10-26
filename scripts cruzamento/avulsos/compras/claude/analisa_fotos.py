"""Indicadores das fotos diarias (saida de carrega_fotos.py).

    python analisa_fotos.py fotos.pkl SAIDA.pkl

- por dia: SKUs em ruptura, SKUs com Comprar > 0 e R$, valor de estoque;
- pedidos de compra (codigo + PEDIDO): primeiro e ultimo dia visto, data de entrega prometida (primeira e ultima),
  quantas vezes a data mudou, data de chegada (ULT_COMPRA quando o pedido some), atraso contra a 1a promessa;
- episodios de ruptura por SKU: inicio, fim, dias, se havia pedido em aberto no inicio.
"""
import datetime as dt
import pickle
import statistics as st
import sys


def num(x):
    try:
        return float(str(x).replace(',', '.')) if isinstance(x, str) else float(x or 0)
    except ValueError:
        return 0.0


def data(x):
    if isinstance(x, dt.datetime):
        return x.date()
    if isinstance(x, dt.date):
        return x
    s = str(x or '').strip()
    try:
        d = dt.datetime.strptime(s[:10], '%d/%m/%Y').date()
        return None if d.year < 2000 else d
    except ValueError:
        return None


def ped(r):
    return str(r.get('ped') or '').strip()


def main(ent, sai):
    fotos = pickle.load(open(ent, 'rb'))
    dias = [d for d, _ in fotos]
    diario = []
    for d, l in fotos:
        rup = [c for c, r in l.items() if num(r.get('rup')) < 0]
        comp = [c for c, r in l.items() if num(r.get('comprar')) > 0]
        diario.append({'dia': d, 'skus': len(l), 'ruptura': len(rup), 'comprar_n': len(comp),
                       'comprar_rs': sum(num(l[c]['comprar']) * num(l[c]['preco']) for c in comp),
                       'estoque_rs': sum(num(r.get('vlr')) for r in l.values()),
                       'pedidos': sum(1 for r in l.values() if ped(r))})

    # pedidos de compra
    po = {}
    for i, (d, l) in enumerate(fotos):
        for c, r in l.items():
            p = ped(r)
            if not p:
                continue
            e = po.setdefault((c, p), {'cod': c, 'ped': p, 'primeiro': d, 'i0': i, 'promessas': [],
                                       'qtd': num(r.get('compr')), 'comprador': str(r.get('comprador') or '').strip(),
                                       'preco': num(r.get('prc')) or num(r.get('preco'))})
            e['ultimo'] = d
            e['i1'] = i
            pr = data(r.get('dtent'))
            if pr and (not e['promessas'] or e['promessas'][-1] != pr):
                e['promessas'].append(pr)
    for e in po.values():
        e['censurado_inicio'] = e['i0'] == 0
        e['aberto_no_fim'] = e['i1'] == len(fotos) - 1
        e['mudou_data'] = max(0, len(e['promessas']) - 1)
        e['chegada'] = None
        if not e['aberto_no_fim']:
            prox = fotos[e['i1'] + 1][1].get(e['cod'])
            uc = data(prox.get('ultc')) if prox else None
            if uc and e['primeiro'] <= uc <= fotos[e['i1'] + 1][0]:
                e['chegada'] = uc
        e['lt_real'] = (e['chegada'] - e['primeiro']).days if e['chegada'] and not e['censurado_inicio'] else None
        e['lt_prometido'] = (e['promessas'][0] - e['primeiro']).days if e['promessas'] and not e['censurado_inicio'] else None
        e['atraso'] = (e['chegada'] - e['promessas'][0]).days if e['chegada'] and e['promessas'] else None

    # episodios de ruptura
    epis, ativo = [], {}
    for i, (d, l) in enumerate(fotos):
        em = {c for c, r in l.items() if num(r.get('rup')) < 0}
        for c in em - set(ativo):
            r = l[c]
            pedindo = 0
            for j in range(i - 1, -1, -1):
                rr = fotos[j][1].get(c)
                if rr and num(rr.get('comprar')) > 0 and not ped(rr):
                    pedindo += 1
                else:
                    break
            ativo[c] = {'cod': c, 'inicio': d, 'i0': i, 'comprador': str(r.get('comprador') or '').strip(),
                        'tinha_pedido': bool(ped(r)), 'rup_ini': num(r.get('rup')), 'pp': num(r.get('pp')),
                        'med': num(r.get('med')), 'preco': num(r.get('preco')), 'pedindo_antes': pedindo}
        for c in set(ativo) - em:
            e = ativo.pop(c)
            e['fim'] = d
            e['dias'] = (d - e['inicio']).days
            e['censurado'] = e['i0'] == 0
            epis.append(e)
    for e in ativo.values():
        e['fim'] = None
        e['dias'] = (dias[-1] - e['inicio']).days
        e['censurado'] = True
        epis.append(e)
    pickle.dump({'diario': diario, 'po': po, 'epis': epis, 'dias': dias}, open(sai, 'wb'))

    print('dias:', len(dias), dias[0], '->', dias[-1])
    print('rupturas por dia: min', min(x['ruptura'] for x in diario), 'max', max(x['ruptura'] for x in diario),
          'media', round(st.mean(x['ruptura'] for x in diario), 1))
    for x in diario[::6] + [diario[-1]]:
        print(f"  {x['dia']} SKUs {x['skus']} ruptura {x['ruptura']} comprar>0 {x['comprar_n']} "
              f"R$ {x['comprar_rs']:,.0f} estoque R$ {x['estoque_rs']:,.0f} pedidos {x['pedidos']}")
    ok = [e for e in po.values() if e['lt_real'] is not None]
    print('pedidos (cod+PC):', len(po), '| com chegada e emissao dentro da janela:', len(ok))
    if ok:
        lt = sorted(e['lt_real'] for e in ok)
        at = [e['atraso'] for e in ok if e['atraso'] is not None]
        print('  lead time real (dias): mediana', st.median(lt), 'p25', lt[len(lt) // 4], 'p75', lt[3 * len(lt) // 4],
              'p90', lt[9 * len(lt) // 10])
        print('  atraso contra a 1a promessa: mediana', st.median(at), '| no prazo ou antes:', sum(a <= 0 for a in at), 'de', len(at))
    mud = [e['mudou_data'] for e in po.values() if e['promessas']]
    print('  pedidos com entrega remarcada:', sum(m > 0 for m in mud), 'de', len(mud), '| 3+ remarcacoes:', sum(m >= 3 for m in mud))
    fe = [e for e in epis if not e['censurado']]
    print('episodios de ruptura:', len(epis), '| completos:', len(fe))
    if fe:
        print('  duracao (dias corridos): mediana', st.median(e['dias'] for e in fe), 'media', round(st.mean(e['dias'] for e in fe), 1))
    print('  sem pedido em aberto no inicio da ruptura:', sum(not e['tinha_pedido'] for e in epis), 'de', len(epis))
    print('  ja pediam compra (Comprar > 0, sem pedido) 3+ fotos antes:', sum(e['pedindo_antes'] >= 3 for e in epis))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
