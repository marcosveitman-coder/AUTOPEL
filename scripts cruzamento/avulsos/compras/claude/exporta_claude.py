"""Gera os arquivos que o grava_claude.ps1 escreve na planilha: parametros por SKU e tabelas da analise.

    python exporta_claude.py PASTA_PKL RELATORIO.xlsx SAIDA_DIR [--hoje 2026-10-26]

Saidas em SAIDA_DIR:
- param_claude.csv  (separado por ';', UTF-8): uma linha por SKU, colunas da tabela ParamClaude;
- analise.json      : blocos de texto e tabelas da aba "analise compras CLAUDE" (numeros do diagnostico);
- referencia.pkl    : resultado da metodologia calculado em Python, para conferir as formulas da planilha.
"""
import collections
import csv
import datetime as dt
import json
import pickle
import statistics as st
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import calcula_claude as cc  # noqa: E402


def num(x):
    try:
        return float(x or 0)
    except (TypeError, ValueError):
        return 0.0


def main(pasta, rel, saida, hoje):
    P = Path(pasta)
    S = Path(saida)
    S.mkdir(parents=True, exist_ok=True)
    par = pickle.load(open(P / 'params.pkl', 'rb'))
    prm = par['params']
    fotos = pickle.load(open(P / 'fotos.pkl', 'rb'))
    ind = pickle.load(open(P / 'fotos_ind.pkl', 'rb'))
    po_forn = pickle.load(open(P / 'po_forn.pkl', 'rb'))
    tv = pickle.load(open(P / 'demanda_totvs.pkl', 'rb'))
    bol = pickle.load(open(P / 'boletins.pkl', 'rb'))

    # 1) parametros
    cab = ['Codigo', 'Classe', 'ABC', 'XYZ', 'PQR', 'CV usado', 'LT (dias)', 'Desvio LT (dias)', 'Multiplo'] + \
          [f'S{i:02d}' for i in range(1, 13)] + ['CV medido', 'Meses com venda (6)', 'Fonte sazonalidade', 'Fonte LT',
                                                 'Fornecedor principal', 'Faturamento cobre a venda']
    with open(S / 'param_claude.csv', 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f, delimiter=';')
        w.writerow(cab)
        for c, p in sorted(prm.items()):
            w.writerow([c, p['classe'], p['abc'], p['xyz'], p['freq'], p['cv_usado'], p['lt'], p['slt'], p['mult']] +
                       p['S'] + ['' if p['cv'] is None else p['cv'], p['meses_venda6'], p['fonte_saz'], p['fonte_lt'],
                                 p['fornecedor'], 'sim' if p['coberto'] else 'nao'])

    # 2) referencia da metodologia
    linhas = cc.le_geral(rel)
    for l in linhas:
        l['claude'] = cc.calcula(l, prm.get(l['cod']), hoje)
        l['param'] = prm.get(l['cod'])
    pickle.dump(linhas, open(S / 'referencia.pkl', 'wb'))

    blocos = {}
    # serie diaria
    blocos['serie'] = [['Dia', 'SKUs em ruptura', 'SKUs com Comprar > 0', 'R$ a comprar (TOTVS)', 'Estoque R$', 'Pedidos de compra em aberto']] + \
        [[x['dia'].strftime('%d/%m/%Y'), x['ruptura'], x['comprar_n'], round(x['comprar_rs']), round(x['estoque_rs']), x['pedidos']] for x in ind['diario']]

    # causas de ruptura
    idx = {d: i for i, (d, _) in enumerate(fotos)}
    import analisa_fotos as af  # noqa
    causa, rs = collections.Counter(), collections.Counter()
    porc = collections.defaultdict(lambda: collections.Counter())
    dur = collections.defaultdict(list)
    for e in ind['epis']:
        i = e['i0']
        r = fotos[i][1][e['cod']]
        ant = fotos[i - 1][1].get(e['cod']) if i > 0 else None
        pp, med, vp = num(r.get('pp')), num(r.get('med')), num(r.get('vprot'))
        vp_ant = num(ant.get('vprot')) if ant else 0
        dent = af.data(r.get('dtent'))
        ped = str(r.get('ped') or '').strip()
        if e['censurado'] and e['i0'] == 0:
            k = 'Já em ruptura no primeiro dia (sem origem)'
        elif pp == 0:
            k = 'PontoPedido 0: item sem estoque por regra (venda sob encomenda)'
        elif vp - vp_ant > max(med, 1):
            k = 'Pedido de venda grande (mais de 1 mês de venda de uma vez)'
        elif ped and dent and dent < fotos[i][0]:
            k = 'Pedido de compra atrasado (data de entrega já passou)'
        elif ped:
            k = 'Pedido de compra a caminho: comprou tarde'
        elif e['pedindo_antes'] >= 3:
            k = 'Sem pedido de compra, com Comprar > 0 havia 3+ dias'
        else:
            k = 'Sem pedido de compra: estoque acabou sem aviso'
        causa[k] += 1
        rs[k] += abs(e['rup_ini']) * e['preco']
        b = e['comprador'] or '(sem comprador)'
        porc[b]['n'] += 1
        porc[b]['sem pedido'] += not e['tinha_pedido']
        porc[b]['pp0'] += pp == 0
        if not e['censurado']:
            dur[b].append(e['dias'])
    tot = sum(causa.values())
    blocos['causas'] = [['Causa da ruptura (início de cada episódio)', 'Episódios', '% dos episódios', 'R$ faltante no início']] + \
        [[k, v, v / tot, round(rs[k])] for k, v in causa.most_common()] + [['Total', tot, 1.0, round(sum(rs.values()))]]
    blocos['rup_comprador'] = [['Comprador', 'Episódios de ruptura', 'Duração mediana (dias corridos)', 'Sem pedido de compra no início', 'Dos quais PontoPedido 0']] + \
        [[b, c['n'], st.median(dur[b]) if dur[b] else None, c['sem pedido'], c['pp0']] for b, c in sorted(porc.items(), key=lambda x: -x[1]['n'])]

    # fornecedores
    fs = collections.defaultdict(list)
    for k, e in ind['po'].items():
        f = po_forn.get(k)
        if f and e['lt_real'] is not None:
            fs[f].append(e)
    lin = []
    for f, l in fs.items():
        if len(l) < 10:
            continue
        lt = sorted(e['lt_real'] for e in l)
        at = [e['atraso'] for e in l if e['atraso'] is not None]
        lin.append([f, len(l), st.median(lt), lt[int(0.9 * (len(lt) - 1))], sum(a <= 0 for a in at) / len(at) if at else None,
                    sum(e['mudou_data'] > 0 for e in l) / len(l), round(sum(e['qtd'] * e['preco'] for e in l))])
    lin.sort(key=lambda x: -x[6])
    blocos['fornecedores'] = [['Fornecedor', 'Pedidos com chegada', 'Prazo real mediano (dias)', 'Prazo real p90 (dias)', 'Chegou até a 1ª data prometida', 'Data de entrega remarcada', 'R$ dos pedidos']] + lin[:30]
    po = list(ind['po'].values())
    ok = [e for e in po if e['lt_real'] is not None]
    at = [e['atraso'] for e in ok if e['atraso'] is not None]
    lt = sorted(e['lt_real'] for e in ok)
    com_d = [x for x in bol if x[2] and isinstance(x[4], dt.date)]
    cump = [(x[4] - x[2]).days for x in com_d]
    blocos['prazos'] = [['Indicador', 'Valor'],
                        ['Pedidos de compra acompanhados nas fotos (código + pedido)', len(po)],
                        ['Prazo real, do pedido à chegada: mediana (dias)', st.median(lt)],
                        ['Prazo real: 75% chegam em até (dias)', lt[int(0.75 * (len(lt) - 1))]],
                        ['Prazo real: 90% chegam em até (dias)', lt[int(0.9 * (len(lt) - 1))]],
                        ['Chegaram até a primeira data prometida', sum(a <= 0 for a in at) / len(at)],
                        ['Pedidos com a data de entrega remarcada', sum(e['mudou_data'] > 0 for e in po if e['promessas']) / sum(1 for e in po if e['promessas'])],
                        ['Pedidos remarcados 3 vezes ou mais', sum(e['mudou_data'] >= 3 for e in po)],
                        ['Boletim de ruptura: linhas com data combinada', len(com_d)],
                        ['Boletim: ruptura resolvida até a data combinada (+1 dia)', sum(c <= 1 for c in cump) / len(cump)],
                        ['Boletim: atraso mediano contra a data combinada (dias)', st.median(cump)]]

    # ABC: estoque, ruptura, alvo
    rup = collections.defaultdict(lambda: [0, 0, 0.0])
    for d, l in fotos:
        for c, r in l.items():
            if num(r.get('pp')) <= 0 or num(r.get('med')) <= 0 or c not in prm:
                continue
            a = rup[prm[c]['abc']]
            a[0] += 1
            a[1] += num(r.get('rup')) < 0
            a[2] += min(12, num(r.get('disp')) / num(r.get('med')))
    ag = collections.defaultdict(lambda: collections.defaultdict(float))
    for l in linhas:
        k = l['param']['abc'] if l['param'] else 'C'
        c = l['claude']
        pr = c['preco']
        a = ag[k]
        a['n'] += 1
        a['est'] += num(l['estoque_rs'])
        a['medc'] += num(l['I']) * pr
        a['tot'] += num(l['E']) * pr
        a['cla'] += c['alvo'][0] * pr
    blocos['abc'] = [['Curva', 'SKUs', 'Venda mensal R$ (MedContr × preço)', 'Estoque hoje R$', 'Estoque hoje (meses de venda)',
                      'Alvo TOTVS (meses)', 'Alvo CLAUDE (meses)', 'Dias em ruptura (27/07 a 26/10)', 'Cobertura média nas fotos (meses)']]
    for k in 'ABC':
        a = ag[k]
        r = rup[k]
        blocos['abc'].append([k, int(a['n']), round(a['medc']), round(a['est']), a['est'] / a['medc'] if a['medc'] else None,
                              a['tot'] / a['medc'] if a['medc'] else None, a['cla'] / a['medc'] if a['medc'] else None,
                              r[1] / r[0] if r[0] else None, r[2] / r[0] if r[0] else None])

    # estoque por cobertura
    d, l = fotos[-1]
    b = collections.defaultdict(lambda: [0, 0.0])
    tt = 0
    for c, r in l.items():
        v = num(r.get('vlr'))
        tt += v
        med, fis = num(r.get('med')), num(r.get('fis'))
        if fis <= 0:
            k = '0 sem estoque'
        elif med <= 0:
            k = '6 com estoque e sem venda (MedContr 0)'
        else:
            x = fis / med
            k = '1 até 1 mês' if x <= 1 else '2 de 1 a 2 meses' if x <= 2 else '3 de 2 a 4 meses' if x <= 4 else '4 de 4 a 12 meses' if x <= 12 else '5 mais de 12 meses'
        b[k][0] += 1
        b[k][1] += v
    blocos['cobertura'] = [['Cobertura do estoque físico (meses de MedContr)', 'SKUs', 'Estoque R$', '% do valor']] + \
        [[k[2:], b[k][0], round(b[k][1]), b[k][1] / tt] for k in sorted(b)] + [['Total', sum(x[0] for x in b.values()), round(tt), 1.0]]

    # cenarios de parametros
    cen = []
    z0, r0 = dict(cc.Z), dict(cc.R_DIAS)
    for nome, z, R in [('Padrão desde 26/10: serviço A 95% / B 90% / C 80%; revisão A 7 / B 15 / C 30 dias', z0, r0),
                       ('Revisão quinzenal (A 15 / B 30 / C 30), serviço padrão', z0, {'A': 15, 'B': 30, 'C': 30}),
                       ('Mais serviço: A 98% / B 95% / C 90%; revisão 15 / 30 / 30', {'A': 2.0537489, 'B': 1.6448536, 'C': 1.2815516}, {'A': 15, 'B': 30, 'C': 30}),
                       ('Revisão mensal para todos (30 dias), serviço padrão', z0, {'A': 30, 'B': 30, 'C': 30})]:
        cc.Z, cc.R_DIAS = z, R
        a = q0 = q1 = q2 = 0.0
        for l in linhas:
            c = cc.calcula(l, l['param'], hoje)
            a += c['alvo'][0] * c['preco']
            q0 += c['rs'][0]
            q1 += c['rs'][1]
            q2 += c['rs'][2]
        cen.append([nome, round(a), round(q0), round(q1), round(q2)])
    cc.Z, cc.R_DIAS = z0, r0
    tot_e = sum(num(l['E']) * l['claude']['preco'] for l in linhas)
    cen.append(['Regra atual (PontoPedido = 2 meses de MedContr)', round(tot_e), round(sum(num(l['Compra mês vigente R$']) for l in linhas)),
                round(sum(num(l['Compra próximo mês R$']) for l in linhas)), round(sum(num(l['Compra mês subsequente R$']) for l in linhas))])
    blocos['cenarios'] = [['Cenário', 'Estoque alvo R$', 'Compra mês vigente R$', 'Compra próximo mês R$', 'Compra mês subsequente R$']] + cen

    # previsao
    T = tv['tot']
    M = ['2026-04', '2026-05', '2026-06', '2026-07', '2026-08', '2026-09']

    def wape(i, f):
        e = t = 0
        for c, s in T.items():
            if all(m in s for m in M) and sum(s[m] for m in M[i - 3:i]) > 0:
                e += abs(s[M[i]] - f([s[m] for m in M[:i]]))
                t += s[M[i]]
        return e / t
    blocos['previsao'] = [['Previsão do mês (erro absoluto ÷ venda, todos os SKUs)', 'Julho', 'Agosto', 'Setembro'],
                          ['MedContr (média dos 3 últimos meses)'] + [wape(i, lambda h: st.mean(h[-3:])) for i in (3, 4, 5)],
                          ['Só o último mês'] + [wape(i, lambda h: h[-1]) for i in (3, 4, 5)],
                          ['Ponderada 3-2-1 (último mês pesa mais)'] + [wape(i, lambda h: (3 * h[-1] + 2 * h[-2] + h[-3]) / 6) for i in (3, 4, 5)],
                          ['Mediana dos 3 últimos meses'] + [wape(i, lambda h: st.median(h[-3:])) for i in (3, 4, 5)]]

    # qualidade dos dados
    cob = tv['cob']
    blocos['dados'] = [['Ponto', 'Números'],
                       ['As bases de faturamento (fat AAAA.N) cobrem só parte da venda do TOTVS', f"{sum(1 for r in cob.values() if .8 <= r <= 1.25)} SKUs cobertos (80%-125%), {sum(1 for r in cob.values() if r < .2)} SKUs quase sem faturamento (< 20%) entre {len(cob)} com venda"],
                       ['A aba Planilha1 do "fat 2025.2.xlsx" é cópia exata de dezembro/2025 da Listagem', '11.211 notas iguais: somar as duas abas dobra dezembro'],
                       ['SKUs da Geral fora do relatório do Protheus', str(sum(1 for p in prm.values() if p['freq'] == '?'))],
                       ['PontoPedido zerado à mão com venda (ESTQ_MININ 0 e NOVO_PONTO > 0)', '266 no relatório do Protheus de 26/10'],
                       ['Disponível sem a vírgula decimal (SKUs de quantidade fracionada)', '5 SKUs (ver LEIA-ME COMPRAS)'],
                       ['Pedidos de compra sem fornecedor identificado nas fontes', f"{sum(1 for k in ind['po'] if not po_forn.get(k))} de {len(ind['po'])}"]]
    json.dump(blocos, open(S / 'analise.json', 'w', encoding='utf-8'), ensure_ascii=False, default=str, indent=0)
    print('ok', S)


if __name__ == '__main__':
    a = sys.argv[1:]
    h = a[a.index('--hoje') + 1] if '--hoje' in a else dt.date.today().isoformat()
    main(a[0], a[1], a[2], dt.date.fromisoformat(h))
