"""STATUS RECEITA / PORTE / STATUS AUTOPEL (05/11/2026) - passo 3: valores finais das 3 colunas.
STATUS RECEITA = situacao do CNPJ da coluna F na copia local da Receita (r19.pkl de busca_receita.py q19.json): ATIVA,
  SUSPENSA, INAPTA, BAIXADA ou NULA; "SEM CNPJ" quando a linha nao tem CNPJ.
PORTE = faixas dos supermercados e farmacias (A >= R$ 1 bi; B R$ 300 mi a 1 bi; C R$ 20 a 300 mi; D < R$ 20 mi):
  1) valor que o usuario ja tinha gravado na coluna (fica como esta);
  2) faturamento publicado: IRTT por bandeira (coluna AA, pela raiz do CNPJ validado ou pelo nome igual);
     na Alessandra, o PORTE da FARMACIA VF / SUPERMERCADOS (reuniao);
  3) estimativa: no de lojas x faturamento medio por loja do segmento (POR_LOJA abaixo, com a fonte).
STATUS AUTOPEL = status_autopel.py (ATIVO < 360 dias; INATIVO; NAO E CLIENTE). Na Alessandra, a empresa que esta numa
  aba FINAL recebe os mesmos 3 valores dessa linha.
Saidas: dados/status_final.pkl e a planilha de conferencia (1o argumento), com a data e o vinculo de cada linha.
Uso: python status_final.py "<conferencia.xlsx>" """
import os, sys, re, pickle, statistics as st, collections as C, datetime as dt
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
HERE = os.path.dirname(os.path.abspath(__file__))
DADOS = os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or os.path.join(HERE, 'dados')
sys.path.insert(0, HERE)
from status_autopel import nucleo, apelidos
D = lambda x: re.sub(r'\D', '', str(x or ''))
s = lambda x: str(x).strip() if x is not None else ''
SIT = {'01': 'NULA', '02': 'ATIVA', '03': 'SUSPENSA', '04': 'INAPTA', '08': 'BAIXADA'}
FAIXAS = [(1e9, 'A – Grande (≥ R$ 1 bi)'), (3e8, 'B – Média (R$ 300 mi a 1 bi)'), (2e7, 'C – Regional (R$ 20 mi a 300 mi)'),
          (0, 'D – Pequena / local (< R$ 20 mi)')]
ABF = 301.7e9 / 202444  # ABF 2025: R$ 301,7 bi / 202.444 unidades (media por unidade de franquia)
POR_LOJA = {  # (R$ por loja/ano, fonte)
    'franquia': (ABF, 'ABF 2025: R$ 301,7 bi ÷ 202.444 unidades'),
    'posto': (15e6, 'ANP 2025: ~108 bi de litros nos postos × ~R$ 6,2 ÷ ~43 mil postos'),
    'estac': (1872.8e6 / 827, 'Estapar 2025: R$ 1,87 bi ÷ 827 operações'),
    'drogaria': (2.2e6, 'mediana do Ranking de farmácias (regra da FARMACIA VF)'),
    'minimercado': (336e6 / 2600, 'Market4U 2025: R$ 336 mi ÷ ~2.600 unidades'),
}
DISTRIB_COMB = (0.44 * 700e9 / 182, 'distribuidora de combustível sem nº de postos: as 182 distribuidoras regionais têm 44% do mercado '
                '(Eixos, 2025), ~R$ 1,7 bi cada em média – estimativa')
porte = lambda v: next(n for lim, n in FAIXAS if v >= lim)


def main(conf_xlsx):
    I = pickle.load(open(os.path.join(DADOS, 'status_itens.pkl'), 'rb'))
    S = pickle.load(open(os.path.join(DADOS, 'status_autopel.pkl'), 'rb'))
    R = pickle.load(open(os.path.join(DADOS, 'r19.pkl'), 'rb'))
    est = {x[0] + x[1] + x[2]: x for x in R['est']}
    out = S['out']
    # faturamento por loja das bandeiras do IRTT (2024/2025): 1o quartil por segmento, para lojas e construcao
    q = C.defaultdict(list)
    for v in I['irtt_raiz'].values():
        for b in v:
            if isinstance(b['lojas'], (int, float)) and b['lojas'] > 0: q[b['seg']].append(b['fat'] / b['lojas'])
    p25 = lambda segs: sorted(x for k in segs for x in q[k])[len([x for k in segs for x in q[k]]) // 4]
    LOJ = ['Moda, Calçados e Artigos Esportivos', 'Lojas de Departamento, Artigos do Lar e Mercadorias em Geral', 'Eletromóveis',
           'Óticas, Jóias, Bijoux, Bolsas e Acessórios', 'Outros Segmentos']
    LOJ.append('Material de Construção')
    POR_LOJA['lojas'] = (p25(LOJ), f'IRTT: 1º quartil do faturamento por loja das bandeiras de varejo não alimentar, com construção ({sum(len(q[k]) for k in LOJ)} bandeiras)')
    POR_LOJA['constr'] = POR_LOJA['lojas']
    for k, (v, f) in POR_LOJA.items(): print(f'  {k:9} R$ {v / 1e6:5.2f} mi/loja  ({f})')
    irtt_nucleo = C.defaultdict(list)
    for b in I['irtt_nome']:
        if isinstance(b['fat'], (int, float)) and b['fat'] > 0 and len(nucleo(b['nome']).replace(' ', '')) >= 5:
            irtt_nucleo[nucleo(b['nome'])].append(b)

    def tipo(it):
        a, sg = it['aba'], common_norm(it.get('segm'))
        if re.search(r'MINIMERCADO|AUTONOMO', sg): return 'minimercado'
        if a == 'POSTOS GAS & CONVENIENCIA FINAL':
            return 'posto' if re.search(r'POSTOS|DISTRIBUIDORA|COOPERATIVA|TROCA DE OLEO', sg) and 'CONVENIENCIA' not in sg else 'franquia'
        return {'REDE DE LOJAS FINAL': 'lojas', 'CONSTRUÇÃO FINAL': 'constr', 'ESTACIONAMENTO FINAL': 'estac'}.get(a, 'franquia')

    def sit_receita(c, alt=''):
        if len(c) != 14: return 'SEM CNPJ'
        x = est.get(c)
        return SIT.get(x[5], x[5]) if x else (s(alt).upper() or 'NÃO ACHADO NA RECEITA')

    fin, conf = {}, []
    for it in I['itens']:
        k = (it['aba'], it['lin'])
        o = out[k]
        U = sit_receita(it['cnpj'] or it['ref'], it.get('sit'))  # sem F: situacao da unidade de referencia (H)
        V0 = s(it.get('V'))
        fat, fonte = None, ''
        r8 = it['cnpj'][:8]
        nucs = {nucleo(x) for x in apelidos(it.get('nome'))} | {nucleo(it.get('nome'))}
        # valor do GRUPO (todas as bandeiras) so vale para a linha que e o proprio grupo
        varias = bool(re.search(r'[(/+]|GRUPO', s(it.get('nome')).upper()))  # a linha e um grupo de marcas
        ok = lambda x: 'valor do GRUPO' not in (x['fonte'] or '') or varias
        bs = [x for x in I['irtt_raiz'].get(r8, []) if nucleo(x['nome']) in nucs and ok(x)] if len(r8) == 8 else []
        if bs:
            fat, fonte = bs[0]['fat'], f"IRTT por bandeira ({bs[0]['nome']}, mesmo CNPJ): {bs[0]['fonte']}"
        else:
            bs = [x for n_ in nucs for x in irtt_nucleo.get(n_, []) if ok(x) and (not x['cnpj'] or len(r8) != 8 or x['cnpj'][:8] == r8)]
            if len({x['lin'] for x in bs}) == 1:
                fat, fonte = bs[0]['fat'], f"IRTT por bandeira ({bs[0]['nome']}, mesmo nome): {bs[0]['fonte']}"
        lojas = it.get('lojas') if isinstance(it.get('lojas'), (int, float)) and it['lojas'] > 0 else None
        if fat is None and not lojas and isinstance(it.get('rede'), (int, float)) and it['rede'] > 0:
            lojas = it['rede']  # sem o no de lojas: unidades ativas com o nome na Receita (coluna Q)
        if fat is None and not lojas and 'DISTRIBUIDORA' in common_norm(it.get('segm')):
            fat, fonte = DISTRIB_COMB
        if fat is None and lojas:
            pl, f = POR_LOJA[tipo(it)]
            fat, fonte = lojas * pl, f'estimativa: {int(lojas)} lojas × R$ {pl / 1e6:.2f} mi ({f})'
        Vc = porte(fat) if fat is not None else 'SEM Nº DE LOJAS'
        V = V0 or Vc
        fin[k] = dict(U=U, V=V, W=o['status'])
        conf.append([it['aba'], it['lin'], s(it.get('nome')), it['cnpj'], U, V, 'usuário (mantido)' if V0 else 'calculado', Vc if V0 else '',
                     fat, fonte, o['status'], s(it.get('W')), o['dt'], o['dias'], o['via'], o['cliente'], D(o['cnpj']), o['fonte'], o['vend'], o['alerta']])
    # Alessandra Dayan
    fale = {}
    for a in I['ale']:
        if a['liga']:
            f = fin[a['liga']]
            fale[a['ref']] = dict(f, via=f"igual à {a['liga'][0]} linha {a['liga'][1]}")
            continue
        o = out[('Alessandra Dayan', a['ref'])]
        c = next((x['cnpj'] for x in a['apoio'] if len(x['cnpj']) == 14), '')
        U = sit_receita(c, next((x['sit'] for x in a['apoio'] if x['sit']), ''))
        V = next((x['porte'] for x in a['apoio'] if x['porte']), '')
        fonte = next((f"PORTE da aba {x['fonte']} (linha {x['lin']})" for x in a['apoio'] if x['porte']), '')
        fat = None
        if not V:
            fx = [x for x in a['apoio'] if isinstance(x.get('fat'), (int, float)) and x['fat'] > 0]
            if fx: fat, fonte = fx[0]['fat'], f"faturamento da aba {fx[0]['fonte']} (linha {fx[0]['lin']})"
            elif isinstance(a['lojas'], (int, float)) and a['lojas'] > 0:
                pl, f = POR_LOJA['drogaria' if a['seg'] == 'DROGARIAS' else 'franquia']
                fat, fonte = a['lojas'] * pl, f"estimativa: {int(a['lojas'])} lojas × R$ {pl / 1e6:.2f} mi ({f})"
            V = porte(fat) if fat is not None else 'SEM Nº DE LOJAS'
        fale[a['ref']] = dict(U=U, V=V, W=o['status'], via='calculado')
        conf.append(['Alessandra Dayan', a['ref'], a['nome'], c, U, V, 'aba de apoio' if fonte.startswith('PORTE') else 'calculado', '', fat, fonte,
                     o['status'], '', o['dt'], o['dias'], o['via'], o['cliente'], D(o['cnpj']), o['fonte'], o['vend'], o['alerta']])
    pickle.dump(dict(fin=fin, ale=fale, gaps=I['gaps'], ale_pos={a['ref']: (a['lin'], a['col'], a['seg']) for a in I['ale']}),
                open(os.path.join(DADOS, 'status_final.pkl'), 'wb'))
    # planilha de conferencia
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'conferencia'
    H = ['ABA', 'LINHA / CÉLULA', 'REDE', 'CNPJ (F)', 'STATUS RECEITA', 'PORTE', 'ORIGEM DO PORTE', 'PORTE CALCULADO (quando o do usuário foi mantido)',
         'FATURAMENTO USADO (R$)', 'FONTE DO FATURAMENTO', 'STATUS AUTOPEL', 'STATUS AUTOPEL ANTERIOR (coluna W)', 'ÚLTIMA COMPRA', 'DIAS SEM COMPRAR',
         'VÍNCULO (COMO FOI ACHADO)', 'CLIENTE QUE COMPROU', 'CNPJ DO CLIENTE', 'FONTE DA DATA', 'VENDEDOR DA CARTEIRA (CADASTRO)', 'ALERTA']
    ws.append(H)
    for r in conf: ws.append(r)
    for c in ws[1]:
        c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F3864'); c.alignment = Alignment(wrap_text=True, vertical='top')
    for row in ws.iter_rows(min_row=2, min_col=13, max_col=13):
        for c in row: c.number_format = 'DD/MM/YYYY'
    for row in ws.iter_rows(min_row=2, min_col=9, max_col=9):
        for c in row: c.number_format = '#,##0'
    for i, w in enumerate([24, 9, 34, 16, 14, 30, 16, 26, 16, 60, 14, 16, 12, 10, 60, 40, 16, 34, 30, 30], start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = 'C2'; ws.auto_filter.ref = ws.dimensions
    r2 = wb.create_sheet('resumo')
    r2.append(['ABA', 'ATIVO', 'INATIVO', 'NÃO É CLIENTE', 'TOTAL'])
    cnt = C.defaultdict(C.Counter)
    for r in conf: cnt[r[0]][r[10]] += 1
    for k, v in cnt.items(): r2.append([k, v['ATIVO'], v['INATIVO'], v['NÃO É CLIENTE'], sum(v.values())])
    r2.append([]); r2.append(['Regras'])
    for t in [f"STATUS AUTOPEL em {S['hoje']:%d/%m/%Y}: ATIVO = última compra há menos de 360 dias; INATIVO = comprou há 360 dias ou mais, ou está "
              "cadastrado no TOTVS sem compra; NÃO É CLIENTE = não achado no cadastro nem no faturamento.",
              f"Datas: bases de faturamento 2019–2026 (até {S['fim']:%d/%m/%Y}) e, depois delas, o Ultimo faturamento do TOTVS de {S['hoje']:%d/%m/%Y}.",
              'Vínculos: CNPJ da rede > razão social > grupo do cadastro com o nome da rede > nome da rede no cliente, no produto personalizado ou no e-mail.',
              'PORTE: A ≥ R$ 1 bi; B R$ 300 mi a 1 bi; C R$ 20 a 300 mi; D < R$ 20 mi. Valor do usuário mantido; senão, faturamento do IRTT; senão, nº de lojas × média por loja:'] + \
             [f'   {k}: R$ {v / 1e6:.2f} mi por loja/ano – {f}' for k, (v, f) in POR_LOJA.items()]:
        r2.append([t])
    r2.column_dimensions['A'].width = 40
    wb.save(conf_xlsx)
    print('conferência:', conf_xlsx)
    for k, v in cnt.items(): print(f'  {k:32} {dict(v)}')
    print('PORTE:', C.Counter(f['V'][:1] for f in fin.values()), '| mantidos do usuário:', sum(1 for r in conf if r[6] == 'usuário (mantido)'),
          '| divergentes:', sum(1 for r in conf if r[7] and r[7] != r[5]))
    print('STATUS RECEITA:', C.Counter(f['U'] for f in fin.values()))


def common_norm(x):
    import unicodedata
    return unicodedata.normalize('NFKD', str(x or '')).encode('ascii', 'ignore').decode().upper()


if __name__ == '__main__':
    main(sys.argv[1])
