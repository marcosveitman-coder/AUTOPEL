"""Atualiza SO o potencial e (desde 10/11/2026) a carteira da aba ORGANOGRAMA, sem regravar a aba (05/11/2026).
A ORGANOGRAMA do arquivo tem um layout anterior ao do org_layout.py atual (resumo e legenda em outra posicao): regravar
pelo org_layout.py mudaria a aba inteira. Este script le a aba como esta e gera as trocas:
  - caixas: texto "Potencial: N oportunidades" (a pessoa e o nome mais proximo acima, na mesma coluna);
  - carteira: coluna do cabecalho POTENCIAL / OPORTUNID. (linha da pessoa pelo nome na coluna B);
  - tabela "2. POTENCIAL POR SEGMENTO": valores das linhas que ja existem e as linhas novas (segmentos novos do
    org_dados.json), inseridas pelo Excel logo depois da ultima linha da tabela, com o formato dela;
  - nota "Fonte: ..." do fim da aba.
Entrada: org_dados.json (org_dados.py) e o MAPA.  Saida: org_potencial.json para org_potencial.ps1.
Uso (na pasta dados): python ../org_potencial.py "<MAPA.xlsx>" """
import sys, re, json, html, zipfile
from openpyxl.utils import column_index_from_string as ci, get_column_letter as L
D = json.load(open('org_dados.json', encoding='utf8'))
ACENTO = {'ALIMENTACAO': 'Alimentação', 'CONSTRUCAO': 'Construção', 'POSTOS E CONVENIENCIAS': 'Postos e Conveniências',
          'DROGARIAS': 'Drogarias', 'REDE DE LOJAS': 'Rede de Lojas', 'ESTACIONAMENTOS': 'Estacionamentos', 'PET SHOP': 'Pet Shop'}
br = lambda n: f'{n:,}'.replace(',', '.')
pc = lambda a, b: f'{100 * a / b:.1f}%'.replace('.', ',') if b else '—'
FONTE = ('Fonte: ' + D.get('fonte_crm', 'telas do CRM (Segmentos & Vendedores e carteira de cada pessoa), recebidas em 04/11/2026') +
         '; potencial: abas ALIMENTACAO, '
         'REDE DE LOJAS, CONSTRUÇÃO, POSTOS GAS & CONVENIENCIA, ESTACIONAMENTO e PET SHOP FINAL (coluna W, STATUS AUTOPEL de 05/11/2026) e '
         '"FARMACIA VF" (coluna P) deste arquivo.')


def celulas(xlsx, aba):
    z = zipfile.ZipFile(xlsx)
    wb = z.read('xl/workbook.xml').decode(); rels = z.read('xl/_rels/workbook.xml.rels').decode()
    rid = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(aba), wb).group(1)
    tgt = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    p = 'xl/' + tgt.group(1).lstrip('/').replace('xl/', '')
    ss = [html.unescape(''.join(re.findall(r'<t[^>]*>(.*?)</t>', si, re.S)))
          for si in re.findall(r'<si>(.*?)</si>', z.read('xl/sharedStrings.xml').decode('utf8'), re.S)]
    out = {}
    for m in re.finditer(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', z.read(p).decode('utf8'), re.S):
        col, row, attr, body = m.groups()
        if not body: continue
        v = re.search(r'<v>(.*?)</v>', body); t = re.search(r'<t[^>]*>(.*?)</t>', body, re.S)
        val = ss[int(v.group(1))] if 't="s"' in attr and v else (html.unescape(t.group(1)) if t else (v.group(1) if v else ''))
        out[(int(row), ci(col))] = val
    return out


def carteira(C, trocas):
    """Numeros de carteira (10/11/2026), com os mesmos textos e contas do org_layout.py: linhas "Clientes:" e a de baixo
    das caixas, o RESUMO DAS EQUIPES e a tabela "1. CARTEIRA DE CADA PESSOA" (H..X) e a nota logo abaixo dela."""
    P = {p['nome']: p for p in D['pessoas']}
    def troca(r, c, novo, quem):
        antes = C.get((r, c))
        if isinstance(novo, str):
            igual = str(antes) == novo
        else:
            try: igual = abs(float(antes) - novo) < 1e-9
            except (TypeError, ValueError): igual = False
        if not igual: trocas.append(dict(ref=f'{L(c)}{r}', valor=novo, antes=antes, quem=quem))
    # caixas
    for (r, c), v in sorted(C.items()):
        if not str(v).startswith('Clientes: '): continue
        dono = next((C[(rr, c)].strip() for rr in range(r - 1, max(0, r - 6), -1) if str(C.get((rr, c), '')).strip() in P), None)
        assert dono, ('caixa sem nome acima', L(c) + str(r))
        p = P[dono]; no, fora = p['no_total'], p['fora_total']
        l4 = f"Clientes: {br(p['total'])}  |  no segmento {br(no)} ({pc(no, no + fora)})" if no + fora else f"Clientes: {br(p['total'])}  |  sem segmento definido"
        l5 = f"No segmento: ativos {pc(p['no'][0], no)} · perdendo {pc(p['no'][1], no)} · a trabalhar {pc(p['no'][2], no)}" if no else \
             f"Ativos {br(p['ativos'])} · perdendo {br(p['perdendo'])} · a trabalhar {br(p['atrab'])}"
        troca(r, c, l4, dono); troca(r + 1, c, l5, dono)
    # resumo das equipes (pessoas = linha que ja esta la)
    tit = [r for (r, c), v in C.items() if str(v).startswith('RESUMO DAS EQUIPES')]
    assert len(tit) == 1
    r0, c0 = next((r, c) for (r, c), v in C.items() if r == tit[0] + 1 and str(v).strip() == 'EQUIPE')
    cab = {str(v).strip(): c for (r, c), v in C.items() if r == r0}
    r = r0 + 1
    while str(C.get((r, c0), '')).startswith('Equipe '):
        nome = next(n for n in P if 'Equipe ' + n.title() == C[(r, c0)].strip())
        p = P[nome]; no, fora = p['no_total'], p['fora_total']
        for k, x in [('CLIENTES', p['total']), ('% NO SEGM.', no / (no + fora)), ('ATIVOS', p['no'][0]), ('% ATIVOS', p['no'][0] / no),
                     ('PERDENDO', p['no'][1]), ('% PERD.', p['no'][1] / no), ('A TRAB.', p['no'][2]), ('% A TRAB.', p['no'][2] / no)]:
            troca(r, cab[k], x, 'resumo ' + nome)
        r += 1
    # tabela 1: colunas H (8) a X (24), como no org_layout.py
    tit = [r for (r, c), v in C.items() if str(v).startswith('1. CARTEIRA DE CADA PESSOA')]
    assert len(tit) == 1
    r = tit[0] + 3
    assert C.get((r - 1, 8)) == 'CLIENTES' and C.get((r - 1, 24)) == '% A TRAB.'
    while str(C.get((r, 2), '')).strip() in P:
        p = P[C[(r, 2)].strip()]; no, fora = p['no_total'], p['fora_total']; tot = no + fora
        vals = [(8, p['total']), (9, no), (10, no / tot if tot else None), (11, fora), (12, fora / tot if tot else None)]
        for base, arr, n in [(13, p['no'], no), (19, p['fora'], fora)]:
            for i in range(3):
                vals += [(base + 2 * i, arr[i]), (base + 2 * i + 1, arr[i] / n if n else None)]
        for c, x in vals: troca(r, c, x if x is not None else '—', p['nome'])
        r += 1
    # nota logo abaixo da tabela
    if str(C.get((r, 2), '')).startswith('CLIENTES = '):
        sup = [p for p in D['pessoas'] if p['perfil'] in ('Supervisor', 'Supervisora')]
        semseg = [p for p in D['pessoas'] if (not p['sup'] or p in sup) and p['no_total'] + p['fora_total'] < p['total']]
        quem = ' e '.join(('na diretoria' if not p['sup'] else 'no supervisor ' + p['nome'].title()) for p in semseg)
        troca(r, 2, f'CLIENTES = ativos + perdendo + a trabalhar (cartões do CRM). {quem[0].upper() + quem[1:]} há clientes sem segmento definido: '
                    'por isso NO + FORA pode ser menor que CLIENTES. Os supervisores mostram a visão do CRM de cada um.', 'nota')


def main(mapa):
    C = celulas(mapa, 'ORGANOGRAMA')
    pot = {p['nome']: sum(p['potencial'].values()) for p in D['pessoas']}
    nomes = set(pot)
    trocas = []
    # caixas
    for (r, c), v in sorted(C.items()):
        if not str(v).startswith('Potencial: '): continue
        dono = next((C[(rr, c)] for rr in range(r - 1, max(0, r - 12), -1) if str(C.get((rr, c), '')).strip() in nomes), None)
        assert dono, ('caixa sem nome acima', L(c) + str(r), v)
        n = pot[dono.strip()]
        novo = ('Potencial: ' + br(n) + ' oportunidades') if n else 'Potencial: sem base nas abas de mercado'
        if novo != v: trocas.append(dict(ref=f'{L(c)}{r}', valor=novo, antes=v, quem=dono.strip()))
    # carteira: coluna com "OPORTUNID." no cabecalho
    hdr = [(r, c) for (r, c), v in C.items() if str(v).strip() == 'OPORTUNID.']
    assert len(hdr) == 1, hdr
    rh, cp = hdr[0]
    r = rh + 1
    while str(C.get((r, 2), '')).strip() in nomes:
        n = pot[C[(r, 2)].strip()]
        antes = C.get((r, cp))
        if str(antes) != (str(n) if n else '—'):
            trocas.append(dict(ref=f'{L(cp)}{r}', valor=n if n else '—', antes=antes, quem=C[(r, 2)].strip()))
        r += 1
    # tabela de potencial por segmento
    tit = [r for (r, c), v in C.items() if str(v).startswith('2. POTENCIAL POR SEGMENTO')]
    assert len(tit) == 1
    rh = tit[0] + 1
    cab = {str(C.get((rh, c), '')).strip(): c for c in range(1, 30) if C.get((rh, c))}
    COLS = {'TOTAL NA LISTA': None, 'ATIVO': 'ATIVO', 'INATIVO': 'INATIVO', 'INATIVANDO': 'INATIVANDO', 'LEAD NO CRM': 'LEAD NO CRM',
            'NÃO É CLIENTE': 'NÃO É CLIENTE', 'A TRABALHAR': 'A TRABALHAR'}
    linhas = {}
    r = rh + 1
    while C.get((r, 2)) and not str(C[(r, 2)]).startswith('POTENCIAL ='):
        linhas[str(C[(r, 2)]).strip()] = r; r += 1
    ultima = r - 1
    def valores(seg, b):
        st = b['status']
        v = {'SEGMENTO': ACENTO.get(seg, seg.title()), 'ABA (COLUNA STATUS AUTOPEL)': f"{b['aba']} (coluna {b['col']})"}
        for k, s in COLS.items(): v[k] = b['total'] if s is None else st.get(s, 0)
        v['POTENCIAL'] = b['potencial']
        v['RESPONSÁVEIS (★ principal)'] = ', '.join(D['segmentos'].get(seg, []))
        return v
    novas = []
    for seg, b in D['potencial'].items():
        v = valores(seg, b)
        if v['SEGMENTO'] in linhas:
            rr = linhas[v['SEGMENTO']]
            for k, x in v.items():
                c = cab[k]
                if str(C.get((rr, c), '')) != str(x):
                    trocas.append(dict(ref=f'{L(c)}{rr}', valor=x, antes=C.get((rr, c)), quem=v['SEGMENTO']))
        else:
            novas.append([dict(col=cab[k], valor=x) for k, x in v.items()])
    carteira(C, trocas)
    # nota da fonte
    fonte = [(r, c) for (r, c), v in C.items() if str(v).startswith('Fonte: ') and 'potencial: abas' in str(v)]
    assert len(fonte) == 1
    if C[fonte[0]] != FONTE: trocas.append(dict(ref=f'{L(fonte[0][1])}{fonte[0][0]}', valor=FONTE, antes=C[fonte[0]], quem='fonte'))
    json.dump(dict(trocas=trocas, apos_linha=ultima, novas=novas), open('org_potencial.json', 'w', encoding='utf8'), ensure_ascii=False, indent=0)
    for t in trocas: print(f"{t['ref']:6} {t['quem'][:24]:24} {str(t['antes'])[:40]:40} -> {str(t['valor'])[:40]}")
    print(len(trocas), 'células trocadas |', len(novas), f'linhas novas depois da linha {ultima}:', [n[0]['valor'] for n in novas])


if __name__ == '__main__':
    main(sys.argv[1])
