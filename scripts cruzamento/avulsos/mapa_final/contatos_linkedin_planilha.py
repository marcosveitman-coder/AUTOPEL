"""Planilha de contatos (LinkedIn) das bandeiras REDE PRÓPRIA da aba "Alessandra Dayan VF" (17/11/2026).
Pedido do usuário: para TODAS as bandeiras com "REDE PRÓPRIA" na coluna A, nome + perfil do LinkedIn de 5 cargos
(gerente de compras, coordenador de compras, diretor de compras, diretor de supply, CEO). Sem e-mail e sem telefone.

Entradas (locais, fora do git, têm nomes de pessoas):
  redes.json  lista das bandeiras [{rede, seg, uf, lojas, cx, busca}] na ordem da aba (índice = posição)
  lnk.tsv     pesquisa: idx, cargo (gerente|coord|dircompras|supply|ceo), nome, url, evidencia (cargo como aparece
              no LinkedIn / buscador). Várias pessoas por cargo: a 1ª é a principal, as outras vão em "Outros".
Como a pesquisa foi feita: buscador (Firecrawl) com site:br.linkedin.com/in "<marca>" + cargos; uma busca por marca
grande e 4-10 marcas pequenas por busca. O lnk.tsv que entra aqui já passou pela regra estrita (r3_estrito, 17/11:
o usuário achou um "diretor de supply" que era supervisor): só o cargo exato, atual, na empresa; o resto fica em branco.
Conferência do cargo atual perfil a perfil: firecrawl_scrape do perfil com formats ["query"] (~14 créditos por perfil).
2ª rodada (opcional, bandeiras sem nenhum nome): sócios e administradores da Receita (QSA pela BrasilAPI / CNPJá aberto)
  qsa.json    {idx: {cnpj, razao, qsa:[{nome, qual}], nota}}
  dec.json    {idx: [{nome, curto, qual}]} decisores (sócio-administrador, administrador, diretor, presidente; sem PJ e menor)
  r2lnk.tsv   idx, nome_receita, url, evidencia: perfil do LinkedIn achado para o sócio (ou outra pessoa do QSA)
  vão na coluna "SÓCIO / ADMINISTRADOR (RECEITA)"; quem tem LinkedIn vem primeiro.
Saída: xlsx novo (openpyxl; NÃO é o MAPA) com as abas CONTATOS (1 linha por bandeira), PESSOAS (1 linha por pessoa)
e COMO LER.
Uso: python contatos_linkedin_planilha.py redes.json lnk.tsv saida.xlsx [qsa.json dec.json r2lnk.tsv]"""
import sys, json, collections as C
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

CARGOS = [('gerente', 'GERENTE DE COMPRAS'), ('coord', 'COORDENADOR DE COMPRAS'), ('dircompras', 'DIRETOR DE COMPRAS'),
          ('supply', 'DIRETOR DE SUPPLY'), ('ceo', 'CEO'), ('socio', 'SÓCIO / ADMINISTRADOR (RECEITA)')]
CORES = ['DDEBF7', 'E2EFDA', 'FFF2CC', 'FCE4D6', 'EDE2F6', 'E7E6E6']
TERC = ['Cargo no LinkedIn'] * 5 + ['Na Receita / no LinkedIn']
NAO = ''


def main(f_redes, f_lnk, saida, f_qsa=None, f_dec=None, f_r2=None):
    redes = json.load(open(f_redes, encoding='utf8'))
    pes = C.defaultdict(list)
    vistos = set()
    for l in open(f_lnk, encoding='utf8'):
        p = l.rstrip('\r\n').split('\t')
        if len(p) < 5 or not p[0].isdigit(): continue
        k = (int(p[0]), p[1], p[3].split('?')[0].rstrip('/') or p[2])
        if k in vistos: continue
        vistos.add(k)
        pes[(int(p[0]), p[1])].append(dict(nome=p[2], url=p[3], ev=p[4]))
    n_socio = 0
    if f_qsa:
        qsa = json.load(open(f_qsa, encoding='utf8')); dec = json.load(open(f_dec, encoding='utf8'))
        r2 = C.defaultdict(dict)
        for l in open(f_r2, encoding='utf8'):
            p = l.rstrip('\r\n').split('\t')
            if len(p) >= 4 and p[0].isdigit(): r2[p[0]].setdefault(p[1].upper(), (p[2], p[3]))
        for i, ps in dec.items():
            x = qsa[i]; c = x['cnpj']; cf = f'{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}'
            base_ev = lambda q: f"{q} na Receita (CNPJ {cf}, {x.get('razao', '')})" + (f" — {x['nota']}" if x.get('nota') else '')
            lst = []
            for pz in ps:
                u, ev = r2[i].pop(pz['nome'].upper(), ('', ''))
                lst.append(dict(nome=pz['nome'], url=u, ev=base_ev(pz['qual']) + (f' | LinkedIn: {ev}' if ev else '')))
            for nome, (u, ev) in r2[i].items():   # pessoa do QSA que não é administradora, mas tem LinkedIn
                q = next((s['qual'] for s in x.get('qsa', []) if s['nome'].upper() == nome), 'Sócio')
                lst.append(dict(nome=nome.title(), url=u, ev=base_ev(q) + f' | LinkedIn: {ev}'))
            lst.sort(key=lambda z: not z['url'])
            if lst: pes[(int(i), 'socio')] = lst; n_socio += 1

    wb = Workbook()
    ws = wb.active; ws.title = 'CONTATOS'
    fino = Side(style='thin', color='BFBFBF'); borda = Border(left=fino, right=fino, top=fino, bottom=fino)
    cab = Font(name='Calibri', bold=True, color='FFFFFF'); fcab = PatternFill('solid', fgColor='1F4E78')
    link = Font(name='Calibri', color='0563C1', underline='single')
    base = ['Nº', 'BANDEIRA', 'SEGMENTO', 'UF', 'Nº DE LOJAS', 'CAIXAS/MÊS', 'CARGOS ACHADOS (de 5)']
    ws.append([''] * len(base) + sum([[t, '', ''] for _, t in CARGOS], []) + ['OUTROS NOMES ACHADOS'])
    ws.append(base + sum([['Nome', 'LinkedIn', t] for t in TERC], []) + [''])
    nb = len(base)
    for j, (_, t) in enumerate(CARGOS):
        c0 = nb + 1 + 3 * j
        ws.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c0 + 2)
        for r in (1, 2):
            for c in range(c0, c0 + 3):
                ws.cell(r, c).fill = PatternFill('solid', fgColor=CORES[j]); ws.cell(r, c).font = Font(bold=True)
    for c in range(1, nb + 1):
        ws.merge_cells(start_row=1, start_column=c, end_row=2, end_column=c)
    co = nb + 3 * len(CARGOS) + 1
    ws.merge_cells(start_row=1, start_column=co, end_row=2, end_column=co)
    for r in (1, 2):
        for c in list(range(1, nb + 1)) + [co]:
            ws.cell(r, c).fill = fcab; ws.cell(r, c).font = cab
        for c in range(1, co + 1):
            ws.cell(r, c).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True); ws.cell(r, c).border = borda

    lista = []
    tot = C.Counter()
    for i, r in enumerate(redes):
        lin = [i + 1, r['rede'], r['seg'], r['uf'], r['lojas'], r['cx']]
        achados = 0; cel = []; outros = []
        for k, t in CARGOS:
            ps = pes.get((i, k), [])
            if ps:
                achados += k != 'socio'; tot[k] += 1
                cel += [ps[0]['nome'], ps[0]['url'], ps[0]['ev']]
                outros += [f"{t.lower()}: {p['nome']} ({p['ev']}) {p['url']}" for p in ps[1:]]
            else:
                cel += [NAO, '', '']
            for p in ps:
                lista.append([r['rede'], r['seg'], r['uf'], t, p['nome'], p['url'], p['ev']])
        tot['com_alguem'] += achados > 0
        tot['com_nome'] += achados > 0 or (i, 'socio') in pes
        ws.append(lin + [achados] + cel + ['\n'.join(outros)])
        rr = ws.max_row
        for j in range(len(CARGOS)):
            c = ws.cell(rr, nb + 2 + 3 * j)
            if c.value: c.hyperlink = c.value; c.font = link
            if ws.cell(rr, nb + 1 + 3 * j).value == NAO: ws.cell(rr, nb + 1 + 3 * j).font = Font(color='A6A6A6', italic=False)
        for c in range(1, co + 1):
            ws.cell(rr, c).border = borda; ws.cell(rr, c).alignment = Alignment(vertical='top', wrap_text=c in (2, co) or (c > nb and (c - nb) % 3 == 0))
        ws.cell(rr, 5).number_format = '#,##0'; ws.cell(rr, 6).number_format = '#,##0'
    larg = [5, 38, 18, 5, 9, 10, 10] + [24, 30, 38] * len(CARGOS) + [60]
    for c, w in enumerate(larg, 1): ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = 'C3'
    ws.auto_filter.ref = f'A2:{get_column_letter(co)}{ws.max_row}'
    ws.row_dimensions[1].height = 22; ws.row_dimensions[2].height = 30

    w2 = wb.create_sheet('PESSOAS')
    w2.append(['BANDEIRA', 'SEGMENTO', 'UF', 'CARGO PROCURADO', 'NOME', 'LINKEDIN', 'CARGO NO LINKEDIN'])
    for x in lista: w2.append(x)
    for c in range(1, 8):
        w2.cell(1, c).font = cab; w2.cell(1, c).fill = fcab
    for rr in range(2, w2.max_row + 1):
        c = w2.cell(rr, 6)
        if c.value: c.hyperlink = c.value; c.font = link
    for c, w in enumerate([38, 18, 5, 24, 30, 55, 70], 1): w2.column_dimensions[get_column_letter(c)].width = w
    w2.freeze_panes = 'A2'; w2.auto_filter.ref = f'A1:G{w2.max_row}'

    w3 = wb.create_sheet('COMO LER')
    txt = [
        'CONTATOS NO LINKEDIN - BANDEIRAS "REDE PRÓPRIA" DA ABA "Alessandra Dayan VF"',
        '',
        f'Bandeiras: {len(redes)} (todas com REDE PRÓPRIA na coluna A). Com pelo menos 1 dos 5 cargos: {tot["com_alguem"]}. '
        f'Com algum nome (cargos ou sócio da Receita): {tot["com_nome"]}. Pessoas: {len(lista)}.',
        'Achados por cargo: ' + ', '.join(f'{t.lower()} {tot[k]}' for k, t in CARGOS) + '.',
        '',
        'Aba CONTATOS: 1 linha por bandeira; para cada cargo, nome, link do perfil e o cargo como aparece no LinkedIn.',
        'Quando há mais de uma pessoa para o mesmo cargo, a 1ª fica na coluna do cargo e as outras em "OUTROS NOMES ACHADOS".',
        'Aba PESSOAS: 1 linha por pessoa (boa para filtrar ou importar no CRM).',
        '',
        'Como foi feito: busca pública do LinkedIn por buscador (perfis br.linkedin.com/in), com o nome da marca e os cargos.',
        'REGRA (estrita): só entra quem tem HOJE, naquela empresa (ou no grupo dono da bandeira), o cargo da coluna:',
        '  GERENTE DE COMPRAS = gerente/gestor de compras, suprimentos ou procurement. Gerente comercial, de categoria ou de logística NÃO.',
        '  COORDENADOR DE COMPRAS = coordenador de compras ou suprimentos. Comprador e analista NÃO.',
        '  DIRETOR DE COMPRAS = diretor/head de compras, suprimentos ou procurement, CPO. Diretor comercial NÃO.',
        '  DIRETOR DE SUPPLY = diretor/head/VP de supply chain ou logística, CSCO. Gerente ou supervisor de supply/logística NÃO.',
        '  CEO = CEO, presidente, diretor-presidente, diretor geral, diretor superintendente, country manager. VP, CFO, COO,',
        '  presidente do conselho, sócio ou diretor sem ser o principal executivo NÃO.',
        'Cargo incerto, perfil que não diz a empresa ou que já saiu = célula EM BRANCO. Empresas do mesmo grupo dividem as pessoas',
        '(ex.: GPA e Minuto Pão de Açúcar; Grupo Boticário e The Beauty Box/Eudora; DIA e DIA%).',
        'Em branco = nenhum perfil público com esse cargo. É comum nas redes pequenas (postos, estacionamentos, lojas regionais),',
        'onde o dono faz as compras: veja a coluna do sócio da Receita.',
        'SÓCIO / ADMINISTRADOR (RECEITA): 2ª rodada, só nas bandeiras em que nenhum dos 5 cargos apareceu no LinkedIn. São os',
        'sócios-administradores, administradores e diretores do quadro societário (QSA) da Receita Federal, com o CNPJ consultado.',
        'Nas redes pequenas, o dono é quem compra. Quando o sócio tem perfil no LinkedIn, o link vem ao lado.',
        'Bandeiras sem CNPJ confirmado ficaram sem sócio. O LinkedIn do sócio só vem quando o perfil mostra a empresa ou o grupo.',
        f'Bandeiras com sócio da Receita: {n_socio}.',
        'Sem e-mail e sem telefone, como pedido.',
    ]
    for t in txt: w3.append([t])
    w3['A1'].font = Font(bold=True, size=13); w3.column_dimensions['A'].width = 130
    wb.save(saida)
    print(f'bandeiras {len(redes)} | com cargo {tot["com_alguem"]} | com algum nome {tot["com_nome"]} | pessoas {len(lista)} | ' +
          ' '.join(f'{k}={tot[k]}' for k, _ in CARGOS))


if __name__ == '__main__':
    main(*sys.argv[1:7])
