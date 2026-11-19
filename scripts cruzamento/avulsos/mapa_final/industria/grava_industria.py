"""Aba FULLFILMIT_INDUSTRIA FINAL, passo 6: grava a aba (vazia, criada pelo usuário) e ACRESCENTA blocos no fim das abas
MAIORES POR REGIÃO e R_E_S_U_M_O (19/11/2026). Só pelo XML: a pasta tem tabelas dinâmicas e macro; nenhuma outra aba é tocada.
Layout igual às abas FINAL desde 06/11 e 16/11: tabela a/b/c/d (W1:X4, 17/12/7/4), cabeçalho na linha 6, dados da linha 7,
colunas A–AA; X = valor da 1ª letra do PORTE × N(nº de unidades), com o valor já calculado; Y–AA do STATUS AUTOPEL.
Entrada: ind_rows.pkl e ind_status.pkl. Uso: python grava_industria.py <pasta de trabalho> <MAPA.xlsm> <saída.xlsm>"""
import os, sys, re, pickle, zipfile, datetime as dt, collections
from xml.sax.saxutils import escape
from openpyxl.utils import get_column_letter
P, MAPA, OUT = sys.argv[1:4]
ABA, REGIAO, RESUMO = 'FULLFILMIT_INDUSTRIA FINAL', 'MAIORES POR REGIÃO', 'R_E_S_U_M_O'
S_HEAD, S_NUM, S_DATA, S_TAB1, S_SEC, S_TAB = 194, 127, 128, 116, 50, 52
DATA = '19/11/2026'
CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
COLS = [("Nº", 6), ("NOME FANTASIA", 34), ("RAZÃO SOCIAL", 46), ("GRUPO ADMINISTRADOR", 30), ("Nº DE UNIDADES (ESTABELECIMENTOS ATIVOS)", 13),
        ("CNPJ", 20), ("CONFERÊNCIA DO CNPJ", 30), ("CNPJ DE REFERÊNCIA (UNIDADE DA REDE)", 20), ("EMPRESA DE REFERÊNCIA", 30), ("UF SEDE", 7),
        ("MUNICÍPIO SEDE", 22), ("REGIÃO DA SEDE", 14), ("RANKING BRASIL (PORTE)", 11), ("RANKING NA REGIÃO (PORTE)", 11), ("SITUAÇÃO NA RECEITA", 12),
        ("FONTE DO Nº DE UNIDADES", 36), ("UNIDADES COM CNAE INDUSTRIAL DO SEGMENTO", 15), ("SEGMENTO", 30), ("ORIGEM", 42), ("OBSERVAÇÃO", 80),
        ("STATUS RECEITA", 16), ("PORTE", 30), ("STATUS AUTOPEL", 16), ("Total de caixas/ mês", 14.7109375), ("ÚLTIMA COMPRA AUTOPEL", 14.7109375),
        ("VENDEDOR (CADASTRO)", 32.7109375), ("COMO FOI ACHADO", 70.7109375)]
CX = {'A': 17, 'B': 12, 'C': 7, 'D': 4}
VIA = {'CNPJ': 'CNPJ (raiz)', 'CODIGO': 'código de cliente', 'RAZAO': 'razão social igual', 'ASSOCIADA': 'associada da rede',
       'GRUPO MAPA': 'grupo do cadastro (mapa)', 'GRUPO ERP': 'grupo do ERP', 'GRUPO': 'grupo do cliente', 'EXATO': 'nome fantasia igual',
       'NOME': 'nome do cliente', 'EMAIL': 'e-mail', 'PRODUTO': 'produto personalizado'}
rows = pickle.load(open(os.path.join(P, 'ind_rows.pkl'), 'rb'))
ST = pickle.load(open(os.path.join(P, 'ind_status.pkl'), 'rb'))
MAN = {}
if os.path.exists(os.path.join(P, 'ind_status_manual.pkl')): MAN = pickle.load(open(os.path.join(P, 'ind_status_manual.pkl'), 'rb'))


def fmt(c): return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}" if len(c) == 14 else c


def cell(ref, v, s=None, f=None):
    st = f' s="{s}"' if s else ''
    if f is not None: return f'<c r="{ref}"{st}><f>{escape(f)}</f><v>{v}</v></c>'
    if v is None or v == "": return f'<c r="{ref}"{st}/>' if s else ''
    if isinstance(v, (int, float)): return f'<c r="{ref}"{st}><v>{v}</v></c>'
    return f'<c r="{ref}"{st} t="inlineStr"><is><t xml:space="preserve">{escape(CTRL.sub("", str(v)))}</t></is></c>'


def row_xml(k, cells, ht=None):
    h = f' ht="{ht}" customHeight="1"' if ht else ''
    return f'<row r="{k}"{h}>' + ''.join(cells) + '</row>'


def status_de(o):
    r = MAN.get(o['raiz']) or ST.get(o['raiz']) or {'status': 'NÃO É CLIENTE'}
    if r['status'] == 'NÃO É CLIENTE': return r['status'], None, '', ''
    cli = re.sub(r'^\S+\s', '', r.get('cliente', ''))
    como = f"{VIA.get(r['via'], r['via'])}: {cli}" + (f" – {r['prova']}" if r['via'] not in ('CNPJ', 'RAZAO', 'CODIGO') and r.get('prova') else '')
    d = r.get('data')
    if isinstance(d, str) and d: d = dt.date.fromisoformat(d)
    return r['status'], d, r.get('vendedor', ''), como


def serial(d): return (d - dt.date(1899, 12, 30)).days if d else None


def aba_xml():
    out = []
    for k, (letra, val) in enumerate(CX.items(), start=1):
        out.append(row_xml(k, [cell(f'W{k}', letra.lower(), S_TAB1), cell(f'X{k}', val, S_TAB1)], 15))
    out.append(row_xml(5, [], 15))
    out.append(row_xml(6, [cell(f'{get_column_letter(j + 1)}6', h, S_HEAD) for j, (h, w) in enumerate(COLS)], 45))
    for i, o in enumerate(rows, start=1):
        k = 6 + i
        st, d, vend, como = status_de(o)
        letra = o['porte'][:1]
        cx = CX.get(letra, 0) * (o['lojas'] or 0)
        f = (f'IF(LEFT($V{k},1)="A",$X$1,IF(LEFT($V{k},1)="B",$X$2,IF(LEFT($V{k},1)="C",$X$3,IF(LEFT($V{k},1)="D",$X$4,0))))*N($E{k})')
        vals = [(i, S_NUM), (o['fant'], None), (o['raz'], None), (o['grupo'], None), (o['lojas'], S_NUM), (fmt(o['cnpj']), None), (o['conf'], None),
                ('', None), ('', None), (o['uf'], None), (o['mun'], None), (o['regiao'], None), (o['rk'], S_NUM), (o['rkr'], S_NUM), (o['sit'], None),
                (o['fonte'], None), (o['q'], S_NUM), (o['segm'], None), (o['origem'], None), (o['obs'], None), (o['sit'], None), (o['porte'], None),
                (st, None)]
        cs = [cell(f'{get_column_letter(j + 1)}{k}', v, s) for j, (v, s) in enumerate(vals)]
        cs.append(cell(f'X{k}', cx, S_NUM, f))
        cs.append(cell(f'Y{k}', serial(d), S_DATA))
        cs.append(cell(f'Z{k}', vend if st != 'NÃO É CLIENTE' else ''))
        cs.append(cell(f'AA{k}', como))
        out.append(row_xml(k, [c for c in cs if c]))
    ult = 6 + len(rows)
    cols = "<cols>" + "".join(f'<col min="{j + 1}" max="{j + 1}" width="{w}" customWidth="1"/>' for j, (h, w) in enumerate(COLS)) + "</cols>"
    return cols + "<sheetData>" + "".join(out) + "</sheetData>" + f'<autoFilter ref="A6:AA{ult}"/>', f"A1:AA{ult}"


SEGORD = ['PNEU', 'FMCG', 'FARMA', 'ELETRO', 'TEXTIL', 'AUTO', 'QUIM', 'CONST', 'OLEO']


def regiao_rows(k0):
    out = []; k = k0
    def lin(cells, ht=None):
        nonlocal k; out.append(row_xml(k, cells, ht)); k += 1
    lin([cell(f'A{k}', f'INDÚSTRIAS — acrescentado em {DATA}: aba FULLFILMIT_INDUSTRIA FINAL (região = UF da sede; ordem = porte pela receita 2025 da Exame e, sem ela, pelo capital social).')]); k += 1
    blocos = [('BRASIL', sorted(rows, key=lambda o: o['rk'])[:30])]
    for g in ['NORTE', 'NORDESTE', 'CENTRO-OESTE', 'SUDESTE', 'SUL']:
        blocos.append((g, sorted([o for o in rows if o['regiao'] == g], key=lambda o: o['rkr'])[:20]))
    cab = ["POSIÇÃO", "NOME FANTASIA", "SEGMENTO", "Nº DE UNIDADES", "UF SEDE", "CNPJ", "PORTE", "STATUS AUTOPEL"]
    for g, lst in blocos:
        tit = f"INDÚSTRIAS — {'MAIORES DO BRASIL (30)' if g == 'BRASIL' else 'MAIORES DA REGIÃO ' + g + ' (20)'}"
        lin([cell(f'A{k}', tit, S_HEAD)] + [cell(f'{get_column_letter(j)}{k}', '', S_HEAD) for j in range(2, 9)], 20)
        lin([cell(f'{get_column_letter(j + 1)}{k}', h, S_HEAD) for j, h in enumerate(cab)])
        for i, o in enumerate(lst, start=1):
            lin([cell(f'A{k}', i, S_NUM), cell(f'B{k}', o['fant']), cell(f'C{k}', o['segm']), cell(f'D{k}', o['lojas'], S_NUM), cell(f'E{k}', o['uf']),
                 cell(f'F{k}', fmt(o['cnpj'])), cell(f'G{k}', o['porte']), cell(f'H{k}', status_de(o)[0])])
        k += 1
    return out, k - 1


def resumo_rows(k0):
    out = []; k = k0
    def lin(vals, ht=None):
        nonlocal k
        out.append(row_xml(k, [cell(f'{get_column_letter(j + 1)}{k}', v, s) for j, (v, s) in enumerate(vals)], ht)); k += 1
    lin([(f'9. INDÚSTRIAS (acrescentado em {DATA})', S_SEC)], 20.1)
    lin([('Aba FULLFILMIT_INDUSTRIA FINAL: mesmas colunas das outras abas FINAL; blocos por região no fim da aba MAIORES POR REGIÃO. Uma linha por empresa (fabricante).', None)])
    lin([('Passo', S_HEAD), ('Como foi feito', S_HEAD)], 30)
    for a, b in [
        ('1. Segmentos pedidos', 'Pneus; bens de consumo não duráveis (FMCG: alimentos, bebidas, higiene, limpeza, cosméticos); farmacêutica e médica; eletroeletrônicos e linha branca; têxtil, calçados e vestuário; automotivo, autopeças e máquinas; química, plásticos, papel e embalagens; móveis e materiais de construção; óleo e lubrificantes (fabricantes e revendas). Fabricantes com selo PROCEL, CONPET ou etiqueta ENCE/PBE marcados na OBSERVAÇÃO.'),
        ('2. Receita Federal por CNAE', 'Varredura das empresas ativas com fábrica (CNAE principal da indústria de cada segmento; lubrificante: refino 1921, lubrificantes 1922 e atacado de lubrificantes 4681-8/05). Filtro contra erro de cadastro: porte "demais", sociedade (fora MEI/EIRELI) e capital social de R$ 5 mi ou mais.'),
        ('3. Corte por segmento e região', 'As maiores pelo capital social no Brasil (FMCG 60, química/papel 50, demais 40) + as 6 maiores de cada segmento em cada região (capital de R$ 20 mi ou mais). Pneus e lubrificantes: todas com capital de R$ 20 mi ou mais.'),
        ('4. Listas de mercado', 'Exame Melhores e Maiores 2026 (receita de 2025) nos setores industriais; associadas da ANFAVEA, da ELETROS e da ANIP; mercado de lubrificantes 2025 (Lubes em Foco); tabelas ENCE/PBE do Inmetro 2025; e multinacionais líderes que não publicam balanço (Unilever, P&G, Colgate, PepsiCo etc.).'),
        ('5. Fora da lista', 'Mineração, siderurgia, saneamento, geração de energia, exploração de petróleo, distribuidoras de combustível, tradings de grãos, usinas de açúcar e etanol a granel, agropecuária, construtoras, varejo, holdings e importadoras sem fábrica.'),
        ('6. CNPJ e unidades', 'CNPJ = matriz da empresa na Receita (todas conferidas: existem e estão ativas). Empresas do mesmo grupo com várias razões sociais viram uma linha (ex.: JBS + Seara; Tupy + MWM). Nº de unidades = estabelecimentos ativos de todas as raízes da linha (fábricas, CDs, filiais e lojas próprias).'),
        ('7. Porte', 'Receita 2025 da Exame: A ≥ R$ 1 bi, B R$ 300 mi a 1 bi, C R$ 20 a 300 mi, D < R$ 20 mi. Sem receita publicada: estimado pelo capital social (A ≥ R$ 500 mi, B R$ 100 a 500 mi, C R$ 10 a 100 mi, D < R$ 10 mi), com aviso na OBSERVAÇÃO.'),
        ('8. STATUS AUTOPEL', 'Mesmo motor das outras abas (16/11): CNPJ (raiz), razão social, grupo do cliente, nome fantasia, e-mail e produto personalizado, com as travas contra homônimo. Indústria só casa por nome com cliente de segmento corporativo, fabricante ou autopeças. ATIVO = compra há menos de 360 dias; INATIVO = comprou antes ou só tem cadastro; NÃO É CLIENTE = nada. Vínculos fracos revistos à mão.')]:
        lin([(a, S_TAB), (b, S_TAB)])
    k += 1
    lin([('Segmento', S_HEAD), ('Empresas', S_HEAD), ('ATIVO', S_HEAD), ('INATIVO', S_HEAD), ('NÃO É CLIENTE', S_HEAD), ('5 maiores (porte)', S_HEAD)], 30)
    for sg in SEGORD:
        lst = sorted([o for o in rows if o['seg'] == sg], key=lambda o: o['rk'])
        if not lst: continue
        c = collections.Counter(status_de(o)[0] for o in lst)
        lin([(lst[0]['segm'], S_TAB), (len(lst), S_NUM), (c['ATIVO'], S_NUM), (c['INATIVO'], S_NUM), (c['NÃO É CLIENTE'], S_NUM),
             (', '.join(o['fant'] for o in lst[:5]), S_TAB)])
    c = collections.Counter(status_de(o)[0] for o in rows)
    lin([('TOTAL', S_TAB), (len(rows), S_NUM), (c['ATIVO'], S_NUM), (c['INATIVO'], S_NUM), (c['NÃO É CLIENTE'], S_NUM), ('', S_TAB)])
    k += 1
    lin([('Fonte', S_HEAD), ('O que foi usado', S_HEAD), ('Endereço', S_HEAD)], 30)
    for a, b, e in [
        ('Receita Federal – dados abertos do CNPJ (out/2026)', 'Empresas industriais por CNAE, capital social, matriz, situação e unidades ativas.', 'https://arquivos.receitafederal.gov.br/'),
        ('Exame – Melhores e Maiores 2026', 'As 1.000 maiores por receita líquida de 2025, com setor e UF: receita e porte das indústrias.', 'https://exame.com/mm/maiores/'),
        ('ANFAVEA – associadas', 'Montadoras de autoveículos e máquinas agrícolas (28 associadas) + BYD e GWM (fábricas no Brasil, não associadas).', 'https://anfavea.com.br/site/associadas/'),
        ('ELETROS – associadas', 'Fabricantes de eletroeletrônicos e linha branca (97% da oferta do mercado brasileiro).', 'https://eletros.org.br/'),
        ('ANIP – Associação Nacional da Indústria de Pneumáticos', 'Fabricantes de pneus com fábrica no Brasil.', 'apresentação da ANIP na Câmara (set/2024) e imprensa'),
        ('Lubes em Foco – mercado de lubrificantes 2025', 'Iconic, Vibra, Blueway (Raízen), Moove, Petronas, Castrol, Quaker, Ultrax e outras.', 'https://portallubes.com.br/2026/05/the-brazilian-lubricant-market-in-2025/'),
        ('Inmetro – tabelas de eficiência energética (PBE/ENCE 2025)', 'Marcas com etiqueta ENCE e selos PROCEL/CONPET: aquecedores a gás, bombas, ar-condicionado, fogões, refrigeradores, lavadoras e veículos.', 'https://www.gov.br/inmetro/pt-br/assuntos/regulamentacao/avaliacao-da-conformidade/programa-brasileiro-de-etiquetagem/tabelas-de-eficiencia-energetica'),
        ('ANAMACO, ABRAMAT, ABCEM, ANR, Abrasel, ABRAS, APAS', 'Consultadas: são de varejo, construção e alimentação fora do lar (já cobertas pelas outras abas FINAL); não publicam lista de indústrias.', 'pesquisa Google')]:
        lin([(a, S_TAB), (b, S_TAB), (e, S_TAB)])
    k += 1
    lin([('• O nº de unidades de uma indústria inclui CDs, escritórios e lojas próprias (ex.: Hering, Veste). Para caixas/mês vale a mesma regra das outras abas (letra do porte × unidades).', None)])
    lin([('• Empresa sem receita publicada tem o porte estimado pelo capital social: conferir antes de usar o número em meta.', None)])
    return out, k - 1


def caminho(z, nome):
    wb = z.read('xl/workbook.xml').decode('utf8'); rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    rid = re.search(r'<sheet name="' + re.escape(escape(nome)) + r'" sheetId="\d+"(?: state="\w+")? r:id="(rId\d+)"', wb).group(1)
    t = re.search(r'<Relationship Id="' + rid + r'"[^>]*Target="([^"]+)"', rels) or re.search(r'Target="([^"]+)"[^>]*Id="' + rid + '"', rels)
    return 'xl/' + t.group(1).lstrip('/').replace('xl/', '')


def acrescenta(x, novas, ultima):
    ult = max(int(n) for n in re.findall(r'<row r="(\d+)"', x))
    assert int(re.search(r'<row r="(\d+)"', novas[0]).group(1)) > ult
    x = x.replace('</sheetData>', ''.join(novas) + '</sheetData>', 1)
    return re.sub(r'<dimension ref="([A-Z]+)1:([A-Z]+)\d+"/>', lambda m: f'<dimension ref="{m.group(1)}1:{m.group(2)}{ultima}"/>', x)


def main():
    z = zipfile.ZipFile(MAPA); novos = {}
    p = caminho(z, ABA); x = z.read(p).decode('utf8')
    assert '<sheetData/>' in x, 'a aba não está vazia'
    corpo, dim = aba_xml()
    x = x.replace('<sheetData/>', corpo, 1)
    x = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="{dim}"/>', x)
    sel = ' tabSelected="1"' if 'tabSelected="1"' in x else ''
    x = re.sub(r'<sheetViews>.*?</sheetViews>', f'<sheetViews><sheetView{sel} workbookViewId="0"><pane ySplit="6" topLeftCell="A7" activePane="bottomLeft" state="frozen"/>'
               '<selection pane="bottomLeft" activeCell="A7" sqref="A7"/></sheetView></sheetViews>', x, flags=re.S)
    novos[p] = x.encode('utf8'); print(ABA, p, len(rows))
    for aba, fn in [(REGIAO, regiao_rows), (RESUMO, resumo_rows)]:
        p = caminho(z, aba); x = z.read(p).decode('utf8')
        ult = max(int(n) for n in re.findall(r'<row r="(\d+)"', x))
        linhas, fim = fn(ult + 3)
        novos[p] = acrescenta(x, linhas, fim).encode('utf8'); print(aba, p, 'linhas', ult + 3, 'a', fim)
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zo:
        for it in z.infolist():
            zo.writestr(it, novos.pop(it.filename, None) or z.read(it.filename))
    assert not novos, novos.keys()
    print('->', OUT)


if __name__ == '__main__':
    main()
