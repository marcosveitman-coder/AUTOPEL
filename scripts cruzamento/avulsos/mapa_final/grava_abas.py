"""2a rodada: regrava as abas ALIMENTACAO FINAL, REDE DE LOJAS FINAL e CONSTRUCAO FINAL (com CNPJ de referencia, regiao e
ranking) e cria a aba MAIORES POR REGIAO no fim da pasta, so pelo XML (a pasta tem tabelas dinamicas).
Uso: python grava_abas.py <MAPA.xlsx> <saida.xlsx>"""
import os as _os
import sys as _sys
_sys.argv[1:] = [_os.path.abspath(a) for a in _sys.argv[1:]]  # caminhos da planilha, antes do chdir
_AQUI = _os.path.dirname(_os.path.abspath(__file__))
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_AQUI, 'dados'))
import re, sys, pickle, zipfile, collections
from xml.sax.saxutils import escape
from openpyxl.utils import get_column_letter
import json
MAPA, OUT = sys.argv[1], sys.argv[2]
S_HEAD, S_TXT, S_NUM = 194, 0, 127
CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
CN = json.load(open('cnae_nomes.json', encoding='utf8'))
REG = {}
for r, ufs in [('NORTE', 'AC AP AM PA RO RR TO'), ('NORDESTE', 'AL BA CE MA PB PE PI RN SE'), ('CENTRO-OESTE', 'DF GO MT MS'),
               ('SUDESTE', 'ES MG RJ SP'), ('SUL', 'PR RS SC')]:
    for u in ufs.split(): REG[u] = r
COLS = [("Nº", 6), ("NOME FANTASIA", 34), ("RAZÃO SOCIAL", 46), ("GRUPO ADMINISTRADOR", 30), ("Nº DE LOJAS", 11), ("CNPJ", 20),
        ("CONFERÊNCIA DO CNPJ", 30), ("CNPJ DE REFERÊNCIA (UNIDADE DA REDE)", 20), ("EMPRESA DE REFERÊNCIA", 50), ("UF SEDE", 7),
        ("MUNICÍPIO SEDE", 22), ("REGIÃO DA SEDE", 14), ("RANKING BRASIL (Nº DE LOJAS)", 11), ("RANKING NA REGIÃO", 11),
        ("SITUAÇÃO NA RECEITA", 12), ("FONTE DO Nº DE LOJAS", 36), ("LOJAS ATIVAS COM O NOME NA RECEITA (TODOS OS CNPJs)", 15),
        ("SEGMENTO", 26), ("ORIGEM", 42), ("OBSERVAÇÃO", 80)]
ABAS = {'ALIM': 'ALIMENTACAO FINAL', 'LOJ': 'REDE DE LOJAS FINAL', 'CONS': 'CONSTRUÇÃO FINAL', 'POST': 'POSTOS GAS & CONVENIENCIA FINAL'}
NOVA = 'MAIORES POR REGIÃO'
RESUMO = 'R_E_S_U_M_O'
DATA = '03/11/2026'
def fmt(c):
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}" if len(c) == 14 else c
def cell(ref, v, s):
    if v is None or v == "": return ""
    if isinstance(v, (int, float)): return f'<c r="{ref}" s="{s}"><v>{v}</v></c>'
    t = escape(CTRL.sub("", str(v)))
    return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{t}</t></is></c>'
def row_xml(k, vals, ht=None):
    n = len(vals)
    h = f' ht="{ht}" customHeight="1"' if ht else ''
    return f'<row r="{k}" spans="1:{max(n,1)}"{h}>' + "".join(cell(f"{get_column_letter(j+1)}{k}", v, s) for j, (v, s) in enumerate(vals)) + "</row>"
def uf_de(r):
    return r.get('uf_sede') or ''
def obs_de(r):
    return '; '.join(x for x in [r['obs'], r.get('uf_sede_fonte') and 'UF SEDE: ' + r['uf_sede_fonte']] if x)
def rede_de(r):
    # contagem pelo nome muito acima do nº de lojas = homônimos (nome comum), não vale como conferência
    v = r.get('rede')
    return None if not v or (r['lojas'] and v > 3 * r['lojas']) else v
rows = pickle.load(open('final_rows.pkl', 'rb'))
import os as _o
if _o.path.exists('postos_rows.pkl'): rows['POST'] = pickle.load(open('postos_rows.pkl', 'rb'))  # aba de postos e conveniencia (build_postos.py)
for seg, rs in rows.items():
    validos = [r for r in rs if r['lojas'] and not r['conf'].startswith('REDE NÃO')]
    for i, r in enumerate(sorted(validos, key=lambda r: -r['lojas']), start=1): r['rk'] = i
    por = collections.defaultdict(list)
    for r in validos:
        g = REG.get(uf_de(r))
        r['regiao'] = g or ''
        if g: por[g].append(r)
    for g, lst in por.items():
        for i, r in enumerate(sorted(lst, key=lambda r: -r['lojas']), start=1): r['rkr'] = i
def aba_xml(rs):
    n = len(COLS)
    out = [row_xml(1, [(h, S_HEAD) for h, w in COLS], ht=45)]
    for i, r in enumerate(rs, start=1):
        u = uf_de(r)
        vals = [(i, S_NUM), (r['fant'], S_TXT), (r['raz'], S_TXT), (r['grupo'], S_TXT), (r['lojas'] or None, S_NUM), (fmt(r['cnpj'] or r.get('ref', '')), S_TXT),
                (r['conf'], S_TXT), (fmt(r.get('ref', '')), S_TXT), (r.get('ref_desc', ''), S_TXT), (uf_de(r), S_TXT), (r['mun'], S_TXT),
                (REG.get(u, ''), S_TXT), (r.get('rk'), S_NUM), (r.get('rkr'), S_NUM), (r['sit'], S_TXT), (r['fonte'], S_TXT),
                (rede_de(r), S_NUM), (r.get('segm') or CN.get(r.get('cnae', ''), ''), S_TXT), (r['origem'], S_TXT), (obs_de(r), S_TXT)]
        out.append(row_xml(i + 1, vals))
    last = get_column_letter(n)
    cols = "<cols>" + "".join(f'<col min="{j+1}" max="{j+1}" width="{w}" customWidth="1"/>' for j, (h, w) in enumerate(COLS)) + "</cols>"
    dim = f"A1:{last}{len(rs)+1}"
    return cols + "<sheetData>" + "".join(out) + "</sheetData>" + f'<autoFilter ref="{dim}"/>', dim
SEGN = {'ALIM': 'ALIMENTAÇÃO', 'LOJ': 'REDE DE LOJAS', 'CONS': 'CONSTRUÇÃO', 'POST': 'POSTOS E CONVENIÊNCIA'}
def regiao_xml():
    C = [("POSIÇÃO", 9), ("NOME FANTASIA", 34), ("GRUPO ADMINISTRADOR", 30), ("Nº DE LOJAS", 11), ("UF SEDE", 7), ("CNPJ", 20),
         ("CONFERÊNCIA DO CNPJ", 30), ("FONTE DO Nº DE LOJAS", 36)]
    out = []; k = 1
    out.append(row_xml(k, [("MAIORES REDES DO BRASIL E POR REGIÃO (região = UF da sede da empresa; ordem = nº de lojas). "
                             "Fonte: abas ALIMENTACAO FINAL, REDE DE LOJAS FINAL e CONSTRUÇÃO FINAL (03/11/2026).", S_TXT)])); k += 2
    for seg in [s for s in ['ALIM', 'LOJ', 'CONS', 'POST'] if s in rows]:
        validos = [r for r in rows[seg] if r['lojas'] and not r['conf'].startswith('REDE NÃO')]
        blocos = [('BRASIL', sorted(validos, key=lambda r: -r['lojas'])[:30])]
        for g in ['NORTE', 'NORDESTE', 'CENTRO-OESTE', 'SUDESTE', 'SUL']:
            blocos.append((g, sorted([r for r in validos if REG.get(uf_de(r)) == g], key=lambda r: -r['lojas'])[:20]))
        for g, lst in blocos:
            titulo = f"{SEGN[seg]} — {'MAIORES DO BRASIL (30)' if g == 'BRASIL' else 'MAIORES DA REGIÃO ' + g + ' (20)'}"
            out.append(row_xml(k, [(titulo, S_HEAD)] + [("", S_HEAD)] * (len(C) - 1), ht=20)); k += 1
            out.append(row_xml(k, [(h, S_HEAD) for h, w in C])); k += 1
            for i, r in enumerate(lst, start=1):
                out.append(row_xml(k, [(i, S_NUM), (r['fant'], S_TXT), (r['grupo'], S_TXT), (r['lojas'], S_NUM), (uf_de(r), S_TXT),
                                       (fmt(r['cnpj'] or r.get('ref', '')), S_TXT), (r['conf'], S_TXT), (r['fonte'], S_TXT)])); k += 1
            k += 1
    cols = "<cols>" + "".join(f'<col min="{j+1}" max="{j+1}" width="{w}" customWidth="1"/>' for j, (h, w) in enumerate(C)) + "</cols>"
    if _o.path.exists('anp_bandeiras.pkl'):  # postos por bandeira e regiao (ANP)
        A = pickle.load(open('anp_bandeiras.pkl', 'rb'))
        RG = ['NORTE', 'NORDESTE', 'CENTRO-OESTE', 'SUDESTE', 'SUL']
        out.append(row_xml(k, [("POSTOS E CONVENIÊNCIA — POSTOS POR BANDEIRA E REGIÃO (ANP, Anuário 2026, tabela 3.17, dados de 2025)", S_HEAD)] + [("", S_HEAD)] * (len(C) - 1), ht=20)); k += 1
        out.append(row_xml(k, [(h, S_HEAD) for h in ['BANDEIRA', 'NORTE', 'NORDESTE', 'CENTRO-OESTE', 'SUDESTE', 'SUL', 'BRASIL']])); k += 1
        NB = {'VIBRA': 'Vibra (Petrobras)', 'IPIRANGA': 'Ipiranga', 'RAIZEN': 'Raízen (Shell)', 'ALESAT': 'Alesat (Ale)', 'OUTRAS': 'Outras bandeiras', 'BANDEIRA BRANCA': 'Bandeira branca (sem bandeira)'}
        for b in ['VIBRA', 'IPIRANGA', 'RAIZEN', 'ALESAT', 'OUTRAS', 'BANDEIRA BRANCA']:
            v = A.get(b, {})
            out.append(row_xml(k, [(NB[b], S_TXT)] + [(v.get(x, 0), S_NUM) for x in RG] + [(sum(v.values()), S_NUM)])); k += 1
        out.append(row_xml(k, [('TOTAL', S_TXT)] + [(sum(A[b].get(x, 0) for b in A), S_NUM) for x in RG] + [(sum(sum(v.values()) for v in A.values()), S_NUM)])); k += 1
    return cols + "<sheetData>" + "".join(out) + "</sheetData>", f"A1:H{k}"
def caminho(z, nome):
    wb = z.read('xl/workbook.xml').decode('utf8'); rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    rid = re.search(r'<sheet name="' + re.escape(escape(nome)) + r'" sheetId="\d+"(?: state="\w+")? r:id="(rId\d+)"', wb).group(1)
    t = re.search(r'<Relationship Id="' + rid + r'"[^>]*Target="([^"]+)"', rels) or re.search(r'Target="([^"]+)"[^>]*Id="' + rid + '"', rels)
    return 'xl/' + t.group(1).lstrip('/').replace('xl/', '')
VIEW = ('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>')
def main():
    z = zipfile.ZipFile(MAPA); novos = {}
    for seg, aba in ABAS.items():
        p = caminho(z, aba); x = z.read(p).decode('utf8')
        assert '<sheetData/>' in x or re.search(r'<c r="B1"', x), f'a aba {aba} tem outro conteúdo'
        corpo, dim = aba_xml(rows[seg])
        x = re.sub(r'(<cols>.*?</cols>)?<sheetData(/>|>.*?</sheetData>)(<autoFilter[^>]*/>)?', lambda m: corpo, x, count=1, flags=re.S)
        x = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="{dim}"/>', x)
        x = re.sub(r'<sheetViews>.*?</sheetViews>', VIEW, x, flags=re.S)
        novos[p] = x.encode('utf8'); print(aba, p, len(rows[seg]))
    wb = z.read('xl/workbook.xml').decode('utf8')
    if f'<sheet name="{RESUMO}"' in wb:  # aba criada pelo usuario para o racional do estudo
        import resumo
        p = caminho(z, RESUMO); x = z.read(p).decode('utf8')
        corpo, dim = resumo.resumo_xml(rows, row_xml, get_column_letter, DATA)
        x = re.sub(r'(<cols>.*?</cols>)?<sheetData(/>|>.*?</sheetData>)', lambda m: corpo, x, count=1, flags=re.S)
        novos[p] = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="{dim}"/>', x).encode('utf8'); print(RESUMO, p)
    corpo, dim = regiao_xml()
    if f'<sheet name="{escape(NOVA)}"' in wb:  # a aba ja existe (atualizacao): troca so o conteudo dela
        p = caminho(z, NOVA); x = z.read(p).decode('utf8')
        x = re.sub(r'(<cols>.*?</cols>)?<sheetData(/>|>.*?</sheetData>)', lambda m: corpo, x, count=1, flags=re.S)
        novos[p] = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="{dim}"/>', x).encode('utf8')
        with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zo:
            for it in z.infolist():
                zo.writestr(it, novos.pop(it.filename, None) or z.read(it.filename))
        assert not novos, novos.keys()
        print('aba', NOVA, 'regravada', p, '->', OUT)
        return
    nums = [int(n) for n in re.findall(r'xl/worksheets/sheet(\d+)\.xml', ' '.join(z.namelist()))]
    parte = f'xl/worksheets/sheet{max(nums)+1}.xml'
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    rid = 'rId' + str(max(int(n) for n in re.findall(r'Id="rId(\d+)"', rels)) + 1)
    sid = max(int(n) for n in re.findall(r'sheetId="(\d+)"', wb)) + 1
    modelo = z.read(caminho(z, 'CONSTRUÇÃO FINAL')).decode('utf8')
    cab = modelo[:modelo.index('<dimension')]
    cab = re.sub(r'xr:uid="\{[^}]+\}"', 'xr:uid="{7A1C2E30-5B4D-4F60-9C11-3D2A6B7E8F90}"', cab)
    sx = (cab + f'<dimension ref="{dim}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>'
          + corpo + '<pageMargins left="0.511811024" right="0.511811024" top="0.78740157499999996" bottom="0.78740157499999996" header="0.31496062000000002" footer="0.31496062000000002"/></worksheet>')
    novos[parte] = sx.encode('utf8')
    novos['xl/workbook.xml'] = wb.replace('</sheets>', f'<sheet name="{escape(NOVA)}" sheetId="{sid}" r:id="{rid}"/></sheets>').encode('utf8')
    novos['xl/_rels/workbook.xml.rels'] = rels.replace('</Relationships>', f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{parte.split("/")[-1]}"/></Relationships>').encode('utf8')
    ct = z.read('[Content_Types].xml').decode('utf8')
    novos['[Content_Types].xml'] = ct.replace('</Types>', f'<Override PartName="/{parte}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>').encode('utf8')
    app = z.read('docProps/app.xml').decode('utf8')
    m = re.search(r'<vt:lpstr>Planilhas</vt:lpstr></vt:variant><vt:variant><vt:i4>(\d+)</vt:i4>', app)
    n = int(m.group(1))
    app = app.replace(m.group(0), m.group(0).replace(f'<vt:i4>{n}</vt:i4>', f'<vt:i4>{n+1}</vt:i4>'))
    app = re.sub(r'<TitlesOfParts><vt:vector size="(\d+)"', lambda q: f'<TitlesOfParts><vt:vector size="{int(q.group(1))+1}"', app)
    app = app.replace('</vt:vector></TitlesOfParts>', f'<vt:lpstr>{escape(NOVA)}</vt:lpstr></vt:vector></TitlesOfParts>')
    novos['docProps/app.xml'] = app.encode('utf8')
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zo:
        for it in z.infolist():
            zo.writestr(it, novos.pop(it.filename, None) or z.read(it.filename))
        zo.writestr(parte, novos.pop(parte))
    assert not novos, novos.keys()
    print('nova aba', NOVA, parte, rid, sid, '->', OUT)
if __name__ == '__main__':
    main()
