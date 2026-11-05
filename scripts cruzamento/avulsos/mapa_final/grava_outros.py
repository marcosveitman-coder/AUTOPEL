"""Cria a aba OUTROS FINAL no fim da pasta (so pelo XML; a pasta tem tabelas dinamicas), no mesmo modelo das abas
PET SHOP  FINAL e ESTACIONAMENTO FINAL, mais as colunas SEGMENTO e SUBSEGMENTO no fim. Le outros_rows.pkl.
Uso: python grava_outros.py <MAPA.xlsx> <saida.xlsx>"""
import os as _os, sys as _sys
_sys.argv[1:] = [_os.path.abspath(a) for a in _sys.argv[1:]]
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'dados'))
import re, sys, pickle, zipfile, collections
from xml.sax.saxutils import escape
from openpyxl.utils import get_column_letter
MAPA, OUT = sys.argv[1], sys.argv[2]
NOVA = 'OUTROS FINAL'
S_HEAD, S_TXT, S_NUM = 194, 0, 127
CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
REG = {}
for r, ufs in [('NORTE', 'AC AP AM PA RO RR TO'), ('NORDESTE', 'AL BA CE MA PB PE PI RN SE'), ('CENTRO-OESTE', 'DF GO MT MS'),
               ('SUDESTE', 'ES MG RJ SP'), ('SUL', 'PR RS SC')]:
    for u in ufs.split(): REG[u] = r
COLS = [("Nº", 6), ("NOME FANTASIA", 34), ("RAZÃO SOCIAL", 46), ("GRUPO ADMINISTRADOR", 30), ("Nº DE LOJAS", 11), ("CNPJ", 20),
        ("CONFERÊNCIA DO CNPJ", 30), ("CNPJ DE REFERÊNCIA (UNIDADE DA REDE)", 20), ("EMPRESA DE REFERÊNCIA", 50), ("UF SEDE", 7),
        ("MUNICÍPIO SEDE", 22), ("REGIÃO DA SEDE", 14), ("RANKING BRASIL (Nº DE LOJAS)", 11), ("RANKING NA REGIÃO", 11),
        ("SITUAÇÃO NA RECEITA", 12), ("FONTE DO Nº DE LOJAS", 36), ("LOJAS ATIVAS COM O NOME NA RECEITA (TODOS OS CNPJs)", 15),
        ("SEGMENTO", 26), ("ORIGEM", 24), ("OBSERVAÇÃO", 80), ("STATUS RECEITA", 16), ("PORTE", 30), ("STATUS AUTOPEL", 16),
        ("SEGMENTO (PESQUISA)", 26), ("SUBSEGMENTO", 30)]
rows = pickle.load(open('outros_rows.pkl', 'rb'))
def fmt(c): return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}" if len(c) == 14 else c
def cell(ref, v, s):
    if v is None or v == "": return ""
    st = f' s="{s}"' if s else ''
    if isinstance(v, (int, float)): return f'<c r="{ref}"{st}><v>{v}</v></c>'
    return f'<c r="{ref}"{st} t="inlineStr"><is><t xml:space="preserve">{escape(CTRL.sub("", str(v)))}</t></is></c>'
def row_xml(k, vals, ht=None):
    h = f' ht="{ht}" customHeight="1"' if ht else ''
    return f'<row r="{k}" spans="1:{len(vals)}"{h}>' + "".join(cell(f"{get_column_letter(j + 1)}{k}", v, s) for j, (v, s) in enumerate(vals)) + "</row>"
validos = [r for r in rows if r['lojas']]
for i, r in enumerate(sorted(validos, key=lambda r: -r['lojas']), start=1): r['rk'] = i
por = collections.defaultdict(list)
for r in validos:
    g = REG.get(r['uf_sede'])
    if g: por[g].append(r)
for g, lst in por.items():
    for i, r in enumerate(sorted(lst, key=lambda r: -r['lojas']), start=1): r['rkr'] = i
def obs_de(r): return '; '.join(x for x in [r['obs'], r.get('uf_sede_fonte') and 'UF SEDE: ' + r['uf_sede_fonte']] if x)
out = [row_xml(1, [(h, S_HEAD) for h, w in COLS], ht=45)]
for i, r in enumerate(rows, start=1):
    out.append(row_xml(i + 1, [(i, S_NUM), (r['fant'], S_TXT), (r['raz'], S_TXT), (r['grupo'], S_TXT), (r['lojas'], S_NUM), (fmt(r['cnpj']), S_TXT),
        (r['conf'], S_TXT), (fmt(r['ref']), S_TXT), (r['ref_desc'], S_TXT), (r['uf_sede'], S_TXT), (r['mun'] if r['uf_sede'] else '', S_TXT),
        (REG.get(r['uf_sede'], ''), S_TXT), (r.get('rk'), S_NUM), (r.get('rkr'), S_NUM), (r['sit'], S_TXT), (r['fonte'], S_TXT), (r['rede'], S_NUM),
        (r['segm'], S_TXT), (r['origem'], S_TXT), (obs_de(r), S_TXT), (r['sit'], S_TXT), ('', S_TXT), ('', S_TXT), (r['seg'], S_TXT), (r['sub'], S_TXT)]))
dim = f"A1:{get_column_letter(len(COLS))}{len(rows) + 1}"
corpo = ("<cols>" + "".join(f'<col min="{j + 1}" max="{j + 1}" width="{w}" customWidth="1"/>' for j, (h, w) in enumerate(COLS)) + "</cols>"
         + "<sheetData>" + "".join(out) + "</sheetData>" + f'<autoFilter ref="{dim}"/>')
VIEW = ('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>')
def caminho(z, nome):
    wb = z.read('xl/workbook.xml').decode('utf8'); rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    rid = re.search(r'<sheet name="' + re.escape(escape(nome)) + r'" sheetId="\d+"(?: state="\w+")? r:id="(rId\d+)"', wb).group(1)
    t = re.search(r'<Relationship Id="' + rid + r'"[^>]*Target="([^"]+)"', rels) or re.search(r'Target="([^"]+)"[^>]*Id="' + rid + '"', rels)
    return 'xl/' + t.group(1).lstrip('/').replace('xl/', '')
z = zipfile.ZipFile(MAPA); novos = {}
wb = z.read('xl/workbook.xml').decode('utf8')
assert f'<sheet name="{NOVA}"' not in wb, 'a aba já existe'
nums = [int(n) for n in re.findall(r'xl/worksheets/sheet(\d+)\.xml', ' '.join(z.namelist()))]
parte = f'xl/worksheets/sheet{max(nums) + 1}.xml'
rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
rid = 'rId' + str(max(int(n) for n in re.findall(r'Id="rId(\d+)"', rels)) + 1)
sid = max(int(n) for n in re.findall(r'sheetId="(\d+)"', wb)) + 1
modelo = z.read(caminho(z, 'ESTACIONAMENTO FINAL')).decode('utf8')
cab = modelo[:modelo.index('<dimension')]
cab = re.sub(r'xr:uid="\{[^}]+\}"', 'xr:uid="{3C6E1A52-7D4B-4E8F-A2C1-5B9D0E7F6A21}"', cab)
sx = (cab + f'<dimension ref="{dim}"/>{VIEW}<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>' + corpo
      + '<pageMargins left="0.511811024" right="0.511811024" top="0.78740157499999996" bottom="0.78740157499999996" header="0.31496062000000002" footer="0.31496062000000002"/></worksheet>')
novos[parte] = sx.encode('utf8')
novos['xl/workbook.xml'] = wb.replace('</sheets>', f'<sheet name="{NOVA}" sheetId="{sid}" r:id="{rid}"/></sheets>').encode('utf8')
novos['xl/_rels/workbook.xml.rels'] = rels.replace('</Relationships>', f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{parte.split("/")[-1]}"/></Relationships>').encode('utf8')
ct = z.read('[Content_Types].xml').decode('utf8')
novos['[Content_Types].xml'] = ct.replace('</Types>', f'<Override PartName="/{parte}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>').encode('utf8')
app = z.read('docProps/app.xml').decode('utf8')
m = re.search(r'<vt:lpstr>Planilhas</vt:lpstr></vt:variant><vt:variant><vt:i4>(\d+)</vt:i4>', app)
if m:
    n = int(m.group(1)); app = app.replace(m.group(0), m.group(0).replace(f'<vt:i4>{n}</vt:i4>', f'<vt:i4>{n + 1}</vt:i4>'))
app = re.sub(r'<TitlesOfParts><vt:vector size="(\d+)"', lambda q: f'<TitlesOfParts><vt:vector size="{int(q.group(1)) + 1}"', app)
app = app.replace('</vt:vector></TitlesOfParts>', f'<vt:lpstr>{NOVA}</vt:lpstr></vt:vector></TitlesOfParts>')
novos['docProps/app.xml'] = app.encode('utf8')
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zo:
    for it in z.infolist():
        zo.writestr(it, novos.pop(it.filename, None) or z.read(it.filename))
    zo.writestr(parte, novos.pop(parte))
assert not novos, novos.keys()
print('nova aba', NOVA, parte, len(rows), 'linhas ->', OUT)
