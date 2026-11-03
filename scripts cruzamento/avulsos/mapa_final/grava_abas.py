"""Grava as abas ALIMENTACAO FINAL, REDE DE LOJAS FINAL e CONSTRUCAO FINAL do MAPA MERCADO trocando so o XML dessas 3 abas
(a pasta tem tabelas dinamicas: nao salvar pelo openpyxl). Reaproveita estilos que ja existem (cabecalho 194, numero 127)."""
import os, re, sys, pickle, zipfile, shutil, json
import os as _os
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'dados'))
from xml.sax.saxutils import escape
from openpyxl.utils import get_column_letter
MAPA = sys.argv[1]
OUT = sys.argv[2]
S_HEAD, S_TXT, S_NUM = 194, 0, 127
CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
COLS = [("Nº", 6), ("NOME FANTASIA", 34), ("RAZÃO SOCIAL", 46), ("GRUPO ADMINISTRADOR", 30), ("Nº DE LOJAS", 11), ("CNPJ", 20),
        ("CONFERÊNCIA DO CNPJ", 27), ("UF SEDE", 7), ("MUNICÍPIO SEDE", 22), ("SITUAÇÃO NA RECEITA", 12), ("FONTE DO Nº DE LOJAS", 36),
        ("LOJAS ATIVAS COM O NOME NA RECEITA (TODOS OS CNPJs)", 15), ("SEGMENTO", 26), ("ORIGEM", 42), ("OBSERVAÇÃO", 80)]
ABAS = {'ALIM': 'ALIMENTACAO FINAL', 'LOJ': 'REDE DE LOJAS FINAL', 'CONS': 'CONSTRUÇÃO FINAL'}
CN = json.load(open('cnae_nomes.json', encoding='utf8'))
def fmt(c):
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}" if len(c) == 14 else c
def cell(ref, v, s):
    if v is None or v == "": return ""
    if isinstance(v, (int, float)): return f'<c r="{ref}" s="{s}"><v>{v}</v></c>'
    t = escape(CTRL.sub("", str(v)))
    return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{t}</t></is></c>'
def sheet_xml(rows):
    n = len(COLS); last = get_column_letter(n)
    out = [f'<row r="1" spans="1:{n}" ht="45" customHeight="1">' + "".join(cell(f"{get_column_letter(j+1)}1", h, S_HEAD) for j, (h, w) in enumerate(COLS)) + "</row>"]
    for i, r in enumerate(rows, start=1):
        k = i + 1
        segm = r.get('segm') or CN.get(r.get('cnae', ''), '')
        vals = [(i, S_NUM), (r['fant'], S_TXT), (r['raz'], S_TXT), (r['grupo'], S_TXT), (r['lojas'] if r['lojas'] else None, S_NUM), (fmt(r['cnpj']), S_TXT),
                (r['conf'], S_TXT), (r['uf'], S_TXT), (r['mun'], S_TXT), (r['sit'], S_TXT), (r['fonte'], S_TXT), (r['rede'] if r.get('rede') else None, S_NUM),
                (segm, S_TXT), (r['origem'], S_TXT), (r['obs'], S_TXT)]
        out.append(f'<row r="{k}" spans="1:{n}">' + "".join(cell(f"{get_column_letter(j+1)}{k}", v, s) for j, (v, s) in enumerate(vals)) + "</row>")
    cols = "<cols>" + "".join(f'<col min="{j+1}" max="{j+1}" width="{w}" customWidth="1"/>' for j, (h, w) in enumerate(COLS)) + "</cols>"
    return cols, "<sheetData>" + "".join(out) + "</sheetData>", f"A1:{last}{len(rows)+1}"
def caminho_aba(z, nome):
    wb = z.read('xl/workbook.xml').decode('utf8'); rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    m = re.search(r'<sheet name="' + re.escape(escape(nome)) + r'" sheetId="\d+"(?: state="\w+")? r:id="(rId\d+)"', wb)
    rid = m.group(1)
    t = re.search(r'<Relationship Id="' + rid + r'"[^>]*Target="([^"]+)"', rels) or re.search(r'Target="([^"]+)"[^>]*Id="' + rid + '"', rels)
    return 'xl/' + t.group(1).lstrip('/').replace('xl/', '')
def main():
    rows = pickle.load(open('final_rows.pkl', 'rb'))
    z = zipfile.ZipFile(MAPA)
    novos = {}
    for seg, aba in ABAS.items():
        p = caminho_aba(z, aba)
        x = z.read(p).decode('utf8')
        assert '<sheetData/>' in x, f'a aba {aba} não está vazia - não sobrescrevo'
        cols, data, dim = sheet_xml(rows[seg])
        x = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="{dim}"/>', x)
        x = re.sub(r'<sheetViews>.*?</sheetViews>', '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
                   '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>', x, flags=re.S)
        x = x.replace('<sheetData/>', cols + data + f'<autoFilter ref="{dim}"/>')
        novos[p] = x.encode('utf8')
        print(aba, p, len(rows[seg]), 'linhas')
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zo:
        for it in z.infolist():
            data = novos.get(it.filename, z.read(it.filename))
            zo.writestr(it, data)
    z.close()
    print('gravado', OUT)
if __name__ == '__main__':
    main()
