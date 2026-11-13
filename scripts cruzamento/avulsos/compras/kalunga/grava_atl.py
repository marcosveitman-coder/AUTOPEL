"""Grava a aba "ATLxKALUNGxML" (vazia) do "Relatorio Revenda atualizado.xlsx", só trocando o XML dessa aba.

    python grava_atl.py PLANILHA.xlsx cruzamento.tsv [SAIDA.xlsx]

Colunas: B, C, D e G da aba Geral exatamente como estão (cabeçalho e valores), depois Familia (MACRO), Familia (MICRO),
EAN, DESCRIÇÃO KALUNGA, produto_url (de cruza_kalunga.py) e duas colunas de apoio: Vínculo Kalunga (EXATO / SIMILAR)
e Como foi cruzado. Estilos reaproveitados da aba Geral (cabeçalho, texto e preço). A aba tem de estar vazia ou ter
sido gravada por este script (aí é refeita); com outro conteúdo, recusa.
Sem SAIDA, grava no próprio arquivo (fechado no Excel).
"""
import csv
import os
import re
import sys
import zipfile
from xml.sax.saxutils import escape

import openpyxl
from openpyxl.utils import get_column_letter

ABA = 'ATLxKALUNGxML'
CTRL = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')
LARG = [12, 60, 18, 12, 18, 34, 16, 70, 60, 12, 40]


def partes(z):
    wb = z.read('xl/workbook.xml').decode('utf-8')
    rel = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rm = {re.search(r'Id="([^"]+)"', e).group(1): re.search(r'Target="([^"]+)"', e).group(1)
          for e in re.findall(r'<Relationship [^>]+>', rel)}
    return {n: 'xl/' + rm[r].lstrip('/').replace('xl/', '', 1) for n, r in
            re.findall(r'<sheet name="([^"]+)" sheetId="\d+"(?: state="\w+")? r:id="(rId\d+)"', wb)}


def estilo(xml, ref):
    m = re.search(rf'<c r="{ref}"[^>]* s="(\d+)"', xml)
    return m.group(1) if m else '0'


def cel(ref, v, s):
    if v is None or v == '':
        return ''
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return f'<c r="{ref}" s="{s}"><v>{v}</v></c>'
    t = escape(CTRL.sub('', str(v)))
    return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{t}</t></is></c>'


def main(arq, tsv, saida=None):
    wb = openpyxl.load_workbook(arq, read_only=True, data_only=True)
    nome_g = next(n for n in wb.sheetnames if n.startswith('Geral (arq.') and n.endswith(')'))
    rows = wb[nome_g].iter_rows(values_only=True)
    next(rows)
    cab = next(rows)
    H = [str(x).strip() if x is not None else '' for x in cab]
    ix = [H.index(h) for h in ('Codigo', 'Descricao', 'Marca', 'Ult.Preco')]
    geral = [[r[i] for i in ix] for r in rows if r and r[ix[0]]]
    wb.close()
    cz = list(csv.reader(open(tsv, encoding='utf-8'), delimiter='\t'))[1:]
    assert len(cz) == len(geral), f'cruzamento com {len(cz)} linhas e Geral com {len(geral)}: refazer o cruzamento'
    for g, c in zip(geral, cz):
        assert str(g[0]).strip() == c[0], f'ordem diferente: {g[0]} x {c[0]}'

    z = zipfile.ZipFile(arq)
    p = partes(z)
    parte, xg = p[ABA], z.read(p[nome_g]).decode('utf-8')
    s_cab, s_txt, s_preco = estilo(xg, 'B2'), estilo(xg, 'B3'), estilo(xg, 'G3')
    x = z.read(parte).decode('utf-8')
    feito = re.search(r'<cols>.*?</cols><sheetData>.*?</sheetData><autoFilter[^>]*/>', x, re.S)
    if feito and 'Familia (MACRO)' in x:          # aba já gravada por este script: refaz
        x = x.replace(feito.group(0), '<sheetData/>', 1)
    assert '<sheetData/>' in x, f'a aba {ABA} tem outro conteúdo: não sobrescrevo'

    titulos = [cab[i] for i in ix] + ['Familia (MACRO)', 'Familia (MICRO)', 'EAN', 'DESCRIÇÃO KALUNGA', 'produto_url',
                                      'Vínculo Kalunga', 'Como foi cruzado']
    nc = len(titulos)
    out = [f'<row r="1" spans="1:{nc}" ht="30" customHeight="1">' +
           ''.join(cel(f'{get_column_letter(j + 1)}1', h, s_cab) for j, h in enumerate(titulos)) + '</row>']
    for n, (g, c) in enumerate(zip(geral, cz), 2):
        como = []
        if c[5]:
            como.append('família pela descrição' if c[6] == 'descrição' else f'família: {c[6]}')
        if c[10]:
            como.append(f'produto: {c[11]}' + (f' (nota {float(c[12]):.2f})'.replace('.', ',') if c[11].startswith('texto') else ''))
        vals = [(g[0], s_txt), (g[1], s_txt), (g[2], s_txt), (g[3], s_preco), (c[4], s_txt), (c[5], s_txt),
                (c[7], s_txt), (c[8], s_txt), (c[9], s_txt), (c[10], s_txt), ('; '.join(como), s_txt)]
        out.append(f'<row r="{n}" spans="1:{nc}">' +
                   ''.join(cel(f'{get_column_letter(j + 1)}{n}', v, s) for j, (v, s) in enumerate(vals)) + '</row>')
    last, nl = get_column_letter(nc), len(geral) + 1
    cols = '<cols>' + ''.join(f'<col min="{j + 1}" max="{j + 1}" width="{w}" customWidth="1"/>' for j, w in enumerate(LARG)) + '</cols>'
    x = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="A1:{last}{nl}"/>', x, count=1)
    x = re.sub(r'<sheetViews>.*?</sheetViews>',
               '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
               '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>', x, count=1, flags=re.S)
    x = x.replace('<sheetData/>', cols + '<sheetData>' + ''.join(out) + '</sheetData>' + f'<autoFilter ref="A1:{last}{nl}"/>', 1)
    destino = saida or arq
    tmp = destino + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for info in z.infolist():
            zo.writestr(info, x.encode('utf-8') if info.filename == parte else z.read(info.filename), compress_type=zipfile.ZIP_DEFLATED)
    z.close()
    os.replace(tmp, destino)
    print(f'gravado: {len(geral)} linhas em {ABA} ({parte})')


if __name__ == '__main__':
    a = sys.argv[1:]
    main(a[0], a[1], a[2] if len(a) > 2 else None)
