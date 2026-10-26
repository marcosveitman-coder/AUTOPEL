"""Confere a gravacao de grava_saude.ps1: arquivo de partida x arquivo gravado.

    python confere_saude.py PARTIDA.xlsx GRAVADO.xlsx PASTA_SAIDA   (PASTA_SAIDA tem saude.json e param_saude.csv)

- todas as abas, menos a analise e a param CLAUDE: valores e formulas iguais celula a celula;
- analise compras CLAUDE: linhas antigas iguais; tabelas novas iguais ao saude.json;
- param CLAUDE: colunas antigas iguais; colunas novas iguais ao param_saude.csv;
- partes do pacote (dinamicas, caches, graficos, segmentacoes) na mesma quantidade.
"""
import collections
import csv
import json
import sys
import zipfile

import openpyxl

ANALISE, PARAM = 'analise compras CLAUDE', 'param CLAUDE'


def celulas(arq, data_only):
    wb = openpyxl.load_workbook(arq, read_only=True, data_only=data_only)
    out = {}
    for ws in wb.worksheets:
        out[ws.title] = {(c.row, c.column): c.value for row in ws.iter_rows() for c in row
                         if getattr(c, 'value', None) is not None}
    wb.close()
    return out


def igual(a, b):
    if hasattr(a, 'text') or hasattr(b, 'text'):        # formula matricial: compara faixa e texto
        return (getattr(a, 'ref', None), getattr(a, 'text', a)) == (getattr(b, 'ref', None), getattr(b, 'text', b))
    if isinstance(a, float) or isinstance(b, float):
        try:
            return abs(float(a) - float(b)) <= 1e-6 * max(1.0, abs(float(a)))
        except (TypeError, ValueError):
            return False
    return a == b


def main(partida, gravado, pasta):
    js = json.load(open(f'{pasta}/saude.json', encoding='utf-8'))
    novas = list(csv.reader(open(f'{pasta}/param_saude.csv', encoding='utf-8-sig'), delimiter=';'))
    erros = collections.Counter()
    for data_only in (False, True):
        A, B = celulas(partida, data_only), celulas(gravado, data_only)
        if set(A) != set(B):
            print('abas diferentes:', set(A) ^ set(B)); erros['abas'] += 1
        for aba in A:
            a, b = A[aba], B.get(aba, {})
            if aba == ANALISE:
                ult = max(r for r, c in a)
                ks = {k for k in set(a) | set(b) if k[0] <= ult}
            elif aba == PARAM:
                ultc = max(c for r, c in a)
                ks = {k for k in set(a) | set(b) if k[1] <= ultc}
            else:
                ks = set(a) | set(b)
            dif = [k for k in ks if not igual(a.get(k), b.get(k))]
            if dif:
                erros[aba] += len(dif)
                print(f'[{"valor" if data_only else "formula"}] {aba}: {len(dif)} celulas diferentes, ex.:',
                      [(k, a.get(k), b.get(k)) for k in sorted(dif)[:3]])
        if data_only:
            # analise: as tabelas novas aparecem na ordem do json
            b = B[ANALISE]
            ult = max(r for r, c in A[ANALISE])
            linhas = {}
            for (r, c), v in b.items():
                if r > ult:
                    linhas.setdefault(r, {})[c] = v
            achadas = 0
            for nome in ('resumo', 'abc', 'comprador', 'grupo', 'top', 'triagem', 'genericos', 'clientes', 'donos', 'concentrados'):
                t = js[nome]
                r0 = next((r for r in sorted(linhas) if all(igual(linhas[r].get(2 + j), t[0][j]) for j in range(len(t[0])))), None)
                if r0 is None:
                    print('tabela nao achada na analise:', nome); erros['analise'] += 1; continue
                for i, lin in enumerate(t[1:], 1):
                    for j, v in enumerate(lin):
                        w = linhas.get(r0 + i, {}).get(2 + j)
                        if v in (None, '') and w in (None, ''):
                            continue
                        if not igual(w, v):
                            erros['analise'] += 1
                            if erros['analise'] <= 5:
                                print('analise', nome, r0 + i, 2 + j, v, w)
                achadas += 1
            print(f'analise: {achadas} de 10 tabelas conferidas')
            # param: colunas novas
            p = B[PARAM]
            ultc = max(c for r, c in A[PARAM])
            hdr = next(r for (r, c), v in p.items() if c == 1 and v == 'Codigo')
            por_cod = {p.get((r, 1)): r for (r, c) in p if c == 1 and r > hdr}
            for j, h in enumerate(novas[0][1:], 1):
                if p.get((hdr, ultc + j)) != h:
                    print('cabecalho da param diferente:', h, p.get((hdr, ultc + j))); erros['param'] += 1
            for lin in novas[1:]:
                r = por_cod.get(lin[0])
                for j, v in enumerate(lin[1:], 1):
                    w = p.get((r, ultc + j))
                    if v == '' and w in (None, ''):
                        continue
                    try:
                        ok = igual(float(v), w)
                    except ValueError:
                        ok = str(w) == v
                    if not ok:
                        erros['param'] += 1
                        if erros['param'] <= 5:
                            print('param', lin[0], novas[0][j], v, w)
            print(f'param: {len(novas) - 1} SKUs conferidos nas colunas novas')
    za, zb = zipfile.ZipFile(partida), zipfile.ZipFile(gravado)
    for tipo in ('pivotTable', 'pivotCacheDefinition', 'chart', 'slicer', 'drawing', 'table'):
        na = sum(1 for n in za.namelist() if f'/{tipo}' in n and n.endswith('.xml'))
        nb = sum(1 for n in zb.namelist() if f'/{tipo}' in n and n.endswith('.xml'))
        print(f'{tipo}: {na} -> {nb}')
        if tipo in ('pivotTable', 'pivotCacheDefinition', 'slicer', 'table') and na != nb:
            erros['pacote'] += 1
    print('ERROS:', dict(erros) if erros else 'nenhum')


if __name__ == '__main__':
    main(*sys.argv[1:4])
