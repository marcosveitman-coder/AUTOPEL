"""Le as fotos diarias do Relatorio de Revenda (TOTVS, layout das colunas A-X da aba Geral) e grava um painel.

    python carrega_fotos.py PASTA_DAS_FOTOS SAIDA.pkl

Cada arquivo "Relatorio de Revenda DD-MM-AAAA.xlsx" vira as linhas de um dia. As colunas sao lidas pelo nome do
cabecalho (a ordem muda entre os arquivos). Saida: lista de (data, {codigo: {campo: valor}}).
"""
import datetime as dt
import pickle
import re
import sys
import warnings
from pathlib import Path

import openpyxl

warnings.filterwarnings('ignore')
CAMPOS = {'Comprador': 'comprador', 'Codigo': 'cod', 'Descricao': 'desc', 'Marca': 'marca', 'PontoPedido': 'pp',
          'Comprar.': 'comprar', 'Ult.Preco': 'preco', 'MedContr': 'med', 'MedVenda': 'medv', 'Vlr Estoque': 'vlr',
          'Fisico': 'fis', 'Reserva': 'res', 'Disponivel': 'disp', 'VendProt': 'vprot', 'Ruptura': 'rup',
          'Comprado': 'compr', 'DT_ENTREGA': 'dtent', 'Prc_Compra.': 'prc', 'ULT_COMPRA': 'ultc', 'PEDIDO': 'ped',
          'EstDias': 'estd', 'Grupo': 'grupo'}


def data_do_nome(p):
    m = re.search(r'(\d{2})-(\d{2})-(\d{4})', p.name)
    return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else None


def le(p):
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    ws = next((w for w in wb.worksheets if w.title.strip().lower().startswith('geral')), wb.worksheets[0])
    out, cab = {}, None
    for r in ws.iter_rows(values_only=True):
        if cab is None:
            if r and any(str(v).strip() == 'Codigo' for v in r if v is not None):
                cab = {CAMPOS[str(v).strip()]: j for j, v in enumerate(r) if v is not None and str(v).strip() in CAMPOS}
            continue
        cod = r[cab['cod']] if cab['cod'] < len(r) else None
        if not cod:
            continue
        out[str(cod).strip()] = {k: (r[j] if j < len(r) else None) for k, j in cab.items()}
    wb.close()
    return out


def main(pasta, saida):
    fotos = []
    for p in sorted(Path(pasta).glob('*.xlsx'), key=lambda p: data_do_nome(p) or dt.date.min):
        d = data_do_nome(p)
        if not d:
            continue
        linhas = le(p)
        fotos.append((d, linhas))
        print(d, len(linhas), p.name, file=sys.stderr)
    # um arquivo por dia (se houver dois do mesmo dia, fica o ultimo da lista)
    por_dia = {}
    for d, l in fotos:
        por_dia[d] = l
    pickle.dump(sorted(por_dia.items()), open(saida, 'wb'))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
