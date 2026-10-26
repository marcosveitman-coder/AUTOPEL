"""Le os relatorios diarios do Protheus (layout EstoqueNecessidade: MES_01..03, fornecedor, pedido) e grava um painel.

    python carrega_protheus.py PASTA SAIDA.pkl

Saida: lista de (data, {(filial, codigo): {campo: valor}}). Colunas pelo nome do cabecalho.
"""
import datetime as dt
import pickle
import re
import sys
import warnings
from pathlib import Path

import openpyxl

warnings.filterwarnings('ignore')


def data_do_nome(p):
    m = re.search(r'(\d{2})-(\d{2})(?:-(\d{2,4}))?', p.name)
    if not m:
        return None
    a = m.group(3) or '2026'
    a = int(a) + 2000 if len(a) == 2 else int(a)
    return dt.date(a, int(m.group(2)), int(m.group(1)))


def main(pasta, saida):
    por_dia = {}
    for p in sorted(Path(pasta).glob('*.xlsx')):
        d = data_do_nome(p)
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        cab, linhas = None, {}
        for r in ws.iter_rows(values_only=True):
            if cab is None:
                if r and 'Codigo' in [str(c).strip() for c in r if c is not None]:
                    cab = [str(c).strip() if c is not None else None for c in r]
                continue
            reg = {k: v for k, v in zip(cab, r) if k}
            cod = str(reg.get('Codigo') or '').strip()
            if cod:
                linhas[(str(reg.get('Filial') or '').strip(), cod)] = reg
        wb.close()
        por_dia[d] = linhas          # dois arquivos no mesmo dia: fica o ultimo
        print(d, len(linhas), p.name, file=sys.stderr)
    pickle.dump(sorted(por_dia.items()), open(saida, 'wb'))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
