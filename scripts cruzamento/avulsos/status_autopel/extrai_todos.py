"""Roda o extrai_fat.py em todos os "faturamento*.xlsx" de uma pasta, 4 ao mesmo tempo (16/11/2026, ~12 min).
Os pickles saem com o nome do arquivo (espaços e parênteses viram "_"), na pasta indicada.
Uso: python extrai_todos.py "<pasta 12.11.2026>" "<pasta dos pickles>" """
import sys, os, glob, re
from concurrent.futures import ProcessPoolExecutor
import extrai_fat


def um(a):
    f, pk = a
    saida = os.path.join(pk, re.sub(r'[ ()]', '_', os.path.basename(f)) + '.pkl')
    extrai_fat.main(f, saida)
    return saida


if __name__ == '__main__':
    F, PK = sys.argv[1], sys.argv[2]
    with ProcessPoolExecutor(4) as ex:
        for s in ex.map(um, [(f, PK) for f in sorted(glob.glob(os.path.join(F, 'faturamento*.xlsx')))]):
            print('ok', s, flush=True)
