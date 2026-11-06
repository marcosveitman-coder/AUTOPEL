"""Grava o MAPA MERCADO como .xlsm com a macro do duplo clique (06/11/2026).
Monta o projeto VBA a partir de vba_ThisWorkbook.txt e vba_DetalheMod.txt (../vba_projeto.py: MS-OVBA + MS-CFB, sem o
Excel), confere lendo de volta e poe no pacote: tipo macro do workbook, relacao, Override e codeName "ThisWorkbook".
Uso: python org_xlsm.py "<entrada.xlsx>" "<saida.xlsm>" """
import os, sys
AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, '..'))
import vba_projeto as vp


def modulos():
    m = []
    for nome, tipo in [('ThisWorkbook', 'doc'), ('DetalheMod', 'std')]:
        cod = open(os.path.join(AQUI, f'vba_{nome}.txt'), encoding='utf8').read()
        assert cod.startswith(f'Attribute VB_Name = "{nome}"'), nome
        cod.encode('cp1252')            # o VBA grava em cp1252: para aqui se houver caractere fora dele
        m.append((nome, tipo, cod))
    return m


if __name__ == '__main__':
    m = modulos()
    b = vp.projeto_vba(m)
    print('projeto VBA:', len(b), 'bytes |', vp.confere(b, m))
    vp.xlsm(sys.argv[1], sys.argv[2], b)
    print('gravado:', sys.argv[2], os.path.getsize(sys.argv[2]), 'bytes')
