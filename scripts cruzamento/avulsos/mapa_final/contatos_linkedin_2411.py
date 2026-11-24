"""Contatos da busca de pessoas no LinkedIn (24/11/2026) -> TSV para o grava_contatos.ps1.
Entrada: o arquivo de decisões (dados/contatos_2311/dec_2411.py: linha da VF -> cargo -> (nome, perfil, cargo no LinkedIn) e
POSTOS: linha da VF -> linha da POSTOS GAS & CONVENIENCIA FINAL), a fila (fila.json, para conferir o nome da linha) e o MAPA atual.
Regras: só preenche o bloco do cargo que está vazio (o que já foi gravado em 23/11 fica); VF em P:AD, FINAL em AB:AP.
Uso: python contatos_linkedin_2411.py "<dec_2411.py>" "<fila.json>" "<MAPA atual>" "<saída.tsv>" """
import sys, json, runpy, re
import openpyxl
from openpyxl.utils import get_column_letter as L, column_index_from_string as I

DEC, FILA, MAPA, SAI = sys.argv[1:5]
g = runpy.run_path(DEC)
D, POSTOS = g['D'], g['POSTOS']
ORD = ['GER', 'COO', 'DIRC', 'DIRS', 'CEO']
emp = {r['linha']: r['empresa'] for r in json.load(open(FILA, encoding='utf-8'))}
wb = openpyxl.load_workbook(MAPA, read_only=True)
out, n = [], {'VF': 0, 'POSTOS': 0, 'mantidos': 0}


def put(aba, c0, row, cells, dec):
    for k, cg in enumerate(ORD):
        if cg not in dec: continue
        if cells[c0 - 1 + 3 * k]:
            n['mantidos'] += 1; continue
        nome, slug, cargo = dec[cg]
        for m, v in enumerate([nome, 'https://www.linkedin.com/in/' + slug, cargo]):
            out.append((aba, L(c0 + 3 * k + m), row, 's', re.sub(r'[\t\r\n]+', ' ', v)))
        n['VF' if aba.startswith('Alessandra') else 'POSTOS'] += 1


vf = {i: r for i, r in enumerate(wb['Alessandra Dayan VF'].iter_rows(min_row=1, max_row=600, max_col=30, values_only=True), 1)}
assert vf[8][15] == 'Nome' and vf[7][15] == 'GERENTE DE COMPRAS'
for row, dec in D.items():
    assert vf[row][3] == emp[row], (row, vf[row][3], emp[row])
    put('Alessandra Dayan VF', I('P'), row, vf[row], dec)
po = {i: r for i, r in enumerate(wb['POSTOS GAS & CONVENIENCIA FINAL'].iter_rows(min_row=1, max_row=60, max_col=42, values_only=True), 1)}
assert po[6][27] == 'Nome' and po[5][27] == 'GERENTE DE COMPRAS'
for row, prow in POSTOS.items():
    put('POSTOS GAS & CONVENIENCIA FINAL', I('AB'), prow, po[prow], D[row])
wb.close()
with open(SAI, 'w', encoding='utf-8-sig', newline='') as f:
    for x in out: f.write('\t'.join(map(str, x)) + '\r\n')
print(len(out), 'células;', n)
