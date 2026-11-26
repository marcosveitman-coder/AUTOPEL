"""Copia os contatos da Alessandra Dayan VF (P:AD) para ESTACIONAMENTO FINAL e POSTOS GAS & CONVENIENCIA FINAL (AB:AP) (26/11/2026).
Mesma ligação do contatos_mapa.py (cluster da coleta.py): a linha FINAL recebe os blocos da linha da VF da mesma rede; várias linhas
da VF na mesma linha FINAL: vale a do mesmo nome. Fica fora a linha FINAL que é outra unidade do grupo (lista FORA).
Só preenche o bloco do cargo (Nome/LinkedIn/Cargo) que está vazio na FINAL: o que já existe fica.
Uso: python contatos_finais_2611.py "<MAPA atual>" "<entidades.pkl da coleta>" "<saída.tsv>" """
import sys, re, pickle, unicodedata, collections as C
import openpyxl
from openpyxl.utils import get_column_letter as L, column_index_from_string as I

MAPA, ENT, SAI = sys.argv[1:4]
ABAS = ['ESTACIONAMENTO FINAL', 'POSTOS GAS & CONVENIENCIA FINAL']
FORA = {'Postos Extra (GPA)', 'Postos Carrefour'}
N = lambda x: re.sub(r'[^A-Z0-9]', '', unicodedata.normalize('NFKD', str(x)).encode('ascii', 'ignore').decode().upper())

wb = openpyxl.load_workbook(MAPA, read_only=True)
vf = list(wb['Alessandra Dayan VF'].iter_rows(max_row=600, max_col=30, values_only=True))
h = next(i for i, r in enumerate(vf) if 'EMPRESA' in [str(x or '').strip() for x in r])
iE = [str(x or '').strip() for x in vf[h]].index('EMPRESA')
p0 = I('P') - 1
assert vf[h][p0] == 'Nome', vf[h][p0]
blocos = {}  # empresa -> 5 blocos (nome, link, cargo) ou None
for r in vf[h + 1:]:
    if not r[iE]: continue
    bl = [tuple(r[p0 + 3 * k:p0 + 3 * k + 3]) for k in range(5)]
    bl = [b if b[0] else None for b in bl]
    if any(bl):
        old = blocos.get(r[iE], [None] * 5)
        blocos[r[iE]] = [o or b for o, b in zip(old, bl)]
E = pickle.load(open(ENT, 'rb'))
cl = C.defaultdict(list)
for e in E: cl[e['cluster']].append(e)
vfcl = {e['extra']['nome']: e['cluster'] for e in E if e['aba'] == 'Alessandra Dayan VF'}
out, n = [], C.Counter()
for aba in ABAS:
    fin = {i: r for i, r in enumerate(wb[aba].iter_rows(max_row=600, max_col=42, values_only=True), 1)}
    assert fin[6][I('AB') - 1] == 'Nome', (aba, fin[6][I('AB') - 1])
    por = C.defaultdict(list)
    for b in blocos:
        for e in cl.get(vfcl.get(b), []):
            if e['aba'] == aba and e['extra']['nome'] not in FORA:
                assert fin[e['linha']][1] == e['extra']['nome'], (aba, e['linha'], fin[e['linha']][1], e['extra']['nome'])
                por[(e['linha'], e['extra']['nome'])].append(b)
    for (row, nm), bs in sorted(por.items()):
        b = next((x for x in bs if N(x).startswith(N(nm)) or N(nm).startswith(N(x))), bs[0])
        for k, blk in enumerate(blocos[b]):
            if not blk: continue
            c0 = I('AB') + 3 * k
            if fin[row][c0 - 1]:
                n['mantidos'] += 1; continue
            n[aba] += 1
            for m, v in enumerate(blk):
                if v: out.append((aba, L(c0 + m), row, 's', re.sub(r'[\t\r\n]+', ' ', str(v))))
wb.close()
with open(SAI, 'w', encoding='utf-8-sig', newline='') as f:
    for x in out: f.write('\t'.join(map(str, x)) + '\r\n')
print(len(out), 'células; blocos:', dict(n))
