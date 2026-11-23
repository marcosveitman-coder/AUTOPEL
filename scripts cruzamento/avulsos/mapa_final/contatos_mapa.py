"""Contatos do LinkedIn no MAPA MERCADO V2 (23/11/2026): gera o TSV para o grava_contatos.ps1.
Fonte: planilha "Contatos LinkedIn - redes proprias (17-11-2026).xlsx", aba CONTATOS (uma linha por bandeira, blocos
Nome / LinkedIn / Cargo no LinkedIn por cargo). Destino:
  - Alessandra Dayan VF: colunas que o usuário montou em P7:AD8 (Gerente, Coordenador e Diretor de Compras, Diretor de
    Supply, CEO), na linha da bandeira (coluna EMPRESA igual);
  - ESTACIONAMENTO FINAL e POSTOS GAS & CONVENIENCIA FINAL: AB:AP (o ps1 copia o cabeçalho da VF para AB5:AP6), na linha
    da mesma rede (cluster da coleta.py). Várias linhas da VF na mesma linha FINAL: vale a do mesmo nome. Fica fora a
    linha FINAL que é outra unidade do grupo (ex.: postos de bandeira de supermercado), lista FORA.
Sócios do QSA e "outros nomes" não vão (não há coluna). O cargo "NÃO conferido" vai como está na planilha.
Uso: python contatos_mapa.py "<planilha contatos>" "<MAPA atual>" "<entidades.pkl da coleta>" "<saída.tsv>" """
import sys, re, pickle, unicodedata, collections as C
import openpyxl
from openpyxl.utils import get_column_letter as L, column_index_from_string as I

CONT, MAPA, ENT, SAI = sys.argv[1:5]
CARG = ['GERENTE DE COMPRAS', 'COORDENADOR DE COMPRAS', 'DIRETOR DE COMPRAS', 'DIRETOR DE SUPPLY', 'CEO']
ABAS = ['ESTACIONAMENTO FINAL', 'POSTOS GAS & CONVENIENCIA FINAL']
FORA = {'Postos Extra (GPA)', 'Postos Carrefour'}
N = lambda x: re.sub(r'[^A-Z0-9]', '', unicodedata.normalize('NFKD', str(x)).encode('ascii', 'ignore').decode().upper())

rows = list(openpyxl.load_workbook(CONT, read_only=True)['CONTATOS'].iter_rows(values_only=True))
col = {v: j for j, v in enumerate(rows[0]) if v}
dados = {r[1]: {c: r[col[c]:col[c] + 3] for c in CARG} for r in rows[2:] if r[1]}
wb = openpyxl.load_workbook(MAPA, read_only=True)
vf = list(wb['Alessandra Dayan VF'].iter_rows(values_only=True))
h = next(i for i, r in enumerate(vf) if 'EMPRESA' in [str(x or '').strip() for x in r])
iE = [str(x or '').strip() for x in vf[h]].index('EMPRESA')
emp = C.defaultdict(list)
for i, r in enumerate(vf[h + 1:], h + 2):
    if r[iE]: emp[r[iE]].append(i)
out = []


def put(aba, c0, row, d):
    for k, cargo in enumerate(CARG):
        for m, v in enumerate(d[cargo]):
            if v: out.append((aba, L(c0 + 3 * k + m), row, 's', re.sub(r'[\t\r\n]+', ' ', str(v))))


tem = {b: d for b, d in dados.items() if any(v[0] for v in d.values())}
for b, d in tem.items():
    for row in emp.get(b, []): put('Alessandra Dayan VF', I('P'), row, d)
E = pickle.load(open(ENT, 'rb'))
cl = C.defaultdict(list)
for e in E: cl[e['cluster']].append(e)
vfcl = {e['extra']['nome']: e['cluster'] for e in E if e['aba'] == 'Alessandra Dayan VF'}
for aba in ABAS:
    nomes = {i: r[1] for i, r in enumerate(wb[aba].iter_rows(values_only=True), 1)}
    por = C.defaultdict(list)
    for b in tem:
        for e in cl.get(vfcl.get(b), []):
            if e['aba'] == aba and e['extra']['nome'] not in FORA:
                assert nomes.get(e['linha']) == e['extra']['nome'], (aba, e['linha'])
                por[(e['linha'], e['extra']['nome'])].append(b)
    for (row, nm), bs in por.items():
        b = next((x for x in bs if N(x).startswith(N(nm)) or N(nm).startswith(N(x))), bs[0])
        put(aba, I('AB'), row, dados[b])
with open(SAI, 'w', encoding='utf-8-sig', newline='') as f:
    for x in out: f.write('\t'.join(map(str, x)) + '\r\n')
print(len(out), C.Counter(x[0] for x in out))
