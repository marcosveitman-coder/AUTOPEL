"""Prepara a gravação do STATUS AUTOPEL no MAPA (16/11/2026): gera gravar.tsv para o grava_status.ps1 (Excel/COM).
Por aba (as colunas são achadas pelo nome do cabeçalho, nunca pela letra):
  - 6 abas FINAL: STATUS AUTOPEL (W) + 3 colunas novas depois da última ("ÚLTIMA COMPRA AUTOPEL", "VENDEDOR (CADASTRO)",
    "COMO FOI ACHADO");
  - Alessandra Dayan (11 tabelas), Alessandra Dayan VF e as 11 (TOP 50): STATUS AUTOPEL;
  - I.R.T.T.__Por bandeira (FINAL): Atende?, Status, Última compra, Como casou, VENDEDOR CARTEIRA e VENDEDOR (CLIENTES ...);
  - "FARMACIA VF": DATA ÚLTIMA COMPRA, STATUS AUTOPEL (ATIVO / INATIVO / NÃO É CLIENTE), COMO FOI ACHADO, CLIENTE ENCONTRADO
    e VENDEDOR (as colunas I e J de nº de lojas são apagadas depois, no Excel).
Confere que o nome da linha no arquivo é o mesmo da coleta (o usuário pode ter mexido no arquivo).
Linhas do tsv: aba, coluna (letra), linha, tipo (s texto, d data, h cabeçalho novo), valor.
Uso: python prepara_gravacao.py "<pasta de trabalho>" "<MAPA atual>" "<saida.tsv>" """
import sys, os, pickle, datetime as dt, re
import openpyxl
from openpyxl.utils import get_column_letter as L

PASTA, MAPA, SAIDA = sys.argv[1:4]
E = pickle.load(open(os.path.join(PASTA, 'entidades.pkl'), 'rb'))
R = pickle.load(open(os.path.join(PASTA, 'resultados.pkl'), 'rb'))
FINAIS = ['ALIMENTACAO FINAL', 'REDE DE LOJAS FINAL', 'CONSTRUÇÃO FINAL', 'PET SHOP  FINAL', 'ESTACIONAMENTO FINAL',
          'POSTOS GAS & CONVENIENCIA FINAL']
VIA = {'CNPJ': 'CNPJ (raiz)', 'CODIGO': 'código de cliente', 'RAZAO': 'razão social igual', 'ASSOCIADA': 'associada da rede',
       'GRUPO MAPA': 'grupo do cadastro (mapa)', 'GRUPO ERP': 'grupo do ERP', 'GRUPO': 'grupo do cliente', 'EXATO': 'nome fantasia igual',
       'NOME': 'nome do cliente', 'EMAIL': 'e-mail', 'PRODUTO': 'produto personalizado'}
s = lambda x: str(x).strip() if x is not None else ''
out = []


def put(aba, col, lin, tipo, val):
    v = '' if val is None else (val.isoformat() if isinstance(val, dt.date) else re.sub(r'[\t\r\n]+', ' ', str(val)))
    out.append((aba, col, lin, tipo, v))


def como(r):
    if r['status'] == 'NÃO É CLIENTE': return ''
    cli = re.sub(r'^\S+\s', '', r['cliente'])           # tira o código do cliente
    return f"{VIA.get(r['via'], r['via'])}: {cli}" + (f" – {r['prova']}" if r['via'] not in ('CNPJ', 'RAZAO', 'CODIGO') else '')


def cabecalho(ws, nomes, ate=12):
    rows = list(ws.iter_rows(min_row=1, max_row=ate, values_only=True))
    for i, r in enumerate(rows, 1):
        txt = [s(x) for x in r]
        if all(any(t.startswith(n) for t in txt) for n in nomes):
            return i, {n: L(next(j for j, t in enumerate(txt, 1) if t.startswith(n))) for n in nomes}, len(r)
    raise SystemExit(f'cabeçalho não achado em {ws.title}: {nomes}')


wb = openpyxl.load_workbook(MAPA, read_only=True, data_only=True)
por_aba = {}
for i, e in enumerate(E): por_aba.setdefault(e['aba'], []).append(i)
nomes_mapa = {}
def confere(aba, ws_rows, col_nome, i):
    e = E[i]
    v = s(ws_rows.get((e['linha'], col_nome)))
    assert v == e['extra']['nome'] or (not v and not e['extra']['nome']), (aba, e['linha'], v, e['extra']['nome'])

# ---------------- FINAL
for aba in FINAIS:
    ws = wb[aba]
    h, cols, ncol = cabecalho(ws, ['NOME FANTASIA', 'STATUS AUTOPEL', 'Total de caixas'])
    NOVAS = ['ÚLTIMA COMPRA AUTOPEL', 'VENDEDOR (CADASTRO)', 'COMO FOI ACHADO']
    cab_txt = [s(x) for x in next(ws.iter_rows(min_row=h, max_row=h, values_only=True))]
    if all(t in cab_txt for t in NOVAS):      # 2ª gravação em diante: as colunas já existem
        nova = [L(cab_txt.index(t) + 1) for t in NOVAS]
    else:
        nova = [L(ncol + 1), L(ncol + 2), L(ncol + 3)]
        for c, t in zip(nova, NOVAS): put(aba, c, h, 'h', t)
    linhas = {r: s(v) for r, (v,) in enumerate(ws.iter_rows(min_row=1, min_col=2, max_col=2, values_only=True), 1)}
    for i in por_aba[aba]:
        e, r = E[i], R[i]
        assert linhas.get(e['linha']) == e['extra']['nome'], (aba, e['linha'], linhas.get(e['linha']), e['extra']['nome'])
        put(aba, cols['STATUS AUTOPEL'], e['linha'], 's', r['status'])
        put(aba, nova[0], e['linha'], 'd', r['data'])
        put(aba, nova[1], e['linha'], 's', r['vendedor'] if r['status'] != 'NÃO É CLIENTE' else '')
        put(aba, nova[2], e['linha'], 's', como(r))
# ---------------- Alessandra Dayan (coluna de status de cada tabela, vinda da coleta)
ws = wb['Alessandra Dayan']
grade = {(ri, ci): v for ri, row in enumerate(ws.iter_rows(values_only=True), 1) for ci, v in enumerate(row, 1)}
for i in por_aba['Alessandra Dayan']:
    e, r = E[i], R[i]
    ce = openpyxl.utils.column_index_from_string(e['extra']['col'])
    assert s(grade.get((e['linha'], ce))) == e['extra']['nome']
    assert s(grade.get((2, openpyxl.utils.column_index_from_string(e['extra']['col_status'])))) == 'STATUS AUTOPEL'
    put('Alessandra Dayan', e['extra']['col_status'], e['linha'], 's', r['status'])
# ---------------- Alessandra Dayan VF e TOP 50
for aba in [a for a in por_aba if a == 'Alessandra Dayan VF' or a.endswith('(TOP 50)')]:
    ws = wb[aba]
    h, cols, _ = cabecalho(ws, ['EMPRESA', 'STATUS AUTOPEL'], 6)
    linhas = {r: s(v) for r, (v,) in enumerate(ws.iter_rows(min_row=1, max_col=1, values_only=True), 1)}
    for i in por_aba[aba]:
        e, r = E[i], R[i]
        assert linhas.get(e['linha']) == e['extra']['nome'], (aba, e['linha'])
        put(aba, cols['STATUS AUTOPEL'], e['linha'], 's', r['status'])
# ---------------- IRTT
aba = 'I.R.T.T.__Por bandeira (FINAL)'
ws = wb[aba]
h, cols, _ = cabecalho(ws, ['Bandeira', 'Atende?', 'Status', 'Última compra', 'Como casou', 'VENDEDOR CARTEIRA', 'VENDEDOR (CLIENTES'])
linhas = {r: s(v) for r, (v,) in enumerate(ws.iter_rows(min_row=1, min_col=2, max_col=2, values_only=True), 1)}
for i in por_aba[aba]:
    e, r = E[i], R[i]
    assert linhas.get(e['linha']) == e['extra']['nome'], (aba, e['linha'])
    st = r['status']
    put(aba, cols['Atende?'], e['linha'], 's', 'Não atendido' if st == 'NÃO É CLIENTE' else 'Atendido')
    put(aba, cols['Status'], e['linha'], 's', {'ATIVO': 'Ativo', 'NÃO É CLIENTE': 'Não atendido'}.get(st, 'Inativo' if r['data'] else 'Cadastrado sem compra'))
    put(aba, cols['Última compra'], e['linha'], 'd', r['data'])
    put(aba, cols['Como casou'], e['linha'], 's', VIA.get(r['via'], r['via']) if st != 'NÃO É CLIENTE' else '')
    vd = r['vendedor'] if st != 'NÃO É CLIENTE' else ''
    put(aba, cols['VENDEDOR CARTEIRA'], e['linha'], 's', vd)
    put(aba, cols['VENDEDOR (CLIENTES'], e['linha'], 's', vd)
# ---------------- FARMACIA VF
aba = '"FARMACIA VF"'
ws = wb[aba]
h, cols, _ = cabecalho(ws, ['NOME FANTASIA', 'DATA ÚLTIMA COMPRA', 'STATUS AUTOPEL', 'COMO FOI ACHADO', 'CLIENTE ENCONTRADO', 'VENDEDOR'])
# 16/11: antes de apagar I e J eram O/P; depois, M/N. O grava_status.ps1 só apaga I:J se ainda forem as de Receita/pesquisa.
assert (cols['STATUS AUTOPEL'], cols['DATA ÚLTIMA COMPRA']) in (('P', 'O'), ('N', 'M')), cols
for i in por_aba[aba]:
    e, r = E[i], R[i]; st = r['status']
    put(aba, cols['DATA ÚLTIMA COMPRA'], e['linha'], 'd', r['data'])
    put(aba, cols['STATUS AUTOPEL'], e['linha'], 's', st)
    put(aba, cols['COMO FOI ACHADO'], e['linha'], 's', VIA.get(r['via'], r['via']) if st != 'NÃO É CLIENTE' else '')
    put(aba, cols['CLIENTE ENCONTRADO'], e['linha'], 's', re.sub(r'^\S+\s', '', r['cliente']) if st != 'NÃO É CLIENTE' else '')
    put(aba, cols['VENDEDOR'], e['linha'], 's', r['vendedor'] if st != 'NÃO É CLIENTE' else '')
wb.close()
with open(SAIDA, 'w', encoding='utf-8-sig', newline='') as f:
    for x in out: f.write('\t'.join(map(str, x)) + '\r\n')
import collections as C
print(len(out), 'células |', dict(C.Counter(x[0] for x in out)))
