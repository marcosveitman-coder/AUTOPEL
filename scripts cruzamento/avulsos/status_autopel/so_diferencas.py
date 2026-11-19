"""Deixa no gravar.tsv (prepara_gravacao.py) só as células que mudam em relação ao MAPA atual (19/11/2026).
Assim a gravação mexe no mínimo de células e a conferência mostra exatamente o que mudou.
Saída: <saida.tsv> no mesmo formato (aba, coluna, linha, tipo, valor) + resumo por aba e coluna.
Uso: python so_diferencas.py "<gravar.tsv>" "<MAPA atual>" "<saida.tsv>" """
import sys, datetime as dt, collections as C
import openpyxl

ENT, MAPA, SAI = sys.argv[1:4]
linhas = [l.rstrip('\r\n').split('\t') for l in open(ENT, encoding='utf-8-sig') if l.strip()]
por_aba = C.defaultdict(list)
for x in linhas: por_aba[x[0]].append(x)
wb = openpyxl.load_workbook(MAPA, read_only=True, data_only=True)
out, cnt = [], C.Counter()


def atual(v):
    if isinstance(v, dt.datetime): return v.date().isoformat()
    if isinstance(v, dt.date): return v.isoformat()
    return '' if v is None else str(v).strip()


for aba, xs in por_aba.items():
    ws = wb[aba]
    alvo = {(int(x[2]), x[1]) for x in xs}
    rmax = max(r for r, _ in alvo)
    grade = {}
    for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=rmax, values_only=True), 1):
        for ci, v in enumerate(row, 1):
            c = openpyxl.utils.get_column_letter(ci)
            if (ri, c) in alvo: grade[(ri, c)] = v
    for x in xs:
        a = atual(grade.get((int(x[2]), x[1])))
        novo = x[4] if len(x) > 4 else ''
        if x[3] == 'h' or a != novo.strip():
            out.append(x); cnt[(aba, x[1])] += 1
wb.close()
with open(SAI, 'w', encoding='utf-8-sig', newline='') as f:
    for x in out: f.write('\t'.join(x) + '\r\n')
print(len(out), 'células mudam de', len(linhas))
for k, v in sorted(cnt.items()): print(' ', k, v)
