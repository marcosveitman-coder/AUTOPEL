"""Compara o STATUS AUTOPEL calculado (resultados.pkl) com o STATUS CRM de um arquivo de conferência do CRM (16/11/2026).
O arquivo do CRM tem uma aba por aba FINAL, com Nº, NOME FANTASIA, ..., STATUS CRM (H), ÚLTIMA COMPRA, COMO CASOU,
CLIENTE DA ÚLTIMA COMPRA. No CRM, INATIVANDO (291-365 dias) é ATIVO pela nossa regra (até 360 dias).
Saída (pasta de trabalho, fora do git, tem nomes): crm_cruz.txt (tabela) e crm_div.txt (cada divergência com as provas).
Uso: python crm_compara.py "<arquivo CRM.xlsx>" [pasta de trabalho]"""
import sys, os, pickle, datetime as dt, collections as C
import openpyxl
from motor import norm

PASTA = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.environ['TEMP'], 'status_1611')
E = pickle.load(open(os.path.join(PASTA, 'entidades.pkl'), 'rb'))
R = pickle.load(open(os.path.join(PASTA, 'resultados.pkl'), 'rb'))
por_aba = C.defaultdict(list)
for n, e in enumerate(E):
    if e['aba'].endswith('FINAL'): por_aba[e['aba']].append(n)

def acha(aba, num, nome):
    ns = por_aba.get(aba, [])
    alvo = norm(nome)
    iguais = [n for n in ns if norm(E[n]['extra'].get('nome', '')) == alvo]
    if not iguais: return None
    return min(iguais, key=lambda n: abs(E[n]['linha'] - 6 - (num or 0)))   # cabeçalho na linha 6

eq = lambda s: 'ATIVO' if s == 'INATIVANDO' else s
cruz = C.Counter(); div = []
wb = openpyxl.load_workbook(sys.argv[1], read_only=True)
for ws in wb.worksheets:
    it = ws.iter_rows(values_only=True)
    cab = [str(x or '').strip() for x in next(it)]
    ix = {h: i for i, h in enumerate(cab)}
    for r in it:
        if not r or r[ix['NOME FANTASIA']] is None: continue
        n = acha(ws.title, r[ix['Nº']], r[ix['NOME FANTASIA']])
        if n is None:
            div.append(f"SEM LINHA NO MAPA: {ws.title} Nº{r[ix['Nº']]} {r[ix['NOME FANTASIA']]}"); continue
        crm = str(r[ix['STATUS CRM']] or '').strip()
        nos = R[n]['status']
        cruz[(nos, crm)] += 1
        if nos != eq(crm):
            res = R[n]
            dcrm = r[ix['ÚLTIMA COMPRA']]
            linhas = [f"{ws.title[:14]} Nº{r[ix['Nº']]} {r[ix['NOME FANTASIA']]} ({r[ix['UF']]}) | NOSSO {nos} {res['data']} | CRM {crm} {dcrm.date() if isinstance(dcrm, dt.datetime) else dcrm}",
                      f"    nosso: {res.get('via','')} {res.get('cliente','')[:60]} – {res.get('prova','')[:90]}",
                      f"    CRM:   {r[ix['COMO CASOU']]} | {str(r[ix['CLIENTE DA ÚLTIMA COMPRA']] or '')[:110]}"]
            div.append('\n'.join(linhas))
with open(os.path.join(PASTA, 'crm_cruz.txt'), 'w', encoding='utf8') as f:
    for (a, b), q in sorted(cruz.items()): f.write(f'{a:14} | {b:14} | {q}\n')
with open(os.path.join(PASTA, 'crm_div.txt'), 'w', encoding='utf8') as f:
    f.write('\n'.join(div) + f'\ntotal {len(div)}\n')
print(open(os.path.join(PASTA, 'crm_cruz.txt'), encoding='utf8').read())
print('divergências:', len(div))
