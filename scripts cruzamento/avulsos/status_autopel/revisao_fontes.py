"""Revisão final (16/11/2026): confere o resultado (resultados.pkl) contra as FONTES BRUTAS, sem passar pelo clientes.pkl:
  1. status x data: ATIVO = compra há menos de 360 dias; INATIVO = mais antiga ou sem nota; NÃO É CLIENTE = sem data,
     sem vendedor; nenhuma data depois de ontem;
  2. vendedor: o do CADASTRO DE CLIENTES.RLT (código -> nome pelo ultimo fat); se o cadastro não tem pessoa (INATIVOS,
     CLIENTE SEM COMPRA, 000000...), o vendedor da nota;
  3. data: o DT_FAT do ultimo fat do cliente mostrado não pode ser mais novo que a data mostrada (fora o caminho por
     produto, em que a data é a do produto com o nome da rede);
  4. a mesma rede tem o mesmo status em todas as abas (cluster).
Saída: resumo na tela e os casos em <pasta>/revisao_fontes.txt (tem nomes, fora do git).
Uso: python revisao_fontes.py "<pasta de trabalho>" "<pasta 12.11.2026>" AAAA-MM-DD"""
import sys, os, pickle, datetime as dt, collections as C
import openpyxl

PASTA, FONTES, HOJE = sys.argv[1], sys.argv[2], dt.date.fromisoformat(sys.argv[3])
LIM = HOJE - dt.timedelta(days=360)
E = pickle.load(open(os.path.join(PASTA, 'entidades.pkl'), 'rb'))
R = pickle.load(open(os.path.join(PASTA, 'resultados.pkl'), 'rb'))
s = lambda x: str(x).strip() if x is not None else ''
NAO_PESSOA = ('INATIVOS', 'CLIENTE SEM COMPRA', 'RD', 'REPRESENTANTE', 'AUTOPEL - VENDA ENCERRADA')

def data(x):
    if isinstance(x, dt.datetime): return x.date()
    if isinstance(x, dt.date): return x
    t = s(x)
    try: return dt.date(int(t[6:10]), int(t[3:5]), int(t[0:2])) if len(t) >= 10 and t[2] == '/' else None
    except ValueError: return None

# cadastro (cabeçalho na linha 2) e ultimo fat (linha 1)
cad = {}
ws = openpyxl.load_workbook(os.path.join(FONTES, 'CADASTRO DE CLIENTES.RLT.xlsx'), read_only=True).worksheets[0]
it = ws.iter_rows(values_only=True); next(it); cab = [s(x) for x in next(it)]
iC, iL, iV = cab.index('Codigo'), cab.index('Loja'), cab.index('Vendedor')
for r in it:
    if r and s(r[iC]): cad[(s(r[iC]), s(r[iL]))] = s(r[iV])
ws = openpyxl.load_workbook(os.path.join(FONTES, 'ultimo fat.xlsx'), read_only=True).worksheets[0]
it = ws.iter_rows(values_only=True); cab = [s(x) for x in next(it)]
iC, iL = cab.index('COD_CLIENT'), cab.index('LOJA')
iVc = cab.index('Codigo'); iVn = iVc + 1; iD = cab.index('DT_FAT')
nome_v, ult = {}, {}
for r in it:
    if not r or not s(r[iC]): continue
    if s(r[iVc]): nome_v.setdefault(s(r[iVc]), s(r[iVn]))
    d = data(r[iD])
    k = (s(r[iC]), s(r[iL]))
    if d and (k not in ult or d > ult[k]): ult[k] = d
print('cadastro', len(cad), '| ultimo fat', len(ult), '| vendedores', len(nome_v))

GRAVA = lambda a: a.endswith('FINAL') and a != 'OUTROS FINAL' or a.startswith('I.R.T.T') or a == '"FARMACIA VF"'
prob = C.defaultdict(list)
for n, e in enumerate(E):
    a = e['aba']
    if a.startswith('SUPERMERCADOS (re') or a == 'OUTROS FINAL': continue
    r = R[n]; st, d = r['status'], r['data']
    rot = f"{a[:16]} L{e['linha']} {s(e['extra'].get('nome'))[:40]}"
    if st not in ('ATIVO', 'INATIVO', 'NÃO É CLIENTE'): prob['status inválido'].append(rot)
    if st == 'ATIVO' and not (d and d > LIM): prob['ATIVO sem compra nos 360 dias'].append(rot)
    if st == 'INATIVO' and d and d > LIM: prob['INATIVO com compra recente'].append(rot)
    if st == 'NÃO É CLIENTE' and (d or r['vendedor']): prob['NÃO É CLIENTE com data/vendedor'].append(rot)
    if d and d >= HOJE: prob['data de hoje ou futura'].append(f'{rot} {d}')
    if not GRAVA(a) or st == 'NÃO É CLIENTE': continue
    k = tuple(r['cliente'].split(' ')[0].split('/'))
    vc = cad.get(k)
    if vc is None:
        prob['cliente mostrado fora do cadastro'].append(f'{rot} {k}'); continue
    vn = nome_v.get(vc, '')
    pessoa = vn and not vn.isdigit() and not any(vn.upper().startswith(x) for x in NAO_PESSOA)
    if pessoa and r['vendedor'] != vn:
        prob['vendedor diferente do cadastro'].append(f"{rot} | arquivo '{r['vendedor']}' x cadastro '{vn}' ({vc})")
    if not pessoa and r.get('vend_fonte') == 'cadastro':
        prob['cadastro sem pessoa mas vendedor do cadastro'].append(f"{rot} | '{r['vendedor']}' ({vc} {vn})")
    if r['vendedor'] in ('', 'SEM VENDEDOR NO CADASTRO'): prob['sem vendedor'].append(rot)
    du = ult.get(k)
    if du and d and du > d and r['via'] != 'PRODUTO':
        prob['ultimo fat mais novo que a data mostrada'].append(f'{rot} | mostrado {d} x ultimo fat {du} ({k})')
# mesma rede, mesmo status
cl = C.defaultdict(set)
for n, e in enumerate(E):
    if not e['extra'].get('so_cnpj') and not e['aba'].startswith('SUPERMERCADOS (re') and e['aba'] != 'OUTROS FINAL':
        cl[e['cluster']].add(R[n]['status'])
div = [c for c, v in cl.items() if len(v) > 1]
# separadas de propósito (marca que não herda do grupo) ficam em clusters com dois status
prob['rede com 2 status entre abas (cluster)'] = [str(c) for c in div]
with open(os.path.join(PASTA, 'revisao_fontes.txt'), 'w', encoding='utf8') as f:
    for k, v in prob.items():
        f.write(f'== {k}: {len(v)}\n' + ''.join(f'   {x}\n' for x in v))
for k, v in prob.items(): print(f'{k}: {len(v)}')
