"""Candidatos a CNPJ (franqueadora / matriz) para as redes ainda sem CNPJ.
F = estabelecimento com o nome da rede e CNAE de franquia/holding/administracao (franqueadora)
R = razao social '<nucleo> FRANCHISING/FRANQUIAS...' ou com o nome e capital alto
I = industria/atacado com o nome (dona da marca)
A = estabelecimento mais antigo com o nome (provavel fundadora)
O = raiz com mais lojas ativas com o nome (operadora)"""
import os as _os
_AQUI = _os.path.dirname(_os.path.abspath(__file__))
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_AQUI, 'dados'))
import pickle, json, collections, re, sys
D = pickle.load(open('r4.pkl', 'rb')); Q = json.load(open('q4.json', encoding='utf8'))
emp, est = D['emp'], D['est']
by_raiz = collections.defaultdict(list)
for i, r in enumerate(est): by_raiz[r[0]].append(i)
fh = collections.defaultdict(set)
for i, t in D['hits_fant']: fh[t].add(i)
rh = collections.defaultdict(set)
for r, t in D['hits_raz']: rh[t].add(r)
FRQ = ('7740300', '8299799', '7020400', '8211300', '6462000', '6463800', '7490104', '8219999', '7319002', '7319003', '6461100')
def cap(r):
    try: return float(emp.get(r, ('', '', '0', ''))[2].replace(',', '.'))
    except: return 0.0
def lin(tag, x, extra=''):
    r = x[0]
    return f"  {tag} {x[0]}{x[1]}{x[2]} {x[10]} {x[8]} ini{x[7][:4]} | {emp.get(r, ('?',))[0][:44]} | cap{cap(r)/1e6:.1f}M | {x[4][:22]} {extra}"
out = open('v4.txt', 'w', encoding='utf8')
for q in Q:
    full = [t for t in q['terms'] if not re.search(r' (FRANCHISING|FRANQUIAS|FRANQUEADORA|FRANQUIA|PARTICIPACOES|HOLDING)$', t)]
    idx = set().union(*[fh.get(t, set()) for t in full])
    at = [est[i] for i in idx if est[i][5] == '02']
    rz_f = collections.Counter(x[0] for x in at)
    out.write(f"## {q['id']} [{q['st']}] {len(at)} lojas ativas com o nome em {len(rz_f)} CNPJs\n")
    seen = set()
    F = sorted([x for x in at if x[8] in FRQ], key=lambda x: -cap(x[0]))
    for x in F[:3]:
        out.write(lin('F', x) + '\n'); seen.add(x[0])
    rr = set().union(*[rh.get(t, set()) for t in q['terms']])
    R = []
    for r in rr:
        rows = [est[i] for i in by_raiz.get(r, [])]
        m = [x for x in rows if x[3] == '1' and x[5] == '02']
        if not m: continue
        rz = emp.get(r, ('',))[0]
        if re.search(r'FRANCHI|FRANQU', rz) or m[0][8] in FRQ: R.append(m[0])
    for x in sorted(R, key=lambda x: -cap(x[0]))[:3]:
        if x[0] not in seen: out.write(lin('R', x) + '\n'); seen.add(x[0])
    I = sorted([x for x in at if x[8][:2] in ('10', '11', '13', '14', '15', '31', '46') and cap(x[0]) >= 1e6], key=lambda x: -cap(x[0]))
    for x in I[:2]:
        if x[0] not in seen: out.write(lin('I', x) + '\n'); seen.add(x[0])
    A = sorted(at, key=lambda x: x[7] or '99999999')
    for x in A[:2]:
        if x[0] not in seen: out.write(lin('A', x) + '\n'); seen.add(x[0])
    for r, n in rz_f.most_common(1):
        if n >= 3 and r not in seen:
            x = next(y for y in at if y[0] == r)
            out.write(lin('O', x, f'({n} lojas)') + '\n')
out.close()
