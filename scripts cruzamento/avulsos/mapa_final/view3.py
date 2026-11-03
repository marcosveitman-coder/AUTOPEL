"""Visao compacta da 3a busca: para cada nome, O = quem mais tem lojas ativas com o nome, F = razao social com o nome
(franqueadora/holding/capital alto)."""
import pickle, json, collections, re, sys
import os as _os
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'dados'))
D = pickle.load(open('r3.pkl', 'rb')); Q = json.load(open('q3.json', encoding='utf8'))
emp, est = D['emp'], D['est']
SIT = {'01': 'NUL', '02': 'ATV', '03': 'SUS', '04': 'INA', '08': 'BAI'}
by_raiz = collections.defaultdict(list)
for i, r in enumerate(est): by_raiz[r[0]].append(i)
fh = collections.defaultdict(set)
for i, t in D['hits_fant']: fh[t].add(i)
rh = collections.defaultdict(set)
for r, t in D['hits_raz']: rh[t].add(r)
def cap(r):
    try: return float(emp.get(r, ('', '', '0', ''))[2].replace(',', '.'))
    except: return 0.0
FRQ = ('7740300', '8299799', '7020400', '6462000', '6463800', '7490104', '8211300')
def desc(r):
    rows = [est[i] for i in by_raiz.get(r, [])]
    m = [x for x in rows if x[3] == '1']; m = m[0] if m else (rows[0] if rows else None)
    at = sum(1 for x in rows if x[5] == '02')
    cc = f"{m[0]}{m[1]}{m[2]}" if m else r + '??????'
    return f"{cc} {SIT.get(m[5], '?') if m else '?'} {m[10] if m else '--'} {m[8] if m else '-------'} {emp.get(r, ('?',))[0][:46]} cap{cap(r)/1e6:.1f}M at{at}"
with open('v3.txt', 'w', encoding='utf8') as f:
    for q in Q:
        if q['id'] == 'RAIZES': continue
        idx = set().union(*[fh.get(t, set()) for t in q['terms']])
        at = [i for i in idx if est[i][5] == '02']
        rz_f = collections.Counter(est[i][0] for i in at)
        rz_r = set().union(*[rh.get(t, set()) for t in q['terms']])
        f.write(f"## {q['id']} | {len(at)} estab ativos com o nome em {len(rz_f)} CNPJs | razao: {len(rz_r)}\n")
        for r, n in rz_f.most_common(2):
            if n >= 2: f.write(f"  O {n}: {desc(r)}\n")
        fr = []
        for r in (rz_r if len(rz_r) <= 800 else []):
            rows = [est[i] for i in by_raiz.get(r, [])]
            m = [x for x in rows if x[3] == '1']
            if not m or m[0][5] != '02' or m[0][10] == 'EX': continue
            rz = emp.get(r, ('',))[0]
            if re.search(r'FRANCHI|FRANQU|PARTICIPA|HOLDING|ADMINISTRA|S\.?A\.?$|S/A', rz) or m[0][8] in FRQ or cap(r) >= 3e6:
                fr.append((cap(r), r))
        for cp, r in sorted(fr, reverse=True)[:4]: f.write(f"  F {desc(r)}\n")
