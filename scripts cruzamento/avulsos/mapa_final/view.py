"""Visao compacta para decisao manual: O=operadora (mais lojas ativas com o nome), F=franqueadora/holding com o nome
na razao social, S=CNPJ da origem, I=CNPJ validado na aba IRTT."""
import pickle, json, collections, re, sys, unicodedata
import os as _os
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'dados'))
D = pickle.load(open(sys.argv[1], 'rb')); Q = json.load(open(sys.argv[2], encoding='utf8'))
G = json.load(open('groups.json', encoding='utf8'))
emp, est = D['emp'], D['est']
SIT = {'01':'NUL','02':'ATV','03':'SUS','04':'INA','08':'BAI'}
by_raiz = collections.defaultdict(list)
for i, r in enumerate(est): by_raiz[r[0]].append(i)
fant_hits = collections.defaultdict(set)
for i, t in D['hits_fant']: fant_hits[t].add(i)
raz_hits = collections.defaultdict(set)
for r, t in D['hits_raz']: raz_hits[t].add(r)
TR = str.maketrans({c: ' ' for c in '!"#$%\'()*+,-./:;<=>?@[\]^_`{|}~’'})
def norm(s):
    s = unicodedata.normalize('NFKD', s or '').encode('ascii', 'ignore').decode().upper().replace('&', ' E ')
    return ' '.join(s.translate(TR).split())
def cap(r):
    try: return float(emp.get(r, ('', '', '0', ''))[2].replace(',', '.'))
    except: return 0.0
IR = collections.defaultdict(list)
d2 = json.load(open('src2.json', encoding='utf8'))['I.R.T.T.__Por bandeira (FINAL)']
for r in d2[4:]:
    if r[24]: IR[norm(r[1])].append(re.sub(r'\D', '', r[24]))
FRQ = ('7740300', '8299799', '7020400', '6462000', '6463800', '7490104', '8211300')
def desc(r, c=None):
    rows = [est[i] for i in by_raiz.get(r, [])]
    if c: m = [x for x in rows if x[1] == c[8:12]]
    else: m = [x for x in rows if x[3] == '1']
    m = m[0] if m else None
    at = sum(1 for x in rows if x[5] == '02')
    cc = c or (f"{m[0]}{m[1]}{m[2]}" if m else r + '??????')
    return f"{cc} {SIT.get(m[5],'?') if m else '?'} {m[10] if m else '--'} {m[8] if m else '-------'} {emp.get(r, ('?',))[0][:44]} cap{cap(r)/1e6:.1f}M at{at}"
with open(sys.argv[3], 'w', encoding='utf8') as f:
    for q in Q:
        mem = G[q['id']]
        ft = set()
        for x in mem:
            if x['lump']: continue
            s = re.sub(r'\(.*?\)', '', x['fant'])
            for p in re.split(r'\s*/\s*|\s+\+\s+', s):
                n = norm(p)
                if n: ft.add(n)
                w = n.split()
                if len(w) > 1 and w[0] in ('LOJAS', 'LOJA', 'REDE', 'GRUPO', 'THE'): ft.add(' '.join(w[1:]))
        ft = {t for t in ft if t in q['terms'] or t.replace(' ', '') in q['terms']} or set(q['terms'][:1])
        ftc = ft | {t.replace(' ', '') for t in ft}
        idx = set().union(*[fant_hits.get(t, set()) for t in ftc])
        at = [i for i in idx if est[i][5] == '02']
        rz_f = collections.Counter(est[i][0] for i in at)
        rz_r = set().union(*[raz_hits.get(t, set()) for t in ftc])
        lj = '; '.join(sorted({f"{x['lojas'][:10]}" + (f"({x['base']})" if x['base'] else '') for x in mem if x['lojas'] and not x['lump']}))
        f.write(f"## {q['id']} [{'/'.join(sorted(ft))[:40]}] lojas fonte: {lj[:60]} | Receita: {len(at)} estab ativos c/ nome em {len(rz_f)} CNPJs\n")
        for c in q['cnpjs']:
            row = [est[i] for i in by_raiz.get(c[:8], []) if est[i][1] == c[8:12] and est[i][2] == c[12:]]
            f.write(f"  S {c} NAO EXISTE\n" if not row else f"  S {desc(c[:8], c)} nome:{'sim' if (c[:8] in rz_r or rz_f.get(c[:8])) else 'NAO'}\n")
        k = norm(q['id'].split('|')[1])
        for kk, cs in IR.items():
            if kk == k or (len(k) >= 5 and (kk.startswith(k + ' ') or k.startswith(kk + ' ') or kk == k)):
                for c in sorted(set(cs)): f.write(f"  I {kk[:18]} {desc(c[:8], c) if by_raiz.get(c[:8]) else c + ' (sem dados)'}\n")
        for r, n in rz_f.most_common(2):
            if n >= 3: f.write(f"  O {n} lojas: {desc(r)}\n")
        fr = []
        for r in rz_r if len(rz_r) <= 500 else []:
            rows = [est[i] for i in by_raiz.get(r, [])]
            m = [x for x in rows if x[3] == '1']
            if not m or m[0][5] != '02' or m[0][10] == 'EX': continue
            rz = emp.get(r, ('',))[0]
            if re.search(r'FRANCHI|FRANQU|PARTICIPA|HOLDING|ADMINISTRA|S\.?A\.?$|S/A', rz) or m[0][8] in FRQ or cap(r) >= 3e6:
                fr.append((cap(r), r))
        for cp, r in sorted(fr, reverse=True)[:3]: f.write(f"  F {desc(r)}\n")
