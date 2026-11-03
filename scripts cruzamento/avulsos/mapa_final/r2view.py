import pickle, json, re
import os as _os
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'dados'))
R = pickle.load(open('r2.pkl', 'rb'))
D1 = pickle.load(open('r1.pkl', 'rb'))
emp = D1['emp']
dec = {}
for f in ['dec_alim1.json', 'dec_2.json', 'dec_3.json']: dec.update(json.load(open(f, encoding='utf8')))
used = {v[0][:8] for v in dec.values() if v[0] and v[0][0].isdigit()}
TH = {'ALIM': 15, 'LOJ': 20, 'CONS': 6}
out = {'ALIM': [], 'LOJ': [], 'CONS': []}
for (r, s), (cn, uf, fa, n) in R['raiz'].items():
    if n >= TH[s] and r not in used:
        out[s].append((n, r, cn, uf, fa))
for s in out:
    out[s].sort(reverse=True)
    with open(f'r2_{s}.txt', 'w', encoding='utf8') as f:
        for n, r, cn, uf, fa in out[s]:
            rz = emp.get(r, ('?',))[0]
            u = ' '.join(f'{k}:{v}' for k, v in sorted(uf.items(), key=lambda x: -x[1])[:4])
            c = max(cn, key=cn.get)
            f.write(f"{n:4} {r} {rz[:45]:45} | {c} | {u} | {'; '.join(x[0][:22] for x in fa[:3])}\n")
    print(s, len(out[s]))
# fantasias em muitas raizes (franquias)
FS = sorted(((v[1], k[0], k[1], v[0], v[2]) for k, v in R['fant'].items()), reverse=True)
with open('r2_fant.txt', 'w', encoding='utf8') as f:
    for nr, fn, s, n, uf in FS:
        if nr >= 25:
            u = ' '.join(f'{k}:{v}' for k, v in sorted(uf.items(), key=lambda x: -x[1])[:3])
            f.write(f"{s:4} {nr:4} cnpjs {n:4} estab | {fn[:40]:40} | {u}\n")
