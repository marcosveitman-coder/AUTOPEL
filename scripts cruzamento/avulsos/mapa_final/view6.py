"""Visao dos candidatos da 6a consulta (redes sem CNPJ das linhas indicadas pelo usuario): v6.txt.
Por rede: lojas ativas com o nome (UFs, CNAEs) e os candidatos a franqueadora/operadora/dona da marca.
Codigos: F = CNAE de franquia/holding/administracao com o nome; R = razao social com o termo; O = raiz com mais lojas
ativas com o nome; U = estabelecimento com o nome na UF da sede (matriz, maior capital, mais antigo)."""
import os as _os
_AQUI = _os.path.dirname(_os.path.abspath(__file__))
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_AQUI, 'dados'))
import pickle, json, collections
D = pickle.load(open('r6.pkl', 'rb')); est, emp = D['est'], D['emp']
MUN = pickle.load(open('mun.pkl', 'rb'))
Q = json.load(open('q6.json', encoding='utf8'))
fh = collections.defaultdict(set)
for i, t in D['hits_fant']: fh[t].add(i)
hr = collections.defaultdict(set)
for r, t in D['hits_raz']: hr[t].add(r)
porraiz = collections.defaultdict(list)
for x in est: porraiz[x[0]].append(x)
FRQ = ('7740300', '8299799', '7020400', '8211300', '6462000', '6463800', '7490104', '8219999', '4618499', '4619200')
SEGP = {'ALIM': ('56', '10', '47', '46'), 'LOJ': ('47', '14', '15', '46', '32'), 'CONS': ('47', '46', '23', '16', '25')}
def cap(r):
    try: return float(emp.get(r, ('', '', '0', ''))[2].replace(',', '.'))
    except Exception: return 0.0
def linha(x, tag, n_nome):
    rs = porraiz[x[0]]; at = sum(1 for y in rs if y[5] == '02')
    return (f"   {tag} {x[0]}{x[1]}{x[2]} {'M' if x[3] == '1' else 'F'} sit{x[5]} {x[8]} ini{(x[7] or '')[:4]} {MUN.get(x[11], x[11])[:18]}/{x[10]} | "
            f"{emp.get(x[0], ('?',))[0][:48]} | cap{cap(x[0]) / 1e6:.2f}M | ativos raiz {at} | c/ nome {n_nome} | fant {x[4][:26]}")
out = open('v6.txt', 'w', encoding='utf8')
for q in Q:
    seg = q['id'].split('|')[0]; uf = q['uf']
    ids = {i for t in q['terms'] for i in fh.get(t, ())}
    ativos = [est[i] for i in ids if est[i][5] == '02']
    seg_at = [x for x in ativos if x[8].startswith(SEGP[seg]) or x[8] in FRQ]
    cnt = collections.Counter(x[0] for x in seg_at)
    ufs = collections.Counter(x[10] for x in seg_at).most_common(6)
    out.write(f"## {q['lin']} {q['id']} | {q['fant']} | lojas origem {q['lojas']} | UF {uf} | {q['st']}\n")
    out.write(f"   com o nome: {len(ativos)} ativos, {len(seg_at)} do segmento, {len(cnt)} raizes | UFs {ufs} | termos {q['terms']}\n")
    vistos = set()
    def put(x, tag):
        c = x[0] + x[1] + x[2]
        if c in vistos: return
        vistos.add(c); out.write(linha(x, tag, cnt.get(x[0], 0)) + '\n')
    for x in sorted([x for x in ativos if x[8] in FRQ], key=lambda x: -cap(x[0]))[:4]: put(x, 'F')
    for t in q['terms']:
        for r in sorted(hr.get(t, ()), key=lambda r: -cap(r))[:5]:
            rs = porraiz[r]
            m = [y for y in rs if y[3] == '1'] or rs
            if m: put(m[0], 'R')
    for r, n in cnt.most_common(3):
        m = [y for y in porraiz[r] if y[3] == '1'] or porraiz[r]
        put(m[0], 'O')
    loc = sorted([x for x in seg_at if x[10] == uf], key=lambda x: (x[3] != '1', -cap(x[0]), x[7] or '9'))
    for x in loc[:3]: put(x, 'U')
out.close()
print('ok')
