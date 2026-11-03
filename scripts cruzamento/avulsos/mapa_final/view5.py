"""Visao compacta dos candidatos (a partir do view4) para decidir a mao o CNPJ das redes ainda sem CNPJ: v5.txt."""
import os as _os
_AQUI = _os.path.dirname(_os.path.abspath(__file__))
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_AQUI, 'dados'))
import pickle, json, collections, re
exec(open(_os.path.join(_AQUI, 'view4.py'), encoding='utf8').read().split("out = open")[0])
import unicodedata
q4 = json.load(open('q4.json', encoding='utf8'))
out = open('v5.txt', 'w', encoding='utf8')
for q in q4:
    if q['st'] == 'NF': continue
    k = q['id']; nm = k.split('|')[1].lstrip('+')
    sufx = r' (FRANCHISING|FRANQUIAS|FRANQUEADORA|FRANQUIA|PARTICIPACOES|HOLDING)$'
    base = [t for t in q['terms'] if not re.search(sufx, t)]
    # nome completo = os termos mais longos (com 2+ palavras) ou unicos de 6+ letras
    full = [t for t in base if len(t.split()) >= 2 or len(t) >= 6]
    idx = set().union(*[fh.get(t, set()) for t in full])
    at = [est[i] for i in idx if est[i][5] == '02']
    rz_f = collections.Counter(x[0] for x in at)
    lines = []
    for x in sorted([x for x in at if x[8] in FRQ], key=lambda x: (x[8] != '7740300', -cap(x[0])))[:3]: lines.append(lin('F', x))
    rr = set().union(*[rh.get(t, set()) for t in q['terms'] if re.search(r' (FRANCHISING|FRANQUIAS|FRANQUEADORA|FRANQUIA)$', t)])
    for r in rr:
        rows = [est[i] for i in by_raiz.get(r, [])]
        m = [x for x in rows if x[3] == '1' and x[5] == '02']
        if m: lines.append(lin('R', m[0]))
    for x in sorted([x for x in at if x[8][:2] in ('10', '11', '13', '14', '15', '31', '46') and cap(x[0]) >= 1e6], key=lambda x: -cap(x[0]))[:1]: lines.append(lin('I', x))
    for x in sorted(at, key=lambda x: x[7] or '9')[:1]: lines.append(lin('A', x))
    for r, n in rz_f.most_common(1):
        if n >= 3: lines.append(lin('O', next(y for y in at if y[0] == r), f'({n})'))
    out.write(f"## {k} [{q['st']}] {len(at)} lojas c/ nome em {len(rz_f)} CNPJs\n" + ''.join(l[:150] + '\n' for l in lines[:5]))
out.close()
