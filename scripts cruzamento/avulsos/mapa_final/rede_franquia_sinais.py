"""Rede própria x associação/franquia nas marcas da Alessandra Dayan (17/11/2026), passo 1: sinais por linha (TIPO da FARMACIA VF,
marca no observatório de franquias, fonte/observação da aba FINAL, CNAE 7740 do CNPJ, lojas ativas na raiz / lojas da marca).
Roda na pasta de trabalho (candidatos.json, cnpj.tsv, mk_slugs.txt, r_*.pkl, cópia do MAPA). Saída: sinais.json (fora do git)."""
import json, pickle, re, unicodedata, collections as C, openpyxl
def norm(x): return re.sub(r'[^a-z0-9]', '', unicodedata.normalize('NFKD', str(x or '')).encode('ascii', 'ignore').decode().lower())
cand = json.load(open('candidatos.json', encoding='utf8'))
slugs = [l.strip() for l in open('mk_slugs.txt', encoding='utf8')]
ns = {norm(s): s for s in slugs}
# CNPJ escolhido (tsv) por linha VF (linha antiga = lin; nova = lin+2? usar nome)
cn = {}
for l in open('cnpj.tsv', encoding='utf-8'):
    p = l.rstrip('\r\n').split('\t')
    if p[0] == 'Alessandra Dayan VF': cn[p[2]] = re.sub(r'\D', '', p[3])
est = C.defaultdict(list)
for f in ['r_alessandra.pkl', 'r_achados.pkl']:
    R = pickle.load(open(f, 'rb'))
    for x in R['est']: est[x[0]].append(x)
wb = openpyxl.load_workbook('v4_atual.xlsm', read_only=True, data_only=True)
vf = {}
for n, r in enumerate(wb['"FARMACIA VF"'].iter_rows(min_row=7, max_col=3, values_only=True), 7): vf[n] = r[2]
fin = {}
for aba in [a for a in wb.sheetnames if a.endswith('FINAL') and not a.startswith('I.R')]:
    rows = list(wb[aba].iter_rows(values_only=True))
    hi = next(i for i, r in enumerate(rows[:12]) if len(r) > 1 and str(r[1] or '').strip() == 'NOME FANTASIA')
    h = [str(x or '').strip() for x in rows[hi]]
    jf, jo, jg = h.index('FONTE DO Nº DE LOJAS'), h.index('OBSERVAÇÃO'), h.index('GRUPO ADMINISTRADOR')
    for k, r in enumerate(rows[hi + 1:], hi + 2):
        if len(r) > jo and r[1]: fin[(aba, k)] = (str(r[jf] or ''), str(r[jo] or ''), str(r[jg] or ''))
wb.close()
out = []
for o in cand:
    nome = o['nome']; s = []
    base = re.sub(r'\(.*?\)', '', nome).split('/')[0].split('+')[0].split(';')[0]
    k = norm(base)
    mk = ns.get(k) or next((v for kk, v in ns.items() if len(k) >= 5 and (kk.startswith(k) or k.startswith(kk)) and len(kk) >= 5), None)
    if mk: s.append('MK:' + mk)
    for c in o['cands'][:1]:
        m = re.match(r'"?FARMACIA VF"? linha (\d+)', c['fonte'].replace('FARMACIA VF', 'FARMACIA VF'))
        if c['fonte'].startswith('FARMACIA VF'): s.append('VF:' + str(vf.get(int(c['fonte'].split()[-1]), ''))[:12])
        if 'UNIDADE' in c['conf'] or 'FRANQUEADORA' in c['conf']: s.append('UNIDADE')
        m = re.match(r'(.+FINAL) linha (\d+)', c['fonte'])
        if m:
            f = fin.get((m.group(1), int(m.group(2))))
            if f:
                t = (f[0] + ' ' + f[1]).upper()
                if re.search(r'ABF|FRANQU|FRANCHIS|LICENCIA|ASSOCIA', t): s.append('FIN:' + re.search(r'ABF|FRANQU\w*|FRANCHIS\w*|LICENCIA\w*|ASSOCIA\w*', t).group(0))
    c = cn.get(nome)
    if c:
        e = est.get(c[:8], []); at = sum(1 for x in e if x[5] == '02'); cnae = next((x[8] for x in e if x[0] + x[1] + x[2] == c), '')
        if cnae == '7740300': s.append('CNAE7740')
        try: lj = int(o['lojas'])
        except Exception: lj = 0
        if lj: s.append(f'raiz {at}/{lj}')
    out.append((o['lin'], o['seg'][:10], nome[:45], o['lojas'], ' '.join(s)))
json.dump(out, open('sinais.json', 'w', encoding='utf8'), ensure_ascii=False)
for x in out: print(*x, sep=' | ')
