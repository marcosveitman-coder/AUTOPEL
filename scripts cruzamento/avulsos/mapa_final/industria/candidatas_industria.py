"""Aba FULLFILMIT_INDUSTRIA FINAL, passo 2: escolhe as candidatas por segmento e por região (19/11/2026).
Entrada: r_ind.pkl (receita_industria.py) e exame_mm26.json (ranking Exame Melhores e Maiores 2026, receita 2025, extraído
da página exame.com/mm/maiores: empresa, setor, posição, receita em R$ mil, cidade, UF).
Filtro da Receita (tira erro de cadastro, MEI e empresa de fachada):
  - porte "demais" (05) e natureza de sociedade (fora empresário individual 2135, EIRELI 2305/2321);
  - capital social >= R$ 5 mi.
Segmento da empresa = segmento do CNAE da matriz, se for industrial; senão, o mais comum entre as unidades.
Corte (decisão nesta sessão, explicada na aba R_E_S_U_M_O):
  - Brasil: as maiores de cada segmento pelo capital social (FMCG 60, QUIM 50, os outros 40);
  - Região: as 6 maiores de cada segmento em cada região (UF da matriz), com capital >= R$ 20 mi;
  - PNEU e OLEO (lubrificantes): todas com capital >= R$ 20 mi, porque são poucas.
Saída: cand.pkl {raiz: {...}} e cand_view.txt para a revisão à mão.
Uso: python candidatas_industria.py <pasta de trabalho>"""
import os, sys, pickle, json, re, unicodedata, collections
P = sys.argv[1]
d = pickle.load(open(os.path.join(P, 'r_ind.pkl'), 'rb')); R, E = d['raiz'], d['emp']
REG = {}
for r, ufs in [('NORTE', 'AC AP AM PA RO RR TO'), ('NORDESTE', 'AL BA CE MA PB PE PI RN SE'), ('CENTRO-OESTE', 'DF GO MT MS'),
               ('SUDESTE', 'ES MG RJ SP'), ('SUL', 'PR RS SC')]:
    for u in ufs.split(): REG[u] = r
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from receita_industria import seg_de


def cap(r):
    try: return float(E[r][2])
    except Exception: return 0.0


def seg(r):
    m = R[r]['matriz']
    if m and seg_de(m[4]): return seg_de(m[4])
    return R[r]['seg'].most_common(1)[0][0]


def uf_sede(r):
    m = R[r]['matriz']
    return m[2] if m else R[r]['uf'].most_common(1)[0][0]


ok = [r for r in R if r in E and E[r][3] == '05' and E[r][1] not in ('2135', '2305', '2321') and cap(r) >= 5e6]
by = collections.defaultdict(list)
for r in ok: by[seg(r)].append(r)
sel = {}
for s, l in by.items():
    l.sort(key=lambda r: -cap(r))
    if s in ('PNEU', 'OLEO'):
        for r in l:
            if cap(r) >= 20e6: sel.setdefault(r, set()).add('BRASIL')
        continue
    for r in l[:{'FMCG': 60, 'QUIM': 50}.get(s, 40)]: sel.setdefault(r, set()).add('BRASIL')
    reg = collections.defaultdict(list)
    for r in l:
        if cap(r) >= 20e6: reg[REG.get(uf_sede(r), '?')].append(r)
    for g, rr in reg.items():
        for r in rr[:6]: sel.setdefault(r, set()).add('REGIAO ' + g)

# receita Exame (R$ mil) por nome + UF, só como apoio (a ligação final é revisada à mão)
def norm(x): return ' '.join(re.sub(r'[^A-Z0-9 ]', ' ', unicodedata.normalize('NFKD', x).encode('ascii', 'ignore').decode().upper()).split())
ex = json.load(open(os.path.join(P, 'exame_mm26.json'), encoding='utf8'))
out = {}
for r, ori in sel.items():
    x = R[r]; m = x['matriz'] or ('', '', uf_sede(r), '', '', '')
    out[r] = dict(raiz=r, razao=E[r][0], natureza=E[r][1], capital=cap(r), seg=seg(r), origem=sorted(ori), un=x['n'],
                  ufs=dict(x['uf']), uf=uf_sede(r), regiao=REG.get(uf_sede(r), '?'), mun=m[3], cnae=m[4],
                  fant=(m[5] or (x['fant'].most_common(1)[0][0] if x['fant'] else '')), ordem=m[0], dv=m[1])
pickle.dump(out, open(os.path.join(P, 'cand.pkl'), 'wb'))
with open(os.path.join(P, 'cand_view.txt'), 'w', encoding='utf8') as f:
    for s in sorted(by):
        f.write(f'===== {s}\n')
        for r in sorted([r for r in out if out[r]['seg'] == s], key=lambda r: -out[r]['capital']):
            o = out[r]
            f.write(f"{r} | {o['razao'][:50]} | cap {o['capital']/1e6:,.0f} mi | un {o['un']} | {o['uf']}/{o['regiao'][:2]} | "
                    f"cnae {o['cnae']} | {o['fant'][:22]} | {','.join(o['origem'])}\n")
print(len(out), collections.Counter(o['seg'] for o in out.values()))
