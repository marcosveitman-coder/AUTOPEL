"""Visao para decidir as redes das abas PET SHOP  FINAL e ESTACIONAMENTO FINAL a partir da varredura por CNAE (r11.pkl,
redes_pet_estac.py) e da busca por nome (r12.pkl, busca_receita.py q12.json).
Mostra, por regiao: empresas (raiz) com mais unidades ativas do segmento e nomes fantasia repetidos em varias raizes
(franquias). Uso: python view_pet_estac.py  ->  v_pet.txt e v_estac.txt"""
import os as _os
_AQUI = _os.path.dirname(_os.path.abspath(__file__))
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_AQUI, 'dados'))
import pickle, collections
R = pickle.load(open('r11.pkl', 'rb')); RA, FA = R['raiz'], R['fant']
EMP = {}
for _pk in ['r12.pkl', 'r13.pkl']:  # razao social das redes citadas (r12) e das candidatas (r13)
    if _os.path.exists(_pk): EMP.update(pickle.load(open(_pk, 'rb'))['emp'])
REG = {}
for g, ufs in [('NORTE', 'AC AP AM PA RO RR TO'), ('NORDESTE', 'AL BA CE MA PB PE PI RN SE'), ('CENTRO-OESTE', 'DF GO MT MS'),
               ('SUDESTE', 'ES MG RJ SP'), ('SUL', 'PR RS SC')]:
    for u in ufs.split(): REG[u] = g
GRUPO = {'PET': ['PET', 'VET', 'BANHO', 'HOTEL', 'VETMED'], 'ESTAC': ['ESTAC']}
for nome, segs in GRUPO.items():
    # soma por raiz das unidades nos CNAEs do grupo
    tot = collections.defaultdict(lambda: [0, collections.Counter(), collections.Counter(), collections.Counter()])
    for (r, s), v in RA.items():
        if s not in segs: continue
        t = tot[r]; t[0] += v[3]; t[1].update(v[1]); t[2][s] += v[3]
        for f, k in v[2]: t[3][f] += k
    linhas = sorted(tot.items(), key=lambda x: -x[1][0])
    por = collections.defaultdict(list)
    for r, t in linhas:
        uf = t[1].most_common(1)[0][0]; por[REG.get(uf, '?')].append((r, t, uf))
    with open(f"v_{'pet' if nome == 'PET' else 'estac'}.txt", 'w', encoding='utf8') as f:
        f.write(f'== {nome}: {len(linhas)} empresas com 3+ unidades\n')
        for g in ['BRASIL', 'NORTE', 'NORDESTE', 'CENTRO-OESTE', 'SUDESTE', 'SUL']:
            lst = [(r, t, t[1].most_common(1)[0][0]) for r, t in linhas] if g == 'BRASIL' else por[g]
            f.write(f'\n--- {g}\n')
            for r, t, uf in lst[:60 if g == 'BRASIL' else 45]:
                raz = EMP.get(r, ('',))[0][:50]
                ufs = ' '.join(f'{u}{k}' for u, k in t[1].most_common(4))
                cn = ' '.join(f'{s}{k}' for s, k in t[2].most_common())
                fan = ' | '.join(f'{x}({k})' for x, k in t[3].most_common(3))
                f.write(f'{r} {t[0]:4d} [{cn}] {ufs} :: {raz} :: {fan[:110]}\n')
        f.write('\n--- NOMES FANTASIA EM VARIAS RAIZES (franquias / homonimos)\n')
        fl = collections.defaultdict(lambda: [0, set(), collections.Counter()])
        for (fn, s), v in FA.items():
            if s not in segs: continue
            x = fl[fn]; x[0] += v[0]; x[1] |= set(v[4]); x[2].update(v[2])
        for fn, x in sorted(fl.items(), key=lambda y: -len(y[1][1]))[:150]:
            if len(x[1]) < 4: break
            ufs = ' '.join(f'{u}{k}' for u, k in x[2].most_common(5))
            f.write(f'{fn[:40]:40s} est={x[0]:4d} raizes={len(x[1]):4d} {ufs}\n')
    print(nome, len(linhas))
