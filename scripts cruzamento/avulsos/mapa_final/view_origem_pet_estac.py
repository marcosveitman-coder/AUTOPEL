"""Para cada rede citada (estacionamentos da aba Alessandra Dayan e redes pet da imprensa), mostra as empresas (raiz) que
a busca por nome achou (r12.pkl): razao social, matriz, unidades ativas no segmento (r11.pkl) e quantas unidades tem o
nome. Serve para decidir a mao o CNPJ de cada rede (dec_pet_estac.json). Uso: python view_origem_pet_estac.py -> v_origem.txt"""
import os as _os
_AQUI = _os.path.dirname(_os.path.abspath(__file__))
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_AQUI, 'dados'))
import pickle, json, collections
R11 = pickle.load(open('r11.pkl', 'rb'))['raiz']
R12 = pickle.load(open('r12.pkl', 'rb'))
EMP = dict(R12['emp'])
for _pk in ['r13.pkl']:
    if _os.path.exists(_pk): EMP.update(pickle.load(open(_pk, 'rb'))['emp'])
Q = json.load(open('q12.json', encoding='utf8'))
SEGS = {'PET': ['PET', 'VET', 'BANHO', 'HOTEL', 'VETMED'], 'EST': ['ESTAC']}
CNS = {'PET': {'4789004', '7500100', '9609208', '9609207', '4771704'}, 'EST': {'5223100'}}
est = R12['est']
por_termo = collections.defaultdict(lambda: collections.defaultdict(lambda: [0, 0]))  # termo -> raiz -> [unid. do segmento com o nome, total com o nome]
for i, t in R12['hits_fant']:
    x = est[i]
    if x[5] != '02': continue
    s = por_termo[t][x[0]]; s[1] += 1
    por_termo[t][x[0]][0] += 0  # contado abaixo por segmento
raz_termo = collections.defaultdict(set)
for r, t in R12['hits_raz']: raz_termo[t].add(r)
estr = collections.defaultdict(list)
for x in est: estr[x[0]].append(x)
def unid(seg, r): return sum(R11[(r, s)][3] for s in SEGS[seg] if (r, s) in R11)
def mat(r):
    m = [x for x in estr.get(r, []) if x[3] == '1']
    return (m[0][0] + m[0][1] + m[0][2], m[0][5], m[0][8], m[0][10], m[0][4]) if m else None
with open('v_origem.txt', 'w', encoding='utf8') as f:
    for q in Q:
        seg = q['id'].split('|')[0]
        cand = collections.Counter()
        for t in q['terms']:
            for r, (a, b) in por_termo[t].items(): cand[r] += b
            for r in raz_termo[t]: cand[r] += 0.5
        f.write(f"\n=== {q['id']}  termos={q['terms']}\n")
        lst = sorted(cand, key=lambda r: (-unid(seg, r), -cand[r]))[:8]
        for r in lst:
            m = mat(r)
            f.write(f"  {r} unid_seg={unid(seg, r):4d} com_nome={cand[r]:5.1f} :: {EMP.get(r, ('?',))[0][:55]} :: matriz={m}\n")
print('ok')
