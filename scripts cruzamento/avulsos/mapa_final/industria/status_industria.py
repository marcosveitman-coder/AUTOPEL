"""Aba FULLFILMIT_INDUSTRIA FINAL, passo 5: STATUS AUTOPEL das indústrias (19/11/2026), com o motor de vínculos de 16/11
(../../status_autopel/motor.py e avalia.py), sem regravar as outras abas.
As linhas novas entram como entidades da família IND ao lado das entidades das outras abas (entidades.pkl de 16/11): assim
a trava "raiz de CNPJ de outra rede não vale por nome" continua valendo contra todas as redes do MAPA.
Entrada: ind_rows.pkl (build_industria.py), <pasta 16/11>/entidades.pkl e clientes.pkl (base_clientes.py).
Saída: <pasta de trabalho>/status/ (entidades.pkl, clientes.pkl -> cópia, resultados.pkl) e ind_status.pkl {raiz: resultado}.
Uso: python status_industria.py <pasta de trabalho> <pasta 16/11> AAAA-MM-DD"""
import os, sys, pickle, shutil, subprocess
P, P16, HOJE = sys.argv[1:4]
AQUI = os.path.dirname(os.path.abspath(__file__))
ST = os.path.join(AQUI, '..', '..', 'status_autopel')
sys.path.insert(0, ST)
from motor import nomes_da_rede

rows = pickle.load(open(os.path.join(P, 'ind_rows.pkl'), 'rb'))
E = pickle.load(open(os.path.join(P16, 'entidades.pkl'), 'rb'))
E = [e for e in E if e['aba'] != 'FULLFILMIT_INDUSTRIA FINAL']
cl0 = max(e['cluster'] for e in E) + 1
idx = {}
for i, o in enumerate(rows):
    nomes = nomes_da_rede(o['fant'])
    if o['grupo'] and o['grupo'] != o['fant']:
        nomes += [n for n in nomes_da_rede(o['grupo']) if n not in nomes]
    idx[o['raiz']] = len(E)
    E.append(dict(aba='FULLFILMIT_INDUSTRIA FINAL', linha=7 + i, nomes=nomes, raizes=sorted(o['raizes']), razoes=[o['raz']],
                  codigos=[], grupos=[], extra=dict(nome=o['fant'], uf=o['uf'], status_antes='', apoio=False, origem=o['origem']),
                  fam='IND', cluster=cl0 + i))
D = os.path.join(P, 'status'); os.makedirs(D, exist_ok=True)
pickle.dump(E, open(os.path.join(D, 'entidades.pkl'), 'wb'))
if not os.path.exists(os.path.join(D, 'clientes.pkl')): shutil.copy(os.path.join(P16, 'clientes.pkl'), os.path.join(D, 'clientes.pkl'))
subprocess.run([sys.executable, os.path.join(ST, 'avalia.py'), D, HOJE], check=True)
R = pickle.load(open(os.path.join(D, 'resultados.pkl'), 'rb'))
out = {r: R.get(n) for r, n in idx.items()}
pickle.dump(out, open(os.path.join(P, 'ind_status.pkl'), 'wb'))
import collections
print(collections.Counter((v or {}).get('status') for v in out.values()))
print(collections.Counter((v or {}).get('via') for v in out.values() if v and v.get('status') != 'NÃO É CLIENTE'))
