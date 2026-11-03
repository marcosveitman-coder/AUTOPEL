"""Candidatos a franqueadora/matriz na UF de origem conhecida da marca (CNAE de franquia/administracao primeiro, depois capital).
Saida na tela; a decisao vai para dados/dec_5.json."""
import os as _os
_AQUI = _os.path.dirname(_os.path.abspath(__file__))
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_AQUI, 'dados'))
import pickle, collections, re
D = pickle.load(open('r4.pkl', 'rb')); emp, est = D['emp'], D['est']
MUN = pickle.load(open('mun.pkl', 'rb'))
fh = collections.defaultdict(set)
for i, t in D['hits_fant']: fh[t].add(i)
FRQ = ('7740300', '8299799', '7020400', '8211300', '6462000', '6463800', '7490104', '8219999')
def cap(r):
    try: return float(emp.get(r, ('', '', '0', ''))[2].replace(',', '.'))
    except: return 0.0
B = [('CHIQUINHO SORVETES', 'SP'), ('MR FIT', 'SP'), ('MAIS1 CAFE', 'PR'), ('MAIS1', 'PR'), ('DIVINO FOGAO', 'SP'), ('FABRICA DE BOLO VO ALZIRA', 'RJ'),
     ('MUNDO VERDE', 'RJ'), ('DOMINO S PIZZA', 'SP'), ('REI DO MATE', 'SP'), ('GELATO BORELLI', 'SP'), ('NUMBER ONE CHICKEN', 'ES'), ('N1 CHICKEN', 'ES'),
     ('HAVANNA', 'SP'), ('MR SHAKE', 'MG'), ('MR MIX MILK SHAKES', 'SP'), ('JAH DO ACAI', 'MG'), ('BOALI', 'SP'), ('CAROL COXINHAS', 'MG'), ('PIZZA CREK', 'SP'),
     ('MINI KALZONE', 'SC'), ('GELA BOCA', 'PR'), ('MANIA DE CHURRASCO', 'SP'), ('CONSTANCE', 'MG'), ('JORGE BISCHOFF', 'RS'), ('MR CAT', 'RJ'),
     ('PETITE JOLIE', 'RS'), ('PUKET', 'SP'), ('MINISO', 'SP'), ('CAPODARTE', 'RJ'), ('SERGIO S', 'MG'), ('REAL CALCADOS', 'BA'), ('CROSBY', 'CE'),
     ('LITORANEUS', 'PE'), ('SAPATARIA NOVA', 'CE'), ('FLORENZA', 'GO'), ('PECA RARA', 'DF'), ('MAGAZINE GOIAS', 'GO'), ('CASA PRADO', 'MT'),
     ('HOT N TENDER', 'RJ'), ('JOHNNIE GRILL', 'MG'), ('GORDAO LANCHES', 'SP'), ('PORKS', 'PR'), ('PATRONI', 'SP'), ('EX BURGUER', 'SP'), ('MAKIS PLACE', 'SP')]
for nome, uf in B:
    idx = fh.get(nome, set())
    rows = [est[i] for i in idx if est[i][5] == '02']
    cnt = collections.Counter(x[0] for x in rows)
    cand = [x for x in rows if x[10] == uf]
    cand.sort(key=lambda x: (x[8] not in FRQ, -cap(x[0]), x[7]))
    print(f"## {nome} ({uf}) {len(rows)} lojas ativas")
    for x in cand[:3]:
        print(f"   {x[0]}{x[1]}{x[2]} {'M' if x[3]=='1' else 'F'} {x[8]} ini{x[7][:4]} {MUN.get(x[11], x[11])[:16]:16} | {emp.get(x[0], ('?',))[0][:42]} | cap{cap(x[0])/1e6:.1f}M | lojas da raiz {cnt[x[0]]} | {x[4][:24]}")
