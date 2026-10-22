import pickle, json, re, collections as C, unicodedata, sys
sys.path.insert(0,'receita')
from receita_l0 import norm
L0=json.load(open('l0.json',encoding='utf8'))
cand=pickle.load(open('receita/cand_razao.pkl','rb'))
est=pickle.load(open('receita/estab_l0.pkl','rb'))
SUP={'4711301','4711302','4712100'}; ATAC={'4691500','4639701','4639702','4637107'}
CDW=re.compile(r'\bCD\b|DEPOSIT|CENTRO DE DISTRIB|ESCRIT|ADMINISTR|LOGISTIC|ARMAZEM GERAL|GALPAO')
byroot=C.defaultdict(list)
for e in est: byroot[e[0]].append(e)
porraz=C.defaultdict(set)
for r,n,raw in cand: porraz[n].add(r)
def conta(root):
    es=byroot.get(root,[])
    at=[e for e in es if e[4].lstrip('0')=='2']
    sup=[e for e in at if e[5] in SUP]
    atac=[e for e in at if e[5] in ATAC and not CDW.search(e[3].upper())]
    ufs=C.Counter(e[6] for e in at)
    return dict(total=len(es), ativos=len(at), sup=len(sup), atac=len(atac), ufs=dict(ufs), cnae=C.Counter(e[5] for e in at).most_common(4), sit=C.Counter(e[4] for e in es))
res=[]
for o in L0:
    r=dict(lin=o['lin'], uf=o['uf'], fant=o['fant'], raz=o['raz'], porte=o['porte'], cnpj=o['cnpj'], st=o['st'], sitW=o['sit'])
    if o['cnpj']:
        r['raiz']=o['cnpj'][:8]; r['como']='CNPJ da aba'; r.update(conta(r['raiz']))
    else:
        roots=porraz.get(norm(o['raz']),set())|porraz.get(norm(o['fant']),set())
        noUF=[x for x in roots if any(e[6]==o['uf'] for e in byroot.get(x,[]))]
        r['n_cand']=len(roots); r['n_cand_uf']=len(noUF)
        if len(noUF)==1:
            r['raiz']=noUF[0]; r['como']='razão social + UF (Receita)'; r.update(conta(r['raiz']))
        elif len(noUF)>1:
            # desempate: o que tem loja de supermercado/atacado ativa na UF
            boas=[x for x in noUF if any(e[6]==o['uf'] and e[4].lstrip('0')=='2' and (e[5] in SUP or e[5] in ATAC) for e in byroot[x])]
            if len(boas)==1: r['raiz']=boas[0]; r['como']='razão social + UF (Receita, desempate pelo CNAE)'; r.update(conta(r['raiz']))
            else: r['como']=f'ambíguo ({len(noUF)} raízes na UF)'
        else:
            r['como']='não achado na Receita'
    res.append(r)
json.dump(res, open('rf_l0.json','w',encoding='utf8'), ensure_ascii=False, indent=0, default=str)
print(C.Counter(r['como'] for r in res))
com=[r for r in res if 'raiz' in r]
print('com raiz', len(com), 'com loja sup', sum(1 for r in com if r['sup']>0), 'só atac', sum(1 for r in com if r['sup']==0 and r['atac']>0), 'nenhuma ativa', sum(1 for r in com if r['ativos']==0))
print('sit estab (raízes sem ativo):', C.Counter(k for r in com if r['ativos']==0 for k in r['sit']))
for r in res:
    if r['porte'] in 'AB' or (('raiz' in r) and r['porte']=='C' and not r['cnpj']):
        print(r['lin'], r['porte'], r['uf'], str(r['fant'])[:28], '|', r['como'][:28], '|', r.get('raiz'), 'sup',r.get('sup'),'atac',r.get('atac'),'ativos',r.get('ativos'), r.get('ufs'), r.get('cnae'))
