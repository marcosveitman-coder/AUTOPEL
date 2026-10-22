import pickle, re, collections as C, unicodedata, json
from openpyxl.utils import column_index_from_string as ci
D=lambda x: re.sub(r'\D','',str(x or ''))
sa=lambda x: unicodedata.normalize('NFKD',str(x or '')).encode('ascii','ignore').decode().upper()
TIPO={'SUPERMERCADO','SUPERMERCADOS','SUPER','MERCADO','MERCADINHO','MERCEARIA','ATACADO','ATACADISTA','ATACAREJO','COMERCIO','COMERCIAL','DE','DO','DA','DOS','DAS','E','LTDA','ME','EPP','EIRELI','SA','S','A','REDE','GRUPO','CIA','ALIMENTOS','HIPERMERCADO','VAREJO','DISTRIBUIDORA'}
def nuc(s):
    w=re.findall(r'[A-Z0-9]+', sa(re.sub(r'\(.*?\)','',str(s or ''))).replace('&',' E '))
    while w and w[0] in TIPO: w.pop(0)
    while w and w[-1] in TIPO: w.pop()
    return ' '.join(w)
a=pickle.load(open('abas.pkl','rb'))['SUPERMERCADOS (reuniao)']
g=lambda r,c: r[ci(c)-1]
rows={i+1:r for i,r in enumerate(a) if i>=4 and any(v is not None for v in r)}
par={k:k for k in rows}
def f(x):
    while par[x]!=x: par[x]=par[par[x]]; x=par[x]
    return x
def u(x,y,why):
    rx,ry=f(x),f(y)
    if rx!=ry: par[ry]=rx
    motivo[(min(x,y),max(x,y))]=why
motivo={}
for key,why in ((lambda r: D(g(r,'F'))[:8] if len(D(g(r,'F')))==14 else None,'CNPJ'),
                (lambda r: ('pos',g(r,'A')),'posição'),
                (lambda r: ('nome',g(r,'C'),nuc(g(r,'D'))) if len(nuc(g(r,'D')).replace(' ',''))>=4 else None,'nome+UF')):
    b=C.defaultdict(list)
    for n,r in rows.items():
        k=key(r)
        if k: b[k].append(n)
    for k,v in b.items():
        for x in v[1:]: u(v[0],x,why)
cl=C.defaultdict(list)
for n in rows: cl[f(n)].append(n)
cl={k:sorted(v) for k,v in cl.items() if len(v)>1}
print('grupos', len(cl), 'linhas', sum(len(v) for v in cl.values()))
out=[]
for k,v in sorted(cl.items()):
    whys=sorted({w for (x,y),w in motivo.items() if x in v and y in v})
    out.append(dict(linhas=v, motivo=whys))
    print('---', whys)
    for n in v:
        r=rows[n]; print('  ', n, g(r,'A'), g(r,'C'), str(g(r,'D'))[:32], '|', str(g(r,'E'))[:28], '|', D(g(r,'F'))[:8] or '-', '| L',g(r,'L'),'AJ',g(r,'AJ'), g(r,'J'), g(r,'AL'), g(r,'AK').date() if g(r,'AK') else None, '|', g(r,'AN'))
json.dump(out, open('dups.json','w'), indent=0)
