import json, re, unicodedata
import os as _os
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'dados'))
from openpyxl.utils import column_index_from_string as ci
d=json.load(open('src.json',encoding='utf8'))
K={n[:8]:n for n in d}
def g(r,c):
    j=ci(c)-1
    return str(r[j]).strip() if j<len(r) and r[j]!='' else ''
def dv_ok(c):
    c=re.sub(r'\D','',c)
    if len(c)!=14 or len(set(c))==1: return False
    def calc(b,w):
        s=sum(int(x)*y for x,y in zip(b,w)); r=s%11; return '0' if r<2 else str(11-r)
    w1=[5,4,3,2,9,8,7,6,5,4,3,2]; w2=[6]+w1
    return calc(c[:12],w1)==c[12] and calc(c[:13],w2)==c[13]
recs=[]
def add(seg,aba,lin,fant,raz,grupo,lojas,uf,segm,cnpj,base=''):
    c=re.sub(r'\D','',cnpj or '')
    recs.append(dict(seg=seg,aba=aba,lin=lin,fant=fant,raz=raz,grupo=grupo,lojas=lojas,uf=uf,segm=segm,cnpj_src=(cnpj or '').strip(),
                     cnpj=c if len(c)==14 else '', dv=dv_ok(c) if len(c)==14 else None, base=base))
# ALIMENTACAO (...)
rows=d[[n for n in d if n.startswith('ALIMENTA') and 'FINAL' not in n][0]]
for i,r in enumerate(rows):
    lin=i+1
    if lin<4: continue
    if g(r,'A')=='x':
        add('ALIM','ALIMENTAÇÃO (...)',lin,g(r,'B'),g(r,'C'),'',g(r,'D'),g(r,'E'),'',g(r,'G'))
    else:
        add('ALIM','ALIMENTAÇÃO (...)',lin,g(r,'B'),'',g(r,'C'),g(r,'D'),g(r,'E'),g(r,'F'),g(r,'G'))
A=d['Alessandra Dayan']
for i,r in enumerate(A):
    lin=i+1
    if lin<3: continue
    for seg,cs in [('CONS','V W X Y'),('ALIM','AC AD AE AF'),('LOJ','AJ AK AL AM')]:
        a,b,c_,dd=cs.split()
        if g(r,a): add(seg,'Alessandra Dayan',f'{a}{lin}',g(r,a),'','',g(r,c_),g(r,b),'','',g(r,dd))
rows=d['VAREJO (MAIORES)_RADES']
for i,r in enumerate(rows):
    lin=i+1
    if lin<2: continue
    add('LOJ','VAREJO (MAIORES)_RADES',lin,g(r,'B'),'',g(r,'C'),g(r,'H'),g(r,'F'),g(r,'D'),g(r,'Q'),g(r,'I'))
rows=d['VAREJO VESTUARIO (IA)']
for i,r in enumerate(rows):
    lin=i+1
    if lin<2: continue
    if g(r,'E').upper() in ('REDE',): continue
    add('LOJ','VAREJO VESTUARIO (IA)',lin,g(r,'E'),g(r,'F'),'',g(r,'H'),g(r,'I'),'',g(r,'G'))
json.dump(recs,open('recs.json','w',encoding='utf8'),ensure_ascii=False,indent=0)
import collections
print(collections.Counter(r['seg'] for r in recs))
print('com cnpj',sum(1 for r in recs if r['cnpj']),'dv ok',sum(1 for r in recs if r['dv']))
