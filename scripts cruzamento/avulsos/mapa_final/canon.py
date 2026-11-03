"""Agrupa as linhas de origem em redes unicas (por segmento) e gera os termos de busca na Receita."""
import json, re, unicodedata, collections
import os as _os
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'dados'))
recs=json.load(open('recs.json',encoding='utf8'))
# chave -> chave canonica (merges manuais, por segmento)
M={'ALIM':{'MC DONALDS':'MCDONALDS','BACCIO DI LATTE':'BACIO DI LATTE','DOMINOS PIZZA':'DOMINOS','KONI STORE':'KONI',
   'MAIS1':'MAIS1','MILKY MOO MILKSHAKES':'MILKY MOO','NUMBER ONE CHICKEN':'N1 CHICKEN','OAKBERRY ACAI BOWLS':'OAKBERRY',
   'OUTBACK ABBRACCIO AUSSIE':'OUTBACK','OUTBACK STEAKHOUSE':'OUTBACK','JIN JIN WOK REGIONAIS':'JIN JIN','KOPENHAGEN MS':'KOPENHAGEN',
   'SPOLETO PROPRIAS':'SPOLETO','IMC':'FRANGO ASSADO VIENA','PIZZA HUT NORDESTE REGIONAIS':'PIZZA HUT','VO ALZIRA':'FABRICA BOLO VO ALZIRA',
   'TEMAKERIA TEMAKERIA':'TEMAKERIA CIA','SHUSHIAKI':'SUSHIAKI','PIZZA PRIME NORTE':'PIZZA PRIME'},
   'LOJ':{'CAEDU OUTLET UNIDADES MODA P':'CAEDU','CASAS PERNAMBUCANAS':'PERNAMBUCANAS','RI HAPPY PBKIDS':'RI HAPPY','CVLB':'CASA VIDEO LE BISCUIT',
   'LE BISCUIT':'CASA VIDEO LE BISCUIT','ESPLANADA BRASIL':'ESPLANADA','TOP FAMA':'TOPFAMA','MEGAMODAS':'MEGA MODAS','FLAMBOYAN':'FLAMBOYAN MODAS'},
   'CONS':{}}
# linhas "de grupo" que se desdobram em redes que ja estao na lista (viram nota, nao linha)
LUMP={('ALIM','CRM'):['KOPENHAGEN','CHOCOLATES BRASIL CACAU'],('ALIM','TRIGO'):['SPOLETO','CHINA IN BOX','CASA PAO QUEIJO'],
      ('ALIM','HALIPAR'):['GRILETTO','MONTANA GRILL','JIN JIN','CROASONHO'],('ALIM','SFB'):['GRILETTO','MONTANA GRILL'],
      ('LOJ','AF'):['AUTHENTIC FEET','ARTWALK'],('LOJ','DI SANTINNI CAPODARTE'):['DI SANTINNI','CAPODARTE']}
for r in recs:
    r['ck']=M[r['seg']].get(r['k'],r['k'])
G=collections.OrderedDict()
for r in recs:
    if (r['seg'],r['ck']) in LUMP: continue
    G.setdefault((r['seg'],r['ck']),[]).append(r)
for (seg,ck),tgts in LUMP.items():
    for r in recs:
        if r['seg']==seg and r['ck']==ck:
            for t in tgts:
                G[(seg,t)].append(dict(r,lump=True))
SUF=r'\b(LTDA|S\.?A\.?|S/A|EIRELI|ME|EPP|LIMITADA)\b\.?'
BAD=re.compile(r'A CONFIRMAR|CONFIRMAR|^—$|^-$|VARIOS|VÁRIOS|NAO DISPONIVEL|NÃO DISPONÍVEL|INFORMACAO|INFORMAÇÃO|EM FORMATA|ESTRUTURA',re.I)
TR=str.maketrans({c:' ' for c in '!"#$%\'()*+,-./:;<=>?@[\]^_`{|}~’'})
def norm(s):
    s=unicodedata.normalize('NFKD',s or '').encode('ascii','ignore').decode().upper().replace('&',' E ')
    return ' '.join(s.translate(TR).split())
def clean_f(s):
    s=re.sub(r'\((Sudeste|Sul|Norte|Nordeste|Centro-Oeste)\)','',s or '',flags=re.I)
    s=re.sub(r'\(.*?\)','',s)
    return s
GEN={'LOJAS','LOJA','REDE','GRUPO','DE','DA','DO','E','THE','CAFE','SORVETES','PIZZA','PIZZARIA','BURGER','SUSHI','BOLOS','ACAI'}
out=[]
for (seg,ck),rs in G.items():
    terms=set()
    for r in rs:
        if r.get('lump'): continue
        for s in re.split(r'\s*/\s*|\s+\+\s+',clean_f(r['fant'])):
            n=norm(s)
            if n: terms.add(n)
            w=n.split()
            if w and w[0] in ('LOJAS','LOJA','REDE','GRUPO','THE') and len(w)>1: terms.add(' '.join(w[1:]))
        for s in re.split(r'\s*/\s*|\s+—\s+|;',r['raz'] or ''):
            if not s or BAD.search(s): continue
            n=norm(re.sub(SUF,'',s,flags=re.I))
            if len(n)>=5: terms.add(n)
    terms={t for t in terms if t and not (len(t.split())==1 and len(t)<4)}
    # variante colada ("MILKY MOO" -> "MILKYMOO")
    for t in list(terms):
        w=t.split()
        if 1<len(w)<=3 and all(x not in GEN for x in w): terms.add(''.join(w))
    cn=sorted({r['cnpj'] for r in rs if r['cnpj'] and not r.get('lump')})
    out.append(dict(seg=seg,ck=ck,terms=sorted(terms),cnpjs=cn,n=len(rs)))
json.dump(out,open('canon.json','w',encoding='utf8'),ensure_ascii=False,indent=0)
json.dump({f'{s}|{k}':[{kk:r[kk] for kk in ('aba','lin','fant','raz','grupo','lojas','uf','segm','cnpj_src','cnpj','dv','base')}|{'lump':r.get('lump',False)} for r in rs] for (s,k),rs in G.items()},
          open('groups.json','w',encoding='utf8'),ensure_ascii=False,indent=0)
print(collections.Counter(o['seg'] for o in out))
print(sum(len(o['terms']) for o in out),'termos', sum(len(o['cnpjs']) for o in out),'cnpjs')
