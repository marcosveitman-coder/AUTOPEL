import json, re, unicodedata
import os as _os
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'dados'))
recs=json.load(open('recs.json',encoding='utf8'))
def nk(s):
    s=unicodedata.normalize('NFKD',s).encode('ascii','ignore').decode().upper()
    s=re.sub(r'\(.*?\)',' ',s)
    s=s.replace('&',' E ').replace("'",'').replace('’','')
    w=re.findall(r'[A-Z0-9]+',s)
    stop={'LOJAS','LOJA','REDE','GRUPO','DE','DA','DO','E','THE','REGIONAL','OPERACOES','SORVETES','CAFE'}
    w2=[x for x in w if x not in stop] or w
    return ' '.join(w2)
for r in recs: r['k']=nk(r['fant'])
json.dump(recs,open('recs.json','w',encoding='utf8'),ensure_ascii=False,indent=0)
for seg in ['ALIM','LOJ','CONS']:
    print('=====',seg)
    for r in sorted([r for r in recs if r['seg']==seg],key=lambda r:r['k']):
        print(f"{r['k'][:28]:28} | {r['fant'][:40]:40} | {r['aba'][:6]} {r['lin']} | {r['lojas'][:14]} | {r['uf'][:10]} | {r['cnpj']} {'' if r['dv'] is None else ('ok' if r['dv'] else 'DV!')}")
