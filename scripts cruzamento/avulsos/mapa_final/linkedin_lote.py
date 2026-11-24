"""Busca de pessoas no LinkedIn (24/11/2026): monta o lote de ações do navegador embutido (browser_batch) para as buscas a..b
de dados/contatos_2311/buscas.json. Cada ação chama __n (extrai a página atual e vai para a próxima) e a última descarrega
o lote com __d (filtrado pela empresa). Nomes trocados na URL: dicionário ov do linkedin_nomes.py.
Uso: python linkedin_lote.py <buscas.json> <a> <b>"""
import json,sys,os,urllib.parse as u
B=json.load(open(sys.argv[1],encoding='utf-8'))
src=open(os.path.join(os.path.dirname(os.path.abspath(__file__)),'linkedin_nomes.py'),encoding='utf-8').read()
ov=eval(src.split('ov=')[1].split('\n')[0])
ov.update({'s025':'ALE Combustíveis','s026':'Diniz','s036':'Help!'})
a,b=int(sys.argv[2]),int(sys.argv[3])
K='&keywords=CEO%20OR%20presidente%20OR%20%22diretor%20geral%22%20OR%20compras%20OR%20suprimentos%20OR%20supply'
L=B[a-1:b]; nm=lambda i,n: ov.get(i,n)
acts=[{"name":"navigate","input":{"tabId":"seed","url":"https://www.linkedin.com/search/results/people/?company="+u.quote(nm(*L[0]),safe='')+K}}]
for j,(i,n) in enumerate(L):
    nx=nm(*L[j+1]) if j+1<len(L) else ''
    acts.append({"name":"javascript_tool","input":{"action":"javascript_exec","tabId":"seed","text":"await eval(localStorage.getItem('__n'))('%s',%s)"%(i,json.dumps(nx,ensure_ascii=False))}})
acts.append({"name":"javascript_tool","input":{"action":"javascript_exec","tabId":"seed","text":"eval(localStorage.getItem('__d'))(%s)"%json.dumps([[i,nm(i,n)] for i,n in L],ensure_ascii=False)}})
print(json.dumps(acts,ensure_ascii=False))
