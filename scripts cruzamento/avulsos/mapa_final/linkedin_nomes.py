import json,sys,urllib.parse as u
B=json.load(open(sys.argv[1],encoding='utf-8'))
ov={'s007':'Ortobom','s020':'CVC','s026':'Óticas Diniz','s028':"Bob's",'s033':'Wizard','s034':'Oakberry','s037':'Touti','s052':'60 Minutos','s056':'Milky Moo','s059':'RE/MAX Brasil','s060':'Linx','s061':'5àsec','s066':'Cheirin Bão','s071':'Fast4You','s075':'Mais1 Café','s078':'Rede Smart','s095':'Lava e Leva','s104':'Hope Lingerie','s114':'Vó Alzira','s119':'Instituto Mix','s130':'Supera Ginástica','s133':'Acqua Lavanderia','s141':'Farmelhor','s145':'Gio Estética','s154':'Hirota','s156':'Number One','s157':'OMO Lavanderia','s158':'Bauducco','s165':'Lave & Pegue','s171':'Maislaser','s173':'Laser Fast','s176':'Amigão','s183':'Maria Pitanga','s188':'Mr Shake','s198':'Gela Boca','s199':'Mr. Mix','s219':'Granado','s223':'Pedpharma','s229':'Maria Gasolina','s232':'Hortifruti','s237':'Hirota Food','s240':'Bamboo Brasil','s241':'Golden Mix','s243':'Iccar','s254':'Extrabom'}
a,b=int(sys.argv[2]),int(sys.argv[3])
K='&keywords=CEO%20OR%20presidente%20OR%20%22diretor%20geral%22%20OR%20compras%20OR%20suprimentos%20OR%20supply'
acts=[];ks=[]
for i,n in B[a-1:b]:
    acts.append({"name":"navigate","input":{"tabId":"seed","url":"https://www.linkedin.com/search/results/people/?company="+u.quote(ov.get(i,n),safe='')+K}})
    acts.append({"name":"javascript_tool","input":{"action":"javascript_exec","tabId":"seed","text":"await eval(localStorage.getItem('__w'))('%s')"%i}})
    ks.append([i,ov.get(i,n)])
acts.append({"name":"javascript_tool","input":{"action":"javascript_exec","tabId":"seed","text":"eval(localStorage.getItem('__d'))(%s)"%json.dumps(ks)}})
print(json.dumps(acts,ensure_ascii=False))
