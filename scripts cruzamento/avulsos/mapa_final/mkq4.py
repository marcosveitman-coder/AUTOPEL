"""4a consulta: redes ainda sem CNPJ. Usa o nucleo do nome (sem palavras de tipo) para achar a franqueadora/matriz
pelo nome fantasia e pela razao social (inclusive '<nucleo> FRANCHISING/FRANQUIAS')."""
import os as _os
_AQUI = _os.path.dirname(_os.path.abspath(__file__))
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_AQUI, 'dados'))
import json, re, unicodedata
dec = {}
for f in ['dec_alim1.json', 'dec_2.json', 'dec_3.json', 'dec_4.json']: dec.update(json.load(open(f, encoding='utf8')))
G = json.load(open('groups.json', encoding='utf8'))
CANON = {f"{o['seg']}|{o['ck']}": o for o in json.load(open('canon.json', encoding='utf8'))}
COMP = json.load(open('complementos.json', encoding='utf8'))
NOMES = json.load(open('nomes.json', encoding='utf8'))
TR = str.maketrans({c: ' ' for c in '!"#$%\'()*+,-./:;<=>?@[\\]^_`{|}~’'})
def norm(s):
    s = unicodedata.normalize('NFKD', s or '').encode('ascii', 'ignore').decode().upper().replace('&', ' E ')
    return ' '.join(s.translate(TR).split())
GEN = set('''PIZZA PIZZARIA PIZZAS PIZZERIA BURGER BURGUER BURGERS BURGUERS HAMBURGUERIA HAMBURGUERES SUSHI SUSHIBAR BAR TEMAKERIA TEMAKI
CAFE CAFES ESPECIAIS COFFEE BOLOS BOLO CASEIROS CASEIRO ARTESANAIS SORVETES SORVETE LANCHES LANCHONETE ACAI ACAITERIA DELIVERY EXPRESS EXPRESSA
RESTAURANTE GRILL STEAKHOUSE LOJAS LOJA REDE GRUPO HOME CENTER MATERIAIS MATERIAL CONSTRUCAO TINTAS MODAS MODA CALCADOS MOVEIS COLCHOES
DE DA DO DAS DOS E THE O A EM NO NA FAST FOOD FRANGO MILK SHAKES SHAKE CHOCOLATES CHOCOLATE COOKIES DOCES DOCERIA FOCACCERIA TRATTORIA
PASTA PANQUECAS JAPANESE RAMEN POKE CIA MR SR SANDUICHES PASTEIS ESFIHA BISTRO MAT REGIONAL OPERACOES NORTE SUL NORDESTE SUDESTE CENTRO OESTE
COMERCIO LTDA ATACADO DISTRIBUIDORA BRASIL'''.split())
def nucleo(n):
    w = [x for x in n.split() if x not in GEN]
    return ' '.join(w)
itens = []
for k, v in dec.items():
    if v[0] not in ('NC', 'NF', 'W'): continue
    mem = [x for x in G.get(k, []) if not x['lump']]
    nomes = {norm(re.sub(r'\(.*?\)', '', x['fant'])) for x in mem} | {norm(NOMES.get(k, ''))}
    itens.append((k, nomes, k.split('|')[0], v[0]))
for c in COMP:
    if not c.get('cnpj'):
        itens.append((c['seg'] + '|+' + norm(c['fant']), {norm(c['fant'])}, c['seg'], 'COMP'))
Q = []
for k, nomes, seg, st in itens:
    t = set()
    for n in nomes:
        if not n: continue
        t.add(n); t.add(n.replace(' ', ''))
        nc = nucleo(n)
        if nc and len(nc.replace(' ', '')) >= 4:
            t.add(nc)
            if ' ' in nc: t.add(nc.replace(' ', ''))
            for suf in ('FRANCHISING', 'FRANQUIAS', 'FRANQUEADORA', 'FRANQUIA', 'PARTICIPACOES', 'HOLDING'):
                t.add(nc + ' ' + suf)
    t = sorted(x for x in t if len(x.replace(' ', '')) >= 4)
    Q.append(dict(id=k, terms=t, cnpjs=[], st=st))
json.dump(Q, open('q4.json', 'w', encoding='utf8'), ensure_ascii=False)
print(len(Q), 'redes', sum(len(q['terms']) for q in Q), 'termos')
