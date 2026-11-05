"""Confere se cada empresa da aba de pesquisa "Alessandra Dayan" (11 tabelas: EMPRESA/UF/LOJAS/BASE) esta no mapeamento
do MAPA MERCADO. Ordem: 1) referencia da celula na coluna ORIGEM das abas FINAL ("Alessandra Dayan: H5");
2) nome igual (nucleo, sem palavras de tipo) em alguma aba de mapeamento; 3) nucleo contido como palavras inteiras.
Decisoes a mao (achado/nao achado) em dados/dec_alessandra.json. Saida: dados/alessandra_cruzamento.json e .txt.
Uso: python confere_alessandra.py <MAPA.xlsx>"""
import os as _os, sys as _sys
_sys.argv[1:] = [_os.path.abspath(a) for a in _sys.argv[1:]]
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'dados'))
import re, json, unicodedata, collections, openpyxl
from openpyxl.utils import get_column_letter as L
MAPA = _sys.argv[1]
TIPO = re.compile(r'\b(LTDA|S ?A|SA|EIRELI|ME|EPP|GRUPO|REDE|REDES|LOJAS?|DROGARIAS?|FARMACIAS?|DROGA|SUPERMERCADOS?|SUPER|HIPER|ATACADO|ATACAREJO|'
                  r'ESTACIONAMENTOS?|PARKING|POSTOS?|AUTO POSTO|COMERCIO|COMERCIAL|E|DE|DO|DA|DOS|DAS|THE|CIA|BRASIL|FRANQUIAS?|OTICAS?|OTICA|'
                  r'SEGUROS|IDIOMAS|LAVANDERIA|MATERIAIS|MATERIAL|CONSTRUCAO|HOME|CENTER|CONVENIENCIA)\b')
def norm(s):
    s = unicodedata.normalize('NFKD', str(s or '')).encode('ascii', 'ignore').decode().upper().replace('&', ' E ')
    s = re.sub(r'\(.*?\)', ' ', s)
    return ' '.join(re.sub(r'[^A-Z0-9 ]', ' ', s).split())
def nucleo(s):
    n = ' '.join(TIPO.sub(' ', norm(s)).split())
    return n or norm(s)
wb = openpyxl.load_workbook(MAPA, read_only=True)
# ---- aba de pesquisa: 11 tabelas
ws = wb['Alessandra Dayan']; rows = list(ws.iter_rows(values_only=True))
itens = []
for j, seg in enumerate(rows[0]):
    if not seg: continue
    for i, r in enumerate(rows[2:], start=3):
        if len(r) > j and r[j]: itens.append(dict(ref=f'{L(j + 1)}{i}', seg=seg, nome=str(r[j]).strip(), uf=r[j + 1], lojas=r[j + 2], base=r[j + 3]))
# ---- abas de mapeamento: (aba, linha do cabecalho, colunas de nome)
ABAS = [('ALIMENTACAO FINAL', 1), ('REDE DE LOJAS FINAL', 1), ('CONSTRUÇÃO FINAL', 1), ('POSTOS GAS & CONVENIENCIA FINAL', 1), ('PET SHOP  FINAL', 1),
        ('ESTACIONAMENTO FINAL', 1), ('"FARMACIA VF"', 6), ('"FARMACIA CONSOLIDADO"', 1), ('SUPERMERCADOS (reuniao)', 4), ('I.R.T.T.__Por bandeira (FINAL)', 4)]
NOMECOL = ['NOME FANTASIA', 'RAZÃO SOCIAL', 'GRUPO ADMINISTRADOR', 'REDE / ASSOCIAÇÃO', 'Nome fantasia', 'Razão social', 'Grupo', 'Bandeira', 'OBSERVAÇÃO']
refs = collections.defaultdict(list); nomes = []  # (aba, linha, texto)
for aba, h in ABAS:
    vals = wb[aba].iter_rows(min_row=h, values_only=True); head = next(vals)
    cols = [k for k, x in enumerate(head) if x in NOMECOL]; corig = [k for k, x in enumerate(head) if x == 'ORIGEM']
    for i, r in enumerate(vals, start=h + 1):
        for k in corig:
            for m in re.findall(r'Alessandra Dayan: ([A-Z]+\d+)', str(r[k] or '')): refs[m].append((aba, i, r[1]))
        for k in cols:
            if k < len(r) and r[k]: nomes.append((aba, i, str(r[k]), nucleo(r[k]) if head[k] != 'OBSERVAÇÃO' else ' ' + norm(r[k]) + ' '))
wb.close()
PERMITE = {'DROGARIAS': ['"FARMACIA VF"', '"FARMACIA CONSOLIDADO"'], 'ESTACIONAMENTOS': ['ESTACIONAMENTO FINAL'],
           'POSTOS DE GASOLINA': ['POSTOS GAS & CONVENIENCIA FINAL'], 'MATERIAL DE CONST.': ['CONSTRUÇÃO FINAL'],
           'ALIMENTAÇÃO': ['ALIMENTACAO FINAL', 'POSTOS GAS & CONVENIENCIA FINAL'], 'REDE DE LOJAS': ['REDE DE LOJAS FINAL', 'PET SHOP  FINAL', 'I.R.T.T.__Por bandeira (FINAL)'],
           'SUPERMERCADOS / ATACAREJO': ['SUPERMERCADOS (reuniao)', 'I.R.T.T.__Por bandeira (FINAL)'],
           'CONVENIÊNCIA / AUTOSSERVIÇO': ['POSTOS GAS & CONVENIENCIA FINAL', 'ALIMENTACAO FINAL'],
           'BELEZA / SAÚDE / BEM-ESTAR': ['REDE DE LOJAS FINAL', 'I.R.T.T.__Por bandeira (FINAL)'], 'ÓTICAS': ['REDE DE LOJAS FINAL', 'I.R.T.T.__Por bandeira (FINAL)'],
           'SERVIÇOS / FRANQUIAS': ['REDE DE LOJAS FINAL', 'ALIMENTACAO FINAL', 'CONSTRUÇÃO FINAL', 'PET SHOP  FINAL', 'I.R.T.T.__Por bandeira (FINAL)']}
def alternativas(nome):
    """nome principal e as marcas entre parenteses ou separadas por + / ;"""
    partes = re.split(r'[+/;()]|E', unicodedata.normalize('NFKD', nome).encode('ascii', 'ignore').decode().upper())
    alts = {nucleo(nome)} | {nucleo(x) for x in partes if len(nucleo(x).replace(' ', '')) >= 4}
    return {a for a in alts if a}
DEC = json.load(open('dec_alessandra.json', encoding='utf8')) if _os.path.exists('dec_alessandra.json') else {}
out = []
for it in itens:
    n = nucleo(it['nome']); achou = []; como = ''; ok = set(PERMITE[it['seg']]); alts = alternativas(it['nome'])
    if it['ref'] in refs: achou = refs[it['ref']]; como = 'referência da origem'
    if not achou:
        achou = [(a, i, t) for a, i, t, nn in nomes if a in ok and nn.strip() in alts]; como = 'nome igual' if achou else ''
    if not achou:
        pats = [re.compile(r'(^| )' + re.escape(x) + r'( |$)') for x in alts if len(x.replace(' ', '')) >= 5]
        achou = [(a, i, t) for a, i, t, nn in nomes if a in ok and any(p.search(nn) for p in pats)][:6]; como = 'nome contido (conferir)' if achou else ''
    d = DEC.get(it['ref'])
    if d: como = d[0]; achou = [tuple(x) for x in d[1]] if len(d) > 1 else []
    out.append(dict(it, nucleo=n, como=como or 'NÃO ACHADO', onde=[f'{a} linha {i} ({t})' for a, i, t in achou[:4]]))
json.dump(out, open('alessandra_cruzamento.json', 'w', encoding='utf8'), ensure_ascii=False, indent=0, default=str)
with open('alessandra_cruzamento.txt', 'w', encoding='utf8') as f:
    for o in out: f.write(f"{o['ref']:5s} {o['seg'][:14]:14s} {o['nome'][:38]:38s} | {o['como']:24s} | {'; '.join(o['onde'])[:160]}\n")
c = collections.Counter(o['como'] for o in out)
print(len(out), dict(c))
print('não achados por tabela:', dict(collections.Counter(o['seg'] for o in out if o['como'] == 'NÃO ACHADO')))
