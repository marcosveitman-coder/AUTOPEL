"""Monta as linhas da aba POSTOS GAS & CONVENIENCIA FINAL (segmento POST), no mesmo formato das outras abas FINAL.
Fontes: varredura da Receita por CNAE (r8.pkl: postos 4731-8/00, conveniencia 4729-6/02, distribuidoras 4681-8),
dados das empresas achadas (r9.pkl), tabela 3.17 do Anuario ANP 2026 (postos por bandeira, regiao e UF, 2025) e
decisoes a mao em dec_postos.json (distribuidoras, redes de conveniencia, grupos, nomes e exclusoes).
Saida: postos_rows.pkl (lista de linhas) e postos_view.txt (para revisar)."""
import os as _os
_AQUI = _os.path.dirname(_os.path.abspath(__file__))
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_AQUI, 'dados'))
import pickle, json, re, collections, unicodedata, csv, io, sys
R8 = pickle.load(open('r8.pkl', 'rb')); RA, FA = R8['raiz'], R8['fant']
D9 = pickle.load(open('r9.pkl', 'rb')); emp = D9['emp']
MUN = pickle.load(open('mun.pkl', 'rb'))
DEC = json.load(open('dec_postos.json', encoding='utf8'))
BAPI = json.load(open('brasilapi.json', encoding='utf8'))
porraiz = collections.defaultdict(list)
for x in D9['est']: porraiz[x[0]].append(x)
for pk in ['r1.pkl', 'r3.pkl', 'r4.pkl', 'r6.pkl']:  # linhas ja vistas (BR Mania, AmPm, Oxxo, Raizen...)
    D = pickle.load(open(pk, 'rb'))
    for k, v in D['emp'].items(): emp.setdefault(k, v)
    for x in D['est']:
        if x not in porraiz[x[0]]: porraiz[x[0]].append(x)
SIT = {'01': 'NULA', '02': 'ATIVA', '03': 'SUSPENSA', '04': 'INAPTA', '08': 'BAIXADA'}
REG = {}
for g, ufs in [('NORTE', 'AC AP AM PA RO RR TO'), ('NORDESTE', 'AL BA CE MA PB PE PI RN SE'), ('CENTRO-OESTE', 'DF GO MT MS'),
               ('SUDESTE', 'ES MG RJ SP'), ('SUL', 'PR RS SC')]:
    for u in ufs.split(): REG[u] = g
TR = str.maketrans({c: ' ' for c in '!"#$%\'()*+,-./:;<=>?@[\\]^_`{|}~'})
def norm(s):
    s = unicodedata.normalize('NFKD', s or '').encode('ascii', 'ignore').decode().upper().replace('&', ' E ')
    return ' '.join(s.translate(TR).split())
def dv(c):
    def k(b, w):
        s = sum(int(x) * y for x, y in zip(b, w)); r = s % 11; return '0' if r < 2 else str(11 - r)
    w1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    return len(c) == 14 and k(c[:12], w1) == c[12] and k(c[:13], [6] + w1) == c[13]
def matriz(r):
    """linha da matriz de uma raiz (ou None)."""
    m = [x for x in porraiz.get(r, []) if x[3] == '1']
    if m: return m[0]
    for c, b in BAPI.items():
        if c[:8] == r and b[4] == '1':
            return (r, c[8:12], c[12:], '1', '', b[3], '', '', '', '', b[1], b[2], '')
    return None
def ativos(r, cn=('4731800',)):
    return sum(1 for x in porraiz.get(r, []) if x[5] == '02' and x[8] in cn)
def limpa(f):
    f = re.sub(r'\b(FILIAL|LOJA|UNIDADE|POSTO|AUTO POSTO)\s*\d+\b', r'\1', f)
    f = re.sub(r'\s+\d+$', '', f).strip()
    return f.title() if f else ''
RUIDO = {'FILIAL', 'LOJA', 'UNIDADE', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'MATRIZ'}
TIPO_RAZ = re.compile(r'(LTDA|S\.?\s?A\.?|S/A|EIRELI|ME|EPP|EM RECUPERACAO JUDICIAL|COMERCIO( E DISTRIBUICAO)? DE (DERIVADOS DE PETROLEO|COMBUSTIVEIS|PETROLEO)( E (LUBRIFICANTES|SERVICOS|DERIVADOS))?|COMBUSTIVEIS|DERIVADOS DE PETROLEO|COMERCIAL|PETROLEO)')
def nome_rede(raz, fans):
    """nome da rede: 'REDE ...' da razao, parte comum dos nomes das unidades, ou a razao sem as palavras de tipo."""
    m = re.match(r'^(REDE (?:DE POSTOS )?)([A-Z0-9 ]+?)\s+(LTDA|S\.?A|S/A|COMERCIO|COMBUSTIVEIS|GESTAO)', raz)
    if m: return ('Rede ' + m.group(2)).title()
    toks = [[w for w in f.split() if not w.isdigit()] for f in fans if f]
    if len(toks) >= 2:
        best = []
        for i in range(len(toks)):
            for j in range(i + 1, len(toks)):
                p = []
                for a, b in zip(toks[i], toks[j]):
                    if a != b: break
                    p.append(a)
                if len(p) > len(best): best = p
        while best and best[-1] in RUIDO: best.pop()
        if best and ' '.join(best) not in ('POSTO', 'AUTO POSTO', 'POSTO DE', 'REDE'): return ' '.join(best).title()
    core = ' '.join(TIPO_RAZ.sub(' ', raz).split())
    return core.title() if core else raz.title()
FIM = re.compile(r'(\s+(LTDA\.?|LIMITADA|S\.?\s?A\.?|S/A|EIRELI|ME|EPP|SA|E|DE|DO|DA|&|CIA|COMERCIO|COMERCIAL|REVENDEDORA|COMBUSTIVEIS?|DERIVADOS|PETROLEO|LUBRIFICANTES|SERVICOS|TRANSPORTES?|DISTRIBUICAO|DISTRIBUIDORA|PRODUTOS|PARA|VEICULOS))+$', re.I)
def limpa_nome(n):
    """tira do fim do nome as palavras de tipo de empresa (Ltda, Comercio de Combustiveis...)."""
    m = FIM.sub('', ' ' + n).strip()
    return m if len(m) >= 3 else n
def linha(r, fant, grupo, lojas, fonte, segm, origem, obs, nivel='C', rede=None):
    m = matriz(r) if r else None
    d = dict(seg='POST', fant=fant, raz='', grupo=grupo, lojas=lojas, cnpj='', cnae='', conf='', uf='', mun='', sit='', at=None,
             fonte=fonte, rede=rede, segm=segm, origem=origem, obs=obs, key='POST|' + norm(fant), uf_origem='', ref='', ref_desc='')
    if m:
        c = m[0] + m[1] + m[2]; assert dv(c), c
        d.update(raz=emp.get(r, ('',))[0].strip(), cnpj=c, cnae=m[8], uf=m[10], mun=MUN.get(m[11], m[11]), sit=SIT.get(m[5], m[5]),
                 conf=('CNPJ PROVÁVEL (CONFERIR)' if nivel == 'P' else 'CNPJ CONFIRMADO NA RECEITA') if m[5] == '02' else f'CNPJ {SIT.get(m[5], m[5])} NA RECEITA')
    else:
        d['conf'] = 'CNPJ NÃO CONFIRMADO'
    d['uf_sede'] = d['uf'] or DEC.get('uf_sede', {}).get(d['key'], ''); d['uf_sede_fonte'] = '' if d['uf'] else ('UF da sede pela pesquisa' if d['uf_sede'] else '')
    return d
# --- ANP: postos por bandeira e regiao (2025)
ANP = collections.defaultdict(collections.Counter)
_b = open('anp_dl/tabela3-17.csv', 'rb').read()
try: txt = _b.decode('utf-8-sig')
except UnicodeDecodeError: txt = _b.decode('latin-1')
for row in csv.reader(io.StringIO(txt), delimiter=';'):
    if len(row) < 4 or not row[3].strip().isdigit(): continue
    g = norm(row[0]).replace('REGIAO ', '').replace('CENTRO OESTE', 'CENTRO-OESTE')
    ANP[norm(re.sub(r'\d$', '', row[2].strip()))][g] += int(row[3])
pickle.dump(dict(ANP), open('anp_bandeiras.pkl', 'wb'))
out = []
usadas = set()
# 1) distribuidoras e redes de conveniencia decididas a mao
for d in DEC['manuais']:
    r = d['raiz']
    lojas, fonte = d.get('lojas'), d.get('fonte', '')
    if d.get('anp'):
        a = ANP[norm(d['anp'])]; assert a, d['anp']; lojas = sum(a.values()); fonte = 'ANP – Anuário 2026, tabela 3.17 (postos com a bandeira, 2025)'
        d['obs'] = '; '.join(x for x in [d.get('obs', ''), 'postos por região: ' + ', '.join(f"{g.title()} {a[g]:,}".replace(',', '.') for g in ['NORTE', 'NORDESTE', 'CENTRO-OESTE', 'SUDESTE', 'SUL'] if a[g])] if x)
    out.append(linha(r, d['fant'], d['grupo'], lojas, fonte, d['segm'], d['origem'], d.get('obs', ''), d.get('nivel', 'C'), ativos(r) or None))
    if r: usadas.add(r)
# 2) redes de postos (operadoras) achadas na Receita por CNAE
NOMES = DEC.get('nomes', {}); GRUPOS = DEC.get('grupos', {}); FORA = set(DEC.get('excluir', [])); NOTAS = DEC.get('notas', {})
cand = []
for (r, s), v in RA.items():
    if s != 'POSTO' or r in usadas or r in FORA: continue
    uf = max(v[1], key=v[1].get)
    cand.append((r, v[3], uf, v))
por_reg = collections.defaultdict(list)
for c in cand: por_reg[REG.get(c[2], '')].append(c)
esc = set()
for g, lst in por_reg.items():
    lst.sort(key=lambda c: -c[1])
    for i, c in enumerate(lst):
        if c[1] >= DEC['corte_brasil'] or (i < DEC['top_regiao'] and c[1] >= DEC['corte_regiao']): esc.add(c[0])
for r, n, uf, v in sorted([c for c in cand if c[0] in esc], key=lambda c: -c[1]):
    raz = emp.get(r, ('',))[0].strip()
    fan = [f for f, k in v[2] if f and not re.search(r'COMBUSTIVEIS? (E|PARA)|LTDA|EIRELI', f)]
    nome = NOMES.get(r) or limpa_nome(nome_rede(raz, [f for f, k in v[2]]))
    ufs = ', '.join(f'{u} {k}' for u, k in sorted(v[1].items(), key=lambda x: -x[1])[:5])
    conv = RA.get((r, 'CONV'), (None, None, None, 0))[3]
    obs = '; '.join(x for x in [f'postos por UF: {ufs}', f'mais {conv} lojas de conveniência com CNPJ da empresa' if conv else '', NOTAS.get(r, '')] if x)
    segm = 'Cooperativa com postos' if re.search(r'COOPERATIVA|COOP\b', raz) else 'Rede de postos (operadora)'
    out.append(linha(r, nome, GRUPOS.get(r, nome), n, 'Receita: postos ativos da empresa (CNAE 4731-8/00, out/2026)', segm,
                     'Complemento: Receita Federal (rede de postos achada por CNAE)', obs, 'C', n))
# 3) redes de conveniencia com CNPJ proprio achadas por CNAE (que nao sao de operadora de postos ja listada)
ja = {d['cnpj'][:8] for d in out if d['cnpj']}
for (r, s), v in sorted(RA.items(), key=lambda x: -x[1][3]):
    if s != 'CONV' or r in ja or r in FORA or v[3] < DEC['corte_conv']: continue
    raz = emp.get(r, ('',))[0].strip()
    fan = [f for f, k in v[2] if f]
    nome = NOMES.get(r) or (limpa(fan[0]) if fan else '') or raz.title()
    ufs = ', '.join(f'{u} {k}' for u, k in sorted(v[1].items(), key=lambda x: -x[1])[:5])
    out.append(linha(r, nome, GRUPOS.get(r, nome), v[3], 'Receita: lojas de conveniência ativas da empresa (CNAE 4729-6/02, out/2026)',
                     'Rede de lojas de conveniência', 'Complemento: Receita Federal (rede de conveniência achada por CNAE)',
                     '; '.join(x for x in [f'lojas por UF: {ufs}', NOTAS.get(r, '')] if x), 'C', v[3]))
# nome repetido com CNPJ diferente: e outra empresa; o municipio da sede entra no nome
cont = collections.Counter(norm(d['fant']) for d in out)
for d in out:
    if cont[norm(d['fant'])] > 1 and d['segm'] != 'Distribuidora de combustível (bandeira)' and d['mun']: d['fant'] = f"{d['fant']} ({d['mun']}/{d['uf']})"; d['key'] = 'POST|' + norm(d['fant'])
# duplicidade
vistos = collections.Counter(norm(d['fant']) for d in out)
dup = [k for k, n in vistos.items() if n > 1]
cn = collections.Counter(d['cnpj'] for d in out if d['cnpj'])
print('linhas', len(out), '| nomes repetidos', dup, '| CNPJ repetido', [k for k, n in cn.items() if n > 1])
ORD = {'CNPJ CONFIRMADO NA RECEITA': 0, 'CNPJ PROVÁVEL (CONFERIR)': 1, 'CNPJ NÃO CONFIRMADO': 3}
out.sort(key=lambda d: (ORD.get(d['conf'], 2), -(d['lojas'] or 0)))
pickle.dump(out, open('postos_rows.pkl', 'wb'))
with open('postos_view.txt', 'w', encoding='utf8') as f:
    for i, d in enumerate(out, start=2):
        f.write(f"{i} | {d['fant'][:32]} | {d['raz'][:44]} | {d['grupo'][:22]} | {d['lojas']} | {d['cnpj']} | {d['uf']} {d['mun'][:14]} | {d['segm'][:22]} | {d['obs'][:90]}\n")
print(collections.Counter(d['segm'] for d in out), collections.Counter(REG.get(d['uf_sede'], '?') for d in out))
