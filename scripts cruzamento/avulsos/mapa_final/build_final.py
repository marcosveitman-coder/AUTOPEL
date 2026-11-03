"""Monta as linhas das abas ALIMENTACAO FINAL, REDE DE LOJAS FINAL e CONSTRUCAO FINAL.
Junta: redes da origem (sem duplicidade) + decisoes de CNPJ (dec_*.json) + complementos (ANAMACO, IRTT, Receita).
Todo CNPJ escrito vem da copia local da Receita (razao, situacao, UF, municipio, matriz)."""
import os as _os
_AQUI = _os.path.dirname(_os.path.abspath(__file__))
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_AQUI, 'dados'))
import json, pickle, re, collections, unicodedata, sys
G = json.load(open('groups.json', encoding='utf8'))
CANON = json.load(open('canon.json', encoding='utf8'))
DEC = {}
for f in ['dec_alim1.json', 'dec_2.json', 'dec_3.json', 'dec_4.json']:
    DEC.update(json.load(open(f, encoding='utf8')))
# 2a rodada: [cnpj, grupo, nota, nivel C/P]
DEC5 = {k: v for k, v in json.load(open('dec_5.json', encoding='utf8')).items() if not k.startswith('_')}
NIVEL = {}
for k, v in DEC5.items():
    NIVEL[k] = v[3]
    if '|+' not in k: DEC[k] = v[:3]
# 3a rodada (03/11): [cnpj, nivel C/P/U/NC/NF/NOTA, nota, UF da sede]; C/P entram como decisao de CNPJ, o resto so como nota
DEC6 = {k: v for k, v in json.load(open('dec_6.json', encoding='utf8')).items() if not k.startswith('_')}
for k, v in DEC6.items():
    if v[1] not in ('C', 'P'): continue
    velho = DEC.get(k, ['', '', ''])
    nota = '; '.join(x for x in [velho[2] if 'não existe' in (velho[2] or '') or 'BAIXADO' in (velho[2] or '') else '', v[2]] if x)
    NIVEL[k] = v[1]
    if '|+' in k: DEC5[k] = [v[0], '', v[2], v[1]]
    else: DEC[k] = [v[0], velho[1], nota]
COMP = json.load(open('complementos.json', encoding='utf8'))
ANA = json.load(open('anamaco.json', encoding='utf8'))
MUN = pickle.load(open('mun.pkl', 'rb'))
NOMES = json.load(open('nomes.json', encoding='utf8'))
BAPI = json.load(open('brasilapi.json', encoding='utf8'))
emp, rows_by = {}, collections.defaultdict(dict)
for pk in ['r1.pkl', 'r3.pkl', 'r4.pkl', 'r5.pkl', 'r6.pkl', 'r7.pkl']:
    D = pickle.load(open(pk, 'rb'))
    for k, v in D['emp'].items(): emp.setdefault(k, v)
    for r in D['est']: rows_by[r[0]][r[0] + r[1] + r[2]] = r
SIT = {'01': 'NULA', '02': 'ATIVA', '03': 'SUSPENSA', '04': 'INAPTA', '08': 'BAIXADA'}
NOMEADAS = {}
CNAE_SEG = {'ALIM': ('56', '10', '47'), 'LOJ': ('47', '14', '15', '95'), 'CONS': ('47', '46', '23', '7732')}
def _rede():
    """Lojas ativas com o nome da rede no nome fantasia, somando todos os CNPJs (franqueados incluidos)."""
    Q1 = {q['id']: q for q in json.load(open('q1.json', encoding='utf8'))}
    Q3 = {q['id']: q for q in json.load(open('q3.json', encoding='utf8'))}
    for pk, QQ in [('r1.pkl', Q1), ('r3.pkl', Q3)]:
        D = pickle.load(open(pk, 'rb')); est = D['est']
        fh = collections.defaultdict(set)
        for i, t in D['hits_fant']: fh[t].add(i)
        for k, q in QQ.items():
            if k == 'RAIZES': continue
            terms = q['terms']
            if pk == 'r1.pkl':
                ft = set()
                for x in G.get(k, []):
                    if x['lump']: continue
                    for p in re.split(r'\s*/\s*|\s+\+\s+', re.sub(r'\(.*?\)', '', x['fant'])):
                        n = norm(p)
                        if n: ft.add(n)
                        w = n.split()
                        if len(w) > 2 and w[0] in ('LOJAS', 'LOJA', 'REDE', 'GRUPO', 'THE'): ft.add(' '.join(w[1:]))
                ftc = {f.replace(' ', '') for f in ft}
                terms = [t for t in terms if t in ft or t.replace(' ', '') in ftc] or terms[:1]
            pref = CNAE_SEG[k.split('|')[0]]
            cs = {est[i][0] + est[i][1] + est[i][2] for t in terms for i in fh.get(t, ()) if est[i][5] == '02' and est[i][8].startswith(pref)}
            if pk == 'r1.pkl': NOMEADAS[k] = len(cs)
            elif cs: NOMEADAS[k] = len(cs)
TR = str.maketrans({c: ' ' for c in '!"#$%\'()*+,-./:;<=>?@[\\]^_`{|}~’'})
def norm(s):
    s = unicodedata.normalize('NFKD', s or '').encode('ascii', 'ignore').decode().upper().replace('&', ' E ')
    return ' '.join(s.translate(TR).split())
def fmt(c):
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}" if len(c) == 14 else c
def resolve(c):
    """CNPJ (14) ou raiz (8) -> linha do estabelecimento (matriz, se for raiz)."""
    r = c[:8]; rs = rows_by.get(r, {})
    if len(c) == 14 and c not in rs and c in BAPI:  # conferido na BrasilAPI
        b = BAPI[c]; emp.setdefault(r, (b[0], '', '0', ''))
        return (r, c[8:12], c[12:], b[4], '', b[3], '', '', '', '', b[1], b[2], ''), 1
    if not rs: return None, 0
    at = sum(1 for x in rs.values() if x[5] == '02')
    if len(c) == 14:
        return rs.get(c), at
    m = [x for x in rs.values() if x[3] == '1']
    return (m[0] if m else None), at
MESES = {'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6, 'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12}
def data_rank(base):
    b = (base or '').lower().strip()
    m = re.match(r'([a-z]{3})/(\d{2,4})', b)
    if m:
        y = int(m.group(2)); y = y + 2000 if y < 100 else y
        return (y, MESES.get(m.group(1), 0))
    m = re.search(r'(20\d\d)', b)
    if m: return (int(m.group(1)), 0 if 'até' not in b else -1)
    return (0, 0)
def num(s):
    s = str(s or '').replace('.', '')
    m = re.search(r'\d+', s)
    return int(m.group(0)) if m else None
PRIO = {'Alessandra Dayan': 3, 'VAREJO (MAIORES)_RADES': 2, 'ALIMENTAÇÃO (...)': 1, 'VAREJO VESTUARIO (IA)': 0}
def escolhe_lojas(mem):
    best = None
    for x in mem:
        if x.get('lump'): continue
        n = num(x['lojas'])
        if n is None or n == 0: continue
        aba = x['aba']
        if aba == 'ALIMENTAÇÃO (...)' and str(x['lin']).isdigit() and int(x['lin']) >= 97: pr = -1
        else: pr = PRIO.get(aba, 0)
        key = (data_rank(x['base']), pr)
        if best is None or key > best[0]:
            aprox = bool(re.search(r'[~+]|mais de|cerca|aprox|–|-| a ', str(x['lojas']).lower()))
            best = (key, n, f"{aba}" + (f" ({x['base']})" if x['base'] else '') + (' - aproximado' if aprox else ''))
    return (best[1], best[2]) if best else (None, '')
def origem_txt(mem):
    out = collections.OrderedDict()
    for x in mem:
        a = x['aba']
        lin = str(x['lin'])
        out.setdefault(a, []).append(('linha ' if lin.isdigit() else '') + lin + (' (linha de grupo)' if x.get('lump') else ''))
    return '; '.join(f"{a}: {', '.join(v)}" for a, v in out.items())
def nome_fantasia(mem, ck):
    for x in mem:
        if not x.get('lump') and x['aba'] != 'Alessandra Dayan':
            s = re.sub(r'\s*\((Sudeste|Sul|Norte|Nordeste|Centro-Oeste)\)', '', x['fant']).strip()
            return s
    s = mem[0]['fant']
    return s.title() if s.isupper() else s
def linha(seg, fant, grupo, cnpj_dec, lojas, fonte_lojas, origem, obs, segm='', status_forcado=None, rede_nome=None, nivel=None):
    row = dict(seg=seg, fant=fant, grupo=grupo, lojas=lojas, fonte=fonte_lojas, origem=origem, obs=obs, segm=segm,
               raz='', cnpj='', uf='', mun='', sit='', at=None, conf='', rede=rede_nome)
    if status_forcado == 'NF':
        row['conf'] = 'REDE NÃO ENCONTRADA NA RECEITA'
        row['obs'] = (obs + '; ' if obs else '') + 'nenhum estabelecimento com esse nome na Receita: conferir se a rede existe'
        return row
    if not cnpj_dec or not cnpj_dec[0].isdigit():
        row['conf'] = 'CNPJ NÃO CONFIRMADO'
        return row
    m, at = resolve(cnpj_dec)
    if m is None:
        row['conf'] = 'CNPJ NÃO CONFIRMADO'
        row['obs'] = (obs + '; ' if obs else '') + f'CNPJ/raiz {fmt(cnpj_dec)} sem dados na Receita'
        return row
    if m[3] != '1' and len(cnpj_dec) == 14:
        mm = [x for x in rows_by.get(m[0], {}).values() if x[3] == '1' and x[5] == '02']
        if mm:
            row['obs'] = (row['obs'] + '; ' if row['obs'] else '') + f'o CNPJ indicado ({fmt(cnpj_dec)}) é de filial; aqui vai a matriz'
            m = mm[0]
    c = m[0] + m[1] + m[2]
    row.update(cnpj=c, cnae=m[8], raz=emp.get(m[0], ('',))[0].strip(), uf=m[10], mun=MUN.get(m[11], m[11]) if m[11].isdigit() else m[11], sit=SIT.get(m[5], m[5]), at=at,
               conf=('CNPJ PROVÁVEL (CONFERIR)' if nivel == 'P' else 'CNPJ CONFIRMADO NA RECEITA') if m[5] == '02' else f'CNPJ {SIT.get(m[5], m[5])} NA RECEITA')
    if m[3] != '1': row['obs'] = (row['obs'] + '; ' if row['obs'] else '') + 'CNPJ de filial (a raiz não tem matriz ativa)'
    return row
_rede()
SEGNOME = {'ALIM': 'ALIMENTAÇÃO', 'LOJ': 'REDE DE LOJAS', 'CONS': 'CONSTRUÇÃO'}
out = {'ALIM': [], 'LOJ': [], 'CONS': []}
notas_lump = collections.defaultdict(list)
for o in CANON:
    key = f"{o['seg']}|{o['ck']}"
    d = DEC.get(key)
    if d and d[0] == 'LUMP':
        for t in d[1].split('|'): notas_lump[f"{o['seg']}|{t}"].append(d[2])
for o in CANON:
    key = f"{o['seg']}|{o['ck']}"; seg = o['seg']
    d = DEC.get(key, ['NC', '', ''])
    if d[0] in ('LUMP', 'DUP'): continue
    mem = list(G[key])
    for o2 in CANON:  # linhas DUP somam a origem na linha alvo
        d2 = DEC.get(f"{o2['seg']}|{o2['ck']}")
        if d2 and d2[0] == 'DUP' and o2['seg'] == seg and d2[1] == o['ck']: mem += G[f"{o2['seg']}|{o2['ck']}"]
    lojas, fonte = escolhe_lojas(mem)
    obs = '; '.join([x for x in [d[2]] + notas_lump.get(key, []) if x])
    segm = next((x['segm'] for x in mem if x.get('segm')), '')
    r = linha(seg, NOMES.get(key) or nome_fantasia(mem, o['ck']), d[1] or '', d[0], lojas, fonte, origem_txt(mem), obs, segm,
              status_forcado='NF' if d[0] == 'NF' else None, nivel=NIVEL.get(key))
    r['key'] = key
    r['uf_origem'] = next((x['uf'][:2] for x in mem if re.match(r'^[A-Z]{2}\b', x['uf'] or '')), '')
    r['rede'] = None if key in ('LOJ|LIMITS GREGORY ANIMALE',) else NOMEADAS.get(key)
    if key in ('ALIM|SODEXO BRASIL', 'ALIM|GRSA') and r['at']:
        r['lojas'], r['fonte'] = r['at'], 'Receita: unidades ativas da empresa (out/2026)'
    if not r['lojas']:
        if r['rede']: r['lojas'], r['fonte'] = r['rede'], 'Receita: lojas ativas com o nome da rede (out/2026)'
        elif r['at']: r['lojas'], r['fonte'] = r['at'], 'Receita: estabelecimentos ativos da empresa (out/2026)'
    if seg == 'CONS' and o['ck'] in ANA['pos']: r['obs'] = (r['obs'] + '; ' if r['obs'] else '') + f"Ranking ANAMACO {ANA['pos'][o['ck']]}"
    if seg == 'CONS' and o['ck'] in ANA['exame'] and not r['lojas']:
        r['lojas'], r['fonte'] = ANA['exame'][o['ck']]
    out[seg].append(r)
COMP_FORA = {'CONS|+CONSTRULAR'}
for c in COMP:
    kc = c['seg'] + '|+' + norm(c['fant'])
    if kc in COMP_FORA: continue
    nv = None
    if kc in DEC5:
        v = DEC5[kc]; c = dict(c, cnpj=v[0], status=None, obs='; '.join(x for x in [c.get('obs', ''), v[2]] if x and 'não aparece' not in x)); nv = v[3]
    r = linha(c['seg'], c['fant'], c.get('grupo', ''), c.get('cnpj', ''), c.get('lojas'), c.get('fonte', ''), c['origem'], c.get('obs', ''),
              c.get('segm', ''), status_forcado=c.get('status'), nivel=nv)
    r['key'] = kc; r['uf_origem'] = ''
    if r['lojas'] is None and r['at']:
        r['lojas'], r['fonte'] = r['at'], 'Receita: estabelecimentos ativos da empresa (out/2026)'
    out[c['seg']].append(r)
# consultas usadas para o CNPJ de referencia: q4 (redes sem CNPJ) e q5 (3 redes da lista ABF)
QREF = {}
for pk, qf in [('r4.pkl', 'q4.json'), ('r5.pkl', 'q5.json')]:
    D = pickle.load(open(pk, 'rb')); F = collections.defaultdict(set)
    for i, t in D['hits_fant']: F[t].add(i)
    for q in json.load(open(qf, encoding='utf8')): QREF[q['id']] = (q, D['est'], F)
# UF da sede (para a regiao): Receita > pesquisa > aba de origem > UF com mais lojas da rede
SEDE_WEB = {'ALIM|MR FIT': ('SP', 'rede criada em Paulínia/SP em 2013 (pesquisa)'),
            'ALIM|+FRUTOS DE GOIAS': ('GO', 'sede e fábrica em Senador Canedo/GO (pesquisa); a fábrica não leva o nome da marca na Receita'),
            'LOJ|+PITICAS': ('SP', 'sede em Guarulhos/SP (pesquisa)'),
            'LOJ|+LIVRARIA NOBEL': ('SP', 'a antiga Livraria Nobel S.A. era de Embu das Artes/SP (Receita, baixada)')}
# nome comum: a unidade achada pode ser de outra empresa; sem referencia
REF_NAO = {'ALIM|SPERANZA', 'CONS|CAETANO', 'CONS|DESTRO', 'CONS|SAO JORGE', 'CONS|+CASA VERDE', 'CONS|+PADOVANI', 'LOJ|RABELO', 'LOJ|SELLER',
           'LOJ|FLORENZA', 'ALIM|GENERAL PRIME BURGER', 'LOJ|CROSBY', 'CONS|C C', 'ALIM|PONTO', 'ALIM|SAKURA SUSHI', 'ALIM|PASTELARIA GAUCHO',
           'ALIM|AU AU LANCHES', 'ALIM|TEMAKI LOUNGE', 'ALIM|PARMEGIANA EXPRESS', 'ALIM|BRAZ PIZZARIA', 'ALIM|BULLS HAMBURGUERIA', 'ALIM|LE BURGER',
           'ALIM|SUSHI PRACA', 'ALIM|CIA PIZZA', 'ALIM|HAI SUSHI'}
SUF = re.compile(r' (FRANCHISING|FRANQUIAS|FRANQUEADORA|FRANQUIA|PARTICIPACOES|HOLDING)$')
def cap(r):
    try: return float(emp.get(r, ('', '', '0', ''))[2].replace(',', '.'))
    except: return 0.0
for seg, rs in out.items():
    for r in rs:
        r.setdefault('ref', ''); r.setdefault('ref_desc', '')
        if r['cnpj'] or r['conf'].startswith('REDE NÃO') or r.get('key') in REF_NAO: continue
        qq = QREF.get(r.get('key'))
        if not qq: continue
        q, E4, F4 = qq
        nomes = {norm(re.sub(r'\(.*?\)', '', x['fant'])) for x in G.get(r['key'], []) if not x['lump']} | {norm(re.sub(r'\(.*?\)', '', r['fant']))}
        nomes = {n for n in nomes if n and (len(n.split()) >= 2 or len(n) >= 6)}
        nomes |= {n.replace(' ', '') for n in nomes if ' ' in n}
        full = [t for t in q['terms'] if t in nomes]
        PREF = {'ALIM': ('56', '10', '47'), 'LOJ': ('47', '14', '15'), 'CONS': ('474', '23')}[seg]
        at = [E4[i] for t in full for i in F4.get(t, ()) if E4[i][5] == '02' and E4[i][8].startswith(PREF)]
        if not at: continue
        r['uf_moda'] = collections.Counter(x[10] for x in at).most_common(1)[0][0]
        uf = r.get('uf_origem', '') or SEDE_WEB.get(r.get('key'), ('',))[0]
        at.sort(key=lambda x: (x[10] != uf if uf else False, x[3] != '1', -cap(x[0]), x[7] or '9'))
        x = at[0]
        r['ref'] = x[0] + x[1] + x[2]
        r['ref_desc'] = f"{emp.get(x[0], ('',))[0].strip()} ({MUN.get(x[11], x[11])}/{x[10]}, nome fantasia {x[4].strip()})"
        r['conf'] = 'CNPJ DE UMA UNIDADE DA REDE (FRANQUEADORA NÃO ACHADA)'
# 3a rodada: CNPJ de uma unidade escolhido a mao (U) e notas das redes que continuam sem CNPJ
for seg, rs in out.items():
    for r in rs:
        v = DEC6.get(r.get('key'))
        if not v or v[1] in ('C', 'P'): continue
        if v[1] == 'U':
            c = v[0]; x = rows_by.get(c[:8], {}).get(c)
            assert x and x[5] == '02', ('unidade sem dado ou inativa', r['key'], c)
            r['ref'] = c
            r['ref_desc'] = f"{emp.get(x[0], ('',))[0].strip()} ({MUN.get(x[11], x[11])}/{x[10]}, nome fantasia {x[4].strip()})"
            r['conf'] = 'CNPJ DE UMA UNIDADE DA REDE (FRANQUEADORA NÃO ACHADA)'
            r['obs'] = '; '.join(x for x in [r['obs'], v[2]] if x)
        else:
            r['obs'] = '; '.join(x for x in [r['obs'], v[2]] if x and v[2] not in (r['obs'] or ''))
        if v[3]: SEDE_WEB[r['key']] = (v[3], 'UF da sede pela pesquisa de 03/11/2026')
NF_NOTA = 'busca refeita em 03/11/2026 pelo núcleo do nome e pela razão social: nada com o nome na UF da origem'
for r in out['ALIM']:
    if r['conf'].startswith('REDE NÃO') and r.get('key') not in DEC6: r['obs'] = '; '.join(x for x in [r['obs'], NF_NOTA] if x)
# UF da sede (para a regiao): Receita > pesquisa > aba de origem > UF com mais lojas da rede
for seg, rs in out.items():
    for r in rs:
        if r['uf']: r['uf_sede'], r['uf_sede_fonte'] = r['uf'], ''
        elif r.get('key') in SEDE_WEB: r['uf_sede'], r['uf_sede_fonte'] = SEDE_WEB[r['key']]
        elif r.get('uf_origem'): r['uf_sede'], r['uf_sede_fonte'] = r['uf_origem'], 'UF da sede pela aba de origem'
        elif r.get('uf_moda'): r['uf_sede'], r['uf_sede_fonte'] = r['uf_moda'], 'UF com mais lojas da rede na Receita (sede não confirmada)'
        else: r['uf_sede'], r['uf_sede_fonte'] = '', ''
# conferencia de duplicidade dentro de cada aba (mesmo CNPJ ou mesmo nome)
for seg, rs in out.items():
    vistos = collections.defaultdict(list)
    for r in rs:
        vistos[norm(r['fant'])].append(r)
    for k, v in vistos.items():
        if len(v) > 1: print('NOME REPETIDO', seg, k, [x['origem'][:40] for x in v])
    porc = collections.defaultdict(list)
    for r in rs:
        if r['cnpj']: porc[r['cnpj']].append(r['fant'])
    for k, v in porc.items():
        if len(v) > 1: print('CNPJ EM MAIS DE UMA REDE', seg, fmt(k), v)
ORD = {'CNPJ CONFIRMADO NA RECEITA': 0, 'CNPJ PROVÁVEL (CONFERIR)': 1, 'CNPJ DE UMA UNIDADE DA REDE (FRANQUEADORA NÃO ACHADA)': 2, 'CNPJ NÃO CONFIRMADO': 3, 'REDE NÃO ENCONTRADA NA RECEITA': 4}
for seg in out:
    out[seg].sort(key=lambda r: (ORD.get(r['conf'], 1), -(r['lojas'] or 0), norm(r['fant'])))
pickle.dump(out, open('final_rows.pkl', 'wb'))
for seg, rs in out.items():
    print(seg, len(rs), collections.Counter(r['conf'] for r in rs))
