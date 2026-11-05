"""Monta as linhas das abas PET SHOP  FINAL (segmento PET) e ESTACIONAMENTO FINAL (segmento EST), no mesmo formato das
outras abas FINAL.
Fontes:
- varredura da Receita por CNAE (r11.pkl, redes_pet_estac.py): pet shop 4789-0/04, veterinaria 7500-1/00, banho e tosa
  9609-2/08, hotel de animais 9609-2/07, farmacia veterinaria 4771-7/04 e estacionamento 5223-1/00;
- busca por nome na Receita (r12.pkl, busca_receita.py q12.json): razao social, matriz e unidades das redes citadas;
- origem: aba "Alessandra Dayan" (estacionamentos H-K), aba REDE DE LOJAS FINAL (redes pet), aba IRTT;
- imprensa e associacoes (pet_estac_fontes.json);
- decisoes a mao em dec_pet_estac.json (CNPJ de cada rede, grupo, nome, exclusoes, cortes).
Saida: pet_rows.pkl, estac_rows.pkl e pet_estac_view.txt (para revisar)."""
import os as _os
_AQUI = _os.path.dirname(_os.path.abspath(__file__))
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_AQUI, 'dados'))
import pickle, json, re, collections, unicodedata
R11 = pickle.load(open('r11.pkl', 'rb')); RA, FA = R11['raiz'], R11['fant']
R12 = pickle.load(open('r12.pkl', 'rb')); emp = R12['emp']
MUN = pickle.load(open('mun.pkl', 'rb'))
DEC = json.load(open('dec_pet_estac.json', encoding='utf8'))
BAPI = json.load(open('brasilapi.json', encoding='utf8'))
R13 = pickle.load(open('r13.pkl', 'rb'))  # razao e estabelecimentos das empresas candidatas (busca por raiz)
for _pk in ['r14.pkl', 'r15.pkl', 'r16.pkl']:  # candidatas que faltaram na r13
    if _os.path.exists(_pk):
        _rx = pickle.load(open(_pk, 'rb')); R13['emp'].update(_rx['emp']); R13['est'] = R13['est'] + _rx['est']
for k, v in R13['emp'].items(): emp.setdefault(k, v)
porraiz = collections.defaultdict(list); _vistos = set()
for x in R12['est'] + R13['est']:
    if x[:3] in _vistos: continue
    _vistos.add(x[:3]); porraiz[x[0]].append(x)
HITS = collections.defaultdict(set)  # termo -> estabelecimentos ativos com o termo no nome fantasia
for i, t in R12['hits_fant']:
    x = R12['est'][i]
    if x[5] == '02': HITS[t].add(x[:3] + (x[8],))
SIT = {'01': 'NULA', '02': 'ATIVA', '03': 'SUSPENSA', '04': 'INAPTA', '08': 'BAIXADA'}
REG = {}
for g, ufs in [('NORTE', 'AC AP AM PA RO RR TO'), ('NORDESTE', 'AL BA CE MA PB PE PI RN SE'), ('CENTRO-OESTE', 'DF GO MT MS'),
               ('SUDESTE', 'ES MG RJ SP'), ('SUL', 'PR RS SC')]:
    for u in ufs.split(): REG[u] = g
SEGS = {'PET': ['PET', 'VET', 'BANHO', 'HOTEL', 'VETMED'], 'EST': ['ESTAC']}
CNAES = {'PET': ('4789004', '7500100', '9609208', '9609207', '4771704'), 'EST': ('5223100',)}
NOME_SEG = {'PET': 'Receita: unidades ativas da empresa (pet shop, veterinária, banho e tosa, hotel e farmácia veterinária; out/2026)',
            'EST': 'Receita: estacionamentos ativos da empresa (CNAE 5223-1/00, out/2026)'}
TR = str.maketrans({c: ' ' for c in '!"#$%\'()*+,-./:;<=>?@[\\]^_`{|}~'})
def norm(s):
    s = unicodedata.normalize('NFKD', s or '').encode('ascii', 'ignore').decode().upper().replace('&', ' E ')
    return ' '.join(s.translate(TR).split())
def dv(c):
    def k(b, w):
        s = sum(int(x) * y for x, y in zip(b, w)); r = s % 11; return '0' if r < 2 else str(11 - r)
    w1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    return len(c) == 14 and c.isdigit() and k(c[:12], w1) == c[12] and k(c[:13], [6] + w1) == c[13]
def matriz(r):
    m = [x for x in porraiz.get(r, []) if x[3] == '1']
    if m: return m[0]
    for c, b in BAPI.items():
        if c[:8] == r and b[4] == '1':
            return (r, c[8:12], c[12:], '1', '', b[3], '', '', '', '', b[1], b[2], '')
    return None
def unidades(seg, r):
    """unidades ativas da raiz nos CNAEs do segmento (varredura r11) e as UFs."""
    n = 0; ufs = collections.Counter(); por = collections.Counter()
    for s in SEGS[seg]:
        v = RA.get((r, s))
        if v: n += v[3]; ufs.update(v[1]); por[s] += v[3]
    return n, ufs, por
def ufs_txt(ufs, rot='unidades'):
    return f'{rot} por UF: ' + ', '.join(f'{u} {k}' for u, k in ufs.most_common(5)) if ufs else ''
def linha(seg, r, fant, grupo, lojas, fonte, segm, origem, obs, nivel='C', rede=None, cnpj=None):
    """uma linha da aba. r = raiz (a matriz vira o CNPJ); cnpj = CNPJ escolhido quando nao e a matriz."""
    m = matriz(r) if r else None
    d = dict(seg=seg, fant=fant, raz='', grupo=grupo, lojas=lojas, cnpj='', cnae='', conf='', uf='', mun='', sit='', at=None,
             fonte=fonte, rede=rede, segm=segm, origem=origem, obs=obs, key=seg + '|' + norm(fant), uf_origem='', ref='', ref_desc='')
    if cnpj:
        e = [x for x in porraiz.get(cnpj[:8], []) if x[0] + x[1] + x[2] == cnpj]
        if e: m = e[0]
    if m:
        c = m[0] + m[1] + m[2]; assert dv(c), c
        d.update(raz=emp.get(m[0], ('',))[0].strip(), cnpj=c, cnae=m[8], uf=m[10], mun=MUN.get(m[11], m[11]),
                 sit=SIT.get(m[5], m[5]))
        if m[5] != '02': d['conf'] = f'CNPJ {SIT.get(m[5], m[5])} NA RECEITA'
        elif nivel == 'P': d['conf'] = 'CNPJ PROVÁVEL (CONFERIR)'
        elif nivel == 'U':
            d['ref'] = c; d['ref_desc'] = f"{d['raz']} ({d['mun']}/{d['uf']})"; d['conf'] = 'CNPJ DE UMA UNIDADE DA REDE (FRANQUEADORA NÃO ACHADA)'
        else: d['conf'] = 'CNPJ CONFIRMADO NA RECEITA'
    else:
        d['conf'] = 'REDE NÃO ENCONTRADA NA RECEITA' if nivel == 'NF' else 'CNPJ NÃO CONFIRMADO'
    d['uf_sede'] = d['uf'] if nivel != 'U' else ''
    d['uf_sede_fonte'] = ''
    return d
def com_nome(seg, termos, r=''):
    """unidades ativas, em qualquer CNPJ, com o nome da rede no nome fantasia e CNAE do segmento;
    sem nome fantasia nas unidades, vale a contagem das unidades ativas da propria empresa."""
    s = set()
    for t in termos:
        s |= {x for x in HITS.get(t, ()) if x[3] in CNAES[seg]}
    return len(s) or (unidades(seg, r)[0] if r else 0) or None
AGRO = re.compile(r'COOPERATIVA|\bCOOP|AGROPECUARI|AGRICOLA|ADUBO|RURAL|AGRO\b|AGROINDUS|LATICINIO|INSUMO|SEMENTE|FERTILIZ|PECUARI|ZOOTEC|NUTRICAO ANIMAL|DISTRIBUIDORA')
PETNOME = re.compile(r'(PET|PETS|ANIMAL|ANIMAIS|VET|VETERINARI\w*|ZOO|DOG|DOGS|CAO|CAES|BICHO|BICHOS|GATO|GATOS|RACAO|RACOES)')
PETFORTE = re.compile(r'(PET|PETS|DOG|DOGS|CAO|CAES|BICHO|BICHOS|GATO|GATOS)')
FORA_PET = re.compile(r'INSTITUTO|ASSOCIACAO|FUNDACAO|UNIVERSIDADE|UNIVERSITARI|FACULDADE|EDUCACIONAL|ENSINO|HOSPITAL GERAL|PREFEITURA|MUNICIPIO')
FORA_EST = re.compile(r'PATIO|LEILO|REMOCAO|GUINCHO|DEPOSITO|TRANSPORTE|TRANSPORTADORA|LOGISTICA|VISTORIA|LAVA ?JATO|HOTEL|CONDOMINIO|HOSPITAL|IGREJA|ASSOCIACAO')
ROT_SEG = {"PET": "pet shop", "VET": "veterinária", "BANHO": "banho e tosa", "HOTEL": "hotel", "VETMED": "farmácia veterinária"}
def nome_auto(r, raz, fans):
    """nome da rede: nome fantasia mais usado nas unidades (sem numero/bairro), ou a razao sem as palavras de tipo."""
    for f, k in fans:
        f = re.sub(r'\s+\d+.*$|\s+(FILIAL|LOJA|UNIDADE)\b.*$', '', f or '').strip()
        if len(f) >= 4 and not re.search(r'LTDA|EIRELI|S\.?A\.?$|^MATRIZ$|^ESTACIONAMENTO$|^PET SHOP$', f): return f.title()
    core = re.sub(r'\b(LTDA\.?|EIRELI|S\.?/?A\.?|ME|EPP|COMERCIO|COMERCIAL|SERVICOS?|E|DE|DO|DA|PRODUTOS|PARA|ANIMAIS|ADMINISTRACAO|ADMINISTRADORA|ESTACIONAMENTOS?|GARAGENS?|PARTICIPACOES)\b', ' ', raz)
    core = ' '.join(core.replace('-', ' ').split())
    return (core or raz).title()
def segm_auto(seg, por, raz):
    if seg == 'EST':
        return 'Estacionamento rotativo (zona azul)' if re.search(r'ZONA AZUL|ROTATIVO', raz) else 'Operadora de estacionamentos'
    s = por.most_common(1)[0][0] if por else 'PET'
    return {'PET': 'Pet shop (lojas próprias)', 'VET': 'Clínica / hospital veterinário', 'BANHO': 'Banho e tosa', 'HOTEL': 'Hotel / creche de animais',
            'VETMED': 'Agropet / farmácia veterinária'}[s]
out = {'PET': [], 'EST': []}
usadas = {'PET': set(), 'EST': set()}
# ---------- PET: redes citadas (REDE DE LOJAS FINAL, IRTT, Alessandra, ABF, imprensa)
LOJ = {norm(x['NOME FANTASIA']): x for x in json.load(open('loj_pet_rows.json', encoding='utf8'))}
for fant, cnpj, niv, grupo, lojas, fonte, segm, origem, obs, termos in DEC['PET']['manuais']:
    r = cnpj[:8] if cnpj else ''
    d = linha('PET', r, fant, grupo, lojas, fonte, segm, origem, obs, niv, com_nome('PET', termos, r if niv != 'U' else ''), cnpj or None)
    if niv == 'U' or not d['uf_sede']:
        d['uf_sede'] = DEC['PET']['sede'].get(fant, ''); d['uf_sede_fonte'] = 'UF da sede pela pesquisa' if d['uf_sede'] else ''
    x = LOJ.get(norm(fant))
    if x and lojas is None:  # mesma linha da REDE DE LOJAS FINAL: mesmo nº de lojas e fonte
        d['lojas'], d['fonte'] = x['Nº DE LOJAS'], x['FONTE DO Nº DE LOJAS']
    if x: d['_status'] = (x.get('PORTE') or '', x.get('STATUS AUTOPEL') or '')
    if r: usadas['PET'].add(r)
    out['PET'].append(d)
# ---------- EST: aba de origem (Alessandra Dayan H-K) + imprensa
ALE = {x['lin']: x for x in json.load(open('alessandra_estac.json', encoding='utf8'))}
por_lin = {}
for lin, fant, cnpj, niv, grupo, lojas, fonte, segm, obs, termos in DEC['EST']['origem']:
    o = ALE[lin]; orig = f'Alessandra Dayan: {lin}'
    if fant is None:  # junta na linha de outra rede (mesmo grupo)
        d = por_lin[cnpj]; d['origem'] += '; ' + orig; d['obs'] = '; '.join(x for x in [d['obs'], obs] if x); continue
    r = cnpj[:8] if cnpj else ''
    if lojas is None and o['lojas']: lojas, fonte = int(o['lojas']), f"Alessandra Dayan ({o['base']})"
    rede = com_nome('EST', termos, r)
    if lin in DEC['EST']['grupo_raizes']: rede = sum(unidades('EST', x)[0] for x in DEC['EST']['grupo_raizes'][lin])
    d = linha('EST', r, fant, grupo, lojas, fonte, segm, orig, obs, niv, rede, cnpj or None)
    if r:
        n, ufs, _ = unidades('EST', r)
        if n: d['obs'] = '; '.join(x for x in [d['obs'], f'Receita: {n} garagens ativas da empresa', ufs_txt(ufs, 'garagens')] if x)
        usadas['EST'].add(r)
    if not d['uf_sede'] and len(str(o['uf'])) == 2: d['uf_sede'] = o['uf']; d['uf_sede_fonte'] = 'UF da sede pela aba de origem'
    por_lin[lin] = d; out['EST'].append(d)
for fant, cnpj, niv, grupo, lojas, fonte, segm, origem, obs, termos in DEC['EST']['extra']:
    d = linha('EST', cnpj[:8] if cnpj else '', fant, grupo, lojas, fonte, segm, origem, obs, niv, com_nome('EST', termos), cnpj or None)
    if not d['uf_sede']: d['uf_sede'] = 'SP'; d['uf_sede_fonte'] = 'UF da sede pela pesquisa'
    out['EST'].append(d)
# ---------- complemento: redes achadas na Receita por CNAE (todas as regioes)
for seg in ['PET', 'EST']:
    cfg = DEC[seg]; FORA = set(cfg['excluir'])
    tot = {}
    for (r, s), v in RA.items():
        if s in SEGS[seg]: tot[r] = unidades(seg, r)
    cand = []
    for r, (n, ufs, por) in tot.items():
        if r in usadas[seg] or r in FORA or not n: continue
        raz = emp.get(r, ('',))[0].strip()
        if seg == 'PET':
            if FORA_PET.search(raz): continue
            # agropecuaria, cooperativa e loja rural ficam fora (a aba e de pet shop); fica a agropet com nome de pet
            if AGRO.search(raz) and not PETNOME.search(raz.replace('AGROPECUARI', '').replace('AGROVETERINARI', '')): continue
            if por.get('VETMED', 0) * 2 > n and not PETFORTE.search(raz): continue  # farmacia veterinaria sem nome de pet = loja agropecuaria
        else:
            fans = ' '.join(f for f, k in RA[(r, 'ESTAC')][2])
            if FORA_EST.search(raz + ' ' + fans): continue
        cand.append((r, n, ufs, por, raz))
    por_reg = collections.defaultdict(list)
    for c in cand: por_reg[REG.get(c[2].most_common(1)[0][0], '')].append(c)
    esc = set()
    for g, lst in por_reg.items():
        lst.sort(key=lambda c: -c[1])
        for i, c in enumerate(lst):
            if c[1] >= cfg['corte_brasil'] or (i < cfg['top_regiao'] and c[1] >= cfg['corte_regiao']): esc.add(c[0])
    for r, n, ufs, por, raz in sorted([c for c in cand if c[0] in esc], key=lambda c: -c[1]):
        fans = collections.Counter()
        for s in SEGS[seg]:
            for f, k in RA.get((r, s), (None, None, []))[2]: fans[f] += k
        nome = cfg['nomes'].get(r) or nome_auto(r, raz, fans.most_common(5))
        rot = 'garagens' if seg == 'EST' else 'unidades'
        det = ', '.join(f'{ROT_SEG.get(s, s)} {k}' for s, k in por.most_common()) if seg == 'PET' else ''
        obs = '; '.join(x for x in [ufs_txt(ufs, rot), det, cfg.get('notas', {}).get(r, '')] if x)
        d = linha(seg, r, nome, cfg['grupos'].get(r, nome), n, NOME_SEG[seg], segm_auto(seg, por, raz),
                  'Complemento: Receita Federal (rede achada por CNAE)', obs, 'C', n)
        if not d['cnpj']: print('sem matriz', seg, r, raz); continue
        out[seg].append(d); usadas[seg].add(r)
# ---------- nome repetido com CNPJ diferente: e outra empresa; o municipio entra no nome
ORD = {'CNPJ CONFIRMADO NA RECEITA': 0, 'CNPJ PROVÁVEL (CONFERIR)': 1, 'CNPJ DE UMA UNIDADE DA REDE (FRANQUEADORA NÃO ACHADA)': 2, 'CNPJ NÃO CONFIRMADO': 3}
for seg, rs in out.items():
    cont = collections.Counter(norm(d['fant']) for d in rs)
    for d in rs:
        if cont[norm(d['fant'])] > 1 and d['mun']: d['fant'] = f"{d['fant']} ({d['mun']}/{d['uf']})"; d['key'] = seg + '|' + norm(d['fant'])
    dup = [k for k, n in collections.Counter(norm(d['fant']) for d in rs).items() if n > 1]
    cn = collections.Counter(d['cnpj'] for d in rs if d['cnpj'])
    rs.sort(key=lambda d: (ORD.get(d['conf'], 4), -(d['lojas'] or 0)))
    print(seg, 'linhas', len(rs), '| nomes repetidos', dup, '| CNPJ repetido', [k for k, n in cn.items() if n > 1])
    print('  ', dict(collections.Counter(d['conf'] for d in rs)))
    print('  ', dict(collections.Counter(REG.get(d['uf_sede'], '?') for d in rs)))
pickle.dump(out['PET'], open('pet_rows.pkl', 'wb')); pickle.dump(out['EST'], open('estac_rows.pkl', 'wb'))
with open('pet_estac_view.txt', 'w', encoding='utf8') as f:
    for seg, rs in out.items():
        f.write(f'===== {seg}\n')
        for i, d in enumerate(rs, start=2):
            f.write(f"{i} | {d['fant'][:34]} | {d['raz'][:42]} | {d['grupo'][:24]} | {d['lojas']} | {d['cnpj']} | {d['conf'][:12]} | {d['uf_sede']} {d['mun'][:14]} | {d['segm'][:24]} | {d['rede']} | {d['obs'][:80]}\n")
