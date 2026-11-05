"""Aba OUTROS FINAL: as 140 empresas da aba de pesquisa (Beleza/Saude, Oticas, Servicos/Franquias) que nao estavam no
mapeamento. Mesmo modelo das abas FINAL + colunas SEGMENTO e SUBSEGMENTO no fim.
CNPJ: busca por nome na Receita (r17.pkl, busca_receita.py q17.json) e escolha:
  1) empresa com o nome na razao social e matriz ativa: franqueadora (CNAE 7740-3/00 ou FRANQUIA/FRANCHISING/MARCAS na
     razao) > a que tem mais unidades com o nome;  nivel C se for franqueadora ou tiver 5+ unidades com o nome, senao P;
  2) sem razao com o nome: a empresa com mais unidades com o nome no fantasia (nivel U = CNPJ de uma unidade);
  3) nada: CNPJ NAO CONFIRMADO.  Decisoes a mao em dec_outros.json (ref -> [cnpj ou '', nivel, nota]).
Saida: outros_rows.pkl e outros_view.txt.  Uso: python build_outros.py"""
import os as _os
_os.chdir(_os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'dados'))
import pickle, json, re, collections
R = pickle.load(open('r17.pkl', 'rb')); emp = R['emp']; est = R['est']
if _os.path.exists('r18.pkl'):  # 2a busca: franqueadoras e raizes escolhidas a mao
    _r = pickle.load(open('r18.pkl', 'rb')); emp.update(_r['emp']); est = est + [x for x in _r['est']]
# CNAE compativel com cada subsegmento (a escolha automatica so aceita esses, franquia ou holding)
FAM = {'Odontologia': ('8630',), 'Academias e pilates': ('9313', '8650', '8591'), 'Depilação e estética': ('9602', '8650', '8690', '8630'),
       'Salões e sobrancelhas': ('9602',), 'Cosméticos e perfumaria': ('4772', '2063', '4646'), 'Cuidadores': ('87', '8690', '9609', '8650'),
       'Clínicas de saúde': ('86',), 'Óticas': ('4774', '3250')}
SEMPRE = ('7740300', '6462000', '6463800', '8299799', '7020400')
def compat(sub, cnae): return sub not in FAM or cnae in SEMPRE or cnae.startswith(FAM[sub])
MUN = pickle.load(open('mun.pkl', 'rb'))
LISTA = json.load(open('outros_lista.json', encoding='utf8'))
ALE = {o['ref']: o for o in json.load(open('alessandra_cruzamento.json', encoding='utf8'))}
DEC = json.load(open('dec_outros.json', encoding='utf8')) if _os.path.exists('dec_outros.json') else {}
SIT = {'01': 'NULA', '02': 'ATIVA', '03': 'SUSPENSA', '04': 'INAPTA', '08': 'BAIXADA'}
REG = {}
for g, ufs in [('NORTE', 'AC AP AM PA RO RR TO'), ('NORDESTE', 'AL BA CE MA PB PE PI RN SE'), ('CENTRO-OESTE', 'DF GO MT MS'),
               ('SUDESTE', 'ES MG RJ SP'), ('SUL', 'PR RS SC')]:
    for u in ufs.split(): REG[u] = g
porraiz = collections.defaultdict(list)
for x in est: porraiz[x[0]].append(x)
def dv(c):
    def k(b, w):
        s = sum(int(x) * y for x, y in zip(b, w)); r = s % 11; return '0' if r < 2 else str(11 - r)
    w1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    return len(c) == 14 and k(c[:12], w1) == c[12] and k(c[:13], [6] + w1) == c[13]
fant_t = collections.defaultdict(lambda: collections.Counter())   # termo -> raiz -> unidades ativas com o nome
for i, t in R['hits_fant']:
    x = est[i]
    if x[5] == '02': fant_t[t][x[0]] += 1
raz_t = collections.defaultdict(set)
for r, t in R['hits_raz']: raz_t[t].add(r)
FRANQ = re.compile(r'FRANQ|FRANCHIS|MARCAS|LICENCIAMENTO')
def matriz(r):
    m = [x for x in porraiz.get(r, []) if x[3] == '1']
    return m[0] if m else None
def ativos(r): return sum(1 for x in porraiz.get(r, []) if x[5] == '02')
out = []; view = []
for ref, it in LISTA.items():
    a = ALE[ref]; termos = it['termos']
    nome_cnt = collections.Counter()
    for t in termos: nome_cnt.update(fant_t[t])
    raizes_raz = set().union(*[raz_t[t] for t in termos])
    cand = []
    for r in raizes_raz | set(nome_cnt):
        m = matriz(r); raz = emp.get(r, ('',))[0]
        if not m: continue
        fr = (m[8] == '7740300' or bool(FRANQ.search(raz))) and not raz.startswith('ASSOCIACAO') and m[8][:4] not in ('9411', '9430', '9499')
        cand.append(dict(r=r, raz=raz, m=m, ativa=m[5] == '02', fr=fr, com_nome=nome_cnt[r], ativos=ativos(r), na_razao=r in raizes_raz))
    cand.sort(key=lambda c: (-(c['ativa'] and c['na_razao']), -c['fr'], -c['com_nome'], -c['ativos']))
    rede = sum(nome_cnt.values()); n_raizes = len(nome_cnt)
    esc, niv, nota = None, 'NC', ''
    rz = [c for c in cand if c['ativa'] and c['na_razao'] and not c['raz'].startswith('ASSOCIACAO') and compat(it['sub'], c['m'][8])]
    if rz:
        esc = rz[0]; niv = 'C' if esc['fr'] or esc['com_nome'] >= 5 else 'P'
    elif nome_cnt:
        _ok = [r for r, k in nome_cnt.most_common() if any(c['r'] == r and compat(it['sub'], c['m'][8]) for c in cand)] or [nome_cnt.most_common(1)[0][0]]
        r0 = _ok[0]; esc = next(c for c in cand if c['r'] == r0) if any(c['r'] == r0 for c in cand) else None; niv = 'U' if esc else 'NC'
    d = DEC.get(ref)
    cnpj = esc['m'][0] + esc['m'][1] + esc['m'][2] if esc else ''
    if d:
        cnpj, niv, nota = d[0], d[1], d[2] if len(d) > 2 else ''
        esc = None
        if cnpj:
            e = [x for x in porraiz.get(cnpj[:8], []) if x[0] + x[1] + x[2] == cnpj] or ([matriz(cnpj[:8])] if matriz(cnpj[:8]) else []) or sorted([x for x in porraiz.get(cnpj[:8], []) if x[5] == '02'], key=lambda x: x[1])
            assert e, (ref, cnpj); m = e[0]; cnpj = m[0] + m[1] + m[2]
            esc = dict(r=cnpj[:8], raz=emp.get(cnpj[:8], ('',))[0], m=m, ativa=m[5] == '02', com_nome=nome_cnt[cnpj[:8]], ativos=ativos(cnpj[:8]))
    lojas = a['lojas'] if isinstance(a['lojas'], (int, float)) and a['lojas'] > 0 else None
    fonte = f"Alessandra Dayan ({a['base']})" if lojas else ''
    if not lojas and rede: lojas, fonte = rede, 'Receita: unidades ativas com o nome (out/2026)'
    row = dict(fant=it['nome'], raz='', grupo=it['nome'], lojas=int(lojas) if lojas else None, cnpj='', conf='', ref='', ref_desc='', uf='', mun='', sit='',
               fonte=fonte, rede=rede or None, segm=it['sub'], origem=f'Alessandra Dayan: {ref}', obs=nota, seg=it['seg'], sub=it['sub'], uf_sede='', uf_sede_fonte='')
    if esc:
        m = esc['m']; c = m[0] + m[1] + m[2]; assert dv(c), c
        row.update(raz=esc['raz'].strip(), uf=m[10], mun=MUN.get(m[11], m[11]), sit=SIT.get(m[5], m[5]))
        if niv == 'U':
            row.update(cnpj=c, ref=c, ref_desc=f"{row['raz']} ({row['mun']}/{row['uf']})", conf='CNPJ DE UMA UNIDADE DA REDE (FRANQUEADORA NÃO ACHADA)')
            extra = f'{n_raizes} empresas com unidades com o nome'
        else:
            row.update(cnpj=c, conf='CNPJ CONFIRMADO NA RECEITA' if niv == 'C' else 'CNPJ PROVÁVEL (CONFERIR)', uf_sede=m[10])
            extra = ('franqueadora (CNAE de franquia ou razão de franquia/marcas); ' if esc.get('fr') else '') + f"{esc['com_nome']} unidades com o nome na empresa, {esc['ativos']} estabelecimentos ativos"
        if m[5] != '02': row['conf'] = f'CNPJ {SIT.get(m[5], m[5])} NA RECEITA'
        row['obs'] = '; '.join(x for x in [nota, extra] if x)
    else:
        row['conf'] = 'CNPJ NÃO CONFIRMADO'
        row['obs'] = nota or 'nenhuma empresa nem unidade ativa com o nome na Receita'
    if not row['uf_sede'] and isinstance(a['uf'], str) and len(a['uf']) == 2: row['uf_sede'] = a['uf']; row['uf_sede_fonte'] = 'UF da sede pela aba de origem'
    out.append(row)
    view.append(f"\n{ref} {it['nome']} | origem {a['lojas']} ({a['base']}) | escolha {niv} {row['cnpj']} {row['raz'][:45]} | rede {rede} em {n_raizes} empresas")
    for c in cand[:5]:
        view.append(f"   {c['r']} {'R' if c['na_razao'] else ' '}{'F' if c['fr'] else ' '} {SIT.get(c['m'][5], c['m'][5])[:3]} {c['m'][8]} {c['m'][10]} nome={c['com_nome']:4d} ativos={c['ativos']:4d} :: {c['raz'][:60]}")
ORD = {'CNPJ CONFIRMADO NA RECEITA': 0, 'CNPJ PROVÁVEL (CONFERIR)': 1, 'CNPJ DE UMA UNIDADE DA REDE (FRANQUEADORA NÃO ACHADA)': 2}
out.sort(key=lambda d: (ORD.get(d['conf'], 3), -(d['lojas'] or 0)))
pickle.dump(out, open('outros_rows.pkl', 'wb'))
open('outros_view.txt', 'w', encoding='utf8').write('\n'.join(view))
print(len(out), dict(collections.Counter(d['conf'] for d in out)))
