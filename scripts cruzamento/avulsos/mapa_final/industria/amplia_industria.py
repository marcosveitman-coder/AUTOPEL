"""Aba FULLFILMIT_INDUSTRIA FINAL, ampliação (19/11/2026, 2º pedido: "foque nas associações / cooperativas / empresas
regionais do Brasil").
Três listas novas, além das 496 linhas da 1ª versão:
  A) REGIONAIS: em cada UF e segmento, as 3 maiores pelo capital social (>= R$ 10 mi), com a MATRIZ numa atividade
     industrial do segmento (assim ficam fora holding, comércio e serviços com uma fábrica pequena), porte "demais",
     sociedade (fora empresário individual, EIRELI e cooperativa);
  B) COOPERATIVAS (natureza 2143) com unidade industrial ativa do segmento: capital >= R$ 5 mi ou 4+ unidades industriais (fora as em liquidação);
  C) ASSOCIADAS das entidades do setor (sites oficiais, 19/11/2026): ABIA (alimentos), ABRAFATI (tintas) e ABRAMAT
     (materiais de construção); o nome é ligado à raiz na base industrial da Receita (razão ou fantasia com o nome inteiro)
     e a escolha final é conferida à mão (ASSOC_RAIZ / ASSOC_FORA em ../dados/dec_industria.py).
Fica fora o que já está na aba, o que está em FORA (dec_industria) e CNAE de CNAE_FORA.
Saída: cand_amp.pkl {raiz: {... como cand.pkl, origem=[...]}} e amp_view.txt (revisão), assoc_match.txt (ligação dos nomes).
Uso: python amplia_industria.py <pasta de trabalho>"""
import os, sys, re, pickle, unicodedata, collections as C
P = sys.argv[1]
AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI); sys.path.insert(0, os.path.join(AQUI, '..', 'dados'))
from receita_industria import seg_de
import dec_industria as DEC

REG = {}
for r_, ufs in [('NORTE', 'AC AP AM PA RO RR TO'), ('NORDESTE', 'AL BA CE MA PB PE PI RN SE'), ('CENTRO-OESTE', 'DF GO MT MS'),
                ('SUDESTE', 'ES MG RJ SP'), ('SUL', 'PR RS SC')]:
    for u in ufs.split(): REG[u] = r_
ASSOC = {
    'ABIA': ('FMCG', '3 Coracoes|AAK|AB Mauri|Abbott|ADM|Ajinomoto|Alispec|Ambev|Arcor|Bakels|Bauducco|Bimbo|Biorigin|Brainfarma|BRF|'
             'Bunge|Cargill|Coamo|Cocamar|Cosmed|Danone|De Marchi|Oetker|Firmenich|Duas Rodas|Dubai Alimentos|Ferrero|Froneri|'
             'Garoto|General Mills|Givaudan|Grupo CRM|Heinz|ICL|IFF|Ingredion|Kerry|La Serenissima|Lapiendrius|M Dias Branco|Mars|'
             'McCain|Mondelez|Nestle|Nissin|Nutrimental|PepsiCo|Perfetti|Piracanjuba|Pramesa|Richs|Roquette|Sakura|Seara|'
             'SL Alimentos|Sococo|Symrise|Tate Lyle|Tetra Pak|Unilever|Vigor|Vitaquima|Yakult'),
    'ABRAFATI': ('QUIM', 'Advance Tintas|AkzoNobel|Alessi|Anjo Tintas|Axalta|Bautech|Brasilux|Dacar|Eucatex|Futura Tintas|Hidracor|'
                 'Hidrotintas|Hydronorth|Ibratin|Indutil|Iquine|Iraja|Isocoat|Jotun|Killing|Lechler|Leinertex|Lukscolor|Maxirubber|'
                 'Maza|Montana Quimica|PPG|Qualyvinil|Resicolor|Sayerlack|Sherwin Williams|Skylack|Suvinil|Universo Tintas|Valspar|'
                 'WEG Tintas|Zinga'),
    'ABRAMAT': ('CONST', 'Tarkett|Tigre|Vedacit|Viapol|Amanco|BASF|Brasilit|Cebrace|Coral|Decortiles|Docol|Eliane|Fortalev|Glasu|'
                'Gypsum|Henkel|Icasa|Isover|Kohler|Knauf|Krona|Legrand|Lorenzetti|Mekal|Pado|Placo|Prysmian|Quartzolit|'
                'Saint Gobain|Sika|Soprano|Soprema|Denver|Steck|Franke|Hoff|ArcelorMittal|Gerdau'),
}


def norm(x): return ' '.join(re.sub(r'[^A-Z0-9 ]', ' ', unicodedata.normalize('NFKD', x or '').encode('ascii', 'ignore').decode().upper().replace('&', ' E ')).split())


d = pickle.load(open(os.path.join(P, 'r_ind.pkl'), 'rb')); R, E = d['raiz'], d['emp']
rows = pickle.load(open(os.path.join(P, 'ind_rows.pkl'), 'rb'))
ja = {r for o in rows for r in o['raizes']}
cap = lambda r: float(str(E[r][2]).replace(',', '.') or 0) if r in E else 0.0
uf = lambda r: R[r]['matriz'][2] if R[r]['matriz'] else R[r]['uf'].most_common(1)[0][0]


def seg(r):
    m = R[r]['matriz']
    if m and seg_de(m[4]): return DEC.SEG.get(r, seg_de(m[4]))
    return DEC.SEG.get(r, R[r]['seg'].most_common(1)[0][0])


def fora(r):
    m = R[r]['matriz']
    rz = E.get(r, ('',))[0]
    if re.search(r'LIQUIDACAO|FALIDO|FALENCIA|\bSCP\b', rz): return True    # em liquidação, falida ou sociedade em conta de participação
    if R[r]['matriz'] and R[r]['matriz'][2] == 'EX': return True              # empresa estrangeira sem matriz no Brasil
    return r in ja or r in DEC.FORA or r in getattr(DEC, 'FORA_AMP', {}) or (m and m[4] in DEC.CNAE_FORA)


sel = C.defaultdict(set)
# A) regionais
by = C.defaultdict(list)
for r, x in R.items():
    if r not in E or fora(r): continue
    e = E[r]
    if e[3] != '05' or e[1] in ('2135', '2305', '2321', '2143') or cap(r) < 10e6: continue
    m = x['matriz']
    if not m or not seg_de(m[4]): continue
    by[(uf(r), seg(r))].append(r)
for (u, s), l in by.items():
    for r in sorted(l, key=lambda r: -cap(r))[:3]: sel[r].add(f'REGIONAL {u}')
# B) cooperativas
for r, x in R.items():
    if r not in E or fora(r) or E[r][1] != '2143': continue
    # 19/11: só as grandes (capital >= R$ 5 mi ou 5+ unidades industriais); as pequenas de agricultura familiar ficam fora
    if (cap(r) >= 5e6 and x['n'] >= 1) or x['n'] >= 4: sel[r].add('COOPERATIVA')
# C) associadas
idx = C.defaultdict(set)
txt = {}
for r, x in R.items():
    if r not in E: continue
    t = ' | '.join({norm(E[r][0]), norm(x['matriz'][5] if x['matriz'] else '')} | {norm(f) for f, _ in x['fant'].most_common(3)})
    txt[r] = t
    for w in set(t.split()): idx[w].add(r)
ASSOC_RAIZ = getattr(DEC, 'ASSOC_RAIZ', {}); ASSOC_FORA = getattr(DEC, 'ASSOC_FORA', {})
with open(os.path.join(P, 'assoc_match.txt'), 'w', encoding='utf8') as f:
    for ent, (sg, nomes) in ASSOC.items():
        for nome in nomes.split('|'):
            n = norm(nome)
            if (ent, nome) in ASSOC_FORA or nome in ASSOC_FORA: f.write(f'{ent} | {nome} | FORA: {ASSOC_FORA.get(nome) or ASSOC_FORA.get((ent, nome))}\n'); continue
            if nome in ASSOC_RAIZ:
                r = ASSOC_RAIZ[nome]
                f.write(f'{ent} | {nome} | MANUAL {r} {E.get(r, ("?",))[0][:50]}' + (' (já na aba)' if r in ja else '') + '\n')
                if r in R and not fora(r): sel[r].add(f'ASSOC {ent}|{nome}')
                continue
            ws = n.split(); pat = re.compile(r'(?<![A-Z0-9])' + r'\s?'.join(map(re.escape, ws)) + r'(?![A-Z0-9])')
            cand = set.intersection(*[idx.get(w, set()) for w in ws]) if ws else set()
            cand = sorted((r for r in cand if pat.search(txt[r])), key=lambda r: -cap(r))
            ja_aba = [r for r in cand if r in ja]
            if ja_aba: f.write(f'{ent} | {nome} | JÁ NA ABA {ja_aba[0]} {E[ja_aba[0]][0][:50]}\n'); continue
            top = cand[:4]
            f.write(f'{ent} | {nome} | ' + (' || '.join(f'{r} {E[r][0][:45]} cap {cap(r)/1e6:,.0f} mi {uf(r)} {seg(r)} un {R[r]["n"]}' for r in top) or 'NADA') + '\n')
            if top and cap(top[0]) >= 1e6 and not fora(top[0]): sel[top[0]].add(f'ASSOC {ent}|{nome}')
out = {}
for r, ori in sel.items():
    x = R[r]; m = x['matriz'] or ('', '', uf(r), '', '', '')
    out[r] = dict(raiz=r, razao=E[r][0], natureza=E[r][1], capital=cap(r), seg=seg(r), origem=sorted(ori), un=x['n'],
                  ufs=dict(x['uf']), uf=uf(r), regiao=REG.get(uf(r), '?'), mun=m[3], cnae=m[4],
                  fant=(m[5] or (x['fant'].most_common(1)[0][0] if x['fant'] else '')), ordem=m[0], dv=m[1])
pickle.dump(out, open(os.path.join(P, 'cand_amp.pkl'), 'wb'))
with open(os.path.join(P, 'amp_view.txt'), 'w', encoding='utf8') as f:
    for r in sorted(out, key=lambda r: (out[r]['seg'], out[r]['uf'], -out[r]['capital'])):
        o = out[r]
        f.write(f"{o['seg']:6} {o['uf']} | {r} | {o['razao'][:55]:55} | cap {o['capital']/1e6:8,.0f} mi | un {o['un']:3} | cnae {o['cnae']} | "
                f"{o['fant'][:22]:22} | {','.join(o['origem'])}\n")
print(len(out), C.Counter(o['seg'] for o in out.values()), C.Counter(x.split()[0] for o in out.values() for x in o['origem']))
