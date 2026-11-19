"""Aba FULLFILMIT_INDUSTRIA FINAL, passo 4: monta as linhas (19/11/2026).
Entradas (pasta de trabalho): cand.pkl (candidatas_industria.py), fontes.json (fontes_industria.py), fontes_match.json
(1ªs candidatas de cada fonte na base industrial), r_fontes*.pkl (busca_receita.py: fontes que não estavam na base
industrial e TODAS as unidades das raízes escolhidas), exame_mm26.json, ../dados/dec_industria.py (local, fora do git).
Regras:
  - entra quem está na lista da Receita (corte por segmento/região) ou numa fonte nominal (Exame, Anfavea, Eletros, ANIP,
    lubrificantes, PBE/Inmetro), menos as raízes de FORA e, só na lista automática, os CNAE de CNAE_FORA;
  - uma linha por empresa (GRUPOS junta raízes da mesma empresa);
  - Nº DE UNIDADES = estabelecimentos ATIVOS de todas as raízes da linha (fábricas, CDs, escritórios), na Receita;
  - PORTE pela receita 2025 da Exame quando a empresa está no ranking (A >= R$ 1 bi, B R$ 300 mi a 1 bi, C R$ 20 a 300 mi,
    D < R$ 20 mi); fora do ranking, estimado pelo capital social (A >= R$ 500 mi, B 100-500, C 10-100, D < 10), com a nota
    "estimado pelo capital social";
  - RANKING BRASIL / NA REGIÃO = posição pela receita (Exame) e, sem ela, pelo capital social.
Saída: ind_rows.pkl (lista de dicts) e ind_view.txt.
Uso: python build_industria.py <pasta de trabalho>"""
import os, sys, re, json, pickle, unicodedata, collections as C
P = sys.argv[1]
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'dados'))  # dec_industria.py: local, fora do git (tem raízes de CNPJ)
import dec_industria as DEC
from receita_industria import seg_de

REG = {}
for r_, ufs in [('NORTE', 'AC AP AM PA RO RR TO'), ('NORDESTE', 'AL BA CE MA PB PE PI RN SE'), ('CENTRO-OESTE', 'DF GO MT MS'),
                ('SUDESTE', 'ES MG RJ SP'), ('SUL', 'PR RS SC')]:
    for u in ufs.split(): REG[u] = r_
SEGN = {'PNEU': 'PNEUS', 'FMCG': 'BENS DE CONSUMO NÃO DURÁVEIS (FMCG)', 'FARMA': 'FARMACÊUTICA E MÉDICA',
        'ELETRO': 'ELETROELETRÔNICOS E LINHA BRANCA', 'TEXTIL': 'TÊXTIL, CALÇADOS E VESTUÁRIO', 'AUTO': 'AUTOMOTIVO, AUTOPEÇAS E MÁQUINAS',
        'QUIM': 'QUÍMICA, PLÁSTICOS, PAPEL E EMBALAGENS', 'CONST': 'MÓVEIS E MATERIAIS DE CONSTRUÇÃO', 'OLEO': 'ÓLEO E LUBRIFICANTES (FABRICANTES E REVENDAS)'}


def norm(x): return ' '.join(re.sub(r'[^A-Z0-9 ]', ' ', unicodedata.normalize('NFKD', x or '').encode('ascii', 'ignore').decode().upper().replace('&', ' E ')).split())


def porte_receita(mi):
    return ('A – Grande (≥ R$ 1 bi)' if mi >= 1000 else 'B – Média (R$ 300 mi a 1 bi)' if mi >= 300 else
            'C – Regional (R$ 20 mi a 300 mi)' if mi >= 20 else 'D – Pequena / local (< R$ 20 mi)')


def porte_capital(cap):
    mi = cap / 1e6
    return ('A – Grande (≥ R$ 1 bi)' if mi >= 500 else 'B – Média (R$ 300 mi a 1 bi)' if mi >= 100 else
            'C – Regional (R$ 20 mi a 300 mi)' if mi >= 10 else 'D – Pequena / local (< R$ 20 mi)')


def main():
    cand = pickle.load(open(os.path.join(P, 'cand.pkl'), 'rb'))
    F = json.load(open(os.path.join(P, 'fontes.json'), encoding='utf8'))
    FM = dict((i, l) for i, l in json.load(open(os.path.join(P, 'fontes_match.json'))))
    RF = {}   # raiz -> (razão, natureza, capital, porte) das buscas completas
    EST = C.defaultdict(list)
    hits = C.defaultdict(set)
    for arq in ('r_fontes.pkl', 'r_fontes2.pkl'):
        p = os.path.join(P, arq)
        if not os.path.exists(p): continue
        d = pickle.load(open(p, 'rb'))
        RF.update(d['emp'])
        for row in d['est']: EST[row[0]].append(row)
        for r, t in d['hits_raz']: hits[t].add(r)
    ind = pickle.load(open(os.path.join(P, 'r_ind.pkl'), 'rb'))
    EMP = dict(ind['emp']); EMP.update(RF)
    ex = json.load(open(os.path.join(P, 'exame_mm26.json'), encoding='utf8'))

    def cap(r):
        try: return float(str(EMP[r][2]).replace(',', '.'))
        except Exception: return 0.0

    linhas = {}   # raiz -> dict

    def novo(r, seg, origem, nome_fonte=''):
        if r in DEC.FORA: return
        L = linhas.setdefault(r, dict(raizes=[r], seg=seg, origem=set(), fontes=set(), nome_fonte='', selo=set()))
        L['origem'].add(origem)
        if nome_fonte and not L['nome_fonte']: L['nome_fonte'] = nome_fonte
        if seg and origem.startswith(('EXAME', 'ANFAVEA', 'ANIP', 'ELETROS', 'LUBES')): L['seg'] = seg
    # lista automática da Receita
    for r, o in cand.items():
        if o['cnae'] in DEC.CNAE_FORA: continue
        for g in o['origem']: novo(r, DEC.SEG.get(r, o['seg']), 'Receita: ' + ('maiores do Brasil' if g == 'BRASIL' else 'maiores da ' + g.title().replace('Regiao', 'região')))
    # ampliação (19/11, 2º pedido): regionais por UF, cooperativas e associadas de ABIA / ABRAFATI / ABRAMAT (amplia_industria.py)
    pa = os.path.join(P, 'cand_amp.pkl')
    if os.path.exists(pa):
        for r, o in pickle.load(open(pa, 'rb')).items():
            for g in o['origem']:
                if g.startswith('REGIONAL'): novo(r, o['seg'], f'Receita: maiores do segmento na UF ({g.split()[1]})')
                elif g == 'COOPERATIVA': novo(r, o['seg'], 'Receita: cooperativa agroindustrial')
                elif g.startswith('ASSOC'):
                    ent, nm = g[6:].split('|', 1)
                    novo(r, o['seg'], f'{ent}: associada (site da entidade, 19/11/2026)', nm)
    # fontes nominais
    sem = []
    for i, f in enumerate(F):
        if f['nome'] in DEC.FORA_FONTE: continue
        r = DEC.RAIZ_FONTE.get(f['nome']) or (FM.get(i) or [None])[0]
        if not r:
            sem.append(f); continue
        novo(r, DEC.SEG.get(r, f['seg']), f['fonte'], f['nome'])
        if f.get('selo') and r in linhas: linhas[r]['selo'].add(f['selo'])
    # líderes achados por pesquisa (fora da Exame e do corte da Receita)
    grupos = list(DEC.GRUPOS)
    for nome, seg, rs in DEC.EXTRA:
        novo(rs[0], seg, 'Pesquisa: líder do setor', nome)
        if len(rs) > 1: grupos.append([nome, nome, rs])
    # grupos: junta as raízes numa linha só (a 1ª é a principal)
    for nome, grupo, rs in grupos:
        presentes = [r for r in rs if r in linhas]
        if not presentes: continue
        alvo = rs[0] if rs[0] in linhas else presentes[0]
        L = linhas[alvo]; L['nome_grupo'] = nome; L['grupo'] = grupo
        for r in rs:
            if r == alvo: continue
            if r in linhas:
                o = linhas.pop(r); L['origem'] |= o['origem']; L['selo'] |= o['selo']
            if r not in L['raizes']: L['raizes'].append(r)
    pickle.dump(dict(linhas=linhas, sem=sem), open(os.path.join(P, 'ind_pre.pkl'), 'wb'))
    print('linhas', len(linhas), 'fontes sem raiz', len(sem))
    for f in sem: print('  sem raiz:', f['fonte'], f['nome'])
    if not os.path.exists(os.path.join(P, 'r_unid.pkl')): return
    # ---------- linhas finais (precisa de r_unid.pkl: todas as unidades das raízes, da Receita) ----------
    U = pickle.load(open(os.path.join(P, 'r_unid.pkl'), 'rb'))
    EMP.update(U['emp'])
    EST = C.defaultdict(list)
    for row in U['est']: EST[row[0]].append(row)
    if os.path.exists(os.path.join(P, 'r_unid_amp.pkl')):         # unidades das raízes da ampliação
        U2 = pickle.load(open(os.path.join(P, 'r_unid_amp.pkl'), 'rb'))
        for k, v in U2['emp'].items(): EMP.setdefault(k, v)
        tem = set(EST)
        for row in U2['est']:
            if row[0] not in tem: EST[row[0]].append(row)
    MUN = municipios()
    rec = receitas(P, F, FM)          # raiz -> (receita R$ mi, nome no ranking)
    SIT = {'01': 'NULA', '02': 'ATIVA', '03': 'SUSPENSA', '04': 'INAPTA', '08': 'BAIXADA'}
    out = []
    for r0, L in linhas.items():
        rs = L['raizes']
        ests = [e for r in rs for e in EST.get(r, [])]
        ativos = [e for e in ests if e[5] == '02']
        mat = next((e for e in EST.get(r0, []) if e[3] == '1'), None)
        cnpj = (mat[0] + mat[1] + mat[2]) if mat else ''
        ind_n = sum(ind['raiz'][r]['n'] for r in rs if r in ind['raiz'])
        rv = next((rec[r] for r in rs if r in rec), None) or next((DEC.RECEITA_EXTRA[r] for r in rs if r in getattr(DEC, 'RECEITA_EXTRA', {})), None)
        capt = sum(cap(r) for r in rs)
        porte = porte_receita(rv[0]) if rv else porte_capital(capt)
        coop = EMP.get(r0, ('', ''))[1] == '2143'
        if coop and not rv:
            # cooperativa: o capital social (quotas dos cooperados) não mede o tamanho; vale o maior entre o capital e o
            # nº de unidades industriais (10+ = B, 4+ = C)
            pu = 'B – Média (R$ 300 mi a 1 bi)' if ind_n >= 10 else 'C – Regional (R$ 20 mi a 300 mi)' if ind_n >= 4 else ''
            if pu and pu[0] < porte[0]: porte = pu
        fan = (mat[4] if mat else '').strip()
        nome = (DEC.NOME.get(r0) or L.get('nome_grupo') or L['nome_fonte'] or
                (fan.title() if fan and len(fan) > 3 and not re.search(r'MATRIZ|FILIAL|UNIDADE|ESCRITORIO|SEDE|FABRICA|\d', fan) else '') or
                re.sub(r'\b(S\.?/?A\.?|LTDA\.?|S/A)\b.*$', '', EMP.get(r0, ('',))[0]).strip().title())
        obs = []
        if len(rs) > 1: obs.append('uma linha para a empresa, que junta as raízes: ' + '; '.join(f"{r} {EMP.get(r, ('?',))[0]}" for r in rs))
        if rv and 'Forbes' in rv[1]: obs.append(f'receita 2025: R$ {rv[0]:,.0f} mi ({rv[1]})'.replace(',', '.'))
        elif rv: obs.append(f'receita 2025: R$ {rv[0]:,.0f} mi (Exame Melhores e Maiores 2026, "{rv[1]}")'.replace(',', '.'))
        elif coop: obs.append(f'cooperativa: porte estimado pelo nº de unidades industriais ({ind_n}) e pelo capital social (R$ {capt / 1e6:,.0f} mi)'.replace(',', '.'))
        else: obs.append(f'porte estimado pelo capital social (R$ {capt / 1e6:,.0f} mi): a empresa não publica balanço no ranking da Exame'.replace(',', '.'))
        if L['selo']: obs.append('Inmetro/PBE: ' + '; '.join(sorted(L['selo'])))
        uf = mat[10] if mat else ''
        out.append(dict(raiz=r0, raizes=rs, fant=nome, raz=EMP.get(r0, ('',))[0], grupo=L.get('grupo') or '', lojas=len(ativos),
                        cnpj=cnpj, conf='CNPJ CONFIRMADO NA RECEITA' if mat and mat[5] == '02' else ('MATRIZ ' + SIT.get(mat[5], mat[5]) + ' NA RECEITA' if mat else 'MATRIZ NÃO ACHADA'),
                        uf=uf, mun=MUN.get(mat[11], '') if mat else '', regiao=REG.get(uf, ''), sit=SIT.get(mat[5], '') if mat else '',
                        fonte='Receita Federal: estabelecimentos ativos de todas as raízes da linha (out/2026)', q=ind_n,
                        segm=SEGN[L['seg']], seg=L['seg'], origem='; '.join(sorted(L['origem'])), obs='; '.join(obs), porte=porte,
                        receita=rv[0] if rv else None, capital=capt))
    # ranking: porte (receita; sem ela, capital) no Brasil e na região
    chave = lambda o: (o['porte'][0], -(o['receita'] if o['receita'] is not None else o['capital'] / 1e6 * 0.5))
    for i, o in enumerate(sorted(out, key=chave), start=1): o['rk'] = i
    por = C.defaultdict(list)
    for o in out: por[o['regiao']].append(o)
    for g, lst in por.items():
        for i, o in enumerate(sorted(lst, key=chave), start=1): o['rkr'] = i
    out.sort(key=lambda o: (list(SEGN).index(o['seg']), o['rk']))
    pickle.dump(out, open(os.path.join(P, 'ind_rows.pkl'), 'wb'))
    with open(os.path.join(P, 'ind_view.txt'), 'w', encoding='utf8') as f:
        for o in out:
            f.write(f"{o['seg']:6} {o['rk']:4} {o['fant'][:40]:40} | {o['raz'][:40]:40} | {o['cnpj']} {o['conf'][:12]} | un {o['lojas']:4} ind {o['q']:3} | "
                    f"{o['uf']} {o['mun'][:18]} | {o['porte'][:1]} {o['receita'] or ''} | {o['origem'][:60]}\n")
    print('linhas finais', len(out), C.Counter(o['conf'] for o in out), C.Counter(o['porte'][0] for o in out))


def municipios():
    import glob, zipfile, io
    d = os.environ.get('AUTOPEL_RECEITA') or sorted(glob.glob(r'C:\dados_receita\*'))[-1]
    out = {}
    with zipfile.ZipFile(os.path.join(d, 'Municipios.zip')) as z, z.open(z.namelist()[0]) as f:
        for raw in io.TextIOWrapper(f, encoding='latin-1'):
            p = raw.strip().split('";"')
            if len(p) == 2: out[p[0].strip('"')] = p[1].strip('"').title()
    return out


def receitas(P, F, FM):
    """raiz -> (receita 2025 em R$ mi, nome no ranking Exame), pela ligação fonte EXAME -> ranking (fonte_exame.json)."""
    fe = json.load(open(os.path.join(P, 'fonte_exame.json'), encoding='utf8'))
    out = {}
    for i, f in enumerate(F):
        x = fe.get(str(i))
        if not x or not f['fonte'].startswith('EXAME') or f['nome'] in DEC.FORA_FONTE: continue
        r = DEC.RAIZ_FONTE.get(f['nome']) or (FM.get(i) or [None])[0]
        if r and r not in out: out[r] = (x[1], x[0])
    return out


if __name__ == '__main__':
    main()
