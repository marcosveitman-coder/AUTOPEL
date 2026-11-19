"""Coleta as redes de todas as abas do MAPA MERCADO que recebem o status (16/11/2026).
Cada linha vira uma "entidade" com: aba, linha, nomes de busca, raízes de CNPJ, razões sociais, códigos de cliente e
grupos do ERP (IRTT). Linhas que são a mesma rede em abas diferentes formam um "cluster" (mesmo núcleo do nome ou mesma
raiz de CNPJ entre as abas FINAL / Alessandra / TOP 50 / IRTT / FARMACIA VF), para que a rede tenha UM status em todas.
Abas: as 6 FINAL pedidas (+ OUTROS FINAL só como apoio de CNPJ), Alessandra Dayan (11 tabelas), Alessandra Dayan VF,
as 11 "(TOP 50)", I.R.T.T.__Por bandeira (FINAL) e "FARMACIA VF" (redes/associações com todos os caminhos; farmácia
avulsa só pela raiz do CNPJ, porque nome de farmácia pequena casa com qualquer coisa).
Saída: entidades.pkl.  Uso: python coleta.py "<MAPA.xlsm>" "<saida.pkl>" """
import sys, re, pickle, collections as C
import openpyxl
from openpyxl.utils import get_column_letter as L
from motor import nomes_da_rede, nucleo, norm

D = lambda x: re.sub(r'\D', '', str(x or ''))
s = lambda x: str(x).strip() if x is not None else ''
FINAIS = ['ALIMENTACAO FINAL', 'REDE DE LOJAS FINAL', 'CONSTRUÇÃO FINAL', 'PET SHOP  FINAL', 'ESTACIONAMENTO FINAL',
          'POSTOS GAS & CONVENIENCIA FINAL', 'OUTROS FINAL', 'FULLFILMIT_INDUSTRIA FINAL']
TOP50 = ['ALIMENTAÇÃO (TOP 50)', 'BELEZA  (TOP 50)', 'CONVENIENCIA  (TOP 50)', 'DROGARIA  (TOP 50)', 'ESTACIONAMENTO (TOP 50)',
         'CONSTRUÇÃO (TOP 50)', 'OTICA  (TOP 50)', 'POSTO GASOLINA  (TOP 50)', 'REDE LOJAS (TOP 50)',
         'SERVIÇOS & FRANQUIAS (TOP 50)', 'SUPERMERCADO (TOP 50)']


FORMATO = set('PIZZARIA CLUBE LOJA LOJAS REDE GRUPO SUPERMERCADO SUPERMERCADOS DROGARIA DROGARIAS FARMACIA FARMACIAS OTICA OTICAS RESTAURANTE RESTAURANTES BOUTIQUE'.split())


def cnpj_ok(c):
    c = D(c)
    if len(c) != 14 or len(set(c)) == 1: return False
    def dv(b, p):
        t = sum(int(x) * y for x, y in zip(b, p)) % 11
        return '0' if t < 2 else str(11 - t)
    p1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]; p2 = [6] + p1
    return c[12] == dv(c[:12], p1) and c[13] == dv(c[:13], p2)


def familia(e):
    """Segmento da linha, para só juntar a mesma rede dentro do mesmo segmento."""
    a, sg = e['aba'], norm(e['extra'].get('seg', ''))
    if a == '"FARMACIA VF"': return 'DROG'
    fin = {'ALIMENTACAO FINAL': 'ALIM', 'REDE DE LOJAS FINAL': 'LOJAS', 'CONSTRUÇÃO FINAL': 'CONSTR', 'PET SHOP  FINAL': 'PET',
           'ESTACIONAMENTO FINAL': 'ESTAC', 'POSTOS GAS & CONVENIENCIA FINAL': 'POSTO', 'OUTROS FINAL': 'OUTROS',
           'FULLFILMIT_INDUSTRIA FINAL': 'IND'}
    if a in fin: return fin[a]
    for chave, fam in [('SUPER', 'SUPER'), ('DROGARIA', 'DROG'), ('ESTACIONAMENTO', 'ESTAC'), ('POSTOS', 'POSTO'), ('CONVENIENCIA', 'POSTO'),
                       ('MATERIAL DE CONST', 'CONSTR'), ('ALIMENTACAO', 'ALIM'), ('FOODSERVICE', 'ALIM'), ('REDE DE LOJAS', 'LOJAS'),
                       ('SUPER', 'SUPER'), ('BELEZA', 'OUTROS'), ('OTICAS', 'OUTROS'), ('SERVICOS', 'OUTROS'), ('MODA', 'LOJAS'),
                       ('ELETROMOVEIS', 'LOJAS'), ('LOJAS DE DEPARTAMENTO', 'LOJAS'), ('OUTROS SEGMENTOS', 'LOJAS')]:
        if chave in sg: return fam
    return 'X'


def ent(aba, linha, nomes, cnpjs=(), razoes=(), codigos=(), grupos=(), extra=None, raizes_ref=()):
    raizes = sorted({D(c)[:8] for c in cnpjs if cnpj_ok(c)})
    ref = sorted({D(c)[:8] for c in raizes_ref if cnpj_ok(c)} - set(raizes))
    return dict(aba=aba, linha=linha, nomes=[n for n in nomes if n], raizes=raizes, raizes_ref=ref, razoes=[r for r in razoes if s(r)],
                codigos=list(codigos), grupos=[g for g in grupos if s(g)], extra=extra or {})


def main(mapa, saida):
    wb = openpyxl.load_workbook(mapa, read_only=True, data_only=True)
    E = []
    # ---- abas FINAL (cabeçalho na linha 6)
    for aba in FINAIS:
        rows = list(wb[aba].iter_rows(values_only=True))
        h = next(i for i, r in enumerate(rows[:12]) if len(r) > 1 and s(r[1]) == 'NOME FANTASIA')
        cab = [s(x) for x in rows[h]]
        ix = {c: cab.index(c) for c in cab if c}
        grupos_cont = C.Counter(nucleo(s(r[ix['GRUPO ADMINISTRADOR']])) for r in rows[h + 1:] if r and s(r[1]))
        # 16/11: o nome do grupo administrador não vira nome de busca se o grupo é outra rede da aba (Tendall Grill x Habib's)
        redes_aba = {x.replace(' ', '') for r in rows[h + 1:] if r and s(r[1]) for x in nomes_da_rede(r[1])[:1]}
        for i, r in enumerate(rows[h + 1:], start=h + 2):
            if not r or not s(r[1]): continue
            # 19/11: o usuário pôs tabelas dinâmicas embaixo dos dados (REDE DE LOJAS FINAL, linha 334): linha de rede tem o Nº na A
            if not isinstance(r[0], (int, float)): continue
            g = s(r[ix['GRUPO ADMINISTRADOR']])
            nomes = nomes_da_rede(r[1])
            if g and grupos_cont[nucleo(g)] == 1:     # grupo com uma rede só
                nomes += [n for n in nomes_da_rede(g) if n not in nomes and (n.replace(' ', '') not in redes_aba or n in nomes_da_rede(r[1]))]
            # 19/11: o CNPJ de referência (H) é de UMA UNIDADE (franqueado): vale como vínculo (raizes_ref), mas não é o CNPJ
            # da rede. Antes ele entrava como CNPJ da rede, e a trava "CNPJs diferentes" separava a rede da mesma rede nas
            # listas da Alessandra (franqueadora): o mesmo nome ficava com dois status.
            ref = r[ix['CNPJ DE REFERÊNCIA (UNIDADE DA REDE)']] if 'CNPJ DE REFERÊNCIA (UNIDADE DA REDE)' in ix else None
            conf = s(r[ix['CONFERÊNCIA DO CNPJ']]).upper() if 'CONFERÊNCIA DO CNPJ' in ix else ''
            unidade = conf.startswith('CNPJ DE UMA UNIDADE') or conf.startswith('FRANQUEADORA NÃO ACHADA')   # F também é de uma unidade
            E.append(ent(aba, i, nomes, [] if unidade else [r[ix['CNPJ']]], [] if unidade else [r[ix['RAZÃO SOCIAL']]],
                         raizes_ref=[x for x in (ref, r[ix['CNPJ']] if unidade else None) if x],
                         extra=dict(nome=s(r[1]), uf=s(r[ix['UF SEDE']]), status_antes=s(r[ix['STATUS AUTOPEL']]) if 'STATUS AUTOPEL' in ix else '',
                                    apoio=aba == 'OUTROS FINAL', origem=s(r[ix['ORIGEM']]) if 'ORIGEM' in ix else '')))
    # ---- Alessandra Dayan: 11 tabelas lado a lado (linha 1 = segmento, linha 2 = cabeçalho)
    ws = wb['Alessandra Dayan']
    rows = list(ws.iter_rows(values_only=True))
    for j, v in enumerate(rows[1]):
        if s(v) != 'EMPRESA': continue
        seg = s(rows[0][j]); cab = [s(x) for x in rows[1][j:j + 8]]
        cst = j + cab.index('STATUS AUTOPEL')
        for i, r in enumerate(rows[2:], start=3):
            if j < len(r) and s(r[j]):
                E.append(ent('Alessandra Dayan', i, nomes_da_rede(r[j]), extra=dict(nome=s(r[j]), seg=seg, col=L(j + 1), col_status=L(cst + 1),
                                                                                 uf=s(r[j + 1]), status_antes=s(r[cst]))))
    # ---- Alessandra Dayan VF e TOP 50 (EMPRESA, UF, ..., STATUS AUTOPEL, SEGMENTO)
    for aba in ['Alessandra Dayan VF'] + TOP50:
        rows = list(wb[aba].iter_rows(values_only=True))
        # 19/11: lê pelo cabeçalho (desde 17/11 o cabeçalho é a linha 6, na VF a 8, e a VF ganhou colunas na frente);
        # a coluna CNPJ (matriz da marca, 17/11) entra como raiz
        h = next(i for i, r in enumerate(rows[:12]) if r and 'EMPRESA' in [s(x) for x in r])
        cab = [s(x) for x in rows[h]]
        ix = {c: cab.index(c) for c in reversed(cab) if c}
        iE, iU, iS, iG, iC = ix['EMPRESA'], ix['UF'], ix['STATUS AUTOPEL'], ix.get('SEGMENTO'), ix.get('CNPJ')
        g = lambda r, j: s(r[j]) if j is not None and j < len(r) else ''
        for i, r in enumerate(rows[h + 1:], start=h + 2):
            if not r or not g(r, iE): break      # 19/11: fim da tabela (embaixo há dinâmicas do usuário)
            cn = [g(r, iC)] if cnpj_ok(g(r, iC)) else []
            E.append(ent(aba, i, nomes_da_rede(r[iE]), cn, extra=dict(nome=g(r, iE), seg=g(r, iG), uf=g(r, iU), status_antes=g(r, iS),
                                                                       col_status=L(iS + 1))))
    # ---- IRTT por bandeira (cabeçalho na linha 4)
    rows = list(wb['I.R.T.T.__Por bandeira (FINAL)'].iter_rows(values_only=True))
    cab = [s(x) for x in rows[3]]
    for i, r in enumerate(rows[4:], start=5):
        if not r or not s(r[1]): continue
        cods = re.findall(r'\b\d{6}\b', s(r[18]))
        E.append(ent('I.R.T.T.__Por bandeira (FINAL)', i, nomes_da_rede(r[1]), [r[24], r[20]], [r[21]], codigos=cods, grupos=[g.strip() for g in s(r[17]).split(',') if g.strip()],
                     extra=dict(nome=s(r[1]), grupo=s(r[0]), seg=s(r[2]), uf=s(r[16]), status_antes=s(r[4]), data_antes=r[6], vend_antes=s(r[32]))))
    # ---- SUPERMERCADOS (reuniao): só apoio de CNPJ (não é gravada); cabeçalho na linha com 'Nome fantasia'
    ws = wb['SUPERMERCADOS (reuniao)']
    rows = list(ws.iter_rows(values_only=True))
    h = next(i for i, r in enumerate(rows[:10]) if r and 'Nome fantasia' in [s(x) for x in r])
    cab = [s(x) for x in rows[h]]
    iN, iR, iC, iU = cab.index('Nome fantasia'), cab.index('Razão social'), cab.index('CNPJ'), cab.index('UF')
    for i, r in enumerate(rows[h + 1:], start=h + 2):
        if r and s(r[iN]):
            E.append(ent('SUPERMERCADOS (reuniao)', i, nomes_da_rede(r[iN]), [r[iC]], [r[iR]],
                         extra=dict(nome=s(r[iN]), seg='SUPERMERCADO', uf=s(r[iU]), apoio=True)))
    # ---- FARMACIA VF (cabeçalho na linha 6)
    # 16/11/2026: colunas pelo cabeçalho. Depois que as colunas I e J saíram, os índices fixos liam a SITUAÇÃO NA RECEITA
    # ("ATIVA") como UF, e a trava de UF das farmácias deixava de valer.
    ws = wb['"FARMACIA VF"']
    cab = [norm(s(x)) for x in next(ws.iter_rows(min_row=6, max_row=6, values_only=True))]
    col = lambda nome: cab.index(norm(nome))
    iR, iT, iF, iZ, iC = col('REDE / ASSOCIAÇÃO'), col('TIPO'), col('NOME FANTASIA'), col('RAZÃO SOCIAL'), col('CNPJ')
    iU, iS, iD = col('UF SEDE'), col('STATUS AUTOPEL'), col('DATA ÚLTIMA COMPRA')
    rows = ws.iter_rows(min_row=7, values_only=True)
    for i, r in enumerate(rows, start=7):
        if not r or (not s(r[iF]) and not s(r[iZ]) and not s(r[iR])): continue
        tipo = s(r[iT])
        if tipo == 'Farmácia':
            E.append(ent('"FARMACIA VF"', i, [], [r[iC]], extra=dict(nome=s(r[iF]) or s(r[iZ]), tipo=tipo, rede=s(r[iR]), uf=s(r[iU]),
                                                                        status_antes=s(r[iS]), data_antes=r[iD], so_cnpj=True)))
        else:
            E.append(ent('"FARMACIA VF"', i, nomes_da_rede(r[iR], r[iF]), [r[iC]], [r[iZ]],
                         extra=dict(nome=s(r[iR]) or s(r[iF]), tipo=tipo, uf=s(r[iU]), status_antes=s(r[iS]), data_antes=r[iD])))
    wb.close()
    # ---- clusters: a mesma rede em abas diferentes, para ter UM status. Trava (16/11): só no mesmo segmento, e nunca
    # duas linhas com CNPJs diferentes (Farmácia Rosário x Rosário da D1000; Center Park x Center Castilho).
    for e in E: e['fam'] = familia(e)
    pai = list(range(len(E)))
    def f(x):
        while pai[x] != x: pai[x] = pai[pai[x]]; x = pai[x]
        return x
    cp = lambda n: n.replace(' ', '')
    raizes_do = {}
    def junta(a, b, forca=False):
        ra, rb = f(a), f(b)
        if ra == rb: return
        za, zb = raizes_do.get(ra, set()), raizes_do.get(rb, set())
        if za and zb and not (za & zb) and not forca: return     # CNPJs diferentes: outra empresa
        pai[ra] = rb; raizes_do[rb] = za | zb
    for n, e in enumerate(E): raizes_do[n] = set(e['raizes'])
    chave = C.defaultdict(list)
    for n, e in enumerate(E):
        if e['extra'].get('so_cnpj') or not e['nomes']: continue
        chave[(e['fam'], cp(e['nomes'][0]))].append(n)
    from motor import valido
    for n, e in enumerate(E):     # "MILKY MOO MILKSHAKES" x "MILKY MOO": prefixo de 2+ palavras ou 6+ letras, não genérico
        if e['extra'].get('so_cnpj') or not e['nomes']: continue
        w = e['nomes'][0].split()
        for i in range(1, len(w)):
            pre = ' '.join(w[:i])
            if (i >= 2 or len(pre) >= 6) and valido(pre) and (e['fam'], cp(pre)) in chave:
                chave[(e['fam'], cp(pre))].append(n)
    for k, ns in chave.items():
        for a in ns[1:]: junta(a, ns[0])
    # 19/11/2026 (caso "Koni Store" x "Koni" da IRTT, mesma raiz): com o MESMO CNPJ, o prefixo curto (4+ letras) também
    # junta. Sem isso a linha "KONI STORE" ficava NÃO É CLIENTE e a "KONI" ATIVO pela bobina "80X100M KONI".
    por_raiz = C.defaultdict(list)
    for n, e in enumerate(E):
        if e['extra'].get('so_cnpj') or not e['nomes']: continue
        for r in e['raizes']: por_raiz[r].append(n)
    for r, ns in por_raiz.items():
        for a in ns:
            for b in ns:
                if a < b:
                    xa, xb = cp(E[a]['nomes'][0]), cp(E[b]['nomes'][0])
                    curto, longo = sorted((xa, xb), key=len)
                    if len(curto) >= 4 and longo.startswith(curto): junta(a, b)
                    # 19/11: "MELISSA" x "CLUBE MELISSA", "SPOLETO" x "PIZZARIA SPOLETO": nome inteiro dentro do outro (5+ letras)
                    na, nb = sorted((E[a]['nomes'][0], E[b]['nomes'][0]), key=len)
                    if (len(na.replace(' ', '')) >= 5 and valido(na) and set(nb.split()) - set(na.split()) <= FORMATO
                            and re.search(r'(?<![A-Z0-9])' + re.escape(na) + r'(?![A-Z0-9])', nb)): junta(a, b)
    # Alessandra Dayan (e VF / TOP 50, mesmos nomes) -> linhas das outras abas do cruzamento de 05/11/2026. As abas FINAL
    # andaram 5 linhas em 06/11 (cabeçalho foi para a linha 6): aceita a linha N ou N+5, conferindo o nome.
    import json, os
    arq = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'mapa_final', 'dados', 'alessandra_cruzamento.json')
    if os.path.exists(arq):
        idx = {(e['aba'], e['linha']): n for n, e in enumerate(E)}
        ale = C.defaultdict(list)
        for n, e in enumerate(E):
            if e['aba'] in ('Alessandra Dayan', 'Alessandra Dayan VF') or e['aba'] in TOP50: ale[norm(e['extra']['nome'])].append(n)
        ligados = 0
        idx.update({(e['aba'], e['linha']): n for n, e in enumerate(E)})
        # 16/11/2026 (conferência com o CRM): a palavra em comum tem de ser própria do nome. Com "PARK", "POSTO", "SORVETES"
        # a linha N (vizinha, 5 acima) também batia, e a rede herdava a compra da vizinha. Palavra de tipo não conta;
        # fica a linha (N ou N+5) com o nome igual ou, senão, com mais palavras próprias em comum.
        VAZIAS = {'DE', 'DA', 'DO', 'DOS', 'DAS', 'E', 'LOJAS', 'LOJA', 'REDE', 'GRUPO', 'POSTO', 'POSTOS', 'AUTO', 'PARK',
                  'PARKING', 'ESTACIONAMENTO', 'ESTACIONAMENTOS', 'SORVETES', 'MATERIAIS', 'CONSTRUCAO', 'SUPERMERCADO',
                  'CONVENIENCIA', 'PET', 'SHOP', 'CENTER', 'CASA', 'EXPRESS', 'COMERCIO', 'BRASIL'}
        def proprias(aba, txt):
            return {w for w in norm(txt).split() if w not in VAZIAS and len(w) > 1}
        for x in json.load(open(arq, encoding='utf8')):
            ns = ale.get(norm(x['nome']), [])
            for o in x.get('onde', []):
                m = re.match(r'(.+?) linha (\d+) \((.*)\)$', o)
                if not m or not ns: continue
                aba_o, lin, nome_o = m.group(1), int(m.group(2)), m.group(3)
                po = proprias(aba_o, nome_o)
                melhor = None
                for dl in ((0, 5) if aba_o.endswith('FINAL') else (0,)):     # só as abas FINAL andaram 5 linhas
                    alvo = idx.get((aba_o, lin + dl))
                    if alvo is None: continue
                    nm = E[alvo]['extra']['nome']
                    q = 100 if norm(nm) == norm(nome_o) else len(proprias(aba_o, nm) & po)
                    if q and (melhor is None or q > melhor[0]): melhor = (q, alvo)
                if melhor:
                    for a2 in ns: junta(a2, melhor[1], forca=True)     # ligação curada em 05/11 (ex.: Clamed + Preço Popular)
                    ligados += 1
        print('cruzamento Alessandra de 05/11:', ligados, 'ligações')
    # a mesma marca em segmentos diferentes (Kopenhagen em ALIMENTAÇÃO e no IRTT; Bob's em ALIMENTAÇÃO e LOJAS): junta pelo
    # nome igual, a não ser que os dois lados tenham CNPJs e eles não se cruzem (Pague Menos drogaria x supermercado)
    por_nome = C.defaultdict(list)
    for (fam, nome), ns in chave.items(): por_nome[nome].append(ns[0])
    for nome, ns in por_nome.items():
        for a in ns[1:]: junta(a, ns[0])
    for n, e in enumerate(E): e['cluster'] = f(n)
    pickle.dump(E, open(saida, 'wb'))
    cnt = C.Counter(e['aba'] for e in E)
    print(len(E), 'linhas |', len({e['cluster'] for e in E if not e['extra'].get('so_cnpj')}), 'redes (clusters)')
    for a, n in cnt.items(): print(f'  {a}: {n}')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
