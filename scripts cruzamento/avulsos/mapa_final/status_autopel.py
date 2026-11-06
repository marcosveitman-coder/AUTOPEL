"""STATUS RECEITA / PORTE / STATUS AUTOPEL (05/11/2026) - passo 2: ultima compra de cada rede na Autopel.
Fontes:
  - bases de faturamento 2019-2026 (cache agg: por raiz de CNPJ, por cliente+UF e por produto+UF), ate o fim delas;
  - relatorio "Ultimo faturamento - SQL.csv" do TOTVS (DT_FAT por cliente): so depois do fim das bases (regra 8 do
    CLAUDE.md), com o grupo do cliente (coluna Descricao), e-mail e segmento;
  - "Clientes - SQL.csv" (cadastro): vendedor da carteira (codigo -> nome pelo Ultimo faturamento).
Vinculos (ordem de forca do CLAUDE.md):
  fortes: raiz do CNPJ da rede (F), raiz do CNPJ de referencia (H, unidade da rede), CNPJ das abas de apoio (FARMACIA VF,
          IRTT, reuniao) e razao social igual (so sem CNPJ da rede ou com a mesma raiz: razao igual nao vence CNPJ);
  grupo:  grupo do cadastro (Descricao do Ultimo faturamento) com o nucleo igual ao da rede, cliente de segmento compativel
          e que nao seja distribuidor;
  fracos: nome da rede no cliente, no fim da descricao do produto e no dominio do e-mail - motor do ../irtt.py (marca
          colada/separada, 5+ letras, marca comum exige UF e frase completa, segmento do comprador compativel).
  Um fraco so substitui o forte se for mais recente e o comprador tiver o nome da rede. Empresa achada por fraco: data da
  ultima venda da raiz dela. CNPJ validado de outra linha (outra rede) nunca vale.
Status: ATIVO = ultima compra ha menos de 360 dias (contados de HOJE); INATIVO = comprou, mas ha 360 dias ou mais;
NAO E CLIENTE = nenhuma compra achada desde 2019.
Entrada: dados/status_itens.pkl (status_coleta.py).  Saida: dados/status_autopel.pkl e status_autopel_view.txt.
Variaveis: AUTOPEL_ULTFAT_CSV, AUTOPEL_CLIENTES_CSV, AUTOPEL_HOJE (AAAA-MM-DD)."""
import os, sys, re, pickle, datetime as dt, collections as C
import openpyxl, pandas as pd
HERE = os.path.dirname(os.path.abspath(__file__))
DADOS = os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or os.path.join(HERE, 'dados')
sys.path.insert(0, os.path.join(HERE, '..')); sys.path.insert(0, os.path.join(HERE, '..', '..'))
import common, match, irtt
from common import norm_legal
BASE = common.BASE
ULTFAT = os.environ.get('AUTOPEL_ULTFAT_CSV', os.path.join(BASE, 'Ultimo faturamento - SQL.csv'))
CLIENTES = os.environ.get('AUTOPEL_CLIENTES_CSV', 'Clientes - SQL.csv')
HOJE = dt.datetime.strptime(os.environ.get('AUTOPEL_HOJE', dt.date.today().isoformat()), '%Y-%m-%d')
DIAS_ATIVO = 360
INICIO = 20190101
D = lambda x: re.sub(r'\D', '', str(x or ''))
s = lambda x: str(x).strip() if x is not None else ''
AP = lambda x: re.sub(r"['’`´]", '', str(x or ''))
ymd = lambda n: dt.datetime.strptime(str(n), '%Y%m%d')

# segmentos do comprador aceitos em cada aba (vinculos fracos e de grupo)
SEGS = {
    'MF_ALIM': {'ALIMENTACAO', 'ENTRETENIMENTO', 'REDE DE LOJAS', 'POSTOS E CONVENIENCIAS'},
    'MF_LOJAS': {'REDE DE LOJAS', 'MAGAZINES', 'E-COMMERCE', 'COSMETICOS', 'PET SHOP'},
    'MF_CONST': {'CONSTRUCAO', 'REDE DE LOJAS', 'MAGAZINES'},
    'MF_POSTOS': {'POSTOS E CONVENIENCIAS', 'ALIMENTACAO', 'REDE DE LOJAS'},
    'MF_PET': {'PET SHOP', 'REDE DE LOJAS', 'CORPORATIVO SAUDE'},
    'MF_ESTAC': {'ESTACIONAMENTOS'},
    'MF_BELEZA': {'COSMETICOS', 'CORPORATIVO SAUDE', 'REDE DE LOJAS', 'ENTRETENIMENTO', 'DROGARIAS'},
    'MF_OTICA': {'REDE DE LOJAS', 'CORPORATIVO SAUDE', 'MAGAZINES'},
    'MF_SERV': {'CORPORATIVO', 'CORPORATIVO FINANCEIRO', 'CORPORATIVO EDUCACIONAL', 'CORPORATIVO SAUDE', 'TRANSPORTE', 'LOGISTICA',
                'AUTOMOTIVO PEÇAS E LOCADORAS', 'REDE DE LOJAS', 'ENTRETENIMENTO', 'E-COMMERCE'},
    'Drogarias e Perfumarias': irtt.SEG['Drogarias e Perfumarias'],
    'Super, Hiper, Atacarejo e Conveniência': irtt.SEG['Super, Hiper, Atacarejo e Conveniência'],
}
SEG_ABA = {'ALIMENTACAO FINAL': 'MF_ALIM', 'REDE DE LOJAS FINAL': 'MF_LOJAS', 'CONSTRUÇÃO FINAL': 'MF_CONST',
           'POSTOS GAS & CONVENIENCIA FINAL': 'MF_POSTOS', 'PET SHOP  FINAL': 'MF_PET', 'ESTACIONAMENTO FINAL': 'MF_ESTAC'}
SEG_OUTROS = {'Beleza / Saúde / Bem-estar': 'MF_BELEZA', 'Óticas': 'MF_OTICA', 'Serviços / Franquias': 'MF_SERV'}
SEG_ALE = {'DROGARIAS': 'Drogarias e Perfumarias', 'SUPERMERCADOS / ATACAREJO': 'Super, Hiper, Atacarejo e Conveniência',
           'BELEZA / SAÚDE / BEM-ESTAR': 'MF_BELEZA', 'ÓTICAS': 'MF_OTICA', 'ALIMENTAÇÃO': 'MF_ALIM'}
GEN = {'GRUPO', 'REDE', 'LOJAS', 'LOJA', 'SUPERMERCADO', 'SUPERMERCADOS', 'DE', 'DO', 'DA', 'DOS', 'DAS', 'E', 'THE', 'S', 'A', 'SA',
       'DROGARIA', 'DROGARIAS', 'FARMACIA', 'FARMACIAS', 'ATACADO', 'ATACADISTA', 'RESTAURANTE', 'RESTAURANTES', 'CALCADOS', 'MODAS',
       'POSTO', 'POSTOS', 'AUTO', 'COMERCIO', 'COMERCIAL', 'OTICA', 'OTICAS', 'ESTACIONAMENTO', 'ESTACIONAMENTOS', 'CIA', 'LTDA', 'ME'}
DISTRIB = re.compile(r'DISTRIBUI|ATACADIST|LOGISTIC|COURIER|TRANSPORT|REPRESENTA')
# comprador que revende papel/embalagem: produto personalizado comprado por ele nao e compra da rede
REVENDE = re.compile(r'DISTRIBUI|ATACADIST|LOGISTIC|SUPRI|PRODUTOS PARA|EMBALAG|PAPEL|PAPEIS|DESCARTAV|REPRESENTA')
GEN_GRP = {'GRUPO', 'REDE', 'LOJAS', 'LOJA', 'DE', 'DO', 'DA', 'DOS', 'DAS', 'E', 'THE', 'S', 'A', 'SA', 'CIA', 'LTDA', 'ME', 'POSTO', 'POSTOS',
           'SUPERMERCADO', 'SUPERMERCADOS', 'DROGARIA', 'DROGARIAS', 'FARMACIA', 'FARMACIAS', 'COMERCIO', 'COMERCIAL'}


def nucleo_grp(x):
    w = re.findall(r'[A-Z0-9]+', common.norm(AP(x)))
    while w and w[0] in GEN_GRP: w.pop(0)
    while w and w[-1] in GEN_GRP: w.pop()
    return ' '.join(w)


def nucleo(x):
    w = re.findall(r'[A-Z0-9]+', common.norm(AP(x)))
    while w and w[0] in GEN: w.pop(0)
    while w and w[-1] in GEN: w.pop()
    return ' '.join(w)


def apelidos(*nomes):
    """nomes de busca: separa "X (+ Y)", "X / Y", "X + Y" e tira o texto entre parenteses que nao e marca"""
    out = []
    for n in nomes:
        n = s(n)
        if not n: continue
        n = re.sub(r'\((?!\s*\+)[^)]*\)', ' ', n)  # "(postos)", "(RD SAUDE)": descricao entre parenteses sai
        for p in re.split(r'\s*\(\+\s*|\)\s*|\s+/\s+|\s+\+\s+|;\s*|,\s*', n):
            p = p.strip(' -–()')
            if len(nucleo(p).replace(' ', '')) >= 5 and p.upper() not in {x.upper() for x in out}:
                out.append(p)
    return out


def main():
    I = pickle.load(open(os.path.join(DADOS, 'status_itens.pkl'), 'rb'))
    # ---- itens a cruzar: linhas das abas FINAL + empresas da Alessandra que nao estao numa aba FINAL
    itens = []
    for it in I['itens']:
        seg = SEG_OUTROS.get(s(it.get('segp')), 'MF_SERV') if it['aba'] == 'OUTROS FINAL' else SEG_ABA[it['aba']]
        if it['aba'] == 'REDE DE LOJAS FINAL' and re.search(r'PET', common.norm(it.get('segm'))): seg = 'MF_PET'
        cn = [(it['cnpj'], 'CNPJ da rede (coluna F)')] if len(it['cnpj']) == 14 else []
        if len(it['ref']) == 14 and it['ref'] != it['cnpj']: cn.append((it['ref'], 'CNPJ de uma unidade da rede (coluna H)'))
        if cn and 'UNIDADE' in s(it.get('conf')).upper():
            cn = [(c, 'CNPJ de uma unidade da rede') for c, _ in cn]
        itens.append(dict(key=(it['aba'], it['lin']), nome=s(it.get('nome')), grupo=s(it.get('grupo')), raz=s(it.get('raz')),
                          seg=seg, ufs={s(it.get('uf'))} - {''}, cnpjs=cn, w_ant=s(it.get('W')), apoio=[]))
    for a in I['ale']:
        if a['liga']: continue
        cn, ufs = [], {a['uf']} if re.fullmatch(r'[A-Z]{2}', a['uf']) else set()
        for x in a['apoio']:
            if len(x['cnpj']) == 14 and x['cnpj'] not in [c for c, _ in cn]: cn.append((x['cnpj'], f"CNPJ da aba {x['fonte']} (linha {x['lin']})"))
            if x.get('uf'): ufs.add(x['uf'])
        itens.append(dict(key=('Alessandra Dayan', a['ref']), nome=a['nome'], grupo='', raz='', seg=SEG_ALE[a['seg']], ufs=ufs, cnpjs=cn,
                          w_ant='', apoio=a['apoio']))
    print(len(itens), 'redes a cruzar')
    dono = C.defaultdict(set)  # raiz validada -> redes
    for it in itens:
        for c, how in it['cnpjs']:
            if 'unidade' not in how: dono[c[:8]].add(it['key'])

    # ---- bases de faturamento (2019+) e Ultimo faturamento de hoje
    root, cli, desc = irtt.load_agg_2019()
    fim = ymd(max(v[0] for v in root.values()))
    U = pd.read_csv(ULTFAT, sep=';', skiprows=1, dtype=str, encoding='utf-8-sig', keep_default_na=False)
    U = U.apply(lambda col: col.str.strip())
    uf_raiz, uf_velha, cadastrada, uf_cli, grp = {}, {}, set(), [], C.defaultdict(list)
    vend_nome = {}
    for r in U.itertuples(index=False):
        cod, cnpj = r[0], D(r[2])
        try: d = dt.datetime.strptime(r[24], '%d/%m/%Y')
        except ValueError: d = None
        if r[14] and r[15]: vend_nome[r[14]] = r[15]
        x = dict(cod=cod, loja=r[1], cnpj=cnpj, grupo=r[4], nome=r[5], fant=r[6], seg=r[13], vend=r[15], dt=d, uf=r[28])
        uf_cli.append(x)
        if len(cnpj) == 14 and d and d > fim and (cnpj[:8] not in uf_raiz or d > uf_raiz[cnpj[:8]]['dt']):
            uf_raiz[cnpj[:8]] = x
        if len(cnpj) == 14 and d and d <= fim and (cnpj[:8] not in uf_velha or d > uf_velha[cnpj[:8]]['dt']):
            uf_velha[cnpj[:8]] = x
        if len(cnpj) == 14: cadastrada.add(cnpj[:8])
        g = nucleo_grp(re.sub(r'\s+-\s+.*$', '', r[4]))
        if g and common.norm(r[4]) not in ('CLIENTES DIVERSOS',): grp[g].append(x)
    print('fim das bases:', fim.date(), '| Ultimo faturamento:', len(U), 'clientes, última nota',
          max(x['dt'] for x in uf_cli if x['dt']).date(), '|', len(uf_raiz), 'raízes com nota depois do fim das bases')
    cad = pd.read_csv(CLIENTES, sep=';', skiprows=1, dtype=str, encoding='utf-8-sig', keep_default_na=False).apply(lambda c: c.str.strip())
    vend_cad, vend_raiz = {}, C.defaultdict(C.Counter)
    for r in cad.itertuples(index=False):
        v = vend_nome.get(r[5], r[5])
        vend_cad[(r[0], r[1])] = v
        if len(D(r[4])) == 14: vend_raiz[D(r[4])[:8]][v] += 1; cadastrada.add(D(r[4])[:8])
    cod_raiz = {}
    for x in uf_cli:
        if len(x['cnpj']) == 14: cod_raiz.setdefault(x['cnpj'][:8], x)

    def data_raiz(r8):
        """(data, cliente, cnpj, fonte) da ultima venda da raiz"""
        best = None
        f = root.get(r8)
        if f and f[0] >= INICIO: best = (ymd(f[0]), f[2], f[3], 'bases de faturamento')
        u = uf_raiz.get(r8)
        if u and (best is None or u['dt'] > best[0]): best = (u['dt'], u['nome'], u['cnpj'], f'Ultimo faturamento (nota de {u["dt"]:%d/%m/%Y})')
        return best

    n_grupo = C.Counter(nucleo(it['grupo']) for it in itens if it['grupo'])
    for it in itens:
        it['busca'] = apelidos(it['nome'], it['grupo'] if it['grupo'] and n_grupo[nucleo(it['grupo'])] == 1 else '')

    # ---- vinculos fracos: motor do irtt.py numa planilha no layout dele (uma linha por nome de busca)
    for k, v in SEGS.items(): irtt.SEG.setdefault(k, v)
    wo = openpyxl.Workbook(); ws = wo.active
    ws.append(['grupo', 'bandeira', 'seg'] + [''] * 18)
    linhas = []
    for n, it in enumerate(itens):
        for ap in it['busca']:
            row = [it['grupo'], ap, it['seg']] + [None] * 18
            row[15] = ', '.join(sorted(it['ufs']))
            row[20] = it['raz'] if not it['cnpjs'] else None  # razao igual so sem CNPJ (com CNPJ, a raiz ja cobre)
            ws.append(row); linhas.append(n)
    tmp = os.path.join(DADOS, 'status_irtt_layout.xlsx'); wo.save(tmp)
    irtt.IRTT, irtt.HERE, irtt.CSV = tmp, DADOS, ULTFAT
    irtt.main()
    R = pickle.load(open(os.path.join(DADOS, 'cand_irtt.pkl'), 'rb'))
    fracos = C.defaultdict(list)
    for k, c in enumerate(R['cand']):
        for how, d in c.items():
            if how == 'cnpj': continue
            fracos[linhas[k]].append((how, d, R['data'][k][1]))

    disputa = C.defaultdict(set)  # raiz do comprador de vinculo fraco -> redes que o reclamam
    for n, l in fracos.items():
        for how, d, ap in l:
            if how != 'razao' and len(d['cnpj']) == 14: disputa[d['cnpj'][:8]].add(n)

    # ---- escolha por rede
    LBL = {'razao': 'razão social igual', 'fantasia': 'nome da rede no cliente', 'desc': 'nome da rede no produto personalizado',
           'email': 'domínio do e-mail do cadastro'}
    out, view, rej, rej_ex = {}, [], C.Counter(), C.Counter()
    for n, it in enumerate(itens):
        fortes, fr, cad_sem = [], [], []
        for c, how in it['cnpjs']:
            b = data_raiz(c[:8])
            if b: fortes.append(dict(dt=b[0], cliente=b[1], cnpj=b[2], fonte=b[3], via=how, brand=True))
            elif c[:8] in cadastrada: cad_sem.append((c, how))
        for a in it['apoio']:  # data ja validada nas abas de apoio (FARMACIA VF / IRTT / reuniao), com associadas e bandeiras
            if isinstance(a.get('data'), dt.datetime):
                fortes.append(dict(dt=a['data'], cliente=a['nome'], cnpj=a['cnpj'], fonte=f"aba {a['fonte']} (linha {a['lin']})",
                                   via=f"data da aba {a['fonte']}", brand=True))
        # grupo do cadastro com o nome da rede
        for ap in it['busca']:
            nb = nucleo_grp(ap)
            if len(nb.replace(' ', '')) < 5 or not match.specific(nb.split()): continue
            for x in grp.get(nb, []):
                if len(x['cnpj']) != 14 or x['seg'] not in SEGS[it['seg']] or DISTRIB.search(common.norm(x['nome'])): continue
                if dono.get(x['cnpj'][:8]) and it['key'] not in dono[x['cnpj'][:8]]:
                    rej['grupo: CNPJ de outra rede'] += 1; rej_ex[(it['nome'][:30], x['grupo'][:30], str(sorted(dono[x['cnpj'][:8]]))[:60])] += 1; continue
                b = data_raiz(x['cnpj'][:8])
                if b: fortes.append(dict(dt=b[0], cliente=b[1], cnpj=b[2], fonte=b[3], via=f"grupo do cadastro {x['grupo']}", brand=True))
                else: cad_sem.append((x['cnpj'], f"grupo do cadastro {x['grupo']}"))
        for how, d, ap in fracos.get(n, []):
            r8 = d['cnpj'][:8] if len(d['cnpj']) == 14 else ''
            if r8 and dono.get(r8) and it['key'] not in dono[r8]: rej[f'{how}: CNPJ de outra rede'] += 1; continue
            if how == 'razao' and it['cnpjs'] and r8 and r8 not in {c[:8] for c, _ in it['cnpjs']}: rej['razão: homônimo'] += 1; continue
            if how != 'razao' and not d.get('buyer_has_brand') and REVENDE.search(common.norm(d['cliente'])):
                rej[f'{how}: comprador revende (distribuidor)'] += 1; continue
            if how != 'razao' and r8 and len(disputa.get(r8, ())) > 1 and d.get('uf') not in it['ufs']:
                rej[f'{how}: comprador disputado por outra rede (UF diferente)'] += 1; continue
            b = data_raiz(r8) if r8 else None
            dd = b[0] if b and b[0] > ymd(d['dt']) else ymd(d['dt'])
            x = dict(dt=dd, cliente=b[1] if b and b[0] > ymd(d['dt']) else d['cliente'], cnpj=d['cnpj'], fonte=b[3] if b and b[0] > ymd(d['dt']) else 'bases de faturamento',
                     via=f"{LBL[how]} ({ap}" + (f": {d.get('produto')}" if how == 'desc' else f": {d.get('dominio')}" if how == 'email' else '') + ')',
                     brand=how == 'razao' or d.get('buyer_has_brand', False))
            (fortes if how == 'razao' else fr).append(x)
        best = max(fortes, key=lambda x: x['dt']) if fortes else None
        if fr:
            wb_ = max(fr, key=lambda x: x['dt'])
            if best is None or (wb_['dt'] > best['dt'] and wb_['brand']):
                best = wb_
        if best:
            dias = (HOJE - best['dt']).days
            st = 'ATIVO' if dias < DIAS_ATIVO else 'INATIVO'
            r8 = D(best['cnpj'])[:8]
            vend = vend_raiz[r8].most_common(1)[0][0] if vend_raiz.get(r8) else ''
        elif cad_sem:
            # cadastrado no TOTVS (codigo de cliente) sem venda desde 2019: INATIVO. Nota antiga do Ultimo faturamento vira a data.
            velhas = [(uf_velha[c[:8]], how) for c, how in cad_sem if c[:8] in uf_velha]
            if velhas:
                u, how = max(velhas, key=lambda v: v[0]['dt'])
                best = dict(dt=u['dt'], cliente=u['nome'], cnpj=u['cnpj'], fonte=f"Ultimo faturamento (nota de {u['dt']:%d/%m/%Y}, fora das bases de venda)", via=how)
                dias = (HOJE - u['dt']).days
            else:
                c, how = cad_sem[0]
                best = dict(dt=None, cliente=cod_raiz.get(c[:8], {}).get('nome', ''), cnpj=c, fonte='cadastrado no TOTVS, sem nota', via=how)
                dias = None
            st = 'INATIVO'
            r8 = D(best['cnpj'])[:8]
            vend = vend_raiz[r8].most_common(1)[0][0] if vend_raiz.get(r8) else ''
        else:
            dias, st, vend = None, 'NÃO É CLIENTE', ''
        # alerta: nota recente no Ultimo faturamento de uma raiz forte que nao aparece nas bases de venda
        alerta = ''
        for c, how in it['cnpjs']:
            u = uf_velha.get(c[:8])
            if u and (HOJE - u['dt']).days < DIAS_ATIVO and st != 'ATIVO':
                alerta = f"nota de {u['dt']:%d/%m/%Y} no Ultimo faturamento que não está nas bases de venda"
        out[it['key']] = dict(status=st, dt=best['dt'] if best else None, dias=dias, cliente=best['cliente'] if best else '',
                              cnpj=best['cnpj'] if best else '', fonte=best['fonte'] if best else '', via=best['via'] if best else '',
                              vend=vend, w_ant=it['w_ant'], n_fortes=len(fortes), n_fracos=len(fr), nome=it['nome'], alerta=alerta)
        o = out[it['key']]
        view.append(f"{it['key'][0][:22]:22} {str(it['key'][1]):>5} {it['nome'][:34]:34} {st:13} {o['dt'].date() if o['dt'] else '':10} {alerta[:12]}"
                    f"ant={it['w_ant'][:12]:12} | {o['via'][:70]} | {o['cliente'][:40]}")
    pickle.dump(dict(out=out, fim=fim, hoje=HOJE, rej=rej), open(os.path.join(DADOS, 'status_autopel.pkl'), 'wb'))
    open(os.path.join(DADOS, 'status_autopel_view.txt'), 'w', encoding='utf8').write('\n'.join(view))
    print('rejeitados:', dict(rej))
    for k, v in rej_ex.most_common(15): print('   ', v, k)
    print('alertas (nota recente fora das bases):', sum(1 for o in out.values() if o['alerta']))
    print('status:', C.Counter(o['status'] for o in out.values()))
    print('anterior x novo:', sorted(C.Counter((o['w_ant'], o['status']) for o in out.values() if o['w_ant']).items()))


if __name__ == '__main__':
    main()
