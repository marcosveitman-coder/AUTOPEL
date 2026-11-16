"""Avalia o status de cada rede (16/11/2026): roda o motor (motor.py) por cluster (a mesma rede nas várias abas) com a
união dos nomes, CNPJs, razões, códigos de cliente e grupos do ERP das linhas do cluster.
Caminhos extras além do motor:
  CODIGO    códigos de cliente que a aba IRTT indica (coluna S);
  GRUPO ERP grupo do ERP indicado na aba IRTT (coluna R) = descrição do grupo do cliente;
  EXATO     para nome curto/genérico ("C&A", "ABC"): razão, fantasia ou grupo do cliente IGUAL ao nome.
Trava: cliente cuja raiz de CNPJ é de OUTRA linha das abas (com CNPJ validado) não vale por nome/e-mail/produto
(Farmácia Rosário x Drogaria Rosário).
Farmácia avulsa da FARMACIA VF: só a raiz do CNPJ.
Entrada: clientes.pkl (base_clientes.py) e entidades.pkl (coleta.py). Saída: resultados.pkl.
Uso: python avalia.py "<pasta de trabalho>" [AAAA-MM-DD de hoje] """
import sys, os, re, pickle, datetime as dt, collections as C
from motor import Motor, norm, norm_legal, COMPAT, GRUPO_GENERICO

PASTA = sys.argv[1]
HOJE = dt.date.fromisoformat(sys.argv[2]) if len(sys.argv) > 2 else dt.date.today()


# vínculos manuais que nenhuma regra acha: nome de busca -> nomes a mais (dados/manuais.json, local)
from motor import ALIAS_MANUAL, RAIZES_ERRADAS, CLIENTES_MANUAIS, EXCLUSOES, UFS, FAM_EXTRA, TIPO, SEPARAR, SO_FORTES, GRUPOS_OK, NACIONAIS


class Proibidas:
    def __init__(self, todas, proprias): self.t, self.p = todas, proprias
    def __contains__(self, r): return bool(r) and r in self.t and r not in self.p


def main():
    B = pickle.load(open(os.path.join(PASTA, 'clientes.pkl'), 'rb'))
    E = pickle.load(open(os.path.join(PASTA, 'entidades.pkl'), 'rb'))
    M = Motor(B['cli'], B['vend'], HOJE)
    cli = B['cli']
    por_cod = C.defaultdict(list)
    for k in cli: por_cod[k[0]].append(k)
    por_grp = C.defaultdict(list)
    for k, c in cli.items():
        if c['_grp']: por_grp[c['_grp']].append(k)
    exato = C.defaultdict(set)
    for k, c in cli.items():
        for t in (c['fant'], c['grp_desc']):
            n = norm_legal(t)
            if n: exato[n].add(k)
    todas = {r for e in E for r in e['raizes'] if not e['extra'].get('so_cnpj')}   # só CNPJ de redes; farmácia avulsa não trava
    # FARMACIA VF: associadas de cada rede (coluna B "REDE / ASSOCIAÇÃO" das farmácias avulsas) e o mapa curado de grupos do
    # cadastro -> rede (avulsos/farmacias/final/grupos_redes.json, local). A rede herda a compra mais recente das associadas.
    assoc = C.defaultdict(set)
    def chaves_rede(txt):
        # 'FARMÁCIA DO TRABALHADOR / ULTRA POPULAR (FARMARCAS)' -> o texto todo, cada parte e a rede-mãe entre parênteses
        ks = {norm(txt)}
        for parte in re.split(r' / ', txt):
            ks.add(norm(parte))
            ks.update(norm(m) for m in re.findall(r'\(([^)]*)\)', parte))
        return {k for k in ks if k}
    for e in E:
        if e['extra'].get('so_cnpj') and e['extra'].get('rede'):
            for k_ in chaves_rede(e['extra']['rede']): assoc[k_].update(e['raizes'])
    mapa_grp = C.defaultdict(set)
    arq_mapa = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'farmacias', 'final', 'grupos_redes.json')
    if os.path.exists(arq_mapa):
        import json
        G = json.load(open(arq_mapa, encoding='utf8'))
        for cod, rede in G.get('grupos', {}).items(): mapa_grp[norm(rede)].add(cod)
        for cod in G.get('farmarcas', []): mapa_grp['FARMARCAS'].add(cod)
    por_grpcod = C.defaultdict(list)
    for k, c in cli.items():
        if c['grp']: por_grpcod[c['grp']].append(k)
    # nomes usados por mais de uma rede (clusters diferentes): vínculo fraco exige a mesma UF
    donos = C.defaultdict(set)
    for e in E:
        if not e['extra'].get('so_cnpj'):
            for x in e['nomes']: donos[x].add(e['cluster'])
    ambiguos = {x for x, cs in donos.items() if len(cs) > 1}
    cl = C.defaultdict(list)
    for n, e in enumerate(E):
        if not e['extra'].get('so_cnpj'): cl[e['cluster']].append(n)
    # marca separada do grupo (duas marcas do mesmo dono, decisão do usuário): vira uma rede à parte, sem o CNPJ e a
    # razão que divide com a outra marca. Só conta compra com o nome dela.
    # As linhas da mesma marca em outras abas (ex.: a mesma marca na aba IRTT) entram juntas, pelo rótulo.
    novo, tira, origem = C.defaultdict(list), {}, C.defaultdict(set)
    for cid, ns in cl.items():
        for n in ns:
            lab = SEPARAR.get(norm(E[n]['extra'].get('nome', '')))
            novo[('SEP', lab) if lab else (cid, None)].append(n)
            if lab: origem[lab].add(cid)
    for lab, cids in origem.items():
        resto = [n for c in cids for n in novo.get((c, None), [])]
        tira[('SEP', lab)] = ({r for n in resto for r in E[n]['raizes']}, {norm_legal(z) for n in resto for z in E[n]['razoes']})
    cl = novo
    R = {}
    t0 = dt.datetime.now()
    # 16/11 (revisão final): regra 9 em todos os grupos do IRTT (medido: 7 de 10 mudanças certas; as 3 erradas ficam em
    # GRUPOS_OK). Código de cliente que a aba IRTT repete em 2+ marcas (o distribuidor do grupo em todas as bandeiras) só
    # vale para a marca com o nome ou o CNPJ no cliente.
    cod_marcas = C.defaultdict(set)
    for e in E:
        if e['aba'].startswith('I.R.T.T'):
            for cod in e['codigos']: cod_marcas[cod].add(norm(e['nomes'][0] if e['nomes'] else e['extra'].get('nome', '')).replace(' ', ''))
    def aval(cid, ns, fams_extra=(), regra9=True):
        es = [E[n] for n in ns]
        nomes = list(dict.fromkeys(x for e in es for x in e['nomes']))
        for x in list(nomes): nomes += [y for y in ALIAS_MANUAL.get(x, []) if y not in nomes]
        erradas = set().union(*[RAIZES_ERRADAS.get(x, set()) for x in nomes]) if nomes else set()
        t_r, t_z = tira.get(cid, (set(), set()))
        raizes = sorted({r for e in es for r in e['raizes']} - erradas - t_r)    # ex.: CNPJ de outra empresa na aba IRTT
        razoes = list(dict.fromkeys(x for e in es for x in e['razoes'] if norm_legal(x) not in t_z))
        fams = sorted({e['fam'] for e in es} | {f for x in nomes for f in FAM_EXTRA.get(x, [])} | set(fams_extra))
        principais = {e['nomes'][0] for e in es if e['nomes']}
        principais |= {y for x in nomes for y in ALIAS_MANUAL.get(x, [])}
        ufs = {u.strip() for e in es for u in str(e['extra'].get('uf', '')).replace(';', ',').split(',')}
        redes_farm = {norm(e['extra']['nome']) for e in es if e['aba'] == '"FARMACIA VF"'}
        r_assoc = set().union(*[assoc.get(x, set()) for x in redes_farm]) if redes_farm else set()
        uf_sempre = set(fams) <= {'SUPER', 'DROG', 'POSTO', 'ESTAC'}   # 16/11: posto e estacionamento também são regionais
        if any(x in NACIONAIS for x in nomes):       # Shell, Ipiranga, Petrobras: o posto com a bandeira é de qualquer UF
            uf_sempre, ufs = False, set()
        v = M.vinculos(nomes, raizes, razoes, Proibidas(todas, set(raizes) | r_assoc), fam=fams, principais=principais,
                       ambiguos=ambiguos, ufs=ufs, uf_sempre=uf_sempre, manuais={y for x in nomes for y in ALIAS_MANUAL.get(x, [])})
        for r in sorted(r_assoc):
            for k in M.raiz.get(r, []): v.append(('ASSOCIADA', k, cli[k]['ult'], f'associada da rede (raiz {r})'))
        for x in redes_farm:
            for gcod in mapa_grp.get(x, ()):
                for k in por_grpcod.get(gcod, ()): v.append(('GRUPO MAPA', k, cli[k]['ult'], f"grupo {gcod} '{cli[k]['grp_desc']}' (mapa curado)"))
        def tem_relacao(k):
            c = cli[k]
            if c['raiz'] and c['raiz'] in raizes: return True
            txt = ' '.join((c['_nome'], c['_fant'], c['_grp']))
            # "HIGA S" (de Higa's) também vale sem o S do apóstrofo: "HIGA PRODUTOS"
            vs = {n for n in nomes} | {n[:-2] for n in nomes if n.endswith(' S')}
            # nome colado só com 2+ palavras (regra 9 do CLAUDE.md: "DROGAL" colado batia em "DROGA LIDER")
            return any(re.search(r'(?<![A-Z0-9])' + re.escape(n) + r'(?![A-Z0-9])', txt) or (' ' in n and n.replace(' ', '') in txt.replace(' ', ''))
                       for n in vs if len(n.replace(' ', '')) >= 4)
        for x in set(nomes) | {norm(e['extra'].get('nome', '')) for e in es}:
            for ck in CLIENTES_MANUAIS.get(x, []):
                k = tuple(ck.split('/'))
                if k in cli: v.append(('NOME', k, cli[k]['ult'], 'vínculo conferido à mão (CRM, 16/11)'))
        for e in es:
            for cod in e['codigos']:      # códigos e grupos da aba IRTT já foram validados (regra 9) quando a aba foi feita
                ks = por_cod.get(cod, [])
                # código de operador logístico com lojas de várias redes (uma loja por rede atendida): só as lojas com o
                # nome da rede. Sem isso uma rede herdava a nota de outra rede atendida pelo mesmo código.
                gs = {cli[k]['_grp'] for k in ks if cli[k]['_grp'] and cli[k]['_grp'] not in GRUPO_GENERICO}
                if len(gs) > 1: ks = [k for k in ks if tem_relacao(k)]
                # código reaproveitado por outra empresa (loja de outra raiz, sem o nome da rede): fica fora
                rz = C.Counter(cli[k]['raiz'] for k in ks if cli[k]['raiz'])
                if rz:
                    princ_r = rz.most_common(1)[0][0]
                    ks = [k for k in ks if not cli[k]['raiz'] or cli[k]['raiz'] in (princ_r, *raizes) or tem_relacao(k)]
                if len(cod_marcas.get(cod, ())) > 1: ks = [k for k in ks if tem_relacao(k)]
                for k in ks: v.append(('CODIGO', k, cli[k]['ult'], f'código {cod} (aba IRTT)'))
            for g in e['grupos']:
                gn = norm(g)
                if gn in GRUPO_GENERICO or len(por_grp.get(gn, ())) > 2000: continue    # 'CLIENTES DIVERSOS' não é grupo de rede
                # regra 9 (CLAUDE.md), para a marca separada do grupo: o grupo do cadastro só vale se tiver o nome da marca,
                # ou o cliente tiver o nome dela ou o CNPJ dela ("GRUPO TRIGO" não é compra do Gurumê)
                nome_no_grupo = any(re.search(r'(?<![A-Z0-9])' + re.escape(x) + r'(?![A-Z0-9])', gn) or (' ' in x and x.replace(' ', '') in gn.replace(' ', ''))
                                    for x in nomes if len(x.replace(' ', '')) >= 4)
                for k in por_grp.get(gn, ()) if gn else ():
                    if regra9 and not nome_no_grupo and gn not in GRUPOS_OK and not tem_relacao(k): continue
                    v.append(('GRUPO ERP', k, cli[k]['ult'], f"grupo ERP '{cli[k]['grp_desc']}'"))
            nome_cheio = norm_legal(re.sub(r'\([^)]*\)', ' ', e['extra'].get('nome', '')))
            if nome_cheio and (not e['nomes'] or len(nome_cheio.replace(' ', '')) <= 5):
                for k in exato.get(nome_cheio, ()):
                    sg = norm(cli[k]['seg'])
                    # 16/11: rede regional (supermercado, farmácia, posto, construção, estacionamento) ou nome com palavra de
                    # tipo ('POSTO CENTRAL', 'CASA DAS TINTAS'): só na mesma UF. Marca nacional curta vale em qualquer UF.
                    regional = set(fams) & {'SUPER', 'DROG', 'POSTO', 'CONSTR', 'ESTAC'} or set(nome_cheio.split()) & TIPO
                    if regional and ufs & UFS and cli[k]['uf'] and cli[k]['uf'] not in ufs: continue
                    if (nome_cheio, k[0]) in EXCLUSOES: continue
                    if not sg or any(sg in COMPAT.get(f, set()) for f in fams):
                        v.append(('EXATO', k, cli[k]['ult'], f"fantasia/grupo igual '{nome_cheio}'"))
        # nome genérico revisado à mão (16/11): a linha só vale por vínculo forte; nome, e-mail, produto e grupo do cliente
        # com o mesmo nome caem (sem isso, tirando um falso entrava o próximo com o mesmo nome)
        if any(norm(e['extra'].get('nome', '')) in SO_FORTES for e in es):
            v = [x for x in v if x[0] in ('CNPJ', 'RAZAO', 'CODIGO', 'ASSOCIADA', 'GRUPO MAPA', 'GRUPO ERP') or x[3].startswith('vínculo conferido')]
        res = M.resultado(v)
        # provas: o vínculo mais recente de cada caminho
        melhor = {}
        for cam, k, d, prova in v:
            if cam not in melhor or (d or dt.date.min) > (melhor[cam][2] or dt.date.min): melhor[cam] = (cam, k, d, prova)
        res['provas'] = [(cam, f"{k[0]}/{k[1]} {cli[k]['nome'][:40]}", d, prova, M.vendedor(k)[0], cli[k]['seg'], cli[k]['grp_desc'])
                         for cam, k, d, prova in sorted(melhor.values(), key=lambda x: x[2] or dt.date.min, reverse=True)]
        res['caminhos'] = sorted({x[0] for x in v})
        res['nomes'], res['raizes'] = nomes, raizes
        return res
    # 16/11 (decisão do usuário: a marca não herda a compra de outra marca do grupo). O cluster junta as marcas de um grupo
    # (pela linha do grupo na Alessandra, "GRUPO X (A / B / C)"). Cada marca das abas FINAL e IRTT é avaliada só com as
    # linhas dela (mesmo nome principal, mesmo CNPJ, ou nome de uma dentro da outra); a linha do grupo fica com o grupo
    # todo. FARMACIA VF não se separa: a rede/associação herda das associadas (regra de 05/11). Marcas do mesmo CNPJ
    # (Pão de Açúcar x Extra) ficam juntas, porque a nota não diz a marca.
    MARCA = lambda a: (a.endswith('FINAL') and a != 'OUTROS FINAL') or a.startswith('I.R.T.T')
    marca = lambda e: (norm(e['nomes'][0]) if e['nomes'] else norm(e['extra'].get('nome', ''))).replace(' ', '')
    n_sep = 0
    for i, (cid, ns) in enumerate(cl.items()):
        bs = [n for n in ns if MARCA(E[n]['aba'])]
        so_drog = {E[n]['fam'] for n in ns} <= {'DROG'}
        if cid[1] is None and len(bs) > 1 and not so_drog:
            pai = {n: n for n in bs}
            def raiz_(x):
                while pai[x] != x: pai[x] = pai[pai[x]]; x = pai[x]
                return x
            def une(a_, b_): pai[raiz_(a_)] = raiz_(b_)
            chave = C.defaultdict(list)
            for n in bs:
                chave[('M', marca(E[n]))].append(n)
                for r in E[n]['raizes']: chave[('R', r)].append(n)
            for (tipo_, _), ms in chave.items():
                if tipo_ != 'M': continue      # 16/11: só pelo nome; marcas do mesmo CNPJ ficam separadas, e as duas levam a
                for m in ms[1:]: une(m, ms[0])  # compra pelo CNPJ (a da empresa), mas o produto com o nome de uma é só dela
            for a_ in bs:
                for b_ in bs:     # "Ri Happy (+ PBKids)" inclui PB Kids; "DOMINOS" e "DOMINOS PIZZA"
                    ma = marca(E[a_])
                    if ma and any(x.replace(' ', '').startswith(ma) for x in E[b_]['nomes'] + [marca(E[b_])]): une(a_, b_)
            grupos = C.defaultdict(list)
            for n in bs: grupos[raiz_(n)].append(n)
        else:
            grupos = {}
        if len(grupos) <= 1:
            res = aval(cid, ns)
            for n in ns: R[n] = res
        else:
            n_sep += 1
            nm_g = {g: {x.replace(' ', '') for n in ms for x in E[n]['nomes']} | {marca(E[n]) for n in ms} for g, ms in grupos.items()}
            fams_cl = {E[n]['fam'] for n in ns}       # o tipo de rede vem do cluster todo (a linha da IRTT só diz "Drogarias")
            rz_g = {g: {r for n in ms for r in E[n]['raizes']} for g, ms in grupos.items()}
            junto, resto = C.defaultdict(list), []
            for n in ns:
                if n in bs: continue
                e = E[n]
                if e['aba'] == '"FARMACIA VF"':
                    ov = [g for g in grupos if set(e['raizes']) & rz_g[g]]
                else:
                    ov = [g for g in grupos if ({x.replace(' ', '') for x in e['nomes']} | {marca(e)}) & nm_g[g]]
                if len(ov) == 1: junto[ov[0]].append(n)
                else: resto.append(n)          # linha do grupo (várias marcas): fica com o grupo todo
            for g, ms in grupos.items():
                res = aval(cid, ms + junto[g], fams_cl, regra9=True)
                for n in ms + junto[g]: R[n] = res
            if resto:
                res = aval(cid, ns)
                for n in resto: R[n] = res
        if i % 300 == 0: print(i, len(cl), (dt.datetime.now() - t0).seconds, 's', flush=True)
    print('clusters com marcas separadas:', n_sep)
    # farmácias avulsas: só a raiz do CNPJ
    for n, e in enumerate(E):
        if not e['extra'].get('so_cnpj'): continue
        v = []
        for r in e['raizes']:
            for k in M.raiz.get(r, []): v.append(('CNPJ', k, cli[k]['ult'], f'raiz {r}'))
        res = M.resultado(v); res['provas'] = []; res['caminhos'] = sorted({x[0] for x in v})
        R[n] = res
    pickle.dump(R, open(os.path.join(PASTA, 'resultados.pkl'), 'wb'))
    print('ok', len(R), (dt.datetime.now() - t0).seconds, 's')


if __name__ == '__main__':
    main()
