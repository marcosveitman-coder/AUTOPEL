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
from motor import ALIAS_MANUAL


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
    R = {}
    t0 = dt.datetime.now()
    for i, (cid, ns) in enumerate(cl.items()):
        es = [E[n] for n in ns]
        nomes = list(dict.fromkeys(x for e in es for x in e['nomes']))
        for x in list(nomes): nomes += [y for y in ALIAS_MANUAL.get(x, []) if y not in nomes]
        raizes = sorted({r for e in es for r in e['raizes']})
        razoes = list(dict.fromkeys(x for e in es for x in e['razoes']))
        fams = sorted({e['fam'] for e in es})
        principais = {e['nomes'][0] for e in es if e['nomes']}
        principais |= {y for x in nomes for y in ALIAS_MANUAL.get(x, [])}
        ufs = {u.strip() for e in es for u in str(e['extra'].get('uf', '')).replace(';', ',').split(',')}
        redes_farm = {norm(e['extra']['nome']) for e in es if e['aba'] == '"FARMACIA VF"'}
        r_assoc = set().union(*[assoc.get(x, set()) for x in redes_farm]) if redes_farm else set()
        uf_sempre = set(fams) <= {'SUPER', 'DROG'}
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
            return any(n in txt or n.replace(' ', '') in txt.replace(' ', '') for n in nomes if len(n.replace(' ', '')) >= 4)
        for e in es:
            for cod in e['codigos']:      # códigos e grupos da aba IRTT já foram validados (regra 9) quando a aba foi feita
                for k in por_cod.get(cod, []): v.append(('CODIGO', k, cli[k]['ult'], f'código {cod} (aba IRTT)'))
            for g in e['grupos']:
                gn = norm(g)
                if gn in GRUPO_GENERICO or len(por_grp.get(gn, ())) > 2000: continue    # 'CLIENTES DIVERSOS' não é grupo de rede
                for k in por_grp.get(gn, ()) if gn else ():
                    v.append(('GRUPO ERP', k, cli[k]['ult'], f"grupo ERP '{cli[k]['grp_desc']}'"))
            nome_cheio = norm_legal(re.sub(r'\([^)]*\)', ' ', e['extra'].get('nome', '')))
            if nome_cheio and (not e['nomes'] or len(nome_cheio.replace(' ', '')) <= 5):
                for k in exato.get(nome_cheio, ()):
                    sg = norm(cli[k]['seg'])
                    if not sg or any(sg in COMPAT.get(f, set()) for f in fams):
                        v.append(('EXATO', k, cli[k]['ult'], f"fantasia/grupo igual '{nome_cheio}'"))
        res = M.resultado(v)
        # provas: o vínculo mais recente de cada caminho
        melhor = {}
        for cam, k, d, prova in v:
            if cam not in melhor or (d or dt.date.min) > (melhor[cam][2] or dt.date.min): melhor[cam] = (cam, k, d, prova)
        res['provas'] = [(cam, f"{k[0]}/{k[1]} {cli[k]['nome'][:40]}", d, prova, M.vendedor(k)[0], cli[k]['seg'], cli[k]['grp_desc'])
                         for cam, k, d, prova in sorted(melhor.values(), key=lambda x: x[2] or dt.date.min, reverse=True)]
        res['caminhos'] = sorted({x[0] for x in v})
        res['nomes'], res['raizes'] = nomes, raizes
        for n in ns: R[n] = res
        if i % 300 == 0: print(i, len(cl), (dt.datetime.now() - t0).seconds, 's', flush=True)
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
