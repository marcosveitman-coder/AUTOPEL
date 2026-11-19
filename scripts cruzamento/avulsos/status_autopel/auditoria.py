"""Auditoria do STATUS AUTOPEL (19/11/2026, depois do caso "Koni Store": NÃO É CLIENTE numa aba e ATIVO na outra).
Duas buscas, para revisão à mão (não muda status sozinha):
 A) CONFLITO: a mesma rede com status diferente em linhas diferentes. Mesma raiz de CNPJ, ou mesmo nome (núcleo sem
    espaços), em clusters diferentes.
 B) PISTA: linha que não ficou ATIVO, mas tem um sinal fraco que as travas do motor não aceitaram:
    - produto personalizado: o 1º nome da rede (4+ letras) logo depois da medida ("80X100M KONI"), em nota ou pedido aberto;
    - cliente com o nome da rede como palavras inteiras na razão, na fantasia ou no grupo;
    - domínio de e-mail igual ao nome colado (5+ letras).
    Só entra pista com compra (ou pedido) MAIS RECENTE que a do status atual.
Entrada: <pasta>/clientes.pkl, entidades.pkl, resultados.pkl. Saída: <pasta>/auditoria.tsv (local: tem clientes).
Uso: python auditoria.py <pasta> AAAA-MM-DD"""
import sys, os, re, pickle, datetime as dt, collections as C
from motor import norm, nucleo, depois_da_medida, GENERICO, TIPO, PROVEDOR

P, HOJE = sys.argv[1], dt.date.fromisoformat(sys.argv[2])
B = pickle.load(open(os.path.join(P, 'clientes.pkl'), 'rb')); cli = B['cli']
E = pickle.load(open(os.path.join(P, 'entidades.pkl'), 'rb'))
R = pickle.load(open(os.path.join(P, 'resultados.pkl'), 'rb'))
GRAVA = lambda e: not e['extra'].get('apoio') and e['aba'] != 'SUPERMERCADOS (reuniao)' and not e['extra'].get('so_cnpj')
ORD = {'ATIVO': 2, 'INATIVO': 1, 'NÃO É CLIENTE': 0}
out = []

# ---- A) conflitos
por_raiz, por_nome = C.defaultdict(set), C.defaultdict(set)
for n, e in enumerate(E):
    if not GRAVA(e): continue
    for r in e['raizes']: por_raiz[r].add(n)
    if e['nomes']: por_nome[e['nomes'][0].replace(' ', '')].add(n)
vistos = set()
for tipo, idx in (('RAIZ', por_raiz), ('NOME', por_nome)):
    for chave, ns in idx.items():
        st = {R[n]['status'] for n in ns}
        cls = {E[n]['cluster'] for n in ns}
        if len(st) > 1 and len(cls) > 1:
            k = frozenset(ns)
            if k in vistos: continue
            vistos.add(k)
            for n in sorted(ns, key=lambda n: -ORD[R[n]['status']]):
                e, r = E[n], R[n]
                out.append(['CONFLITO ' + tipo, chave, e['aba'], e['linha'], e['extra'].get('nome', ''), r['status'], r['data'] or '',
                            r['via'], r.get('cliente', ''), r.get('prova', '')])

# ---- B) pistas
prod_idx = C.defaultdict(list)     # 1ª palavra depois da medida -> (texto, cliente, data, origem)
for k, c in cli.items():
    for des, (d, _) in c['prods'].items():
        resto = depois_da_medida(norm(des))
        if resto: prod_idx[resto.split()[0]].append((resto, k, d, 'nota'))
    for des, d in c.get('ped', {}).items():
        resto = depois_da_medida(norm(des))
        if resto: prod_idx[resto.split()[0]].append((resto, k, d, 'pedido aberto'))
tok = C.defaultdict(set)
for k, c in cli.items():
    for t in (c['_nome'] if '_nome' in c else norm(c['nome']), norm(c['fant']), norm(c['grp_desc'])):
        for w in set(t.split()): tok[w].add(k)
dom = C.defaultdict(set)
for k, c in cli.items():
    for d in c['dom']: dom[d.split('.')[0].replace('-', '').upper()].add(k)
feito = set()
for n, e in enumerate(E):
    if not GRAVA(e) or not e['nomes']: continue
    r = R[n]
    if r['status'] == 'ATIVO': continue
    chave = (e['cluster'], tuple(e['nomes']))
    if chave in feito: continue
    feito.add(chave)
    base = r['data'] or dt.date(1900, 1, 1)
    pistas = {}
    for nome in e['nomes']:
        w = [x for x in nome.split() if x not in TIPO]
        if not w: continue
        w0 = w[0]
        # nome inteiro (colado ou separado) logo depois da medida; ou o nome sem a palavra genérica do fim
        # ("KONI STORE" -> "KONI"), com 4+ letras
        alvo = {''.join(w)}
        ww = list(w)
        while len(ww) > 1 and (ww[-1] in GENERICO or ww[-1] in TIPO): ww.pop()
        if len(''.join(ww)) >= 4: alvo.add(''.join(ww))
        if len(w0) >= 3:
            for resto, k, d, org in prod_idx.get(w0, ()):
                rc = resto.replace(' ', '')
                if not any(rc == a_ or (rc.startswith(a_) and re.match(r'^(C?\d+|[A-Z]{2})?$', rc[len(a_):][:3]) is not None) or
                           resto.startswith(' '.join(ww) + ' ') for a_ in alvo): continue
                if d and d > base:
                    p = pistas.get(('PRODUTO', k))
                    if not p or d > p[0]: pistas[('PRODUTO', k)] = (d, f"{org}: '{resto[:40]}'")
        fr = re.compile(r'(?<![A-Z0-9])' + r'\s?'.join(map(re.escape, nome.split())) + r'(?![A-Z0-9])')
        raro = min(nome.split(), key=lambda x: len(tok.get(x, ())))
        if len(nome.replace(' ', '')) >= 5 and not all(x in GENERICO or x in TIPO for x in nome.split()) and len(tok.get(raro, ())) < 400:
            for k in tok.get(raro, ()):
                c = cli[k]
                txt = ' | '.join((norm(c['nome']), norm(c['fant']), norm(c['grp_desc'])))
                if fr.search(txt) and c['ult'] and c['ult'] > base:
                    pistas.setdefault(('NOME', k), (c['ult'], f"cliente '{c['nome'][:35]}' / '{c['fant'][:25]}' / grupo '{c['grp_desc'][:25]}'"))
        cp = nome.replace(' ', '')
        if len(cp) >= 5 and cp not in PROVEDOR:
            for k in dom.get(cp, ()):
                c = cli[k]
                if c['ult'] and c['ult'] > base: pistas.setdefault(('EMAIL', k), (c['ult'], f"e-mail @{cp.lower()}"))
    for (cam, k), (d, prova) in sorted(pistas.items(), key=lambda x: x[1][0], reverse=True)[:6]:
        c = cli[k]
        out.append(['PISTA ' + cam, e['nomes'][0], e['aba'], e['linha'], e['extra'].get('nome', ''), r['status'], r['data'] or '',
                    f"{d}", f"{k[0]}/{k[1]} {c['nome'][:40]} | UF {c['uf']} | {c['seg']} | grupo {c['grp_desc']}", prova])
with open(os.path.join(P, 'auditoria.tsv'), 'w', encoding='utf-8-sig') as f:
    f.write('\t'.join(['TIPO', 'CHAVE', 'ABA', 'LINHA', 'NOME', 'STATUS NOVO', 'DATA DO STATUS', 'VIA / DATA DA PISTA', 'CLIENTE', 'PROVA']) + '\n')
    for x in out: f.write('\t'.join(str(v) for v in x) + '\n')
print(C.Counter(x[0] for x in out))
print('linhas com pista:', len({(x[2], x[3]) for x in out if x[0].startswith('PISTA')}))
