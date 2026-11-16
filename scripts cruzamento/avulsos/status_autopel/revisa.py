"""Lista para revisão manual (16/11/2026): redes cujo status final depende de vínculo fraco (nome, e-mail, produto),
ou seja, os só-fortes (CNPJ, código, razão, grupo) dariam outro status. Mostra as provas de cada caminho.
Uso: python revisa.py "<pasta de trabalho>" [nomes para conferir, separados por ;] """
import sys, os, pickle, datetime as dt

PASTA = sys.argv[1]
E = pickle.load(open(os.path.join(PASTA, 'entidades.pkl'), 'rb'))
R = pickle.load(open(os.path.join(PASTA, 'resultados.pkl'), 'rb'))
FORTE = {'CNPJ', 'CODIGO', 'RAZAO', 'ASSOCIADA', 'GRUPO MAPA', 'GRUPO ERP', 'GRUPO', 'EXATO'}
lim = dt.date(2026, 11, 16) - dt.timedelta(days=360)


def mostra(e, r, n):
    print(f"#{n} [{e['fam']}] {e['aba'][:20]} L{e['linha']} {e['extra']['nome'][:45]} | nomes={r['nomes'][:4]} | antes={e['extra'].get('status_antes')} agora={r['status']} {r['data']}")
    for p in r['provas'][:6]:
        print(f"     {p[0]:9} {p[2]} {p[1][:42]:42} | {p[3][:58]:58} | seg={p[5][:14]} grp={p[6][:22]}")


if len(sys.argv) > 2:
    alvo = [x.strip().upper() for x in sys.argv[2].split(';')]
    vis = set()
    for i, e in enumerate(E):
        if e['extra'].get('so_cnpj'): continue
        if any(a in e['extra']['nome'].upper() for a in alvo) and id(R[i]) not in vis:
            vis.add(id(R[i])); mostra(e, R[i], len(vis))
    sys.exit()
vis = set(); n = 0
for i, e in enumerate(E):
    if e['extra'].get('so_cnpj'): continue
    r = R[i]
    if r['status'] == 'NÃO É CLIENTE' or id(r) in vis: continue
    vis.add(id(r))
    forte = [p for p in r['provas'] if p[0] in FORTE]
    mf = max([p[2] for p in forte if p[2]], default=None)
    st = 'ATIVO' if mf and mf > lim else ('INATIVO' if forte else 'NÃO É CLIENTE')
    if st != r['status']:
        n += 1; mostra(e, r, n)
print('total a revisar:', n)
