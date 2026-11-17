"""CNPJ matriz das marcas da Alessandra Dayan (17/11/2026), passo 3: escolhe o CNPJ e acha a matriz.
- candidato aceito: CNPJ CONFIRMADO / PROVÁVEL / INAPTA da aba FINAL, ou o CNPJ das abas de apoio (FARMACIA VF,
  SUPERMERCADOS (reuniao), IRTT); a matriz é o estabelecimento com indicador 1 da raiz na Receita;
- decisões à mão (dados/dec_cnpj_alessandra.json, local): {"<linha>": "<CNPJ 14 dígitos>" | "" | {"cnpj", "nota"}} trocam ou
  tiram o CNPJ (CNPJ dado à mão e conferido na BrasilAPI vale mesmo sem a matriz na busca);
- CNPJ PROVÁVEL das abas FINAL (e decisão com nota) leva um comentário na célula: "CNPJ provável (conferir)";
- marcas sem CNPJ: lista as empresas achadas pelos termos (razão social e lojas ativas com o nome), para decidir à mão.
Saídas na pasta: revisar.txt (conferência), cnpj.tsv (para o top50_cnpj_caixas.ps1: aba, linha, EMPRESA, CNPJ, nota). Fora do git.
Uso: python cnpj_alessandra_decide.py "<pasta de trabalho>" "<MAPA.xlsm>" """
import os, sys, json, pickle, collections as C
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
DADOS = os.environ.get('AUTOPEL_MAPA_FINAL_DADOS') or os.path.join(HERE, 'dados')
sys.path.insert(0, HERE)
from cnpj_alessandra import norm, dv_ok, s
ACEITOS = {'CNPJ CONFIRMADO NA RECEITA', 'CNPJ PROVÁVEL (CONFERIR)', 'CNPJ INAPTA NA RECEITA', 'apoio'}
SIT = {'01': 'NULA', '02': 'ATIVA', '03': 'SUSPENSA', '04': 'INAPTA', '08': 'BAIXADA'}
fmt = lambda c: f'{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}'


def main(pasta, mapa):
    cand = json.load(open(os.path.join(pasta, 'candidatos.json'), encoding='utf8'))
    R = pickle.load(open(os.path.join(pasta, 'r_alessandra.pkl'), 'rb'))
    arq_dec = os.path.join(DADOS, 'dec_cnpj_alessandra.json')
    DEC = json.load(open(arq_dec, encoding='utf8')) if os.path.exists(arq_dec) else {}
    emp = R['emp']
    est = C.defaultdict(list)
    for x in R['est']: est[x[0]].append(x)
    def matriz(raiz):
        m = [x for x in est.get(raiz, []) if x[3] == '1']
        return (m[0][0] + m[0][1] + m[0][2], SIT.get(m[0][5], m[0][5]), m[0][4], m[0][10]) if m else None
    def ativas(raiz): return sum(1 for x in est.get(raiz, []) if x[5] == '02')
    termos = C.defaultdict(set)
    for q in json.load(open(os.path.join(pasta, 'q_alessandra.json'), encoding='utf8')):
        if q['id'].startswith('LIN|'):
            for t in q['terms']: termos[t].add(int(q['id'][4:]))
    achados = C.defaultdict(C.Counter)          # linha -> raiz -> nº de estabelecimentos ativos com o termo no fantasia
    for r, t in R['hits_raz']:
        for lin in termos.get(t, ()): achados[lin][r] += 0
    for i, t in R['hits_fant']:
        for lin in termos.get(t, ()): achados[lin][R['est'][i][0]] += 1
    out, rev, cont = {}, [], C.Counter()
    for o in cand:
        lin = o['lin']
        c0 = next((c for c in o['cands'] if c['conf'] in ACEITOS and len(c['cnpj']) == 14), None)
        cn, via = (c0['cnpj'], c0['conf'] + ' – ' + c0['fonte']) if c0 else ('', '')
        nota = 'CNPJ provável (conferir)' if c0 and c0['conf'] == 'CNPJ PROVÁVEL (CONFERIR)' else ''
        if str(lin) in DEC:
            d = DEC[str(lin)]
            cn, nota = (d['cnpj'], d.get('nota', '')) if isinstance(d, dict) else (d, '')
            via = 'decisão manual'
        if cn:
            assert dv_ok(cn), (lin, cn)
            m = matriz(cn[:8])
            if not m and via == 'decisão manual': m = (cn, 'ATIVA?', '', '')   # conferido à mão na BrasilAPI
            if not m: rev.append(f'!! {lin} {o["nome"]}: raiz {cn[:8]} sem matriz na Receita'); cont['sem matriz'] += 1; continue
            out[lin] = (m[0], nota)
            e = emp.get(cn[:8], ('',))
            outros = {c['cnpj'][:8] for c in o['cands'] if len(c['cnpj']) == 14} - {cn[:8]}
            rev.append(f'OK {lin:3} {o["seg"][:12]:12} | {o["nome"][:45]:45} | {fmt(m[0])} {m[1]:7} {m[3]} | {e[0][:50]} | fantasia "{m[2][:25]}" | '
                       f'{ativas(cn[:8])} ativos | {via[:60]}' + (f' | OUTRAS RAÍZES {sorted(outros)}' if outros else ''))
            cont['com CNPJ'] += 1
            if m[0] != cn: cont['trocado pela matriz'] += 1
        else:
            cont['sem CNPJ'] += 1
            rev.append(f'-- {lin:3} {o["seg"][:12]:12} | {o["nome"][:45]:45} | {o["cands"][0]["conf"] if o["cands"] else "SEM CANDIDATO"}')
            for r, n in achados.get(lin, C.Counter()).most_common(8):
                m = matriz(r); e = emp.get(r, ('', '', '', ''))
                rev.append(f'      {r} {fmt(m[0]) if m else "sem matriz":18} {m[1] if m else "":7} {m[3] if m else "":2} | {e[0][:55]:55} | '
                           f'nat {e[1]} | {n} lojas c/ nome | {ativas(r)} ativos')
    open(os.path.join(pasta, 'revisar.txt'), 'w', encoding='utf8').write('\n'.join(rev))
    # tsv para o Excel: Alessandra Dayan VF e as 11 TOP 50 (mesmas empresas, pelo nome)
    por_nome = C.defaultdict(set)
    for o in cand:
        if o['lin'] in out: por_nome[norm(o['nome'])].add(out[o['lin']])
    wb = openpyxl.load_workbook(mapa, read_only=True, data_only=True)
    tsv = []
    for aba in [a for a in wb.sheetnames if a.endswith('(TOP 50)') or a == 'Alessandra Dayan VF']:
        for n, (v,) in enumerate(wb[aba].iter_rows(min_col=1, max_col=1, values_only=True), 1):
            if aba == 'Alessandra Dayan VF':
                c = out.get(n)
            else:
                cs = por_nome.get(norm(v)) if v else None
                c = next(iter(cs)) if cs and len(cs) == 1 else None
                if cs and len(cs) > 1: print('ATENÇÃO: nome com 2 CNPJs', aba, n, v)
            if c and v and s(v) != 'EMPRESA': tsv.append((aba, n, s(v), fmt(c[0]), c[1]))
    wb.close()
    with open(os.path.join(pasta, 'cnpj.tsv'), 'w', encoding='utf-8', newline='') as f:
        for x in tsv: f.write('\t'.join(map(str, x)) + '\r\n')
    print(dict(cont), '| linhas no tsv:', len(tsv), C.Counter(x[0] for x in tsv))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
