"""Conferência no sentido contrário (16/11/2026): de quem COMPROU para as listas do MAPA.
Soma as vendas (valor > 0) dos últimos 360 dias por raiz de CNPJ nos arquivos brutos de faturamento e, para cada um dos
maiores compradores, procura a rede das listas que o comprador menciona (raiz do CNPJ da linha, nome principal da rede no
nome / fantasia / grupo do cliente). Se a rede achada NÃO está ATIVO, o caso vai para revisão manual.
Saída: <pasta>/confere_compradores.txt (tem nomes, fora do git).
Uso: python confere_compradores.py "<pasta de trabalho>" "<pasta 12.11.2026>" AAAA-MM-DD [quantos, padrão 1500]"""
import sys, os, glob, pickle, re, datetime as dt, collections as C
from concurrent.futures import ProcessPoolExecutor
import openpyxl
from confere_bruto import data, norm, s, D


def soma(arg):
    arq, lim = arg
    wb = openpyxl.load_workbook(arq, read_only=True)
    val, info = C.Counter(), {}
    for ws in wb.worksheets:
        it = ws.iter_rows(values_only=True)
        cab = [s(x) for x in next(it)]
        while 'EMISSAO' not in cab: cab = [s(x) for x in next(it)]
        iE, iC, iL, iN, iV, iG, iR = (cab.index(c) for c in ('EMISSAO', 'COD_CLI', 'COD_LOJA', 'CNPJ', 'VLR_TOTAL', 'GRP_CLI', 'CLIENTE'))
        for r in it:
            if not r or len(r) <= iV: continue
            d = data(r[iE])
            if not d or d <= lim: continue
            try: v = float(r[iV] or 0)
            except (TypeError, ValueError): continue
            if v <= 0: continue
            rz = D(r[iN])[:8] if len(D(r[iN])) == 14 else 'COD' + s(r[iC])
            val[rz] += v
            info.setdefault(rz, (s(r[iC]), s(r[iL]), s(r[iR])))
    wb.close()
    return val, info


def main():
    PASTA, FONTES, HOJE = sys.argv[1], sys.argv[2], dt.date.fromisoformat(sys.argv[3])
    N = int(sys.argv[4]) if len(sys.argv) > 4 else 1500
    lim = HOJE - dt.timedelta(days=360)
    arqs = [f for f in sorted(glob.glob(os.path.join(FONTES, 'faturamento*.xlsx'))) if re.search(r'202[56]', f)]
    cache = os.path.join(PASTA, f'compradores_{HOJE}.pkl')
    if os.path.exists(cache):
        val, info = pickle.load(open(cache, 'rb'))
    else:
        val, info = C.Counter(), {}
        with ProcessPoolExecutor(3) as ex:
            for v, i in ex.map(soma, [(f, lim) for f in arqs]):
                val.update(v)
                for k, x in i.items(): info.setdefault(k, x)
        pickle.dump((val, info), open(cache, 'wb'))
    E = pickle.load(open(os.path.join(PASTA, 'entidades.pkl'), 'rb'))
    R = pickle.load(open(os.path.join(PASTA, 'resultados.pkl'), 'rb'))
    B = pickle.load(open(os.path.join(PASTA, 'clientes.pkl'), 'rb'))
    cli = B['cli']
    por_rz, por_cod = C.defaultdict(list), C.defaultdict(list)
    for k, c in cli.items():
        if c['raiz']: por_rz[c['raiz']].append(k)
        por_cod[k[0]].append(k)
    # linhas por raiz e por nome principal (só abas gravadas, sem farmácia avulsa)
    lin_rz, lin_nome = C.defaultdict(set), C.defaultdict(set)
    for n, e in enumerate(E):
        a = e['aba']
        if a.startswith('SUPERMERCADOS (re') or a == 'OUTROS FINAL' or e['extra'].get('so_cnpj'): continue
        for r in e['raizes']: lin_rz[r].add(n)
        if e['nomes'] and len(e['nomes'][0].replace(' ', '')) >= 5: lin_nome[e['nomes'][0]].add(n)
    pal = re.compile(r'(?<![A-Z0-9])(' + '|'.join(sorted((re.escape(x) for x in lin_nome), key=len, reverse=True)) + r')(?![A-Z0-9])')
    out, rev = [], 0
    for rz, v in val.most_common(N):
        cod, loja, nome = info[rz]
        ks = por_rz.get(rz, []) or por_cod.get(cod, [])[:20]
        # comprador que já é o CNPJ de uma linha: vale a linha dele. Senão, o nome da rede na razão, no grupo ou no
        # começo de uma fantasia (cidade e bairro na fantasia de filial não contam)
        achou = set(lin_rz.get(rz, set()))
        razao = C.Counter(norm(cli[k]['nome']) for k in ks).most_common(1)[0][0] if ks else norm(nome)
        txt = razao + ' | ' + ' | '.join(sorted({norm(cli[k]['grp_desc']) for k in ks} - {'CLIENTES DIVERSOS', ''}))
        if not achou:
            for m in pal.finditer(txt): achou |= lin_nome[m.group(1)]
            for k in ks:
                f = re.sub(r'^(LOJA|LOJAS|REDE|GRUPO|SUPERMERCADO|SUPERMERCADOS|DROGARIA|DROGARIAS|FARMACIA|FARMACIAS|POSTO|AUTO POSTO) ', '', norm(cli[k]['fant']))
                m = pal.match(f)
                if m: achou |= lin_nome[m.group(1)]
        ruins = sorted({(E[n]['aba'][:16], s(E[n]['extra'].get('nome'))[:36], R[n]['status']) for n in achou if R[n]['status'] != 'ATIVO'})
        if ruins:
            rev += 1
            out.append(f"R$ {v:,.0f} | raiz {rz} | {nome[:40]} | {txt[:160]}\n" + ''.join(f"      -> {a} | {nm} | {st}\n" for a, nm, st in ruins))
    open(os.path.join(PASTA, 'confere_compradores.txt'), 'w', encoding='utf8').write(''.join(out))
    print(f'compradores conferidos: {min(N, len(val))} de {len(val)} | com rede das listas que não está ATIVO: {rev}')


if __name__ == '__main__':
    main()
