"""Conferência independente do STATUS AUTOPEL contra os ARQUIVOS BRUTOS (16/11/2026), sem usar os pickles do status:
lê de novo os 10 "faturamento*.xlsx", o "ultimo fat.xlsx" e o "CADASTRO DE CLIENTES.RLT.xlsx" e confere, linha a linha:
  1. data mostrada = última nota do cliente mostrado (bases + ultimo fat); no caminho por produto, a última nota daquele
     produto para aquele cliente;
  2. status = regra dos 360 dias aplicada à data (ATIVO < 360 dias; INATIVO = mais antiga ou cadastro sem nota);
  3. a data mostrada é nota de venda (valor > 0), e não só devolução/valor zero;
  4. linha com CNPJ: nenhuma compra pela raiz do CNPJ é mais nova que a data mostrada (compra perdida), e raiz que está no
     cadastro nunca fica NÃO É CLIENTE (fora as raízes erradas e as marcas separadas, decisões registradas);
  5. farmácia avulsa da FARMACIA VF: status e data refeitos só pela raiz.
Saída: resumo na tela e os casos em <pasta>/confere_bruto.txt (tem nomes, fora do git).
Uso: python confere_bruto.py "<pasta de trabalho>" "<pasta 12.11.2026>" AAAA-MM-DD"""
import sys, os, glob, pickle, re, datetime as dt, collections as C
from concurrent.futures import ProcessPoolExecutor
import openpyxl

s = lambda x: str(x).strip() if x is not None else ''
D = lambda x: re.sub(r'\D', '', s(x))


def data(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    t = s(v)
    for f, n in (('%Y-%m-%d', 10), ('%d/%m/%Y', 10)):
        try: return dt.datetime.strptime(t[:n], f).date()
        except ValueError: pass
    return None


def norm(x):
    import unicodedata
    x = unicodedata.normalize('NFKD', s(x)).encode('ascii', 'ignore').decode().upper()
    x = re.sub(r"['’`´]", '', x).replace('&', ' E ')
    return re.sub(r'\s+', ' ', re.sub(r'[^A-Z0-9 ]', ' ', x)).strip()


def le_fat(arg):
    arq, alvo = arg
    wb = openpyxl.load_workbook(arq, read_only=True)
    cl, cl_pos, rz, prod, neg, lin = {}, {}, {}, {}, 0, 0
    for ws in wb.worksheets:
        it = ws.iter_rows(values_only=True)
        cab = [s(x) for x in next(it)]
        while 'EMISSAO' not in cab: cab = [s(x) for x in next(it)]
        iE, iC, iL, iN, iD, iV = (cab.index(c) for c in ('EMISSAO', 'COD_CLI', 'COD_LOJA', 'CNPJ', 'DES_PROD', 'VLR_TOTAL'))
        for r in it:
            if not r or len(r) <= iV: continue
            d = data(r[iE])
            if d is None: continue
            lin += 1
            k = (s(r[iC]), s(r[iL]))
            try: v = float(r[iV] or 0)
            except (TypeError, ValueError): v = 0.0
            if v <= 0: neg += 1
            if k not in cl or d > cl[k]: cl[k] = d
            if v > 0:
                if k not in cl_pos or d > cl_pos[k]: cl_pos[k] = d
                ra = D(r[iN])[:8]
                if len(D(r[iN])) == 14 and (ra not in rz or d > rz[ra]): rz[ra] = d
                if k in alvo:
                    pk = (k, norm(r[iD]))
                    if pk not in prod or d > prod[pk]: prod[pk] = d
    wb.close()
    return os.path.basename(arq), lin, neg, cl, cl_pos, rz, prod


def junta(dst, src):
    for k, d in src.items():
        if k not in dst or d > dst[k]: dst[k] = d


def main():
    PASTA, FONTES, HOJE = sys.argv[1], sys.argv[2], dt.date.fromisoformat(sys.argv[3])
    LIM = HOJE - dt.timedelta(days=360)
    E = pickle.load(open(os.path.join(PASTA, 'entidades.pkl'), 'rb'))
    R = pickle.load(open(os.path.join(PASTA, 'resultados.pkl'), 'rb'))
    alvo = set()
    for n, e in enumerate(E):
        r = R[n]
        if r.get('cliente') and r.get('via') == 'PRODUTO': alvo.add(tuple(r['cliente'].split(' ')[0].split('/')))
    arqs = sorted(glob.glob(os.path.join(FONTES, 'faturamento*.xlsx')))
    cl, cl_pos, rz, prod = {}, {}, {}, {}
    with ProcessPoolExecutor(4) as ex:
        for nome, lin, neg, a, b, c, p in ex.map(le_fat, [(f, alvo) for f in arqs]):
            print(f'{nome}: {lin} linhas, {neg} com valor <= 0', flush=True)
            junta(cl, a); junta(cl_pos, b); junta(rz, c); junta(prod, p)
    # ultimo fat (DT_FAT) e cadastro (CNPJ de cada cliente+loja)
    uf, rz_cad, cnpj_k = {}, C.defaultdict(set), {}
    ws = openpyxl.load_workbook(os.path.join(FONTES, 'ultimo fat.xlsx'), read_only=True).worksheets[0]
    it = ws.iter_rows(values_only=True); cab = [s(x) for x in next(it)]
    iC, iL, iN, iD = cab.index('COD_CLIENT'), cab.index('LOJA'), cab.index('CNPJ'), cab.index('DT_FAT')
    for r in it:
        if not r or not s(r[iC]): continue
        k = (s(r[iC]), s(r[iL])); d = data(r[iD])
        if d and (k not in uf or d > uf[k]): uf[k] = d
        if len(D(r[iN])) == 14: cnpj_k.setdefault(k, D(r[iN]))
    ws = openpyxl.load_workbook(os.path.join(FONTES, 'CADASTRO DE CLIENTES.RLT.xlsx'), read_only=True).worksheets[0]
    it = ws.iter_rows(values_only=True); next(it); cab = [s(x) for x in next(it)]
    iC, iL, iN = cab.index('Codigo'), cab.index('Loja'), cab.index('CNPJ/CPF')
    for r in it:
        if r and s(r[iC]) and len(D(r[iN])) == 14: cnpj_k[(s(r[iC]), s(r[iL]))] = D(r[iN])
    for k, c in cnpj_k.items(): rz_cad[c[:8]].add(k)
    rz_uf = {}
    for k, d in uf.items():
        c = cnpj_k.get(k)
        if c and (c[:8] not in rz_uf or d > rz_uf[c[:8]]): rz_uf[c[:8]] = d
    print('ultimo fat', len(uf), '| cadastro com CNPJ', len(cnpj_k), '| raízes', len(rz_cad))
    ult = lambda k: max([d for d in (cl_pos.get(k), uf.get(k)) if d], default=None)     # só nota de venda (valor > 0)
    ult_rz = lambda r: max([d for d in (rz.get(r), rz_uf.get(r)) if d], default=None)
    st = lambda d: 'ATIVO' if d and d > LIM else 'INATIVO'
    from motor import RAIZES_ERRADAS, SEPARAR
    erradas = set().union(*RAIZES_ERRADAS.values()) if RAIZES_ERRADAS else set()
    P = C.defaultdict(list); vistos = set(); n_linhas = 0
    for n, e in enumerate(E):
        a = e['aba']
        if a.startswith('SUPERMERCADOS (re') or a == 'OUTROS FINAL': continue
        r = R[n]; n_linhas += 1
        rot = f"{a[:16]} L{e['linha']} {s(e['extra'].get('nome'))[:40]}"
        if e['extra'].get('so_cnpj'):          # farmácia avulsa: só a raiz
            raizes = e['raizes']
            esp = max([ult_rz(x) for x in raizes if ult_rz(x)], default=None)
            cad = any(rz_cad.get(x) for x in raizes)
            st_esp = st(esp) if esp else ('INATIVO' if cad else 'NÃO É CLIENTE')
            if (st_esp, esp) != (r['status'], r['data']): P['farmácia avulsa diferente'].append(f"{rot} | arquivo {r['status']} {r['data']} x bruto {st_esp} {esp}")
            continue
        if id(r) in vistos and not e['raizes']: continue
        vistos.add(id(r))
        if r['status'] != 'NÃO É CLIENTE':
            k = tuple(r['cliente'].split(' ')[0].split('/'))
            if r['data']:
                if r['via'] == 'PRODUTO':
                    pr = r['prova'].split("'")[1] if "'" in r['prova'] else ''
                    cand = [d for (kk, des), d in prod.items() if kk == k and des.startswith(pr[:55])]
                    esp = max(cand, default=None)
                else:
                    esp = ult(k)
                if esp != r['data']: P['data mostrada diferente do bruto'].append(f"{rot} | {r['via']} {k} arquivo {r['data']} x bruto {esp}")
                if r['via'] != 'PRODUTO' and cl.get(k) and cl.get(k) == r['data'] and cl_pos.get(k) != r['data'] and uf.get(k) != r['data']:
                    P['data mostrada é só devolução/valor zero'].append(f"{rot} | {k} {r['data']} (última venda > 0: {cl_pos.get(k)})")
                if st(r['data']) != r['status']: P['status contra a data'].append(f"{rot} | {r['status']} {r['data']}")
            elif r['status'] != 'INATIVO':
                P['status sem data e não INATIVO'].append(rot)
        sep = SEPARAR.get(norm(e['extra'].get('nome', '')))
        for x in e['raizes']:
            if x in erradas or sep: continue
            dx = ult_rz(x)
            if dx and (not r['data'] or dx > r['data']):
                P['compra pela raiz do CNPJ mais nova que a mostrada'].append(f"{rot} | raiz {x} {dx} x arquivo {r['status']} {r['data']}")
            if rz_cad.get(x) and r['status'] == 'NÃO É CLIENTE':
                P['raiz no cadastro mas NÃO É CLIENTE'].append(f"{rot} | raiz {x} ({len(rz_cad[x])} cliente+loja)")
    with open(os.path.join(PASTA, 'confere_bruto.txt'), 'w', encoding='utf8') as f:
        for k, v in P.items(): f.write(f'== {k}: {len(v)}\n' + ''.join(f'   {x}\n' for x in v))
    print('linhas conferidas:', n_linhas)
    for k, v in P.items(): print(f'{k}: {len(v)}')
    if not P: print('nenhuma divergência')


if __name__ == '__main__':
    main()
