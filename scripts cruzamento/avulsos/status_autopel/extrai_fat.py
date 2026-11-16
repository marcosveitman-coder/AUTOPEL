"""Extrai um arquivo "faturamento AAAA*.xlsx" do TOTVS (16/11/2026) para um pickle enxuto.
Lê as colunas pelo nome do cabeçalho (linha 1). Guarda uma linha por cliente+loja+produto, com a nota mais recente
(data, vendedor, valor) e os dados do cliente: CNPJ, grupo (GRP_CLI), razão (CLIENTE), UF, município e segmento, além
da descrição do produto.
Uso: python extrai_fat.py "<faturamento.xlsx>" "<saida.pkl>" """
import sys, pickle, datetime as dt, time
import openpyxl

COLS = ['EMISSAO', 'COD_CLI', 'COD_LOJA', 'COD_VENDEDOR', 'NOME_VENDEDOR', 'CNPJ', 'Estado', 'GRP_CLI', 'CLIENTE',
        'MUNICIPIO', 'COD_PROD', 'DES_PROD', 'VLR_TOTAL', 'SEGMENTO']


def data(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    s = str(v or '').strip()
    for f in ('%Y-%m-%d %H:%M:%S', '%d/%m/%Y', '%Y-%m-%d'):
        try: return dt.datetime.strptime(s[:19] if f.startswith('%Y-%m-%d %H') else s[:10], f).date()
        except ValueError: pass
    return None


def main(arq, saida):
    t = time.time()
    wb = openpyxl.load_workbook(arq, read_only=True)
    out, linhas, sem_data = {}, 0, 0
    for ws in wb.worksheets:
        it = ws.iter_rows(values_only=True)
        cab = [str(x or '').strip() for x in next(it)]
        idx = {c: cab.index(c) for c in COLS if c in cab}
        falta = [c for c in COLS if c not in idx]
        assert not falta or falta == ['MUNICIPIO'], (arq, ws.title, falta)
        g = lambda r, c: (str(r[idx[c]]).strip() if c in idx and r[idx[c]] is not None else '')
        for r in it:
            if r is None or len(r) < len(cab) // 2: continue
            d = data(r[idx['EMISSAO']])
            if d is None: sem_data += 1; continue
            linhas += 1
            k = (g(r, 'COD_CLI'), g(r, 'COD_LOJA'), g(r, 'COD_PROD'))
            try: v = float(r[idx['VLR_TOTAL']] or 0)
            except (TypeError, ValueError): v = 0.0
            ant = out.get(k)
            if ant is None or d > ant[0] or (d == ant[0] and v > ant[3]):
                out[k] = (d, g(r, 'COD_VENDEDOR'), g(r, 'NOME_VENDEDOR'), v, g(r, 'CNPJ'), g(r, 'GRP_CLI'), g(r, 'CLIENTE'),
                          g(r, 'Estado'), g(r, 'MUNICIPIO'), g(r, 'DES_PROD'), g(r, 'SEGMENTO'))
    wb.close()
    dmin = min(v[0] for v in out.values()); dmax = max(v[0] for v in out.values())
    pickle.dump(dict(arquivo=arq, linhas=linhas, sem_data=sem_data, dados=out), open(saida, 'wb'))
    print(f'{arq}: {linhas} linhas ({sem_data} sem data), {len(out)} cliente+loja+produto, {dmin} a {dmax}, {time.time() - t:.0f}s', flush=True)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
