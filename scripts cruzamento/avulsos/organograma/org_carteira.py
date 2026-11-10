"""Carteira de cada pessoa do organograma a partir da exportacao do CRM (10/11/2026), no lugar das telas.
Entrada: o arquivo "carteira-AAAA-MM-DD-HHMMSS.xlsx" baixado do CRM: uma linha por loja de cliente, com Vendedor, Status
(Ativo / Perdendo / A trabalhar) e Aderencia (No segmento / Fora do segmento / Sem segmento definido).
Conta como o CRM: um cliente por vendedor + codigo de cliente do TOTVS (o codigo vem do "Ultimo faturamento - SQL.csv",
pelo CNPJ); status do cliente = o melhor das suas lojas; aderencia = a mais frequente. Conferido com as telas de 04/11.
  - pessoa = linhas com o nome dela na coluna Vendedor (nome do CRM, igual ao do equipe.json);
  - supervisao = soma das pessoas da equipe;
  - diretoria = a carteira inteira menos a carteira "INATIVOS".
Atualiza dados/equipe.json (campo carteira de cada pessoa e _fonte) e guarda a versao anterior em equipe_AAAAMMDD.json.
Uso (na pasta dados): AUTOPEL_ULTFAT_CSV=<CSV> python ../org_carteira.py "<carteira.xlsx>" """
import os, sys, re, json, shutil, datetime as dt, collections as C
import openpyxl, pandas as pd

ST = {'Ativo': 0, 'Perdendo': 1, 'A trabalhar': 2}
D = lambda s: re.sub(r'\D', '', str(s or ''))


def main(arq):
    m = re.search(r'(\d{4})-(\d{2})-(\d{2})', os.path.basename(arq))
    data = f'{m.group(3)}/{m.group(2)}/{m.group(1)}' if m else dt.date.today().strftime('%d/%m/%Y')
    ws = openpyxl.load_workbook(arq, read_only=True, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    h = [str(x or '').strip() for x in rows[0]]
    iv, ic, ist, iad, icl = h.index('Vendedor'), h.index('CNPJ'), h.index('Status'), h.index('Aderência'), h.index('Cliente')
    U = pd.read_csv(os.environ['AUTOPEL_ULTFAT_CSV'], sep=';', skiprows=1, dtype=str, encoding='utf-8-sig', keep_default_na=False)
    cod = {}
    for c, cn in zip(U.iloc[:, 0].str.strip(), U.iloc[:, 2]):
        cod.setdefault(D(cn), c)
    por = C.defaultdict(dict)          # vendedor -> codigo -> [melhor status, Counter de aderencia]
    for r in rows[1:]:
        v = str(r[iv] or '').strip()
        k = cod.get(D(r[ic])) or '?' + D(r[ic]) + str(r[icl])
        g = por[v].setdefault(k, [9, C.Counter()])
        g[0] = min(g[0], ST[r[ist]])
        g[1][r[iad]] += 1
    def carteira(vendedores):
        cards, no, fora = [0, 0, 0], [0, 0, 0], [0, 0, 0]
        for v in vendedores:
            for s, ad in por.get(v, {}).values():
                cards[s] += 1
                a = ad.most_common(1)[0][0]
                if a == 'No segmento': no[s] += 1
                elif a == 'Fora do segmento': fora[s] += 1
        return dict(cards=cards, no_segmento=no, fora_segmento=fora)
    arq_eq = os.environ.get('AUTOPEL_ORGANOGRAMA_EQUIPE') or 'equipe.json'
    EQ = json.load(open(arq_eq, encoding='utf8'))
    copia = arq_eq.replace('.json', '_' + dt.date.today().strftime('%Y%m%d') + '.json')
    if not os.path.exists(copia):
        shutil.copyfile(arq_eq, copia)
    P = {p['id']: p for p in EQ['pessoas']}
    novo = {}
    for p in EQ['pessoas']:
        if p['perfil'] not in ('Diretoria', 'Supervisor', 'Supervisora'):
            if p['nome'] not in por:
                print('ATENÇÃO: sem linhas na carteira:', p['id'])
            novo[p['id']] = carteira([p['nome']])
    for p in EQ['pessoas']:
        if p['perfil'] in ('Supervisor', 'Supervisora'):
            soma = dict(cards=[0, 0, 0], no_segmento=[0, 0, 0], fora_segmento=[0, 0, 0])
            for q in EQ['pessoas']:
                if q['sup'] == p['id']:
                    for f in soma:
                        soma[f] = [a + b for a, b in zip(soma[f], novo[q['id']][f])]
            novo[p['id']] = soma
        elif p['perfil'] == 'Diretoria':
            novo[p['id']] = carteira([v for v in por if v != 'INATIVOS'])
    for p in EQ['pessoas']:
        antes = p['carteira']
        p['carteira'] = novo[p['id']]
        print(f"{p['perfil'][:13]:13} clientes {sum(antes['cards']):6} -> {sum(p['carteira']['cards']):6} | ativos {antes['cards'][0]:5} -> {p['carteira']['cards'][0]:5}")
    EQ['_fonte'] = (f'carteira exportada do CRM em {data} (uma linha por loja; contada por vendedor + código de cliente); '
                    'equipe e segmentos: telas do CRM (Segmentos & Vendedores) de 04/11/2026')
    json.dump(EQ, open(arq_eq, 'w', encoding='utf8'), ensure_ascii=False, indent=1)
    print('equipe.json atualizado; cópia anterior em', copia)


if __name__ == '__main__':
    main(sys.argv[1])
