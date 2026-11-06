"""Aba DETALHE do MAPA MERCADO: as listas por tras dos numeros da "ORGANOGRAMA impressao" (05/11/2026).
O usuario queria "clicar no numero e abrir a lista", como o duplo clique de uma tabela dinamica. Sem macro (o arquivo
e .xlsx), cada numero vira um hiperlink para o ponto certo desta aba.
Layout:
  1. POTENCIAL POR PESSOA E EQUIPE: um bloco por pessoa (ancora P|id) com os segmentos que formam o numero e um link
     "ver lista" para a lista de potencial de cada segmento;
  2. um bloco por segmento (ancora S|seg|TOTAL): primeiro os ATIVOS (S|seg|ATIVO), depois o POTENCIAL = tudo que nao e
     ATIVO (S|seg|POTENCIAL), separado por status (S|seg|INATIVO, ...). Cada linha da lista aparece uma vez so.
Colunas: RAZAO SOCIAL, NOME FANTASIA, No DE LOJAS, CNPJ (pedido do usuario) + STATUS AUTOPEL e ORIGEM (aba e linha).
Le o MAPA (abas do potencial, mesmas colunas do org_dados.py), org_dados.json e org_impressao.json; grava detalhe.tsv e detalhe.json
(para org_detalhe.ps1) e acrescenta 'alvo' (linha na DETALHE) nos numeros com 'link' do org_impressao.json.
Uso (na pasta dados, depois do org_dados.py e do org_impressao.py): python ../org_detalhe.py "<MAPA.xlsx>" """
import sys, re, json, collections
import openpyxl
from openpyxl.utils import column_index_from_string as ci
D = json.load(open('org_dados.json', encoding='utf8'))
IMP = json.load(open('org_impressao.json', encoding='utf8'))
ACENTO = {'ALIMENTACAO': 'Alimentação', 'CONSTRUCAO': 'Construção', 'POSTOS E CONVENIENCIAS': 'Postos e Conveniências', 'DROGARIAS': 'Drogarias',
          'REDE DE LOJAS': 'Rede de Lojas', 'ESTACIONAMENTOS': 'Estacionamentos', 'PET SHOP': 'Pet Shop'}
ORDEM = ['ATIVO', 'INATIVO', 'INATIVANDO', 'LEAD NO CRM', 'NÃO É CLIENTE', 'A TRABALHAR']
s = lambda x: str(x).strip() if x is not None else ''
br = lambda n: f'{n:,}'.replace(',', '.')


def fmt_cnpj(v):
    d = re.sub(r'\D', '', s(v))
    return f'{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}' if len(d) == 14 else s(v)


def main(mapa):
    wb = openpyxl.load_workbook(mapa, read_only=True, data_only=True)
    listas = {}
    for seg, b in D['potencial'].items():
        aba = '"FARMACIA VF"' if b['aba'] == 'FARMACIA VF' else b['aba']
        rows = list(wb[aba].iter_rows(values_only=True))
        jst = ci(b['col']) - 1
        hi = next(i for i, r in enumerate(rows[:12]) if len(r) > jst and s(r[jst]).upper() == 'STATUS AUTOPEL')
        h = [s(x).upper() for x in rows[hi]]
        col = lambda nome, ini=False: next(j for j, x in enumerate(h) if (x.startswith(nome) if ini else x == nome))
        jr, jf, jl, jc = col('RAZÃO SOCIAL'), col('NOME FANTASIA'), col('Nº DE LOJAS', True), col('CNPJ')
        L = []
        for n, r in enumerate(rows[hi + 1:], start=hi + 2):
            st = s(r[jst]).upper() if jst < len(r) else ''
            if not st: continue
            lj = r[jl] if jl < len(r) and isinstance(r[jl], (int, float)) else None
            L.append([s(r[jr]), s(r[jf]), lj, fmt_cnpj(r[jc]), st, f"{b['aba']} – linha {n}"])
        cnt = collections.Counter(x[4] for x in L)
        assert dict(cnt) == {k.upper(): v for k, v in b['status'].items()}, ('contagem diferente do org_dados', seg, dict(cnt), b['status'])
        listas[seg] = L
    wb.close()
    linhas, estilos, ancora, links = [], [], {}, []   # links: (linha, coluna, alvo_linha, texto da dica)
    def add(v, est=None):
        linhas.append(v); estilos.append(est); return len(linhas)
    dt = '05/11/2026'
    r = add(['DETALHE – listas por trás dos números da aba "ORGANOGRAMA impressao"', '', '', '', '', '← voltar ao organograma'], 'titulo')
    links.append((r, 6, 'ORGANOGRAMA', 'Voltar à aba ORGANOGRAMA impressao'))
    add([f'Clique num número de potencial ou da tabela por segmento da ORGANOGRAMA impressao para chegar aqui. STATUS AUTOPEL de {dt}: '
         'ATIVO = comprou há menos de 360 dias. POTENCIAL = tudo que não é ATIVO. Cada empresa aparece uma vez, no bloco do seu segmento.',
         '', '', '', '', ''], 'nota')
    add([''] * 6)
    # ---- 1. por pessoa
    r = add(['1. POTENCIAL POR PESSOA E EQUIPE (soma dos segmentos; na supervisão e na diretoria, cada segmento da equipe uma vez)', '', '', '', '', ''], 'secao')
    add(['PESSOA / EQUIPE', 'PERFIL', 'SEGMENTO', 'POTENCIAL', 'ABA', 'LISTA'], 'cab')
    pend = []
    for p in D['pessoas']:
        if not p['potencial']: continue
        r0 = add([p['nome'], p['perfil'], 'TOTAL', sum(p['potencial'].values()), '', ''], 'pessoa')
        ancora[f"P|{p['id']}"] = r0
        for seg, v in sorted(p['potencial'].items(), key=lambda kv: -kv[1]):
            rr = add(['', '', ACENTO.get(seg, seg.title()), v, D['potencial'][seg]['aba'], 'ver lista →'], 'dado')
            pend.append((rr, seg))
    add([''] * 6)
    # ---- 2. por segmento
    add(['2. LISTAS POR SEGMENTO (primeiro os ATIVOS, depois o POTENCIAL, por status)', '', '', '', '', ''], 'secao')
    CAB = ['RAZÃO SOCIAL', 'NOME FANTASIA', 'Nº DE LOJAS', 'CNPJ', 'STATUS AUTOPEL', 'ORIGEM (aba e linha)']
    for seg, L in listas.items():
        b = D['potencial'][seg]
        add([''] * 6)
        r = add([f"{ACENTO.get(seg, seg.title()).upper()} – {b['aba']} (coluna {b['col']}): {br(len(L))} na lista, {br(b['potencial'])} de potencial",
                 '', '', '', '', '← voltar ao organograma'], 'bloco')
        links.append((r, 6, 'ORGANOGRAMA', 'Voltar à aba ORGANOGRAMA impressao'))
        ancora[f'S|{seg}|TOTAL'] = r
        add(CAB, 'cab')
        por = collections.defaultdict(list)
        for x in L: por[x[4]].append(x)
        for st in [x for x in ORDEM if x in por] + sorted(set(por) - set(ORDEM)):
            if st != 'ATIVO' and f'S|{seg}|POTENCIAL' not in ancora:
                ancora[f'S|{seg}|POTENCIAL'] = add([f'POTENCIAL (tudo que não é ATIVO): {br(b["potencial"])}', '', '', '', '', ''], 'pot')
            ancora[f'S|{seg}|{st}'] = add([f'{st}: {br(len(por[st]))}', '', '', '', '', ''], 'status')
            for x in sorted(por[st], key=lambda x: (-(x[2] or 0), x[1] or x[0])):
                add(x, None)
    for rr, seg in pend:
        links.append((rr, 6, ancora[f'S|{seg}|POTENCIAL'], 'Ver a lista de potencial do segmento'))
    # numeros da ORGANOGRAMA impressao -> linha da DETALHE
    sem = []
    for o in IMP['ops']:
        if o.get('link'):
            if o['link'] in ancora: o['alvo'] = ancora[o['link']]
            else: sem.append(o['link'])
    assert not sem, ('link sem destino na DETALHE', sem)
    json.dump(IMP, open('org_impressao.json', 'w', encoding='utf8'), ensure_ascii=False)
    # linhas num TSV (o ConvertFrom-Json do PowerShell 5.1 nao aguenta 90 mil linhas); estilos e links no json
    limpa = lambda v: '' if v is None else re.sub(r'[\t\r\n]+', ' ', str(v))
    with open('detalhe.tsv', 'w', encoding='utf-8-sig', newline='\n') as f:
        for l in linhas: f.write('\t'.join(limpa(v) for v in l) + '\n')
    por_estilo = collections.defaultdict(list)
    for i, e in enumerate(estilos, start=1):
        if e: por_estilo[e].append(i)
    json.dump(dict(n=len(linhas), estilos=por_estilo, links=links), open('detalhe.json', 'w', encoding='utf8'), ensure_ascii=False)
    print('DETALHE:', len(linhas), 'linhas |', len(ancora), 'âncoras |', sum(1 for o in IMP['ops'] if o.get('alvo')), 'números com link |',
          {seg: len(L) for seg, L in listas.items()})


if __name__ == '__main__':
    main(sys.argv[1])
