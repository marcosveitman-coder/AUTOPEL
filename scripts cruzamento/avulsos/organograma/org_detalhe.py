"""Aba oculta "CONFIG DETALHE" do MAPA MERCADO (06/11/2026): o que a macro do duplo clique precisa saber.
Com duplo clique num numero da "ORGANOGRAMA impressao", a macro (vba_DetalheMod.txt) monta a aba DETALHE com a lista
por tras do numero. Nenhuma lista fica guardada no arquivo: a macro le na hora
  - carteira (clientes, no segmento, ativos, perdendo, a trabalhar): o CSV "Ultimo faturamento" do TOTVS (um cliente
    por vendedor + codigo, como o CRM; status pela nota mais recente: ativo <= 290 dias, perdendo <= 365, a trabalhar);
  - potencial e tabela por segmento: as abas FINAL e "FARMACIA VF" do proprio arquivo.
Esta etapa grava em config_detalhe.tsv (para o org_detalhe.ps1):
  - CSV, aba reserva, dias, data do CRM;
  - PESSOAS: id, nome, perfil, escopo (VEND = vendedor, EQUIPE = vendedores da equipe, DIR = diretores no TOTVS),
    nomes no TOTVS, segmentos, supervisor e clientes no CRM;
  - SEGMENTOS: aba e coluna do STATUS AUTOPEL de cada lista de mercado;
  - CELULAS: endereco de cada numero clicavel, tipo (C/P/S/E), pessoa, segmento, status, rotulo e o numero mostrado.
Pessoa -> vendedor no TOTVS: dados/detalhe_vendedores.json (local, tem nomes). Se nao existir, e criado pelas regras:
nome do CRM contido no nome do TOTVS (o mais perto do total do CRM), venda interna = carteira "VENDA INDIRETA INTERNA",
supervisao = os vendedores da equipe, diretoria = diretores com 10%+ dos clientes.
Uso (na pasta dados): AUTOPEL_ULTFAT_CSV=<CSV do dia> python ../org_detalhe.py "<MAPA>" """
import os, sys, re, json, collections, datetime as dt
import pandas as pd, openpyxl
from openpyxl.utils import get_column_letter as L
D = json.load(open('org_dados.json', encoding='utf8'))
IMP = json.load(open('org_impressao.json', encoding='utf8'))
EQ = json.load(open(os.environ.get('AUTOPEL_ORGANOGRAMA_EQUIPE') or 'equipe.json', encoding='utf8'))
CSV = os.environ['AUTOPEL_ULTFAT_CSV']
DIAS_ATIVO, DIAS_PERDENDO = 290, 365
ACENTO = {'ALIMENTACAO': 'Alimentação', 'CONSTRUCAO': 'Construção', 'POSTOS E CONVENIENCIAS': 'Postos e Conveniências', 'DROGARIAS': 'Drogarias',
          'REDE DE LOJAS': 'Rede de Lojas', 'ESTACIONAMENTOS': 'Estacionamentos', 'PET SHOP': 'Pet Shop'}
br = lambda n: f'{n:,}'.replace(',', '.')
pct = lambda x: f'{100 * x:.1f}%'.replace('.', ',')


def ultfat():
    U = pd.read_csv(CSV, sep=';', skiprows=1, dtype=str, encoding='utf-8-sig', keep_default_na=False).apply(lambda c: c.str.strip())
    U.columns = [f'c{i}' for i in range(U.shape[1])]
    return U


def vendedores(U):
    arq = 'detalhe_vendedores.json'
    if os.path.exists(arq):
        return json.load(open(arq, encoding='utf8'))
    cli = U.groupby('c15').c0.nunique().to_dict()             # clientes (codigos) por vendedor no TOTVS
    tot = U.groupby(['c15', 'c0']).ngroups
    dirs = U.groupby(['c15', 'c0']).c21.first().value_counts()
    M = {}
    for p in EQ['pessoas']:
        crm = sum(p['carteira']['cards'])
        if p['perfil'] == 'Diretoria':
            M[p['id']] = dict(escopo='DIR', totvs=sorted(n for n, c in dirs.items() if n and c >= 0.10 * tot))
        elif p['perfil'] in ('Supervisor', 'Supervisora'):
            M[p['id']] = dict(escopo='EQUIPE', totvs=[])
        elif p['perfil'] == 'Venda interna':
            M[p['id']] = dict(escopo='VEND', totvs=[n for n in cli if 'VENDA INDIRETA INTERNA' in n])
        else:
            cand = [n for n in cli if all(w in n.split() for w in p['nome'].split())]
            M[p['id']] = dict(escopo='VEND', totvs=sorted(cand, key=lambda n: abs(cli[n] - crm))[:1])
    for p in EQ['pessoas']:
        if M[p['id']]['escopo'] == 'EQUIPE':
            M[p['id']]['totvs'] = sorted({n for q in EQ['pessoas'] if q['sup'] == p['id'] for n in M[q['id']]['totvs']})
    json.dump(M, open(arq, 'w', encoding='utf8'), ensure_ascii=False, indent=1)
    return M


def main(mapa):
    U = ultfat()
    V = vendedores(U)
    data_crm = re.search(r'\d{2}/\d{2}/\d{4}', EQ['_fonte']).group(0)
    caminho = os.path.abspath(CSV)
    perfil = os.environ.get('USERPROFILE', '')
    if perfil and caminho.lower().startswith(perfil.lower()):
        caminho = '%USERPROFILE%' + caminho[len(perfil):]
    linhas = [['CONFIGURAÇÃO DA ABA DETALHE (gerada pelo org_detalhe.py em ' + dt.date.today().strftime('%d/%m/%Y') +
               '). Usada pela macro do duplo clique na ORGANOGRAMA impressao. Só o caminho do CSV pode ser trocado à mão.'],
              ['CSV', caminho], ['ABA RESERVA', 'ULTIMO FATURAMENTO'], ['DIAS ATIVO', str(DIAS_ATIVO)],
              ['DIAS PERDENDO', str(DIAS_PERDENDO)], ['DATA CRM', data_crm], []]
    # pessoas
    linhas += [['PESSOAS'], ['ID', 'NOME', 'PERFIL', 'ESCOPO', 'TOTVS', 'SEGMENTOS', 'SUPERVISOR', 'CLIENTES CRM']]
    for p in EQ['pessoas']:
        v = V[p['id']]
        linhas.append([p['id'], p['nome'], p['perfil'], v['escopo'], ';'.join(v['totvs']), ';'.join(s.lstrip('*') for s in p['segmentos']),
                       p['sup'], str(sum(p['carteira']['cards']))])
    # segmentos (listas de mercado)
    wb = openpyxl.load_workbook(mapa, read_only=True, data_only=True)
    linhas += [[], ['SEGMENTOS'], ['CHAVE', 'NOME', 'ABA', 'COLUNA STATUS', 'LINHA CABECALHO']]
    for seg, b in D['potencial'].items():
        aba = '"FARMACIA VF"' if b['aba'] == 'FARMACIA VF' else b['aba']
        rows = list(wb[aba].iter_rows(min_row=1, max_row=12, values_only=True))
        j = openpyxl.utils.column_index_from_string(b['col']) - 1
        hr = next(i + 1 for i, r in enumerate(rows) if len(r) > j and str(r[j] or '').strip().upper() == 'STATUS AUTOPEL')
        linhas.append([seg, ACENTO.get(seg, seg.title()), aba, b['col'], str(hr)])
    wb.close()
    # celulas clicaveis
    P = {p['id']: p for p in EQ['pessoas']}
    linhas += [[], ['CELULAS'], ['CELULA', 'TIPO', 'QUEM', 'SEGMENTO', 'STATUS', 'ROTULO', 'NUMERO']]
    vistos = set()
    for o in IMP['ops']:
        d = o.get('det')
        if not d:
            continue
        cel = f"{L(o['c'])}{o['r']}"
        assert cel not in vistos, cel
        vistos.add(cel)
        n = d['n']
        if d['t'] == 'C':
            num = f"{br(n)} clientes" + (f" ({pct(o['v'])})" if o.get('fmt') == 'p' else '') + f" no CRM de {data_crm}"
            linhas.append([cel, 'C', d['q'], d['s'], d['st'], f"{P[d['q']]['nome']} – {d['rot']}", num])
        elif d['t'] == 'P':
            num = f"{br(n)} oportunidades (" + ', '.join(ACENTO.get(s, s) for s in d['segs'].split(';')) + ')'
            linhas.append([cel, 'P', d['q'], d['segs'], 'POTENCIAL', f"{P[d['q']]['nome']} – {d['rot']}", num])
        elif d['t'] == 'S':
            linhas.append([cel, 'S', '', d['segs'], d['st'], d['rot'], f'{br(n)} na tabela por segmento'])
        elif d['t'] == 'E':
            linhas.append([cel, 'E', d['q'], '', '', f"{P[d['q']]['nome']} – {d['rot']}", f'{n} pessoas'])
    limpa = lambda v: re.sub(r'[\t\r\n]+', ' ', str(v))
    with open('config_detalhe.tsv', 'w', encoding='utf-8-sig', newline='') as f:
        for l in linhas:
            f.write('\t'.join(limpa(v) for v in (l + [''] * 8)[:8]) + '\r\n')
    print('CONFIG DETALHE:', len(linhas), 'linhas |', len(vistos), 'números clicáveis')
    for p in EQ['pessoas']:
        print(f"   {p['perfil'][:13]:13} {V[p['id']]['escopo']:6} {len(V[p['id']]['totvs'])} nome(s) no TOTVS")


if __name__ == '__main__':
    main(sys.argv[1])
