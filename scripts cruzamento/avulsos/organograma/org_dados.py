"""Dados da aba ORGANOGRAMA ("constante organograma") do MAPA MERCADO.
Estrutura e segmentos: telas do CRM (Segmentos & Vendedores). Carteira de cada pessoa: exportacao do CRM (org_carteira.py)
(ativos <= 290 dias, perdendo 291-365, a trabalhar > 365 ou sem compra; no segmento x fora do segmento).
Potencial: linhas das abas ALIMENTACAO/REDE DE LOJAS/CONSTRUCAO/POSTOS/ESTACIONAMENTO/PET SHOP FINAL (coluna W) e "FARMACIA VF" (coluna P)
com STATUS AUTOPEL diferente de ATIVO. Ale Dayan: abas "(TOP 50)". Uso: python org_dados.py <MAPA.xlsx>  ->  org_dados.json (le equipe.json: pessoas, perfis, segmentos e carteira)"""
import sys, json, collections, openpyxl
MAPA = sys.argv[1]
import os
EQ = json.load(open(os.environ.get('AUTOPEL_ORGANOGRAMA_EQUIPE') or 'equipe.json', encoding='utf8'))  # local, fora do git (nomes e carteira)
CART = {p['id']: (p['carteira']['cards'], p['carteira']['no_segmento'], p['carteira']['fora_segmento']) for p in EQ['pessoas']}
PES = [(p['id'], p['nome'], p['perfil'], p['sup'], p['segmentos']) for p in EQ['pessoas']]
# potencial: segmento do CRM -> (aba, coluna, linha do cabecalho). Abas FINAL: cabecalho na linha 6 desde 05/11/2026 (caixas_final.ps1)
BASE = {'ALIMENTACAO': ('ALIMENTACAO FINAL', 'W', 6), 'REDE DE LOJAS': ('REDE DE LOJAS FINAL', 'W', 6),
        'CONSTRUCAO': ('CONSTRUÇÃO FINAL', 'W', 6), 'DROGARIAS': ('"FARMACIA VF"', 'P', 6),
        # 05/11/2026: abas FINAL novas com STATUS AUTOPEL (ATIVO < 360 dias / INATIVO / NAO E CLIENTE)
        'POSTOS E CONVENIENCIAS': ('POSTOS GAS & CONVENIENCIA FINAL', 'W', 6), 'ESTACIONAMENTOS': ('ESTACIONAMENTO FINAL', 'W', 6),
        'PET SHOP': ('PET SHOP  FINAL', 'W', 6)}
wb = openpyxl.load_workbook(MAPA, read_only=True)
POT = {}
for seg, (aba, col, h) in BASE.items():
    ci = openpyxl.utils.column_index_from_string(col)
    rows = list(wb[aba].iter_rows(min_row=h, min_col=ci, max_col=ci, values_only=True))
    assert str(rows[0][0]).strip().upper() == 'STATUS AUTOPEL', (aba, rows[0][0])
    c = collections.Counter(str(r[0]).strip().upper() for r in rows[1:] if r[0] not in (None, ''))
    POT[seg] = {'aba': aba.strip('"'), 'col': col, 'total': sum(c.values()), 'status': dict(c), 'potencial': sum(v for k, v in c.items() if k != 'ATIVO')}
# 10/11/2026: listas TOP 50 da Ale Dayan (abas "<SEGMENTO> (TOP 50)", cabecalho com EMPRESA e STATUS AUTOPEL), na ordem do usuario.
# A trabalhar = tudo que nao e ATIVO (inativo + nao e cliente), como no potencial.
ORDEM_ALE = [('SUPERMERCADO', 'Supermercado'), ('SERVIÇOS & FRANQUIAS', 'Serviços e franquias'), ('REDE LOJAS', 'Rede de lojas'),
             ('POSTO GASOLINA', 'Posto de gasolina'), ('OTICA', 'Ótica'), ('CONSTRUÇÃO', 'Construção'), ('ESTACIONAMENTO', 'Estacionamento'),
             ('DROGARIA', 'Drogaria'), ('CONVENIENCIA', 'Conveniência'), ('BELEZA', 'Beleza'), ('ALIMENTAÇÃO', 'Alimentação')]
abas = {' '.join(n.split()).upper(): n for n in wb.sheetnames}
ALE = []
for base, nome in ORDEM_ALE:
    aba = abas.get(f'{base} (TOP 50)')
    if not aba:
        print('ATENÇÃO: aba não achada:', base, '(TOP 50)'); continue
    rows = list(wb[aba].iter_rows(values_only=True))
    hi = next(i for i, r in enumerate(rows[:12]) if any(str(x or '').strip().upper() == 'STATUS AUTOPEL' for x in r))
    h = [str(x or '').strip().upper() for x in rows[hi]]
    js = h.index('STATUS AUTOPEL')
    st = [str(r[js]).strip().upper() for r in rows[hi + 1:] if len(r) > js and r[js] not in (None, '')]
    c = collections.Counter(st)
    ALE.append(dict(chave='TOP50 ' + base, nome=nome, aba=aba, col=openpyxl.utils.get_column_letter(js + 1), linha=hi + 1,
                    total=len(st), ativo=c.get('ATIVO', 0), atrab=len(st) - c.get('ATIVO', 0), status=dict(c)))
wb.close()
out = []
for k, nome, perfil, sup, segs in PES:
    top, no, fora = CART[k]
    limpos = [s.lstrip('*') for s in segs]
    pot = {s: POT[s]['potencial'] for s in limpos if s in POT}
    out.append(dict(id=k, nome=nome, perfil=perfil, sup=sup, segmentos=segs, total=sum(top), ativos=top[0], perdendo=top[1], atrab=top[2],
                    no=no, fora=fora, no_total=sum(no), fora_total=sum(fora), potencial=pot))
# potencial das equipes: segmentos atendidos por alguem da equipe (sem somar o mesmo segmento duas vezes)
for o in out:
    if any(x['sup'] == o['id'] for x in out):  # diretoria e supervisores: segmentos de toda a equipe abaixo
        eq = [x for x in out if x['sup'] == o['id'] or any(y['sup'] == o['id'] and x['sup'] == y['id'] for y in out)]
        o['potencial'] = {s: POT[s]['potencial'] for x in eq for s in x['potencial']}
segpes = collections.OrderedDict()
for o in out:
    for s in o['segmentos']:
        segpes.setdefault(s.lstrip('*'), []).append(('★ ' if s.startswith('*') else '') + o['nome'])
json.dump(dict(pessoas=out, potencial=POT, segmentos=segpes, sem_responsavel=EQ['segmentos_sem_responsavel'], ale=ALE, fonte_crm=EQ['_fonte']), open('org_dados.json', 'w', encoding='utf8'), ensure_ascii=False, indent=1)
for o in out: print(o['id'], o['total'], o['no_total'], o['fora_total'], o['potencial'])
print(json.dumps(POT, ensure_ascii=False))
for a in ALE: print('ALE', a['nome'], a['total'], 'ativos', a['ativo'], 'a trabalhar', a['atrab'], '|', a['aba'], a['col'], a['linha'])
