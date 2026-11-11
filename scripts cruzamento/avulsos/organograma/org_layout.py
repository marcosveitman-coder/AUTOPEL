"""Desenho da aba ORGANOGRAMA: gera org_layout.json com as operacoes de celula (mesclar, valor, formato) que o
org_grava.ps1 aplica pelo Excel. Colunas B..Y (24 colunas iguais): metade esquerda (B..M) = equipe menor,
metade direita (N..Y) = equipe maior; caixas de 4 colunas, 3 por linha, centralizadas em cada metade."""
import json
D = json.load(open('org_dados.json', encoding='utf8'))
P = {p['id']: p for p in D['pessoas']}
POT = D['potencial']
NAVY, BLUE, LIGHT, GREEN, GRAY, WHITE, DARK = (31, 58, 95), (46, 117, 182), (221, 235, 247), (226, 239, 218), (237, 237, 237), (255, 255, 255), (31, 58, 95)
COR = {'Diretoria': (NAVY, WHITE), 'Supervisora': (BLUE, WHITE), 'Supervisor': (BLUE, WHITE), 'Vendedora': (LIGHT, DARK), 'Vendedor': (LIGHT, DARK),
       'Venda interna': (GREEN, DARK), 'Representante': (GRAY, DARK)}
ops = []
def put(r, c, v, c2=None, **f):
    ops.append(dict(r=r, c=c, c2=c2 or c, v=v, **f))
def br(n): return f'{n:,}'.replace(',', '.')
def pc(a, b): return f'{100 * a / b:.1f}%'.replace('.', ',') if b else '—'
ACENTO = {'ALIMENTACAO': 'Alimentação', 'CONSTRUCAO': 'Construção', 'CORPORATIVO SAUDE': 'Corporativo Saúde', 'POSTOS E CONVENIENCIAS': 'Postos e Conveniências',
          'LOGISTICA': 'Logística', 'ORGAO PUBLICO': 'Órgão Público', 'PEDAGIO': 'Pedágio', 'AUTOMOTIVO PEÇAS E LOCADORAS': 'Automotivo, Peças e Locadoras',
          'REDE DE LOJAS': 'Rede de Lojas', 'E-COMMERCE': 'E-commerce', 'FABRICANTES DE EQUIPAMENTOS': 'Fabricantes de Equipamentos', 'AUTOVIAÇÃO': 'Autoviação',
          'CORPORATIVO EDUCACIONAL': 'Corporativo Educacional', 'REDE SMART': 'Rede Smart (grupos)', 'COSMETICOS': 'Cosméticos e Perfumaria'}
def nomeseg(s):
    s = s.lstrip('*'); return ACENTO.get(s, s.title().replace(' E ', ' e ').replace(' De ', ' de '))
def card(p, r, c):
    fill, font = COR[p['perfil']]
    segs = ', '.join(('★ ' if s.startswith('*') else '') + nomeseg(s) for s in p['segmentos']) or ('Todos os segmentos' if not p['sup'] else 'Equipe inteira' if p['id'] in SUP else 'Sem segmento no CRM')
    no, fora = p['no_total'], p['fora_total']
    pot = sum(p['potencial'].values())
    potxt = ('Potencial: ' + br(pot) + ' oportunidades') if pot else 'Potencial: sem base nas abas de mercado'
    l4 = f"Clientes: {br(p['total'])}  |  no segmento {br(no)} ({pc(no, no + fora)})" if no + fora else f"Clientes: {br(p['total'])}  |  sem segmento definido" if p['total'] else "Vendedor novo: ainda sem carteira"
    l5 = f"No segmento: ativos {pc(p['no'][0], no)} · perdendo {pc(p['no'][1], no)} · a trabalhar {pc(p['no'][2], no)}" if no else \
         f"Ativos {br(p['ativos'])} · perdendo {br(p['perdendo'])} · a trabalhar {br(p['atrab'])}"
    linhas = [(p['nome'], dict(bold=True, size=11, fill=fill, font=font, h=20)), (p['perfil'].upper(), dict(italic=True, size=8, fill=fill, font=font, h=14)),
              (segs, dict(size=8, wrap=True, h=34)), (l4, dict(size=8, h=15)), (l5, dict(size=8, wrap=True, h=24)), (potxt, dict(size=8, bold=True, h=15, font=(0, 97, 0) if pot else (128, 128, 128)))]
    for i, (t, f) in enumerate(linhas):
        put(r + i, c, t, c + 3, halign='center', valign='center', **f)
    ops.append(dict(box=[r, c, r + len(linhas) - 1, c + 3]))
    return r + len(linhas) - 1
def vline(col, r1, r2): ops.append(dict(vline=[col, r1, r2]))
def hline(row, c1, c2): ops.append(dict(hline=[row, c1, c2]))
# ---------- 1) titulo
put(1, 2, 'CONSTANTE ORGANOGRAMA – EQUIPE COMERCIAL AUTOPEL', 25, bold=True, size=16, font=NAVY, halign='center', h=28)
put(2, 2, 'Quem está embaixo de quem, os segmentos de cada pessoa, a carteira no CRM (dentro e fora do segmento) e o potencial de cada segmento.', 25, size=9, italic=True, halign='center')
put(3, 2, 'Status da carteira: ATIVO = comprou há até 290 dias · PERDENDO = 291 a 365 dias · A TRABALHAR = mais de 365 dias ou sem compra.  ★ = segmento principal da pessoa no CRM.', 25, size=9, halign='center')
# ---------- 2) organograma em caixas
DIR = next(p['id'] for p in D['pessoas'] if not p['sup'])
SUP = [p['id'] for p in D['pessoas'] if p['sup'] == DIR]
MEMB = {s: sorted([p['id'] for p in D['pessoas'] if p['sup'] == s], key=lambda k: -P[k]['total']) for s in SUP}
SUP.sort(key=lambda s: len(MEMB[s]))  # equipe menor na metade esquerda
fim = card(P[DIR], 5, 12)                       # diretoria no centro (colunas L..O)
vline(13, fim + 1, fim + 1); hline(fim + 1, 8, 19); vline(7, fim + 2, fim + 2); vline(19, fim + 2, fim + 2)
rs = fim + 3
card(P[SUP[0]], rs, 6); fim_s = card(P[SUP[1]], rs, 18)
EQ_S = MEMB[SUP[0]]
VEND_A = [k for k in MEMB[SUP[1]] if P[k]['perfil'] != 'Representante']
REP_A = [k for k in MEMB[SUP[1]] if P[k]['perfil'] == 'Representante']
def rotulo(ids):
    return 'VENDEDORAS' if all(P[k]['perfil'] == 'Vendedora' for k in ids) else 'VENDEDORES E VENDA INTERNA' if any(P[k]['perfil'] == 'Venda interna' for k in ids) else 'VENDEDORES'
def bloco(ids, r, c0, titulo):
    """caixas de 4 colunas, 3 por linha, centralizadas na metade que comeca na coluna c0 (12 colunas)."""
    put(r, c0, titulo, c0 + 11, bold=True, size=9, fill=NAVY, font=WHITE, halign='center', h=16); r += 2
    for i in range(0, len(ids), 3):
        lin = ids[i:i + 3]; start = c0 + (12 - 4 * len(lin)) // 2
        ult = r
        for j, k in enumerate(lin): ult = card(P[k], r, start + 4 * j)
        r = ult + 2
    return r
vline(7, fim_s + 1, fim_s + 1); vline(19, fim_s + 1, fim_s + 1)
r0 = fim_s + 2
fs = bloco(EQ_S, r0, 2, f"EQUIPE {P[SUP[0]]['nome']} – {rotulo(EQ_S)} ({len(EQ_S)})")
fa = bloco(VEND_A, r0, 14, f"EQUIPE {P[SUP[1]]['nome']} – {rotulo(VEND_A)} ({len(VEND_A)})")
if REP_A: fa = bloco(REP_A, fa, 14, f"EQUIPE {P[SUP[1]]['nome']} – REPRESENTANTES ({len(REP_A)})")
# resumo das equipes e legenda, embaixo da equipe menor (equilibra as duas metades)
rr = fs
put(rr, 2, 'RESUMO DAS EQUIPES – NO SEGMENTO (CRM)', 13, bold=True, size=9, fill=NAVY, font=WHITE, halign='center'); rr += 1
for c1, c2, t in [(2, 4, 'EQUIPE'), (5, 5, 'PESSOAS'), (6, 6, 'CLIENTES'), (7, 7, '% NO SEGM.'), (8, 8, 'ATIVOS'), (9, 9, '% ATIVOS'),
                  (10, 10, 'PERDENDO'), (11, 11, '% PERD.'), (12, 12, 'A TRAB.'), (13, 13, '% A TRAB.')]:
    put(rr, c1, t, c2, bold=True, size=8, fill=LIGHT, font=DARK, halign='center', border=True)
rr += 1
for k, n in [(SUP[0], len(EQ_S)), (SUP[1], len(VEND_A) + len(REP_A))]:
    p = P[k]; no, fora = p['no_total'], p['fora_total']; f = dict(size=8, border=True, halign='center')
    put(rr, 2, 'Equipe ' + p['nome'].title(), 4, bold=True, size=8, border=True)
    for c, v, t in [(5, n, 'n'), (6, p['total'], 'n'), (7, no / (no + fora), 'p'), (8, p['no'][0], 'n'), (9, p['no'][0] / no, 'p'), (10, p['no'][1], 'n'),
                    (11, p['no'][1] / no, 'p'), (12, p['no'][2], 'n'), (13, p['no'][2] / no, 'p')]:
        put(rr, c, v, c, fmt=t, **f)
    rr += 1
rr += 1
put(rr, 2, 'LEGENDA DAS CAIXAS', 13, bold=True, size=9, fill=NAVY, font=WHITE, halign='center'); rr += 1
for i, (t, perf) in enumerate([('Diretoria', 'Diretoria'), ('Supervisão', 'Supervisor'), ('Vendedor(a)', 'Vendedor'), ('Venda interna', 'Venda interna')]):
    fill, font = COR[perf]; put(rr, 2 + 3 * i, t, 4 + 3 * i, bold=True, size=8, fill=fill, font=font, halign='center', border=True)
rr += 1
fill, font = COR['Representante']; put(rr, 2, 'Representante', 4, bold=True, size=8, fill=fill, font=font, halign='center', border=True)
put(rr, 5, 'Verde = há potencial (lista de mercado do segmento); cinza = segmento ainda sem lista de mercado.', 13, size=8, italic=True)
r = max(rr + 2, fa) + 1
# ---------- 3) tabela de metricas
put(r, 2, '1. CARTEIRA DE CADA PESSOA – DENTRO E FORA DO SEGMENTO (CRM)', 25, bold=True, size=11, fill=NAVY, font=WHITE, h=20); r += 1
H1 = [(2, 7, ''), (8, 12, 'CARTEIRA'), (13, 18, 'NO SEGMENTO'), (19, 24, 'FORA DO SEGMENTO'), (25, 25, 'POTENCIAL')]
for c1, c2, t in H1: put(r, c1, t, c2, bold=True, size=9, fill=BLUE, font=WHITE, halign='center', border=True)
r += 1
H2 = [(2, 3, 'NOME'), (4, 4, 'PERFIL'), (5, 7, 'SEGMENTOS (★ principal)'), (8, 8, 'CLIENTES'), (9, 9, 'NO SEGM.'), (10, 10, '% NO SEGM.'), (11, 11, 'FORA SEGM.'),
      (12, 12, '% FORA'), (13, 13, 'ATIVOS'), (14, 14, '% ATIVOS'), (15, 15, 'PERDENDO'), (16, 16, '% PERD.'), (17, 17, 'A TRABALHAR'), (18, 18, '% A TRAB.'),
      (19, 19, 'ATIVOS'), (20, 20, '% ATIVOS'), (21, 21, 'PERDENDO'), (22, 22, '% PERD.'), (23, 23, 'A TRABALHAR'), (24, 24, '% A TRAB.'), (25, 25, 'OPORTUNID.')]
for c1, c2, t in H2: put(r, c1, t, c2, bold=True, size=8, fill=LIGHT, font=DARK, halign='center', wrap=True, border=True, h=30)
r += 1
ORDEM = [DIR, SUP[0]] + EQ_S + [SUP[1]] + VEND_A + REP_A
for k in ORDEM:
    p = P[k]; no, fora = p['no_total'], p['fora_total']; tot = no + fora
    destaque = k in [DIR] + SUP
    f = dict(size=8, border=True, bold=destaque, fill=GRAY if destaque else None)
    nome = p['nome'] if destaque else '   ' + p['nome']
    put(r, 2, nome, 3, wrap=True, **f); put(r, 4, p['perfil'], 4, wrap=True, **{**f, 'size': 7})
    put(r, 5, ', '.join(('★ ' if s.startswith('*') else '') + nomeseg(s) for s in p['segmentos']) or '—', 7, wrap=True, **f)
    vals = [(8, p['total'], 'n'), (9, no, 'n'), (10, no / tot if tot else None, 'p'), (11, fora, 'n'), (12, fora / tot if tot else None, 'p')]
    for base, arr, n in [(13, p['no'], no), (19, p['fora'], fora)]:
        for i in range(3):
            vals += [(base + 2 * i, arr[i], 'n'), (base + 2 * i + 1, arr[i] / n if n else None, 'p')]
    pot = sum(p['potencial'].values())
    vals.append((25, pot if pot else None, 'n'))
    for c, v, t in vals: put(r, c, v if v is not None else '—', c, fmt=t if v is not None else None, halign='center', **f)
    ops.append(dict(rowh=[r, 24])); r += 1
semseg = [P[k] for k in [DIR] + SUP if P[k]['no_total'] + P[k]['fora_total'] < P[k]['total']]
quem = ' e '.join(('na diretoria' if not p['sup'] else 'no supervisor ' + p['nome'].title()) for p in semseg)
put(r, 2, f'CLIENTES = ativos + perdendo + a trabalhar (cartões do CRM). {quem[0].upper() + quem[1:]} há clientes sem segmento definido: por isso NO + FORA pode ser menor que CLIENTES. Os supervisores mostram a visão do CRM de cada um.', 25, size=8, italic=True, wrap=True, h=24); r += 2
# ---------- 4) potencial por segmento
put(r, 2, '2. POTENCIAL POR SEGMENTO – OPORTUNIDADES NAS LISTAS DE MERCADO (TUDO QUE NÃO É CLIENTE ATIVO)', 25, bold=True, size=11, fill=NAVY, font=WHITE, h=20); r += 1
H3 = [(2, 4, 'SEGMENTO'), (5, 7, 'ABA (COLUNA STATUS AUTOPEL)'), (8, 8, 'TOTAL NA LISTA'), (9, 9, 'ATIVO'), (10, 10, 'INATIVO'), (11, 11, 'INATIVANDO'),
      (12, 12, 'LEAD NO CRM'), (13, 13, 'NÃO É CLIENTE'), (14, 14, 'A TRABALHAR'), (15, 15, 'POTENCIAL'), (16, 25, 'RESPONSÁVEIS (★ principal)')]
for c1, c2, t in H3: put(r, c1, t, c2, bold=True, size=8, fill=LIGHT, font=DARK, halign='center', wrap=True, border=True, h=30)
r += 1
for seg, b in POT.items():
    st = b['status']; f = dict(size=8, border=True)
    put(r, 2, nomeseg(seg), 4, bold=True, **f); put(r, 5, f"{b['aba']} (coluna {b['col']})", 7, **f)
    for c, k in [(8, None), (9, 'ATIVO'), (10, 'INATIVO'), (11, 'INATIVANDO'), (12, 'LEAD NO CRM'), (13, 'NÃO É CLIENTE'), (14, 'A TRABALHAR')]:
        v = b['total'] if k is None else st.get(k, 0)
        put(r, c, v, c, fmt='n', halign='center', **f)
    put(r, 15, b['potencial'], 15, fmt='n', halign='center', bold=True, fill=GREEN, **{x: y for x, y in f.items()})
    put(r, 16, ', '.join(D['segmentos'].get(seg, [])), 25, wrap=True, **f)
    ops.append(dict(rowh=[r, 26])); r += 1
put(r, 2, 'POTENCIAL = linhas da lista do segmento com STATUS AUTOPEL diferente de ATIVO (inativo, inativando, lead no CRM, não é cliente, a trabalhar). '
          'O potencial é do segmento: quem divide o segmento divide as mesmas oportunidades (não somar entre pessoas). Segmentos sem lista de mercado nestas abas '
          '(Supermercadista, Revenda, Transporte, Corporativo…) aparecem sem potencial.', 25, size=8, italic=True, wrap=True, h=36); r += 2
# ---------- 5) segmentos x responsaveis
put(r, 2, '3. SEGMENTOS E RESPONSÁVEIS (CRM)', 25, bold=True, size=11, fill=NAVY, font=WHITE, h=20); r += 1
for c1, c2, t in [(2, 6, 'SEGMENTO'), (7, 7, 'PESSOAS'), (8, 25, 'RESPONSÁVEIS (★ principal)')]:
    put(r, c1, t, c2, bold=True, size=8, fill=LIGHT, font=DARK, halign='center', border=True)
r += 1
SEM = D['sem_responsavel']
segs = sorted(D['segmentos'].items(), key=lambda x: (-len(x[1]), x[0]))
for seg, ps in segs + [(s, []) for s in SEM]:
    f = dict(size=8, border=True)
    put(r, 2, nomeseg(seg), 6, bold=True, **f); put(r, 7, len(ps), 7, fmt='n', halign='center', **f)
    put(r, 8, ', '.join(ps) if ps else 'sem responsável', 25, font=(128, 128, 128) if not ps else None, **f); r += 1
put(r, 2, 'Fonte: ' + D.get('fonte_crm', 'telas do CRM') + '; potencial: abas ALIMENTACAO, REDE DE LOJAS, CONSTRUÇÃO, POSTOS GAS & CONVENIENCIA, ESTACIONAMENTO e PET SHOP FINAL (coluna W, STATUS AUTOPEL de 05/11/2026) e "FARMACIA VF" (coluna P) deste arquivo.', 25, size=8, italic=True, h=15)
json.dump(dict(ops=ops, ultima=r), open('org_layout.json', 'w', encoding='utf8'), ensure_ascii=False)
print('operacoes', len(ops), 'ultima linha', r)
