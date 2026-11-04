"""Desenho da aba "ORGANOGRAMA impressao": as mesmas informacoes da aba ORGANOGRAMA, em folhas A4 paisagem:
folha 1 e 2 = uma equipe por folha (supervisor no centro, resumo da equipe e legenda dos lados, equipe em caixas de
6 colunas, 4 por linha); folha 3 = carteira de cada pessoa; folha 4 = potencial por segmento e segmentos x
responsaveis. 30 colunas iguais (B..AE). Gera org_impressao.json para o org_grava.ps1."""
import json
D = json.load(open('org_dados.json', encoding='utf8'))
P = {p['id']: p for p in D['pessoas']}
POT = D['potencial']
NAVY, BLUE, LIGHT, GREEN, GRAY, WHITE, DARK = (31, 58, 95), (46, 117, 182), (221, 235, 247), (226, 239, 218), (237, 237, 237), (255, 255, 255), (31, 58, 95)
VERDE, CINZA = (0, 97, 0), (128, 128, 128)
COR = {'Diretoria': (NAVY, WHITE), 'Supervisora': (BLUE, WHITE), 'Supervisor': (BLUE, WHITE), 'Vendedora': (LIGHT, DARK), 'Vendedor': (LIGHT, DARK),
       'Venda interna': (GREEN, DARK), 'Representante': (GRAY, DARK)}
C0, NC = 2, 30                      # colunas B..AE
CF = C0 + NC - 1
ops, H = [], {}                     # H: altura de cada linha (todas fixas, para calcular as folhas)
def put(r, c, v, c2=None, **f):
    ops.append(dict(r=r, c=c, c2=c2 or c, v=v, **f))
def alt(r, h): H[r] = max(H.get(r, 0), h)
def br(n): return f'{n:,}'.replace(',', '.')
def pc(a, b): return f'{100 * a / b:.1f}%'.replace('.', ',') if b else '—'
ACENTO = {'ALIMENTACAO': 'Alimentação', 'CONSTRUCAO': 'Construção', 'CORPORATIVO SAUDE': 'Corporativo Saúde', 'POSTOS E CONVENIENCIAS': 'Postos e Conveniências',
          'COSMETICOS': 'Cosméticos', 'LOGISTICA': 'Logística', 'ORGAO PUBLICO': 'Órgão Público', 'PEDAGIO': 'Pedágio', 'AUTOMOTIVO PEÇAS E LOCADORAS': 'Automotivo, Peças e Locadoras',
          'REDE DE LOJAS': 'Rede de Lojas', 'E-COMMERCE': 'E-commerce', 'FABRICANTES DE EQUIPAMENTOS': 'Fabricantes de Equipamentos', 'AUTOVIAÇÃO': 'Autoviação',
          'CORPORATIVO EDUCACIONAL': 'Corporativo Educacional'}
def nomeseg(s):
    s = s.lstrip('*'); return ACENTO.get(s, s.title().replace(' E ', ' e ').replace(' De ', ' de '))
def segtxt(p): return ', '.join(('★ ' if s.startswith('*') else '') + nomeseg(s) for s in p['segmentos'])
DIR = next(p['id'] for p in D['pessoas'] if not p['sup'])
SUP = [p['id'] for p in D['pessoas'] if p['sup'] == DIR]
MEMB = {s: sorted([p['id'] for p in D['pessoas'] if p['sup'] == s], key=lambda k: -P[k]['total']) for s in SUP}
SUP.sort(key=lambda s: len(MEMB[s]))            # mesma ordem da aba ORGANOGRAMA (equipe menor primeiro)
CARD_H = [17, 12, 24, 11, 14, 11, 14, 13]        # alturas das 8 linhas de uma caixa
def partes(c, larg, n):
    """divide `larg` colunas a partir de c em n partes quase iguais (a do meio fica com a sobra)."""
    base, sobra = divmod(larg, n); tam = [base] * n
    for i in range(sobra): tam[(n // 2 + i) % n] += 1
    out = []
    for t in tam: out.append((c, c + t - 1)); c += t
    return out
def card(p, r, c):
    fill, font = COR[p['perfil']]
    for i, h in enumerate(CARD_H): alt(r + i, h)
    segs = segtxt(p) or ('Todos os segmentos' if not p['sup'] else 'Equipe inteira' if p['id'] in SUP else 'Sem segmento no CRM')
    no, fora = p['no_total'], p['fora_total']
    put(r, c, p['nome'], c + 5, bold=True, size=11, fill=fill, font=font, halign='center', wrap=True)
    put(r + 1, c, p['perfil'].upper(), c + 5, italic=True, size=8, fill=fill, font=font, halign='center')
    put(r + 2, c, segs, c + 5, size=8, wrap=True, halign='center')
    (a1, a2), (b1, b2) = partes(c, 6, 2)
    lab = dict(size=7, bold=True, font=CINZA, halign='center')
    put(r + 3, a1, 'CLIENTES', a2, **lab); put(r + 3, b1, 'NO SEGMENTO' if no else 'SEGMENTO', b2, **lab)
    put(r + 4, a1, p['total'], a2, fmt='n', size=10, bold=True, halign='center')
    put(r + 4, b1, f"{br(no)} ({pc(no, no + fora)})" if no else 'não definido', b2, size=10 if no else 8, bold=bool(no), halign='center')
    base, arr = (no, p['no']) if no else (p['total'], [p['ativos'], p['perdendo'], p['atrab']])
    for (x1, x2), t, v in zip(partes(c, 6, 3), ['% ATIVOS', '% PERDENDO', '% A TRAB.'], arr):
        put(r + 5, x1, t, x2, **lab); put(r + 6, x1, v / base if base else '—', x2, fmt='p' if base else None, size=10, bold=True, halign='center')
    pot = sum(p['potencial'].values())
    put(r + 7, c, ('Potencial: ' + br(pot) + ' oportunidades') if pot else 'Potencial: sem lista de mercado', c + 5, size=8, bold=True,
        halign='center', font=VERDE if pot else CINZA, fill=GREEN if pot else None)
    ops.append(dict(box=[r, c, r + 7, c + 5]))
    return r + 7
def titulo(r, texto, sub):
    put(r, C0, texto, CF, bold=True, size=16, font=NAVY, halign='center'); alt(r, 26)
    put(r + 1, C0, sub, CF, size=8, italic=True, halign='center', font=CINZA); alt(r + 1, 13)
    alt(r + 2, 6)
    return r + 3
def faixa(r, texto, h=16, size=10):
    put(r, C0, texto, CF, bold=True, size=size, fill=NAVY, font=WHITE, halign='center'); alt(r, h)
def rotulo(ids):
    return 'VENDEDORAS' if all(P[k]['perfil'] == 'Vendedora' for k in ids) else 'VENDEDORES E VENDA INTERNA' if any(P[k]['perfil'] == 'Venda interna' for k in ids) else 'VENDEDORES'
def caixas(ids, r):
    """caixas de 6 colunas, 4 por linha, 2 colunas entre elas, centralizadas nas 30 colunas."""
    for i in range(0, len(ids), 4):
        lin = ids[i:i + 4]; larg = 6 * len(lin) + 2 * (len(lin) - 1); start = C0 + (NC - larg) // 2
        for j, k in enumerate(lin): ult = card(P[k], r, start + 8 * j)
        r = ult + 1; alt(r, 8); r += 1
    return r
quebras = []
# ---------- folhas 1 e 2: uma equipe por folha
for s in SUP:
    sp = P[s]; r = 1 if not quebras and not ops else r
    if ops: quebras.append(r)
    eq = MEMB[s]; vend = [k for k in eq if P[k]['perfil'] != 'Representante']; rep = [k for k in eq if P[k]['perfil'] == 'Representante']
    r = titulo(r, f"CONSTANTE ORGANOGRAMA – EQUIPE {sp['nome']}",
               f"Equipe comercial Autopel · {sp['perfil'].lower()} {sp['nome'].title()}, abaixo da diretoria · carteira do CRM e potencial das listas de mercado")
    # faixa do supervisor: legenda (esquerda) | caixa do supervisor (centro) | resumo da equipe (direita)
    cc = C0 + (NC - 6) // 2                    # colunas 14..19
    L1, L2, R1, R2 = C0, cc - 2, cc + 7, CF     # 11 colunas de cada lado, 1 coluna de folga
    card(sp, r, cc)
    dp = P[DIR]
    put(r, L1, 'DIRETORIA (A QUEM A EQUIPE RESPONDE)', L2, bold=True, size=8, fill=NAVY, font=WHITE, halign='center')
    put(r + 1, L1, f"{dp['nome']} · carteira {br(dp['total'])} clientes", L2, size=8, halign='center')
    put(r + 2, L1, 'ATIVO = comprou há até 290 dias · PERDENDO = 291 a 365 dias · A TRABALHAR = mais de 365 dias ou sem compra · ★ = segmento principal no CRM',
        L2, size=8, wrap=True, halign='center')
    put(r + 3, L1, 'LEGENDA DAS CAIXAS', L2, bold=True, size=8, fill=NAVY, font=WHITE, halign='center')
    chips = [('Supervisão', COR['Supervisor']), ('Vendedor(a)', COR['Vendedor']), ('Venda interna', COR['Venda interna']),
             ('Representante', COR['Representante']), ('Com potencial', (GREEN, VERDE)), ('Sem lista', ((255, 255, 255), CINZA))]
    pos = [(L1, L1 + 2), (L1 + 4, L1 + 6), (L1 + 8, L2)]
    for i, (t, (fl, fo)) in enumerate(chips):
        x1, x2 = pos[i % 3]; put(r + 4 + i // 3, x1, t, x2, bold=True, size=8, fill=fl, font=fo, halign='center', border=True)
    put(r + 6, L1, '% = clientes no segmento do CRM (quem não tem segmento: carteira toda)', L2, size=7, italic=True, halign='center')
    put(r + 7, L1, 'Potencial = lista do segmento com status ≠ ATIVO; não somar entre pessoas', L2, size=7, italic=True, halign='center')
    no = sp['no_total']; fora = sp['fora_total']
    put(r, R1, 'RESUMO DA EQUIPE – NO SEGMENTO (CRM)', R2, bold=True, size=8, fill=NAVY, font=WHITE, halign='center')
    tres = partes(R1, R2 - R1 + 1, 3)
    lab = dict(bold=True, size=7, fill=LIGHT, font=DARK, halign='center', border=True)
    val = dict(size=10, bold=True, halign='center', border=True)
    for (x1, x2), t, v, fm in zip(tres, ['PESSOAS', 'CLIENTES', '% NO SEGMENTO'], [len(eq), sp['total'], no / (no + fora)], ['n', 'n', 'p']):
        put(r + 1, x1, t, x2, **lab); put(r + 2, x1, v, x2, fmt=fm, **{**val, 'size': 12})
    for (x1, x2), t, v in zip(tres, ['ATIVOS', 'PERDENDO', 'A TRABALHAR'], sp['no']):
        put(r + 3, x1, t, x2, **lab); put(r + 4, x1, v, x2, fmt='n', **val)
        put(r + 5, x1, '% ' + t, x2, **lab); put(r + 6, x1, v / no, x2, fmt='p', **val)
    pot = sum(sp['potencial'].values())
    put(r + 7, R1, f"Potencial dos segmentos da equipe: {br(pot)} oportunidades", R2, size=8, bold=True, font=VERDE, fill=GREEN, halign='center', border=True)
    r += 8
    ops.append(dict(vline=[cc + 2, r, r])); alt(r, 10); r += 1
    faixa(r, f"EQUIPE {sp['nome']} – {rotulo(vend)} ({len(vend)})"); alt(r + 1, 6); r = caixas(vend, r + 2)
    if rep:
        faixa(r, f"EQUIPE {sp['nome']} – REPRESENTANTES ({len(rep)})"); alt(r + 1, 6); r = caixas(rep, r + 2)
    H.pop(r - 1); r -= 1             # sem a folga depois da ultima linha de caixas
# ---------- folha 3: carteira de cada pessoa
quebras.append(r)
r = titulo(r, 'CONSTANTE ORGANOGRAMA – CARTEIRA DE CADA PESSOA', 'Clientes no CRM, dentro e fora do segmento de cada pessoa, e o potencial dos seus segmentos')
faixa(r, '1. CARTEIRA DE CADA PESSOA – DENTRO E FORA DO SEGMENTO (CRM)', 20, 11); r += 1
for c1, c2, t in [(2, 11, ''), (12, 17, 'CARTEIRA (QUANTIDADE E %)'), (18, 23, 'NO SEGMENTO (QUANTIDADE E %)'), (24, 29, 'FORA DO SEGMENTO (QUANTIDADE E %)'), (30, 31, 'POTENCIAL')]:
    put(r, c1, t, c2, bold=True, size=8, fill=BLUE, font=WHITE, halign='center', border=True)
alt(r, 16); r += 1
H2 = [(2, 5, 'NOME'), (6, 7, 'PERFIL'), (8, 11, 'SEGMENTOS (★ principal)'), (12, 13, 'CLIENTES'), (14, 15, 'NO SEGMENTO'), (16, 17, 'FORA DO SEGMENTO'),
      (18, 19, 'ATIVOS'), (20, 21, 'PERDENDO'), (22, 23, 'A' + chr(10) + 'TRABALHAR'), (24, 25, 'ATIVOS'), (26, 27, 'PERDENDO'), (28, 29, 'A' + chr(10) + 'TRABALHAR'), (30, 31, 'OPORTU-' + chr(10) + 'NIDADES')]
for c1, c2, t in H2: put(r, c1, t, c2, bold=True, size=7, fill=LIGHT, font=DARK, halign='center', wrap=True, border=True)
alt(r, 24); r += 1
def qp(a, b): return br(a) + chr(10) + pc(a, b) if b else '—'   # quantidade em cima, % embaixo
for k in [DIR] + [x for s in SUP for x in [s] + MEMB[s]]:
    p = P[k]; no, fora = p['no_total'], p['fora_total']; tot = no + fora
    destaque = k in [DIR] + SUP
    f = dict(size=8, border=True, bold=destaque, fill=GRAY if destaque else None)
    put(r, 2, p['nome'] if destaque else '   ' + p['nome'], 5, wrap=True, **f); put(r, 6, p['perfil'], 7, wrap=True, **{**f, 'size': 7})
    put(r, 8, segtxt(p) or '—', 11, wrap=True, **{**f, 'size': 7})
    vals = [(12, br(p['total'])), (14, qp(no, tot)), (16, qp(fora, tot))]
    for base, arr, n in [(18, p['no'], no), (24, p['fora'], fora)]:
        vals += [(base + 2 * i, qp(arr[i], n)) for i in range(3)]
    pot = sum(p['potencial'].values())
    vals.append((30, br(pot) if pot else '—'))
    for c, v in vals: put(r, c, v, c + 1, halign='center', wrap=True, **f)
    alt(r, 24); r += 1
semseg = [P[k] for k in [DIR] + SUP if P[k]['no_total'] + P[k]['fora_total'] < P[k]['total']]
quem = ' e '.join(('na diretoria' if not p['sup'] else 'no supervisor ' + p['nome'].title()) for p in semseg)
put(r, 2, f'CLIENTES = ativos + perdendo + a trabalhar (cartões do CRM). {quem[0].upper() + quem[1:]} há clientes sem segmento definido: por isso NO + FORA pode ser menor que CLIENTES. '
          'Os supervisores mostram a visão do CRM de cada um, que não é a soma da equipe. Potencial: soma dos segmentos da pessoa (diretoria e supervisão: cada segmento da equipe uma vez).',
    CF, size=8, italic=True, wrap=True); alt(r, 26); r += 1
# ---------- folha 4: potencial por segmento + segmentos x responsaveis
quebras.append(r)
r = titulo(r, 'CONSTANTE ORGANOGRAMA – POTENCIAL E SEGMENTOS', 'Oportunidades nas listas de mercado deste arquivo e quem responde por cada segmento no CRM')
faixa(r, '2. POTENCIAL POR SEGMENTO – OPORTUNIDADES NAS LISTAS DE MERCADO (TUDO QUE NÃO É CLIENTE ATIVO)', 20, 11); r += 1
H3 = [(2, 4, 'SEGMENTO'), (5, 8, 'ABA (COLUNA STATUS AUTOPEL)'), (9, 10, 'TOTAL NA LISTA'), (11, 12, 'ATIVO'), (13, 14, 'INATIVO'), (15, 16, 'INATIVANDO'),
      (17, 18, 'LEAD NO CRM'), (19, 20, 'NÃO É CLIENTE'), (21, 22, 'A TRABALHAR'), (23, 24, 'POTENCIAL'), (25, 31, 'RESPONSÁVEIS (★ principal)')]
for c1, c2, t in H3: put(r, c1, t, c2, bold=True, size=8, fill=LIGHT, font=DARK, halign='center', wrap=True, border=True)
alt(r, 28); r += 1
for seg, b in POT.items():
    st = b['status']; f = dict(size=8, border=True)
    put(r, 2, nomeseg(seg), 4, bold=True, **f); put(r, 5, f"{b['aba']} (coluna {b['col']})", 8, wrap=True, **f)
    for (c1, c2), k in zip([(9, 10), (11, 12), (13, 14), (15, 16), (17, 18), (19, 20), (21, 22)], [None, 'ATIVO', 'INATIVO', 'INATIVANDO', 'LEAD NO CRM', 'NÃO É CLIENTE', 'A TRABALHAR']):
        put(r, c1, b['total'] if k is None else st.get(k, 0), c2, fmt='n', halign='center', **f)
    put(r, 23, b['potencial'], 24, fmt='n', halign='center', bold=True, fill=GREEN, **f)
    put(r, 25, ', '.join(D['segmentos'].get(seg, [])), 31, wrap=True, **{**f, 'size': 7})
    alt(r, 30); r += 1
put(r, 2, 'POTENCIAL = linhas da lista do segmento com STATUS AUTOPEL diferente de ATIVO (inativo, inativando, lead no CRM, não é cliente, a trabalhar). '
          'O potencial é do segmento: quem divide o segmento divide as mesmas oportunidades (não somar entre pessoas). Segmentos sem lista de mercado nestas abas '
          '(Supermercadista, Revenda, Transporte, Corporativo…) aparecem sem potencial.', CF, size=8, italic=True, wrap=True); alt(r, 26); r += 1
alt(r, 10); r += 1
faixa(r, '3. SEGMENTOS E RESPONSÁVEIS (CRM)', 20, 11); r += 1
segs = sorted(D['segmentos'].items(), key=lambda x: (-len(x[1]), x[0])) + [(s, []) for s in D['sem_responsavel']]
meio = (len(segs) + 1) // 2
lados = [(2, 5, 6, 8, 15), (18, 21, 22, 24, 31)]   # segmento | pessoas | responsaveis (14 colunas de cada lado, 2 de folga)
for a, b_, pc_, rc, rf in lados:
    for c1, c2, t in [(a, b_, 'SEGMENTO'), (pc_, pc_ + 1, 'PESSOAS'), (rc, rf, 'RESPONSÁVEIS (★ principal)')]:
        put(r, c1, t, c2, bold=True, size=7, fill=LIGHT, font=DARK, halign='center', wrap=True, border=True)
alt(r, 22); r += 1
for i in range(meio):
    for (a, b_, pc_, rc, rf), j in zip(lados, [i, i + meio]):
        if j >= len(segs): continue
        seg, ps = segs[j]; f = dict(size=8, border=True)
        put(r + 0, a, nomeseg(seg), b_, bold=True, wrap=True, **f); put(r, pc_, len(ps), pc_ + 1, fmt='n', halign='center', **f)
        put(r, rc, ', '.join(ps) if ps else 'sem responsável', rf, wrap=True, **{**f, 'size': 7, 'font': CINZA if not ps else None})
    alt(r, 20); r += 1
put(r, 2, 'Fonte: telas do CRM (Segmentos & Vendedores e carteira de cada pessoa), recebidas em 04/11/2026; potencial: abas ALIMENTACAO, REDE DE LOJAS e CONSTRUÇÃO FINAL (coluna W) '
          'e "FARMACIA VF" (coluna P) deste arquivo.', CF, size=8, italic=True, wrap=True); alt(r, 24)
ultima = r
for o in ops: o.pop('valign', None)
ops += [dict(rowh=[k, v]) for k, v in sorted(H.items())]
# altura de cada folha (pontos) -> zoom para caber em A4 paisagem
lim = [1] + quebras + [ultima + 1]
folhas = [sum(H.get(x, 15) for x in range(a, b)) for a, b in zip(lim, lim[1:])]
# A4 paisagem 842 x 595 pt; margens 0,6 cm dos lados e 0,8 cm em cima e embaixo; colunas de 4,3 (35 px) e 1 (12 px)
larg = (NC * 35 + 2 * 12) * 0.75
zoom = int(min(1, (595 - 2 * 22.7) / max(folhas), (842 - 2 * 17) / larg) * 100)
json.dump(dict(ops=ops, ultima=ultima, aba='ORGANOGRAMA impressao', colunas=dict(primeira=1, ultima=CF + 1, largura=4.3, margem=1), quebras=quebras,
               folhas=folhas, zoom=zoom), open('org_impressao.json', 'w', encoding='utf8'), ensure_ascii=False)
print('operacoes', len(ops), 'ultima', ultima, 'quebras', quebras, 'alturas das folhas (pt)', folhas, 'zoom', zoom)
