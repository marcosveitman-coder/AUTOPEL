"""Desenho da aba "ORGANOGRAMA impressao": as mesmas informacoes da aba ORGANOGRAMA, para imprimir em A4 paisagem a
100%, com letra Verdana de tamanho 10 ou maior e sem italico.
Folhas: uma por grupo de caixas (equipe menor; equipe maior - vendedores e venda interna; equipe maior -
representantes), depois a carteira de cada pessoa (uma folha por equipe), o potencial por segmento (folha 6) e os
segmentos x responsaveis (folha 7, desde 05/11/2026).
Cada numero leva 'det' (o que ele conta: carteira C, potencial P, tabela por segmento S, pessoas da equipe E): o
org_detalhe.py grava esse significado na aba oculta CONFIG DETALHE, e a macro do duplo clique monta a lista na aba DETALHE. Grade de 62 colunas estreitas (B..BK); caixas de 14 colunas, 4 por linha.
As alturas das linhas sao estimadas pela largura das letras da Verdana (texto quebrado em celula mesclada nao tem
autoajuste) e cada folha e conferida contra a altura util da pagina. Gera org_impressao.json para o org_grava.ps1."""
import json, math
D = json.load(open('org_dados.json', encoding='utf8'))
P = {p['id']: p for p in D['pessoas']}
POT = D['potencial']
NAVY, BLUE, LIGHT, GREEN, GRAY, WHITE, DARK = (31, 58, 95), (46, 117, 182), (221, 235, 247), (226, 239, 218), (237, 237, 237), (255, 255, 255), (31, 58, 95)
OURO, OURO_T, CINZA = (255, 242, 204), (127, 96, 0), (89, 89, 89)
COR = {'Diretoria': (NAVY, WHITE), 'Supervisora': (BLUE, WHITE), 'Supervisor': (BLUE, WHITE), 'Vendedora': (LIGHT, DARK), 'Vendedor': (LIGHT, DARK),
       'Venda interna': (GREEN, DARK), 'Representante': (GRAY, DARK)}
C0, NC = 2, 63                      # colunas B..BL (4 caixas de 15 colunas + 3 de folga)
CF = C0 + NC - 1
GAPS = {C0 + 15, C0 + 31, C0 + 47}  # colunas estreitas entre as caixas
def px(c): return 5 if c in GAPS else 16   # 16 px = 12 pt; a impressora so imprime ~757 pt de largura (margem fisica)
def span(c1, c2): return sum(px(c) for c in range(c1, c2 + 1))
FOLGA = 1.10                        # o Excel desenha a letra mais larga que a medida (escala de tela)
LIM = 556                           # altura util (pt) de A4 paisagem com margens de 0,6/0,7 cm
ops, H = [], {}
# ---------- largura das letras (Verdana, fracao do tamanho) e estimativa de linhas
import os as _os
VERD = {'Regular': {}, 'Bold': {}}
for _l in open(_os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'verdana10.txt'), encoding='utf-8-sig'):
    _st, _cp, _w = _l.split('\t'); VERD[_st][chr(int(_cp))] = float(_w)
def larg_txt(t, size, bold):
    """largura em px (96 dpi) do texto em Verdana, com a folga."""
    tab = VERD['Bold' if bold else 'Regular']
    return sum(tab.get(c, tab['M']) for c in t) * size / 10 * FOLGA
FALHAS = []
def linhas(t, larg, size=10, bold=False, indent=0):
    util = larg - 7 - indent * 9
    for w in str(t).replace('\n', ' ').split():
        if larg_txt(w, size, bold) > util: FALHAS.append((w, larg))
    n = 0
    for par in str(t).split('\n'):
        atual, n_par = '', 1
        for w in par.split(' '):
            tent = (atual + ' ' + w).strip()
            if larg_txt(tent, size, bold) <= util: atual = tent
            else:
                n_par += 1 if atual else 0; atual = w
                while larg_txt(atual, size, bold) > util: n_par += 1; atual = atual[len(atual) // 2:]
        n += n_par
    return n
def altura(t, larg, size=10, bold=False, indent=0): return max(15, round(linhas(t, larg, size, bold, indent) * size * 1.275 + 3, 1))
def put(r, c, v, c2=None, size=10, auto=True, **f):
    c2 = c2 or c
    _txt = pc(v, 1) if f.get('fmt') == 'p' and isinstance(v, float) else br(v) if f.get('fmt') == 'n' and isinstance(v, int) else str(v)
    if f.pop('uma', False) and larg_txt(_txt, size, f.get('bold', False)) > span(c, c2) - 7 - f.get('indent', 0) * 9:
        FALHAS.append((str(v), span(c, c2)))
    ops.append(dict(r=r, c=c, c2=c2, v=v, size=size, **f))
    if auto and f.get('wrap', True): alt(r, altura(_txt, span(c, c2), size, f.get('bold', False), f.get('indent', 0)))
def alt(r, h): H[r] = max(H.get(r, 0), h)
def br(n): return f'{n:,}'.replace(',', '.')
def pc(a, b): return f'{100 * a / b:.1f}%'.replace('.', ',') if b else '—'
def qp(a, b): return f'{br(a)} ({pc(a, b)})' if b else '—'
ACENTO = {'ALIMENTACAO': 'Alimentação', 'CONSTRUCAO': 'Construção', 'CORPORATIVO SAUDE': 'Corporativo Saúde', 'POSTOS E CONVENIENCIAS': 'Postos e Conveniências',
          'COSMETICOS': 'Cosméticos', 'LOGISTICA': 'Logística', 'ORGAO PUBLICO': 'Órgão Público', 'PEDAGIO': 'Pedágio', 'AUTOMOTIVO PEÇAS E LOCADORAS': 'Automotivo, Peças e Locadoras',
          'REDE DE LOJAS': 'Rede de Lojas', 'E-COMMERCE': 'E-commerce', 'FABRICANTES DE EQUIPAMENTOS': 'Fabricantes de Equipamentos', 'AUTOVIAÇÃO': 'Autoviação',
          'CORPORATIVO EDUCACIONAL': 'Corporativo Educacional'}
def nomeseg(s):
    s = s.lstrip('*'); return ACENTO.get(s, s.title().replace(' E ', ' e ').replace(' De ', ' de '))
def segtxt(p): return ', '.join(('★ ' if s.startswith('*') else '') + nomeseg(s) for s in p['segmentos'])
def partes(c, larg, n):
    base, sobra = divmod(larg, n); tam = [base] * n
    for i in range(sobra): tam[(n // 2 + i) % n] += 1
    out = []
    for t in tam: out.append((c, c + t - 1)); c += t
    return out
DIR = next(p['id'] for p in D['pessoas'] if not p['sup'])
SUP = [p['id'] for p in D['pessoas'] if p['sup'] == DIR]
MEMB = {s: sorted([p['id'] for p in D['pessoas'] if p['sup'] == s], key=lambda k: -P[k]['total']) for s in SUP}
SUP.sort(key=lambda s: len(MEMB[s]))            # mesma ordem da aba ORGANOGRAMA (equipe menor primeiro)
def sup_txt(p): return p['perfil'].lower()
# ---------- Ale Dayan (10/11/2026): segmento do CRM -> abas "(TOP 50)" (chave do org_dados.json: 'TOP50 ' + nome da aba)
ALE_K = {a['chave']: a for a in D.get('ale', [])}
ALE_SEG = {'SUPERMERCADISTA': ['SUPERMERCADO'], 'ALIMENTACAO': ['ALIMENTAÇÃO'], 'DROGARIAS': ['DROGARIA'], 'REDE DE LOJAS': ['REDE LOJAS'],
           'POSTOS E CONVENIENCIAS': ['POSTO GASOLINA', 'CONVENIENCIA'], 'CONSTRUCAO': ['CONSTRUÇÃO'], 'ESTACIONAMENTOS': ['ESTACIONAMENTO'],
           'COSMETICOS': ['BELEZA']}   # Cosmeticos = Beleza (decisao do usuario, 10/11)
ALE_ST = [('ATIVO', 'ativos', GREEN), ('INATIVO', 'inativos', OURO), ('NÃO É CLIENTE', 'não é cliente', GRAY)]
PERFIS_ALE = ('Vendedora', 'Vendedor', 'Venda interna', 'Representante')
def ale_da_pessoa(p):
    out = []
    for s in p['segmentos']:
        for aba in ALE_SEG.get(s.lstrip('*'), []):
            k = 'TOP50 ' + aba
            if k in ALE_K and k not in out: out.append(k)
    return out
# ---------- caixas
def card(p, r, c):
    fill, font = COR[p['perfil']]
    no, fora = p['no_total'], p['fora_total']
    put(r, c, p['nome'], c + 14, size=11, bold=True, fill=fill, font=font, halign='center', wrap=True)
    put(r + 1, c, p['perfil'], c + 14, fill=fill, font=font, halign='center')
    put(r + 2, c, segtxt(p) or 'Sem segmento no CRM', c + 14, halign='center', wrap=True)
    base, arr = (no, p['no']) if no else (p['total'], [p['ativos'], p['perdendo'], p['atrab']])
    pot = sum(p['potencial'].values())
    C = lambda s, st, rot, n: dict(t='C', q=p['id'], s=s, st=st, rot=rot, n=n)
    put(r + 3, c, ' Clientes', c + 6, border=True, uma=True)
    put(r + 3, c + 7, p['total'], c + 14, fmt='n', bold=True, halign='center', border=True, uma=True, det=C('TODOS', 'TODOS', 'Clientes', p['total']))
    put(r + 4, c, ' No segmento', c + 6, border=True, uma=True)
    if no:
        put(r + 4, c + 7, no, c + 10, fmt='n', bold=True, halign='center', border=True, uma=True, det=C('NO', 'TODOS', 'Clientes no segmento', no))
        put(r + 4, c + 11, no / (no + fora), c + 14, fmt='p', bold=True, halign='center', border=True, uma=True, det=C('NO', 'TODOS', 'Clientes no segmento', no))
    else:
        put(r + 4, c + 7, 'sem segmento', c + 14, bold=True, halign='center', border=True, uma=True)
    for i, (t, v) in enumerate(zip(['Ativos', 'Perdendo', 'A trabalhar'], arr)):
        put(r + 5 + i, c, ' ' + t, c + 6, border=True, uma=True)
        put(r + 5 + i, c + 7, v / base if base else '—', c + 14, fmt='p' if base else None, bold=True, halign='center', border=True, uma=True,
            det=C('NO' if no else 'TODOS', 'APT'[i], t + (' no segmento' if no else ''), v) if base else None)
    extra = dict(fill=OURO, font=OURO_T) if pot else dict(font=CINZA)
    put(r + 8, c, ' Potencial', c + 6, border=True, uma=True, **extra)
    put(r + 8, c + 7, pot if pot else 'sem lista', c + 14, fmt='n' if pot else None, bold=True, halign='center', border=True, uma=True,
        det=dict(t='P', q=p['id'], segs=';'.join(p['potencial']), rot='Potencial (listas de mercado)', n=pot) if pot else None, **extra)
    if p['perfil'] not in PERFIS_ALE:
        ops.append(dict(box=[r, c, r + 8, c + 14]))
        return r + 8
    # 10/11/2026: linha Ale Dayan = listas TOP 50 dos segmentos da pessoa (ativos, inativos, nao e cliente; cores na legenda)
    ch = ale_da_pessoa(p)
    put(r + 9, c, ' Ale Dayan', c + 5, border=True, uma=True, **(dict() if ch else dict(font=CINZA)))
    if ch:
        nomes = ', '.join(ALE_K[k]['nome'] for k in ch)
        for (x1, x2), (st, rot, fl) in zip(partes(c + 6, 9, 3), ALE_ST):
            v = sum(ALE_K[k]['status'].get(st, 0) for k in ch)
            put(r + 9, x1, v, x2, fmt='n', bold=True, halign='center', border=True, uma=True, fill=fl, font=DARK,
                det=dict(t='S', segs=';'.join(ch), st=st, rot=f"{p['nome']} – Ale Dayan TOP 50 ({nomes}) – {rot}", n=v) if v else None)
    else:
        put(r + 9, c + 6, 'sem lista', c + 14, bold=True, halign='center', border=True, uma=True, font=CINZA)
    ops.append(dict(box=[r, c, r + 9, c + 14]))
    return r + 9
def caixas(ids, r):
    for i in range(0, len(ids), 4):
        lin = ids[i:i + 4]; larg = 15 * len(lin) + 1 * (len(lin) - 1); start = C0 + (NC - larg) // 2
        for j, k in enumerate(lin): ult = card(P[k], r, start + 16 * j)
        r = ult + 1
        if i + 4 < len(ids): alt(r, 10); r += 1
    return r
def titulo(r, texto, sub):
    put(r, C0, texto, CF, size=16, bold=True, font=NAVY, halign='center'); alt(r, 26)
    if sub is None: alt(r + 1, 6); return r + 2
    put(r + 1, C0, sub, CF, font=CINZA, halign='center')
    alt(r + 2, 6)
    return r + 3
def faixa(r, texto, size=11):
    put(r, C0, texto, CF, size=size, bold=True, fill=NAVY, font=WHITE, halign='center'); return r + 1
def legenda(r):
    alt(r, 10); r += 1
    chips = [('Supervisão', COR['Supervisor']), ('Vendedor(a)', COR['Vendedor']), ('Venda interna', COR['Venda interna']),
             ('Representante', COR['Representante']), ('Potencial', (OURO, OURO_T))]
    for (x1, x2), (t, (fl, fo)) in zip(partes(C0, NC, 5), chips): put(r, x1, t, x2, bold=True, fill=fl, font=fo, halign='center', border=True)
    put(r + 1, C0, 'ATIVO = comprou há até 290 dias · PERDENDO = 291 a 365 dias · A TRABALHAR = mais de 365 dias ou sem compra · ★ = segmento principal',
        CF, halign='center', wrap=True)
    put(r + 2, C0, 'Percentuais: sobre os clientes no segmento (sem segmento: carteira toda) · Potencial: lista do segmento com status ≠ ATIVO',
        CF, halign='center', wrap=True)
    # 10/11/2026: cores da linha Ale Dayan das caixas
    ale = [('Ale Dayan (TOP 50 do segmento):', None)] + [(rot.capitalize(), fl) for _, rot, fl in ALE_ST]
    for (x1, x2), (tx, fl) in zip(partes(C0, NC, 4), ale):
        put(r + 3, x1, tx, x2, bold=True, halign='center', **(dict(fill=fl, font=DARK, border=True) if fl else {}))
    return r + 4
def banner(r, sp, eq):
    no, fora = sp['no_total'], sp['fora_total']; pot = sum(sp['potencial'].values())
    put(r, C0, f"{sp['nome']} – {sp['perfil'].upper()}", CF, size=12, bold=True, fill=BLUE, font=WHITE, halign='center')
    cel = [('Pessoas', len(eq)), ('Clientes', sp['total']), ('No segmento', qp(no, no + fora))] + \
          [(t, qp(v, no)) for t, v in zip(['Ativos', 'Perdendo', 'A trabalhar'], sp['no'])] + [('Potencial', pot)]
    DET = {'Pessoas': dict(t='E', q=sp['id'], rot='Pessoas da equipe', n=len(eq)), 'Clientes': dict(t='C', q=sp['id'], s='TODOS', st='TODOS', rot='Clientes', n=sp['total']),
           'No segmento': dict(t='C', q=sp['id'], s='NO', st='TODOS', rot='Clientes no segmento', n=no),
           'Potencial': dict(t='P', q=sp['id'], segs=';'.join(sp['potencial']), rot='Potencial (listas de mercado)', n=pot)}
    for i, t in enumerate(['Ativos', 'Perdendo', 'A trabalhar']): DET[t] = dict(t='C', q=sp['id'], s='NO', st='APT'[i], rot=t + ' no segmento', n=sp['no'][i])
    for (x1, x2), (t, v) in zip(partes(C0, NC, 7), cel):
        extra = dict(fill=OURO, font=OURO_T) if t == 'Potencial' else {}
        put(r + 1, x1, t, x2, bold=True, fill=extra.get('fill', LIGHT), font=extra.get('font', DARK), halign='center', border=True, wrap=True)
        put(r + 2, x1, v, x2, fmt='n' if isinstance(v, int) else None, size=10, bold=True, halign='center', border=True, uma=True,
            det=DET[t] if v else None, **extra)
    ops.append(dict(box=[r, C0, r + 2, CF]))
    r += 3
    ops.append(dict(vline=[C0 + NC // 2 - 1, r, r])); alt(r, 12)
    return r + 1
def rotulo(ids):
    return 'VENDEDORAS' if all(P[k]['perfil'] == 'Vendedora' for k in ids) else 'VENDEDORES E VENDA INTERNA' if any(P[k]['perfil'] == 'Venda interna' for k in ids) else 'VENDEDORES'
quebras, folhas_ini = [], []
def nova_folha(r):
    if ops: quebras.append(r)
    folhas_ini.append(r); return r
dp = P[DIR]
r = 1
# ---------- folhas das equipes (caixas)
for s in SUP:
    sp = P[s]; eq = MEMB[s]
    vend = [k for k in eq if P[k]['perfil'] != 'Representante']; rep = [k for k in eq if P[k]['perfil'] == 'Representante']
    grupos = [(vend, rotulo(vend))] + ([(rep, 'REPRESENTANTES')] if rep else [])
    for n, (ids, rot) in enumerate(grupos):
        r = nova_folha(r)
        parte = f" ({n + 1} de {len(grupos)})" if len(grupos) > 1 else ''
        r = titulo(r, f"CONSTANTE ORGANOGRAMA – EQUIPE {sp['nome']}{parte}",
                   f"{sup_txt(sp).capitalize()} {sp['nome'].title()} · responde à {dp['nome'].title()} (carteira de {br(dp['total'])} clientes no CRM)")
        r = banner(r, sp, eq)
        r = faixa(r, f"{rot} ({len(ids)})"); alt(r, 8); r += 1
        r = caixas(ids, r)
        if n == len(grupos) - 1:   # 10/11/2026: legenda so na ultima folha da equipe (a linha Ale Dayan nao cabia com ela)
            r = legenda(r)
# ---------- carteira de cada pessoa: uma folha por equipe
T1 = [(2, 14)] + [(15 + 5 * i, 19 + 5 * i) for i in range(9)] + [(60, 64)]
def t1_cab(r):
    for c1, c2, t in [(2, 14, ''), (15, 29, 'CARTEIRA'), (30, 44, 'NO SEGMENTO'), (45, 59, 'FORA DO SEGMENTO'), (60, 64, '')]:
        put(r, c1, t, c2, bold=True, fill=BLUE, font=WHITE, halign='center', border=True)
    r += 1
    for (c1, c2), t in zip(T1, ['Nome e perfil', 'Clientes', 'No segmento', 'Fora do segmento', 'Ativos', 'Perdendo', 'A trabalhar',
                                'Ativos', 'Perdendo', 'A trabalhar', 'Potencial']):
        put(r, c1, t, c2, bold=False, fill=LIGHT, font=DARK, halign='center', wrap=True, border=True)
    return r + 1
def t1_linha(r, k):
    p = P[k]; no, fora = p['no_total'], p['fora_total']; tot = no + fora
    destaque = k in [DIR] + SUP
    f = dict(border=True, bold=destaque, fill=GRAY if destaque else None, wrap=True)
    put(r, 2, p['nome'] + '\n' + p['perfil'], 13, indent=1, **f)
    q = lambda a, b: f'{br(a)}\n{pc(a, b)}' if b else '—'
    vals = [p['total'], q(no, tot), q(fora, tot)] + [q(p['no'][i], no) for i in range(3)] + [q(p['fora'][i], fora) for i in range(3)]
    pot = sum(p['potencial'].values())
    vals.append(pot if pot else '—')
    C = lambda s, st, rot, n: dict(t='C', q=k, s=s, st=st, rot=rot, n=n)
    STS = ['Ativos', 'Perdendo', 'A trabalhar']
    DETS = [C('TODOS', 'TODOS', 'Clientes', p['total']), C('NO', 'TODOS', 'Clientes no segmento', no), C('FORA', 'TODOS', 'Clientes fora do segmento', fora)]
    DETS += [C('NO', 'APT'[i], t + ' no segmento', p['no'][i]) for i, t in enumerate(STS)]
    DETS += [C('FORA', 'APT'[i], t + ' fora do segmento', p['fora'][i]) for i, t in enumerate(STS)]
    DETS += [dict(t='P', q=k, segs=';'.join(p['potencial']), rot='Potencial (listas de mercado)', n=pot)]
    for i, ((c1, c2), v) in enumerate(zip(T1[1:], vals)):
        put(r, c1, v, c2, fmt='n' if isinstance(v, int) else None, halign='center', det=DETS[i] if v != '—' else None, **f)
    return r + 1
semseg = [P[k] for k in [DIR] + SUP if P[k]['no_total'] + P[k]['fora_total'] < P[k]['total']]
quem = ' e '.join(('na diretoria' if not p['sup'] else 'no supervisor ' + p['nome'].title()) for p in semseg)
NOTA1 = (f'CLIENTES = ativos + perdendo + a trabalhar (CRM). {quem[0].upper() + quem[1:]} há clientes sem segmento: NO + FORA pode ser menor que CLIENTES. '
         'Supervisão = visão do CRM de cada um (não é a soma da equipe). Potencial: soma dos segmentos (na supervisão, cada segmento da equipe uma vez).')
for n, (rot, ids) in enumerate([('DIRETORIA E EQUIPE ' + P[SUP[0]]['nome'], [DIR, SUP[0]] + MEMB[SUP[0]])] +
                               [('EQUIPE ' + P[s]['nome'], [s] + MEMB[s]) for s in SUP[1:]]):
    r = nova_folha(r)
    r = titulo(r, 'CONSTANTE ORGANOGRAMA – CARTEIRA DE CADA PESSOA', None)
    r = faixa(r, f'1. CARTEIRA DE CADA PESSOA – {rot}')
    r = t1_cab(r)
    for k in ids: r = t1_linha(r, k)
    alt(r, 6); r += 1
    put(r, C0, NOTA1, CF, wrap=True); r += 1
# ---------- potencial por segmento + segmentos x responsaveis
r = nova_folha(r)
r = titulo(r, 'CONSTANTE ORGANOGRAMA – POTENCIAL E SEGMENTOS', None)
r = faixa(r, '2. POTENCIAL POR SEGMENTO – LISTAS DE MERCADO DESTE ARQUIVO')
T2 = [(2, 11), (12, 23)] + [(24 + 5 * i, 28 + 5 * i) for i in range(3)] + [(39, 44)] + [(45 + 5 * i, 49 + 5 * i) for i in range(4)]
for (c1, c2), t in zip(T2, ['Segmento', 'Aba (coluna do status)', 'Total na lista', 'Ativo', 'Inativo', 'Inativando', 'Lead no CRM', 'Não é cliente', 'A trabalhar', 'Potencial']):
    put(r, c1, t, c2, bold=False, fill=LIGHT, font=DARK, halign='center', wrap=True, border=True)
r += 1
for seg, b in POT.items():
    st = b['status']; f = dict(border=True)
    put(r, 2, nomeseg(seg), 11, bold=True, **f); put(r, 12, f" {b['aba']} ({b['col']})", 23, **f)
    for (c1, c2), k in zip(T2[2:9], [None, 'ATIVO', 'INATIVO', 'INATIVANDO', 'LEAD NO CRM', 'NÃO É CLIENTE', 'A TRABALHAR']):
        v = b['total'] if k is None else st.get(k, 0)
        put(r, c1, v, c2, fmt='n', halign='center', det=dict(t='S', segs=seg, st=k or 'TOTAL', rot=f"{nomeseg(seg)} – {(k or 'total na lista').lower()}", n=v) if v else None, **f)
    put(r, 60, b['potencial'], 64, fmt='n', halign='center', bold=True, fill=OURO, font=OURO_T,
        det=dict(t='S', segs=seg, st='POTENCIAL', rot=f'{nomeseg(seg)} – potencial (tudo que não é ATIVO)', n=b['potencial']) if b['potencial'] else None, **f)
    r += 1
alt(r, 6); r += 1
put(r, C0, 'POTENCIAL = linhas da lista do segmento com STATUS AUTOPEL diferente de ATIVO. O potencial é do segmento: quem divide o segmento divide as mesmas '
           'oportunidades (não somar entre pessoas). Segmentos sem lista de mercado nestas abas aparecem sem potencial.', CF, wrap=True); r += 1
alt(r, 12); r += 1
# 05/11/2026: com POSTOS, ESTACIONAMENTO e PET SHOP no potencial, a folha 6 passou da pagina -> segmentos x responsaveis na folha 7
r = nova_folha(r)
r = titulo(r, 'CONSTANTE ORGANOGRAMA – SEGMENTOS E RESPONSÁVEIS', None)
r = faixa(r, '3. SEGMENTOS E RESPONSÁVEIS (CRM)')
T3 = [(2, 15), (16, 21), (22, 64)]
for (c1, c2), t in zip(T3, ['Segmento', 'Pessoas', 'Responsáveis (★ = segmento principal da pessoa)']):
    put(r, c1, t, c2, bold=True, fill=LIGHT, font=DARK, halign='center', wrap=True, border=True)
r += 1
for seg, ps in sorted(D['segmentos'].items(), key=lambda x: (-len(x[1]), x[0])):
    f = dict(border=True)
    put(r, 2, nomeseg(seg), 15, bold=True, indent=1, **f); put(r, 16, len(ps), 21, fmt='n', halign='center', **f)
    put(r, 22, ', '.join(ps), 64, indent=1, wrap=True, **f); r += 1
SEM = D['sem_responsavel']
put(r, 2, 'Sem responsável', 15, bold=True, indent=1, border=True, font=CINZA); put(r, 16, len(SEM), 21, fmt='n', halign='center', border=True, font=CINZA)
put(r, 22, ', '.join(nomeseg(s) for s in SEM), 64, indent=1, wrap=True, border=True, font=CINZA); r += 1
alt(r, 6); r += 1
put(r, C0, 'Fonte: ' + D.get('fonte_crm', 'telas do CRM') + '; potencial: abas ALIMENTACAO, REDE DE LOJAS, CONSTRUÇÃO, POSTOS GAS & '
           'CONVENIENCIA, ESTACIONAMENTO e PET SHOP FINAL (coluna W, STATUS AUTOPEL) e "FARMACIA VF" (coluna P); Ale Dayan: abas "(TOP 50)" deste arquivo.', CF, wrap=True)
ultima = r
for o in ops: o.setdefault('wrap', True)
ops += [dict(rowh=[k, v]) for k, v in sorted(H.items())]
lim = folhas_ini + [ultima + 1]
folhas = [round(sum(H.get(x, 16) for x in range(a, b))) for a, b in zip(lim, lim[1:])]
assert not FALHAS, ('texto que nao cabe na celula', sorted(set(FALHAS))[:20])
assert max(folhas) <= LIM, ('folha maior que a pagina', folhas)
assert min(o['size'] for o in ops if 'size' in o) >= 10
json.dump(dict(ops=ops, ultima=ultima, aba='ORGANOGRAMA impressao', fonte='Verdana', zoom=100, quebras=quebras, folhas=folhas,
               colunas=dict(primeira=1, ultima=CF + 1, largura=1.58, margem=0.5, px={str(c): px(c) for c in range(C0, CF + 1) if px(c) != 16})), open('org_impressao.json', 'w', encoding='utf8'), ensure_ascii=False)
print('operacoes', len(ops), 'ultima', ultima, 'quebras', quebras, 'alturas das folhas (pt)', folhas)
