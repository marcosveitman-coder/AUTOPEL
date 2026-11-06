"""Conta, em Python e com as mesmas regras da macro (vba_DetalheMod.txt), quantas linhas a aba DETALHE deve ter para cada
numero clicavel da CONFIG DETALHE (06/11/2026). Serve para conferir o teste da macro (org_detalhe_teste.ps1).
Carteira: CSV do TOTVS (um cliente por vendedor + codigo, nota mais recente, ativo <= 290 dias, perdendo <= 365).
Potencial e tabela por segmento: o numero da propria celula (org_impressao.json). Pessoas: linhas da equipe.
Uso (na pasta dados): AUTOPEL_ULTFAT_CSV=<CSV> python ../org_detalhe_esperado.py <resultado da macro .tsv> [AAAA-MM-DD]"""
import os, sys, json, datetime as dt, collections

HOJE = dt.date.fromisoformat(sys.argv[2]) if len(sys.argv) > 2 else dt.date.today()
cfg = [l.rstrip('\r\n').split('\t') for l in open('config_detalhe.tsv', encoding='utf-8-sig')]
val = {l[0]: l[1] for l in cfg[:30] if len(l) > 1}
DA, DP = int(val['DIAS ATIVO']), int(val['DIAS PERDENDO'])
def secao(t):
    i = next(k for k, l in enumerate(cfg) if l[0] == t) + 2
    out = []
    while i < len(cfg) and cfg[i][0]:
        out.append(cfg[i]); i += 1
    return out
PES = {l[0]: l for l in secao('PESSOAS')}
CEL = secao('CELULAS')
seg_vend = {}
for l in PES.values():
    if l[3] == 'VEND' and l[4]:
        for n in l[4].split(';'):
            seg_vend[n] = set(l[5].split(';')) if l[5] else set()

def norm(s, nd):
    s = s.strip(' ')
    return s.zfill(nd) if s and len(s) <= 9 and s.isdigit() else s
def data(s):
    s = s.strip(' ')
    if len(s) >= 10 and s[2] == '/' and s[5] == '/':
        try:
            return dt.date(int(s[6:10]), int(s[3:5]), int(s[0:2]))
        except ValueError:
            return None
    return None

txt = open(os.environ['AUTOPEL_ULTFAT_CSV'], encoding='utf-8').read()
U = []
for lin in txt.split('\n')[2:]:
    if len(lin) > 20:
        c = lin.split(';')
        if len(c) >= 25:
            U.append(dict(cod=norm(c[0], 6), loja=norm(c[1], 4), seg=c[13].strip(' '), vend=c[15].strip(' '), dir=c[21].strip(' '), dt=data(c[24])))

cache = {}
def carteira(quem):
    if quem in cache:
        return cache[quem]
    l = PES[quem]; esc, alvo = l[3], set(l[4].split(';')) if l[4] else set()
    ag = {}
    for u in U:
        k = u['dir'] if esc == 'DIR' else u['vend']
        if not k or k not in alvo:
            continue
        ch = (u['vend'], u['cod'])
        a = ag.get(ch)
        if a is None:
            a = ag[ch] = dict(vend=u['vend'], seg=u['seg'], dt=None)
        if u['dt'] and (a['dt'] is None or u['dt'] > a['dt']):
            a['dt'] = u['dt']
    for a in ag.values():
        dias = (HOJE - a['dt']).days if a['dt'] else None
        a['st'] = 'T' if dias is None else ('A' if dias <= DA else ('P' if dias <= DP else 'T'))
        a['no'] = a['seg'] in seg_vend.get(a['vend'], set()) if a['vend'] in seg_vend else False
    cache[quem] = (esc, list(ag.values()))
    return cache[quem]

IMP = json.load(open('org_impressao.json', encoding='utf8'))
from openpyxl.utils import get_column_letter as Lt
n_cel = {f"{Lt(o['c'])}{o['r']}": o['det']['n'] for o in IMP['ops'] if o.get('det')}
esp = {}
for l in CEL:
    cel, tipo, quem, seg, st = l[0], l[1], l[2], l[3], l[4]
    if tipo == 'C':
        esc, A = carteira(quem)
        f = (seg in ('NO', 'FORA')) and esc != 'DIR'
        esp[cel] = sum(1 for a in A if (st == 'TODOS' or a['st'] == st) and (not f or (a['no'] if seg == 'NO' else not a['no'])))
    elif tipo == 'E':
        esp[cel] = sum(1 for p in PES.values() if p[6] == quem)
    else:
        esp[cel] = n_cel[cel]
res = {}
for l in open(sys.argv[1], encoding='utf-8-sig'):
    p = l.rstrip('\r\n').split('\t')
    if p[0].startswith('#'):
        print(' '.join(p)); continue
    res[p[0]] = p
ok = dif = 0
crm = collections.Counter()
for l in CEL:
    cel = l[0]
    r = res.get(cel)
    if not r:
        print('SEM RESULTADO', cel); dif += 1; continue
    n_macro = int(r[1].split('|')[1]) if r[1].startswith('OK|') else None
    n_aba = int(r[2]) if r[2].lstrip('-').isdigit() else None
    if n_macro == esp[cel] and (n_aba == n_macro or (n_macro == 0 and n_aba in (0, 1))):
        ok += 1
    else:
        dif += 1; print('DIFERENTE', l[:6], '| macro', r[1:], '| esperado', esp[cel])
    if l[1] == 'C':
        crm[l[2]] = crm[l[2]]
print(f'células: {len(CEL)} | iguais ao esperado: {ok} | diferentes: {dif}')
