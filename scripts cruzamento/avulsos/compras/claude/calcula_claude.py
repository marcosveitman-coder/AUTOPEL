"""Metodologia CLAUDE de compra (referencia em Python; a planilha usa as mesmas regras em formulas).

    python calcula_claude.py RELATORIO.xlsx params.pkl SAIDA.pkl [--hoje 2026-10-26]

Para cada SKU da aba Geral (Table1):
- previsao do mes m = MedContr x S(m) / media de S nos 3 meses que formam a MedContr (os 3 ultimos meses completos);
- protecao T = LT do fornecedor + R (revisao: 7 dias para a curva A, 15 para a B, 30 para a C);
- estoque de seguranca SS(m) = z(ABC) x raiz( T/30 x (CV x previsao)^2 + (previsao/30 x desvio do LT)^2 );
- estoque alvo(m) = previsao x T/30 + SS(m), no maximo 4 meses da previsao;
- saldo hoje = PontoPedido - Comprar do TOTVS (= Disponivel - VendProt + Comprado);
- compra no mes vigente = alvo(0) - saldo; saldo no inicio do proximo mes = saldo + compra - previsao x dias que
  faltam no mes / dias do mes; compra no proximo = alvo(1) - esse saldo; no subsequente, idem com 1 mes de venda;
- arredonda para cima no multiplo de compra; nunca negativo;
- PontoPedido 0 (regra do usuario: sem estoque): so compra no mes vigente o que ja foi vendido sem estoque;
- comprador INATIVO: zero.
"""
import calendar
import datetime as dt
import math
import pickle
import sys

import openpyxl

R_DIAS = {'A': 7, 'B': 15, 'C': 30}          # revisao: semanal (A), quinzenal (B), mensal (C) - padrao desde 26/10
Z = {'A': 1.6448536269514722, 'B': 1.2815515655446004, 'C': 0.8416212335729143}   # NORM.S.INV de 95% / 90% / 80%
COB_MAX = 4                                   # estoque alvo no maximo 4 meses da previsao
SLT_PADRAO = 5.9                              # mesmos padroes da aba param CLAUDE
LT_PADRAO = 8.0
CV_PADRAO = 0.5


def num(x):
    try:
        return float(x or 0)
    except (TypeError, ValueError):
        return 0.0


def arred(x, mult):
    if x <= 1e-9:
        return 0
    m = max(1, mult)
    return math.ceil(round(x / m, 9)) * m


def calcula(linha, p, hoje):
    """linha: dict com E, F, I, G, T, comprador; p: parametros do SKU (ou None)."""
    E, Fc, I = num(linha['E']), num(linha['F']), max(0.0, num(linha['I']))
    preco = num(linha['T']) if num(linha['T']) > 0 else num(linha['G'])
    S = p['S'] if p else [1.0] * 12
    meses = [((hoje.month - 1 + k) % 12) for k in range(3)]                 # indice 0-11 de m0, m1, m2
    janela = [((hoje.month - 1 - k) % 12) for k in (1, 2, 3)]               # 3 ultimos meses completos
    base = sum(S[i] for i in janela) / 3
    prev = [I * S[i] / base for i in meses]
    lt = p['lt'] if p else LT_PADRAO
    slt = p['slt'] if p else SLT_PADRAO
    cv = p['cv_usado'] if p else CV_PADRAO
    z = Z.get(p['abc'] if p else 'C', 1.282)
    t = lt + R_DIAS.get(p['abc'] if p else 'C', 30)
    ss = [z * math.sqrt(t / 30 * (cv * f) ** 2 + (f / 30 * slt) ** 2) for f in prev]
    alvo = [min(f * t / 30 + s, f * COB_MAX) for f, s in zip(prev, ss)]
    mult = p['mult'] if p else 1
    saldo = E - Fc
    dm = calendar.monthrange(hoje.year, hoje.month)[1]
    dr = dm - hoje.day
    inativo = str(linha.get('comprador') or '').upper().startswith('INATIVO')
    if inativo:
        q = [0, 0, 0]
    elif E == 0:
        q = [arred(max(0.0, -saldo), 1), 0, 0]
    else:
        q0 = arred(max(0.0, alvo[0] - saldo), mult)
        s1 = saldo + q0 - prev[0] * dr / dm
        q1 = arred(max(0.0, alvo[1] - s1), mult)
        s2 = s1 + q1 - prev[1]
        q2 = arred(max(0.0, alvo[2] - s2), mult)
        q = [q0, q1, q2]
    return {'q': q, 'rs': [x * preco for x in q], 'prev': prev, 'alvo': alvo, 'ss': ss[0], 'lt': lt, 'preco': preco,
            'cobertura': alvo[0] / (prev[0] / 30) if prev[0] > 0 else None, 'saldo': saldo}


def le_geral(arq):
    wb = openpyxl.load_workbook(arq, read_only=True, data_only=True)
    ws = next(w for w in wb.worksheets if w.title.startswith('Geral (arq') and 'O.N.T.E.M' not in w.title)
    linhas = list(ws.iter_rows(values_only=True))
    wb.close()
    cab = {str(c).strip(): j for j, c in enumerate(linhas[1]) if c is not None}
    out = []
    for i, r in enumerate(linhas[2:], start=3):
        cod = r[cab['Codigo']]
        if not cod:
            continue
        reg = {'linha': i, 'cod': str(cod).strip(), 'comprador': r[cab['Comprador']], 'E': r[cab['PontoPedido']],
               'F': r[cab['Comprar.']], 'I': r[cab['MedContr']], 'G': r[cab['Ult.Preco']], 'T': r[cab['Prc_Compra.']],
               'estoque_rs': r[cab['Vlr Estoque']], 'grupo': r[cab['Grupo']], 'marca': r[cab['Marca']],
               'desc': r[cab['Descricao']], 'ruptura': r[cab['Ruptura']], 'O': r[cab['Disponivel']]}
        for k in ('Compra mês vigente QTD', 'Compra próximo mês QTD', 'Compra mês subsequente QTD',
                  'Compra mês vigente R$', 'Compra próximo mês R$', 'Compra mês subsequente R$'):
            reg[k] = r[cab[k]] if k in cab else None
        out.append(reg)
    return out


def main(arq, par, saida, hoje):
    P = pickle.load(open(par, 'rb'))['params']
    linhas = le_geral(arq)
    for l in linhas:
        l['claude'] = calcula(l, P.get(l['cod']), hoje)
        l['param'] = P.get(l['cod'])
    pickle.dump(linhas, open(saida, 'wb'))
    print('SKUs:', len(linhas), '| hoje', hoje)


if __name__ == '__main__':
    a = sys.argv[1:]
    h = a[a.index('--hoje') + 1] if '--hoje' in a else dt.date.today().isoformat()
    main(a[0], a[1], a[2], dt.date.fromisoformat(h))
