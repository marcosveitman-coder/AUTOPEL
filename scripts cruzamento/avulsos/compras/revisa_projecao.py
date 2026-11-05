"""Revisa a projecao de compra de 3 meses (colunas "Compra mes vigente / proximo mes / mes subsequente") com os numeros
atuais da aba Geral do "Relatorio Revenda atualizado.xlsx". So le; nao grava nada.

    python revisa_projecao.py PLANILHA.xlsx [AAAA-MM-DD]

1. Refaz em Python o saldo (E - F) e os 3 meses e compara com a planilha (deve dar 0 divergencia).
2. Mostra onde o racional depende de hipoteses:
   - Comprado com entrega vencida ou sem data, contado como se fosse chegar;
   - SKUs que so ficam no ponto por causa do Comprado (fisico abaixo do ponto) e rupturas que o Comprado ja cobre;
   - composicao dos meses (SKUs ja no ponto x SKUs acima do ponto que vao cair) e PontoPedido manual;
   - cenarios: PP = 2 x MedContr de hoje e saldo sem os pedidos vencidos ou sem data.
"""
import collections
import datetime as dt
import sys

import openpyxl


def num(x):
    return float(x) if isinstance(x, (int, float)) else 0.0


def data(s):
    try:
        return dt.datetime.strptime(str(s or '').strip(), '%d/%m/%Y').date()
    except ValueError:
        return None


def plano(E, s, I):
    """Regra da planilha: mes vigente volta o saldo ao ponto; depois 1 MedContr por mes. PP 0 = so o ja vendido."""
    at = max(0.0, -s) if E == 0 else max(0.0, E - s)
    av = 0.0 if E == 0 else max(0.0, E - (s + at - I))
    ax = 0.0 if E == 0 else max(0.0, E - (s + at + av - 2 * I))
    return at, av, ax


def main(arq, hoje):
    wb = openpyxl.load_workbook(arq, read_only=True, data_only=True)
    ws = wb[next(n for n in wb.sheetnames if n.startswith('Geral (arq.') and n.endswith(')'))]
    rows = ws.iter_rows(values_only=True)
    next(rows)
    H = [str(x).strip() if x is not None else '' for x in next(rows)]
    ix = {h: i for i, h in enumerate(H) if h}
    D = [r for r in rows if r and r[ix['Codigo']]]
    wb.close()
    v = lambda r, c: num(r[ix[c]])
    fim_mes = (hoje.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)

    dif = collections.Counter()
    A, C = collections.defaultdict(float), collections.Counter()
    cen = collections.defaultdict(lambda: [0.0, 0.0, 0.0])
    for r in D:
        E, F, I, O, P, R, G = [v(r, c) for c in ('PontoPedido', 'Comprar.', 'MedContr', 'Disponivel', 'VendProt', 'Comprado', 'Ult.Preco')]
        s, fis = E - F, O - P
        base = plano(E, s, I)
        for c, x in zip(('Compra mês vigente QTD', 'Compra próximo mês QTD', 'Compra mês subsequente QTD'), base):
            if abs(v(r, c) - x) > 1e-6:
                dif[c] += 1
        if abs(v(r, 'Saldo hoje (Disp-VendProt+Comprado)') - s) > 1e-6:
            dif['saldo'] += 1
        if abs(F - (E - (O - P + R))) > 0.5:
            dif['Comprar diferente de E-(O-P+R)'] += 1
        de = data(r[ix['DT_ENTREGA']])
        vencido = R > 0 and (de is None or de < hoje)
        if R > 0:
            k = 'sem data' if de is None else 'vencido' if de < hoje else 'ate o fim do mes' if de <= fim_mes else 'depois'
            A['Comprado ' + k] += R * G; C['Comprado ' + k] += 1
        if fis < 0:
            C['ruptura'] += 1
            if base[0] == 0:
                C['ruptura que o Comprado ja cobre (cobrar PC)'] += 1
            if E == 0:
                C['ruptura com PontoPedido 0'] += 1
        if E > 0 and fis < E <= s:
            C['no ponto so por causa do Comprado'] += 1
        if E > 0 and fis < I and base[0] == 0:
            C['fisico < 1 MedContr e sem compra no mes'] += 1
        if E > 0 and abs(E - 2 * I) > 1:
            C['PontoPedido diferente de 2 x MedContr'] += 1
        if E == 0 and I > 0:
            C['PontoPedido 0 com MedContr'] += 1
        if E > 0:
            k = 'ja no ponto' if s <= E else 'acima do ponto'
            A[f'proximo mes: {k}'] += base[1] * G
            A[f'mes subsequente: {k}'] += base[2] * G
        E2 = 2 * I if E > 0 and abs(E - 2 * I) > 1 else E
        for nome, p in (('regra atual', base), ('PP = 2 x MedContr de hoje', plano(E2, s, I)),
                        ('sem PC vencido ou sem data', plano(E, s - R if vencido else s, I))):
            for j in range(3):
                cen[nome][j] += p[j] * G
    print(f'{len(D)} SKUs; divergencias Python x planilha: {dict(dif) or 0}')
    for k in sorted(A):
        print(f'  {k:40s} R$ {A[k] / 1e6:6.2f} mi  ({C.get(k, "")} SKUs)')
    for k in sorted(C):
        if k not in A:
            print(f'  {k:40s} {C[k]} SKUs')
    print('Cenarios (mes vigente / proximo / subsequente, R$ mi):')
    for k, t in cen.items():
        print(f'  {k:30s} ' + ' / '.join(f'{x / 1e6:.2f}' for x in t) + f' = {sum(t) / 1e6:.2f}')


if __name__ == '__main__':
    main(sys.argv[1], dt.date.fromisoformat(sys.argv[2]) if len(sys.argv) > 2 else dt.date.today())
