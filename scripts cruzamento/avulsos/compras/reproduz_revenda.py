"""Refaz o "Relatorio de Revenda" (saida do Power BI) so com o EstoqueNecessidade e compara com o relatorio de verdade.

    python reproduz_revenda.py "EstoqueNecessidade (N).xlsx" "Relatorio de Revenda dd-mm-aaaa.xlsx"

Os 3 relatorios do TOTVS que entram no Power BI sao EstoqueNecessidade, HistoricodeCompras e OrdemCompra. Este script
mostra que a aba Geral da saida sai so do EstoqueNecessidade (HistoricodeCompras e OrdemCompra nao entram nas 24 colunas),
e quais colunas nao saem das 3 entradas.

Regras (conferidas contra a saida de 25/11/2026, 4.266 SKUs):
- Universo: codigos com linha na filial 05, armazem 40 (ou 10 nos PA), ATIVO e Tipo RV ou PA. Os valores somam as linhas
  da filial 05 do codigo (65 codigos tem mais de uma linha).
- Comprador = Usuario (10 primeiras letras); Marca, Descricao, Grupo, DT_ENTREGA e PEDIDO = copia.
- PontoPedido = ESTQ_MININ; Fisico, Vlr Estoque, Reserva, Disponivel = EST_FISICO, VLR_ESTQ, Qtd. Reserva, EST_DISP;
  VendProt = Qtd.Ped.Vend; Comprado = PED_EMITID; Comprar. = TT_COMPRAR.
- MedContr = TETO da media de MES_01..03. Ult.Preco e Prc_Compra. = TETO a 2 casas de "Ult. Preco" e PRC_COMPRA.
- Valor Total = Comprar. x Ult.Preco. Ruptura = MIN(0; Disponivel - VendProt).
  EstDias = MAX(0; Disponivel - VendProt) / MedContr x 30 (0 sem MedContr). ULT_COMPRA = a mais recente das linhas.
- NAO sai das entradas: MedVenda (e Total Media = MedVenda x Ult.Preco). Vem de outra fonte (vendas).
"""
import collections
import datetime as dt
import math
import sys

import openpyxl


def le(arq, aba):
    wb = openpyxl.load_workbook(arq, read_only=True, data_only=True)
    rows = list(wb[aba].iter_rows(values_only=True))
    wb.close()
    H = [str(x).strip() if x is not None else '' for x in rows[0]]
    return {h: i for i, h in enumerate(H) if h}, [r for r in rows[1:] if r and any(x not in (None, '') for x in r)]


s = lambda x: str(x).strip() if x is not None else ''
n = lambda x: float(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else 0.0


def data(x):
    try:
        return dt.datetime.strptime(s(x), '%d/%m/%Y').date()
    except ValueError:
        return None


def main(est, saida):
    ie, E = le(est, 'Listagem')
    io, O = le(saida, 'Geral')
    linhas = collections.defaultdict(list)
    for r in E:
        if (s(r[ie['Filial']]) == '05' and s(r[ie['ATIVO']]) == 'ATIVO' and s(r[ie['Tipo']]) in ('RV', 'PA')
                and s(r[ie['Armazem']]) in ('40', '10', '01')):
            linhas[s(r[ie['Codigo']])].append(r)
    # armazem 01 so conta se o codigo tem outra linha (as 2 excecoes tem so a do armazem 01)
    linhas = {c: rs for c, rs in linhas.items() if any(s(r[ie['Armazem']]) != '01' for r in rs)}
    Om = {s(r[io['Codigo']]): r for r in O}
    print(f'saida {len(Om)} SKUs | refeito {len(linhas)} | so na saida {len(set(Om) - set(linhas))} | so no refeito {len(set(linhas) - set(Om))}')
    S = lambda c, col: sum(n(r[ie[col]]) for r in linhas[c])
    ok, ruim, ex = collections.Counter(), collections.Counter(), collections.defaultdict(list)

    def cmp(col, c, o, v, tol=0.005):
        bom = (abs(n(o) - n(v)) <= tol) if isinstance(v, (int, float)) and not isinstance(v, bool) else s(o) == s(v)
        (ok if bom else ruim)[col] += 1
        if not bom and len(ex[col]) < 3:
            ex[col].append((c, o, v))

    for c, o in Om.items():
        if c not in linhas:
            continue
        g = lambda k: o[io[k]]
        p = linhas[c][0]
        mc = math.ceil(round((S(c, 'MES_01') + S(c, 'MES_02') + S(c, 'MES_03')) / 3, 9))
        preco = math.ceil(round(n(p[ie['Ult. Preco']]) * 100, 6)) / 100
        disp, vp = S(c, 'EST_DISP'), S(c, 'Qtd.Ped.Vend')
        comprar = S(c, 'TT_COMPRAR')
        ds = [data(r[ie['ULT_COMPRA']]) for r in linhas[c]]
        ds = [d for d in ds if d]
        for col, v in [('Comprador', s(p[ie['Usuario']])[:10]), ('Descricao', s(p[ie['Descricao']])), ('Grupo', s(p[ie['Grupo']])),
                       ('DT_ENTREGA', s(p[ie['DT_ENTREGA']])), ('PEDIDO', s(p[ie['PEDIDO']])), (' Prc_Compra.', math.ceil(round(n(p[ie['PRC_COMPRA']]) * 100, 6)) / 100)]:
            k = col if col in io else col.strip()
            cmp(k, c, o[io[k]] if k in io else o[io[col]], v)
        cmp('Marca', c, s(g('Marca')).upper(), s(p[ie['Marca']]).upper())
        for col, ec in [('PontoPedido', 'ESTQ_MININ'), ('Fisico', 'EST_FISICO'), ('Vlr Estoque', 'VLR_ESTQ'), ('Reserva', 'Qtd. Reserva'),
                        ('Disponivel', 'EST_DISP'), ('VendProt', 'Qtd.Ped.Vend'), ('Comprado', 'PED_EMITID'), ('Comprar.', 'TT_COMPRAR')]:
            cmp(col, c, g(col), S(c, ec), 0.01)
        cmp('MedContr', c, g('MedContr'), mc)
        cmp('Ult.Preco', c, g('Ult.Preco'), preco)
        cmp('Valor Total', c, g('Valor Total'), comprar * preco, 0.01)
        cmp('Ruptura', c, g('Ruptura'), min(0, disp - vp))
        cmp('EstDias', c, g('EstDias'), max(0, disp - vp) / mc * 30 if mc > 0 else 0, 0.001)
        cmp('ULT_COMPRA', c, g('ULT_COMPRA'), max(ds).strftime('%d/%m/%Y') if ds else '01/01/1900')
        cmp('Total Media (= MedVenda x preco, MedVenda vem de fora)', c, g('Total Media'), n(g('MedVenda')) * preco, 0.01)
    print('coluna: bateu / nao bateu')
    for col in list(ok) + [k for k in ruim if k not in ok]:
        print(f'  {col:58s} {ok[col]:5d} / {ruim[col]:3d}', ex[col] if ruim[col] else '')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
