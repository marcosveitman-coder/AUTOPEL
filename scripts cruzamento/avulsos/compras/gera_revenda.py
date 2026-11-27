"""Gera o "Relatorio de Revenda" so a partir do EstoqueNecessidade (sem o Power BI).

    python gera_revenda.py "EstoqueNecessidade (N).xlsx" "SAIDA.xlsx" [--modelo "Relatorio de Revenda dd-mm-aaaa.xlsx"]

Mesmo layout do relatorio que o Power BI exporta: aba Geral (24 colunas) e uma aba por comprador, com subtotal por marca
e "Total Geral". Regras (conferidas em 25/11/2026, ver reproduz_revenda.py e a sessao de 25/11):
- Universo: filial 05, ATIVO, Tipo RV ou PA; valores = soma das linhas da filial 05 do codigo.
- Copia: Comprador (Usuario, 10 letras), Descricao, Marca, Grupo, DT_ENTREGA, PEDIDO, PontoPedido, Fisico, Vlr Estoque,
  Reserva, Disponivel, VendProt, Comprado, Comprar. (TT_COMPRAR).
- MedContr = teto da media de MES_01..03; Ult.Preco e Prc_Compra. = teto a 2 casas; Valor Total = Comprar. x Ult.Preco;
  Ruptura = MIN(0; Disponivel - VendProt); EstDias = MAX(0; Disponivel - VendProt) / MedContr x 30.
- **MedVenda e Total Media (decisao de 26/11/2026):** a MedVenda do Power BI vem de colunas fixas e desatualizadas (Mes -1..-3).
  Aqui MedVenda = MedContr (venda media dos 3 ultimos meses fechados, do TOTVS; vazia quando 0) e
  Total Media = MedContr x Ult.Preco.
- Ordem: marca (sem maiusculas, acentos como desempate, hifen ignorado), depois codigo. Abas por comprador: quem tem 5 ou
  mais SKUs e nao e INATIVO/OCIOSO; o nome da aba e o primeiro nome do comprador.
Com --modelo, confere contra o relatorio do Power BI (todas as colunas menos MedVenda e Total Media).
"""
import collections
import datetime as dt
import math
import re
import sys
import unicodedata

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

CAB = ['Comprador', 'Codigo      ', 'Descricao   ', 'Marca       ', ' PontoPedido', 'Comprar.', ' Ult.Preco', 'Valor Total ',
       ' MedContr', 'MedVenda', ' Total Media', ' Vlr Estoque', ' Fisico', ' Reserva', ' Disponivel', ' VendProt',
       ' Ruptura', ' Comprado', 'DT_ENTREGA', ' Prc_Compra.', 'ULT_COMPRA', 'PEDIDO', ' EstDias', 'Grupo       ']
RS = '_-"R$"\\ * #,##0.00_-;\\-"R$"\\ * #,##0.00_-;_-"R$"\\ * "-"??_-;_-@_-'
FORMATOS = {7: RS, 8: RS, 11: RS, 12: RS, 13: '0'}
LARG = {1: 11, 2: 9, 3: 50, 4: 18, 5: 9, 6: 9, 7: 11.5, 8: 15.5, 9: 9, 10: 9, 11: 15, 12: 16, 13: 8, 14: 8, 15: 10, 16: 9,
        17: 8, 18: 9, 19: 12, 20: 10, 21: 12, 22: 8, 23: 10, 24: 8}
s = lambda x: str(x).strip() if x is not None else ''
n = lambda x: float(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else 0.0
teto2 = lambda x: math.ceil(round(x * 100, 6)) / 100


def data(x):
    try:
        return dt.datetime.strptime(s(x), '%d/%m/%Y').date()
    except ValueError:
        return None


def chave_marca(m):
    """Ordem do Power BI (cultura pt-BR): ignora maiusculas e hifen, acento so desempata, o espaco do inicio conta."""
    sem = ''.join(c for c in unicodedata.normalize('NFD', m) if unicodedata.category(c) != 'Mn')
    return (re.sub(r"[-]", '', sem.casefold()), m.casefold())


def le_estoque(arq):
    wb = openpyxl.load_workbook(arq, read_only=True, data_only=True)
    rows = list(wb['Listagem'].iter_rows(values_only=True))
    wb.close()
    ie = {str(h).strip(): i for i, h in enumerate(rows[0]) if h is not None}
    return ie, [r for r in rows[1:] if r and any(x not in (None, '') for x in r)]


def monta(arq):
    """Devolve as linhas da Geral (listas de 24 valores), ja na ordem do relatorio."""
    ie, E = le_estoque(arq)
    linhas = collections.defaultdict(list)
    for r in E:
        if (s(r[ie['Filial']]) == '05' and s(r[ie['ATIVO']]) == 'ATIVO' and s(r[ie['Tipo']]) in ('RV', 'PA')
                and s(r[ie['Armazem']]) in ('40', '10', '01')):
            linhas[s(r[ie['Codigo']])].append(r)
    linhas = {c: rs for c, rs in linhas.items() if any(s(r[ie['Armazem']]) != '01' for r in rs)}   # so armazem 01 fica fora
    # a mesma marca com maiusculas diferentes aparece com uma grafia so (a mais frequente; empate: a de maiusculas)
    graf = collections.defaultdict(collections.Counter)
    for rs in linhas.values():
        m = str(rs[0][ie['Marca']]).rstrip()
        graf[m.casefold()][m] += 1
    canon = {k: sorted(v.items(), key=lambda kv: (-kv[1], kv[0]))[0][0] for k, v in graf.items()}
    out = []
    for c, rs in linhas.items():
        S = lambda col: sum(n(r[ie[col]]) for r in rs)
        p = rs[0]
        mc = math.ceil(round((S('MES_01') + S('MES_02') + S('MES_03')) / 3, 9))
        preco = teto2(n(p[ie['Ult. Preco']]))
        disp, vp, comprar = S('EST_DISP'), S('Qtd.Ped.Vend'), S('TT_COMPRAR')
        ds = [d for d in (data(r[ie['ULT_COMPRA']]) for r in rs) if d]
        out.append([s(p[ie['Usuario']])[:10], c, str(p[ie['Descricao']]).rstrip(' \t'), canon[str(p[ie['Marca']]).rstrip().casefold()],
                    S('ESTQ_MININ'), comprar, preco, comprar * preco, mc, mc if mc > 0 else None, mc * preco, S('VLR_ESTQ'),
                    S('EST_FISICO'), S('Qtd. Reserva'), disp, vp, min(0, disp - vp), S('PED_EMITID'), s(p[ie['DT_ENTREGA']]),
                    teto2(n(p[ie['PRC_COMPRA']])), max(ds).strftime('%d/%m/%Y') if ds else '', s(p[ie['PEDIDO']]),
                    (max(0, disp - vp) / mc * 30 if mc > 0 else 0), s(p[ie['Grupo']])])
    out.sort(key=lambda r: (chave_marca(r[3]), r[1]))
    return out


def escreve(linhas, saida):
    wb = openpyxl.Workbook()
    borda = Border(*(Side(style='thin'),) * 4)
    negrito = Font(name='Calibri', size=10, bold=True)
    normal = Font(name='Calibri', size=10)

    def folha(ws, dados):
        ws.append(CAB)
        for r in dados:
            ws.append(r)
        for j in range(1, 25):
            ws.column_dimensions[get_column_letter(j)].width = LARG[j]
        for row in ws.iter_rows():
            for c in row:
                c.border = borda
                c.font = negrito if c.row == 1 else normal
                if c.row > 1 and c.column in FORMATOS:
                    c.number_format = FORMATOS[c.column]
        ws.freeze_panes = 'A2'

    ws = wb.active
    ws.title = 'Geral'
    folha(ws, linhas)
    ws.auto_filter.ref = f'A1:X{len(linhas) + 1}'
    por = collections.defaultdict(list)
    for r in linhas:
        por[r[0]].append(r)
    for comp, rs in sorted(por.items(), key=lambda kv: kv[0].split()[0].casefold()):
        if comp.upper().startswith(('INATIVO', 'OCIOSO')) or len(rs) < 5:
            continue
        w = wb.create_sheet(comp.split()[0].capitalize())
        dados, tot = [], [0.0, 0.0, 0.0]
        grupos = collections.OrderedDict()
        for r in rs:
            grupos.setdefault(r[3], []).append(r)
        for marca, g in grupos.items():
            dados += g
            v, m, e = sum(x[7] for x in g), sum(x[10] for x in g), sum(x[11] for x in g)
            dados.append([None, None, None, f'{marca} Total', None, None, None, v, None, None, m, e] + [None] * 12)
            tot = [tot[0] + v, tot[1] + m, tot[2] + e]
        dados.append([None, None, None, 'Total Geral', None, None, None, tot[0], None, None, tot[1], tot[2]] + [None] * 12)
        folha(w, dados)
        for row in w.iter_rows(min_row=2):
            if row[0].value is None:
                for c in row:
                    c.font = negrito
    wb.save(saida)


def confere(linhas, modelo):
    wb = openpyxl.load_workbook(modelo, read_only=True, data_only=True)
    M = [list(r) for r in wb['Geral'].iter_rows(values_only=True)][1:]
    wb.close()
    mm = {s(r[1]): r for r in M}
    gg = {r[1]: r for r in linhas}
    print(f'modelo {len(M)} linhas | gerado {len(linhas)} | so no modelo {len(set(mm) - set(gg))} | so no gerado {len(set(gg) - set(mm))}')
    print('mesma ordem das linhas:', sum(1 for a, b in zip(M, linhas) if s(a[1]) == b[1]), 'de', len(M))
    dif = collections.Counter()
    ex = collections.defaultdict(list)
    for c in set(mm) & set(gg):
        for j in range(24):
            if j in (9, 10):
                continue            # MedVenda e Total Media mudam por decisao
            a, b = mm[c][j], gg[c][j]
            if isinstance(a, (int, float)) or isinstance(b, (int, float)):
                bom = abs(n(a) - n(b)) <= 0.01
            else:
                bom = (str(a or '') == str(b or '')) if j in (2, 3) else s(a) == s(b)
            if not bom:
                dif[CAB[j].strip()] += 1
                if len(ex[CAB[j].strip()]) < 3:
                    ex[CAB[j].strip()].append((c, a, b))
    print('colunas diferentes (sem MedVenda e Total Media):', dict(dif) or 'nenhuma')
    for k, v in ex.items():
        print('  ', k, v)


if __name__ == '__main__':
    a = sys.argv[1:]
    modelo = a[a.index('--modelo') + 1] if '--modelo' in a else None
    linhas = monta(a[0])
    escreve(linhas, a[1])
    print(f'gravado {a[1]}: {len(linhas)} SKUs')
    if modelo:
        confere(linhas, modelo)
