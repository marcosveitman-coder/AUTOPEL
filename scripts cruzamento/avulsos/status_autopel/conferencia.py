"""Planilha de conferência do STATUS AUTOPEL (16/11/2026), para o usuário aprovar ANTES de gravar no MAPA.
Abas: LEIA (regras e decisões), RESUMO (antes x agora por aba), REDES (uma linha por linha de aba, com a prova),
FARMÁCIAS AVULSAS (só as que mudaram), FARMACIA VF LOJAS (colunas H, I, J).
Uso: python conferencia.py "<pasta de trabalho>" "<MAPA.xlsm>" "<saida.xlsx>" [AAAA-MM-DD] """
import sys, os, pickle, datetime as dt, collections as C
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L

PASTA, MAPA, SAIDA = sys.argv[1:4]
HOJE = dt.date.fromisoformat(sys.argv[4]) if len(sys.argv) > 4 else dt.date.today()
E = pickle.load(open(os.path.join(PASTA, 'entidades.pkl'), 'rb'))
R = pickle.load(open(os.path.join(PASTA, 'resultados.pkl'), 'rb'))
FRACO = {'NOME', 'EMAIL', 'PRODUTO'}
APOIO = {'OUTROS FINAL', 'SUPERMERCADOS (reuniao)'}
AZUL, CINZA = PatternFill('solid', fgColor='1F3A5F'), PatternFill('solid', fgColor='EDEDED')
VERDE, AMAR, VERM = PatternFill('solid', fgColor='E2EFDA'), PatternFill('solid', fgColor='FFF2CC'), PatternFill('solid', fgColor='FCE4D6')
COR = {'ATIVO': VERDE, 'INATIVO': AMAR, 'NÃO É CLIENTE': VERM}
VIA = {'CNPJ': 'CNPJ (raiz)', 'CODIGO': 'código de cliente (aba IRTT)', 'RAZAO': 'razão social igual', 'ASSOCIADA': 'associada da rede',
       'GRUPO MAPA': 'grupo do cadastro (mapa curado)', 'GRUPO ERP': 'grupo do ERP (aba IRTT)', 'GRUPO': 'grupo do cliente no cadastro',
       'EXATO': 'nome fantasia igual', 'NOME': 'nome do cliente', 'EMAIL': 'domínio do e-mail', 'PRODUTO': 'produto personalizado'}


def norm_st(a):
    a = str(a or '').strip().upper()
    return {'A TRABALHAR': 'NÃO ATIVO (A TRABALHAR)', 'NÃO ATENDIDO': 'NÃO É CLIENTE', 'CADASTRADO SEM COMPRA': 'INATIVO',
            'PERDENDO': 'ATIVO'}.get(a, a)


def cab(ws, cols, larg):
    ws.append(cols)
    for j, w in enumerate(larg, 1):
        c = ws.cell(1, j); c.font = Font(bold=True, color='FFFFFF'); c.fill = AZUL
        c.alignment = Alignment(wrap_text=True, vertical='center'); ws.column_dimensions[L(j)].width = w
    ws.freeze_panes = 'A2'; ws.auto_filter.ref = f'A1:{L(len(cols))}1'


def main():
    wb = openpyxl.Workbook()
    # ---------------- LEIA
    ws = wb.active; ws.title = 'LEIA'
    txt = [
        f'CONFERÊNCIA DO STATUS AUTOPEL – {HOJE:%d/%m/%Y} (para aprovar antes de gravar no MAPA MERCADO V2)',
        '',
        'FONTES (pasta Downloads\\12.11.2026): faturamento 2020 a 2026 (até 15/11/2026), ultimo fat.xlsx (última nota de cada',
        'cliente, até 12/11/2026), CADASTRO DE CLIENTES.RLT (vendedor do cadastro) e CADASTRO DE PRODUTO.',
        '',
        'REGRA DE STATUS: ATIVO = última compra há menos de 360 dias (contados de hoje); INATIVO = comprou antes disso ou só',
        'tem cadastro sem nota; NÃO É CLIENTE = nenhum vínculo.',
        '',
        'A COMPRA DE QUALQUER CLIENTE LIGADO À REDE CONTA (combinado em 16/11): a própria rede, franqueado, associada,',
        'distribuidor que atende a rede, cliente no grupo com o nome da rede, comprador de produto personalizado com o nome.',
        'Caminhos usados (todos ao mesmo tempo): CNPJ (raiz), código de cliente e grupo do ERP da aba IRTT, razão social igual,',
        'associadas da FARMACIA VF (coluna B), mapa curado de grupos de farmácia, grupo do cliente no cadastro, nome do',
        'cliente, domínio do e-mail e nome da rede na descrição do produto (logo depois da medida, ex.: 80X40M MILKY MOO).',
        '',
        'TRAVAS CONTRA FALSO VÍNCULO (revisão manual de 16/11): nome tem de estar no começo do nome do cliente e o segmento',
        'do cliente tem de combinar; palavra de tipo de outro ramo barra ("SUPERMERCADO OPCAO" não é a loja Opção); nome',
        'curto ou genérico só vale igual; supermercado e farmácia: vínculo fraco só na mesma UF (regra 2); nome usado por',
        '2+ redes: mesma UF; CNPJ de outra rede da lista não vale por nome; provedor de e-mail (terra, uol...) não vale.',
        '',
        'VENDEDOR = vendedor do CADASTRO DE CLIENTES do cliente da compra mais recente (código traduzido pelo ultimo fat).',
        'Se a carteira não é pessoa (INATIVOS, CLIENTE SEM COMPRA...), vai o vendedor daquela nota.',
        '',
        'COMO CONFERIR: aba REDES, filtrar "MUDOU" = SIM e "SÓ VÍNCULO FRACO" = SIM (os que merecem um olho). A coluna PROVA',
        'diz o cliente e o caminho. Me diga as linhas erradas que eu corrijo antes de gravar.',
        '',
        'DECISÕES PENDENTES (responder antes de gravar):',
        '1) FARMACIA VF, colunas H/I/J (aba FARMACIA VF LOJAS): proposta = manter só a H "Nº DE LOJAS (CONSIDERADO)" (é a',
        '   que o total de caixas usa) e apagar I (Receita) e J (pesquisa). As dinâmicas da aba não usam I nem J.',
        '2) FARMACIA VF, coluna P: hoje usa ATIVO / A TRABALHAR. Proposta: ATIVO / INATIVO / NÃO É CLIENTE, como as outras.',
        '3) Abas FINAL não têm coluna de data nem de vendedor. Proposta: 3 colunas novas depois da X (Total de caixas):',
        '   ÚLTIMA COMPRA AUTOPEL, VENDEDOR (CADASTRO) e COMO FOI ACHADO.',
        '4) Alessandra Dayan / VF / TOP 50: só o STATUS AUTOPEL (as tabelas não têm data nem vendedor).',
        '5) IRTT por bandeira: Status (E), Última compra (G; F e H são fórmulas e se recalculam), VENDEDOR CARTEIRA (W),',
        '   VENDEDOR (CLIENTES) (AG) e Como casou (T).',
    ]
    for t in txt: ws.append([t])
    ws['A1'].font = Font(bold=True, size=13); ws.column_dimensions['A'].width = 125
    # ---------------- REDES
    wr = wb.create_sheet('REDES')
    cols = ['ABA', 'LINHA', 'REDE (como está na aba)', 'STATUS ANTES', 'STATUS AGORA', 'MUDOU', 'ÚLTIMA COMPRA', 'DIAS',
            'VENDEDOR (CADASTRO)', 'FONTE DO VENDEDOR', 'CLIENTE DA ÚLTIMA COMPRA', 'COMO FOI ACHADO', 'PROVA', 'SÓ VÍNCULO FRACO',
            'CLIENTES LIGADOS', 'OUTRAS PROVAS (mais recente de cada caminho)', 'NOMES BUSCADOS']
    cab(wr, cols, [24, 7, 34, 15, 15, 8, 12, 7, 28, 14, 40, 22, 55, 10, 9, 90, 40])
    resumo = C.Counter()
    for i, e in enumerate(E):
        if e['extra'].get('so_cnpj') or e['aba'] in APOIO: continue
        r = R[i]
        antes = norm_st(e['extra'].get('status_antes'))
        agora = r['status']
        mudou = 'SIM' if antes.split(' (')[0] != agora else ''
        forte_d = max([p[2] for p in r.get('provas', []) if p[0] not in FRACO and p[2]], default=None)
        so_fraco = 'SIM' if r['via'] in FRACO and (not forte_d or forte_d < r['data']) and agora == 'ATIVO' else ''
        outras = ' || '.join(f"{VIA.get(p[0], p[0])} {p[2]:%d/%m/%Y} {p[1][:35]} ({p[3][:45]})" if p[2] else
                             f"{VIA.get(p[0], p[0])} sem nota {p[1][:35]}" for p in r.get('provas', [])[:4])
        wr.append([e['aba'], e['linha'], e['extra']['nome'], antes, agora, mudou, r['data'], (HOJE - r['data']).days if r['data'] else None,
                   r['vendedor'], r.get('vend_fonte', ''), r['cliente'], VIA.get(r['via'], r['via']), r['prova'], so_fraco, r['n'], outras,
                   ', '.join(r.get('nomes', [])[:6])])
        n = wr.max_row
        wr.cell(n, 5).fill = COR.get(agora, CINZA)
        if r['data']: wr.cell(n, 7).number_format = 'DD/MM/YYYY'
        resumo[(e['aba'], antes, agora)] += 1
    # ---------------- RESUMO
    wsr = wb.create_sheet('RESUMO', 1)
    cab(wsr, ['ABA', 'STATUS ANTES', 'STATUS AGORA', 'LINHAS'], [34, 26, 18, 10])
    for (a, b, c), n in sorted(resumo.items()): wsr.append([a, b, c, n])
    # ---------------- FARMÁCIAS AVULSAS
    wf = wb.create_sheet('FARMÁCIAS AVULSAS')
    cab(wf, ['LINHA (FARMACIA VF)', 'FARMÁCIA', 'REDE / ASSOCIAÇÃO', 'UF', 'STATUS ANTES', 'DATA ANTES', 'STATUS AGORA', 'ÚLTIMA COMPRA',
             'VENDEDOR (CADASTRO)', 'CLIENTE'], [10, 40, 30, 5, 22, 12, 15, 12, 28, 40])
    cont = C.Counter()
    for i, e in enumerate(E):
        if not e['extra'].get('so_cnpj'): continue
        r = R[i]; antes = norm_st(e['extra'].get('status_antes'))
        cont[(antes, r['status'])] += 1
        mud = (antes == 'ATIVO') != (r['status'] == 'ATIVO') or (antes != 'ATIVO' and r['status'] == 'INATIVO' and r['data'])
        if not mud: continue
        da = e['extra'].get('data_antes')
        wf.append([e['linha'], e['extra']['nome'], e['extra'].get('rede', ''), e['extra'].get('uf', ''), antes,
                   da.date() if isinstance(da, dt.datetime) else da, r['status'], r['data'], r['vendedor'], r['cliente']])
        for c_ in (6, 8): wf.cell(wf.max_row, c_).number_format = 'DD/MM/YYYY'
    for (a, b), n in sorted(cont.items()): wsr.append(['"FARMACIA VF" (farmácias avulsas)', a, b, n])
    # ---------------- FARMACIA VF LOJAS (H, I, J)
    wl = wb.create_sheet('FARMACIA VF LOJAS')
    src = openpyxl.load_workbook(MAPA, read_only=True, data_only=True)['"FARMACIA VF"']
    cab(wl, ['LINHA', 'REDE / NOME', 'TIPO', 'H – CONSIDERADO', 'I – RECEITA', 'J – PESQUISA', 'H = ?', 'OBSERVAÇÃO (coluna N)'],
        [8, 40, 24, 14, 12, 12, 14, 80])
    st = C.Counter()
    for i, r in enumerate(src.iter_rows(min_row=7, values_only=True), start=7):
        if not r or (r[4] is None and r[1] is None): continue
        H, I, J = r[7], r[8], r[9]
        k = 'H = I' if H == I else 'H = J' if H == J else 'outro'
        st[(str(r[2]), k)] += 1
        if r[2] != 'Farmácia': wl.append([i, r[1] or r[4], r[2], H, I, J, k, str(r[13] or '')[:200]])
    wl.append([]); wl.append(['', 'CONTAGEM (todas as linhas)'])
    for (t, k), n in sorted(st.items()): wl.append(['', t, k, n])
    wb.save(SAIDA)
    print('gravado', SAIDA, '| linhas REDES:', wr.max_row - 1, '| farmácias avulsas que mudaram:', wf.max_row - 1)


if __name__ == '__main__':
    main()
