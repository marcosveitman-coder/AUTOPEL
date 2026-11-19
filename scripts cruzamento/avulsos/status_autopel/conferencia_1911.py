"""Planilha de conferência da revisão do STATUS AUTOPEL de 19/11/2026 (caso "Koni Store"), para o usuário aprovar.
Abas:
  LEIA      o que foi revisto e as regras novas;
  MUDANÇAS  toda linha das abas do MAPA cujo STATUS muda (antes x depois), com a última compra, o cliente e a prova;
  DATAS     linhas que continuam com o mesmo status, mas com outra última compra ou outro cliente (só abas FINAL);
  CONFLITOS a mesma empresa (raiz do CNPJ) com status diferentes em linhas de MARCAS diferentes: é a regra de 16/11
            ("a marca não herda a compra de outra marca do grupo"), mostrada para o usuário confirmar.
Entrada: <pasta>/entidades.pkl e resultados.pkl (coleta.py + avalia.py) e auditoria.tsv (auditoria.py).
Uso: python conferencia_1911.py "<pasta>" "<saída.xlsx>" """
import sys, os, pickle, csv, re, collections as C
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

P, SAI = sys.argv[1:3]
E = pickle.load(open(os.path.join(P, 'entidades.pkl'), 'rb'))
R = pickle.load(open(os.path.join(P, 'resultados.pkl'), 'rb'))
MP = {'Ativo': 'ATIVO', 'Inativo': 'INATIVO', 'Cadastrado sem compra': 'INATIVO', 'Não atendido': 'NÃO É CLIENTE'}
VIA = {'CNPJ': 'CNPJ (raiz)', 'CODIGO': 'código de cliente (aba IRTT)', 'RAZAO': 'razão social igual', 'ASSOCIADA': 'associada da rede',
       'GRUPO MAPA': 'grupo do cadastro (mapa curado)', 'GRUPO ERP': 'grupo do ERP (aba IRTT)', 'GRUPO': 'grupo do cliente no cadastro',
       'EXATO': 'nome fantasia igual', 'NOME': 'nome do cliente', 'EMAIL': 'domínio do e-mail', 'PRODUTO': 'produto personalizado'}
AZUL = PatternFill('solid', fgColor='1F3A5F')
COR = {'ATIVO': PatternFill('solid', fgColor='E2EFDA'), 'INATIVO': PatternFill('solid', fgColor='FFF2CC'),
       'NÃO É CLIENTE': PatternFill('solid', fgColor='FCE4D6')}
GRAVA = lambda e: not e['extra'].get('apoio') and e['aba'] != 'SUPERMERCADOS (reuniao)' and e['aba'] != 'FULLFILMIT_INDUSTRIA FINAL'


def cab(ws, cols):
    ws.append([c for c, _ in cols])
    for j, (c, w) in enumerate(cols, 1):
        x = ws.cell(1, j); x.font = Font(bold=True, color='FFFFFF'); x.fill = AZUL; x.alignment = Alignment(wrap_text=True, vertical='center')
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    ws.freeze_panes = 'A2'


wb = openpyxl.Workbook()
ws = wb.active; ws.title = 'LEIA'
for t in [
    'Revisão do STATUS AUTOPEL — 19/11/2026',
    '',
    'Por que: a linha "Koni Store" (ALIMENTACAO FINAL 131) estava NÃO É CLIENTE, mas a mesma marca, com o mesmo CNPJ, está ATIVO na aba IRTT',
    '(bobina "80X100M KONI" comprada pela empresa do grupo). Foi revisto o status de TODAS as abas, com os relatórios do TOTVS de hoje.',
    '',
    'O que mudou nas regras:',
    '1. Base do dia: cadastro, último faturamento, notas do mês e pedidos em aberto do TOTVS de 19/11 (antes, a foto era de 16/11).',
    '2. Mesma marca em abas diferentes: linhas com o mesmo CNPJ e o nome de uma no começo da outra ("KONI" x "KONI STORE"), ou com só',
    '   uma palavra de formato a mais ("PIZZARIA SPOLETO" = SPOLETO; "CLUBE MELISSA" = MELISSA), têm agora UM status.',
    '3. CNPJ de UMA UNIDADE (franqueado, coluna H, ou "CNPJ DE UMA UNIDADE" na G) vale como vínculo, mas não separa mais a rede da mesma',
    '   rede nas listas da Alessandra (We Pink, Santo Remédio etc. tinham dois status).',
    '4. Linha de grupo (ex.: "IMC (FRANGO ASSADO / PIZZA HUT / KFC)") não empresta à marca o nome de OUTRA marca que tem linha própria',
    '   (o Frango Assado estava ATIVO por uma bobina da Pizza Hut; agora vale a bobina "FRANGO ASSADO POSTOS" da própria IMC).',
    '5. Produtos personalizados e clientes conferidos à mão (Tendall, Calvin Klein, Natura, Apoio Mineiro, Rei do Mate, Unissul,',
    '   Mega Descontão, Petrox) e exclusões de homônimos (cooperativa de crédito na Coop; e-mail genérico "sistemas").',
    '6. A aba de indústrias deixou de criar nomes ambíguos para as outras abas (NATURA + AVON voltou a ATIVO).',
    '',
    'CNPJs corrigidos nas listas da Alessandra / TOP 50 (preenchidos em 17/11):',
    '- AZZAS 2154: estava o CNPJ do Magazine Luiza; foi trocado pelo CNPJ da Azzas 2154 S.A.',
    '- St. Marché: o CNPJ da aba SUPERMERCADOS (reuniao) é do Spani Atacadista (Taubaté). O St. Marché é a Hortus, cadastrada no',
    '  TOTVS sem nota: fica INATIVO. A aba SUPERMERCADOS (reuniao) é sua: o CNPJ dela não foi alterado.',
    '',
    'Continua valendo (decisão de 16/11): a marca não herda a compra de outra marca do mesmo grupo. A aba CONFLITOS lista esses casos',
    '(ex.: Burger King INATIVO e Subway ATIVO, ambos da Zamp). Se quiser que a marca herde do grupo, é só pedir.',
    '',
    'Pedidos em aberto não contam como compra (a regra é nota fiscal com valor). Ficam como informação na auditoria.']:
    ws.append([t])
ws.column_dimensions['A'].width = 140
ws2 = wb.create_sheet('MUDANÇAS')
cols = [('ABA', 26), ('LINHA', 7), ('NOME NA ABA', 34), ('STATUS ANTES', 15), ('STATUS AGORA', 15), ('ÚLTIMA COMPRA', 12),
        ('COMO FOI ACHADO', 26), ('CLIENTE', 40), ('PROVA', 60), ('VENDEDOR', 26)]
cab(ws2, cols)
n = C.Counter()
for i, e in enumerate(E):
    if not GRAVA(e): continue
    a = MP.get(e['extra'].get('status_antes', ''), e['extra'].get('status_antes', ''))
    r = R[i]
    if a == r['status']: continue
    if e['extra'].get('so_cnpj') and not a: continue
    ws2.append([e['aba'], e['linha'], e['extra'].get('nome', ''), a, r['status'], r['data'], VIA.get(r['via'], r['via']),
                re.sub(r'^\S+\s', '', r.get('cliente', '')), r.get('prova', ''), r['vendedor'] if r['status'] != 'NÃO É CLIENTE' else ''])
    k = ws2.max_row
    ws2.cell(k, 5).fill = COR.get(r['status'], AZUL); ws2.cell(k, 6).number_format = 'DD/MM/YYYY'
    n[(a, r['status'])] += 1
ws3 = wb.create_sheet('CONFLITOS')
cab(ws3, [('RAIZ DO CNPJ', 12), ('ABA', 26), ('LINHA', 7), ('NOME NA ABA', 34), ('STATUS', 15), ('ÚLTIMA COMPRA', 12), ('COMO', 24), ('CLIENTE', 40)])
au = os.path.join(P, 'auditoria.tsv')
if os.path.exists(au):
    for r in list(csv.reader(open(au, encoding='utf-8-sig'), delimiter='\t'))[1:]:
        if r[0] == 'CONFLITO RAIZ':
            ws3.append([r[1], r[2], int(r[3]), r[4], r[5], r[6], VIA.get(r[7], r[7]), re.sub(r'^\S+\s', '', r[8])])
            ws3.cell(ws3.max_row, 5).fill = COR.get(r[5], AZUL)
for w in (ws2, ws3): w.auto_filter.ref = w.dimensions
wb.save(SAI)
print('mudanças:', sum(n.values()), dict(n))
