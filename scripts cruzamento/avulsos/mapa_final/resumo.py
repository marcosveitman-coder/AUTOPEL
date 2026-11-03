"""Conteudo da aba R_E_S_U_M_O do MAPA MERCADO: metodologia, fontes pesquisadas e numeros finais das abas FINAL.
Usa so estilos que ja existem na pasta: 1 titulo, 50 faixa de secao, 194 cabecalho de tabela, 52 texto com borda, 127 numero,
0 texto corrido (passa para as celulas vazias da direita)."""
import collections
T_TIT, T_SEC, T_CAB, T_TAB, T_NUM, T_TXT = 1, 50, 194, 52, 127, 0
NCOL = 7
COLS = [34, 62, 70, 14, 14, 14, 14]
SEG = [('ALIM', 'ALIMENTACAO FINAL'), ('LOJ', 'REDE DE LOJAS FINAL'), ('CONS', 'CONSTRUÇÃO FINAL'), ('POST', 'POSTOS GAS & CONVENIENCIA FINAL')]
NIVEIS = [
    ('CNPJ CONFIRMADO NA RECEITA', 'É o CNPJ de quem compra pela rede: a operadora das lojas próprias, a franqueadora ou a dona da marca. Existe, está ativo e tem o nome da rede.'),
    ('CNPJ PROVÁVEL (CONFERIR)', 'Indício forte, sem prova completa (ex.: empresa de franquia com o nome na cidade de origem da marca). Confirmar antes de usar em contrato.'),
    ('CNPJ DE UMA UNIDADE DA REDE (FRANQUEADORA NÃO ACHADA)', 'A franqueadora não usa o nome da rede na Receita. O CNPJ é de uma loja da rede (colunas H e I): serve para achar a rede no cadastro, não é o CNPJ da rede.'),
    ('CNPJ NÃO CONFIRMADO', 'Nenhum candidato seguro (nome comum, várias empresas sem ligação). A coluna OBSERVAÇÃO diz o que foi achado.'),
    ('REDE NÃO ENCONTRADA NA RECEITA', 'Nenhum estabelecimento com o nome na UF da origem. Quase todas vêm da lista regional da aba "ALIMENTAÇÃO (...)", que tinha CNPJs inventados.'),
]
METODO = [
    ('1. Juntar as listas', 'Cada aba FINAL junta as abas de origem pedidas (704 linhas no total). Nomes iguais ou parecidos viram uma linha só, e as linhas de grupo (ex.: "GRUPO TRIGO") viram nota nas marcas, para não contar lojas duas vezes.'),
    ('2. Conferir o CNPJ da origem', 'Todo CNPJ informado passou pelo dígito verificador e pela Receita (existe, está ativo, tem o nome da rede). Dos 276 informados, 104 falharam no dígito: muitos eram inventados (listas feitas por IA).'),
    ('3. Buscar na Receita Federal', 'Cópia completa dos dados abertos do CNPJ (out/2026), varrida 7 vezes: pelo nome fantasia, pela razão social, pelo núcleo do nome (sem "pizzaria", "sushi"...) e pela razão informada nos rankings.'),
    ('4. Escolher o CNPJ, rede a rede', 'Os candidatos foram vistos à mão: operadora com muitas lojas, franqueadora (CNAE 7740-3/00, "franchising" na razão), holding ou indústria dona da marca. Nunca o CNPJ de um franqueado. CNPJ de filial vira o da matriz.'),
    ('5. Nº de lojas', 'Vale o número mais recente e com data das abas de origem ou dos rankings (ANAMACO, ABF, IRTT). Sem número, entra a contagem de lojas ativas na Receita, só com o nome completo e o CNAE do segmento.'),
    ('6. Complementar as listas', 'Entraram as redes que faltavam dos rankings (ANAMACO, IRTT, ABF) e as redes regionais achadas na própria Receita por CNAE e região (cortes: lojas 30+, alimentação 20+, construção 9+ unidades).'),
    ('7. Região e ranking', 'A região é a da UF da sede (Receita; sem CNPJ, a pesquisa ou a aba de origem). O ranking é pelo nº de lojas: no Brasil e em cada região. A aba MAIORES POR REGIÃO mostra as 30 maiores do Brasil e as 20 maiores de cada região.'),
    ('8. Postos e conveniência', 'Sem lista de origem: as redes vieram da Receita (empresas com muitos postos, CNAE 4731-8/00, e lojas de conveniência, 4729-6/02), das distribuidoras (CNAE 4681-8) e da ANP (postos por bandeira e região). Nº de postos: ANP para as bandeiras, Receita para as redes e imprensa quando havia número recente.'),
    ('9. Conferência final', 'Antes de gravar: todo CNPJ com dígito válido e ativo, nenhuma rede repetida na mesma aba, e a planilha aberta no Excel sem erro. A gravação mexe só nestas abas; as tabelas dinâmicas não são tocadas.'),
]
FONTES = [
    ('Receita Federal – dados abertos do CNPJ (out/2026)', 'Base oficial: CNPJ, razão social, nome fantasia, situação, CNAE, município e lojas de cada empresa.', 'https://arquivos.receitafederal.gov.br/'),
    ('BrasilAPI – consulta de CNPJ', 'Conferência pontual de CNPJ (situação, matriz, endereço, sócios).', 'https://brasilapi.com.br/'),
    ('Revista ANAMACO – Top 50 Varejo Nacional 2026', 'Lojas de material de construção, posição e razão social.', 'https://www.revistaanamaco.com.br/top-50-varejo-nacional---2026'),
    ('Revista ANAMACO – Top 50 Varejo Nacional 2025', 'Idem, edição anterior.', 'https://www.revistaanamaco.com.br/top-50-varejo-nacional---2025'),
    ('Ranking IRTT Top 300 (aba I.R.T.T. desta planilha)', 'Bandeiras de lojas físicas que não estavam nas origens.', 'aba "I.R.T.T.__Por bandeira (FINAL)"'),
    ('Exame – maiores empresas de material de construção', 'Ranking de material de construção (IRTT 2024): nº de lojas.', 'https://exame.com/negocios/quais-sao-as-maiores-empresas-de-material-de-construcao-do-brasil-veja-quanto-elas-faturam/'),
    ('ABF / Portal do Franchising – maiores franquias', 'Nº de unidades das franquias; incluiu Fast4You, Piticas, Multicoisas e Livraria Nobel.', 'https://www.portaldofranchising.com.br/franquias/maiores-franquias-em-unidades/'),
    ('Economic News Brasil – maiores franquias 2025', 'Ranking ABF 2025.', 'https://economicnewsbrasil.com.br/2026/03/07/maiores-franquias-do-brasil-ranking-2025/'),
    ('Nosso Meio – nordestinas no ranking IRTT', 'Maiores varejistas do Nordeste (conferência regional).', 'https://nossomeio.com.br/33-empresas-nordestinas-aparecem-entre-as-300-maiores-do-varejo-brasileiro-segundo-o-irtt/'),
    ('Diário do Nordeste – maiores varejistas do Ceará', 'Conferência regional (Ceará).', 'https://diariodonordeste.verdesmares.com.br/negocios/ceara-tem-12-empresas-entre-as-maiores-varejistas-do-pais-veja-ranking-1.3787486'),
    ('ABRAMAT', 'Consultada. É associação de indústrias de materiais: não lista lojas.', 'https://www.abramat.org.br/'),
    ('ABCEM', 'Consultada. É associação de construção metálica: não lista lojas.', 'https://www.abcem.org.br/'),
    ('ANR – Associação Nacional de Restaurantes', 'Consultada. A lista de associados não é pública.', 'https://anrbrasil.org.br/associados/'),
    ('Abrasel', 'Consultada. Não publica lista de redes associadas.', 'https://abrasel.com.br/'),
    ('ABVTEX, FEBRAMAT e SBVC', 'Consultadas. ABVTEX e FEBRAMAT não publicam associados; o ranking SBVC mais recente é de 2024.', 'pesquisa Google'),
    ('Notícias (Gironews, Mercado&Consumo, Jornal Opção, Seu Dinheiro, Baguete e outras)', 'Grupo administrador, sede e nº de lojas quando a lista não trazia (ex.: KFC, Pizza Hut, Domino\'s, Mr. Fit, Frutos de Goiás, Piticas).', 'pesquisa Google'),
    ('ANP – Anuário Estatístico 2026, tabela 3.17', 'Postos por bandeira (Vibra, Ipiranga, Raízen, Alesat, outras, bandeira branca) por região e UF, 2025.', 'https://www.gov.br/anp/pt-br/centrais-de-conteudo/dados-abertos/anuario/anuario-estatistico-2026-dados-abertos'),
    ('Fecombustíveis', 'Consultada. Não publica lista de redes de postos; o número oficial de postos por bandeira é o da ANP.', 'https://www.fecombustiveis.org.br/'),
    ('ABIEPS', 'Consultada. Reúne empresas de equipamentos e serviços para postos e conveniência (ExpoPostos): não lista redes.', 'pesquisa Google'),
    ('Imprensa de postos (InvestNews, Suno, NovaCana, Jornal do Comércio, InfoMoney, Movimento Econômico)', 'Vibra 7.456 postos (4T25), Shell Select 1.256 lojas (set/2025), SIM Rede 191 unidades, Charrua 500+, Rodoil 600, Petrobahia 113, Ale (Glencore).', 'pesquisa Google'),
    ('Abas desta planilha', 'Origem das redes: "ALIMENTAÇÃO (...)", "Alessandra Dayan", "VAREJO (MAIORES)_RADES" e "VAREJO VESTUARIO (IA)".', 'abas da pasta'),
]
CUIDADOS = [
    'Listas feitas por IA trazem CNPJ e até rede inventados: CNPJ com dígito certo não basta, precisa existir na Receita, estar ativo e ter o nome.',
    '"Ativa" na Receita não quer dizer loja aberta (ex.: Ricardo Eletro, Saraiva, Mappin ainda aparecem ativas).',
    'Em franquia, cada loja é uma empresa: o CNPJ de uma unidade não compra pela rede. Para negociar com a rede, vale a franqueadora.',
    'Os números de lojas têm datas diferentes (coluna FONTE DO Nº DE LOJAS). Nomes comuns (sobrenomes, "Praça", "Amigão") inflam a contagem pelo nome.',
    'O cruzamento com o faturamento e o cadastro de clientes (cliente atendido, última compra, vendedor) ainda não foi feito.',
]
def resumo_xml(rows, row_xml, get_column_letter, rotulo_data):
    out = []; k = 1
    def put(vals, ht=None):
        nonlocal k
        out.append(row_xml(k, vals, ht)); k += 1
    def sec(t):
        put([(t, T_SEC)] + [('', T_SEC)] * (NCOL - 1), ht=20)
    def pula():
        nonlocal k
        k += 1
    put([('RESUMO DO ESTUDO DE REDES – ALIMENTAÇÃO, REDE DE LOJAS E CONSTRUÇÃO', T_TIT)], ht=24)
    put([(f'Atualizado em {rotulo_data}. Abas: ALIMENTACAO FINAL, REDE DE LOJAS FINAL, CONSTRUÇÃO FINAL, POSTOS GAS & CONVENIENCIA FINAL e MAIORES POR REGIÃO.', T_TXT)])
    pula()
    sec('1. OBJETIVO')
    put([('Montar, para cada segmento, uma lista única de redes com nome fantasia, razão social, grupo administrador, nº de lojas e CNPJ,', T_TXT)])
    put([('e mostrar as maiores do Brasil e de cada região, para o planejamento comercial da Autopel.', T_TXT)])
    pula()
    sec('2. METODOLOGIA (PASSO A PASSO)')
    put([('Passo', T_CAB), ('Como foi feito', T_CAB)] + [('', T_CAB)] * (NCOL - 2), ht=18)
    for a, b in METODO:
        put([(a, T_TAB), (b, T_TAB)], ht=48)
    pula()
    sec('3. COMO LER A COLUNA "CONFERÊNCIA DO CNPJ"')
    put([('Nível', T_CAB), ('O que quer dizer', T_CAB), ('', T_CAB), ('Alimentação', T_CAB), ('Rede de lojas', T_CAB), ('Construção', T_CAB), ('Postos e conveniência', T_CAB)], ht=30)
    SEGR = [(s, a) for s, a in SEG if s in rows]
    cnt = {s: collections.Counter(r['conf'] for r in rows[s]) for s, _ in SEGR}
    for n, d in NIVEIS:
        put([(n, T_TAB), (d, T_TAB), ('', T_TAB)] + [(cnt[s].get(n, 0), T_NUM) for s, _ in SEGR], ht=45)
    put([('TOTAL DE REDES', T_TAB), ('', T_TAB), ('', T_TAB)] + [(len(rows[s]), T_NUM) for s, _ in SEGR])
    pula()
    sec('4. LUGARES ONDE A PESQUISA FOI FEITA')
    put([('Fonte', T_CAB), ('O que foi usado', T_CAB), ('Endereço', T_CAB)] + [('', T_CAB)] * (NCOL - 3), ht=18)
    for a, b, c in FONTES:
        put([(a, T_TAB), (b, T_TAB), (c, T_TAB)], ht=32)
    pula()
    sec('5. MAIORES REDES DE CADA SEGMENTO (Nº DE LOJAS)')
    put([('Segmento', T_CAB), ('5 maiores do Brasil', T_CAB), ('Maior de cada região', T_CAB)] + [('', T_CAB)] * (NCOL - 3), ht=18)
    for s, aba in SEGR:
        v = sorted([r for r in rows[s] if r['lojas'] and not r['conf'].startswith('REDE NÃO')], key=lambda r: -r['lojas'])
        top = ', '.join(f"{r['fant']} ({r['lojas']:,})".replace(',', '.') for r in v[:5])
        reg = []
        for g in ['NORTE', 'NORDESTE', 'CENTRO-OESTE', 'SUDESTE', 'SUL']:
            x = next((r for r in v if r.get('regiao') == g), None)
            if x: reg.append(f"{g.title()}: {x['fant']}")
        put([(aba, T_TAB), (top, T_TAB), ('; '.join(reg), T_TAB)], ht=48)
    put([('A lista completa está na aba MAIORES POR REGIÃO (30 do Brasil e 20 por região, em cada segmento).', T_TXT)])
    pula()
    sec('6. CUIDADOS AO USAR')
    for t in CUIDADOS:
        put([('• ' + t, T_TXT)])
    pula()
    sec('7. PRÓXIMOS PASSOS')
    put([('• Cruzar as redes com o faturamento e o cadastro de clientes (já atendida, última compra, vendedor).', T_TXT)])
    put([('• Confirmar os CNPJs "prováveis" e os de unidade com a COF (circular de franquia) ou com contato direto com a rede.', T_TXT)])
    put([('• Decidir se as redes "não encontradas" saem da lista ou se a fonte da aba de origem confirma que existem.', T_TXT)])
    cols = '<cols>' + ''.join(f'<col min="{j+1}" max="{j+1}" width="{w}" customWidth="1"/>' for j, w in enumerate(COLS)) + '</cols>'
    return cols + '<sheetData>' + ''.join(out) + '</sheetData>', f'A1:{get_column_letter(NCOL)}{k - 1}'
