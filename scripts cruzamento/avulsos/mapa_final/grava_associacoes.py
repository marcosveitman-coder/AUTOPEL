"""Cria a aba ASSOCIASSOES MAPA (nome pedido pelo usuario) no fim do MAPA MERCADO, so pelo XML (a pasta tem tabelas
dinamicas): associacoes nacionais e regionais de cada segmento do mapeamento, com o que representam, se ja foram
usadas no estudo, se foram conferidas na internet em 05/11/2026, site e observacao. Dados publicos (sem CNPJ).
Uso: python grava_associacoes.py <MAPA.xlsx> <saida.xlsx>"""
import os as _os, sys as _sys
_sys.argv[1:] = [_os.path.abspath(a) for a in _sys.argv[1:]]
import re, sys, zipfile
from xml.sax.saxutils import escape
from openpyxl.utils import get_column_letter
MAPA, OUT = sys.argv[1], sys.argv[2]
NOVA = 'ASSOCIASSOES MAPA'
S_HEAD, S_TXT, S_NUM = 194, 0, 127
REG = {}
for r, ufs in [('NORTE', 'AC AP AM PA RO RR TO'), ('NORDESTE', 'AL BA CE MA PB PE PI RN SE'), ('CENTRO-OESTE', 'DF GO MT MS'),
               ('SUDESTE', 'ES MG RJ SP'), ('SUL', 'PR RS SC')]:
    for u in ufs.split(): REG[u] = r
COLS = [("Nº", 6), ("ABRANGÊNCIA", 13), ("SEGMENTO", 28), ("SIGLA", 22), ("NOME COMPLETO", 58), ("UF / BASE", 16), ("REGIÃO", 14),
        ("O QUE REPRESENTA", 60), ("SITUAÇÃO NO ESTUDO", 34), ("ONDE FOI USADA", 34), ("CONFERIDA NA INTERNET (05/11/2026)", 16),
        ("SITE", 30), ("OBSERVAÇÃO", 70)]
USADA, CONSULT, FALTA, NAOENC, REGIO = ('JÁ USADA NO MAPEAMENTO', 'CONSULTADA – NÃO PUBLICA LISTA DE REDES', 'FALTOU – A CONSULTAR',
                                         'NÃO ENCONTRADA', 'REGIONAL – A CONSULTAR')
S, N = 'SIM', 'NÃO (entidade conhecida)'
# (abrangencia, segmento, sigla, nome, uf, o que representa, situacao, onde, conferida, site, obs)
NAC = [
 ('Supermercados e atacarejo', 'ABRAS', 'Associação Brasileira de Supermercados', 'Supermercados de todo o país; publica o Ranking ABRAS', USADA,
  'SUPERMERCADOS (reuniao); I.R.T.T.', S, 'abras.com.br', 'Ranking ABRAS 2026 (faturamento de 2025); 27 associações estaduais filiadas'),
 ('Franquias (vários segmentos)', 'ABF', 'Associação Brasileira de Franchising', 'Franquias; ranking das maiores redes por nº de unidades', USADA,
  'ALIMENTACAO FINAL; REDE DE LOJAS FINAL; PET SHOP  FINAL', S, 'abf.com.br', 'nº de unidades via ABF / Portal do Franchising'),
 ('Varejo (vários segmentos)', 'IRTT', '', 'Ranking Top 300 do varejo (por grupo e bandeira)', USADA, 'I.R.T.T.__Por bandeira (FINAL)', S, '',
  'nome completo da entidade não conferido'),
 ('Material de construção', 'ANAMACO', 'Associação Nacional dos Comerciantes de Material de Construção', 'Varejo de material de construção; ranking Top 50 Varejo Nacional',
  USADA, 'CONSTRUÇÃO FINAL', S, 'revistaanamaco.com.br', 'reúne 38 Acomacs e 3 Fecomacs regionais'),
 ('Material de construção', 'ABRAMAT', 'Associação Brasileira da Indústria de Materiais de Construção', 'Indústria de materiais', CONSULT, '', S, 'abramat.org.br', 'não lista lojas'),
 ('Material de construção', 'ABCEM', 'Associação Brasileira da Construção Metálica', 'Construção metálica', CONSULT, '', S, 'abcem.org.br', 'não lista lojas'),
 ('Material de construção', 'FEBRAMAT', '', 'Comércio de material de construção', CONSULT, '', S, '', 'não publica associados; nome completo não conferido'),
 ('Rede de lojas', 'ABVTEX', 'Associação Brasileira do Varejo Têxtil', 'Varejo de moda', CONSULT, '', S, '', 'não publica associados'),
 ('Rede de lojas', 'SBVC', 'Sociedade Brasileira de Varejo e Consumo', 'Varejo; ranking das maiores varejistas', CONSULT, '', S, '', 'o ranking mais recente é de 2024'),
 ('Farmácias', 'Abrafarma', 'Associação Brasileira de Redes de Farmácias e Drogarias', 'Grandes redes de farmácias', USADA, '"FARMACIA VF" (coluna ENTIDADE / FEDERAÇÃO)', S, '', ''),
 ('Farmácias', 'Febrafar', 'Federação Brasileira das Redes Associativistas e Independentes de Farmácias', 'Redes associativistas (bandeiras de farmácias independentes)',
  USADA, '"FARMACIA VF" (coluna ENTIDADE / FEDERAÇÃO)', S, '', ''),
 ('Farmácias', 'Abrafad', '', 'Entidade de farmácias citada na base de farmácias', USADA, '"FARMACIA VF" (coluna ENTIDADE / FEDERAÇÃO)', 'NÃO', '', 'nome completo não conferido'),
 ('Farmácias', 'Fecofar', '', 'Entidade de farmácias citada na base de farmácias', USADA, '"FARMACIA VF" (coluna ENTIDADE / FEDERAÇÃO)', 'NÃO', '', 'nome completo não conferido'),
 ('Alimentação', 'Abrasel', 'Associação Brasileira de Bares e Restaurantes', 'Bares e restaurantes', CONSULT, '', S, 'abrasel.com.br', 'tem seccionais estaduais'),
 ('Alimentação', 'ANR', 'Associação Nacional de Restaurantes', 'Redes de restaurantes', CONSULT, '', S, 'anrbrasil.org.br', 'lista de associados não é pública'),
 ('Postos e conveniência', 'Fecombustíveis', 'Federação Nacional do Comércio de Combustíveis e de Lubrificantes', 'Sindicatos patronais dos postos (revendedores)',
  CONSULT, '', S, 'fecombustiveis.org.br', 'diz reunir 34 sindicatos e a Abragás; em abr/2026 os sindicatos de BA, MG, SP, Santos, PR, SC, RS, MT e DF '
  'anunciaram a saída para criar outra entidade (Metrópoles, 10/04/2026)'),
 ('Postos e conveniência', 'ABIEPS', 'Associação Brasileira das Empresas de Equipamentos e de Serviços para o Mercado de Combustíveis e de Conveniência',
  'Fornecedores de equipamentos e serviços para postos e conveniência; organiza a ExpoPostos', CONSULT, '', S, 'abieps.com.br', 'não lista redes'),
 ('Postos e conveniência', 'ANP', 'Agência Nacional do Petróleo, Gás Natural e Biocombustíveis', 'Agência reguladora (não é associação): postos por bandeira',
  USADA, 'POSTOS GAS & CONVENIENCIA FINAL; MAIORES POR REGIÃO', S, 'gov.br/anp', 'Anuário 2026, tabela 3.17'),
 ('Pet', 'IPB', 'Instituto Pet Brasil', 'Números do mercado pet (faturamento, nº de pet shops)', CONSULT, '', S, '', 'não lista redes'),
 ('Pet', 'ABINPET', 'Associação Brasileira da Indústria de Produtos para Animais de Estimação', 'Indústria pet (fabricantes)', CONSULT, '', S, 'abinpet.org.br', 'não é varejo'),
 ('Pet', 'ANCLIVEPA', 'Associação Nacional de Clínicos Veterinários de Pequenos Animais', 'Veterinários de pequenos animais', CONSULT, '', S, '', 'não lista redes de clínicas'),
 ('Estacionamento', 'ABRAPARK', 'Associação Brasileira de Estacionamentos', 'Operadoras de estacionamento', CONSULT, '', S, 'abrapark.com.br', 'não publica associados'),
 ('Locação de veículos', 'ABLA', 'Associação Brasileira das Locadoras de Automóveis', 'Locadoras de veículos (Localiza, Movida, Unidas…)', FALTA, '', S, '',
  'estava na lista de pet, mas é de locadoras: serve ao segmento Automotivo, Peças e Locadoras do CRM'),
 ('Pet', 'ABRASE', '', 'Não existe entidade do setor pet com essa sigla', NAOENC, '', S, '', 'talvez seja ABRASEL (bares e restaurantes) ou ABASE (supermercados da Bahia)'),
 ('Atacarejo', 'ABAAS', 'Associação Brasileira dos Atacadistas de Autosserviço', 'Atacarejos (Assaí, Atacadão, Roldão…)', FALTA, '', S, 'abaas.com.br',
  'publica o Ranking ABAAS com a NielsenIQ (2026, dados de 2025)'),
 ('Atacado distribuidor', 'ABAD', 'Associação Brasileira de Atacadistas e Distribuidores de Produtos Industrializados', 'Atacado distribuidor (segmento Revenda do CRM)', FALTA, '', N, '', ''),
 ('Alimentação', 'ABERC', 'Associação Brasileira das Empresas de Refeições Coletivas', 'Refeição coletiva (Sodexo, GRSA…)', FALTA, '', S, 'aberc.com.br',
  'a ALIMENTACAO FINAL tem o subsegmento de refeição coletiva'),
 ('Alimentação', 'ABIP', 'Associação Brasileira da Indústria de Panificação e Confeitaria', 'Padarias e confeitarias', FALTA, '', N, '', ''),
 ('Rede de lojas', 'IDV', 'Instituto para Desenvolvimento do Varejo', 'As cerca de 70 maiores varejistas do país', FALTA, '', S, 'idv.org.br', ''),
 ('Rede de lojas', 'CNDL', 'Confederação Nacional de Dirigentes Lojistas', 'Lojistas; rede das CDLs municipais', FALTA, '', N, '', ''),
 ('Shopping center', 'ABRASCE', 'Associação Brasileira de Shopping Centers', 'Shopping centers (lista dos shoppings do país)', FALTA, '', N, '', 'é nacional'),
 ('Lojistas de shopping', 'Alshop', 'Associação Brasileira de Lojistas de Shopping', 'Lojistas de shopping (diz reunir 54 mil lojistas)', FALTA, '', S, '', ''),
 ('Lojistas de shopping', 'ABLOS', 'Associação Brasileira de Lojistas Satélites', 'Lojas pequenas e médias de shopping', FALTA, '', S, '', 'criada em 2019, a partir de um racha com a Alshop'),
 ('Farmácias', 'Abcfarma', 'Associação Brasileira do Comércio Farmacêutico', 'Farmácias independentes', FALTA, '', N, '', ''),
 ('Postos (distribuidoras)', 'Brasilcom', 'Federação Nacional das Distribuidoras de Combustíveis, Gás Natural e Biocombustíveis', 'Distribuidoras regionais de combustível',
  FALTA, '', S, 'federacaobrasilcom.org.br', 'reúne 8 sindicatos de distribuidoras (linhas regionais abaixo)'),
 ('Beleza / Saúde / Bem-estar', 'ACAD Brasil', 'Associação Brasileira de Academias', 'Academias (Smart Fit, Bluefit, Selfit…)', FALTA, '', S, '', ''),
 ('Óticas', 'Abióptica', 'Associação Brasileira das Indústrias Ópticas', 'Indústria óptica', FALTA, '', S, '', 'o varejo óptico de SP tem o Sindióptica (linha regional)'),
 ('Serviços / Franquias', 'ANEL', 'Associação Nacional das Empresas de Lavanderia', 'Lavanderias (cerca de 2 mil empresas)', FALTA, '', S, '', ''),
 ('Automação comercial', 'AFRAC', 'Associação Brasileira de Tecnologia para o Comércio e Serviços (antiga Associação Brasileira de Automação Comercial)',
  'Fabricantes de PDV e periféricos, softwares, revendas e fornecedores de papel (bobinas): o setor da própria Autopel', FALTA, '', S, 'afrac.com.br', 'mais de 300 associadas; organiza a feira Autocom'),
 ('Pedágio', 'ABCR', 'Associação Brasileira de Concessionárias de Rodovias', 'Concessionárias de rodovias (praças de pedágio)', FALTA, '', N, '', 'segmento do CRM ainda sem lista'),
 ('Transporte / Autoviação', 'ABRATI', 'Associação Brasileira das Empresas de Transporte Terrestre de Passageiros', 'Empresas de ônibus rodoviário', FALTA, '', N, '', 'segmento do CRM ainda sem lista'),
 ('Logística', 'ABOL', 'Associação Brasileira de Operadores Logísticos', 'Operadores logísticos', FALTA, '', N, '', 'segmento do CRM ainda sem lista'),
 ('Corporativo Saúde', 'ANAHP', 'Associação Nacional de Hospitais Privados', 'Hospitais privados', FALTA, '', N, '', 'segmento do CRM ainda sem lista'),
 ('E-commerce', 'ABComm', 'Associação Brasileira de Comércio Eletrônico', 'Lojas virtuais', FALTA, '', N, '', 'segmento do CRM ainda sem lista'),
 ('Entretenimento', 'Abraplex', 'Associação Brasileira das Empresas Exibidoras Cinematográficas Operadoras de Multiplex', 'Redes de cinema', FALTA, '', N, '', 'segmento do CRM ainda sem lista'),
]
# regionais: (segmento, sigla, nome, uf, o que representa, situacao, onde, conferida, site, obs)
SUPER = [('Asas', 'Associação Acreana de Supermercados', 'AC'), ('Amaps', 'Associação Amapaense de Supermercados', 'AP'),
         ('Amase', 'Associação Amazonense de Supermercados', 'AM'), ('Aspas', 'Associação Paraense de Supermercados', 'PA'),
         ('Asmeron', 'Associação de Supermercados de Rondônia', 'RO'), ('Atos', 'Associação Tocantinense de Supermercados', 'TO'),
         ('Asa', 'Associação de Supermercados de Alagoas', 'AL'), ('Abase', 'Associação Bahiana de Supermercados', 'BA'),
         ('Acesu', 'Associação Cearense de Supermercados', 'CE'), ('Amasp', 'Associação Maranhense de Supermercados', 'MA'),
         ('Aspb', 'Associação de Supermercados da Paraíba', 'PB'), ('Apes', 'Associação Pernambucana de Supermercados', 'PE'),
         ('Apisu', 'Associação Piauiense de Supermercados', 'PI'), ('Assurn', 'Associação de Supermercados do Rio Grande do Norte', 'RN'),
         ('Ases', 'Associação Sergipana de Supermercados', 'SE'), ('Asbra', 'Associação de Supermercados de Brasília', 'DF'),
         ('Agos', 'Associação Goiana de Supermercados', 'GO'), ('Asmat', 'Associação de Supermercados de Mato Grosso', 'MT'),
         ('Amas', 'Associação Sul-Mato-Grossense de Supermercados', 'MS'), ('Acaps', 'Associação Capixaba de Supermercados', 'ES'),
         ('Asserj', 'Associação de Supermercados do Estado do Rio de Janeiro', 'RJ'), ('Amis', 'Associação Mineira de Supermercados', 'MG'),
         ('Apas', 'Associação Paulista de Supermercados', 'SP'), ('Apras', 'Associação Paranaense de Supermercados', 'PR'),
         ('Agas', 'Associação Gaúcha de Supermercados', 'RS'), ('Acats', 'Associação Catarinense de Supermercados', 'SC')]
REGS = []
for sig, nome, uf in SUPER:
    REGS.append(('Supermercados', sig.upper(), nome, uf, 'Supermercados do estado (filiada à ABRAS)', 'CITADA NO MAPEAMENTO' if sig == 'Apas' else REGIO,
                 'SUPERMERCADOS (reuniao)' if sig == 'Apas' else '', S, '', 'lista da página de associações estaduais da ABRAS (a de Roraima não aparece)'))
POSTOS = [('MINASPETRO', 'MG'), ('PARANAPETRO', 'PR'), ('SCPETRO', 'SC'), ('SINDICOMBUSTÍVEIS DF', 'DF'), ('SINDICOMBUSTÍVEIS BAHIA', 'BA'),
          ('SINDIPETRÓLEO MT', 'MT'), ('SINDIPOSTOGO', 'GO'), ('SINDIPOSTOS PI', 'PI'), ('SINDIPOSTOS ES', 'ES'), ('SINPETRO MS', 'MS'),
          ('SINDIPETRO PB', 'PB'), ('SINDICOMBUSTÍVEIS AL', 'AL'), ('SINDICOMB RJ', 'RJ'), ('SINDIPOSTOS RN', 'RN'), ('SULPETRO', 'RS'),
          ('SINDICOMBUSTÍVEIS PE', 'PE'), ('SINDICOMBUSTÍVEIS PA', 'PA'), ('SINDIPOSTO-TO', 'TO'), ('SINDIPETRO-RO', 'RO'), ('SINDICOMBUSTÍVEIS MA', 'MA'),
          ('SINDIPESE', 'SE'), ('RESAN', 'SP (Baixada Santista)'), ('RECAP', 'SP (Campinas)')]
for sig, uf in POSTOS:
    obs = 'assinou a carta da Fecombustíveis de 30/03/2026'
    if sig == 'SINDIPESE': obs += '; UF pelo nome (a carta não indica)'
    if sig in ('RESAN', 'RECAP'): obs += '; base regional pelo nome do sindicato (a carta não indica)'
    REGS.append(('Postos e conveniência', sig, '', uf, 'Sindicato patronal dos postos de combustível', REGIO, '', S, '', obs))
REGS.append(('Postos e conveniência', 'SINCOPETRO', '', 'SP', 'Sindicato patronal dos postos de combustível de SP', REGIO, '', N, '', 'não está entre os signatários da carta de 30/03/2026'))
REGS.append(('Postos e conveniência', 'SINDÍPOLIS, SINPEB, SINDPAC, SINDTRR', '', '', 'Outros signatários da carta da Fecombustíveis (UF não indicada)', REGIO, '', S, '',
             'SINDTRR = transportadores revendedores retalhistas (TRR)'))
for sig, uf in [('SINDIMINAS', 'MG'), ('SINDISUL', 'RS'), ('SINDESC', 'SC'), ('SICOMPAR', 'PR'), ('SINDIGOIÁS', 'GO'), ('SIDICOESP', 'SP'), ('SINDIMAT', 'MT'), ('SINDIPE', 'PE')]:
    REGS.append(('Postos (distribuidoras)', sig, '', uf, 'Sindicato de distribuidoras de combustível (filiado à Brasilcom)', REGIO, '', S, '', ''))
REGS += [
 ('Material de construção', 'ACOMACs e FECOMACs', 'Associações dos Comerciantes de Material de Construção (38) e federações (3)', 'vários',
  'Lojas de material de construção de cada região (ligadas à ANAMACO)', REGIO, '', S, '', 'exemplos nas linhas seguintes'),
 ('Material de construção', 'ACOMAC PE', 'Associação dos Comerciantes de Materiais de Construção de Pernambuco', 'PE', 'Lojas de material de construção de PE',
  REGIO, '', S, '', 'organiza o Ecomac Nordeste'),
 ('Material de construção', 'ACOMAC JUNDIAÍ', 'Associação dos Comerciantes de Material de Construção de Jundiaí', 'SP (Jundiaí)', 'Lojas de material de construção da região',
  REGIO, '', S, '', 'fundada em 1968'),
 ('Material de construção', 'SINCOMAVI', 'Sindicato do Comércio Varejista de Material de Construção da Grande São Paulo', 'SP (Grande SP)',
  'Lojas de material de construção da Grande SP', REGIO, '', S, '', 'apoiou a criação da Rede Construir (1997)'),
 ('Farmácias', 'SINCOFARMA-SP', 'Sindicato do Comércio Varejista de Produtos Farmacêuticos no Estado de São Paulo', 'SP', 'Farmácias e drogarias de SP',
  REGIO, '', S, '', 'filiado à Fecomercio SP'),
 ('Farmácias', 'SINCOFARMA-CE', 'Sindicato do Comércio Varejista de Produtos Farmacêuticos do Estado do Ceará', 'CE', 'Farmácias e drogarias do CE', REGIO, '', S, '', ''),
 ('Farmácias', 'SINCOFAC', '', 'AC', 'Farmácias do Acre', REGIO, '', S, '', 'nome completo não conferido'),
 ('Óticas', 'SINDIÓPTICA', '', 'SP', 'Sindicato patronal do varejo óptico de SP', REGIO, '', S, '', ''),
 ('Estacionamento', 'SINDEPARK', 'Sindicato das Empresas de Garagens e Estacionamentos do Estado de São Paulo', 'SP', 'Estacionamentos de SP', CONSULT, '', S, '', 'não publica associados'),
 ('Estacionamento', 'SINDEPARK/RS', 'Sindicato das Empresas de Garagens, Estacionamento e de Limpeza e Conservação de Veículos no Estado do Rio Grande do Sul',
  'RS', 'Estacionamentos do RS', REGIO, '', S, '', ''),
 ('Estacionamento', 'SIESGAVAMO', '', 'SP (São José do Rio Preto)', 'Sindicato patronal de garagens e estacionamentos', REGIO, '', S, '', 'base municipal'),
 ('Alimentação', 'ABRASEL (seccionais)', 'Seccionais estaduais da Associação Brasileira de Bares e Restaurantes', 'vários', 'Bares e restaurantes de cada estado',
  REGIO, '', N, '', ''),
 ('Rede de lojas', 'CDLs', 'Câmaras de Dirigentes Lojistas (ligadas à CNDL)', 'vários', 'Lojistas de cada cidade', REGIO, '', N, '', ''),
 ('Rede de lojas', 'FECOMÉRCIOS', 'Federações do Comércio de Bens, Serviços e Turismo de cada estado', 'vários', 'Comércio de cada estado (sindicatos patronais)', REGIO, '', N, '', ''),
]
CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
def cell(ref, v, s):
    if v is None or v == "": return ""
    st = f' s="{s}"' if s else ''
    if isinstance(v, (int, float)): return f'<c r="{ref}"{st}><v>{v}</v></c>'
    return f'<c r="{ref}"{st} t="inlineStr"><is><t xml:space="preserve">{escape(CTRL.sub("", str(v)))}</t></is></c>'
def row_xml(k, vals, ht=None):
    h = f' ht="{ht}" customHeight="1"' if ht else ''
    return f'<row r="{k}" spans="1:{len(vals)}"{h}>' + "".join(cell(f"{get_column_letter(j + 1)}{k}", v, s) for j, (v, s) in enumerate(vals)) + "</row>"
linhas = [('NACIONAL', seg, sig, nome, 'BR', '', oq, sit, onde, conf, site, obs) for seg, sig, nome, oq, sit, onde, conf, site, obs in NAC]
for seg, sig, nome, uf, oq, sit, onde, conf, site, obs in REGS:
    linhas.append(('REGIONAL', seg, sig, nome, uf, REG.get(uf[:2], '') if uf[:2] in REG else '', oq, sit, onde, conf, site, obs))
siglas = [l[2] for l in linhas]
assert len(siglas) == len(set(siglas)), [s for s in siglas if siglas.count(s) > 1]
out = [row_xml(1, [(h, S_HEAD) for h, w in COLS], ht=45)]
for i, l in enumerate(linhas, start=1):
    out.append(row_xml(i + 1, [(i, S_NUM)] + [(v, S_TXT) for v in l]))
dim = f"A1:{get_column_letter(len(COLS))}{len(linhas) + 1}"
corpo = ("<cols>" + "".join(f'<col min="{j + 1}" max="{j + 1}" width="{w}" customWidth="1"/>' for j, (h, w) in enumerate(COLS)) + "</cols>"
         + "<sheetData>" + "".join(out) + "</sheetData>" + f'<autoFilter ref="{dim}"/>')
VIEW = ('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>')
def caminho(z, nome):
    wb = z.read('xl/workbook.xml').decode('utf8'); rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    rid = re.search(r'<sheet name="' + re.escape(escape(nome)) + r'" sheetId="\d+"(?: state="\w+")? r:id="(rId\d+)"', wb).group(1)
    t = re.search(r'<Relationship Id="' + rid + r'"[^>]*Target="([^"]+)"', rels) or re.search(r'Target="([^"]+)"[^>]*Id="' + rid + '"', rels)
    return 'xl/' + t.group(1).lstrip('/').replace('xl/', '')
z = zipfile.ZipFile(MAPA); novos = {}
wb = z.read('xl/workbook.xml').decode('utf8')
assert f'<sheet name="{NOVA}"' not in wb, 'a aba já existe'
nums = [int(n) for n in re.findall(r'xl/worksheets/sheet(\d+)\.xml', ' '.join(z.namelist()))]
parte = f'xl/worksheets/sheet{max(nums) + 1}.xml'
rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
rid = 'rId' + str(max(int(n) for n in re.findall(r'Id="rId(\d+)"', rels)) + 1)
sid = max(int(n) for n in re.findall(r'sheetId="(\d+)"', wb)) + 1
modelo = z.read(caminho(z, 'ESTACIONAMENTO FINAL')).decode('utf8')
cab = re.sub(r'xr:uid="\{[^}]+\}"', 'xr:uid="{6B2D9E41-3A7C-4F15-9E08-2C4B7A1D5F63}"', modelo[:modelo.index('<dimension')])
sx = (cab + f'<dimension ref="{dim}"/>{VIEW}<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>' + corpo
      + '<pageMargins left="0.511811024" right="0.511811024" top="0.78740157499999996" bottom="0.78740157499999996" header="0.31496062000000002" footer="0.31496062000000002"/></worksheet>')
novos[parte] = sx.encode('utf8')
novos['xl/workbook.xml'] = wb.replace('</sheets>', f'<sheet name="{NOVA}" sheetId="{sid}" r:id="{rid}"/></sheets>').encode('utf8')
novos['xl/_rels/workbook.xml.rels'] = rels.replace('</Relationships>', f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{parte.split("/")[-1]}"/></Relationships>').encode('utf8')
ct = z.read('[Content_Types].xml').decode('utf8')
novos['[Content_Types].xml'] = ct.replace('</Types>', f'<Override PartName="/{parte}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>').encode('utf8')
app = z.read('docProps/app.xml').decode('utf8')
m = re.search(r'<vt:lpstr>Planilhas</vt:lpstr></vt:variant><vt:variant><vt:i4>(\d+)</vt:i4>', app)
if m:
    n = int(m.group(1)); app = app.replace(m.group(0), m.group(0).replace(f'<vt:i4>{n}</vt:i4>', f'<vt:i4>{n + 1}</vt:i4>'))
app = re.sub(r'<TitlesOfParts><vt:vector size="(\d+)"', lambda q: f'<TitlesOfParts><vt:vector size="{int(q.group(1)) + 1}"', app)
app = app.replace('</vt:vector></TitlesOfParts>', f'<vt:lpstr>{NOVA}</vt:lpstr></vt:vector></TitlesOfParts>')
novos['docProps/app.xml'] = app.encode('utf8')
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zo:
    for it in z.infolist():
        zo.writestr(it, novos.pop(it.filename, None) or z.read(it.filename))
    zo.writestr(parte, novos.pop(parte))
assert not novos, novos.keys()
import collections
print('nova aba', NOVA, parte, len(linhas), 'linhas:', dict(collections.Counter(l[0] for l in linhas)), dict(collections.Counter(l[7] for l in linhas)))
