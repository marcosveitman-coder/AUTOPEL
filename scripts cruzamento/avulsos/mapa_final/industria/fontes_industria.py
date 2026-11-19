"""Aba FULLFILMIT_INDUSTRIA FINAL: listas nominais das fontes de mercado (19/11/2026). Gera fontes.json.
Cada item: nome da empresa como a fonte publica, segmento, fonte, palavras para achar a razão social na Receita
(todas precisam estar na razão) e a UF da sede quando a fonte diz.
Fontes:
  EXAME   Melhores e Maiores 2026 (receita 2025), exame.com/mm/maiores: só os setores industriais pedidos. Ficam fora
          distribuidoras de combustível e gás, petroleiras de exploração, mineração/siderurgia, usinas de etanol puras,
          tradings de grãos, varejistas e restaurantes.
  ANFAVEA associadas (anfavea.com.br/site/associadas), ELETROS associadas (eletros.org.br), ANIP fabricantes de pneus,
  LUBES (Lubes em Foco, mercado de lubrificantes 2025), PBE/INMETRO (tabelas ENCE 2025: marcas com mais modelos).
Uso: python fontes_industria.py <pasta de trabalho>"""
import json, os, sys

F = []


def add(fonte, seg, nome, termos, uf=''):
    F.append(dict(fonte=fonte, seg=seg, nome=nome, termos=termos.split('|'), uf=uf))


# ---------- EXAME (setores industriais) ----------
for n, t, uf in [('Braskem', 'BRASKEM S A', 'BA'), ('BASF', 'BASF S A', 'SP'), ('White Martins', 'WHITE MARTINS GASES', 'RJ'),
                 ('Unipar Carbocloro', 'UNIPAR CARBOCLORO', 'SP'), ('Innova', 'VIDEOLAR INNOVA', 'AM'), ('Vipal Borrachas', 'BORRACHAS VIPAL S A', 'RS'),
                 ('Química Anastácio', 'QUIMICA ANASTACIO', 'SP'), ('Arlanxeo', 'ARLANXEO BRASIL', 'RJ'), ('Nitro Química', 'NITRO QUIMICA', 'SP'),
                 ('Moeve Química', 'MOEVE QUIMICA', 'BA'), ('Grupo Krona', 'KRONA TUBOS', 'SC'), ('Elekeiroz', 'ELEKEIROZ', 'SP'),
                 ('GPC Química', 'GPC QUIMICA', 'PR'), ('Sansuy', 'SANSUY', 'SP'), ('Tronox Pigmentos', 'TRONOX PIGMENTOS', 'BA'),
                 ('Borealis Brasil', 'BOREALIS', 'SP'), ('Fábrica Carioca de Catalisadores', 'FABRICA CARIOCA DE CATALISADORES', 'RJ'),
                 ('Suzano', 'SUZANO S A', 'BA'), ('Klabin', 'KLABIN S A', 'SP'), ('Santher', 'SANTHER', 'SP'), ('Irani', 'IRANI PAPEL', 'RS'),
                 ('Grupo Penha', 'PENHA PAPEIS', 'SP'), ('Ibema', 'IBEMA', 'PR'), ('Santa Maria Papel', 'SANTA MARIA CIA DE PAPEL', 'PR'),
                 ('Polo Films', 'POLO FILMS', 'SP'), ('Cia Canoinhas', 'CANOINHAS DE PAPEL', 'SC'), ('Bragagnolo', 'BRAGAGNOLO', 'SC'),
                 ('Brasilata', 'BRASILATA', 'SP'), ('UPL', 'UPL DO BRASIL', 'SP'), ('Iharabras', 'IHARABRAS', 'SP'), ('Adama', 'ADAMA BRASIL', 'PR'),
                 ('Vittia', 'VITTIA S A', 'SP'), ('CP Kelco', 'CP KELCO BRASIL', 'SP')]:
    add('EXAME', 'QUIM', n, t, uf)
for n, t, uf in [('JBS (Friboi / Seara)', 'JBS S A', 'SP'), ('MBRF (Marfrig + BRF)', 'BRF S A', 'SC'), ('Cargill', 'CARGILL ALIMENTOS', 'SP'),
                 ('Ambev', 'AMBEV S A', 'SP'), ('Minerva Foods', 'MINERVA S A', 'SP'), ('Louis Dreyfus (LDC Sucos)', 'LOUIS DREYFUS', 'SP'),
                 ('3corações', 'TRES CORACOES ALIMENTOS', 'CE'), ('Solar Coca-Cola', 'SOLAR BEBIDAS', 'CE'), ('M. Dias Branco', 'M DIAS BRANCO', 'CE'),
                 ('Laticínios Bela Vista (Piracanjuba)', 'LATICINIOS BELA VISTA', 'GO'), ('Norsa Coca-Cola', 'NORSA REFRIGERANTES', 'CE'),
                 ('Brejeiro Alimentos', 'BREJEIRO', 'SP'), ('Mataboi Alimentos', 'MATABOI', 'MG'), ('Bem Brasil Alimentos', 'BEM BRASIL ALIMENTOS', 'MG'),
                 ('Estrela Alimentos', 'ALIMENTOS ESTRELA', 'SP'), ('J. Macêdo', 'J MACEDO', 'CE'), ('Barra Mansa Alimentos', 'BARRA MANSA', 'SP'),
                 ('Kraft Heinz', 'HEINZ BRASIL', 'SP'), ('Granja Faria', 'GRANJA FARIA', 'SC'), ('Adoro', 'ADORO S A', 'SP'), ('Alibem', 'ALIBEM', 'RS'),
                 ('Pamplona', 'PAMPLONA ALIMENTOS', 'SC'), ('Café Cacique', 'CACIQUE DE CAFE', 'PR'), ('Pastifício Selmi', 'PASTIFICIO SELMI', 'SP'),
                 ('Marilan', 'MARILAN ALIMENTOS', 'SP'), ('São Braz', 'SAO BRAZ S A', 'PA'), ('GDC Alimentos', 'GDC ALIMENTOS', 'GO'),
                 ('Leite Jussara', 'JUSSARA', 'SP'), ('Orquídea Alimentos', 'ORQUIDEA', 'RS'), ('Vidroporto', 'VIDROPORTO', 'SP'),
                 ('Cervejaria Cidade Imperial', 'CERVEJARIA CIDADE IMPERIAL', 'MG'), ('Vilma Alimentos', 'VILMA ALIMENTOS', 'MG'),
                 ('Anaconda', 'ANACONDA', 'SP'), ('Cia Müller de Bebidas', 'MULLER DE BEBIDAS', 'SP'), ('Mogiana Alimentos', 'MOGIANA ALIMENTOS', 'SP'),
                 ('Moinho Paulista', 'MOINHO PAULISTA', 'SP'), ('Natural One', 'NATURAL ONE', 'SP'), ('Santa Helena', 'SANTA HELENA INDUSTRIA DE ALIMENTOS', 'SP'),
                 ('Ocrim', 'OCRIM', 'SP'), ('Peccin', 'PECCIN', 'RS'), ('Yakult', 'YAKULT', 'SP'), ('Moinho Globo', 'MOINHO GLOBO', 'RJ'),
                 ('Real Café', 'REAL CAFE', 'ES'), ('Bunge', 'BUNGE ALIMENTOS', 'SC'), ('3tentos', 'TRES TENTOS', 'RS'), ('Camil', 'CAMIL ALIMENTOS', 'SP'),
                 ('Caramuru Alimentos', 'CARAMURU ALIMENTOS', 'GO'), ('Bianchini', 'BIANCHINI', 'RS'), ('Frigol', 'FRIGOL', 'SP'),
                 ('Josapar (Tio João)', 'JOSAPAR', 'RS'), ('Rivelli', 'RIVELLI', 'MG'), ('Grande Moinho Cearense', 'GRANDE MOINHO CEARENSE', 'CE'),
                 ('Leitesol', 'LEITESOL', 'SP'), ('Baldo', 'BALDO S A', 'RS'), ('Cooxupé', 'COOXUPE', 'MG'), ('Cocamar', 'COCAMAR', 'PR'),
                 ('Copacol', 'COPACOL', 'PR'), ('Aurora / Cooperalfa', 'COOPERALFA', 'SC'), ('Integrada', 'INTEGRADA COOPERATIVA', 'PR'),
                 ('Castrolanda', 'CASTROLANDA', 'PR'), ('Frimesa', 'FRIMESA', 'PR'), ('Natura', 'NATURA COSMETICOS', 'SP'),
                 ('Flora (J&F)', 'FLORA PRODUTOS DE HIGIENE', 'SP'), ('Granado', 'GRANADO', 'RJ')]:
    add('EXAME', 'FMCG', n, t, uf)
for n, t, uf in [('Eurofarma', 'EUROFARMA LABORATORIOS', 'SP'), ('EMS', 'EMS S A', 'SP'), ('Hypera', 'HYPERA S A', 'SP'), ('Bayer', 'BAYER S A', 'RJ'),
                 ('Novartis', 'NOVARTIS', 'SP'), ('Aché', 'ACHE LABORATORIOS', 'SP'), ('Roche', 'PRODUTOS ROCHE', 'SP'),
                 ('União Química', 'UNIAO QUIMICA FARMACEUTICA', 'SP'), ('Cimed', 'CIMED INDUSTRIA', 'SP'), ('Merck', 'MERCK S A', 'RJ'),
                 ('Bionovis', 'BIONOVIS', 'SP'), ('Apsen', 'APSEN', 'SP'), ('Blau Farmacêutica', 'BLAU FARMACEUTICA', 'SP'),
                 ('Ourofino', 'OURO FINO SAUDE ANIMAL', 'SP'), ('Zodiac', 'ZODIAC', 'SP'), ('Supera Farma', 'SUPERA FARMA', 'SP'),
                 ('Althaia', 'ALTHAIA', 'SP'), ('Blanver', 'BLANVER', 'SP'), ('Lafepe', 'LABORATORIO FARMACEUTICO DO ESTADO DE PERNAMBUCO', 'PE')]:
    add('EXAME', 'FARMA', n, t, uf)
for n, t, uf in [('Guararapes (Riachuelo)', 'GUARARAPES CONFECCOES', 'RN'), ('Cidade Maravilhosa', 'CIDADE MARAVILHOSA', 'RJ'),
                 ('Alpargatas', 'ALPARGATAS S A', 'SP'), ('Calçados Beira Rio', 'CALCADOS BEIRA RIO', 'RS'), ('Vulcabras', 'VULCABRAS S A', 'SP'),
                 ('Grendene', 'GRENDENE', 'CE'), ('Vicunha Têxtil', 'VICUNHA TEXTIL', 'CE'), ('Lunelli', 'LUNELLI INDUSTRIA', 'SC'),
                 ('Pettenati', 'PETTENATI', 'RS'), ('Dakota Calçados', 'DAKOTA CALCADOS', 'RS'), ('Azzas 2154 (Arezzo&Co)', 'AZZAS 2154', 'RS')]:
    add('EXAME', 'TEXTIL', n, t, uf)
for n, t, uf, s in [('WEG', 'WEG EQUIPAMENTOS ELETRICOS', 'SC', 'ELETRO'), ('Iochpe-Maxion', 'IOCHPE MAXION', 'SP', 'AUTO'),
                    ('Randoncorp', 'RANDONCORP', 'RS', 'AUTO'), ('Whirlpool (Brastemp / Consul)', 'WHIRLPOOL S A', 'SP', 'ELETRO'),
                    ('Tupy', 'TUPY S A', 'SC', 'AUTO'), ('Marcopolo', 'MARCOPOLO', 'RS', 'AUTO'), ('Mahle', 'MAHLE METAL LEVE', 'SP', 'AUTO'),
                    ('Gestamp', 'GESTAMP BRASIL', 'PR', 'AUTO'), ('Elgin', 'ELGIN', 'SP', 'ELETRO'), ('Facchini', 'FACCHINI', 'SP', 'AUTO'),
                    ('Metalfrio', 'METALFRIO', 'SP', 'ELETRO'), ('Lorenzetti', 'LORENZETTI', 'SP', 'ELETRO'), ('Nadir Figueiredo', 'NADIR FIGUEIREDO', 'SP', 'CONST'),
                    ('Agrale', 'AGRALE', 'RS', 'AUTO'), ('Plascar', 'PLASCAR', 'SP', 'AUTO'), ('S Riko', 'S RIKO AUTOMOTIVE HOSE TECALON', 'RJ', 'AUTO'),
                    ('Mangels', 'MANGELS', 'MG', 'AUTO'), ('Mundial', 'MUNDIAL S A PRODUTOS DE CONSUMO', 'SP', 'FMCG'), ('Sika', 'SIKA S A', 'SP', 'CONST'),
                    ('Tramontina', 'TRAMONTINA', 'RS', 'CONST'), ('Panatlântica', 'PANATLANTICA', 'RS', 'CONST'),
                    ('Tuper', 'TUPER', 'SC', 'AUTO'), ('Schulz', 'SCHULZ', 'SC', 'ELETRO')]:
    add('EXAME', s, n, t, uf)
for n, t, uf in [('Petrobras (Lubrax via Vibra / refino)', 'PETROLEO BRASILEIRO', 'RJ'), ('Vibra Energia (Lubrax)', 'VIBRA ENERGIA', 'RJ'),
                 ('Ipiranga (Ipiranga Lubrificantes)', 'IPIRANGA PRODUTOS DE PETROLEO', 'RJ'), ('Moove (Mobil, Cosan)', 'COSAN LUBRIFICANTES', 'RJ'),
                 ('Iconic Lubrificantes (Texaco / Ipiranga)', 'ICONIC LUBRIFICANTES', 'RJ'), ('Petronas Lubrificantes', 'PETRONAS LUBRIFICANTES', 'MG'),
                 ('Refinaria Riograndense', 'REFINARIA DE PETROLEO RIOGRANDENSE', 'RS'), ('Petrocoque', 'PETROCOQUE', 'SP')]:
    add('EXAME/LUBES', 'OLEO', n, t, uf)
# ---------- LUBES (Lubes em Foco 2025) ----------
for n, t in [('Blueway (Shell, Raízen)', 'RAIZEN LUBRIFICANTES|SHELL LUBRIFICANTES'), ('Castrol', 'CASTROL'), ('Quaker Houghton', 'QUAKER'),
             ('Ultrax', 'ULTRAX'), ('YPF Brasil', 'YPF'), ('TotalEnergies Lubrificantes', 'TOTALENERGIES'), ('Lwart (rerrefino)', 'LWART')]:
    add('LUBES', 'OLEO', n, t)
# ---------- ANFAVEA associadas ----------
for n, t in [('AGCO', 'AGCO DO BRASIL'), ('Agrale', 'AGRALE'), ('Audi', 'AUDI DO BRASIL'), ('BMW', 'BMW DO BRASIL'), ('CAOA', 'CAOA MONTADORA'),
             ('Caterpillar', 'CATERPILLAR BRASIL'), ('CNH Industrial', 'CNH INDUSTRIAL'), ('DAF', 'DAF CAMINHOES'), ('Ford', 'FORD MOTOR COMPANY BRASIL|FORD BRASIL'),
             ('General Motors', 'GENERAL MOTORS DO BRASIL'), ('Honda Automóveis', 'HONDA AUTOMOVEIS DO BRASIL'), ('HPE Automotores (Mitsubishi)', 'HPE AUTOMOTORES'),
             ('Hyundai', 'HYUNDAI MOTOR BRASIL'), ('Iveco', 'IVECO'), ('Jaguar Land Rover', 'JAGUAR'), ('JCB', 'JCB DO BRASIL'), ('Komatsu', 'KOMATSU DO BRASIL'),
             ('Mercedes-Benz', 'MERCEDES BENZ DO BRASIL'), ('Mercedes-Benz Cars & Vans', 'MERCEDES BENZ CARS'), ('Nissan', 'NISSAN DO BRASIL'),
             ('Renault Geely', 'RENAULT DO BRASIL'), ('Scania', 'SCANIA'), ('Stellantis', 'STELLANTIS AUTOMOVEIS'), ('Toyota', 'TOYOTA DO BRASIL'),
             ('Volkswagen', 'VOLKSWAGEN DO BRASIL'), ('Volkswagen Caminhões e Ônibus', 'VOLKSWAGEN TRUCK'), ('Volvo', 'VOLVO DO BRASIL'),
             ('BYD (não associada, fábrica em Camaçari)', 'BYD AUTO DO BRASIL'), ('GWM (não associada, fábrica em Iracemápolis)', 'GWM')]:
    add('ANFAVEA', 'AUTO', n, t)
# ---------- ANIP (pneus) ----------
for n, t in [('Bridgestone', 'BRIDGESTONE'), ('Continental Pneus', 'CONTINENTAL DO BRASIL PRODUTOS AUTOMOTIVOS|CONTINENTAL PNEUS'),
             ('Dunlop (Sumitomo)', 'SUMITOMO RUBBER'), ('Goodyear', 'GOODYEAR DO BRASIL'), ('Maggion', 'MAGGION'), ('Michelin', 'MICHELIN'),
             ('Pirelli', 'PIRELLI PNEUS'), ('Prometeon', 'PROMETEON'), ('Rinaldi', 'RINALDI'), ('Titan Pneus', 'TITAN PNEUS'), ('Tortuga', 'TORTUGA')]:
    add('ANIP', 'PNEU', n, t)
# ---------- ELETROS associadas ----------
for n, t in [('Belmicro', 'BELMICRO'), ('Bosch (BSH)', 'BSH CONTINENTAL'), ('Electrolux', 'ELECTROLUX'), ('Elgin', 'ELGIN'), ('Elsys', 'ELSYS'),
             ('Friovix', 'FRIOVIX'), ('Gree', 'GREE ELECTRIC'), ('Grupo Multi', 'GRUPO MULTILASER'), ('Hisense', 'HISENSE'), ('LG Electronics', 'LG ELECTRONICS DO BRASIL'),
             ('Midea Carrier', 'MIDEA CARRIER'), ('Mondial', 'MONDIAL'), ('Panasonic', 'PANASONIC DO BRASIL'), ('Philco / Britânia', 'BRITANIA ELETRODOMESTICOS'),
             ('Positivo', 'POSITIVO TECNOLOGIA'), ('Samsung', 'SAMSUNG ELETRONICA DA AMAZONIA'), ('TCL Semp', 'SEMP AMAZONAS|TCL SEMP'),
             ('TPV (Philips TV / AOC)', 'TPV DO BRASIL|ENVISION'), ('Whirlpool', 'WHIRLPOOL S A'), ('Newell Brands (Oster)', 'NEWELL'),
             ('Black & Decker', 'STANLEY BLACK'), ('Philips Versuni (Walita)', 'VERSUNI|PHILIPS DO BRASIL'), ('Esmaltec', 'ESMALTEC'), ('Mallory', 'MALLORY'),
             ('Atlas Eletrodomésticos', 'ATLAS ELETRODOMESTICOS'), ('Itatiaia', 'ITATIAIA MOVEIS'), ('Imbera', 'IMBERA'), ('Culligan', 'CULLIGAN'), ('Mueller', 'MUELLER')]:
    add('ELETROS', 'ELETRO', n, t)
# ---------- PBE / INMETRO (marcas com mais modelos etiquetados, tabelas 2025) ----------
for n, t, sel in [('Komeco', 'KOMECO', 'ENCE aquecedor a gás / Selo CONPET'), ('Rinnai', 'RINNAI', 'ENCE aquecedor a gás / Selo CONPET'),
                  ('Lorenzetti', 'LORENZETTI', 'ENCE aquecedor a gás'), ('Rheem', 'RHEEM', 'ENCE aquecedor a gás'),
                  ('Famac', 'FAMAC', 'ENCE bombas / Selo PROCEL'), ('Dancor', 'DANCOR', 'ENCE bombas / Selo PROCEL'), ('Grundfos', 'GRUNDFOS', 'ENCE bombas'),
                  ('Thebe', 'THEBE', 'ENCE bombas'), ('Ebara', 'EBARA', 'ENCE bombas'), ('Jacuzzi', 'JACUZZI', 'ENCE bombas'), ('Franklin Electric', 'FRANKLIN ELECTRIC', 'ENCE bombas'),
                  ('Schneider Motobombas', 'SCHNEIDER MOTOBOMBAS', 'ENCE bombas'), ('Gree', 'GREE ELECTRIC', 'ENCE ar-condicionado / Selo PROCEL'),
                  ('Agratto', 'AGRATTO', 'ENCE ar-condicionado'), ('Springer Carrier', 'SPRINGER CARRIER|MIDEA CARRIER', 'ENCE ar-condicionado'),
                  ('Elgin', 'ELGIN', 'ENCE ar-condicionado'), ('Daikin', 'DAIKIN', 'ENCE ar-condicionado'), ('Itatiaia', 'ITATIAIA', 'ENCE fogão / Selo CONPET'),
                  ('Esmaltec', 'ESMALTEC', 'ENCE fogão e refrigerador / Selo CONPET'), ('Atlas', 'ATLAS ELETRODOMESTICOS', 'ENCE fogão / Selo CONPET'),
                  ('Dako', 'DAKO', 'ENCE fogão / Selo CONPET'), ('Suggar', 'SUGGAR', 'ENCE fogão'), ('Venax', 'VENAX', 'ENCE fogão e refrigerador'),
                  ('Realce', 'REALCE', 'ENCE fogão'), ('Clarice', 'CLARICE ELETRODOMESTICOS', 'ENCE fogão'), ('Fischer', 'FISCHER', 'ENCE fogão')]:
    add('PBE', 'ELETRO', n, t)
    F[-1]['selo'] = sel

json.dump(F, open(os.path.join(sys.argv[1], 'fontes.json'), 'w', encoding='utf8'), ensure_ascii=False, indent=0)
print(len(F))
