# Sessão 05/11/2026 — abas PET SHOP  FINAL e ESTACIONAMENTO FINAL (MAPA MERCADO)

## Pedido
O usuário criou as abas **PET SHOP  FINAL** (com dois espaços no nome) e **ESTACIONAMENTO FINAL**. Pediu:
- o mesmo racional e as mesmas colunas das abas ALIMENTACAO, REDE DE LOJAS e CONSTRUÇÃO FINAL: nome fantasia, razão
  social, grupo administrador, nº de lojas e CNPJ ("informação crucial");
- completar as listas pelas associações: IPB, ABINPET, ABLA, ABRASE e ANCLIVEPA (pet); ABRAPARK e Sindepark
  (estacionamento);
- usar o Google e pesquisar por região (Norte, Nordeste, Centro-Oeste, Sudeste e Sul), para não perder redes
  regionais;
- "não podemos errar".

## Entregas (MAPA MERCADO)

| Aba | O que tem |
|---|---|
| PET SHOP  FINAL | 110 redes (A–W): 99 CNPJ confirmado, 7 provável e 4 franquias com CNPJ de uma unidade (franqueadora não achada) |
| ESTACIONAMENTO FINAL | 133 redes: 120 confirmado, 5 provável, 7 não confirmado e 1 não encontrada |
| MAIORES POR REGIÃO | Acrescentados no fim (linhas 581–839) os blocos de pet e de estacionamento: 30 maiores do Brasil e 20 por região |
| R_E_S_U_M_O | Acrescentada a seção 8 no fim (linhas 74–113): passo a passo, níveis do CNPJ, fontes e maiores |

**Redes por região da sede:**
- **Pet:** Sudeste 45, Sul 22, Nordeste 17, Centro-Oeste 14 e Norte 12.
- **Estacionamento:** Sudeste 73, Sul 21, Nordeste 21, Centro-Oeste 14 e Norte 4.

**Colunas:**
- A–T: as mesmas das outras abas FINAL.
- U–W: as que o usuário criou nas abas antigas.
  - U (STATUS RECEITA) vem da Receita.
  - V (PORTE) e W (STATUS AUTOPEL) foram copiadas só das redes pet que já estavam na REDE DE LOJAS FINAL. Nas outras
    linhas ficaram em branco, porque a regra do usuário para essas colunas não está documentada.

## Como foi feito
1. **Origem.**
   - Estacionamentos: tabela H–K da aba de pesquisa de mercado (aba "Alessandra Dayan"), com 50 empresas.
   - Pet:
     - as 6 redes pet que já estavam na REDE DE LOJAS FINAL, com o mesmo nº de lojas e a mesma fonte;
     - a aba IRTT;
     - a Pet Anjo (mesma aba de pesquisa, coluna BK).
2. **Receita por CNAE** (`redes_pet_estac.py`, cerca de 13 min):
   - pet shop 4789-0/04, veterinária 7500-1/00, banho e tosa 9609-2/08, hotel 9609-2/07, farmácia veterinária
     4771-7/04 e estacionamento 5223-1/00;
   - resultado: 644 empresas pet e 1.063 de estacionamento com 3 ou mais unidades ativas.
3. **Busca por nome e por raiz** (`busca_receita.py`, consultas q12 a q16): razão social, matriz e unidades de cada
   candidata.
4. **Imprensa e listas:**
   - Revista PetCenter, "100 Top Pet Shops das maiores cidades" (mar/2025): redes por cidade e UF;
   - IRTT 2026: Petz 265, Cobasi 256, Petlove 30;
   - ABF/Portal do Franchising: Petland 168, Pet Anjo 275, Pop Pet Center 135, Fórmula Animal 92;
   - Estapar, prévia do 2T26: 851 operações;
   - Forbes e IstoÉ Dinheiro: Indigo com 320+ operações (nov/2025).
5. **CNPJ rede a rede**, à mão (`dados/mk_dec_pet_estac.py`): a mesma regra das outras abas — operadora, franqueadora
   ou dona da marca, nunca franqueado.
6. **Complementos regionais da Receita:**
   - pet: 8+ unidades, ou entre as 15 maiores da região com 4+;
   - estacionamento: 15+ garagens, ou entre as 15 maiores da região com 5+.

**Gravação** (`grava_pet_estac.py`):
- as 2 abas novas foram gravadas pelo XML, e nas abas MAIORES POR REGIÃO e R_E_S_U_M_O só se acrescentou no fim;
- conferência com `confere_pet_estac.py`:
  - as outras 40 abas ficaram iguais célula a célula, e a MAIORES e a R_E_S_U_M_O iguais até a última linha antiga;
  - só mudaram as 4 partes dessas abas;
  - dinâmicas 21 = 21, caches 9 = 9, nomes definidos iguais;
  - todo CNPJ tem dígito válido e está ativo na Receita local, sem nome repetido;
- a cópia abriu no Excel (42 abas, 21 dinâmicas); o original foi trocado depois de conferir que estava igual ao backup.

## Decisões e regras
- **Uma linha por grupo:**
  - Indigo = Administradora Geral de Estacionamentos + PB Administradora (antiga PareBem) + Gepark + Center Park,
    todas com garagens de nome INDIGO. As linhas PB Estacionamentos e Gepark da origem entraram na Indigo;
  - a Rio Park da origem é a marca Riopark da Estapar;
  - a Netpark junta NAA e NEE, que a origem trazia à parte;
  - a Hora Park ficou em linha própria, com grupo Estapar (36 das suas garagens usam o nome ESTAPAR).
- **Sede da Indigo:** a matriz da empresa é em Porto Alegre (antiga Moving), por isso a região da sede é Sul. A
  observação avisa que o escritório da Indigo Brasil é em São Paulo.
- **Franquias sem franqueadora com o nome:** Petland, Rede Pop Pet Center, Apaixonados por 4 Patas e DrogaVet ficam com
  o CNPJ de uma unidade, marcado como tal. A franqueadora da DrogaVet está baixada.
- **Fora da aba de pet shop:**
  - agropecuárias, cooperativas e lojas rurais (farmácia veterinária sem nome de pet). Exemplo: uma rede agropecuária
    de 98 lojas;
  - hospitais universitários e franqueados de redes já listadas.
- **Fora da aba de estacionamento:** pátios de veículos apreendidos, guinchos, transportadoras, hotéis e hospitais.
- **Sem CNPJ (origem):**
  - empresa baixada ou inapta: Wellpark, Servi Park e Simopark;
  - nome comum, sem rede: Brasil Park, Center Park e Park Service;
  - sem empresa com o nome: Vezani e Sandrini;
  - Maxipark (imprensa de 2018): nenhuma empresa ativa com o nome.
- **Associações:**
  - IPB, ABINPET (indústria), ANCLIVEPA, ABRAPARK e Sindepark não publicam lista de redes;
  - ABLA e ABRASE não são entidades do setor pet (ABLA é a associação das locadoras de automóveis).

## Pendências
- Colunas V (PORTE) e W (STATUS AUTOPEL) das linhas novas: o usuário precisa dizer como as preencheu nas outras abas,
  ou pedir o cruzamento com o faturamento.
- Confirmar as franqueadoras de Petland, Pop Pet Center, Apaixonados por 4 Patas e DrogaVet (COF ou contato).
- Petz, Cobasi, Pet Camp, American Pet, Pet Fun e Petlove estão também na REDE DE LOJAS FINAL. O usuário decide se saem
  de lá.

## Lições
- **Não regravar abas FINAL antigas:** elas têm colunas do usuário (U–W). Aba nova se grava sozinha. Em aba de resumo,
  só se acrescenta no fim.
- **Corte regional "top N com X+":** ao excluir uma candidata, entra a próxima, que pode não ter dados na busca. Buscar
  de uma vez todas as candidatas acima do corte mínimo (`q16`).
- **CNPJ digitado à mão:** usar só a raiz e deixar a matriz vir da Receita. O build ignora dígitos digitados que não
  existem.
- **Bash com aspas:** texto com apóstrofo dentro de heredoc ou de `python -c` quebra o comando. Gravar o script com a
  ferramenta de escrita de arquivo.
- **BrasilAPI:** pode ficar fora do ar (serviço indisponível ou excesso de pedidos). A cópia local da Receita resolve.

## 2ª parte: conferência da origem dos estacionamentos
As 50 empresas da tabela H–K da aba de pesquisa (linhas 3 a 52) estão todas na ESTACIONAMENTO FINAL. Nenhuma faltou.
- 44 têm linha própria.
- 6 entraram na linha do grupo: PB Estacionamentos e Gepark na Indigo, Rio Park na Estapar, NAA e NEE na Netpark.

## 3ª parte: ORGANOGRAMA impressao com qualidade de resumo executivo
**Pedido:** na impressão, "No segmento" quebrava em duas linhas e ficava cortado. Atualizar a aba e acabar com o
desleixo.

**Causas:**
1. **Letra mais larga que a estimativa:** a largura das letras era estimada. "No segmento" ocupava 93 px numa célula de
   102 px e quebrou na tela do usuário (escala de tela).
2. **Página mais larga que a área útil:** a impressora tem margem física de ~9 mm. A área útil é de ~757 pt e a aba
   tinha 808 pt. A última coluna ia para outra folha: eram **12 páginas, e não 6**. Isso já acontecia na versão
   anterior e não tinha sido conferido.

**Correção:**
- **Largura real das letras:**
  - a Verdana 10 foi medida no Windows (`mede_verdana.ps1` → `verdana10.txt`), com 10% de folga;
  - o `org_impressao.py` para com erro se alguma palavra não couber na célula.
- **Grade:** 60 colunas de 16 px, e as 3 colunas entre as caixas com 5 px. Largura total de 740 pt.
- **Caixas:** "No segmento" com a quantidade e o % em células separadas. Sem segmento, mostra "sem segmento".
- **Tabelas:** cabeçalhos sem negrito nas colunas estreitas e colunas de segmento e de pessoas mais largas.
- **Margens laterais:** 0,4 cm.

**Conferência no Excel:**
- 6 páginas, nenhuma quebra vertical, zoom 100%;
- imagens das 6 folhas sem texto cortado;
- as outras abas iguais célula a célula, dinâmicas 21 = 21;
- nos nomes definidos, só mudou a área de impressão.

Os dados do organograma foram refeitos a partir do arquivo atual e não mudaram.

**Lição:** conferir sempre `PageSetup.Pages.Count` e `VPageBreaks.Count` pelo COM, porque a margem física da impressora
reduz a área útil. A imagem por `CopyPicture` não mostra esse problema.

## 4ª parte: empresas da aba de pesquisa que não estão no mapeamento (vermelho)
**Pedido:** cruzar as empresas da aba "Alessandra Dayan" com todas as bases do arquivo e pintar de vermelho as que não
constam no mapeamento.

**Como foi feito** (`confere_alessandra.py`, decisões em `dados/dec_alessandra.json`):
- As 550 empresas das 11 tabelas da aba foram conferidas em três etapas:
  1. referência da célula na coluna ORIGEM das abas FINAL (299 casos);
  2. nome igual, sem palavras de tipo (106 casos);
  3. revisão à mão (5 casos).
- Só valem bases do mesmo segmento:
  - drogarias: FARMACIA VF e FARMACIA CONSOLIDADO;
  - supermercados: SUPERMERCADOS (reuniao) e IRTT;
  - beleza, óticas e serviços: REDE DE LOJAS FINAL e IRTT;
  - os outros segmentos: a aba FINAL do segmento.

  Sem essa trava, "Óticas São José" casava com "Farmácia São José".
- **Revisão à mão:**
  - estão no mapeamento: Mart Minas, Spani, Koch/Komprão, Bacio di Latte e OJO/Ray-Ban;
  - foram casos falsos: Óticas Martins, Raiz e Ipanema, que tinham batido com outras empresas.

**Resultado:** 140 empresas não estão no mapeamento. Elas foram pintadas de vermelho (fundo vermelho, letra branca)
nas 4 células de cada uma.
- Beleza/Saúde: 44.
- Óticas: 47.
- Serviços/Franquias: 49.
- Nas tabelas de drogarias, estacionamentos, postos, construção, alimentação, rede de lojas, supermercados e
  conveniência, nenhuma faltou.

**Gravação:**
- feita pelo Excel numa cópia (`pinta_alessandra.ps1`);
- conferência: as outras abas e os valores ficaram iguais, dinâmicas 21 = 21 e 560 células vermelhas.

**Pendência:** não há aba de mapeamento para beleza/estética, óticas e serviços/franquias. Criar essas abas, se o
usuário quiser, com o mesmo racional das abas FINAL.

## 5ª parte: aba OUTROS FINAL
**Pedido:** criar uma aba só, OUTROS FINAL, para as 140 empresas que não estavam no mapeamento. O modelo é o mesmo da
PET SHOP e da ESTACIONAMENTO FINAL, com o CNPJ e com segmento e subsegmento no fim.

**Entrega:** aba nova no fim da pasta, com 140 linhas e as colunas A–W das outras abas FINAL.
- **Colunas novas:**
  - X: SEGMENTO (PESQUISA), com Beleza / Saúde / Bem-estar, Óticas ou Serviços / Franquias;
  - Y: SUBSEGMENTO, com 19 subsegmentos, como Odontologia, Academias e pilates, Depilação e estética, Cosméticos e
    perfumaria, Idiomas e cursos, Lavanderias e Seguros e crédito.
- **CNPJ:**
  - 52 confirmados (franqueadoras com CNAE 7740-3/00 ou empresas da marca);
  - 31 prováveis;
  - 42 com CNPJ de uma unidade (franqueadora não achada);
  - 14 não confirmados (nome genérico ou franqueadora sem o nome da marca: Wizard, CNA, Seguralta, Velox…);
  - 1 com a franqueadora inapta (Via Lorran).
- **Nº de lojas:** o da aba de pesquisa. Sem número, as unidades com o nome na Receita.

**Como foi feito:**
- `dados/outros_lista.py`: nomes, subsegmentos e termos de busca.
- `busca_receita.py q17` e `q18`: busca na Receita.
- `build_outros.py`: escolha automática do CNPJ, com estas travas:
  - CNAE compatível com o subsegmento;
  - nunca uma associação de franqueados;
  - franqueadora antes de unidade.
- `dados/mk_dec_outros.py`: decisões à mão. Exemplos:
  - Espaçolaser é a Corpóreos S.A.;
  - Hinode é o Grupo Hinode Participações;
  - CCAA é o CCAA Centro de Cultura Anglo Americana;
  - homônimos de outro ramo foram descartados (Hinode Telecom, Bamboo Pneus, Granado Imóveis).
- `grava_outros.py`: cria a aba pelo XML. Mudaram só `workbook.xml`, `rels`, `Content_Types` e `app.xml`, e entrou a
  `sheet43`.

**Conferência:**
- 168 CNPJs (CNPJ e referência), todos com dígito válido;
- a pasta abriu no Excel com 43 abas e 21 dinâmicas.

## 6ª parte: associações nacionais que faltaram e as grandes regionais (só pesquisa)
**Pedido:** dizer se faltou alguma associação nacional e listar as grandes regionais de cada segmento.

**Já usadas no arquivo:** ABRAS, APAS, ABF, ANAMACO, ABRAMAT, ABCEM, FEBRAMAT, ABVTEX, SBVC, IRTT, Abrafarma,
Febrafar, Abrafad, Fecofar, Abrasel, ANR, Fecombustíveis, ABIEPS, ANP, IPB, ABINPET, ANCLIVEPA, ABRAPARK e Sindepark.

**Nacionais que faltaram (conferidas na internet):**
- ABAAS: atacarejo, com ranking próprio;
- ABERC: refeições coletivas;
- IDV: as cerca de 70 maiores varejistas;
- Alshop e ABLOS: lojistas de shopping;
- ACAD Brasil: academias;
- ANEL: lavanderias;
- AFRAC: automação comercial, o setor de PDV e bobinas;
- Brasilcom: distribuidoras regionais de combustível;
- Abióptica (indústria óptica) e Sindióptica (varejo óptico de SP).

**Conhecidas, mas não conferidas agora:**
- ABAD (atacado distribuidor), Abcfarma (farmácias independentes), ABIP (padarias), ABRASCE (shopping centers) e CNDL
  (CDLs);
- dos segmentos do CRM sem lista: ABCR (pedágio), ABRATI (ônibus), ABOL (logística), ANAHP (hospitais), ABComm
  (e-commerce) e Abraplex (cinemas).

**Regionais:**
- **Supermercados:** 26 associações estaduais filiadas à ABRAS, da página de contatos da ABRAS. A ABRAS fala em 27; a
  de Roraima não aparece na página.
- **Postos:** sindicatos estaduais, como Minaspetro, Sulpetro, Paranapetro, Sincopetro e os Sindicombustíveis. Em
  abril/2026, os sindicatos de 9 bases anunciaram a saída da Fecombustíveis: BA, MG, SP, Santos, PR, SC, RS, MT e DF.
- **Distribuidoras:** 8 sindicatos ligados à Brasilcom.
- **Material de construção:** 38 Acomacs e 3 Fecomacs ligadas à ANAMACO, e o Sincomavi (Grande SP).
- **Farmácias:** Sincofarmas estaduais.
- **Estacionamentos:** Sindepark (SP) e o sindicato de garagens e estacionamentos do RS.
- **Óticas:** Sindióptica (SP).

**Observações:**
- ABLA é a associação das locadoras de automóveis: serve ao segmento Automotivo, Peças e Locadoras, não ao pet.
- ABRASE não existe no setor pet. Talvez seja ABRASEL (bares e restaurantes) ou ABASE (supermercados da Bahia).
- A ABRASCE é nacional.
