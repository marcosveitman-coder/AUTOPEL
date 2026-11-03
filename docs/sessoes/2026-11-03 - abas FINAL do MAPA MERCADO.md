# Sessão 03/11/2026 — abas FINAL do MAPA MERCADO (alimentação, rede de lojas, construção)

Registro sem CNPJs, códigos de cliente, datas de compra nem nomes de vendedores (repositório público). Os nomes de
redes abaixo são listas de mercado, não dados de clientes.

## Pedido
O usuário criou três abas vazias no `MAPA MERCADO.xlsx` e pediu para preenchê-las com nome fantasia, razão social,
grupo administrador, nº de lojas e CNPJ (o mais importante), sem duplicidade:

| Aba | Origem |
|---|---|
| **ALIMENTACAO FINAL** | aba "ALIMENTAÇÃO (...)" + Alessandra Dayan, colunas AC–AF |
| **REDE DE LOJAS FINAL** | "VAREJO (MAIORES)_RADES" + Alessandra Dayan, colunas AJ–AM + "VAREJO VESTUARIO (IA)" |
| **CONSTRUÇÃO FINAL** | Alessandra Dayan, colunas V–Y |

Depois disso, completar as listas com as referências do setor (ANAMACO, ABRAMAT, ABCEM, Exame, ANR, Abrasel e Google),
inclusive por região.

## Entregas (MAPA MERCADO, só as 3 abas)

| Aba | Linhas | CNPJ confirmado | Não confirmado | Rede não encontrada |
|---|---|---|---|---|
| ALIMENTACAO FINAL | 368 | 208 | 112 | 48 |
| REDE DE LOJAS FINAL | 324 | 288 | 36 | — |
| CONSTRUÇÃO FINAL | 251 | 228 | 23 | — |

Colunas (A–O): Nº, NOME FANTASIA, RAZÃO SOCIAL, GRUPO ADMINISTRADOR, Nº DE LOJAS, CNPJ, CONFERÊNCIA DO CNPJ, UF SEDE,
MUNICÍPIO SEDE, SITUAÇÃO NA RECEITA, FONTE DO Nº DE LOJAS, LOJAS ATIVAS COM O NOME NA RECEITA (TODOS OS CNPJs),
SEGMENTO, ORIGEM (aba e linha, ou complemento), OBSERVAÇÃO.

- **Formato:** cabeçalho congelado, filtro, estilos que já existiam (cabeçalho 194, número 127). Nenhum estilo novo.
- **Ordem:** confirmados primeiro, por nº de lojas; depois os não confirmados; por fim os não encontrados.
- **Gravação:** só o XML das 3 abas (`grava_abas.py`). Conferido que, das 136 partes do arquivo, só elas mudaram. A cópia
  abriu no Excel sem reparo, antes de substituir o original.
- **Backup:** está no scratchpad da sessão, que é apagado. O OneDrive tem o histórico de versões.
- **Conferência final no arquivo gravado** (`confere_final.py`): os 724 CNPJs têm dígito verificador válido, existem
  na Receita e estão ATIVOS. Nenhum nome se repete dentro de uma aba.

## Como foi feito
1. **Origem:** 704 linhas (340 de alimentação, 314 de lojas, 50 de construção), juntadas em 479 redes únicas.
   - **Duplicatas:** chave do nome mais junções manuais, por exemplo "MC DONALDS" × "McDonald's", "BACCIO DI LATTE" ×
     "Bacio di Latte", "Vó Alzira (Sudeste)" × "Fábrica de Bolo Vó Alzira" e as dezenas de linhas repetidas de
     Lojas Avenida, Riachuelo e Renner na aba de vestuário.
   - **Linhas de grupo da Alessandra** ("GRUPO TRIGO (Spoleto / China in Box / Casa do Pão de Queijo)", "GRUPO CRM",
     "GRUPO AF", "AZZAS 2154", "Halipar"): viraram nota nas linhas das marcas, para não contar as lojas duas vezes.
     A Azzas foi desdobrada nas marcas da aba IRTT.
2. **CNPJ da origem:** dos 276 informados, 104 falhavam no dígito verificador. Muitos da aba "ALIMENTAÇÃO (...)", da
   linha 97 em diante, eram claramente inventados, com sequências como 22.333.555 e 14.555.777.
3. **Receita local (2026-10), três varreduras:**
   - **nome e CNPJ:** razão social e nome fantasia das redes, CNPJs da origem;
   - **redes por CNAE e região:** raízes com muitas lojas ativas de alimentação, de varejo ou de construção, e nomes
     fantasia repetidos em muitos CNPJs (franquias);
   - **matrizes e complementos:** matrizes das raízes escolhidas, nomes refinados e os complementos.
4. **Escolha do CNPJ, rede a rede e à mão:** as regras automáticas erraram (holding estrangeira para o McDonald's, uma
   empresa qualquer com "Kopenhagen" no nome). A regra está no `LEIA-ME.md` da pasta: operadora com lojas próprias,
   franqueadora ou dona da marca, ou CNPJ da origem conferido. **Nunca franqueado.** Sem certeza, fica em branco.
5. **Nº de lojas:** vale o número mais recente e com data das abas de origem. Quando a origem não tem número:
   - para as redes da origem, as lojas ativas com o nome na Receita;
   - para os complementos, o número da IRTT ou as lojas ativas da empresa na Receita.

   A coluna L traz sempre a contagem de lojas ativas com o nome (franqueados incluídos), para conferência.

## Complementos (fonte na coluna ORIGEM)
- **Ranking ANAMACO, Top 50 Varejo Nacional 2025 e 2026** (Revista Anamaco): 35 lojas de material de construção que não
  estavam na lista, com a posição no ranking.
- **Aba "I.R.T.T.__Por bandeira (FINAL)"** (ranking IRTT Top 300): 29 bandeiras de lojas físicas, como as marcas da
  Azzas, Carter's, Pormenos, iPlace e Brunella.
- **Receita, redes regionais por CNAE:** 385 redes com sede em todas as regiões.
  - Exemplos: Lojas Benoit (RS/SC), Salfer (SC/PR), Eletro Mateus (MA/PA), Móveis Estrela (GO), Laser Eletro (PE),
    Bela Tintas e Leo Madeiras (SP), Glaçaí (PB/PE), refeição coletiva (Refeições Naturas, Exal, Apetit).
  - Cortes e exclusões: no `LEIA-ME.md` da pasta.
- **Receita, franquias pelo nome fantasia repetido:** 17 redes, como Cosechas, Frutos de Goiás e Kidstok.
- **Ranking ABF 2025:** a microfranquia Fast4You.
- **ABRAMAT e ABCEM** são associações de indústrias (fabricantes e construção metálica) e não listam lojas. Foram
  consultadas, mas não renderam linhas. O texto da Exame é o ranking IRTT de material de construção de 2024 (13 redes,
  todas já na lista, com o nº de lojas usado quando faltava).

## Correções que a pesquisa trouxe
- **Pizza Hut:** a master franqueada continua sendo a IMC. A BFFC (Bob's) é a maior franqueada. A aba de origem dizia
  "BFFC (antes IMC)".
- **KFC:** passou para a Horizonte Restaurantes (Kentucky Foods Chile + IMC) em jun/2026.
- **Domino's:** é da Vinci Partners (master franqueada). O CNPJ da origem era do Grupo Trigo.
- **Casa do Pão de Queijo:** o Grupo Trigo comprou a operação em 2025, na recuperação judicial.
- **Grupo de Griletto, Montana Grill, Jin Jin e Croasonho:** é a Halipar. A Alessandra escreveu "Grupo SFB".
- **Rede Graal:** o CNPJ que a aba IRTT trazia é, na Receita, de um restaurante no RJ. A linha ficou "não confirmado".
- **Outback:** o grupo é a Vinci Partners (ex-Bloomin' Brands).
- **Kopenhagen e Brasil Cacau:** o Grupo CRM é da Nestlé desde 2023.
- **Bob's na REDE DE LOJAS** (Alessandra AJ41): é rede de alimentação. A linha ficou com um alerta.

## Pendências
- **112 redes de alimentação, 36 de lojas e 23 de construção** ficaram com CNPJ não confirmado. Muitas são franquias
  cuja franqueadora não leva o nome da rede na razão social, como Chiquinho, Casa do Construtor e Domino's. Dá para
  fechar com a COF (circular de franquia) ou com contato direto.
- **48 redes de alimentação "não encontradas"** (nenhum estabelecimento com o nome na Receita). Quase todas vêm da lista
  regional da aba "ALIMENTAÇÃO (...)", que tinha CNPJs inventados. Conferir se existem antes de usar.
- Não foi feito o cruzamento com o faturamento e o cadastro (cliente atendido, última compra, vendedor). Fica para
  quando o usuário pedir.

## Lições
- **Lista vinda de IA:** CNPJ com dígito certo não basta. É preciso existir na Receita, estar ativo e ter o nome da
  rede. Nome de rede inventado também aparece. Sinal: nenhum estabelecimento com o nome em todo o Brasil.
- **Franquia:** a busca pelo nome traz centenas de franqueados. A franqueadora costuma ter CNAE 7740-3/00 e razão com
  "FRANCHISING" ou "FRANQUIAS". Capital social alto ajuda a separar dona da marca de franqueado.
- **"Ativa" na Receita não quer dizer loja aberta:** Ricardo Eletro, Saraiva e Mappin ainda aparecem com centenas de
  estabelecimentos ativos.
- A matriz de uma raiz pode não ser a 0001 (ex.: Quero-Quero, Todimo). Usar o estabelecimento com identificador de
  matriz = 1.
- Na ferramenta Bash daqui, heredoc com muitas aspas às vezes quebra ("unexpected EOF"). Para arquivos grandes, usar
  a ferramenta Write.

## Scripts
`scripts cruzamento/avulsos/mapa_final/`, com a ordem de execução no `LEIA-ME.md`. Os dados e as decisões ficam em
`mapa_final/dados/`, fora do git.

---

# 2ª parte (mesmo dia): fechar os CNPJs e mostrar as maiores por região

## Pedido
- O cruzamento com o faturamento e o cadastro fica **para depois**.
- "Acabar o que não foi acabado": trazer **todos os CNPJs**. Em rede de franquias, pode ser o CNPJ da matriz.
- Procurar no Google listas e associações que ainda não tinham sido usadas.
- Mostrar as **maiores do Brasil** e as **maiores de cada região** (Norte, Nordeste, Centro-Oeste, Sudeste, Sul) em
  alimentação, rede de lojas e construção.

## Entregas (MAPA MERCADO)
- **As 3 abas FINAL foram regravadas** com 20 colunas (A–T). Colunas novas:
  - **CNPJ DE REFERÊNCIA (UNIDADE DA REDE)** e **EMPRESA DE REFERÊNCIA**;
  - **REGIÃO DA SEDE**;
  - **RANKING BRASIL (Nº DE LOJAS)** e **RANKING NA REGIÃO**.
- **Aba nova "MAIORES POR REGIÃO"**, no fim da pasta. Para cada segmento:
  - as 30 maiores do Brasil;
  - as 20 maiores de cada região.

  A região é a UF da sede da empresa, e a ordem é o nº de lojas.

**Conferência do CNPJ**, por aba:

| Aba | Linhas | Confirmado | Provável (conferir) | Só CNPJ de referência | Não confirmado | Rede não encontrada |
|---|---|---|---|---|---|---|
| ALIMENTACAO FINAL | 368 | 234 | 30 | 34 | 22 | 48 |
| REDE DE LOJAS FINAL | 324 | 293 | 13 | 8 | 10 | — |
| CONSTRUÇÃO FINAL | 250 | 232 | 1 | 1 | 16 | — |

Antes desta parte eram 208 / 288 / 228 confirmados. Agora são **803 CNPJs de rede** (confirmados + prováveis) e
**43 de referência**.

- **Gravação:** só pelo XML (`mapa_final/grava_abas.py`). Mudaram só:
  - as 3 abas e a aba nova;
  - `workbook.xml`, `workbook.xml.rels`, `[Content_Types].xml` e `docProps/app.xml`.

  A cópia abriu no Excel (36 abas, com as 21 tabelas dinâmicas inteiras) antes de substituir o original.
- **`confere_final.py`:** todo CNPJ e todo CNPJ de referência têm dígito verificador válido e estão ATIVOS (Receita
  local ou BrasilAPI). Nenhum nome se repete dentro de uma aba.

## Níveis de conferência do CNPJ (coluna G)
- **CNPJ CONFIRMADO NA RECEITA:** é a operadora das lojas próprias, a franqueadora ou dona da marca, ou o CNPJ da
  origem conferido.
- **CNPJ PROVÁVEL (CONFERIR):** franqueadora achada por indício forte, mas sem prova completa. Exemplos de indício:
  - CNAE de franquia ou de administração, com o nome da rede, na cidade de origem da marca;
  - empresa mais antiga com o nome.
- **FRANQUEADORA NÃO ACHADA: VER CNPJ DE REFERÊNCIA:**
  - a franqueadora não leva o nome da rede na Receita;
  - a coluna H traz o CNPJ de **uma unidade** da rede, para achar a rede no cadastro e no faturamento, e **não** como
    CNPJ da rede;
  - a unidade é escolhida nesta ordem: na UF de origem, matriz, maior capital, mais antiga, CNAE do segmento e nome
    completo.
- **CNPJ NÃO CONFIRMADO:** nenhum candidato seguro. Exemplos: Rede Graal, Puket, Crosby, Milon, Rabelo e algumas lojas
  de construção pequenas.
- **REDE NÃO ENCONTRADA NA RECEITA:**
  - as 48 linhas regionais da aba "ALIMENTAÇÃO (...)", da linha 97 em diante;
  - nenhum estabelecimento com o nome em todo o Brasil;
  - **recomendação: tirar da lista ou confirmar com a fonte.**

## Correções feitas nesta parte
- **Contagem de lojas pelo nome:** o nome "LOJAS X" também era contado como só "X". Em "Lojas Rezende", isso pegava
  2.125 empresas com o sobrenome Rezende (drogarias, transportadoras, construtoras).
  - Agora a contagem usa o nome completo e só CNAE do segmento.
  - Lojas Rezende = 20 lojas, todas em TO/PA.
  - Mudaram mais 5 números de lojas: KFC 110 → 72, Popeyes 19 → 17, Milon 116 → 109, Cia. Paulista de Pizza 9 → 8 e
    Grupo Tropy 2 → sem número.
  - A coluna "lojas ativas com o nome" fica em branco quando passa de 3× o nº de lojas, porque aí são homônimos.
- **Constrular saiu da construção:**
  - o nome aparece em 1.112 estabelecimentos de 1.036 empresas, espalhadas por todas as UFs;
  - não se achou rede nem associação com esse nome, só uma holding em GO;
  - é nome genérico ("constru" + "lar").
- **UF da sede das redes sem CNPJ:** a UF da unidade de referência punha a rede na região errada (Mr. Fit no Nordeste).
  Agora a UF vem, nesta ordem, de:
  1. Receita;
  2. pesquisa (Mr. Fit = Paulínia/SP; Frutos de Goiás = Senador Canedo/GO);
  3. aba de origem;
  4. a UF com mais lojas da rede, marcada "sede não confirmada" na observação.

## Listas e associações procuradas no Google
- **Usadas:**
  - matérias regionais sobre o ranking IRTT (Nordeste e Ceará). As redes de lojas e de construção do Nordeste que
    aparecem nelas já estavam na lista;
  - ranking de franquias por unidades do Portal do Franchising.
- **Sem lista pública de associados:** ABVTEX, FEBRAMAT, ANR e Abrasel. O ranking da SBVC mais recente é de 2024.
- **Faltam 3 redes de varejo da lista de franquias:** Piticas, Multicoisas e Livraria Nobel. Não foram incluídas
  ainda.

## Pendências
- Piticas, Multicoisas e Livraria Nobel: procurar o CNPJ na Receita e incluir na REDE DE LOJAS FINAL.
- Fechar os 30 + 13 + 1 "prováveis" e os 43 de referência com a COF (circular de franquia) ou com contato direto.
- As 48 redes "não encontradas": decidir com o usuário se saem da aba.
- Cruzamento com o faturamento e o cadastro: quando o usuário pedir. A coluna H (referência) ajuda a achar a rede.

## Lições
- **Contagem pelo nome:** tirar "LOJAS/REDE/GRUPO" do começo só vale se sobrarem 2+ palavras. Filtrar pelo CNAE do
  segmento. Se a contagem passar muito do nº de lojas, são homônimos.
- **Nome genérico de loja** (Constrular, "Elétrica", "São Jorge", "Avenida"): centenas de empresas sem ligação têm
  o mesmo nome. Com muitas raízes espalhadas por todas as UFs e sem rede na web, não é rede.
- **Região:** usar a UF da sede, nunca a da unidade usada como referência.
- **Aba nova pelo XML numa pasta com dinâmicas:**
  - pôr a aba no **fim** de `<sheets>`, porque os nomes definidos usam `localSheetId` (a posição da aba);
  - criar a parte da aba, a relação em `workbook.xml.rels` e o `Override` em `[Content_Types].xml`;
  - somar 1 nos dois contadores do `docProps/app.xml`;
  - o gravador troca só o conteúdo quando a aba já existe.

## 3ª parte: as 3 redes da lista de franquias
O usuário mandou incluir as 3 redes de varejo que faltavam. Elas vieram de uma busca nova na Receita (`q5.json`) e da
pesquisa na web:

| Rede | Nº de lojas (fonte) | CNPJ | Região |
|---|---|---|---|
| Multicoisas | 212 (ranking ABF 2019) | **confirmado**: a franqueadora "Multicoisas Franquias", ativa desde 1985 | Centro-Oeste (matriz em MS; filial em SP) |
| Piticas | 220 (Portal do Franchising, set/2025) | **referência**: a franqueadora original está baixada e a holding com o nome não tem os fundadores como sócios | Sudeste (sede em Guarulhos/SP) |
| Livraria Nobel | 121 (ranking ABF 2019) | **referência**: a Livraria Nobel S.A. está baixada e a holding dona da marca não aparece na Receita | Sudeste |

- **REDE DE LOJAS FINAL:** 327 linhas (294 confirmados, 13 prováveis, 10 de referência, 10 não confirmados).
- **Unidade de referência:** agora é escolhida na UF da sede quando a pesquisa a conhece. Mr. Fit passou a ter
  referência em SP, e Frutos de Goiás em GO.
- **Gravação:** mudaram só ALIMENTACAO FINAL, REDE DE LOJAS FINAL e MAIORES POR REGIÃO. A CONSTRUÇÃO ficou igual.
  `confere_final.py` não achou erros, e a pasta abriu no Excel (36 abas, 21 dinâmicas) antes de substituir o original.
- **Pendência:** os números da Multicoisas e da Nobel são de 2019, os mais recentes com data que foram achados. A
  Receita mostra 33 lojas ativas com o nome Nobel. Confirmar com a ABF ou com as redes.

## 4ª parte: CNPJs que ainda faltavam e aba R_E_S_U_M_O
**Pedido.** O usuário apontou as redes sem CNPJ e pediu esforço extra:
- ALIMENTACAO FINAL da linha 300 (Rede Graal) em diante;
- REDE DE LOJAS FINAL da 309 (Constance) em diante;
- CONSTRUÇÃO FINAL da 235 (Beira Rio) em diante.

Pediu também o racional do estudo na aba **R_E_S_U_M_O**, que ele mesmo criou: metodologia e lugares pesquisados.

**Antes de gravar.** O usuário tinha copiado à mão, para a coluna F, o CNPJ de referência das linhas 266 a 299.
- Decisão: seguir o mesmo critério para todas as redes nesse nível. A F traz o CNPJ de uma unidade.
- A coluna G passou a dizer "CNPJ DE UMA UNIDADE DA REDE (FRANQUEADORA NÃO ACHADA)".
- Conferido por nome de rede: nenhum CNPJ que já estava na aba mudou (só a linha, porque a ordem foi refeita).

**Busca.** Duas varreduras novas na Receita:
- `q6`: 107 redes, pelo núcleo do nome e pela razão informada no ranking;
- `q7`: as donas da Bráz e da Rede Graal.

As decisões, rede a rede, ficaram em `dados/dec_6.json`.

**Resultado**

| Aba | Confirmado | Provável | CNPJ de uma unidade | Não confirmado | Não encontrada |
|---|---|---|---|---|---|
| ALIMENTACAO FINAL (368) | 237 | 33 | 42 | 13 | 43 |
| REDE DE LOJAS FINAL (327) | 296 | 16 | 9 | 6 | — |
| CONSTRUÇÃO FINAL (250) | 237 | 4 | 0 | 9 | — |

- **Confirmados pela razão ou pelo CNAE de franquia:**
  - franqueadoras: Parmê, Jorge Bischoff e Peça Rara;
  - operadoras: Tortelê e Lucca Cafés;
  - razão igual à do ranking ANAMACO: Cocil, Copafer, Ferragem Lorenzet, Nicom e Padovani. Essas não tinham entrado na busca da 1ª rodada.
- **Rede Graal:** cada posto/restaurante é uma empresa. Não existe franqueadora nem holding com o nome. Fica com o CNPJ da unidade de origem da rede (o posto "Graal Petropen", de 1974).
- **Continuam sem CNPJ:** nomes comuns (Matsuri, Bulls, Rabelo, Seller, São Jorge, Caetano) e redes sem nenhuma loja na UF da origem. A observação de cada linha diz o que foi achado.
- **Aba R_E_S_U_M_O:** gravada com o `resumo.py`, chamado pelo `grava_abas.py`. Usa só estilos que já existiam na pasta. Seções:
  1. objetivo;
  2. metodologia em 8 passos;
  3. níveis da conferência do CNPJ, com as contagens por aba;
  4. fontes pesquisadas, com os endereços;
  5. maiores de cada segmento e região;
  6. cuidados;
  7. próximos passos.
- **Gravação:** mudaram só as 3 abas FINAL, a MAIORES POR REGIÃO e a R_E_S_U_M_O. `confere_final.py` não achou erros. A pasta abriu no Excel (37 abas, 21 dinâmicas).

**Lições**
- **O Excel do usuário salva com strings compartilhadas:** depois disso o XML da aba já não tem o texto do cabeçalho. Conferir a aba pela célula B1, não pelo texto.
- **O usuário pode salvar a pasta entre um pedido e outro:** comparar o arquivo atual com o backup antes de trocar, e refazer a cópia sobre a versão nova.
- **Na exportação de imagem pelo COM com o Excel oculto,** o gráfico colado pode sair em branco. Para conferir o conteúdo, ler o texto da aba.
