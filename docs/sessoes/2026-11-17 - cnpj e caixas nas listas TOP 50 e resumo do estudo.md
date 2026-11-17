# 2026-11-17 – CNPJ e caixas nas listas TOP 50 e resumo do estudo para a diretoria

Arquivo: `COMERCIAL (NOVO)\MAPA MERCADO V2.xlsm`, trocado às 09:38. O backup de antes está no scratchpad da sessão,
fora do git.

## Pedidos

1. **Tarefa 1**, nas 12 abas "Alessandra Dayan VF" e "<SEGMENTO> (TOP 50)":
   - inserir linhas no topo, como na FARMACIA VF ("56 linhas", lido como "o cabeçalho vai para a linha 6");
   - criar a tabela a/b/c/d de caixas por porte;
   - criar, no fim da tabela, a coluna CNPJ (matriz de cada marca);
   - criar a coluna "Total de caixas/ mês", com a fórmula da coluna V da FARMACIA VF.
2. **Tarefa 2**: na "ORGANOGRAMA impressao", abaixo da linha 182, um resumo do estudo para nível C-level, com o racional,
   como foi feito, os sites e associações usados (nº de lojas, UF, CNPJ) e as premissas de caixas por porte e segmento.

## Entregas

### Tarefa 1

**Linhas e cabeçalho**
- TOP 50: 3 linhas novas, e o título "Detalhes do..." foi para a linha 4.
- Alessandra Dayan VF: 5 linhas novas.
- Nas 12 abas, o cabeçalho ficou na linha 6 e os dados começam na 7.

**Tabela de caixas por porte**
- Nas linhas 1–4: letra na I e valor na J.
- Valores: 17/12/7/4. Na SUPERMERCADO (TOP 50): 20/15/10/8, a premissa de supermercado.
- Na Alessandra Dayan VF, que mistura segmentos: J = demais segmentos e K = supermercado. A fórmula escolhe a coluna pelo
  SEGMENTO.

**Colunas novas no fim da tabela**
- I "CNPJ" e J "Total de caixas/ mês". Nas TOP 50, a tabela do Excel foi aumentada (`ListObject.Resize`).
- A fórmula é a mesma das abas FINAL: `SE(ESQUERDA(F;1)="A";$J$1;…;0)*N(C)`. Porte sem letra ou SEM Nº DE LOJAS dá 0, sem
  erro.

**CNPJ**
- 419 das 550 marcas da Alessandra Dayan VF e 417 linhas nas TOP 50.
- 42 são "prováveis", com o comentário "CNPJ provável (conferir)" na célula.
- 131 ficaram em branco:
  - franquias sem franqueadora identificada (só se achou o CNPJ de uma loja);
  - CNPJs "não confirmados" nas abas FINAL;
  - vínculos errados descartados: farmácia avulsa com o mesmo nome da rede, ou supermercado de outra UF com o mesmo nome.
    Exemplos: Super Nosso (MG) ligado a uma empresa do RN; Grupo Amigão ligado a uma lotérica do MS.
- Quando o CNPJ da aba FINAL era de filial, entrou a matriz da raiz (7 casos). No Grupo Mateus a matriz não é a 0001.
- Achados novos na busca por nome:
  - Grupo Boticário: O Boticário pela dona da marca; Quem Disse Berenice, Beauty Box e Eudora pela operadora das lojas
    próprias;
  - Natura;
  - We Pink (participações);
  - Acquazero Franchising;
  - Maria Brasileira Franchising;
  - SGH Brasil (Sunglass Hut / GrandVision), marcada como provável.

**Abas que dependem destas**
- A dinâmica da Alessandra Dayan VF andou sozinha, de A1:H551 para A6:H556.
- A CONFIG DETALHE foi refeita (org_dados → org_impressao → org_detalhe → org_detalhe.ps1). Só mudaram a data e a linha
  do cabeçalho das 11 TOP 50 (3 → 6).

### Tarefa 2

Duas folhas novas na "ORGANOGRAMA impressao", linhas 183–234. Agora são 11 folhas A4 paisagem.

- **Folha 10, "Como os números foram feitos"**:
  - o que o estudo responde;
  - método em 4 passos (mapear, validar na Receita, cruzar com a Autopel, medir o potencial);
  - premissas de caixas por porte (A 20/17, B 15/12, C 10/7, D 8/4; supermercado / demais);
  - critério do porte (faturamento publicado ou lojas × média por loja);
  - regra do STATUS AUTOPEL.
- **Folha 11, "Fontes e cuidados"**:
  - fontes por tipo de dado: Receita (dados abertos), BrasilAPI, CNPJ.ws, ABRAS, ranking de farmácias, IRTT, ABF, ANAMACO,
    ANP, Estapar, PetCenter, Panorama PetVet e base Autopel;
  - associações consultadas que não publicam lista de redes;
  - 4 cuidados de leitura.

## Conferências

| Conferência | Resultado |
|---|---|
| `confere_top50.py` | Células antigas iguais e descidas 3 ou 5 linhas. Caixas recalculadas iguais. CNPJ igual ao tsv. Nenhum erro nas 12 abas. As outras 43 abas iguais célula a célula. |
| ORGANOGRAMA impressao | Linhas 1–182 iguais. 11 páginas, `VPageBreaks` = 0. Imagens das folhas 10 e 11 vistas. |
| Macro do duplo clique | 390 de 390 números iguais ao esperado (`org_detalhe_teste.ps1` + `org_detalhe_esperado.py`). |
| `confere_final.py` | Tudo certo, menos as 380 "#N/A" já existentes na "I.R.T.T.__ (MAIORES 300 GRUPOS)", fora do pedido. Os exemplos continuam ATIVO. |
| Amostra de 10 CNPJs gravados | 10 de 10 são matriz ativa com o nome da marca. Para 2, a BrasilAPI disse "não encontrado"; o CNPJ.ws confirmou os dois. |

## Decisões tomadas (avisar o usuário)

- "56 linhas" foi lido como "cabeçalho na linha 6, como na FARMACIA VF".
- Supermercado usa a premissa 20/15/10/8, e os outros segmentos 17/12/7/4, como nas outras abas. Os valores ficam nas
  linhas 1–4 e podem ser trocados.
- CNPJ provável entra com comentário. Unidade de franquia não entra.

## Pendências

- 131 marcas sem CNPJ, a maioria franquias de serviços, beleza e óticas sem franqueadora com o nome da marca na Receita.
  Dá para pesquisar uma a uma (COF, site da marca), se o usuário quiser.
- As 380 "#N/A" da aba original do IRTT continuam lá (oferta de 16/11 ainda em aberto).

## Lições

- A REDE DE LOJAS FINAL ganhou linhas depois de 05/11: a ligação curada ("linha N") tem de ser procurada pelo nome na aba.
- A linha de apoio da FARMACIA VF com TIPO "Farmácia" (posição alta) é uma loja avulsa com o mesmo nome, e não a rede.
- A ligação da SUPERMERCADOS (reuniao) a linhas de leads (posição 800+) pode ser homônimo de outra UF: conferir a UF.
- A BrasilAPI pode não ter um CNPJ ativo. Conferir no CNPJ.ws antes de concluir que não existe.

## 2ª parte: pesquisa das marcas sem CNPJ (pedido "Sim, pesquisar")

Arquivo trocado de novo às 10:06. Foi refeito a partir do original desta manhã: as 12 abas, a CONFIG DETALHE e a impressão
(estas duas ficaram iguais às da 1ª gravação).

**Como foi feito**
1. Observatório público de franquias da Marktech (`cnpj_publicado_site.py`): mostra o CNPJ que cada rede publica no
   próprio site. Saíram 40 candidatos.
2. Internet: site oficial (rodapé, "quem somos", política de privacidade), Reclame Aqui e cadastro de franquias da FGV.
   Usado para supermercados, drogarias e franquias grandes.
3. Todo candidato conferido na cópia local da Receita: razão social, nome fantasia, matriz, situação e CNAE. Quase todas
   as franqueadoras têm CNAE 7740-3/00.

**Regras de decisão**
- Aceito: franqueadora, dona da marca ou operadora das lojas com o nome da marca.
- Recusado:
  - CNPJ publicado que é de uma unidade (Azul Empréstimo Maracanaú) ou da loja virtual (Chilli Beans; ficou a Luz
    Franquias, da política de privacidade);
  - empresa baixada (Chiquinho).
- "Provável" com comentário na célula: Grupo É, Wizard (Pearson), Nagumo, Amigão, Carvalho, ADCOS e Capilé (esta com o
  CNPJ inapto na Receita).

**Resultado**
- 469 das 550 marcas com CNPJ, eram 419 de manhã. Nas TOP 50 são 458 linhas.
- Comentário "provável" em 49 marcas da Alessandra Dayan VF.
- As outras 81 continuam em branco, sem franqueadora identificável:
  - franquias pequenas de óticas, lavanderias, conveniência e alimentação;
  - estacionamentos, postos e lojas de construção regionais, com nome comum e sem CNPJ seguro;
  - Seguralta, Domino's, Divino Fogão, Unifarma e Drogaria Total, entre outras.

**Conferências**
- `confere_top50.py`: tudo certo.
- Impressão e CONFIG DETALHE iguais às da manhã.
- Macro: 390 de 390.

**Lições**
- O "CNPJ publicado no site da rede" nem sempre é da franqueadora. Às vezes é de uma unidade ou da loja virtual.
- A BrasilAPI pode dizer "não encontrado" para um CNPJ que existe (Selfit). A cópia local da Receita resolve.

## 3ª parte: caixas por segmento na Alessandra Dayan VF

- O usuário montou na aba um quadro de caixas por porte para cada segmento:
  - títulos dos segmentos em F1:P1, letras a–d em E2:E5 e valores em F2:P5;
  - o cabeçalho dos dados foi para a linha 8 (dados de 9 a 558).
- Gravado pelo Excel numa cópia, com a fórmula na coluna J (linhas 9–558):
  `=SEERRO(ÍNDICE($F$2:$P$5;CORRESP(ESQUERDA($F9;1);$E$2:$E$5;0);CORRESP($H9;$F$1:$P$1;0));0)*N($C9)`.
  Ela busca a coluna do SEGMENTO e a linha da letra do PORTE. Faz o mesmo que vários SE, numa fórmula só.
- As duas dinâmicas da aba (L8 e L28) foram atualizadas.
- Conferência: as 550 linhas recalculadas em Python estão iguais. As outras colunas e abas não mudaram. Total de
  2.168.035 caixas/mês.
- Arquivo trocado às 10:37. O backup está no scratchpad.
- Os 11 títulos do quadro precisam ser iguais ao texto da coluna SEGMENTO: se mudar um título, a linha dá 0.

## 4ª parte: rede própria x associação/franquia (Alessandra Dayan VF, colunas A e B)

- O usuário inseriu duas colunas no começo da aba:
  - A, "rede / associação ou franquia";
  - B, "quem atende na autopel", com 9 linhas já preenchidas (nomes da equipe).
- **A** foi preenchida nas 550 linhas: REDE PRÓPRIA (266) ou ASSOCIAÇÃO / FRANQUIA (284). Os sinais usados, em ordem:
  1. decisões à mão, com pesquisa;
  2. TIPO da FARMACIA VF;
  3. lojas ativas na raiz do CNPJ da marca (50% ou mais das lojas = rede própria);
  4. sinal de franquia: a marca está no observatório de franquias, a aba FINAL fala em franqueadora/ABF, ou o CNPJ tem
     CNAE 7740-3/00.
- Critérios:
  - posto com bandeira de distribuidora (Petrobras, Shell, Ipiranga, Ale) = associação/franquia, porque é revendedor
    independente;
  - rede mista vai pelo modelo predominante: McDonald's, Burger King e Smart Fit = rede; Azzas, Habib's e Subway =
    franquia.
- **B**: o que já estava ficou como estava. Linha de associação/franquia com B vazio recebeu "IA + 1 ou 2 closer"
  (274 linhas).
- Pesquisa: Drogaria Total (associativista, Gironews), Boa Farma e Cityfarma (redes da Febrafar) e Legítima (junção de
  donos de farmácias).
- Conferência:
  - nenhuma célula de C em diante mudou;
  - nenhuma B já preenchida mudou;
  - nenhuma linha de rede própria recebeu "IA";
  - as outras abas ficaram iguais.
- Arquivo trocado às 12:13, com backup no scratchpad.

## 5ª parte: lista de contatos das redes próprias (interrompida pelo usuário)

**Pedido:** para as redes próprias da coluna A, nome, LinkedIn, e-mail e telefone de cinco cargos: gerente, coordenador e
diretor de compras, diretor de supply e CEO, usando Apollo, Lusha e LinkedIn.

**Limites encontrados (17/11):**
- **Apollo:** 170 créditos de e-mail e 0 de telefone. O plano bloqueia a busca de pessoas, tanto a da API quanto a do
  agente de IA. Só enriquece quem já se tem por nome ou LinkedIn.
- **Lusha:** plano gratuito, 4 créditos. Telefone custa 5.
- **ZoomInfo:**
  - a busca é gratuita mas devolve só IDs;
  - o resgate custa 1 crédito por contato, e o plano não libera telefone nem LinkedIn;
  - o limite acabou depois de 10 resgates.
- **LinkedIn Hiring Pro:** só tem prévia de vaga de emprego.

**O que foi feito:**
- Escolhidas as 30 maiores redes próprias por caixas/mês.
- A contagem gratuita no ZoomInfo achou 133 contatos em 23 redes. Sete redes não estão no ZoomInfo.
- Resgatados 10 contatos (10 créditos) de uma rede; 3 são dos cargos pedidos, só com e-mail. O arquivo com eles é local,
  fora do git.

**Decisão do usuário:** parar aqui. Não houve pesquisa pública e nenhuma planilha foi gravada.

**Para retomar:**
- pesquisa pública de nome e LinkedIn por cargo;
- depois, e-mails pelo Apollo por nome ou LinkedIn, até 170 créditos, com aprovação antes de gastar;
- telefone exige comprar créditos ou fazer upgrade.

**Lições:**
- Checar créditos e o que o plano libera (campos e buscas) antes de prometer o resultado.
- No ZoomInfo, "presidente" no filtro de cargo traz todos os vice-presidentes. Usar "CEO", "Chief Executive" e
  "Diretor Presidente".
