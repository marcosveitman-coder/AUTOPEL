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
