# Sessão 26/10/2026 — universo de compras, metodologia CLAUDE e aba de análise

Registro sem nomes de clientes, compradores ou fornecedores (repositório público). Valores em R$ e contagens são da
foto de 26/10/2026.

## Pedido
1. Entender o contexto de compras pela pasta `COMPRAS\IA (compras)`:
   - prints e exportações do Power BI "Necessidade de Compras";
   - TOTVS: EstoqueNecessidade, HistoricodeCompras e OrdemCompra;
   - fotos diárias do Relatório de Revenda, relatórios do Protheus e boletins de ruptura;
   - `rupturas (evolução) 2.xlsx`;
   - o MRP em HTML (`Downloads\Autopel_MRP_23102026_atualizado.html`).
2. Na aba Geral, a partir da CM: metodologia própria de compra, em QTD e R$, para o mês vigente, o próximo e o
   subsequente ("CLAUDE …").
3. Contexto: o PLANO de AÇAO impresso para os compradores fica na aba PLANO de AÇAO.
4. Aba "analise compras CLAUDE" (criada vazia pelo usuário): conclusões, tabelas e sugestões.

## O que as fontes mostraram
- **TOTVS:**
  - MedContr = média dos 3 últimos meses completos (MES_01 = mês anterior), igual em 100% dos SKUs;
  - PontoPedido = 2 × MedContr em 95% dos SKUs. 266 estão zerados à mão com NOVO_PONTO > 0.
- **Fotos diárias (56 dias, 27/07 a 26/10):**
  - ruptura de 45 a 151 SKUs por dia (média 93);
  - 5.180 pedidos de compra (código + PC). Prazo real do 1º dia visto até a chegada (ULT_COMPRA): mediana de 7 dias,
    p75 12, p90 21;
  - 36% chegam até a 1ª data prometida; 49% têm a data remarcada.
- **Causas dos 1.314 episódios de ruptura:**
  - PontoPedido 0 (sob encomenda): 30%;
  - pedido de compra a caminho, compra tardia: 28%;
  - pedido de venda grande (mais de 1 mês de venda de uma vez): 19%;
  - Comprar > 0 havia 3 dias ou mais, sem pedido: 10%;
  - pedido de compra atrasado: 4%.
- **Curva ABC (MedContr × preço):**
  - A: 132 SKUs, 8,8% dos dias em ruptura, 1,1 mês de estoque;
  - C: 3.743 SKUs, 9,8 meses de estoque.
  - R$ 11 mi (37% do estoque) estão com mais de 4 meses de cobertura ou sem venda.
- **Boletins de ruptura:** 39% resolvem até a data combinada; atraso mediano de 3 dias.
- **Previsão:** a MedContr erra 29% a 54% do volume no mês (WAPE, julho a setembro). Último mês, ponderada e mediana
  não foram melhores.
- **Bases de faturamento (`fat AAAA.N`) cobrem só parte da venda do TOTVS:**
  - 565 SKUs cobertos (80% a 125%) e 919 quase sem faturamento (abaixo de 20%), sobretudo informática;
  - em valor, 68% da venda de maio a setembro.
  - O MRP em HTML usa essas bases como demanda e herda o ponto cego.
- **`fat 2025.2.xlsx`: a aba Planilha1 é cópia exata de dez/2025 da Listagem** (11.211 notas). Somar as duas dobra
  dezembro. `extrai_demanda.py` ignora a cópia.

## Metodologia CLAUDE (planilha)
1. **Previsão do mês:** MedContr × S(mês) ÷ média de S nos 3 meses da MedContr.
   - S sai da mediana de 2023 a 2025 dos SKUs que o faturamento cobre: por grupo, misturado meio a meio com o total;
     com o próprio SKU quando ele tem 2 anos ou mais.
   - Entra amortecido (metade do efeito), limitado a 0,6–1,6.
   - Em 26/10: outubro +18%, novembro +25%, dezembro +9%.
2. **Proteção** = LT do fornecedor + revisão.
   - LT: mediana do fornecedor com 8 ou mais pedidos; senão o prazo prometido nas OCs; senão 8 dias. Máximo 45.
   - Desvio do LT: 3 a 15 dias; padrão 5,9.
   - Revisão: A 15 dias, B 30, C 30.
3. **Segurança** = z × √(T/30 × (CV × prev)² + (prev/30 × desvio do LT)²).
   - z = NORM.S.INV do nível de serviço: A 95%, B 90%, C 80%.
   - CV da série TOTVS de abril a setembro, sem sazonalidade, puxada para a mediana da curva; máximo 1,0. Venda rara
     = CV 0.
4. **Alvo** = prev × T/30 + segurança, no máximo 4 meses de previsão.
5. **Compras:**
   - mês vigente = alvo − saldo (E − F);
   - próximo mês: saldo projetado = saldo + compra − prev × dias que faltam ÷ dias do mês;
   - subsequente: idem, com 1 mês inteiro.
   - Arredonda no múltiplo de compra: MDC das entradas de 12 meses, até 1 mês de venda.
   - PontoPedido 0 = só o que já foi vendido. Comprador INATIVO = 0.
- **Calibração:**
  - com revisão de 30 dias e serviço 98/95/90, o alvo ia a R$ 58 mi;
  - os problemas eram desvio de LT global (11,7 dias, que mistura fornecedores) e prazos de pedidos programados de
    85–90 dias. Corrigidos: mediana, teto e desvio por fornecedor;
  - revisão quinzenal na curva A.
- **Cenários (alvo; compra out / nov / dez):**

  | Cenário | Alvo | Out | Nov | Dez |
  |---|---|---|---|---|
  | Padrão | R$ 40,3 mi | 7,9 | 2,9 | 10,4 |
  | Revisão semanal na curva A | R$ 32,9 mi | 4,8 | 2,1 | 9,6 |
  | Regra atual | R$ 29,0 mi | 2,5 | 7,4 | 10,8 |

  A compra de outubro mais alta é a recomposição da curva A.

## Entregas (cópia conferida; a gravação no arquivo ficou pendente)
- **Aba "param CLAUDE":**
  - parâmetros editáveis (revisão, nível de serviço → z, cobertura máxima, LT/desvio/CV padrão), com nomes
    `CLAUDE_R_A`… `CLAUDE_CV_PADRAO`;
  - tabela `ParamClaude`: 4.224 SKUs com classe, CV, LT, desvio, múltiplo, S01..S12, fontes e fornecedor.
- **Table1 (Geral):**
  - CM–CR: as 6 colunas pedidas;
  - CS–DE: 13 auxiliares (previsões, alvos, segurança, cobertura, classe, LT, desvio, CV, múltiplo);
  - todas são fórmula (LET/XLOOKUP) e o filtro da tabela é devolvido.
  - A tabela é estendida com `ListObject.Resize`: `ListColumns.Add` falha, porque há formatação até a XEP e um
    valor solto na XEQ1.
- **Aba "analise compras CLAUDE":**
  - 1 contexto; 2 diagnóstico com 8 tabelas e gráfico de ruptura por dia; 3 metodologia;
  - 4 regra atual × CLAUDE por comprador (fórmulas) e cenários; 5 qualidade dos dados; 6 dez sugestões.
- **Scripts em `scripts cruzamento/avulsos/compras/claude/`:**
  - `carrega_fotos.py`, `analisa_fotos.py`, `carrega_protheus.py`, `analisa_compras.py`, `extrai_demanda.py`;
  - `gera_parametros.py`, `calcula_claude.py` (referência em Python), `exporta_claude.py`, `grava_claude.ps1`.

## Conferência
- As 19 colunas calculadas pelo Excel × `calcula_claude.py`: **0 diferença** nos 4.224 SKUs (com o z exato do
  NORM.S.INV).
- Valores de todas as outras abas iguais ao arquivo das 10h05; Geral A–CL igual.
- As 65 dinâmicas ficaram com formatação idêntica.
- Alturas automáticas recalculadas foram devolvidas com `devolve_layout.py`. As células vazias sem estilo da
  "Estudo Ocioso" o Excel descarta ao salvar; não muda nada.
- Imagens da análise, da param e das colunas novas conferidas.

## Pendente
- **Gravar no arquivo:** às ~11h45 ele estava em uso por outro processo (o Excel tinha aberto "Relatorio Revenda
  atualizadoXXX.xlsx"). A versão conferida está no scratchpad desta sessão (`saida\final.xlsx`). Gravar só depois de
  confirmar que o arquivo não mudou; se mudou, rodar de novo os passos sobre o arquivo atual.
- O usuário decide o cenário (padrão × revisão semanal da A) e se o PLANO de AÇAO passa a mostrar a coluna CLAUDE.
- Pedir ao TI a venda mensal por SKU de 24 a 36 meses em todos os canais. Hoje a sazonalidade é estimada pelos SKUs
  que o faturamento cobre.

## Lições
- **PowerShell não diferencia maiúsculas:** `$J` (JSON) e `$j` (laço), `$V` (nomes) e `$v`, `$T` e `$t` são a mesma
  variável. Usar nomes longos.
- **Matriz grande no COM:** `Range.Value2 = object[,]` deu "não converte para String" dentro do script, mas
  `Range.Value` funcionou.
- **`ListColumns.Add` falha** quando há formatação até o fim da planilha. Escrever os cabeçalhos ao lado e usar
  `ListObject.Resize`.
- **Nomes no LET:** não podem parecer célula (`s1`, `r`, `c`). Usar `saldo1`, `revis` etc.
- **O fator z arredondado (1,645)** muda o arredondamento no lote em alguns SKUs. A referência em Python usa o mesmo
  NORM.S.INV da planilha.
