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

## Entregas (gravadas no arquivo às 11h47)
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
- ~~Gravar no arquivo~~: feito às 11h47, depois que o usuário fechou o Excel.
  - O arquivo ainda ficou travado alguns segundos pelo antivírus (Kaspersky, achado com o Restart Manager do
    Windows).
  - Antes de gravar: o arquivo estava inalterado desde as 10h05.
  - Depois: idêntico à cópia conferida; abre sem reparo (42 abas, 65 dinâmicas, Table1 com 109 colunas).
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
- **Arquivo travado sem Excel aberto:** listar quem segura o arquivo com o Restart Manager (`rstrtmgr.dll`,
  RmGetList). Aqui era o antivírus examinando o arquivo; bastou tentar de novo em alguns segundos.

---

## 2ª parte (12h05): novo padrão de revisão

- **Decisão do usuário:** revisão semanal para a curva A (7 dias), 15 para a B e 30 para a C.
- **Aplicado com `ajusta_revisao_claude.ps1`:**
  - param CLAUDE: B5:D5 = 7 / 15 / 30 e a descrição;
  - análise: 5 textos e as tabelas 2.3 (curva ABC) e 4.1 (cenários), refeitas pelo `exporta_claude.py`.
- `calcula_claude.py`, `grava_claude.ps1` e `exporta_claude.py` passam a usar 7 / 15 / 30 como padrão.
- **Números em 26/10:**
  - estoque alvo R$ 32,9 mi (regra atual R$ 29,0 mi; estoque hoje R$ 30,0 mi);
  - compras outubro / novembro / dezembro: R$ 4,8 / 2,1 / 9,6 mi;
  - alvo em meses de venda: A 1,8 · B 2,1 · C 2,7.
- **Conferência:**
  - colunas CLAUDE × Python: 0 diferença;
  - só mudaram 3 células da param e 63 da análise;
  - layout devolvido; gravado às 12h05, idêntico à cópia conferida.

---

## 3ª parte: exploração completa do MRP em HTML

- **Pedido:** abrir `Downloads\Autopel_MRP_23102026_atualizado (1).html` no navegador, clicar em todos os botões e
  analisar 100% do conteúdo. O arquivo é idêntico ao `..._atualizado.html` (mesmo MD5).
- **Como foi aberto:** o navegador embutido não abre `file://`. Uma cópia foi servida em `http://localhost:8765` com
  `python -m http.server` (configuração `mrp-html` em `.claude/launch.json`, fora do git).
- **O que foi testado:**
  - os 2 modos (Política/MRP e Saúde/IQR), as 32 abas e os 7 atalhos da "Rotina do comprador" (que só levam a abas
    que já existem);
  - os 6 cartões do painel e os 14 focos do diagnóstico;
  - os 5 recortes de saúde, a busca, a paginação (84 páginas de 50) e as 8 ordenações;
  - o detalhe do SKU (janela com memória de cálculo, ordens de compra e campos editáveis por SKU);
  - o simulador e o desconto de ordens de compra;
  - as 30 exportações, salvar, carregar e restaurar cenário, e imprimir.
  - Os downloads foram interceptados na página (nome, tamanho e cabeçalho), sem gravar arquivos.

### Números do MRP (posição de 23/10; faturamento até 14/10)
- **Recorte e estoque:** 4.174 SKUs RV ativos, filial 05 / armazém 40; estoque de R$ 25,6 mi.
- **Política:** calculável em 1.750 SKUs (R$ 21,3 mi de estoque); 2.424 ficam sem política (R$ 4,4 mi):
  - 1.392 sem vínculo no faturamento;
  - 1.027 sem saída recente;
  - 5 sem custo.
- **Compras:**
  - compra técnica R$ 7,6 mi; compra proposta (após o filtro de concentração de clientes) R$ 1,8 mi, em 239 SKUs;
  - estoque-alvo dos calculáveis R$ 20,5 mi;
  - excesso físico R$ 8,4 mi; 231 SKUs abaixo do ponto.
- **Saúde do inventário:**
  - IQR estrito 48,9%, ampliado (+2 meses) 59,2%;
  - excesso R$ 4,7 mi; slow (6 meses sem venda) R$ 1,4 mi; phase out (12 meses) R$ 1,2 mi; não classificado R$ 3,2 mi.
- **Outras visões:**
  - excedente livre R$ 8,4 mi, dos quais R$ 1,9 mi são genéricos candidatos a e-commerce (359 SKUs);
  - ordens de compra: 1.306 linhas, 520 PCs, R$ 27,9 mi a receber, 68 SKUs com PC vencido;
  - carteira de vendas: R$ 6,4 mi, dos quais R$ 6,2 mi com data prevista vencida (base antiga).

### Como o MRP calcula
- **Demanda:** saída positiva diária das bases de faturamento nos últimos 3 meses completos (jul–set). Sem vínculo,
  usa a MédiaVenda ou a QTD Fat do relatório de estoque (570 SKUs).
- **Segurança e ponto de pedido:**
  - segurança = z × desvio × √(LT/30), com serviço A 99%, B 97,5% e C 95%;
  - ROP = demanda diária × LT + segurança.
- **Alvo:** máximo(ROP + segurança; demanda do horizonte de 22 dias + segurança).
- **Compra:** alvo − disponível protegido (descontada a carteira), elevada a um MOQ de 10% da compra média mensal.
- **Prazo:** LT único de 15 dias para todos os SKUs. O simulador aceita no máximo 30.
- **Classes:** ABC × XYZ × PQR × concentração de clientes (1–4). A classe 4 vira MTO/contrato.

### Problemas achados (testados no simulador e no código da página)
1. **O campo "Revisão • dias" não faz nada.** Com 7, 15 ou 30 o resultado é idêntico. No código, `review` só aparece
   nos padrões e na validação; a função da política não o usa.
2. **A segurança entra duas vezes no alvo.** O ROP já contém a segurança, e o alvo soma de novo. A própria página
   chama isso de "simulação conservadora solicitada".
3. **As ordens de compra não são descontadas por padrão.**
   - O CSV vem sem filial. Atribuindo à 05, a compra técnica cai de R$ 7,6 mi para R$ 3,2 mi; usando o PED_EMITID do
     estoque, para R$ 2,1 mi.
   - Exemplo: um SKU com 5.425 un já pedidas recebe sugestão de mais 3.790.
4. **O prazo de 15 dias é o dobro do real** (mediana de 7 dias nas fotos diárias).
   - LT 8 → compra técnica R$ 4,5 mi, alvo R$ 15,7 mi.
   - LT 8 com serviço 95/90/80 → R$ 3,4 mi e R$ 13,8 mi.
   - LT 30 → R$ 16,4 mi e R$ 31,9 mi.
5. **A demanda vem das bases de faturamento.** Elas não cobrem toda a venda: 1.392 SKUs ficam sem política. O piso
   pela MedContr (opção do simulador, desligada) leva os calculáveis a 1.924 e a compra técnica a R$ 8,2 mi.
6. **Datas misturadas:**
   - estoque de 23/10; faturamento até 14/10; pedidos em aberto de 19/10;
   - boletim de 21/10; carteira SQL antiga, com 98% do valor vencido.
7. **Curva ABC definida de dois jeitos:**
   - a nota da classificação fala em valor de estoque;
   - a aba Grupos MRP fala em consumo × custo;
   - 2.888 SKUs receberam classificação "conservadora".
8. A comparação anual mostra +56% (2026 × 2025, jan–set), mas as linhas por mês dobraram. Pode ser mudança de
   cobertura das bases, não crescimento. A aba Planilha1 de `fat 2025.2` não foi duplicada (repetições exatas: 0).

### O que o MRP tem de bom
- Saúde do inventário com IQR e cortes de slow e phase out;
- triagem de excedente para e-commerce;
- concentração de clientes por raiz de CNPJ;
- detalhe por SKU com memória de cálculo;
- exportações completas e documentação de cada coluna (dicionário com 126 campos).
- Vale aproveitar essas visões ao lado da metodologia CLAUDE.

### Comparação com a metodologia CLAUDE

| Ponto | MRP HTML | CLAUDE (planilha) |
|---|---|---|
| Demanda | faturamento diário jul–set (bases parciais) | MedContr do TOTVS × sazonalidade |
| SKUs com política | 1.750 de 4.174 | 4.224 |
| Prazo | 15 dias para todos | mediana do fornecedor (padrão 8, teto 45) |
| Revisão | campo existe, sem efeito | A 7, B 15, C 30 |
| Serviço | 99 / 97,5 / 95 | 95 / 90 / 80 |
| Segurança | conta duas vezes | uma vez, com desvio do prazo |
| Compras abertas | não descontadas por padrão | descontadas (saldo E − F inclui o Comprado) |

### Lições
- **Navegador embutido e `file://`:** ele não abre arquivo local. Copiar o HTML para o scratchpad e servir com
  `python -m http.server` (configuração em `.claude/launch.json` + `preview_start`).
- **Página que redesenha a cada mudança:**
  - buscar o elemento de novo depois de cada render; o `select` antigo fica solto e a troca não tem efeito;
  - quebrar os testes em chamadas curtas, porque o script morre em 45 s;
  - guardar os resultados em `window` para ler depois.
- **Botões de download:** substituir `URL.createObjectURL` e `HTMLAnchorElement.prototype.click` dá nome, tamanho e
  conteúdo do arquivo sem baixar nada.

---

## 4ª parte: saúde do estoque, e-commerce e concentração de clientes na aba de análise

- **Pedido:** levar para a aba "analise compras CLAUDE" a saúde do inventário (IQR, slow, phase out), a triagem de
  excedente para e-commerce e a concentração de clientes, calculadas com os dados da planilha.
- **Partida:** o arquivo salvo pelo usuário às 14h55 (diferente do das 12h05: o usuário mexeu depois).
- **Scripts novos em `avulsos/compras/claude/`:**
  - `saude_ecommerce.py`: calcula por SKU e gera `saude.json` e `param_saude.csv`;
  - `grava_saude.ps1`: grava pelo Excel; recusa se a seção 7 ou as colunas já existirem;
  - `confere_saude.py`: confere partida × gravado.

### Regras
- **Venda recente:** venda no TOTVS (todos os canais, abril a setembro, das fotos diárias) ou nas bases de faturamento
  nos últimos 6 meses.
- **Saúde do estoque:**
  - o comprometido (Reserva + VendProt, até o físico) conta como saudável;
  - saudável estrito = comprometido + alvo CLAUDE do mês vigente; ampliado = + 2 meses da previsão CLAUDE;
  - acima disso é excesso;
  - sem venda em 6 meses: slow moving se as bases mostram venda nos 12 meses, senão phase out potencial;
  - sem físico ou sem valor: não classificado.
- **Excedente livre** = excesso + slow + phase out. Triagem:
  - personalizado: nome de cliente da venda na descrição, PERSONALIZ/LOGO ou grupo 0015;
  - concentrado: 1 cliente com 80% ou mais em 12 meses;
  - revisar requisitos: alimento, EPI/saúde, elétrico ou volumoso;
  - genérico: o resto.
- **Preço de mercado:** colunas Oceano/Kalunga/Mercado Livre da Geral, só entre 0,3 e 5 vezes o custo (fora disso a
  pesquisa achou outro produto).
- **Concentração de clientes:**
  - quantidade por raiz de CNPJ, de outubro/2025 a setembro/2026, nas bases de faturamento, sem a cópia de dezembro
    da Planilha1;
  - coluna de cobertura = venda nas bases ÷ venda do TOTVS (abril a setembro).
- **Nomes dos grupos** (dados por esta análise, olhando as descrições):
  - 0001 informática; 0002 papelaria; 0003 bobinas; 0004 limpeza; 0005 papéis; 0006 etiquetas e ribbons;
  - 0013 mobiliário e eletro; 0015 personalizados; 0016 EPI; 0017 copa e cozinha.

### Números (26/10)
- **Saúde do estoque** (R$ 30,0 mi): IQR estrito 65,8% e ampliado 78,8%.
  - Fora da política: R$ 6,3 mi = excesso R$ 3,9 mi + slow R$ 0,5 mi + phase out R$ 1,9 mi.
  - Curva C: R$ 5,4 mi do que está fora (IQR ampliado 35%); A 98% e B 91%.
  - O MRP em HTML dá 49% / 59% porque só calcula 1.750 SKUs.
- **E-commerce:**
  - genérico R$ 2,6 mi (1.447 SKUs; os 40 maiores somam R$ 1,3 mi);
  - revisar requisitos R$ 0,8 mi; concentrado R$ 1,7 mi; personalizado R$ 1,3 mi.
- **Concentração de clientes:**
  - 1.638 SKUs têm cliente nas bases;
  - 661 vendem para 1 cliente só (R$ 5,4 mi de estoque, R$ 3,4 mi da compra CLAUDE de outubro a dezembro);
  - 1.143 SKUs têm 1 cliente com 80% ou mais (R$ 17,8 mi de estoque);
  - uma única rede de varejo concentra 426 SKUs, R$ 11,8 mi de estoque (39%) e R$ 7,5 mi da compra CLAUDE.

### Entregas (no arquivo `final.xlsx` do scratchpad, conferido)
- **Aba de análise:** seções 7, 8 e 9 nas linhas 234 a 445, depois da seção 6.
  - 7: resumo, 7.1 curva, 7.2 comprador, 7.3 grupo, 7.4 30 SKUs fora da política;
  - 8: triagem, 8.1 40 genéricos;
  - 9: classes de clientes, 9.1 clientes que concentram, 9.2 30 SKUs concentrados.
- **ParamClaude:** 9 colunas novas no fim:
  - Última venda (mês), Saúde do estoque, Fora da política QTD/R$, Triagem e-commerce;
  - Clientes 12m, % maior cliente, Maior cliente, Classe clientes.
- **Conferência:**
  - todas as outras abas iguais, em valores e fórmulas;
  - linhas antigas da análise e colunas antigas da param iguais;
  - 10 tabelas e 4.224 SKUs conferidos;
  - 65 dinâmicas, 12 tabelas e 1 gráfico mantidos;
  - layout das outras abas devolvido com `devolve_layout.py --exceto` (só 1 altura de linha voltou).

### Pendente
- **Gravar no arquivo original:** o usuário estava com o arquivo aberto no Excel. Antes de gravar, conferir que o
  arquivo ainda é o das 14h55. Se mudou, refazer a partir do atual:
  1. `saude_ecommerce.py`;
  2. `grava_saude.ps1`;
  3. `devolve_layout.py --exceto "analise compras CLAUDE" "param CLAUDE"`;
  4. `confere_saude.py`.

### Lições
- **.ps1 com acento** precisa ser salvo com BOM (UTF-8): sem ele o PowerShell 5.1 lê "SUGESTÕES" errado e a busca
  pelo texto falha.
- **openpyxl e fórmula matricial:** compara objetos `ArrayFormula`; comparar `ref` e `text`.
- **Texto `AAAA-MM` gravado pelo COM:** vira data. Formatar a coluna como texto (`@`) ou prefixar com apóstrofo.
