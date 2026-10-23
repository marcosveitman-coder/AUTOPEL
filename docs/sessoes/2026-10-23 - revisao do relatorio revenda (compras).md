# Sessão 23/10/2026 — revisão do "Relatorio Revenda atualizado" (projeção de compras)

Registro sem nomes de compradores, fornecedores ou valores em R$ (repositório público).

## Pedido
Planilha `Planejamento comercial_AUTOPEL\2026\COMPRAS\Relatorio Revenda atualizado.xlsx`, aba "Geral (arq. …)":
- revisar o arquivo inteiro e corrigir fórmulas e contas, inclusive a programação de compra do trimestre (mês vigente,
  próximo e subsequente), respeitando 2 meses de estoque = PontoPedido;
- fórmula condicional (SE) de compra por mês com as variáveis Comprado (R), Disponível (O), VendProt (P),
  MedContr (I) e PontoPedido (E); com PontoPedido 0 não precisa ter estoque;
- revisar a aba PLANO de AÇAO (dinâmica impressa para os compradores) e o tracking da aba dinamicas (a partir de AC);
- explicar a coluna F "Comprar" (por que a linha 5 dá −1416 e não +1416). Liberou criar abas e dinâmicas novas.
- As colunas A–X vêm do TOTVS e não podem ser alteradas.

## Como o TOTVS calcula as colunas (conferido em todas as linhas)
- `F Comprar = E − (O − P + R)`: bate em 4.218 de 4.223 SKUs. O sinal é "o que preciso ter − o que vou ter":
  positivo = falta (comprar); negativo = sobra acima do ponto de pedido. Linha 5: saldo 571 − 255 + 1.100 = 1.416 e
  PontoPedido 0, então 0 − 1.416 = −1.416. Linha 29: 98 − 87 = +11.
- `O Disponível = M Físico − N Reserva`; `Q Ruptura = MÍNIMO(0; O − P)`; `W EstDias = (O − P) / I × 30`;
  `H = F × G`; `K = J × G`; `L` usa custo médio (≠ M × G). G (Ult.Preco) é o último preço de compra.
- 5 SKUs de quantidade fracionada vieram com a **vírgula perdida na coluna O** (163,66 → 16366). Neles só o F está
  certo.

## Problemas encontrados
1. **Tracking (dinamicas AF:AH) somava só os SKUs com Comprar > 0 hoje.** Ele somava a dinâmica filtrada em
   "AUX COMPRAR = SIM", e os 934 SKUs que só precisam de compra nos meses seguintes ficavam de fora. Por isso novembro e
   dezembro davam iguais (quem já está abaixo do ponto compra 1 MedContr por mês). Corrigido, novembro subiu ~37% e
   dezembro ~86%.
2. **Mês vigente com PontoPedido 0 dava sempre 0**, mesmo com venda feita sem estoque. Eram 9 SKUs; o TOTVS manda
   comprar esses.
3. A conta das colunas de previsão usava a coluna O (vírgula perdida) e zerava a compra de 3 SKUs.
4. As colunas GROK (AN:AS), GPT (AT:AY) e deepSEAK (AZ:BE) eram **a mesma conta** com outra escrita. As dinâmicas usam
   só a GPT.
5. "divergencia COMPRAR" (AK) marcava estoque acima de 3 meses, não divergência.
6. BW/BX (% contra Oceano/Kalunga): 481 #VALOR!, porque `"" > 0` é VERDADEIRO no Excel.
7. BK/BL (família): 1 #N/D.
8. dinamicas AJ (rupturas): 0 digitado no lugar da fórmula para quem não tem ruptura.
9. Não corrigidos, só registrados na aba LEIA-ME:
   - 177 SKUs com PontoPedido 0 e venda;
   - 137 com PontoPedido ≠ 2 × MedContr (regra da AJ);
   - 215 com MedContr ≥ PontoPedido;
   - VendProt negativo (18) e Comprado negativo (1);
   - a área V:AI da PLANO de AÇAO é de uma rodada antiga (agosto–outubro);
   - os rótulos de cabeçalho das colunas CB/CC e AI não batem com a fórmula.

## Entregas (no próprio arquivo, pelo Excel COM)
Script: `scripts cruzamento/avulsos/compras/revisa_relatorio_revenda.ps1`.
- **Table1** (aba Geral):
  - **Saldo hoje (Disp-VendProt+Comprado)** (CG) = PontoPedido − Comprar. É o mesmo número do TOTVS e não herda a
    vírgula perdida da O.
  - **AT (mês vigente)** = `SE(E=0; MÁXIMO(0; −Saldo); MÁXIMO(0; E − Saldo))`. É o Comprar quando positivo.
  - **AV (próximo)** = `SE(E=0; 0; MÁXIMO(0; E − (Saldo + AT − I)))`.
  - **AX (subsequente)** = `SE(E=0; 0; MÁXIMO(0; E − (Saldo + AT + AV − 2I)))`. AU/AW/AY (R$) = qtd × G, sem mudança.
  - Colunas novas: Compra trimestre QTD (CH) e R$ (CI), Situacao compra (CJ: 1 ruptura, 2 comprar agora,
    3 comprar nos próximos meses, 4 estoque ok, 5 sem ponto de pedido). Totais SUBTOTAL na linha 1.
  - AK passou a marcar "divergente" quando F ≠ E − (O − P + R): marca os 5 SKUs da vírgula perdida.
  - BW/BX com `N()`; BK/BL com "SEM FAMILIA".
  - GROK e deepSEAK ficaram como estavam (podem ser apagadas).
- **dinamicas**: AF:AH = `SOMASES` na Table1 por "Comprador 2" (todos os SKUs); AJ com `SEERRO(…;0)`. As colunas
  Marcas e SKUs continuam vindo das dinâmicas (Comprar > 0), para comparar com os dias anteriores digitados.
- **Aba nova "LEIA-ME COMPRAS"**:
  - calculadora: digita-se o código do SKU e a aba mostra E, I, M, N, O, P, R, o saldo, a conta do Comprar, o
    Comprar do TOTVS, a leitura e a compra dos 3 meses;
  - explicação do −1416 montada por fórmula;
  - regras da programação, lista do que foi corrigido e contagens de pontos de atenção.
- **Aba nova "PROGRAMACAO TRIMESTRE"**:
  - resumo por comprador com matriz dinâmica (`CLASSIFICAR(ÚNICO(FILTRO()))`) e mês no cabeçalho por `HOJE()`;
  - tabela dinâmica "TD Programacao trimestre" (comprador × marca, filtro de situação) na cache da Table1.
- A cache da Table1 foi atualizada; as dinâmicas (vf), PLANO de AÇAO e dinamicas já mostram a conta nova.

## Conferência (`compras/confere_revenda.py`)
- Geral A:X: 0 célula diferente.
- Programação refeita em Python para as 4.223 linhas: 0 divergência (saldo, 3 meses, R$, trimestre, situação, AK).
- Valores mudaram só onde era esperado:
  - colunas GPT: 11 SKUs no mês vigente, 2 no próximo e 3 no subsequente;
  - dinâmicas das abas (vf), PLANO de AÇAO e dinamicas (AT:AY e o tracking);
  - FABI (oculta): o item #N/D virou "SEM FAMILIA".
- Nenhuma outra aba mudou. Abas: 38 + 2. Dinâmicas: 64 + 1.
- Visual conferido por imagem (CopyPicture → gráfico → PNG, porque não há renderizador de PDF aqui).

## Decisões
- Cabeçalhos com OUTUBRO/NOVEMBRO/DEZEMBRO **não foram renomeados**: renomear coluna da tabela tira o campo das
  ~12 dinâmicas que o usam. A conta é sempre mês vigente/próximo/subsequente; o rótulo é só nome.
- Na pasta há também `Relatorio Revenda atualizado....xlsx` (com reticências), aberto no Excel com alterações não
  salvas. Não foi mexido.
- Backup do original só no scratchpad. O OneDrive guarda o histórico de versões.

## Pendências / sugestões
- Tornar os cabeçalhos neutros ("mês vigente/próximo/subsequente") e recolocar os campos nas dinâmicas pelo Excel.
- Revisar no TOTVS os 177 SKUs com PontoPedido 0 e venda, e a importação da coluna O (vírgula).
- Apagar GROK/deepSEAK se não forem mais usadas.

## Lições
- **Formato de número pelo PowerShell (COM):** o Excel lê `NumberFormat` no idioma local, mesmo com a cultura da
  thread em en-US: `'#,##0'` virou `1416,0`. Montar o formato com `$xl.International(4)` (milhar) e
  `International(3)` (decimal) e gravar em `NumberFormatLocal`.
- openpyxl lê fórmula matricial como objeto `ArrayFormula`, sem `==`. Comparar `.text` para não acusar diferença falsa.
- Somar uma tabela dinâmica filtrada para fazer total de outro critério esconde linhas: preferir `SOMASES` direto na
  tabela de origem.
