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

---

## 2ª parte: cabeçalhos sem nome de mês

### Pedido
"Renomeie os cabeçalhos das colunas de previsão para mês vigente / próximo / subsequente e recoloque os campos em todas
as dinâmicas, mantendo o layout das abas impressas."

### Por que não pelo Excel
Renomear a coluna da tabela e atualizar tira o campo das dinâmicas. Recolocar o campo pelo Excel apagaria a formatação
própria daquelas colunas: as dinâmicas das abas (vf) e do PLANO têm de 600 a 42 mil formatações gravadas cada uma.

### Como foi feito
1. `compras/renomeia_previsao.py` edita o XML, tudo junto:
   - nome das 18 colunas na Table1 e cabeçalho da linha 2 (texto em linha, para não mexer em texto compartilhado com
     as tabelas antigas das abas ocultas "Ruptura … 14", que têm colunas de mesmo nome);
   - **nome do campo na cache da dinâmica da Table1**. As dinâmicas guardam o campo pela posição na cache, então
     continuam com ele no mesmo lugar e com a mesma formatação;
   - referências `Table1[...]` nas fórmulas: 19 por linha da tabela, mais dinamicas, LEIA-ME e PROGRAMACAO. Só as
     referências à Table1; as outras tabelas ficaram como estavam;
   - títulos dos campos de valor nas 11 dinâmicas do usuário.
2. `compras/renomeia_previsao.ps1` (Excel):
   - confere que as 12 dinâmicas (11 + a da PROGRAMACAO) reconhecem os campos, na mesma posição;
   - atualiza a cache e devolve as larguras de coluna que a atualização ajusta (70);
   - troca os textos da LEIA-ME e põe o nome do mês por `HOJE()` nos cabeçalhos do tracking (AF4:AH4 e AG21:AI21).

### Nomes novos
| Antes | Agora |
|---|---|
| GPT previsao OUTUBRO / NOVEMBRO / DEZEMVRO QNTDD | Compra mês vigente / próximo mês / mês subsequente QTD |
| GPT previsao … R$ | Compra mês vigente / próximo mês / mês subsequente R$ |
| GROK / deepSEAK previsao … | GROK / deepSEAK mês vigente / próximo mês / mês subsequente QTD ou R$ |
| Títulos nas dinâmicas: "Soma de GPT previsao …" e ".previsao …" | .Mês vigente QTD, .Mês vigente R$, .Próximo mês QTD, .Próximo mês R$, .Mês subsequente QTD, .Mês subsequente R$ |

### Conferência (`compras/confere_renomeia.py`)
- Valores: só mudaram os 18 cabeçalhos, os 6 títulos de cada dinâmica e 4 textos da LEIA-ME. Todos os números ficaram
  iguais.
- Abas do usuário:
  - estilo de cada célula igual, comparado pelo conteúdo do estilo, porque o Excel renumera os estilos ao salvar e só
    juntou fontes que diferiam na dica `<family>`;
  - larguras de coluna e alturas de linha iguais;
  - formatação e campos das 65 dinâmicas iguais.
- Nenhuma fórmula ficou com o nome antigo. Imagens antes × depois da (vf), do PLANO e do tracking conferidas.

### Lições
- **Trocar o nome de uma coluna que alimenta dinâmicas sem perder campo nem formatação:** renomear junto, no XML, a
  coluna da tabela e o `cacheField` da cache. As dinâmicas apontam para o índice do campo.
- **No PowerShell, `$pt.DataFields` não enumera:** usar `$pt.DataFields()` e `.Item(i)`.
- **Não gravar `RowHeight` para "devolver" altura:** isso fixa a altura (`customHeight`) de linhas que eram
  automáticas. Devolver só as larguras.
- **`Select-Object -First` num comando que chama o script do Excel interrompe o script antes de salvar** e deixa um
  Excel invisível aberto. Conferir `Get-Process EXCEL` e fechar só o processo sem janela que o script abriu.

---

## 3ª parte (15h): por que 5 linhas dão FALSO na conferência do usuário

### Situação do arquivo
- O usuário apagou o `Relatorio Revenda atualizado....xlsx` (com reticências). O arquivo que ficou (salvo às 15h21) é a
  **versão de antes da revisão** mais 2 colunas novas do usuário na Table1:
  - CG "PB" = `−O + P − R + E`;
  - CH "Coluna1" = `PB = Comprar.`.
- Não estão mais no arquivo:
  - as abas LEIA-ME COMPRAS e PROGRAMACAO TRIMESTRE;
  - as colunas novas (Saldo hoje, Compra trimestre, Situação);
  - as fórmulas corrigidas e os nomes "Compra mês vigente…": voltou "GPT previsao OUTUBRO…";
  - a dinâmica da programação (agora são 64).
- Ficou para o usuário decidir se reaplica a revisão. O `revisa_relatorio_revenda.ps1` recusa a Table1 com 86 colunas
  e precisa aceitar as colunas novas do usuário.

### Diagnóstico (`compras/diagnostico_disponivel.py`)
- Os 5 FALSO (linhas 529, 658, 1665, 1780, 2808) são **exatamente os 5 SKUs, entre 4.223, com Disponível
  fracionado**. Na cópia de ontem (aba oculta O.N.T.E.M), são os mesmos 5 SKUs.
- A coluna O vem do TOTVS **sem o separador decimal, com os dígitos colados**: 163,66 → 16366 (×100),
  622,6 → 6226 (×10), 67,4001 → 674001 (×10000).
- O TOTVS tem o valor certo:
  - o Comprar (F) sai com decimal e bate com o Disponível verdadeiro (E − F + P − R);
  - o Vlr Estoque usa o estoque fracionado (ex.: 67,4001 × preço);
  - Físico e Reserva vêm arredondados (164 × 163,66).
- O EstDias (W) também vem errado pelo mesmo fator (ex.: 54.180 dias em vez de 5.418). O erro já está na saída do
  relatório, antes do EstDias ser calculado; não é da colagem no Excel nem da fórmula do usuário.
- Na cópia de ontem, 13 SKUs tinham decimal no PontoPedido (E) e bateram. O problema é só na coluna Disponível.

### Recomendação dada
- Pedir ao TI/TOTVS para revisar o formato do campo Disponível (e do EstDias) no relatório. Algo tira o separador
  decimal, por exemplo uma troca de "." para tirar o ponto de milhar.
- Na planilha, o saldo confiável é E − F. Para a CH acusar só divergência de verdade, comparar com o Disponível
  verdadeiro ou com tolerância.
- Conferir no cadastro por que caneta, lápis e envelope têm estoque fracionado (fator de conversão de unidade).

---

## 4ª parte (16h): revisão reaplicada no arquivo atual

### Pedido
"Reaplique a revisão (fórmulas do trimestre, tracking, abas LEIA-ME e PROGRAMACAO, nomes mês vigente/próximo/
subsequente) neste arquivo, mantendo as minhas colunas PB e Coluna1."

### Ponto de partida
- O arquivo das 15h21 é **idêntico ao original da manhã**, fora as colunas do usuário CG "PB" e CH "Coluna1". Isso foi
  conferido célula a célula em todas as abas.
- A Table1 estava **filtrada em Coluna1 = VERDADEIRO**, com as 5 linhas FALSO ocultas.
- Na linha 529 a PB tem uma fórmula diferente da coluna (com "+ Fisico"), um teste do usuário. Foi mantida.

### O que deu errado na 1ª tentativa e foi corrigido
Com o filtro ligado, gravar a fórmula de uma **coluna nova** pelo Excel (`DataBodyRange.Formula`) **pula as linhas
ocultas**:
- as 4 colunas novas ficaram sem fórmula nas 5 linhas;
- a tabela deixou de tratá-las como coluna calculada;
- as colunas que já existiam foram gravadas por inteiro.

O `revisa_relatorio_revenda.ps1` agora:
- guarda o filtro, mostra tudo, grava e devolve o mesmo filtro (4.218 linhas visíveis antes e depois);
- confere que toda linha tem a fórmula nas colunas que grava;
- aceita colunas do usuário depois da CF.

### Passos
1. `revisa_relatorio_revenda.ps1`. As colunas novas ficaram depois das do usuário: CI Saldo hoje, CJ Compra trimestre
   QTD, CK Compra trimestre R$, CL Situacao compra.
2. `renomeia_previsao.py` + `renomeia_previsao.ps1`: as mesmas contagens da 1ª vez (19 referências nas fórmulas da
   tabela, 80.237 nas células, 12 dinâmicas).
3. **`devolve_layout.py` (novo)**: copia do arquivo do usuário as larguras de coluna e as alturas de linha. A
   atualização das dinâmicas tinha mexido em 4 larguras (Planilha2, dinamicas) e em 7 alturas automáticas (abas (vf)).

### Conferência
- Contra o arquivo do usuário:
  - A–X, PB e Coluna1 iguais;
  - a programação refeita em Python bate em todas as linhas;
  - layout igual em todas as abas: estilo de cada célula, larguras, alturas e formatação das 65 dinâmicas. Só as
    colunas novas da Geral têm formato novo.
- Contra a entrega validada das 13h56: todas as abas iguais. Na Geral, só diferem:
  - a letra dos totais da linha 1 das colunas de trimestre;
  - 8 totais da linha 1, que usam SUBTOTAL e agora somam só as linhas visíveis do filtro.
- O Excel abre sem reparo: 40 abas, 65 dinâmicas, Table1 com 90 colunas, filtro ligado.

### Lições
- **Antes de gravar coluna de tabela pelo Excel, ver se a tabela tem filtro**: no XML, `<filterColumn>` e linhas
  `hidden="1"`. Com filtro, a coluna nova fica sem fórmula nas linhas ocultas e deixa de ser coluna calculada.
- **Atualizar dinâmicas pelo Excel mexe em larguras e alturas automáticas**: no fim, devolver as do arquivo do usuário
  pelo XML (`devolve_layout.py`). Copiar só os atributos, porque o `style` da coluna é índice e muda de número.

---

## 5ª parte: conferência depois do desligamento forçado do computador

- O arquivo na pasta foi salvo pelo usuário às 17h17, depois da gravação das 16h01. O zip está íntegro e o Excel abre
  sem reparo: 40 abas, 65 dinâmicas, Table1 com 90 colunas.
- Comparado célula a célula com a versão conferida das 16h01:
  - A–X iguais;
  - fórmulas e valores iguais em todas as abas;
  - a programação refeita em Python bate em todas as linhas;
  - formatação e campos das dinâmicas iguais.
- Diferenças, todas do usuário:
  - o filtro Coluna1 = VERDADEIRO foi tirado e todas as linhas estão visíveis, então os 8 totais SUBTOTAL da linha 1
    voltaram ao total cheio;
  - as larguras das colunas T:CG da Geral foram reajustadas;
  - uma altura automática de linha mudou na GERAL (vf).
- Conclusão: tudo o que foi pedido está salvo no arquivo.
