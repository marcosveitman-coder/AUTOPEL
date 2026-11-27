# Sessão 25/11/2026 — dá para refazer o "Relatorio de Revenda" sem o Power BI?

Registro sem nomes de pessoas, clientes ou fornecedores (repositório público).

## Pergunta do usuário
O TOTVS gera 3 relatórios por dia (EstoqueNecessidade, HistoricodeCompras, OrdemCompra). O Power BI os importa, faz as contas
e exporta o "Relatorio de Revenda dd-mm-aaaa.xlsx" (aba Geral + 1 aba por comprador). Com os 3 relatórios de entrada, dá para
chegar na saída sem o Power BI?

## Resposta: quase (22 das 24 colunas; falta a MedVenda)
Testado com os arquivos de 25/11 (`avulsos/compras/reproduz_revenda.py`). Os números abaixo são das 4.266 linhas da Geral.

- **Universo (4.266 SKUs, igual à saída):** código com linha na filial 05, ATIVO, Tipo RV ou PA, armazém 40 (ou 10/01 se o
  código tem também outra linha). Os 957 códigos que ficam fora são MP, BN, PI, INATIVO e códigos só das filiais 07/01.
  Valores = soma das linhas da filial 05 (65 códigos têm mais de uma linha).
- **Cópia direta do EstoqueNecessidade (100%):** Comprador (Usuario, 10 letras), Descricao, Marca, Grupo, DT_ENTREGA, PEDIDO,
  PontoPedido (ESTQ_MININ), Fisico, Vlr Estoque, Reserva, Disponivel, VendProt (Qtd.Ped.Vend), Comprado (PED_EMITID) e
  Comprar. (TT_COMPRAR).
- **Contas que o Power BI faz (100%):**
  - MedContr = **teto** da média de MES_01..03;
  - Ult.Preco e Prc_Compra. = **teto a 2 casas**;
  - Valor Total = Comprar. × Ult.Preco;
  - Ruptura = MÍNIMO(0; Disponivel − VendProt);
  - EstDias = MÁXIMO(0; Disponivel − VendProt) ÷ MedContr × 30 (0 sem MedContr);
  - ULT_COMPRA = a data mais recente das linhas (4.262 de 4.266; as 4 diferenças são o campo em branco).
- **Abas por comprador:** o mesmo conjunto de SKUs da Geral por comprador (conferido), organizado por marca com subtotal.
- **O que não sai das entradas:** **MedVenda** (preenchida em 55% dos SKUs; tem valor em 260 SKUs sem venda nos MES e fica
  vazia em 312 com venda) e **Total Media** (= MedVenda × Ult.Preco). Não bate com MES_01..03, MEDIA, nem com a média das
  entradas do HistoricodeCompras em nenhuma janela de 3 a 36 meses. Vem de outra fonte (vendas), que o EstoqueNecessidade
  novo não traz (o relatório antigo do Protheus trazia).
- **HistoricodeCompras e OrdemCompra não entram nas 24 colunas da Geral:** Comprado, PEDIDO e DT_ENTREGA vêm do
  EstoqueNecessidade. Esses dois relatórios alimentam outras páginas do Power BI (follow-up de pedidos, últimas entradas),
  que não foram testadas aqui.
- **Uso na planilha de compras:** a metodologia CLAUDE e a projeção de 3 meses usam MedContr, não MedVenda. Só a coluna CF
  (relação venda/estoque) usa a Total Media.

## Para fechar 100%
Descobrir de onde vem a MedVenda (qual medida do Power BI ou qual relatório de vendas). Com ela, a saída inteira se refaz em
Python a partir dos arquivos do TOTVS.

## Pista nova sobre a MedVenda (mesmo dia)
- No Power BI o campo se chama **MediaVenda** (ao lado de **QTD Fat**) e está nas telas de indicadores e de rastrear vendas.
- **A MedVenda não muda com o tempo:** nos 4.212 códigos comuns, o valor da aba Geral de 13/11 (arquivo de trabalho) é
  igual ao do Relatorio de Revenda de 25/11, e as 1.875 células vazias são as mesmas nos dois. O V30481 tem 814 no export do Power BI
  de 26/10 e 814 em 25/11. Uma média móvel de 3 meses mudaria todo mês: logo é uma **tabela estática** (importada uma vez ou
  digitada), e não uma medida calculada sobre as vendas do dia.
- Próximo passo: o usuário ver no Power BI de que tabela/consulta vem a coluna (painel Dados, Transformar dados).

## MedVenda descoberta (export `data.xlsx` da página Indicadores do Power BI, 25/11)
- O export traz 45 colunas, entre elas **MediaVenda**, **QTD Fat** e três colunas de venda mensal: **Mês -3**, **Mês -2**,
  **Mês -1**.
- **MedVenda (Revenda) = MediaVenda (Power BI) = TETO da média de Mês -3, Mês -2 e Mês -1.** Bate nos 2.342 SKUs com valor;
  nos 1.871 sem valor, os três meses também estão vazios. A hipótese do usuário (média dos últimos 3 meses) está certa na
  conta.
- **Os três meses são estáticos:** iguais em 100% dos 2.334 SKUs entre o export de 26/10 e o de 25/11. Também não são o
  MES_01..03 do EstoqueNecessidade de 25/11 (13 de 2.342 iguais) nem as bases de faturamento. São de uma tabela do Power BI
  carregada uma vez e não atualizada; a MedVenda do Revenda está defasada.
- **Falta:** de qual tabela/arquivo vêm as colunas Mês -3..-1 e a QTD Fat (o usuário vai olhar em Transformar dados). Com ela,
  e sabendo quais meses são, dá para atualizar a MedVenda e fechar a reprodução do Revenda.

## O que o modelo semântico do Power BI mostra (visto pelo navegador, só leitura)
- Relatório "Necessidade de Compras", modelo semântico de mesmo nome (21 tabelas, atualizado em 25/11 às 09h28). O modelo é de outro
  usuário da empresa; o usuário que pediu só tem leitura (sem permissão de gravação): a fórmula (DAX) e a consulta (Power Query)
  **não aparecem** pelo navegador, e a aba Linhagem só mostra o relatório a jusante (a fonte não aparece sem permissão).
- **Onde está a MediaVenda:** tabela **Estoque Necessidade**, como **coluna** (inteiro), não como medida.
- **Colunas vizinhas da mesma tabela:** `Mes-1`, `Mes-2`, `Mes-3` (e a Total Media), e também `Mes Julho 1`, `Mes Agosto 2`, `Mes Setembro 3`
  (com versões em R$). Os nomes com **julho, agosto e setembro** explicam por que a MediaVenda não muda: os meses estão
  escritos fixos no modelo (média de jul–set, não dos últimos 3 meses de verdade).
- A tabela de vendas do modelo é a **FatMensal** (e FatConsolidado, Fat2025, Fat2025A): EMISSAO, COD_PROD, QTD Fat, FILIAL.
  O EstoqueNecessidade, o HistoricodeCompras e a OrdemCompra aparecem como tabelas do modelo.
- Pendente: a fórmula/consulta exata (pedir ao dono do modelo ou pedir permissão de gravação) e confirmar se Mes-1..-3 são
  jul/ago/set 2026. Se forem, a MedVenda do Revenda = teto da média de jul–set (vendas da FatMensal por produto) e dá para
  refazer com as bases de faturamento de jul–set; hoje isso bate menos de 4% (as bases `fat` não são a FatMensal).

## 26/11: exploração da página "Rastrear Venda" do relatório (Chrome do usuário, só leitura, sem exportar nada)
- A página tem filtros (Codigo, Descricao, GRPVEN, Tipo, Filial, RAZAO_SOCIAL) e 3 tabelas:
  - posição de estoque (colunas do EstoqueNecessidade);
  - **pedidos de venda** (`dt_emissao`, `pedido`, `produto`, `quantidade`, `filial`): tem dados de outubro/2026 (atual);
  - **faturamento** (`EMISSAO`, descrição, `QTD Fat`, custos, razão social): as colunas da tabela **FatMensal** do modelo.
- **A tabela de faturamento está desatualizada:** filtrando o produto V30481, as notas mais recentes são de **16/01/2026**
  (de 05 a 16/01); o export antigo dela (`PowerB_restrear.venda3`, 150 mil linhas mais recentes) vai de 02/12/2025 a
  **19/01/2026**. Ou seja, a FatMensal não recebe vendas desde meados de janeiro de 2026. A coluna `QTD Fat` aparece vazia nessas
  linhas filtradas por código.
- Conclusão provisória: a MedVenda (Mes-1..-3) não pode vir de vendas recentes da FatMensal; são valores gravados no modelo
  (colunas `Mes Julho 1`, `Mes Agosto 2`, `Mes Setembro 3`), de origem ainda não identificada.
- Não deu para ver DAX/Power Query (usuário só tem leitura). Capturas de tela do Chrome travaram (janela sem renderizar); o
  DOM da página funcionou e foi lido pela árvore de acessibilidade.

## 26/11: e se o dono do modelo não responder? (alternativas)
- **Teste feito:** os valores de Mês -1, -2 e -3 (export de 25/11) foram comparados com MES_01..03 das 19 fotos diárias do
  EstoqueNecessidade (13/08 a 26/10), em 3 ordens possíveis. O melhor resultado foi 315 de 2.342 SKUs (13%): **não são uma cópia do
  MES do TOTVS de nenhuma data**. A origem continua desconhecida.
- **Caminhos sem depender do dono:**
  1. a TI ou o administrador do Power BI abre o modelo (ou assume a propriedade) e lê a fórmula;
  2. o usuário pede a permissão pelo botão "Solicitar acesso" do modelo, ou por um superior;
  3. conferir o outro relatório do mesmo espaço de trabalho ("Base Relatorio Geral…"), que pode ter modelo próprio;
  4. **redefinir a MedVenda aqui:** a MedContr (média dos 3 últimos meses fechados, do próprio TOTVS) já é a venda média atual;
     o modelo tem até a coluna `VlrUltPrecoMedContr` (MedContr × preço). Sem a MedVenda parada, a Total Media vira
     MedContr × Ult.Preco. Só a coluna CF (relação venda/estoque) usa a Total Media.

## 26/11: relatório refeito sem o Power BI (decisão do usuário: MedContr no lugar da MedVenda)
- **Decisão:** a MedVenda do Power BI vem de colunas fixas e desatualizadas (origem não identificada). No relatório novo,
  **MedVenda = MedContr** (vazia quando 0) e **Total Media = MedContr × Ult.Preco**.
- **Script:** `avulsos/compras/gera_revenda.py` (entrada: só o EstoqueNecessidade). Gera a aba Geral (24 colunas, mesma ordem de
  colunas, formatos e ordem de linhas do Power BI) e uma aba por comprador com subtotal por marca e "Total Geral".
- **Arquivo gerado (25/11):** `Downloads\COMPRAS[[\Relatorio de Revenda 25-11-2026 (sem Power BI).xlsx` (9 abas, abre no Excel).
- **Conferência contra o relatório do Power BI:**
  - 4.266 SKUs, os mesmos, **na mesma ordem** (4.266 de 4.266);
  - todas as colunas iguais, exceto MedVenda e Total Media (por decisão) e **1 célula** de ULT_COMPRA (um código com
    data em um armazém e vazio no outro: o Power BI deixa vazio, o novo traz a data);
  - abas por comprador: mesmos nomes (em ordem alfabética), mesmas linhas e subtotais (Valor Total e Vlr Estoque iguais).
- **O que mudou de verdade:** MedVenda igual nos dois em 1.851 dos 4.266 SKUs (os que não vendem). Total Media somada: R$ 14,4 mi
  (Power BI, parada) → R$ 18,3 mi (MedContr × preço). Exemplos: V30481 814 → 1.082; V15982 1.286 → 1.350; V13790 22 → vazio.
- **Regras do que vira aba:** comprador com 5 ou mais SKUs e que não seja INATIVO/OCIOSO; nome da aba = primeiro nome.
- **Como usar todo dia:** baixar o EstoqueNecessidade do TOTVS e rodar
  `python gera_revenda.py "EstoqueNecessidade (N).xlsx" "Relatorio de Revenda dd-mm-aaaa.xlsx"`.
