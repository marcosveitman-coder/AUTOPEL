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
