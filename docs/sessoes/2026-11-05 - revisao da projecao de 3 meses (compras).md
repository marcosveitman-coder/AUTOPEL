# Sessão 05/11/2026 — revisão da projeção de compra de 3 meses com os números atuais

Registro sem nomes de compradores, clientes ou fornecedores (repositório público).

## Pedido
- O usuário confirmou que eu lembrava da projeção de 3 meses (ruptura, disponível, comprado) e perguntou quais colunas
  são minhas e quais são dele, na aba "Geral (arq. …)".
- Depois: "revise o seu racional com base nos números atuais". A aba Geral foi atualizada em 04/11 (15h08).

## Colunas
- **Minhas:**
  - AT–AY: compra do mês vigente, do próximo mês e do mês subsequente, em QTD e R$. Eram as colunas "GPT previsao" do
    usuário; a conta foi corrigida em 23/10;
  - CI Saldo hoje, CJ/CK Compra trimestre, CL Situação compra;
  - CM–CR CLAUDE e CS–DE auxiliares.
- **Do usuário:**
  - CG PB, CH Coluna1;
  - AN–AS GROK e AZ–BE deepSEAK: cópias antigas, com a mesma conta da AT–AY, mas sem o caso PontoPedido 0 e usando a
    coluna O;
  - colunas novas DF "aux estoque" e DG "TROCA". A DF tem 1.588 #N/D (procura a marca na aba consultas).

## O que mudou nos dados (26/10 → 04/11)
- 4.241 SKUs: 18 novos e 1 que saiu.
- MedContr mudou em 2.058 SKUs (agora agosto–outubro) e o PontoPedido em 2.011.
- Valores:
  - MedContr × preço: R$ 17,2 mi → R$ 18,2 mi;
  - PontoPedido × preço: R$ 29,0 mi → R$ 30,9 mi;
  - estoque: R$ 30,0 mi → R$ 32,6 mi;
  - Comprado: R$ 28,7 mi → R$ 29,2 mi;
  - **VendProt: R$ 1,3 mi → R$ 2,8 mi** (116 SKUs com pedido de venda maior que 1 mês).
- A coluna O já não tem a vírgula perdida: o Comprar bate com E − (O − P + R) em todas as linhas.

## Conferência
`avulsos/compras/revisa_projecao.py` refaz o saldo e os 3 meses em Python: 0 divergência nas 4.241 linhas. As
fórmulas estão intactas e atualizaram sozinhas.

## Números de hoje (regra atual × CLAUDE)
| | Novembro | Dezembro | Janeiro | Total |
|---|---|---|---|---|
| Regra atual (AT–AY) | R$ 3,0 mi | R$ 7,2 mi | R$ 12,4 mi | R$ 22,6 mi |
| CLAUDE (CM–CR) | R$ 4,7 mi | R$ 6,9 mi | R$ 7,2 mi | R$ 18,7 mi |

Situação: 74 em ruptura, 381 comprar agora, 888 comprar nos próximos meses, 898 estoque ok, 2.000 sem ponto de pedido.

## Revisão do racional
A conta está certa, mas depende de hipóteses que os números de hoje põem à prova.

1. **Todo o Comprado conta como se fosse chegar.**
   - R$ 3,75 mi (92 SKUs) estão com a entrega vencida e R$ 4,0 mi (25 SKUs) estão sem data;
   - 320 SKUs só estão no ponto por causa do Comprado (o físico está abaixo);
   - 133 têm físico menor que 1 mês e nenhuma compra no mês;
   - sem os pedidos vencidos ou sem data, a compra de novembro iria de R$ 3,0 mi para R$ 7,4 mi.
   - Não tirar automaticamente (são pedidos reais), mas mostrar.
2. **Ruptura nem sempre é compra.** Das 74 rupturas:
   - 35 não têm compra porque o Comprado já cobre: a ação é cobrar o fornecedor;
   - 25 são PontoPedido 0 (sob encomenda).
3. **O crescimento de dezembro e janeiro está certo.**
   - R$ 4,1 mi por mês vêm de SKUs que já estão no ponto e repõem 1 MedContr;
   - o resto (R$ 3,1 mi em dezembro e R$ 8,3 mi em janeiro) é de SKUs acima do ponto que vão cair quando os R$ 20,3 mi
     com entrega em novembro forem vendidos;
   - a regra atual não tem sazonalidade nem desconta os dias já passados do mês. A CLAUDE tem as duas coisas.
4. **PontoPedido manual.**
   - 121 SKUs têm PontoPedido diferente de 2 × MedContr, a maioria itens grandes com números redondos;
   - a regra respeita o TOTVS. Com PontoPedido = 2 × MedContr, novembro subiria R$ 1,8 mi;
   - 191 SKUs têm PontoPedido 0 e venda (eram 177 em 23/10).

## Sugestões (aguardando decisão do usuário)
- Na Situação compra, separar:
  - "1 RUPTURA – cobrar PC" (o Comprado já cobre) de "1 RUPTURA – comprar";
  - criar "AGUARDANDO PC VENCIDO" para os SKUs que só estão no ponto por causa de pedido vencido ou sem data.
- Coluna de sensibilidade: compra do mês vigente se o pedido vencido não chegar.
- Na regra atual, descontar só os dias que faltam do mês vigente (como a CLAUDE).

## Pendente
- Nada foi gravado na planilha: só leitura (cópia no scratchpad).
- A aba de análise CLAUDE (seções 2 a 9) tem os números de 26/10. As colunas CLAUDE recalculam sozinhas, e a
  ParamClaude deve ser refeita no fim de novembro.

## Lições
- Conferir PontoPedido × MedContr com `round` do Python dá falso alarme (arredonda 80,5 para 80). Usar |E − 2I| ≤ 1.
