# 11/11/2026 – Nova divisão de segmentos na equipe maior (ORGANOGRAMA e ORGANOGRAMA impressao)

## Pedido

O usuário mandou a foto da "Divisão Segmentos Mercado" (atualização de 03/11/2026) de **uma** equipe, a maior, e pediu
para reformular as abas ORGANOGRAMA e ORGANOGRAMA impressao do `MAPA MERCADO V2.xlsm`.

A lista:
- 7 vendedores atuais, com os segmentos novos;
- 3 "vendedores novos" ainda sem nome: Rede de Lojas, Construção e Alimentação;
- "REDE SMART" como segmento de uma pessoa;
- Drogarias, Cosméticos e Perfumaria juntos;
- Autoviação + Entretenimento + Estacionamentos para uma pessoa.

## Decisões do usuário

- **Muda só a equipe da lista.** A outra equipe fica como está e será feita depois.
- Venda interna e representantes não estão na lista e ficaram como estavam.
- **Rede Smart** não é segmento do CRM: é o grupo de clientes "REDE SMART …" (944 lojas no arquivo do CRM, quase todas
  supermercadistas).
  - Quem tem Rede Smart conta os grupos Rede Smart, e quem tem Supermercados conta os supermercados sem a Rede Smart.
  - Na prática não muda número: supermercado não tem lista de potencial, e a lista TOP 50 de supermercado (sem CNPJ)
    vai inteira para quem tem Supermercados. Rede Smart fica "sem lista".
- **Vendedores novos:** caixa com carteira zero ("sem carteira"), o potencial e a linha Ale Dayan do segmento.
- **"No segmento" continua sendo a Aderência do arquivo do CRM** (segmentos antigos do CRM). O usuário não quis
  recalcular com os segmentos novos.
- Cosméticos = Beleza (TOP 50), decisão de 10/11.

## Como foi feito

- `organograma/org_divisao.py` (novo) lê `dados/divisao_<SUP>.json` (local, tem nomes) e muda o `equipe.json`:
  - quem já existe: nome da lista, perfil e segmentos novos, mantendo a carteira;
  - quem é novo: entra com carteira zero;
  - quem não está na lista fica como está;
  - os segmentos ficam sem estrela, porque a lista não marca segmento principal;
  - refaz os "segmentos sem responsável" e a fonte;
  - guarda a versão anterior em `equipe_AAAAMMDD_antes_divisao.json`.
- Nomes: os da lista (nome completo). Os ids continuam os mesmos, e o duplo clique continua achando o vendedor no TOTVS.
- `detalhe_vendedores.json` (local): os 3 novos sem vendedor no TOTVS. Número 0 não é clicável.
- **Aba ORGANOGRAMA regravada inteira pelo `org_layout.py`.** Com 3 caixas novas, não dava para só trocar números (o
  `org_potencial.py` só troca células). O layout é o do `org_layout.py`, com as caixas em 3 por linha.
  - Antes, nenhuma outra aba fazia referência a células da ORGANOGRAMA.
  - Fonte com a data da divisão; "Vendedor novo: ainda sem carteira".
- **ORGANOGRAMA impressao** (`org_impressao.py`): passou a ter 8 folhas.
  - No máximo 8 caixas (2 linhas) por folha. A sobra de vendedores (até 4) divide a folha com os representantes (até 4).
  - A legenda inteira entra onde couber. Nas folhas sem lugar, entra só a linha de cores da Ale Dayan.
  - A tabela da carteira tem no máximo 13 linhas por folha, então a equipe maior ocupa 2 folhas.
  - Folhas: [406, 516, 528, 321, 549, 226, 324, 403] pt (limite 556).
  - Segmentos: "Rede Smart (grupos)" e "Cosméticos e Perfumaria".

## Conferência

- **Macro:** 356 números clicáveis, os 356 iguais ao esperado, com duplo clique, fechar e reabrir.
- **Arquivo:**
  - as outras 54 abas estão iguais célula a célula;
  - 22 dinâmicas e 10 caches iguais;
  - nos nomes definidos, só mudaram as áreas de impressão da ORGANOGRAMA (até a linha 131) e da impressão (até a 168).
- Imagens das folhas 2, 3 e 6 conferidas.
- Backup do arquivo anterior no scratchpad.

## Pendências

- A outra equipe vai receber a sua divisão depois. Usar o mesmo `org_divisao.py`, com um `divisao_<SUP>.json`.
- Quando os vendedores novos tiverem nome e carteira no CRM: trocar o nome no json da divisão, rodar o
  `org_carteira.py` com o arquivo novo e mapear o vendedor no `detalhe_vendedores.json`.
- O "no segmento" ainda é o do CRM (segmentos antigos), até o CRM ser atualizado.

## Lições

- O `org_confere.py` acumulava o contador de diferenças: depois da primeira aba diferente, todas apareciam como "dif".
  Conferir aba por aba (corrigido no script).
- O `ConvertTo-Json` do PowerShell 5.1 pode trocar lista de 1 item por texto. Conferir o json depois (aqui ficou
  certo).
