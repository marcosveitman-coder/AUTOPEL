# 10/11/2026 – Carteira do CRM por arquivo e quadro da Ale Dayan na ORGANOGRAMA impressao

## Pedidos

1. Pergunta do usuário: o que é preciso para atualizar a ORGANOGRAMA impressao do `MAPA MERCADO V2.xlsm`?
   Respondido:
   - telas ou dados novos do CRM para a carteira;
   - nada para o potencial;
   - o arquivo fechado.
2. Usar o arquivo `carteira-2026-11-10-110700.xlsx` (exportação do CRM) para atualizar a carteira.
3. Embaixo do quadro de cada uma das duas equipes, uma linha **ALE DAYAN**, com os clientes das 11 abas
   "(TOP 50)" divididos em **ativos** e **a trabalhar**:
   - abas: Supermercado, Serviços & Franquias, Rede Lojas, Posto Gasolina, Ótica, Construção, Estacionamento, Drogaria,
     Conveniência, Beleza e Alimentação;
   - a mesma informação também no quadro "CONSTANTE ORGANOGRAMA – POTENCIAL E SEGMENTOS".

## Carteira pelo arquivo do CRM

O arquivo tem uma linha por loja de cliente (93.707), com Vendedor (nome do CRM), Status (Ativo / Perdendo / A
trabalhar) e Aderência (No segmento / Fora / Sem segmento definido).

**Como o CRM conta** (conferido contra as telas de 04/11): um cliente por **vendedor + código de cliente do TOTVS**.
- O código vem do `Ultimo faturamento - SQL.csv`, pelo CNPJ.
- O status do cliente é o melhor entre as lojas; a aderência é a mais frequente.
- Os números batem, com a diferença dos 6 dias. Exemplos: 773 × 782, 454 × 454, 424 × 425; diretoria 39.966 × 40.072.

Regras:
- **supervisão** = soma das pessoas da equipe;
- **diretoria** = a carteira inteira menos a carteira "INATIVOS".

Mudança grande: a **venda interna** caiu de 78 para 18 clientes. É a carteira do CRM de hoje.

Script: `organograma/org_carteira.py`.
- Atualiza o `dados/equipe.json` (local) e guarda a versão anterior em `equipe_20261110.json`.
- A data da carteira vai para a fonte da folha 7.

**Duplo clique na carteira:** o nome do CRM é o **nome reduzido** do vendedor no TOTVS (coluna "Nome Reduzid" do
Ultimo faturamento). Com isso, a venda interna passou a apontar para os 2 vendedores de nome reduzido igual, e um
representante que não tinha vendedor foi achado. Mapa em `dados/detalhe_vendedores.json` (local).

## Ale Dayan (listas TOP 50)

- As abas "(TOP 50)" são detalhes de tabela dinâmica feitos pelo usuário: cabeçalho na linha 3, com EMPRESA, UF, LOJAS,
  BASE, STATUS RECEITA, PORTE, STATUS AUTOPEL e SEGMENTO.
- Só têm as empresas com situação ATIVA na Receita, por isso de 41 a 50 por aba.
- **Ativo** = STATUS AUTOPEL ATIVO. **A trabalhar** = inativo + não é cliente, como no potencial.

| Segmento | Clientes | Ativos | A trabalhar |
|---|---|---|---|
| Supermercado | 48 | 34 | 14 |
| Serviços e franquias | 41 | 2 | 39 |
| Rede de lojas | 50 | 17 | 33 |
| Posto de gasolina | 48 | 11 | 37 |
| Ótica | 45 | 1 | 44 |
| Construção | 41 | 4 | 37 |
| Estacionamento | 43 | 16 | 27 |
| Drogaria | 48 | 24 | 24 |
| Conveniência | 46 | 9 | 37 |
| Beleza | 47 | 4 | 43 |
| Alimentação | 50 | 7 | 43 |

**Desenho:**
- quadro "ALE DAYAN – CLIENTES DAS LISTAS TOP 50 POR SEGMENTO" em duas metades lado a lado (6 + 5 segmentos), com as
  colunas Segmento, Clientes, Ativos e A trabalhar;
- aparece embaixo das caixas da 1ª equipe (folha 1) e da 2ª equipe (folha 3, depois dos representantes) e na
  folha 6, depois da tabela de potencial;
- os mesmos números nos 3 lugares (a lista é uma só);
- 7 folhas, a maior com 541 pt (limite 556).

**Duplo clique:** os 99 números novos abrem a lista da aba TOP 50. A macro passou a aceitar EMPRESA e LOJAS (sem NOME
FANTASIA e CNPJ). O `org_xlsm.py` agora troca só a macro num .xlsm que já existe (`vba_projeto.troca_vba`).

## Conferência

- **Teste da macro em Excel oculto:**
  - 420 números em 203 s; os 420 iguais ao recálculo independente;
  - fechar apaga a DETALHE;
  - salvar com a lista aberta e reabrir apaga a DETALHE sem marcar alteração;
  - a macro continua no arquivo.
- **Arquivo:**
  - só a ORGANOGRAMA impressao e a CONFIG DETALHE (oculta) mudaram;
  - as outras abas estão iguais célula a célula;
  - 22 dinâmicas e 10 caches iguais;
  - nos nomes definidos, só mudou a área de impressão (até a linha 171).
- Imagens das folhas 1, 3 e 6 conferidas.

## 2ª parte: carteira na aba ORGANOGRAMA

O usuário pediu para atualizar também a aba ORGANOGRAMA (a que não é de impressão).
- O `org_potencial.py` passou a trocar também os números de carteira, com os mesmos textos e contas do `org_layout.py`:
  - as 2 linhas de carteira de cada caixa;
  - o RESUMO DAS EQUIPES;
  - a tabela "1. CARTEIRA DE CADA PESSOA" (H–X);
  - a nota abaixo dela, que agora cita só a diretoria com clientes sem segmento;
  - a fonte.
- Foram 292 células, gravadas pelo `org_potencial.ps1` numa cópia.
- **Conferência:**
  - rodar de novo dá 0 trocas;
  - todas as outras abas e as células não trocadas estão iguais;
  - as 22 dinâmicas são as mesmas: o Excel só renumerou arquivos e 2 estilos (dxfId);
  - os nomes definidos estão iguais;
  - a macro foi testada em 5 números, com duplo clique, fechar e reabrir.
- A venda interna agora mostra "no segmento 5 (27,8%)", porque o arquivo do CRM dá aderência a 5 clientes dela. A caixa
  continua dizendo "Sem segmento no CRM", porque a pessoa não tem segmento cadastrado.

## 3ª parte: linha Ale Dayan dentro das caixas (substitui os quadros)

O usuário não gostou dos quadros e mandou uma foto do formato que queria: **uma linha "Ale Dayan" no fim de cada caixa**
de vendedor(a), venda interna e representante.
- Os quadros da 1ª parte saíram (embaixo das equipes e na folha 6).
- A linha mostra **3 números** (decisão do usuário): ativos (verde), inativos (amarelo) e não é cliente (cinza). São
  as listas TOP 50 dos segmentos da pessoa, somadas.
- **Segmento do CRM → aba TOP 50:**

  | Segmento do CRM | Aba TOP 50 |
  |---|---|
  | Supermercadista | Supermercado |
  | Alimentação | Alimentação |
  | Drogarias | Drogaria |
  | Rede de Lojas | Rede Lojas |
  | Postos e Conveniências | Posto Gasolina + Conveniência |
  | Construção | Construção |
  | Estacionamentos | Estacionamento |
  | Cosméticos | Beleza (decisão do usuário) |

  Ótica e Serviços & Franquias não têm responsável. Sem lista = "sem lista".
- A folha 2 (1ª parte da equipe maior) não cabia: 601 pt. Agora a legenda (com as cores da Ale Dayan) fica só na
  última folha de cada equipe, e as folhas têm [419, 502, 406, 321, 536, 324, 387] pt.
- **Conferência:**
  - 354 números clicáveis; a macro deu os 354 iguais ao esperado, com duplo clique, fechar e reabrir;
  - só a ORGANOGRAMA impressao e a CONFIG DETALHE mudaram;
  - dinâmicas e caches iguais; nos nomes definidos, só a área de impressão (até a linha 145).
- O `org_dados.py` ainda lê as abas TOP 50 (`ale` no json). A função dos quadros foi apagada do `org_impressao.py`.

## Pendências

- O duplo clique nos números de carteira lista os clientes pelo TOTVS do dia. O número do organograma é do arquivo do CRM
  (10/11), e a DETALHE mostra os dois.

## Lições

- O arquivo de carteira do CRM é melhor que as telas: dá todos os números de uma vez, e o nome do vendedor é o nome
  reduzido do TOTVS.
- Abas criadas pelo usuário mudam a ordem e os nomes. Achar as abas TOP 50 pelo nome com os espaços normalizados (há abas
  com 2 espaços antes de "(TOP 50)").
