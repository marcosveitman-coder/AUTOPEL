# 16/11/2026 – Refazer o STATUS AUTOPEL de todas as abas do MAPA MERCADO V2

## Contexto

O usuário abriu o MAPA MERCADO V2 diante da diretoria e encontrou status errados. Exemplos: três redes de alimentação
marcadas como INATIVO ou NÃO É CLIENTE que compram toda semana, por franqueado, por distribuidor e por produto
personalizado com o nome da rede.

**Causa (erro meu, 05/11):** regras rígidas demais. Compra de franqueado e de distribuidor não contava, e uma regra
contra quem revende produto personalizado derrubava vínculos verdadeiros. Também não houve revisão rede a rede antes de
entregar.

## Pedido

- Refazer status, data da última venda e último vendedor (do CADASTRO DE CLIENTES.RLT) com os arquivos de
  `Downloads\12.11.2026`:
  - faturamento 2020–2026, até 31/07/2026;
  - ultimo fat, até 12/11/2026;
  - cadastro de clientes e de produtos.
- Esgotar os caminhos: e-mail, descrição do produto e grupo de clientes.
- Abas: Alessandra Dayan, as 11 "(TOP 50)", Alessandra Dayan VF, IRTT por bandeira, IRTT resumo, FARMACIA VF, FARMACIA VF
  RESUMO, as 6 abas FINAL e ORGANOGRAMA impressao.
- FARMACIA VF: deixar só uma das colunas H/I/J de nº de lojas.
- Regra combinada: a compra de qualquer cliente ligado à rede conta (franqueado, associada, distribuidor que atende a rede,
  cliente no grupo com o nome da rede, comprador de produto personalizado).

## Feito (scripts em `avulsos/status_autopel/`)

- `extrai_fat.py`: lê cada base pelo cabeçalho, em paralelo (9 arquivos, ~9 min). Guarda a última nota por
  cliente+loja+produto. Os pickles ficam em `%TEMP%\status_1611\pk`.
- `base_clientes.py`: junta cadastro, ultimo fat e bases. São 93.845 clientes+loja; 37.920 têm compra; a mais recente é
  de 12/11/2026.
- `coleta.py`:
  - lê as linhas de todas as abas; OUTROS FINAL e SUPERMERCADOS (reuniao) entram só como apoio de CNPJ;
  - forma clusters (a mesma rede nas várias abas): mesmo segmento, sem CNPJs diferentes, nome igual ou começando igual,
    e o cruzamento curado da Alessandra de 05/11 (linhas das abas FINAL +5).
- `motor.py`: os caminhos e as travas. As regras estão no topo do arquivo e na aba LEIA da conferência.
- `avalia.py`: roda por cluster. Inclui associadas da FARMACIA VF (coluna B, com várias redes separadas por " / " e a
  rede-mãe entre parênteses), o mapa curado `farmacias/final/grupos_redes.json` e os códigos e grupos da aba IRTT (já
  validados em 23/10).
- `revisa.py`: lista para revisão manual (status que depende só de vínculo fraco; perdas de ATIVO).
- `conferencia.py`: planilha `COMERCIAL (NOVO)\conferencia STATUS AUTOPEL 16-11-2026.xlsx`, com as abas LEIA, RESUMO,
  REDES (3.958 linhas, com a prova), FARMÁCIAS AVULSAS (962 que mudaram) e FARMACIA VF LOJAS.
- Vínculos manuais e exclusões (têm código de cliente): `status_autopel/dados/manuais.json`, local, fora do git.

## Revisão manual (lições)

Falsos vínculos achados e as regras que os barram:
- nome igual de outro ramo ("supermercado X" para a loja X; posto com o nome de uma marca de chocolate). Correção:
  palavra de tipo de outro ramo na fantasia (qualquer posição) ou na 1ª palavra da razão;
- sobrenome ou palavra comum (lista GENERICO; nome curto só vale igual);
- grupo administrador ou nome entre parênteses (nome secundário) casando com o começo de outra razão social;
- provedor de e-mail (terra, uol…) e "CLIENTES DIVERSOS" como grupo;
- nome regional repetido em outra UF: para supermercado e farmácia, vínculo fraco só na mesma UF (regra 2 do CLAUDE.md);
- a FARMACIA VF tem quase todas as farmácias do país como linhas. Uma trava "CNPJ de outra linha" barra associadas e
  franqueadas: a trava vale só para CNPJ de outras REDES;
- `\b` dentro de heredoc do bash virou caractere de controle no .py: escrever os patches com a ferramenta de arquivo.

## 2ª parte: gravação (autorizada pelo usuário)

O usuário pôs na pasta o `faturamento 2026 (ago_ontem).xlsx`, de 03/08 a 15/11/2026. Ele tem uma linha de título antes
do cabeçalho, e o `extrai_fat.py` agora acha o cabeçalho pela palavra EMISSAO. Com esse arquivo só mudaram 7 linhas,
porque o ultimo fat já trazia a nota mais recente; o ganho foi no caminho por produto. O usuário autorizou corrigir,
conferir e gravar, com as propostas da aba LEIA.

**Gravado no `MAPA MERCADO V2.xlsm`** (`prepara_gravacao.py` → `grava_status.ps1`, numa cópia):
- **6 abas FINAL:** STATUS AUTOPEL (W) e 3 colunas novas: Y "ÚLTIMA COMPRA AUTOPEL", Z "VENDEDOR (CADASTRO)" e
  AA "COMO FOI ACHADO", com o formato da W;
- **Alessandra Dayan** (11 tabelas), **Alessandra Dayan VF** e as **11 (TOP 50)**: STATUS AUTOPEL;
- **IRTT por bandeira:** Atende?, Status, Última compra, Como casou, VENDEDOR CARTEIRA e VENDEDOR (CLIENTES). F e H são
  fórmulas e se recalcularam;
- **FARMACIA VF:**
  - data, status (agora ATIVO / INATIVO / NÃO É CLIENTE), como foi achado, cliente e vendedor;
  - **colunas I e J apagadas**: fica só a H "Nº DE LOJAS (CONSIDERADO)";
  - o status foi da P para a **N**, e a fórmula de caixas passou a `...U...$V$1...*H`;
- **dinâmicas atualizadas:** só as 5 abas que usam as abas gravadas. As larguras foram devolvidas pelo `devolve_layout.py`;
- **ORGANOGRAMA e ORGANOGRAMA impressao:** regeradas. O `org_dados.py` agora acha a coluna STATUS AUTOPEL pelo
  cabeçalho.

**Conferência:**
- `confere_gravacao.py`: 466.570 células gravadas, **0 erradas**; as outras abas estão iguais célula a célula;
- FARMACIA VF: o total de caixas é o mesmo antes e depois (1.092.584);
- as dinâmicas batem com os dados. A da Alessandra Dayan VF tem filtro STATUS RECEITA = ATIVA, e com ele os 11 segmentos
  batem;
- as TOP 50 estão iguais à VF;
- 22 dinâmicas e 10 caches; nos nomes definidos, só o filtro da FARMACIA VF (A6:X6 → A6:V6);
- macro: 357 de 357 números iguais ao esperado, com duplo clique, fechar e reabrir.

**Resultado** (ATIVO / INATIVO / NÃO É CLIENTE):

| Aba | Ativo | Inativo | Não é cliente |
|---|---|---|---|
| Alimentação | 53 | 54 | 261 |
| Rede de lojas | 57 | 102 | 167 |
| Construção | 17 | 36 | 197 |
| Pet | 10 | 5 | 95 |
| Estacionamento | 33 | 38 | 62 |
| Postos | 43 | 50 | 117 |
| Alessandra | 174 | 131 | 245 |
| IRTT | 268 | 210 | 92 |
| Redes da FARMACIA VF | 147 | 94 | 143 |

839 linhas mudaram.

**Armadilhas desta parte:**
- a aba IRTT estava com **filtro ligado**: gravar a coluna em bloco pulou as linhas ocultas e deslocou valores (achado na
  conferência). Nas abas com `FilterMode`, gravar célula por célula;
- o código de vendedor "000000" virou número 0: vendedor só número → vendedor da nota;
- o `--exceto` do `devolve_layout.py` compara o nome como está no XML: `&quot;FARMACIA VF&quot;`;
- o código de cliente do IRTT cobre todas as lojas do código (um CD que atende duas redes). Para rede ATIVA, a prova
  mostrada é a mais específica entre as compras dos últimos 360 dias (`ORDEM_PROVA`).

## Observação para o usuário

O total de caixas da FARMACIA VF já estava em 1.092.584 antes de hoje. Pelo CLAUDE.md, com as exceções das redes ele
seria 667.594: alguém arrastou a fórmula de novo. Não foi mexido; perguntar se quer reaplicar as exceções.
