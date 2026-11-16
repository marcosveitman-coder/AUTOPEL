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

## 3ª parte: exceções de caixas da FARMACIA VF reaplicadas (pedido do usuário)

- O total estava em 1.092.584 desde antes de hoje (fórmula arrastada de novo em algum momento).
- Os 112 valores de lojas sem contagem dupla saíram do `MAPA MERCADO (MARCOS).xlsx` (06/11), que ainda os tinha. O
  `vf_wx_23.json` de 23/10 não existia mais.
- As linhas foram ligadas por posição + rede + nome + CNPJ, e as 112 estão nas mesmas linhas.
- `resumos/vf_excecoes.ps1` agora recebe a coluna no spec (`"col": "V"`, porque I e J saíram).
- **Total: 667.594**, igual ao de 23/10. As duas dinâmicas de resumo da FARMACIA VF foram atualizadas e as larguras
  devolvidas.
- Conferência: só mudaram as 112 fórmulas da coluna V; as outras abas estão iguais. O resumo fecha (409.759 + 61.640 +
  196.195). Macro, 22 dinâmicas e 10 caches presentes.
- Lição: guardar o spec das exceções no repositório local (`resumos/dados`, fora do git). Se a fórmula for arrastada de
  novo, os valores estão no `MAPA MERCADO (MARCOS).xlsx`.

## 4ª parte: potencial em 3 métricas no organograma (pedido do usuário)

- **Pergunta do usuário:** o "Potencial" é PDV ou caixas? Nenhum dos dois: era a contagem de linhas das listas com STATUS
  AUTOPEL diferente de ATIVO.
- **Pedido:** renomear para **POTENCIAL REDES** e pôr embaixo **POTENCIAL PDV** (soma do nº de lojas) e **POTENCIAL Nº
  DE CAIXAS** (soma do "Total de caixas/ mês"), sempre só nas linhas que não são ATIVO, em todos os quadros.
- `org_dados.py` (por segmento: redes, `pdv`, `caixas`; por pessoa: `pot_pdv`, `pot_caixas`):
  - colunas achadas pelo cabeçalho: abas FINAL E/X; FARMACIA VF H/V;
  - nas 112 redes/associações da FARMACIA VF, o PDV usa o nº sem contagem dupla do fim da fórmula de caixas;
  - supervisão e diretoria: cada segmento da equipe uma vez.
- **ORGANOGRAMA impressao** (`org_impressao.py`):
  - caixas: 3 linhas de potencial;
  - faixa do supervisor em 2 linhas (6 métricas da carteira e embaixo os 3 potenciais);
  - tabela da carteira: os 3 potenciais numa coluna, um embaixo do outro (13 colunas não cabem com letra 10);
  - tabela por segmento: saíram Inativando, Lead no CRM e A trabalhar (sempre 0) e entraram PDV e caixas;
  - **paginação pela altura real** (linha de caixas ou da tabela que não cabe vai para uma folha "(continuação)"):
    9 folhas, a maior com 526 pt.
- **Aba ORGANOGRAMA** (`org_layout.py`): 3 linhas nas caixas e 3 colunas nas tabelas (vai até a coluna AA).
- **Duplo clique:**
  - os números de PDV e caixas abrem a mesma lista do potencial de redes;
  - a macro ganhou a coluna CAIXAS/MÊS e a soma de lojas e caixas na nota (`vba_DetalheMod.txt`, macro trocada com
    `org_xlsm.py`);
  - o `org_detalhe_esperado.py` compara PDV e caixas com a quantidade de linhas do potencial de redes.
- **Conferência:**
  - macro: 390 de 390;
  - só mudaram ORGANOGRAMA, ORGANOGRAMA impressao e CONFIG DETALHE;
  - Drogarias: 471.399 caixas = inativo + não é cliente do resumo da FARMACIA VF;
  - imagens das folhas conferidas.

| Segmento | Redes | PDV | Nº de caixas |
|---|---|---|---|
| Alimentação | 315 | 32.355 | 374.717 |
| Drogarias | 90.352 | 106.337 | 471.399 |
| Diretoria (todos) | 91.436 | 198.746 | 1.664.516 |

## 5ª parte: validação com o arquivo do CRM (pedido do usuário)

- **Pedido:** usar o arquivo "MAPA MERCADO V2 - status CRM" (Downloads; uma aba por aba FINAL, com a coluna H STATUS CRM)
  para corrigir qualquer divergência no MAPA V2, "o máximo criterioso possível", perguntando na dúvida.
- **Comparação:** `status_autopel/crm_compara.py` (liga cada linha do CRM à linha da aba FINAL pelo nome; INATIVANDO do
  CRM, 291–365 dias, é ATIVO pela nossa regra). Saída em `%TEMP%\status_1611\crm_div.txt` (tem nomes, fora do git).
- **Antes:** 1.282 de 1.397 linhas iguais. **Depois:** 1.302 iguais; as 95 que sobram foram revistas uma a uma.
- **Erros nossos achados e corrigidos no motor:**
  1. ligação curada da Alessandra (`alessandra_cruzamento.json`, linha N ou N+5): bastava uma palavra em comum, e
     "PARK", "POSTO", "SORVETES" faziam a linha vizinha (5 acima) bater. Estacionamentos e postos herdavam a compra
     da rede de cima. Agora: nome igual vence; senão, mais palavras próprias em comum; N+5 só nas abas FINAL;
  2. **FARMACIA VF lida por índice fixo** depois de apagar I e J: a UF vinha da SITUAÇÃO NA RECEITA ("ATIVA") e a trava
     de UF das farmácias sumia. A `coleta.py` lê a aba pelo cabeçalho;
  3. grupo do cliente igual ao nome (7+ letras) passava por cima de tudo. Agora respeita CNPJ de outra rede, outro
     ramo (também pela 1ª palavra do grupo: "SUPERMERCADO X", "POSTO X"), segmento (ou mesma UF) e UF quando o nome é
     de 2+ redes ou a rede é regional; nome secundário (cidade entre parênteses) só pelo grupo na mesma UF;
  4. "nome curto igual" (EXATO): mesma UF para rede regional (supermercado, farmácia, posto, construção,
     estacionamento) ou nome com palavra de tipo. Marca nacional curta vale em qualquer UF;
  5. razão social igual com raiz diferente da linha = homônimo (regra 7);
  6. código de operador logístico com lojas de várias redes: só as lojas com o nome da rede;
  7. posto também exige a mesma UF nos vínculos fracos; cadastro sem segmento, só na mesma UF;
  8. "nome + palavra curta" no grupo só vale se a palavra for UF ou região.
- **O CRM achou certo e nós não** (4 casos): entraram como vínculo manual (`dados/manuais.json`, local: alias,
  clientes por código/loja).
- **O CRM errou** (a maioria das ~40 linhas em que ele achava compra): nome genérico ou de outra UF, grupo de
  supermercado com nome de loja, produto de papelaria ("LÁPIS 12 CORES") e outra empresa com o mesmo nome.
- **Decisões do usuário (16/11):**
  - marca do mesmo grupo **não herda** a compra da outra marca (`manuais.json` → `separar`): fica só com o que tem o nome
    dela; a linha do grupo inteiro (Alessandra/TOP 50) continua com a compra do grupo;
  - operador de franquia com o nome da marca no cadastro **conta**;
  - revendedor de papel colocado no grupo de uma rede **não conta**.
- **Gravação:** `prepara_gravacao.py` → `grava_status.ps1` (cópia) → `confere_gravacao.py` → `devolve_layout.py` →
  organograma (`org_dados` → `org_layout` → `org_impressao` → `org_detalhe` → `org_detalhe.ps1` → `org_grava` ×2).
  Os 3 scripts de gravação agora servem para regravar: usam as colunas novas das FINAL se já existem e só apagam I:J
  da FARMACIA VF se ainda forem as de Receita/pesquisa.
- **Conferência:** 466.570 células, 0 erradas, nenhuma outra célula diferente; 22 dinâmicas, 10 caches, macro; no
  organograma só mudaram ORGANOGRAMA, ORGANOGRAMA impressao e CONFIG DETALHE; macro 390 de 390; 9 folhas.
- **Mudanças de status nas abas gravadas:** 175 linhas (52 ATIVO→INATIVO, 42 ATIVO→NÃO É CLIENTE, 63 INATIVO→NÃO É
  CLIENTE, 10 NÃO É CLIENTE→INATIVO, 8 para ATIVO).

| Potencial (redes / PDV / caixas) | Antes | Depois |
|---|---|---|
| Alimentação | 315 / 32.355 / 374.717 | 321 / 34.329 / 400.755 |
| Rede de lojas | 269 / 29.020 / 414.099 | 272 / 29.145 / 415.559 |
| Construção | 233 / 5.901 / 65.203 | 236 / 5.940 / 65.476 |
| Drogarias | 90.352 / 106.337 / 471.399 | 90.365 / 106.804 / 475.598 |
| Postos e conveniências | 167 / 23.025 / 324.753 | 175 / 30.026 / 441.984 |
| Estacionamentos | 100 / 2.108 / 14.345 | 105 / 2.309 / 15.752 |
| Pet shop | 100 / 1.685 / 11.844 | 104 / 1.710 / 11.944 |

- **Lições:** depois de apagar coluna, conferir todo script que lê a aba por índice; ligação curada por linha precisa
  de palavra própria do nome, não de palavra de tipo; comparar com uma fonte de fora (CRM) achou erros que a revisão
  interna não achou.

## 6ª parte: revisão final (pedido do usuário, apresentação C-level)

- `revisao_fontes.py`: confere com as fontes brutas (CADASTRO DE CLIENTES.RLT e ultimo fat). Nenhum vendedor diferente
  do cadastro, nenhuma data mais velha que a última nota, nenhum status contra a data.
- Corrigido: a "última compra" da rede ATIVA agora é a compra mais recente (antes era o vínculo mais específico dos 360
  dias); código de cliente reaproveitado por outra empresa sai; loja sem segmento (E001) herda o da loja principal;
  estacionamento exige a mesma UF; ~25 vínculos fracos falsos revisados à mão (`manuais.json`: exclusões, `so_fortes`).
- Decisão do usuário aplicada a todas as redes: marca não herda de outra marca do grupo (cada marca das abas FINAL/IRTT
  é avaliada só com as linhas dela; a linha do grupo fica com o grupo); regra 9 em todos os grupos do IRTT (exceções em
  `grupos_ok`); código que o IRTT repete em várias marcas só vale com o nome da marca.
- Gravado (só as 623 células que mudaram), conferido célula a célula (0 erradas, nada mais mudou), organograma regerado.
- O pendente desta parte foi feito na 7ª parte.

## 7ª parte: fecho da revisão final

- **Macro do duplo clique:** 390 de 390 números iguais ao recálculo, com salvar e reabrir.
- **Dinâmicas:** todas batem com os dados. FARMACIA VF: 825 ATIVO / 1.343 INATIVO / 89.023 NÃO É CLIENTE, caixas
  191.996 + 61.675 + 413.923 = 667.594; portes A e B: 35 / 22 / 19, 172.462 + 44.421 + 4.773 = 221.656; IRTT: 232 ativas
  e 338 a trabalhar (coluna F = coluna E); Alessandra Dayan VF: os 11 segmentos iguais (151 / 124 / 232).
- **Organograma:** potencial (redes, PDV, caixas) dos 7 segmentos igual ao recálculo direto das abas.
- **Varredura:** nenhuma célula de erro (#REF!, #N/D…) nas abas do pedido. Os 3 exemplos da reclamação inicial estão ATIVO
  com data recente. A coluna de data é data de verdade.
- **Achado:** os textos fixos do topo de "FARMACIA VF RESUMO", "… port A e B" e "I.R.T. T. resumo" ainda eram de 23/10
  (contagens antigas, colunas O/P e X1:X4). Foram reescritos com os números de hoje (35 células, `grava_status.ps1` com o
  tipo novo `n` = número), e também uma linha da R_E_S_U_M_O sobre o STATUS AUTOPEL. Conferido célula a célula; nada mais
  mudou. O título da coluna AG da IRTT ("VENDEDOR (CLIENTES 20.10.26)") ficou: uma cache de dinâmica usa o campo, e
  renomear tiraria o campo da tabela. O texto do resumo avisa que os dados são de 16/11.
- **Lição:** resumo com texto fixo não se atualiza com a dinâmica. Depois de mudar status, reescrever os números do texto.

## 8ª parte: revisão de certeza (pedido do usuário: "100% de acuracidade, pode demorar")

Em camadas, da mais objetiva para a de julgamento, sem reaproveitar o processamento onde dava:
1. **`confere_bruto.py`** relê os 10 arquivos de faturamento, o ultimo fat e o cadastro **brutos** e refaz, para as 94.765
   linhas de todas as abas: data mostrada, status x data, compra pela raiz do CNPJ e raiz no cadastro.
   - Achou: **nota com valor zero ou negativo (devolução) contava como compra.** O `extrai_fat.py` agora guarda só nota
     de venda (valor > 0); o `extrai_todos.py` roda os 10 arquivos (no Windows o paralelo precisa de script em arquivo).
     Com isso uma rede do IRTT saiu de ATIVO e algumas dezenas de datas mudaram. O ultimo fat não tem nota com valor <= 0.
   - Depois da correção: nenhuma data, status ou compra pela raiz diferente do bruto.
2. **`revisao_fontes.py`**: vendedor = CADASTRO DE CLIENTES.RLT e data contra o ultimo fat — sem diferença.
3. **`confere_compradores.py`** (de quem compra para o mapa): os 1.500 maiores compradores dos últimos 360 dias; a rede
   que o comprador menciona (CNPJ da linha; nome na razão principal, no grupo ou no começo da fantasia) tem de estar
   ATIVO. 81 redes revistas à mão: coincidência de palavra, ou o comprador é de outra empresa ou de outra UF.
4. **Maiores redes não ATIVAS de cada aba** (30 por aba, com os clientes de mesmo nome): revistas à mão.
5. **Todas as diferenças com o CRM refeitas no estado final** e os vínculos fracos novos revistos à mão.

**Erros achados e corrigidos (meus, das correções da tarde):**
- a trava "posto só na mesma UF" derrubou a **bandeira nacional**: postos com a bandeira Shell (grupo do cadastro
  "POSTOS SHEL", nome cortado no TOTVS) são da rede. Bandeiras nacionais de combustível (`nacionais` no `manuais.json`)
  valem em qualquer UF;
- **marcas do mesmo CNPJ** ficaram juntas e o produto com o nome de uma (Pizza Hut) ia para a outra. A união das marcas
  agora é só pelo nome: a compra pelo CNPJ vale para todas, o produto com o nome de uma é só dela;
- **"DROGAL" colado batia em "DROGA LIDER"** (armadilha já registrada) na nova função de relação: nome colado só com
  2+ palavras;
- data mais nova vinda de vínculo falso (ex.: "ATACADÃO DAS EMBALAGENS"; e-mail com "nosso"): exclusões;
- franqueadora de Hering pelo grupo do franqueado (mesmo endereço da empresa com fantasia HERING): vínculo à mão.
- **Decisão sem mudança:** a data do ultimo fat que não está nas bases (notas de canais que faltam nas bases e notas
  de 2018–2019, antes do início das bases) fica: o CRM, que vem do TOTVS, mostra a mesma data em 29 de 35 casos e nenhum
  status depende disso. A regra 8 era para quando as bases terminavam antes do relatório.

**Resultado:** 7 redes mudaram de status (21 linhas nas abas) e ~80 linhas de data/vendedor. Gravadas só as 131
células que mudaram + 42 do texto dos resumos (`textos_resumos.py`, que recalcula os números pelo próprio arquivo).
- `confere_gravacao.py`: 0 erradas, nada mais mudou; macro 390 de 390; `confere_final.py`: dinâmicas, tabelas fixas,
  textos e organograma batem com as abas; nenhuma célula de erro nas abas do pedido.
- A aba "I.R.T.T.__ (MAIORES 300 GRUPOS)" (fonte original do ranking, fora do pedido) tem 380 "#N/A" nas colunas G/H
  desde antes de hoje. Não mexi.
- **O que é julgamento:** vínculo só por nome, e-mail ou produto (franqueado, distribuidor) foi revisto à mão e a prova
  está na coluna "COMO FOI ACHADO". O resto (CNPJ, código, razão, grupo com o nome) é fato do TOTVS.
