# 05–06/11/2026 – Caixas por mês nas abas FINAL e aba DETALHE (números clicáveis da ORGANOGRAMA impressao)

## Pedido

1. Levar o racional de potencial em caixas da aba "FARMACIA VF" para as 7 abas FINAL (ALIMENTACAO, REDE DE LOJAS,
   CONSTRUÇÃO, POSTOS GAS & CONVENIENCIA, PET SHOP, ESTACIONAMENTO, OUTROS). Na FARMACIA VF, o racional é:
   - a tabela a/b/c/d em W1:X4 (17, 12, 7, 4 caixas/mês por loja), feita para simular na hora;
   - a fórmula da coluna X: valor do porte × nº de lojas.

   Instruções do usuário:
   - coluna "Total de caixas/ mês" na última coluna de cada aba;
   - pode inserir de 5 a 7 linhas no topo para a tabela de portes;
   - valores iguais aos da FARMACIA VF.
2. Na "ORGANOGRAMA impressao", todos os números clicáveis, "como numa tabela dinâmica": com 2 cliques, abrir uma aba
   DETALHE com RAZÃO SOCIAL, NOME FANTASIA, Nº DE LOJAS e CNPJ.

**Decisões do usuário** (perguntadas antes de começar):
- **Link sem macro**: o arquivo continua .xlsx. Duplo clique precisaria de macro e de um arquivo .xlsm;
- **Só os números de potencial e das listas**: os de carteira vieram das telas do CRM, que não têm a lista de clientes.

## Entregas (`COMERCIAL (NOVO)\MAPA MERCADO.xlsx`)

**Abas FINAL** (as 7), por `mapa_final/caixas_final.ps1`, pelo Excel numa cópia:
- 5 linhas inseridas no topo: **o cabeçalho passou da linha 1 para a 6**, como na FARMACIA VF. O filtro foi refeito
  (A6 até a coluna nova) e o painel ficou congelado abaixo da linha 6;
- linhas 1–4: a tabela a/b/c/d copiada da FARMACIA VF W1:X4 (valores e formato). A letra fica na última coluna antiga
  (W; na OUTROS, Y) e o valor na coluna nova. Os valores ficam em cada aba, para simular por segmento;
- coluna nova "Total de caixas/ mês": X (na OUTROS, Z). A fórmula é
  `=SE(ESQUERDA($V7;1)="A";$X$1;SE(…="B";$X$2;SE(…="C";$X$3;SE(…="D";$X$4;0))))*N($E7)`:
  - o PORTE das abas FINAL é texto ("A – Grande…"), por isso a fórmula usa a 1ª letra;
  - "SEM Nº DE LOJAS" e linha sem lojas dão 0.

| Aba | Caixas/mês |
|---|---|
| ALIMENTACAO FINAL | 674.706 |
| REDE DE LOJAS FINAL | 643.120 |
| CONSTRUÇÃO FINAL | 74.920 |
| POSTOS GAS & CONVENIENCIA FINAL | 612.782 |
| PET SHOP  FINAL | 22.286 |
| ESTACIONAMENTO FINAL | 41.601 |
| OUTROS FINAL | 689.089 |

As redes pet também estão na REDE DE LOJAS: ao somar as abas, contar uma vez.

**Aba DETALHE** (nova, no fim da pasta; `organograma/org_detalhe.py` + `org_detalhe.ps1`), 92.682 linhas:
1. "Potencial por pessoa e equipe":
   - um bloco por pessoa, com o total e os segmentos que formam o número;
   - cada segmento tem o link "ver lista →".
2. Um bloco por segmento:
   - primeiro os ATIVOS, depois "POTENCIAL (tudo que não é ATIVO)", separado por status;
   - colunas: RAZÃO SOCIAL, NOME FANTASIA, Nº DE LOJAS e CNPJ (pedidas), mais STATUS AUTOPEL e ORIGEM (aba e linha);
   - cada empresa aparece uma vez;
   - o bloco de Drogarias tem as 91.191 linhas da FARMACIA VF;
   - todo título de bloco tem o link "← voltar ao organograma".

**ORGANOGRAMA impressao**: 56 números viraram link para a DETALHE:
- o potencial das caixas, das faixas de supervisor e da carteira, que leva ao bloco da pessoa;
- na tabela por segmento, total, ativo, inativo, não é cliente, a trabalhar e potencial, que levam à lista certa.

Os números com link não mudaram de aparência (Verdana, mesma cor, sem sublinhado), para a impressão continuar igual. O
clique é **simples**: o Excel não deixa link de célula com duplo clique sem macro.

## Conferência

- **Abas FINAL:**
  - toda célula antiga está 5 linhas abaixo;
  - as linhas 1–5 só têm a tabela;
  - o cabeçalho está certo;
  - toda linha bate com porte × lojas recalculado à parte.
- **Outras abas** iguais célula a célula. As 21 dinâmicas continuam no arquivo. Nos nomes definidos, só mudaram os
  filtros das abas FINAL (`_FilterDatabase`, agora a partir da linha 6).
- **DETALHE:** conferida contra o arquivo gerado; as contagens por status batem com o `org_dados.json`. Diferem só 20
  nomes da FARMACIA VF que começam com apóstrofo: o Excel esconde o 1º apóstrofo.
- **ORGANOGRAMA impressao:**
  - mesmos valores de antes;
  - 7 páginas, nenhuma quebra vertical;
  - os 56 links em Verdana 10, sem sublinhado, cada um conferido contra a linha de destino na DETALHE;
  - imagens conferidas.

## Scripts

- `mapa_final/caixas_final.ps1`: só uma vez. Para se o cabeçalho já estiver na linha 6.
- `mapa_final/status_coleta.py`, `status_grava.py` e `status_alessandra.py`, e `organograma/org_dados.py`: acham o
  cabeçalho das abas FINAL sozinhos (`linha_cab`), na linha 1 ou na 6.
- `organograma/org_impressao.py`: marca os números com `link` (P|pessoa, S|segmento|status).
- `organograma/org_detalhe.py`:
  - monta a DETALHE (`detalhe.tsv` + `detalhe.json`);
  - grava no `org_impressao.json` a linha de destino (`alvo`) de cada link.
- `organograma/org_detalhe.ps1`: grava a DETALHE pelo Excel, em blocos de 10 mil linhas, em cerca de 1 min.
- `organograma/org_grava.ps1`:
  - grava o hiperlink dos números com `alvo` e refaz a aparência;
  - apaga os links antigos antes de redesenhar.

Ordem para refazer (na pasta `organograma/dados`):
1. `org_dados.py`;
2. `org_impressao.py`;
3. `org_detalhe.py`;
4. `org_detalhe.ps1`;
5. `org_grava.ps1` com o `org_impressao.json`.

**A DETALHE e a impressão andam juntas:** os links apontam para números de linha da DETALHE.

## Lições

- Script .ps1 com acento ("Nº", "CONSTRUÇÃO") precisa de UTF-8 **com BOM**: sem ele, o PowerShell 5.1 lê errado e o
  nome da aba ou do cabeçalho não é achado.
- O `ConvertFrom-Json` do PowerShell 5.1 não lê JSON grande (90 mil linhas): passar as linhas por TSV.
- No Git Bash, `\t` e `\n` dentro de um heredoc com Python viraram caracteres de verdade no arquivo gerado. Corrigir com
  `chr(92)` ou pelas ferramentas de edição.
- `Hyperlinks.Add` aplica o estilo "Hyperlink" (fonte, cor, sublinhado). Gravar o link antes do formato e refazer a
  fonte depois.

## 2ª parte: duplo clique de verdade (macro) e arquivo leve

**Pedido:** a DETALHE com links não era o que o usuário queria, e o arquivo ficou mais pesado (46 → 50 MB). Ele quer:
- duplo clique em **todos** os números, inclusive os da carteira (ex.: clientes 782, ativos 13,7% e a trabalhar 86,0%
  de uma vendedora);
- uma aba nova, DETALHE, aberta automaticamente com a lista (razão social, nome fantasia, nº de lojas, CNPJ);
- arquivo sem ficar muito mais pesado.

**Decisões técnicas:**
- **Só macro abre aba nova com duplo clique.** O arquivo passa a ser `COMERCIAL (NOVO)\MAPA MERCADO.xlsm`.
- **Nenhuma lista fica guardada.** A macro monta a lista na hora:
  - carteira: CSV do dia do TOTVS;
  - potencial e tabela por segmento: as próprias abas.
  - Saíram a DETALHE fixa de 92 mil linhas e os 91 links. O .xlsm salvo pelo Excel tem 46,3 MB, praticamente o
    tamanho do .xlsx antes da DETALHE fixa (46,2 MB); com ela eram 50,0 MB.
- **Carteira do CRM reproduzida pelo TOTVS:**
  - um cliente por vendedor + código, com a nota mais recente;
  - ativo ≤ 290 dias, perdendo ≤ 365;
  - no segmento = segmento do cliente entre os do vendedor.

  | Pessoa | CRM | TOTVS | Ativos CRM × TOTVS |
  |---|---|---|---|
  | Vendedora A | 782 | 779 | 149 × 150 |
  | Vendedora B | 3.185 | 3.185 | 302 × 302 |
  | Vendedora C | 616 | 616 | 186 × 186 |

  - supervisão = soma da equipe (1.818 × 1.804; 8.889 × 8.707);
  - diretoria = diretores no TOTVS (40.444 × 40.072). A divisão no/fora do segmento da diretoria não dá para refazer,
    e a lista avisa;
  - venda interna = carteira "VENDA INDIRETA INTERNA".
- **O Excel daqui não deixa gravar macro por automação** ("Confiar no acesso ao modelo de objeto do projeto do VBA"
  desligada; opção de segurança não se muda). O `vbaProject.bin` foi montado sem o Excel, por `avulsos/vba_projeto.py`
  (especificações MS-OVBA e MS-CFB, com autoteste de ida e volta).
- **Primeiro teste numa pasta pequena:** o Excel abriu, compilou, rodou função, disparou o evento de salvar e manteve a
  macro depois de salvar e reabrir.

**Problemas achados no teste e corrigidos:**
- **`ADODB.Stream.ReadText` levava mais de 4 minutos** para os 84 MB do CSV. Troquei por leitura binária +
  `MultiByteToWideChar` (UTF-8): 1,2 s.
- **O Excel ignora, durante o salvamento (`BeforeSave`), a exclusão de aba e a limpeza de células**, sem dar erro. Além
  disso, o salvamento automático do OneDrive dispara o `BeforeSave` a cada alteração, e mexer na lista ali a apagaria
  enquanto o usuário lê. Solução:
  - a DETALHE é apagada no `Workbook_Open` e no `Workbook_BeforeClose`;
  - montar ou apagar a lista devolve o estado "salvo" do arquivo, então quem só consultou fecha sem a pergunta de salvar;
  - se alguém salvar com a lista aberta, ela some na próxima abertura.
- O formato de número ia até uma linha depois da última da lista. Corrigido.

**Teste completo** (Excel oculto próprio, com tempo limite):
- a macro montou a lista dos 311 números em 134 s (0,4 s em média; a maior, 3,8 s; a 1ª leitura do CSV leva cerca de
  1 s);
- as 311 quantidades bateram com o recálculo independente em Python (`org_detalhe_esperado.py`);
- a simulação do duplo clique abriu a DETALHE;
- fechar apaga a DETALHE, salvar com a lista aberta e reabrir apaga a DETALHE, e a macro continua no arquivo.

**Conferência do arquivo:**
- todas as abas com os mesmos valores do .xlsx de partida;
- saíram a DETALHE fixa e os 91 links que eu tinha criado; os outros links do arquivo ficaram;
- entrou a CONFIG DETALHE (oculta);
- 21 dinâmicas, 9 caches e nomes definidos iguais.

**Scripts:**
- `avulsos/vba_projeto.py`;
- `organograma/vba_ThisWorkbook.txt` e `vba_DetalheMod.txt` (código da macro);
- `org_detalhe.py`, `org_detalhe.ps1`, `org_xlsm.py`, `org_detalhe_teste.ps1` e `org_detalhe_esperado.py`.

**Lições:**
- Antes de propor uma alternativa sem macro, confirmar que ela atende ao pedido: "como tabela dinâmica, com 2 cliques,
  abrindo aba nova" só se faz com macro.
- Teste de macro por automação precisa de tempo limite e de encerrar só o Excel do teste: um erro de compilação abre
  janela modal e trava a chamada.
- O CRM da Autopel é o TOTVS: a carteira se refaz por vendedor + código de cliente.

**Entrega:** `COMERCIAL (NOVO)\MAPA MERCADO.xlsm`, salvo uma vez pelo próprio Excel (a macro fica compilada).

Conferência final contra o .xlsx de partida:
- todas as abas com os mesmos valores;
- só saíram a DETALHE fixa e os links; só entrou a CONFIG DETALHE;
- 21 dinâmicas, 9 caches e nomes definidos iguais;
- testes depois de salvar: clientes de uma vendedora 778 em 1 s; potencial 461; total de farmácias 91.191 em 4 s.

**Nomes finais** (decisão do usuário, que não quis apagar o antigo; só renomear):
- `MAPA MERCADO.xlsx` → **`MAPA MERCADO (MARCOS).xlsx`**: a versão anterior, com a DETALHE fixa e os links;
- `MAPA MERCADO.xlsm` → **`MAPA MERCADO V2.xlsm`**: a versão de trabalho, com a macro.

Conteúdo idêntico antes e depois da troca de nome (SHA-256). A macro não depende do nome do arquivo.

**Como usar** (explicado ao usuário):
- ao abrir o .xlsm, clicar em "Habilitar Conteúdo";
- dar duplo clique num número da ORGANOGRAMA impressao;
- a aba DETALHE é apagada ao fechar o arquivo.
