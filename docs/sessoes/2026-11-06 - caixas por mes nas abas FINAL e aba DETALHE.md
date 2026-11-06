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
