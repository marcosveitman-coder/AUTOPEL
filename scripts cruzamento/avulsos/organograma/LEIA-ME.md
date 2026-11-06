# Organograma da equipe comercial (aba ORGANOGRAMA do MAPA MERCADO)

Monta a aba **ORGANOGRAMA**, o "constante organograma":
- quem está embaixo de quem (diretoria → supervisores → vendedores, venda interna e representantes);
- os segmentos de cada pessoa (★ = segmento principal no CRM);
- a carteira de cada pessoa no CRM, dentro e fora do segmento: ativos (até 290 dias), perdendo (291 a 365) e a
  trabalhar (mais de 365 dias ou sem compra), em número e em %;
- o potencial de cada segmento: linhas das listas de mercado com STATUS AUTOPEL diferente de ATIVO.

## Dados

- `dados/equipe.json` (local, fora do git, porque tem nomes): pessoas, perfil, supervisor, segmentos e os números
  das telas de carteira do CRM. Atualizar à mão quando chegarem telas novas. Outra pasta:
  `AUTOPEL_ORGANOGRAMA_EQUIPE`.
- **Potencial**, lido do próprio MAPA MERCADO:
  - coluna W (STATUS AUTOPEL) das abas ALIMENTACAO FINAL, REDE DE LOJAS FINAL e CONSTRUÇÃO FINAL;
  - coluna P da "FARMACIA VF", cabeçalho na linha 6;
  - desde 05/11/2026, também a coluna W da POSTOS GAS & CONVENIENCIA FINAL, da ESTACIONAMENTO FINAL e da PET SHOP  FINAL.
- **Segmento do CRM → aba:**
  - ALIMENTACAO → ALIMENTACAO FINAL;
  - REDE DE LOJAS → REDE DE LOJAS FINAL;
  - CONSTRUCAO → CONSTRUÇÃO FINAL;
  - DROGARIAS → FARMACIA VF.

  - POSTOS E CONVENIENCIAS → POSTOS GAS & CONVENIENCIA FINAL;
  - ESTACIONAMENTOS → ESTACIONAMENTO FINAL;
  - PET SHOP → PET SHOP  FINAL (segmento sem responsável no CRM).

  Os outros segmentos ainda não têm lista de mercado e aparecem sem potencial. A OUTROS FINAL mistura beleza, óticas e
  serviços e não corresponde a um segmento do CRM.

**A aba ORGANOGRAMA do arquivo tem um layout anterior ao do `org_layout.py`** (resumo das equipes e legenda depois das
caixas, não ao lado). Não regravar a aba com o `org_layout.py` sem perguntar. Para atualizar os números de potencial:
1. `python ../org_dados.py "<MAPA>"` (na pasta `dados`);
2. `python ../org_potencial.py "<MAPA>"`:
   - lê a aba como está e gera `org_potencial.json`;
   - a caixa é ligada à pessoa pelo nome acima dela, e a linha da carteira pelo nome na coluna B;
   - segmento novo vira linha nova da tabela "2. POTENCIAL POR SEGMENTO";
3. `powershell -File ../org_potencial.ps1 -Arquivo "<cópia>"`: troca as células e insere as linhas novas com o formato
   da última linha da tabela;
4. conferir com `org_confere.py` (outras abas) e célula a célula na própria aba (linhas abaixo da tabela descem).

## Ordem

1. `python org_dados.py "<MAPA.xlsx>"` (na pasta `dados`): carteira + potencial → `org_dados.json`.
2. `python org_layout.py`: desenho da aba → `org_layout.json`.
   - **Caixas:** 24 colunas iguais (B..Y), metade esquerda para a equipe menor. Caixas de 4 colunas, 3 por linha,
     centralizadas.
   - **Abaixo das caixas:** resumo das equipes, carteira de cada pessoa, potencial por segmento, e segmentos ×
     responsáveis.
3. `powershell -File org_grava.ps1 -Arquivo "<cópia.xlsx>" -Layout org_layout.json`: grava pelo Excel numa cópia.
4. Conferir a cópia:
   - as outras abas célula a célula;
   - as dinâmicas e os caches;
   - nos nomes definidos, só pode entrar a área de impressão da ORGANOGRAMA.

   Depois trocar o original.

## Aba "ORGANOGRAMA impressao" (versão para imprimir)

As mesmas informações, em **7 folhas A4 paisagem impressas a 100%** (6 até 04/11/2026). A letra é Verdana, sempre com tamanho 10 ou
maior e sem itálico (pedido do usuário):
1. equipe menor;
2. equipe maior, parte 1: vendedores e venda interna;
3. equipe maior, parte 2: representantes;
4. carteira de cada pessoa: diretoria e equipe menor;
5. carteira de cada pessoa: equipe maior;
6. potencial por segmento;
7. segmentos × responsáveis (separada da folha 6 em 05/11, quando o potencial ganhou 3 segmentos).

Em cada folha de equipe:
- no topo, uma faixa do supervisor: pessoas, clientes, no segmento, ativos, perdendo, a trabalhar e potencial;
- abaixo, as caixas: 14 colunas cada, 4 por linha, com rótulo e valor em linhas separadas;
- no rodapé, a legenda.

Com Verdana 10, a equipe maior (12 pessoas) não cabe em uma folha: por isso ela ocupa 2.

**Grade:** 62 colunas estreitas (B..BK, largura 1,72 = 17 px).
**Alturas:** texto quebrado em célula mesclada não se ajusta sozinho. O script estima as linhas pela largura das letras
da Verdana, fixa a altura de cada linha e para com erro se alguma folha passar de 556 pt (altura útil da página).

1. `python org_impressao.py` (depois do `org_dados.py`) → `org_impressao.json`.
2. `powershell -File org_grava.ps1 -Arquivo "<cópia.xlsx>" -Layout org_impressao.json`. O gravador lê do json a aba,
   a letra, as colunas, as quebras e o zoom. A saída mostra as quebras gravadas, que devem ser só as do json.
3. `powershell -File org_png.ps1 -Arquivo "<cópia.xlsx>" -Layout org_impressao.json -Saida <pasta>`: uma imagem por
   folha. Conferir se há texto cortado.
4. `python org_confere.py "<partida.xlsx>" "<cópia.xlsx>" "ORGANOGRAMA impressao"`: as outras abas célula a célula,
   as dinâmicas, os caches e os nomes definidos.

## Cuidados

- **O potencial é do segmento:** quem divide o segmento divide as mesmas oportunidades; não somar entre pessoas. No
  supervisor e na diretoria, cada segmento da equipe conta uma vez só.
- **CLIENTES** = ativos + perdendo + a trabalhar (cartões do CRM). Onde há clientes sem segmento, NO + FORA fica menor
  que CLIENTES.
- **Carteira dos supervisores:** é a visão do CRM de cada um e não bate com a soma das pessoas da equipe.
- **Impressão legível:** Verdana 10 é bem mais larga que Calibri. Rótulo e valor em células separadas, número
  gravado como número (senão aparece o triângulo verde) e nada de zoom abaixo de 100%.
- **Faixa compartilhada:** caixas e quadros na mesma faixa de linhas dividem a altura. O quadro ao lado das caixas
  fica sem altura fixa, senão estica os cartões.

### Legibilidade e largura da página (05/11/2026)
- **Largura das letras:** a largura de cada letra vem de `verdana10.txt`, medida por `mede_verdana.ps1`, com 10% de
  folga, porque o Excel desenha a letra mais larga conforme a escala da tela. O `org_impressao.py` para se alguma
  palavra não couber na célula. Corrigir o layout, nunca diminuir a letra.
- **Grade:** 60 colunas de 16 px e 3 colunas de folga de 5 px entre as caixas, num total de 740 pt. A impressora tem
  margem física e só imprime ~757 pt.
- **Conferir pelo COM:** `PageSetup.Pages.Count` tem de ser 6 e `VPageBreaks.Count` tem de ser 0.

## Aba DETALHE e números clicáveis (06/11/2026)

Os números de potencial (caixas, faixas de supervisor, carteira) e da tabela por segmento da "ORGANOGRAMA impressao" são
hiperlinks para a aba **DETALHE**, que traz as listas: RAZÃO SOCIAL, NOME FANTASIA, Nº DE LOJAS, CNPJ, STATUS AUTOPEL e
origem. É clique simples e sem macro. Os números de carteira vêm do CRM e não têm lista.

Ordem (na pasta `dados`):
1. `python ../org_dados.py "<MAPA>"`;
2. `python ../org_impressao.py`: marca os números com `link`;
3. `python ../org_detalhe.py "<MAPA>"`:
   - gera `detalhe.tsv` e `detalhe.json`;
   - põe a linha de destino (`alvo`) no `org_impressao.json`;
   - para se a contagem por status for diferente da do `org_dados.json`;
4. `powershell -File ../org_detalhe.ps1 -Arquivo "<cópia>"`: cria ou regrava a DETALHE, no fim da pasta;
5. `powershell -File ../org_grava.ps1 -Arquivo "<cópia>" -Layout org_impressao.json`: grava a impressão com os links,
   na mesma letra, cor e sem sublinhado.

A DETALHE e a impressão andam juntas, porque os links apontam para números de linha da DETALHE.
