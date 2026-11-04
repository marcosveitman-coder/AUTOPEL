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
  - coluna P da "FARMACIA VF", cabeçalho na linha 6.
- **Segmento do CRM → aba:**
  - ALIMENTACAO → ALIMENTACAO FINAL;
  - REDE DE LOJAS → REDE DE LOJAS FINAL;
  - CONSTRUCAO → CONSTRUÇÃO FINAL;
  - DROGARIAS → FARMACIA VF.

  Os outros segmentos ainda não têm lista de mercado e aparecem sem potencial.

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

As mesmas informações, em **6 folhas A4 paisagem impressas a 100%**. A letra é Verdana, sempre com tamanho 10 ou
maior e sem itálico (pedido do usuário):
1. equipe menor;
2. equipe maior, parte 1: vendedores e venda interna;
3. equipe maior, parte 2: representantes;
4. carteira de cada pessoa: diretoria e equipe menor;
5. carteira de cada pessoa: equipe maior;
6. potencial por segmento e segmentos × responsáveis.

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
