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

## Cuidados

- **O potencial é do segmento:** quem divide o segmento divide as mesmas oportunidades; não somar entre pessoas. No
  supervisor e na diretoria, cada segmento da equipe conta uma vez só.
- **CLIENTES** = ativos + perdendo + a trabalhar (cartões do CRM). Onde há clientes sem segmento, NO + FORA fica menor
  que CLIENTES.
- **Carteira dos supervisores:** é a visão do CRM de cada um e não bate com a soma das pessoas da equipe.
- **Faixa compartilhada:** caixas e quadros na mesma faixa de linhas dividem a altura. O quadro ao lado das caixas
  fica sem altura fixa, senão estica os cartões.
