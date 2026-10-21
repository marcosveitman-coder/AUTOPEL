# Sessão 21/10/2026 (tarde) — aba "FARMACIA VF" (EM ANDAMENTO)

Sessão interrompida pelo limite de uso. Nada foi gravado na planilha ainda.

## Pedido
Cruzar a aba "FARMACIA CONSOLIDADO" (MAPA MERCADO) com a aba Ranking de `ranking_redes_farmacias_brasil_2025.xlsx`.
Repetir os passos 1-3 da sessão anterior: completar e validar, tirar duplicatas e trazer a última compra. Gravar só na
aba vazia "FARMACIA VF".

A aba deve sair rankeada pelo nº de lojas e ter duas colunas no fim:
- "valor faturado em 2025". Vou entregar duas colunas: o faturamento estimado da farmácia (Ranking) e quanto a
  Autopel faturou para ela em 2025;
- o porte da Receita.

## Estado
- O usuário apagou da CONSOLIDADO as 840 empresas BAIXADAS. Manter essa decisão na VF.
- Ranking: 131 redes.
  - `vf/liga_ranking.py` ligou 77 delas às redes de ontem; as correções manuais estão em `vf/redes_vf.py`.
  - As 54 novas têm bandeira e restrição de UF, e as lojas já foram contadas na Receita (cache de ontem).
- CNPJs do Ranking validados (cache da Receita + BrasilAPI). Farma Conde no Ranking = holding; manter a
  operação e citar a holding.
- `vf/soma2025.py`: soma do VLR_TOTAL de 2025 por raiz, COD_CLI e produto (rodou, `soma2025.pkl` no scratchpad).
- Empresas*.zip da Receita foram baixados de novo em `%TEMP%\receita_cnpj`, para buscar o CNPJ das associações
  novas pela razão social.

## Falta
1. Buscar o CNPJ das associações novas em Empresas*.zip e depois apagar os zips.
2. Montar a base VF: CONSOLIDADO + dados do Ranking (tipo, federação, lojas, faturamento estimado) + linhas das redes
   novas; associadas rotuladas.
3. Cruzar as redes novas com o faturamento (produto, nome, e-mail, associadas).
4. Colunas finais: faturamento 2025 (Ranking), valor faturado pela Autopel em 2025, porte da Receita.
5. Rankear pelo nº de lojas e gravar via XML na aba "FARMACIA VF".
