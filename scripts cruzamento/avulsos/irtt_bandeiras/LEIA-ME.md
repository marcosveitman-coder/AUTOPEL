# IRTT por bandeira (aba "I.R.T.T.__Por bandeira (FINAL)" do MAPA MERCADO)

Revisa a última compra (coluna G) e acrescenta, no fim da tabela (Y:AE): CNPJ validado, validação, faturamento 2025 da
bandeira e fonte, valor faturado pela Autopel em 2025, nº de lojas e fonte. Rodar numa pasta de trabalho (os `.pkl`
ficam nela), nesta ordem:

| # | Script | O que faz |
|---|---|---|
| 1 | `prep_irtt.py` | lê a aba, o IRTT (venda e lojas 2024 por grupo), o ABRAS 2026 e o ranking de farmácias |
| 2 | `receita_irtt.py` | busca as bandeiras na cópia local da Receita (`C:\dados_receita\AAAA-MM`), ~15 min |
| 3 | `valida.py` / `valida2.py` | classifica o CNPJ da coluna U (confere / não confere / vazio) e lista candidatos |
| 4 | `detalha.py` | detalha raízes na Receita (razão, matriz, ativos, nomes fantasia), ~3 min |
| 5 | revisão manual | decisões em `decisoes_irtt.json` (**local, fora do git**: tem CNPJ), lidas por `decisoes.py` |
| 6 | `final_cnpj.py` | CNPJ final por bandeira (+ CNPJ do grupo para bandeira sem CNPJ próprio) |
| 7 | `passo1.py` + `passo1b.py` | última compra: CNPJ validado, códigos da col. S, nome, produto e e-mail (regras do `../irtt.py`) |
| 8 | `passo34.py` | faturamento de mercado 2025 (web → ranking de farmácias → ABRAS → IRTT 2024), Autopel 2025 e lojas |
| 9 | `grava_irtt_vf.py` | grava G e Y:AE pelo XML (`../gravar_coluna_xml.gravar_celulas`) |

`web2025.py`: faturamento 2025 achado na internet para grupos fora de supermercado (fonte e métrica anotadas).

Variáveis: `AUTOPEL_MAPA`, `AUTOPEL_SCRIPTS` (pasta `scripts cruzamento`), `AUTOPEL_COD_DIR`, `AUTOPEL_CLIENTES_CSV`,
`AUTOPEL_IRTT_DIR`, `AUTOPEL_ABRAS`, `AUTOPEL_FARM_RANKING`, `AUTOPEL_SOMA2025`, `AUTOPEL_RECEITA`.

Regras principais (detalhes em `docs/sessoes/2026-10-21 - IRTT por bandeira.md`):
- CNPJ confere se a raiz tem lojas ativas com o nome da bandeira, ou se é empresa do grupo (razão social).
- Grupo com várias bandeiras: o faturamento só é dividido pelas lojas quando as lojas por bandeira cobrem 60% do grupo;
  senão o valor do grupo vai numa linha só (as outras dizem "incluído na linha de…"). Carrefour: divisão do release 1T25.
- Valor Autopel 2025: soma por raiz de CNPJ; raiz compartilhada entre bandeiras conta uma vez só.
