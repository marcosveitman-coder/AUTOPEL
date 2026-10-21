# Sessão 21/10/2026 (tarde) — aba "FARMACIA VF"

## Pedido
Cruzar a base de farmácias com a aba Ranking de `ranking_redes_farmacias_brasil_2025.xlsx`, com as mesmas premissas
da base consolidada:
1. completar e validar nome fantasia, razão social, CNPJ e nº de lojas;
2. tirar as duplicatas;
3. trazer a última compra (faturamento, produto personalizado e e-mail do cadastro).

Gravar **só** na aba vazia "FARMACIA VF", rankeada pelo nº de lojas. A aba "FARMACIA CONSOLIDADO" não é mexida. No
fim da aba vão duas colunas de valor 2025 e uma de porte da Receita. A busca na internet foi autorizada.

A sessão foi interrompida uma vez pelo limite de uso e retomada com "retome".

## Entrega
`MAPA MERCADO.xlsx`, aba "FARMACIA VF" (`sheet19.xml`), gravada pelo XML. As outras partes do arquivo ficaram
idênticas ao backup (conferido parte a parte).

Linhas: 91.193, uma por raiz de CNPJ, mais as linhas das redes e associações. Cabeçalho congelado, com filtro.

Colunas:
- A Posição
- B Rede/associação
- C Tipo
- D Entidade/federação
- E Nome fantasia
- F Razão social
- G CNPJ
- H Nº de lojas (considerado, que é o critério do ranking)
- I Nº de lojas (Receita)
- J Nº de lojas (pesquisa/Ranking)
- K UF
- L Município
- M Situação na Receita
- N Validação/observação
- O Data da última compra
- P Status
- Q Como foi achado
- R Cliente encontrado
- S Faturamento da farmácia em 2025 (R$, estimativa do Ranking)
- T Valor faturado pela Autopel em 2025 (R$)
- U Porte (Receita)

## Números de conferência
- Ranking: 131 redes, todas com os dados levados para a base. 44 delas viraram linhas novas, porque não estavam na
  base consolidada.
- CNPJ repetido: 0. A linha da DSG Farma ficou uma só: o CNPJ do Ranking, que é a matriz, foi unido à linha da
  associação.
- Com data de compra: 1.784. ATIVO (até 360 dias): 790.
- Com valor faturado em 2025: 780, somando cerca de R$ 203,1 milhões.
- Porte:

  | Porte | Linhas |
  |---|---|
  | ME | 70.627 |
  | EPP | 17.123 |
  | DEMAIS | 3.270 |
  | Sem porte (linha de rede sem CNPJ) | 173 |

## Regras e decisões
- A origem é a base consolidada. O usuário escreveu "FARMACIA VF", mas a VF era a aba vazia de destino.
- As 840 BAIXADAS que o usuário apagou da CONSOLIDADO ficam fora da VF também.
- "Valor faturado em 2025" virou duas colunas: o faturamento estimado da farmácia (Ranking) e o valor que a Autopel
  faturou para ela.
- **Nº de lojas considerado:**
  - o do Ranking, quando existe;
  - senão, para rede ou associação, o maior entre o da pesquisa e o da Receita;
  - para uma farmácia, o da Receita.
- **Rede com mais de uma raiz** (ex.: DPSP = Pacheco + São Paulo): a 1ª raiz leva os dados do Ranking e as outras
  ficam com as lojas da Receita e uma observação. Assim a rede não é contada duas vezes.
- **Rede-mãe (Farmarcas):** herda a compra mais recente das suas bandeiras.
- **CNPJ das associações novas:** buscado na Receita pela razão social. As dúvidas foram resolvidas pela BrasilAPI.

## Pendências / observações
- Nomes iguais com raízes diferentes (ex.: franqueados "Ultra Popular", "Farma Fácil") são empresas distintas e
  ficam em linhas próprias.
- O nº de lojas das associações na Receita fica abaixo do real, porque muitas associadas não levam a bandeira no nome.
  Por isso a coluna da pesquisa fica ao lado.

## Lições
- Ao comparar o rótulo de uma rede, usar o nome **inteiro**. Cortar em " (" separou "DROGARIA TOTAL (+ TOTAL POPULAR)"
  do "Grupo Total" e criou uma linha duplicada.
- O rollup da rede-mãe era feito só no gravador da CONSOLIDADO. Numa base nova é preciso refazer.
- Na `<sheetView>` da aba vazia vinha um `<selection>`. Para congelar o cabeçalho, trocar o bloco `<sheetViews>`
  inteiro.
- Os zips da Receita foram apagados no fim (`%TEMP%\receita_cnpj`).
