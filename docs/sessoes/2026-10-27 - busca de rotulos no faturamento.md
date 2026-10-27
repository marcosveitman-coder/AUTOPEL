# Sessão 27/10/2026 — Busca de rótulos nas bases de faturamento

## Pedido
Pesquisar nas descrições de produto (`DES_PROD`) de todas as bases de faturamento (2018–2026) se a Autopel vendeu
algum rótulo.

## Entrega
`Bases Claude/rotulos vendidos 2018-2026.xlsx` (fora do git), com as abas:
- **Resumo:** valor por grupo e por ano;
- **Por produto;**
- **Por cliente:** com a última compra e o vendedor;
- **Linhas:** as 4.197 linhas encontradas.

## Regras
- Todas as abas de todos os `fat *.xlsx`, menos a Planilha1 de `fat 2025.2.xlsx`, que é cópia de dez/2025.
- Descrição sem acento e em maiúsculas. Três grupos:
  - **"Rótulo"** = a palavra ROTUL* na descrição, sem ser fita;
  - **"Etiqueta BOPP"** = ETIQUETA + BOPP, o material do rótulo adesivo;
  - **"Fita para rotulador"** = fitas Brother e Dymo. É suprimento, não é rótulo.

## Números de conferência

| Grupo | Linhas | Produtos | Valor |
|---|---|---|---|
| Rótulo (palavra na descrição) | 4 | 2 | R$ 25,2 mil (2020, 2024 e 2025) |
| Etiqueta BOPP | 3.014 | 61 | R$ 2,18 mi (687 clientes; vendas todo ano desde 2018, R$ 485 mil em 2025) |
| Fita para rotulador | 1.179 | 10 | R$ 266 mil |

## Lições
- Quase nada é vendido com a palavra "rótulo" na descrição: o rótulo adesivo aparece como "ETIQUETA BOPP".
- Em 2026, a coluna `DESC_FAMILIA` traz as famílias "ETIQUETA BOPP ACRILICO" e "ROTULADORES".

Scripts: `scripts cruzamento/avulsos/rotulos/` (`busca_rotulos.py` → `planilha_rotulos.py`).
