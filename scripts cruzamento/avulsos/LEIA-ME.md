# Scripts avulsos

Scripts das entregas pontuais. Ficam guardados como **referência**: dependem de arquivos intermediários (`.pkl`) e de
planilhas que mudam a cada pedido. Para reaproveitar, ajuste os caminhos ou use as variáveis de ambiente:
`AUTOPEL_CLIENTES_CSV` (Clientes - SQL.csv), `AUTOPEL_IRTT` e `AUTOPEL_ABRAS`. Os detalhes de cada entrega estão em
`docs/sessoes/`.

| Script | O que faz |
|---|---|
| `gravar_coluna_xml.py` | **Reutilizável.** Grava uma coluna de texto numa aba editando só o XML daquela aba. Preserva tabelas dinâmicas, gráficos e as outras abas. |
| `extract_cod.py` | Lê as bases `fat *.xlsx` (2019–2026) e guarda a última venda por **código de cliente** (`COD_CLI`). |
| `irtt.py` | Cruza as bandeiras do ranking IRTT com o faturamento: CNPJ/razão indicados, nome do cliente, produto personalizado e domínio de e-mail. Usa as faixas de marca exclusiva × comum. |
| `choose_irtt.py` | Escolha da data final: forte > fraco, e o fraco só supera o forte se o comprador tiver o nome da rede. |
| `write_irtt.py` | Junta os códigos de cliente do IRTT aos vínculos, confere com "dias sem comprar" e grava a data (`--gravar`). |
| `vend_irtt.py` | Vendedor da carteira (cadastro) para cada bandeira do IRTT. |
| `vend_reuniao.py` | Vendedor do cadastro (com recurso ao da última venda) para a lista de supermercados. |
| `abras_cru.py` | Cruza o ranking ABRAS (razão social + UF) com o faturamento, incluindo domínio de e-mail. |
| `resumos/` | Abas de resumo do MAPA MERCADO: `gerar_spec.py` calcula textos e números a partir dos dados; `td_resumos.ps1` grava pelo Excel (COM), com textos, tabelas e **tabelas dinâmicas**; `conta_sup.py` conta lojas de supermercado por raiz na Receita; `confere_td.py` e `confere_tudo.py` fazem as conferências. Nº de lojas e duplicidades da aba da reunião: `receita/` (3 rodadas na Receita), `consolida_reuniao.py` (decide lojas, fonte e linhas que saem), `ajusta_reuniao.ps1` (grava e exclui pelo Excel) e `confere_reuniao.py`. `web_lojas.json`/`raiz_manual.json` ficam locais. |
| `conta_filiais.py` | Conta estabelecimentos ativos por raiz de CNPJ nos dados abertos da Receita (`Estabelecimentos*.zip`). |

A extração do PDF do Ranking ABRAS (2025 e 2026) usa `pypdf`: texto por linha, depois uma regex de
`posição  razão social  UF  faturamento`. O passo a passo está registrado em `docs/sessoes/`.
