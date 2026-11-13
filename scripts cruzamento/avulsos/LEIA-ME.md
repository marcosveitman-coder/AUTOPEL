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
| `compras/` | Planilha de compras "Relatorio Revenda atualizado". `revisa_relatorio_revenda.ps1` grava pelo Excel (COM) a programação do trimestre corrigida na Table1, o tracking da aba dinamicas e as abas LEIA-ME COMPRAS e PROGRAMACAO TRIMESTRE. Rodar numa cópia; recusa se a Table1 já tiver as colunas novas. `confere_revenda.py ORIGINAL REVISADO` confere célula a célula e refaz a conta em Python. Troca de nome das colunas sem perder campo nas dinâmicas: `renomeia_previsao.py` (XML: tabela + cache + fórmulas) → `renomeia_previsao.ps1` (Excel: atualiza e devolve larguras) → `confere_renomeia.py` (valores, estilos, larguras e formatação das dinâmicas). `devolve_layout.py` devolve larguras e alturas de um arquivo de referência pelo XML. `diagnostico_disponivel.py` lista os SKUs com o Disponível sem vírgula. |
| `compras/claude/` | Metodologia CLAUDE de compra (26/10/2026). Lê as fotos diárias, o Protheus, o histórico de entradas e as ordens de compra do TOTVS e as bases de faturamento. Gera os parâmetros por SKU (sazonalidade, CV, prazo do fornecedor, ABC, múltiplo) e grava pelo Excel a aba param CLAUDE, as colunas CLAUDE da Table1 e a aba de análise. `calcula_claude.py` é a referência em Python para conferir as fórmulas. `saude_ecommerce.py` + `grava_saude.ps1` + `confere_saude.py`: saúde do estoque, triagem para e-commerce e concentração de clientes (seções 7–9 da análise). |
| `compras/kalunga/` | Produtos Autopel × Kalunga (13/11/2026): `cruza_kalunga.py` acha família MACRO/MICRO e produto Kalunga (EAN, descrição, URL) de cada SKU da Geral; `grava_atl.py` grava a aba ATLxKALUNGxML pelo XML; `familias_site.py` baixa a lista de produtos de cada família do site da Kalunga (json usado pelo `cruza_kalunga.py`). |

| `mapa_final/` | Abas ALIMENTACAO FINAL, REDE DE LOJAS FINAL e CONSTRUÇÃO FINAL do MAPA MERCADO (03/11/2026): junta as abas de origem sem duplicidade, confere ou acha o CNPJ de cada rede na Receita local (operadora/franqueadora, nunca franqueado) e completa com ANAMACO, IRTT e redes regionais por CNAE. Ordem de execução e regras no `LEIA-ME.md` da pasta; dados e decisões em `dados/`, fora do git. |
| `kraft_heinz/` | Prospecção de um grupo (29/10/2026). `receita_grupo.py SAIDA.json "REGEX"` acha na Receita as empresas pela razão social e todos os estabelecimentos das raízes (filtrar depois: o regex pega homônimos). `linhas_cliente.py SAIDA.json RAIZ "fat ….xlsx"…` tira todas as linhas de uma raiz das bases de faturamento, com todas as colunas. |
| `pesquisa_cliente/` | Pesquisa completa de um cliente (09/11/2026): roteiro no `LEIA-ME.md` da pasta (Receita pontual + IE, TOTVS, bases, empresas ligadas, domínio do e-mail, internet). `receita_vinculos.py` acha na Receita local a raiz e as empresas ligadas por razão, telefone, e-mail, endereço ou fantasia; `resumo_cliente.py` junta com as linhas do `linhas_cliente.py`. |
| `midia/` | Ver fotos e vídeos sem ffmpeg nem Pillow. `folha_contato.ps1 -Pasta -Saida` monta grades numeradas das fotos únicas (System.Drawing). `serve_range.py PORTA PASTA [SAIDA]` serve vídeos com suporte a Range para abrir no navegador embutido (entrada em `.claude/launch.json`) e tirar quadros num `<canvas>`; com SAIDA, grava o que chegar por PUT. `remove_marca.html` (local, fora do git) tira a marca d'água translúcida do Movavi e desfoca um logo, gravando o vídeo pelo navegador: `processar()`, depois `gravar()`. `mp4info.py ARQ.mp4` conta amostras e duração por trilha. |

A extração do PDF do Ranking ABRAS (2025 e 2026) usa `pypdf`: texto por linha, depois uma regex de
`posição  razão social  UF  faturamento`. O passo a passo está registrado em `docs/sessoes/`.
