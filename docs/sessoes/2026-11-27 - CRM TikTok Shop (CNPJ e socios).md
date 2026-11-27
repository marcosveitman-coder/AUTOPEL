# 2026-11-27 — CRM de prospecção TikTok Shop: CNPJ e sócios

## Pedido
Lista de prospecção do TikTok Shop (`Downloads\CRM_Prospeccao_TikTokShop_StarLive.xlsx`, aba Leads, 1.606 lojas com menos
de 100 afiliados, dados do FastMoss). Com o nome da loja (B), Categoria (AD), Subcategoria (AE) e ID FastMoss (AM),
preencher AN–AT: CNPJ, nome CEO, LinkedIn CEO, nome diretor, LinkedIn diretor, nome gerente, LinkedIn gerente.

## O que foi feito
- Busca do nome da loja na cópia local da Receita (`mapa_final/busca_receita.py`) e 2ª passada com o nome sem espaços
  (`tiktok_shop/busca_colado.py`). Só nome igual, e CNAE compatível com a categoria.
- QSA pela BrasilAPI nos CNPJs achados; nomes do CEO e do diretor pelo QSA (MEI/EI: o titular, pela razão social).
- LinkedIn: testes com WebSearch e Firecrawl não acharam nenhum perfil. Ficou um link de busca de pessoas do LinkedIn.
- Gravado pelo XML da aba Leads (backup no scratchpad). Conferido: só `sheet3.xml` mudou, 0 células fora de AN–AT, e o
  Excel abre o arquivo.

## Números
| Resultado | Lojas |
|---|---|
| CNPJ achado (ativo) | 317 |
| nome CEO | 314 |
| nome diretor | 33 |
| ambíguo (várias empresas com o nome) | 120 |
| nome igual só em outro ramo | 44 |
| CNPJ baixado/inapto | 4 |
| não achado | 1.101 |
| nome curto demais | 20 |

Planilha de conferência (razão, fantasia, CNAE, UF, como foi achado, qualificação no QSA, candidatos dos ambíguos):
`Downloads\CRM_TikTokShop_conferencia_CNPJ.xlsx` (fora do git).

## Pendências / ideias
- Os 120 ambíguos podem ser resolvidos olhando a loja (cidade, produto) ou pela página da loja no TikTok, que não abre aqui.
- Gerente (AS/AT) ficou em branco: o QSA não tem e loja pequena não tem gerente público.

## Lições
- Loja de TikTok Shop com nome genérico ("Wave Shop") casa com várias empresas: sem a trava de ramo, entravam homônimos.
- `gravar_celulas` agora aceita fórmula (`{"f": ..., "v": ...}`).
