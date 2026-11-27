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

## 2º pedido: resolver as 120 ambíguas
- Pesquisa na web (Firecrawl) de ~40 lojas: cidade no Instagram/site, nome do vendedor na Shopee/Mercado Livre.
  17 resolvidas como "PROVÁVEL (web)"; 5 pela regra "todos os candidatos na mesma cidade". Decisões em
  `tiktok_shop/dados/dec_ambiguos.json` (local).
- Os créditos diários do Firecrawl acabaram (~45 buscas); o WebSearch não traz loja brasileira pequena. Ficaram **98**
  ambíguas.
- Planilha regravada: 339 CNPJs, 336 CEOs, 34 diretores. Conferido: só AN–AT mudaram (68 células), Excel abre.
  A conferência mostra o nível (CONFIRMADO / PROVÁVEL (web) / PROVÁVEL (mesma cidade)) e a fonte de cada desempate.

## 3ª rodada (mesmo dia): Bing por script
- `tiktok_shop/busca_bing.py`: 2 buscas no Bing por loja (192 consultas, sem crédito; o DuckDuckGo bloqueia script) e cruzamento
  automático com a cidade dos candidatos. Só 2 sugestões, aceitas após revisão (as duas de São Paulo).
- Planilha final do dia: **341 CNPJs** (317 confirmados + 24 prováveis), 338 CEOs, 34 diretores; 96 ambíguas.
  Conferido: só AN–AT mudaram, Excel abre.
- Tarefa agendada para 28/11 9h15 (`tiktok-shop-desempate-ambiguos`) tenta as 96 com o crédito novo do Firecrawl.

## 4º pedido: link do perfil de verdade no LinkedIn (conta do usuário no Chrome)
- O usuário pediu o link do **perfil** do decisor, não a busca por nome. Conta conferida pelo e-mail principal na página de
  e-mails do LinkedIn (Claude in Chrome, sessão já logada pelo usuário).
- Busca por nome do sócio não serve (cartão só mostra nome e cargo; dezenas de homônimos: 0 de 12). Busca **pela loja**
  (`"<loja>"`, `tiktok_shop/linkedin_lista.py` + `linkedin_funcoes.js`, lotes de 12 com 8–10 s entre buscas,
  `li_passo.py` grava cada lote em `dados/li_result.json`, local):
  - CEO = sócio do QSA (1º nome + um sobrenome) com a loja no cartão, ou cargo CEO/fundador/sócio/proprietário;
  - diretor/gerente pelo cargo; quem não é sócio precisa de "na/no/at/em <loja>" e do estado da empresa;
  - cargo "Anterior:" não vale; "Product Owner" não é dono; revisão à mão de cada lote (homônimos de outra empresa caíram).
- 310 das 340 lojas pesquisadas; **35 lojas com perfil** (27 CEO, 6 diretores, 9 gerentes). A sessão caiu na busca 311
  (tela de cadastro): parado sem mexer. Faltam 30 lojas de prioridade C.
- Retomada após novo login: mais 20 lojas (o LinkedIn já mostrava só 3 cartões por busca = limite mensal), +5 lojas com
  perfil. A sessão caiu de novo; **faltam 10 lojas** (L1569–L1604). Total: **40 lojas com perfil** (30 CEO, 9 diretores,
  12 gerentes aprox.). O localStorage some quando a sessão cai: a função é recarregada de `linkedin_funcoes.js`.
- Planilha: AP/AR/AT = "Ver perfil" (link do perfil) ou vazio; AS = nome do gerente. Quando o decisor achado não é o sócio
  do QSA, a AO/AQ passa a ter o nome da pessoa do link. Conferência com nome, cargo no perfil, como achou e link.
- Bugs da rotina: `''` dentro de string virou backspace (cargo nunca casava); lote de 20 estourou o tempo do navegador.

## Pendências / ideias
- Os 98 ambíguos restantes podem ser resolvidos (Firecrawl renova o crédito no dia seguinte) olhando a loja (cidade, produto) ou pela página da loja no TikTok, que não abre aqui.
- Gerente (AS/AT) ficou em branco: o QSA não tem e loja pequena não tem gerente público.

## Lições
- Loja de TikTok Shop com nome genérico ("Wave Shop") casa com várias empresas: sem a trava de ramo, entravam homônimos.
- `gravar_celulas` agora aceita fórmula (`{"f": ..., "v": ...}`).
