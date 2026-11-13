# Sessão 13/11/2026 — produtos da Autopel × famílias e EAN da Kalunga (aba ATLxKALUNGxML)

Registro sem nomes de clientes, compradores ou fornecedores (repositório público).

## Pedido
No "Relatorio Revenda atualizado.xlsx" (COMPRAS):
- na aba **ATLxKALUNGxML** (vazia), trazer exatamente as colunas B, C, D e G da aba Geral (Codigo, Descricao, Marca,
  Ult.Preco);
- cruzar com a aba **kalunga familias** e criar **Familia (MACRO)** e **Familia (MICRO)**, para organizar o universo da
  Autopel como o da Kalunga;
- cruzar com a aba **kalunga_ean** (C ean, D nome, F produto_url) e criar **EAN**, **DESCRIÇÃO KALUNGA** e **produto_url**.

## O que as abas têm
- **kalunga familias:** 1.348 pares MACRO/MICRO (28 MACRO) do mapa do site. Só os nomes, sem ligar produto a família.
  O mesmo nome MICRO aparece em várias MACRO (ex.: "Apontadores" em Escolar, Escrita e Escritório).
- **kalunga_ean:** 15.896 produtos com produto_id, EAN, nome, marca e URL. O nome vem com o acento quebrado (UTF-8 lido
  como latin-1) e foi consertado na leitura.
- **Geral:** 4.251 SKUs. A descrição não tem padrão (abreviações como ADES., ESF., PLAST., OF., KN; cores AM/AZ/VM).
- Já existiam vínculos pesquisados:
  - "… (KALUNGA)" (pesquisa de preços): 478 SKUs com EXATO/SIMILAR e confiança;
  - "MERCADO LIVRE (…)": vínculo Kalunga por CODIGO, MEDIA ou BAIXA.

## Regras usadas (`avulsos/compras/kalunga/cruza_kalunga.py`)
- **Produto Kalunga (EAN, descrição, URL)**, em ordem:
  1. vínculo já pesquisado: pesquisa Kalunga EXATO, ou SIMILAR com confiança ≥ 70; Mercado Livre por CODIGO ou MEDIA (BAIXA
     tinha erros, ex.: 76x102 ligado a 76x76);
  2. código do fabricante (TN116BR, TZE231, GI190): todos os códigos da descrição no nome Kalunga, e o mesmo
     substantivo. Medida e quantidade (20FLS, 300MTS, 48X50M) não contam como código;
  3. texto:
     - o 1º substantivo da descrição tem de estar no nome Kalunga, nas 3 primeiras palavras;
     - nota = cosseno das palavras pesadas pela raridade;
     - os números (medida, gramatura, quantidade) têm de bater, inclusive a unidade: 25MM ≠ 25FL;
     - cor diferente e "SEM GAS" × "com gás" tiram nota; a mesma marca soma.
- **Classes do vínculo:**
  - EXATO: pesquisa EXATO, código ou mesma marca com nota ≥ 0,75;
  - SIMILAR: nota ≥ 0,55 com os números iguais (pode ser outra marca ou cor).
- **Família:**
  - a MICRO que contém o 1º substantivo da descrição (ou sinônimo: PINCEL ↔ MARCADOR, REGISTRADOR ↔ PASTA) e mais
    palavras em comum. Família de várias palavras precisa de outra palavra em comum além do substantivo, para
    "ESPELHO" não cair em "Papel Dobradura (Espelho)" e "ÁGUA SANITÁRIA" em "Água Mineral";
  - algumas palavras vão direto para uma família (BOTA/COLETE → EPI, ESPIRAL → Encadernação, CARTOLINA → Papel
    Cartolina);
  - família genérica só quando nada casou (ETIQUETA → Etiquetas Adesivas, CADEIRA → Cadeiras Giratórias);
  - MICRO repetida em várias MACRO: vale uma ordem de prioridade (Cartuchos, Impressoras, Papéis, Envelopes, Escrita,
    Escritório…).

## Entrega
- **Aba ATLxKALUNGxML**, gravada pelo XML da aba (`grava_atl.py`):
  - A–D: Codigo, Descricao, Marca e Ult.Preco da Geral, iguais (cabeçalho e valores);
  - E Familia (MACRO), F Familia (MICRO), G EAN (texto), H DESCRIÇÃO KALUNGA, I produto_url;
  - J Vínculo Kalunga (EXATO/SIMILAR) e K Como foi cruzado (fonte e nota): colunas de apoio para filtrar o que
    conferir;
  - filtro, cabeçalho congelado e estilos da Geral.
- **Números:**
  - família em 3.250 dos 4.251 SKUs (76%), em 382 famílias MICRO;
  - produto Kalunga em 735: 170 EXATO e 565 SIMILAR; 727 com EAN;
  - fontes do vínculo: texto 413, pesquisa Kalunga 234, pesquisa Mercado Livre 71, código do fabricante 17.
- **Sem família (1.001):** a Kalunga não tem o item, ou é personalizado. Principais grupos:
  - sacolas plásticas de rede (122), envelopes personalizados, lacres;
  - camisas e camisetas;
  - refrigeradores, baguetes, placas e displays de loja.

## Conferência
- Só a parte `sheet39.xml` (ATLxKALUNGxML) mudou; todas as outras partes do pacote são iguais byte a byte.
- As 4.251 linhas de A–D batem com B, C, D e G da Geral.
- Abre no Excel com 45 abas e 67 dinâmicas.
- O arquivo ficou menor (30,0 → 24,7 MB) só pela compressão do zip.
- Gravado às 11h23 e regravado às 11h29, depois de conferir que o original não tinha mudado desde a cópia. Backup no
  scratchpad.
  - Motivo da regravação: quando dois produtos Kalunga empatavam na nota, a escolha variava entre execuções
    (4 SKUs). Agora vale o menor produto_id e o script dá sempre o mesmo resultado.
  - Na regravação, de novo só a sheet39 mudou em relação ao arquivo das 11h06, e as colunas E–J batem com o cruzamento
    nas 4.251 linhas.

## Pendente
- ~~Família pelo próprio site~~ (feito na 2ª parte): `familias_site.py` baixa, pelo mapa do site e pela API das páginas de família
  (`/api/obterProdutosDepartamento/<posição>`), a lista de produtos de cada família. Com isso, o SKU que tem produto
  Kalunga herda a família do site, em vez da regra de texto.
  - A API foi testada (15 produtos por página).
  - O download completo (873 famílias, ~30 min) **foi bloqueado pela permissão automática da sessão**: falta
    autorização do usuário.
- Os vínculos SIMILAR e a família pela descrição são aproximados: conferir por amostra, filtrando as colunas J e K.

## Lições
- No PowerShell, `$S` e o laço `foreach ($s ...)` são a mesma variável (de novo): usar nomes longos (`$PASTA`).
- Nome de família da Kalunga no plural: o radical precisa tratar ÕES → ÃO, AIS → AL, EIS → EL e ORES → OR, sem
  cortar "LACRES" em "LACR".
- **Desempate em busca por nota:** percorrer um `set` de strings muda de ordem a cada execução (hash aleatório). Ordenar
  antes, para o resultado ser reproduzível.

---

## 2ª parte: família pelo site da Kalunga (autorizada pelo usuário)
- **Download (`familias_site.py`):**
  - 1.279 entradas do mapa do site, contando as famílias repetidas em mais de uma MACRO;
  - 12.971 produtos com família, sem nenhum erro e nenhuma família vazia;
  - levou cerca de 50 minutos, com pausa de 0,4 s entre as consultas.
  - O json ficou no scratchpad, fora do git.
- **Família no cruzamento, em ordem:**
  1. a do produto Kalunga achado (EXATO ou SIMILAR), segundo o site: 668 SKUs;
  2. sem produto equivalente, a do produto Kalunga mais parecido, com o mesmo substantivo e nota ≥ 0,40: 377 SKUs;
  3. senão, pela descrição: 2.249 SKUs.
- **Resultado:** família em 3.294 dos 4.251 SKUs (77%; eram 3.250). Sem família: 957. A coluna K diz de onde veio
  cada família.
- **Gravação (regravada às 11h56):**
  - o original não tinha mudado desde as 11h29;
  - só a sheet39 mudou em relação ao arquivo das 11h06;
  - A–D batem com a Geral e E–J com o cruzamento nas 4.251 linhas;
  - abre no Excel com 45 abas e 67 dinâmicas.
- **Erros vistos por amostra no "produto parecido":** "COPO DE VIDRO" → copos plásticos; "ÁLCOOL LÍQUIDO" → álcool em
  gel. Conferir filtrando a coluna K.
