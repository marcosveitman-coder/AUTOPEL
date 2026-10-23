# CLAUDE.md — Bases Claude (Autopel)

Planejamento comercial da Autopel. O trabalho aqui é cruzar listas de mercado (supermercados mapeados, ranking ABRAS,
ranking IRTT, listas do CRM) com as **bases de faturamento** e com o **cadastro de clientes** do TOTVS, para dizer se
o cliente já é atendido, quando foi a última compra e qual vendedor o atende. O usuário fala português e trabalha em
Excel: as entregas são planilhas, e as explicações são em português simples.

---

## ⭐ Golden rules

### Golden rule nº 1 — fechar toda sessão documentada e versionada
**Ao final de cada sessão de trabalho**, antes da mensagem final:

1. **Documentar** tudo o que foi feito em `docs/sessoes/AAAA-MM-DD - <assunto>.md`, seguindo o modelo das sessões
   que já existem. Registrar: pedidos, entregas (arquivo, aba, coluna), regras usadas, números de conferência,
   decisões do usuário, o que ficou pendente e lições aprendidas.
2. **Atualizar este CLAUDE.md** se surgir regra nova, armadilha nova ou mudança de layout das bases.
3. **Salvar os scripts** que geraram as entregas em `scripts cruzamento/` (ou `scripts cruzamento/avulsos/`).
   O scratchpad é apagado; o que não estiver no repositório se perde.
4. **Commit e push**: `git status` → conferir que **só código e documentação** entraram → `git commit` → `git push`.

Se o usuário encerrar no meio de uma tarefa, documentar o estado (o que está feito, o que falta) e versionar mesmo assim.

---

## 🔒 Repositório público: o que nunca pode subir

O remote `origin` (github.com/marcosveitman-coder/AUTOPEL) é **PÚBLICO**.

- O `.gitignore` é lista de permissão: só entram `*.py`, `*.bat`, `*.txt`, `*.md`, `*.sql`. Planilhas e CSVs
  nunca entram. Não forçar (`git add -f`) arquivo de dados.
- **Nunca** escrever em arquivo versionado nem em mensagem de commit:
  - nomes de clientes junto com data de compra ou vendedor;
  - CNPJs, códigos de cliente ou nomes de vendedores;
  - caminhos com nome de outras pessoas.
  Use contagens e exemplos genéricos ("marca colada × separada").
- Caminhos de arquivos de dados nos scripts: relativos à pasta `Bases Claude` ou por variável de ambiente
  (`AUTOPEL_CLIENTES_CSV`, `AUTOPEL_IRTT`, `AUTOPEL_ABRAS`).
- PowerShell 5.1: `git commit -F -` com here-string não funciona → gravar a mensagem num arquivo e usar `-F arquivo`.
  O push já tem credencial salva.

---

## 📂 Estrutura

```
Bases Claude/                       (fora do git: fat *.xlsx, supermercados*.xlsx, Ultimo faturamento - SQL.csv)
├── CLAUDE.md                       este arquivo
├── docs/sessoes/                   registro de cada sessão (golden rule nº 1)
└── scripts cruzamento/
    ├── rodar_cruzamento.bat        supermercados.xlsx × faturamento (extract → match → build)
    ├── extract.py / match.py / build.py / common.py / conferir.py
    ├── LEIA-ME.txt                 como rodar e regras, para o usuário
    └── avulsos/                    scripts das entregas pontuais (IRTT, ABRAS, vendedor, XML…) + LEIA-ME.md
        └── farmacias/              base consolidada de farmácias (Receita + redes + faturamento) + LEIA-ME.md
```

Fontes de dados fora desta pasta (só leitura):
- `Clientes - SQL.csv`: cadastro de clientes do TOTVS (Codigo, Loja, Nome, N Fantasia, CNPJ/CPF, **Vendedor = código**,
  Estado, E-mail NF-e, Segmento…). Fica numa pasta compartilhada de relatórios TOTVS.
- `Ultimo faturamento - SQL.csv` (nesta pasta): por cliente, a última nota e o vendedor (código + nome).
  É a fonte da tradução **código → nome do vendedor**.
- Planilhas de trabalho do usuário: OneDrive `Planejamento comercial_AUTOPEL\2026\COMERCIAL (NOVO)\` e `Downloads\`.

---

## 📊 Bases de faturamento (`fat AAAA.N.xlsx`)

- São cerca de 6,3 milhões de linhas (2018–2026). Ler **pelo nome do cabeçalho, nunca pela letra**: o usuário descreve
  colunas por letra e o layout muda.
- Colunas usadas: `EMISSAO`, `COD_CLI`, `COD_LOJA`, `COD_VENDEDOR`, `NOME_VENDEDOR`, `CNPJ`, `Estado`, `CLIENTE`
  (razão social), `DES_PROD`, `VLR_TOTAL`, `SEGMENTO`.
- 2026 tem uma coluna `MUNICIPIO` extra, então `DES_PROD` vai da P para a Q. `fat 2025.2.xlsx` tem **2 abas**. A coluna
  D é `COD_LOJA`, não o nome fantasia.
- `EMISSAO` vem às vezes como texto `dd/mm/aaaa   ` e às vezes como data.
- A leitura com openpyxl `read_only` leva cerca de 11–13 min com 4 processos. Os agregados ficam em cache em
  `%TEMP%\cruzamento_supermercados\agg\*.pkl` (por raiz de CNPJ, por cliente+UF, por produto+UF). Reaproveitar o
  cache em vez de reler.
- Quando o pedido for "2019 até 2026", pular o pkl de 2018.
- **IRTT (Ranking-Top300-IRTT.xlsx):** venda e lojas vêm com o ponto de milhar lido como decimal (120,6 = R$ 120.600 mi;
  1,007 = 1.007 lojas). Valores abaixo de R$ 50 mi são dado não divulgado. Os números são de 2024 por grupo, não por bandeira.
- O `SEGMENTO` do faturamento é pouco confiável: há rede de lojas classificada como DROGARIAS e distribuidor como
  REVENDA. Não usar como única trava.

---

## 🏛️ Dados abertos do CNPJ (Receita Federal)

- **Fonte oficial** para completar e validar CNPJ, razão social, nome fantasia, situação, município e nº de lojas.
  O usuário autoriza o download.
- **Onde baixar:** `https://arquivos.receitafederal.gov.br/public.php/webdav/AAAA-MM/` (WebDAV, usuário
  `YggdBLfdninEJX9`, senha vazia).
- **Arquivos:** `Estabelecimentos0-9.zip` (~5,3 GB), `Empresas0-9.zip` (~1,3 GB) e `Municipios.zip`.
- **Cópia local permanente** (decisão do usuário, 21/10/2026): `C:\dados_receita\AAAA-MM\` (hoje `2026-10`, 6,4 GB),
  fora do OneDrive e do git. Os scripts leem de `AUTOPEL_RECEITA` ou da pasta mais nova em `C:\dados_receita`.
  **Não baixar de novo nem apagar.** Só baixar o mês novo quando o usuário pedir ou a cópia passar de um mês; aí
  apagar a pasta do mês antigo.
  - Não vai para o GitHub: arquivos de até 2,3 GB, acima do limite de 100 MB, e o repositório é público.
  - Inclui `Municipios.zip` (código → nome do município).
- **Estabelecimentos**, separados por `";"`, em latin-1. Índices das colunas:
  - 0 raiz, 1 ordem, 2 DV, 3 matriz/filial (1/2), 4 nome fantasia;
  - 5 situação: 2 ativa, 3 suspensa, 4 inapta, 8 baixada;
  - 11 CNAE principal, 19 UF, 20 código do município, 27 e-mail.
- **Empresas:** 0 raiz, 1 razão social, 2 natureza, 4 capital, 5 porte.
- **Farmácia** = CNAE principal 4771-7/01, /02 ou /03. O 4771-7/04 é veterinária e fica fora. O universo ativo tem
  cerca de 120 mil lojas e 92 mil empresas.
- **Nº de lojas** de uma empresa = estabelecimentos **ativos** da raiz, filtrados pelo CNAE do segmento.
  - **Supermercado:** CNAEs 4711-3/01, 4711-3/02, 4712-1/00, 4729-6/99 **e os de atacarejo** (4691-5/00, 4639-7/01,
    4639-7/02, 4693-1/00), sem estabelecimento chamado CD, depósito ou escritório. Atacarejo usa CNAE de atacado: sem eles,
    a Tenda dá 0 loja.
  - **Achar a raiz pelo nome:** razão igual + UF; depois o miolo igual (sem palavras de tipo, com iniciais, sem o CPF
    do fim). A raiz só vale se tiver loja ativa na UF, porque razão genérica de grupo casa com holding ou empresa de
    serviços. Para busca em massa, indexar pela 1ª palavra (`avulsos/resumos/receita/`).
- Scripts em `scripts cruzamento/avulsos/farmacias/` (pipeline completo da base de farmácias).
- **Consulta pontual de CNPJ:** BrasilAPI `https://brasilapi.com.br/api/cnpj/v1/<14 dígitos>`. Usar via `curl`, porque o
  urllib do Python leva 403.
- **CNPJ vindo de base de terceiros** (ex.: Neoway): validar o dígito verificador. Cerca de 28% vieram corrompidos;
  a razão social costuma estar certa e serve para corrigir.

---

## 🔗 Regras de cruzamento (validadas com o usuário)

Ordem de força: **CNPJ raiz (8 dígitos) / código de cliente > razão social igual > nome fantasia, produto
personalizado e domínio de e-mail**.

1. **Fortes** (CNPJ raiz, código de cliente, razão social normalizada sem LTDA/SA/EIRELI/ME/EPP, com `&` = `E`):
   valem em qualquer UF.
2. **Fracos** (nome do cliente, marca no fim da `DES_PROD` depois da medida tipo `80X40M`, domínio do e-mail no cadastro):
   - comparar o **núcleo** da marca (sem palavras de tipo nas pontas: SUPERMERCADO, ATACADO, COMERCIO…);
   - aceitar marca **colada ou separada** ("SUPER X" = "SUPERX", só com prefixos SUPER/HIPER/MINI/MEGA/MAXI);
   - "SUPER X" nunca vira só "X", senão casa com outra empresa que só tenha o "X";
   - **mesma UF** (lista de supermercados) ou **UF listada pelo sistema** (IRTT);
   - marca com 5+ letras;
   - marca **comum ou franquia** (no nome de mais de 5 empresas do cadastro) exige UF **e** a frase completa da
     bandeira. Marca exclusiva vale em qualquer lugar;
   - bandeira com "supermercado/atacado/hiper" no nome só aceita comprador supermercadista; revenda fica fora.
3. Empresa achada por um vínculo fraco → a data e o vendedor vêm da **última venda para a raiz do CNPJ dela**
   (qualquer produto).
4. Um fraco só **substitui** um forte se for mais recente **e** o próprio comprador tiver o nome da rede.
5. **Vendedor**:
   - "vendedor da última venda" = `NOME_VENDEDOR` do dia (empate → maior valor);
   - "vendedor da carteira" = vendedor do **cadastro** (Clientes - SQL) do cliente que comprou por último. Se não houver,
     usar a raiz e depois os códigos de cliente;
   - ignorar carteiras que não são pessoas: INATIVOS, CLIENTE SEM COMPRA, RD/REPRESENTANTE, AUTOPEL - VENDA ENCERRADA,
     código 000000.
6. Sempre **conferir** contra o que já existe na planilha (datas do CRM, "dias sem comprar", "Vendedor principal") e
   explicar as divergências.
7. **Razão social igual não vence CNPJ:** se a linha tem CNPJ válido (dígito verificador) e achado na Receita, e o
   comprador é de **outra raiz**, é homônimo e o vínculo cai. Já vimos 24 casos em farmácias. Exceção: CNPJ da linha
   corrompido (base de terceiros), em que o comprador tem o CNPJ verdadeiro.
8. **Relatório "ultimo fat" do TOTVS** (`ultimo fat (dd.mm.aa).xlsx`: por cliente, DT_FAT, vendedor e grupo):
   - usar a data **só depois do fim das bases de faturamento**. Antes disso vale a data das bases: ~54 raízes tinham
     nota anterior que não aparece nas bases de venda;
   - a coluna E "Descrição" é o **grupo do cliente** (ex.: o grupo de uma rede de farmácias). Ligar grupo → rede por um
     **mapa curado**: os CNPJs dos clientes do grupo caem nas linhas da rede **e** o nome bate; na dúvida, a cidade/UF
     decide. Grupos têm sufixo de vendedor/região ("GRUPO X - FULANO"). Mapa de farmácias:
     `avulsos/farmacias/final/grupos_redes.json`, local e fora do git.
9. **Cliente achado por código ou por grupo do cadastro** (IRTT, 23/10/2026):
   - **código**: cai se não tem relação de nome (bandeira, grupo, razão indicada) **e** é de outro segmento. Sem relação de
     nome, só vale comprador do **segmento principal** da bandeira (supermercado → SUPERMERCADISTA). REDE DE LOJAS sozinho
     deixava passar loja de roupas para supermercado;
   - **grupo do cadastro**: vale só com o CNPJ da bandeira, o nome da bandeira no cliente, ou um grupo com o nome da
     bandeira + segmento compatível + cliente que não seja distribuidor (franqueado). O grupo pode ser uma central de
     compras ("REDE BRASIL") ou incluir uma empresa vendida (Sendas em "PAO DE ACUCAR");
   - CNPJ validado de **outro** grupo nunca vale;
   - **marca colada** só com nome de mais de uma palavra: "DROGAL" colado batia em "DROGA LIDER".

Armadilhas reais já vistas:
- drogaria e supermercado com o mesmo nome;
- sobrenomes (LOPES, ALMEIDA) e palavras comuns (VERDE, MODELO, TORRE) casando com qualquer cliente;
- loja de calçados com o mesmo nome de supermercado;
- código de cliente reaproveitado por outra empresa;
- lista de códigos do IRTT cortada com "…";
- bandeira "contida" no nome de outra empresa ("Drogaria Rosário" × "Farmácia N. S. do Rosário", "Araujo" ×
  "Araujo & Tomaz"). Exigir **núcleo igual**, e em rede com CNPJ próprio só aceitar comprador do próprio grupo.

**Associações e franquias de farmácia** (Farmarcas, Farmácias Associadas, Agafarma…):
- cada associada é uma empresa;
- a base leva **uma linha para a rede** (CNPJ da associação, quando existir) e as associadas nas suas próprias linhas,
  com a coluna "Rede/associação";
- a Receita só identifica a associada que leva a bandeira no nome, então manter ao lado o nº de lojas da pesquisa;
- a rede-mãe (ex.: Farmarcas) herda a compra mais recente das suas bandeiras.
- **Potencial (caixas) numa base-universo:** a linha da associação/franquia e a da rede-mãe ficam com 0 loja,
  porque toda associada já é uma linha, com ou sem a bandeira no nome. A associação com CNPJ próprio fica só com as lojas
  da sua raiz, e a rede com várias raízes conta o nº considerado uma vez só. Conferir a soma de lojas contra o
  universo (~120,7 mil farmácias ativas).
- Premissa de caixas/mês por loja: supermercados A 20, B 15, C 10, D 8; farmácias A 17, B 12, C 7, D 4.

**Aba "SUPERMERCADOS (reuniao)" do MAPA MERCADO** (decisões de 22/10/2026):
- vale o cálculo das colunas AI/AJ. As colunas N–T são um modelo antigo, que o usuário mandou esquecer;
- **uma linha por empresa.** É duplicidade: mesma raiz de CNPJ, mesma posição no ranking (regra da aba) ou bandeira
  da mesma rede. Fica a linha do ranking, com o nº de lojas da empresa e a compra mais recente. Nome igual com CNPJ
  diferente é outra empresa e fica. A aba junta o Ranking ABRAS (sem CNPJ) e uma lista de leads (com CNPJ), e a mesma
  rede aparece nas duas;
- a coluna AO guarda a fonte do nº de lojas e a unificação;
- empresa do ranking sem CNPJ localizado: mínimo de 1 loja, com a nota. Associação (ex.: rede de associados) e
  minimercado autônomo ficam com 0, com a nota;
- a aba "TAB DIN. SUPERMERCADOS final" (do usuário) usa esta aba: atualizar as tabelas dela depois de mexer.

**Aba "FARMACIA VF"** (23/10/2026):
- o cabeçalho está na linha 6, e as linhas 1–4 têm os parâmetros de caixas por porte (W1:X4 = a 17, b 12, c 7, d 4);
- a X é `=SE(W="A";$X$1;…)*H`, mas nas 112 linhas de rede/associação o `*H` é trocado pelo nº de lojas sem contagem
  dupla (decisão do usuário; `resumos/vf_excecoes.ps1`). Total 667.594. Se o usuário arrastar a fórmula de novo, as
  exceções somem e o total vai para 1.092.584: perguntar antes de reaplicar;
- a aba "I.R.T.T.__Por bandeira (FINAL)" tem a F como fórmula de status, a AG com o vendedor do CLIENTES e a AH com a
  UF sede; os resumos dela estão em "I.R.T. T. resumo".

---

## 🛒 Planilha de compras ("Relatorio Revenda atualizado.xlsx", 23/10/2026)

- Fica em `Planejamento comercial_AUTOPEL\2026\COMPRAS\`. Na mesma pasta há um `Relatorio Revenda atualizado....xlsx`
  (com reticências), que o usuário deixa aberto: não confundir os dois.
- Aba "Geral (arq. …)" = tabela `Table1`:
  - **A–X vêm do TOTVS e não se alteram.** As outras colunas são fórmulas;
  - quase todas as ~65 dinâmicas do arquivo (abas "(vf)", PLANO de AÇAO, dinamicas…) usam a cache da Table1.
- Conta do TOTVS: `Comprar (F) = PontoPedido (E) − (Disponível O − VendProt P + Comprado R)`:
  - positivo = falta, negativo = sobra;
  - PontoPedido = 2 meses de MedContr (I); PontoPedido 0 = não ter estoque;
  - a coluna O pode vir com a vírgula perdida (quantidade fracionada): o saldo confiável é E − F.
- Programação do trimestre nas colunas **GPT** (AT/AV/AX qtd, AU/AW/AY R$):
  - cada mês volta o saldo ao PontoPedido e desconta 1 MedContr entre um mês e o seguinte;
  - com PontoPedido 0 só compra no mês vigente, e só o que já foi vendido sem estoque;
  - GROK e deepSEAK são cópias antigas;
  - os cabeçalhos dizem OUTUBRO/NOVEMBRO/DEZEMBRO, mas a conta é mês vigente/próximo/subsequente. **Renomear o
    cabeçalho tira o campo das dinâmicas.**
- Tracking da aba dinamicas (AC:AJ):
  - R$ por mês = `SOMASES` na Table1 (todos os SKUs);
  - Marcas e SKUs = dinâmicas filtradas em Comprar > 0;
  - os dias anteriores são digitados pelo usuário.
- Script: `avulsos/compras/revisa_relatorio_revenda.ps1` (+ `confere_revenda.py`). Sessão: `docs/sessoes/2026-10-23 -
  revisao do relatorio revenda (compras).md`.

---

## 📝 Gravar em planilha do usuário

- **Sempre** fazer backup no scratchpad antes. Confirmar que o arquivo **não está aberto no Excel** (abrir `r+b`).
  Se estiver, pedir para fechar.
- Gravar **só** as colunas pedidas. Depois conferir célula a célula que as outras colunas e abas ficaram iguais.
- Pasta com **tabela dinâmica, gráfico ou segmentação** (ex.: MAPA MERCADO): **não salvar pelo openpyxl**. Gravar
  direto no XML da aba com `scripts cruzamento/avulsos/gravar_coluna_xml.py` e conferir que só a parte daquela aba mudou.
  - **Aba inteira nova** (dezenas de milhares de linhas): trocar o `<sheetData/>` da aba vazia (modelo em
    `avulsos/farmacias/grava_consolidado.py`, com cabeçalho congelado e filtro).
  - O nome da aba pode ter aspas: no `workbook.xml` vem como `&quot;`.
- Ler com openpyxl `read_only` e gravar no mesmo processo: fechar o workbook (`wb.close()`) antes de trocar o arquivo.
  Senão o Windows nega o `os.replace`.
- Várias colunas (texto, número, data) de uma vez: `gravar_celulas` em `avulsos/gravar_coluna_xml.py`. Estilo de
  número novo entra no fim de `cellXfs`; aí `styles.xml` também muda, e é preciso avisar.
- **Criar tabela dinâmica:** pelo próprio Excel (COM, `avulsos/resumos/td_resumos.ps1`), numa cópia, depois de gravar
  os dados pelo XML. Depois conferir todas as outras abas célula a célula e recalcular cada célula das tabelas a partir
  dos dados.
  - Contar clientes por um campo sempre preenchido que **não** seja o campo da coluna (status). Senão o Excel tira o
    campo da coluna.
  - A fonte não pode ter cabeçalho vazio (ex.: coluna D da CONSOLIDADO): começar a faixa depois dela.
  - O Python daqui não tem `pywin32`: usar PowerShell.
  - **Formato de número pelo COM:** o Excel daqui lê `NumberFormat` em português, e `'#,##0'` aparece como `1416,0`.
    Montar com `$xl.International(4)` (milhar) e `International(3)` (decimal) e gravar em `NumberFormatLocal`.
  - Para ver como ficou (não há leitor de PDF aqui): `Range.CopyPicture` → colar num gráfico → `Chart.Export` PNG.
  - Refazer um resumo: `td_resumos.ps1 -Limpar` (apaga as tabelas e o conteúdo da aba antes).
- **Excluir linhas** de aba usada por tabela dinâmica: pelo Excel (`avulsos/resumos/ajusta_reuniao.ps1`), de baixo para
  cima. Depois conferir linha velha → linha nova (`confere_reuniao.py`). Para copiar formato, usar `Range.Copy(destino)`:
  `CutCopyMode` falha no PowerShell.
- Datas no formato `DD/MM/YYYY`. Copiar o estilo do cabeçalho vizinho.
- Se o nome da coluna que o usuário citou não bater com o pedido (ex.: pediu a data na "E", que se chama "Status"),
  perguntar antes de gravar.
- O usuário pode editar a planilha entre um pedido e outro (inserir coluna, apagar valores). Reler o arquivo atual
  antes de cada gravação; nunca regravar a partir de uma cópia antiga.
