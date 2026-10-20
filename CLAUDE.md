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
- O `SEGMENTO` do faturamento é pouco confiável: há rede de lojas classificada como DROGARIAS e distribuidor como
  REVENDA. Não usar como única trava.

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
   - ignorar carteiras que não são pessoas: INATIVOS, CLIENTE SEM COMPRA, RD/REPRESENTANTE, código 000000.
6. Sempre **conferir** contra o que já existe na planilha (datas do CRM, "dias sem comprar", "Vendedor principal") e
   explicar as divergências.

Armadilhas reais já vistas:
- drogaria e supermercado com o mesmo nome;
- sobrenomes (LOPES, ALMEIDA) e palavras comuns (VERDE, MODELO, TORRE) casando com qualquer cliente;
- loja de calçados com o mesmo nome de supermercado;
- código de cliente reaproveitado por outra empresa;
- lista de códigos do IRTT cortada com "…".

---

## 📝 Gravar em planilha do usuário

- **Sempre** fazer backup no scratchpad antes. Confirmar que o arquivo **não está aberto no Excel** (abrir `r+b`).
  Se estiver, pedir para fechar.
- Gravar **só** as colunas pedidas. Depois conferir célula a célula que as outras colunas e abas ficaram iguais.
- Pasta com **tabela dinâmica, gráfico ou segmentação** (ex.: MAPA MERCADO): **não salvar pelo openpyxl**. Gravar
  direto no XML da aba com `scripts cruzamento/avulsos/gravar_coluna_xml.py` e conferir que só a parte daquela aba mudou.
- Datas no formato `DD/MM/YYYY`. Copiar o estilo do cabeçalho vizinho.
- Se o nome da coluna que o usuário citou não bater com o pedido (ex.: pediu a data na "E", que se chama "Status"),
  perguntar antes de gravar.
- O usuário pode editar a planilha entre um pedido e outro (inserir coluna, apagar valores). Reler o arquivo atual
  antes de cada gravação; nunca regravar a partir de uma cópia antiga.
