# Sessão 19–20/10/2026 — cruzamentos de faturamento, mercado e carteira

Primeira sessão registrada. Começou em 19/10 e terminou em 20/10 com a criação do `CLAUDE.md` e da golden rule nº 1.
Este registro **não tem dados de clientes**, porque o repositório é público: só contagens, regras e decisões.

---

## 1. Supermercados mapeados × faturamento 2018–2026

**Pedido:** saber, para cada supermercado de `supermercados.xlsx` (2.445 redes), se já é cliente, a data da última
venda e quem vendeu.

**Entrega:** `supermercados - cruzamento faturamento.xlsx`, uma cópia com 4 colunas no fim:
- AA: última venda;
- AB: vendedor;
- AC: critério;
- AD: cliente encontrado.

O original não foi alterado.

**Resultado:** 599 das 2.445 redes com venda.

| Critério | Redes |
|---|---|
| CNPJ raiz | 437 |
| Razão social | 68 |
| Produto personalizado | 52 |
| Nome fantasia + UF | 42 |

**Conferência com o CRM:** em 457 de 478 redes a data bate com até 1 mês de diferença, e em 21 o faturamento é mais
recente. Não sobrou nenhum caso em que o CRM tivesse uma venda mais nova que a nossa.

**Decisões do usuário:**
- não usar o `Ultimo faturamento - SQL.csv`;
- incluir colunas de conferência;
- a busca por descrição do produto é obrigatória (bobina personalizada).

**Como as regras evoluíram:**
1. A busca por substring gerou centenas de falsos positivos: sobrenomes, palavras comuns e drogaria × supermercado
   com o mesmo nome.
2. Passou a exigir núcleo igual, mesma UF e segmento.
3. Um fraco só supera um forte se o comprador tiver o nome da rede.
4. Nome comparado também sem espaços (marca colada).
5. "&" passou a valer como "E".
6. Empresa identificada → última venda da raiz do CNPJ dela.

**Scripts:** `scripts cruzamento/` (`extract.py`, `match.py`, `build.py`, `rodar_cruzamento.bat`, `LEIA-ME.txt`). O `.bat`
foi testado de ponta a ponta: 13 minutos e resultado idêntico, célula a célula.

## 2. Lista "maiores por segmento" do CRM (12.268 linhas)

- **Colunas K e L** ("Quantidade de lojas" e "Faturamento 2025"):
  - Fonte: dados de mercado do mapa de supermercados (Ranking ABRAS).
  - 1ª versão: CNPJ + nome. O usuário pediu **só CNPJ raiz**, com os casos por nome numa aba de **validação manual**.
  - Depois de validado, tudo voltou para uma aba só.
  - Linhas com dois valores confirmados: **maior valor**, decisão do usuário.
- **Coluna G** ("Filiais"):
  - O usuário pediu fonte externa padronizada: **dados abertos do CNPJ da Receita Federal** (`Estabelecimentos0–9.zip`,
    out/2026, 5,3 GB).
  - Contagem de estabelecimentos **ativos, incluindo a matriz**, por raiz. As 12.260 raízes foram encontradas.
  - Arquivos baixados e apagados no fim.
- **Faturamento ampliado:**
  - Ranking ABRAS 2026 completo (1.058 empresas) e ABRAS 2025 (1.017 empresas, faturamento de 2024), extraídos do PDF
    com `pypdf`.
  - 193 vínculos por nome foram para validação, e o usuário mandou aplicar todos.
  - Final: 959 linhas com faturamento (870 ABRAS 2026, 84 ABRAS 2025, 5 outras fontes), com a coluna M "Fonte do faturamento".
  - Faturamento só existe publicamente para cerca de 8% das raízes. O usuário recusou usar o porte da Receita como faixa.

## 3. "final final.xlsx" — última compra na coluna AH (faturamento 2019–2026)

- 581 de 2.445 com data. As datas batem com o CRM em 443 de 464.
- **Ajuste pedido pelo usuário:** marca colada × separada ("SUPER X" = "SUPERX"). Entrou no `match.py` (`glued_keys`).
  A forma colada **soma** ao vínculo normal, não substitui. Quando substituía, outros vínculos se perdiam.

## 4. "ABRAS 2026 - CRU.xlsx" — última compra na coluna M

- Razão social + UF da lista, mais o **domínio de e-mail** do cadastro (exige mesma UF e segmento supermercadista).
- Depois o usuário pediu para usar também a coluna K ("Cliente no CRM").
- Final: 378 de 1.058. As datas batem com o CRM (até 1 mês) em 270 de 302.

## 5. "IRTT (claude).xlsx" — 570 bandeiras de vários segmentos, data na coluna F

- **Decisão do usuário:** data na F ("Última compra"); a coluna E ("Status") fica vazia.
- **Vínculos:**
  - códigos de cliente da coluna R: foi preciso uma 2ª leitura das bases para tirar a última venda por `COD_CLI`;
  - CNPJ e razão indicados;
  - bandeira no nome, no produto e no e-mail.
- **Faixas de marca:**
  - exclusiva (até 5 empresas no cadastro): vale em qualquer lugar;
  - comum ou franquia: exige UF listada no IRTT e a frase completa da bandeira.
- **Código de cliente descartado** só quando a empresa não tem relação de nome com bandeira/grupo/razão **e** está em
  outro segmento. O segmento do faturamento sozinho derrubava vínculos certos.
- **Resultado:** 317 de 570 com data. Conferência com "Dias sem comprar" (referência 13/10/2026): 251 de 288
  batem exatamente.
- **Divergências explicadas:**
  - lista de códigos cortada com "…";
  - compras mais novas achadas por CNPJ ou produto;
  - bandeiras cuja última compra foi em 2018, fora do período.
- Duas bandeiras ficaram sinalizadas ao usuário para conferência manual.

## 6. "MAPA MERCADO.xlsx" — vendedor (pasta com 23 abas e 12 tabelas dinâmicas)

- **Aba "SUPERMERCADOS (reuniao)", coluna AJ "Vendedor":** 692 de 2.434.
  - Decisão do usuário: **vendedor do cadastro** (dono da conta hoje). Recurso: vendedor da última venda.
  - 688 vieram do cadastro e 4 do faturamento.
  - As 579 datas da coluna AG bateram 100% com o nosso cruzamento.
- **Aba "I.R.T.T.__Por bandeira (FINAL)", coluna W "VENDEDOR CARTEIRA":** 395 de 570.
  - Ordem de busca: cliente da última compra → códigos da coluna S → CNPJ da coluna U.
  - Batem com o "Vendedor principal" em 296 bandeiras e divergem em 62.
- **Técnica:** gravação direta no XML da aba, porque o openpyxl corromperia as tabelas dinâmicas. Só a parte da aba
  mudou. Virou o script reutilizável `scripts cruzamento/avulsos/gravar_coluna_xml.py`, testado contra as duas
  gravações reais.

## 7. Organização do repositório

- Criado o repositório git na pasta `Bases Claude` (ramo `main`) e enviado para o GitHub (`origin`). O `.gitignore` é
  lista de permissão (só código e documentação). Conferido no primeiro commit: nenhuma planilha ou CSV entrou.
- O repositório é **público por decisão do usuário**, tomada depois do aviso de que os scripts mostram as regras de
  cruzamento.
- Criados o `CLAUDE.md` (contexto, regras de cruzamento, cuidados com planilhas, repo público) e a
  **golden rule nº 1**: documentar e versionar ao fim de cada sessão.
- Scripts avulsos da sessão salvos em `scripts cruzamento/avulsos/`, sem caminhos pessoais e sem nomes de clientes.
- O `LEIA-ME.txt` usava clientes reais como exemplo. Trocados por exemplos genéricos. As versões antigas continuam no
  histórico do primeiro commit.

---

## Pendências
- Duas bandeiras do IRTT sinalizadas para conferência manual (vínculo por produto de outra UF e um código de cliente
  compartilhado).
- Regras do sistema interno ainda não reproduzidas aqui: produto com mais de 300 compradores, marca em mais de 100
  produtos (só como trava de nome comum) e revenda que compra produto de duas redes.

## Lições
- O usuário cita colunas por letra, mas confere pelo nome. Se letra e nome não baterem, perguntar.
- O usuário edita as planilhas entre um pedido e outro. Sempre reler o arquivo atual antes de gravar.
- Pasta com tabela dinâmica → gravar via XML. Arquivo aberto no Excel → pedir para fechar.
- Validar sempre contra um dado que já existe (datas do CRM, "dias sem comprar", "Vendedor principal").
