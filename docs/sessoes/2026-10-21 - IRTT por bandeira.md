# Sessão 21/10/2026 (fim da tarde) — IRTT por bandeira: data, CNPJ, faturamento 2025 e lojas

## Pedidos
Na aba "I.R.T.T.__Por bandeira (FINAL)" do MAPA MERCADO (300 grupos, 570 bandeiras):
1. revisar a coluna G ("Última compra") pelas bases de faturamento (descrição do produto) e pelo e-mail do cadastro;
2. validar se o CNPJ da coluna U ("CNPJ indicado no IRTT") é da bandeira, e buscar quando não for;
3. trazer numa coluna nova o faturamento 2025 da bandeira;
4. trazer numa coluna nova o nº de lojas de cada bandeira.

Decisões do usuário:
- faturamento em **duas colunas**: o de mercado da bandeira e o valor faturado pela Autopel em 2025;
- CNPJ correto numa **coluna nova**, sem mexer na U.

No meio da sessão, o usuário pediu para guardar a base da Receita no git, para não baixá-la toda vez. Não foi possível:
os arquivos têm até 2,3 GB, o limite do GitHub é 100 MB e o repositório é público. A alternativa aceita foi uma cópia
local permanente em `C:\dados_receita\2026-10\`, registrada no CLAUDE.md.

## Entrega (MAPA MERCADO, aba "I.R.T.T.__Por bandeira (FINAL)", `sheet14.xml`)
- **G** revisada: 20 células mudaram (17 estavam vazias e 3 trocaram de data). As outras 550 linhas ficaram iguais.
- Colunas novas:

  | Coluna | Conteúdo |
  |---|---|
  | Y | CNPJ VALIDADO |
  | Z | VALIDAÇÃO DO CNPJ |
  | AA | FATURAMENTO 2025 DA BANDEIRA (R$) |
  | AB | FONTE DO FATURAMENTO |
  | AC | VALOR FATURADO AUTOPEL 2025 (R$) |
  | AD | Nº DE LOJAS DA BANDEIRA |
  | AE | FONTE DO Nº DE LOJAS |

- **Gravação** pelo XML com `gravar_celulas`, função nova de `avulsos/gravar_coluna_xml.py`.
  - Mudaram só `sheet14.xml` e `styles.xml`: este último ganhou um estilo de número `#,##0` no fim de `cellXfs`.
  - Conferido: nenhuma célula de A:X mudou fora da G, e as outras abas ficaram idênticas.

## Números de conferência
**Passo 2 – CNPJ:**

| Situação | Bandeiras |
|---|---|
| Confere | 324 (308 automáticas + 16 na revisão manual) |
| Corrigido (era de outra empresa ou a raiz não existe) | 19 |
| Encontrado (U vazia ou PENDENTE) | 150 |
| CNPJ do grupo (bandeira sem CNPJ próprio) | 30 |
| Não encontrado | 47 |

16 ficaram marcados "(conferir)".

**Passo 1 – última compra:** 334 bandeiras com data. Correções que se destacam:
- **São João:** a data antiga, de 28/05/2025, vinha de outra farmácia com o mesmo nome. A data certa é 24/07/2024, a
  mesma do cruzamento de farmácias.
- **Guanabara:** a linha do Guanabara RJ estava com o CNPJ do Guanabara RS.

**Passo 3 – faturamento de mercado:**

| Fonte | Bandeiras |
|---|---|
| Ranking ABRAS 2026 (2025) | 112 |
| Internet 2025 | 17 |
| Ranking de farmácias (2025) | 14 |
| IRTT 2024 (o IRTT 2025 ainda não foi publicado) | 167 |
| Valor incluído na linha do grupo | 238 |
| Sem valor | 22 |

**Valor Autopel 2025:** 123 bandeiras, cerca de R$ 288 milhões, sem contar duas vezes a mesma raiz.

**Passo 4 – lojas:** 342 bandeiras.
- 186 vêm do IRTT 2024, 7 de fontes 2025 e 149 da Receita.
- As 228 sem número são bandeiras de grupos com várias bandeiras, em que a Receita não separa as lojas, e lojas on-line.

## Regras usadas
- **CNPJ confere:** quando a raiz tem lojas ativas com o nome da bandeira no nome fantasia, ou quando é empresa do grupo
  (razão social com o nome do grupo ou operadora conhecida).
- **Candidatos:** raízes com estabelecimentos ativos que trazem o nome da bandeira, com CNAE do segmento.
- **Revisão manual:** 210 bandeiras, em `decisoes_irtt.json` (local, tem CNPJ).
- **Bandeira sem CNPJ:** recebe o CNPJ de outra bandeira do mesmo grupo, exceto em rede associativa (Plurix), onde cada
  bandeira é uma empresa.
- **Faturamento de mercado:** a ordem de preferência é internet 2025 > ranking de farmácias > ABRAS 2026 > IRTT 2024.
  - Valores do IRTT abaixo de R$ 50 mi são dado não divulgado (Amazon, IMC, São João…).
  - Valores do IRTT vêm com o ponto de milhar lido como decimal: 120,6 é R$ 120.600 mi; 1,007 lojas são 1.007 lojas.
- **Grupo com várias bandeiras:**
  - só divide o faturamento pelas lojas se as lojas por bandeira cobrirem 60% do total do grupo;
  - senão, o valor do grupo vai na bandeira com mais lojas, e as outras dizem "incluído…";
  - Carrefour: divisão das vendas brutas do 1T25 (Cash & Carry, Varejo e Clube).
- **Lojas iguais entre bandeiras irmãs** (nome em comum na Receita, como MATEUS): a contagem fica só na bandeira cujo
  nome está contido nas outras ("Riachuelo" ⊂ "Casa Riachuelo"); se nenhuma, fica em branco.

## Pendências
- 47 bandeiras sem CNPJ e 16 marcadas "(conferir)" na coluna Z.
- Faturamento 2025 por bandeira dos grandes grupos (GPA, Mateus, Boticário, Magalu…) não é público. Hoje o valor do
  grupo fica numa linha só.
- Quando o IRTT 2025 sair, trocar os 167 valores de 2024.

## Lições
- `detalha.py` testa a raiz direto no texto da linha (`raw[1:9]`) antes de quebrar os campos. Isso reduz a leitura da
  Receita para cerca de 3 min.
- `receita_irtt.py` com chaves genéricas (sobrenomes, MOVEIS, NACIONAL…) traz milhões de razões sociais. Funciona, mas o
  `.pkl` passa de 800 MB.
- Na revisão, conferir a UF: o IRTT tem grupos com o mesmo nome em UFs diferentes (Guanabara RJ × RS).
