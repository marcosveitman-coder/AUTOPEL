# Sessão 22/10/2026 — resumo da reunião de supermercados e fechamento das farmácias

Registro sem dados de clientes (repositório público): só contagens, regras e decisões.

## Pedidos
Tudo no `MAPA MERCADO.xlsx`:
1. **"RESUMO SUPERMERCADO (REUNIAO)"** (aba vazia): diagnóstico da aba "SUPERMERCADOS (reuniao)".
   - Parte 1: o racional das contas das colunas AI e AJ.
   - Parte 2: tabelas dinâmicas de clientes ativos × a trabalhar + caixas (AJ), por UF, por porte (I) e por vendedor (AN).
2. **"FARMACIA CONSOLIDADO"**:
   - rever a última compra usando também o relatório `ultimo fat (20.10.26).xlsx`, incluindo a coluna E "Descrição", que é o **grupo do cliente** no cadastro;
   - trazer o vendedor da carteira (`CLIENTES (20.10.26).xlsx`);
   - criar a coluna PORTE A–D;
   - criar a coluna "Total de caixas/ mês" (A 17, B 12, C 7, D 4 caixas);
   - montar as mesmas tabelas em "FARMACIA CONSOLIDADO RESUMO".
3. **"FARMACIA VF"**: mesma revisão da data, vendedor e "Total de caixas/ mês" (o porte já existia), com as tabelas em "FARMACIA VF RESUMO".

A base da Receita não foi baixada de novo: foi usada a cópia de `C:\dados_receita\2026-10`. A busca na internet foi autorizada.

## Entregas (arquivo, aba, coluna)
| Aba | Colunas gravadas |
|---|---|
| "FARMACIA CONSOLIDADO" | N data, O status, P/Q (só onde a data mudou); novas: **R VENDEDOR, S PORTE, T Total de caixas/ mês** |
| "FARMACIA VF" | O data, P status, Q/R (só onde a data mudou); novas: **V VENDEDOR, W Total de caixas/ mês** |
| RESUMO SUPERMERCADO (REUNIAO) | Parte 1 (texto + 2 tabelas) e Parte 2 (3 tabelas dinâmicas `TD_SUP_*`) |
| "FARMACIA CONSOLIDADO RESUMO" / "FARMACIA VF RESUMO" | "como foram feitas as contas" + tabela por porte + 3 tabelas dinâmicas `TD_CO_*` / `TD_VF_*` |

**Como foi gravado:**
1. As colunas de farmácia foram gravadas pelo XML (`gravar_celulas`), numa cópia do arquivo. Só as partes das duas abas mudaram, e nenhuma célula fora das colunas pedidas mudou.
2. Os resumos foram montados pelo **próprio Excel (COM)**, com o script `resumos/td_resumos.ps1`.
3. Conferência depois do Excel salvar:
   - as outras 26 abas ficaram idênticas, célula a célula;
   - as 6 tabelas dinâmicas antigas foram preservadas (agora são 15);
   - o `calcPr` não mudou;
   - cada célula das 9 tabelas novas foi recalculada direto dos dados, com 0 divergência.

## Números de conferência

### Supermercados (reunião)

**Visão geral**
- 2.427 linhas.
- AJ = 178.301 caixas/mês: 90.606 de ativos (316 linhas) e 87.695 a trabalhar (2.111 linhas).
- Porte A: 99 linhas e 63% das caixas.

**Como a aba calcula**
- AI = premissa por porte (A 20, B 15, C 10, D 8 caixas por loja/mês).
- AJ = AI × L.
- Status AL = AK com até 360 dias. Confere em 100% das linhas.

**Pontos de atenção (estão no resumo)**

1. Linhas com 0 loja na coluna L, e por isso 0 caixa: 448.
   - 264 não têm CNPJ.
   - 169 têm CNPJ, mas nenhuma loja de supermercado ativa na Receita.
   - 15 têm 26 lojas na Receita (+214 caixas).
   - 42 delas são redes A/B. Estimativa: faltam ~18,7 mil caixas, ~10% do total, calculadas pelo faturamento por loja mediano do porte.
2. A mesma raiz de CNPJ aparece em várias linhas: 29 empresas em 59 linhas.
   - 3.704 caixas contadas duas vezes.
   - 2 dessas empresas têm status em conflito entre as linhas.
3. Premissa AI × modelo N–T (consumo calculado), em caixas por loja/mês:

   | Porte | Premissa AI | Modelo N–T (ticket R$ 450) | Modelo N–T (ticket R$ 135) |
   |---|---|---|---|
   | A | 20 | 4,4 | 14,6 |
   | B | 15 | 2,2 | 7,4 |
   | C | 10 | 0,9 | 3,0 |

   - Com o ticket da planilha, a premissa fica de 4 a 11 vezes acima do modelo.
   - O modelo usa ticket de R$ 450 (coluna P). O mercado gira em torno de R$ 125 a 140 (Neogrid 2023 R$ 125,34; Scanntech 2025 +~6%).
4. Coluna L × Receita: batem exatamente em 1.736 de 1.945 empresas com CNPJ. A L é a fonte mais confiável; a M soma só 4.611 lojas, contra 12.439 da L.
5. Outros ajustes de leitura:
   - 78 linhas têm o status AD diferente do AL; as tabelas usam o AL;
   - o rótulo "a tarbalhar" aparece como "a trabalhar" nas tabelas (só a legenda; a aba não foi alterada);
   - 23 empresas da lista também estão nas abas de farmácia (2.677 caixas).

### Farmácias

**Data e status**

| Mudança | Linhas |
|---|---|
| Datas novas pela raiz do CNPJ (notas de 15 a 21/10) | 39 |
| Pelo comprador já identificado | 1 |
| Pelo grupo do cadastro | 12 |
| Por associada da rede | 2 |
| Vínculos por razão social removidos (homônimos) | 24 |

ATIVO passou de 782 para 783 na CONSOLIDADO e de 790 para 791 na VF. Os homônimos removidos tiraram 7 ATIVOs indevidos.

**Vendedor:** 2.027 linhas na CONSOLIDADO e 2.040 na VF. Todas as linhas com compra têm vendedor.

**Porte e caixas**

| Porte | Linhas | Caixas/mês |
|---|---|---|
| A | 38 | 208.216 |
| B | 38 | 13.440 |
| C | 297 | 41.426 |
| D | 90.818 | 404.512 |

- O porte da CONSOLIDADO é igual ao da VF para a mesma empresa (sem par na VF: 2 linhas, pela regra de lojas).
- Total de caixas/mês: VF 667.594 (169.574 de ATIVOS); CONSOLIDADO 667.602.
- Lojas somadas: 120.414, contra ~120,7 mil farmácias ativas na Receita.

## Regras e decisões novas
- **Grupo do cadastro (`Grp.Vendas`/Descrição do "ultimo fat")** → rede:
  - **mapa curado** em `farmacias/final/grupos_redes.json` (local, fora do git, porque tem nomes de grupos), com 35 grupos para 29 redes;
  - critério: os CNPJs dos clientes do grupo caem nas linhas daquela rede **e** o nome bate. Na dúvida, a UF/cidade dos clientes decide;
  - recusados: grupos homônimos de outro estado, grupo só com distribuidora/promotora e grupos sem rede na base;
  - vale só para a **linha-cabeça** (associação, rede-mãe, rede com várias raízes). A linha de uma empresa já se resolve pela raiz;
  - data do cliente do grupo = última venda da raiz dele (bases de faturamento; depois de 14/10, o "ultimo fat").
- **"ultimo fat" só depois do fim das bases de faturamento.**
  - Validação por raiz: 19.254 raízes idênticas às bases e 420 mais novas, todas depois de 14/10.
  - ~54 raízes tinham nota anterior que não aparece nas bases de venda (quase todas no dia 12 do mês). Essas notas não foram usadas.
- **Homônimo por razão social:** se a linha tem CNPJ com dígito verificador válido e achado na Receita (com município), e o comprador é de **outra raiz**, o vínculo cai.
  - Os 7 vínculos cujo CNPJ da aba veio corrompido (Neoway, dígitos deslocados) foram mantidos: o comprador tem o CNPJ verdadeiro.
- **Caixas sem contagem dupla numa base que é o universo da Receita:**
  - linha de associação/franquia = 0 caixa;
  - associação com CNPJ próprio = só as lojas da própria raiz;
  - Farmarcas (rede-mãe) = 0;
  - rede com CNPJ próprio = nº considerado menos as outras raízes da rede que têm linha;
  - motivo: toda associada já é uma linha (com ou sem a bandeira no nome).
  - A primeira versão, que dava à associação as "lojas restantes", somava ~140 mil lojas, acima do universo, e foi descartada.
- **Carteira que não é pessoa:** entrou também "AUTOPEL - VENDA ENCERRADA", além de INATIVOS, CLIENTE SEM COMPRA, RD e 000000.
- Linhas com 0 loja ficam com 0 caixa: CONSOLIDADO 961, VF 968 (inaptas, sem farmácia ativa na raiz ou sem CNPJ válido).

## Scripts
- `scripts cruzamento/avulsos/farmacias/final/farm_final.py`: data, status, vendedor, porte e caixas das duas abas.
  - Sem argumento, só confere; com `--gravar`, grava.
  - `confere_xml.py` confere partes e células.
- `scripts cruzamento/avulsos/resumos/`:
  - `conta_sup.py`: lojas ativas por raiz na Receita, total e só supermercado;
  - `gerar_spec.py`: textos e números dos resumos, todos calculados a partir dos dados;
  - `td_resumos.ps1`: Excel COM, escreve os textos e cria as tabelas dinâmicas;
  - `confere_td.py` e `confere_tudo.py`: conferências.
- `.gitignore` passou a aceitar `*.ps1`.

## Pendências
- Supermercados:
  - preencher a coluna L das 42 redes A/B sem nº de lojas;
  - tirar as 29 empresas repetidas (ou zerar a L das linhas extras);
  - decidir a premissa AI × modelo N–T (ticket de R$ 450).
- Farmácias:
  - associação "a trabalhar" cujas associadas a Receita não identifica pelo nome continua com contagem baixa (pendência de 21/10);
  - o grupo "GRUPO UNIFABRA" (~300 farmácias clientes) não tem rede correspondente na base.

## Lições
- Ao contar com o **mesmo campo** que está na coluna da tabela dinâmica (`AddDataField` do campo de status), o Excel tira o campo da coluna. Conte por outro campo sempre preenchido.
- `PivotItem.Position` falha nesse caso; para pôr "ativo" primeiro, use `AutoSort` decrescente no próprio campo.
- No PowerShell 5.1, passar número vindo do JSON direto para `Range.Value2` dá erro de tipo: converter para `[double]` ou `[string]`.
- Somar o potencial de uma base-universo exige conferir a soma de lojas contra o universo (aqui, ~120,7 mil).
- Validar uma fonte nova contra as bases antes de usar (data por raiz). Aqui isso mostrou as notas fora das bases de venda.

---

## 2ª rodada (tarde): nº de lojas, duplicidades e resumos refeitos

### Pedidos
1. Supermercados: esquecer as colunas N–T (modelo antigo). O cálculo vale pelas colunas AI e AJ.
2. Buscar o nº de lojas das 448 linhas da aba "SUPERMERCADOS (reuniao)" que estavam com L = 0. Busca na internet liberada.
3. As "29 empresas repetidas" eram de supermercados. O usuário pediu para excluir as duplicidades e deixar cada empresa
   uma vez só.
4. Corrigir as três abas de resumo, se preciso.

### Entrega
`MAPA MERCADO.xlsx`. As alterações na aba "SUPERMERCADOS (reuniao)" foram feitas pelo Excel (COM):
- coluna L preenchida;
- nova coluna **AO "Fonte do nº de lojas (L) / linha unificada"**;
- 51 linhas excluídas;
- data, status e vendedor da compra mais recente levados para 2 linhas unificadas.

Depois disso:
- as tabelas dinâmicas da aba "TAB DIN. SUPERMERCADOS final" (do usuário) foram atualizadas. Elas já estavam
  desatualizadas antes: mostravam 2.434 linhas; agora mostram 2.376;
- as três abas de resumo foram limpas e refeitas (`td_resumos.ps1 -Limpar`). O texto de supermercados não fala mais do
  modelo N–T e ganhou a seção "ajustes feitos".

### Nº de lojas das linhas com L = 0
**Regra (a mesma da coluna L, que bate com a Receita em 1.736 de 1.945 empresas):**
- conta os estabelecimentos **ativos** da raiz;
- CNAEs aceitos: supermercado/hiper (4711-3/01, /02), minimercado (4712-1/00), varejo alimentício (4729-6/99) e
  atacarejo (4691-5/00, 4639-7/01, 4639-7/02, 4693-1/00);
- fica fora estabelecimento com nome de CD, depósito, escritório ou manutenção.
- O atacarejo costuma ter CNAE de atacado. Com o filtro antigo, que só aceitava os CNAEs de supermercado, a Tenda dava 0.

**Como achar a raiz de quem não tinha CNPJ (264 linhas):**
1. razão social igual + UF;
2. miolo da razão igual;
3. miolo dentro do nome fantasia, na UF e em estabelecimento de loja;
4. razão com iniciais e sem o CPF do fim (empresário individual).

A raiz só vale se tiver loja ativa na UF. Exemplos de nome que casou com a empresa errada e foi recusado: uma razão
genérica de grupo que casou com uma empresa de serviços de escritório, e a holding sem lojas.

**Internet** (fonte e data na coluna AO), para grupos com várias bandeiras, holding sem lojas ou CNPJ sem
estabelecimento ativo: 9 linhas. A Market4u (2.554 minimercados autônomos) ficou com 0, com a nota.

**Mínimo de 1 loja:** para a empresa com faturamento no Ranking ABRAS cujo CNPJ não foi localizado. Na internet só
apareciam homônimos de outras cidades, e nenhum foi usado.

**Resultado:**

| Situação | Linhas |
|---|---|
| Saíram por duplicidade | 6 |
| Ganharam lojas (1.403 lojas, +21.527 caixas) | 289 |
| — pela Receita | 228 |
| — pela internet | 8 |
| — com o mínimo de 1 loja | 53 |
| Continuam com 0 | 153 |
| — CNPJ inapto na Receita | 129 |
| — CNPJ ativo de outro ramo | 21 |
| — outras (associação Rede Paraíba, Market4u, CNPJ baixado) | 3 |

### Duplicidades (51 linhas excluídas, 45 empresas)
**Critérios:**
- mesma raiz de CNPJ (da aba ou achada na Receita): 12 casos eram a mesma empresa na lista do Ranking ABRAS e na lista
  de leads;
- mesma posição no ranking (pela regra da própria aba, "linhas da mesma empresa dividem a mesma posição");
- bandeiras da mesma rede confirmadas na internet: Plurix (Superpão, Paraná), Mart Minas (Dom Atacadista MG/RJ) e
  Costazul Alimentos = Super Bom/Grupo Barcelos.

**Qual linha fica:**
1. a da rede confirmada;
2. senão, a que já tem lojas;
3. senão, a de maior porte (ranking) e a compra mais recente.

Ela recebe o nº de lojas da empresa e a última compra do grupo.

**Ficaram de fora de propósito:** nomes iguais com **CNPJs diferentes** (ex.: vários "Supermercado Santiago" em RO, cada
um com 1 loja). São empresas distintas, e em 79 de 105 grupos a L de cada linha bate com a Receita daquele CNPJ.

### Números de conferência
- Linhas: 2.427 → 2.376.
- Caixas/mês: 178.301 + 21.527 (lojas preenchidas) − 6.042 (linhas repetidas) = **193.786**. Ativos: 95.986 (304 linhas).
- Aba da reunião conferida célula a célula contra o original (linha velha → linha nova): só mudaram L (289), AJ (289,
  sempre AI × L), AO (472) e AK/AL/AN (2 linhas unificadas). As fórmulas continuam iguais. 0 erro.
- As outras abas ficaram idênticas, menos a "TAB DIN. SUPERMERCADOS final", que foi atualizada. As 15 tabelas
  dinâmicas estão preservadas. Nas 9 tabelas dos resumos, cada célula foi recalculada a partir dos dados: 0 divergência.
- Farmácias: sem mudança de dados. Os resumos foram refeitos com os mesmos números.

### Pontos de atenção (estão no resumo)
- A Rede Supermarket (143 lojas) é uma associação, e associadas com linha própria (ex.: uma com 43 lojas) podem estar
  contadas duas vezes.
- Market4u fora do potencial. Se for contada: +38.310 caixas.
- 26 grupos com nome igual e CNPJs diferentes têm L diferente da Receita. A L não foi mexida.

### Scripts
Pasta `scripts cruzamento/avulsos/resumos/`:
- `receita/receita_l0.py`, `receita/pass2.py`, `receita/pass3.py`: rodadas na Receita;
- `rf_l0.py`, `dups.py`: diagnóstico;
- `consolida_reuniao.py`: decide o nº de lojas, a fonte e as linhas que saem;
- `ajusta_reuniao.ps1`: grava e exclui pelo Excel;
- `confere_reuniao.py`: conferência linha velha → nova.

Ficam locais, fora do git, por terem nomes de redes com números: `web_lojas.json` (números da internet com a fonte) e
`raiz_manual.json` (2 raízes conferidas à mão).

### Lições
- A busca de nome na Receita sem índice fica inviável (60 milhões de razões × 200 nomes). Indexar pela 1ª palavra e
  comparar o miolo igual.
- Ao juntar listas de estabelecimentos de duas rodadas, remover repetição pelo **número do estabelecimento**. Pelo nome
  fantasia, lojas com o mesmo nome viravam uma só (uma rede de 43 lojas deu 3).
- `Range.Copy()` + `CutCopyMode` falha no PowerShell. Usar `Range.Copy(destino)`.
- A lista da reunião junta o Ranking ABRAS (sem CNPJ) com uma lista de leads (com CNPJ). A mesma rede aparece nas duas,
  e a Receita (raiz pela razão social) é o que revela isso.

### Decisão do usuário (fim da sessão)
- Market4u fica **fora do potencial** (L = 0, com a nota na coluna AO).

## 3ª rodada: "FARMACIA VF" com porte em W e caixas por fórmula em X
**Pedido.** Na coluna W, só a letra do porte (da coluna U). Na coluna X, o "Total de caixas/ mês" por fórmula SE, no
mesmo racional da coluna AI de supermercados (A 17, B 12, C 7, D 4). Atualizar o resumo da VF.

**Entrega**, feita pelo Excel com `resumos/vf_porte_caixas.ps1`:
- W1 = "P" e, nas linhas, a letra do porte;
- X1 = "Total de caixas/ mês", com `=SE(W="A";17;SE(W="B";12;SE(W="C";7;SE(W="D";4))))*H`;
- o filtro passou a ir até a X;
- nas **112 linhas** de rede/associação em que o nº de lojas da conta não é o da H (para não contar loja duas vezes:
  101 associações, 6 redes com mais de uma raiz e 5 associações com CNPJ próprio), a fórmula multiplica por esse número
  no lugar da H (ex.: `*0` numa associação). Se fosse `*H` em todas, o total iria a 1.092.584 e contaria lojas em dobro;
- o resumo da VF foi refeito com a fonte `A1:X`.

**Mudança do usuário respeitada.** O cabeçalho P da VF passou a "STATUS AUTOPEL"; a tabela dinâmica usa esse nome.

**Conferência:**
- W bate com a 1ª letra da U em todas as linhas;
- X tem o mesmo valor da W antiga, a soma continua 667.594 e as fórmulas estão como previsto;
- as colunas A:V, as outras abas e os outros dois resumos ficaram idênticos;
- nas 3 tabelas da VF, cada célula foi recalculada a partir dos dados: 0 divergência.

**Lição.** No Excel em português, `NumberFormat = 'General'` falha ("Geral"); é melhor não mexer no formato.
