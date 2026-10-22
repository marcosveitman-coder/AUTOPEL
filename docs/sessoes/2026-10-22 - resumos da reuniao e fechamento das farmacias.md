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
