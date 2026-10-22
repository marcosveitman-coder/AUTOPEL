# Base consolidada de farmácias (aba "FARMACIA CONSOLIDADO" do MAPA MERCADO)

Pipeline usado em 21/10/2026. Rodar nesta ordem, numa pasta de trabalho (os `.pkl` intermediários ficam nela):

| # | Script | O que faz | Tempo |
|---|---|---|---|
| 0 | baixar `Estabelecimentos0-9.zip`, `Empresas0-9.zip` e `Municipios.zip` | dados abertos do CNPJ da Receita (~6,6 GB), em `%TEMP%\receita_cnpj` | ~10 min |
| 1 | `receita_farm.py` | estabelecimentos de farmácia (CNAE 4771-7) + empresas + associações/cooperativas de farmácia | ~9 min |
| 2 | `redes_classif.py` | classifica cada rede de `redes.py`: **REDE** (CNPJ próprio) ou **ASSOC** (associação/franquia) | 1 min |
| 3 | `busca_assoc.py` e depois `busca_est.py` | busca ampla de razão social/fantasia para achar o CNPJ das associações | ~5 min |
| 4 | `consolida.py` | uma linha por raiz de CNPJ + linhas das associações; valida a aba FARMACIA | 2 min |
| 5 | `cruza_farm.py` | última compra (faturamento 2019–2026) e status | 2 min |
| 6 | `grava_consolidado.py` | grava a aba pelo XML (preserva tabelas dinâmicas) | 1 min |

**Arquivos de configuração:**
- `redes.py`: lista das redes com as palavras da bandeira, a UF de referência e as lojas informadas na pesquisa. Também tem a restrição de UF para nomes genéricos e a rede-mãe (as bandeiras da Farmarcas).
- `assoc_cnpj.json`: CNPJ (raiz) da associação ou administradora de cada rede. É um arquivo **local, fora do git**, porque o repositório é público.

**Variáveis de ambiente:**
- `AUTOPEL_CLIENTES_CSV`: o Clientes - SQL.csv;
- `AUTOPEL_COD_DIR`: os agregados por `COD_CLI`, gerados por `../extract_cod.py`;
- `AUTOPEL_MAPA`: o MAPA MERCADO.xlsx.

**Regras principais** (detalhes no `CLAUDE.md` e em `docs/sessoes/`):
- **Universo:** estabelecimentos **ativos** com CNAE principal 4771-7/01, /02 ou /03. O 4771-7/04 (veterinária) fica fora.
- **Nº de lojas da Receita:**
  - para empresas e redes com CNPJ próprio, são os estabelecimentos ativos da raiz;
  - para associações, são as associadas cuja razão social ou nome fantasia traz a bandeira. Fica abaixo do número real, porque muitas associadas mantêm nome próprio.
- **CNPJ da aba com dígito verificador inválido:** corrigido pela razão social (com UF para desempatar).
- **Cruzamento:**
  - empresa: raiz do CNPJ, ou razão social + UF;
  - rede ou associação: produto personalizado, nome do cliente com núcleo igual à bandeira, domínio de e-mail do cadastro e compras das associadas;
  - rede com CNPJ próprio: só conta comprador do próprio grupo.
- **Status:** ATIVO se comprou nos últimos 360 dias.

## Aba "FARMACIA VF" (pasta `vf/`): consolidada + Ranking de redes 2025

Junta a base consolidada (sem as BAIXADAS que o usuário apagou) com a aba Ranking de
`ranking_redes_farmacias_brasil_2025.xlsx` e grava na aba vazia "FARMACIA VF", rankeada pelo nº de lojas.

| # | Script | O que faz |
|---|---|---|
| 1 | `vf/liga_ranking.py` | liga cada rede do Ranking às redes de `redes.py` (núcleo do nome). Saída: `ranking_lig.pkl` |
| 2 | `vf/soma2025.py` | soma o VLR_TOTAL de 2025 por raiz, COD_CLI e produto (lê `fat 2025*`/`fat 2026*`, ~5 min) |
| 3 | `vf/busca_vf.py` | busca na Receita (precisa dos zips) a razão/fantasia das redes novas do Ranking |
| 4 | `vf/vf_build.py` | monta a base: dados do Ranking, linhas das redes novas, última compra, valor 2025, porte. Saída: `vf.pkl` |
| 5 | `vf/grava_vf.py` | grava a aba "FARMACIA VF" pelo XML (só essa parte do arquivo muda) |

- `vf/redes_vf.py`: bandeiras e UF das redes do Ranking que não estavam em `redes.py`, mais as correções de ligação.
- `assoc_cnpj_vf.json` (local, fora do git): CNPJ das associações novas.
- Variável extra: `AUTOPEL_FARM_CACHE` = pasta dos `.pkl` gerados pela base consolidada (`cruz.pkl`, `rf_*.pkl`, `municipios.pkl`).
- **Nº de lojas considerado:** o do Ranking quando existe. Senão, para rede ou associação, o maior entre o da pesquisa e o
  da Receita; para uma farmácia, o da Receita.
- **Faturamento da farmácia 2025:** o estimado do Ranking (R$). **Valor faturado Autopel 2025:** soma do VLR_TOTAL de
  2025 da raiz. Para uma associação, soma as raízes das compradoras com a bandeira, sem as associadas que já têm linha própria.
- **Rede com mais de uma raiz** (ex.: grupo com duas empresas): a 1ª leva os dados do Ranking e as outras ficam
  marcadas na observação, com as lojas da Receita, para não contar a rede duas vezes.

## Fechamento das duas abas (pasta `final/`, 22/10/2026)

`final/farm_final.py` revê a última compra e o status, e grava as colunas novas nas duas abas:
- CONSOLIDADO: VENDEDOR, PORTE e Total de caixas/ mês;
- VF: VENDEDOR e Total de caixas/ mês.

Uso: sem argumento, só confere; com `--gravar`, grava pelo XML.

**Fontes novas:**
- relatório `ultimo fat` (`AUTOPEL_ULTFAT`): DT_FAT só depois do fim das bases e o grupo do cliente (coluna Descrição);
- `CLIENTES` (`AUTOPEL_CLIENTES_XLSX`): carteira;
- `grupos_redes.json` (local, fora do git): mapa grupo → rede.

**Regras:**
- Homônimo por razão social: CNPJ válido na Receita e comprador de outra raiz → o vínculo cai.
- Caixas = caixas por loja do porte (A 17, B 12, C 7, D 4) × lojas, sem contar loja duas vezes:
  - associação/franquia e Farmarcas ficam com 0 (ou com as lojas da própria raiz);
  - rede com várias raízes conta o nº considerado uma vez.
- `final/confere_xml.py`: confere que só as duas abas e as colunas pedidas mudaram.
