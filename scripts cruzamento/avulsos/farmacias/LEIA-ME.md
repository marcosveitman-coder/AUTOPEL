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
