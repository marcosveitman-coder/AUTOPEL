# CRM de prospecção TikTok Shop: CNPJ e sócios (27/11/2026)

Preenche as colunas AN–AT da aba **Leads** (CNPJ, nome/LinkedIn do CEO, do diretor e do gerente) a partir do nome da loja.

```
python prep_termos.py <CRM.xlsx> tt_q.json                       # nome da loja -> termos
AUTOPEL_MAPA_FINAL_DADOS=<pasta> python ../mapa_final/busca_receita.py tt_q.json tt_rec.pkl   # ~8 min
python busca_colado.py tt_q.json tt_col.pkl                       # 2a passada, nome sem espacos (~8 min)
python casa_cnpj.py tt_q.json tt_rec.pkl tt_cnpj.json <CRM.xlsx> tt_col.pkl
python desempata.py tt_cnpj.json dados/dec_ambiguos.json tt_cnpj3.json   # ambiguos (opcional)
python qsa.py tt_cnpj3.json tt_qsa.json                            # BrasilAPI (curl), cache
python pessoas.py tt_cnpj.json tt_qsa.json tt_pes.json
python grava_crm.py <CRM.xlsx> tt_pes.json tt_cnpj.json <conferencia.xlsx> <pasta_backup>
```

Regras:
- **Só nome igual** (normalizado; também colado: "lemoiselleintima" = "LE MOISELLE INTIMA"): nome fantasia de
  estabelecimento ativo ou razão social sem LTDA/ME/EIRELI e sem o CPF do MEI.
- O **CNAE** (principal ou secundário) tem de combinar com a Categoria da loja (`ramo.py`). Sem isso, homônimos de outro
  ramo entravam (chocolateria × calça masculina, mercearia × roupa infantil, autopeças × utensílios de cozinha).
- Várias empresas compatíveis → AMBÍGUO, fica em branco (os candidatos vão na planilha de conferência).
- CNPJ baixado/inapto na BrasilAPI não entra.
- CEO = Presidente > Sócio-Administrador > Administrador > sócio único > titular do Empresário Individual (razão sem o
  CPF). Diretor = Diretor; senão o 2º sócio-administrador. Gerente: o QSA não tem.
- **LinkedIn (perfil, 27/11 à tarde):** buscar **pelo nome da loja** na conta do usuário (Claude in Chrome):
  `linkedin_lista.py` → lotes com `linkedin_funcoes.js` (localStorage `__tm`/`__tb`/`__tr`) → `li_passo.py` (grava o lote em
  `dados/li_result.json`) → `grava_crm.py ... dados/li_result.json`. Revisar cada lote à mão. Busca pelo nome do sócio não acha.
  Cerca de 1 loja em 10 tem alguém no LinkedIn. A sessão caiu depois de ~340 buscas.
- **LinkedIn (1ª versão):** dono de loja pequena quase nunca aparece nos buscadores (0 de 6 no teste, WebSearch e Firecrawl). A
  coluna leva um link "Buscar no LinkedIn" (busca de pessoas com o nome), no padrão do "Buscar no TikTok" da planilha.
- O arquivo tem tabela e validação de dados: gravar pelo XML (`gravar_celulas`, que aceita fórmula `{"f","v"}`), nunca
  salvar pelo openpyxl.
- TikTok não abre nesta rede (certificado) e o FastMoss não mostra a empresa do vendedor.

**Desempate dos ambíguos** (`desempata.py`, decisões em `dados/dec_ambiguos.json`, local, tem CNPJs):
- pela web (Instagram, site, Shopee/Mercado Livre): a cidade ou o nome do vendedor que só um candidato tem → "PROVÁVEL (web)";
- todos os candidatos na mesma cidade: o de razão com o nome, senão o mais novo → "PROVÁVEL (mesma cidade)";
- os sites de consulta de CNPJ (cnpj.biz, Serasa, Econodata) listam os mesmos homônimos e não desempatam;
- o Firecrawl gasta 2 créditos por busca e o crédito diário acaba em ~45 buscas; o WebSearch não acha loja brasileira pequena;
- sem crédito: `busca_bing.py` (Bing por script) sugere a cidade, mas acerta pouco (2 de 96). Revisar antes de aceitar.
