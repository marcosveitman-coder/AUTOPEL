# CRM de prospecção TikTok Shop: CNPJ e sócios (27/11/2026)

Preenche as colunas AN–AT da aba **Leads** (CNPJ, nome/LinkedIn do CEO, do diretor e do gerente) a partir do nome da loja.

```
python prep_termos.py <CRM.xlsx> tt_q.json                       # nome da loja -> termos
AUTOPEL_MAPA_FINAL_DADOS=<pasta> python ../mapa_final/busca_receita.py tt_q.json tt_rec.pkl   # ~8 min
python busca_colado.py tt_q.json tt_col.pkl                       # 2a passada, nome sem espacos (~8 min)
python casa_cnpj.py tt_q.json tt_rec.pkl tt_cnpj.json <CRM.xlsx> tt_col.pkl
python qsa.py tt_cnpj.json tt_qsa.json                            # BrasilAPI (curl), cache
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
- **LinkedIn:** dono de loja pequena quase nunca aparece nos buscadores (0 de 6 no teste, WebSearch e Firecrawl). A
  coluna leva um link "Buscar no LinkedIn" (busca de pessoas com o nome), no padrão do "Buscar no TikTok" da planilha.
- O arquivo tem tabela e validação de dados: gravar pelo XML (`gravar_celulas`, que aceita fórmula `{"f","v"}`), nunca
  salvar pelo openpyxl.
- TikTok não abre nesta rede (certificado) e o FastMoss não mostra a empresa do vendedor.
