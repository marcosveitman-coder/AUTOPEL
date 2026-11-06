# Abas FINAL do MAPA MERCADO (alimentação, rede de lojas, construção, postos, pet shop e estacionamento)

Monta as abas **ALIMENTACAO FINAL**, **REDE DE LOJAS FINAL** e **CONSTRUÇÃO FINAL**:

- junta as abas de origem;
- tira as duplicatas;
- confere ou acha o CNPJ de cada rede na cópia local da Receita;
- completa a lista com o ranking ANAMACO, a aba IRTT e as redes regionais achadas na própria Receita.

Sessão: `docs/sessoes/2026-11-03 - abas FINAL do MAPA MERCADO.md`.

## Dados

Tudo o que é dado fica em `dados/`, local e fora do git (o `.gitignore` só deixa entrar código e documentação). Lá estão:

- as decisões rede a rede (`dec_*.json`);
- as regras dos complementos (`comp_rules.json`, `complementos_manuais.json`, `raizes_extra.json`);
- o ranking ANAMACO (`anamaco.json`);
- os nomes de exibição (`nomes.json`);
- da 2ª rodada: as decisões com nível C/P (`dec_5.json`), os CNPJs conferidos na BrasilAPI (`brasilapi.json`) e a
  4ª consulta (`q4.json`).

Os resultados das varreduras da Receita (`r1.pkl` a `r5.pkl`) não foram guardados: são grandes e se refazem em cerca de
7 min cada. Para rodar com os dados em outra pasta, usar a variável `AUTOPEL_MAPA_FINAL_DADOS`.

## Ordem para refazer

1. `python extrai_abas.py "<MAPA MERCADO.xlsx>"`: abas de origem para `src.json` e `src2.json`.
2. `python parse.py`, depois `python keys.py`, depois `python canon.py`:
   - uma linha por rede e por segmento (`canon.json`, `groups.json`);
   - as junções manuais e as linhas de grupo ("LUMP") ficam no `canon.py`.
3. `python busca_receita.py q1.json r1.pkl`: busca nome e CNPJ das redes da origem na Receita.
4. `python redes_cnae.py r2.pkl`: redes por CNAE e região. Conta os estabelecimentos ativos por raiz e os nomes fantasia repetidos em muitos CNPJs (franquias).
5. `python view.py r1.pkl q1.json v1.txt`: visão para decidir à mão o CNPJ de cada rede. As decisões ficam em `dec_*.json`.
6. `python mkq3.py` e depois `python busca_receita.py q3.json r3.pkl`: matrizes das raízes escolhidas, nomes refinados e complementos. Revisar com `python view3.py`.
7. `python mk_comp.py`: lista de complementos (`complementos.json`).
8. 2ª rodada, para as redes ainda sem CNPJ:
   - `python mkq4.py` e depois `python busca_receita.py q4.json r4.pkl`;
   - `python view4.py`, `python view5.py` (visão compacta, `v5.txt`) e `python sede.py`: candidatos a franqueadora;
   - decisões em `dec_5.json`, no formato `[cnpj, grupo, nota, nível]`. O nível é C (confirmado) ou P (provável,
     conferir).
   - redes novas de listas de fora (ex.: Piticas, Multicoisas, Livraria Nobel, da lista ABF):
     - entram em `complementos_manuais.json`, chave `ABF`;
     - a busca na Receita é `busca_receita.py q5.json r5.pkl`;
     - o `build_final.py` usa `q4` e `q5` para a unidade de referência.
   - 3ª rodada (redes ainda sem CNPJ): `python mkq6.py`, `busca_receita.py q6.json r6.pkl`, `python view6.py` (`v6.txt`);
     decisões em `dec_6.json` (`[cnpj, nível C/P/U/NC/NF, nota, UF da sede]`, U = CNPJ de uma unidade).
   - O `grava_abas.py` também regrava a aba **R_E_S_U_M_O** (texto em `resumo.py`) quando ela existe.
9. `python build_final.py`:
   - gera as linhas finais (`final_rows.pkl`) e mostra nome repetido e CNPJ usado em mais de uma rede;
   - monta também o CNPJ de referência e a UF da sede.
10. `python grava_abas.py "<MAPA>" "<cópia nova>"`:
    - regrava só o XML das 3 abas (vazias ou já preenchidas);
    - cria ou atualiza a aba **MAIORES POR REGIÃO** no fim da pasta;
    - reaproveita os estilos 194 (cabeçalho) e 127 (número).
11. `python confere_zip.py <backup> <cópia nova>`: mostra as partes novas e as que mudaram. Só podem mudar:
    - as 3 abas e a aba de região;
    - `workbook.xml`, `workbook.xml.rels`, `[Content_Types].xml` e `docProps/app.xml`.
12. Abrir a cópia no Excel em modo leitura, pelo COM, e depois `python confere_final.py "<cópia nova>"`. Confere:
    - dígito verificador e situação ativa do CNPJ e do CNPJ de referência;
    - nomes repetidos;
    - blocos da aba de região.

## Regras do CNPJ

- O CNPJ da rede é o de quem compra pela rede:
  - a **operadora** das lojas próprias, quando ela tem muitas lojas com o nome;
  - a **franqueadora** ou dona da marca (razão com o nome da rede e CNAE de franquia 7740-3/00, de holding ou de indústria);
  - o CNPJ da origem, quando confere (existe, está ativo e tem o nome da rede).
- **Nunca** o CNPJ de um franqueado.
- Sem regra que feche, o CNPJ fica em branco, como "CNPJ NÃO CONFIRMADO". Se não há nenhum estabelecimento com o nome, a linha fica como "REDE NÃO ENCONTRADA NA RECEITA".
- CNPJ de filial vira a matriz da raiz, com nota.
- Níveis da coluna "CONFERÊNCIA DO CNPJ":
  - **confirmado**;
  - **provável (conferir)**;
  - **franqueadora não achada: ver CNPJ de referência**. O CNPJ de **uma unidade** vai numa coluna separada e não vale
    como CNPJ da rede;
  - **não confirmado**;
  - **rede não encontrada**.
- **Unidade de referência:** só com o nome completo da rede e CNAE do segmento. Nomes genéricos ficam fora
  (`REF_NAO` no `build_final.py`).

## Nº de lojas pelo nome e região
- **Lojas com o nome na Receita:**
  - só o nome completo e só CNAE do segmento;
  - "LOJAS X" só vira "X" se sobrarem 2+ palavras ("LOJAS REZENDE" virava "REZENDE" e pegava 2 mil empresas);
  - acima de 3× o nº de lojas, a coluna fica em branco (homônimos).
- **Região** = UF da sede, nesta ordem: Receita, pesquisa (`SEDE_WEB` no `build_final.py`), aba de origem, UF com mais
  lojas da rede ("sede não confirmada"). Nunca a UF da unidade de referência.

## Cortes dos complementos da Receita

- **Lojas:** 30 ou mais lojas ativas com CNAE do segmento.
- **Alimentação:** 20 ou mais unidades. Na refeição coletiva, 50 ou mais.
- **Construção:** 9 ou mais lojas.

Ficam fora:

- cooperativas, locadoras de máquinas, telefonia e informática;
- lojas online, free shops e massa falida;
- franqueados de redes que já estão na lista;
- redes que fecharam as lojas, embora a Receita ainda mostre estabelecimentos ativos: Ricardo Eletro, Saraiva, Mappin.

## Aba POSTOS GAS & CONVENIENCIA FINAL
1. `python redes_postos.py r8.pkl`: empresas com postos (4731-8/00), conveniência (4729-6/02) e distribuidoras (4681-8) ativas, por raiz e por nome fantasia.
2. `python busca_receita.py q9.json r9.pkl`: razão social e matriz das candidatas.
3. Tabela 3.17 do Anuário da ANP em `dados/anp_dl/tabela3-17.csv` (postos por bandeira, região e UF).
4. `python build_postos.py`:
   - gera `postos_rows.pkl`, usando as decisões de `dec_postos.json` (distribuidoras, conveniência, nomes, grupos, exclusões e cortes);
   - cortes: 15+ postos no Brasil, ou entre as 25 maiores da região com 8+ postos; conveniência com 8+ lojas.
5. `grava_abas.py` grava a aba, os blocos de postos na MAIORES POR REGIÃO, a tabela ANP e a R_E_S_U_M_O.

## Abas PET SHOP  FINAL e ESTACIONAMENTO FINAL (05/11/2026)

O nome da aba de pet tem **dois espaços** ("PET SHOP  FINAL"), como o usuário criou.

**Ordem:**
1. `python redes_pet_estac.py r11.pkl`: varre a Receita por CNAE e conta as unidades ativas por empresa (raiz) e por nome
   fantasia. Leva cerca de 13 min.
   - Pet: pet shop 4789-0/04, veterinária 7500-1/00, banho e tosa 9609-2/08, hotel 9609-2/07, farmácia veterinária
     4771-7/04.
   - Estacionamento: 5223-1/00.
2. `busca_receita.py` com três consultas:
   - `q12.json`: nomes das redes citadas (aba de origem e imprensa, em `dados/pet_estac_fontes.json`);
   - `q13.json`: raízes das candidatas;
   - `q14`/`q15`: raízes que faltaram.

   As saídas são `r12.pkl` a `r15.pkl`.
3. `python view_pet_estac.py`: candidatas por região, em `v_pet.txt` e `v_estac.txt`.
4. `python view_origem_pet_estac.py`: candidatas de cada rede citada, em `v_origem.txt`.
5. Decisões à mão em `dados/mk_dec_pet_estac.py`, que gera `dec_pet_estac.json`:
   - CNPJ e nível de cada rede citada;
   - junções de grupo;
   - exclusões, nomes, notas e cortes.
6. `python build_pet_estac.py`: gera `pet_rows.pkl`, `estac_rows.pkl` e `pet_estac_view.txt`.
7. `python grava_pet_estac.py "<MAPA>" "<cópia>"`:
   - grava as 2 abas (vazias), com as colunas A–W (U–W como nas outras abas FINAL);
   - **acrescenta** blocos no fim da MAIORES POR REGIÃO e da R_E_S_U_M_O;
   - não regrava nenhuma outra aba. As abas FINAL antigas têm colunas U–W do usuário e não podem ser regravadas pelo
     `grava_abas.py`.
8. `python confere_pet_estac.py "<MAPA>" "<cópia>"`:
   - outras abas iguais célula a célula;
   - dígito verificador e situação ativa de todo CNPJ;
   - nomes repetidos;
   - dinâmicas e nomes definidos.

**Regras próprias:**
- **Uma linha por grupo:**
  - Indigo = Administradora Geral de Estacionamentos + PB Administradora (antiga PareBem) + Gepark + Center Park;
  - Estapar inclui a marca Riopark; a Hora Park fica numa linha própria, com grupo Estapar;
  - Netpark = NAA + NEE + NEP.
- **Pet:**
  - ficam fora agropecuárias, cooperativas e lojas rurais (farmácia veterinária sem nome de pet), hospitais
    universitários e franqueados de redes já listadas;
  - corte: 8+ unidades no Brasil, ou entre as 15 maiores da região com 4+.
- **Estacionamento:**
  - ficam fora pátios de veículos apreendidos, guinchos, transportadoras, hotéis e hospitais;
  - corte: 15+ garagens, ou entre as 15 maiores da região com 5+.
- **Redes pet também na REDE DE LOJAS FINAL** (Petz, Cobasi, Pet Camp, American Pet, Pet Fun, Petlove): mesmo nº de
  lojas e mesma fonte. PORTE e STATUS AUTOPEL foram copiados de lá. Nas outras linhas, U = situação na Receita, e V e
  W ficam em branco.

## STATUS RECEITA, PORTE e STATUS AUTOPEL (05/11/2026)

Colunas U–W das 7 abas FINAL e 3 colunas ao lado de cada tabela da aba Alessandra Dayan.
Sessão: `docs/sessoes/2026-11-05 - status receita, porte e status autopel.md`.

**Ordem:**
1. `python status_coleta.py "<MAPA>"`: gera `status_itens.pkl` e `q19.json`.
2. `python busca_receita.py q19.json r19.pkl`: cerca de 9 min.
3. Rodar o `status_autopel.py`:
   - com as variáveis `AUTOPEL_ULTFAT_CSV` e `AUTOPEL_CLIENTES_CSV` apontando para cópias do dia dos CSVs do TOTVS, e
     `AUTOPEL_HOJE=AAAA-MM-DD`;
   - revisar à mão em `status_autopel_view.txt` os vínculos fracos e os status que mudaram.
4. `python status_final.py "<conferencia.xlsx>"`: valores finais e a planilha de conferência para o usuário.
5. `python status_grava.py "<MAPA>" "<cópia>"`: U–W das abas FINAL pelo XML. Na coluna V, só grava onde está vazia.
6. `python status_alessandra.py "<cópia>"` e depois `powershell -File status_alessandra.ps1 -Arquivo "<cópia>"`: o Excel insere as colunas
   que faltam e grava a Alessandra. **Só uma vez:** depois de 05/11, as colunas já existem e os endereços mudaram (ver
   `dados/alessandra_refs_0511.json`). Para refazer, gravar os valores nas colunas que já existem.
7. Conferir:
   - `confere_zip.py`;
   - `../organograma/org_confere.py` (outras abas, dinâmicas e nomes);
   - todas as células U–W contra o `status_final.pkl`.
