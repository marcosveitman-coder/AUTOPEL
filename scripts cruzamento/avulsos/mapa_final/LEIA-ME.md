# Abas FINAL do MAPA MERCADO (alimentação, rede de lojas, construção)

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
