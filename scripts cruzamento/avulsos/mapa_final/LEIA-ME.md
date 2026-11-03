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
- os nomes de exibição (`nomes.json`).

Os resultados das varreduras da Receita (`r1.pkl`, `r2.pkl`, `r3.pkl`) não foram guardados: são grandes e se refazem em cerca de 7 min cada.

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
8. `python build_final.py`: linhas finais (`final_rows.pkl`). Mostra nome repetido e CNPJ usado em mais de uma rede.
9. `python grava_abas.py "<MAPA>" "<cópia nova>"`:
   - grava só o XML das 3 abas, que precisam estar vazias (`<sheetData/>`);
   - reaproveita os estilos 194 (cabeçalho) e 127 (número).
10. `python confere_zip.py <backup> <cópia nova>`: só podem mudar as 3 abas.
11. Abrir a cópia no Excel em modo leitura, pelo COM, e depois `python confere_final.py "<MAPA>"`. Confere dígito verificador, situação ativa na Receita e nomes repetidos.

## Regras do CNPJ

- O CNPJ da rede é o de quem compra pela rede:
  - a **operadora** das lojas próprias, quando ela tem muitas lojas com o nome;
  - a **franqueadora** ou dona da marca (razão com o nome da rede e CNAE de franquia 7740-3/00, de holding ou de indústria);
  - o CNPJ da origem, quando confere (existe, está ativo e tem o nome da rede).
- **Nunca** o CNPJ de um franqueado.
- Sem regra que feche, o CNPJ fica em branco, como "CNPJ NÃO CONFIRMADO". Se não há nenhum estabelecimento com o nome, a linha fica como "REDE NÃO ENCONTRADA NA RECEITA".
- CNPJ de filial vira a matriz da raiz, com nota.

## Cortes dos complementos da Receita

- **Lojas:** 30 ou mais lojas ativas com CNAE do segmento.
- **Alimentação:** 20 ou mais unidades. Na refeição coletiva, 50 ou mais.
- **Construção:** 9 ou mais lojas.

Ficam fora:

- cooperativas, locadoras de máquinas, telefonia e informática;
- lojas online, free shops e massa falida;
- franqueados de redes que já estão na lista;
- redes que fecharam as lojas, embora a Receita ainda mostre estabelecimentos ativos: Ricardo Eletro, Saraiva, Mappin.
