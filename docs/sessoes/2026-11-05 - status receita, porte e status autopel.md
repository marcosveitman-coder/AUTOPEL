# 05/11/2026 – STATUS RECEITA, PORTE e STATUS AUTOPEL nas abas FINAL, na Alessandra Dayan e no organograma

## Pedido

O usuário pediu três colunas em todas as abas criadas na semana:
- **PORTE**;
- **STATUS RECEITA** (ATIVA, BAIXADA, INAPTA);
- **STATUS AUTOPEL**: ATIVO se a última compra foi há menos de 360 dias; senão, INATIVO.

Depois, as mesmas informações na aba **Alessandra Dayan**, para todas as tabelas de todos os segmentos. Por fim, atualizar
a aba **ORGANOGRAMA impressao**.

Antes de começar, o usuário perguntou:
1. quais abas;
2. se a pesquisa do STATUS AUTOPEL incluiria o cadastro (nome do grupo, e-mail), o "ultimo fat" e o faturamento deste
   ano (descrição de produto).

Também perguntou se a base da Receita estava salva. Resposta: está em `C:\dados_receita\2026-10`, fora do git.

Decisões do usuário:
- seguir com os relatórios que já existem;
- liberou pesquisa na internet;
- não precisava subir o MAPA: o arquivo é atualizado direto na pasta dele.

## Entregas (arquivo `COMERCIAL (NOVO)\MAPA MERCADO.xlsx`)

| Aba | O que mudou |
|---|---|
| ALIMENTACAO, REDE DE LOJAS, CONSTRUÇÃO, PET SHOP, ESTACIONAMENTO, OUTROS FINAL | U (STATUS RECEITA) e W (STATUS AUTOPEL) regravadas; V (PORTE) só onde estava vazia |
| POSTOS GAS & CONVENIENCIA FINAL | colunas U–W criadas (cabeçalho e largura iguais às outras abas FINAL) |
| PET SHOP, POSTOS e OUTROS FINAL | coluna ORIGEM (S): só o endereço "Alessandra Dayan: XX" foi atualizado (191 células), porque as colunas da Alessandra andaram |
| Alessandra Dayan | STATUS RECEITA, PORTE e STATUS AUTOPEL ao lado de cada uma das 11 tabelas; 2 colunas inseridas depois de SUPERMERCADOS, CONVENIÊNCIA, BELEZA e ÓTICAS |
| ORGANOGRAMA impressao | potencial recalculado; entram POSTOS E CONVENIENCIAS, ESTACIONAMENTOS e PET SHOP; agora 7 folhas |

Conferência para o usuário: `COMERCIAL (NOVO)\conferencia STATUS AUTOPEL (05.11.26).xlsx`. Para cada linha, a planilha
traz:
- a data da última compra e os dias sem comprar;
- o vínculo (como foi achado), o cliente e o CNPJ dele;
- a fonte da data e o vendedor da carteira;
- o STATUS AUTOPEL anterior;
- a origem e o faturamento usados no PORTE.

A aba "resumo" traz as regras. A planilha fica fora do git, porque tem nomes e CNPJs.

A aba **ORGANOGRAMA** (não a de impressão) **não foi regravada**. O layout dela é anterior ao do `org_layout.py` (resumo
e legenda em outra posição), e regravar mudaria a aba inteira. Os números de potencial dela foram atualizados depois, só
nas células, na 2ª parte, a pedido do usuário. Na 1ª parte, chamei Z71 de "anotação do usuário" por engano: é o
cabeçalho POTENCIAL da tabela de carteira.

## Regras usadas

**STATUS AUTOPEL** (data de referência 05/11/2026):
- **ATIVO**: última compra há menos de 360 dias.
- **INATIVO**: comprou há 360 dias ou mais, **ou** está cadastrado no TOTVS (tem código de cliente) sem nota. As 143
  redes que o usuário tinha marcado como INATIVO caíram aqui; 72 delas são cadastradas sem nota.
- **NÃO É CLIENTE**: não achada no cadastro nem no faturamento.
- "LEAD NO CRM" e "INATIVANDO" saíram, porque vinham do CRM e não da compra. O valor anterior ficou na planilha de conferência.

**Datas:**
- bases de faturamento 2019–2026 (cache agg, até 14/10/2026);
- depois disso, o `Ultimo faturamento - SQL.csv` de 05/11 (pasta de relatórios TOTVS, cópia de hoje);
- nota do Ultimo faturamento anterior ao fim das bases só vale para raiz sem venda nas bases (regra 8 do CLAUDE.md).

**Vínculos**, por ordem de força (`status_autopel.py`):
1. raiz do CNPJ da rede (F), do CNPJ de referência (H) e das abas de apoio (FARMACIA VF, IRTT, reuniao);
2. razão social igual;
3. grupo do cadastro (coluna Descrição do Ultimo faturamento) com o mesmo núcleo da rede;
4. vínculos fracos pelo motor do `../irtt.py`: nome da rede no cliente, no fim da descrição do produto e no domínio do
   e-mail.

Ajustes nos vínculos depois da revisão manual (todos viraram regra no script):
- parênteses sem "+" são descrição e não viram nome de busca. Antes, "(postos)" casava com qualquer "REDE POSTOS" e
  "(estacionamentos)" com o domínio @estacionamentos;
- o "grupo administrador" só é nome de busca quando o grupo tem uma rede só. Antes, a compra da empresa-mãe deixava
  ATIVAS as outras marcas do grupo;
- o núcleo do grupo do cadastro não tira palavra de produto: o grupo "X CALÇADOS" não é a rede "LOJAS X";
- produto personalizado ou e-mail de comprador que revende (distribuidor, logística, suprimentos, "produtos para…")
  não vale, a menos que o comprador tenha o nome da rede. Isso pegou 21 casos;
- comprador de vínculo fraco disputado por 2 redes só vale para a rede da mesma UF;
- CNPJ validado de outra linha nunca vale. Na mesma rede em duas linhas, a outra linha ainda leva a data pela aba de apoio.

**PORTE** (faixas de supermercados e farmácias: A ≥ R$ 1 bi; B R$ 300 mi a 1 bi; C R$ 20 a 300 mi; D < R$ 20 mi):
1. o valor que o usuário já tinha gravado fica (147 linhas). Em 63 delas o cálculo daria outra faixa, quase sempre a
   vizinha;
2. faturamento publicado: IRTT por bandeira (AA), pelo mesmo CNPJ ou pelo mesmo nome. O "valor do GRUPO (todas as
   bandeiras)" só vale para a linha que é o grupo inteiro. Na Alessandra, o PORTE da FARMACIA VF ou da reuniao;
3. estimativa, nº de lojas × faturamento médio por loja:

   | Tipo | R$ por loja/ano | Fonte |
   |---|---|---|
   | franquia (alimentação, beleza, serviços, óticas, pet, conveniência) | 1,49 mi | ABF 2025: R$ 301,7 bi ÷ 202.444 unidades |
   | rede de lojas e construção | 7,42 mi | IRTT: 1º quartil do faturamento por loja de 96 bandeiras de varejo não alimentar |
   | posto de combustível | 15 mi | ANP 2025: ~108 bi de litros nos postos × ~R$ 6,2 ÷ ~43 mil postos (estimativa) |
   | estacionamento | 2,26 mi | Estapar 2025: R$ 1,87 bi ÷ 827 operações |
   | minimercado autônomo | 0,13 mi | Market4U 2025: R$ 336 mi ÷ ~2.600 unidades |
   | drogaria | 2,2 mi | mediana do Ranking de farmácias (regra da FARMACIA VF) |
   | distribuidora de combustível sem nº de postos | ~1,7 bi por empresa | 182 distribuidoras regionais = 44% do mercado (Eixos, 2025) |

   Sem nº de lojas na coluna E, usa a coluna Q (lojas com o nome na Receita). As 14 linhas sem nenhum dos dois ficam
   "SEM Nº DE LOJAS" (8 delas são redes do ranking ANAMACO).

**STATUS RECEITA**:
- situação do CNPJ da coluna F na Receita local (busca nova, `r19.pkl`);
- sem F, a situação da unidade de referência (H), como já estava;
- sem nenhum dos dois, "SEM CNPJ".

## Números de conferência

| Aba | ATIVO | INATIVO | NÃO É CLIENTE |
|---|---|---|---|
| ALIMENTACAO FINAL | 21 | 41 | 306 |
| REDE DE LOJAS FINAL | 43 | 90 | 194 |
| CONSTRUÇÃO FINAL | 13 | 31 | 206 |
| POSTOS GAS & CONVENIENCIA FINAL | 33 | 37 | 140 |
| PET SHOP  FINAL | 8 | 2 | 100 |
| ESTACIONAMENTO FINAL | 26 | 28 | 79 |
| OUTROS FINAL | 4 | 8 | 128 |
| Alessandra (110 que não estão numa aba FINAL) | 63 | 35 | 12 |

As outras 440 empresas da Alessandra copiam os 3 valores da linha da aba FINAL.

Comparação com a coluna W anterior (do usuário):
- ATIVO → ATIVO: 62. Nenhuma rede marcada ATIVO deixou de ser cliente;
- INATIVO → INATIVO: 143; INATIVO → ATIVO: 9 (compras novas, ex.: franqueado no grupo do cadastro);
- LEAD NO CRM → NÃO É CLIENTE: 47;
- NÃO É CLIENTE → ATIVO: 6; NÃO É CLIENTE → INATIVO: 18. Quase todas achadas pelo grupo do cadastro ou pelo e-mail
  de um franqueado.

Vínculo das redes com compra ou cadastro:

| Vínculo | Redes |
|---|---|
| CNPJ | 384 |
| grupo do cadastro | 39 |
| data das abas de apoio | 28 |
| e-mail | 16 |
| produto personalizado | 11 |
| nome do cliente | 5 |

Conferências do arquivo:
- as abas que não foram gravadas estão iguais célula a célula (`org_confere.py` e `confere_zip.py`);
- as 21 dinâmicas e os 9 caches continuam no arquivo. Nos nomes definidos, só mudou a área de impressão da ORGANOGRAMA
  impressao, que passou a ir até a linha 143;
- nas abas FINAL, nenhuma célula fora de U–W mudou, e todos os valores conferem (`conf_uvw`);
- na Alessandra, as 3.938 células estão no lugar esperado depois da inserção;
- na ORGANOGRAMA impressao: 7 páginas, nenhuma quebra vertical, zoom 100%, só Verdana 10 ou maior e sem itálico. As
  imagens das folhas foram conferidas.

## Scripts (em `scripts cruzamento/avulsos/mapa_final/`)

1. `status_coleta.py <MAPA>`: linhas das 7 abas FINAL e da Alessandra, com as ligações. Gera `status_itens.pkl` e
   `q19.json`.
2. `busca_receita.py q19.json r19.pkl`: situação e porte dos 1.535 CNPJs, em cerca de 9 min.
3. `status_autopel.py`: última compra e status. Variáveis: `AUTOPEL_ULTFAT_CSV`, `AUTOPEL_CLIENTES_CSV`, `AUTOPEL_HOJE`.
4. `status_final.py <conferencia.xlsx>`: valores das 3 colunas e a planilha de conferência.
5. `status_grava.py <MAPA> <cópia>`: grava U–W das abas FINAL pelo XML.
6. `status_alessandra.py <cópia>`: atualiza os endereços da ORIGEM e gera `status_alessandra.json`.
7. `status_alessandra.ps1 -Arquivo <cópia>`: o Excel insere as colunas e grava a Alessandra.
8. Organograma (`../organograma`):
   - `org_dados.py`, agora com as 3 abas novas no potencial;
   - `org_impressao.py`, agora com 7 folhas;
   - `org_grava.ps1` com o `org_impressao.json`.

Correção no `org_grava.ps1`: gravar a largura coluna a coluna dava `InvalidCastException` na aba ORGANOGRAMA. Agora a
largura é gravada numa faixa única.

## Pendências

- **OUTROS FINAL** fica fora do potencial do organograma: mistura beleza, óticas e serviços e não corresponde a um
  segmento do CRM.
- Possível duplicidade no ESTACIONAMENTO FINAL: o cadastro da Autopel põe a empresa de uma linha no grupo de outra rede
  da aba. Conferir se as duas linhas são a mesma rede (detalhe na planilha de conferência).
- 1 CNPJ da aba de apoio dos supermercados tem dígito verificador errado e aparece como "NÃO ACHADO NA RECEITA".

## Lições

- "Cadastrado sem nota" é INATIVO, não "não é cliente": era assim que o usuário classificava no CRM.
- O endereço de célula escrito como texto (ORIGEM "Alessandra Dayan: BA7") anda quando se insere coluna. Antes de inserir,
  procurar referências em texto, não só em fórmulas.
- Antes de regravar uma aba gerada por script, comparar com o que o script produziria: ela pode ter outro layout (versão
  antiga ou mudança do usuário). Nesse caso, trocar só as células que mudam.
- O PORTE estimado pelo nº de lojas precisa de média por tipo de negócio: a média da ABF dava A para minimercado autônomo
  e para franquia de locação.

## 2ª parte: potencial da aba ORGANOGRAMA (só as células)

**Pedido:** "sim, atualize": atualizar os números de potencial da ORGANOGRAMA sem regravar a aba.

**Entrega** (`MAPA MERCADO.xlsx`, aba ORGANOGRAMA), 35 células trocadas:
- 9 caixas "Potencial: N oportunidades";
- 9 linhas da coluna Y (OPORTUNID.) da carteira;
- 16 números da tabela "2. POTENCIAL POR SEGMENTO" (Alimentação, Rede de Lojas, Construção);
- a nota "Fonte".

A tabela ganhou 3 linhas depois de Drogarias (Postos e Conveniências, Estacionamentos, Pet Shop), com o formato da
linha de cima e quebra de texto para o nome longo da aba. Tudo abaixo desceu 3 linhas, e a área de impressão foi até a
linha 130.

**Números:** diretoria e supervisor da equipe maior 91.552; equipe menor 91.031. Os números de cada pessoa são os mesmos
da ORGANOGRAMA impressao.

**Conferência:**
- as outras abas estão iguais célula a célula;
- dinâmicas e caches continuam iguais;
- nos nomes definidos, só a área de impressão mudou;
- na ORGANOGRAMA, as 725 células estão no lugar esperado (velha → nova, com as trocas), sem diferença nem célula a mais;
- a imagem da tabela foi conferida.

**Scripts:** `organograma/org_potencial.py` (gera as trocas lendo a aba atual) e `organograma/org_potencial.ps1`
(grava pelo Excel).
