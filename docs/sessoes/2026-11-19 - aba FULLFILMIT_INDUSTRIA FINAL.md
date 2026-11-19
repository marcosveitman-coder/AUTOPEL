# 19/11/2026 — aba FULLFILMIT_INDUSTRIA FINAL (MAPA MERCADO V2)

## Pedido
O usuário criou a aba vazia **FULLFILMIT_INDUSTRIA FINAL** no `MAPA MERCADO V2.xlsm` e pediu o estudo das grandes indústrias,
no mesmo padrão das abas FINAL:
- pneus; FMCG; farmacêutica e médica; eletroeletrônicos e linha branca; têxtil, calçados e vestuário;
- automotivo e autopeças; química, plásticos e embalagens; móveis e materiais de construção;
- fabricantes e revendas de óleo;
- fabricantes com selo Procel, Conpet ou ENCE/PBE.

Também pediu o STATUS AUTOPEL cruzado com as bases (e-mail, produto, nome, razão, grupo etc.). A pesquisa na internet foi
liberada.

## Entrega
`COMERCIAL (NOVO)\MAPA MERCADO V2.xlsm`, gravado pelo XML (só 3 partes do zip mudaram; backup no scratchpad da sessão):

- **FULLFILMIT_INDUSTRIA FINAL**: **496 linhas**, uma por empresa.
  - Layout das abas FINAL: tabela a/b/c/d em W1:X4, cabeçalho na linha 6 (A–AA), filtro e painel congelado.
  - Colunas adaptadas: E = nº de unidades (estabelecimentos ativos de todas as raízes da linha), P = fonte do nº,
    Q = unidades com CNAE industrial do segmento, M/N = ranking Brasil e na região pelo porte.
  - X = caixas/mês (fórmula das abas FINAL), Y = última compra, Z = vendedor do cadastro, AA = como foi achado.
- **MAIORES POR REGIÃO**: bloco novo no fim (linhas 841–990), com o top 30 do Brasil e o top 20 por região.
- **R_E_S_U_M_O**: seção "9. INDÚSTRIAS" no fim (linhas 116–151), com o método e as fontes.

### Conferência
| Item | Resultado |
|---|---|
| Linhas | 496 |
| CNPJs | todos com DV válido e "CNPJ CONFIRMADO NA RECEITA" |
| Nomes repetidos | nenhum |
| Excel | abriu a cópia e recalculou: 0 erros |
| Caixas/mês | soma 198.212 |
| Outras abas | XML idêntico; nas 2 abas com bloco novo, o conteúdo antigo ficou igual |

### Números
| Segmento | Linhas | ATIVO | INATIVO | NÃO É CLIENTE |
|---|---:|---:|---:|---:|
| FMCG | 92 | 1 | 20 | 71 |
| Eletroeletrônicos e linha branca | 73 | 2 | 5 | 66 |
| Química, plásticos, papel e embalagens | 65 | 1 | 10 | 54 |
| Automotivo, autopeças e máquinas | 63 | 1 | 9 | 53 |
| Farmacêutica e médica | 59 | 1 | 8 | 50 |
| Têxtil, calçados e vestuário | 44 | 4 | 5 | 35 |
| Móveis e materiais de construção | 44 | 0 | 4 | 40 |
| Óleo e lubrificantes | 42 | 0 | 1 | 41 |
| Pneus | 14 | 0 | 1 | 13 |
| **Total** | **496** | **10** | **63** | **423** |

- **Região da sede:** Sudeste 282, Sul 104, Nordeste 46, Norte 38, Centro-Oeste 26.
- **Porte:** A 342, B 98, C 48, D 8.
  - 154 linhas pela receita da Exame Melhores e Maiores 2026: A ≥ R$ 1 bi, B ≥ 300 mi, C ≥ 20 mi.
  - 342 linhas pelo capital social, com nota (A ≥ 500 mi, B ≥ 100 mi, C ≥ 10 mi).
- **Vínculos do STATUS:** 64 por CNPJ, mais 9 fracos aceitos na revisão à mão (5 nome, 2 e-mail, 1 grupo, 1 produto).
  8 fracos foram rejeitados (homônimos e e-mail genérico).

## Como foi feito (scripts em `scripts cruzamento/avulsos/mapa_final/industria/`)
1. `receita_industria.py`:
   - procura na Receita as raízes industriais com estabelecimento ativo e CNAE do segmento;
   - prefixos: pneu 2211; FMCG 10/11/12/2061-2063/1742; farma 2110/2121-2123/3250; eletro 26/27; têxtil 13/14/15;
     auto 29/3091; química 20/22/17/2312; construção 31/23/1621/1622/2511/2512/2542; óleo 1921/1922/4681-8/05.
2. `candidatas_industria.py`:
   - fica só porte "demais" e capital ≥ R$ 5 mi, sem cooperativa de crédito nem órgão público;
   - cortes: top por capital no Brasil por segmento e top 6 por região com capital ≥ R$ 20 mi; para pneu e óleo,
     todas acima de R$ 20 mi. Resultado: 444 candidatas.
3. `fontes_industria.py`: listas com nome, em 260 entradas.
   - Fontes: Exame Melhores e Maiores 2026 (setores industriais), associadas da ANFAVEA, ANIP e ELETROS, fabricantes
     de lubrificantes e marcas com selo Procel, Conpet ou ENCE/PBE.
4. `build_industria.py`:
   - junta tudo com as decisões à mão de `../dados/dec_industria.py` (local, fora do git, porque tem raízes de CNPJ);
   - aplica exclusões (mineração, saneamento, construtora, trading, holding, varejo, distribuidora de combustível,
     usina) e grupos numa linha só;
   - corrige homônimos das fontes, pesquisados um a um na web e na Receita;
   - acrescenta 19 líderes de mercado de capital fechado que não estão na Exame.
5. `status_industria.py`: roda o motor de 16/11 (`status_autopel/motor.py` + `avalia.py`).
   - As linhas entram como família nova **IND** ao lado das entidades das outras abas, então valem as travas contra
     raiz de outra rede.
   - Revisão dos fracos em `revisao_status.txt`; a decisão fica em `ind_status_manual.pkl` (local).
6. `grava_industria.py <pasta> <MAPA> <saída>`: grava as 3 abas pelo XML.

Mudanças no motor:
- `motor.py`: COMPAT da família IND (segmentos do TOTVS compatíveis com indústria).
- `coleta.py`: a aba nova entra nas FINAIS, com a família IND.

## Lições
- A Exame Melhores e Maiores 2026 vem dentro do HTML de `exame.com/mm/maiores` (JSON do Next.js, 997 empresas, receita em
  R$ mil). O corte mínimo do ranking é ~R$ 582 mi.
- O capital social da cópia local da Receita vem com vírgula decimal: sem trocar por ponto, tudo vira porte D.
- Nome de fonte casa com homônimo: um biscoito com o nome de uma fundição, uma montadora espanhola, adubo com o nome de um
  fabricante de eletroportáteis. Conferir a raiz de cada fonte.
- O `busca_receita` não acha razão de 1 palavra com menos de 6 letras: buscar o CNPJ na web.
- Multinacional de capital fechado (higiene, limpeza, alimentos) não aparece na Exame: complementar à mão.

## Pendente
- Contatos LinkedIn (17/11): conferir os perfis 64 a 220 quando houver crédito do Firecrawl e refazer a planilha.

---

# 2ª parte (19/11/2026): revisão do STATUS AUTOPEL em todas as abas + ampliação da indústria

## Pedido
O usuário mostrou um erro e pediu dois trabalhos:
- **o erro:** a linha "Koni Store" (ALIMENTACAO FINAL 131) estava NÃO É CLIENTE, mas a marca tem pedido programado e compra
  (bobina "80X100M KONI" do cliente do grupo da marca);
- **revisão:** ter certeza do STATUS AUTOPEL em todas as abas, sem conflitos internos;
- **ampliação:** a aba de indústrias ficou pequena; focar em associações, cooperativas e empresas regionais.

## Causa do caso Koni
A mesma marca estava em duas linhas com o mesmo CNPJ: "KONI" na IRTT, ATIVO, e "KONI STORE" na ALIMENTACAO FINAL, NÃO É CLIENTE.
Como o prefixo tem só 4 letras, as duas linhas não foram ligadas, e o produto "KONI" não bate com o nome "KONI STORE".

## O que a auditoria achou (e foi corrigido no motor)
Ferramenta nova: `auditoria.py`, que procura conflitos de status na mesma raiz ou no mesmo nome, mais as pistas fracas não aceitas.
- **Mesma marca com dois status:** Koni, Pizzaria Spoleto, Melissa × Clube Melissa.
  - Regra nova: mesmo CNPJ + prefixo de 4+ letras, ou só uma palavra de formato a mais.
- **CNPJ de franqueado separava a rede da mesma rede nas listas da Alessandra**: We Pink, Santo Remédio e outras.
  - Regra nova: o CNPJ de uma unidade (coluna H, ou a G que avisa "unidade") vale só como vínculo.
- **Linha de grupo emprestando nome de outra marca:** o Frango Assado estava ATIVO por uma bobina da Pizza Hut. Agora fica ATIVO
  pela bobina "FRANGO ASSADO POSTOS" da própria IMC.
- **A aba nova de indústrias criou nomes ambíguos** (NATURA), e a linha NATURA + AVON perdeu o vínculo pela UF. A IND não cria
  mais ambiguidade para as outras abas.
- **Vínculos errados ou que faltavam:**
  - código de cooperativa de crédito na lista da Coop Supermercado (IRTT), excluído;
  - produtos personalizados e clientes conferidos à mão: Tendall, Calvin Klein, Natura, Apoio Mineiro, Rei do Mate, Unissul,
    Mega Descontão e Petrox.
- **CNPJs das listas da Alessandra / TOP 50, gravados em 17/11, conferidos na Receita** (468 raízes):
  - AZZAS 2154 estava com o CNPJ do Magazine Luiza. Foi corrigido, e o status vinha das compras do Magalu;
  - o CNPJ do St. Marché na aba SUPERMERCADOS (reuniao) é do Spani Atacadista, e o St. Marché fica INATIVO (Hortus, cadastro sem
    nota). Essa aba é do usuário e não foi alterada: avisado.
- **Base do dia:** `atualiza_clientes.py` usa os CSVs do TOTVS de 19/11. Foram 4.888 datas de última nota mais novas que as do 16/11.
- **Abas editadas pelo usuário desde 17/11:**
  - REDE DE LOJAS FINAL tem dinâmicas embaixo dos dados (linha 334) e a Alessandra Dayan VF também (linha 565);
  - a VF ganhou 2 colunas na frente e o cabeçalho foi para a linha 8;
  - a coleta e a gravação agora leem pelo cabeçalho e param no fim da tabela.

Fica como estava (regra de 16/11): a marca não herda a compra de outra marca do mesmo grupo. Esses casos estão na aba CONFLITOS
da planilha de conferência, para o usuário confirmar.

## Resultado da revisão do STATUS
- **29 linhas mudam de status** (de 98 mil avaliadas), em 266 células.
  - 17 sobem para ATIVO (marca igual em duas abas, produto personalizado da própria marca, associada, franqueado);
  - 4 descem de ATIVO para INATIVO (vínculo de cooperativa de crédito, CNPJ de outra empresa e uma farmácia pela data do dia);
  - 8 passam de NÃO É CLIENTE para INATIVO.
- **Planilha de conferência:** `COMERCIAL (NOVO)\CONFERENCIA STATUS AUTOPEL 19.11.2026.xlsx` (LEIA, MUDANÇAS, CONFLITOS).
- **CNPJ da AZZAS 2154** corrigido na Alessandra Dayan VF (linha 260) e na REDE LOJAS (TOP 50) (linha 8).
- **Conferência do arquivo:** as 266 células planejadas estão certas e nenhuma outra célula mudou. A exceção é a IRTT coluna H,
  "Dias sem comprar" (fórmula HOJE()−G, que acompanha as datas novas). Macro e as 52 dinâmicas mantidas.

## Resultado da ampliação da indústria
- **De 496 para 1.037 linhas:**

  | Origem das linhas novas | Linhas |
  |---|---:|
  | Regionais (3 maiores por UF e segmento) | 378 |
  | Cooperativas | 92 |
  | Associadas da ABIA, ABRAFATI e ABRAMAT | 71 |

- **Por segmento:**

  | Segmento | Linhas |
  |---|---:|
  | FMCG | 276 |
  | Química, plásticos e embalagens | 168 |
  | Eletro | 119 |
  | Construção | 115 |
  | Automotivo | 99 |
  | Têxtil | 98 |
  | Farma | 89 |
  | Óleo | 56 |
  | Pneus | 17 |

- **Por região:** Sudeste 439, Sul 232, Nordeste 186, Norte 90, Centro-Oeste 90.
- **STATUS:** 18 ATIVO, 86 INATIVO, 933 NÃO É CLIENTE, calculado junto com as outras abas (um cluster por empresa).
- **Exclusões das listas novas** (`FORA_AMP`): construtoras, holdings, importadoras, algodoeiras, usinas, projetos sem operação,
  empresas em liquidação/falidas e SCP.
- **MAIORES POR REGIÃO e R_E_S_U_M_O:** os blocos da indústria foram regravados no mesmo lugar, sem duplicar. O Excel passa os
  textos para a tabela compartilhada ao salvar, e o `grava_industria.py` procura o título nos dois lugares.

## Pendente
- ~~Aprovação do usuário~~: aprovado; o MAPA foi trocado às 14:48 (backup do MAPA de 12:05 no scratchpad da sessão).
- Os resumos com texto fixo e as dinâmicas não foram atualizados: o usuário atualiza ao abrir.
- Contatos LinkedIn (perfis 64 a 220) quando houver crédito do Firecrawl.

## Lições
- Antes de gravar CNPJ numa lista, conferir na Receita a razão e a fantasia contra o nome da linha. O erro da AZZAS (CNPJ do
  Magalu) passou em 17/11 e dava ATIVO pelas compras do Magalu.
- Aba nova numa família de nomes comuns cria ambiguidade para as abas antigas: rodar a auditoria depois de toda aba nova.
- O usuário edita o MAPA durante a sessão (colunas novas na VF, dinâmicas embaixo dos dados): sempre recoletar do arquivo atual.
- No PowerShell, `Range.Value2 = $null` dá erro de referência nula: usar `ClearContents()`.
