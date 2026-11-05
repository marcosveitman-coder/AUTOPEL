# Sessão 05/11/2026 — abas PET SHOP  FINAL e ESTACIONAMENTO FINAL (MAPA MERCADO)

## Pedido
O usuário criou as abas **PET SHOP  FINAL** (com dois espaços no nome) e **ESTACIONAMENTO FINAL**. Pediu:
- o mesmo racional e as mesmas colunas das abas ALIMENTACAO, REDE DE LOJAS e CONSTRUÇÃO FINAL: nome fantasia, razão
  social, grupo administrador, nº de lojas e CNPJ ("informação crucial");
- completar as listas pelas associações: IPB, ABINPET, ABLA, ABRASE e ANCLIVEPA (pet); ABRAPARK e Sindepark
  (estacionamento);
- usar o Google e pesquisar por região (Norte, Nordeste, Centro-Oeste, Sudeste e Sul), para não perder redes
  regionais;
- "não podemos errar".

## Entregas (MAPA MERCADO)

| Aba | O que tem |
|---|---|
| PET SHOP  FINAL | 110 redes (A–W): 99 CNPJ confirmado, 7 provável e 4 franquias com CNPJ de uma unidade (franqueadora não achada) |
| ESTACIONAMENTO FINAL | 133 redes: 120 confirmado, 5 provável, 7 não confirmado e 1 não encontrada |
| MAIORES POR REGIÃO | Acrescentados no fim (linhas 581–839) os blocos de pet e de estacionamento: 30 maiores do Brasil e 20 por região |
| R_E_S_U_M_O | Acrescentada a seção 8 no fim (linhas 74–113): passo a passo, níveis do CNPJ, fontes e maiores |

**Redes por região da sede:**
- **Pet:** Sudeste 45, Sul 22, Nordeste 17, Centro-Oeste 14 e Norte 12.
- **Estacionamento:** Sudeste 73, Sul 21, Nordeste 21, Centro-Oeste 14 e Norte 4.

**Colunas:**
- A–T: as mesmas das outras abas FINAL.
- U–W: as que o usuário criou nas abas antigas.
  - U (STATUS RECEITA) vem da Receita.
  - V (PORTE) e W (STATUS AUTOPEL) foram copiadas só das redes pet que já estavam na REDE DE LOJAS FINAL. Nas outras
    linhas ficaram em branco, porque a regra do usuário para essas colunas não está documentada.

## Como foi feito
1. **Origem.**
   - Estacionamentos: tabela H–K da aba de pesquisa de mercado (aba "Alessandra Dayan"), com 50 empresas.
   - Pet:
     - as 6 redes pet que já estavam na REDE DE LOJAS FINAL, com o mesmo nº de lojas e a mesma fonte;
     - a aba IRTT;
     - a Pet Anjo (mesma aba de pesquisa, coluna BK).
2. **Receita por CNAE** (`redes_pet_estac.py`, cerca de 13 min):
   - pet shop 4789-0/04, veterinária 7500-1/00, banho e tosa 9609-2/08, hotel 9609-2/07, farmácia veterinária
     4771-7/04 e estacionamento 5223-1/00;
   - resultado: 644 empresas pet e 1.063 de estacionamento com 3 ou mais unidades ativas.
3. **Busca por nome e por raiz** (`busca_receita.py`, consultas q12 a q16): razão social, matriz e unidades de cada
   candidata.
4. **Imprensa e listas:**
   - Revista PetCenter, "100 Top Pet Shops das maiores cidades" (mar/2025): redes por cidade e UF;
   - IRTT 2026: Petz 265, Cobasi 256, Petlove 30;
   - ABF/Portal do Franchising: Petland 168, Pet Anjo 275, Pop Pet Center 135, Fórmula Animal 92;
   - Estapar, prévia do 2T26: 851 operações;
   - Forbes e IstoÉ Dinheiro: Indigo com 320+ operações (nov/2025).
5. **CNPJ rede a rede**, à mão (`dados/mk_dec_pet_estac.py`): a mesma regra das outras abas — operadora, franqueadora
   ou dona da marca, nunca franqueado.
6. **Complementos regionais da Receita:**
   - pet: 8+ unidades, ou entre as 15 maiores da região com 4+;
   - estacionamento: 15+ garagens, ou entre as 15 maiores da região com 5+.

**Gravação** (`grava_pet_estac.py`):
- as 2 abas novas foram gravadas pelo XML, e nas abas MAIORES POR REGIÃO e R_E_S_U_M_O só se acrescentou no fim;
- conferência com `confere_pet_estac.py`:
  - as outras 40 abas ficaram iguais célula a célula, e a MAIORES e a R_E_S_U_M_O iguais até a última linha antiga;
  - só mudaram as 4 partes dessas abas;
  - dinâmicas 21 = 21, caches 9 = 9, nomes definidos iguais;
  - todo CNPJ tem dígito válido e está ativo na Receita local, sem nome repetido;
- a cópia abriu no Excel (42 abas, 21 dinâmicas); o original foi trocado depois de conferir que estava igual ao backup.

## Decisões e regras
- **Uma linha por grupo:**
  - Indigo = Administradora Geral de Estacionamentos + PB Administradora (antiga PareBem) + Gepark + Center Park,
    todas com garagens de nome INDIGO. As linhas PB Estacionamentos e Gepark da origem entraram na Indigo;
  - a Rio Park da origem é a marca Riopark da Estapar;
  - a Netpark junta NAA e NEE, que a origem trazia à parte;
  - a Hora Park ficou em linha própria, com grupo Estapar (36 das suas garagens usam o nome ESTAPAR).
- **Sede da Indigo:** a matriz da empresa é em Porto Alegre (antiga Moving), por isso a região da sede é Sul. A
  observação avisa que o escritório da Indigo Brasil é em São Paulo.
- **Franquias sem franqueadora com o nome:** Petland, Rede Pop Pet Center, Apaixonados por 4 Patas e DrogaVet ficam com
  o CNPJ de uma unidade, marcado como tal. A franqueadora da DrogaVet está baixada.
- **Fora da aba de pet shop:**
  - agropecuárias, cooperativas e lojas rurais (farmácia veterinária sem nome de pet). Exemplo: uma rede agropecuária
    de 98 lojas;
  - hospitais universitários e franqueados de redes já listadas.
- **Fora da aba de estacionamento:** pátios de veículos apreendidos, guinchos, transportadoras, hotéis e hospitais.
- **Sem CNPJ (origem):**
  - empresa baixada ou inapta: Wellpark, Servi Park e Simopark;
  - nome comum, sem rede: Brasil Park, Center Park e Park Service;
  - sem empresa com o nome: Vezani e Sandrini;
  - Maxipark (imprensa de 2018): nenhuma empresa ativa com o nome.
- **Associações:**
  - IPB, ABINPET (indústria), ANCLIVEPA, ABRAPARK e Sindepark não publicam lista de redes;
  - ABLA e ABRASE não são entidades do setor pet (ABLA é a associação das locadoras de automóveis).

## Pendências
- Colunas V (PORTE) e W (STATUS AUTOPEL) das linhas novas: o usuário precisa dizer como as preencheu nas outras abas,
  ou pedir o cruzamento com o faturamento.
- Confirmar as franqueadoras de Petland, Pop Pet Center, Apaixonados por 4 Patas e DrogaVet (COF ou contato).
- Petz, Cobasi, Pet Camp, American Pet, Pet Fun e Petlove estão também na REDE DE LOJAS FINAL. O usuário decide se saem
  de lá.

## Lições
- **Não regravar abas FINAL antigas:** elas têm colunas do usuário (U–W). Aba nova se grava sozinha. Em aba de resumo,
  só se acrescenta no fim.
- **Corte regional "top N com X+":** ao excluir uma candidata, entra a próxima, que pode não ter dados na busca. Buscar
  de uma vez todas as candidatas acima do corte mínimo (`q16`).
- **CNPJ digitado à mão:** usar só a raiz e deixar a matriz vir da Receita. O build ignora dígitos digitados que não
  existem.
- **Bash com aspas:** texto com apóstrofo dentro de heredoc ou de `python -c` quebra o comando. Gravar o script com a
  ferramenta de escrita de arquivo.
- **BrasilAPI:** pode ficar fora do ar (serviço indisponível ou excesso de pedidos). A cópia local da Receita resolve.
