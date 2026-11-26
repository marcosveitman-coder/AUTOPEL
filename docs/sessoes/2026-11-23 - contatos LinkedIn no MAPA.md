# 23/11/2026 — contatos do LinkedIn dentro do MAPA MERCADO V2

## Pedido
- **Onde gravar:** na aba Alessandra Dayan VF, que fica filtrada como está, preencher as colunas que o usuário montou em
  P7:AD8. São 5 cargos (Gerente de Compras, Coordenador de Compras, Diretor de Compras, Diretor de Supply e CEO), cada um com
  Nome, LinkedIn e Cargo no LinkedIn.
- **Sem arquivo à parte e sem aba nova.**
- **Outras abas:** se a bandeira também está na ESTACIONAMENTO FINAL ou na POSTOS GAS & CONVENIENCIA FINAL, pôr os contatos na
  linha da empresa lá, com as mesmas colunas.

## Entrega
- **Alessandra Dayan VF:** 600 células em P:AD, nas linhas das bandeiras (EMPRESA igual à da planilha de contatos de 17/11).
  O filtro continua igual (mesma condição e 526 linhas ocultas).
- **ESTACIONAMENTO FINAL** (9 redes) **e POSTOS GAS & CONVENIENCIA FINAL** (8 redes): o cabeçalho da VF foi copiado para
  AB5:AP6 com formato e larguras, e os contatos ocupam 72 células.
- **Conferência:** as 672 células estão certas e nenhuma outra célula do arquivo mudou. Backup no scratchpad.

## Regras
- A fonte é a aba CONTATOS da planilha de 17/11. Sócios (QSA) e "outros nomes" não entram, porque não há coluna para eles.
- O cargo "NÃO conferido no perfil" vai como está: são os perfis que a conferência de 17/11 não alcançou.
- **Várias linhas da VF caindo na mesma linha FINAL** (ex.: Indigo, PB e Gepark): vale a linha com o mesmo nome.
- **Ficam fora** os postos de bandeira de supermercado (Postos Extra, Postos Carrefour): os contatos são das compras do varejo.
- Scripts: `mapa_final/contatos_mapa.py` → `mapa_final/grava_contatos.ps1` (cópia, Excel, célula por célula).

## Pendente
- Conferir no perfil os nomes "NÃO conferido" quando houver crédito do Firecrawl e regravar.

## Lição
- Arquivo de texto gravado em modo texto no Windows com '\r\n' vira '\r\r\n'. Gravar com `newline=''`.

## 2º pedido (23/11): lista completa da Alessandra Dayan VF pela aba "Pessoas" do LinkedIn — EM ANDAMENTO
- Pedido: preencher P:AD das 550 linhas da VF (coluna B: rede própria e associação/franquia) com as pessoas dos cargos pedidos,
  buscando como o usuário mostrou (página da empresa no LinkedIn → Pessoas → cargo).
- Método: busca de pessoas do LinkedIn com o filtro de empresa, no navegador embutido, com a conta do usuário:
  `search/results/people/?company=<NOME>&keywords=CEO OR presidente OR "diretor geral" OR compras OR suprimentos OR supply`.
  - No máximo ~6 termos com OR (8 dá "Nenhum resultado"). O parâmetro `title` com OR vira frase e falha. `/company/<slug>/people/?keywords=` é ignorado.
  - Extrator no localStorage do linkedin.com: guarda nome, título, cargo atual, local e perfil só dos cartões com cargo-alvo.
  - A CSP do LinkedIn bloqueia `eval` depois de um `await`: a função de espera transforma o extrator em função antes de esperar.
- Estado na parada (limite de uso do Claude): 180 de 539 buscas feitas, sem bloqueio do LinkedIn. Nada foi gravado no MAPA.
  O passo a passo para retomar está no ESTADO.md da pasta local `mapa_final/dados/contatos_2311/` (fora do git).
- Falta: as 359 buscas restantes, uma 2ª busca com outros nomes para quem não teve resultado, a classificação estrita dos cargos
  (empresa igual; na dúvida, em branco), a conferência no perfil, a gravação na VF e nas abas ESTACIONAMENTO/POSTOS FINAL (cópia +
  conferência célula a célula) e a planilha de conferência.

### Continuação em 24/11 — busca parada pelo limite mensal do LinkedIn
- O navegador embutido reabriu deslogado e o localStorage veio vazio: as 180 buscas de 23/11 se perderam (não havia cópia em disco).
  Desde então, cada lote é gravado em `mapa_final/dados/contatos_2311/brutos/bNNN.json` (local, fora do git), já filtrado
  pela empresa da busca.
- Refeitas e salvas as buscas 1–59; as 60–71 vieram parciais.
- Lotes encadeados de 24 buscas (~4 s cada) derrubaram a sessão (2ª queda). No ritmo lento (7–14 s), a conta bateu o
  **limite mensal de pesquisas** da conta gratuita ("Você atingiu seu limite mensal para pesquisas"): a busca passa a mostrar
  só 3 cartões. O limite volta em 01/12/2026.
- Scripts: `mapa_final/linkedin_lote.py` (monta o lote), `linkedin_nomes.py` (nomes trocados na URL), `linkedin_funcoes.txt`
  (funções do navegador e cuidados).
- Falta: a partir de 01/12, refazer as buscas 60–71 e fazer as 72–539. Depois, a classificação estrita, a conferência no perfil e a
  gravação na VF e nas abas ESTACIONAMENTO/POSTOS FINAL. Nada foi gravado no MAPA.

### Gravação parcial em 24/11 (pedido do usuário: "Write now")
- Das 71 buscas feitas, as pessoas foram conferidas à mão (decisões locais em `dados/contatos_2311/dec_2411.py`). Só entrou
  quem tem o cargo pedido atual na própria rede ou franqueadora. Ficaram fora: franqueado de unidade (vários "CEO" de franquias),
  ex-CEO, VP sem o cargo pedido, CFO e conselho, empresa homônima, cargo só no título sem a empresa, e rede com dois "CEO".
- Gravado com `contatos_linkedin_2411.py` → `grava_contatos.ps1` (cópia, Excel, célula por célula, filtro da VF intacto):
  - 216 células: 61 blocos de cargo na Alessandra Dayan VF (41 linhas) e 11 na POSTOS GAS & CONVENIENCIA FINAL (6 redes);
  - 7 blocos que já tinham contato de 23/11 ficaram como estavam.
- Conferência: só as 216 células mudaram em todo o arquivo. O MAPA estava fechado e igual à cópia de partida ao ser trocado.
- Falta: as buscas 60–71 (parciais) e 72–539, a partir de 01/12, quando volta o limite mensal do LinkedIn.

### Continuação em 25/11 — buscas 60–297 e 2ª gravação
- O limite liberou em 25/11 (antes do dia 1º). Feitas as buscas 60 a 297 em ritmo lento, com uma fila guardada na página: cada
  chamada extrai a página atual e abre a próxima busca depois de 7–14 s.
- A partir da busca 281, os resultados voltaram com 0–3 pessoas e a página passou a oferecer o Premium: sinal de que o limite
  está chegando. As buscas pararam ali. Faltam refazer as 281–297 e fazer as 298–539.
- Candidatos conferidos à mão a cada lote, com a mesma regra estrita de 24/11. Lista local em `dados/contatos_2311/brutos/c*.tsv`.
- Gravado com `contatos_linkedin_2411.py` (decisões em `dados/contatos_2311/dec_2511.py`) → `grava_contatos.ps1`:
  - 330 células: 104 blocos de cargo na Alessandra Dayan VF e 6 na POSTOS GAS & CONVENIENCIA FINAL;
  - 1 bloco que já tinha contato ficou como estava.
- Conferência: só as 330 células mudaram. O MAPA estava fechado e igual à cópia de partida ao ser trocado.

### 25/11 (noite) — filtro da VF devolvido e foco nas drogarias do filtro
- O usuário limpou o filtro da VF para ver o resultado. Devolvido com `mapa_final/filtro_vf.ps1` (cópia, Excel):
  STATUS AUTOPEL = INATIVO ou NÃO É CLIENTE e SEGMENTO = DROGARIA, D8:O558, 24 linhas visíveis. Conferido contra o backup
  (mesmo autoFilter, 526 linhas ocultas) e nenhum valor mudou.
- Pedido novo: foco total nessas 24 drogarias. A aba Pessoas da página da empresa lista os funcionários mesmo com o limite de
  pesquisa. A sessão do LinkedIn caiu de novo após 4 buscas de empresas; à espera de novo login.
- Depois do novo login, a primeira busca abriu uma "Verificação de segurança" do LinkedIn (não resolvida por automação). O usuário
  decidiu parar por hoje. Pendente: as 24 drogarias do filtro (busca de empresas + aba Pessoas, 30–60 s por página) e as buscas
  281–539 da lista geral.
- Lições: o LinkedIn derrubou a sessão 3 vezes e pediu verificação com navegação automática seguida; a aba Pessoas da empresa não
  sofre o corte de 3 cartões do limite de pesquisa.

### 25/11 (noite, 2ª parte) — 24 drogarias do filtro pela busca pública
- O usuário criou outra conta no LinkedIn e pediu para continuar por ela, informando a senha no chat. Recusado: não se digita senha e
  não se usa conta extra para contornar as travas do LinkedIn. As buscas seguiram pela busca pública de perfis (Firecrawl,
  `site:linkedin.com/in "<rede>" + cargos`), sem conta, conferindo o cargo atual (data "o momento") no trecho ou no perfil.
- Gravados 6 cargos novos (18 células) na Alessandra Dayan VF: Farmes, Vida Farmácias, Unifarma, Farmagnus, Nissei e Venâncio.
  O filtro (INATIVO/NÃO É CLIENTE + DROGARIA) continua igual. Conferência: só as 18 células mudaram.
- Das 24 linhas do filtro, 14 têm pelo menos um contato e 10 seguem em branco (redes pequenas sem perfil com o cargo pedido:
  Farmácias Associadas, D1000, Capilé, Grupofarma, Nossa Rede, Uai Farma, Afiminas, Cityfarma, Pedpharma/Redebella, Legítima, PMP).
- Pedido "preciso de 100%": as 10 drogarias do filtro sem ninguém com o cargo pedido receberam a pessoa mais próxima achada na busca
  pública (diretor de outra área, dono, gerente, comprador), com o cargo real marcado "[aproximado]" na coluna "Cargo no LinkedIn"
  (decisões locais em `dados/contatos_2311/dec_2511f.py`). 48 células; conferido: só elas mudaram, filtro igual. Agora as 24 linhas
  do filtro têm pelo menos um contato, todos com o link do LinkedIn.
- Pedido seguinte: nas drogarias sem ninguém de compras, procurar tudo de compras, supply, suprimentos, estoque e logística. Gravadas
  10 pessoas (30 células, "[aproximado]" no cargo) em Unifarma, Vida, Drogamais, Maxifarma, Cityfarma, Legítima e PMP
  (`dados/contatos_2311/dec_2511g.py`). Sem ninguém dessas áreas no LinkedIn: Augefarma, Farmagnus, Capilé, Afiminas, Uai Farma.
  Conferido: só as 30 células mudaram, filtro igual.
- Pedido de buscar as pendentes pela conta nova: recusado de novo; combinado fazer as 118 linhas das redes grandes (buscas 425–539)
  pela busca pública. Parou na 1ª consulta: acabaram os créditos diários do Firecrawl (renovam em ~11 h). Nada gravado nessa etapa.
