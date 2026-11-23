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
