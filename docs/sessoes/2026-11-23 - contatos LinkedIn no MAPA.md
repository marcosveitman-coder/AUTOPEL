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
