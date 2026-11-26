# 26/11/2026 — Relatórios do TOTVS → CSV (2ª rodada)

## Pedido
Mesmo trabalho de 25/11 (6 relatórios do Gera Dados Excel → CSV), com duas mudanças do usuário:
- o 005 agora vai para `produtos\PROTUDOS TOTVS - SQL.csv` (antes `CSV\etc\005 - ...RLT.csv`);
- usar os tempos de 25/11 para checar a pasta `C:\Excel` com menos consultas; manter as capturas em 1253 × 698.
198 e 232 com a data "até" = 25/11/2026 (D-1).

## Como foi feito
- Zoom medido: 125% (`devicePixelRatio` 1,25, Windows a 100%). A área de 1253 × 698 de ontem era a janela menor
  (painel do Claude aberto); hoje, maximizada, deu 1530 × 731. O Chrome maximizado não aceita redimensionar.
- O usuário pediu para trabalhar no **navegador embutido (barra lateral)**. Lá o WebAgent também conecta depois que o
  usuário autoriza a abertura; login feito pelo usuário.
- O script esperou cada listagem em segundo plano: as conferências da tela só foram feitas ao fim de cada relatório.

## Tempos (do OK até a listagem gravada; ±20 s)
| Relatório | Início | Pronta | Hoje | 25/11 | CSV |
|---|---|---|---|---|---|
| 005 | 09:13:28 | 09:21:18 | ~7 min 50 s | ~1 min | 27,0 MB |
| 198 (01/11–25/11) | 09:22:56 | 10:02:40 | ~39 min 45 s | ~5 min 45 s | 114,0 MB |
| 199 | 10:18:46 | 10:47:05 | ~28 min 20 s | ~6 min 50 s | 80,2 MB |
| 200 | 10:50:14 | 10:56:58 | ~6 min 45 s | ~1 min 15 s | 24,6 MB |
| 210 | 10:59:04 | 11:08:28 | ~9 min 25 s | ~1 min 30 s | 52,9 MB |
| 232 (01/11–25/11) | 11:12:47 | 11:49:07 | ~36 min 20 s | ~5 min 25 s | 118,8 MB |

Total de geração ≈ 2 h 08 min (25/11: ≈ 22 min). O TOTVS estava lento o dia todo: cada clique levava 15–30 s para
aparecer na tela.

## Problemas e lições
- **Listagem aberta no Excel do usuário:** com outro Excel aberto (Relatorio de Revenda), a listagem abriu dentro dele e o
  script não achava a janela pelo processo. Corrigido: o script procura as janelas `XLMAIN` com o título "listagem" e
  manda WM_CLOSE só nelas (cada pasta tem a própria janela). Testado no 199, 200, 210 e 232.
- **TOTVS lento = cliques atrasados:** o menu abria e fechava, e no 232 a data foi digitada no campo "De" (corrigido antes
  do OK: De 01/11, Até 25/11). Com o TOTVS lento, esperar ~20 s depois de cada clique e conferir o campo antes do OK.
- Depois do 198 o TOTVS voltou para a tela de parâmetros iniciais (já autenticado): só clicar em Entrar.
- A listagem do 198 foi gravada de novo 2 min depois; o CSV refeito saiu do mesmo tamanho.
