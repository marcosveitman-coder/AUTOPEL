# 25/11/2026 — Relatórios do TOTVS → CSV (pelo Chrome)

## Pedido
Entrar no TOTVS pelo Chrome (autopel.com → ERP – Autopel), gerar 6 relatórios no "Gera Dados Excel" (005, 198,
199, 200, 210, 232; 198 e 232 com a data "até" = D-1 = 24/11/2026), salvar cada `C:\Excel\listagem.xlsx` como CSV
UTF-8 por cima dos arquivos da pasta de relatórios TOTVS e informar o tempo de cada relatório.

## Como foi feito
- A extensão Claude in Chrome demorou a conectar (o usuário reinstalou/entrou de novo). O login no TOTVS foi feito
  pelo usuário.
- Script novo `scripts cruzamento/avulsos/totvs_relatorios/salva_listagem_csv.ps1` + LEIA-ME na mesma pasta.

## Tempos (do OK até a listagem gravada; ±10 s)
| Relatório | Início | Listagem pronta | Tempo | CSV |
|---|---|---|---|---|
| 005 | 11:13:05 | 11:14:09 | ~1 min 05 s | 27,0 MB |
| 198 | 11:16:11 | 11:21:56 | ~5 min 45 s | 113,4 MB |
| 199 | 11:23:56 | 11:30:46 | ~6 min 50 s | 80,2 MB |
| 200 | 11:32:30 | 11:33:44 | ~1 min 15 s | 27,9 MB |
| 210 | 11:34:54 | 11:36:26 | ~1 min 30 s | 52,9 MB |
| 232 | 11:38:07 | 11:43:31 | ~5 min 25 s | 118,6 MB |

Total de geração ≈ 22 min; com navegação e conversão, 11:13 → 11:45.

## Lições
- O WebAgent abre a listagem no Excel; com ela aberta o TOTVS recusa gerar a próxima ("arquivo em uso").
  O script fecha a janela no fim.
- Os CSVs existentes usam ";" (Excel pt-BR), não vírgula, apesar do nome do formato. Mantido ";".
- Campo de data do webapp: clicar, `ctrl+a`, digitar `dd/mm/aaaa`.
