# Relatórios do TOTVS (Gera Dados Excel) → CSV

Roteiro para atualizar os CSVs da pasta de relatórios TOTVS (`...\RELATORIOS TOTVS\CSV`, `CSV\etc` e `Pedidos emitidos`).

## Caminho no TOTVS (Chrome, com a extensão Claude in Chrome)
1. autopel.com → "ERP – Autopel" (abre o webapp do Protheus).
2. Login: **feito pelo usuário** (o Claude não digita senha nem clica em Entrar com a senha).
3. Faturamento → Atualizações (2) → Gera Excel (1) → Gera Dados Excel → "Confirmar" na tela de parâmetros
   iniciais (data base, grupo 01, filial 05, ambiente 05) → lista "Relatorios" → selecionar → OK.
4. 198 e 232 pedem datas: a data "até" = **D-1**. No campo de data, clicar, `ctrl+a` e digitar `dd/mm/aaaa`
   (clicar três vezes + Home não funciona).

| Relatório | Destino (sobrescreve) |
|---|---|
| 005 - CADASTRO DE PRODUTO COM NCM | `produtos\PROTUDOS TOTVS - SQL.csv` (desde 26/11; antes `CSV\etc\005 - ...RLT.csv`) |
| 198 - FATURAMENTO EQUIPE (até D-1) | `CSV\FAT - SQL.csv` |
| 199 - ULTIMO FATURAMENTO CLIENTE | `CSV\Ultimo faturamento - SQL.csv` |
| 200 - PEDIDOS EM ABERTO COM STATUS | `CSV\Pedidos abertos - SQL.csv` |
| 210 - CADASTRO DE CLIENTES | `CSV\Clientes - SQL.csv` |
| 232 - CONSULTA DE PEDIDOS EMITIDOS META DE VENDAS (até D-1) | `Pedidos emitidos\Pedidos emitidos - MMAAAA - SQL.csv` |

## Salvar o CSV: `salva_listagem_csv.ps1`
Rodar em segundo plano **antes** de clicar no OK do relatório:

```
salva_listagem_csv.ps1 -Destino "<arquivo.csv>" -Desde (Get-Date)
```

- Espera o `C:\Excel\listagem.xlsx` ficar mais novo que `-Desde`, copia, salva a cópia como CSV UTF-8 e
  sobrescreve o destino.
- **Separador ";"**: o "CSV UTF-8 (delimitado por vírgulas)" do Excel em português grava com ";" (separador de
  lista do Windows). Os CSVs antigos e os scripts que os leem usam ";". O script grava com `Local=True`.
- **O WebAgent abre a listagem no Excel** depois de gerar. Com ela aberta, o próximo relatório dá
  "O arquivo C:\EXCEL\Listagem.XLSx esta em uso, feche o excel!". O script fecha a janela "listagem - Excel"
  (WM_CLOSE) no fim. Esse Excel não aparece para `GetActiveObject`.
- Se outro Excel do usuário estiver aberto, a listagem abre dentro dele: o script fecha só a janela `XLMAIN` com o
  título "listagem" (WM_CLOSE), sem mexer nas outras pastas.
- Tempos: 25/11 ≈ 22 min no total; 26/11 ≈ 2 h 08 min (TOTVS lento; 198 e 232 ~40 min). Com o TOTVS lento, os cliques
  aparecem 15–30 s depois: esperar ~20 s por clique e conferir o campo de data antes do OK.
- Funciona também no navegador embutido do app (barra lateral), depois que o usuário autoriza o WebAgent e faz o login.
  Zoom do Chrome do usuário: 125%.
- Conferir: a 1ª linha de cada CSV traz o nome do relatório (`005 - ...RLT;;;`).
