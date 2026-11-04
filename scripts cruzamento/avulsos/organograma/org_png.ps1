# Imagem (PNG) de cada folha da aba "ORGANOGRAMA impressao", para conferir sem imprimir (nao ha leitor de PDF aqui).
# Uso: powershell -File org_png.ps1 -Arquivo <copia.xlsx> -Layout org_impressao.json -Saida <pasta>
# CopyPicture -> grafico -> Export so funciona com o Excel visivel (oculto, a imagem sai em branco).
param([string]$Arquivo, [string]$Layout, [string]$Saida)
$dados = Get-Content -Raw -Encoding UTF8 $Layout | ConvertFrom-Json
$lim = @(1) + @($dados.quebras) + @([int]$dados.ultima + 1)
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($Arquivo, 0, $true); $ws = $wb.Worksheets.Item($dados.aba); $ws.Activate()
  $colFim = $ws.Cells.Item(1, [int]$dados.colunas.ultima).Address($false, $false) -replace '1$', ''
  for ($i = 0; $i -lt $lim.Count - 1; $i++) {
    $rg = $ws.Range("A$($lim[$i]):$colFim$($lim[$i+1]-1)")
    $rg.CopyPicture(1, -4147) | Out-Null
    $ch = $ws.ChartObjects().Add(0, 0, $rg.Width, $rg.Height); $ch.Activate() | Out-Null
    $ch.Chart.Paste() | Out-Null; $ch.Chart.Export("$Saida\folha_$($i+1).png") | Out-Null; $ch.Delete() | Out-Null
  }
  $wb.Close($false)
} finally { $xl.Quit() }
