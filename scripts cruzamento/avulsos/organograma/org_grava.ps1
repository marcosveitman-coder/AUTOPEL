# Grava a aba ORGANOGRAMA pelo Excel (COM) numa copia do MAPA MERCADO, a partir do org_layout.json.
# Uso: powershell -File org_grava.ps1 -Arquivo <copia.xlsx> -Layout <org_layout.json>
param([string]$Arquivo, [string]$Layout)
$ErrorActionPreference = 'Stop'
$dados = Get-Content -Raw -Encoding UTF8 $Layout | ConvertFrom-Json
function Cor($rgb) { return [int]$rgb[0] + 256 * [int]$rgb[1] + 65536 * [int]$rgb[2] }
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
  $wb = $xl.Workbooks.Open($Arquivo)
  $ws = $wb.Worksheets.Item('ORGANOGRAMA')
  $ws.Cells.Clear() | Out-Null
  $milhar = $xl.International(4); $decimal = $xl.International(3)
  $fmtNum = "#${milhar}##0"; $fmtPct = "0${decimal}0%"
  $ws.Columns.Item(1).ColumnWidth = 2
  for ($c = 2; $c -le 25; $c++) { $ws.Columns.Item($c).ColumnWidth = 9.3 }
  $ws.Columns.Item(26).ColumnWidth = 2
  $alturas = @{}
  foreach ($op in $dados.ops) {
    if ($op.box) {
      $b = $op.box; $rg = $ws.Range($ws.Cells.Item($b[0], $b[1]), $ws.Cells.Item($b[2], $b[3]))
      $rg.BorderAround(1, -4138, 0, (Cor @(31, 58, 95))) | Out-Null; continue
    }
    if ($op.vline) {
      $v = $op.vline; $rg = $ws.Range($ws.Cells.Item($v[1], $v[0]), $ws.Cells.Item($v[2], $v[0]))
      $bd = $rg.Borders.Item(10); $bd.LineStyle = 1; $bd.Weight = -4138; $bd.Color = (Cor @(46, 117, 182)); continue
    }
    if ($op.hline) {
      $h = $op.hline; $rg = $ws.Range($ws.Cells.Item($h[0], $h[1]), $ws.Cells.Item($h[0], $h[2]))
      $bd = $rg.Borders.Item(9); $bd.LineStyle = 1; $bd.Weight = -4138; $bd.Color = (Cor @(46, 117, 182)); continue
    }
    if ($op.rowh) { $k = [int]$op.rowh[0]; if (-not $alturas.ContainsKey($k) -or $alturas[$k] -lt $op.rowh[1]) { $alturas[$k] = [double]$op.rowh[1] }; continue }
    $rg = $ws.Range($ws.Cells.Item($op.r, $op.c), $ws.Cells.Item($op.r, $op.c2))
    if ($op.c2 -gt $op.c) { $rg.Merge() | Out-Null }
    $cel = $ws.Cells.Item($op.r, $op.c)
    if ($op.fmt -eq 'n') { $rg.NumberFormatLocal = $fmtNum; $cel.Value2 = [double]$op.v }
    elseif ($op.fmt -eq 'p') { $rg.NumberFormatLocal = $fmtPct; $cel.Value2 = [double]$op.v }
    else { $rg.NumberFormat = '@'; $cel.Value2 = [string]$op.v }
    if ($op.bold) { $rg.Font.Bold = $true }
    if ($op.italic) { $rg.Font.Italic = $true }
    if ($op.size) { $rg.Font.Size = $op.size }
    if ($op.font) { $rg.Font.Color = (Cor $op.font) }
    if ($op.fill) { $rg.Interior.Color = (Cor $op.fill) }
    if ($op.halign -eq 'center') { $rg.HorizontalAlignment = -4108 } else { $rg.HorizontalAlignment = -4131 }
    $rg.VerticalAlignment = -4108
    if ($op.wrap) { $rg.WrapText = $true }
    if ($op.border) { $rg.Borders.LineStyle = 1; $rg.Borders.Weight = 2; $rg.Borders.Color = (Cor @(191, 191, 191)) }
    if ($op.h) { $k = [int]$op.r; if (-not $alturas.ContainsKey($k) -or $alturas[$k] -lt $op.h) { $alturas[$k] = [double]$op.h } }
  }
  foreach ($k in $alturas.Keys) { $ws.Rows.Item($k).RowHeight = $alturas[$k] }
  $ws.PageSetup.Orientation = 2; $ws.PageSetup.Zoom = $false; $ws.PageSetup.FitToPagesWide = 1; $ws.PageSetup.FitToPagesTall = $false
  $ws.PageSetup.PrintArea = '$A$1:$Z$' + $dados.ultima
  $ws.Activate(); $xl.ScreenUpdating = $true
  $janela = $wb.Windows.Item(1); $janela.DisplayGridlines = $false; $janela.Zoom = 90
  $ws.Range('A1').Select() | Out-Null
  $wb.Save()
  'gravado: ' + $ws.UsedRange.Address(0, 0)
  $wb.Close($false)
} finally { $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
