# Grava a aba ORGANOGRAMA pelo Excel (COM) numa copia do MAPA MERCADO, a partir do org_layout.json.
# Uso: powershell -File org_grava.ps1 -Arquivo <copia.xlsx> -Layout <org_layout.json | org_impressao.json>
# O org_impressao.json traz aba, colunas, quebras de pagina e zoom (aba "ORGANOGRAMA impressao", uma folha A4 por parte).
param([string]$Arquivo, [string]$Layout)
$ErrorActionPreference = 'Stop'
$dados = Get-Content -Raw -Encoding UTF8 $Layout | ConvertFrom-Json
function Cor($rgb) { return [int]$rgb[0] + 256 * [int]$rgb[1] + 65536 * [int]$rgb[2] }
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
  $wb = $xl.Workbooks.Open($Arquivo)
  $aba = 'ORGANOGRAMA'; if ($dados.aba) { $aba = $dados.aba }
  $ws = $wb.Worksheets.Item($aba)
  $ws.ResetAllPageBreaks() | Out-Null
  $ws.Cells.Clear() | Out-Null
  $milhar = $xl.International(4); $decimal = $xl.International(3)
  $fmtNum = "#${milhar}##0"; $fmtPct = "0${decimal}0%"
  $cIni = 1; $cFim = 26; $larg = 9.3; $marg = 2
  if ($dados.colunas) { $cIni = [int]$dados.colunas.primeira; $cFim = [int]$dados.colunas.ultima; $larg = [double]$dados.colunas.largura; $marg = [double]$dados.colunas.margem }
  $ws.Columns.Item($cIni).ColumnWidth = $marg
  for ($c = $cIni + 1; $c -lt $cFim; $c++) { $ws.Columns.Item($c).ColumnWidth = $larg }
  $ws.Columns.Item($cFim).ColumnWidth = $marg
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
  $ps = $ws.PageSetup; $ps.Orientation = 2
  if ($dados.quebras) {
    # uma folha A4 paisagem por parte: zoom fixo, quebras manuais, margens estreitas, centralizado, numero da folha no rodape
    $ps.PaperSize = 9; $ps.Zoom = [int]$dados.zoom
    $ps.LeftMargin = $xl.CentimetersToPoints(0.6); $ps.RightMargin = $xl.CentimetersToPoints(0.6)
    $ps.TopMargin = $xl.CentimetersToPoints(0.8); $ps.BottomMargin = $xl.CentimetersToPoints(0.8)
    $ps.HeaderMargin = $xl.CentimetersToPoints(0.3); $ps.FooterMargin = $xl.CentimetersToPoints(0.3)
    $ps.CenterHorizontally = $true; $ps.CenterFooter = '&8Folha &P de &N'
    foreach ($q in $dados.quebras) { $ws.HPageBreaks.Add($ws.Cells.Item([int]$q, 1)) | Out-Null }
  } else { $ps.Zoom = $false; $ps.FitToPagesWide = 1; $ps.FitToPagesTall = $false }
  $colFim = $ws.Cells.Item(1, $cFim).Address($true, $true) -replace '\$1$', ''
  $ps.PrintArea = '$A$1:' + $colFim + '$' + $dados.ultima
  $ws.Activate(); $xl.ScreenUpdating = $true
  $janela = $wb.Windows.Item(1); $janela.DisplayGridlines = $false; $janela.Zoom = 90
  $ws.Range('A1').Select() | Out-Null
  $wb.Save()
  'gravado: ' + $ws.UsedRange.Address(0, 0) + ' | quebras: ' + (@(foreach ($b in $ws.HPageBreaks) { $b.Location.Row }) -join ',')
  $wb.Close($false)
} finally { $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
