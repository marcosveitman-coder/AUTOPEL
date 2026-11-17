# CNPJ e caixas/mes nas listas da Alessandra Dayan (17/11/2026), pelo Excel (COM), numa COPIA do MAPA MERCADO.
# Nas 11 abas "(TOP 50)" e na "Alessandra Dayan VF", igual a FARMACIA VF:
#  1. insere linhas no topo para o cabecalho ir para a linha 6 (TOP 50: 3 linhas; Alessandra Dayan VF: 5 linhas);
#  2. tabela de caixas por porte nas linhas 1-4 (letra a-d na I, valor na J; na Alessandra Dayan VF tambem a K = supermercado);
#  3. colunas novas no fim da tabela: I "CNPJ" (matriz da marca, texto 00.000.000/0000-00) e J "Total de caixas/ mês"
#     = valor da 1a letra do PORTE (F) x N(LOJAS) (C), a mesma formula das abas FINAL. Tabela do Excel: ListObject.Resize.
# Entrada: tsv (UTF-8) aba, linha de hoje (antes da insercao), EMPRESA, CNPJ formatado, nota (comentario na celula: CNPJ provavel).
# Confere o nome antes de gravar.
# Uso: powershell -File top50_cnpj_caixas.ps1 -Arquivo "<copia.xlsm>" -Tsv "<cnpj.tsv>"
param([string]$Arquivo, [string]$Tsv)
$ErrorActionPreference = 'Stop'
$PARAM_DEMAIS = @(17, 12, 7, 4); $PARAM_SUPER = @(20, 15, 10, 8)   # premissas: farmacia/demais e supermercado
$SUPER = 'SUPERMERCADO / ATACAREJO'
$linhas = @(Get-Content -Encoding UTF8 $Tsv | Where-Object { $_ -ne '' } | ForEach-Object { , ($_ -split "`t") })
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
  $wb = $xl.Workbooks.Open($Arquivo)
  $fmtCx = $wb.Worksheets.Item('"FARMACIA VF"').Range('V7').NumberFormat
  $abas = @($wb.Worksheets | Where-Object { $_.Name -like '*(TOP 50)' -or $_.Name -eq 'Alessandra Dayan VF' })
  foreach ($ws in $abas) {
    $nome = $ws.Name; $vf = ($nome -eq 'Alessandra Dayan VF')
    $antes = $(if ($vf) { 1 } else { 3 })
    if ([string]$ws.Cells.Item($antes, 1).Value2 -ne 'EMPRESA') { throw "cabecalho fora do lugar em $nome" }
    $ins = 6 - $antes
    $null = $ws.Range("1:$ins").EntireRow.Insert()
    $ult = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
    # tabela de caixas por porte (linhas 1-4)
    $val = $(if ($nome -like 'SUPERMERCADO*') { $PARAM_SUPER } else { $PARAM_DEMAIS })
    for ($k = 0; $k -lt 4; $k++) {
      $ws.Cells.Item($k + 1, 9).Value2 = [string]('abcd'[$k]); $ws.Cells.Item($k + 1, 10).Value2 = [double]$val[$k]
      if ($vf) { $ws.Cells.Item($k + 1, 11).Value2 = [double]$PARAM_SUPER[$k] }
    }
    if ($vf) { $ws.Cells.Item(5, 10).Value2 = 'demais'; $ws.Cells.Item(5, 11).Value2 = 'supermercado' }
    # colunas novas
    $ws.Cells.Item(6, 9).Value2 = 'CNPJ'; $ws.Cells.Item(6, 10).Value2 = 'Total de caixas/ mês'
    if ($ws.ListObjects.Count -gt 0) {
      $lo = $ws.ListObjects.Item(1)
      $lo.Resize($ws.Range("A6:J$ult"))
    } else {
      $null = $ws.Range('H6').Copy($ws.Range('I6:J6'))
      $ws.Cells.Item(6, 9).Value2 = 'CNPJ'; $ws.Cells.Item(6, 10).Value2 = 'Total de caixas/ mês'
    }
    $se = { param($c) "SE(ESQUERDA(`$F7;1)=""A"";`$$c`$1;SE(ESQUERDA(`$F7;1)=""B"";`$$c`$2;SE(ESQUERDA(`$F7;1)=""C"";`$$c`$3;SE(ESQUERDA(`$F7;1)=""D"";`$$c`$4;0))))" }
    if ($vf) { $f = "=SE(`$H7=""$SUPER"";" + (& $se 'K') + ';' + (& $se 'J') + ')*N($C7)' }
    else { $f = '=' + (& $se 'J') + '*N($C7)' }
    $rg = $ws.Range("J7:J$ult")
    $rg.FormulaLocal = $f
    $rg.NumberFormat = $fmtCx
    $ws.Range("I7:I$ult").NumberFormat = '@'
    $ws.Columns.Item(9).ColumnWidth = 20; $ws.Columns.Item(10).ColumnWidth = 14
    if ($vf) { $ws.Columns.Item(11).ColumnWidth = 13 }
    # CNPJ da marca (confere o nome da linha)
    $n = 0
    foreach ($l in $linhas) {
      if ($l[0] -ne $nome) { continue }
      $r = [int]$l[1] + $ins
      if ([string]$ws.Cells.Item($r, 1).Value2 -ne $l[2]) { throw "nome diferente em $nome linha $r" }
      $cel = $ws.Cells.Item($r, 9); $cel.Value2 = [string]$l[3]; $n++
      if ($l.Count -gt 4 -and $l[4]) { if ($cel.Comment) { $cel.Comment.Delete() }; $null = $cel.AddComment([string]$l[4]) }   # CNPJ provavel
    }
    "$nome : +$ins linhas, cabecalho na 6, dados 7-$ult, CNPJ em $n linhas"
  }
  $wb.Save(); $wb.Close($false)
} finally { $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
