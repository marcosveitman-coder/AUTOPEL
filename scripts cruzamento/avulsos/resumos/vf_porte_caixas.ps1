# Aba "FARMACIA VF": coluna W = letra do porte (da coluna U); coluna X = Total de caixas/mes por formula SE, no
# racional da coluna AI de supermercados (A 17, B 12, C 7, D 4 caixas por loja) x nº de lojas da coluna H.
# Nas linhas de rede/associacao em que a conta usa outro nº de lojas (sem contar loja duas vezes), a formula leva
# esse numero no lugar de H.  Use numa COPIA; feche o arquivo no Excel.
#   powershell -ExecutionPolicy Bypass -File vf_porte_caixas.ps1 -Xlsx C:\...\copia.xlsx -Spec C:\...\vf_wx.json
param([Parameter(Mandatory = $true)][string]$Xlsx, [Parameter(Mandatory = $true)][string]$Spec)
$ErrorActionPreference = 'Stop'
$S = Get-Content -Raw -Encoding UTF8 -Path $Spec | ConvertFrom-Json
$se = 'IF(W{0}="A",17,IF(W{0}="B",12,IF(W{0}="C",7,IF(W{0}="D",4))))'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
try {
    $wb = $xl.Workbooks.Open($Xlsx, 0, $false)
    $ws = $wb.Worksheets.Item($S.aba)
    $n = [int]$S.ult                                   # ultima linha de dados
    $c = $(if ($S.cab) { [int]$S.cab } else { 1 })      # linha do cabecalho (o usuario pode inserir linhas no topo)
    $p1 = $c + 1
    $tot = [string][char]0x54 + 'otal de caixas/ m' + [char]0xEA + 's'
    if ([string]$ws.Range("U$c").Value2 -ne 'PORTE') { throw "layout inesperado: U$c" }
    if ([string]$ws.Range("W$c").Value2 -ne 'P') {
        # W: letra do porte (da coluna U)
        $u = $ws.Range("U${p1}:U$n").Value2
        $w = New-Object 'object[,]' ($n - $c), 1
        for ($i = 1; $i -le $n - $c; $i++) { $t = [string]$u[$i, 1]; $w[($i - 1), 0] = $(if ($t.Length -gt 0) { $t.Substring(0, 1) } else { $null }) }
        $ws.Range("W${p1}:W$n").Value2 = $w
        $ws.Range("W${p1}:W$n").HorizontalAlignment = -4108
        $ws.Range("W$c").Value2 = 'P'
    }
    if ($ws.Range("X$c").Value2 -ne $null -and [string]$ws.Range("X$c").Value2 -ne $tot) { throw "a coluna X tem outro conteudo: $($ws.Range("X$c").Value2)" }
    # X: formula no racional da coluna AI de supermercados
    $ws.Range("V$c").Copy($ws.Range("X$c")) | Out-Null
    $ws.Range("X$c").Value2 = $tot
    $ws.Range("X${p1}:X$n").Formula = '=' + ($se -f $p1) + '*H' + $p1
    foreach ($p in $S.excecoes.PSObject.Properties) {
        $r = [int]$p.Name
        $ws.Range("X$r").Formula = '=' + ($se -f $r) + '*' + [int]$p.Value
    }
    $ws.Range("X${p1}:X$n").NumberFormat = '#,##0'
    $ws.Columns.Item("W").ColumnWidth = 5
    $ws.Columns.Item("X").ColumnWidth = 14
    if ($ws.AutoFilterMode) { $ws.AutoFilterMode = $false }
    $ws.Range("A${c}:X$n").AutoFilter() | Out-Null
    $xl.Calculate()
    Write-Output ("soma X = {0}" -f $xl.WorksheetFunction.Sum($ws.Range("X${p1}:X$n")))
    $wb.Save(); $wb.Close($false)
}
catch {
    Write-Output ("ERRO na linha {0}: {1}" -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message); throw
}
finally {
    $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
