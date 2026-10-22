# Ajusta a aba "SUPERMERCADOS (reuniao)" pelo proprio Excel: grava celulas (nº de lojas, data/status/vendedor das
# linhas unificadas, fonte) e exclui as linhas repetidas (de baixo para cima, para o Excel acertar formulas e
# tabelas dinamicas).  Depois atualiza as tabelas dinamicas que usam a aba.  Use numa COPIA; feche o arquivo no Excel.
#   powershell -ExecutionPolicy Bypass -File ajusta_reuniao.ps1 -Xlsx C:\...\copia.xlsx -Spec C:\...\ajuste_spec.json
param([Parameter(Mandatory = $true)][string]$Xlsx, [Parameter(Mandatory = $true)][string]$Spec)
$ErrorActionPreference = 'Stop'
$S = Get-Content -Raw -Encoding UTF8 -Path $Spec | ConvertFrom-Json

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
try {
    $wb = $xl.Workbooks.Open($Xlsx, 0, $false)
    $ws = $wb.Worksheets.Item($S.aba)
    foreach ($c in $S.celulas) {
        $r = $ws.Range($c.cel)
        switch ($c.tipo) {
            'num'  { $r.Value2 = [double]$c.valor }
            'data' { $r.Value2 = [double]$c.valor; $r.NumberFormat = 'dd/mm/yyyy' }   # numero de serie do Excel
            'txt'  { $r.Value2 = [string]$c.valor }
            'vazio' { $r.ClearContents() | Out-Null }
        }
    }
    if ($S.cabecalho) {
        $h = $ws.Range($S.cabecalho.cel)
        $ws.Range($S.cabecalho.modelo).Copy($h) | Out-Null     # traz o formato do cabecalho vizinho (sem area de transferencia)
        $h.Value2 = [string]$S.cabecalho.valor
        $ws.Columns.Item($S.cabecalho.col).ColumnWidth = 60
    }
    $linhas = @($S.excluir | ForEach-Object { [int]$_ } | Sort-Object -Descending)
    foreach ($n in $linhas) { $ws.Rows.Item($n).Delete() | Out-Null }
    Write-Output ("excluidas: {0}" -f $linhas.Count)
    foreach ($nome in $S.atualizar_abas) {
        foreach ($pt in $wb.Worksheets.Item($nome).PivotTables()) { $pt.PivotCache().Refresh() | Out-Null; Write-Output ("atualizada: {0} / {1}" -f $nome, $pt.Name) }
    }
    $wb.Save()
    $wb.Close($false)
}
catch {
    Write-Output ("ERRO na linha {0}: {1}" -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message)
    Write-Output $_.InvocationInfo.Line
    throw
}
finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
