# Aba "I.R.T.T.__Por bandeira (FINAL)", coluna H "Dias sem comprar": formula =SE(G="";"";HOJE()-G), no mesmo padrao
# da coluna F (status), para acompanhar a data da G.  Use numa COPIA; feche o arquivo no Excel.
#   powershell -ExecutionPolicy Bypass -File irtt_dias.ps1 -Xlsx C:\...\copia.xlsx
param([Parameter(Mandatory = $true)][string]$Xlsx, [string]$Aba = 'I.R.T.T.__Por bandeira (FINAL)', [int]$Cab = 4)
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
try {
    $wb = $xl.Workbooks.Open($Xlsx, 0, $false)
    $ws = $wb.Worksheets.Item($Aba)
    if ([string]$ws.Range("H$Cab").Value2 -ne 'Dias sem comprar' -or [string]$ws.Range("G$Cab").Value2 -notmatch 'compra') { throw "layout inesperado em G/H$Cab" }
    $p = $Cab + 1
    $n = $ws.Cells.Item($ws.Rows.Count, 2).End(-4162).Row     # ultima linha com bandeira (col. B)
    $ws.Range("H${p}:H$n").Formula = "=IF(G$p="""","""",TODAY()-G$p)"
    $ws.Range("H${p}:H$n").NumberFormat = '0'
    $xl.Calculate()
    Write-Output ("linhas {0}-{1}; soma H = {2}" -f $p, $n, $xl.WorksheetFunction.Sum($ws.Range("H${p}:H$n")))
    $wb.Save(); $wb.Close($false)
}
catch { Write-Output ("ERRO na linha {0}: {1}" -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message); throw }
finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
