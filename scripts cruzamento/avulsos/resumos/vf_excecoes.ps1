# Aba "FARMACIA VF", coluna X: mantem a formula do usuario (caixas por porte nos parametros $X$1:$X$4) e, so nas linhas
# de rede/associacao em que o nº de lojas da conta nao e o da coluna H (para nao contar loja duas vezes), troca o "*H<n>"
# do fim da formula pelo nº de lojas sem duplicidade.  Use numa COPIA; feche o arquivo no Excel.
#   powershell -ExecutionPolicy Bypass -File vf_excecoes.ps1 -Xlsx C:\...\copia.xlsx -Spec C:\...\vf_wx_23.json
param([Parameter(Mandatory = $true)][string]$Xlsx, [Parameter(Mandatory = $true)][string]$Spec)
$ErrorActionPreference = 'Stop'
$S = Get-Content -Raw -Encoding UTF8 -Path $Spec | ConvertFrom-Json
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
try {
    $wb = $xl.Workbooks.Open($Xlsx, 0, $false)
    $ws = $wb.Worksheets.Item($S.aba)
    $n = 0
    foreach ($p in $S.excecoes.PSObject.Properties) {
        $r = [int]$p.Name
        $f = [string]$ws.Range("X$r").Formula
        if ($f -notmatch "\*H$r$") { throw "X${r}: formula inesperada: $f" }
        $ws.Range("X$r").Formula = ($f -replace "\*H$r$", ('*' + [int]$p.Value)); $n++
    }
    $xl.Calculate()
    Write-Output ("trocadas: {0}; soma X = {1}" -f $n, $xl.WorksheetFunction.Sum($ws.Range("X$([int]$S.cab + 1):X$([int]$S.ult)")))
    $wb.Save(); $wb.Close($false)
}
catch { Write-Output ("ERRO na linha {0}: {1}" -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message); throw }
finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
