# Grava, célula por célula, as diferenças do STATUS AUTOPEL (so_diferencas.py) numa CÓPIA do MAPA pelo Excel (COM), 19/11/2026.
# Célula por célula porque as abas têm filtro ligado (em bloco o Excel pula as linhas ocultas). Datas como data (dd/mm/aaaa).
# Não apaga, não insere e não move nada; não atualiza dinâmicas (o usuário atualiza ao abrir, se quiser).
# Uso: powershell -File grava_celulas.ps1 -Arquivo "<cópia.xlsm>" -Tsv "<diferencas.tsv>"
param([Parameter(Mandatory = $true)][string]$Arquivo, [Parameter(Mandatory = $true)][string]$Tsv)
$ErrorActionPreference = "Stop"
$linhas = [IO.File]::ReadAllLines($Tsv, (New-Object Text.UTF8Encoding $true))
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
$xl.AutomationSecurity = 3
try {
    $wb = $xl.Workbooks.Open((Resolve-Path $Arquivo).Path)
    $xl.Calculation = -4135
    $fmtData = "dd" + $xl.International(17) + "mm" + $xl.International(17) + "aaaa"
    $n = 0
    foreach ($l in $linhas) {
        if (-not $l) { continue }
        $c = $l.Split("`t")
        $ws = $wb.Worksheets.Item($c[0])
        $cel = $ws.Range($c[1] + $c[2])
        $valor = $(if ($c.Count -gt 4) { $c[4] } else { "" })
        if ($c[3] -eq 'd') {
            if ($valor) {
                $p = $valor.Split('-')
                $cel.Value2 = ([datetime]::new([int]$p[0], [int]$p[1], [int]$p[2])).ToOADate()
                $cel.NumberFormatLocal = $fmtData
            } else { [void]$cel.ClearContents() }
        } elseif ($c[3] -eq 'n') {
            $cel.Value2 = [double]$valor
        } else {
            if ($valor) { $cel.Value2 = [string]$valor } else { [void]$cel.ClearContents() }
        }
        $n++
    }
    $xl.Calculation = -4105
    $xl.CalculateFull()
    $wb.Save()
    "gravadas $n celulas"
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
