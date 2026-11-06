# Grava a aba oculta "CONFIG DETALHE" do MAPA MERCADO pelo Excel (COM), numa copia (06/11/2026).
# Le dados\config_detalhe.tsv (org_detalhe.py). Apaga a aba DETALHE se existir (a lista e montada pela macro na hora
# do duplo clique; nao fica guardada no arquivo). A CONFIG DETALHE fica no fim da pasta, oculta, tudo como texto.
# Uso: powershell -File org_detalhe.ps1 -Arquivo "<copia.xlsx>"
param([Parameter(Mandatory = $true)][string]$Arquivo, [string]$Pasta = "")
$ErrorActionPreference = "Stop"
if (-not $Pasta) { $Pasta = Join-Path $PSScriptRoot "dados" }
$linhas = [IO.File]::ReadAllLines((Join-Path $Pasta "config_detalhe.tsv"), [Text.Encoding]::UTF8)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
    $wb = $xl.Workbooks.Open((Resolve-Path $Arquivo).Path)
    $ativa = $wb.ActiveSheet.Name
    if ($ativa -eq 'DETALHE' -or $ativa -eq 'CONFIG DETALHE') { $ativa = 'ORGANOGRAMA impressao' }
    $null = $wb.Worksheets.Item($ativa).Activate()
    $cfg = $null; $det = $null
    foreach ($s in $wb.Worksheets) { if ($s.Name -eq 'CONFIG DETALHE') { $cfg = $s }; if ($s.Name -eq 'DETALHE') { $det = $s } }
    if ($det) { $det.Visible = -1; $det.Delete() }
    if ($cfg) { $cfg.Visible = -1; $null = $cfg.Cells.Clear() }
    else { $cfg = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)); $cfg.Name = 'CONFIG DETALHE' }
    $cfg.Cells.NumberFormat = '@'
    $n = $linhas.Count
    $arr = New-Object 'object[,]' $n, 8
    for ($i = 0; $i -lt $n; $i++) {
        $p = $linhas[$i].Split("`t")
        for ($j = 0; $j -lt 8; $j++) { if ($j -lt $p.Length -and $p[$j] -ne '') { $arr[$i, $j] = $p[$j] } }
    }
    $cfg.Range($cfg.Cells.Item(1, 1), $cfg.Cells.Item($n, 8)).Value2 = $arr
    foreach ($t in 'PESSOAS', 'SEGMENTOS', 'CELULAS') {
        for ($i = 1; $i -le $n; $i++) {
            if ([string]$cfg.Cells.Item($i, 1).Value2 -eq $t) { $cfg.Rows.Item($i).Font.Bold = $true; $cfg.Rows.Item($i + 1).Font.Bold = $true }
        }
    }
    $cfg.Rows.Item(1).Font.Bold = $true
    $larg = @(14, 40, 16, 12, 60, 50, 14, 40)
    for ($j = 0; $j -lt 8; $j++) { $cfg.Columns.Item($j + 1).ColumnWidth = $larg[$j] }
    $null = $wb.Worksheets.Item($ativa).Activate()
    $cfg.Visible = 0
    $wb.Save()
    Write-Output ("CONFIG DETALHE: {0} linhas, oculta, posicao {1} de {2}; DETALHE antiga apagada: {3}" -f $n, $cfg.Index, $wb.Worksheets.Count, [bool]$det)
    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
