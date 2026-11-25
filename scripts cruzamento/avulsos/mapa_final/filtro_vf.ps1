# Devolve o filtro da Alessandra Dayan VF (25/11/2026) numa CÓPIA do MAPA, pelo Excel (COM):
# STATUS AUTOPEL (coluna J) = INATIVO ou NÃO É CLIENTE e SEGMENTO (coluna K) = DROGARIA, na faixa D8:O558
# (o filtro que estava no arquivo antes de o usuário limpar). Não mexe em valor nenhum.
# Uso: powershell -File filtro_vf.ps1 -Arquivo "<cópia.xlsm>"
param([Parameter(Mandatory = $true)][string]$Arquivo)
$ErrorActionPreference = "Stop"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
$xl.AutomationSecurity = 3
try {
    $wb = $xl.Workbooks.Open((Resolve-Path $Arquivo).Path)
    $ws = $wb.Worksheets.Item("Alessandra Dayan VF")
    $faixa = $ws.Range("D8:O558")
    if ($ws.AutoFilterMode -and $ws.AutoFilter.Range.Address() -ne $faixa.Address()) { $ws.AutoFilterMode = $false }
    if ($ws.FilterMode) { $ws.ShowAllData() }
    [void]$faixa.AutoFilter(7, [object[]]@("INATIVO", "NÃO É CLIENTE"), 7)
    [void]$faixa.AutoFilter(8, "DROGARIA")
    $vis = 0
    for ($r = 9; $r -le 558; $r++) { if (-not $ws.Rows($r).Hidden) { $vis++ } }
    $wb.Save()
    "filtro aplicado: $vis linhas visiveis"
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
