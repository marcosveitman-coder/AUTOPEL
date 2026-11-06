# Atualiza SO o potencial da aba ORGANOGRAMA pelo Excel (COM), numa copia (05/11/2026).
# Le org_potencial.json (org_potencial.py): troca as celulas indicadas e insere as linhas novas da tabela de potencial
# logo depois da ultima linha dela, copiando o formato (e as mesclagens) dessa linha.
# Uso: powershell -File org_potencial.ps1 -Arquivo "<copia.xlsx>" [-Json dados\org_potencial.json]
param([Parameter(Mandatory = $true)][string]$Arquivo, [string]$Json = "")
$ErrorActionPreference = "Stop"
if (-not $Json) { $Json = Join-Path $PSScriptRoot "dados\org_potencial.json" }
$cfg = Get-Content -Raw -Encoding UTF8 $Json | ConvertFrom-Json
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
    $wb = $xl.Workbooks.Open((Resolve-Path $Arquivo).Path)
    $ws = $wb.Worksheets.Item("ORGANOGRAMA")
    foreach ($t in $cfg.trocas) {
        $cel = $ws.Range($t.ref)
        if ($t.valor -is [string]) { $cel.Value2 = [string]$t.valor } else { $cel.Value2 = [double]$t.valor }
    }
    $base = [int]$cfg.apos_linha
    $qtd = @($cfg.novas).Count
    if ($qtd -gt 0) {
        $ini = $base + 1; $fim = $base + $qtd
        $null = $ws.Range("${ini}:${fim}").EntireRow.Insert()
        $null = $ws.Range("${base}:${base}").EntireRow.Copy($ws.Range("${ini}:${fim}"))
        $alt = $ws.Rows.Item($base).RowHeight
        $i = 0
        foreach ($linha in $cfg.novas) {
            $r = $ini + $i
            $ws.Rows.Item($r).RowHeight = $alt
            $ws.Rows.Item($r).WrapText = $true  # nome de aba longo (POSTOS GAS & CONVENIENCIA FINAL) quebra em 2 linhas
            foreach ($c in $linha) {
                $cel = $ws.Cells.Item($r, [int]$c.col)
                if ($c.valor -is [string]) { $cel.Value2 = [string]$c.valor } else { $cel.Value2 = [double]$c.valor }
            }
            $i++
        }
    }
    $wb.Save()
    $wb.Close($false)
    Write-Output ("trocas: {0} | linhas novas: {1} depois da linha {2}" -f @($cfg.trocas).Count, $qtd, $base)
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
