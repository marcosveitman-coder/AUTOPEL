# Contatos do LinkedIn no MAPA MERCADO V2 (23/11/2026), numa CÓPIA, pelo Excel (COM):
#  1. nas abas FINAL pedidas (ESTACIONAMENTO, POSTOS), copia o cabeçalho de contatos que o usuário montou na
#     Alessandra Dayan VF (P7:AD8: 5 cargos x Nome / LinkedIn / Cargo no LinkedIn) para AB5:AP6, com formato e larguras;
#  2. grava os valores célula por célula (a VF fica filtrada como está: não mexe em filtro, linha ou coluna).
# Entrada: TSV (aba, coluna, linha, tipo, valor) de contatos_mapa.py.
# Uso: powershell -File grava_contatos.ps1 -Arquivo "<cópia.xlsm>" -Tsv "<contatos.tsv>" -Abas "ESTACIONAMENTO FINAL|POSTOS GAS & CONVENIENCIA FINAL"
param([Parameter(Mandatory = $true)][string]$Arquivo, [Parameter(Mandatory = $true)][string]$Tsv, [string]$Abas = "")
$ErrorActionPreference = "Stop"
$linhas = [IO.File]::ReadAllLines($Tsv, (New-Object Text.UTF8Encoding $true))
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
$xl.AutomationSecurity = 3
try {
    $wb = $xl.Workbooks.Open((Resolve-Path $Arquivo).Path)
    $xl.Calculation = -4135
    $vf = $wb.Worksheets.Item("Alessandra Dayan VF")
    foreach ($nomeAba in ($Abas.Split('|') | Where-Object { $_ })) {
        $ws = $wb.Worksheets.Item($nomeAba)
        [void]$vf.Range("P7:AD8").Copy($ws.Range("AB5"))
        for ($k = 0; $k -lt 15; $k++) { $ws.Columns(28 + $k).ColumnWidth = $vf.Columns(16 + $k).ColumnWidth }
        $ws.Rows(5).RowHeight = $vf.Rows(7).RowHeight
    }
    $n = 0
    foreach ($l in $linhas) {
        if (-not $l) { continue }
        $c = $l.Split("`t")
        $wb.Worksheets.Item($c[0]).Range($c[1] + $c[2]).Value2 = [string]$c[4]
        $n++
    }
    $xl.Calculation = -4105
    $wb.Save()
    "gravadas $n celulas"
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
