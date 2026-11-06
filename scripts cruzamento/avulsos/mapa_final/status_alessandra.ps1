# STATUS RECEITA / PORTE / STATUS AUTOPEL (05/11/2026) - passo 6: grava a aba Alessandra Dayan pelo Excel (COM), numa copia.
# Le dados/status_alessandra.json (status_alessandra.py): insere 2 colunas antes das colunas indicadas (da direita para a
# esquerda), copia o formato da coluna BASE de cada tabela para as 3 colunas novas e grava cabecalho e valores.
# Uso: powershell -File status_alessandra.ps1 -Arquivo "<copia.xlsx>" [-Json dados\status_alessandra.json]
param([Parameter(Mandatory = $true)][string]$Arquivo, [string]$Json = "")
$ErrorActionPreference = "Stop"
if (-not $Json) { $Json = Join-Path $PSScriptRoot "dados\status_alessandra.json" }
$cfg = Get-Content -Raw -Encoding UTF8 $Json | ConvertFrom-Json
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
    $wb = $xl.Workbooks.Open((Resolve-Path $Arquivo).Path)
    $ws = $wb.Worksheets.Item("Alessandra Dayan")
    foreach ($p in $cfg.inserir) {
        $null = $ws.Range($ws.Cells(1, $p), $ws.Cells(1, $p + 1)).EntireColumn.Insert()
    }
    $larguras = @(14, 30, 16)
    foreach ($t in $cfg.tabelas) {
        $linhas = @($t.valores[0] | ForEach-Object { [int]$_[0] })
        $ultima = ($linhas | Measure-Object -Maximum).Maximum
        $baseNome = [string]$ws.Cells(2, [int]$t.col_base).Value2
        if ($baseNome.Trim() -ne "BASE") { throw "coluna BASE fora do lugar em $($t.seg): '$baseNome'" }
        for ($k = 0; $k -lt 3; $k++) {
            $c = [int]$t.cols[$k]
            $src = $ws.Range($ws.Cells(2, [int]$t.col_base), $ws.Cells($ultima, [int]$t.col_base))
            $dst = $ws.Range($ws.Cells(2, $c), $ws.Cells($ultima, $c))
            $null = $src.Copy($dst)
            $n = $ultima - 1
            $arr = New-Object 'object[,]' $n, 1
            for ($i = 0; $i -lt $n; $i++) { $arr[$i, 0] = $null }
            $arr[0, 0] = [string]$cfg.cab[$k]
            foreach ($par in $t.valores[$k]) { $arr[([int]$par[0] - 2), 0] = [string]$par[1] }
            $dst.Value2 = $arr
            $ws.Columns.Item($c).ColumnWidth = $larguras[$k]
        }
        Write-Output ("{0}: colunas {1}-{2}, linhas 3-{3}" -f $t.seg, $t.cols[0], $t.cols[2], $ultima)
    }
    $wb.Save()
    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
