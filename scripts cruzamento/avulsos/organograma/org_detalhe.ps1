# Grava a aba DETALHE do MAPA MERCADO pelo Excel (COM), numa copia (05/11/2026).
# Le dados\detalhe.tsv (linhas) e dados\detalhe.json (estilos e links), gerados pelo org_detalhe.py.
# Cria a aba no fim da pasta (os nomes definidos usam a posicao das abas) ou limpa a que existe e regrava.
# Uso: powershell -File org_detalhe.ps1 -Arquivo "<copia.xlsx>"
param([Parameter(Mandatory = $true)][string]$Arquivo, [string]$Pasta = "")
$ErrorActionPreference = "Stop"
if (-not $Pasta) { $Pasta = Join-Path $PSScriptRoot "dados" }
$linhas = [IO.File]::ReadAllLines((Join-Path $Pasta "detalhe.tsv"), [Text.Encoding]::UTF8)
$cfg = Get-Content -Raw -Encoding UTF8 (Join-Path $Pasta "detalhe.json") | ConvertFrom-Json
if ($linhas.Count -ne [int]$cfg.n) { throw "detalhe.tsv tem $($linhas.Count) linhas e o json diz $($cfg.n)" }
function Cor($r, $g, $b) { return [int]$r + 256 * [int]$g + 65536 * [int]$b }
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
    $wb = $xl.Workbooks.Open((Resolve-Path $Arquivo).Path)
    $ativa = $wb.ActiveSheet.Name
    $ws = $null
    foreach ($s in $wb.Worksheets) { if ($s.Name -eq 'DETALHE') { $ws = $s } }
    if ($ws) { $null = $ws.Hyperlinks.Delete(); $null = $ws.Cells.Clear() }
    else { $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)); $ws.Name = 'DETALHE' }
    $ws.Cells.Font.Name = 'Calibri'; $ws.Cells.Font.Size = 10
    $ws.Range('D:D').NumberFormat = '@'
    $mil = $xl.International(4)
    # valores em blocos de 10 mil linhas; numero de lojas e potencial (colunas C e D, so digitos) vao como numero
    $bloco = 10000
    for ($ini = 0; $ini -lt $linhas.Count; $ini += $bloco) {
        $n = [Math]::Min($bloco, $linhas.Count - $ini)
        $arr = New-Object 'object[,]' $n, 6
        for ($i = 0; $i -lt $n; $i++) {
            $p = $linhas[$ini + $i].Split("`t")
            for ($j = 0; $j -lt 6; $j++) {
                $v = if ($j -lt $p.Length) { $p[$j] } else { '' }
                if (($j -eq 2 -or $j -eq 3) -and $v -match '^\d+(\.\d+)?$') { $arr[$i, $j] = [double]::Parse($v, [Globalization.CultureInfo]::InvariantCulture) }
                elseif ($v -ne '') { $arr[$i, $j] = $v }
            }
        }
        $rg = $ws.Range($ws.Cells.Item($ini + 1, 1), $ws.Cells.Item($ini + $n, 6))
        try { $rg.Value2 = $arr } catch { $rg.Value = $arr }
    }
    $ws.Range('C:C').NumberFormatLocal = "#${mil}##0"
    # numeros do bloco 1 (potencial por pessoa) estao na coluna D
    $ult1 = ($cfg.estilos.pessoa + $cfg.estilos.dado | Measure-Object -Maximum).Maximum
    if ($ult1) { $ws.Range("D5:D$ult1").NumberFormatLocal = "#${mil}##0"; $ws.Range("D5:D$ult1").HorizontalAlignment = -4152 }
    $NAVY = Cor 31 58 95; $BLUE = Cor 46 117 182; $LIGHT = Cor 221 235 247; $GRAY = Cor 242 242 242; $OURO = Cor 255 242 204
    $estilo = @{
        titulo = @{ bold = $true; size = 14; fill = $NAVY; font = (Cor 255 255 255) }
        nota   = @{ italic = $true; size = 9 }
        secao  = @{ bold = $true; size = 12; fill = $NAVY; font = (Cor 255 255 255) }
        cab    = @{ bold = $true; fill = $LIGHT }
        pessoa = @{ bold = $true; fill = $GRAY }
        bloco  = @{ bold = $true; size = 12; fill = $BLUE; font = (Cor 255 255 255) }
        pot    = @{ bold = $true; fill = $OURO }
        status = @{ bold = $true; fill = $LIGHT }
    }
    foreach ($pr in $cfg.estilos.PSObject.Properties) {
        $e = $estilo[$pr.Name]; if (-not $e) { continue }
        foreach ($r in $pr.Value) {
            $rg = $ws.Range("A${r}:F${r}")
            if ($e.bold) { $rg.Font.Bold = $true }
            if ($e.italic) { $rg.Font.Italic = $true }
            if ($e.size) { $rg.Font.Size = $e.size }
            if ($e.fill) { $rg.Interior.Color = $e.fill }
            if ($e.font) { $rg.Font.Color = $e.font }
        }
    }
    foreach ($l in $cfg.links) {
        $r = [int]$l[0]; $c = [int]$l[1]; $alvo = $l[2]
        $sub = if ([string]$alvo -eq 'ORGANOGRAMA') { "'ORGANOGRAMA impressao'!A1" } else { "'DETALHE'!A$([int]$alvo)" }
        $cel = $ws.Cells.Item($r, $c); $txt = [string]$cel.Value2
        $null = $ws.Hyperlinks.Add($cel, '', $sub, [string]$l[3], $txt)
    }
    $larg = @(48, 34, 22, 22, 22, 34)
    for ($j = 0; $j -lt 6; $j++) { $ws.Columns.Item($j + 1).ColumnWidth = $larg[$j] }
    $null = $ws.Activate(); $win = $wb.Windows.Item(1); $win.DisplayGridlines = $true; $null = $ws.Range('A1').Select()
    $null = $wb.Worksheets.Item($ativa).Activate()
    $wb.Save()
    Write-Output ("DETALHE: {0} linhas, {1} links internos, aba na posicao {2} de {3}" -f $linhas.Count, @($cfg.links).Count, $ws.Index, $wb.Worksheets.Count)
    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
