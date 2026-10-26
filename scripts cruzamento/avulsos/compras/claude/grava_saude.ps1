# Grava no "Relatorio Revenda atualizado.xlsx" (usar numa COPIA fechada) o resultado de saude_ecommerce.py:
#   - aba "analise compras CLAUDE": secoes 7 (saude do estoque), 8 (triagem do excedente para e-commerce) e
#     9 (concentracao de clientes), depois da secao 6 (saude.json);
#   - tabela ParamClaude (aba param CLAUDE): 9 colunas por SKU no fim da tabela (param_saude.csv).
#   Recusa se a secao 7 ou as colunas ja existirem.
#   powershell -ExecutionPolicy Bypass -File grava_saude.ps1 -Xlsx C:\...\copia.xlsx -Dir C:\...\saida
param([Parameter(Mandatory = $true)][string]$Xlsx, [Parameter(Mandatory = $true)][string]$Dir)
$ErrorActionPreference = 'Stop'
$JS = Get-Content -Raw -Encoding UTF8 -Path (Join-Path $Dir 'saude.json') | ConvertFrom-Json
$CSV = Import-Csv -Path (Join-Path $Dir 'param_saude.csv') -Delimiter ';' -Encoding UTF8
$INV = [Globalization.CultureInfo]::InvariantCulture

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
try {
    $wb = $xl.Workbooks.Open($Xlsx, 0, $false)
    $xl.Calculation = -4135
    $mil = $xl.International(4); $dec = $xl.International(3)
    $FI = "#${mil}##0"; $FD2 = "0${dec}00"; $FP = "0${dec}0%"; $FR = "`"R`$`" #${mil}##0"; $FR2 = "`"R`$`" #${mil}##0${dec}00"

    $wa = $wb.Worksheets.Item('analise compras CLAUDE')
    $wp = $wb.Worksheets.Item('param CLAUDE')
    $tp = $wp.ListObjects.Item('ParamClaude')
    $novas = @($CSV[0].PSObject.Properties.Name | Select-Object -Skip 1)
    foreach ($lc in $tp.ListColumns) { if ($novas -contains $lc.Name) { throw "ParamClaude ja tem a coluna '$($lc.Name)'" } }
    $ult = $wa.Cells.Item($wa.Rows.Count, 2).End(-4162).Row
    for ($r = 1; $r -le $ult; $r++) {
        $t = [string]$wa.Cells.Item($r, 2).Value2
        if ($t.StartsWith('7. SAÚDE')) { throw 'a secao 7 ja existe' }
        if ($t.StartsWith('6. SUGESTÕES')) { $r6 = $r }
    }
    if (-not $r6) { throw 'secao 6 nao encontrada na analise' }
    "analise: ultima linha $ult (secao 6 na linha $r6)"

    # --- ParamClaude: colunas novas no fim -------------------------------------------------------------------------
    $cods = $tp.ListColumns.Item('Codigo').DataBodyRange.Value2
    $nl = $cods.GetLength(0)
    $mapa = @{}; foreach ($l in $CSV) { $mapa[$l.Codigo] = $l }
    $num = @('Fora da política QTD', 'Fora da política R$', 'Clientes 12m', '% maior cliente')
    $arr = New-Object 'object[,]' $nl, $novas.Count
    $acha = 0
    for ($i = 1; $i -le $nl; $i++) {
        $l = $mapa[[string]$cods[$i, 1]]
        if ($l) { $acha++ }
        for ($j = 0; $j -lt $novas.Count; $j++) {
            $v = if ($l) { $l.($novas[$j]) } else { '' }
            $arr[($i - 1), $j] = $(if ($num -contains $novas[$j] -and $v -ne '') { [double]::Parse($v, $INV) } else { [string]$v })
        }
    }
    "ParamClaude: $acha de $nl SKUs com dados novos"
    $c1 = $tp.Range.Column + $tp.Range.Columns.Count
    $h0 = $tp.HeaderRowRange.Row
    for ($j = 0; $j -lt $novas.Count; $j++) { $wp.Cells.Item($h0, $c1 + $j).Value2 = $novas[$j] }
    $tp.Resize($wp.Range($wp.Cells.Item($h0, $tp.Range.Column), $wp.Cells.Item($h0 + $nl, $c1 + $novas.Count - 1)))
    $col = @{}; for ($j = 0; $j -lt $novas.Count; $j++) { $col[$novas[$j]] = $c1 + $j }
    # texto fica texto (o Excel leria '2026-10' como data)
    foreach ($n in $novas) { if ($num -notcontains $n) { $wp.Range($wp.Cells.Item($h0 + 1, $col[$n]), $wp.Cells.Item($h0 + $nl, $col[$n])).NumberFormat = '@' } }
    $wp.Range($wp.Cells.Item($h0 + 1, $c1), $wp.Cells.Item($h0 + $nl, $c1 + $novas.Count - 1)).Value = $arr
    $wp.Range($wp.Cells.Item($h0 + 1, $col['Fora da política QTD']), $wp.Cells.Item($h0 + $nl, $col['Fora da política QTD'])).NumberFormatLocal = $FI
    $wp.Range($wp.Cells.Item($h0 + 1, $col['Fora da política R$']), $wp.Cells.Item($h0 + $nl, $col['Fora da política R$'])).NumberFormatLocal = $FR
    $wp.Range($wp.Cells.Item($h0 + 1, $col['% maior cliente']), $wp.Cells.Item($h0 + $nl, $col['% maior cliente'])).NumberFormatLocal = $FP
    foreach ($n in $novas) { $wp.Columns.Item($col[$n]).ColumnWidth = 14 }
    foreach ($n in @('Triagem e-commerce', 'Maior cliente')) { $wp.Columns.Item($col[$n]).ColumnWidth = 34 }

    # --- aba de analise: secoes 7, 8 e 9 ---------------------------------------------------------------------------
    $script:r = $ult + 3
    function Tit([string]$t, [int]$tam = 13) { $c = $wa.Cells.Item($script:r, 2); $c.Value2 = $t; $c.Font.Bold = $true; $c.Font.Size = $tam; $script:r++ }
    function Par([string]$t, [int]$linhas = 1) {
        $m = $wa.Range($wa.Cells.Item($script:r, 2), $wa.Cells.Item($script:r, 12)); $m.Merge(); $m.Value2 = $t
        $m.WrapText = $true; $m.VerticalAlignment = -4160; $m.RowHeight = 15 * $linhas; $script:r++
    }
    function Linhas([string]$t) { [math]::Max(1, [math]::Ceiling($t.Length / 205)) }
    function Txt($lista) { foreach ($t in $lista) { if ($t) { Par $t (Linhas $t) } } }
    function Tab($dados, $fmts) {
        $r0 = $script:r; $nlin = $dados.Count; $ncol = $dados[0].Count
        $a = New-Object 'object[,]' $nlin, $ncol
        for ($i = 0; $i -lt $nlin; $i++) {
            for ($j = 0; $j -lt $ncol; $j++) {
                $v = $dados[$i][$j]
                $a[$i, $j] = $(if ($null -eq $v) { '' } elseif ($v -is [long] -or $v -is [int] -or $v -is [decimal]) { [double]$v }
                               elseif ($v -is [string] -and $v -match '^\d{4}-\d{2}$') { "'" + $v } else { $v })
            }
        }
        $rgt = $wa.Range($wa.Cells.Item($r0, 2), $wa.Cells.Item($r0 + $nlin - 1, 1 + $ncol)); $rgt.Value = $a
        $rgt.Borders.LineStyle = 1; $rgt.Borders.Weight = 2
        $h = $wa.Range($wa.Cells.Item($r0, 2), $wa.Cells.Item($r0, 1 + $ncol)); $h.Font.Bold = $true; $h.Interior.Color = 14599344; $h.Font.Color = 16777215; $h.WrapText = $true
        for ($j = 0; $j -lt $ncol; $j++) {
            $f = $fmts[$j]
            if ($f) { $wa.Range($wa.Cells.Item($r0 + 1, 2 + $j), $wa.Cells.Item($r0 + $nlin - 1, 2 + $j)).NumberFormatLocal = $f }
        }
        if ([string]$dados[$nlin - 1][0] -eq 'Total') { $wa.Range($wa.Cells.Item($r0 + $nlin - 1, 2), $wa.Cells.Item($r0 + $nlin - 1, 1 + $ncol)).Font.Bold = $true }
        $script:r = $r0 + $nlin + 1
    }
    $TX = $JS.textos
    $saude = @($null, $FI, $FR, $FP, $FP, $FR, $FR, $FR, $FR)

    Tit "7. SAÚDE DO ESTOQUE: IQR, SLOW MOVING E PHASE OUT (foto de $($JS.info.data))"
    Par $TX.s7_intro (Linhas $TX.s7_intro)
    $script:r++
    Tab $JS.resumo @($null, $FI, $FR, $FP)
    Txt $TX.s7
    $script:r++
    Tit '7.1 Por curva ABC'
    Tab $JS.abc $saude
    Tit '7.2 Por comprador'
    Tab $JS.comprador $saude
    Tit '7.3 Por grupo (coluna Grupo da Geral)'
    Tab $JS.grupo $saude
    Tit '7.4 Os 30 SKUs com mais valor fora da política'
    Tab $JS.top @($null, $null, $null, $null, $null, $FR, $FR, $null)
    $script:r++

    Tit '8. TRIAGEM DO EXCEDENTE PARA E-COMMERCE'
    Par $TX.s8_intro (Linhas $TX.s8_intro)
    $script:r++
    Tab $JS.triagem @($null, $FI, $FR, $FP, $FI, $null)
    Txt $TX.s8
    $script:r++
    Tit '8.1 Os 40 genéricos com mais excedente (candidatos ao piloto)'
    Tab $JS.genericos @($null, $null, $null, $null, $FI, $FR2, $FR, $FR2, $FD2, $null)
    $script:r++

    Tit '9. CONCENTRAÇÃO DE CLIENTES'
    Par $TX.s9_intro (Linhas $TX.s9_intro)
    $script:r++
    Tab $JS.clientes @($null, $FI, $FR, $FR, $FR, $FI, $null)
    Txt $TX.s9
    $script:r++
    Tit '9.1 Clientes que concentram SKUs (80% ou mais da venda do SKU)'
    Tab $JS.donos @($null, $FI, $FR, $FR)
    Tit '9.2 Os 30 SKUs concentrados com mais estoque + compra CLAUDE'
    Tab $JS.concentrados @($null, $null, $null, $FI, $FP, $null, $FR, $FR, $FP)
    "analise: secoes 7 a 9 nas linhas $($ult + 3) a $($script:r - 1)"

    $xl.Calculation = -4105
    $xl.CalculateFull()
    $wb.Save()
    'ok'
}
catch {
    "ERRO na linha $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"
    throw
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
