# 2o passo da troca de nomes das colunas de previsao (o 1o e renomeia_previsao.py, que edita o XML).
# Abre pelo Excel, confere que as dinamicas reconhecem os campos renomeados, atualiza a cache da Table1 (reescreve os
# titulos nas celulas), devolve as larguras de coluna das abas com dinamica e ajusta os textos.
# As alturas de linha nao sao mexidas: as que se ajustam ao texto continuam automaticas (gravar a altura a fixaria).
#   powershell -ExecutionPolicy Bypass -File renomeia_previsao.ps1 -Xlsx C:\...\copia.xlsx
param([Parameter(Mandatory = $true)][string]$Xlsx)
$ErrorActionPreference = 'Stop'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
try {
    $wb = $xl.Workbooks.Open($Xlsx, 0, $false)
    $novos = @('Compra mês vigente QTD', 'Compra mês vigente R$', 'Compra próximo mês QTD', 'Compra próximo mês R$',
               'Compra mês subsequente QTD', 'Compra mês subsequente R$')

    # campos de valor de uma dinamica (no PowerShell a colecao so enumera pelo indice)
    function Dados($pt) { $d = $pt.DataFields(); for ($i = 1; $i -le $d.Count; $i++) { $d.Item($i) } }
    # dinamicas da Table1 e seus campos de valor (antes de atualizar)
    $alvo = @(); $total = 0
    foreach ($ws in $wb.Worksheets) {
        foreach ($pt in $ws.PivotTables()) {
            $total++
            if ([string]$pt.PivotCache().SourceData -ne 'Table1') { continue }
            if (@(Dados $pt | Where-Object { $novos -contains $_.SourceName }).Count) { $alvo += , @($ws, $pt) }
        }
    }
    "dinamicas no arquivo: $total; com as colunas renomeadas: $($alvo.Count)"
    if ($total -ne 65 -or $alvo.Count -ne 12)   # 11 do usuario + a da PROGRAMACAO TRIMESTRE
    { throw 'numero de dinamicas inesperado (arquivo reparado?)' }
    function Campos($pt) { (Dados $pt | ForEach-Object { "$($_.Position):$($_.Caption)<-$($_.SourceName)" }) -join ' | ' }
    $antes = @{}
    foreach ($a in $alvo) { $antes["$($a[0].Name)/$($a[1].Name)"] = Campos $a[1] }

    # larguras das abas com dinamica da Table1 (a atualizacao ajusta as colunas da dinamica ao conteudo)
    $medidas = @{}
    foreach ($ws in $wb.Worksheets) {
        $tem = $false; foreach ($pt in $ws.PivotTables()) { if ([string]$pt.PivotCache().SourceData -eq 'Table1') { $tem = $true } }
        if (-not $tem) { continue }
        $w = @{}; foreach ($c in 1..90) { $w[$c] = $ws.Columns.Item($c).ColumnWidth }
        $medidas[$ws.Name] = @($ws, $w)
    }

    $pc = $null
    foreach ($p in $wb.PivotCaches()) { if ([string]$p.SourceData -eq 'Table1') { $pc = $p } }
    $pc.Refresh()

    foreach ($a in $alvo) {
        $k = "$($a[0].Name)/$($a[1].Name)"; $depois = Campos $a[1]
        if ($depois -ne $antes[$k]) { throw "campos mudaram em ${k}:`n  antes : $($antes[$k])`n  depois: $depois" }
        "$k -> $depois"
    }
    $mud = 0
    foreach ($m in $medidas.Values) {
        $ws, $w = $m
        foreach ($c in $w.Keys) { if ($ws.Columns.Item($c).ColumnWidth -ne $w[$c]) { $ws.Columns.Item($c).ColumnWidth = $w[$c]; $mud++ } }
    }
    "larguras devolvidas: $mud"

    # textos da LEIA-ME
    $wl = $wb.Worksheets.Item('LEIA-ME COMPRAS')
    $wl.Range('E20').Value2 = 'Coluna AT da Geral ("Compra mês vigente QTD").'
    $wl.Range('E21').Value2 = 'Coluna AV da Geral ("Compra próximo mês QTD").'
    $wl.Range('E22').Value2 = 'Coluna AX da Geral ("Compra mês subsequente QTD").'
    $ok = $false
    foreach ($r in 30..60) {
        if ([string]$wl.Range("B$r").Value2 -like 'Os nomes OUTUBRO*') {
            $wl.Range("B$r").Value2 = 'Cabeçalhos sem nome de mês: "Compra mês vigente / próximo mês / mês subsequente" (QTD e R$). Não precisam mudar quando o mês vira; o nome do mês aparece no tracking e na PROGRAMACAO TRIMESTRE, pela data de hoje.'
            $ok = $true
        }
    }
    if (-not $ok) { throw 'linha dos nomes dos meses nao achada na LEIA-ME' }

    # tracking: nome do mes pela data de hoje
    $wd = $wb.Worksheets.Item('dinamicas')
    $rot = @('=PROPER(TEXT(TODAY(),"mmmm"))&" R$"', '=PROPER(TEXT(EDATE(TODAY(),1),"mmmm"))&" R$"', '=PROPER(TEXT(EDATE(TODAY(),2),"mmmm"))&" R$"')
    foreach ($par in @(@('AF4', 'AG4', 'AH4'), @('AG21', 'AH21', 'AI21'))) {
        for ($i = 0; $i -lt 3; $i++) {
            $c = $wd.Range($par[$i]); $esp = @('Outubro R$', 'Novembro R$', 'Dezembro R$')[$i]
            if ([string]$c.Value2 -ne $esp) { throw "dinamicas $($par[$i]) = '$($c.Value2)' (esperado '$esp')" }
            $c.Formula = $rot[$i]
        }
    }

    $xl.CalculateFull()
    $wb.Save()
    'ok'
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
