# "Total de caixas/ mes" nas 7 abas FINAL, com o mesmo racional da aba "FARMACIA VF" (05/11/2026). Pelo Excel (COM), numa copia.
# Em cada aba:
#  - insere 5 linhas no topo (o cabecalho vai da linha 1 para a 6, como na FARMACIA VF) e limpa o formato delas;
#  - copia a tabela de caixas por porte da FARMACIA VF (W1:X4: a 17, b 12, c 7, d 4) para as linhas 1-4, com a letra na
#    ultima coluna atual e o valor na coluna nova (os valores ficam em cada aba, para simular por segmento);
#  - coluna nova no fim: "Total de caixas/ mes" = SE(1a letra do PORTE = A; X$1; ...; D; X$4; 0) x No DE LOJAS;
#  - refaz o filtro (cabecalho na linha 6, ate a coluna nova) e congela o painel abaixo da linha 6.
# Para se a aba ja tiver o cabecalho na linha 6 ou estiver com filtro aplicado.
# Uso: powershell -File caixas_final.ps1 -Arquivo "<copia.xlsx>"
param([Parameter(Mandatory = $true)][string]$Arquivo)
$ErrorActionPreference = "Stop"
$abas = @('ALIMENTACAO FINAL', 'REDE DE LOJAS FINAL', 'CONSTRUÇÃO FINAL', 'POSTOS GAS & CONVENIENCIA FINAL', 'PET SHOP  FINAL',
          'ESTACIONAMENTO FINAL', 'OUTROS FINAL')
function Letra([int]$n) { $s = ''; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [char](65 + $m) + $s; $n = [int](($n - $m - 1) / 26) }; return $s }
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
    $wb = $xl.Workbooks.Open((Resolve-Path $Arquivo).Path)
    $ativa = $wb.ActiveSheet.Name
    $vf = $wb.Worksheets.Item('"FARMACIA VF"')
    $tituloCx = [string]$vf.Range('X6').Value2
    if ($tituloCx -notlike 'Total de caixas*') { throw "X6 da FARMACIA VF mudou: '$tituloCx'" }
    foreach ($nome in $abas) {
        $ws = $wb.Worksheets.Item($nome)
        if ([string]$ws.Cells.Item(1, 2).Value2 -ne 'NOME FANTASIA') { throw "$nome : cabecalho nao esta na linha 1 (ja foi feito?)" }
        if ($ws.AutoFilterMode -and $ws.AutoFilter.FilterMode) { throw "$nome : filtro aplicado; tirar o filtro antes" }
        # colunas pelo cabecalho
        $ultCol = 1; while ([string]$ws.Cells.Item(1, $ultCol + 1).Value2 -ne '') { $ultCol++ }
        $cPorte = 0; $cLojas = 0
        for ($c = 1; $c -le $ultCol; $c++) {
            $h = [string]$ws.Cells.Item(1, $c).Value2
            if ($h -eq 'PORTE') { $cPorte = $c }
            if ($h -like 'Nº DE LOJAS*') { $cLojas = $c }
        }
        if (-not $cPorte -or -not $cLojas) { throw "$nome : PORTE ou No DE LOJAS nao achado" }
        $ultLin = $ws.Cells.Item($ws.Rows.Count, 2).End(-4162).Row
        # 5 linhas no topo
        $null = $ws.Range('1:5').EntireRow.Insert()
        $null = $ws.Range('1:5').EntireRow.ClearFormats()
        $ws.Range('1:5').EntireRow.RowHeight = $ws.StandardHeight
        $hr = 6; $ini = 7; $fim = $ultLin + 5; $nc = $ultCol + 1; $lc = $ultCol
        $L = Letra $nc; $LP = Letra $cPorte; $LL = Letra $cLojas
        # tabela de caixas por porte (valores e formato da FARMACIA VF)
        $null = $vf.Range('W1:X4').Copy($ws.Range("$(Letra $lc)1:${L}4"))
        # cabecalho e coluna nova
        $null = $ws.Cells.Item($hr, $lc).Copy($ws.Cells.Item($hr, $nc))
        $ws.Cells.Item($hr, $nc).Value2 = $tituloCx
        $null = $ws.Range("${LL}${ini}:${LL}${fim}").Copy($ws.Range("${L}${ini}:${L}${fim}"))
        $f = "=IF(LEFT(`$${LP}${ini},1)=""A"",`$${L}`$1,IF(LEFT(`$${LP}${ini},1)=""B"",`$${L}`$2,IF(LEFT(`$${LP}${ini},1)=""C"",`$${L}`$3," +
             "IF(LEFT(`$${LP}${ini},1)=""D"",`$${L}`$4,0))))*N(`$${LL}${ini})"
        $ws.Range("${L}${ini}:${L}${fim}").Formula = $f
        $mil = $xl.International(4)
        $ws.Range("${L}${ini}:${L}${fim}").NumberFormatLocal = "#${mil}##0"
        $ws.Columns.Item($nc).ColumnWidth = 14
        # filtro e painel
        if ($ws.AutoFilterMode) { $ws.AutoFilterMode = $false }
        $null = $ws.Range("A${hr}:${L}${fim}").AutoFilter()
        $null = $ws.Activate()
        $win = $wb.Windows.Item(1)
        $win.FreezePanes = $false; $win.ScrollRow = 1; $win.ScrollColumn = 1; $win.SplitColumn = 0; $win.SplitRow = $hr; $win.FreezePanes = $true
        $tot = $xl.WorksheetFunction.Sum($ws.Range("${L}${ini}:${L}${fim}"))
        Write-Output ("{0}: cabecalho na linha 6, coluna nova {1} (porte {2}, lojas {3}), linhas {4}-{5}, total {6:N0} caixas/mes" -f $nome, $L, $LP, $LL, $ini, $fim, $tot)
    }
    $null = $wb.Worksheets.Item($ativa).Activate()
    $wb.Save()
    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
