# Grava as abas de resumo do MAPA MERCADO pelo proprio Excel (textos, tabelas e tabelas dinamicas).
# Le a especificacao gerada por gerar_spec.py.  Use numa COPIA do arquivo; feche o arquivo no Excel antes.
#   powershell -ExecutionPolicy Bypass -File td_resumos.ps1 -Xlsx C:\...\copia.xlsx -Spec C:\...\resumos_spec.json
param([Parameter(Mandatory = $true)][string]$Xlsx, [Parameter(Mandatory = $true)][string]$Spec)
$ErrorActionPreference = 'Stop'
$S = Get-Content -Raw -Encoding UTF8 -Path $Spec | ConvertFrom-Json

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
try {
    $wb = $xl.Workbooks.Open($Xlsx, 0, $false)
    foreach ($a in $S.abas) {
        $ws = $wb.Worksheets.Item($a.aba)
        if ($xl.WorksheetFunction.CountA($ws.Cells) -gt 0 -or $ws.PivotTables().Count -gt 0) { throw "aba nao esta vazia: $($a.aba)" }

        foreach ($l in $a.linhas) {
            $c = $ws.Range($l.cel); $c.Value2 = [string]$l.txt
            if ($l.neg) { $c.Font.Bold = $true }
            if ($l.tam) { $c.Font.Size = $l.tam }
        }
        foreach ($t in $a.tabelas) {
            $r0 = $ws.Range($t.cel).Row; $c0 = $ws.Range($t.cel).Column
            for ($i = 0; $i -lt $t.linhas.Count; $i++) {
                $lin = $t.linhas[$i]
                for ($j = 0; $j -lt $lin.Count; $j++) {
                    $cel = $ws.Cells.Item($r0 + $i, $c0 + $j)
                    if ($lin[$j] -is [int] -or $lin[$j] -is [long] -or $lin[$j] -is [double]) {
                        $cel.Value2 = [double]$lin[$j]; $cel.NumberFormat = '#,##0'
                    } else { $cel.Value2 = [string]$lin[$j] }
                    $cel.HorizontalAlignment = $(if ($j -eq 0) { -4131 } else { -4152 })
                }
            }
            $rg = $ws.Range($ws.Cells.Item($r0, $c0), $ws.Cells.Item($r0 + $t.linhas.Count - 1, $c0 + $t.linhas[0].Count - 1))
            $rg.Borders.LineStyle = 1; $rg.Borders.Weight = 2
            $h = $ws.Range($ws.Cells.Item($r0, $c0), $ws.Cells.Item($r0, $c0 + $t.linhas[0].Count - 1))
            $h.Font.Bold = $true; $h.Interior.Color = 14599344; $h.Font.Color = 16777215; $h.WrapText = $true
            if ($t.linhas[-1][0] -eq 'Total') {
                $ws.Range($ws.Cells.Item($r0 + $t.linhas.Count - 1, $c0), $ws.Cells.Item($r0 + $t.linhas.Count - 1, $c0 + $t.linhas[0].Count - 1)).Font.Bold = $true
            }
        }

        $src = $wb.Worksheets.Item($a.fonte_aba).Range($a.fonte_faixa)
        $pc = $wb.PivotCaches().Create(1, $src, 6)      # xlDatabase, versao 2016+
        foreach ($t in $a.tds) {
            $tc = $ws.Range($t.titulo_cel); $tc.Value2 = [string]$t.titulo; $tc.Font.Bold = $true
            $pt = $pc.CreatePivotTable($ws.Range($t.cel), $t.nome)
            $pt.RowAxisLayout(1)                        # tabular: o cabecalho mostra o nome do campo
            $pt.TableStyle2 = 'PivotStyleMedium2'
            $rf = $pt.PivotFields($t.campo); $rf.Orientation = 1
            $sf = $pt.PivotFields($a.campo_status); $sf.Orientation = 2
            $d1 = $pt.AddDataField($pt.PivotFields($a.campo_conta), 'Clientes', -4112)   # campo sempre preenchido, fora da tabela; $d1.NumberFormat = '#,##0'
            $d2 = $pt.AddDataField($pt.PivotFields($a.campo_caixas), ([string][char]0x43 + 'aixas/m' + [char]0xEA + 's'), -4157); $d2.NumberFormat = '#,##0'
            foreach ($it in $sf.PivotItems()) {
                foreach ($p in $a.renomeia.PSObject.Properties) { if ($it.Name -eq $p.Name) { $it.Caption = $p.Value } }
            }
            $sf.AutoSort(2, $sf.Name)                   # decrescente: "ativo" antes de "a trabalhar"
            foreach ($it in $rf.PivotItems()) { if ($it.Name -eq '(blank)' -or $it.Name -eq '(vazio)') { $it.Caption = $a.vazio } }
            if ($t.ordena) { $rf.AutoSort(2, $d2.Name) }
            $ws.Columns.Item($t.col).ColumnWidth = $t.larg
            Write-Output ("{0} {1}: {2} linhas" -f $a.aba, $t.nome, $pt.TableRange1.Rows.Count)
        }
    }
    $wb.Save()
    $wb.Close($false)
}
catch {
    Write-Output ("ERRO na linha {0}: {1}" -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message)
    Write-Output $_.InvocationInfo.Line
    throw
}
finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
