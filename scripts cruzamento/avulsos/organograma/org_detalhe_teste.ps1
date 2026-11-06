# Teste da macro do duplo clique (06/11/2026), numa copia do .xlsm, num Excel oculto proprio:
#  1. DetalheMod.MontarDetalhe em cada celula da CONFIG DETALHE -> grava "celula<TAB>resultado<TAB>linhas na aba DETALHE";
#  2. simula o duplo clique (ThisWorkbook.TesteDuploClique) em 3 celulas;
#  3. fecha (BeforeClose apaga a DETALHE), reabre, salva com a lista aberta e reabre (Workbook_Open apaga a DETALHE).
# Grava o PID do Excel em -PidArq (para encerrar so este Excel se travar).
# Uso: powershell -File org_detalhe_teste.ps1 -Arquivo "<copia.xlsm>" -Saida "<resultado.tsv>" -PidArq "<pid.txt>" [-Celulas "A1,B2"]
param([string]$Arquivo, [string]$Saida, [string]$PidArq, [string]$Celulas = "")
Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr h, out int p);'
$xl = New-Object -ComObject Excel.Application
$p = 0; [void][W.U]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$p); Set-Content $PidArq $p
$xl.Visible = $false; $xl.DisplayAlerts = $false
$out = New-Object System.Collections.Generic.List[string]
function Tem($w) { foreach ($s in $w.Worksheets) { if ($s.Name -eq 'DETALHE') { return $true } }; return $false }
[IO.File]::WriteAllText($Saida, '')
function Add-Linha($s) { $script:out.Add($s); [IO.File]::AppendAllText($Saida, $s + [Environment]::NewLine) }
try {
    $wb = $xl.Workbooks.Open($Arquivo)
    $nome = $wb.Name
    Add-Linha ("#INFO`tHasVBProject=$($wb.HasVBProject)`tCodeName=$($wb.CodeName)`tFileFormat=$($wb.FileFormat)")
    $cfg = $wb.Worksheets.Item('CONFIG DETALHE')
    $lista = @()
    if ($Celulas) { $lista = $Celulas.Split(',') }
    else {
        $ult = $cfg.Cells.Item($cfg.Rows.Count, 1).End(-4162).Row
        $ini = 0
        for ($r = 1; $r -le $ult; $r++) { if ([string]$cfg.Cells.Item($r, 1).Value2 -eq 'CELULAS') { $ini = $r + 2; break } }
        for ($r = $ini; $r -le $ult; $r++) { $v = [string]$cfg.Cells.Item($r, 1).Value2; if ($v) { $lista += $v } }
    }
    $t0 = Get-Date
    foreach ($c in $lista) {
        $t1 = Get-Date
        $res = [string]$xl.Run("'" + $nome + "'!DetalheMod.MontarDetalhe", $c)
        $linhasAba = ''
        try { $det = $wb.Worksheets.Item('DETALHE'); $ur = $det.UsedRange; $linhasAba = $ur.Row + $ur.Rows.Count - 1 - 5 } catch { $linhasAba = 'sem aba' }
        Add-Linha ("$c`t$res`t$linhasAba`t$([int]((Get-Date) - $t1).TotalMilliseconds)")
    }
    Add-Linha ("#TEMPO`t$([int]((Get-Date) - $t0).TotalSeconds)s para $($lista.Count) celulas")
    Add-Linha ("#ALTERADO`tdepois das listas o arquivo esta salvo (Saved): $($wb.Saved)")
    foreach ($c in @($lista | Select-Object -First 3)) {
        $r = [string]$xl.Run("'" + $nome + "'!ThisWorkbook.TesteDuploClique", $c)
        Add-Linha ("#DUPLO`t$c`taba ativa=$($wb.ActiveSheet.Name)`tSaved=$($wb.Saved)")
    }
    $wb.Close($false)                                   # BeforeClose apaga a DETALHE; nada e salvo
    $wb = $xl.Workbooks.Open($Arquivo)
    Add-Linha ("#REABERTO`tHasVBProject=$($wb.HasVBProject)`tDETALHE=$(Tem $wb)")
    $r = [string]$xl.Run("'" + $wb.Name + "'!DetalheMod.MontarDetalhe", $lista[0])
    $wb.Save()                                          # salvar com a lista aberta (Ctrl+S)
    Add-Linha ("#SALVO COM LISTA`t$($lista[0]) -> $r`tDETALHE=$(Tem $wb)")
    $wb.Close($false)
    $wb = $xl.Workbooks.Open($Arquivo)                  # Workbook_Open apaga a DETALHE sem marcar alterado
    Add-Linha ("#REABERTO 2`tDETALHE=$(Tem $wb)`tSaved=$($wb.Saved)`tHasVBProject=$($wb.HasVBProject)")
    $wb.Close($false)
}
catch { Add-Linha ("#ERRO`t$($_.Exception.Message)") }
finally {
    $xl.Quit()
}
