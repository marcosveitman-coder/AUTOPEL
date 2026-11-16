# Grava o STATUS AUTOPEL (16/11/2026) numa COPIA do MAPA MERCADO V2 pelo Excel (COM), a partir do gravar.tsv do
# prepara_gravacao.py (aba, coluna, linha, tipo s/d/h/n, valor):
#  1. valores por coluna em bloco (Range.Value2 com a matriz da coluna); datas como data (dd/mm/aaaa);
#  2. cabecalhos novos das abas FINAL (tipo h): copia o formato da coluna STATUS AUTOPEL (W) e poe o titulo;
#  3. "FARMACIA VF": apaga as colunas I e J (no de lojas Receita e pesquisa; fica a H, considerado);
#  4. atualiza so as dinamicas que usam as abas gravadas (Alessandra Dayan VF, IRTT, FARMACIA VF).
# Uso: powershell -File grava_status.ps1 -Arquivo "<copia.xlsm>" -Tsv "<gravar.tsv>" [-PidArq "<pid.txt>"]
param([Parameter(Mandatory = $true)][string]$Arquivo, [Parameter(Mandatory = $true)][string]$Tsv, [string]$PidArq = "")
$ErrorActionPreference = "Stop"
Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr h, out int p);'
$linhas = [IO.File]::ReadAllLines($Tsv, (New-Object Text.UTF8Encoding $true))
$grupos = @{}
$ordem = New-Object System.Collections.Generic.List[string]
foreach ($l in $linhas) {
    if (-not $l) { continue }
    $c = $l.Split("`t")
    $chave = $c[0] + "`t" + $c[1]
    if (-not $grupos.ContainsKey($chave)) { $grupos[$chave] = New-Object System.Collections.Generic.List[object]; $ordem.Add($chave) }
    $grupos[$chave].Add(@([int]$c[2], $c[3], $(if ($c.Count -gt 4) { $c[4] } else { "" })))
}
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
$pidX = 0; [void][W.U]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$pidX)
if ($PidArq) { Set-Content -Path $PidArq -Value $pidX }
try {
    $wb = $xl.Workbooks.Open((Resolve-Path $Arquivo).Path)
    $xl.Calculation = -4135   # manual durante a gravacao
    $n = 0
    foreach ($chave in $ordem) {
        $aba, $col = $chave.Split("`t")
        $ws = $wb.Worksheets.Item($aba)
        $itens = $grupos[$chave]
        $cab = @($itens | Where-Object { $_[1] -eq 'h' })
        $dados = @($itens | Where-Object { $_[1] -ne 'h' })
        foreach ($h in $cab) {
            # coluna nova: formato da coluna STATUS AUTOPEL (W) da linha do cabecalho ate a ultima linha de dados
            $ult = ($dados | ForEach-Object { $_[0] } | Measure-Object -Maximum).Maximum
            if (-not $ult) { $ult = ($grupos[$aba + "`tW"] | ForEach-Object { $_[0] } | Measure-Object -Maximum).Maximum }
            $null = $ws.Range("W$($h[0]):W$ult").Copy($ws.Range("$col$($h[0])"))
            $ws.Range("$col$($h[0])").Value2 = [string]$h[2]
        }
        if ($dados.Count -eq 0) { continue }
        if ($ws.FilterMode) {
            # aba com filtro ligado (ex.: IRTT): gravar a coluna em bloco pula as linhas ocultas -> celula por celula
            foreach ($d in $dados) {
                $cel = $ws.Range("$col$($d[0])")
                if ($d[1] -eq 'd') {
                    if ($d[2]) { $cel.Value2 = [DateTime]::ParseExact($d[2], 'yyyy-MM-dd', $null).ToOADate(); $cel.NumberFormatLocal = 'dd/mm/aaaa' }
                    else { $cel.Value2 = $null }
                } elseif ($d[1] -eq 'n') { $cel.Value2 = [double]::Parse($d[2], [Globalization.CultureInfo]::InvariantCulture) }
                else { $cel.Value2 = $(if ($d[2]) { [string]$d[2] } else { $null }) }
            }
            $n += $dados.Count
            continue
        }
        $r1 = ($dados | ForEach-Object { $_[0] } | Measure-Object -Minimum).Minimum
        $r2 = ($dados | ForEach-Object { $_[0] } | Measure-Object -Maximum).Maximum
        $rng = $ws.Range("$col$($r1):$col$r2")
        if ($r1 -eq $r2) { $vals = New-Object 'object[,]' 1, 1; $vals[0, 0] = $rng.Value2; $base = 0 }
        else { $vals = $rng.Value2; $base = 1 }
        $temData = $false
        foreach ($d in $dados) {
            $i = $d[0] - $r1 + $base
            if ($d[1] -eq 'd') {
                $temData = $true
                $vals[$i, $base] = $(if ($d[2]) { [DateTime]::ParseExact($d[2], 'yyyy-MM-dd', $null).ToOADate() } else { $null })
            } elseif ($d[1] -eq 'n') { $vals[$i, $base] = [double]::Parse($d[2], [Globalization.CultureInfo]::InvariantCulture) }   # tipo n: numero
            else { $vals[$i, $base] = $(if ($d[2]) { [string]$d[2] } else { $null }) }
        }
        if ($r1 -eq $r2) { $rng.Value2 = $vals[0, 0] } else { $rng.Value2 = $vals }
        if ($temData) { $rng.NumberFormatLocal = 'dd/mm/aaaa' }
        $n += $dados.Count
    }
    # larguras das colunas novas das abas FINAL
    foreach ($chave in $ordem) {
        $aba, $col = $chave.Split("`t")
        if (@($grupos[$chave] | Where-Object { $_[1] -eq 'h' }).Count -gt 0) {
            $ws = $wb.Worksheets.Item($aba)
            $t = [string]($grupos[$chave] | Where-Object { $_[1] -eq 'h' } | Select-Object -First 1)[2]
            $ws.Columns.Item($col).ColumnWidth = $(if ($t -like 'COMO*') { 70 } elseif ($t -like 'VENDEDOR*') { 32 } else { 14 })
        }
    }
    # FARMACIA VF: so a coluna H de lojas (apaga I = Receita e J = pesquisa)
    $wf = $wb.Worksheets.Item('"FARMACIA VF"')
    $hI = [string]$wf.Range("I6").Value2; $hJ = [string]$wf.Range("J6").Value2
    # 2a gravacao em diante (I e J ja apagadas): nao apaga nada
    $apagou = $false
    if ($hI -like '*RECEITA*' -and $hJ -like '*PESQUISA*') { $null = $wf.Range("I:J").EntireColumn.Delete(); $apagou = $true }
    $xl.Calculation = -4105   # automatico
    $xl.Calculate()
    foreach ($nome in @('Alessandra Dayan VF', 'TAB DIN. I.R.T.T.', 'I.R.T. T. resumo', '"FARMACIA VF RESUMO"', '"FARMACIA VF RESUMO" port A e B')) {
        $ws = $wb.Worksheets.Item($nome)
        foreach ($pt in $ws.PivotTables()) { $null = $pt.RefreshTable() }
    }
    $wb.Save()
    $wb.Close($false)
    Write-Output ("celulas gravadas: {0} | I:J da FARMACIA VF apagadas agora: {1} ({2} | {3}) | dinamicas atualizadas" -f $n, $apagou, $hI, $hJ)
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
