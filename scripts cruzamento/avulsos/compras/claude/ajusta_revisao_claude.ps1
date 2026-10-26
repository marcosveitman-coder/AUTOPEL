# Muda o padrao de revisao da metodologia CLAUDE num arquivo ja gravado (26/10/2026: A 7 / B 15 / C 30 dias):
#   - aba param CLAUDE: celulas B5:D5 e a descricao E5;
#   - aba analise compras CLAUDE: textos que citam a revisao (por inicio do texto) e as tabelas 2.3 (curva ABC) e
#     4.1 (cenarios), reescritas com o analise.json gerado de novo por exporta_claude.py.
#   powershell -ExecutionPolicy Bypass -File ajusta_revisao_claude.ps1 -Xlsx C:\...\copia.xlsx -Dir C:\...\saida
param([Parameter(Mandatory = $true)][string]$Xlsx, [Parameter(Mandatory = $true)][string]$Dir)
$ErrorActionPreference = 'Stop'
$JSON = Get-Content -Raw -Encoding UTF8 -Path (Join-Path $Dir 'analise.json') | ConvertFrom-Json
$textos = [ordered]@{
    '2) Proteção = prazo real'                       = '2) Proteção = prazo real do fornecedor (mediana dos pedidos observados; sem histórico, o prometido nas ordens de compra ou 8 dias) + período de revisão (7 dias para a curva A, 15 para a B, 30 para a C).'
    '• No mês vigente a compra CLAUDE'               = '• No mês vigente a compra CLAUDE é maior porque recompõe a curva A, que está com 1,1 mês de estoque contra alvo de cerca de 1,8. Sugestão: dividir entre o fim de outubro e novembro, começando pelos itens em ruptura e de pedido grande.'
    '• O estoque alvo CLAUDE é maior'                = '• Com revisão semanal da curva A (padrão desde 26/10), o estoque alvo CLAUDE fica perto do estoque de hoje: a curva A sobe de 1,1 mês (hoje) para 1,8 de alvo e a C cai de 9,8 meses para 2,7; o excesso da C vai sendo vendido sem compra nova.'
    '• Para gastar menos capital'                    = '• Comprar a curva A toda semana exige mais pedidos, mas são só 132 SKUs. Para menos pedidos, voltar a revisão da A para 15 dias na aba param CLAUDE (alvo maior: cenário de revisão quinzenal na tabela abaixo).'
    '3. Curva A com revisão quinzenal'               = '3. Curva A com revisão semanal (padrão da metodologia desde 26/10) e prioridade no PLANO de AÇAO: 132 SKUs, 80% da venda, 8,8% dos dias em ruptura e 1,1 mês de estoque.'
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
try {
    $wb = $xl.Workbooks.Open($Xlsx, 0, $false)
    $wp = $wb.Worksheets.Item('param CLAUDE')
    if ([string]$wp.Range('A5').Value2 -ne 'Revisão da compra (dias)') { throw 'param CLAUDE mudou de lugar' }
    $wp.Range('B5').Value2 = 7; $wp.Range('C5').Value2 = 15; $wp.Range('D5').Value2 = 30
    $wp.Range('E5').Value2 = 'De quanto em quanto tempo o comprador revisa e compra o item. Curva A: semanal; B: quinzenal; C: mensal.'

    $wa = $wb.Worksheets.Item('analise compras CLAUDE')
    $ult = $wa.UsedRange.Row + $wa.UsedRange.Rows.Count
    $feitos = @{}
    for ($r = 1; $r -le $ult; $r++) {
        $t = [string]$wa.Cells.Item($r, 2).Value2
        foreach ($k in $textos.Keys) {
            if ($t.StartsWith($k)) { $wa.Cells.Item($r, 2).Value2 = $textos[$k]; $feitos[$k] = $r }
        }
    }
    foreach ($k in $textos.Keys) { if (-not $feitos.ContainsKey($k)) { throw "texto nao encontrado: $k" } }

    # tabelas: localiza o cabecalho pela 1a coluna e reescreve as linhas de dados
    $tabelas = @(@('Curva', $JSON.abc), @('Cenário', $JSON.cenarios))
    foreach ($tb in $tabelas) {
        $cab, $dados = $tb
        $r0 = $null
        for ($r = 1; $r -le $ult; $r++) { if ([string]$wa.Cells.Item($r, 2).Value2 -eq $cab) { $r0 = $r; break } }
        if (-not $r0) { throw "tabela $cab nao encontrada" }
        for ($i = 1; $i -lt $dados.Count; $i++) {
            for ($j = 0; $j -lt $dados[$i].Count; $j++) {
                $v = $dados[$i][$j]
                $c = $wa.Cells.Item($r0 + $i, 2 + $j)
                if ($v -is [string]) { $c.Value2 = $v } elseif ($null -ne $v) { $c.Value2 = [double]$v }
            }
        }
        "tabela $cab reescrita na linha $r0 ($($dados.Count - 1) linhas)"
    }
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
