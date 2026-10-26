# Grava a metodologia CLAUDE no "Relatorio Revenda atualizado.xlsx" pelo proprio Excel (usar numa COPIA fechada).
#   - aba "param CLAUDE": parametros editaveis (revisao, nivel de servico, cobertura maxima, padroes) e a tabela
#     ParamClaude, um SKU por linha (gerada por exporta_claude.py: param_claude.csv);
#   - Table1 da aba Geral: 6 colunas pedidas (CLAUDE mes vigente / proximo / subsequente, QTD e R$) e 13 auxiliares,
#     todas formula; o filtro da tabela e tirado para gravar e devolvido no fim;
#   - aba "analise compras CLAUDE" (ja criada pelo usuario, vazia): diagnostico, metodologia, comparacao TOTVS x
#     CLAUDE por comprador (formulas), cenarios, qualidade dos dados e sugestoes (analise.json).
#   powershell -ExecutionPolicy Bypass -File grava_claude.ps1 -Xlsx C:\...\copia.xlsx -Dir C:\...\saida
param([Parameter(Mandatory = $true)][string]$Xlsx, [Parameter(Mandatory = $true)][string]$Dir)
$ErrorActionPreference = 'Stop'
$JS = Get-Content -Raw -Encoding UTF8 -Path (Join-Path $Dir 'analise.json') | ConvertFrom-Json
$CSV = Import-Csv -Path (Join-Path $Dir 'param_claude.csv') -Delimiter ';' -Encoding UTF8
$INV = [Globalization.CultureInfo]::InvariantCulture

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
try {
    $wb = $xl.Workbooks.Open($Xlsx, 0, $false)
    $xl.Calculation = -4135
    $mil = $xl.International(4); $dec = $xl.International(3)
    $FI = "#${mil}##0"; $FD1 = "0${dec}0"; $FD2 = "0${dec}00"; $FP = "0${dec}0%"; $FR = "`"R`$`" #${mil}##0"

    $ws = $null
    foreach ($s in $wb.Worksheets) { foreach ($t in $s.ListObjects) { if ($t.Name -eq 'Table1') { $ws = $s } } }
    $lo = $ws.ListObjects.Item('Table1')
    $nm = @{}; foreach ($lc in $lo.ListColumns) { $nm[$lc.Index] = $lc.Name }
    $esp = @{ 1 = 'Comprador'; 2 = 'Codigo      '; 5 = ' PontoPedido'; 7 = ' Ult.Preco'; 9 = ' MedContr'; 20 = ' Prc_Compra.'; 12 = ' Vlr Estoque' }
    foreach ($k in $esp.Keys) { if ($nm[$k] -ne $esp[$k]) { throw "coluna $k da Table1 mudou: '$($nm[$k])'" } }
    foreach ($n in @('Saldo hoje (Disp-VendProt+Comprado)', 'Compra mês vigente QTD', 'Compra mês vigente R$')) {
        if ($nm.Values -notcontains $n) { throw "falta a coluna '$n' (revisao anterior)" }
    }
    if ($nm.Values -contains 'CLAUDE mês vigente QTD') { throw 'Table1 ja tem as colunas CLAUDE' }
    foreach ($s in $wb.Worksheets) { if ($s.Name -eq 'param CLAUDE') { throw 'aba param CLAUDE ja existe' } }
    $wa = $wb.Worksheets.Item('analise compras CLAUDE')
    if ($xl.WorksheetFunction.CountA($wa.Cells) -gt 0) { throw 'aba analise compras CLAUDE nao esta vazia' }

    # filtro da tabela (gravar coluna com filtro ligado pula as linhas ocultas)
    $filtros = @(); $visiveis = $null
    if ($lo.ShowAutoFilter -and $lo.AutoFilter.FilterMode) {
        $visiveis = $lo.DataBodyRange.Columns.Item(1).SpecialCells(12).Count
        $fs = $lo.AutoFilter.Filters
        for ($i = 1; $i -le $fs.Count; $i++) { $fl = $fs.Item($i); if ($fl.On) { $filtros += , @($i, $fl.Criteria1, $fl.Operator) } }
        $lo.AutoFilter.ShowAllData()
    }

    # --- aba param CLAUDE -------------------------------------------------------------------------------------------
    $wp = $wb.Worksheets.Add([Type]::Missing, $wa); $wp.Name = 'param CLAUDE'
    $wp.Cells.Font.Name = 'Calibri'; $wp.Cells.Font.Size = 11
    $wp.Range('A1').Value2 = 'PARÂMETROS DA METODOLOGIA CLAUDE'; $wp.Range('A1').Font.Bold = $true; $wp.Range('A1').Font.Size = 14
    $wp.Range('A2').Value2 = 'Células amarelas podem ser alteradas: as colunas CLAUDE da aba Geral recalculam sozinhas. A tabela por SKU (abaixo) é gerada pelos scripts e deve ser refeita 1 vez por mês.'
    $g = @(
        @('Parâmetro', 'Curva A', 'Curva B', 'Curva C', 'O que é'),
        @('Revisão da compra (dias)', 15, 30, 30, 'De quanto em quanto tempo o comprador revisa e compra o item. Curva A: quinzenal.'),
        @('Nível de serviço', 0.95, 0.90, 0.80, 'Chance de não faltar entre uma compra e a chegada da seguinte.'),
        @('Fator z (calculado)', '=NORM.S.INV(B6)', '=NORM.S.INV(C6)', '=NORM.S.INV(D6)', 'Quantos desvios de venda o estoque de segurança cobre.'),
        @('', '', '', '', ''),
        @('Cobertura máxima (meses de previsão)', 4, '', '', 'Teto do estoque alvo de qualquer SKU.'),
        @('Prazo padrão do fornecedor (dias)', 8, '', '', 'Usado quando o fornecedor não tem histórico (mediana dos pedidos observados).'),
        @('Desvio padrão do prazo (dias)', 5.9, '', '', 'Usado quando o fornecedor não tem histórico (mediana dos desvios por fornecedor).'),
        @('Variação da demanda (CV) padrão', 0.5, '', '', 'Usado quando o SKU não está na tabela abaixo.')
    )
    for ($i = 0; $i -lt $g.Count; $i++) {
        for ($j = 0; $j -lt 5; $j++) {
            $v = $g[$i][$j]; $c = $wp.Cells.Item(4 + $i, 1 + $j)
            if ([string]$v -like '=*') { $c.Formula = $v } elseif ($v -is [string]) { $c.Value2 = $v } else { $c.Value2 = [double]$v }
        }
    }
    $wp.Range('A4:E4').Font.Bold = $true; $wp.Range('A4:E4').Interior.Color = 14599344; $wp.Range('A4:E4').Font.Color = 16777215
    foreach ($a in @('B5:D6', 'B9:B12')) { $wp.Range($a).Interior.Color = 10284031 }
    $wp.Range('B6:D6').NumberFormatLocal = "0%"; $wp.Range('B7:D7').NumberFormatLocal = $FD2
    $wp.Range('A4:E7').Borders.LineStyle = 1; $wp.Range('A9:E12').Borders.LineStyle = 1
    $nomes = @{ 'CLAUDE_R_A' = 'B5'; 'CLAUDE_R_B' = 'C5'; 'CLAUDE_R_C' = 'D5'; 'CLAUDE_Z_A' = 'B7'; 'CLAUDE_Z_B' = 'C7'; 'CLAUDE_Z_C' = 'D7'
               'CLAUDE_COB_MAX' = 'B9'; 'CLAUDE_LT_PADRAO' = 'B10'; 'CLAUDE_DLT_PADRAO' = 'B11'; 'CLAUDE_CV_PADRAO' = 'B12' }
    foreach ($k in $nomes.Keys) {
        $c = $nomes[$k]; $ref = "='param CLAUDE'!`$" + $c.Substring(0, 1) + "`$" + $c.Substring(1)
        $wb.Names.Add($k, $ref) | Out-Null
    }
    $wp.Range('A14').Value2 = 'Parâmetros por SKU (scripts cruzamento/avulsos/compras/claude: exporta_claude.py; S01..S12 = fator sazonal de jan a dez, média 1)'
    $wp.Range('A14').Font.Bold = $true
    $cols = $CSV[0].PSObject.Properties.Name
    $txt = @('Codigo', 'Classe', 'ABC', 'XYZ', 'PQR', 'Fonte sazonalidade', 'Fonte LT', 'Fornecedor principal', 'Faturamento cobre a venda')
    $nl = $CSV.Count; $nc = $cols.Count
    $arr = New-Object 'object[,]' ($nl + 1), $nc
    for ($j = 0; $j -lt $nc; $j++) { $arr[0, $j] = $cols[$j] }
    for ($i = 0; $i -lt $nl; $i++) {
        $r = $CSV[$i]
        for ($j = 0; $j -lt $nc; $j++) {
            $v = $r.($cols[$j])
            if ($txt -contains $cols[$j] -or $v -eq '') { $arr[($i + 1), $j] = [string]$v }
            else { $arr[($i + 1), $j] = [double]::Parse($v, $INV) }
        }
    }
    $ini = 16
    $rg = $wp.Range($wp.Cells.Item($ini, 1), $wp.Cells.Item($ini + $nl, $nc))
    $wp.Range($wp.Cells.Item($ini + 1, 1), $wp.Cells.Item($ini + $nl, 1)).NumberFormat = '@'
    "tabela de parametros: $($arr.GetLength(0)) x $($arr.GetLength(1)) -> $($rg.Address(0, 0)) ($($rg.GetType().Name))"
    try { $rg.Value = $arr }
    catch {
        # gravacao em blocos de 500 linhas (alguns ambientes recusam a matriz inteira)
        "matriz inteira recusada ($($_.Exception.Message)); gravando em blocos"
        for ($b0 = 0; $b0 -le $nl; $b0 += 500) {
            $n = [math]::Min(500, $nl + 1 - $b0)
            $bl = New-Object 'object[,]' $n, $nc
            for ($i = 0; $i -lt $n; $i++) { for ($j = 0; $j -lt $nc; $j++) { $bl[$i, $j] = $arr[($b0 + $i), $j] } }
            $wp.Range($wp.Cells.Item($ini + $b0, 1), $wp.Cells.Item($ini + $b0 + $n - 1, $nc)).Value = $bl
        }
    }
    $tp = $wp.ListObjects.Add(1, $rg, [Type]::Missing, 1); $tp.Name = 'ParamClaude'; $tp.TableStyle = 'TableStyleLight9'
    $wp.Columns.Item(1).ColumnWidth = 36; foreach ($j in 2..5) { $wp.Columns.Item($j).ColumnWidth = 12 }; $wp.Columns.Item(5).ColumnWidth = 14
    $wp.Range('E5:E12').WrapText = $false
    foreach ($j in 6..$nc) { $wp.Columns.Item($j).ColumnWidth = 10 }
    $wp.Columns.Item($nc - 1).ColumnWidth = 40

    # --- colunas CLAUDE na Table1 -----------------------------------------------------------------------------------
    function L([string]$n) { '[@[' + $n + ']]' }
    $COD = L $nm[2]; $MED = L $nm[9]; $PP = L $nm[5]; $PRC = L $nm[20]; $ULT = L $nm[7]; $COMP = L $nm[1]
    $SALDO = L 'Saldo hoje (Disp-VendProt+Comprado)'
    $CN = @{ q0 = 'CLAUDE mês vigente QTD'; r0 = 'CLAUDE mês vigente R$'; q1 = 'CLAUDE próximo mês QTD'; r1 = 'CLAUDE próximo mês R$'
            q2 = 'CLAUDE mês subsequente QTD'; r2 = 'CLAUDE mês subsequente R$'
            f0 = 'CLAUDE previsão mês vigente QTD'; f1 = 'CLAUDE previsão próximo mês QTD'; f2 = 'CLAUDE previsão mês subsequente QTD'
            a0 = 'CLAUDE alvo mês vigente'; a1 = 'CLAUDE alvo próximo mês'; a2 = 'CLAUDE alvo mês subsequente'
            ss = 'CLAUDE estoque segurança'; cob = 'CLAUDE cobertura alvo (dias)'; cl = 'CLAUDE classe (ABC-XYZ-PQR)'
            lt = 'CLAUDE LT fornecedor (dias)'; dlt = 'CLAUDE desvio LT (dias)'; cv = 'CLAUDE variação demanda (CV)'; mu = 'CLAUDE múltiplo de compra' }
    $INAT = "LEFT($COMP,7)=""INATIVO"""
    $PRECO = "IF($PRC>0,$PRC,$ULT)"
    function Prev([int]$k) { "=LET(sz,XLOOKUP($COD,ParamClaude[Codigo],ParamClaude[[S01]:[S12]],""""),temsz,ISNUMBER(INDEX(sz,1)),basesz,IF(temsz,(INDEX(sz,MONTH(EDATE(TODAY(),-1)))+INDEX(sz,MONTH(EDATE(TODAY(),-2)))+INDEX(sz,MONTH(EDATE(TODAY(),-3))))/3,1),fatm,IF(temsz,INDEX(sz,MONTH(EDATE(TODAY(),$k))),1),MAX(0,$MED)*fatm/basesz)" }
    function Seg([string]$f) { "curva,LEFT($(L $CN.cl),1),revis,IF(curva=""A"",CLAUDE_R_A,IF(curva=""B"",CLAUDE_R_B,CLAUDE_R_C)),zval,IF(curva=""A"",CLAUDE_Z_A,IF(curva=""B"",CLAUDE_Z_B,CLAUDE_Z_C)),prot,$(L $CN.lt)+revis,varcv,$(L $CN.cv),dlt,$(L $CN.dlt),prevm,$(L $f),segur,zval*SQRT(prot/30*(varcv*prevm)^2+(prevm/30*dlt)^2)" }
    function Alvo([string]$f) { "=LET($(Seg $f),MIN(prevm*prot/30+segur,prevm*CLAUDE_COB_MAX))" }
    $ARR = 'IF(falta<=0,0,ROUNDUP(falta/mult,0)*mult)'
    $DIAS = 'dm,DAY(EOMONTH(TODAY(),0)),dr,dm-DAY(TODAY())'
    $defs = [ordered]@{}
    $defs[$CN.q0] = "=LET(ponto,$PP,saldo,$SALDO,mult,MAX(1,$(L $CN.mu)),falta,MAX(0,$(L $CN.a0)-saldo),IF($INAT,0,IF(ponto=0,ROUNDUP(MAX(0,-saldo),0),$ARR)))"
    $defs[$CN.r0] = "=$(L $CN.q0)*$PRECO"
    $defs[$CN.q1] = "=LET(ponto,$PP,saldo,$SALDO,mult,MAX(1,$(L $CN.mu)),$DIAS,saldo1,saldo+$(L $CN.q0)-$(L $CN.f0)*dr/dm,falta,MAX(0,$(L $CN.a1)-saldo1),IF(OR($INAT,ponto=0),0,$ARR))"
    $defs[$CN.r1] = "=$(L $CN.q1)*$PRECO"
    $defs[$CN.q2] = "=LET(ponto,$PP,saldo,$SALDO,mult,MAX(1,$(L $CN.mu)),$DIAS,saldo1,saldo+$(L $CN.q0)-$(L $CN.f0)*dr/dm,saldo2,saldo1+$(L $CN.q1)-$(L $CN.f1),falta,MAX(0,$(L $CN.a2)-saldo2),IF(OR($INAT,ponto=0),0,$ARR))"
    $defs[$CN.r2] = "=$(L $CN.q2)*$PRECO"
    $defs[$CN.f0] = Prev 0
    $defs[$CN.f1] = Prev 1
    $defs[$CN.f2] = Prev 2
    $defs[$CN.a0] = Alvo $CN.f0
    $defs[$CN.a1] = Alvo $CN.f1
    $defs[$CN.a2] = Alvo $CN.f2
    $defs[$CN.ss] = "=LET($(Seg $CN.f0),segur)"
    $defs[$CN.cob] = "=IF($(L $CN.f0)>0,$(L $CN.a0)/($(L $CN.f0)/30),0)"
    $defs[$CN.cl] = "=XLOOKUP($COD,ParamClaude[Codigo],ParamClaude[Classe],""sem parâmetro"")"
    $defs[$CN.lt] = "=XLOOKUP($COD,ParamClaude[Codigo],ParamClaude[[LT (dias)]],CLAUDE_LT_PADRAO)"
    $defs[$CN.dlt] = "=XLOOKUP($COD,ParamClaude[Codigo],ParamClaude[[Desvio LT (dias)]],CLAUDE_DLT_PADRAO)"
    $defs[$CN.cv] = "=XLOOKUP($COD,ParamClaude[Codigo],ParamClaude[[CV usado]],CLAUDE_CV_PADRAO)"
    $defs[$CN.mu] = "=XLOOKUP($COD,ParamClaude[Codigo],ParamClaude[Multiplo],1)"
    # cria todas as colunas primeiro (as formulas citam umas as outras) e depois grava. Nao usa ListColumns.Add: ele
    #   insere celulas e o Excel recusa quando ha formatacao ate o fim da planilha; escreve os cabecalhos a direita
    #   da tabela (colunas vazias) e estende a tabela sobre elas
    $c1 = $lo.Range.Column + $lo.Range.Columns.Count; $l1 = $lo.Range.Row; $l2 = $l1 + $lo.Range.Rows.Count - 1
    $area = $ws.Range($ws.Cells.Item($l1, $c1), $ws.Cells.Item($l2, $c1 + $defs.Count - 1))
    if ($xl.WorksheetFunction.CountA($area) -gt 0) { throw "colunas a direita da Table1 nao estao vazias: $($area.Address(0, 0))" }
    $j = 0; foreach ($n in $defs.Keys) { $ws.Cells.Item($l1, $c1 + $j).Value2 = [string]$n; $j++ }
    $lo.Resize($ws.Range($ws.Cells.Item($l1, $lo.Range.Column), $ws.Cells.Item($l2, $c1 + $defs.Count - 1)))
    foreach ($n in $defs.Keys) { if ($lo.ListColumns.Item($n).Name -ne $n) { throw "coluna $n nao criada" } }
    $fmtQ = $lo.ListColumns.Item('Compra mês vigente QTD').DataBodyRange.Cells.Item(1, 1).NumberFormat
    $fmtR = $lo.ListColumns.Item('Compra mês vigente R$').DataBodyRange.Cells.Item(1, 1).NumberFormat
    foreach ($n in $defs.Keys) {
        $rgc = $lo.ListColumns.Item($n).DataBodyRange
        $rgc.Formula = $defs[$n]
        if ($n -like '*R$') { $rgc.NumberFormat = $fmtR }
        elseif ($n -eq $CN.cv) { $rgc.NumberFormatLocal = $FD2 }
        elseif ($n -eq $CN.cl) { }
        else { $rgc.NumberFormatLocal = $FI }
    }
    foreach ($n in $defs.Keys) { $lo.ListColumns.Item($n).Range.EntireColumn.ColumnWidth = $(if ($n -like '*R$') { 19 } else { 15 }) }
    foreach ($p in @(@($CN.q0, 'AX1'), @($CN.r0, 'AY1'), @($CN.q1, 'AX1'), @($CN.r1, 'AY1'), @($CN.q2, 'AX1'), @($CN.r2, 'AY1'))) {
        $col = $lo.ListColumns.Item($p[0]).Range.Column
        $ws.Range($p[1]).Copy($ws.Cells.Item(1, $col)) | Out-Null
    }

    # (nomes longos de proposito: o PowerShell nao diferencia maiusculas e $j, $v, $t sao usados em lacos)
    # --- aba analise compras CLAUDE -------------------------------------------------------------------------------
    $wa.Cells.Font.Name = 'Calibri'; $wa.Cells.Font.Size = 11
    $wa.Columns.Item(1).ColumnWidth = 2; $wa.Columns.Item(2).ColumnWidth = 62
    foreach ($j in 3..12) { $wa.Columns.Item($j).ColumnWidth = 16 }
    $script:r = 1
    function Tit([string]$t, [int]$tam = 13) { $c = $wa.Cells.Item($script:r, 2); $c.Value2 = $t; $c.Font.Bold = $true; $c.Font.Size = $tam; $script:r++ }
    function Par([string]$t, [int]$linhas = 1) {
        $m = $wa.Range($wa.Cells.Item($script:r, 2), $wa.Cells.Item($script:r, 12)); $m.Merge(); $m.Value2 = $t
        $m.WrapText = $true; $m.VerticalAlignment = -4160; $m.RowHeight = 15 * $linhas; $script:r++
    }
    function Linhas([string]$t) { [math]::Max(1, [math]::Ceiling($t.Length / 205)) }
    function Tab($dados, $fmts) {
        $r0 = $script:r; $nlin = $dados.Count; $ncol = $dados[0].Count
        $a = New-Object 'object[,]' $nlin, $ncol
        for ($i = 0; $i -lt $nlin; $i++) {
            for ($j = 0; $j -lt $ncol; $j++) {
                $v = $dados[$i][$j]
                # numero inteiro do JSON vem como Int64, que o Excel nao aceita: vira double
                $a[$i, $j] = $(if ($null -eq $v) { '' } elseif ($v -is [long] -or $v -is [int] -or $v -is [decimal]) { [double]$v }
                               elseif ($i -gt 0 -and $v -is [string] -and $v -match '^\d{2}/\d{2}/\d{4}$') { [datetime]::ParseExact($v, 'dd/MM/yyyy', $INV) }
                               else { $v })
            }
        }
        $rgt = $wa.Range($wa.Cells.Item($r0, 2), $wa.Cells.Item($r0 + $nlin - 1, 1 + $ncol)); $rgt.Value = $a
        $rgt.Borders.LineStyle = 1; $rgt.Borders.Weight = 2
        $h = $wa.Range($wa.Cells.Item($r0, 2), $wa.Cells.Item($r0, 1 + $ncol)); $h.Font.Bold = $true; $h.Interior.Color = 14599344; $h.Font.Color = 16777215; $h.WrapText = $true
        for ($j = 1; $j -lt $ncol; $j++) {
            $f = $fmts[$j]
            if ($f) { $wa.Range($wa.Cells.Item($r0 + 1, 2 + $j), $wa.Cells.Item($r0 + $nlin - 1, 2 + $j)).NumberFormatLocal = $f }
        }
        if ([string]$dados[$nlin - 1][0] -eq 'Total') { $wa.Range($wa.Cells.Item($r0 + $nlin - 1, 2), $wa.Cells.Item($r0 + $nlin - 1, 1 + $ncol)).Font.Bold = $true }
        $script:r = $r0 + $nlin + 1
    }
    function Txt($lista) { foreach ($t in $lista) { Par $t (Linhas $t) } }

    Tit 'ANÁLISE DE COMPRAS — CLAUDE (dados de 26/10/2026)' 16
    Par 'Base: 56 fotos diárias do Relatório de Revenda (27/07 a 26/10), 21 relatórios do Protheus (13/08 a 26/10), histórico de entradas e ordens de compra do TOTVS, exportações e telas do Power BI, 22 boletins de ruptura, a planilha de rupturas (evolução) e o MRP em HTML de 23/10. Os números das seções 2, 5 e dos cenários são da data acima; a seção 4 é fórmula e atualiza com a Geral.' 2
    $script:r++
    Tit '1. COMO FUNCIONA HOJE (o que entendi)'
    Txt @(
        '• O TOTVS calcula por SKU: MedContr = média da venda dos 3 últimos meses completos (MES_01 a MES_03 do relatório do Protheus) e PontoPedido = 2 × MedContr (95% dos SKUs; 266 estão zerados à mão). Comprar = PontoPedido − (Disponível − VendProt + Comprado).',
        '• O Relatório de Revenda é salvo todo dia (aba Geral, colunas A a X). Esta planilha calcula a programação do mês vigente e dos 2 seguintes, alimenta o PLANO de AÇAO (impresso para cada comprador) e o tracking da aba dinamicas (hoje × dias anteriores).',
        '• O relatório do Protheus (estoque, necessidade, fornecedor, venda dos 3 meses) é acompanhado todo dia, e o boletim diário de ruptura registra o que foi combinado: data de chegada, chamado, retorno. A planilha "rupturas (evolução)" junta os boletins dia a dia.',
        '• O Power BI "Necessidade de Compras" mostra ruptura, comprar e crítico por comprador, follow-up dos pedidos (atrasados), agenda do dia e de amanhã, últimas entradas e custos, pedidos de venda, mix e mobiliário. O painel de indicadores tem a curva (ex.: CAA+) e os pedidos dos grandes clientes (CEF, RD, FDE, PGMN...).',
        '• O MRP em HTML (23/10) propõe política com ABC/XYZ/PQR, nível de serviço por curva e prazo fixo de 15 dias. Ele usa as bases de faturamento como demanda, e elas só têm parte da venda do TOTVS (seção 5).',
        '• Objetivo do plano de ação: o comprador programa o mês e os 2 seguintes, liga menos para o fornecedor e ganha tempo para negociar e buscar novos fornecedores.'
    )
    $script:r++
    Tit '2. DIAGNÓSTICO'
    Txt @(
        '• Ruptura: entre 45 e 151 SKUs por dia (média 93) de 27/07 a 26/10; hoje 45, o menor nível do período.',
        '• 30% dos episódios de ruptura são de itens com PontoPedido 0 (venda sob encomenda): não é falha de estoque, mas pesa no indicador e no boletim.',
        '• 28% das rupturas começam com o pedido de compra já a caminho (comprou tarde) e 10% com o TOTVS pedindo compra havia 3 dias ou mais sem pedido: quase 4 em cada 10 dependem do tempo de reação.',
        '• 19% vêm de um pedido de venda grande (mais de 1 mês de venda de uma vez). Estoque de segurança não cobre isso: precisa de aviso da carteira e dos contratos.',
        '• A curva A (132 SKUs, 80% da venda em R$) ficou 8,8% dos dias em ruptura e roda com 1,1 mês de estoque; a curva C (3.743 SKUs) tem 9,8 meses de venda em estoque (R$ 8,4 mi). R$ 11 mi (37% do estoque) estão em SKUs com mais de 4 meses de cobertura ou sem venda.',
        '• Prazo real do pedido à chegada: mediana de 7 dias e 90% em até 21 dias; só 36% chegam na primeira data prometida e 49% têm a data remarcada. No boletim, 39% das rupturas se resolvem até a data combinada.',
        '• A previsão pela MedContr erra de 29% a 54% do volume no mês (venda B2B irregular). Outras médias testadas não ficaram melhores: o erro tem que ser coberto pelo estoque de segurança, não por outra média.'
    )
    $script:r++
    Tit '2.1 Causas das rupturas (episódios de 27/07 a 26/10)'
    Tab $JS.causas @($null, $FI, $FP, $FR)
    Tit '2.2 Rupturas por comprador'
    Tab $JS.rup_comprador @($null, $FI, $FD1, $FI, $FI)
    Tit '2.3 Curva ABC: estoque, alvo e ruptura'
    Tab $JS.abc @($null, $FI, $FR, $FR, $FD2, $FD2, $FD2, $FP, $FD2)
    Tit '2.4 Estoque por cobertura (26/10)'
    Tab $JS.cobertura @($null, $FI, $FR, $FP)
    Tit '2.5 Prazos de entrega e pontualidade'
    $pz = $JS.prazos
    Tab $pz @($null, $null)
    foreach ($i in 1..($pz.Count - 1)) { $v = $pz[$i][1]; $c = $wa.Cells.Item($script:r - $pz.Count - 1 + $i, 3); if (($v -is [double] -or $v -is [decimal]) -and $v -lt 1) { $c.NumberFormatLocal = $FP } else { $c.NumberFormatLocal = $FI } }
    Tit '2.6 Fornecedores (10+ pedidos com chegada nas fotos, por R$)'
    Tab $JS.fornecedores @($null, $FI, $FD1, $FI, $FP, $FP, $FR)
    Tit '2.7 Erro da previsão do mês (venda real do TOTVS)'
    Tab $JS.previsao @($null, $FP, $FP, $FP)
    $rSerie = $script:r
    Tit '2.8 Rupturas e compra por dia'
    $rTab = $script:r
    Tab $JS.serie @($null, $FI, $FI, $FR, $FR, $FI)
    $ch = $wa.ChartObjects().Add($wa.Cells.Item($rTab, 9).Left, $wa.Cells.Item($rTab, 9).Top, 620, 300)
    $ch.Chart.ChartType = 4
    $ch.Chart.SetSourceData($wa.Range($wa.Cells.Item($rTab, 2), $wa.Cells.Item($rTab + $JS.serie.Count - 1, 3)))
    $ch.Chart.HasTitle = $true; $ch.Chart.ChartTitle.Text = 'SKUs em ruptura por dia'
    $script:r++

    Tit '3. METODOLOGIA CLAUDE (aba Geral, colunas CM em diante; parâmetros na aba param CLAUDE)'
    Txt @(
        '1) Previsão do mês = MedContr × fator sazonal do mês ÷ fator médio dos 3 meses que formam a MedContr. A sazonalidade vem da mediana de 2023 a 2025, do grupo e do próprio SKU quando o faturamento cobre a venda, e entra amortecida (metade do efeito). Hoje: outubro +18%, novembro +25%, dezembro +9% sobre a MedContr.',
        '2) Proteção = prazo real do fornecedor (mediana dos pedidos observados; sem histórico, o prometido nas ordens de compra ou 8 dias) + período de revisão (15 dias para a curva A, 30 para B e C).',
        '3) Estoque de segurança = z × raiz( proteção/30 × (CV × previsão)² + (previsão/30 × desvio do prazo)² ): cobre a variação da venda e o atraso do fornecedor. z sai do nível de serviço: A 95%, B 90%, C 80%.',
        '4) Estoque alvo = previsão × proteção/30 + segurança, no máximo 4 meses da previsão.',
        '5) Compra do mês vigente = alvo − saldo de hoje (Disponível − VendProt + Comprado). Próximo mês: saldo projetado = saldo + compra − venda prevista nos dias que faltam do mês; compra = alvo do próximo mês − saldo projetado. Mês subsequente: igual, com 1 mês inteiro de venda prevista.',
        '6) Arredonda para cima no múltiplo de compra (lote histórico, até 1 mês de venda) e nunca fica negativo. Regra mantida: PontoPedido 0 = sem estoque, só compra o que já foi vendido. Comprador INATIVO = zero.',
        'Diferenças para a regra atual: o alvo varia por SKU conforme prazo e variação da venda (não é 2 meses fixos); entra a sazonalidade; o mês vigente considera só os dias que faltam; respeita o lote; protege mais a curva A e menos a C.'
    )
    $script:r++
    Tit '4. REGRA ATUAL × CLAUDE, POR COMPRADOR (fórmulas: atualiza com a Geral)'
    $compradores = @()
    $vals = $lo.ListColumns.Item($nm[1]).DataBodyRange.Value2
    $set = New-Object 'System.Collections.Generic.SortedSet[string]'
    for ($i = 1; $i -le $vals.GetLength(0); $i++) { $v = [string]$vals[$i, 1]; if ($v.Trim()) { [void]$set.Add($v) } }
    $compradores = @($set)
    $cab = @('Comprador', 'SKUs a comprar no mês vigente: regra atual', 'SKUs a comprar no mês vigente: CLAUDE', 'Mês vigente R$: regra atual', 'Mês vigente R$: CLAUDE',
             'Próximo mês R$: regra atual', 'Próximo mês R$: CLAUDE', 'Mês subsequente R$: regra atual', 'Mês subsequente R$: CLAUDE', 'Estoque hoje R$', 'Estoque alvo R$: regra atual', 'Estoque alvo R$: CLAUDE')
    $r0 = $script:r
    for ($j = 0; $j -lt $cab.Count; $j++) { $wa.Cells.Item($r0, 2 + $j).Value2 = $cab[$j] }
    $TB = 'Table1'
    $CMP = "$TB[[$($nm[1])]]"
    $PRV = "IF($TB[[$($nm[20])]]>0,$TB[[$($nm[20])]],$TB[[$($nm[7])]])"
    for ($i = 0; $i -lt $compradores.Count; $i++) {
        $rr = $r0 + 1 + $i; $x = "`$B$rr"
        $wa.Cells.Item($rr, 2).Value2 = $compradores[$i]
        $f = @("=COUNTIFS($CMP,$x,$TB[[Compra mês vigente QTD]],"">0"")", "=COUNTIFS($CMP,$x,$TB[[$($CN.q0)]],"">0"")",
               "=SUMIFS($TB[[Compra mês vigente R`$]],$CMP,$x)", "=SUMIFS($TB[[$($CN.r0)]],$CMP,$x)",
               "=SUMIFS($TB[[Compra próximo mês R`$]],$CMP,$x)", "=SUMIFS($TB[[$($CN.r1)]],$CMP,$x)",
               "=SUMIFS($TB[[Compra mês subsequente R`$]],$CMP,$x)", "=SUMIFS($TB[[$($CN.r2)]],$CMP,$x)",
               "=SUMIFS($TB[[$($nm[12])]],$CMP,$x)",
               "=SUMPRODUCT(($CMP=$x)*$TB[[$($nm[5])]]*$PRV)", "=SUMPRODUCT(($CMP=$x)*$TB[[$($CN.a0)]]*$PRV)")
        for ($j = 0; $j -lt $f.Count; $j++) {
            try { $wa.Cells.Item($rr, 3 + $j).Formula2 = $f[$j] } catch { throw "formula recusada: $($f[$j])" }
        }
    }
    $rt = $r0 + 1 + $compradores.Count
    $wa.Cells.Item($rt, 2).Value2 = 'Total'
    for ($j = 3; $j -le 13; $j++) { $c = $wa.Cells.Item($rt, $j); $c.Formula = '=SUM(' + $wa.Cells.Item($r0 + 1, $j).Address(0, 0) + ':' + $wa.Cells.Item($rt - 1, $j).Address(0, 0) + ')' }
    $tb = $wa.Range($wa.Cells.Item($r0, 2), $wa.Cells.Item($rt, 13)); $tb.Borders.LineStyle = 1
    $h = $wa.Range($wa.Cells.Item($r0, 2), $wa.Cells.Item($r0, 13)); $h.Font.Bold = $true; $h.Interior.Color = 14599344; $h.Font.Color = 16777215; $h.WrapText = $true
    $wa.Range($wa.Cells.Item($r0 + 1, 3), $wa.Cells.Item($rt, 4)).NumberFormatLocal = $FI
    $wa.Range($wa.Cells.Item($r0 + 1, 5), $wa.Cells.Item($rt, 13)).NumberFormatLocal = $FR
    $wa.Range($wa.Cells.Item($rt, 2), $wa.Cells.Item($rt, 13)).Font.Bold = $true
    $script:r = $rt + 2
    Txt @(
        '• No mês vigente a compra CLAUDE é maior porque recompõe a curva A, que está com 1,1 mês de estoque contra alvo de cerca de 2,2. Sugestão: dividir entre o fim de outubro e novembro, começando pelos itens em ruptura e de pedido grande.',
        '• O estoque alvo CLAUDE é maior que o da regra atual porque a curva A sobe de 1,6 para 2,2 meses. A curva C cai de 9,8 meses (hoje) para 2,7 de alvo: o excesso vai sendo vendido sem compra nova.',
        '• Para gastar menos capital com a mesma proteção, a curva A pode ter revisão semanal (cenário "menos capital": alvo de R$ 33 mi, perto do estoque de hoje). Os parâmetros ficam na aba param CLAUDE.'
    )
    Tit '4.1 Cenários de parâmetros (26/10)'
    Tab $JS.cenarios @($null, $FR, $FR, $FR, $FR)
    $script:r++
    Tit '5. QUALIDADE DOS DADOS'
    Tab $JS.dados @($null, $null)
    $script:r++
    Tit '6. SUGESTÕES DE MELHORIA (por ordem de impacto)'
    Txt @(
        '1. Separar no indicador de ruptura os itens com PontoPedido 0 (sob encomenda): 30% dos episódios e R$ 2,4 mi faltantes no início. Acompanhar à parte, pelo prazo de compra do pedido de venda.',
        '2. Rever os 266 PontoPedido zerados à mão que têm venda (o NOVO_PONTO do TOTVS sugere estoque): quem vende todo mês deveria ter ponto.',
        '3. Curva A com revisão quinzenal (ou semanal) e prioridade no PLANO de AÇAO: 132 SKUs, 80% da venda, 8,8% dos dias em ruptura e 1,1 mês de estoque.',
        '4. Não comprar curva C com cobertura acima de 4 meses e tratar o excesso: R$ 11 mi com mais de 4 meses ou sem venda (marketplace, transferência, devolução ao fornecedor, ação comercial).',
        '5. Combinar com o comercial um aviso de pedido grande (carteira e contratos: CEF, RD, FDE, PGMN...): 19% das rupturas começam com um pedido de mais de 1 mês de venda.',
        '6. Follow-up por fornecedor com o prazo real medido (tabela 2.6): os 3 maiores em R$ chegam no prazo em 0% a 20% dos pedidos e 80% a 100% são remarcados. Renegociar prazo, pedir antes ou dividir entre fornecedores.',
        '7. Trazer no relatório do TOTVS a data de emissão do pedido de compra e a quantidade já recebida: o prazo real passa a ser medido direto, sem precisar das fotos diárias.',
        '8. Pedir ao TI um relatório de venda mensal por SKU de 24 a 36 meses (todos os canais, na unidade do estoque): as bases de faturamento não trazem todos os canais, e sem isso a sazonalidade é estimada pelos SKUs que elas cobrem.',
        '9. Corrigir no relatório do TOTVS o Disponível sem vírgula (5 SKUs fracionados) e incluir no relatório do Protheus os 49 SKUs da Geral que faltam.',
        '10. No PLANO de AÇAO, mostrar também a compra CLAUDE e a cobertura alvo ao lado da regra atual, para o comprador decidir com os dois números.'
    )

    # devolve o filtro e salva
    $xl.Calculation = -4105
    $xl.CalculateFull()
    foreach ($fl in $filtros) {
        if ($fl[2] -eq 7) { $lo.Range.AutoFilter($fl[0], $fl[1], 7) | Out-Null } else { $lo.Range.AutoFilter($fl[0], $fl[1]) | Out-Null }
    }
    if ($filtros.Count) { "filtro devolvido: $($lo.DataBodyRange.Columns.Item(1).SpecialCells(12).Count) linhas visiveis (antes $visiveis)" }
    $sem = 0
    foreach ($n in $defs.Keys) { $rgc = $lo.ListColumns.Item($n).DataBodyRange; $sem += $rgc.Count - @($rgc.SpecialCells(-4123)).Count }
    if ($sem) { throw "$sem celulas CLAUDE sem formula" }
    $wa.Activate()
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
