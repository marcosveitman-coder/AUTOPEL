# Revisao do "Relatorio Revenda atualizado.xlsx" (aba Geral, tabela Table1) pelo proprio Excel.
# Use numa COPIA do arquivo, fechada no Excel.  Grava:
#   - Table1: corrige a programacao do trimestre (colunas GPT de mes vigente / proximo / subsequente),
#     a checagem "divergencia COMPRAR", os % contra Oceano/Kalunga (#VALOR!) e a FAMILIA sem cadastro (#N/D);
#     acrescenta no fim: Saldo hoje, Compra trimestre QTD e R$, Situacao compra.
#   - dinamicas: o tracking (AF:AH) passa a somar o trimestre de TODOS os SKUs, direto da Table1.
#   - abas novas: LEIA-ME COMPRAS (calculadora do Comprar + regras) e PROGRAMACAO TRIMESTRE (resumo + dinamica).
#   - atualiza a cache das dinamicas que usam a Table1.
#   powershell -ExecutionPolicy Bypass -File revisa_relatorio_revenda.ps1 -Xlsx C:\...\copia.xlsx
param([Parameter(Mandatory = $true)][string]$Xlsx,
      [string]$Exemplo1 = 'V30481',     # SKU com Comprar negativo (sobra)
      [string]$Exemplo2 = 'V31087')     # SKU com Comprar positivo (falta)
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
try {
    $wb = $xl.Workbooks.Open($Xlsx, 0, $false)
    # formatos de numero com os separadores do Excel da maquina: chamado pelo PowerShell, o Excel le o formato no
    #   idioma local (em pt-BR '#,##0' vira numero com decimal: 1416,0)
    $mil = $xl.International(4); $dec = $xl.International(3)
    $FI = "#${mil}##0"; $FD = "#${mil}##0${dec}00"
    $xl.Calculation = -4135                                  # manual durante a gravacao
    $ws = $null                                               # aba Geral = a que tem a Table1
    foreach ($s in $wb.Worksheets) { foreach ($t in $s.ListObjects) { if ($t.Name -eq 'Table1') { $ws = $s } } }
    if (-not $ws) { throw 'Table1 nao encontrada' }
    $lo = $ws.ListObjects.Item('Table1')

    # nomes das colunas pela posicao (os cabecalhos do TOTVS tem espacos nas pontas)
    $nm = @{}; foreach ($lc in $lo.ListColumns) { $nm[$lc.Index] = $lc.Name }
    $esperado = @{ 5 = ' PontoPedido'; 6 = 'Comprar.'; 9 = ' MedContr'; 15 = ' Disponivel'; 16 = ' VendProt'; 17 = ' Ruptura'
                   18 = ' Comprado'; 25 = 'Comprador  2'; 37 = 'divergencia COMPRAR'; 46 = 'GPT previsao OUTUBRO QNTDD'
                   47 = 'GPT  previsao OUTUBRO R$'; 48 = 'GPT previsao NOVEMBRO QNTDD'; 49 = 'GPT  previsao NOVEMBRO R$'
                   50 = 'GPT  previsao DEZEMVRO QNTDD'; 51 = 'GPT previsao DEZEMVRO R$'; 63 = 'FAMILIA'
                   65 = 'Oceano_Preco'; 67 = 'Kalunga_Preco'; 75 = 'AUTOPEL vs. Oceano B2B'; 76 = 'AUTOPEL vs. Kalunga' }
    foreach ($k in $esperado.Keys) { if ($nm[$k] -ne $esperado[$k]) { throw "coluna $k da Table1 mudou: '$($nm[$k])'" } }
    if ($lo.ListColumns.Count -ne 84) { throw "Table1 tem $($lo.ListColumns.Count) colunas (esperado 84): ja foi revisada?" }

    function L([int]$i) { '[@[' + $nm[$i] + ']]' }          # celula da linha, dentro da tabela
    function T([string]$n) { 'Table1[[' + $n + ']]' }        # coluna inteira, fora da tabela
    function SetCol([int]$i, [string]$f) { $lo.ListColumns.Item($i).DataBodyRange.Formula = $f }

    $PP = L 5; $COMPRAR = L 6; $MED = L 9; $DISP = L 15; $VPROT = L 16; $COMPRADO = L 18

    # --- colunas novas no fim da Table1 ---------------------------------------------------------------
    $SALDO_N = 'Saldo hoje (Disp-VendProt+Comprado)'
    $c = $lo.ListColumns.Add(); $c.Name = $SALDO_N; $nm[$c.Index] = $SALDO_N
    # = Disponivel - VendProt + Comprado.  Lido como PontoPedido - Comprar: e o mesmo numero do TOTVS e nao herda
    #   a virgula perdida da coluna O (SKUs de quantidade fracionada).
    $c.DataBodyRange.Formula = "=$PP-$COMPRAR"
    $SALDO = L $c.Index

    # --- programacao do trimestre (colunas GPT, as que alimentam as dinamicas) -------------------------
    # mes vigente: volta o saldo ao ponto de pedido; com ponto 0 so cobre o que ja foi vendido sem estoque
    SetCol 46 "=IF($PP=0,MAX(0,-$SALDO),MAX(0,$PP-$SALDO))"
    # proximo mes: saldo depois da compra do mes vigente e de 1 mes de venda
    SetCol 48 "=IF($PP=0,0,MAX(0,$PP-($SALDO+$(L 46)-$MED)))"
    # mes subsequente: saldo depois das duas compras e de 2 meses de venda
    SetCol 50 "=IF($PP=0,0,MAX(0,$PP-($SALDO+$(L 46)+$(L 48)-2*$MED)))"

    $c = $lo.ListColumns.Add(); $c.Name = 'Compra trimestre QTD'; $nm[$c.Index] = $c.Name
    $c.DataBodyRange.Formula = "=$(L 46)+$(L 48)+$(L 50)"
    $c.DataBodyRange.NumberFormat = $lo.ListColumns.Item(46).DataBodyRange.Cells.Item(1, 1).NumberFormat
    $iQtd = $c.Index
    $c = $lo.ListColumns.Add(); $c.Name = 'Compra trimestre R$'; $nm[$c.Index] = $c.Name
    $c.DataBodyRange.Formula = "=$(L 47)+$(L 49)+$(L 51)"
    $c.DataBodyRange.NumberFormat = $lo.ListColumns.Item(47).DataBodyRange.Cells.Item(1, 1).NumberFormat
    $c = $lo.ListColumns.Add(); $c.Name = 'Situacao compra'; $nm[$c.Index] = $c.Name
    $c.DataBodyRange.Formula = "=IF($(L 17)<0,""1 RUPTURA"",IF($(L 46)>0,""2 COMPRAR AGORA"",IF($(L $iQtd)>0,""3 COMPRAR NOS PROXIMOS MESES"",IF($PP=0,""5 SEM PONTO DE PEDIDO"",""4 ESTOQUE OK""))))"
    $lo.ListColumns.Item($SALDO_N).DataBodyRange.NumberFormat = $lo.ListColumns.Item(46).DataBodyRange.Cells.Item(1, 1).NumberFormat

    # totais da linha 1, no mesmo padrao das colunas de previsao
    foreach ($p in @(@('Compra trimestre QTD', 'AX1'), @('Compra trimestre R$', 'AY1'))) {
        $col = $lo.ListColumns.Item($p[0]).Range.Column
        $ws.Range($p[1]).Copy($ws.Cells.Item(1, $col)) | Out-Null
    }

    # --- checagens que davam resultado errado -----------------------------------------------------------
    # "divergencia COMPRAR": marcava estoque acima de 3 meses; agora marca quando o Comprar do TOTVS nao bate
    #   com PontoPedido - (Disponivel - VendProt + Comprado)
    SetCol 37 "=IF(ABS($COMPRAR-($PP-($DISP-$VPROT+$COMPRADO)))>0.5,""divergente"","""")"
    # % contra concorrente: "" > 0 e VERDADEIRO no Excel e a divisao dava #VALOR!
    SetCol 75 "=IF(N($(L 65))>0,($(L 7)/$(L 65))-1,"""")"
    SetCol 76 "=IF(N($(L 67))>0,($(L 7)/$(L 67))-1,"""")"
    # FAMILIA de SKU sem cadastro na aba de familias: #N/D
    SetCol 63 "=XLOOKUP($(L 2),'familias skus PADRONIZACAO)'!AE:AE,'familias skus PADRONIZACAO)'!AC:AC,""SEM FAMILIA"")"
    SetCol 64 "=XLOOKUP($(L 2),'familias skus PADRONIZACAO)'!AE:AE,'familias skus PADRONIZACAO)'!`$AB:`$AB,""SEM FAMILIA"")"

    # --- tracking da aba dinamicas ----------------------------------------------------------------------
    # AF:AH somavam a dinamica filtrada em "AUX COMPRAR = SIM" (so SKUs com Comprar > 0 hoje): novembro e dezembro
    # perdiam os SKUs que so precisam de compra nos proximos meses.  Agora somam a Table1 inteira por comprador.
    $wd = $wb.Worksheets.Item('dinamicas')
    if ([string]$wd.Range('AC4').Value2 -notlike 'Comprador*' -or [string]$wd.Range('AC13').Value2 -ne 'TOTAL') { throw 'tracking da dinamicas mudou de lugar' }
    $wd.Range('AF5:AF12').Formula = "=SUMIFS($(T $nm[47]),$(T $nm[25]),`$AC5)"
    $wd.Range('AG5:AG12').Formula = "=SUMIFS($(T $nm[49]),$(T $nm[25]),`$AC5)"
    $wd.Range('AH5:AH12').Formula = "=SUMIFS($(T $nm[51]),$(T $nm[25]),`$AC5)"
    $wd.Range('AJ5:AJ12').Formula = '=IFERROR(XLOOKUP(AC5,$BD$4:$BD$1048576,$BE$4:$BE$1048576),0)'

    # --- aba LEIA-ME COMPRAS ------------------------------------------------------------------------------
    $wl = $wb.Worksheets.Add([Type]::Missing, $ws); $wl.Name = 'LEIA-ME COMPRAS'
    $wl.Cells.Font.Name = 'Calibri'; $wl.Cells.Font.Size = 11
    $wl.Columns.Item(1).ColumnWidth = 2; $wl.Columns.Item(2).ColumnWidth = 46
    $wl.Columns.Item(3).ColumnWidth = 30; $wl.Columns.Item(4).ColumnWidth = 30; $wl.Columns.Item(5).ColumnWidth = 80
    function W([string]$cel, $v, [switch]$Neg, [int]$Tam = 0) {
        $r = $wl.Range($cel)
        if ([string]$v -like '=*') { $r.Formula = $v } else { $r.Value2 = [string]$v }
        if ($Neg) { $r.Font.Bold = $true }
        if ($Tam) { $r.Font.Size = $Tam }
    }
    # paragrafo: texto na B, juntando B:E com quebra de linha (Linhas = altura, para textos de formula)
    function P([int]$lin, $v, [int]$Linhas = 0) {
        W "B$lin" $v
        $m = $wl.Range("B${lin}:E$lin"); $m.Merge(); $m.WrapText = $true; $m.VerticalAlignment = -4160
        if (-not $Linhas) { $Linhas = [math]::Max(1, [math]::Ceiling(([string]$v).Length / 165)) }
        $m.RowHeight = 15 * $Linhas
    }
    W 'B1' 'COMO A PLANILHA CALCULA A COMPRA' -Neg -Tam 16
    W 'B2' ('Aba ' + $ws.Name + ': as colunas A a X vêm do TOTVS e não são alteradas. Tudo aqui é fórmula: atualiza sozinho quando a Geral muda.')

    W 'B4' '1. CALCULADORA: troque o código nas células amarelas' -Neg -Tam 13
    $lin = @(
        @('Código do SKU', $null, 'Exemplo 1: Comprar negativo (sobra).  Exemplo 2: Comprar positivo (falta).'),
        @('Descrição', 3, ''),
        @('Comprador', 1, ''),
        @('PontoPedido (E)', 5, 'Estoque-alvo: 2 meses de venda. Se for 0, o item não precisa ter estoque.'),
        @('MedContr (I)', 9, 'Quanto se vende por mês.'),
        @('Físico (M)', 13, 'Estoque na prateleira.'),
        @('Reserva (N)', 14, 'Parte do físico já separada para pedidos.'),
        @('Disponível (O)', 15, 'Físico menos Reserva.'),
        @('VendProt (P)', 16, 'Vendido no sistema e ainda não entregue: vai sair do estoque.'),
        @('Comprado (R)', 18, 'Pedido de compra em aberto: vai entrar no estoque.'),
        @('SALDO HOJE = Disponível − VendProt + Comprado', 'saldo', 'O que vou ter depois de entregar o que já vendi e receber o que já comprei.'),
        @('Conta do Comprar = PontoPedido − Saldo', 'conta', 'O que preciso ter MENOS o que vou ter.'),
        @('Comprar (F) que veio do TOTVS', 6, 'Tem que bater com a linha de cima.'),
        @('Leitura', 'leitura', ''),
        @('Compra no mês vigente (QTD)', 46, 'Coluna AT da Geral ("GPT previsao OUTUBRO QNTDD").'),
        @('Compra no próximo mês (QTD)', 48, 'Coluna AV da Geral ("GPT previsao NOVEMBRO QNTDD").'),
        @('Compra no mês subsequente (QTD)', 50, 'Coluna AX da Geral ("GPT  previsao DEZEMVRO QNTDD").'),
        @('Compra no trimestre (R$)', 'Compra trimestre R$', 'Quantidade × Ult.Preco, somando os 3 meses.'),
        @('Situação', 'Situacao compra', '')
    )
    $r0 = 6
    for ($i = 0; $i -lt $lin.Count; $i++) {
        $r = $r0 + $i; $l = $lin[$i]
        W "B$r" $l[0]; W "E$r" $l[2]
        foreach ($col in @('C', 'D')) {
            $cod = "`$$col`$$r0"
            $f = $null
            if ($l[1] -is [int]) { $f = "=XLOOKUP($cod,$(T $nm[2]),$(T $nm[$l[1]]),""código não achado"")" }
            elseif ($l[1] -eq 'saldo') { $f = "=${col}13-${col}14+${col}15" }
            elseif ($l[1] -eq 'conta') { $f = "=${col}9-${col}16" }
            elseif ($l[1] -eq 'leitura') { $f = "=IF(${col}18>0,""FALTA: comprar ""&FIXED(${col}18,0),IF(${col}18<0,""SOBRA ""&FIXED(-${col}18,0)&"": não comprar"",""no ponto de pedido""))" }
            elseif ($l[1]) { $f = "=XLOOKUP($cod,$(T $nm[2]),$(T $l[1]),""código não achado"")" }
            if ($f) { $wl.Range("$col$r").Formula = $f }
        }
    }
    $wl.Range('C6').Value2 = $Exemplo1; $wl.Range('D6').Value2 = $Exemplo2
    $wl.Range('C6:D6').Interior.Color = 10284031; $wl.Range('C6:D6').Font.Bold = $true
    $wl.Range('B16:D18').Font.Bold = $true; $wl.Range('B16:D16').Interior.Color = 15921906; $wl.Range('B18:D18').Interior.Color = 15921906
    $wl.Range("C9:D$($r0 + $lin.Count - 1)").NumberFormatLocal = "$FI;-$FI"
    $wl.Range('C23:D23').NumberFormatLocal = $FD
    $wl.Range("C6:D$($r0 + $lin.Count - 1)").HorizontalAlignment = -4152
    foreach ($t in @(7, 8, 19, 24)) { $wl.Range("C${t}:D$t").HorizontalAlignment = -4131 }
    $wl.Range('C6:D6').HorizontalAlignment = -4108
    $wl.Range('C7:D7').WrapText = $true; $wl.Range('C7:D7').VerticalAlignment = -4160
    $tab = $wl.Range("B6:E$($r0 + $lin.Count - 1)"); $tab.Borders.LineStyle = 1; $tab.Borders.Weight = 2
    $wl.Range("E6:E$($r0 + $lin.Count - 1)").Font.Color = 8421504

    $r = 27
    W "B$r" '2. POR QUE O COMPRAR DÁ NEGATIVO (exemplo 1, coluna C)' -Neg -Tam 13; $r++
    P $r '="Disponível = Físico − Reserva = "&FIXED(C11,0)&" − "&FIXED(C12,0)&" = "&FIXED(C13,0)' 1; $r++
    P $r '="Saldo = Disponível − VendProt + Comprado = "&FIXED(C13,0)&" − "&FIXED(C14,0)&" + "&FIXED(C15,0)&" = "&FIXED(C16,0)&"   ← é o número positivo que você chega"' 1; $r++
    P $r '="Comprar = PontoPedido − Saldo = "&FIXED(C9,0)&" − "&FIXED(C16,0)&" = "&FIXED(C17,0)' 1; $r++
    P $r 'O TOTVS sempre faz "o que preciso ter" MENOS "o que vou ter". Positivo = falta (comprar). Negativo = sobra acima do ponto de pedido (não comprar).'; $r++
    P $r '="No exemplo 2 (coluna D) é a mesma conta: "&FIXED(D9,0)&" − "&FIXED(D16,0)&" = "&FIXED(D17,0)&". Por isso ali o resultado é positivo."' 1; $r++
    P $r '=IF(AND(C9=0,C10>0),"Atenção: este SKU vende "&FIXED(C10,0)&" por mês e está com PontoPedido 0 no TOTVS (a coluna AL da Geral marca ""equivoco""). Com 2 meses de estoque o ponto seria "&FIXED(2*C10,0)&" e o Comprar daria "&FIXED(2*C10-C16,0)&".","")' 2; $r++
    $r++

    W "B$r" '3. PROGRAMAÇÃO DO TRIMESTRE (colunas AT a AY da Geral)' -Neg -Tam 13; $r++
    $txt = @(
        'Regra: no começo de cada mês o saldo volta ao PontoPedido (2 meses de estoque). Entre uma compra e a outra sai 1 MedContr (1 mês de venda).',
        'Saldo hoje (coluna nova no fim da tabela) = Disponível − VendProt + Comprado.',
        'Mês vigente  =SE(PontoPedido=0; MÁXIMO(0; −Saldo); MÁXIMO(0; PontoPedido − Saldo))',
        '      É o próprio Comprar do TOTVS quando positivo. Com PontoPedido 0 só compra o que já foi vendido sem estoque (VendProt maior que Disponível + Comprado).',
        'Próximo mês  =SE(PontoPedido=0; 0; MÁXIMO(0; PontoPedido − (Saldo + Compra mês vigente − MedContr)))',
        'Mês subsequente  =SE(PontoPedido=0; 0; MÁXIMO(0; PontoPedido − (Saldo + Compra mês vigente + Compra próximo mês − 2 × MedContr)))',
        'R$ de cada mês = quantidade × Ult.Preco (colunas AU, AW, AY).  Trimestre = soma dos 3 meses (colunas novas "Compra trimestre QTD" e "R$").',
        'Quando o SKU já está abaixo do ponto de pedido, a compra do próximo mês e do subsequente é igual à MedContr: repõe o que se vende em 1 mês. Não é erro.',
        'Quando o SKU tem sobra, a compra começa no mês em que o saldo, descontada a venda, fica abaixo do ponto de pedido.',
        'Os nomes OUTUBRO, NOVEMBRO e DEZEMBRO nos cabeçalhos são só rótulos: a conta é sempre mês vigente, próximo e subsequente. Renomear o cabeçalho tira o campo das dinâmicas.'
    )
    foreach ($t in $txt) { P $r $t; $r++ }
    $r++

    W "B$r" '4. O QUE FOI CORRIGIDO NESTA REVISÃO' -Neg -Tam 13; $r++
    $txt = @(
        '• Tracking da aba dinamicas (AF:AH): somava só os SKUs com Comprar > 0 hoje (dinâmica filtrada em "AUX COMPRAR = SIM"). Os SKUs que só precisam de compra em novembro ou dezembro ficavam de fora, e por isso novembro e dezembro davam iguais. Agora soma todos os SKUs, direto da tabela.',
        '• Mês vigente (coluna AT): com PontoPedido 0 dava sempre 0, mesmo com venda feita sem estoque. Agora compra o que já foi vendido (o TOTVS também manda comprar esses).',
        '• Saldo: passa a vir do próprio TOTVS (PontoPedido − Comprar), igual a Disponível − VendProt + Comprado. Em SKUs de quantidade fracionada a coluna O veio sem a vírgula (ex.: 163,66 virou 16366) e zerava a compra.',
        '• "divergencia COMPRAR" (coluna AK): marcava estoque acima de 3 meses, não divergência. Agora marca "divergente" quando o Comprar do TOTVS não bate com a conta.',
        '• "AUTOPEL vs. Oceano B2B" e "vs. Kalunga" (BW e BX): davam #VALOR! quando não havia preço do concorrente. Agora ficam em branco.',
        '• FAMILIA (BK e BL): SKU sem cadastro na aba de famílias dava #N/D. Agora mostra "SEM FAMILIA".',
        '• RUPTURAS no tracking (dinamicas AJ): comprador sem ruptura dava #N/D (estava digitado 0). Agora a fórmula devolve 0.',
        '• Colunas novas no fim da tabela: Saldo hoje, Compra trimestre QTD e R$, Situação compra (1 ruptura, 2 comprar agora, 3 comprar nos próximos meses, 4 estoque ok, 5 sem ponto de pedido).',
        '• Aba nova PROGRAMACAO TRIMESTRE: resumo por comprador (todos os SKUs) e tabela dinâmica por comprador e marca, com filtro de situação.'
    )
    foreach ($t in $txt) { P $r $t; $r++ }
    $r++

    W "B$r" '5. PONTOS DE ATENÇÃO NOS DADOS DO TOTVS (não alterados)' -Neg -Tam 13; $r++
    $chk = @(
        @('SKUs na tabela', "=ROWS(Table1)"),
        @('PontoPedido = 0, mas com venda (MedContr > 0)', "=COUNTIFS($(T $nm[5]),0,$(T $nm[9]),"">0"")"),
        @('PontoPedido diferente de 2 × MedContr (coluna AJ = FALSO)', "=COUNTIF($(T $nm[36]),FALSE)"),
        @('MedContr maior ou igual ao PontoPedido (cobre menos de 1 mês)', "=COUNTIF($(T $nm[39]),""equivoco"")"),
        @('Comprar do TOTVS não bate com a conta (coluna O sem vírgula)', "=COUNTIF($(T $nm[37]),""divergente"")"),
        @('VendProt negativo', "=COUNTIF($(T $nm[16]),""<0"")"),
        @('Comprado negativo', "=COUNTIF($(T $nm[18]),""<0"")"),
        @('Estoque acima de 4 meses de venda (coluna CF)', "=COUNTIF($(T $nm[84]),""sim"")"),
        @('SKUs em ruptura (Disponível menor que VendProt)', "=COUNTIF($(T $nm[17]),""<0"")")
    )
    $r1 = $r
    foreach ($k in $chk) {
        W "B$r" $k[0]; $wl.Range("B${r}:C$r").Merge()
        $wl.Range("D$r").Formula = $k[1]; $wl.Range("D$r").NumberFormatLocal = $FI; $r++
    }
    $tab = $wl.Range("B${r1}:D$($r - 1)"); $tab.Borders.LineStyle = 1; $tab.Borders.Weight = 2
    $txt = @(
        'As colunas GROK (AN:AS) e deepSEAK (AZ:BE) fazem a mesma conta da GPT antiga e não foram corrigidas: as dinâmicas usam só a GPT. Podem ser apagadas.',
        'Na aba PLANO de AÇAO, a área V:AI (previsão de agosto a outubro) é de uma rodada antiga e está fora das dinâmicas atuais.'
    )
    foreach ($t in $txt) { P $r $t; $r++ }

    # --- aba PROGRAMACAO TRIMESTRE ------------------------------------------------------------------------
    $wp = $wb.Worksheets.Add([Type]::Missing, $wl); $wp.Name = 'PROGRAMACAO TRIMESTRE'
    $wp.Cells.Font.Name = 'Calibri'; $wp.Cells.Font.Size = 11
    $wp.Range('A1').Value2 = 'PROGRAMAÇÃO DE COMPRA DO TRIMESTRE, POR COMPRADOR'; $wp.Range('A1').Font.Bold = $true; $wp.Range('A1').Font.Size = 16
    $wp.Range('A2').Value2 = 'Todos os SKUs da aba Geral, sem filtro. Mês vigente = mês de hoje. Regra: 2 meses de estoque (PontoPedido).'
    $cab = @('Comprador', 'SKUs em ruptura', 'SKUs comprar agora', 'SKUs só nos próximos meses')
    for ($j = 0; $j -lt $cab.Count; $j++) { $wp.Cells.Item(4, 1 + $j).Value2 = $cab[$j] }
    $wp.Range('E4').Formula = '=PROPER(TEXT(TODAY(),"mmmm"))&" R$"'
    $wp.Range('F4').Formula = '=PROPER(TEXT(EDATE(TODAY(),1),"mmmm"))&" R$"'
    $wp.Range('G4').Formula = '=PROPER(TEXT(EDATE(TODAY(),2),"mmmm"))&" R$"'
    $wp.Range('H4').Value2 = 'Trimestre R$'
    $CMP = T $nm[25]
    $wp.Range('A5').Formula2 = "=SORT(UNIQUE(FILTER($CMP,$CMP<>"""")))"
    $wp.Range('B5').Formula2 = "=COUNTIFS($CMP,A5#,$(T $nm[17]),""<0"")"
    $wp.Range('C5').Formula2 = "=COUNTIFS($CMP,A5#,$(T $nm[46]),"">0"")"
    $wp.Range('D5').Formula2 = "=COUNTIFS($CMP,A5#,$(T 'Situacao compra'),""3*"")"
    $wp.Range('E5').Formula2 = "=SUMIFS($(T $nm[47]),$CMP,A5#)"
    $wp.Range('F5').Formula2 = "=SUMIFS($(T $nm[49]),$CMP,A5#)"
    $wp.Range('G5').Formula2 = "=SUMIFS($(T $nm[51]),$CMP,A5#)"
    $wp.Range('H5').Formula2 = '=E5#+F5#+G5#'
    $wp.Range('A3').Value2 = 'TOTAL'
    foreach ($col in @('B', 'C', 'D', 'E', 'F', 'G', 'H')) { $wp.Range("${col}3").Formula2 = "=SUM(${col}5#)" }
    $wp.Range('A3:H3').Font.Bold = $true; $wp.Range('A3:H3').Interior.Color = 15921906
    $h = $wp.Range('A4:H4'); $h.Font.Bold = $true; $h.Interior.Color = 14599344; $h.Font.Color = 16777215; $h.WrapText = $true
    $wp.Range('B3:H40').NumberFormatLocal = $FI
    $wp.Columns.Item(1).ColumnWidth = 16
    foreach ($j in 2..8) { $wp.Columns.Item($j).ColumnWidth = 15 }
    # notas em linhas curtas: as linhas sao as mesmas da dinamica ao lado, entao sem quebra de texto nem altura maior
    $notas = @('Como ler: SKUs comprar agora = Comprar do TOTVS positivo.',
               'SKUs só nos próximos meses = hoje tem saldo, mas ele fica abaixo do ponto de pedido em um dos 2 meses seguintes.',
               'O número de SKUs e de marcas do tracking (aba dinamicas) continua contando só quem tem Comprar > 0 hoje,',
               'para comparar com os dias anteriores.')
    $r = 18
    foreach ($t in $notas) { $wp.Range("A$r").Value2 = $t; $wp.Range("A$r").Font.Color = 8421504; $r++ }

    # --- atualiza a cache da Table1 e cria a dinamica de programacao ------------------------------------
    $xl.Calculation = -4105                                  # automatico
    $xl.CalculateFull()
    $pc = $null
    foreach ($p in $wb.PivotCaches()) { if ([string]$p.SourceData -eq 'Table1') { $pc = $p } }
    if (-not $pc) { throw 'cache da Table1 nao encontrada' }
    $pc.Refresh()
    $wp.Range('J3').Value2 = 'Tabela dinâmica: comprador e marca (filtre a situação no campo de cima)'; $wp.Range('J3').Font.Bold = $true
    $pt = $pc.CreatePivotTable($wp.Range('J6'), 'TD Programacao trimestre')
    $pt.RowAxisLayout(1)
    $pt.RepeatAllLabels(2)
    $pt.TableStyle2 = 'PivotStyleMedium2'
    $pt.PivotFields('Situacao compra').Orientation = 3
    $f = $pt.PivotFields($nm[25]); $f.Orientation = 1; $f.Position = 1
    $f = $pt.PivotFields($nm[26]); $f.Orientation = 1; $f.Position = 2
    $dados = @(@(47, 'Mês vigente R$'), @(49, 'Próximo mês R$'), @(51, 'Mês subsequente R$'), @('Compra trimestre R$', 'Trimestre R$'))
    foreach ($d in $dados) {
        $campo = $(if ($d[0] -is [int]) { $nm[$d[0]] } else { $d[0] })
        $df = $pt.AddDataField($pt.PivotFields($campo), $d[1], -4157); $df.NumberFormat = $FI     # campo de dinamica nao tem NumberFormatLocal; aqui o Excel ja le no idioma local
    }
    $wp.Columns.Item(10).ColumnWidth = 16; $wp.Columns.Item(11).ColumnWidth = 24
    foreach ($j in 12..15) { $wp.Columns.Item($j).ColumnWidth = 16 }

    $xl.CalculateFull()
    $wb.Worksheets.Item('LEIA-ME COMPRAS').Activate()
    $wb.Save()
    'ok'
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
