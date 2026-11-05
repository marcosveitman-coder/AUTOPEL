# Mede a largura (px, 96 dpi) de cada letra da Verdana 10 normal e negrito e grava verdana10.txt (estilo, codigo, largura),
# usado pelo org_impressao.py para saber se um texto cabe na celula. Rodar de novo so se mudar a letra.
Add-Type -AssemblyName System.Drawing
$bmp = New-Object System.Drawing.Bitmap 10, 10; $bmp.SetResolution(96, 96); $g = [System.Drawing.Graphics]::FromImage($bmp)
$fmt = [System.Drawing.StringFormat]::GenericTypographic; $fmt.FormatFlags = $fmt.FormatFlags -bor [System.Drawing.StringFormatFlags]::MeasureTrailingSpaces
$chars = ' !"#$%&''()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~·–—…ÀÁÂÃÇÉÊÍÓÔÕÚÜàáâãçéêíóôõúü★≠≤'
$lines = @()
foreach ($st in @('Regular', 'Bold')) {
  $f = New-Object System.Drawing.Font('Verdana', 10, [System.Drawing.FontStyle]::$st, [System.Drawing.GraphicsUnit]::Point)
  foreach ($c in $chars.ToCharArray()) { $s = [string]$c * 4; $lines += "$st`t$([int]$c)`t$([math]::Round($g.MeasureString($s, $f, 2000, $fmt).Width / 4, 3))" }
}
$lines | Out-File -Encoding utf8 (Join-Path $PSScriptRoot 'verdana10.txt')
