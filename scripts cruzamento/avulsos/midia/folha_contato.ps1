# Monta folhas de contato (grade de fotos numeradas) com System.Drawing.
# Uso: folha_contato.ps1 -Pasta <fotos> -Saida <pasta> [-Colunas 5] [-Linhas 2] [-Largura 380] [-Altura 676]
param(
    [Parameter(Mandatory=$true)][string]$Pasta,
    [Parameter(Mandatory=$true)][string]$Saida,
    [int]$Colunas = 5,
    [int]$Linhas = 2,
    [int]$Largura = 380,
    [int]$Altura = 676,
    [string]$Filtro = "*.jpg"
)
Add-Type -AssemblyName System.Drawing
New-Item -ItemType Directory -Force $Saida | Out-Null

# so fotos unicas (mesmo conteudo = mesma foto)
$md5 = [System.Security.Cryptography.MD5]::Create()
$vistos = @{}
$fotos = @()
foreach ($arquivo in (Get-ChildItem $Pasta -Filter $Filtro | Sort-Object Name)) {
    $hash = [BitConverter]::ToString($md5.ComputeHash([IO.File]::ReadAllBytes($arquivo.FullName)))
    if (-not $vistos.ContainsKey($hash)) { $vistos[$hash] = 1; $fotos += $arquivo }
}

$porFolha = $Colunas * $Linhas
$fonte = New-Object System.Drawing.Font('Arial', 22, [System.Drawing.FontStyle]::Bold)
$indice = @()
for ($inicio = 0; $inicio -lt $fotos.Count; $inicio += $porFolha) {
    $folha = New-Object System.Drawing.Bitmap ($Colunas * $Largura), ($Linhas * $Altura)
    $grafico = [System.Drawing.Graphics]::FromImage($folha)
    $grafico.Clear([System.Drawing.Color]::Black)
    $grafico.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    for ($k = 0; $k -lt $porFolha -and ($inicio + $k) -lt $fotos.Count; $k++) {
        $numero = $inicio + $k + 1
        $foto = [System.Drawing.Image]::FromFile($fotos[$inicio + $k].FullName)
        $escala = [Math]::Min(($Largura - 8) / $foto.Width, ($Altura - 8) / $foto.Height)
        $larguraFoto = [int]($foto.Width * $escala); $alturaFoto = [int]($foto.Height * $escala)
        $x = ($k % $Colunas) * $Largura + [int](($Largura - $larguraFoto) / 2)
        $y = [Math]::Floor($k / $Colunas) * $Altura + [int](($Altura - $alturaFoto) / 2)
        $grafico.DrawImage($foto, $x, $y, $larguraFoto, $alturaFoto)
        $grafico.FillRectangle([System.Drawing.Brushes]::Black, $x, $y, 52, 34)
        $grafico.DrawString([string]$numero, $fonte, [System.Drawing.Brushes]::Yellow, $x + 2, $y + 1)
        $indice += "{0}`t{1}`t{2}x{3}" -f $numero, $fotos[$inicio + $k].Name, $foto.Width, $foto.Height
        $foto.Dispose()
    }
    $nome = Join-Path $Saida ("folha_{0:D2}.jpg" -f [int]([Math]::Floor($inicio / $porFolha) + 1))
    $folha.Save($nome, [System.Drawing.Imaging.ImageFormat]::Jpeg)
    $grafico.Dispose(); $folha.Dispose()
}
$indice | Set-Content -Encoding utf8 (Join-Path $Saida 'indice.txt')
$indice
