# Pinta de vermelho, na aba "Alessandra Dayan", as empresas que nao estao no mapeamento (lista de confere_alessandra.py).
# Uso: powershell -File pinta_alessandra.ps1 -Arquivo <copia.xlsx> -Lista dados\alessandra_vermelho.txt
param([string]$Arquivo, [string]$Lista)
$ErrorActionPreference = 'Stop'
$faixas = (Get-Content -Raw $Lista).Trim().Split(',')
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
  $wb = $xl.Workbooks.Open($Arquivo)
  $ws = $wb.Worksheets.Item('Alessandra Dayan')
  foreach ($f in $faixas) { $rg = $ws.Range($f); $rg.Interior.Color = 255; $rg.Font.Color = 16777215; $rg.Font.Bold = $true }  # fundo vermelho, letra branca
  $wb.Save(); 'pintadas: ' + $faixas.Count
  $wb.Close($false)
} finally { $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
