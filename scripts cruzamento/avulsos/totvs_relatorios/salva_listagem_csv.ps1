# Relatorios do TOTVS (Gera Dados Excel): o WebAgent grava C:\Excel\listagem.xlsx e o abre no Excel.
# Este script espera a listagem ficar mais nova que -Desde, salva uma copia como CSV UTF-8
# (separador ";" do pt-BR, igual ao "Salvar como CSV UTF-8" do Excel) em -Destino e fecha a
# janela da listagem no Excel, para o TOTVS poder gravar o proximo relatorio.
# Uso: salva_listagem_csv.ps1 -Destino "<arquivo.csv>" -Desde "2026-11-25 11:00:00" [-TimeoutMin 60]
#      salva_listagem_csv.ps1 -SoFechar      (so fecha a listagem aberta no Excel)
param(
  [string]$Destino,
  [datetime]$Desde = [datetime]::MinValue,
  [int]$TimeoutMin = 60,
  [string]$Origem = "C:\Excel\listagem.xlsx",
  [switch]$SoFechar
)

function Close-Listagem {
  # O Excel aberto pelo WebAgent nao aparece para GetActiveObject: fecha a janela "listagem - Excel"
  # com WM_CLOSE (se houver alteracao, o proprio Excel pergunta; nada se perde).
  param([string]$Caminho)
  $nome = [IO.Path]::GetFileNameWithoutExtension($Caminho)
  $ps = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like "$nome*" })
  foreach ($p in $ps) { [void]$p.CloseMainWindow() }
  Start-Sleep -Seconds 4
  try { $fs = [IO.File]::Open($Caminho, 'Open', 'Read', 'None'); $fs.Close(); $livre = $true } catch { $livre = $false }
  Write-Output ("FECHAR: janelas=" + $ps.Count + " | listagem " + $(if ($livre) { "livre" } else { "AINDA EM USO" }))
}

if ($SoFechar) { Close-Listagem $Origem; exit 0 }

$limite = (Get-Date).AddMinutes($TimeoutMin)
while ($true) {
  $it = Get-Item $Origem -ErrorAction SilentlyContinue
  if ($it -and $it.LastWriteTime -gt $Desde) { break }
  if ((Get-Date) -gt $limite) { Write-Output "TIMEOUT"; exit 1 }
  Start-Sleep -Seconds 2
}
$gerado = (Get-Item $Origem).LastWriteTime
# espera o tamanho parar de mudar (gravacao terminada)
$t0 = -1
while ($true) { Start-Sleep -Seconds 3; $t1 = (Get-Item $Origem).Length; if ($t1 -eq $t0) { break }; $t0 = $t1 }

$copia = Join-Path $env:TEMP ("listagem_" + [guid]::NewGuid().ToString('N') + ".xlsx")
$ok = $false
for ($i = 0; $i -lt 20 -and -not $ok; $i++) { try { Copy-Item $Origem $copia -ErrorAction Stop; $ok = $true } catch { Start-Sleep -Seconds 3 } }
if (-not $ok) { Write-Output "ERRO: nao consegui copiar a listagem"; exit 1 }

$tmp = [IO.Path]::ChangeExtension($copia, ".csv")
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($copia, 0, $true)
  $m = [Type]::Missing
  $wb.SaveAs($tmp, 62, $m, $m, $m, $m, 1, $m, $m, $m, $m, $true)   # 62 = xlCSVUTF8, Local
  $wb.Close($false)
} finally {
  $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
Copy-Item $tmp $Destino -Force; Remove-Item $tmp, $copia
$d = Get-Item $Destino
Write-Output ("LISTAGEM_GERADA=" + $gerado.ToString('yyyy-MM-dd HH:mm:ss'))
Write-Output ("CSV=" + $d.FullName + " | " + $d.Length + " bytes | " + $d.LastWriteTime.ToString('HH:mm:ss'))
Start-Sleep -Seconds 5
Close-Listagem $Origem
