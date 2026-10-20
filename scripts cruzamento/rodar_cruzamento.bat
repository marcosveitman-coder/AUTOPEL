@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

echo === 1/3 Lendo as bases de faturamento (fat *.xlsx) - leva uns 10 a 15 minutos ===
python extract.py 4 || goto erro

echo.
echo === 2/3 Cruzando com supermercados.xlsx ===
python match.py || goto erro

echo.
echo === 3/3 Gerando "supermercados - cruzamento faturamento.xlsx" ===
python build.py || goto erro

echo.
echo Concluido.
pause
exit /b 0

:erro
echo.
echo *** Ocorreu um erro - veja a mensagem acima. ***
pause
exit /b 1
