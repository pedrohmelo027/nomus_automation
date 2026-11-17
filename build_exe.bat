@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo Verificando PyInstaller...
py -m pip install pyinstaller
echo Empacotando executavel...
py -m PyInstaller --noconfirm --onedir --windowed --name "Automacao_BOM_Nomus" ^
    --add-data "config;config" ^
    --add-data "dados_referencia;dados_referencia" ^
    src\app.py
echo.
echo Executavel gerado em: dist\Automacao_BOM_Nomus\Automacao_BOM_Nomus.exe
pause
