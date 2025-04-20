@echo off
echo Iniciando Autocofin - m10
cd /d "%~dp0"
python main.py --title="m10"
if %ERRORLEVEL% NEQ 0 (
    echo Erro ao executar. Pressione qualquer tecla para sair...
    pause >nul
)
