@echo off
echo Iniciando Autocofin - M11
cd /d "%~dp0"
python main.py --title="M11"
if %ERRORLEVEL% NEQ 0 (
    echo Erro ao executar. Pressione qualquer tecla para sair...
    pause >nul
)
