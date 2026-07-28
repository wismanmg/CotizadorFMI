@echo off
title Cotizador FMI - servidor web
cd /d "%~dp0"

echo ==========================================================
echo   COTIZADOR FMI - version web
echo ==========================================================
echo.
echo   Iniciando... no cierres esta ventana mientras trabajas.
echo.

REM Instala Flask la primera vez si no esta
python -c "import flask" 2>nul
if errorlevel 1 (
    echo   Instalando componentes por unica vez, espera...
    python -m pip install --quiet flask
)

REM Abre el navegador y arranca el servidor
start "" http://localhost:8000
python app.py

echo.
echo   El servidor se detuvo. Podes cerrar esta ventana.
pause
