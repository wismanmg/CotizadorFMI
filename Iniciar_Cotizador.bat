@echo off
REM Lanzador del Cotizador FMI (sin necesidad de compilar a .exe)
cd /d "%~dp0"
python app.py
if errorlevel 1 pause
