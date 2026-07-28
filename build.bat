@echo off
REM ================================================================
REM  Compila el Cotizador FMI a un unico ejecutable Windows (.exe)
REM  Usa un entorno virtual limpio para que el .exe sea liviano.
REM ================================================================
cd /d "%~dp0"
echo Preparando entorno de compilacion...

if not exist ".buildenv\Scripts\python.exe" (
    python -m venv .buildenv
)
call ".buildenv\Scripts\python.exe" -m pip install --upgrade pip
call ".buildenv\Scripts\python.exe" -m pip install openpyxl pdfplumber reportlab customtkinter pypdfium2 rapidocr-onnxruntime pywin32 pyinstaller

echo Compilando Cotizador FMI...
REM  --onedir (carpeta) en vez de --onefile: los antivirus casi no lo bloquean
REM  porque el .exe NO se auto-desempaqueta en una carpeta temporal al abrir.
call ".buildenv\Scripts\python.exe" -m PyInstaller --noconfirm --onedir --windowed ^
  --name "CotizadorFMI" ^
  --add-data "datos;datos" ^
  --add-data "assets;assets" ^
  --add-data "plantilla;plantilla" ^
  --collect-all customtkinter ^
  --collect-all pdfplumber ^
  --collect-all pdfminer ^
  --collect-all reportlab ^
  --collect-all rapidocr_onnxruntime ^
  --collect-all pypdfium2 ^
  --collect-all onnxruntime ^
  --hidden-import win32com ^
  --hidden-import win32com.client ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --hidden-import win32timezone ^
  app.py

REM ---- Instala la app en la carpeta de uso diario --------------------
REM  Copia el .exe y su carpeta _internal SIN tocar 'datos' (config y
REM  bitacora) ni 'salidas' (tus cotizaciones), que viven en esa carpeta.
set "APP=%USERPROFILE%\Documents\CotizadorFMI"
if not exist "%APP%" mkdir "%APP%"
copy /Y "dist\CotizadorFMI\CotizadorFMI.exe" "%APP%\CotizadorFMI.exe" >nul
if exist "%APP%\_internal" rmdir /S /Q "%APP%\_internal"
xcopy /E /I /Y /Q "dist\CotizadorFMI\_internal" "%APP%\_internal" >nul
copy /Y "INSTRUCCIONES.md" "%APP%\INSTRUCCIONES.md" >nul
if exist "dist\CotizadorFMI\Desbloquear (ejecutar primero).bat" copy /Y "dist\CotizadorFMI\Desbloquear (ejecutar primero).bat" "%APP%\" >nul

echo.
echo ================================================================
echo  Listo. App instalada (formato carpeta) en:
echo    %APP%
echo  Abrir con:  %APP%\CotizadorFMI.exe
echo.
echo  Para llevarla a otra PC: copia la carpeta COMPLETA.
echo  Tus cotizaciones: %APP%\salidas   Config: %APP%\datos
echo ================================================================
pause
