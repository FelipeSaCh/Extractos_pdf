@echo off
setlocal
cd /d "%~dp0"
set ICONO=assets\logo.ico

if exist ".venv\Scripts\activate.bat" call ".venv\Scripts\activate.bat"
for /f "delims=" %%v in ('python -c "from version import __version__; print(__version__)"') do set VERSION=%%v
echo Compilando version %VERSION%...

python -m nuitka app_gui.py ^
  --standalone ^
  --msvc=latest ^
  --assume-yes-for-downloads ^
  --enable-plugin=tk-inter ^
  --include-package-data=pdfminer ^
  --windows-console-mode=disable ^
  --windows-icon-from-ico=%ICONO% ^
  --include-data-dir=assets=assets ^
  --output-filename=ExtractosPDF.exe ^
  --output-dir=dist ^
  --remove-output ^
  --company-name="Asesores Contables y Tributarios" ^
  --product-name="Extractos PDF" ^
  --file-description="Extractos PDF" ^
  --product-version=%VERSION% ^
  --file-version=%VERSION%

if errorlevel 1 (
  echo.
  echo *** La compilacion fallo. Revisa los mensajes anteriores. ***
  pause
  exit /b 1
)

set ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe
if not exist "%ISCC%" set ISCC=%ProgramFiles%\Inno Setup 7\ISCC.exe
if not exist "%ISCC%" set ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe
if not exist "%ISCC%" (
  echo.
  echo Programa compilado en dist\app_gui.dist. No se encontro Inno Setup: compila instalador.iss manualmente.
  pause
  exit /b 0
)
"%ISCC%" /DVersion=%VERSION% instalador.iss
if errorlevel 1 (
  echo *** Fallo la creacion del instalador. ***
) else (
  echo.
  echo Instalador listo: instaladores\ExtractosPDF_Setup_%VERSION%.exe
)
pause
