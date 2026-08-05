@echo off
setlocal
cd /d "%~dp0"

echo ======================================================
echo GENERADOR DEL EJECUTABLE - CREDITOS Y CORREOS V2
echo ======================================================

where py >nul 2>nul
if %errorlevel%==0 (
  set "PY=py"
) else (
  where python >nul 2>nul
  if %errorlevel%==0 (
    set "PY=python"
  ) else (
    echo.
    echo No se encontro Python instalado.
    echo Instale Python 3.11 o superior y marque Add Python to PATH.
    pause
    exit /b 1
  )
)

%PY% -m pip install --upgrade pip
if errorlevel 1 goto :error
%PY% -m pip install openpyxl pywin32 pyinstaller
if errorlevel 1 goto :error

%PY% -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name Actualizar_Creditos_y_Correos_V2 ^
  --hidden-import=pythoncom ^
  --hidden-import=pywintypes ^
  --hidden-import=win32com ^
  --hidden-import=win32com.client ^
  actualizar_creditos_app_con_correos_v2.pyw
if errorlevel 1 goto :error

echo.
echo EJECUTABLE CREADO:
echo %CD%\dist\Actualizar_Creditos_y_Correos_V2.exe
echo.
pause
exit /b 0

:error
echo.
echo Ocurrio un error durante la generacion.
pause
exit /b 1
