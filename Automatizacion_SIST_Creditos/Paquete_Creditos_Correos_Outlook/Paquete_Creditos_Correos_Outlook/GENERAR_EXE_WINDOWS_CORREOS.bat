@echo off
setlocal
cd /d "%~dp0"

echo =========================================================
echo  GENERADOR DEL EJECUTABLE - CREDITOS Y CORREOS OUTLOOK
echo =========================================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    set PY=py
) else (
    where python >nul 2>nul
    if %errorlevel% neq 0 (
        echo No se encontro Python instalado.
        echo Instale Python 3.11 o superior desde python.org y marque
        echo la opcion "Add Python to PATH" durante la instalacion.
        pause
        exit /b 1
    )
    set PY=python
)

%PY% -m pip install --upgrade pip
if %errorlevel% neq 0 goto :error

%PY% -m pip install openpyxl pywin32 pyinstaller
if %errorlevel% neq 0 goto :error

%PY% -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "Actualizar_Creditos_y_Correos" ^
  --hidden-import win32com.client ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  actualizar_creditos_app_con_correos.pyw
if %errorlevel% neq 0 goto :error

echo.
echo =========================================================
echo EJECUTABLE CREADO CORRECTAMENTE
echo Ubicacion: %~dp0dist\Actualizar_Creditos_y_Correos.exe
echo =========================================================
start "" "%~dp0dist"
pause
exit /b 0

:error
echo.
echo No se pudo generar el ejecutable. Revise el mensaje anterior.
pause
exit /b 1
