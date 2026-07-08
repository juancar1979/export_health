@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

cd /d "%~dp0"

echo ============================================================
echo  Exportador de Apple Salud a Excel
echo ============================================================
echo.

set "SCRIPT=exportar_salud_completo.py"

if not exist "%SCRIPT%" (
    echo ERROR: No encuentro %SCRIPT% en esta carpeta.
    echo Copia este .bat en la misma carpeta que %SCRIPT%.
    echo.
    pause
    exit /b 1
)

where python >nul 2>nul
if %errorlevel%==0 (
    set "PY=python"
) else (
    where py >nul 2>nul
    if %errorlevel%==0 (
        set "PY=py -3"
    ) else (
        echo ERROR: No encuentro Python instalado.
        echo Instala Python desde https://www.python.org/downloads/
        echo Importante: marca la opcion "Add Python to PATH" durante la instalacion.
        echo.
        pause
        exit /b 1
    )
)

echo Comprobando dependencia openpyxl...
%PY% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo openpyxl no esta instalado. Instalando...
    %PY% -m pip install openpyxl
    if errorlevel 1 (
        echo ERROR: No se ha podido instalar openpyxl.
        echo Prueba manualmente: python -m pip install openpyxl
        echo.
        pause
        exit /b 1
    )
)

echo.
echo Version del script:
%PY% "%SCRIPT%" --version
echo.
echo Deja en blanco cualquier campo para usar el valor por defecto.
echo.

set /p "XML=Ruta a export.xml o carpeta apple_health_export [Enter = buscar automaticamente]: "
set "XML=%XML:"=%"

set /p "INICIO=Fecha inicio YYYY-MM-DD [Enter = sin limite]: "
set /p "FIN=Fecha fin YYYY-MM-DD [Enter = sin limite]: "
set /p "DETALLE=Detalle dia/hora/min/todos [Enter = hora]: "
if "%DETALLE%"=="" set "DETALLE=hora"

set /p "SALIDA=Nombre Excel salida .xlsx [Enter = automatico]: "
set "SALIDA=%SALIDA:"=%"

echo.
echo Ejecutando exportacion...
echo.

%PY% "%SCRIPT%" --xml "%XML%" --inicio "%INICIO%" --fin "%FIN%" --detalle "%DETALLE%" --salida "%SALIDA%"

set "STATUS=%errorlevel%"
echo.
if "%STATUS%"=="0" (
    echo Exportacion finalizada correctamente.
) else (
    echo La exportacion ha terminado con error. Codigo: %STATUS%
)
echo.
pause
exit /b %STATUS%
