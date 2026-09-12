@echo off
rem === Instalador (lo ejecuta una sola vez la persona que configura el sistema) ===
cd /d "%~dp0"
where python >nul 2>nul
if not %errorlevel%==0 (
    echo Primero instala Python desde https://www.python.org/downloads/
    echo IMPORTANTE: marcar la casilla "Add Python to PATH" durante la instalacion.
    pause
    exit /b
)
echo Instalando componentes necesarios...
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo Hubo un problema instalando los componentes. Revisa la conexion a internet.
    pause
    exit /b
)
echo.
echo Listo. Ya se puede abrir el sistema con doble clic en "Iniciar.bat".
pause
