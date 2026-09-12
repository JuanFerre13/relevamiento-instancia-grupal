@echo off
rem === Relevamiento de Instancia Grupal - doble clic para abrir ===
cd /d "%~dp0"
where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw codigo\app.py
    exit /b
)
where python >nul 2>nul
if %errorlevel%==0 (
    python codigo\app.py
    if errorlevel 1 pause
    exit /b
)
echo No se encontro Python en esta computadora.
echo Pedile a quien instalo el sistema que ejecute "instalar.bat".
pause
