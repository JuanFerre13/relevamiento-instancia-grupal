@echo off
rem === Arma Relevamiento_para_entregar.zip (sin credenciales ni ajustes) ===
setlocal
cd /d "%~dp0.."

if not exist "dist\Relevamiento\Relevamiento.exe" (
    echo No esta construido el programa.
    echo Primero hay que ejecutar: desarrollo\construir_exe.bat
    pause
    exit /b 1
)

python desarrollo\preparar_entrega.py
if errorlevel 1 (
    echo.
    echo No se pudo armar el ZIP.
    pause
    exit /b 1
)

echo.
pause
