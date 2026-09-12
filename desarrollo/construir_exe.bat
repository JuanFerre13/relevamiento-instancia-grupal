@echo off
rem === Construye Relevamiento.exe en dist\Relevamiento (modo onedir) ===
rem Se ejecuta desde la carpeta desarrollo, pero trabaja parado en la RAIZ:
rem el .spec resuelve sus rutas solo, pero dist\ y build\ se crean en el
rem directorio actual.
setlocal
cd /d "%~dp0.."

rem PyInstaller borra dist\Relevamiento ENTERA al reconstruir, y ahi adentro
rem viven los archivos del usuario. Se resguardan antes y se reponen despues.
rem Correr PyInstaller a mano NO tiene esta red: asi se perdieron una vez las
rem credenciales.
set "RESGUARDO=%TEMP%\relevamiento_resguardo"
if exist "%RESGUARDO%" rd /s /q "%RESGUARDO%"
mkdir "%RESGUARDO%" >nul 2>&1
for %%A in (credenciales.json ajustes.json entrevistas_completadas.json) do (
    if exist "dist\Relevamiento\%%A" copy /y "dist\Relevamiento\%%A" "%RESGUARDO%\%%A" >nul
)

echo Construyendo el programa. Tarda unos minutos...
echo.
python -m PyInstaller --noconfirm --clean desarrollo\relevamiento.spec
if errorlevel 1 (
    echo.
    echo FALLO la construccion. Revisar los mensajes de arriba.
    pause
    exit /b 1
)

rem El acceso directo viaja adentro de la carpeta que se entrega, para que el
rem cliente no vea nunca esta carpeta.
copy /y "desarrollo\crear_acceso_directo.bat" "dist\Relevamiento\" >nul

rem Reponer los archivos del usuario que se resguardaron.
for %%A in (credenciales.json ajustes.json entrevistas_completadas.json) do (
    if exist "%RESGUARDO%\%%A" copy /y "%RESGUARDO%\%%A" "dist\Relevamiento\%%A" >nul
)
rd /s /q "%RESGUARDO%" >nul 2>&1

echo.
echo Listo: dist\Relevamiento\Relevamiento.exe
echo Para armar el ZIP del cliente: desarrollo\preparar_entrega.bat
pause
