@echo off
rem === Crea el acceso directo "Relevamiento" en el Escritorio ===
rem Se ejecuta UNA sola vez, en la PC del cliente, despues de copiar la carpeta.
cd /d "%~dp0"

if not exist "%~dp0Relevamiento.exe" (
    echo No se encontro Relevamiento.exe en esta carpeta.
    echo Copia este archivo dentro de la carpeta del programa y volve a ejecutarlo.
    pause
    exit /b 1
)

set "CARPETA=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$c = $env:CARPETA;" ^
  "$s = New-Object -ComObject WScript.Shell;" ^
  "$destino = Join-Path $s.SpecialFolders('Desktop') 'Relevamiento.lnk';" ^
  "$l = $s.CreateShortcut($destino);" ^
  "$l.TargetPath = Join-Path $c 'Relevamiento.exe';" ^
  "$l.WorkingDirectory = $c;" ^
  "$l.IconLocation = (Join-Path $c 'Relevamiento.exe') + ',0';" ^
  "$l.Description = 'Relevamiento de Instancia Grupal';" ^
  "$l.Save();" ^
  "Write-Output ('Acceso directo creado en: ' + $destino)"

if errorlevel 1 (
    echo.
    echo No se pudo crear el acceso directo.
    echo Alternativa: clic derecho en Relevamiento.exe ^> Enviar a ^> Escritorio.
    pause
    exit /b 1
)

echo.
echo Listo. Ya podes abrir el programa desde el icono "Relevamiento" del Escritorio.
pause
