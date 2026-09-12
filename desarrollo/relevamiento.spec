# -*- mode: python ; coding: utf-8 -*-
r"""
Empaquetado del programa con PyInstaller (modo ONEDIR).

Se arma una carpeta `dist\Relevamiento\` con el .exe y un `_internal\` al
lado. Es onedir y no onefile a propósito: onefile se descomprime entero a un
temporal en cada arranque y la ventana tardaba 3,1 s en aparecer sin ninguna
señal; onedir abre en menos de 1 s. La contra (hay que entregar una carpeta y
no un archivo) se resuelve con `crear_acceso_directo.bat`.

RUTAS — el detalle que hace fallar el build si se toca:
PyInstaller resuelve el SCRIPT relativo al .spec, pero los ORÍGENES de `datas`
relativos al directorio actual. Como este .spec vive en `desarrollo/` y el
programa se construye parado en la RAÍZ, mezclar los dos criterios hacía que
buscara `desarrollo/codigo/app.py`. Por eso acá TODAS las rutas se arman
absolutas desde SPECPATH (la carpeta del .spec), y nada depende de dónde se
haya parado quien lanza el build.

Los destinos de `datas` reproducen la estructura de la raíz dentro del .exe
(`recursos/` y `web/` cuelgan de sys._MEIPASS), que es lo que espera
`rutas.recurso()`: se llama `recurso("recursos/config.json")` y `recurso("web")`,
así que la ruta es idéntica empaquetado y como script.
"""

import os

from PyInstaller.utils.hooks import collect_all

# SPECPATH es la carpeta de este .spec (desarrollo/); la raíz es su padre.
RAIZ = os.path.dirname(SPECPATH)


def r(*partes):
    return os.path.join(RAIZ, *partes)


# --- Recursos embebidos de solo lectura -------------------------------------
# `web/` va acá y no en `recursos/` aunque para PyInstaller sean lo mismo: son
# 1.600 líneas de código fuente y la convención del proyecto es que los cambios
# visuales van en esa carpeta.
datas = [
    (r("recursos"), "recursos"),
    (r("web"), "web"),
]
binaries = []
hiddenimports = []

# webview/clr_loader/pythonnet cargan cosas por reflexión (el backend WebView2
# es .NET), así que hay que recolectarlos enteros o el .exe abre sin ventana.
# Las librerías de datos van igual porque traen archivos que no son .py
# (openpyxl sus plantillas, anthropic/certifi los certificados).
for paquete in (
    "webview",
    "clr_loader",
    "pythonnet",
    "gspread",
    "google.auth",
    "openpyxl",
    "anthropic",
    "docx",
):
    try:
        d, b, h = collect_all(paquete)
    except Exception as e:  # un paquete que no esté no debe voltear el build
        print(f"[spec] aviso: no se pudo recolectar {paquete}: {e}")
        continue
    datas += d
    binaries += b
    hiddenimports += h


a = Analysis(
    [r("codigo", "app.py")],
    pathex=[r("codigo"), RAIZ],  # los 4 módulos de codigo/ se importan planos
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Sin exclusiones: el build que se entregó incluye PIL y setuptools
    # (los arrastra alguna de las librerías recolectadas), así que no se
    # podan acá para no divergir de lo ya probado en la PC del cliente.
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,  # onedir: los binarios quedan afuera, en _internal/
    name="Relevamiento",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # UPX apagado: comprimir el .exe dispara falsos positivos de antivirus
    console=False,  # sin consola: es una ventana, no un programa de terminal
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=r("recursos", "icono.ico"),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="Relevamiento",
)
