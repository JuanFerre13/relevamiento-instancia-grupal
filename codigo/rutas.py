# -*- coding: utf-8 -*-
"""
Resolución de rutas que funciona igual como script de Python y como .exe
empaquetado con PyInstaller.

Hay dos tipos de archivo con ubicaciones distintas cuando el programa corre
empaquetado:

  * RECURSOS embebidos (solo lectura): todo lo que hay en `recursos/`
    (config.json, estilo_informe.md, plantilla.docx, informes_ejemplo/) y la
    interfaz en `web/`. PyInstaller los descomprime a una carpeta temporal
    (sys._MEIPASS) CONSERVANDO esas subcarpetas, así que se leen igual en los
    dos modos: `recurso("recursos/config.json")`, `recurso("web")`.

  * ARCHIVOS EXTERNOS que el usuario pone/edita: credenciales.json y
    ajustes.json. Viven JUNTO al .exe. Se acceden con `archivo_externo(...)`.

  * RESULTADOS (los informes y planillas generados): van a la carpeta
    "InformesIntervenciones" del Escritorio del usuario. Se obtiene con
    `carpeta_informes(...)`.

Corriendo como script normal (sin empaquetar) `recurso(...)` y `archivo_externo(...)`
devuelven la carpeta del proyecto; `carpeta_informes(...)` siempre apunta al
Escritorio.
"""

import os
import sys
from pathlib import Path

# Nombre de la carpeta de resultados en el Escritorio.
CARPETA_RESULTADOS = "InformesIntervenciones"


def _empaquetado():
    return getattr(sys, "frozen", False)


def _escritorio():
    """Ruta real del Escritorio del usuario.

    En Windows contempla la redirección a OneDrive consultando el registro
    (User Shell Folders), que es la fuente confiable. Si algo falla, prueba
    OneDrive\\Desktop y finalmente ~/Desktop."""
    if sys.platform.startswith("win"):
        try:
            import winreg
            clave = r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, clave) as k:
                valor, _ = winreg.QueryValueEx(k, "Desktop")
                ruta = Path(os.path.expandvars(valor))
                if ruta.parent.exists():
                    return ruta
        except Exception:
            pass
    hogar = Path.home()
    for candidata in (hogar / "OneDrive" / "Desktop", hogar / "Desktop"):
        if candidata.exists():
            return candidata
    return hogar / "Desktop"


def carpeta_informes():
    """Carpeta de resultados en el Escritorio (se crea si no existe)."""
    carpeta = _escritorio() / CARPETA_RESULTADOS
    carpeta.mkdir(parents=True, exist_ok=True)
    return carpeta


def _raiz_proyecto():
    """Raíz del proyecto corriendo como script.

    Este módulo vive en `codigo/`, así que hay que subir un nivel: la raíz es
    la carpeta que tiene `codigo/`, `recursos/` y `web/` adentro. Si algún día
    se mueve rutas.py de carpeta, ESTE es el lugar a corregir."""
    return Path(__file__).resolve().parent.parent


def dir_recursos():
    """Carpeta de los recursos embebidos (solo lectura).

    Es la RAÍZ, no `recursos/`: adentro cuelgan tanto `recursos/` como `web/`,
    y los nombres que se le pasan a `recurso(...)` incluyen esa subcarpeta.
    Empaquetado, la misma estructura se reproduce dentro de sys._MEIPASS
    (ver los destinos de `datas` en relevamiento.spec)."""
    if _empaquetado():
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    return _raiz_proyecto()


def dir_externo():
    """Carpeta junto al ejecutable (o al proyecto) para archivos del usuario y salida."""
    if _empaquetado():
        return Path(sys.executable).resolve().parent
    return _raiz_proyecto()


def recurso(nombre):
    """Ruta a un recurso embebido de solo lectura, relativa a la raíz.

    El nombre lleva la subcarpeta: `recurso("recursos/config.json")`."""
    return dir_recursos() / nombre


def archivo_externo(nombre):
    """Ruta a un archivo externo (junto al .exe) que el usuario pone o se genera."""
    return dir_externo() / nombre
