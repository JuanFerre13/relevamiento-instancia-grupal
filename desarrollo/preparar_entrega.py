# -*- coding: utf-8 -*-
r"""
Arma el ZIP que se le manda al cliente.

Existe porque zipear `dist\Relevamiento` a mano manda la cuenta de servicio de
Google y la clave de la API por el canal que sea (mail, WhatsApp, Drive). La
convención del proyecto es que esos archivos NUNCA viajan en un ZIP de
distribución, así que acá se filtran por nombre (constante EXCLUIR) y el
filtrado no depende de en qué carpeta hayan quedado.

Todo cuelga de una carpeta `Relevamiento/` adentro del ZIP para que al
descomprimir no se desparrame en la carpeta de Descargas.

Uso:  python desarrollo\preparar_entrega.py
"""

import sys
import zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
ORIGEN = RAIZ / "dist" / "Relevamiento"
DESTINO = RAIZ / "Relevamiento_para_entregar.zip"

# Carpeta raíz dentro del ZIP.
CARPETA = "Relevamiento"

# Archivos que NO se distribuyen. Se compara por NOMBRE de archivo, no por
# ruta, así que da igual dónde estén:
#   * credenciales.json / ajustes.json  -> secretos (cuenta de servicio, api_key)
#   * registro.log                      -> lleva el ID de la planilla
#   * entrevistas_completadas.json      -> datos de trabajo de esta PC
#   * autotest_resultado.txt            -> salida de diagnóstico
EXCLUIR = {
    "credenciales.json",
    "ajustes.json",
    "registro.log",
    "entrevistas_completadas.json",
    "autotest_resultado.txt",
}

LEEME = """RELEVAMIENTO DE INSTANCIA GRUPAL
================================

Para dejarlo funcionando (una sola vez):

1. Descomprimir esta carpeta en un lugar definitivo de la computadora.
   Por ejemplo:  C:\\Relevamiento
   Si mas adelante se mueve la carpeta de lugar, hay que repetir el paso 3.

2. Copiar dentro de esta carpeta el archivo "credenciales.json", que se
   entrega aparte. Tiene que quedar al lado de Relevamiento.exe.

3. Doble clic en "crear_acceso_directo.bat".
   Eso deja el icono "Relevamiento" en el Escritorio.

Listo: a partir de ahi el programa se abre siempre desde ese icono.

La primera vez pide la configuracion (el ID de la planilla de Google y la
clave de la IA). Se cargan una sola vez y quedan guardadas.

Los informes se guardan en la carpeta "InformesIntervenciones" del Escritorio.

Notas
-----
* Si Windows muestra un aviso azul diciendo que no reconoce el programa,
  elegir "Mas informacion" y despues "Ejecutar de todas formas".
* Si el antivirus lo bloquea, hay que autorizarlo: es un falso positivo
  habitual con este tipo de programas.
* No hace falta instalar Python ni ningun otro programa.
"""


def main():
    if not ORIGEN.is_dir():
        print(f"No existe {ORIGEN}.")
        print(r"Primero hay que construir el programa: desarrollo\construir_exe.bat")
        return 1

    if DESTINO.exists():
        DESTINO.unlink()

    incluidos = 0
    omitidos = []
    with zipfile.ZipFile(DESTINO, "w", zipfile.ZIP_DEFLATED) as z:
        for archivo in sorted(ORIGEN.rglob("*")):
            if archivo.is_dir():
                continue
            if archivo.name in EXCLUIR:
                omitidos.append(archivo.name)
                continue
            relativa = archivo.relative_to(ORIGEN)
            z.write(archivo, f"{CARPETA}/{relativa.as_posix()}")
            incluidos += 1
        # El LEEME se escribe con BOM porque el cliente lo abre con el Bloc de
        # notas de Windows.
        z.writestr(f"{CARPETA}/LEEME.txt", "\ufeff" + LEEME)
        incluidos += 1

    mb = DESTINO.stat().st_size / (1024 * 1024)
    print(f"Listo: {DESTINO}")
    print(f"  {incluidos} archivos, {mb:.0f} MB")
    if omitidos:
        print(f"  Excluidos (no se distribuyen): {', '.join(sorted(set(omitidos)))}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
