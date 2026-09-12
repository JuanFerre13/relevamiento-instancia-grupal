# -*- coding: utf-8 -*-
r"""
Genera el icono del programa (recursos\icono.ico).

El diseño es el semáforo del relevamiento: tres barras sobre fondo azul, con
los MISMOS tres colores que usa el XLSX para los porcentajes (rojo FF7C5C,
amarillo FFFF66, verde 92D050), de menor a mayor. Así el icono del Escritorio
y la planilla dicen lo mismo.

Pillow es dependencia SOLO de desarrollo: no está en requirements.txt y el
programa no la importa nunca. El .ico queda versionado en recursos/, así que
esto se corre únicamente si se quiere cambiar el icono.

Uso:
    python desarrollo\crear_icono.py                 -> escribe icono_nuevo.ico
    python desarrollo\crear_icono.py --sobrescribir  -> pisa recursos\icono.ico

Se pide --sobrescribir a propósito: el .ico versionado es el que ya está
probado a 16 px, y no conviene pisarlo por accidente al pasar por acá.
"""

import sys
from pathlib import Path

try:
    from PIL import Image, ImageDraw
except ImportError:
    print("Falta Pillow (solo hace falta para regenerar el icono):")
    print("    pip install Pillow")
    sys.exit(1)

RAIZ = Path(__file__).resolve().parent.parent

# Paleta. Los tres primeros son los del semáforo del XLSX (procesar.py).
AZUL = (31, 78, 121, 255)       # 1F4E79, el azul de la papelería institucional
ROJO = (255, 124, 92, 255)      # FF7C5C
AMARILLO = (255, 255, 102, 255) # FFFF66
VERDE = (146, 208, 80, 255)     # 92D050
BASE = (229, 234, 238, 255)     # E5EAEE, la línea de apoyo de las barras

# Tamaños que guarda el .ico. Windows toma 16 px para la barra de tareas y 256
# para la vista de iconos grandes.
TAMANOS = [16, 24, 32, 48, 64, 128, 256]

# Todo se dibuja en fracciones de 1 y después se escala, así el diseño es el
# mismo en todos los tamaños. Los valores salen de dividir por 32 el diseño
# original, que se hizo sobre una grilla de 32x32.
RADIO = 6 / 32          # redondeo de las esquinas
MARGEN = 1 / 32         # aire entre el borde del icono y el cuadrado
PISO = 25 / 32          # dónde apoyan las barras
ALTO_BASE = 1.2 / 32    # grosor de la línea de apoyo
BASE_X = (5 / 32, 27 / 32)

# (x inicial, x final, y del tope, color) — de menor a mayor, como el semáforo.
BARRAS = [
    (6 / 32, 11 / 32, 18 / 32, ROJO),
    (13 / 32, 18 / 32, 14 / 32, AMARILLO),
    (21 / 32, 26 / 32, 10 / 32, VERDE),
]

# Se dibuja 8 veces más grande y se achica: es la forma barata de conseguir
# bordes suaves sin depender del antialiasing de ImageDraw, que no lo tiene
# para rectángulos.
ESCALA = 8


def dibujar(lado):
    grande = lado * ESCALA
    im = Image.new("RGBA", (grande, grande), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)

    def p(v):
        return v * grande

    d.rounded_rectangle(
        [p(MARGEN), p(MARGEN), grande - p(MARGEN) - 1, grande - p(MARGEN) - 1],
        radius=p(RADIO),
        fill=AZUL,
    )

    for x0, x1, tope, color in BARRAS:
        d.rectangle([p(x0), p(tope), p(x1), p(PISO)], fill=color)

    d.rectangle(
        [p(BASE_X[0]), p(PISO), p(BASE_X[1]), p(PISO + ALTO_BASE)],
        fill=BASE,
    )

    return im.resize((lado, lado), Image.LANCZOS)


def main():
    sobrescribir = "--sobrescribir" in sys.argv
    destino = RAIZ / "recursos" / "icono.ico" if sobrescribir else RAIZ / "icono_nuevo.ico"

    base = dibujar(max(TAMANOS))
    base.save(destino, format="ICO", sizes=[(t, t) for t in TAMANOS])

    print(f"Icono escrito: {destino}")
    print(f"  tamanos: {', '.join(f'{t}x{t}' for t in TAMANOS)}")
    if not sobrescribir:
        print()
        print("Se escribio en un archivo aparte para no pisar el icono que ya esta")
        print("probado. Para reemplazarlo de verdad:")
        print(r"    python desarrollo\crear_icono.py --sobrescribir")
        print("Despues hay que reconstruir el .exe para que tome el icono nuevo.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
