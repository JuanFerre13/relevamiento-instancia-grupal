# -*- coding: utf-8 -*-
"""
Aplicación de escritorio para el Relevamiento de Instancia Grupal.

Es una ventana nativa de Windows cuyo interior se dibuja con HTML/CSS/JS
(carpeta `web/`), usando pywebview sobre el motor WebView2 que ya viene con
el sistema. No hay servidor, ni puerto, ni navegador: se abre con doble clic
igual que antes.

Pensada para usuarios sin conocimientos técnicos: se completan dos o tres
datos y se aprieta un botón.

Reparto de tareas:
  * `web/`  → todo lo visual y la interacción (ver web/app.js).
  * este archivo → el puente con la lógica de negocio, que sigue viviendo
    intacta en procesar.py y generar_informe.py.

La configuración técnica (planilla de Google, clave de la IA) se carga una
sola vez desde el botón "Configuración" y queda guardada en ajustes.json.
"""

import json
import logging
import logging.handlers
import os
import subprocess
import sys
import threading
import traceback
import webbrowser
from datetime import datetime

import webview

from rutas import archivo_externo, carpeta_informes, dir_externo, recurso

# Archivos que el usuario pone o edita: viven junto al .exe (o al proyecto,
# cuando corre como script).
ARCHIVO_AJUSTES = archivo_externo("ajustes.json")
ARCHIVO_CREDENCIALES = archivo_externo("credenciales.json")
# Qué entrevistas individuales ya se hicieron. Va junto al programa y no en la
# planilla porque la cuenta de servicio de Google entra como LECTORA: el
# programa no puede escribir en el Sheets. Guarda sólo la clave (un hash de la
# fila) y cuándo se marcó, así no quedan nombres ni teléfonos en el archivo.
ARCHIVO_ENTREVISTAS = archivo_externo("entrevistas_completadas.json")
# Los resultados van a la carpeta "InformesIntervenciones" del Escritorio.
CARPETA_SALIDA = carpeta_informes()
CONFIG = recurso("recursos/config.json")
INTERFAZ = recurso("web") / "index.html"

AJUSTES_VACIOS = {"sheet_id": "", "worksheet": "", "api_key": "", "modelo": ""}

# Dirección de la planilla de respuestas, que se arma con el sheet_id guardado en
# ajustes.json: así el botón de la ventana apunta siempre a la planilla que está
# configurada, sin ninguna dirección escrita a mano en el código.
URL_PLANILLA = "https://docs.google.com/spreadsheets/d/{}/edit"

TITULO = "Relevamiento de Instancia Grupal"

ARCHIVO_REGISTRO = archivo_externo("registro.log")
registro = logging.getLogger("relevamiento")
# Sin esto, mientras el registro no esté configurado Python usa su handler de
# último recurso y escupe los errores por stderr. No rompe nada, pero ensucia la
# consola de quien corre el programa como script.
registro.addHandler(logging.NullHandler())
registro.propagate = False


def configurar_registro():
    """Deja constancia en `registro.log`, junto al programa, de todo lo que se
    muestra en la ventana más los errores completos.

    Existe para poder diagnosticar a distancia: el recuadro de actividad se borra
    al cerrar el programa, así que cuando el cliente dice «no anduvo» no queda
    nada que mirar. Rota a los 3 archivos de 1 MB para que no crezca sin fin."""
    registro.setLevel(logging.INFO)
    # Se ignora el NullHandler que se agrega al importar: lo que interesa es si
    # ya hay un handler que escriba al archivo.
    if any(isinstance(h, logging.FileHandler) for h in registro.handlers):
        return
    try:
        manejador = logging.handlers.RotatingFileHandler(
            str(ARCHIVO_REGISTRO), maxBytes=1_000_000, backupCount=3, encoding="utf-8")
        manejador.setFormatter(logging.Formatter("%(asctime)s  %(message)s",
                                                 datefmt="%d/%m/%Y %H:%M:%S"))
        registro.addHandler(manejador)
        registro.info("=" * 60)
        registro.info("Programa abierto (%s)", "empaquetado" if getattr(sys, "frozen", False) else "script")
    except Exception:
        # Si no se puede escribir (carpeta de solo lectura, permisos), el
        # programa tiene que funcionar igual: el registro es una ayuda, no un
        # requisito.
        registro.addHandler(logging.NullHandler())


def ya_esta_abierto():
    """True si el programa ya está corriendo (y en ese caso le trae la ventana
    al frente).

    Sin esto, el doble clic repetido —muy habitual cuando alguien no está
    seguro de que haya arrancado— abre dos copias, y las dos escriben en la
    misma carpeta de resultados."""
    if not sys.platform.startswith("win"):
        return False
    try:
        import ctypes

        k32 = ctypes.windll.kernel32
        # El handle queda abierto a propósito mientras viva el proceso: es lo
        # que hace que la segunda copia encuentre el mutex ya creado.
        k32.CreateMutexW(None, True, "Relevamiento_InstanciaGrupal")
        if k32.GetLastError() != 183:      # ERROR_ALREADY_EXISTS
            return False
        u32 = ctypes.windll.user32
        ventana = u32.FindWindowW(None, TITULO)
        if ventana:
            u32.ShowWindow(ventana, 9)     # SW_RESTORE, por si estaba minimizada
            u32.SetForegroundWindow(ventana)
        return True
    except Exception:
        return False  # ante la duda, dejar abrir


def cargar_ajustes():
    if ARCHIVO_AJUSTES.exists():
        try:
            guardados = json.loads(ARCHIVO_AJUSTES.read_text(encoding="utf-8"))
            return {**AJUSTES_VACIOS, **guardados}
        except Exception:
            pass
    return dict(AJUSTES_VACIOS)


def guardar_ajustes(ajustes):
    ARCHIVO_AJUSTES.write_text(json.dumps(ajustes, ensure_ascii=False, indent=2), encoding="utf-8")


def cargar_entrevistas_completadas():
    """{clave: fecha en que se marcó}. Si el archivo no está o está roto, se
    arranca de cero: perder las marcas es molesto, pero no puede impedir que se
    vea la lista."""
    if ARCHIVO_ENTREVISTAS.exists():
        try:
            guardado = json.loads(ARCHIVO_ENTREVISTAS.read_text(encoding="utf-8"))
            if isinstance(guardado, dict):
                return guardado
        except Exception:
            registro.warning("No se pudo leer %s; se empieza con la lista vacía.",
                             ARCHIVO_ENTREVISTAS.name)
    return {}


def guardar_entrevistas_completadas(completadas):
    ARCHIVO_ENTREVISTAS.write_text(
        json.dumps(completadas, ensure_ascii=False, indent=2), encoding="utf-8")


def _orden_marca(marca):
    """Clave para ordenar por cuándo se marcó ('05/08/2026 11:51').

    Se guarda en el formato que se le muestra a la persona, así que para ordenar
    hay que darlo vuelta. Lo que no se entienda va último y no rompe nada."""
    try:
        return datetime.strptime(str(marca), "%d/%m/%Y %H:%M")
    except (ValueError, TypeError):
        return datetime.min


def procesar_modulo():
    """Import diferido: mantiene liviano el arranque de la ventana."""
    import procesar
    return procesar


def abrir_enlace(url):
    """Abre una dirección web en el navegador del sistema. True si pudo.

    Tiene que salir por acá y no por un enlace del HTML: la interfaz vive adentro
    de la ventana del programa, así que un <a href> la haría navegar a la página
    y el cliente se quedaría sin interfaz y sin botón para volver."""
    try:
        return bool(webbrowser.open(url))
    except Exception:
        registro.error("No se pudo abrir el navegador en %s\n%s", url, traceback.format_exc())
        return False


def abrir_carpeta(ruta):
    ruta = str(ruta)
    try:
        if sys.platform.startswith("win"):
            os.startfile(ruta)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", ruta])
        else:
            subprocess.Popen(["xdg-open", ruta])
    except Exception:
        pass


class Puente:
    """Métodos que la interfaz llama como `pywebview.api.<nombre>(...)`.

    Los que arrancan trabajo pesado devuelven enseguida y siguen en un hilo
    aparte, informando el progreso con `self.log(...)`; así la ventana nunca
    se congela.

    IMPORTANTE: los atributos internos van con guión bajo adelante. pywebview
    recorre los atributos públicos de esta clase para publicarlos al JS, y si
    encuentra la ventana se mete en los objetos COM de WebView2 hasta agotar
    la recursión, llenando la consola de errores."""

    def __init__(self):
        self._ajustes = cargar_ajustes()
        self._ventana = None
        self._trabajando = False

    # ------------------------------------------------------------- hacia el JS
    def _js(self, funcion, *args):
        """Llama a una función global del JS con argumentos ya serializados."""
        if self._ventana is None:
            return
        argumentos = ", ".join(json.dumps(a, ensure_ascii=False) for a in args)
        try:
            self._ventana.evaluate_js(f"window.{funcion}({argumentos})")
        except Exception:
            pass  # la ventana se cerró mientras trabajábamos

    def log(self, texto):
        registro.info("%s", texto)
        self._js("agregarLog", str(texto))

    def _terminar(self, ok, mensaje):
        self._trabajando = False
        self._js("terminar", ok, mensaje)

    def _fallo(self, e):
        """Un error que termina el trabajo. En la ventana va el mensaje en
        español; al archivo va además el traceback completo, que es lo único
        que sirve para diagnosticar después."""
        registro.error("FALLÓ: %s\n%s", e, traceback.format_exc())
        self.log(f"✖ Ocurrió un problema: {e}")
        self._terminar(False, str(e))

    # ------------------------------------------------------- llamadas desde JS
    def estado_inicial(self):
        """Estado para la interfaz. Se vuelve a pedir cada vez que se abre la
        ventana de Configuración: `credenciales.json` se puede haber copiado
        con el programa ya abierto, y antes esa foto quedaba vieja para siempre."""
        import generar_informe
        return {
            "ajustes": self._ajustes,
            "credenciales_ok": ARCHIVO_CREDENCIALES.exists(),
            "carpeta_credenciales": str(ARCHIVO_CREDENCIALES.parent),
            "modelo_default": generar_informe.MODELO,
        }

    def guardar_ajustes(self, ajustes):
        # Se mezcla sobre lo que ya había para no borrar claves que la ventana
        # de configuración no muestra.
        self._ajustes = {**AJUSTES_VACIOS, **self._ajustes, **(ajustes or {})}
        guardar_ajustes(self._ajustes)
        return {"ok": True}

    def cargar_opciones(self, ruta_csv=None):
        """Lee la planilla (o un CSV) y devuelve qué fechas, unidades, regiones y
        servicios hay cargados, para que la interfaz muestre listas para elegir."""
        try:
            opciones = procesar_modulo().opciones_disponibles(
                csv=ruta_csv,
                sheet_id=None if ruta_csv else (self._ajustes.get("sheet_id") or None),
                worksheet=self._ajustes.get("worksheet") or None,
                credenciales=str(ARCHIVO_CREDENCIALES),
                config_path=str(CONFIG),
                log=self.log,
            )
        except Exception as e:
            registro.error("No se pudieron leer los datos: %s\n%s", e, traceback.format_exc())
            return {"ok": False, "error": str(e)}
        return {
            "ok": True,
            "opciones": opciones,
            "nombre_archivo": os.path.basename(ruta_csv) if ruta_csv else None,
        }

    def entrevistas_pendientes(self, ruta_csv=None):
        """Quiénes pidieron una entrevista individual y todavía no se marcó como
        hecha. Se lee la planilla de nuevo cada vez que se abre la ventana, para
        que aparezcan los pedidos nuevos."""
        try:
            pedidos = procesar_modulo().entrevistas_solicitadas(
                csv=ruta_csv,
                sheet_id=None if ruta_csv else (self._ajustes.get("sheet_id") or None),
                worksheet=self._ajustes.get("worksheet") or None,
                credenciales=str(ARCHIVO_CREDENCIALES),
                config_path=str(CONFIG),
                log=self.log,
            )
        except Exception as e:
            registro.error("No se pudo leer las entrevistas: %s\n%s", e, traceback.format_exc())
            return {"ok": False, "error": str(e)}

        completadas = cargar_entrevistas_completadas()
        pendientes = [p for p in pedidos if p["clave"] not in completadas]
        # Las ya hechas viajan también, para poder deshacer una marca desde la
        # ventana. Van con la fecha en que se marcaron y las últimas primero,
        # que es el orden en que se busca la que se marcó por error.
        hechas = [{**p, "marcada_el": completadas[p["clave"]]}
                  for p in pedidos if p["clave"] in completadas]
        hechas.sort(key=lambda p: _orden_marca(p["marcada_el"]), reverse=True)
        return {"ok": True, "pendientes": pendientes, "hechas": hechas}

    def marcar_entrevista(self, clave, hecha=True):
        """Anota (o desanota) una entrevista como realizada."""
        if not clave:
            return {"ok": False, "error": "Falta la clave de la entrevista."}
        completadas = cargar_entrevistas_completadas()
        if hecha:
            completadas[clave] = datetime.now().strftime("%d/%m/%Y %H:%M")
        else:
            completadas.pop(clave, None)
        try:
            guardar_entrevistas_completadas(completadas)
        except OSError as e:
            registro.error("No se pudo guardar %s: %s", ARCHIVO_ENTREVISTAS.name, e)
            return {"ok": False,
                    "error": "No se pudo guardar la marca. Revisá que la carpeta del "
                             "programa no sea de sólo lectura."}
        return {"ok": True}

    def abrir_resultados(self):
        abrir_carpeta(CARPETA_SALIDA)
        return {"ok": True}

    def abrir_planilla(self):
        """Abre en el navegador la planilla de Google con las respuestas del
        formulario (el botón verde de la barra de arriba)."""
        sheet_id = (self._ajustes.get("sheet_id") or "").strip()
        if not sheet_id:
            return {"ok": False,
                    "error": "Todavía no está cargado el ID de la planilla de Google "
                             "en Configuración."}
        if not abrir_enlace(URL_PLANILLA.format(sheet_id)):
            return {"ok": False,
                    "error": "No se pudo abrir el navegador. Entrá a la planilla de "
                             "Google desde el navegador como lo hacés siempre."}
        self.log("Se abrió la planilla de respuestas en el navegador.")
        return {"ok": True}

    def elegir_archivo(self, tipo):
        """Diálogo nativo de Windows. Devuelve la ruta elegida o None."""
        if tipo == "csv":
            tipos = ("Archivo CSV (*.csv)", "Todos los archivos (*.*)")
        else:
            # El cliente trabaja con LibreOffice, así que .ods va primero.
            tipos = ("Planillas (*.ods;*.xlsx;*.xlsm)",
                     "LibreOffice Calc (*.ods)",
                     "Excel (*.xlsx;*.xlsm)",
                     "Todos los archivos (*.*)")
        elegido = self._ventana.create_file_dialog(webview.OPEN_DIALOG, file_types=tipos)
        if not elegido:
            return None
        return elegido[0] if isinstance(elegido, (list, tuple)) else str(elegido)

    def ejecutar(self, datos):
        """Arranca el trabajo en segundo plano. Valida antes de largar el hilo."""
        if self._trabajando:
            return {"ok": False, "error": "Ya hay un trabajo en curso, esperá a que termine."}

        origen = datos.get("origen", "sheets")
        if origen == "sheets" and not self._ajustes.get("sheet_id"):
            return {
                "ok": False,
                "abrir_config": True,
                "error": "Primero hay que cargar el ID de la planilla de Google en Configuración.",
            }

        self._trabajando = True
        objetivo = self._trabajo_xlsx if origen == "xlsx" else self._trabajo
        threading.Thread(target=objetivo, args=(datos,), daemon=True).start()
        return {"ok": True}

    # ----------------------------------------------------------------- trabajo
    def _trabajo(self, datos):
        """Flujo normal: Google Sheets (o un CSV descargado) → XLSX + informe."""
        try:
            procesar = procesar_modulo()

            tipo = datos.get("tipo") or "instancia"
            # El tipo real depende de cuántas fechas se marcaron (una instancia
            # dictada en varios días es otro informe): se resuelve igual que
            # adentro de procesar_datos, sólo para que el registro no mienta.
            tipo_real, _f, _fs = procesar.resolver_tipo(
                tipo, datos.get("fecha"), datos.get("fechas"))
            self.log(f"— {procesar.TIPOS_INFORME.get(tipo_real, tipo_real)}: leyendo respuestas…")
            resultados = procesar.procesar_datos(
                csv=datos.get("archivo") if datos.get("origen") == "csv" else None,
                sheet_id=self._ajustes.get("sheet_id") or None,
                worksheet=self._ajustes.get("worksheet") or None,
                credenciales=str(ARCHIVO_CREDENCIALES),
                config_path=str(CONFIG),
                tipo=tipo,
                unidad=datos.get("unidad") or None,
                fecha=datos.get("fecha") or None,
                # Jornadas marcadas de una instancia de fecha múltiple.
                fechas=datos.get("fechas") or None,
                servicios=datos.get("servicios") or None,
                # Cómo se llama el servicio en el informe (lo elige la persona).
                servicio=datos.get("servicio_nombre") or "",
                asistentes=datos.get("asistentes") or 0,
                salida=str(CARPETA_SALIDA),
                log=self.log,
            )
            self.log(f"Planillas Excel generadas: {len(resultados)}.")

            if datos.get("con_informe"):
                self.generar_informes([rj for _, rj in resultados],
                                      self.indicaciones_de(datos.get("antecedentes")),
                                      datos.get("entrevistas_conf") or None)

            self._exito()
        except Exception as e:
            self._fallo(e)

    def _trabajo_xlsx(self, datos):
        """Genera el informe a partir de un XLSX (formato oficial) hecho a mano."""
        try:
            import shutil

            procesar = procesar_modulo()

            ruta_xlsx = datos["archivo"]
            self.log(f"— Leyendo la planilla {os.path.splitext(ruta_xlsx)[1].lower()}…")
            resumen, faltantes = procesar.datos_desde_xlsx(ruta_xlsx, config_path=str(CONFIG))
            if faltantes:
                self.log(f"⚠ {len(faltantes)} pregunta(s) del formulario no se encontraron en la planilla (van en 0).")

            unidad, fecha = resumen["unidad"], resumen["fecha"]
            base = procesar.nombre_archivo_seguro(f"relevamiento_{unidad}_{fecha}")
            sub = CARPETA_SALIDA / procesar.nombre_carpeta_instancia(unidad, fecha)
            sub.mkdir(parents=True, exist_ok=True)

            # La planilla del usuario se copia a la carpeta de la instancia (es el
            # entregable). Se conserva su extensión: puede ser .ods o .xlsx.
            extension = os.path.splitext(ruta_xlsx)[1].lower() or ".xlsx"
            try:
                shutil.copy2(ruta_xlsx, sub / f"{base}{extension}")
            except Exception as e:
                self.log(f"⚠ No se pudo copiar la planilla a la carpeta: {e}")

            ruta_json = sub / f"{base}.json"
            ruta_json.write_text(json.dumps(resumen, ensure_ascii=False, indent=2), encoding="utf-8")
            self.log(f"✔ {unidad}: datos leídos de la planilla → {sub.name}/")

            # Este botón existe para generar el informe, así que siempre se genera.
            self.generar_informes([str(ruta_json)], self.indicaciones_de(datos.get("antecedentes")),
                                  datos.get("entrevistas_conf") or None)

            self._exito()
        except Exception as e:
            self._fallo(e)

    def _exito(self):
        mensaje = "Listo. Los archivos están en la carpeta «InformesIntervenciones» del Escritorio."
        self.log("✅ " + mensaje)
        self._terminar(True, mensaje)
        abrir_carpeta(CARPETA_SALIDA)

    # -------------------------------------------------------------- el informe
    def indicaciones_de(self, antecedentes):
        """Arma el texto de 'indicaciones' para la IA a partir del campo antecedentes."""
        if antecedentes:
            return (
                "Información de antecedentes y contexto aportada por la Unidad "
                "(usala para redactar los Antecedentes y la Metodología, sin inventar nada más): "
                + antecedentes
            )
        return None

    def generar_informes(self, rutas_json, indicaciones, entrevistas_conf):
        """Genera el informe para cada JSON (con API, o el texto para claude.ai si
        no hay clave). Con DOCX OK deja solo los entregables (XLSX + DOCX)."""
        import generar_informe
        modelo = self._ajustes.get("modelo") or generar_informe.MODELO
        if self._ajustes.get("api_key"):
            os.environ["ANTHROPIC_API_KEY"] = self._ajustes["api_key"]
            for ruta_json in rutas_json:
                ruta_md, ruta_docx = generar_informe.generar(
                    ruta_json, modelo=modelo, indicaciones=indicaciones,
                    entrevistas_confidenciales=entrevistas_conf, log=self.log)
                if ruta_docx:
                    self.dejar_solo_entregables(ruta_json, ruta_md)
        else:
            self.log("ℹ No hay clave de IA cargada: se creó un texto para generar el informe GRATIS en claude.ai:")
            for ruta_json in rutas_json:
                generar_informe.exportar_prompt(ruta_json, indicaciones=indicaciones,
                                                entrevistas_confidenciales=entrevistas_conf,
                                                log=self.log)
            self.log("   Cómo usarlo: 1) entrá a https://claude.ai con tu cuenta")
            self.log("   2) abrí el archivo _para_claude.txt, copiá TODO su contenido")
            self.log("   3) pegalo en el chat de Claude y envialo")
            self.log("   4) guardá la respuesta: ese es el informe.")

    def dejar_solo_entregables(self, ruta_json, ruta_md):
        """Borra los intermedios (JSON y MD) para que en la carpeta queden solo
        los entregables: el XLSX y el DOCX. Se llama únicamente cuando el DOCX
        se generó bien."""
        for ruta in (ruta_json, ruta_md):
            try:
                if ruta and os.path.exists(ruta):
                    os.remove(ruta)
            except OSError:
                pass  # si no se puede borrar, no es grave: el informe ya está


def _autotest():
    """Verificación silenciosa para el empaquetado: confirma que los recursos
    embebidos (config, estilo, plantilla, ejemplos, interfaz web) se encuentran
    al correr como .exe. Escribe el resultado en un archivo (la app es de
    ventana, sin consola) y devuelve el código de salida.
    Se invoca con: Relevamiento.exe --autotest"""
    recursos = {
        "config.json": CONFIG,
        "estilo_informe.md": recurso("recursos/estilo_informe.md"),
        "plantilla.docx": recurso("recursos/plantilla.docx"),
        "informes_ejemplo": recurso("recursos/informes_ejemplo"),
        "web/index.html": INTERFAZ,
        "web/estilo.css": recurso("web") / "estilo.css",
        "web/app.js": recurso("web") / "app.js",
    }
    faltan = [n for n, p in recursos.items() if not p.exists()]
    lineas = [f"{'OK ' if p.exists() else 'FALTA'}  {n}  ->  {p}" for n, p in recursos.items()]
    lineas.append(f"Carpeta externa (junto al ejecutable): {dir_externo()}")
    lineas.append("RESULTADO: " + ("todos los recursos OK" if not faltan else "FALTAN recursos"))
    (dir_externo() / "autotest_resultado.txt").write_text("\n".join(lineas), encoding="utf-8")
    return 0 if not faltan else 1


def area_util():
    """Escritorio disponible (sin la barra de tareas) como (x, y, ancho, alto),
    en las mismas unidades lógicas que usa create_window. None si no se puede.

    Se consulta a Windows en vez de usar `webview.screens` porque esa propiedad
    devuelve píxeles físicos o lógicos según si el proceso ya declaró que
    entiende de escalado, cosa que pywebview recién hace dentro de start().
    La cuenta de acá sale bien en los dos casos: sin escalado declarado Windows
    ya reporta todo achicado y el DPI da 96 (escala 1), y con escalado declarado
    reporta píxeles reales y el DPI verdadero."""
    if not sys.platform.startswith("win"):
        return None
    try:
        import ctypes
        from ctypes import wintypes

        u = ctypes.windll.user32
        rect = wintypes.RECT()
        if not u.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0):  # SPI_GETWORKAREA
            return None
        escala = ctypes.windll.gdi32.GetDeviceCaps(u.GetDC(0), 88) / 96    # LOGPIXELSX
        if escala <= 0:
            return None
        return (int(rect.left / escala), int(rect.top / escala),
                int((rect.right - rect.left) / escala),
                int((rect.bottom - rect.top) / escala))
    except Exception:
        return None


def geometria_ventana(ancho=920, alto=860, margen=40):
    """Tamaño y posición de la ventana RESTAURADA, es decir cuando el usuario
    la achica con el botón del medio de la barra de título (arranca maximizada).

    Va recortada para que entre entera en el escritorio: en una pantalla común
    con escalado al 125% una ventana de alto fijo se pasa del área visible y el
    botón EJECUTAR queda tapado por la barra de tareas. Devuelve (ancho, alto, x, y)."""
    area = area_util()
    if not area:
        return ancho, alto, None, None
    ax, ay, ancho_util, alto_util = area
    ancho = max(720, min(ancho, ancho_util - margen))
    alto = max(560, min(alto, alto_util - margen))
    return ancho, alto, ax + (ancho_util - ancho) // 2, ay + (alto_util - alto) // 2


def main():
    if ya_esta_abierto():
        return   # ya hay una ventana abierta: se la trajo al frente y listo

    configurar_registro()
    puente = Puente()
    ancho, alto, x, y = geometria_ventana()
    opciones = {} if x is None else {"x": x, "y": y}
    ventana = webview.create_window(
        TITULO,
        str(INTERFAZ),
        js_api=puente,
        width=ancho,
        height=alto,
        min_size=(720, 560),
        background_color="#eef1f5",
        text_select=False,
        # Arranca ocupando toda la pantalla. Se usa `maximized` y NO `fullscreen`
        # a propósito: fullscreen saca la barra de título, y el cliente se
        # quedaría sin la X para cerrar el programa.
        maximized=True,
        **opciones,
    )
    # Se asigna después de crear la ventana, y con guión bajo: ver el comentario
    # en Puente sobre por qué no puede ser un atributo público.
    puente._ventana = ventana
    # --debug abre las herramientas de desarrollo (sólo para quien programa).
    webview.start(debug="--debug" in sys.argv)


if __name__ == "__main__":
    if "--autotest" in sys.argv:
        sys.exit(_autotest())
    main()
