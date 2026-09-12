# Informes de ejemplo

`generar_informe.py` lee esta carpeta y le pasa lo que encuentre al modelo como
ejemplo de estilo, estructura, tono y extensión (la parte *few-shot* del
prompt). Es lo que hace que el informe generado se parezca a los que la
organización ya venía escribiendo, más allá de lo que describa
`../estilo_informe.md`.

## Los de este repositorio son ficticios

Las unidades, los servicios, las fechas, los porcentajes y los testimonios que
aparecen acá están **inventados**. No corresponden a ninguna organización ni a
ninguna persona real. Se generan con `desarrollo/crear_ejemplos.py` y existen
para que el proyecto funcione de punta a punta sin publicar documentos internos
de nadie.

En una instalación real esta carpeta se reemplaza por informes propios: cuantos
más y más representativos, mejor sale el resultado.

## Formatos

Se leen `.md`, `.txt`, `.docx` y `.odt`; cualquier otra extensión se ignora, y
los archivos llamados `README` o `LEEME` se saltean. Los cuatro ejemplos cubren
los tres lectores a propósito: los `.md` se leen directo, el `.docx` trae una
tabla (que se convierte a Markdown conservando su lugar en el texto) y el `.odt`
trae una nota al pie (que se descarta al leer).
