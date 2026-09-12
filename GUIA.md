# Guía del proyecto — Relevamiento de Instancia Grupal

Automatización para la Unidad de Mediación y Convivencia de un servicio público de salud,
desarrollada junto a Claude. **El cliente final no sabe de computación**: toda
mejora tiene que mantener la simplicidad de uso (doble clic → completar campos →
Ejecutar).

Esta es la **única documentación del proyecto**. Está ordenada del uso diario
hacia adentro: las secciones 1 a 5 alcanzan para usar, instalar y entregar el
programa; de la 6 en adelante está el detalle técnico y el porqué de cada
decisión, que es lo que hay que leer antes de tocar código.

---

# 1. Qué hace (3 etapas)

1. **Recopilación**: Google Forms → Google Sheets (ya existe, no es parte de
   este código).
2. **Procesamiento** (`codigo\procesar.py`): lee las respuestas (Google Sheets
   API o CSV) y genera un XLSX con el formato oficial del relevamiento
   (cantidades + porcentajes con semáforo de colores) y un JSON de resumen.
3. **Informe** (`codigo\generar_informe.py`): con el JSON + los informes reales
   de ejemplo + la guía de estilo, la API de Claude redacta el Informe Técnico,
   que se convierte a Word sobre la hoja membretada institucional.

---

# 2. Uso diario

1. Doble clic en **Iniciar.bat** (o en el ícono del Escritorio, si se entregó el
   `.exe`).
2. La ventana lee sola la planilla y arriba muestra cuántas respuestas encontró.
3. Elegir **qué informe se quiere generar** y completar los datos que pida.
4. Poner la **cantidad de asistentes**. Opcionalmente, escribir en **Antecedentes
   / contexto** quién solicitó la intervención y qué actividades previas hubo: la
   IA lo usa para redactar esa sección del informe.
5. Apretar **▶ EJECUTAR**.
6. Al terminar se abre sola la carpeta con los resultados: la planilla Excel y el
   informe en Word.

Si algo falta o falla, la ventana lo explica en lenguaje simple. El botón **"Usar
archivo descargado…"** permite trabajar con un CSV bajado a mano de Google
Sheets, por si algún día no hay conexión: al elegirlo, las listas se completan
con lo que tenga ese archivo.

## El botón verde: abrir la planilla de respuestas

En la barra celeste de arriba, el botón verde **"Abrir la planilla de
respuestas"** abre el Google Sheets del formulario **en el navegador**, para
mirar las respuestas crudas sin salir a buscarlo a mano. Usa el ID de la planilla
que está cargado en Configuración, así que siempre apunta a la que se está
usando; si todavía no hay ID cargado, el botón no aparece.

## Los tipos de informe

La ventana muestra **2 opciones**:

| Tipo | Qué junta | Qué hay que elegir |
|---|---|---|
| **Instancia** | Un taller de una unidad y su servicio, se haya dictado en uno o en varios días | Unidad Ejecutora, la(s) fecha(s) y Servicio/Área |
| **Por Unidad Ejecutora** | Toda la unidad: todos sus servicios y todas sus fechas | Unidad Ejecutora |

**Sobre las fechas de la Instancia**: se elige primero la Unidad Ejecutora y
abajo aparecen sus fechas con un tilde cada una, de la más reciente a la más
antigua. Lo normal es marcar **una**: sale el informe del taller de ese día,
igual que siempre. Si la misma instancia se dictó en **más de un día** (por
ejemplo, el mismo taller repetido para el turno de la mañana y el de la tarde),
se marcan todas esas fechas: las respuestas se suman y sale **un solo informe**,
redactado como una única instancia con sus jornadas enumeradas.

No hay que elegir de antemano entre "un día" y "varios días": el programa se da
cuenta por la cantidad de fechas marcadas. Marcar una fecha más no sólo suma
respuestas — también cambia cómo se cuentan las entrevistas individuales (con un
solo día se informa el total de ese día; con varios, sólo las de esa unidad y
servicio).

**Sobre Servicio/Área**: en el formulario es una pregunta de texto libre, así que
la misma área aparece escrita de varias maneras («Emergencia», «Emergencia
adultos», «Emergencia de adulto»…). La ventana muestra la lista de lo que hay
cargado con la cantidad de respuestas de cada uno: se marcan todas las que
correspondan al mismo servicio y se juntan en un solo informe. Si no se marca
nada, entran todas.

**Las listas se van acotando entre sí.** Cada una ofrece únicamente lo que existe
para lo que ya elegiste, así no se puede pedir una combinación sin respuestas. El
orden es **Unidad Ejecutora → fechas → Servicio/Área**:

- Al elegir la **unidad**, la lista de fechas queda con los días en que esa
  unidad respondió.
- Al marcar las **fechas**, la lista de Servicio/Área queda con los servicios de
  esas jornadas. Si el 28/7 no hubo Laboratorio, no aparece.

El número que aparece a la derecha de cada fecha y de cada servicio es la
cantidad de respuestas **dentro de lo que elegiste**, no del total de la planilla.

**El campo "Nombre para el informe"** define cómo se va a llamar ese servicio en
el Excel y en el título del informe. Se completa solo con el nombre más corto de
los que marcaste, y se puede corregir. Por ejemplo, si marcás `Laboratorio` y
`Laboratorio/ pediatría`, sugiere **Laboratorio**, y el informe queda titulado
*HOSPITAL CENTRAL – LABORATORIO*. Sin ese campo diría *LABORATORIO,
LABORATORIO/ PEDIATRÍA*, que en un documento firmado parece un error de carga.
Las variantes originales quedan igual registradas en la planilla, así que no se
pierde el detalle de qué se incluyó.

> Lo que pasa por dentro cuando se marcan una o varias fechas está en
> «Cómo se decide el tipo», más abajo.

Los informes que juntan varias fechas suman las respuestas, no las separan por
fecha. Por eso el informe no dice que algo mejoró o empeoró en el tiempo: el de
varias jornadas habla de la instancia y sus días, y el de unidad, de una
tendencia general del período.

## Entrevistas pendientes

El botón **"Ver entrevistas pendientes"**, debajo de EJECUTAR, abre la lista de
las personas que pidieron una entrevista individual. Muestra una tabla con la
**U.E.**, la **fecha**, el **nombre** y el **contacto**, ordenada **de la más
antigua a la más reciente**: arriba queda el pedido que lleva más tiempo
esperando, que es el primero que conviene atender.

> **Qué se considera un pedido de entrevista.** Tienen que darse las dos cosas:
> haber respondido **que sí** a la pregunta *"¿Considera usted necesario solicitar
> una entrevista individual confidencial por algún tema de su especial
> interés?"* **y** haber dejado nombre y celular. Es el mismo criterio con el que
> se cuentan las entrevistas en el informe, así que la lista y el número del
> informe siempre coinciden.
>
> Se pide la confirmación explícita porque antes alcanzaba con dejar los datos de
> contacto, y eso contaba como pedido a alguien que había completado esos campos
> sin querer una entrevista. Al revés también se descarta: quien dice que sí pero
> no deja cómo contactarlo no entra, porque no se lo podría llamar.

Al terminar una entrevista se aprieta **"Completada"** en esa fila: desaparece de
la lista y no vuelve a aparecer, ni siquiera al cerrar y abrir el programa. La
lista se lee de la planilla cada vez que se abre la ventana, así que los pedidos
nuevos aparecen solos.

La ventana tiene dos pestañas: **Pendientes** y **Completadas**. En Completadas
están todas las que se fueron marcando, con la más reciente arriba, y cada una
tiene un botón **"Deshacer"** que la devuelve a pendientes. Así, si se marcó una
por equivocación, se arregla con un clic.

> Las marcas se guardan en el archivo `entrevistas_completadas.json`, al lado del
> programa (en la planilla de Google no se puede escribir: el programa entra como
> lector). El archivo guarda un código, no los nombres ni los teléfonos.

## Informe desde una planilla hecha a mano

El botón **"Informe desde planilla…"** genera el informe a partir de una planilla
en el formato oficial llenada a mano (sin pasar por Google Sheets): lee los datos
de la planilla y arma el informe con esos números, usando igual los campos de
antecedentes y entrevistas confidenciales.

> **Sirve tanto para LibreOffice como para Excel.** Acepta `.ods` (LibreOffice
> Calc) y `.xlsx` / `.xlsm` (Excel), y no hace falta convertir nada: se elige el
> archivo tal como está. Ojo con la diferencia: `.ods` es la **planilla** de
> LibreOffice (Calc), mientras que `.odt` es un **documento de texto** (Writer) y
> no sirve para esto. Si se elige un `.odt` por error, el programa lo explica en
> vez de dar un error técnico.

Por consola: `python codigo\generar_informe.py --desde-xlsx archivo.ods`.

---

# 3. Instalación y configuración

Requiere **Windows** y Python 3.9 o superior:

1. Instalá Python desde https://www.python.org/downloads/ marcando la casilla
   **"Add Python to PATH"**.
2. Copiá esta carpeta completa a la PC (por ejemplo a Documentos).
3. Doble clic en **instalar.bat** (instala los componentes; requiere internet).
4. Abrí **Iniciar.bat**: la primera vez se abre sola la ventana de
   **Configuración**, donde se carga el ID de la planilla de Google y la clave de
   la IA (quedan guardados en `ajustes.json` y no hay que volver a tocarlos). El
   archivo `credenciales.json` de Google se copia a esta misma carpeta.

Dependencias: gspread, google-auth, openpyxl, anthropic, python-docx, pywebview
(están en `requirements.txt`).

> La ventana se dibuja con HTML/CSS/JS adentro de una ventana nativa (pywebview).
> En Windows usa **WebView2**, que ya viene instalado con Windows 10 y 11: no hay
> que instalar nada extra, ni se abre un navegador, ni se levanta ningún servidor.

> **Sólo Windows.** La aplicación de ventana está probada únicamente en Windows.
> En Mac o Linux pywebview necesita otro motor (WebKit o Qt) que nunca se probó
> acá, así que no hay lanzador para esos sistemas. Lo que sí funciona en cualquier
> sistema es el uso por consola (`procesar.py` y `generar_informe.py`), que no
> depende de la ventana.

## Configurar la API de Google (una sola vez)

Se usa una **cuenta de servicio**, que permite leer la planilla sin intervención
manual:

1. Entrá a https://console.cloud.google.com y creá un proyecto (o usá uno
   existente).
2. En "APIs y servicios" → "Biblioteca", habilitá **Google Sheets API** y
   **Google Drive API**.
3. En "APIs y servicios" → "Credenciales" → "Crear credenciales" → **Cuenta de
   servicio**. Completá un nombre y creala.
4. Abrí la cuenta de servicio creada → pestaña "Claves" → "Agregar clave" →
   "Crear clave nueva" → tipo **JSON**. Se descarga un archivo: guardalo en esta
   carpeta con el nombre `credenciales.json`.
5. Copiá el correo de la cuenta de servicio (termina en
   `...iam.gserviceaccount.com`) y **compartí tu Google Sheets con ese correo**
   como Lector, igual que compartirías con una persona.

El ID de la planilla es la parte larga de la URL:
`https://docs.google.com/spreadsheets/d/`**`ESTE_ES_EL_ID`**`/edit`.

## Configurar la API de Claude (una sola vez)

1. Creá una clave en https://console.anthropic.com (sección API Keys). El uso se
   paga por consumo: generar un informe típico cuesta unos pocos centavos de
   dólar.
2. Cargala en el botón **⚙ Configuración** de la aplicación (queda guardada en
   `ajustes.json`). Para uso por consola también sirve la variable de entorno
   `ANTHROPIC_API_KEY`.
3. En esa misma ventana, el campo **Modelo de IA** permite cambiar el modelo sin
   tocar código. Dejalo vacío para usar el recomendado (`claude-opus-5`);
   modelos y precios vigentes en
   https://docs.claude.com/en/docs/about-claude/models.

## Probar sin pagar la API (modo sin clave)

Mientras no haya una clave de IA cargada en Configuración, la aplicación no
falla: al ejecutar, además del Excel genera un archivo `..._para_claude.txt` con
todo lo necesario (instrucciones + informes de ejemplo + datos procesados). Para
obtener el informe gratis:

1. Entrá a https://claude.ai con una cuenta (la gratuita sirve).
2. Copiá todo el contenido del archivo `_para_claude.txt` y pegalo en el chat.
3. La respuesta de Claude es el informe: guardala en Word.

Es el mismo modelo y las mismas instrucciones que usaría la API, así que el
resultado es representativo de lo que produce el modo automático. Cuando se
contrate la API, se carga la clave en Configuración y el paso manual desaparece.
Por consola: `python codigo\generar_informe.py salida\archivo.json --sin-clave`.

---

# 4. Entrega al cliente (el `.exe`)

Si se prefiere que el cliente **no** tenga que instalar Python ni correr
`instalar.bat`, se empaqueta todo con PyInstaller.

1. Instalar PyInstaller una vez en la PC de desarrollo: `pip install pyinstaller`
   (el build de referencia se hizo con **PyInstaller 6.22.2** sobre Python
   3.14.7). **`construir_exe.bat` no lo instala solo**: si falta, el build falla
   con "No module named PyInstaller".
2. Doble clic en **desarrollo\construir_exe.bat**. Demora unos minutos la primera
   vez. El resultado queda en **`dist\Relevamiento\`** (modo *onedir*, ~91 MB).
3. Doble clic en **desarrollo\preparar_entrega.bat**: arma
   **`Relevamiento_para_entregar.zip`** (~52 MB), que es lo que se le manda al
   cliente.
4. En la PC del cliente: descomprimir el ZIP en un lugar definitivo, copiar
   `credenciales.json` adentro de la carpeta y doble clic en
   `crear_acceso_directo.bat`. Esos tres pasos están escritos en el `LEEME.txt`
   que viene dentro del ZIP.
5. A partir de ahí el cliente abre el programa desde el ícono del Escritorio y no
   necesita entrar nunca a la carpeta. La primera vez se abre la ventana de
   **Configuración** (ID de la planilla y clave de la IA). La app crea sola
   `ajustes.json` junto al `.exe` y guarda los resultados en la carpeta
   **`InformesIntervenciones`** del **Escritorio**.

> **No zipear `dist\Relevamiento` a mano.** Si se probó el programa desde esa
> carpeta, adentro quedaron el `credenciales.json` (la cuenta de servicio de
> Google) y el `ajustes.json` (con la clave de la API de Claude).
> `desarrollo\preparar_entrega.bat` existe justamente para eso: los excluye del
> ZIP y lo recuerda por pantalla.

> **`credenciales.json` va aparte.** No viaja en el ZIP: se le hace llegar al
> cliente por un canal donde no quede dando vueltas (en persona, en un pendrive, o
> por un medio que después se pueda borrar). Lo mismo con la clave de la API:
> mejor cargarla uno mismo en la ventana de Configuración que mandarla por escrito.

> **El ZIP pesa ~52 MB**, así que no entra en un correo común (el límite habitual
> es 25 MB). Pendrive o servicio de transferencia de archivos.

Qué va **adentro** del programa (no hay que copiarlo aparte): la carpeta
`recursos\` (config.json, estilo_informe.md, plantilla.docx, informes_ejemplo/) y
`web\` (la interfaz). Si más adelante se cambia alguno de esos archivos, hay que
volver a correr `desarrollo\construir_exe.bat`.

> **Por qué una carpeta y no un solo archivo**: se puede empaquetar todo en un
> `.exe` suelto, pero ese modo se descomprime entero en cada arranque y la ventana
> tardaba unos 3 segundos en aparecer, sin ninguna señal de que estaba cargando —
> un usuario no técnico piensa que no funcionó y vuelve a hacer doble clic. En
> modo carpeta abre en menos de un segundo. Como el cliente usa el acceso directo,
> la carpeta ni la ve.

> **Si hace doble clic dos veces**: no pasa nada. El programa detecta que ya está
> abierto y trae al frente la ventana que ya existe, en vez de abrir una segunda
> copia.

> **Si algo falla, pedirle el registro.** El programa deja un archivo
> **`registro.log`** dentro de su propia carpeta, al lado de `Relevamiento.exe`.
> Ahí queda todo lo que apareció en el recuadro de actividad más el detalle
> técnico de los errores. Cuando el cliente diga que algo no anduvo, hay que
> pedirle ese archivo: es lo único que permite ver qué pasó sin estar frente a la
> máquina. Se limita solo a unos pocos MB, así que no hay que borrarlo.

> **Ojo al reconstruir**: `dist\Relevamiento` es una carpeta *generada*, y
> PyInstaller la borra entera cada vez que se construye de nuevo.
> `desarrollo\construir_exe.bat` guarda antes y repone después
> `credenciales.json`, `ajustes.json` y `entrevistas_completadas.json` — pero si
> se corre PyInstaller a mano (`python -m PyInstaller --noconfirm --clean
> desarrollo\relevamiento.spec`) **no hay red de contención** y se pierden (ya
> pasó una vez). Correrlo a mano es lo que hay que hacer para automatizar, porque
> el `.bat` termina en `pause`; en ese caso hay que copiar esos tres archivos
> antes y reponerlos después, más `desarrollo\crear_acceso_directo.bat`.

> Notas: algunos antivirus marcan falsos positivos con ejecutables de PyInstaller;
> conviene probarlo en una PC parecida a la del cliente antes de entregarlo.
> `dist\` y `build\` se pueden borrar en cualquier momento: se regeneran.

## El ícono

El ícono del programa está en `recursos\icono.ico` y se genera con
`python desarrollo\crear_icono.py` (son tres barras con los colores del semáforo
del relevamiento sobre el azul institucional). Para cambiarlo, se edita ese
script y se vuelve a construir; necesita `pip install pillow`, que se usa sólo
para eso y no es una dependencia del programa. También se puede reemplazar
`recursos\icono.ico` directamente por otro `.ico` con los tamaños 16, 32, 48 y
256.

---

# 5. Estructura del proyecto

Reorganizada en carpetas el 2026-08-13 (antes era todo plano en la raíz). El
criterio: la raíz sólo tiene lo que el cliente toca o lo que tiene que quedar
junto al `.exe`; el resto se agrupa por rol.

```
relevamiento/
├── Iniciar.bat            Doble clic del cliente → lanza `python codigo\app.py`
├── instalar.bat           Instala las dependencias (una vez)
├── credenciales.json      (lo pone el usuario) cuenta de servicio de Google — SECRETO
├── ajustes.json           (lo crea la GUI) sheet_id, worksheet, api_key, servicio, modelo
├── requirements.txt / GUIA.md / CLAUDE.md
│
├── codigo/                Toda la lógica (los 4 módulos se importan planos entre sí)
│   ├── app.py             Ventana pywebview + puente con el JS
│   ├── procesar.py        Etapa 2 (función reutilizable procesar_datos + CLI)
│   ├── generar_informe.py Etapa 3 (generar con API / exportar_prompt sin clave + CLI)
│   └── rutas.py           recurso() / archivo_externo() / carpeta_informes()
│
├── web/                   La interfaz: index.html + estilo.css + app.js
│
├── recursos/              Embebidos en el .exe, sólo lectura
│   ├── config.json        Secciones y textos de las preguntas del formulario
│   ├── estilo_informe.md  Guía de estilo v2 (se anexa al prompt en <guia_de_estilo>)
│   ├── plantilla.docx     Hoja membretada institucional (base de todos los informes Word)
│   ├── icono.ico          Ícono del programa (lo genera crear_icono.py)
│   └── informes_ejemplo/  16 informes reales de estilo: 12 .odt + 4 .docx
│
├── desarrollo/            Nada de esto le llega al cliente
│   ├── construir_exe.bat / relevamiento.spec   Empaquetado
│   ├── preparar_entrega.bat/.py  Arma el ZIP, sin los archivos con claves
│   ├── crear_icono.py     Regenera el ícono (necesita Pillow)
│   ├── crear_acceso_directo.bat  construir_exe.bat lo copia junto al .exe
│   └── ejemplo_respuestas.csv    Datos ficticios para probar sin conexión
│
└── dist/Relevamiento/     El programa empaquetado (lo genera construir_exe.bat)
    ├── Relevamiento.exe + _internal/
    ├── crear_acceso_directo.bat
    └── credenciales.json / ajustes.json / entrevistas_completadas.json / registro.log
                            Los crea o los copia el usuario; NO van en el ZIP
```

Todo lo demás que aparezca en la carpeta es **generado** y se puede borrar sin
miedo: `build\` (caché de PyInstaller, la rehace el próximo build), `salida\`
(resultados del modo consola, `--salida` por defecto), `registro.log`,
`autotest_resultado.txt`, los `__pycache__` y `Relevamiento_para_entregar.zip`
(lo rehace `desarrollo\preparar_entrega.bat`). El único de esa lista que conviene
NO borrar a la ligera es `entrevistas_completadas.json`: se regenera vacío, y con
eso se pierde el registro de qué entrevistas ya se hicieron.

**`web/` queda en la raíz y NO adentro de `recursos/`**, aunque para PyInstaller
sean lo mismo (los dos viajan embebidos de sólo lectura): son 1.600 líneas de
código fuente y la convención del proyecto es que los cambios visuales van ahí.

**`credenciales.json` y `ajustes.json` NO se movieron a una subcarpeta** a
propósito: `archivo_externo()` los busca junto al `.exe` y el cliente tiene que
poner `credenciales.json` ahí a mano. Una carpeta `privado/` no agregaría
seguridad —`EXCLUIR` de `preparar_entrega.py` filtra por NOMBRE de archivo, así
que funciona estén donde estén— y sí le agregaría un paso a quien no es técnico.
Que estén en la raíz **y** en `dist\Relevamiento\` no es un duplicado por
descuido: una copia la usa el programa corriendo como script y la otra el `.exe`.

Las tres piezas que sostienen el layout, por si hay que moverlo de nuevo:

1. **`rutas._raiz_proyecto()`** es `parent.parent` porque `rutas.py` vive en
   `codigo/`. Si se mueve de carpeta, ESE es el único lugar a corregir.
2. **`recurso(nombre)` es relativo a la RAÍZ**, no a `recursos/`: se llama
   `recurso("recursos/config.json")` y `recurso("web")`. El destino de `datas`
   en el `.spec` reproduce esa misma estructura adentro del `.exe`, así que la
   ruta es idéntica empaquetado y como script.
3. **El `.spec` arma todas sus rutas absolutas desde `SPECPATH`**, porque
   PyInstaller resuelve el script relativo al `.spec` pero los orígenes de
   `datas` relativos al directorio actual. Mezclar los dos criterios hacía
   fallar el build buscando `desarrollo/codigo/app.py`.

---

# 6. Cómo funciona por dentro

## Cómo se decide el tipo: 2 opciones en la ventana, 3 tipos internos

La ventana muestra **2 opciones** (Instancia · Por Unidad Ejecutora).
Internamente hay **3 tipos** (`TIPOS_INFORME` en `procesar.py`, parámetro `tipo`
de `procesar_datos`), porque "Instancia" se desdobla según cuántas fechas se
marcaron:

| tipo | filtra por | resultado |
|---|---|---|
| `instancia` | 1 fecha + U.E. + servicio(s) | el informe de siempre |
| `multifecha` | 2+ fechas marcadas + U.E. + servicio(s) | un solo informe: es UNA instancia dictada en varias jornadas |
| `unidad` | U.E. entera (todos los servicios y fechas) | un solo informe |

**`resolver_tipo(tipo, fecha, fechas)` es quien decide** entre los dos primeros, y
`procesar_datos()` lo llama antes que nada. La ventana manda siempre
`tipo="instancia"` más la lista de fechas marcadas; con una sola queda
`instancia` (y `fecha` = esa), con dos o más pasa a `multifecha`. Se fusionaron
las dos tarjetas el 2026-08-05, a pedido del usuario: quien usa el programa no
tiene por qué clasificar de antemano si su taller fue de un día o de varios. **Lo
que NO se fusionó es el comportamiento**, porque tres cosas dependen de verdad
del tipo y no son cosméticas:

1. **Las entrevistas individuales** (`contar_entrevistas()`): `instancia` informa
   el total del DÍA sobre todas las unidades (convención histórica de la
   Unidad), `multifecha` cuenta dentro del conjunto filtrado. Medido con
   `ejemplo_respuestas.csv`: Hospital Central el 18/6/2026 da **2** como instancia
   y **1** como fecha múltiple, porque la otra la pidió el Hospital del Norte ese
   mismo día. Ese número entra al informe firmado. (Ese archivo es anterior a la
   pregunta de confirmación de 2026-08-13, así que ejercita el camino de
   compatibilidad de `detector_de_entrevistas()`; con la columna cargada el mismo
   caso da **1**.)
2. **La carpeta de salida** (`nombre_carpeta_salida()`).
3. **El bloque `<alcance_del_informe>`** que va a la IA: con `multifecha` afirma
   que la instancia se dictó en más de una jornada, lo que con una sola fecha
   sería falso.

Verificado que con UNA fecha marcada el JSON sale **idéntico byte a byte** al del
camino viejo (`--fecha` en vez de `--fechas`), misma carpeta y mismo nombre de
archivo.

`multifecha` se llamaba `historico` y juntaba TODAS las fechas de la U.E. y el
servicio, sin poder elegirlas. `ALIAS_TIPOS` y `normalizar_tipo()` siguen
aceptando `historico` para no romper la línea de comandos; `historico` SIN
`--fechas` sigue juntando todas (por eso `resolver_tipo()` sólo degrada a
`instancia` cuando hay exactamente una elegida). Se redacta como informe de
instancia y no como seguimiento en el tiempo: ver la regla de `multifecha` en
`estilo_informe.md` y la rama propia de `_bloque_alcance()`.

**No existe la opción "(todas) las U.E."** Existió hasta el 2026-08-05 en el
informe puntual y generaba un informe por cada unidad de esa fecha (era el único
tipo que devolvía más de un resultado). El usuario indicó que no se usa. El
fan-out sigue vivo en `procesar_datos` (tipo `instancia` sin `unidad`) y por
consola, pero la ventana ya no lo ofrece.

**Hubo un cuarto tipo, `region`** (todas las U.E. de una región en un solo
informe), sacado el 2026-08-05 por pedido del usuario. Se borró entero: el tipo,
el filtro, el parámetro `region=` de `procesar_datos`/`filtrar_registros`, el
`--region` de la consola, la tarjeta y el `<select>` de la ventana, y la regla de
la guía de estilo. **Lo que NO se tocó es la lectura de la región**
(`region_de_registro()` → `_region`): sigue viajando como dato en el JSON de
cualquier informe, la muestra `--listar` y forma parte de la clave de
`combinaciones`. Si algún día se repone el informe regional, los datos ya están.

Los tipos que juntan varias fechas suman las respuestas: el JSON **no** las abre
por fecha ni por unidad, y por eso la guía de estilo y el bloque
`<alcance_del_informe>` le prohíben a la IA hablar de evolución en el tiempo o
atribuirle un porcentaje a una unidad o a una jornada concreta.

## procesar.py

- Google Sheets vía cuenta de servicio (gspread) o `--csv`; comparaciones de
  texto con `normalizar()` (minúsculas, sin tildes).
- **Hallazgos sobre la planilla real (2026-08-03), que motivaron varios cambios:**
  - La columna **`Fecha` del formulario está vacía en todas las filas**; la fecha
    real está en `Marca temporal` (`28/7/2026 17:46:26`). `fecha_de_registro()`
    usa `Fecha` y cae en la marca temporal cuando viene vacía; `solo_fecha()` le
    saca la hora. Sin esto el filtro por fecha no encontraba nada.
  - **`Servicio/Área` sí existe como columna** (índice 31) y no estaba en
    `config.json`: por eso antes se tipeaba a mano. Es texto libre y está
    desprolijo (11 variantes para ~3 servicios: "Emergencia", "Emergencia
    adultos", "Emergencia de adulto"…). Por eso el filtro es una **lista de
    casillas para marcar varias**, no un campo de texto.
  - Hay una columna `U.E (UNIFICADO)` vacía que `detectar_columnas_unidad()`
    igual detecta (empieza con "ue"). Es inofensiva porque `unidad_de_registro()`
    toma la primera con valor, pero si algún día se llena, va a mandar.
- `anotar_registros()` agrega a cada fila `_unidad`, `_fecha`, `_region` y
  `_servicio`; lo comparten `procesar_datos` y `opciones_disponibles`.
  `region_de_registro()` usa la columna `Región` y, si está vacía, deduce la
  región del nombre de la columna de U.E. que vino completa ("U.E. Región Norte"
  → Región Norte).
- `opciones_disponibles()` devuelve qué hay cargado (fechas, unidades, servicios,
  y también regiones, que hoy sólo usa `--listar`) para poblar las listas de la
  ventana. **Los servicios se agrupan por su forma normalizada**, que es el mismo
  criterio con el que después se filtra, así el conteo que se muestra es el que
  se va a obtener; cada grupo trae `etiqueta` (la forma más usada), `respuestas` y
  `variantes`.
- `nombre_carpeta_salida()` decide la subcarpeta: la instancia puntual conserva
  `{U.E.}-{mm}-{aa}` y los demás llevan sufijo (`-VariasFechas`, `-Completo`)
  para no pisarse. El de fecha múltiple mete además el mes de las jornadas
  (`{U.E.}-{Servicio}-{03-26a06-26}-VariasFechas`): como ahora las fechas se
  eligen, puede haber dos instancias distintas de la misma U.E. y servicio y sin
  eso la segunda pisaba a la primera.
- `notas_de_alcance()` escribe al pie del XLSX qué fechas, unidades y servicios
  se juntaron. **Van al pie y no como filas del encabezado a propósito**: el
  bloque de arriba tiene posiciones fijas y los porcentajes se calculan sobre
  `$B$8`.
- `_mensaje_error_google()` traduce los errores de gspread a lenguaje simple. Sin
  eso, un ID mal copiado le mostraba al cliente literalmente `<Response [404]>`.
- Columnas de unidad: `detectar_columnas_unidad()` toma toda columna cuyo
  encabezado empiece con "U.E" o "RAP" (los nombres reales son tipo "U.E. Región
  Sur", distintos a lo especificado originalmente). `unidad_de_registro()` saltea
  el valor intermedio "RAP" (select dependiente del formulario) y toma el nombre
  concreto de la RAP.
- Fechas con tolerancia de formato: `misma_fecha()` (18/7/2026 == 18/07/26).
- Si los filtros no coinciden, el error lista las fechas y unidades que sí
  existen (RuntimeError con mensaje en lenguaje simple: la GUI lo muestra tal
  cual).
- XLSX (openpyxl): bloque de cantidades en columnas A–F y de porcentajes en H–M;
  los % son fórmulas sobre la celda "Respuestas" `$B$8`; colores
  rojo/amarillo/verde FF7C5C, FFFF66, 92D050; Total por pregunta = SUM de las 3
  opciones; nota al pie sobre respuestas omitidas.
- El JSON agrega pct_* redondeados por pregunta (insumo de la etapa 3).
- "Asistentes" es un dato manual (no sale del formulario). En los informes
  agregados se espera el total de todas las instancias.

### Qué cuenta como entrevista solicitada

`detector_de_entrevistas()` (2026-08-13): hay que haber respondido **que SÍ** a la
pregunta de confirmación del formulario («¿Considera usted necesario solicitar una
entrevista individual confidencial…?», columna `solicita_entrevista` de
`config.json`) **y** haber dejado `Nombre completo` y `Celular de contacto`. Antes
alcanzaba con los datos de contacto. Es el criterio más conservador de los
posibles, elegido por el usuario: ninguna entrevista contada en un informe firmado
queda incontactable, y el número no se infla con quien dejó sus datos sin querer
una entrevista.

- **Está factorizado en UNA sola función a propósito.** De ahí salen dos cosas
  que no pueden contradecirse: el número que va al informe firmado
  (`contar_entrevistas()`) y la lista del botón «Ver entrevistas pendientes»
  (`entrevistas_solicitadas()`). Antes el criterio estaba escrito dos veces.
- `es_afirmativa()` compara normalizado (sin tildes ni mayúsculas), así que toma
  `Sí`, `SI`, `si` y también una opción larga tipo «Sí, considero necesario…». El
  prefijo se acepta **sólo si lo sigue un separador**, para que `siempre` no
  cuente como `sí`. Cubierto con 19 casos borde.
- **Compatibilidad hacia atrás**: si la columna de confirmación NO existe en la
  planilla, avisa por el log y vuelve al criterio viejo (nombre + celular). Eso
  mantiene andando los CSV anteriores, `ejemplo_respuestas.csv` incluido.
- **La celda vacía NO cuenta**, y eso es deliberado. Al agregar la pregunta al
  formulario, las respuestas ya cargadas quedaban con esa celda vacía y habrían
  dejado de contarse de golpe (incluidos pedidos de entrevista todavía
  pendientes). **El usuario lo resolvió en la planilla, completando a mano las
  filas históricas con "Si"/"No"** (2026-08-13), así que no hizo falta ninguna
  regla de transición en el código y el criterio quedó parejo para toda la
  planilla. Si alguna vez reaparecen celdas vacías, van a ser respuestas nuevas
  que saltearon la pregunta: conviene que sea obligatoria en el Google Form,
  porque si no, quien la saltee no se cuenta y no hay ningún aviso.

### La lista de entrevistas solicitadas

`entrevistas_solicitadas()`: las filas que cumplen el criterio de arriba,
devueltas como `{clave, unidad, fecha, nombre, contacto, orden}` **de la más
antigua a la más reciente** (pedido del usuario el 2026-08-13): es una cola de
trabajo, y lo que más esperó es lo primero que hay que atender. Dentro del mismo
día se ordena por nombre. Alimenta el botón «Ver entrevistas pendientes» y
`--entrevistas` por consola.

- **`orden` es el índice cronológico** que calcula Python y viaja hasta el JS.
  Existe porque al apretar «Deshacer» la fila vuelve de Completadas a Pendientes:
  antes se hacía `unshift()` y aterrizaba arriba de todo, lo que con este orden la
  mostraría como la más antigua. Ahora el JS la inserta en su lugar (`splice`
  sobre el primer `orden` mayor). **Se manda el índice y no la fecha porque quien
  ordena es Python**: el JS no parsea fechas (misma convención que la lista de
  fechas de la ventana).
- Las dos pestañas se ordenan por criterios distintos y eso es a propósito:
  **Pendientes** por fecha del pedido (más antigua arriba) y **Completadas** por
  cuándo se marcaron (última arriba, `_orden_marca()`), que es el orden en que se
  busca la que se marcó por error. La fecha es `_fecha`, o sea la marca temporal
  mientras la columna `Fecha` siga viniendo vacía.
- **`clave_entrevista()`**: hash corto de `unidad|fecha|nombre|contacto`
  normalizados. La planilla no tiene número de fila estable (se agregan respuestas
  todo el tiempo y la hoja se puede reordenar), así que la fila se identifica por
  su contenido. **Se guarda el hash y no los datos** para que el archivo de
  completadas no quede con nombres y teléfonos adentro. Contra: dos filas
  idénticas (misma persona, misma unidad, mismo día) colapsan en una sola clave;
  es una respuesta duplicada y tratarla como una sola es lo razonable.
- **Dos criterios distintos según el tipo** (`contar_entrevistas()`) — qué filas
  son elegibles lo decide `detector_de_entrevistas()`; esto es sobre qué CONJUNTO
  se cuenta: en `instancia` es el total del DÍA sobre todas las unidades, como se
  viene informando históricamente; en los otros —`multifecha` incluido— se cuenta
  dentro del propio conjunto filtrado, porque el criterio del día mezclaría
  instancias de otras unidades o regiones que cayeron el mismo día. Por eso marcar
  una fecha más en la ventana no sólo agrega respuestas: también cambia cómo se
  cuentan las entrevistas. Es la razón principal por la que `resolver_tipo()`
  distingue 1 fecha de 2+ en vez de tratar todo como `multifecha`.
- `contar_entrevistas_dia()` cuenta sobre TODOS los registros filtrando solo por
  la fecha del reporte (no por unidad), así que el mismo valor va en todas las
  unidades de esa fecha. Aparece como `entrevistas_individuales` en el JSON y como
  una fila en el XLSX **después** de Médicos/No Médicos (a propósito, para no
  correr la celda `$B$8` "Respuestas" sobre la que se calculan los %). Si faltan
  las columnas, avisa y usa 0. `estilo_informe.md` indica a la IA que lo mencione
  en la Metodología cuando es > 0.

### Planillas de LibreOffice y planillas hechas a mano

- **Lee planillas de LibreOffice (.ods) además de Excel (.xlsx)**: el cliente usa
  LibreOffice, así que `filas_de_planilla()` despacha por extensión a
  `_filas_xlsx` (openpyxl) o `_filas_ods`. El lector de .ods es propio y **no
  agrega ninguna dependencia**: un .ods es un ZIP con `content.xml`, y se parsea
  con `zipfile` + `xml.etree` de la biblioteca estándar. Contempla
  `number-columns-repeated` y saltea las filas vacías (Calc declara ~1.048.000
  filas vacías al final con `number-rows-repeated`; el parseo busca por contenido,
  no por número de fila, así que saltearlas es seguro). Verificado que un .ods y
  su .xlsx equivalente dan un resumen idéntico. Un `.odt` (documento de Writer) da
  un error que explica la diferencia entre planilla y documento de texto.
- **Informe desde una planilla hecha a mano**: `datos_desde_xlsx(ruta,
  config_path)` —el nombre quedó por compatibilidad, acepta .ods y .xlsx— lee una
  planilla en el formato oficial y reconstruye el MISMO dict/JSON que produce
  `procesar_datos` (unidad, fecha, respuestas, médicos, entrevistas, secciones con
  conteos y pct_*), para poder generar el informe sin pasar por Sheets/CSV. Parsea
  el bloque de cantidades (col A–F): fila de encabezado = etiqueta de texto en A +
  valor en B; fila de pregunta = número en A, texto en B, conteos en C/D/E (un 0 se
  guarda como celda vacía → se toma como 0). Empareja las preguntas contra
  `config.json` (normalizado, con match parcial de respaldo). Si no reconoce
  ninguna pregunta, lanza RuntimeError pidiendo el formato oficial. Round-trip
  verificado (genera→lee→idéntico).
- **Agrupación en subcarpetas**: cada instancia (unidad + fecha) escribe sus
  archivos (xlsx, json, `_para_claude.txt`, `_informe.*`) dentro de una subcarpeta
  propia de la carpeta de salida, con nombre `{U.E.}-{mm}-{aa}`
  (`nombre_carpeta_instancia()`): unidad sin tildes ni espacios y conservando
  mayúsculas (p. ej. `HospitalDelNorte`), fecha mes-año a 2
  dígitos (`fecha_mm_aa()`, 20/7/26 → `07-26`). Si la fecha no es reconocible, usa
  un fallback sanitizado sin romper.

## generar_informe.py

- `MODELO = "claude-opus-5"` (se sobreescribe con `ajustes.json["modelo"]` o
  `--modelo`). Cuesta lo mismo que el Opus 4.8 que se usaba antes: US$ 5 por
  millón de tokens de entrada y US$ 25 de salida.
- **No se manda el parámetro `thinking`, y es a propósito.** Opus 5 razona solo
  antes de escribir (viene activado de fábrica); Opus 4.8 y los anteriores no, y
  la forma de configurarlo cambia de un modelo a otro —lo que sirve para uno da
  error 400 en el otro—. No mandarlo es lo único que funciona con cualquier
  modelo que se cargue en `ajustes.json` sin tocar código. Dos consecuencias:
  el `MAX_TOKENS` ahora se reparte entre lo que el modelo razona y lo que
  escribe, y los tokens de salida del registro (y el costo estimado) incluyen
  ese razonamiento. La lectura de la respuesta ya lo contempla: filtra por
  `b.type == "text"`, así que los bloques de razonamiento no se cuelan en el
  informe.
- **`MAX_TOKENS = 16000` y chequeo de `stop_reason`** (2026-08-04). Estaba en 4000
  y no se miraba `stop_reason`: si el modelo llegaba al tope, el informe se
  cortaba a mitad de frase y se convertía a Word como si estuviera completo.
  Medido después del cambio: los informes reales usan **~3.600 tokens de salida**,
  o sea el 90% del tope viejo — se venía salvando por poco, y un informe regional
  o con la sección de entrevistas confidenciales se habría cortado. El tope no
  cuesta nada (se paga lo generado, no el tope). `max_tokens` → avisa fuerte por
  `log()` y guarda igual los archivos; `refusal` y respuesta vacía →
  `RuntimeError` con mensaje en español.
- **Caché de prompt**: `_armar_contenido()` devuelve `(estable, variable)` y
  `generar()` manda el mensaje en dos bloques de contenido con `cache_control:
  ephemeral` en el primero. El prefijo estable son los informes de ejemplo; como
  el `system` se renderiza antes que los `messages`, ese único corte cachea
  instrucciones + guía + ejemplos. **El corte va exactamente ahí porque el caché
  es por prefijo**: si algo variable quedara antes del punto de corte, se pierde
  todo lo que viene después. Verificado contra la API real: el 2º informe de una
  tanda lee 15.789 tokens del caché (76% menos de entrada). Un informe suelto sale
  apenas más caro (la escritura cuesta 1,25×); el equilibrio son 2.
  `exportar_prompt()` simplemente concatena las dos partes — el modo sin clave no
  cambia.
- `_registrar_consumo()` deja en el registro los tokens y un costo estimado. Las
  tarifas están en `PRECIO_ENTRADA_POR_MILLON` / `PRECIO_SALIDA_POR_MILLON`, en un
  solo lugar y con la advertencia de que cambian.
- Prompt = `INSTRUCCIONES` (rol, no inventar, formato Markdown) +
  `estilo_informe.md` en `<guia_de_estilo>` + ejemplos en `<informe_ejemplo>` +
  JSON en `<datos_nueva_instancia>` + indicaciones opcionales.
- `exportar_prompt()` = modo sin clave: arma el mismo contenido en
  `..._para_claude.txt` para pegar gratis en claude.ai. La GUI usa este modo
  automáticamente si no hay api_key.
- `markdown_a_docx()`: genera SIEMPRE sobre `plantilla.docx` (vacía el cuerpo,
  conserva encabezado/pie en todas las páginas). Estética: Calibri 11 justificado;
  `#` centrado negrita subrayado (las DOS líneas del título del informe van con
  `#`); `##` negrita subrayado; **negrita**/*cursiva* inline; viñetas, numeradas y
  tablas.
- ⚠️ **`plantilla.docx` no define NINGÚN estilo de Word que necesitemos** — ni los
  de lista (`List Bullet`, `List Number`) ni los de tabla (ni siquiera `Table
  Grid`). Verificado dos veces, con dos bugs distintos. Como `markdown_a_docx()`
  se llama dentro de un `try` en `generar()`, un `KeyError` por estilo faltante no
  explota: se traga, no se genera el `.docx` y el cliente se queda sólo con el
  `.md`. **Regla para cualquier elemento nuevo: no pedirle un estilo a la
  plantilla; armarlo a mano.** Las viñetas usan sangría + marca de texto
  (`parrafo_lista()`); las tablas dibujan los bordes por XML (`tblBorders` en
  `agregar_tabla()`).
- **Tablas** (2026-08-04): `estilo_informe.md` le pide a la IA tablas para las
  dinámicas de columnas («¿Qué voy a hacer más? / diferente / dejar de hacer»),
  pero no había ninguna rama que las manejara: las líneas caían al `else` y los
  caracteres `|` se escribían crudos en el Word firmado. El parseo detecta la fila
  de celdas + la separadora (`|---|---|`), consume el bloque entero (por eso el
  bucle recorre con índice y no con `for linea in`) y emite una tabla real con el
  encabezado en negrita.
- **Lectura de los ejemplos, la otra mitad del mismo bug** (2026-08-06):
  `cargar_ejemplos()` extraía los .docx con `doc.paragraphs`, que **no devuelve
  los párrafos que están dentro de celdas de tabla**. O sea que el modelo recibía
  los cuatro informes de ejemplo SIN una sola tabla, mientras `estilo_informe.md`
  le pedía producirlas: la dinámica de columnas se le exigía sin mostrarle nunca
  una. Los cuatro ejemplos tenían exactamente una tabla cada uno y se perdían
  entre 1.000 y 3.000 caracteres por informe (26.854 → 33.702 en total, ~8.400
  tokens). Ahora `_texto_de_docx()` recorre `doc.element.body.iterchildren()` y
  despacha por tag (`w:p` / `w:tbl`): **se recorre el cuerpo y no `doc.paragraphs`
  + `doc.tables` para no perder el orden**, porque la tabla tiene que quedar donde
  está en el informe (justo después de "se propuso una reflexión personal y
  confidencial…") y no amontonada al final. `_tabla_a_markdown()` la emite en
  pipes, el mismo formato que `markdown_a_docx()` sabe leer de vuelta; las celdas
  con varios párrafos se aplastan a una línea con `" ".join(texto.split())` porque
  un salto adentro de la fila partiría la tabla Markdown en dos. Verificado: el
  texto fuera de las tablas queda idéntico al anterior en los cuatro, el
  round-trip leer→escribir→leer es idempotente y el Word regenerado no tiene
  ningún `|` crudo.
- **Los ejemplos también pueden ser .odt** (2026-08-06): el cliente escribe los
  informes en LibreOffice, así que los que aportó llegaron en `.odt` y
  `cargar_ejemplos()` los **ignoraba en silencio** (la extensión no matcheaba
  ninguna rama: no hay un `else` que avise). `_texto_de_odt()` los lee con
  `zipfile` + `xml.etree`, **sin agregar dependencias**, igual que `_filas_ods()`
  en `procesar.py`: un .odt es un ZIP con `content.xml`. Recorre el cuerpo en
  orden despachando por tag (`text:p`/`text:h` → texto, `table:table` → Markdown
  de pipes, `text:list`/`text:section` → se entra a buscar adentro), con el mismo
  criterio que `_texto_de_docx()`. Detalles que importan: el texto se saca con
  `itertext()` porque Writer parte los párrafos en `text:span` en cada cambio de
  formato (con `.text` se perdían pedazos); las filas se buscan con `iter()` y no
  `findall()` para tomar también las que Writer envuelve en
  `table:table-header-rows`; el membrete no molesta porque en OpenDocument el
  encabezado y el pie están en `styles.xml`, no en `content.xml`.
  `_filas_a_markdown()` es común a los dos lectores y **rellena las filas cortas
  al ancho del encabezado**, porque en .odt las celdas vacías del final no se
  escriben y la tabla salía torcida.
- Extensiones aceptadas en `informes_ejemplo/`: `.txt`, `.md`, `.docx` y `.odt`.
  Los archivos llamados `leeme` o `readme` se saltean, para poder dejar una nota
  en esa carpeta sin que entre al prompt.
- **Desglose médicos / no médicos en la Metodología** (2026-08-06, pedido expreso
  del usuario). **Ninguno de los 16 informes de ejemplo lo trae** —la fórmula de
  la casa es «a la cual concurren N funcionarios del referido Servicio», y cuando
  desglosan lo hacen por sector («5 del Block Quirúrgico y 5 del…»)—, así que el
  corpus empuja en contra y una regla suelta en `estilo_informe.md` perdería
  contra 16 ejemplos que lo omiten. Por eso va en las dos partes: la regla en la
  guía (editable por el usuario) **y** un bloque en la parte VARIABLE de
  `_armar_contenido()` que nombra los números y aclara textualmente que los
  ejemplos no lo traen y que hay que incluirlo igual. Verificado contra la API:
  aparece.
  - **El desglose es de los FORMULARIOS recibidos, no de los asistentes.**
    `no_medicos` es `respuestas - medicos` (`procesar_grupo()`), así que médicos +
    no médicos siempre da la cantidad de respuestas; pero `asistentes` es un dato
    MANUAL e independiente y puede ser mayor. Si el bloque no lo aclarara, el
    modelo pegaría el desglose al número de concurrentes y firmaría una cuenta que
    no cierra («concurren 12 funcionarios, de los cuales 3 médicos y 4 no
    médicos»). El aviso extra sólo se agrega cuando `asistentes` existe y difiere.
    Probado con asistentes=12 y 7 formularios: separa bien las dos frases.
  - Ojo con el sentido de `no_medicos`: es un residuo (todo el que no marcó
    exactamente "medico", incluido quien dejó el campo vacío), no una
    clasificación afirmativa. Es el mismo criterio con el que ya se informa en el
    XLSX, así que el informe y la planilla dicen lo mismo.
- **Notas al pie: se descartan al leer y no se generan nunca** (2026-08-06). La
  nota «Ver Cuadro Relevamiento de Intervención Grupal» es una convención de la
  Unidad (estaba en 3 de los 4 `.docx` y en 11 de los 12 `.odt` del corpus original), y causaba dos
  problemas encadenados. **Al leer**: en OpenDocument la nota va INLINE dentro del
  párrafo, así que `itertext()` la metía en medio de la frase — «…los siguientes
  datos**1Ver cuadro adjunto anexo.**:» en 11 de los 12 ejemplos.
  `_texto_de_odt()` ahora saca los `text:note` del árbol antes de leer,
  **preservando el `tail`** (el texto pegado después de la nota, que sí es del
  informe: sin eso se perdían los dos puntos). Se descartan y no se reubican
  porque no son contenido, son punteros al Excel adjunto, y porque el lector de
  `.docx` tampoco las ve (allí viven en `footnotes.xml`, fuera del cuerpo): así
  los dos formatos quedan con el mismo criterio. **Al escribir**: el modelo
  copiaba el patrón y emitía `[^1]`, que `markdown_a_docx()` no conoce y escribía
  crudo en el Word firmado — la misma fuga que tenían los `|` de las tablas. Tres
  capas: los ejemplos ya no traen notas, `INSTRUCCIONES` las prohíbe
  explícitamente explicando por qué, y `_sin_notas_al_pie()` es la red final —saca
  el marcador de la oración y conserva el texto como un párrafo «Nota: …» al
  final, que es donde una nota al pie se termina leyendo. **No se generan notas al
  pie de verdad a propósito**: python-docx no tiene API y la plantilla no define
  estilos (ver la regla de no pedirle estilos a la plantilla).
- **Los 12 informes `.odt` no traen la dinámica de columnas**; los únicos 4 que la
  tienen son los `.docx` originales. Si algún día se depura el corpus, esos 4 no
  se sacan: son la única muestra del formato de tabla que `estilo_informe.md` le
  exige producir al modelo.

## plantilla.docx

Historia importante: la hoja original tenía el logo como imagen flotante
(wrapNone, behindDoc, offsets negativos) que se superponía al texto según el
visor. Se reconstruyó el encabezado como tabla invisible 1×2: logo recortado
(6,9 cm, tamaño original) a la izquierda, texto Verdana 10/7/7/8 pt gris 808080
alineado a la derecha. El pie original quedó intacto. **NO volver a imágenes
flotantes en el encabezado.**

Para cambiar el membrete se reemplaza ese archivo (se puede convertir un `.odt` a
`.docx` con LibreOffice: Archivo → Guardar como).

## estilo_informe.md

Destilado de un corpus de informes previos de la Unidad: estructura INFORME TÉCNICO → Antecedentes → Metodología
de trabajo → Análisis (párrafos A- B- C- D- por dimensión, porcentajes en prosa,
"percepción negativa o neutra") → Conclusiones y recomendaciones por eje → "En
suma," → "Montevideo, [fecha].". Prohíbe marcadores tipo "(completar)"; sin
antecedentes aportados usa fórmula genérica. **Editable por el usuario sin tocar
código.**

Tiene además una sección **"Tipos de informe"** con las reglas de cada uno
(encabezado, Metodología y foco del Análisis). `generar_informe._bloque_alcance()`
agrega un `<alcance_del_informe>` con el período, las fechas, las unidades y los
servicios concretos, y la advertencia de que los datos vienen sumados. Sin eso,
como todos los ejemplos de estilo son de un taller de un día, el modelo escribe
"la instancia del día de hoy" aunque el informe resuma meses.

## app.py + web/ (la interfaz)

- **Ventana nativa con HTML adentro** (pywebview sobre WebView2, que ya viene con
  Windows 10/11). No hay servidor, ni puerto, ni navegador. Migrado desde tkinter
  en 2026-08-03.
- Reparto: `web/` tiene todo lo visual y la interacción; `app.py` es sólo el
  puente con `procesar.py` / `generar_informe.py`, que no se tocaron.
- La clase `Puente` es la `js_api`: sus métodos públicos se llaman desde JS como
  `pywebview.api.<nombre>(...)` → `estado_inicial`, `guardar_ajustes`,
  `elegir_archivo`, `abrir_resultados`, `abrir_planilla`, `ejecutar`. Python le
  habla al JS con `_js(funcion, *args)`, que serializa con `json.dumps` y llama
  `evaluate_js`; el JS expone `window.agregarLog(texto)` y `window.terminar(ok,
  mensaje)`.
- **Los atributos internos de `Puente` van con guión bajo** (`_ventana`,
  `_ajustes`, `_trabajando`). pywebview recorre los atributos públicos de la
  js_api para publicarlos al JS; si encuentra la ventana se mete en los objetos
  COM de WebView2 hasta agotar la recursión y llena la consola de errores. No
  sacarles el guión bajo.
- `ejecutar(datos)` valida, arranca un hilo (`_trabajo` o `_trabajo_xlsx`) y
  devuelve enseguida `{ok: True}`; si falla la validación devuelve `{ok: False,
  error, abrir_config}` y el JS muestra el aviso rojo y abre la configuración.
- **Arranca maximizada** (`maximized=True`). Se usa `maximized` y NO `fullscreen`
  a propósito: fullscreen saca la barra de título y el cliente se quedaría sin la
  X para cerrar. Ojo al medir: Windows extiende las ventanas maximizadas 9 px por
  lado (borde invisible de redimensionado), así que `GetWindowRect` da
  (-9,-9)-(1929,1029) en una pantalla de 1920x1080; no es un bug.
- `geometria_ventana()` / `area_util()` definen el tamaño de la ventana
  **restaurada** (cuando el usuario la achica), recortado al escritorio disponible
  consultando Win32 (SPI_GETWORKAREA + LOGPIXELSX). No usar `webview.screens`:
  devuelve píxeles físicos o lógicos según si el proceso ya declaró DPI awareness,
  cosa que pywebview recién hace dentro de `start()`. Sin este recorte, con
  escalado al 125 % la ventana restaurada se pasa de la pantalla y el botón
  EJECUTAR queda tapado.
- El CSS acompaña los dos tamaños: `.app` va centrada con `max-width: 1000px`
  (maximizada no queda una columna angosta con vacío a los costados) y el registro
  usa `height: clamp(118px, 16vh, 240px)` para aprovechar la pantalla grande sin
  comerse el formulario.
- **CSS**: la regla `[hidden] { display: none !important; }` es necesaria — el
  atributo `hidden` lo aplica el navegador con la prioridad más baja y cualquier
  clase con `display: flex` lo pisa (pasaba con el spinner y el aviso de
  resultado). Layout de alto fijo: sólo `.tarjeta` (el formulario) scrollea, así
  EJECUTAR y el registro quedan siempre visibles.
- `python codigo\app.py --debug` abre las DevTools. `--autotest` verifica los
  recursos embebidos (incluye `web/`).
- **`registro.log` junto al programa** (`configurar_registro()`,
  `RotatingFileHandler` 1 MB × 3): todo lo que pasa por `Puente.log()` va también
  al archivo, y `_fallo()` agrega el traceback completo (en la ventana sólo se ve
  `str(e)`). Existe para diagnosticar a distancia: el recuadro de actividad se
  borra al cerrar y no quedaba nada que pedirle al cliente. Si no se puede
  escribir el archivo, se usa un `NullHandler` y el programa sigue igual — el
  registro es una ayuda, no un requisito. Está en `EXCLUIR` de
  `preparar_entrega.py` porque lleva el ID de la planilla.
- **Selector de tipo de informe** arriba de todo (2 tarjetas con radio: Instancia
  y Por Unidad Ejecutora). Cada `.campo` lleva `data-para="instancia unidad"` y
  `aplicarTipo()` muestra sólo los que corresponden. La U.E. es un `<select>`
  poblado desde la planilla; Servicio/Área y las fechas son **listas de casillas**
  (`#lista-servicios` / `#lista-fechas`), que comparten las clases CSS `.opcion`,
  `.opcion-nombre` y `.opcion-cuenta`. **El JS no conoce `multifecha`**: manda
  siempre `tipo="instancia"` y deja que `resolver_tipo()` decida en Python.
- **Botón "Abrir la planilla de respuestas"** (2026-09-09), el verde de la barra
  de origen: abre el Google Sheets del formulario. **Lo abre Python en el
  navegador del sistema (`Puente.abrir_planilla()` → `abrir_enlace()` →
  `webbrowser.open`) y NO puede ser un `<a href>`**: la interfaz vive adentro de
  la ventana de pywebview, así que un enlace común la haría navegar a Google y el
  cliente se quedaría sin interfaz y sin botón para volver. La dirección se arma
  con el `sheet_id` de `ajustes.json` (`URL_PLANILLA`), no está escrita a mano: si
  se cambia de planilla, el botón sigue apuntando a la que está configurada. Sin
  `sheet_id` cargado el botón no se muestra (`mostrarBotonPlanilla()` en el JS,
  desde `volcarAjustes()` y al guardar la configuración), y es el único de la
  barra que NO se bloquea mientras se trabaja: mirar las respuestas no interfiere
  con el informe que se está generando.
- **Fechas de la instancia** (`poblarFechas()` / `fechasMarcadas()` → `fechas` en
  el dict `datos` → parámetro `fechas` de `procesar_datos`): una casilla por día,
  acotadas a la U.E. elegida, con la cantidad de respuestas de cada día y en el
  orden que ya trae `opciones.fechas` (lo ordena Python, no el JS). Se exige al
  menos una marcada (`queFalta()`). Ya no hay un `<select>` de fecha: la lista de
  casillas es la única.
- **Las fechas de la lista van de la más reciente a la más antigua**
  (`opciones_disponibles()` devuelve `ordenar_fechas()` invertido): el informe que
  se pide es casi siempre el de lo último relevado, así que arriba de todo queda
  la fecha que se va a marcar. Es sólo el orden de lo que se muestra:
  `fechas_incluidas`, `periodo` y `texto_periodo()` siguen en orden cronológico,
  que es lo que necesitan el Excel, las notas de alcance y el informe.
- **Las listas se encadenan U.E. → fechas → servicios**: cada una ofrece sólo lo
  que existe para lo ya elegido. `opciones_disponibles()` devuelve `combinaciones`
  (fecha + unidad + región + servicio normalizado, con su conteo) y el JS cruza
  contra eso en `combinaciones()`/`filtroActual()`. La U.E. encabeza la cadena y
  ofrece todas las unidades; elegirla acota las fechas ofrecidas, y las fechas
  marcadas acotan los servicios. **La U.E. va primero a propósito**: las fechas
  son de multiselección, y si se marcaran antes se podrían juntar días de
  hospitales distintos y salir un informe con menos respuestas de las esperadas,
  sin aviso. El conteo que se muestra al lado de cada servicio es el del
  subconjunto elegido, no el de toda la planilla. Reportado el 2026-08-04: con el
  28/7 elegido aparecía igual "Laboratorio/ pediatría", que es del 29/7.
- Al repoblar las listas de servicios y de fechas se conserva lo que estaba
  marcado si sigue existiendo en la nueva selección (los servicios se cruzan por
  `clave`, la forma normalizada; las fechas, por su texto).
- **Campo "Nombre para el informe"** (`servicio_nombre` → parámetro `servicio` de
  `procesar_datos`): es el nombre con el que el servicio figura en el renglón
  SERVICIO/ÁREA del Excel, en el título del informe y en la carpeta de salida. Se
  sugiere solo (la etiqueta más corta de las marcadas) y deja de sugerirse en
  cuanto se escribe a mano. `servicios_incluidos` sigue guardando las variantes
  reales: es el registro de qué entró de verdad. El valor se calcula en
  `datosDelFormulario()` con `nombreSugerido()` si el campo está vacío, para que no
  dependa de que se haya disparado el evento `change`.
- **Las listas se llenan leyendo la planilla al abrir** (`cargar_opciones()` →
  `procesar.opciones_disponibles`). La barra de origen arriba dice de dónde salen
  los datos y cuántas respuestas hay; si falla, se pinta de rojo con el motivo y
  EJECUTAR queda bloqueado. El botón "Usar archivo descargado…" ya **no ejecuta**:
  cambia el origen a un CSV y recarga las listas; después se ejecuta normalmente
  con EJECUTAR.
- Ya no hay campo de texto "Servicio/Área": ahora sale de la planilla. El
  parámetro `servicio=` de `procesar_datos` se conserva para la consola.
- La validación de qué falta según el tipo está en el JS (`queFalta()`), para dar
  el mensaje sin ir y volver a Python.
- **`abrirConfiguracion()` vuelve a pedir `estado_inicial()` cada vez.** El estado
  de `credenciales.json` se leía una sola vez al arrancar, así que si el archivo
  se copiaba con el programa abierto la ventana seguía diciendo que faltaba para
  siempre (reportado y corregido el 2026-08-04). El mensaje ahora incluye además
  la carpeta exacta donde hay que ponerlo (`carpeta_credenciales`).
- Campos comunes: asistentes, **antecedentes/contexto** (textarea → viaja como
  `indicaciones` a la IA en ambos modos), **entrevistas individuales
  confidenciales** (textarea debajo del contexto → viaja como
  `entrevistas_confidenciales` a `generar`/`exportar_prompt`), checkbox "Generar
  también el informe con IA".
- **Entrevistas confidenciales inventadas** (2026-08-06, arreglado): con
  `entrevistas_individuales > 0` en el JSON y el campo de la ventana VACÍO, el
  modelo escribía igual la sección entera y **fabricaba qué habían dicho las
  personas** en las entrevistas. Prosa inventada sobre material confidencial, en
  un informe que se firma, y se disparaba solo: alcanzaba con que hubiera
  entrevistas contadas en la planilla y que no se pegara el texto. Medido con los
  dos corpus de ejemplos (4 y 16), así que **no depende de cuántos ejemplos haya**.
  Lo importante: `estilo_informe.md` ya lo prohibía DOS veces (la regla de la
  sección y «no inventar testimonios») **y no alcanzó** — la presión de ver el
  número en el JSON y la sección en los informes de ejemplo le gana a la guía de
  estilo. La prohibición está ahora en la rama `elif` de `_armar_contenido()`,
  pegada a los datos y nombrando explícitamente lo que no hay que hacer (incluido
  «no la tomes de los informes de ejemplo, que sí traen esa sección porque a ellos
  se les aportó el material»). Va en la parte VARIABLE porque depende del JSON.
  Verificado contra la API: sin material no aparece la sección y la cantidad sigue
  mencionándose en la Metodología; el camino con material aportado es byte a byte
  el de antes (la rama nueva es un `elif`, no se toca el `if`). **Moraleja para
  reglas que de verdad no se pueden violar: no dejarlas sólo en
  `estilo_informe.md`.**
- **Entrevistas confidenciales**: si el campo tiene texto, `_armar_contenido()`
  inyecta un bloque `<entrevistas_individuales_confidenciales>` con la instrucción
  de armar una sección DESPUÉS del Análisis (tras "D- Motivación") y ANTES de
  Conclusiones, que empieza textualmente con "Es dable señalar que se solicitaron
  X entrevistas Individuales Confidenciales, de las cuales surge:" (X =
  `entrevistas_individuales` del JSON). La regla estructural también está en
  `estilo_informe.md`. Si el campo queda vacío, no se agrega la sección.
- El JS junta todos los campos en un dict `datos` (con `origen`: `sheets` | `csv`
  | `xlsx`) y lo manda en una sola llamada; el hilo trabajador de Python lo lee de
  ahí. El log llega por `agregarLog` a medida que ocurre, sin cola ni polling.
- `app.py` no hace `os.chdir`: todo se resuelve con rutas absolutas (`rutas.py` y
  los `config_path`/`credenciales`/`salida` explícitos). Había un chdir heredado
  de antes del empaquetado; se sacó el 2026-08-04 tras verificar que nada dependía
  del directorio actual.
- **Botón "Ver entrevistas pendientes"** (debajo de EJECUTAR) → modal
  `#modal-entrevistas` con **dos pestañas** (Pendientes / Completadas) que
  comparten una tabla U.E. / Fecha / Nombre / Contacto. El botón al final de cada
  fila cambia según la pestaña: **«Completada»** la marca como hecha,
  **«Deshacer»** la devuelve a pendientes. `Puente.entrevistas_pendientes(ruta_csv)`
  relee la planilla cada vez que se abre (así aparecen los pedidos nuevos) y
  devuelve las dos listas ya separadas; `Puente.marcar_entrevista(clave, hecha)`
  anota o desanota. **Se guarda en `entrevistas_completadas.json`, junto al
  programa, y NO en el Sheets**: la cuenta de servicio de Google entra como
  Lectora, el programa no puede escribir en la planilla. El archivo es `{clave:
  cuándo se marcó}` y está en `EXCLUIR` de `preparar_entrega.py`.
  - Las completadas vienen ordenadas por cuándo se marcaron, **las últimas
    primero** (`_orden_marca()`), que es el orden en que se busca la que se marcó
    por error. Una marca con formato ilegible va al final en vez de romper el
    orden.
  - **La fila se mueve de lista recién cuando Python confirmó que guardó.** Si se
    moviera antes, un cambio que no se pudo escribir (carpeta de sólo lectura) se
    vería como aplicado y esa entrevista se perdía.
  - El JS mantiene las dos listas en memoria (`entrevistas.pendientes` /
    `.hechas`) y al marcar mueve el elemento de una a la otra **sin releer la
    planilla**: releerla en cada clic haría esperar un segundo cada vez.
  - El botón de cada fila lleva `.btn-fila` además de `.btn-mini`: `.btn-mini` es
    `inline-block` y ahí el ícono y el texto caen en dos renglones, con filas de
    70 px en vez de 45.
- Botón "Usar archivo descargado…" = modo CSV. **Botón "Informe desde planilla…"**
  = genera el informe a partir de una planilla en formato oficial hecha a mano,
  `.ods` o `.xlsx` (`_trabajo_xlsx` → `procesar.datos_desde_xlsx` → escribe el
  JSON y copia la planilla a la subcarpeta de la instancia conservando su
  extensión → informe; siempre genera informe, ignora el checkbox; usa los campos
  antecedentes/entrevistas). Los dos piden el archivo con `elegir_archivo(tipo)`,
  que abre el diálogo nativo de Windows. "⚙ Configuración" es un modal HTML que
  edita ajustes.json (se abre sola la primera vez, y también cuando falta el
  sheet_id al ejecutar); incluye el campo **modelo**. Al terminar OK abre la
  carpeta de resultados.
- La generación del informe está factorizada en `generar_informes(rutas_json,
  indicaciones, entrevistas_conf)` e `indicaciones_de(antecedentes)`, compartidas
  por el flujo normal (`_trabajo`) y el de planilla (`_trabajo_xlsx`).
- **Solo entregables (XLSX + DOCX)**: en el flujo con API, tras generar el DOCX
  OK, `dejar_solo_entregables()` borra el JSON y el MD intermedios de esa
  instancia. Recaudos: si el DOCX falla se conserva el MD (único informe); en modo
  sin clave no se borra nada (el JSON es insumo del `_para_claude.txt`). El CLI de
  `generar_informe.py` sigue dejando los 4 archivos (uso dev).
- **Carpeta de resultados**: la app escribe en `<Escritorio>\InformesIntervenciones`
  (`rutas.carpeta_informes()`, constante `CARPETA_RESULTADOS`), no en `salida\`. En
  Windows el Escritorio se resuelve por el registro (contempla redirección a
  OneDrive); fallback a `~/OneDrive/Desktop` y `~/Desktop`. El CLI de `procesar.py`
  mantiene `--salida` con default `salida` (uso consola/dev). Dentro, cada
  instancia queda agrupada en su propia subcarpeta `{Unidad}-{mes}-{año}` (ej.:
  `HospitalDelNorte-07-26`).

## config.json (importante al cambiar el formulario)

En `recursos\config.json` están definidas las secciones y el **texto exacto de
cada pregunta**, que debe coincidir con los encabezados de columna que genera el
formulario (la comparación ignora mayúsculas y tildes). Si al procesar aparece el
aviso "preguntas no encontradas", hay que copiar el encabezado tal cual figura en
el Sheets.

En la sección `columnas` están los nombres de las columnas de clasificación:
`fecha`, `marca_temporal`, `region`, `servicio`, `tipo_encuestado`,
`nombre_completo`, `celular`, `solicita_entrevista` y `respuesta_afirmativa`. Si
en el formulario cambia el título de alguna de esas preguntas, hay que
actualizarlo acá.

> **Sobre la confirmación de entrevista**: `solicita_entrevista` es el texto
> exacto de la pregunta que confirma si la persona quiere una entrevista
> individual, y `respuesta_afirmativa` es cómo está escrita la opción que
> significa que sí (por defecto `"Sí"`). Si en el formulario esa opción dice otra
> cosa —por ejemplo `"Sí, me interesa"`— alcanza con corregir
> `respuesta_afirmativa`; no hace falta tocar el programa. La comparación ignora
> mayúsculas y tildes.
>
> Si esa columna **no existe** en la planilla, el programa avisa y vuelve al
> criterio anterior: cuenta como pedido de entrevista toda fila con nombre y
> celular cargados. Eso permite seguir procesando planillas viejas, anteriores a
> que se agregara la pregunta.

> **Sobre la fecha**: el programa usa la columna `Fecha` del formulario y, si
> viene vacía, la fecha de la **marca temporal** que Google completa sola. Al día
> de hoy la pregunta de fecha del formulario no se está respondiendo, así que en
> la práctica se usa la marca temporal. Si se quiere que mande la fecha declarada
> (por ejemplo, porque se cargan respuestas en papel días después), hay que hacer
> obligatoria esa pregunta en el Google Form.

## Retocar la interfaz

La ventana es HTML: está en la carpeta **`web\`** y se puede editar con cualquier
editor de texto, sin tocar Python.

- `index.html` — los campos, los botones y sus textos de ayuda.
- `estilo.css` — colores, tamaños y espaciados (el azul institucional está
  definido una sola vez, en `--azul`).
- `app.js` — qué hace cada botón; se comunica con Python llamando a
  `pywebview.api.<lo que sea>`.

Se guarda el archivo, se cierra y se vuelve a abrir la aplicación y listo. Si se
trabaja con el `.exe`, hay que correr `desarrollo\construir_exe.bat` de nuevo.
Para depurar, `python codigo\app.py --debug` abre las herramientas de desarrollo.

---

# 7. Decisiones tomadas con el usuario

- Ejecución local a demanda en la PC del cliente; API de Google con cuenta de
  servicio; GUI porque el cliente no es técnico.
- **Ventana nativa con HTML adentro, no web app** (2026-08-03): se evaluó pasar a
  FastAPI + navegador y se descartó para este caso (un solo usuario, local, no
  técnico), porque agregaba puerto, posible proceso huérfano al cerrar la pestaña
  y confusión entre descargas y carpeta de resultados, a cambio de nada. pywebview
  da la libertad visual del HTML sin ninguna de esas contras. Si algún día la usan
  varias personas o desde otra PC, ahí sí conviene la web app de verdad (y cambia
  el modelo de seguridad: `credenciales.json` y la api_key pasan a ser
  compartidas).
- Modelo **Opus 5** por pedido explícito (2026-09-10; antes Opus 4.8, también
  por pedido). Más caro que Sonnet, y reversible sin tocar código desde el campo
  «Modelo de IA» de `ajustes.json`.
- Modo sin clave para evaluar el resultado antes de pagar la API.
- Todo el código, comentarios y mensajes en español rioplatense (voseo en la UI).
- **Sólo Windows** (2026-08-04): se descartó el lanzador para Mac. Se borró
  `iniciar.sh`, que existía pero nunca se probó con pywebview (en Mac usa
  WebKit/Cocoa o Qt, otro backend). El uso por consola sigue siendo
  multiplataforma porque no toca la ventana.

---

# 8. Convenciones para seguir desarrollando

- Errores de negocio como `RuntimeError` con mensajes claros en español (la
  interfaz los muestra sin traducir, en el aviso rojo).
- `procesar.py` y `generar_informe.py` deben seguir funcionando por consola además
  de desde la interfaz (main() delgados sobre funciones reutilizables), y no deben
  saber nada de la UI. Ese desacople fue lo que permitió cambiar tkinter por
  pywebview sin tocarlos.
- Los cambios visuales van en `web\`, no en Python. `app.py` sólo debería crecer
  si hace falta una función nueva del puente.
- **Los `.bat` se guardan con saltos CRLF y sólo ASCII.** Con LF sueltos
  `cmd.exe` se rompe adentro de los bloques `if ... ( ... )` (los tres `.bat`
  venían así y `Iniciar.bat` tiraba `"m" no se reconoce como un comando` antes de
  abrir; corregido 2026-08-03). Los caracteres acentuados o guiones largos en un
  `.bat` también salen como basura, porque cmd no los lee como UTF-8.
- No romper la compatibilidad del formato XLSX (es el formato oficial que ya
  usan) ni quitar el modo CSV (plan B sin conexión).
- `credenciales.json` y `ajustes.json` contienen secretos: nunca subirlos a
  repositorios ni incluirlos en ZIPs de distribución.

---

# 9. Uso por consola (opcional)

La aplicación de ventana cubre el uso normal. Los mismos pasos también se pueden
ejecutar por consola, parado en cualquier carpeta (los defaults se resuelven con
`recurso()` / `archivo_externo()`).

**Procesar** (después de cada instancia):

```
python codigo\procesar.py --sheet-id TU_SHEET_ID --fecha 10/02/26 --asistentes 15 --servicio "Enfermería"
```

Opciones útiles:

- `--tipo instancia|multifecha|unidad` → qué informe generar (por defecto
  `instancia`). Se sigue aceptando el nombre viejo `historico` como sinónimo de
  `multifecha`. Entre `instancia` y `multifecha` no hace falta acertar: si se
  pasan varias fechas se usa `multifecha`, y si se pasa una sola, `instancia`.
- `--unidad "Hospital Central"` → procesa solo esa U.E. En el
  tipo `instancia`, si no se indica, genera un archivo por cada unidad que tenga
  respuestas.
- `--fecha 10/02/26` → filtra por la fecha de la instancia (solo con `--tipo
  instancia`).
- `--fechas 28/07/26 04/08/26` → las jornadas de la instancia (solo con `--tipo
  multifecha`). Si no se indican, entran todas las fechas de esa U.E. y servicio.
- `--servicios "Emergencia" "Emergencia adultos"` → uno o más valores de
  Servicio/Área; sirve para juntar las variantes de escritura de un mismo
  servicio.
- `--asistentes 15` → cantidad de asistentes (dato manual, no sale del
  formulario).
- `--csv archivo.csv` → alternativa sin API: se descarga la planilla como CSV y se
  procesa localmente. Probarlo con el incluido:
  `python codigo\procesar.py --csv desarrollo\ejemplo_respuestas.csv --asistentes 15`
- `--worksheet "Respuestas de formulario 1"` → si la hoja no es la primera de la
  planilla.
- `--listar` → no genera nada: muestra las fechas, unidades, regiones y servicios
  que hay cargados. Es la forma rápida de ver qué se puede pedir.
- `--entrevistas` → no genera nada: muestra quiénes pidieron una entrevista
  individual. Las lista todas, sin descontar las ya realizadas (eso lo lleva la
  aplicación en `entrevistas_completadas.json`).

Ejemplos de los otros tipos:

```
python codigo\procesar.py --csv desarrollo\ejemplo_respuestas.csv --tipo multifecha --unidad "Hospital Central" --fechas 18/7/2026 25/7/2026 --servicios "Emergencia" "Emergencia adultos"
python codigo\procesar.py --csv desarrollo\ejemplo_respuestas.csv --tipo unidad --unidad "Hospital Central"
```

Por consola los resultados quedan en `salida\` (se cambia con `--salida`): un
`.xlsx` (el relevamiento con formato) y un `.json` (resumen para el informe). La
**aplicación de escritorio**, en cambio, los guarda en la carpeta
**`InformesIntervenciones`** del Escritorio.

**Generar el informe:**

```
python codigo\generar_informe.py salida\relevamiento_hospital_central_100226.json
```

Opcional: `--indicaciones "destacar el eje de motivación"` para dar instrucciones
puntuales, `--modelo` para cambiar el modelo, `--sin-clave` para el modo sin API y
`--desde-xlsx archivo.ods` para partir de una planilla hecha a mano.

El informe queda junto al JSON como `..._informe.md` y `..._informe.docx`
(revisarlo y ajustarlo antes de distribuirlo: la redacción final es
responsabilidad de quien lo firma).

---

# 10. Estado y pendientes

**`config.json` calibrado contra el formulario real y probado end-to-end con API
key real** (2026-07-20).

**Migración a pywebview verificada en la PC del cliente** (2026-08-03), con Python
3.14.6 y pywebview 6.2.1 sobre WebView2. Probado: pipeline de datos con CSV de
estructura real (2 unidades, entrevistas individuales, respuestas omitidas),
round-trip XLSX→resumen idéntico, Word sobre la plantilla membretada, flujo CSV
completo por la interfaz con el log llegando en vivo, estado "trabajando", los
tres caminos de error (fecha inexistente, falta sheet_id, Excel en formato
equivocado), guardado de configuración, ver/ocultar la clave y cierre con Escape.
Y lo mismo desde el `.exe` empaquetado: recursos embebidos OK y ventana abriendo
con la interfaz HTML.

**Empaquetado verificado** (2026-09-09): PyInstaller 6.22.2 sobre Python 3.14.7
construye `dist\Relevamiento` desde `desarrollo\relevamiento.spec` sin errores en
~41 s; `--autotest` da "todos los recursos OK" y el `.exe` abre la ventana y lee
la planilla real.

Pendientes conocidos:

1. ~~Calibrar `config.json`~~ ✅ Hecho: los textos de las preguntas coinciden con
   los encabezados reales del Sheets.
2. ~~Probar end-to-end con API key real~~ ✅ Hecho.
3. ~~Empaquetar .exe con PyInstaller~~ ✅ Hecho. ~~Campo "modelo" en
   Configuración~~ ✅ Hecho.
4. ~~Lanzador de doble clic para Mac~~ ❌ Descartado (2026-08-04): el proyecto es
   sólo Windows.
5. `desarrollo\crear_icono.py` **nunca se probó** (falta Pillow en esta PC) y el
   diseño se dedujo decodificando el `.ico` existente, así que regenera algo
   parecido pero no idéntico. Por eso escribe en `icono_nuevo.ico` salvo que se le
   pase `--sobrescribir`. El `.ico` bueno está versionado en `recursos\`.
6. Idea pendiente (mejora de UX, no bug): ya no aplica al uso normal —fecha y
   unidad salen de listas pobladas desde la planilla—, pero por consola siguen
   siendo texto libre.
