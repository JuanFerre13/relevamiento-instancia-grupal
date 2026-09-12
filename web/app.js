/* Relevamiento de Instancia Grupal — lógica de la interfaz.
 *
 * Todo el trabajo pesado vive en Python (procesar.py / generar_informe.py).
 * Acá sólo se juntan los datos del formulario, se los manda por el puente
 * `pywebview.api` y se muestra lo que Python va informando.
 *
 * Python llama a estas funciones globales mientras trabaja:
 *   window.agregarLog(texto)      una línea del registro de actividad
 *   window.terminar(ok, mensaje)  fin del trabajo (ok = true | false)
 */

const $ = (id) => document.getElementById(id);

let trabajando = false;
let modeloPorDefecto = "";
let opciones = null;          // lo que hay cargado en la planilla
let origenCsv = null;         // ruta del CSV elegido, o null = planilla de Google

/* ───────────────────────────── Arranque ───────────────────────────── */

window.addEventListener("pywebviewready", iniciar);

async function iniciar() {
  conectarBotones();
  const estado = await pywebview.api.estado_inicial();
  modeloPorDefecto = estado.modelo_default || "";
  $("modelo-default").textContent = modeloPorDefecto;
  volcarAjustes(estado);
  aplicarTipo();

  if (!estado.ajustes.sheet_id) {
    abrirConfiguracion();
    marcarOrigen("Falta configurar la planilla de Google.", true);
    bloquearEjecutar(true);
    return;
  }
  cargarOpciones();
}

function conectarBotones() {
  $("btn-ejecutar").addEventListener("click", ejecutar);
  $("btn-xlsx").addEventListener("click", ejecutarDesdeExcel);
  $("btn-resultados").addEventListener("click", () => pywebview.api.abrir_resultados());
  $("btn-recargar").addEventListener("click", () => { origenCsv = null; cargarOpciones(); });
  $("btn-csv").addEventListener("click", elegirCsv);
  $("btn-abrir-sheet").addEventListener("click", abrirPlanilla);

  for (const radio of document.querySelectorAll("input[name=tipo]")) {
    radio.addEventListener("change", aplicarTipo);
  }
  // Las listas se encadenan U.E. → fechas → servicios: cada una muestra sólo lo
  // que existe para lo ya elegido.
  $("unidad").addEventListener("change", () => { poblarFechas(); poblarServicios(); });
  $("btn-todos-servicios").addEventListener("click", () => marcarServicios(true));
  $("btn-ningun-servicio").addEventListener("click", () => marcarServicios(false));
  $("btn-todas-fechas").addEventListener("click", () => marcarFechas(true));
  $("btn-ninguna-fecha").addEventListener("click", () => marcarFechas(false));
  // Si lo escribe a mano, la sugerencia automática deja de pisarlo.
  $("nombre-servicio").addEventListener("input", () => {
    nombreServicioEditado = $("nombre-servicio").value.trim() !== "";
  });

  $("btn-entrevistas").addEventListener("click", abrirEntrevistas);
  $("btn-cerrar-entrevistas").addEventListener("click", cerrarEntrevistas);
  $("tab-pendientes").addEventListener("click", () => mostrarPestanaEntrevistas("pendientes"));
  $("tab-hechas").addEventListener("click", () => mostrarPestanaEntrevistas("hechas"));

  $("btn-config").addEventListener("click", abrirConfiguracion);
  $("btn-cancelar").addEventListener("click", cerrarConfiguracion);
  $("btn-guardar").addEventListener("click", guardarConfiguracion);
  $("btn-ver-clave").addEventListener("click", alternarClave);

  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape") return;
    if (!$("modal").hidden) cerrarConfiguracion();
    else if (!$("modal-entrevistas").hidden) cerrarEntrevistas();
  });
}

/* ──────────────────── Tipo de informe: qué campos van ───────────────── */

function tipoElegido() {
  return document.querySelector("input[name=tipo]:checked").value;
}

function aplicarTipo() {
  const tipo = tipoElegido();
  for (const campo of document.querySelectorAll(".campo[data-para]")) {
    campo.hidden = !campo.dataset.para.split(" ").includes(tipo);
  }
  // La ayuda sobre las fechas sólo tiene sentido donde se eligen fechas.
  $("ayuda-unidad").hidden = tipo !== "instancia";
  poblarUnidades();
  poblarFechas();
  poblarServicios();
}

/* ─────────────────── Opciones que hay en la planilla ────────────────── */

async function cargarOpciones() {
  marcarOrigen(origenCsv ? "Leyendo el archivo…" : "Leyendo la planilla de Google…", false);
  bloquearEjecutar(true);

  const respuesta = await pywebview.api.cargar_opciones(origenCsv);
  if (!respuesta.ok) {
    opciones = null;
    marcarOrigen(respuesta.error, true);
    vaciarListas();
    bloquearEjecutar(true);
    return;
  }

  opciones = respuesta.opciones;
  const fuente = origenCsv ? `Archivo: ${respuesta.nombre_archivo}` : "Planilla de Google";
  marcarOrigen(`${fuente} · ${opciones.respuestas} respuestas · `
    + `${opciones.fechas.length} fecha(s) · ${opciones.unidades.length} unidad(es)`, false);

  poblarUnidades();
  poblarFechas();
  poblarServicios();
  bloquearEjecutar(false);
}

function marcarOrigen(texto, esError) {
  $("origen-texto").textContent = texto;
  $("origen").classList.toggle("con-error", !!esError);
}

function opcionesDe(select, valores, primera) {
  const previo = select.value;
  select.innerHTML = "";
  if (primera !== null && primera !== undefined) {
    select.appendChild(new Option(primera, ""));
  }
  for (const v of valores) select.appendChild(new Option(v, v));
  if (previo && valores.includes(previo)) select.value = previo;
}

/* Las listas se encadenan: cada una ofrece sólo lo que existe para lo ya
 * elegido. En el informe de instancia acotan las fechas marcadas (sin marcar
 * ninguna, entran todas). */
function filtroActual() {
  return {
    fechas: tipoElegido() === "instancia" ? fechasMarcadas() : [],
    unidad: $("unidad").value,
  };
}

function combinaciones({ unidad, fechas }) {
  const marcadas = fechas && fechas.length ? new Set(fechas) : null;
  return (opciones.combinaciones || []).filter(
    (c) => (!marcadas || marcadas.has(c.fecha)) && (!unidad || c.unidad === unidad));
}

/* La U.E. encabeza la cadena, así que se ofrecen todas las que hay cargadas.
 * Siempre hay que elegir una: no existe "(todas)" (nunca se pidió un informe
 * por cada unidad de una fecha). */
function poblarUnidades() {
  if (!opciones) return;
  opcionesDe($("unidad"), opciones.unidades || [], null);
}

/* ─────────────────────── Fechas de la instancia ─────────────────────── */

/* Una casilla por día, acotadas a la Unidad Ejecutora elegida y de la fecha más
 * reciente a la más antigua, que es el orden en que las devuelve Python. Se
 * marcan los días en que se dictó la instancia: con una sola marcada sale el
 * informe puntual de siempre, y con varias, uno solo que las junta (el tipo lo
 * deriva `resolver_tipo()` en Python, acá no se elige). */
function poblarFechas() {
  if (!opciones) return;
  const caja = $("lista-fechas");
  const marcadas = new Set(
    [...caja.querySelectorAll("input:checked")].map((c) => c.value));

  const conteo = {};
  for (const c of combinaciones({ unidad: $("unidad").value })) {
    if (c.fecha) conteo[c.fecha] = (conteo[c.fecha] || 0) + c.respuestas;
  }
  const disponibles = opciones.fechas.filter((f) => conteo[f]);

  caja.innerHTML = "";
  if (!disponibles.length) {
    caja.innerHTML = '<p class="lista-vacia">No hay fechas cargadas para esa '
      + 'Unidad Ejecutora.</p>';
    actualizarResumenFechas();
    return;
  }

  for (const f of disponibles) {
    const fila = document.createElement("label");
    fila.className = "opcion";
    const casilla = document.createElement("input");
    casilla.type = "checkbox";
    casilla.value = f;
    casilla.checked = marcadas.has(f);   // se respeta lo que ya estaba marcado
    // Los servicios se acotan a las fechas marcadas, así que se rehacen.
    casilla.addEventListener("change", () => {
      actualizarResumenFechas();
      poblarServicios();
    });
    const nombre = document.createElement("span");
    nombre.className = "opcion-nombre";
    nombre.textContent = f;
    const cuenta = document.createElement("span");
    cuenta.className = "opcion-cuenta";
    cuenta.textContent = conteo[f];
    fila.append(casilla, nombre, cuenta);
    caja.appendChild(fila);
  }
  actualizarResumenFechas();
}

function fechasMarcadas() {
  return [...document.querySelectorAll("#lista-fechas input:checked")].map((c) => c.value);
}

function marcarFechas(valor) {
  for (const c of document.querySelectorAll("#lista-fechas input")) c.checked = valor;
  actualizarResumenFechas();
  poblarServicios();
}

function actualizarResumenFechas() {
  const marcadas = fechasMarcadas().length;
  $("resumen-fechas").textContent = marcadas
    ? `${marcadas} fecha(s) marcada(s)`
    : "marcá las fechas de la instancia";
}

function poblarServicios() {
  if (!opciones) return;
  const caja = $("lista-servicios");
  const marcados = new Set(
    [...caja.querySelectorAll("input:checked")].map((c) => c.dataset.clave));

  // Cuántas respuestas tiene cada servicio DENTRO de lo elegido hasta ahora.
  const conteo = {};
  for (const c of combinaciones(filtroActual())) {
    if (c.servicio) conteo[c.servicio] = (conteo[c.servicio] || 0) + c.respuestas;
  }
  const disponibles = opciones.servicios.filter((s) => conteo[s.clave]);

  caja.innerHTML = "";
  if (!disponibles.length) {
    caja.innerHTML = '<p class="lista-vacia">No hay Servicio/Área cargado para lo que '
      + 'elegiste. El informe va a incluir todas las respuestas que encuentre.</p>';
    actualizarResumenServicios();
    return;
  }

  for (const s of disponibles) {
    const fila = document.createElement("label");
    fila.className = "opcion";
    const casilla = document.createElement("input");
    casilla.type = "checkbox";
    casilla.value = s.etiqueta;
    casilla.dataset.clave = s.clave;
    // Se mandan todas las formas de escribirlo, para que el filtro las agarre todas.
    casilla.dataset.variantes = JSON.stringify(s.variantes);
    casilla.checked = marcados.has(s.clave);   // se respeta lo que ya estaba marcado
    casilla.addEventListener("change", actualizarResumenServicios);
    const nombre = document.createElement("span");
    nombre.className = "opcion-nombre";
    nombre.textContent = s.etiqueta;
    const cuenta = document.createElement("span");
    cuenta.className = "opcion-cuenta";
    cuenta.textContent = conteo[s.clave];
    fila.append(casilla, nombre, cuenta);
    if (s.variantes.length > 1) {
      fila.title = "Se escribió también como: " + s.variantes.join(" · ");
      nombre.textContent = `${s.etiqueta}  (${s.variantes.length} formas)`;
    }
    caja.appendChild(fila);
  }
  actualizarResumenServicios();
}

function serviciosMarcados() {
  const elegidos = [];
  for (const c of document.querySelectorAll("#lista-servicios input:checked")) {
    elegidos.push(...JSON.parse(c.dataset.variantes));
  }
  return elegidos;
}

function marcarServicios(valor) {
  for (const c of document.querySelectorAll("#lista-servicios input")) c.checked = valor;
  actualizarResumenServicios();
}

function actualizarResumenServicios() {
  const marcados = document.querySelectorAll("#lista-servicios input:checked").length;
  $("resumen-servicios").textContent = marcados
    ? `${marcados} marcado(s)`
    : "sin marcar: entran todos";
  sugerirNombreServicio();
}

/* El nombre con el que el servicio va a figurar en el Excel y en el título del
 * informe. Se sugiere solo, pero deja de sugerirse en cuanto la persona lo
 * escribe a mano: juntar los textos del formulario da cosas como
 * "Laboratorio, Laboratorio/ pediatría", que en un informe oficial queda mal. */
let nombreServicioEditado = false;

/* De las marcadas, la más corta suele ser el nombre "limpio" del servicio
 * ("Laboratorio" antes que "Laboratorio/ pediatría"). */
function nombreSugerido() {
  const etiquetas = [...document.querySelectorAll("#lista-servicios input:checked")]
    .map((c) => c.value);
  etiquetas.sort((a, b) => a.length - b.length || a.localeCompare(b));
  return etiquetas[0] || "";
}

function sugerirNombreServicio() {
  if (nombreServicioEditado) return;
  $("nombre-servicio").value = nombreSugerido();
}

function vaciarListas() {
  $("unidad").innerHTML = "";
  $("unidad").appendChild(new Option("—", ""));
  $("lista-servicios").innerHTML = '<p class="lista-vacia">Sin datos.</p>';
  $("lista-fechas").innerHTML = '<p class="lista-vacia">Sin datos.</p>';
  actualizarResumenServicios();
  actualizarResumenFechas();
}

function bloquearEjecutar(valor) {
  $("btn-ejecutar").disabled = valor || trabajando;
}

/* ───────────────────────────── Ejecución ──────────────────────────── */

/* La planilla se abre en el navegador del sistema, no acá adentro: la abre
 * Python (ver Puente.abrir_planilla). No se bloquea mientras se trabaja, porque
 * mirar las respuestas no interfiere con el informe que se está generando. */
async function abrirPlanilla() {
  const respuesta = await pywebview.api.abrir_planilla();
  if (respuesta && respuesta.ok === false) mostrarAviso(false, respuesta.error);
}

function mostrarBotonPlanilla(sheetId) {
  $("btn-abrir-sheet").hidden = !sheetId;
}

async function elegirCsv() {
  const ruta = await pywebview.api.elegir_archivo("csv");
  if (!ruta) return;
  origenCsv = ruta;
  cargarOpciones();
}

function datosDelFormulario(tipo) {
  // Los campos de Servicio/Área quedan OCULTOS en el informe por unidad, pero
  // ocultar no es vaciar: si se configuró una instancia y después se cambió de
  // tipo, seguían viajando y el informe de la unidad entera terminaba titulado
  // con el nombre de un servicio suelto. Sólo se mandan donde se usan.
  const porServicio = tipo === "instancia";
  return {
    tipo: tipo,
    origen: origenCsv ? "csv" : "sheets",
    archivo: origenCsv,
    // Las fechas marcadas: una sola da el informe puntual, varias uno que las
    // junta. El tipo real lo decide `resolver_tipo()` en Python.
    fechas: fechasMarcadas(),
    unidad: $("unidad").value,
    servicios: porServicio ? serviciosMarcados() : [],
    // Si el campo está vacío se calcula acá mismo: así el nombre que se manda no
    // depende de que se haya disparado el evento que actualiza la sugerencia.
    servicio_nombre: porServicio
      ? ($("nombre-servicio").value.trim() || nombreSugerido())
      : "",
    asistentes: parseInt($("asistentes").value, 10) || 0,
    antecedentes: $("antecedentes").value.trim(),
    entrevistas_conf: $("entrevistas").value.trim(),
    con_informe: $("con-informe").checked,
  };
}

async function ejecutar() {
  if (trabajando) return;
  const tipo = tipoElegido();
  const datos = datosDelFormulario(tipo);

  const falta = queFalta(tipo, datos);
  if (falta) {
    mostrarAviso(false, falta);
    return;
  }
  await lanzar(datos);
}

/* Validación en el navegador: mensajes inmediatos, sin ir y volver a Python. */
function queFalta(tipo, datos) {
  if (!opciones) return "Todavía no se pudieron leer los datos. Probá con «Actualizar».";
  if (!datos.unidad) return "Elegí la Unidad Ejecutora.";
  if (tipo === "instancia" && !datos.fechas.length) {
    return "Marcá la fecha de la instancia (o varias, si se dictó en más de un día).";
  }
  return null;
}

async function ejecutarDesdeExcel() {
  if (trabajando) return;
  const ruta = await pywebview.api.elegir_archivo("xlsx");
  if (!ruta) return;
  const datos = datosDelFormulario("instancia");
  datos.origen = "xlsx";
  datos.archivo = ruta;
  await lanzar(datos);
}

async function lanzar(datos) {
  marcarTrabajando(true);
  limpiarRegistro();
  const respuesta = await pywebview.api.ejecutar(datos);
  if (respuesta && respuesta.ok === false) {
    marcarTrabajando(false);
    mostrarAviso(false, respuesta.error);
    if (respuesta.abrir_config) abrirConfiguracion();
  }
}

function marcarTrabajando(valor) {
  trabajando = valor;
  $("btn-ejecutar").disabled = valor;
  $("txt-ejecutar").textContent = valor ? "TRABAJANDO…" : "EJECUTAR";
  $("hilando").hidden = !valor;
  for (const id of ["btn-xlsx", "btn-csv", "btn-recargar"]) $(id).disabled = valor;
}

/* Python nos avisa cada paso del trabajo. */
window.agregarLog = function (texto) {
  const registro = $("registro");
  registro.textContent += texto + "\n";
  registro.scrollTop = registro.scrollHeight;
};

window.terminar = function (ok, mensaje) {
  marcarTrabajando(false);
  mostrarAviso(ok, mensaje);
};

function limpiarRegistro() {
  $("registro").textContent = "";
  $("aviso").hidden = true;
}

function mostrarAviso(ok, mensaje) {
  const aviso = $("aviso");
  aviso.className = "aviso " + (ok ? "ok" : "error");
  aviso.querySelector("use").setAttribute("href", ok ? "#i-check" : "#i-alerta");
  $("aviso-texto").textContent = mensaje;
  aviso.hidden = false;
}

/* ────────────────────── Entrevistas individuales ──────────────────── */

/* Quiénes pidieron una entrevista individual, en dos pestañas: las que faltan
 * hacer y las ya hechas (de donde se pueden recuperar con «Deshacer»).
 *
 * Las dos listas se guardan acá y al marcar o deshacer se mueve el elemento de
 * una a la otra, sin volver a leer la planilla: releerla en cada clic haría
 * esperar un segundo cada vez. La planilla se lee al abrir la ventana, así
 * aparecen los pedidos nuevos sin reiniciar el programa. */
let entrevistas = { pendientes: [], hechas: [] };
let pestanaEntrevistas = "pendientes";

async function abrirEntrevistas() {
  $("modal-entrevistas").hidden = false;
  $("entrevistas-filas").innerHTML = "";
  $("entrevistas-vacio").hidden = true;
  entrevistas = { pendientes: [], hechas: [] };
  mostrarPestanaEntrevistas("pendientes");
  $("entrevistas-resumen").textContent = origenCsv
    ? "Leyendo el archivo…"
    : "Leyendo la planilla de Google…";
  $("btn-cerrar-entrevistas").focus();

  const respuesta = await pywebview.api.entrevistas_pendientes(origenCsv);
  if ($("modal-entrevistas").hidden) return;   // se cerró mientras leía la planilla
  if (!respuesta.ok) {
    $("entrevistas-resumen").textContent = "⚠ " + respuesta.error;
    return;
  }
  entrevistas = { pendientes: respuesta.pendientes, hechas: respuesta.hechas || [] };
  dibujarEntrevistas();
}

function cerrarEntrevistas() {
  $("modal-entrevistas").hidden = true;
}

function mostrarPestanaEntrevistas(cual) {
  pestanaEntrevistas = cual;
  $("tab-pendientes").classList.toggle("activa", cual === "pendientes");
  $("tab-hechas").classList.toggle("activa", cual === "hechas");
  if (entrevistas.pendientes.length || entrevistas.hechas.length) dibujarEntrevistas();
}

function dibujarEntrevistas() {
  const hechas = pestanaEntrevistas === "hechas";
  const lista = hechas ? entrevistas.hechas : entrevistas.pendientes;
  const cuerpo = $("entrevistas-filas");
  cuerpo.innerHTML = "";
  for (const p of lista) cuerpo.appendChild(filaEntrevista(p, hechas));

  $("tab-pendientes").textContent = `Pendientes (${entrevistas.pendientes.length})`;
  $("tab-hechas").textContent = `Completadas (${entrevistas.hechas.length})`;

  const vacio = $("entrevistas-vacio");
  vacio.hidden = lista.length > 0;
  vacio.textContent = hechas
    ? "Todavía no marcaste ninguna entrevista como completada."
    : (entrevistas.hechas.length
        ? "No queda ninguna entrevista pendiente."
        : "Nadie pidió una entrevista individual en la planilla.");
  $("entrevistas-resumen").textContent =
    `${entrevistas.pendientes.length} pendiente(s) · ${entrevistas.hechas.length} completada(s)`;
}

function filaEntrevista(p, hecha) {
  const fila = document.createElement("tr");
  for (const [valor, clase] of [[p.unidad, "celda-unidad"], [p.fecha, "celda-fecha"],
                                [p.nombre, "celda-nombre"], [p.contacto, "celda-contacto"]]) {
    const celda = document.createElement("td");
    celda.className = clase;
    celda.textContent = valor || "—";
    fila.appendChild(celda);
  }
  if (hecha && p.marcada_el) fila.title = `Marcada como completada el ${p.marcada_el}`;

  const celdaBoton = document.createElement("td");
  const boton = document.createElement("button");
  boton.type = "button";
  boton.className = "btn-mini btn-fila";
  boton.title = hecha
    ? "Volver a ponerla como pendiente"
    : "Marcar esta entrevista como ya realizada";
  boton.innerHTML = hecha
    ? '<svg class="ico"><use href="#i-deshacer"></use></svg><span>Deshacer</span>'
    : '<svg class="ico"><use href="#i-listo"></use></svg><span>Completada</span>';
  boton.addEventListener("click", () => cambiarEstadoEntrevista(p, !hecha, boton));
  celdaBoton.appendChild(boton);
  fila.appendChild(celdaBoton);
  return fila;
}

async function cambiarEstadoEntrevista(pedido, hecha, boton) {
  boton.disabled = true;
  const respuesta = await pywebview.api.marcar_entrevista(pedido.clave, hecha);
  if (!respuesta.ok) {
    boton.disabled = false;
    $("entrevistas-resumen").textContent = "⚠ " + respuesta.error;
    return;
  }
  // Las listas se tocan recién cuando Python confirmó que lo pudo guardar: si
  // la fila se moviera antes, un cambio no guardado se vería como aplicado.
  const desde = hecha ? "pendientes" : "hechas";
  const hacia = hecha ? "hechas" : "pendientes";
  entrevistas[desde] = entrevistas[desde].filter((p) => p.clave !== pedido.clave);
  const movido = { ...pedido };
  if (hecha) movido.marcada_el = "recién";     // lo exacto lo pone Python al releer
  else delete movido.marcada_el;
  entrevistas[hacia].unshift(movido);
  dibujarEntrevistas();
}

/* ─────────────────────────── Configuración ────────────────────────── */

function volcarAjustes(estado) {
  const a = estado.ajustes;
  $("cfg-sheet").value = a.sheet_id || "";
  $("cfg-hoja").value = a.worksheet || "";
  $("cfg-clave").value = a.api_key || "";
  $("cfg-modelo").value = a.modelo || "";
  // Sin ID cargado no hay planilla que abrir, así que el botón no se muestra.
  mostrarBotonPlanilla(a.sheet_id);

  const caja = $("estado-cred");
  caja.className = "estado-credenciales " + (estado.credenciales_ok ? "ok" : "falta");
  caja.textContent = estado.credenciales_ok
    ? "✔ Archivo credenciales.json encontrado (acceso a Google)."
    : "✖ Falta el archivo credenciales.json. Copialo en esta carpeta:\n"
      + (estado.carpeta_credenciales || "")
      + "\nSin él no se puede leer la planilla de Google (igual podés usar el "
      + "botón «Usar archivo descargado…»).";
}

async function abrirConfiguracion() {
  // El estado se vuelve a consultar cada vez, porque credenciales.json se puede
  // haber copiado a la carpeta con el programa ya abierto. Antes se leía una
  // sola vez al arrancar y seguía diciendo que faltaba aunque ya estuviera.
  try {
    volcarAjustes(await pywebview.api.estado_inicial());
  } catch (e) {
    // si no se pudo consultar, queda lo último que se sabía
  }
  $("modal").hidden = false;
  $("cfg-sheet").focus();
}

function cerrarConfiguracion() {
  $("modal").hidden = true;
  $("cfg-clave").type = "password";
  $("btn-ver-clave").textContent = "Ver";
}

async function guardarConfiguracion() {
  const ajustes = {
    sheet_id: $("cfg-sheet").value.trim(),
    worksheet: $("cfg-hoja").value.trim(),
    api_key: $("cfg-clave").value.trim(),
    modelo: $("cfg-modelo").value.trim(),
  };
  await pywebview.api.guardar_ajustes(ajustes);
  mostrarBotonPlanilla(ajustes.sheet_id);   // aparece apenas se carga el ID
  cerrarConfiguracion();
  window.agregarLog("Configuración guardada.");
  if (ajustes.sheet_id && !origenCsv) cargarOpciones();
}

function alternarClave() {
  const campo = $("cfg-clave");
  const oculta = campo.type === "password";
  campo.type = oculta ? "text" : "password";
  $("btn-ver-clave").textContent = oculta ? "Ocultar" : "Ver";
}
