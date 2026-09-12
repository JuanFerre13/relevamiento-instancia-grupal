# -*- coding: utf-8 -*-
r"""
Genera el corpus FICTICIO de informes de ejemplo en `recursos\informes_ejemplo\`.

Esos informes son la parte few-shot del prompt: `generar_informe.py` se los pasa
al modelo como ejemplo de estilo, estructura, tono y extensión. Sin ellos el
informe sale sin forma, por más que la guía de estilo la describa.

En una instalación real la carpeta se llena con informes propios de la
organización. Los de este repositorio están **inventados**: las unidades, los
servicios, las fechas, los porcentajes y los testimonios no corresponden a
ninguna organización ni persona real.

Los cuatro informes se escriben una sola vez, en Markdown, y de ahí salen los
tres formatos que `cargar_ejemplos()` sabe leer:

  * dos `.md`   — se leen con `read_text`; además GitHub los muestra formateados.
  * uno `.docx` — con una tabla de dinámica de columnas, que es lo que ejercita
                  el recorrido elemento por elemento de `_texto_de_docx()`.
  * uno `.odt`  — con una nota al pie, que es lo que ejercita el descarte de
                  `text:note` de `_texto_de_odt()`.

Uso:  python desarrollo\crear_ejemplos.py [--sobrescribir]
"""

import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

RAIZ = Path(__file__).resolve().parent.parent
DESTINO = RAIZ / "recursos" / "informes_ejemplo"


# --------------------------------------------------------------------------
# Los informes. Markdown como fuente única: `#` encabezado, `##` sección,
# `- ` viñeta, `| a | b |` tabla.
# --------------------------------------------------------------------------

INFORME_1 = """\
# INFORME TÉCNICO
# HOSPITAL CENTRAL – BLOCK QUIRÚRGICO

## Antecedentes

A petición de la Dirección del Hospital Central y de su Equipo de Gestión se
inicia un proceso por parte de esta Unidad para abordar una situación planteada
en el Block Quirúrgico. Con fecha 4 de marzo del corriente se mantuvo una
reunión virtual con ambos equipos a efectos de conocer el contexto y solicitar
la información pertinente mediante el Formulario de solicitud de intervención,
de modo de determinar el tipo de abordaje a implementar.

De la información recibida se desprende la existencia de dificultades de
relacionamiento entre los funcionarios del Servicio, con especial impacto en la
coordinación de las agendas quirúrgicas. La gestión refiere que las medidas
internas adoptadas fueron la designación de una referente de enfermería con
formación específica y la realización de reuniones periódicas con el personal.
Asimismo, entiende que la situación afecta la calidad del desempeño del equipo
de trabajo.

## Metodología de trabajo

Analizada la información relevada se define aplicar un Abordaje de
*Intervención* Grupal, consistente en la realización de un Taller presencial
denominado "Cómo mejorar la comunicación en el equipo de trabajo", con énfasis
en los diversos estilos de comunicación y su impacto relacional.

El día 19 de marzo, a la hora 13, en el Hospital Central, se realiza la
instancia propuesta, a la cual asisten 24 funcionarios del referido Servicio. En
dicha oportunidad se aplica el *Formulario de Relevamiento de Instancia Grupal*,
de carácter anónimo, método amplio de tipo encuesta para medir a nivel de grupo
las siguientes variables y sus respectivos indicadores: *Comunicación* (escucha
e información), *Relacionamiento Interpersonal* (respeto, confianza), *Trabajo
en equipo* (colaboración, apoyo, cumplimiento de roles) y *Motivación*
(reconocimiento y feedback).

Se recibieron 21 formularios, de los cuales 8 corresponden al personal médico y
13 al personal no médico. En el marco de la instancia se solicitaron 2
entrevistas individuales confidenciales de seguimiento.

Entre los comentarios anónimos de valoración de la actividad se transcriben:

- "Sirvió para escucharnos entre turnos, que es algo que no hacemos nunca."
- "Estuvo bien, pero el problema de fondo es la falta de personal."
- "Ojalá se repita y participe también la jefatura."

## Análisis

De la información relevada del personal del Block Quirúrgico en el Taller,
surgen como más relevantes los siguientes datos:

**A- Comunicación.** El 71 % del personal considera que la comunicación con sus
compañeros de sector es buena y el 67 % valora positivamente la comunicación con
su jefe inmediato. Sin embargo, la comunicación con el Equipo de Gestión reúne
un 62 % de percepción negativa o neutra, el porcentaje más alto de esta
dimensión. Asimismo, un 48 % manifiesta no sentirse debidamente informado sobre
las tareas y responsabilidades necesarias para el desempeño de sus funciones, lo
que sugiere que la información circula de manera desigual entre los turnos.

**B- Relaciones interpersonales.** Es la dimensión mejor evaluada. El 86 %
señala que sus compañeros lo tratan con respeto y el 81 % considera que se lleva
bien con ellos. No obstante, la confianza entre pares desciende al 57 %, brecha
que suele indicar que la convivencia cotidiana se sostiene sin que exista
todavía un vínculo de trabajo consolidado.

**C- Trabajo en Equipo.** El 67 % entiende que puede contar con sus compañeros
y/o con el jefe del sector ante un problema. En contrapartida, un 52 % consigna
una percepción negativa o neutra respecto de la definición de tareas y roles, y
el mismo porcentaje respecto del incentivo al trabajo en equipo. La conjunción
de ambos datos orienta hacia una dificultad de organización del trabajo antes
que hacia un problema vincular.

**D- Motivación.** Es la dimensión más comprometida. El 62 % no se siente
reconocido ni valorado por sus jefes en el desempeño de sus funciones y un 57 %
consigna que no recibe orientación sobre cómo se percibe su desempeño ni pautas
para mejorar. La satisfacción con las propias tareas, en cambio, alcanza el
67 %, lo que permite suponer que el malestar no se dirige al contenido del
trabajo sino a su reconocimiento.

Respecto a la valoración de la actividad llevada a cabo por esta Unidad, un
90 % manifiesta conformidad.

Es dable señalar que se solicitaron 2 entrevistas Individuales Confidenciales,
de las cuales surge: la percepción de que las decisiones sobre la distribución
de tareas no siguen criterios explícitos, y la existencia de situaciones
puntuales de trato descortés que no fueron canalizadas por ninguna vía formal.
En ambos casos se trata de aspectos que las personas entrevistadas vinculan al
desgaste acumulado del Servicio más que a conductas deliberadas.

## Conclusiones y recomendaciones

Los datos relevados permiten identificar causas subyacentes y factores
desencadenantes de eventuales conflictos: una circulación irregular de la
información entre turnos, una definición difusa de tareas y roles y una práctica
de reconocimiento escasa o poco visible. Cabe recordar que los conflictos no
resueltos son una fuente importante de estrés en el lugar de trabajo y que su
resolución efectiva mejora el bienestar general del equipo.

## Comunicación

**Comunicación Formal:** a) establecer circulares y/o memorándums para las
definiciones que alcanzan a todo el Servicio, de modo que la información no
dependa del turno en que se esté; b) sostener reuniones periódicas con mandos
medios y/o referentes de cada turno, con temario y registro breve de acuerdos.

**Comunicación Informal:** promover prácticas de apertura al inicio de cada
jornada —saludo, repaso de la agenda quirúrgica— y sostener el trato asertivo y
respetuoso en un marco de asertividad y respeto.

## Relacionamiento interpersonal

Podría valorarse la realización de instancias breves de encuentro entre turnos,
que habiliten el conocimiento mutuo del personal que hoy no coincide en el
horario de trabajo, así como espacios de intercambio sobre la tarea que
permitan construir confianza a partir de lo concreto.

## Motivación y trabajo en equipo

Se sugiere avanzar en la definición escrita de tareas y roles del Servicio
mediante protocolos, flujogramas y/o procedimientos, de modo que la
distribución del trabajo responda a criterios objetivos y verificables.
Asimismo, se recomienda incorporar prácticas sistemáticas de feedback, con
instancias pautadas en las que cada funcionario reciba una devolución concreta
sobre su desempeño.

En suma, entendemos que el fortalecimiento de los canales formales de
comunicación y la explicitación de los criterios de organización del trabajo
permitirían reducir la tensión relevada y mejorar el clima del Servicio. Le
proponemos al Equipo de Gestión un acompañamiento inicial para definir
estrategias para llevar a cabo las recomendaciones sugeridas en este Informe.

Montevideo, 27 de marzo de 2026.
"""


INFORME_2 = """\
# INFORME TÉCNICO
# HOSPITAL DEL NORTE – FARMACIA

## Antecedentes

A solicitud del Equipo de Gestión, la Unidad se puso en contacto a efectos de
coordinar el tipo de abordaje a implementar de modo de colaborar para mejorar el
relacionamiento interpersonal en el servicio de Farmacia del Hospital del Norte.
De la información aportada mediante el Formulario de solicitud de intervención
surge la preocupación por aspectos vinculares que afectan el trabajo en equipo,
particularmente en lo relativo a la distribución de las tareas de dispensación.

## Metodología de trabajo

Se define aplicar un Abordaje de *Sensibilización*, consistente en un Taller
presencial denominado "El trabajo en equipo como práctica cotidiana".

El día 8 de abril, a la hora 11, en el Hospital del Norte, se realiza la
instancia propuesta. La plantilla del Servicio es de 19 funcionarios; asisten 6
y se reciben 5 formularios, de los cuales 1 corresponde al personal médico y 4
al personal no médico. Se aplica el *Formulario de Relevamiento de Instancia
Grupal*, de carácter anónimo, método amplio de tipo encuesta para medir a nivel
de grupo las siguientes variables y sus respectivos indicadores: *Comunicación*
(escucha e información), *Relacionamiento Interpersonal* (respeto, confianza),
*Trabajo en equipo* (colaboración, apoyo, cumplimiento de roles) y *Motivación*
(reconocimiento y feedback).

Atento a la cantidad de formularios recibidos, se dejaron formularios en blanco
a disposición del personal que no concurrió, sin que se recibieran respuestas
adicionales dentro del plazo previsto.

## Análisis

De la información relevada del personal de Farmacia en el Taller, surgen como
más relevantes los siguientes datos. Corresponde señalar con carácter previo que
la información **no es representativa, dado que no hubo adherencia a la
actividad**: los 5 formularios recibidos equivalen a poco más de la cuarta parte
de la plantilla del Servicio, por lo que el análisis que sigue se limita a
consignar lo relevado, sin extrapolar al conjunto.

**A- Comunicación.** De los 5 formularios recibidos, 4 consignan que la
comunicación con el jefe inmediato es buena y 3 que lo es con los compañeros de
sector. La comunicación con el Equipo de Gestión reúne 3 respuestas de
percepción negativa o neutra.

**B- Relaciones interpersonales.** La totalidad de los formularios consigna
sentirse tratado con respeto por sus compañeros. La confianza entre pares reúne
2 respuestas neutras.

**C- Trabajo en Equipo.** 3 formularios señalan que en el sector están bien
definidas las tareas y roles de cada uno; los 2 restantes consignan una
percepción neutra. Ante un problema, 4 entienden que pueden contar con
compañeros y/o con el jefe del sector.

**D- Motivación.** 3 formularios consignan sentirse motivados y satisfechos con
el desempeño de sus tareas. El reconocimiento por parte de las jefaturas reúne 3
respuestas de percepción negativa o neutra, y la orientación sobre el propio
desempeño, 4.

Respecto a la valoración de la actividad llevada a cabo por esta Unidad, la
totalidad de quienes respondieron manifiesta conformidad.

## Conclusiones y recomendaciones

La baja adherencia a la instancia constituye, en sí misma, el dato más
relevante del relevamiento, y podría estar vinculada a la organización de los
turnos, a la carga de trabajo del Servicio o al grado de expectativa depositado
en este tipo de actividades. No es posible, con la información disponible,
establecer cuál de esos factores predomina. Se recuerda que los conflictos no
resueltos son una fuente importante de estrés en el lugar de trabajo y que su
resolución efectiva mejora el bienestar general.

## Comunicación

**Comunicación Formal:** a) comunicar con antelación suficiente la realización
de este tipo de instancias, precisando su carácter anónimo y su finalidad; b)
coordinar la actividad con la organización de los turnos, de modo que la
asistencia no dependa de la disponibilidad individual.

**Comunicación Informal:** sostener prácticas de apertura y trato asertivo y
respetuoso en el intercambio cotidiano del Servicio.

## Relacionamiento interpersonal

Se sugiere valorar la realización de una nueva instancia, con convocatoria
ampliada, que permita relevar la percepción del conjunto del Servicio antes de
definir acciones específicas.

## Motivación y trabajo en equipo

Podría valorarse la incorporación de prácticas sistemáticas de feedback, dado
que el reconocimiento y la orientación sobre el desempeño concentran, aun en
este número reducido de respuestas, la mayor proporción de percepción negativa o
neutra.

En suma, entendemos que una nueva convocatoria en condiciones que favorezcan la
participación permitiría contar con información representativa sobre la cual
definir acciones. La Unidad queda a disposición del Equipo de Gestión para
definir estrategias que permitan llevar a cabo las recomendaciones sugeridas en
este Informe.

Montevideo, 15 de abril de 2026.
"""


INFORME_3 = """\
# INFORME TÉCNICO
# CENTRO DEPARTAMENTAL ESTE – LABORATORIO

## Antecedentes

A petición de la Dirección del Centro Departamental Este se inicia un proceso
por parte de esta Unidad para abordar una situación planteada en el servicio de
Laboratorio. Con fecha 12 de mayo del corriente se mantuvo una reunión con el
Equipo de Gestión a efectos de conocer el contexto y solicitar la información
pertinente mediante el Formulario de solicitud de intervención.

De la información recibida surge la existencia de dificultades en el
cumplimiento de los roles definidos y en la coordinación entre el turno matutino
y el vespertino, que la gestión consigna de alta afectación en los tiempos de
respuesta del Servicio.

## Metodología de trabajo

Analizada la información relevada se define aplicar un Abordaje de
*Intervención* Grupal, consistente en la realización de un Taller presencial
denominado "Roles, tareas y acuerdos de trabajo".

El día 28 de mayo, a la hora 14, en el Centro Departamental Este, se realiza la
instancia propuesta, a la cual asisten 17 funcionarios del referido Servicio. Se
recibieron 16 formularios, de los cuales 3 corresponden al personal médico y 13
al personal no médico. Se aplica el *Formulario de Relevamiento de Instancia
Grupal*, de carácter anónimo, método amplio de tipo encuesta para medir a nivel
de grupo las siguientes variables y sus respectivos indicadores: *Comunicación*
(escucha e información), *Relacionamiento Interpersonal* (respeto, confianza),
*Trabajo en equipo* (colaboración, apoyo, cumplimiento de roles) y *Motivación*
(reconocimiento y feedback).

En el marco del Taller se aplicó una dinámica de trabajo en subgrupos, cuyos
emergentes se consignan en el Análisis.

## Análisis

De la información relevada del personal del Laboratorio en el Taller, surgen
como más relevantes los siguientes datos:

**A- Comunicación.** El 75 % considera buena la comunicación con sus compañeros
de sector y el 69 % con su jefe inmediato. Sin embargo, un 56 % consigna una
percepción negativa o neutra respecto de la escucha de sus ideas, aportes y
sugerencias para mejorar el trabajo.

**B- Relaciones interpersonales.** El 88 % señala que sus compañeros lo tratan
con respeto y el 81 % que se lleva bien con ellos. El trato respetuoso por parte
de la jefatura del sector alcanza el 69 %.

**C- Trabajo en Equipo.** Es la dimensión que concentra las mayores
dificultades. Un 63 % consigna una percepción negativa o neutra respecto de la
definición de tareas y roles, y un 56 % respecto del incentivo al trabajo en
equipo. No obstante, el 75 % entiende que puede contar con sus compañeros y/o
con el jefe del sector cuando tiene un problema, lo que sugiere que la
solidaridad cotidiana está compensando una organización del trabajo poco
explícita.

**D- Motivación.** El 69 % se siente motivado y satisfecho con el desempeño de
sus tareas. En cambio, un 56 % consigna no sentirse reconocido ni valorado por
sus jefes y un 50 % no recibir orientación sobre cómo se percibe su desempeño.

De la dinámica aplicada en subgrupos surgen los siguientes emergentes:

| ¿Qué voy a hacer más? | ¿Qué voy a hacer diferente? | ¿Qué voy a dejar de hacer? |
| Dejar registro escrito del pase de turno | Plantear las diferencias en el momento y no después | Dar por sobreentendido que el otro turno sabe |
| Consultar antes de reasignar una muestra | Usar el cuaderno de novedades para todo lo que no es urgente | Resolver por mensajería informal lo que corresponde al Servicio |
| Reconocer el trabajo del compañero | Pedir ayuda antes de que el problema escale | Comentar con terceros lo que no se habló con el interesado |

Asimismo, los subgrupos identificaron como causas de las dificultades
relevadas: 1) **la superposición de tareas** entre turnos (particularmente en la
recepción y el procesamiento de muestras); 2) **la ausencia de registro** de lo
resuelto en cada jornada (lo que obliga a reconstruir verbalmente el estado del
Servicio); y 3) **la falta de un ámbito** donde plantear las diferencias antes
de que se instalen.

Respecto a la valoración de la actividad llevada a cabo por esta Unidad, un
94 % manifiesta conformidad.

## Conclusiones y recomendaciones

Los datos relevados y los emergentes de la dinámica convergen en un mismo punto:
las dificultades del Servicio se vinculan a causas subyacentes de orden
organizativo —superposición de tareas, ausencia de registro, falta de un ámbito
de planteo— antes que a un deterioro del vínculo entre las personas, que se
mantiene en niveles altos. Los conflictos no resueltos son una fuente importante
de estrés en el lugar de trabajo y su resolución efectiva mejora el bienestar
general.

## Comunicación

**Comunicación Formal:** a) instituir un registro escrito del pase de turno, con
formato breve y de uso obligatorio; b) sostener reuniones periódicas con
referentes de ambos turnos, con temario previo y registro de acuerdos.

**Comunicación Informal:** promover el planteo directo y oportuno de las
diferencias, en un marco de asertividad y respeto, evitando que las situaciones
se procesen por fuera del Servicio.

## Relacionamiento interpersonal

Se sugiere sostener los acuerdos surgidos de la dinámica mediante una instancia
de seguimiento a los tres meses, que permita revisar qué se incorporó
efectivamente a la práctica cotidiana.

## Motivación y trabajo en equipo

Se recomienda avanzar en protocolos, flujogramas y/o procedimientos que definan
con criterios objetivos el alcance de cada rol en la recepción y el
procesamiento de muestras, así como incorporar prácticas sistemáticas de
feedback sobre el desempeño.

En suma, entendemos que la explicitación de los acuerdos de trabajo y la
incorporación de un registro compartido permitirían sostener la buena base
vincular del Servicio y mejorar sus tiempos de respuesta. Le proponemos al
Equipo de Gestión un acompañamiento inicial para definir estrategias para llevar
a cabo las recomendaciones sugeridas en este Informe.

Montevideo, 5 de junio de 2026.
"""


INFORME_4 = """\
# INFORME TÉCNICO – UNIDAD EJECUTORA
# HOSPITAL CENTRAL

## Antecedentes

A solicitud de la Dirección del Hospital Central, esta Unidad desarrolló durante
el primer semestre del corriente un conjunto de instancias de relevamiento en
distintos servicios de la Unidad Ejecutora, en el marco de la política
institucional de gestión de conflictos vigente en la organización. El presente
Informe reúne lo relevado en ese período.

## Metodología de trabajo

Las actividades se desarrollaron entre los meses de marzo y junio y comprendieron
los servicios de Block Quirúrgico, Emergencia y Laboratorio. En cada uno se
aplicó el *Formulario de Relevamiento de Instancia Grupal*, de carácter anónimo,
método amplio de tipo encuesta para medir a nivel de grupo las siguientes
variables y sus respectivos indicadores: *Comunicación* (escucha e información),
*Relacionamiento Interpersonal* (respeto, confianza), *Trabajo en equipo*
(colaboración, apoyo, cumplimiento de roles) y *Motivación* (reconocimiento y
feedback).

Se recibieron en total 63 formularios, de los cuales 19 corresponden al personal
médico y 44 al personal no médico. Corresponde advertir que los datos que siguen
corresponden al conjunto de las instancias relevadas en el período, por lo que
expresan una tendencia general de la Unidad Ejecutora y no la situación de un
servicio ni de una jornada en particular.

## Análisis

De la información relevada del personal del Hospital Central en las instancias
del período, surgen como más relevantes los siguientes datos:

**A- Comunicación.** El 73 % considera buena la comunicación con sus compañeros
de sector y el 68 % con su jefe inmediato. La comunicación con el Equipo de
Gestión, en cambio, reúne un 59 % de percepción negativa o neutra, y un 44 %
manifiesta no sentirse debidamente informado sobre las tareas y
responsabilidades necesarias para el desempeño de sus funciones.

**B- Relaciones interpersonales.** Es la dimensión mejor evaluada de la Unidad
Ejecutora. El 84 % consigna ser tratado con respeto por sus compañeros y el 79 %
llevarse bien con ellos. La confianza entre pares se ubica en el 62 %.

**C- Trabajo en Equipo.** El 70 % entiende que puede contar con sus compañeros
y/o con el jefe del sector ante un problema. Sin embargo, un 57 % consigna una
percepción negativa o neutra respecto de la definición de tareas y roles,
porcentaje que se repite con escasa variación en los tres servicios relevados.

**D- Motivación.** Es la dimensión más comprometida del conjunto. Un 59 % no se
siente reconocido ni valorado por sus jefes y un 54 % no recibe orientación
sobre cómo se percibe su desempeño. La satisfacción con las propias tareas, en
cambio, alcanza el 68 %.

Respecto a la valoración de la actividad llevada a cabo por esta Unidad, un
91 % manifiesta conformidad.

## Conclusiones y recomendaciones

Leído en conjunto, el relevamiento describe una Unidad Ejecutora con vínculos
interpersonales sólidos y una organización del trabajo poco explícita. Las
causas subyacentes y factores desencadenantes de eventuales conflictos se
concentran, de manera consistente entre servicios, en la definición de tareas y
roles y en las prácticas de reconocimiento, antes que en el trato cotidiano
entre las personas. Cabe recordar que los conflictos no resueltos son una fuente
importante de estrés en el lugar de trabajo y que su resolución efectiva mejora
el bienestar general.

## Comunicación

**Comunicación Formal:** a) establecer un canal institucional regular entre el
Equipo de Gestión y los servicios, que no dependa de la iniciativa de cada
jefatura; b) sostener reuniones periódicas con mandos medios y/o referentes, con
temario previo y registro de acuerdos.

**Comunicación Informal:** promover prácticas de apertura, saludo y trato
asertivo y respetuoso como criterio transversal a la Unidad Ejecutora.

## Relacionamiento interpersonal

Dado que es la dimensión mejor evaluada, se sugiere sostener las prácticas
existentes y utilizarlas como punto de apoyo para las acciones de organización
del trabajo, cuyo impacto suele ser mayor cuando la base vincular es favorable.

## Motivación y trabajo en equipo

Se recomienda avanzar, a nivel de la Unidad Ejecutora, en criterios objetivos y
comunes para la definición de tareas y roles mediante protocolos, flujogramas
y/o procedimientos, así como incorporar prácticas sistemáticas de feedback que
alcancen a todos los servicios y no dependan del estilo de cada jefatura.

En suma, entendemos que un abordaje de alcance institucional sobre la
organización del trabajo y el reconocimiento permitiría capitalizar la buena
base vincular relevada. La Unidad queda a disposición del Equipo de Gestión para
definir estrategias que permitan llevar a cabo las recomendaciones sugeridas en
este Informe.

Montevideo, 30 de junio de 2026.
"""


LEEME = """\
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
"""


# --------------------------------------------------------------------------
# Conversores. El Markdown de arriba es la fuente; de acá salen los formatos.
# --------------------------------------------------------------------------

def bloques(markdown):
    """Markdown → lista de (clase, contenido).

    Clases: 'titulo' (`#`), 'seccion' (`##`), 'vineta' (`- `), 'tabla' (lista de
    filas, cada una lista de celdas) y 'parrafo'. Los párrafos vienen con los
    saltos de línea internos aplastados: en el .docx y el .odt el ancho lo
    resuelve el procesador de texto, no el archivo."""
    salida = []
    for crudo in markdown.split("\n\n"):
        crudo = crudo.strip()
        if not crudo:
            continue
        lineas = crudo.split("\n")
        if lineas[0].startswith("| "):
            filas = [[c.strip() for c in ln.strip().strip("|").split("|")] for ln in lineas]
            salida.append(("tabla", filas))
        elif lineas[0].startswith("- "):
            for ln in lineas:
                salida.append(("vineta", ln[2:].strip()))
        elif crudo.startswith("## "):
            salida.append(("seccion", crudo[3:].strip()))
        elif crudo.startswith("# "):
            for ln in lineas:
                salida.append(("titulo", ln.lstrip("# ").strip()))
        else:
            salida.append(("parrafo", " ".join(x.strip() for x in lineas)))
    return salida


def _sin_marcas(texto):
    """Saca los `**` y `*` del Markdown: en .docx y .odt el ejemplo va en texto
    plano, que es como el modelo lo va a recibir de todos modos (los lectores
    devuelven `paragraph.text`, sin formato)."""
    return texto.replace("**", "").replace("*", "")


def escribir_docx(markdown, destino):
    """Informe de ejemplo en Word, con su tabla en el lugar que le toca."""
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt

    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)

    for clase, contenido in bloques(markdown):
        if clase == "tabla":
            tabla = doc.add_table(rows=len(contenido), cols=len(contenido[0]))
            tabla.style = "Table Grid"
            for i, fila in enumerate(contenido):
                for j, celda in enumerate(fila):
                    tabla.cell(i, j).text = _sin_marcas(celda)
            continue
        p = doc.add_paragraph()
        corrida = p.add_run(_sin_marcas(contenido))
        if clase == "titulo":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            corrida.bold = corrida.underline = True
        elif clase == "seccion":
            corrida.bold = corrida.underline = True
        elif clase == "vineta":
            p.paragraph_format.left_indent = Pt(21)
            corrida.text = "• " + corrida.text
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    doc.save(str(destino))


# Un .odt es un ZIP con un `content.xml` adentro, y `_texto_de_odt()` sólo lee
# ese archivo. Igual se escriben `styles.xml`, `meta.xml` y el manifiesto para
# que el documento también abra en LibreOffice.
NS_ODT = (
    'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
    'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
    'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
    'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
    'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"'
)

MANIFIESTO = """<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
 <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
 <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
"""


def escribir_odt(markdown, destino, nota_al_pie=None):
    """Informe de ejemplo en LibreOffice Writer.

    `nota_al_pie` mete un `<text:note>` INLINE en el primer párrafo del
    Análisis, que es donde lo pone la convención de los informes ("Ver Cuadro
    Relevamiento de Intervención Grupal"). Es el caso que ejercita el descarte
    de notas de `_texto_de_odt()`: sin ese descarte, el texto de la nota queda
    pegado en medio de la frase."""
    partes = []
    nota_pendiente = nota_al_pie
    for clase, contenido in bloques(markdown):
        if clase == "tabla":
            filas = []
            for fila in contenido:
                celdas = "".join(
                    f'<table:table-cell office:value-type="string">'
                    f"<text:p>{escape(_sin_marcas(c))}</text:p></table:table-cell>"
                    for c in fila)
                filas.append(f"<table:table-row>{celdas}</table:table-row>")
            partes.append(
                f'<table:table table:name="Dinamica">'
                f'<table:table-column table:number-columns-repeated="{len(contenido[0])}"/>'
                + "".join(filas) + "</table:table>")
            continue

        texto = escape(_sin_marcas(contenido))
        if clase == "titulo":
            partes.append(f'<text:h text:outline-level="1">{texto}</text:h>')
        elif clase == "seccion":
            partes.append(f'<text:h text:outline-level="2">{texto}</text:h>')
        elif clase == "vineta":
            partes.append(f"<text:list><text:list-item><text:p>{texto}"
                          f"</text:p></text:list-item></text:list>")
        else:
            if nota_pendiente and contenido.startswith("De la información relevada"):
                # La nota va antes del ":" final, como en los informes: el texto
                # que queda después de la nota es el `tail` que el lector
                # preserva al sacarla.
                cuerpo, sep, cola = texto.rpartition(":")
                nota = (
                    '<text:note text:note-class="footnote" text:id="ftn1">'
                    '<text:note-citation>1</text:note-citation>'
                    f"<text:note-body><text:p>{escape(nota_pendiente)}</text:p>"
                    "</text:note-body></text:note>")
                partes.append(f"<text:p>{cuerpo}{nota}{sep}{cola}</text:p>")
                nota_pendiente = None
            else:
                partes.append(f"<text:p>{texto}</text:p>")

    contenido_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        f'<office:document-content {NS_ODT} office:version="1.2">'
        "<office:body><office:text>" + "".join(partes) +
        "</office:text></office:body></office:document-content>")

    estilos_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        f'<office:document-styles {NS_ODT} office:version="1.2">'
        "<office:styles/></office:document-styles>")

    meta_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.2">'
        "<office:meta><meta:generator>crear_ejemplos.py</meta:generator></office:meta>"
        "</office:document-meta>")

    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
        # `mimetype` va primero y SIN comprimir: lo pide la especificación ODF.
        z.writestr(zipfile.ZipInfo("mimetype"),
                   "application/vnd.oasis.opendocument.text",
                   compress_type=zipfile.ZIP_STORED)
        z.writestr("META-INF/manifest.xml", MANIFIESTO)
        z.writestr("content.xml", contenido_xml)
        z.writestr("styles.xml", estilos_xml)
        z.writestr("meta.xml", meta_xml)


def main():
    if DESTINO.exists() and any(DESTINO.iterdir()) and "--sobrescribir" not in sys.argv:
        print(f"{DESTINO} ya tiene archivos.")
        print("Pasale --sobrescribir si querés regenerar el corpus de ejemplo.")
        return 1

    DESTINO.mkdir(parents=True, exist_ok=True)

    (DESTINO / "README.md").write_text(LEEME, encoding="utf-8")
    (DESTINO / "01-hospital-central-block-quirurgico.md").write_text(
        INFORME_1, encoding="utf-8")
    (DESTINO / "02-hospital-del-norte-farmacia.md").write_text(
        INFORME_2, encoding="utf-8")
    escribir_docx(INFORME_3, DESTINO / "03-centro-departamental-este-laboratorio.docx")
    escribir_odt(INFORME_4, DESTINO / "04-hospital-central-unidad-ejecutora.odt",
                 nota_al_pie="Ver Cuadro Relevamiento de Instancia Grupal.")

    for archivo in sorted(DESTINO.iterdir()):
        print(f"  {archivo.name}  ({archivo.stat().st_size // 1024 or 1} KB)")
    print(f"Listo: {DESTINO}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
