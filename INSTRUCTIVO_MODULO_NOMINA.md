# Instructivo del Módulo de Nómina y Control de Asistencia

**Sistema de Control de Flota Automotriz - Collie Software**
Versión del módulo: 1.0 · Última actualización: setiembre 2026
Dirigido a: administración, recursos humanos y jefaturas de operaciones.

---

## 1. ¿Para qué sirve este módulo?

El módulo toma el archivo que exporta el **reloj biométrico (captahuella)**, identifica la **primera y la última marcación de cada empleado en cada día**, las cruza con el **turno y el horario asignado** y calcula automáticamente:

- Tardanzas y salidas anticipadas (en minutos y en soles).
- Faltas injustificadas y días con marcación incompleta.
- Horas extra con sus recargos (25% y 35%).
- Descansos, feriados y trabajo en día de descanso.
- Vacaciones, permisos, licencias y descansos médicos.
- La **liquidación de la planilla del mes**: haberes, descuentos, aportes del empleador y neto a pagar.
- Boletas de pago por trabajador y reportes exportables a Excel.

Todo queda guardado en la base de datos del sistema (la misma que usan los demás módulos), por lo que **todos los equipos ven la misma información**.

### Capacidades resumidas

| Área | Qué resuelve |
|---|---|
| Reloj biométrico | Importa el Excel del captahuella con detección automática de formato y columnas |
| Marcaciones | Evita duplicados al reimportar, conserva el histórico y audita cada importación |
| Empleados | Toma el padrón de la tabla de Choferes y agrega los datos de nómina |
| Turnos | Catálogo de turnos con tolerancia, refrigerio, jornada, días y turno nocturno |
| Horarios | Plantilla semanal por día (Lunes a Domingo) reutilizable |
| Asignaciones | Cambios por rango de fechas para turnos rotativos |
| Asistencia | Grilla mensual con código de color por estado y detalle día por día |
| Incidencias | Vacaciones, permisos, licencias y descansos médicos que justifican la falta |
| Horas extra | Cálculo automático, edición manual y aprobación |
| Planilla | Liquidación mensual, boletas, estados (borrador, aprobada, pagada) y exportación |
| Reportes | Puntualidad, tardanzas, faltas, horas extra y bitácora de marcaciones |
| Calendario | Feriados nacionales del Perú y días no laborables |

---

## 2. Cómo entrar al módulo

1. Inicie el sistema e ingrese con su usuario y contraseña.
2. En el menú lateral busque el grupo **FINANZAS Y REPORTES**.
3. Haga clic en **🧾 Nómina y Asistencia**.

El módulo se abre dentro de la misma ventana, como todos los demás.

### Permisos (importante la primera vez)

El permiso del módulo es nuevo, así que **los usuarios que no sean Super Administrador deben recibirlo**:

1. Vaya a **Ajustes de Sistema → Configurar Usuarios**.
2. Seleccione el usuario en la lista.
3. Marque la casilla **🧾 Nómina y Asistencia**.
4. Pulse **Guardar Usuario**.

Si al abrir el módulo aparece el mensaje "No tiene permisos", es exactamente esto.

### Pantalla completa (modo trabajo)

- Botón **⛶ Pantalla completa** en la cabecera del módulo y también en la barra lateral del sistema.
- Al activarla, **los botones del sistema se ocultan**: la barra lateral desaparece y el módulo pasa a ocupar todo el ancho de la pantalla.
- Para volver, tiene tres opciones: el botón flotante **🗗 Salir de pantalla completa** que aparece en la esquina inferior derecha, la tecla **F11** o la tecla **Esc**.
- Al salir, la barra lateral vuelve sola a su lugar y la ventana queda maximizada.
- Es ideal para la grilla mensual de asistencia, que tiene una columna por cada día.

### Aviso mientras el sistema trabaja

Cuando el módulo **lee la base de datos o hace un cálculo**, aparece un **recuadro centrado** con el icono de reloj, el mensaje de lo que está haciendo y una barra de progreso:

- «Leyendo la base de datos…» al abrir el módulo, cargar la asistencia de un mes, ver una boleta o generar un reporte.
- «Calculando la asistencia…» al recalcular el mes, con el avance («Avance: 100 de 150»).
- «Calculando la planilla…» al liquidar el período.
- «Importando las marcaciones…» con el avance fila por fila («Fila 300 de 864»).
- «Preparando la pantalla…» al abrir el módulo por primera vez.

El aviso **desaparece solo** al terminar y la ventana sigue respondiendo mientras tanto: puede moverla o cambiar de tamaño. Nunca queda pegado en pantalla.

### Ayuda: este instructivo dentro del sistema

El botón **❓ Ayuda**, arriba a la derecha junto a Pantalla completa, abre **este mismo instructivo sin salir del sistema**:

**Búsqueda rápida.** Escriba una palabra o frase en el buscador y pulse **Buscar** (o Enter). El sistema recorre todo el instructivo y **resalta en amarillo** las coincidencias dentro de la página; la coincidencia actual se marca en naranja.

- Con **◀** y **▶** salta a la coincidencia anterior o siguiente, recorriendo todo el instructivo.
- **No distingue mayúsculas ni acentos**: escribir *nomina* encuentra *nómina*, y *horas extra* encuentra la frase completa.
- Debajo del buscador indica en qué punto va: «*turno*: coincidencia 3 de 49 · página 4 · 1 en esta página».
- **Limpiar** quita el resaltado y vacía el buscador.
- El atajo **Ctrl+F** lleva el cursor directo al buscador.

Ejemplos de búsquedas útiles: *turno*, *tolerancia*, *horas extra*, *tardanza*, *falta*, *AFP*, *quinta*, *feriado*, *incompleto*, *boleta*, *permiso*.

**Navegación por el documento**

- Página por página, con **◀ Anterior** y **Siguiente ▶** o escribiendo el número de página.
- Acercar y alejar con **➕ ➖**, o usar **Ajustar al ancho**.
- Con el teclado: **← →** cambian de página, **Inicio** y **Fin** van a la primera o a la última, **Esc** cierra la ayuda.
- El botón **🔍 Abrir con el visor del sistema** lo abre en su lector de PDF si prefiere leerlo aparte.

Si el equipo no puede mostrar el PDF dentro del sistema, el botón lo abre directamente con el visor de PDF de Windows o Mac, sin interrumpir su trabajo.

---

## 3. Flujo de trabajo recomendado

Este es el orden correcto la primera vez. Después, el trabajo de cada mes se reduce a los últimos cuatro pasos: importar las marcaciones, vincular las personas nuevas, recalcular la asistencia y calcular la planilla.

| # | Paso | Dónde |
|---|---|---|
| 1 | Revisar que los empleados existan en Control de Choferes | Módulo Choferes |
| 2 | Sincronizar el padrón al módulo de nómina | Personal → Empleados |
| 3 | Crear o ajustar los **turnos** de la empresa | Personal → Turnos |
| 4 | Armar los **horarios** semanales y asignarlos a cada empleado | Personal → Horarios y Empleados |
| 5 | Completar la **ficha de nómina**: sueldo, jornada, pensión, AFP | Personal → Empleados |
| 6 | Cargar los **feriados del año** y revisar los **parámetros** | Configuración |
| 7 | **Importar el Excel del reloj** y vincular las personas que falten | Marcaciones |
| 8 | **Recalcular la asistencia** del mes | Asistencia |
| 9 | Revisar la grilla, corregir lo necesario y registrar incidencias | Asistencia |
| 10 | **Calcular la planilla**, revisar boletas, aprobar y exportar | Planilla |

> Regla de oro: **cada vez que cambie un turno, un horario, un sueldo, un parámetro o una incidencia, vuelva a recalcular la asistencia y luego la planilla.** Los cálculos no se actualizan solos.

### 3.1 El asistente paso a paso (la forma más fácil)

Si no quiere memorizar el orden de los pasos, use el botón **🧭 ASISTENTE PASO A PASO** del Panel. Se abre una ventana que **revisa el estado real del sistema** y le dice exactamente qué falta, sin que usted tenga que buscarlo.

- **Barra de avance**: el porcentaje completado y cuántos pasos críticos están listos.
- **Lista de pasos** (izquierda): cada paso con su estado — ✅ listo, ⚠️ revisar, ⛔ pendiente. Haga clic en cualquiera para verlo.
- **Detalle** (derecha): qué encontró el sistema (con nombres y cantidades reales), por qué importa ese paso y qué debe hacer.
- **➡️ Ir a esa pantalla**: lo lleva directo a la pestaña y a la vista exacta donde se resuelve.
- **⚡ Ejecutar**: en los pasos que se pueden automatizar, hace la tarea desde el mismo asistente: crear las fichas de nómina que faltan, cargar los feriados del año, reparar parámetros dañados, recalcular la asistencia o calcular la planilla.
- **🔄 Verificar de nuevo**: vuelve a revisar todo después de cada cambio.

Los pasos están agrupados en dos fases:

**Puesta en marcha** (una sola vez)

1. Padrón de empleados
2. Turnos de trabajo
3. Horarios semanales
4. Asignar horario a cada empleado
5. Sueldos y datos de pago
6. Parámetros de cálculo
7. Feriados del año

**Cierre del mes** (cada mes)

8. Importar las marcaciones del mes
9. Vincular el reloj con los DNI (las personas nuevas que aparezcan en el reloj)
10. Recalcular la asistencia
11. Calcular la planilla

Las pestañas del módulo están ordenadas **exactamente igual** que estos pasos, así que basta con avanzar de izquierda a derecha.

Ejemplos reales de lo que le dirá el asistente: «45% completado, 4 de 8 pasos críticos listos», «5 empleados sin sueldo básico: Ángel, Eduardo, Isaac y 2 más», «hay 23 marcas en Setiembre 2026, pero el archivo cargado llega solo hasta el 09/09/2026».

El Panel muestra además un resumen permanente: cuántos pasos faltan y cuál es el siguiente.

---

## 4. Descripción de cada pestaña

**Las pestañas están ordenadas igual que los pasos del asistente**, para que usted avance de
izquierda a derecha sin tener que pensar qué toca hacer:

| Pestaña | Pasos del asistente que se resuelven aquí |
|---|---|
| 📊 Panel | Inicio: estado general del mes y botón del asistente |
| 👥 Personal | 1 padrón · 2 turnos · 3 horarios · 4 asignar horario · 5 sueldos |
| ⚙️ Configuración | 6 parámetros de cálculo · 7 feriados |
| 📥 Marcaciones | 8 importar las marcaciones · 9 vincular las personas del reloj |
| 🗓️ Asistencia | 10 recalcular y revisar la asistencia |
| 💵 Planilla | 11 calcular la planilla y emitir boletas |
| 📈 Reportes | Consultas y exportación de resultados |

### 4.1 Panel

Es el resumen general. Muestra seis indicadores del mes seleccionado:

- **Empleados activos**
- **Marcaciones cargadas** (total histórico en la base)
- **Tardanzas del mes**
- **Faltas del mes**
- **Horas extra del mes** (lo que realmente se paga)
- **Neto de la planilla** del mes

Debajo aparece una franja amarilla con **avisos automáticos**. Es la sección más importante para detectar problemas antes de que afecten la planilla. Avisa cuando:

- Hay personas del reloj sin vincular a un DNI.
- Hay empleados sin turno ni horario asignado (se calcularían como descanso).
- No hay turnos configurados.
- El mes todavía no se ha calculado.
- Algún parámetro tiene un valor inválido o sospechoso.
- Cuál fue la última importación y cuántas marcas nuevas trajo.

El Panel incluye también **acciones rápidas** que llevan directo a cada pestaña — entre ellas el botón **🧭 ASISTENTE PASO A PASO** — la lista de las **últimas importaciones** del reloj, y un resumen del asistente que indica cuántos pasos faltan y cuál es el siguiente.


### 4.2 Personal

**a) Empleados** — la lista toma el padrón de la tabla **Choferes** y le agrega los datos de nómina. Seleccione un empleado y complete su ficha:

- Sueldo básico, jornada en horas por día y cargo.
- Régimen, sistema de pensión (ONP, AFP o ninguno), nombre de la AFP, comisión más prima de la AFP y CUSPP.
- Fecha de ingreso, banco y cuenta de haberes, código ESSALUD.
- Asignación familiar (10% de la RMV), discapacidad, personal de confianza.
- Otros ingresos, otros descuentos, adelanto mensual y porcentaje de retención judicial.
- **Horario** y **turno** asignados.
- Estado (activo, cesado, suspendido).

El botón **🔄 Sincronizar desde Choferes** crea las fichas que falten. La **asignación masiva** permite aplicar un horario o un turno a **todos** los empleados activos de una sola vez.

La ficha tiene sus botones **💾 GUARDAR FICHA** y **🧹 Limpiar** arriba a la derecha, así siempre están a la vista: si la ventana es baja, la ficha se desplaza con su propia barra sin que los botones desaparezcan.

**b) Turnos** — el catálogo de horarios de trabajo. Cada turno define:

- Nombre, hora de entrada y hora de salida.
- Tolerancia en minutos (después de este margen la tardanza se descuenta).
- Minutos de refrigerio (se descuentan del tiempo trabajado).
- Jornada en horas.
- Días de la semana en que aplica.
- Si cruza la medianoche y si es nocturno (por ejemplo 22:00 a 06:00).

El sistema crea cinco turnos típicos la primera vez: Día (8:00-17:00), Mañana (6:00-14:00), Tarde (14:00-22:00), Noche (22:00-06:00) y Partido (8:00-19:00).

**c) Horarios** — una plantilla semanal: para cada día de la semana se elige un turno o Descanso. Se puede duplicar un horario para crear variantes (por ejemplo cambiar solo el sábado).

**d) Asignaciones** — cambios por rango de fechas. Sirve para turnos rotativos o reemplazos temporales: se indica el empleado, el horario o turno, la fecha desde y, opcionalmente, la fecha hasta.

**Orden de prioridad al calcular un día:**

1. Asignación por fecha vigente (la más reciente gana).
2. Horario semanal del empleado.
3. Turno fijo del empleado.
4. Si no hay nada de lo anterior, el día se considera **descanso**.


### 4.3 Configuración

**a) Parámetros de cálculo** — todos los valores que usa el motor. Vea el capítulo 8.

**b) Feriados y calendario** — registre feriados, días no laborables o días laborables especiales. El botón **🇵🇪 Cargar feriados del año** registra automáticamente los feriados nacionales del Perú, incluyendo Jueves y Viernes Santo.

**c) Herramientas** — tareas de mantenimiento:

- Sincronizar el padrón desde Choferes.
- Borrar la asistencia calculada de un período (no borra las marcaciones).
- Recalcular todos los meses que tengan marcaciones cargadas.
- Recrear o verificar las tablas del módulo.

---


### 4.4 Marcaciones

Tiene tres vistas internas.

**a) Importar Excel del reloj** — el proceso completo en tres pasos:

1. **Seleccionar el archivo** exportado por el captahuella (botón 📂) y pulsar 🔎 Analizar.
2. **Verificar las columnas detectadas**: el sistema las reconoce solo y las muestra en nueve casillas (Código, Nombres, Apellidos, Nombre completo, Fecha, Hora, Tipo de pase, Método de verificación, Departamento). Si algo quedó mal, se corrige con el desplegable.
3. **Vista previa**: muestra las primeras 150 filas válidas y, en la última columna (Vinculación), indica si esa persona ya está vinculada a un DNI.

Después pulse **✅ IMPORTAR MARCACIONES A LA BASE DE DATOS**. Al terminar informa cuántas filas leyó, cuántas marcas nuevas entraron, cuántas se descartaron por estar repetidas y cuántas personas quedaron sin vincular.

**b) Personas del reloj** — aquí se conecta cada código del reloj con el DNI real del empleado:

- La tabla lista todas las personas detectadas con su código, su nombre, cuántas marcas tienen y su rango de fechas.
- La columna **Sugerencia del sistema** propone un DNI cuando coincide el número o el nombre.
- Seleccione una fila, elija el empleado en el desplegable y pulse **🔗 Vincular**.
- El botón **🤖 Vincular automáticamente** resuelve de una sola vez todas las coincidencias confiables (DNI exacto, DNI normalizado o nombre idéntico).

Mientras una persona no esté vinculada, **sus marcaciones no entran en el cálculo de asistencia**.

**c) Historial** — todas las importaciones realizadas, con archivo, formato detectado, filas, marcas nuevas, duplicadas, rango de fechas y usuario. Sirve de auditoría.

También hay un botón para **🗑️ eliminar las marcaciones de un rango de fechas**, útil cuando se exportó mal un período y se quiere reimportar limpio.


### 4.5 Asistencia

**a) Matriz mensual** — la vista principal. Una fila por empleado y una columna por cada día del mes. Cada celda muestra el **código de estado** y la hora de entrada. Las filas se colorean según la situación del día.

Leyenda de códigos:

| Código | Estado | Significado |
|---|---|---|
| P | Puntual | Marcó dentro del horario y con tolerancia |
| T | Tardanza | Llegó después de la hora de entrada más la tolerancia |
| F | Falta | Tenía turno y no registró ninguna marcación |
| I | Incompleto | Registró solo una marcación (entrada o salida) |
| D | Descanso | No tenía turno asignado ese día |
| DT | Descanso trabajado | Día de descanso, pero asistió: se paga con recargo |
| FE | Feriado | Feriado del calendario, sin trabajo |
| FT | Feriado trabajado | Feriado en el que sí trabajó: se paga con recargo |
| V | Vacaciones | Cubierto por una incidencia de vacaciones |
| PE | Permiso | Cubierto por un permiso aprobado |
| LI | Licencia | Licencia registrada |
| DM | Descanso médico | Descanso médico registrado |
| CO | Comisión | Comisión de servicio |
| FJ | Falta justificada | Falta con sustento aprobado |
| JU | Justificado | Corrección manual del día |
| - | Pendiente | Día sin calcular |

Los botones de la barra superior permiten cambiar de mes (◀ ▶), ver el período, **🧮 recalcular la asistencia del mes** y **📊 exportar a Excel**.

**b) Detalle por día** — elija un empleado y vea sus días uno por uno con entrada, salida, número de marcas, tardanza, salida anticipada, tiempo trabajado, horas extra y estado. Puede filtrar por estado.

- **Doble clic** en una fila muestra **todas las marcaciones crudas del reloj** de ese día, con su tipo de pase y método de verificación.
- El botón **✏️ Corregir el día seleccionado** permite fijar manualmente la entrada, la salida, los minutos y el estado. Esa corrección **queda marcada como manual y no se sobrescribe** en los siguientes recálculos.

**c) Incidencias** — registre aquí lo que justifica una ausencia o un retraso:

- Tipo (vacaciones, permiso, licencia, descanso médico, comisión, falta justificada).
- Fechas desde y hasta.
- Motivo, si tiene goce de haber, horas (para permisos parciales) y quién lo aprobó.

Los días cubiertos por una incidencia **no generan descuento**. Después de guardar, el sistema le ofrece recalcular el mes para aplicarla.

**d) Horas extra** — lista las horas extra del mes con su origen:

- **AUTO**: calculadas por el sistema al recalcular la asistencia.
- **MANUAL**: registradas a mano en esta pantalla.

Puede agregar o corregir registros y marcar si están aprobadas. Solo las aprobadas se pagan.


### 4.6 Planilla

**a) Resumen del período** — seleccione el mes y pulse **🧮 CALCULAR / RECALCULAR PLANILLA**. La tabla muestra por empleado: días laborables, tardanzas, faltas, horas extra, total de ingresos, total de descuentos, neto a pagar y aportes del empleador. Abajo aparece el estado de la planilla y los totales generales.

Estados posibles: **BORRADOR**, **CALCULADA**, **APROBADA**, **PAGADA** y **CERRADA**. Una planilla marcada como pagada o cerrada no se puede recalcular hasta reabrirla.

**b) Boleta de pago** — elija el empleado y vea su boleta completa, con tres bloques (ingresos, descuentos y aportes del empleador), **la fórmula usada en cada línea** y el neto a pagar. Desde aquí puede exportar esa boleta o todas a Excel.

**c) Historial** — todos los períodos calculados con sus totales, estado, usuario y fecha de cálculo.


### 4.7 Reportes

Cinco reportes exportables a Excel:

| Reporte | Contenido |
|---|---|
| Puntualidad | Ranking por empleado con días laborables, puntuales, tardanzas, faltas, incompletos, minutos de tardanza, horas extra, horas trabajadas y porcentaje de puntualidad |
| Tardanzas | Detalle día por día de tardanzas y salidas anticipadas, ordenado de mayor a menor |
| Faltas | Todas las inasistencias del período con el día de la semana y la observación |
| Horas extra | Minutos, horas, tipo, origen, si está aprobada y el monto estimado |
| Marcaciones del reloj | Bitácora cruda de marcaciones (con rango de fechas libre) para auditoría |

## 5. El archivo del captahuella

### 5.1 Formatos soportados

El sistema **reconoce el formato automáticamente**. Busca la fila de encabezados en las primeras 40 filas y compara los títulos con una lista de nombres conocidos en español e inglés.

Formatos probados:

- **Hikvision** (equipos de control de acceso, por ejemplo el modelo DS-K1T321EFWX-B).
- **ZKTeco / ZKTime** (exportaciones de asistencia).
- **CSV genérico** y archivos con columnas propias.

### 5.2 Estructura del export de Hikvision

El archivo Transacciones tiene títulos en las primeras filas y los encabezados en la **fila 8**:

| Fila | Contenido |
|---|---|
| 1 a 3 | Vacías |
| 4 | Título del reporte |
| 5 | Hora de exportación |
| 6 | Operador |
| 7 | Período exportado |
| 8 | Encabezados de columna |
| 9 en adelante | Las marcaciones |

Columnas del encabezado: **Nombre, Apellido, ID, Departamento, Fecha, Día de la semana, Tiempo, Temperatura en la superficie de la piel, Estado de la temperatura, Tipo de pase de tarjeta, Método de verificación, Nombre personalizado, Fuente de datos, Tipo de gestión, Comentario**.

El sistema usa **ID** (código), **Nombre** y **Apellido** (para identificar a la persona), **Fecha** y **Tiempo**; además guarda el tipo de pase y el método de verificación para auditoría.

### 5.3 Recomendaciones al exportar

- Exporte en **.xlsx** o **.csv**. El formato .xls antiguo solo funciona si la librería pandas está instalada.
- **No modifique la disposición de las columnas** ni inserte filas encima de los encabezados.
- Puede exportar un rango amplio (por ejemplo todo el año): el sistema descarta las marcas repetidas y **no duplica nada** si vuelve a importar el mismo archivo.
- Puede exportar períodos que se solapan sin problema.

### 5.4 Códigos del reloj frente a DNI

Los equipos guardan el identificador que se les cargó, que no siempre es el DNI:

- Un código de **8 dígitos** suele ser el DNI y se vincula solo.
- Un código de **9 dígitos con ceros a la izquierda** (por ejemplo 004291054) es un código interno del equipo. En ese caso el sistema propone el vínculo por similitud de nombre y usted lo confirma en **Marcaciones → Personas del reloj**.

---

## 6. Cómo se calcula la asistencia

Estas son las reglas exactas que aplica el motor al pulsar Recalcular asistencia.

### 6.1 Marcaciones del día

1. Se toman **todas** las marcaciones del empleado en esa fecha.
2. Se ordenan por hora y se **descartan las lecturas repetidas**: si dos lecturas están a menos de **2 minutos** una de otra, se consideran la misma marcación. Los relojes biométricos suelen grabar dos o tres veces la misma lectura (reintentos de huella o de rostro); sin esta depuración, un día con una sola entrada parecería tener entrada y salida.
3. La **primera** marca es la entrada y la **última** es la salida.
4. Si el turno **cruza la medianoche**, se suman también las marcas del día siguiente hasta la hora de salida.

Ejemplo real del archivo del reloj: el 4 de setiembre un empleado registró 05:46, 05:47 y 05:48. Las tres son la misma lectura al ingresar, así que el día queda como **Incompleto** (una sola marcación) y no como un día completo trabajado.

### 6.2 Cálculo según el turno

| Situación | Resultado |
|---|---|
| Sin turno asignado ese día | Descanso (o Descanso trabajado si hay marcas) |
| Turno asignado y sin ninguna marca | Falta (se descuenta la jornada completa) |
| Una sola marca | Incompleto |
| Primera y última marca | Se calculan tardanza, salida anticipada y tiempo trabajado |

Fórmulas:

    tardanza          = primera marca - hora de entrada del turno - tolerancia
    salida anticipada = hora de salida del turno - última marca
    tiempo trabajado  = última marca - primera marca - minutos de refrigerio
    horas extra       = última marca - hora de salida del turno

Las horas extra se redondean a bloques de 15 minutos, se reconocen solo si superan los 15 minutos y tienen un tope de 300 minutos por día.

### 6.3 Feriados e incidencias

- Un día marcado como **FERIADO** o **NO LABORABLE** en el calendario nunca genera falta.
- Si el empleado **trabaja un feriado o un día de descanso**, el día se registra como Feriado trabajado o Descanso trabajado y se paga con recargo (ver 7.3). Ese día **no se cuenta además como horas extra**, para no pagarlo dos veces.
- Si el día está cubierto por una **incidencia** (vacaciones, permiso, licencia, descanso médico), el estado toma el nombre de la incidencia y **no hay descuento**.

### 6.4 Correcciones y recálculos

- Las correcciones manuales de un día quedan marcadas y **no se sobrescriben** al recalcular.
- Si recalcula un mes cuyas marcaciones todavía no se han importado, **todos los días saldrán como falta**. El sistema le avisa mostrando hasta qué fecha llegan las marcaciones cargadas antes de confirmar.

---

## 7. Cómo se calcula la planilla

### 7.1 Valores base

    valor día    = sueldo básico / días base del mes (30 por defecto)
    valor hora   = valor día / jornada en horas
    valor minuto = valor día / (jornada en horas x 60)

Con un sueldo de S/ 1,500 y jornada de 8 horas: valor día S/ 50.00, valor hora S/ 6.25, valor minuto S/ 0.1042.

### 7.2 Ingresos

| Concepto | Fórmula |
|---|---|
| Sueldo básico | El sueldo mensual registrado en la ficha |
| Asignación familiar | RMV x 10% (S/ 113.00 con RMV de S/ 1,130), si está marcada en la ficha |
| Horas extra 25% | Las **dos primeras horas** del día x valor hora x 1.25 |
| Horas extra 35% | Las horas **siguientes** del día x valor hora x 1.35 |
| Horas extra nocturnas | Valor hora x (25% + 35% de recargo nocturno) |
| Trabajo en descanso o feriado | Días x valor día x 2 (recargo del 100%) |
| Bonos y otros ingresos | Monto fijo de la ficha del empleado |

**Remuneración afecta** = sueldo básico + asignación familiar + horas extra + bonos y otros ingresos. Sobre esa base se calculan la ONP, la AFP, la renta de quinta categoría y el ESSALUD.

### 7.3 Descuentos

| Concepto | Fórmula |
|---|---|
| Tardanzas | Minutos de tardanza x valor minuto |
| Salidas anticipadas | Minutos de salida anticipada x valor minuto |
| Faltas injustificadas | Días de falta x valor día |
| ONP | Remuneración afecta x 13% |
| AFP | Remuneración afecta x (10% + comisión y prima configuradas para el empleado) |
| Renta de quinta categoría | Opcional, desactivada por defecto (ver 7.5) |
| Retención judicial | Remuneración afecta x porcentaje de la ficha |
| Adelanto de sueldo | Monto fijo de la ficha |
| Otros descuentos | Monto fijo de la ficha |

**Neto a pagar** = total de ingresos - total de descuentos.

### 7.4 Aportes del empleador (informativos)

| Concepto | Fórmula |
|---|---|
| ESSALUD | 9% de la remuneración afecta, con base mínima igual a la RMV |

Este importe no se descuenta al trabajador: es el costo que asume la empresa y aparece en la boleta solo como información.

### 7.5 Renta de quinta categoría (opcional)

Viene **desactivada**. Si se activa, el sistema proyecta la renta anual del trabajador, resta **7 UIT** y aplica la escala progresiva vigente: 8% hasta 5 UIT, 14% hasta 20 UIT, 17% hasta 35 UIT, 20% hasta 45 UIT y 30% por el exceso.

### 7.6 Ejemplo real de boleta

Empleado con sueldo de S/ 1,500, jornada de 8 horas, régimen ONP, en un mes con 19 faltas, 770 minutos de tardanza, 9 horas 45 minutos de horas extra repartidas en tres días de 195 minutos y un día de descanso trabajado:

| Tipo | Concepto | Monto |
|---|---|---|
| Ingreso | Sueldo Básico | 1,500.00 |
| Ingreso | Horas Extra 25% (6.00 h x 6.25 x 1.25) | 46.88 |
| Ingreso | Horas Extra 35% (3.75 h x 6.25 x 1.35) | 31.64 |
| Ingreso | Trabajo en Descanso / Feriado (1 día x 50 x 2) | 100.00 |
| | **Total ingresos** | **1,678.52** |
| Descuento | Tardanzas (770 min x 0.1042) | 80.21 |
| Descuento | Faltas Injustificadas (19 días x 50.00) | 950.00 |
| Descuento | ONP 13% de 1,578.52 | 205.21 |
| | **Total descuentos** | **1,235.42** |
| | **NETO A PAGAR** | **443.10** |
| Aporte | ESSALUD 9% de 1,578.52 | 142.07 |

Observe cómo los 195 minutos de cada día se parten en 120 minutos al 25% y 75 minutos al 35%.

---

## 8. Parámetros de cálculo

Todos los valores se editan en **Configuración → Parámetros de cálculo** y se guardan en la base de datos.

### 8.1 Jornada y marcaciones

| Parámetro | Por defecto | Qué hace |
|---|---|---|
| tolerancia_min | 10 | Minutos de gracia antes de considerar tardanza |
| refrigerio_min | 45 | Minutos que se descuentan del tiempo trabajado |
| jornada_horas | 8 | Jornada usada cuando el turno no la define |
| dias_base_mes | 30 | Días base para calcular el valor día |
| minutos_minimos_entre_marcas | 2 | Lecturas más cercanas que esto se consideran la misma marcación |

### 8.2 Horas extra y recargos

| Parámetro | Por defecto | Qué hace |
|---|---|---|
| he_25_pct | 25 | Recargo de las dos primeras horas extra del día |
| he_35_pct | 35 | Recargo de las horas extra siguientes |
| he_desde_hora | 2 | Cuántas horas se pagan al 25% antes de pasar al 35% |
| he_tope_dia_min | 300 | Tope de minutos extra reconocidos por día |
| he_minimo_min | 15 | Mínimo de minutos para reconocer una hora extra |
| he_bloque_min | 15 | Redondeo del tiempo extra en bloques de minutos |
| recargo_nocturno_pct | 35 | Recargo adicional por trabajo nocturno |
| recargo_feriado_pct | 100 | Recargo por trabajar en descanso o feriado |
| hora_inicio_nocturno | 22:00 | Inicio del horario nocturno |
| hora_fin_nocturno | 06:00 | Fin del horario nocturno |
| pagar_descanso_trabajado | SI | Pagar con recargo el trabajo en día de descanso |

### 8.3 Remuneraciones

| Parámetro | Por defecto | Qué hace |
|---|---|---|
| rmv | 1130 | Remuneración Mínima Vital vigente |
| uit | 5500 | UIT vigente |
| onp_pct | 13 | Aporte ONP del trabajador |
| essalud_pct | 9 | Aporte ESSALUD del empleador |
| asignacion_familiar_pct | 10 | Porcentaje de la RMV por asignación familiar |
| descontar_tardanza | SI | Descontar las tardanzas en la planilla |
| descontar_falta | SI | Descontar las faltas injustificadas |
| descontar_anticipo | SI | Descontar las salidas anticipadas |

### 8.4 Renta de quinta categoría

| Parámetro | Por defecto | Qué hace |
|---|---|---|
| calcular_renta_5ta | NO | Activar la retención mensual de quinta categoría |
| deduccion_uit_5ta | 7 | Cuántas UIT se deducen antes de aplicar la escala |

### 8.5 Validación de parámetros (lea esto)

El sistema **valida los parámetros antes de guardarlos**:

- No acepta campos vacíos, textos donde va un número, valores negativos, horas mal escritas ni respuestas distintas de SI/NO.
- Si algo está mal, muestra la lista de campos por corregir y **no guarda nada** (no deja la configuración a medias).
- Si un valor es válido pero sospechoso, pide confirmación antes de guardarlo.

**Cuidado con los ceros.** El sistema rechaza un 0 en el tope diario de horas extra (he_tope_dia_min), en la Remuneración Mínima Vital, en la UIT, en la jornada y en los días base del mes, porque dejarían el cálculo sin efecto. En los demás casos el 0 se acepta pero se avisa antes de guardar: por ejemplo un 0 en recargo_feriado_pct significa que trabajar en feriado no tendría ningún recargo, y un 0 en he_desde_hora haría que todas las horas extra se pagaran con el recargo del 35%. El Panel también muestra estos avisos.

Si alguna vez un valor quedó vacío o mal escrito, use el botón **🩹 Reparar valores inválidos**: restaura a su valor por defecto únicamente los campos dañados, sin tocar los demás.

---

## 9. Usuarios, equipos y conexión

- Toda la información del módulo vive en la **misma base de datos** que el resto del sistema, por lo que **todos los equipos ven lo mismo** apenas se guarda un cambio.
- Cada acción que modifica datos queda registrada en la **Bitácora de Auditoría** con el usuario, la fecha y la hora.
- Si no hay conexión, el sistema avisa y trabaja en modo lectura: podrá ver los datos cargados, pero no guardar.
- Los permisos por usuario se administran en **Ajustes de Sistema → Configurar Usuarios**.

---

## 10. Solución de problemas

| Síntoma | Causa probable | Solución |
|---|---|---|
| Al abrir el módulo dice "No tiene permisos" | El permiso es nuevo y no está marcado para ese usuario | Configurar Usuarios, marcar Nómina y Asistencia y guardar |
| Todos los días aparecen como Descanso | Los empleados no tienen turno ni horario asignado | Personal → Empleados (o asignación masiva) y recalcular |
| Muchos días aparecen como Incompleto | El reloj solo registró una marca ese día (muy común: solo hay entradas) | Verifique en Detalle por día las marcas crudas. Si falta la salida, el día queda incompleto a propósito |
| Aparecen faltas en días recientes | Las marcaciones cargadas llegan solo hasta cierta fecha | Importe el export más reciente y recalcule. El sistema avisa hasta qué fecha hay marcaciones |
| Una persona del reloj no cuenta en la asistencia | Su código no está vinculado a un DNI | Marcaciones → Personas del reloj → Vincular |
| Un empleado no aparece en la lista | No está registrado en el módulo de Choferes | Regístrelo en Choferes y pulse Sincronizar desde Choferes |
| Las horas extra salen en cero | El parámetro he_tope_dia_min está en 0, o las horas extra están desaprobadas | Configuración → Parámetros (ponga 300) y recalcular |
| La planilla sale en cero | Los sueldos de la ficha están en 0 | Personal → Empleados, cargar el sueldo real y recalcular la planilla |
| El neto cambió de un mes a otro sin razón aparente | Cambió un parámetro, un turno o un sueldo | Recalcule asistencia y luego planilla con los valores actuales |
| Al reimportar el mismo archivo dice que todas son duplicadas | Es correcto: el sistema no repite marcas ya cargadas | Nada que hacer; el archivo ya estaba importado |
| Las tardanzas son enormes | El turno asignado no corresponde al horario real de esa persona | Revise el turno en Personal → Empleados o use una asignación por fecha |
| Un día de descanso trabajado no aparece como hora extra | Es correcto: se paga como día completo con recargo | Lo verá como concepto Trabajo en Descanso / Feriado en la boleta |
| Los cambios no se reflejan | Falta recalcular | Asistencia → Recalcular; después Planilla → Calcular |
| No veo el botón en el menú | Falta el permiso, o la sesión es anterior al cambio | Cierre sesión y vuelva a entrar; revise permisos |
| Quiero más espacio en pantalla | La grilla mensual es ancha | Botón ⛶ Pantalla completa o la tecla F11: los botones del sistema se ocultan solos |
| Me quedé sin el menú lateral | Está en pantalla completa (por eso se ocultó) | Pulse F11, Esc o el botón flotante 🗗 de la esquina inferior derecha |
| No sé por dónde empezar o qué me falta | Hay varios pasos de configuración pendientes | Pulse **🧭 ASISTENTE PASO A PASO** en el Panel: le dice qué falta y lo lleva a cada pantalla |

---

## 11. Mantenimiento y respaldo

### 11.1 Rutina mensual sugerida

1. Exportar el período desde el reloj.
2. Importar el archivo y revisar el informe de importación.
3. Vincular las personas nuevas.
4. Recalcular la asistencia del mes.
5. Revisar la grilla: tardanzas, faltas e incompletos.
6. Registrar incidencias del mes (vacaciones, permisos, descansos médicos).
7. Confirmar las horas extra.
8. Calcular la planilla y revisar las boletas.
9. Exportar la planilla y las boletas a Excel y archivarlas.
10. Cambiar el estado de la planilla a APROBADA y luego PAGADA.

### 11.2 Respaldos

- Toda la información está en la base de datos del sistema.
- Además, exporte cada mes la **planilla**, las **boletas** y la **asistencia** a Excel con los botones de exportación. Esos archivos son su respaldo de lectura rápida.
- Las marcaciones crudas se conservan completas; reimportar un export antiguo no las altera.

### 11.3 Herramientas de mantenimiento

En **Configuración → Herramientas**:

- Sincronizar el padrón desde Choferes.
- Borrar la asistencia calculada de un período (útil antes de un recálculo limpio).
- Recalcular todos los meses con marcaciones.
- Recrear o verificar las tablas del módulo.

### 11.4 Pruebas de regresión (uso técnico)

En la carpeta del sistema hay dos scripts que permiten comprobar que todo sigue funcionando después de una actualización:

- **\_smoke_gui.py**: abre el módulo y recorre todas las pestañas y vistas, informando cualquier error.
- **\_test_pantalla.py**: comprueba el botón de pantalla completa y los atajos.
- **\_test_barra.py**: comprueba que en pantalla completa se oculte la barra lateral y que vuelva a su sitio al salir.
- **\_test_asistente.py**: imprime el estado que detecta el asistente para un período (útil para revisar por qué marca un paso como pendiente).
- **\_test_layout.py**: comprueba que ningún botón o campo quede fuera de la ventana en ninguna vista (útil si se cambia el diseño o se trabaja en pantallas pequeñas).
- **\_test_letrero.py**: comprueba que el aviso de "leyendo base de datos / calculando" aparezca y desaparezca en cada operación, y que nunca quede pegado.
- **\_test_ayuda.py**: comprueba que el botón de Ayuda encuentre el instructivo, lo muestre dentro del sistema y permita navegar y hacer zoom.

Se ejecutan con el comando python seguido del nombre del archivo, desde la carpeta del programa.

---

## 12. Anexo A: tablas de la base de datos

| Tabla | Contenido |
|---|---|
| nom_parametros | Valores de configuración del módulo |
| nom_turnos | Catálogo de turnos |
| nom_horarios | Plantillas de horario semanal |
| nom_horario_detalle | Turno asignado a cada día de la semana de un horario |
| nom_empleado | Datos de nómina de cada empleado (se complementa con la tabla choferes) |
| nom_asignaciones | Asignaciones de horario o turno por rango de fechas |
| nom_calendario | Feriados y días no laborables |
| nom_mapeo_reloj | Vínculo entre el código del reloj y el DNI del empleado |
| nom_marcaciones | Marcaciones crudas importadas del reloj |
| nom_asistencia | Asistencia diaria calculada |
| nom_incidencias | Vacaciones, permisos, licencias y descansos médicos |
| nom_horas_extra | Horas extra automáticas y manuales |
| nom_importaciones | Historial de importaciones del reloj |
| nom_planilla | Cabecera de la planilla por período |
| nom_planilla_detalle | Detalle de conceptos de la planilla |

---

## 13. Anexo B: archivos del módulo

| Archivo | Función |
|---|---|
| nomina_core.py | Motor principal: esquema, catálogos, lectura del reloj y cálculo de asistencia |
| nomina_planilla.py | Cálculo de la planilla, boletas y reportes |
| modulo_nomina.py | Interfaz gráfica del módulo |
| control_general.py | Sistema principal, donde está el botón del módulo |

---

## 14. Anexo C: atajos y gestos frecuentes

| Acción | Cómo |
|---|---|
| Entrar a pantalla completa (oculta la barra lateral) | F11, o el botón ⛶ |
| Salir de pantalla completa (devuelve la barra lateral) | Esc, F11 o el botón flotante 🗗 |
| Elegir una fecha | Botón 📅 junto al campo |
| Ver el mes anterior o siguiente | Botones ◀ ▶ |
| Ver las marcas crudas de un día | Doble clic en la fila del detalle por día |
| Ver la boleta de un empleado | Doble clic en su fila del resumen de planilla |
| Saber qué falta configurar | Botón 🧭 ASISTENTE PASO A PASO en el Panel |

---

*Documento generado para el Sistema de Control de Flota Automotriz. Ante cualquier duda sobre un cálculo, revise la boleta del empleado: cada línea indica la fórmula exacta que se utilizó.*
