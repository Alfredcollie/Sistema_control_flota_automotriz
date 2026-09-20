# -*- coding: utf-8 -*-
"""
NOMINA_PLANILLA.PY - Motor de Liquidacion de Planilla
======================================================
Calcula la planilla mensual a partir de la asistencia ya procesada por
nomina_core (tardanzas, faltas, horas extra, descansos trabajados) mas los
datos contractuales del empleado (sueldo, asignacion familiar, pension).

Genera:
  * Cabecera del periodo en 'nom_planilla'.
  * Un detalle por concepto en 'nom_planilla_detalle' (ingresos, descuentos
    y aportes del empleador), con la formula usada en cada linea.
  * La boleta de pago de cada empleado.
  * Reportes de tardanzas, faltas, horas extra y puntualidad.

Reglas peruanas aplicadas (parametrizables desde la GUI):
    valor_dia    = sueldo_basico / dias_base_mes (30)
    valor_hora   = valor_dia / jornada_horas
    valor_minuto = valor_dia / (jornada_horas * 60)
    tardanza     = minutos_tardanza * valor_minuto
    falta        = dias_falta * valor_dia
    hora_extra   = horas * valor_hora * (1 + recargo/100)
    ONP          = 13% de la remuneracion afecta
    AFP          = (10% + comision + prima) de la remuneracion afecta
    ESSALUD      = 9% de la remuneracion afecta (aporte del empleador)
    renta_5ta    = proyeccion anual - 7 UIT, escala progresiva (opcional)

Autor: Collie Software
"""
from datetime import datetime

import nomina_core as nc


# =========================================================================
# CONSTANTES DE CONCEPTOS
# =========================================================================
TIPO_INGRESO = "INGRESO"
TIPO_DESCUENTO = "DESCUENTO"
TIPO_APORTE = "APORTE"

CAT_SUELDO = "SUELDO"
CAT_ASIGNACION = "ASIGNACION FAMILIAR"
CAT_HORAS_EXTRA = "HORAS EXTRA"
CAT_DESCANSO = "DESCANSO TRABAJADO"
CAT_BONO = "BONOS"
CAT_TARDANZA = "TARDANZA"
CAT_ANTICIPO = "SALIDA ANTICIPADA"
CAT_FALTA = "FALTA"
CAT_PENSION = "PENSION"
CAT_RENTA = "RENTA 5TA"
CAT_JUDICIAL = "RETENCION JUDICIAL"
CAT_ADELANTO = "ADELANTO"
CAT_OTROS = "OTROS"
CAT_ESSALUD = "ESSALUD"

ESTADOS_PLANILLA = ["BORRADOR", "CALCULADA", "APROBADA", "PAGADA", "CERRADA"]

# Escala progresiva de renta de quinta categoria (tramos en UIT)
ESCALA_QUINTA = [(0, 5, 8.0), (5, 20, 14.0), (20, 35, 17.0), (35, 45, 20.0), (45, None, 30.0)]


# =========================================================================
# UTILIDADES DE CALCULO
# =========================================================================
def valor_dia(sueldo, parametros):
    base = nc.a_float(parametros.get("dias_base_mes"), 30) or 30
    return nc.a_float(sueldo) / base


def valor_hora(sueldo, jornada, parametros):
    dia = valor_dia(sueldo, parametros)
    horas = nc.a_float(jornada, 8) or 8
    return dia / horas


def aplicar_escala_quinta(renta_neta, uit):
    """Aplica la escala progresiva del impuesto a la renta de 5ta categoria."""
    impuesto = 0.0
    resto = max(0.0, renta_neta)
    for inferior, superior, tasa in ESCALA_QUINTA:
        if resto <= 0:
            break
        base_inferior = inferior * uit
        base_superior = (superior * uit) if superior is not None else None
        ancho = (base_superior - base_inferior) if base_superior is not None else resto
        gravado = min(resto, ancho)
        if gravado <= 0:
            continue
        impuesto += gravado * tasa / 100.0
        resto -= gravado
    return impuesto


def calcular_retencion_quinta(remuneracion_afecta, mes, sueldo, parametros,
                              acumulado_remuneracion=0.0, acumulado_retencion=0.0):
    """Retencion mensual de renta de 5ta categoria por proyeccion anual."""
    uit = nc.a_float(parametros.get("uit"), 5500) or 5500
    deduccion_uit = nc.a_float(parametros.get("deduccion_uit_5ta"), 7) or 7
    meses_restantes = 13 - int(mes)
    if meses_restantes <= 0:
        meses_restantes = 1
    proyeccion = remuneracion_afecta * meses_restantes
    # Gratificaciones de julio y diciembre que aun no se pagan
    gratificacion = nc.a_float(sueldo)
    gratificaciones_futuras = 0.0
    if mes <= 6:
        gratificaciones_futuras += gratificacion * 2
    elif mes <= 11:
        gratificaciones_futuras += gratificacion
    renta_anual = acumulado_remuneracion + proyeccion + gratificaciones_futuras
    renta_neta = renta_anual - (deduccion_uit * uit)
    if renta_neta <= 0:
        return 0.0, {"renta_anual": renta_anual, "renta_neta": 0.0, "impuesto_anual": 0.0,
                     "meses_restantes": meses_restantes, "uit": uit}
    impuesto_anual = aplicar_escala_quinta(renta_neta, uit)
    retencion = max(0.0, (impuesto_anual - acumulado_retencion) / meses_restantes)
    memoria = {"renta_anual": renta_anual, "renta_neta": renta_neta,
               "impuesto_anual": impuesto_anual, "meses_restantes": meses_restantes, "uit": uit}
    return nc.redondear(retencion), memoria


# =========================================================================
# CALCULO DE LA PLANILLA
# =========================================================================
def liquidar_empleado(empleado, periodo, resumen, horas_extra, parametros, mes,
                      acumulado_remuneracion=0.0, acumulado_retencion=0.0):
    """Calcula las lineas de la boleta de un empleado. Devuelve (lineas, totales)."""
    sueldo = nc.a_float(empleado.get("sueldo_basico"))
    jornada = nc.a_float(empleado.get("jornada_horas"), 8) or 8
    v_dia = valor_dia(sueldo, parametros)
    v_hora = valor_hora(sueldo, jornada, parametros)
    v_minuto = v_dia / (jornada * 60.0) if jornada else 0.0
    rmv = nc.a_float(parametros.get("rmv"), 1130) or 1130

    lineas = []
    orden = [0]

    def agregar(concepto, tipo, categoria, monto, base=0.0, cantidad=0.0, unidad="", formula=""):
        if abs(nc.a_float(monto)) < 0.005:
            return
        orden[0] += 1
        lineas.append({
            "concepto": concepto, "tipo": tipo, "categoria": categoria,
            "base": nc.redondear(base), "cantidad": nc.redondear(cantidad, 2),
            "unidad": unidad, "monto": nc.redondear(monto), "formula": formula, "orden": orden[0],
        })

    # ---------- INGRESOS ----------
    agregar("Sueldo Basico", TIPO_INGRESO, CAT_SUELDO, sueldo, sueldo, 1, "MES",
            "Sueldo basico mensual")
    asignacion = 0.0
    if empleado.get("asignacion_familiar"):
        asignacion = nc.redondear(rmv * nc.a_float(parametros.get("asignacion_familiar_pct"), 10) / 100.0)
        agregar("Asignacion Familiar", TIPO_INGRESO, CAT_ASIGNACION, asignacion, rmv, 1, "MES",
                "RMV x %s%%" % parametros.get("asignacion_familiar_pct", 10))

    # Horas extra. Regla peruana (D.S. 007-2002-TR): las DOS PRIMERAS horas extra del
    # dia se pagan con recargo del 25% y las siguientes con el 35%. La tabla guarda un
    # registro por dia, asi que el tramo se separa aqui minuto a minuto.
    extra_25 = nc.a_int(parametros.get("he_25_pct"), 25)
    extra_35 = nc.a_int(parametros.get("he_35_pct"), 35)
    recargo_noche = nc.a_int(parametros.get("recargo_nocturno_pct"), 35)
    limite_25 = nc.a_float(parametros.get("he_desde_hora"), 2) * 60
    if limite_25 <= 0:
        limite_25 = 120
    minutos_25 = minutos_35 = minutos_nocturnos = 0
    for registro in horas_extra:
        if not registro.get("aprobado", True):
            continue
        minutos = nc.a_int(registro.get("minutos"))
        if minutos <= 0:
            continue
        if str(registro.get("tipo", "")).startswith("NOCTURNA"):
            minutos_nocturnos += minutos
            continue
        minutos_25 += min(minutos, limite_25)
        minutos_35 += max(0, minutos - limite_25)

    total_horas_extra = 0.0
    tramos = [
        ("Horas Extra %d%% (primeras %.0f h del dia)" % (extra_25, limite_25 / 60.0), minutos_25,
         extra_25, "primeras %.0f h del dia" % (limite_25 / 60.0)),
        ("Horas Extra %d%% (horas siguientes)" % extra_35, minutos_35, extra_35, "horas siguientes"),
        ("Horas Extra Nocturna %d%%" % (extra_25 + recargo_noche), minutos_nocturnos,
         extra_25 + recargo_noche, "turno nocturno"),
    ]
    for concepto, minutos, recargo, detalle in tramos:
        if minutos <= 0:
            continue
        horas = minutos / 60.0
        monto = horas * v_hora * (1 + recargo / 100.0)
        total_horas_extra += monto
        agregar(concepto, TIPO_INGRESO, CAT_HORAS_EXTRA, monto, v_hora, horas, "HORAS",
                "%.2f h (%s) x %s x (1 + %d%%)" % (horas, detalle, nc.redondear(v_hora, 4), recargo))

    # Trabajo en descanso o feriado (se paga con recargo)
    recargo_descanso = nc.a_int(parametros.get("recargo_feriado_pct"), 100)
    if nc.normalizar_texto(parametros.get("pagar_descanso_trabajado", "SI")) == "SI":
        dias_descanso_trab = nc.a_int(resumen.get("descansos_trabajados"))
        if dias_descanso_trab > 0:
            monto = dias_descanso_trab * v_dia * (1 + recargo_descanso / 100.0)
            agregar("Trabajo en Descanso / Feriado", TIPO_INGRESO, CAT_DESCANSO, monto, v_dia,
                    dias_descanso_trab, "DIAS", "%d dia(s) x %s x (1 + %d%%)" %
                    (dias_descanso_trab, nc.redondear(v_dia, 2), recargo_descanso))

    otros_ingresos = nc.a_float(empleado.get("otros_ingresos"))
    agregar("Bonos y Otros Ingresos", TIPO_INGRESO, CAT_BONO, otros_ingresos, otros_ingresos, 1, "MONTO",
            "Monto fijo registrado en la ficha del empleado")

    total_ingresos = sum(l["monto"] for l in lineas if l["tipo"] == TIPO_INGRESO)
    remuneracion_afecta = nc.redondear(sueldo + asignacion + total_horas_extra + otros_ingresos)

    # ---------- DESCUENTOS ----------
    if nc.normalizar_texto(parametros.get("descontar_tardanza", "SI")) == "SI":
        minutos_tardanza = nc.a_int(resumen.get("minutos_tardanza"))
        if minutos_tardanza > 0:
            monto = minutos_tardanza * v_minuto
            agregar("Descuento por Tardanzas", TIPO_DESCUENTO, CAT_TARDANZA, monto, v_dia,
                    minutos_tardanza, "MINUTOS", "%d min x %s" % (minutos_tardanza, nc.redondear(v_minuto, 4)))
    if nc.normalizar_texto(parametros.get("descontar_anticipo", "SI")) == "SI":
        minutos_anticipo = nc.a_int(resumen.get("minutos_anticipo"))
        if minutos_anticipo > 0:
            monto = minutos_anticipo * v_minuto
            agregar("Descuento por Salidas Anticipadas", TIPO_DESCUENTO, CAT_ANTICIPO, monto, v_dia,
                    minutos_anticipo, "MINUTOS", "%d min x %s" % (minutos_anticipo, nc.redondear(v_minuto, 4)))
    if nc.normalizar_texto(parametros.get("descontar_falta", "SI")) == "SI":
        dias_falta = nc.a_int(resumen.get("faltas"))
        if dias_falta > 0:
            monto = dias_falta * v_dia
            agregar("Descuento por Faltas Injustificadas", TIPO_DESCUENTO, CAT_FALTA, monto, v_dia,
                    dias_falta, "DIAS", "%d dia(s) x %s" % (dias_falta, nc.redondear(v_dia, 2)))

    # Sistema de pensiones
    sistema = nc.normalizar_texto(empleado.get("sistema_pension") or "ONP")
    aporte_pension = 0.0
    if sistema == "AFP":
        obligatorio = 10.0
        comision = nc.a_float(empleado.get("afp_comision_pct"))
        tasa = obligatorio + comision
        aporte_pension = remuneracion_afecta * tasa / 100.0
        agregar("AFP %s (%.2f%%)" % (empleado.get("afp_nombre") or "", tasa), TIPO_DESCUENTO,
                CAT_PENSION, aporte_pension, remuneracion_afecta, tasa, "%",
                "Remuneracion afecta x %.2f%% (10%% aporte + %.2f%% comision/prima)" % (tasa, comision))
    elif sistema == "ONP":
        tasa = nc.a_float(parametros.get("onp_pct"), 13) or 13
        aporte_pension = remuneracion_afecta * tasa / 100.0
        agregar("ONP %.2f%%" % tasa, TIPO_DESCUENTO, CAT_PENSION, aporte_pension,
                remuneracion_afecta, tasa, "%", "Remuneracion afecta x %.2f%%" % tasa)

    # Renta de quinta categoria (opcional)
    memoria_renta = None
    if nc.normalizar_texto(parametros.get("calcular_renta_5ta", "NO")) == "SI":
        retencion, memoria_renta = calcular_retencion_quinta(
            remuneracion_afecta, mes, sueldo, parametros, acumulado_remuneracion, acumulado_retencion)
        if retencion > 0:
            agregar("Renta 5ta Categoria", TIPO_DESCUENTO, CAT_RENTA, retencion, remuneracion_afecta,
                    mes, "MES", "Proyeccion anual con deduccion de %s UIT" %
                    parametros.get("deduccion_uit_5ta", 7))

    retencion_judicial = remuneracion_afecta * nc.a_float(empleado.get("retencion_judicial_pct")) / 100.0
    agregar("Retencion Judicial", TIPO_DESCUENTO, CAT_JUDICIAL, retencion_judicial, remuneracion_afecta,
            nc.a_float(empleado.get("retencion_judicial_pct")), "%", "Remuneracion x %s%%" %
            empleado.get("retencion_judicial_pct"))

    adelanto = nc.a_float(empleado.get("adelanto_mensual"))
    agregar("Adelanto de Sueldo", TIPO_DESCUENTO, CAT_ADELANTO, adelanto, adelanto, 1, "MONTO",
            "Adelanto registrado en la ficha del empleado")
    otros_descuentos = nc.a_float(empleado.get("otros_descuentos"))
    agregar("Otros Descuentos", TIPO_DESCUENTO, CAT_OTROS, otros_descuentos, otros_descuentos, 1, "MONTO",
            "Descuentos adicionales de la ficha del empleado")

    total_descuentos = sum(l["monto"] for l in lineas if l["tipo"] == TIPO_DESCUENTO)

    # ---------- APORTES DEL EMPLEADOR ----------
    essalud_pct = nc.a_float(parametros.get("essalud_pct"), 9) or 9
    base_essalud = max(remuneracion_afecta, rmv)
    aporte_essalud = base_essalud * essalud_pct / 100.0
    agregar("ESSALUD %.2f%%" % essalud_pct, TIPO_APORTE, CAT_ESSALUD, aporte_essalud, base_essalud,
            essalud_pct, "%", "Base minima RMV. Remuneracion x %.2f%%" % essalud_pct)
    total_aportes = sum(l["monto"] for l in lineas if l["tipo"] == TIPO_APORTE)

    totales = {
        "total_ingresos": nc.redondear(total_ingresos),
        "total_descuentos": nc.redondear(total_descuentos),
        "neto_pagar": nc.redondear(total_ingresos - total_descuentos),
        "total_aportes": nc.redondear(total_aportes),
        "remuneracion_afecta": remuneracion_afecta,
        "aporte_pension": nc.redondear(aporte_pension),
        "memoria_renta": memoria_renta,
        "valor_dia": nc.redondear(v_dia, 2),
        "valor_hora": nc.redondear(v_hora, 4),
    }
    return lineas, totales


def calcular_planilla(periodo, usuario="sistema", dnis=None, observacion="", recalcular=True):
    """Calcula (o recalcula) la planilla completa de un periodo 'YYYY-MM'."""
    ok, error = nc.inicializar_esquema_nomina()
    if not ok:
        return {"ok": False, "error": error}
    dias = nc.dias_del_periodo(periodo)
    if not dias:
        return {"ok": False, "error": "Periodo invalido. Use el formato AAAA-MM."}
    estado_actual = nc._consultar("SELECT estado FROM nom_planilla WHERE periodo = %s", (periodo,), uno=True)
    if estado_actual and estado_actual[0] in ("PAGADA", "CERRADA") and recalcular:
        return {"ok": False, "error": "La planilla de %s esta %s. Reabrala antes de recalcular."
                % (periodo, estado_actual[0])}

    parametros = nc.obtener_parametros()
    mes = int(str(periodo).split("-")[1])
    empleados = nc.listar_empleados(solo_activos=False)
    if dnis:
        empleados = [e for e in empleados if str(e["dni"]) in set(str(d) for d in dnis)]
    resumen = nc.resumen_asistencia(dias[0], dias[-1])
    extra = nc.listar_horas_extra(desde=dias[0], hasta=dias[-1])

    por_dni_extra = {}
    for registro in extra:
        por_dni_extra.setdefault(str(registro["dni"]), []).append(registro)

    lineas_guardar = []
    resumen_empleados = []
    totales_periodo = {"ingresos": 0.0, "descuentos": 0.0, "neto": 0.0, "aportes": 0.0}
    calculados = 0
    # Acumulados del anio para la renta de 5ta (meses anteriores del mismo ejercicio)
    anio = str(periodo).split("-")[0]
    for empleado in empleados:
        dni = str(empleado["dni"])
        datos_asistencia = resumen.get(dni, {})
        if not datos_asistencia.get("dias") and nc.a_float(empleado.get("sueldo_basico")) <= 0:
            continue  # sin asistencia ni sueldo: no entra en la planilla
        acumulado_remun = 0.0
        acumulado_ret = 0.0
        if nc.normalizar_texto(parametros.get("calcular_renta_5ta", "NO")) == "SI":
            fila = nc._consultar(
                "SELECT COALESCE(SUM(CASE WHEN concepto = 'Sueldo Basico' THEN monto ELSE 0 END),0), "
                "COALESCE(SUM(CASE WHEN categoria = %s THEN monto ELSE 0 END),0) "
                "FROM nom_planilla_detalle WHERE dni = %s AND periodo LIKE %s AND periodo < %s",
                (CAT_RENTA, dni, anio + "-%", periodo), uno=True)
            if fila:
                acumulado_remun = nc.a_float(fila[0])
                acumulado_ret = nc.a_float(fila[1])
        lineas, totales = liquidar_empleado(empleado, periodo, datos_asistencia,
                                            por_dni_extra.get(dni, []), parametros, mes,
                                            acumulado_remun, acumulado_ret)
        for linea in lineas:
            lineas_guardar.append((
                periodo, dni, empleado.get("nombre") or dni, linea["concepto"], linea["tipo"],
                linea["categoria"], linea["base"], linea["cantidad"], linea["unidad"],
                linea["monto"], linea["formula"], linea["orden"]))
        totales_periodo["ingresos"] += totales["total_ingresos"]
        totales_periodo["descuentos"] += totales["total_descuentos"]
        totales_periodo["neto"] += totales["neto_pagar"]
        totales_periodo["aportes"] += totales["total_aportes"]
        resumen_empleados.append({
            "dni": dni, "empleado": empleado.get("nombre") or dni, "cargo": empleado.get("cargo") or "",
            "sueldo": nc.a_float(empleado.get("sueldo_basico")),
            "dias_laborables": nc.a_int(datos_asistencia.get("dias_laborables")),
            "tardanzas": nc.a_int(datos_asistencia.get("tardanzas")),
            "faltas": nc.a_int(datos_asistencia.get("faltas")),
            "incompletos": nc.a_int(datos_asistencia.get("incompletos")),
            "horas_extra": nc.redondear(nc.a_int(datos_asistencia.get("minutos_extra")) / 60.0, 2),
            **totales,
        })
        calculados += 1

    ahora = datetime.now().strftime("%d/%m/%Y %H:%M")
    conn = nc.conectar_db(silencioso=True)
    if not conn:
        return {"ok": False, "error": "Sin conexion a la base de datos"}
    try:
        with conn.cursor() as cursor:
            cursor.execute("DELETE FROM nom_planilla_detalle WHERE periodo = %s", (periodo,))
            cursor.execute(
                "INSERT INTO nom_planilla (periodo, estado, fecha_calculo, usuario, total_ingresos, "
                "total_descuentos, total_neto, total_aportes, n_empleados, observacion) "
                "VALUES (%s, 'CALCULADA', %s, %s, %s, %s, %s, %s, %s, %s) "
                "ON CONFLICT (periodo) DO UPDATE SET estado = 'CALCULADA', "
                "fecha_calculo = EXCLUDED.fecha_calculo, usuario = EXCLUDED.usuario, "
                "total_ingresos = EXCLUDED.total_ingresos, total_descuentos = EXCLUDED.total_descuentos, "
                "total_neto = EXCLUDED.total_neto, total_aportes = EXCLUDED.total_aportes, "
                "n_empleados = EXCLUDED.n_empleados, observacion = EXCLUDED.observacion",
                (periodo, ahora, usuario, nc.redondear(totales_periodo["ingresos"]),
                 nc.redondear(totales_periodo["descuentos"]), nc.redondear(totales_periodo["neto"]),
                 nc.redondear(totales_periodo["aportes"]), calculados, observacion))
            if lineas_guardar:
                cursor.executemany(
                    "INSERT INTO nom_planilla_detalle (periodo, dni, empleado, concepto, tipo, "
                    "categoria, base, cantidad, unidad, monto, formula, orden) "
                    "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", lineas_guardar)
        conn.commit()
    except Exception as e:
        try:
            conn.rollback()
        except Exception:
            pass
        return {"ok": False, "error": str(e)}
    finally:
        nc.liberar_conexion(conn)

    nc.registrar_auditoria(usuario, "Nomina", "Calculo la planilla de %s (%d empleados, neto %s)"
                           % (periodo, calculados, nc.redondear(totales_periodo["neto"])))
    return {
        "ok": True, "error": "", "periodo": periodo, "empleados": calculados,
        "lineas": len(lineas_guardar),
        "total_ingresos": nc.redondear(totales_periodo["ingresos"]),
        "total_descuentos": nc.redondear(totales_periodo["descuentos"]),
        "total_neto": nc.redondear(totales_periodo["neto"]),
        "total_aportes": nc.redondear(totales_periodo["aportes"]),
        "detalle": resumen_empleados,
    }


# =========================================================================
# CONSULTAS DE PLANILLA
# =========================================================================
def obtener_planilla(periodo):
    fila = nc._consultar(
        "SELECT periodo, estado, fecha_calculo, usuario, total_ingresos, total_descuentos, "
        "total_neto, total_aportes, n_empleados, observacion FROM nom_planilla WHERE periodo = %s",
        (periodo,), uno=True)
    if not fila:
        return None
    columnas = ["periodo", "estado", "fecha_calculo", "usuario", "total_ingresos", "total_descuentos",
                "total_neto", "total_aportes", "n_empleados", "observacion"]
    datos = dict(zip(columnas, fila))
    for clave in ("total_ingresos", "total_descuentos", "total_neto", "total_aportes"):
        datos[clave] = nc.a_float(datos.get(clave))
    datos["n_empleados"] = nc.a_int(datos.get("n_empleados"))
    return datos


def listar_periodos():
    """Periodos con planilla calculada, mas los ultimos periodos con marcaciones."""
    columnas = ["periodo", "estado", "fecha_calculo", "usuario", "total_ingresos", "total_descuentos",
                "total_neto", "total_aportes", "n_empleados"]
    periodos = []
    for fila in nc._consultar(
            "SELECT periodo, estado, fecha_calculo, usuario, total_ingresos, total_descuentos, "
            "total_neto, total_aportes, n_empleados FROM nom_planilla ORDER BY periodo DESC") or []:
        datos = dict(zip(columnas, fila))
        for clave in ("total_ingresos", "total_descuentos", "total_neto", "total_aportes"):
            datos[clave] = nc.a_float(datos.get(clave))
        datos["n_empleados"] = nc.a_int(datos.get("n_empleados"))
        periodos.append(datos)
    conocidos = {p["periodo"] for p in periodos}
    for fila in nc._consultar("SELECT DISTINCT SUBSTRING(fecha, 1, 7) FROM nom_marcaciones "
                           "ORDER BY 1 DESC") or []:
        periodo = fila[0]
        if periodo and periodo not in conocidos:
            periodos.append({"periodo": periodo, "estado": "SIN CALCULAR", "fecha_calculo": "",
                             "usuario": "", "total_ingresos": 0, "total_descuentos": 0, "total_neto": 0,
                             "total_aportes": 0, "n_empleados": 0})
            conocidos.add(periodo)
    periodos.sort(key=lambda p: p["periodo"], reverse=True)
    return periodos


def detalle_planilla(periodo, dni=None):
    columnas = ["dni", "empleado", "concepto", "tipo", "categoria", "base", "cantidad", "unidad",
                "monto", "formula", "orden"]
    sql = ("SELECT dni, empleado, concepto, tipo, categoria, base, cantidad, unidad, monto, formula, "
           "orden FROM nom_planilla_detalle WHERE periodo = %s ")
    params = [periodo]
    if dni:
        sql += "AND dni = %s "
        params.append(str(dni))
    sql += "ORDER BY empleado, tipo DESC, orden"
    filas = []
    for fila in nc._consultar(sql, tuple(params)) or []:
        registro = dict(zip(columnas, fila))
        # PostgreSQL devuelve NUMERIC como Decimal: se normaliza a float para la GUI
        for clave in ("base", "cantidad", "monto"):
            registro[clave] = nc.a_float(registro.get(clave))
        filas.append(registro)
    return filas


def resumen_planilla(periodo):
    """Totales por empleado de la planilla calculada (para la tabla principal)."""
    filas = nc._consultar(
        "SELECT dni, MAX(empleado), "
        "COALESCE(SUM(CASE WHEN tipo = 'INGRESO' THEN monto ELSE 0 END),0), "
        "COALESCE(SUM(CASE WHEN tipo = 'DESCUENTO' THEN monto ELSE 0 END),0), "
        "COALESCE(SUM(CASE WHEN tipo = 'APORTE' THEN monto ELSE 0 END),0) "
        "FROM nom_planilla_detalle WHERE periodo = %s GROUP BY dni ORDER BY MAX(empleado)", (periodo,)) or []
    resultado = []
    for dni, empleado, ingresos, descuentos, aportes in filas:
        resultado.append({"dni": dni, "empleado": empleado,
                          "ingresos": nc.redondear(ingresos), "descuentos": nc.redondear(descuentos),
                          "neto": nc.redondear(nc.a_float(ingresos) - nc.a_float(descuentos)),
                          "aportes": nc.redondear(aportes)})
    return resultado


def boleta_empleado(periodo, dni):
    """Estructura completa de la boleta de pago de un empleado."""
    empleado = nc.obtener_empleado(dni) or {"dni": dni, "nombre": dni}
    planilla = obtener_planilla(periodo) or {"periodo": periodo, "estado": "SIN CALCULAR"}
    lineas = detalle_planilla(periodo, dni)
    ingresos = [l for l in lineas if l["tipo"] == TIPO_INGRESO]
    descuentos = [l for l in lineas if l["tipo"] == TIPO_DESCUENTO]
    aportes = [l for l in lineas if l["tipo"] == TIPO_APORTE]
    total_ingresos = nc.redondear(sum(l["monto"] for l in ingresos))
    total_descuentos = nc.redondear(sum(l["monto"] for l in descuentos))
    total_aportes = nc.redondear(sum(l["monto"] for l in aportes))
    dias = nc.dias_del_periodo(periodo)
    asistencia = nc.resumen_asistencia(dias[0], dias[-1], dnis=[dni]).get(str(dni), {}) if dias else {}
    return {
        "periodo": periodo, "periodo_nombre": nc.nombre_mes(periodo), "estado": planilla.get("estado"),
        "empleado": empleado, "dni": str(dni),
        "ingresos": ingresos, "descuentos": descuentos, "aportes": aportes,
        "total_ingresos": total_ingresos, "total_descuentos": total_descuentos,
        "neto_pagar": nc.redondear(total_ingresos - total_descuentos), "total_aportes": total_aportes,
        "asistencia": asistencia,
        "valor_dia": nc.redondear(valor_dia(empleado.get("sueldo_basico"), nc.obtener_parametros()), 2),
    }


def cambiar_estado_planilla(periodo, estado, usuario="sistema"):
    if estado not in ESTADOS_PLANILLA:
        return False, "Estado no valido"
    ok, error = nc._ejecutar("UPDATE nom_planilla SET estado = %s WHERE periodo = %s", (estado, periodo))
    if ok:
        nc.registrar_auditoria(usuario, "Nomina", "Cambio la planilla de %s a %s" % (periodo, estado))
    return ok, error


def eliminar_planilla(periodo, usuario="sistema"):
    ok, error = nc._ejecutar("DELETE FROM nom_planilla WHERE periodo = %s", (periodo,))
    if ok:
        nc._ejecutar("DELETE FROM nom_planilla_detalle WHERE periodo = %s", (periodo,))
        nc.registrar_auditoria(usuario, "Nomina", "Elimino la planilla de %s" % periodo)
    return ok, error


# =========================================================================
# REPORTES
# =========================================================================
def reporte_tardanzas(periodo):
    """Tardanzas detalladas del periodo, ordenadas por minutos."""
    dias = nc.dias_del_periodo(periodo)
    if not dias:
        return []
    registros = nc.listar_asistencia(dias[0], dias[-1])
    empleados = {str(e["dni"]): e for e in nc.listar_empleados()}
    filas = []
    for registro in registros:
        minutos = nc.a_int(registro.get("minutos_tardanza"))
        if minutos <= 0 and nc.a_int(registro.get("minutos_anticipo")) <= 0:
            continue
        empleado = empleados.get(str(registro["dni"]), {})
        filas.append({
            "dni": registro["dni"], "empleado": empleado.get("nombre", registro["dni"]),
            "fecha": registro["fecha"], "entrada": registro.get("hora_entrada") or "",
            "salida": registro.get("hora_salida") or "",
            "minutos_tardanza": minutos, "minutos_anticipo": nc.a_int(registro.get("minutos_anticipo")),
            "estado": registro.get("estado"),
        })
    filas.sort(key=lambda f: (f["minutos_tardanza"] + f["minutos_anticipo"]), reverse=True)
    return filas


def reporte_faltas(periodo):
    """Faltas e inasistencias del periodo."""
    dias = nc.dias_del_periodo(periodo)
    if not dias:
        return []
    registros = nc.listar_asistencia(dias[0], dias[-1], estados=[nc.EST_FALTA])
    empleados = {str(e["dni"]): e for e in nc.listar_empleados()}
    return [{"dni": r["dni"], "empleado": empleados.get(str(r["dni"]), {}).get("nombre", r["dni"]),
             "fecha": r["fecha"], "dia": nc.DIAS_SEMANA[nc.dia_semana(r["fecha"]) - 1],
             "observacion": r.get("observacion") or "Sin marcacion registrada"}
            for r in registros]


def reporte_horas_extra(periodo):
    dias = nc.dias_del_periodo(periodo)
    if not dias:
        return []
    registros = nc.listar_horas_extra(desde=dias[0], hasta=dias[-1])
    empleados = {str(e["dni"]): e for e in nc.listar_empleados()}
    parametros = nc.obtener_parametros()
    limite_25 = nc.a_float(parametros.get("he_desde_hora"), 2) * 60 or 120
    recargo_noche = nc.a_int(parametros.get("recargo_nocturno_pct"), 35)
    filas = []
    for registro in registros:
        empleado = empleados.get(str(registro["dni"]), {})
        v_hora = valor_hora(empleado.get("sueldo_basico"), empleado.get("jornada_horas"), parametros)
        minutos = nc.a_int(registro["minutos"])
        nocturna = str(registro.get("tipo", "")).startswith("NOCTURNA")
        if nocturna:
            monto = (minutos / 60.0) * v_hora * (1 + (nc.a_int(parametros.get("he_25_pct"), 25)
                                                      + recargo_noche) / 100.0)
        else:
            # 2 primeras horas al 25% y las siguientes al 35%
            monto = ((min(minutos, limite_25) / 60.0) * v_hora
                     * (1 + nc.a_int(parametros.get("he_25_pct"), 25) / 100.0))
            monto += ((max(0, minutos - limite_25) / 60.0) * v_hora
                      * (1 + nc.a_int(parametros.get("he_35_pct"), 35) / 100.0))
        filas.append({
            "dni": registro["dni"], "empleado": empleado.get("nombre", registro["dni"]),
            "fecha": registro["fecha"], "minutos": minutos,
            "horas": nc.redondear(minutos / 60.0, 2), "tipo": registro["tipo"],
            "origen": registro.get("origen", ""), "aprobado": registro.get("aprobado", True),
            "monto_estimado": nc.redondear(monto),
        })
    filas.sort(key=lambda f: (f["dni"], f["fecha"]))
    return filas


def reporte_puntualidad(periodo):
    """Ranking de puntualidad por empleado (base para el tablero de indicadores)."""
    dias = nc.dias_del_periodo(periodo)
    if not dias:
        return []
    resumen = nc.resumen_asistencia(dias[0], dias[-1])
    empleados = {str(e["dni"]): e for e in nc.listar_empleados()}
    filas = []
    for dni, datos in resumen.items():
        laborables = nc.a_int(datos.get("dias_laborables")) or 1
        filas.append({
            "dni": dni, "empleado": empleados.get(str(dni), {}).get("nombre", dni),
            "dias_laborables": nc.a_int(datos.get("dias_laborables")),
            "puntuales": nc.a_int(datos.get("puntuales")),
            "tardanzas": nc.a_int(datos.get("tardanzas")),
            "faltas": nc.a_int(datos.get("faltas")),
            "incompletos": nc.a_int(datos.get("incompletos")),
            "minutos_tardanza": nc.a_int(datos.get("minutos_tardanza")),
            "horas_extra": nc.redondear(nc.a_int(datos.get("minutos_extra")) / 60.0, 2),
            "horas_trabajadas": nc.redondear(nc.a_int(datos.get("minutos_trabajados")) / 60.0, 2),
            "porcentaje_puntualidad": nc.redondear(100.0 * nc.a_int(datos.get("puntuales")) / laborables, 1),
        })
    filas.sort(key=lambda f: f["porcentaje_puntualidad"], reverse=True)
    return filas


def reporte_marcaciones(desde, hasta, dni=None):
    """Bitacora cruda de marcaciones del reloj (para auditoria)."""
    columnas = ["dni", "codigo_reloj", "nombre_reloj", "fecha", "hora", "tipo_pase", "metodo", "fuente"]
    sql = ("SELECT dni, codigo_reloj, nombre_reloj, fecha, hora, tipo_pase, metodo, fuente "
           "FROM nom_marcaciones WHERE fecha BETWEEN %s AND %s ")
    params = [nc.fecha_iso(desde), nc.fecha_iso(hasta)]
    if dni:
        sql += "AND dni = %s "
        params.append(str(dni))
    sql += "ORDER BY fecha DESC, hora, nombre_reloj"
    return [dict(zip(columnas, fila)) for fila in (nc._consultar(sql, tuple(params)) or [])]


def exportar_planilla_excel(ruta, periodo):
    """Exporta planilla, detalle por concepto y boletas resumidas a un Excel."""
    planilla = obtener_planilla(periodo)
    if not planilla:
        return False, "No hay planilla calculada para %s" % periodo
    hojas = {}

    resumen = resumen_planilla(periodo)
    hojas["Resumen"] = (
        ["DNI", "Empleado", "Total Ingresos", "Total Descuentos", "Neto a Pagar", "Aportes Empleador"],
        [[r["dni"], r["empleado"], r["ingresos"], r["descuentos"], r["neto"], r["aportes"]]
         for r in resumen])
    totales = ["", "TOTALES", sum(r["ingresos"] for r in resumen),
               sum(r["descuentos"] for r in resumen), sum(r["neto"] for r in resumen),
               sum(r["aportes"] for r in resumen)]
    hojas["Resumen"][1].append(totales)

    detalle = detalle_planilla(periodo)
    hojas["Detalle por Concepto"] = (
        ["DNI", "Empleado", "Concepto", "Tipo", "Categoria", "Base", "Cantidad", "Unidad", "Monto",
         "Formula"],
        [[d["dni"], d["empleado"], d["concepto"], d["tipo"], d["categoria"], d["base"], d["cantidad"],
          d["unidad"], d["monto"], d["formula"]] for d in detalle])

    puntualidad = reporte_puntualidad(periodo)
    hojas["Puntualidad"] = (
        ["DNI", "Empleado", "Dias laborables", "Puntuales", "Tardanzas", "Faltas", "Incompletos",
         "Minutos tardanza", "Horas extra", "Horas trabajadas", "% Puntualidad"],
        [[p["dni"], p["empleado"], p["dias_laborables"], p["puntuales"], p["tardanzas"], p["faltas"],
          p["incompletos"], p["minutos_tardanza"], p["horas_extra"], p["horas_trabajadas"],
          p["porcentaje_puntualidad"]] for p in puntualidad])

    extra = reporte_horas_extra(periodo)
    if extra:
        hojas["Horas Extra"] = (
            ["DNI", "Empleado", "Fecha", "Minutos", "Horas", "Tipo", "Origen", "Aprobado",
             "Monto estimado"],
            [[e["dni"], e["empleado"], e["fecha"], e["minutos"], e["horas"], e["tipo"], e["origen"],
              "SI" if e["aprobado"] else "NO", e["monto_estimado"]] for e in extra])

    tardanzas = reporte_tardanzas(periodo)
    if tardanzas:
        hojas["Tardanzas"] = (
            ["DNI", "Empleado", "Fecha", "Entrada", "Salida", "Min. tardanza", "Min. anticipo", "Estado"],
            [[t["dni"], t["empleado"], t["fecha"], t["entrada"], t["salida"], t["minutos_tardanza"],
              t["minutos_anticipo"], t["estado"]] for t in tardanzas])

    faltas = reporte_faltas(periodo)
    if faltas:
        hojas["Faltas"] = (
            ["DNI", "Empleado", "Fecha", "Dia", "Observacion"],
            [[f["dni"], f["empleado"], f["fecha"], f["dia"], f["observacion"]] for f in faltas])
    return nc.exportar_excel(ruta, hojas)


def exportar_boletas_excel(ruta, periodo, dnis=None):
    """Exporta una hoja por empleado con su boleta de pago."""
    hojas = {}
    lista = dnis or [r["dni"] for r in resumen_planilla(periodo)]
    for dni in lista:
        boleta = boleta_empleado(periodo, dni)
        nombre_hoja = (str(boleta["empleado"].get("nombre") or dni))[:28]
        filas = [["Periodo", boleta["periodo_nombre"]],
                 ["Empleado", boleta["empleado"].get("nombre", "")],
                 ["DNI", boleta["dni"]],
                 ["Cargo", boleta["empleado"].get("cargo", "")],
                 ["Sueldo basico", boleta["empleado"].get("sueldo_basico", 0)],
                 ["Valor dia", boleta["valor_dia"]], []]
        filas.append(["INGRESOS", "Monto"])
        for linea in boleta["ingresos"]:
            filas.append([linea["concepto"], linea["monto"]])
        filas.append(["Total ingresos", boleta["total_ingresos"]])
        filas.append([])
        filas.append(["DESCUENTOS", "Monto"])
        for linea in boleta["descuentos"]:
            filas.append([linea["concepto"], linea["monto"]])
        filas.append(["Total descuentos", boleta["total_descuentos"]])
        filas.append([])
        filas.append(["NETO A PAGAR", boleta["neto_pagar"]])
        filas.append([])
        asistencia = boleta.get("asistencia") or {}
        filas.append(["Asistencia del periodo", ""])
        filas.append(["Dias laborables", asistencia.get("dias_laborables", 0)])
        filas.append(["Puntuales", asistencia.get("puntuales", 0)])
        filas.append(["Tardanzas", asistencia.get("tardanzas", 0)])
        filas.append(["Faltas", asistencia.get("faltas", 0)])
        filas.append(["Minutos de tardanza", asistencia.get("minutos_tardanza", 0)])
        filas.append(["Horas extra", nc.redondear(nc.a_int(asistencia.get("minutos_extra")) / 60.0, 2)])
        hojas[nombre_hoja] = (["Concepto", "Detalle"], filas)
    return nc.exportar_excel(ruta, hojas)


# =========================================================================
# PRUEBA DIRECTA
# =========================================================================
if __name__ == "__main__":
    pass
