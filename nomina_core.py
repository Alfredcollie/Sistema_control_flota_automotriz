# -*- coding: utf-8 -*-
"""
NOMINA_CORE.PY - Motor de Control de Asistencia y Nomina
=========================================================
Backend del modulo de Nomina del Sistema de Control de Flota Automotriz.

Responsabilidades:
  * Esquema de base de datos del modulo (tablas nom_*).
  * Catalogos: parametros, turnos, horarios, asignaciones, calendario,
    horas extra e incidencias.
  * Padron de empleados: toma la tabla 'choferes' (maestro existente) y
    guarda los datos de nomina en la tabla satelite 'nom_empleado'.
  * Importacion del Excel del reloj biometrico (captahuella) con deteccion
    automatica de formato (Hikvision / ZKTeco / generico) y mapeo manual.
  * Motor de asistencia: primera y ultima marcacion del dia, cruce con el
    turno/horario asignado y calculo de tardanzas, faltas, horas extra,
    descansos, feriados e incidencias.

Todas las fechas se guardan como texto ISO 'YYYY-MM-DD' y las horas como
texto 'HH:MM' (o 'HH:MM:SS' en las marcaciones crudas), respetando el estilo
del resto del sistema.

Autor: Collie Software
"""
import os
import re
import csv
import json
import hashlib
import unicodedata
import threading
from datetime import datetime, date, time, timedelta

try:
    from conexion import conectar_db, liberar_conexion, registrar_auditoria
except Exception:  # pragma: no cover - permite importar el modulo sin BD
    def conectar_db(silencioso=True):
        return None

    def liberar_conexion(conn):
        return None

    def registrar_auditoria(usuario, modulo, accion):
        return None

try:
    import openpyxl
    OPENPYXL_OK = True
except Exception:  # pragma: no cover
    openpyxl = None
    OPENPYXL_OK = False


# =========================================================================
# CONSTANTES
# =========================================================================
DIAS_SEMANA = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado", "Domingo"]
DIAS_CORTOS = ["Lun", "Mar", "Mie", "Jue", "Vie", "Sab", "Dom"]

# Estados de asistencia
EST_PUNTUAL = "PUNTUAL"
EST_TARDANZA = "TARDANZA"
EST_FALTA = "FALTA"
EST_INCOMPLETO = "INCOMPLETO"
EST_DESCANSO = "DESCANSO"
EST_DESCANSO_TRAB = "DESCANSO TRABAJADO"
EST_FERIADO = "FERIADO"
EST_FERIADO_TRAB = "FERIADO TRABAJADO"
EST_VACACIONES = "VACACIONES"
EST_PERMISO = "PERMISO"
EST_LICENCIA = "LICENCIA"
EST_DESCANSO_MEDICO = "DESCANSO MEDICO"
EST_COMISION = "COMISION"
EST_FALTA_JUST = "FALTA JUSTIFICADA"
EST_JUSTIFICADO = "JUSTIFICADO"
EST_PENDIENTE = "PENDIENTE"

ESTADOS_JUSTIFICADOS = {EST_VACACIONES, EST_PERMISO, EST_LICENCIA, EST_DESCANSO_MEDICO,
                        EST_COMISION, EST_FALTA_JUST, EST_JUSTIFICADO, EST_FERIADO}
ESTADOS_TRABAJADOS = {EST_PUNTUAL, EST_TARDANZA, EST_INCOMPLETO, EST_DESCANSO_TRAB, EST_FERIADO_TRAB}

TIPOS_INCIDENCIA = [EST_VACACIONES, EST_PERMISO, EST_LICENCIA, EST_DESCANSO_MEDICO,
                    EST_COMISION, EST_FALTA_JUST, EST_JUSTIFICADO]

COLORES_ESTADO = {
    EST_PUNTUAL: "#27ae60",
    EST_TARDANZA: "#e67e22",
    EST_FALTA: "#c0392b",
    EST_INCOMPLETO: "#8e44ad",
    EST_DESCANSO: "#7f8c8d",
    EST_DESCANSO_TRAB: "#16a085",
    EST_FERIADO: "#2980b9",
    EST_FERIADO_TRAB: "#1f538d",
    EST_VACACIONES: "#3498db",
    EST_PERMISO: "#f39c12",
    EST_LICENCIA: "#95a5a6",
    EST_DESCANSO_MEDICO: "#9b59b6",
    EST_COMISION: "#00bcd4",
    EST_FALTA_JUST: "#d35400",
    EST_JUSTIFICADO: "#f1c40f",
    EST_PENDIENTE: "#bdc3c7",
}

TIPOS_HORA_EXTRA = ["25", "35", "NOCTURNA_25", "NOCTURNA_35"]
ETIQUETA_HORA_EXTRA = {
    "25": "Extra 25% (2 primeras horas)",
    "35": "Extra 35% (siguientes horas)",
    "NOCTURNA_25": "Extra nocturna 25%",
    "NOCTURNA_35": "Extra nocturna 35%",
}

# Parametros por defecto del modulo (editables desde la GUI)
PARAMETROS_DEFECTO = {
    "tolerancia_min": ("10", "Minutos de tolerancia antes de considerar tardanza"),
    "refrigerio_min": ("45", "Minutos de refrigerio que se descuentan de la jornada"),
    "jornada_horas": ("8", "Horas de jornada laboral por defecto"),
    "dias_base_mes": ("30", "Dias base para el valor dia en planilla"),
    "rmv": ("1130", "Remuneracion Minima Vital vigente (S/)"),
    "uit": ("5500", "UIT vigente (S/)"),
    "onp_pct": ("13", "Aporte ONP del trabajador (%)"),
    "essalud_pct": ("9", "Aporte ESSALUD del empleador (%)"),
    "asignacion_familiar_pct": ("10", "Asignacion familiar (% de la RMV)"),
    "he_25_pct": ("25", "Recargo de las 2 primeras horas extra (%)"),
    "he_35_pct": ("35", "Recargo de las horas extra siguientes (%)"),
    "he_desde_hora": ("2", "Horas extra que se pagan al 25% antes de pasar al 35%"),
    "he_tope_dia_min": ("300", "Tope de minutos extra reconocidos por dia"),
    "he_minimo_min": ("15", "Minutos extra minimos para que se registren"),
    "he_bloque_min": ("15", "Redondeo de minutos extra (bloques de N minutos)"),
    "hora_inicio_nocturno": ("22:00", "Inicio del horario nocturno"),
    "hora_fin_nocturno": ("06:00", "Fin del horario nocturno"),
    "recargo_nocturno_pct": ("35", "Recargo por trabajo nocturno (%)"),
    "recargo_feriado_pct": ("100", "Recargo por trabajo en feriado (%)"),
    "descontar_tardanza": ("SI", "Descontar tardanzas en la planilla"),
    "descontar_falta": ("SI", "Descontar faltas injustificadas en la planilla"),
    "descontar_anticipo": ("SI", "Descontar salidas anticipadas en la planilla"),
    "calcular_renta_5ta": ("NO", "Calcular retencion de renta de quinta categoria"),
    "deduccion_uit_5ta": ("7", "UIT de deduccion para la renta de quinta categoria"),
    "pagar_descanso_trabajado": ("SI", "Pagar con recargo el trabajo en dia de descanso"),
    "minutos_minimos_entre_marcas": ("2", "Minutos minimos entre marcas: las mas cercanas se consideran la misma lectura repetida"),
}

# Clasificacion de los parametros para poder validarlos antes de guardarlos.
# Sin esta validacion, un campo vacio o mal escrito se guardaba tal cual y podia
# dejar el calculo en cero (por ejemplo un tope de horas extra en 0).
PARAMETROS_SOLO_POSITIVOS = {"rmv", "uit", "jornada_horas", "dias_base_mes", "he_bloque_min",
                            "he_minimo_min", "he_tope_dia_min"}
PARAMETROS_SI_NO = {"descontar_tardanza", "descontar_falta", "descontar_anticipo",
                    "calcular_renta_5ta", "pagar_descanso_trabajado"}
PARAMETROS_HORA = {"hora_inicio_nocturno", "hora_fin_nocturno"}
PARAMETROS_NO_CERO = {
    "he_tope_dia_min": "Con 0 minutos de tope NO se registraria ninguna hora extra",
    "he_minimo_min": "Con 0 no habria minimo para reconocer una hora extra",
    "he_bloque_min": "Con 0 no se podria redondear el tiempo extra",
    "he_desde_hora": "Con 0 todas las horas extra se pagarian con el recargo del 35%",
    "rmv": "La Remuneracion Minima Vital no puede ser 0",
    "uit": "La UIT no puede ser 0",
    "jornada_horas": "La jornada no puede ser 0",
    "dias_base_mes": "Los dias base del mes no pueden ser 0",
}

# Alias de columnas del Excel del reloj biometrico (normalizados sin acentos)
ALIAS_COLUMNAS = {
    "codigo": ["id", "no", "n", "nro", "numero", "numero de empleado", "codigo", "codigo de empleado",
               "codigo de persona", "employee id", "employee no", "person id", "personnel id",
               "enroll id", "enrollment id", "user id", "userid", "ac no", "acno", "ac number",
               "pin", "dni", "documento", "numero de documento", "documento de identidad",
               "card no", "no de tarjeta", "id de empleado", "identificador"],
    "nombre": ["nombre", "nombres", "name", "first name", "primer nombre", "given name"],
    "apellido": ["apellido", "apellidos", "last name", "surname", "apellido paterno"],
    "nombre_completo": ["nombre completo", "nombre y apellidos", "apellidos y nombres",
                        "employee name", "full name", "person name", "nombre del empleado",
                        "nombres y apellidos"],
    "fecha": ["fecha", "date", "fecha de registro", "fecha de marcacion", "att date",
              "punch date", "dia", "fecha de transaccion", "fecha y hora", "datetime", "date time"],
    "hora": ["tiempo", "hora", "time", "clock time", "hora de registro", "hora de marcacion",
             "punch time", "hora de transaccion", "hora de entrada", "check time"],
    "tipo_pase": ["tipo de pase de tarjeta", "tipo de pase", "pass type", "tipo", "tipo de evento",
                  "event type", "estado", "state", "sentido", "tipo de marcacion"],
    "metodo": ["metodo de verificacion", "verification mode", "verify mode", "metodo",
               "modo de verificacion"],
    "departamento": ["departamento", "department", "area", "seccion", "grupo"],
    "dispositivo": ["dispositivo", "device", "device name", "nombre del dispositivo", "equipo"],
}

PERFILES_FORMATO = {
    "HIKVISION": "Hikvision (Transacciones)",
    "ZKTECO": "ZKTeco / ZKTime (Registros de asistencia)",
    "GENERICO": "Generico / CSV",
}


# =========================================================================
# UTILIDADES GENERALES
# =========================================================================
def normalizar_texto(valor):
    """Mayusculas sin acentos ni signos, para comparar nombres y encabezados."""
    if valor is None:
        return ""
    texto = unicodedata.normalize("NFKD", str(valor))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"[^A-Za-z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip().upper()


def solo_digitos(valor):
    return re.sub(r"\D", "", str(valor or ""))


def normalizar_documento(valor):
    """DNI/documento comparable: solo digitos y sin ceros a la izquierda."""
    digitos = solo_digitos(valor)
    limpio = digitos.lstrip("0")
    return limpio if limpio else digitos


def parse_fecha(valor):
    """Convierte texto/date/datetime/serial de Excel a date. Devuelve None si no puede."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, (int, float)):
        try:
            return (datetime(1899, 12, 30) + timedelta(days=float(valor))).date()
        except Exception:
            return None
    texto = str(valor).strip()
    if not texto:
        return None
    # Recorta la parte de hora si viene junto a la fecha
    texto = texto.replace("T", " ")
    candidato = texto.split(" ")[0]
    formatos = ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y", "%d.%m.%Y",
                "%m/%d/%Y", "%Y%m%d", "%d/%m/%y")
    for fmt in formatos:
        try:
            return datetime.strptime(candidato, fmt).date()
        except ValueError:
            continue
    return None


def parse_hora(valor):
    """Convierte texto/time/datetime a 'HH:MM'. Devuelve None si no puede."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, datetime):
        return valor.strftime("%H:%M")
    if isinstance(valor, time):
        return valor.strftime("%H:%M")
    if isinstance(valor, (int, float)):
        # Fraccion de dia de Excel (0.5 = 12:00)
        try:
            total = int(round(float(valor) * 24 * 60)) % (24 * 60)
            return "%02d:%02d" % (total // 60, total % 60)
        except Exception:
            return None
    texto = str(valor).strip().upper()
    if not texto:
        return None
    # Si trae fecha y hora juntas, toma la hora
    if " " in texto:
        partes = texto.split(" ")
        for parte in reversed(partes):
            if ":" in parte:
                texto = parte
                break
    ampm = None
    for sufijo in ("A.M.", "P.M.", "AM", "PM", "A M", "P M"):
        if texto.endswith(sufijo):
            ampm = "PM" if "P" in sufijo else "AM"
            texto = texto[: -len(sufijo)].strip()
            break
    partes = re.split(r"[:.]", texto)
    try:
        hora = int(partes[0])
        minuto = int(partes[1]) if len(partes) > 1 else 0
    except (ValueError, IndexError):
        return None
    if ampm == "PM" and hora < 12:
        hora += 12
    if ampm == "AM" and hora == 12:
        hora = 0
    if not (0 <= hora <= 23 and 0 <= minuto <= 59):
        return None
    return "%02d:%02d" % (hora, minuto)


def minutos_de_hora(texto_hora):
    """'HH:MM' o 'HH:MM:SS' -> minutos desde la medianoche."""
    if not texto_hora:
        return None
    partes = str(texto_hora).split(":")
    try:
        hora = int(partes[0])
        minuto = int(partes[1]) if len(partes) > 1 else 0
    except (ValueError, IndexError):
        return None
    return hora * 60 + minuto


def hora_de_minutos(minutos):
    """Minutos desde la medianoche -> 'HH:MM' (acepta valores > 1440)."""
    if minutos is None:
        return None
    minutos = int(minutos) % (24 * 60)
    return "%02d:%02d" % (minutos // 60, minutos % 60)


def formato_duracion(minutos):
    """Minutos -> '8h 30m'."""
    try:
        minutos = int(minutos or 0)
    except (TypeError, ValueError):
        minutos = 0
    signo = "-" if minutos < 0 else ""
    minutos = abs(minutos)
    return "%s%dh %02dm" % (signo, minutos // 60, minutos % 60)


def fecha_iso(valor):
    f = parse_fecha(valor)
    return f.strftime("%Y-%m-%d") if f else None


def rango_fechas(desde, hasta):
    """Genera las fechas (date) entre dos extremos, inclusive."""
    f1, f2 = parse_fecha(desde), parse_fecha(hasta)
    if not f1 or not f2 or f2 < f1:
        return []
    dias = []
    actual = f1
    while actual <= f2:
        dias.append(actual)
        actual += timedelta(days=1)
    return dias


def dias_del_periodo(periodo):
    """'YYYY-MM' -> lista de fechas del mes."""
    try:
        anio, mes = [int(x) for x in str(periodo).split("-")[:2]]
        primero = date(anio, mes, 1)
    except Exception:
        return []
    ultimo = date(anio + (mes // 12), (mes % 12) + 1, 1) - timedelta(days=1)
    return rango_fechas(primero, ultimo)


def periodo_de_fecha(valor):
    f = parse_fecha(valor)
    return f.strftime("%Y-%m") if f else None


def nombre_mes(periodo):
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
             "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre"]
    try:
        anio, mes = [int(x) for x in str(periodo).split("-")[:2]]
        return "%s %d" % (meses[mes - 1], anio)
    except Exception:
        return str(periodo)


def dia_semana(fecha):
    """1 = lunes ... 7 = domingo."""
    f = parse_fecha(fecha)
    return f.isoweekday() if f else None


def parsear_dias_semana(texto):
    """'1,2,3,4,5' o 'LUN-VIE' -> set de enteros 1..7."""
    if not texto:
        return set()
    if isinstance(texto, (list, tuple, set)):
        return {int(x) for x in texto if str(x).strip().isdigit()}
    texto = str(texto)
    dias = {int(x) for x in re.findall(r"[1-7]", texto)}
    if dias:
        return dias
    equivalencias = {"LUN": 1, "MAR": 2, "MIE": 3, "JUE": 4, "VIE": 5, "SAB": 6, "DOM": 7}
    return {v for k, v in equivalencias.items() if k in normalizar_texto(texto)}


def dias_semana_texto(dias):
    """set de enteros -> 'Lun a Vie' compacto."""
    if not dias:
        return "Sin dias"
    dias = sorted(int(d) for d in dias)
    if dias == list(range(1, 8)):
        return "Todos los dias"
    if dias == [1, 2, 3, 4, 5]:
        return "Lunes a Viernes"
    if dias == [1, 2, 3, 4, 5, 6]:
        return "Lunes a Sabado"
    return ", ".join(DIAS_CORTOS[d - 1] for d in dias)


def redondear(valor, decimales=2):
    try:
        return round(float(valor or 0), decimales)
    except (TypeError, ValueError):
        return 0.0


def a_float(valor, por_defecto=0.0):
    if valor is None or valor == "":
        return por_defecto
    try:
        return float(str(valor).replace(",", "").replace("S/", "").strip())
    except (TypeError, ValueError):
        return por_defecto


def a_int(valor, por_defecto=0):
    try:
        return int(round(a_float(valor, por_defecto)))
    except (TypeError, ValueError):
        return por_defecto


def _si_no(valor, por_defecto=False):
    if valor is None:
        return por_defecto
    if isinstance(valor, bool):
        return valor
    return normalizar_texto(valor) in ("SI", "S", "YES", "Y", "TRUE", "VERDADERO", "1", "X")


def nombre_archivo_seguro(texto):
    limpio = re.sub(r"[^A-Za-z0-9_\-]+", "_", str(texto or ""))
    return limpio.strip("_") or "archivo"


# =========================================================================
# ACCESO A DATOS (atajos)
# =========================================================================
def _consultar(sql, params=None, uno=False, muchos=False):
    """Ejecuta un SELECT. Devuelve fila unica, lista de filas o [] / None."""
    conn = conectar_db(silencioso=True)
    if not conn:
        return None if uno else []
    try:
        with conn.cursor() as cursor:
            cursor.execute(sql, params or ())
            if uno:
                return cursor.fetchone()
            return cursor.fetchall()
    except Exception as e:
        print("[Nomina] Error en consulta:", e)
        return None if uno else []
    finally:
        liberar_conexion(conn)


def _ejecutar(sql, params=None):
    """Ejecuta INSERT/UPDATE/DELETE. Devuelve (exito, mensaje)."""
    conn = conectar_db(silencioso=True)
    if not conn:
        return False, "Sin conexion a la base de datos"
    try:
        with conn.cursor() as cursor:
            cursor.execute(sql, params or ())
        conn.commit()
        return True, ""
    except Exception as e:
        try:
            conn.rollback()
        except Exception:
            pass
        return False, str(e)
    finally:
        liberar_conexion(conn)


def _ejecutar_retorna_id(sql, params=None):
    conn = conectar_db(silencioso=True)
    if not conn:
        return None, "Sin conexion a la base de datos"
    try:
        with conn.cursor() as cursor:
            cursor.execute(sql, params or ())
            nuevo_id = None
            try:
                fila = cursor.fetchone()
                nuevo_id = fila[0] if fila else None
            except Exception:
                nuevo_id = None
        conn.commit()
        return nuevo_id, ""
    except Exception as e:
        try:
            conn.rollback()
        except Exception:
            pass
        return None, str(e)
    finally:
        liberar_conexion(conn)


def _ejecutar_lote(sql, filas):
    """Inserta/actualiza muchas filas en una sola transaccion."""
    if not filas:
        return True, 0, ""
    conn = conectar_db(silencioso=True)
    if not conn:
        return False, 0, "Sin conexion a la base de datos"
    try:
        with conn.cursor() as cursor:
            cursor.executemany(sql, filas)
            afectadas = cursor.rowcount
        conn.commit()
        if afectadas is None or afectadas < 0:
            afectadas = len(filas)
        return True, afectadas, ""
    except Exception as e:
        try:
            conn.rollback()
        except Exception:
            pass
        return False, 0, str(e)
    finally:
        liberar_conexion(conn)


# =========================================================================
# ESQUEMA DE BASE DE DATOS
# =========================================================================
_ESQUEMA_OK = False
_ESQUEMA_LOCK = threading.Lock()

TABLAS_NOMINA = [
    """
    CREATE TABLE IF NOT EXISTS nom_parametros (
        clave VARCHAR(80) PRIMARY KEY,
        valor TEXT,
        descripcion TEXT,
        actualizado VARCHAR(30)
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS nom_turnos (
        id SERIAL PRIMARY KEY,
        nombre VARCHAR(120) NOT NULL,
        hora_entrada VARCHAR(8) NOT NULL DEFAULT '08:00',
        hora_salida VARCHAR(8) NOT NULL DEFAULT '17:00',
        tolerancia_min INTEGER DEFAULT 10,
        refrigerio_min INTEGER DEFAULT 45,
        horas_jornada NUMERIC(5,2) DEFAULT 8,
        dias_semana VARCHAR(30) DEFAULT '1,2,3,4,5,6',
        cruza_medianoche BOOLEAN DEFAULT FALSE,
        aplica_nocturno BOOLEAN DEFAULT FALSE,
        color VARCHAR(20) DEFAULT '#1f538d',
        activo BOOLEAN DEFAULT TRUE,
        observacion TEXT
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS nom_horarios (
        id SERIAL PRIMARY KEY,
        nombre VARCHAR(120) NOT NULL,
        descripcion TEXT,
        activo BOOLEAN DEFAULT TRUE
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS nom_horario_detalle (
        id SERIAL PRIMARY KEY,
        horario_id INTEGER NOT NULL,
        dia_semana INTEGER NOT NULL,
        turno_id INTEGER,
        UNIQUE (horario_id, dia_semana)
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS nom_empleado (
        id SERIAL PRIMARY KEY,
        dni VARCHAR(20) UNIQUE NOT NULL,
        sueldo_basico NUMERIC(12,2) DEFAULT 0,
        asignacion_familiar BOOLEAN DEFAULT FALSE,
        jornada_horas NUMERIC(5,2) DEFAULT 8,
        sistema_pension VARCHAR(20) DEFAULT 'ONP',
        afp_nombre VARCHAR(60),
        afp_comision_pct NUMERIC(6,4) DEFAULT 0,
        cuspp VARCHAR(30),
        tipo_documento VARCHAR(5) DEFAULT '1',
        cargo VARCHAR(120),
        regimen VARCHAR(20) DEFAULT 'GENERAL',
        fecha_ingreso VARCHAR(20),
        horario_id INTEGER,
        turno_id INTEGER,
        banco_haberes VARCHAR(80),
        cuenta_haberes VARCHAR(60),
        essalud_codigo VARCHAR(40),
        discapacidad BOOLEAN DEFAULT FALSE,
        confianza BOOLEAN DEFAULT FALSE,
        otros_ingresos NUMERIC(12,2) DEFAULT 0,
        otros_descuentos NUMERIC(12,2) DEFAULT 0,
        retencion_judicial_pct NUMERIC(6,4) DEFAULT 0,
        adelanto_mensual NUMERIC(12,2) DEFAULT 0,
        estado VARCHAR(20) DEFAULT 'ACTIVO',
        observacion TEXT,
        actualizado VARCHAR(30)
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS nom_asignaciones (
        id SERIAL PRIMARY KEY,
        dni VARCHAR(20) NOT NULL,
        horario_id INTEGER,
        turno_id INTEGER,
        fecha_desde VARCHAR(10) NOT NULL,
        fecha_hasta VARCHAR(10),
        observacion TEXT,
        activo BOOLEAN DEFAULT TRUE,
        creado VARCHAR(30)
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS nom_calendario (
        id SERIAL PRIMARY KEY,
        fecha VARCHAR(10) UNIQUE NOT NULL,
        tipo VARCHAR(20) DEFAULT 'FERIADO',
        descripcion VARCHAR(150),
        creado VARCHAR(30)
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS nom_mapeo_reloj (
        id SERIAL PRIMARY KEY,
        codigo_reloj VARCHAR(40) UNIQUE NOT NULL,
        dni VARCHAR(20) NOT NULL,
        nombre_reloj VARCHAR(200),
        creado VARCHAR(30)
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS nom_marcaciones (
        id SERIAL PRIMARY KEY,
        codigo_reloj VARCHAR(40),
        dni VARCHAR(20),
        nombre_reloj VARCHAR(200),
        fecha VARCHAR(10) NOT NULL,
        hora VARCHAR(8) NOT NULL,
        tipo_pase VARCHAR(60),
        metodo VARCHAR(60),
        departamento VARCHAR(120),
        fuente VARCHAR(300),
        hash_marca VARCHAR(64) UNIQUE,
        importado VARCHAR(30)
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS nom_asistencia (
        id SERIAL PRIMARY KEY,
        dni VARCHAR(20) NOT NULL,
        fecha VARCHAR(10) NOT NULL,
        turno_id INTEGER,
        hora_entrada VARCHAR(8),
        hora_salida VARCHAR(8),
        primera_marca VARCHAR(8),
        ultima_marca VARCHAR(8),
        n_marcas INTEGER DEFAULT 0,
        minutos_tardanza INTEGER DEFAULT 0,
        minutos_anticipo INTEGER DEFAULT 0,
        minutos_trabajados INTEGER DEFAULT 0,
        minutos_extra INTEGER DEFAULT 0,
        minutos_falta INTEGER DEFAULT 0,
        estado VARCHAR(30) DEFAULT 'PENDIENTE',
        manual BOOLEAN DEFAULT FALSE,
        incidencia_id INTEGER,
        observacion TEXT,
        actualizado VARCHAR(30),
        UNIQUE (dni, fecha)
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS nom_incidencias (
        id SERIAL PRIMARY KEY,
        dni VARCHAR(20) NOT NULL,
        fecha_desde VARCHAR(10) NOT NULL,
        fecha_hasta VARCHAR(10) NOT NULL,
        tipo VARCHAR(30) NOT NULL,
        motivo TEXT,
        con_goce BOOLEAN DEFAULT TRUE,
        horas NUMERIC(6,2),
        aprobado_por VARCHAR(120),
        archivo VARCHAR(300),
        creado VARCHAR(30)
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS nom_horas_extra (
        id SERIAL PRIMARY KEY,
        dni VARCHAR(20) NOT NULL,
        fecha VARCHAR(10) NOT NULL,
        minutos INTEGER DEFAULT 0,
        tipo VARCHAR(20) DEFAULT '25',
        origen VARCHAR(20) DEFAULT 'AUTO',
        aprobado BOOLEAN DEFAULT TRUE,
        observacion TEXT,
        creado VARCHAR(30),
        UNIQUE (dni, fecha)
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS nom_importaciones (
        id SERIAL PRIMARY KEY,
        archivo VARCHAR(300),
        hash_archivo VARCHAR(64),
        formato VARCHAR(60),
        total_filas INTEGER DEFAULT 0,
        nuevas INTEGER DEFAULT 0,
        duplicadas INTEGER DEFAULT 0,
        sin_empleado INTEGER DEFAULT 0,
        fecha_min VARCHAR(10),
        fecha_max VARCHAR(10),
        usuario VARCHAR(120),
        fecha VARCHAR(30),
        detalle TEXT
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS nom_planilla (
        id SERIAL PRIMARY KEY,
        periodo VARCHAR(7) UNIQUE NOT NULL,
        estado VARCHAR(20) DEFAULT 'BORRADOR',
        fecha_calculo VARCHAR(30),
        usuario VARCHAR(120),
        total_ingresos NUMERIC(14,2) DEFAULT 0,
        total_descuentos NUMERIC(14,2) DEFAULT 0,
        total_neto NUMERIC(14,2) DEFAULT 0,
        total_aportes NUMERIC(14,2) DEFAULT 0,
        n_empleados INTEGER DEFAULT 0,
        observacion TEXT
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS nom_planilla_detalle (
        id SERIAL PRIMARY KEY,
        periodo VARCHAR(7) NOT NULL,
        dni VARCHAR(20) NOT NULL,
        empleado VARCHAR(200),
        concepto VARCHAR(120) NOT NULL,
        tipo VARCHAR(20) NOT NULL,
        categoria VARCHAR(40),
        base NUMERIC(14,2) DEFAULT 0,
        cantidad NUMERIC(14,2) DEFAULT 0,
        unidad VARCHAR(20),
        monto NUMERIC(14,2) DEFAULT 0,
        formula TEXT,
        orden INTEGER DEFAULT 0
    )
    """,
]

INDICES_NOMINA = [
    "CREATE INDEX IF NOT EXISTS idx_nom_marcaciones_dni ON nom_marcaciones (dni, fecha)",
    "CREATE INDEX IF NOT EXISTS idx_nom_marcaciones_cod ON nom_marcaciones (codigo_reloj, fecha)",
    "CREATE INDEX IF NOT EXISTS idx_nom_asistencia_periodo ON nom_asistencia (fecha, dni)",
    "CREATE INDEX IF NOT EXISTS idx_nom_detalle_periodo ON nom_planilla_detalle (periodo, dni)",
    "CREATE INDEX IF NOT EXISTS idx_nom_asignaciones_dni ON nom_asignaciones (dni, fecha_desde)",
    "CREATE INDEX IF NOT EXISTS idx_nom_incidencias_dni ON nom_incidencias (dni, fecha_desde)",
]


def inicializar_esquema_nomina(forzar=False):
    """Crea las tablas e indices del modulo. Devuelve (ok, mensaje)."""
    global _ESQUEMA_OK
    if _ESQUEMA_OK and not forzar:
        return True, ""
    with _ESQUEMA_LOCK:
        if _ESQUEMA_OK and not forzar:
            return True, ""
        conn = conectar_db(silencioso=True)
        if not conn:
            return False, "Sin conexion a la base de datos"
        try:
            with conn.cursor() as cursor:
                for ddl in TABLAS_NOMINA:
                    cursor.execute(ddl)
                for indice in INDICES_NOMINA:
                    try:
                        cursor.execute(indice)
                    except Exception:
                        conn.rollback()
                conn.commit()
            _sembrar_parametros(conn)
            _sembrar_turnos_base(conn)
            _ESQUEMA_OK = True
            return True, ""
        except Exception as e:
            try:
                conn.rollback()
            except Exception:
                pass
            return False, str(e)
        finally:
            liberar_conexion(conn)


def _sembrar_parametros(conn):
    """Inserta los parametros por defecto que aun no existan."""
    try:
        with conn.cursor() as cursor:
            for clave, (valor, descripcion) in PARAMETROS_DEFECTO.items():
                cursor.execute(
                    "INSERT INTO nom_parametros (clave, valor, descripcion, actualizado) "
                    "VALUES (%s, %s, %s, %s) ON CONFLICT (clave) DO NOTHING",
                    (clave, valor, descripcion, datetime.now().strftime("%d/%m/%Y %H:%M"))
                )
        conn.commit()
    except Exception as e:
        print("[Nomina] Error sembrando parametros:", e)
        try:
            conn.rollback()
        except Exception:
            pass


def _sembrar_turnos_base(conn):
    """Crea turnos tipicos la primera vez (solo si no hay ninguno)."""
    try:
        with conn.cursor() as cursor:
            cursor.execute("SELECT COUNT(*) FROM nom_turnos")
            if cursor.fetchone()[0] > 0:
                return
            base = [
                ("Turno Dia (8:00 - 17:00)", "08:00", "17:00", 10, 45, 8, "1,2,3,4,5,6", False, False, "#1f538d"),
                ("Turno Manana (6:00 - 14:00)", "06:00", "14:00", 10, 30, 8, "1,2,3,4,5,6", False, False, "#27ae60"),
                ("Turno Tarde (14:00 - 22:00)", "14:00", "22:00", 10, 30, 8, "1,2,3,4,5,6", False, False, "#e67e22"),
                ("Turno Noche (22:00 - 06:00)", "22:00", "06:00", 10, 45, 8, "1,2,3,4,5,6", True, True, "#2c3e50"),
                ("Turno Partido (8:00 - 13:00 / 15:00 - 19:00)", "08:00", "19:00", 10, 120, 9, "1,2,3,4,5,6", False, False, "#8e44ad"),
            ]
            for fila in base:
                cursor.execute(
                    "INSERT INTO nom_turnos (nombre, hora_entrada, hora_salida, tolerancia_min, "
                    "refrigerio_min, horas_jornada, dias_semana, cruza_medianoche, aplica_nocturno, color) "
                    "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", fila
                )
            conn.commit()
    except Exception as e:
        print("[Nomina] Error sembrando turnos:", e)
        try:
            conn.rollback()
        except Exception:
            pass


# =========================================================================
# PARAMETROS
# =========================================================================
def obtener_parametros():
    """Devuelve el diccionario de parametros del modulo (con valores por defecto)."""
    datos = {clave: valor for clave, (valor, _) in PARAMETROS_DEFECTO.items()}
    filas = _consultar("SELECT clave, valor FROM nom_parametros")
    for clave, valor in filas or []:
        datos[clave] = valor
    return datos


def _texto_numero(numero):
    """Convierte un numero a texto corto ('30', '1130', '0.5')."""
    try:
        flotante = float(numero)
    except (TypeError, ValueError):
        return str(numero)
    if flotante == int(flotante):
        return str(int(flotante))
    return ("%.4f" % flotante).rstrip("0").rstrip(".")


def validar_parametros(diccionario):
    """Revisa los parametros antes de guardarlos.

    Devuelve (limpios, errores, advertencias):
      * limpios      : {clave: valor normalizado} listo para guardar.
      * errores      : lista de textos con el problema de cada campo (no se guarda nada).
      * advertencias : valores validos pero probablemente equivocados (por ejemplo un 0
                       en el tope de horas extra). Se guardan, pero se avisa.
    """
    limpios, errores, advertencias = {}, [], []
    for clave, valor in (diccionario or {}).items():
        if clave not in PARAMETROS_DEFECTO:
            continue
        texto = "" if valor is None else str(valor).strip()
        if texto == "":
            errores.append("%s: el campo no puede quedar vacio" % clave)
            continue
        if clave in PARAMETROS_SI_NO:
            marcado = normalizar_texto(texto)
            if marcado in ("SI", "S", "YES", "Y", "TRUE", "VERDADERO", "1"):
                limpios[clave] = "SI"
            elif marcado in ("NO", "N", "FALSE", "FALSO", "0"):
                limpios[clave] = "NO"
            else:
                errores.append("%s: escriba SI o NO (valor recibido: %s)" % (clave, texto))
            continue
        if clave in PARAMETROS_HORA:
            hora = parse_hora(texto)
            if not hora:
                errores.append("%s: escriba la hora como HH:MM (valor recibido: %s)" % (clave, texto))
            else:
                limpios[clave] = hora
            continue
        numero = a_float(texto, None)
        if numero is None:
            errores.append("%s: debe ser un numero (valor recibido: %s)" % (clave, texto))
            continue
        if numero < 0:
            errores.append("%s: no puede ser negativo" % clave)
            continue
        if clave in PARAMETROS_SOLO_POSITIVOS and numero <= 0:
            errores.append("%s: debe ser mayor que cero" % clave)
            continue
        if numero == 0 and clave in PARAMETROS_NO_CERO:
            advertencias.append("%s = 0. %s." % (clave, PARAMETROS_NO_CERO[clave]))
        limpios[clave] = _texto_numero(numero)
    return limpios, errores, advertencias


def reparar_parametros(usuario="sistema"):
    """Restaura a su valor por defecto los parametros guardados vacios o invalidos.

    Devuelve la lista de correcciones [(clave, valor_anterior, valor_nuevo), ...].
    """
    actuales = obtener_parametros()
    correcciones = {}
    detalle = []
    for clave in PARAMETROS_DEFECTO:
        limpios, errores, _advertencias = validar_parametros({clave: actuales.get(clave)})
        if errores:
            valor_defecto = PARAMETROS_DEFECTO[clave][0]
            correcciones[clave] = valor_defecto
            detalle.append((clave, actuales.get(clave), valor_defecto))
    if correcciones:
        guardar_parametros(correcciones, usuario)
        registrar_auditoria(usuario, "Nomina", "Reparo %d parametros invalidos de nomina" % len(detalle))
    return detalle


def guardar_parametros(diccionario, usuario="sistema"):
    """Guarda los parametros. Rechaza la operacion completa si algun valor es invalido."""
    limpios, errores, advertencias = validar_parametros(diccionario)
    if errores:
        return False, "Corrija estos campos antes de guardar:\n- " + "\n- ".join(errores)
    if not limpios:
        return False, "No hay parametros que guardar"
    ahora = datetime.now().strftime("%d/%m/%Y %H:%M")
    filas = [(clave, str(valor), ahora) for clave, valor in limpios.items()]
    ok, _, error = _ejecutar_lote(
        "INSERT INTO nom_parametros (clave, valor, actualizado) VALUES (%s, %s, %s) "
        "ON CONFLICT (clave) DO UPDATE SET valor = EXCLUDED.valor, actualizado = EXCLUDED.actualizado",
        filas
    )
    if ok:
        registrar_auditoria(usuario, "Nomina", "Actualizo parametros de nomina")
    return ok, error


# =========================================================================
# TURNOS
# =========================================================================
def listar_turnos(solo_activos=False):
    sql = ("SELECT id, nombre, hora_entrada, hora_salida, tolerancia_min, refrigerio_min, "
           "horas_jornada, dias_semana, cruza_medianoche, aplica_nocturno, color, activo, observacion "
           "FROM nom_turnos ")
    if solo_activos:
        sql += "WHERE activo = TRUE "
    sql += "ORDER BY hora_entrada, nombre"
    columnas = ["id", "nombre", "hora_entrada", "hora_salida", "tolerancia_min", "refrigerio_min",
                "horas_jornada", "dias_semana", "cruza_medianoche", "aplica_nocturno", "color",
                "activo", "observacion"]
    return [dict(zip(columnas, fila)) for fila in (_consultar(sql) or [])]


def obtener_turno(turno_id):
    if not turno_id:
        return None
    for turno in listar_turnos():
        if int(turno["id"]) == int(turno_id):
            return turno
    return None


def guardar_turno(datos, usuario="sistema"):
    dias = datos.get("dias_semana") or "1,2,3,4,5,6"
    if isinstance(dias, (list, tuple, set)):
        dias = ",".join(str(d) for d in sorted(int(x) for x in dias))
    campos = (
        datos.get("nombre", "").strip(),
        datos.get("hora_entrada", "08:00"),
        datos.get("hora_salida", "17:00"),
        a_int(datos.get("tolerancia_min", 10), 10),
        a_int(datos.get("refrigerio_min", 45), 45),
        a_float(datos.get("horas_jornada", 8), 8),
        dias,
        bool(datos.get("cruza_medianoche", False)),
        bool(datos.get("aplica_nocturno", False)),
        datos.get("color", "#1f538d"),
        bool(datos.get("activo", True)),
        datos.get("observacion", ""),
    )
    if not campos[0]:
        return None, "El nombre del turno es obligatorio"
    if datos.get("id"):
        sql = ("UPDATE nom_turnos SET nombre=%s, hora_entrada=%s, hora_salida=%s, tolerancia_min=%s, "
               "refrigerio_min=%s, horas_jornada=%s, dias_semana=%s, cruza_medianoche=%s, "
               "aplica_nocturno=%s, color=%s, activo=%s, observacion=%s WHERE id=%s")
        ok, error = _ejecutar(sql, campos + (int(datos["id"]),))
        if ok:
            registrar_auditoria(usuario, "Nomina", "Edito turno %s" % campos[0])
            return int(datos["id"]), ""
        return None, error
    sql = ("INSERT INTO nom_turnos (nombre, hora_entrada, hora_salida, tolerancia_min, refrigerio_min, "
           "horas_jornada, dias_semana, cruza_medianoche, aplica_nocturno, color, activo, observacion) "
           "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id")
    nuevo_id, error = _ejecutar_retorna_id(sql, campos)
    if nuevo_id:
        registrar_auditoria(usuario, "Nomina", "Creo turno %s" % campos[0])
    return nuevo_id, error


def eliminar_turno(turno_id, usuario="sistema"):
    ok, error = _ejecutar("DELETE FROM nom_turnos WHERE id = %s", (int(turno_id),))
    if ok:
        _ejecutar("UPDATE nom_horario_detalle SET turno_id = NULL WHERE turno_id = %s", (int(turno_id),))
        _ejecutar("UPDATE nom_empleado SET turno_id = NULL WHERE turno_id = %s", (int(turno_id),))
        registrar_auditoria(usuario, "Nomina", "Elimino turno id %s" % turno_id)
    return ok, error


# =========================================================================
# HORARIOS (plantilla semanal)
# =========================================================================
def listar_horarios(solo_activos=False):
    sql = "SELECT id, nombre, descripcion, activo FROM nom_horarios "
    if solo_activos:
        sql += "WHERE activo = TRUE "
    sql += "ORDER BY nombre"
    horarios = []
    for fila in _consultar(sql) or []:
        horarios.append({"id": fila[0], "nombre": fila[1], "descripcion": fila[2] or "",
                         "activo": fila[3], "detalle": {}})
    return horarios


def obtener_horario(horario_id):
    if not horario_id:
        return None
    fila = _consultar("SELECT id, nombre, descripcion, activo FROM nom_horarios WHERE id = %s",
                      (int(horario_id),), uno=True)
    if not fila:
        return None
    detalle = {}
    for dia, turno_id in _consultar(
            "SELECT dia_semana, turno_id FROM nom_horario_detalle WHERE horario_id = %s", (int(horario_id),)) or []:
        detalle[int(dia)] = turno_id
    return {"id": fila[0], "nombre": fila[1], "descripcion": fila[2] or "", "activo": fila[3],
            "detalle": detalle}


def guardar_horario(datos, detalle=None, usuario="sistema"):
    """detalle: {1..7: turno_id o None}. None = dia de descanso."""
    nombre = (datos.get("nombre") or "").strip()
    if not nombre:
        return None, "El nombre del horario es obligatorio"
    campos = (nombre, datos.get("descripcion", ""), bool(datos.get("activo", True)))
    if datos.get("id"):
        horario_id = int(datos["id"])
        ok, error = _ejecutar(
            "UPDATE nom_horarios SET nombre=%s, descripcion=%s, activo=%s WHERE id=%s",
            campos + (horario_id,))
        if not ok:
            return None, error
    else:
        horario_id, error = _ejecutar_retorna_id(
            "INSERT INTO nom_horarios (nombre, descripcion, activo) VALUES (%s,%s,%s) RETURNING id", campos)
        if not horario_id:
            return None, error
    if detalle is not None:
        filas = [(horario_id, int(dia), turno_id) for dia, turno_id in detalle.items()]
        ok_det, _, error_det = _ejecutar_lote(
            "INSERT INTO nom_horario_detalle (horario_id, dia_semana, turno_id) VALUES (%s,%s,%s) "
            "ON CONFLICT (horario_id, dia_semana) DO UPDATE SET turno_id = EXCLUDED.turno_id", filas)
        if not ok_det:
            return None, error_det
    registrar_auditoria(usuario, "Nomina", "Guardo horario %s" % nombre)
    return horario_id, ""


def eliminar_horario(horario_id, usuario="sistema"):
    ok, error = _ejecutar("DELETE FROM nom_horarios WHERE id = %s", (int(horario_id),))
    if ok:
        _ejecutar("DELETE FROM nom_horario_detalle WHERE horario_id = %s", (int(horario_id),))
        _ejecutar("UPDATE nom_empleado SET horario_id = NULL WHERE horario_id = %s", (int(horario_id),))
        registrar_auditoria(usuario, "Nomina", "Elimino horario id %s" % horario_id)
    return ok, error


def duplicar_horario(horario_id, nuevo_nombre, usuario="sistema"):
    original = obtener_horario(horario_id)
    if not original:
        return None, "Horario no encontrado"
    return guardar_horario({"nombre": nuevo_nombre, "descripcion": original["descripcion"],
                            "activo": True}, original["detalle"], usuario)


# =========================================================================
# ASIGNACIONES DE HORARIO / TURNO POR RANGO DE FECHAS
# =========================================================================
def listar_asignaciones(dni=None):
    columnas = ["id", "dni", "horario_id", "turno_id", "fecha_desde", "fecha_hasta",
                "observacion", "activo", "creado"]
    sql = ("SELECT id, dni, horario_id, turno_id, fecha_desde, fecha_hasta, observacion, activo, creado "
           "FROM nom_asignaciones ")
    params = None
    if dni:
        sql += "WHERE dni = %s "
        params = (str(dni),)
    sql += "ORDER BY fecha_desde DESC, dni"
    return [dict(zip(columnas, fila)) for fila in (_consultar(sql, params) or [])]


def guardar_asignacion(datos, usuario="sistema"):
    dni = str(datos.get("dni") or "").strip()
    desde = fecha_iso(datos.get("fecha_desde"))
    hasta = fecha_iso(datos.get("fecha_hasta")) if datos.get("fecha_hasta") else None
    if not dni:
        return None, "Debe indicar el empleado (DNI)"
    if not desde:
        return None, "La fecha de inicio no es valida"
    horario_id = int(datos["horario_id"]) if datos.get("horario_id") else None
    turno_id = int(datos["turno_id"]) if datos.get("turno_id") else None
    if not horario_id and not turno_id:
        return None, "Debe seleccionar un horario o un turno"
    campos = (dni, horario_id, turno_id, desde, hasta, datos.get("observacion", ""),
              bool(datos.get("activo", True)), datetime.now().strftime("%d/%m/%Y %H:%M"))
    if datos.get("id"):
        ok, error = _ejecutar(
            "UPDATE nom_asignaciones SET dni=%s, horario_id=%s, turno_id=%s, fecha_desde=%s, "
            "fecha_hasta=%s, observacion=%s, activo=%s WHERE id=%s", campos + (int(datos["id"]),))
        if not ok:
            return None, error
        registrar_auditoria(usuario, "Nomina", "Edito asignacion de %s" % dni)
        return int(datos["id"]), ""
    nuevo_id, error = _ejecutar_retorna_id(
        "INSERT INTO nom_asignaciones (dni, horario_id, turno_id, fecha_desde, fecha_hasta, "
        "observacion, activo, creado) VALUES (%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id", campos)
    if nuevo_id:
        registrar_auditoria(usuario, "Nomina", "Asigno horario a %s desde %s" % (dni, desde))
    return nuevo_id, error


def eliminar_asignacion(asignacion_id, usuario="sistema"):
    ok, error = _ejecutar("DELETE FROM nom_asignaciones WHERE id = %s", (int(asignacion_id),))
    if ok:
        registrar_auditoria(usuario, "Nomina", "Elimino asignacion id %s" % asignacion_id)
    return ok, error


# =========================================================================
# EMPLEADOS (maestro: tabla choferes + datos de nomina en nom_empleado)
# =========================================================================
CAMPOS_NOM_EMPLEADO = [
    "dni", "sueldo_basico", "asignacion_familiar", "jornada_horas", "sistema_pension",
    "afp_nombre", "afp_comision_pct", "cuspp", "tipo_documento", "cargo", "regimen",
    "fecha_ingreso", "horario_id", "turno_id", "banco_haberes", "cuenta_haberes",
    "essalud_codigo", "discapacidad", "confianza", "otros_ingresos", "otros_descuentos",
    "retencion_judicial_pct", "adelanto_mensual", "estado", "observacion",
]


def listar_empleados(solo_activos=False, incluir_sin_chofer=True):
    """Padron de empleados: choferes (maestro) + datos de nomina.

    Devuelve una lista de diccionarios con los datos de la tabla 'choferes'
    combinados con la tabla satelite 'nom_empleado'.
    """
    filas = _consultar(
        "SELECT c.dni, c.nombres, c.estado, c.telefono, c.correo, c.movil_asignado, "
        "       c.fecha_inicio_contrato, c.fecha_nacimiento, c.numero_hijos, c.sexo, "
        "       n.dni, n.sueldo_basico, n.asignacion_familiar, n.jornada_horas, n.sistema_pension, "
        "       n.afp_nombre, n.afp_comision_pct, n.cuspp, n.tipo_documento, n.cargo, n.regimen, "
        "       n.fecha_ingreso, n.horario_id, n.turno_id, n.banco_haberes, n.cuenta_haberes, "
        "       n.essalud_codigo, n.discapacidad, n.confianza, n.otros_ingresos, n.otros_descuentos, "
        "       n.retencion_judicial_pct, n.adelanto_mensual, n.estado, n.observacion "
        "FROM choferes c LEFT JOIN nom_empleado n ON n.dni = c.dni "
        "ORDER BY c.nombres"
    ) or []
    empleados = []
    dnis_vistos = set()
    for fila in filas:
        dni = str(fila[0] or "").strip()
        dnis_vistos.add(dni)
        empleados.append({
            "dni": dni,
            "nombre": fila[1] or "",
            "estado_chofer": fila[2] or "Activo",
            "telefono": fila[3] or "",
            "correo": fila[4] or "",
            "movil_asignado": fila[5] or "",
            "fecha_inicio_contrato": fila[6] or "",
            "fecha_nacimiento": fila[7] or "",
            "numero_hijos": fila[8] or "0",
            "sexo": fila[9] or "",
            "en_nomina": bool(fila[10]),
            "sueldo_basico": a_float(fila[11]),
            "asignacion_familiar": bool(fila[12]),
            "jornada_horas": a_float(fila[13], 8) or 8,
            "sistema_pension": fila[14] or "ONP",
            "afp_nombre": fila[15] or "",
            "afp_comision_pct": a_float(fila[16]),
            "cuspp": fila[17] or "",
            "tipo_documento": fila[18] or "1",
            "cargo": fila[19] or "",
            "regimen": fila[20] or "GENERAL",
            "fecha_ingreso": fila[21] or "",
            "horario_id": fila[22],
            "turno_id": fila[23],
            "banco_haberes": fila[24] or "",
            "cuenta_haberes": fila[25] or "",
            "essalud_codigo": fila[26] or "",
            "discapacidad": bool(fila[27]),
            "confianza": bool(fila[28]),
            "otros_ingresos": a_float(fila[29]),
            "otros_descuentos": a_float(fila[30]),
            "retencion_judicial_pct": a_float(fila[31]),
            "adelanto_mensual": a_float(fila[32]),
            "estado": fila[33] or "ACTIVO",
            "observacion": fila[34] or "",
        })
    if incluir_sin_chofer:
        # Empleados creados directamente en nom_empleado (no son choferes)
        for fila in _consultar(
                "SELECT dni, sueldo_basico, asignacion_familiar, jornada_horas, sistema_pension, "
                "afp_nombre, afp_comision_pct, cuspp, cargo, regimen, fecha_ingreso, horario_id, "
                "turno_id, estado, otros_ingresos, otros_descuentos, retencion_judicial_pct, "
                "adelanto_mensual, observacion, banco_haberes, cuenta_haberes "
                "FROM nom_empleado ORDER BY dni") or []:
            dni = str(fila[0] or "").strip()
            if dni in dnis_vistos:
                continue
            empleados.append({
                "dni": dni, "nombre": (fila[8] or dni), "estado_chofer": "No es chofer",
                "telefono": "", "correo": "", "movil_asignado": "", "fecha_inicio_contrato": "",
                "fecha_nacimiento": "", "numero_hijos": "0", "sexo": "", "en_nomina": True,
                "sueldo_basico": a_float(fila[1]), "asignacion_familiar": bool(fila[2]),
                "jornada_horas": a_float(fila[3], 8) or 8, "sistema_pension": fila[4] or "ONP",
                "afp_nombre": fila[5] or "", "afp_comision_pct": a_float(fila[6]), "cuspp": fila[7] or "",
                "tipo_documento": "1", "cargo": fila[8] or "", "regimen": fila[9] or "GENERAL",
                "fecha_ingreso": fila[10] or "", "horario_id": fila[11], "turno_id": fila[12],
                "banco_haberes": fila[19] or "", "cuenta_haberes": fila[20] or "",
                "essalud_codigo": "", "discapacidad": False, "confianza": False,
                "otros_ingresos": a_float(fila[14]), "otros_descuentos": a_float(fila[15]),
                "retencion_judicial_pct": a_float(fila[16]), "adelanto_mensual": a_float(fila[17]),
                "estado": fila[13] or "ACTIVO", "observacion": fila[18] or "",
            })
        empleados.sort(key=lambda e: normalizar_texto(e["nombre"]))
    if solo_activos:
        empleados = [e for e in empleados if normalizar_texto(e["estado"]) not in ("CESADO", "INACTIVO")]
    return empleados


def sincronizar_empleados(usuario="sistema"):
    """Crea en nom_empleado una fila por cada chofer que aun no la tenga."""
    filas = _consultar("SELECT dni FROM choferes") or []
    dnis = [str(f[0]).strip() for f in filas if f[0]]
    if not dnis:
        return 0
    ahora = datetime.now().strftime("%d/%m/%Y %H:%M")
    datos = [(dni, ahora) for dni in dnis]
    ok, _, error = _ejecutar_lote(
        "INSERT INTO nom_empleado (dni, actualizado) VALUES (%s, %s) ON CONFLICT (dni) DO NOTHING", datos)
    if not ok:
        print("[Nomina] Error sincronizando empleados:", error)
        return 0
    registrar_auditoria(usuario, "Nomina", "Sincronizo padron de empleados desde choferes")
    return len(datos)


def obtener_empleado(dni):
    for empleado in listar_empleados(incluir_sin_chofer=True):
        if str(empleado["dni"]) == str(dni):
            return empleado
    return None


def guardar_empleado_nomina(dni, datos, usuario="sistema"):
    """Guarda (upsert) los datos de nomina de un empleado."""
    dni = str(dni or "").strip()
    if not dni:
        return False, "DNI requerido"
    valores = (
        a_float(datos.get("sueldo_basico", 0)),
        bool(datos.get("asignacion_familiar", False)),
        a_float(datos.get("jornada_horas", 8), 8),
        datos.get("sistema_pension", "ONP") or "ONP",
        datos.get("afp_nombre", ""),
        a_float(datos.get("afp_comision_pct", 0)),
        datos.get("cuspp", ""),
        datos.get("tipo_documento", "1") or "1",
        datos.get("cargo", ""),
        datos.get("regimen", "GENERAL") or "GENERAL",
        fecha_iso(datos.get("fecha_ingreso")) or datos.get("fecha_ingreso", "") or "",
        int(datos["horario_id"]) if datos.get("horario_id") else None,
        int(datos["turno_id"]) if datos.get("turno_id") else None,
        datos.get("banco_haberes", ""),
        datos.get("cuenta_haberes", ""),
        datos.get("essalud_codigo", ""),
        bool(datos.get("discapacidad", False)),
        bool(datos.get("confianza", False)),
        a_float(datos.get("otros_ingresos", 0)),
        a_float(datos.get("otros_descuentos", 0)),
        a_float(datos.get("retencion_judicial_pct", 0)),
        a_float(datos.get("adelanto_mensual", 0)),
        datos.get("estado", "ACTIVO") or "ACTIVO",
        datos.get("observacion", ""),
        datetime.now().strftime("%d/%m/%Y %H:%M"),
        dni,
    )
    sql = ("INSERT INTO nom_empleado (sueldo_basico, asignacion_familiar, jornada_horas, "
           "sistema_pension, afp_nombre, afp_comision_pct, cuspp, tipo_documento, cargo, regimen, "
           "fecha_ingreso, horario_id, turno_id, banco_haberes, cuenta_haberes, essalud_codigo, "
           "discapacidad, confianza, otros_ingresos, otros_descuentos, retencion_judicial_pct, "
           "adelanto_mensual, estado, observacion, actualizado, dni) "
           "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) "
           "ON CONFLICT (dni) DO UPDATE SET sueldo_basico=EXCLUDED.sueldo_basico, "
           "asignacion_familiar=EXCLUDED.asignacion_familiar, jornada_horas=EXCLUDED.jornada_horas, "
           "sistema_pension=EXCLUDED.sistema_pension, afp_nombre=EXCLUDED.afp_nombre, "
           "afp_comision_pct=EXCLUDED.afp_comision_pct, cuspp=EXCLUDED.cuspp, "
           "tipo_documento=EXCLUDED.tipo_documento, cargo=EXCLUDED.cargo, regimen=EXCLUDED.regimen, "
           "fecha_ingreso=EXCLUDED.fecha_ingreso, horario_id=EXCLUDED.horario_id, "
           "turno_id=EXCLUDED.turno_id, banco_haberes=EXCLUDED.banco_haberes, "
           "cuenta_haberes=EXCLUDED.cuenta_haberes, essalud_codigo=EXCLUDED.essalud_codigo, "
           "discapacidad=EXCLUDED.discapacidad, confianza=EXCLUDED.confianza, "
           "otros_ingresos=EXCLUDED.otros_ingresos, otros_descuentos=EXCLUDED.otros_descuentos, "
           "retencion_judicial_pct=EXCLUDED.retencion_judicial_pct, "
           "adelanto_mensual=EXCLUDED.adelanto_mensual, estado=EXCLUDED.estado, "
           "observacion=EXCLUDED.observacion, actualizado=EXCLUDED.actualizado")
    ok, error = _ejecutar(sql, valores)
    if ok:
        registrar_auditoria(usuario, "Nomina", "Actualizo datos de nomina de %s" % dni)
    return ok, error


def guardar_empleados_nomina_lote(lista_datos, usuario="sistema"):
    """Guarda varios empleados de una sola vez (usado por el asistente masivo)."""
    guardados = 0
    for datos in lista_datos:
        ok, _ = guardar_empleado_nomina(datos.get("dni"), datos, usuario)
        if ok:
            guardados += 1
    return guardados


def empleados_activos_cache():
    """{dni: empleado} de empleados vigentes, para cruces rapidos."""
    cache = {}
    for empleado in listar_empleados(solo_activos=True):
        cache[str(empleado["dni"])] = empleado
    return cache


# =========================================================================
# CALENDARIO (feriados y dias no laborables)
# =========================================================================
def listar_calendario(anio=None):
    columnas = ["id", "fecha", "tipo", "descripcion"]
    sql = "SELECT id, fecha, tipo, descripcion FROM nom_calendario "
    params = None
    if anio:
        sql += "WHERE fecha LIKE %s "
        params = (str(anio) + "-%",)
    sql += "ORDER BY fecha"
    return [dict(zip(columnas, fila)) for fila in (_consultar(sql, params) or [])]


def guardar_dia_calendario(fecha, tipo="FERIADO", descripcion="", usuario="sistema"):
    fecha_ok = fecha_iso(fecha)
    if not fecha_ok:
        return False, "Fecha invalida"
    ok, error = _ejecutar(
        "INSERT INTO nom_calendario (fecha, tipo, descripcion, creado) VALUES (%s,%s,%s,%s) "
        "ON CONFLICT (fecha) DO UPDATE SET tipo=EXCLUDED.tipo, descripcion=EXCLUDED.descripcion",
        (fecha_ok, tipo, descripcion, datetime.now().strftime("%d/%m/%Y %H:%M")))
    if ok:
        registrar_auditoria(usuario, "Nomina", "Registro %s el %s" % (tipo, fecha_ok))
    return ok, error


def eliminar_dia_calendario(fecha, usuario="sistema"):
    return _ejecutar("DELETE FROM nom_calendario WHERE fecha = %s", (fecha_iso(fecha) or str(fecha),))


def sembrar_feriados_peru(anio, usuario="sistema"):
    """Carga los feriados nacionales del Peru para el anio indicado."""
    anio = int(anio)
    feriados = [
        ("01-01", "Anio Nuevo"),
        ("05-01", "Dia del Trabajo"),
        ("06-29", "San Pedro y San Pablo"),
        ("07-23", "Dia de la Fuerza Aerea del Peru"),
        ("07-28", "Fiestas Patrias"),
        ("07-29", "Fiestas Patrias"),
        ("08-06", "Batalla de Junin"),
        ("08-30", "Santa Rosa de Lima"),
        ("10-08", "Combate de Angamos"),
        ("11-01", "Todos los Santos"),
        ("12-08", "Inmaculada Concepcion"),
        ("12-09", "Batalla de Ayacucho"),
        ("12-25", "Navidad del Senor"),
    ]
    # Jueves y Viernes Santo (fechas moviles, calculadas con el algoritmo de Gauss)
    a = anio % 19
    b = anio // 100
    c = anio % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    dia = ((h + l - 7 * m + 114) % 31) + 1
    pascua = date(anio, mes, dia)
    jueves_santo = pascua - timedelta(days=3)
    viernes_santo = pascua - timedelta(days=2)
    total = 0
    for mmdd, nombre in feriados:
        mes_f, dia_f = [int(x) for x in mmdd.split("-")]
        ok, _ = guardar_dia_calendario(date(anio, mes_f, dia_f), "FERIADO", nombre, usuario)
        total += 1 if ok else 0
    for fecha, nombre in ((jueves_santo, "Jueves Santo"), (viernes_santo, "Viernes Santo")):
        ok, _ = guardar_dia_calendario(fecha, "FERIADO", nombre, usuario)
        total += 1 if ok else 0
    registrar_auditoria(usuario, "Nomina", "Cargo feriados %s (%d)" % (anio, total))
    return total


# =========================================================================
# MAPEO RELOJ <-> EMPLEADO
# =========================================================================
def listar_mapeos():
    """{codigo_reloj: {'dni':..., 'nombre_reloj':...}}"""
    mapeos = {}
    for fila in _consultar("SELECT codigo_reloj, dni, nombre_reloj FROM nom_mapeo_reloj") or []:
        mapeos[str(fila[0])] = {"dni": str(fila[1]), "nombre_reloj": fila[2] or ""}
    return mapeos


def guardar_mapeo(codigo_reloj, dni, nombre_reloj="", usuario="sistema"):
    codigo_reloj = str(codigo_reloj or "").strip()
    dni = str(dni or "").strip()
    if not codigo_reloj or not dni:
        return False, "Codigo del reloj y DNI son obligatorios"
    ok, error = _ejecutar(
        "INSERT INTO nom_mapeo_reloj (codigo_reloj, dni, nombre_reloj, creado) VALUES (%s,%s,%s,%s) "
        "ON CONFLICT (codigo_reloj) DO UPDATE SET dni=EXCLUDED.dni, nombre_reloj=EXCLUDED.nombre_reloj",
        (codigo_reloj, dni, nombre_reloj, datetime.now().strftime("%d/%m/%Y %H:%M")))
    if ok:
        # Reasigna las marcaciones ya importadas con este codigo
        _ejecutar("UPDATE nom_marcaciones SET dni = %s WHERE codigo_reloj = %s", (dni, codigo_reloj))
        registrar_auditoria(usuario, "Nomina", "Vinculo codigo de reloj %s con DNI %s" % (codigo_reloj, dni))
    return ok, error


def eliminar_mapeo(codigo_reloj, usuario="sistema"):
    return _ejecutar("DELETE FROM nom_mapeo_reloj WHERE codigo_reloj = %s", (str(codigo_reloj),))


def similitud_nombres(a, b):
    """Similitud simple por tokens (0 a 1) entre dos nombres de persona."""
    ta = set(normalizar_texto(a).split())
    tb = set(normalizar_texto(b).split())
    if not ta or not tb:
        return 0.0
    comunes = len(ta & tb)
    return comunes / float(max(len(ta), len(tb)))


def sugerir_dni(codigo_reloj, nombre_reloj, empleados=None, mapeos=None):
    """Sugiere el DNI de un codigo del reloj. Devuelve (dni, metodo, confianza)."""
    codigo_reloj = str(codigo_reloj or "").strip()
    mapeos = mapeos if mapeos is not None else listar_mapeos()
    if codigo_reloj in mapeos:
        return mapeos[codigo_reloj]["dni"], "MAPEO GUARDADO", 1.0
    empleados = empleados if empleados is not None else listar_empleados()
    por_dni = {str(e["dni"]).strip(): e for e in empleados}
    por_doc = {}
    for empleado in empleados:
        por_doc.setdefault(normalizar_documento(empleado["dni"]), empleado)
    if codigo_reloj in por_dni:
        return codigo_reloj, "DNI EXACTO", 1.0
    normalizado = normalizar_documento(codigo_reloj)
    if normalizado and normalizado in por_doc:
        return str(por_doc[normalizado]["dni"]), "DNI NORMALIZADO", 0.95
    # Comparacion por nombre
    mejor, mejor_puntaje = None, 0.0
    for empleado in empleados:
        puntaje = similitud_nombres(nombre_reloj, empleado["nombre"])
        if puntaje > mejor_puntaje:
            mejor, mejor_puntaje = empleado, puntaje
    if mejor and mejor_puntaje >= 0.6:
        return str(mejor["dni"]), "NOMBRE SIMILAR (%.0f%%)" % (mejor_puntaje * 100), mejor_puntaje
    return None, "SIN COINCIDENCIA", 0.0


def personas_del_reloj():
    """Personas detectadas en las marcaciones importadas, con su estado de mapeo."""
    filas = _consultar(
        "SELECT codigo_reloj, MAX(nombre_reloj), COUNT(*), MIN(fecha), MAX(fecha), "
        "       MAX(dni) FROM nom_marcaciones GROUP BY codigo_reloj ORDER BY MAX(nombre_reloj)") or []
    resultado = []
    for codigo, nombre, marcas, fmin, fmax, dni in filas:
        resultado.append({"codigo_reloj": codigo, "nombre_reloj": nombre or "", "marcas": marcas,
                          "fecha_min": fmin, "fecha_max": fmax, "dni": dni or ""})
    return resultado


# =========================================================================
# LECTURA DEL EXCEL DEL RELOJ BIOMETRICO
# =========================================================================
def _normalizar_encabezado(valor):
    return normalizar_texto(valor).lower()


# Valores de relleno que algunos relojes escriben cuando el dato no existe
_RELLENOS = {"", "-", "--", "---", "N/A", "NA", "NINGUNO", "NINGUNA", "SIN NOMBRE", "0", "NULL", "NONE"}


def es_nombre_util(valor):
    """True si el valor sirve como nombre de persona (no es un relleno del reloj)."""
    if valor is None:
        return False
    texto = str(valor).strip()
    if not texto or len(texto) < 2:
        return False
    return normalizar_texto(texto) not in _RELLENOS


def mapear_columnas(encabezados):
    """Detecta que columna del archivo corresponde a cada dato. Devuelve {campo: indice}."""
    mapeo = {}
    normalizados = [_normalizar_encabezado(h) for h in encabezados]
    for campo, alias in ALIAS_COLUMNAS.items():
        for indice, titulo in enumerate(normalizados):
            if indice in mapeo.values():
                continue
            if titulo in alias:
                mapeo[campo] = indice
                break
        if campo in mapeo:
            continue
        # Coincidencia parcial (el titulo contiene el alias). Se exige un minimo de
        # caracteres para que alias cortos ("no", "id") no confundan otras columnas.
        for indice, titulo in enumerate(normalizados):
            if not titulo or indice in mapeo.values():
                continue
            if any(len(alias_txt) >= 5 and alias_txt in titulo for alias_txt in alias):
                mapeo[campo] = indice
                break
    return mapeo


def detectar_formato(encabezados, mapeo):
    """Identifica el origen probable del archivo segun sus columnas."""
    titulos = " ".join(_normalizar_encabezado(h) for h in encabezados if h)
    if "tipo de pase" in titulos or "temperatura en la superficie" in titulos or "metodo de verificacion" in titulos:
        return "HIKVISION"
    if "verify mode" in titulos or "att date" in titulos or "punch" in titulos:
        return "ZKTECO"
    if "ac no" in titulos or "enroll" in titulos:
        return "ZKTECO"
    return "GENERICO"


def _buscar_fila_encabezado(matriz, max_filas=40):
    """Devuelve (indice_fila, mapeo) del encabezado mas probable dentro del archivo."""
    mejor_indice, mejor_mapeo, mejor_puntaje = None, {}, 0
    for indice, fila in enumerate(matriz[:max_filas]):
        celdas = [c for c in fila if c not in (None, "")]
        if len(celdas) < 2:
            continue
        mapeo = mapear_columnas(fila)
        puntaje = len(mapeo)
        # Un encabezado valido debe identificar al menos fecha/hora y un identificador
        tiene_tiempo = ("fecha" in mapeo) and ("hora" in mapeo or "fecha" in mapeo)
        tiene_persona = any(c in mapeo for c in ("codigo", "nombre_completo", "nombre", "apellido"))
        if tiene_tiempo and tiene_persona and puntaje > mejor_puntaje:
            mejor_indice, mejor_mapeo, mejor_puntaje = indice, mapeo, puntaje
    return mejor_indice, mejor_mapeo


def leer_archivo_marcaciones(ruta, fila_encabezado=None, mapeo=None):
    """Lee el archivo del reloj y devuelve su contenido normalizado.

    Devuelve un diccionario:
        {'ok': bool, 'error': str, 'encabezados': [...], 'mapeo': {...},
         'formato': 'HIKVISION', 'fila_encabezado': int, 'filas': [ {...} ],
         'total': int, 'hojas': [...]}
    Cada fila normalizada contiene: codigo, nombre_reloj, fecha, hora, tipo_pase,
    metodo, departamento y 'valida' (bool).
    """
    resultado = {"ok": False, "error": "", "encabezados": [], "mapeo": {},
                 "formato": "GENERICO", "fila_encabezado": 0, "filas": [], "total": 0,
                 "hojas": [], "ruta": ruta}
    if not ruta or not os.path.exists(ruta):
        resultado["error"] = "El archivo no existe"
        return resultado
    extension = os.path.splitext(ruta)[1].lower()
    matriz = []
    if extension in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        if not OPENPYXL_OK:
            resultado["error"] = "Falta la libreria openpyxl. Ejecute: pip install openpyxl"
            return resultado
        try:
            libro = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
        except Exception as e:
            resultado["error"] = "No se pudo abrir el archivo de Excel: %s" % e
            return resultado
        try:
            resultado["hojas"] = list(libro.sheetnames)
            hoja = libro[libro.sheetnames[0]]
            for fila in hoja.iter_rows(values_only=True):
                matriz.append(list(fila))
        finally:
            try:
                libro.close()
            except Exception:
                pass
    elif extension in (".csv", ".txt"):
        try:
            with open(ruta, "r", encoding="utf-8-sig", errors="replace") as manejador:
                muestra = manejador.read(4096)
            dialecto = csv.Sniffer().sniff(muestra, delimiters=",;\t|") if muestra else csv.excel
        except Exception:
            dialecto = csv.excel
        try:
            with open(ruta, "r", encoding="utf-8-sig", errors="replace") as manejador:
                for fila in csv.reader(manejador, dialecto):
                    matriz.append(list(fila))
        except Exception as e:
            resultado["error"] = "No se pudo leer el archivo CSV: %s" % e
            return resultado
    else:
        # Formatos antiguos (.xls) o exportaciones HTML con extension .xls
        try:
            import pandas as pd
            tabla = pd.read_excel(ruta, header=None)
            matriz = tabla.values.tolist()
        except Exception:
            try:
                import pandas as pd
                tablas = pd.read_html(ruta)
                matriz = tablas[0].values.tolist() if tablas else []
            except Exception as e:
                resultado["error"] = ("Formato no soportado (%s). Exporte el reloj a .xlsx o .csv. Detalle: %s"
                                      % (extension, e))
                return resultado
    if not matriz:
        resultado["error"] = "El archivo no contiene datos"
        return resultado

    if fila_encabezado is None:
        indice, mapeo_detectado = _buscar_fila_encabezado(matriz)
        if indice is None:
            resultado["error"] = ("No se pudo detectar el encabezado del archivo. "
                                  "Verifique que el reporte tenga columnas de fecha y hora.")
            return resultado
    else:
        indice = int(fila_encabezado)
        mapeo_detectado = mapear_columnas(matriz[indice] if indice < len(matriz) else [])
    if mapeo:
        mapeo_detectado = dict(mapeo_detectado)
        mapeo_detectado.update({k: v for k, v in mapeo.items() if v is not None})
    encabezados = [("" if c is None else str(c)) for c in matriz[indice]]
    resultado["encabezados"] = encabezados
    resultado["mapeo"] = mapeo_detectado
    resultado["fila_encabezado"] = indice
    resultado["formato"] = detectar_formato(encabezados, mapeo_detectado)

    def celda(fila, campo):
        posicion = mapeo_detectado.get(campo)
        if posicion is None or posicion >= len(fila):
            return None
        return fila[posicion]

    filas_normalizadas = []
    for fila in matriz[indice + 1:]:
        if not any(c not in (None, "") for c in fila):
            continue
        fecha_valor = celda(fila, "fecha")
        hora_valor = celda(fila, "hora")
        if hora_valor is None and fecha_valor is not None and " " in str(fecha_valor):
            partes = str(fecha_valor).split(" ")
            fecha_valor = partes[0]
            hora_valor = partes[1] if len(partes) > 1 else None
        fecha = fecha_iso(fecha_valor)
        hora = parse_hora(hora_valor)
        codigo = celda(fila, "codigo")
        partes = [celda(fila, "nombre"), celda(fila, "apellido")]
        nombre = " ".join(str(p).strip() for p in partes if p not in (None, "") and es_nombre_util(p))
        if not nombre:
            # La columna de nombre completo (si existe) sirve de respaldo
            alterno = celda(fila, "nombre_completo")
            nombre = str(alterno).strip() if es_nombre_util(alterno) else ""
        filas_normalizadas.append({
            "codigo": "" if codigo is None else str(codigo).strip(),
            "nombre_reloj": str(nombre or "").strip(),
            "fecha": fecha,
            "hora": hora,
            "tipo_pase": str(celda(fila, "tipo_pase") or "").strip(),
            "metodo": str(celda(fila, "metodo") or "").strip(),
            "departamento": str(celda(fila, "departamento") or "").strip(),
            "valida": bool(fecha and hora and (codigo not in (None, "") or nombre)),
            "crudo": [("" if c is None else str(c)) for c in fila],
        })
    resultado["filas"] = filas_normalizadas
    resultado["total"] = len(filas_normalizadas)
    resultado["ok"] = True
    return resultado


# =========================================================================
# IMPORTACION DE MARCACIONES A LA BASE DE DATOS
# =========================================================================
def importar_marcaciones(ruta, mapeo=None, fila_encabezado=None, usuario="sistema",
                         crear_mapeos_automaticos=True, progreso=None):
    """Importa el archivo del reloj a nom_marcaciones (sin duplicados)."""
    lectura = leer_archivo_marcaciones(ruta, fila_encabezado=fila_encabezado, mapeo=mapeo)
    if not lectura["ok"]:
        return {"ok": False, "error": lectura["error"]}, None

    empleados = listar_empleados()
    mapeos = listar_mapeos()
    archivo = os.path.basename(ruta)
    ahora = datetime.now().strftime("%d/%m/%Y %H:%M")
    try:
        with open(ruta, "rb") as manejador:
            hash_archivo = hashlib.md5(manejador.read()).hexdigest()
    except Exception:
        hash_archivo = ""

    pendientes = []
    personas = {}
    fechas = []
    invalidas = 0
    duplicadas = 0
    nuevas = 0
    sin_empleado_codigos = {}
    # Huellas de las marcas que ya estan en la base: permite informar con exactitud
    # cuantas marcas son nuevas y cuantas ya se habian importado antes.
    existentes = {fila[0] for fila in (_consultar("SELECT hash_marca FROM nom_marcaciones") or [])}
    for posicion, fila in enumerate(lectura["filas"]):
        if progreso and posicion % 200 == 0:
            try:
                progreso(posicion, lectura["total"])
            except Exception:
                pass
        if not fila["valida"]:
            invalidas += 1
            continue
        codigo = fila["codigo"] or fila["nombre_reloj"]
        dni, metodo, _ = sugerir_dni(codigo, fila["nombre_reloj"], empleados, mapeos)
        if dni and crear_mapeos_automaticos and metodo in ("DNI EXACTO", "DNI NORMALIZADO", "NOMBRE SIMILAR (100%)"):
            mapeos[codigo] = {"dni": dni, "nombre_reloj": fila["nombre_reloj"]}
            guardar_mapeo(codigo, dni, fila["nombre_reloj"], usuario)
        if not dni:
            sin_empleado_codigos[codigo] = fila["nombre_reloj"]
        firma = "%s|%s|%s|%s" % (codigo, fila["fecha"], fila["hora"], fila["tipo_pase"])
        huella = hashlib.md5(firma.encode("utf-8")).hexdigest()
        if huella in existentes:
            duplicadas += 1
        else:
            nuevas += 1
            existentes.add(huella)
        pendientes.append((
            codigo, dni, fila["nombre_reloj"], fila["fecha"], fila["hora"],
            fila["tipo_pase"], fila["metodo"], fila["departamento"], archivo, huella, ahora,
        ))
        personas[codigo] = fila["nombre_reloj"]
        fechas.append(fila["fecha"])

    ok, insertadas, error = _ejecutar_lote(
        "INSERT INTO nom_marcaciones (codigo_reloj, dni, nombre_reloj, fecha, hora, tipo_pase, "
        "metodo, departamento, fuente, hash_marca, importado) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) "
        # Si la marca ya existia solo se refrescan el nombre y la vinculacion: asi una
        # reimportacion del mismo archivo corrige nombres mal leidos o nuevos mapeos.
        "ON CONFLICT (hash_marca) DO UPDATE SET nombre_reloj = EXCLUDED.nombre_reloj, "
        "dni = COALESCE(EXCLUDED.dni, nom_marcaciones.dni), codigo_reloj = EXCLUDED.codigo_reloj",
  pendientes)
    if not ok:
        return {"ok": False, "error": "Error guardando las marcaciones: %s" % error}, None

    resultado = {
        "ok": True,
        "error": "",
        "archivo": archivo,
        "formato": PERFILES_FORMATO.get(lectura["formato"], lectura["formato"]),
        "total_filas": lectura["total"],
        "nuevas": nuevas,
        "duplicadas": duplicadas,
        "invalidas": invalidas,
        "sin_empleado_codigos": sin_empleado_codigos,
        "personas": sorted(personas.values()),
        "fecha_min": min(fechas) if fechas else None,
        "fecha_max": max(fechas) if fechas else None,
        "encabezados": lectura["encabezados"],
        "mapeo": lectura["mapeo"],
    }
    detalle = json.dumps({"sin_empleado": list(sin_empleado_codigos.keys())[:50],
                          "personas": resultado["personas"][:100]}, ensure_ascii=False)
    _ejecutar(
        "INSERT INTO nom_importaciones (archivo, hash_archivo, formato, total_filas, nuevas, "
        "duplicadas, sin_empleado, fecha_min, fecha_max, usuario, fecha, detalle) "
        "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
        (archivo, hash_archivo, resultado["formato"], lectura["total"], nuevas, duplicadas,
         len(sin_empleado_codigos), resultado["fecha_min"], resultado["fecha_max"], usuario, ahora, detalle))
    registrar_auditoria(usuario, "Nomina",
                        "Importo marcaciones de %s (%d nuevas, %d duplicadas)" % (archivo, nuevas, duplicadas))
    return resultado, lectura


def listar_importaciones(limite=50):
    columnas = ["id", "archivo", "formato", "total_filas", "nuevas", "duplicadas", "sin_empleado",
                "fecha_min", "fecha_max", "usuario", "fecha"]
    return [dict(zip(columnas, fila)) for fila in (_consultar(
        "SELECT id, archivo, formato, total_filas, nuevas, duplicadas, sin_empleado, fecha_min, "
        "fecha_max, usuario, fecha FROM nom_importaciones ORDER BY id DESC LIMIT %s", (int(limite),)) or [])]


def eliminar_marcaciones_periodo(desde, hasta, usuario="sistema"):
    """Borra las marcaciones de un rango (para reimportar limpio)."""
    desde_iso, hasta_iso = fecha_iso(desde), fecha_iso(hasta)
    if not desde_iso or not hasta_iso:
        return False, "Rango de fechas invalido"
    conn = conectar_db(silencioso=True)
    if not conn:
        return False, "Sin conexion a la base de datos"
    try:
        with conn.cursor() as cursor:
            cursor.execute("DELETE FROM nom_marcaciones WHERE fecha BETWEEN %s AND %s",
                           (desde_iso, hasta_iso))
            borradas = cursor.rowcount
        conn.commit()
        registrar_auditoria(usuario, "Nomina",
                            "Elimino %d marcaciones entre %s y %s" % (borradas, desde_iso, hasta_iso))
        return True, "Se eliminaron %d marcaciones" % borradas
    except Exception as e:
        try:
            conn.rollback()
        except Exception:
            pass
        return False, str(e)
    finally:
        liberar_conexion(conn)


# =========================================================================
# INCIDENCIAS Y HORAS EXTRA
# =========================================================================
def listar_incidencias(dni=None, desde=None, hasta=None):
    columnas = ["id", "dni", "fecha_desde", "fecha_hasta", "tipo", "motivo", "con_goce",
                "horas", "aprobado_por", "archivo", "creado"]
    sql = ("SELECT id, dni, fecha_desde, fecha_hasta, tipo, motivo, con_goce, horas, "
           "aprobado_por, archivo, creado FROM nom_incidencias WHERE 1=1 ")
    params = []
    if dni:
        sql += "AND dni = %s "
        params.append(str(dni))
    if desde:
        sql += "AND fecha_hasta >= %s "
        params.append(fecha_iso(desde))
    if hasta:
        sql += "AND fecha_desde <= %s "
        params.append(fecha_iso(hasta))
    sql += "ORDER BY fecha_desde DESC, dni"
    return [dict(zip(columnas, fila)) for fila in (_consultar(sql, tuple(params)) or [])]


def guardar_incidencia(datos, usuario="sistema"):
    dni = str(datos.get("dni") or "").strip()
    desde = fecha_iso(datos.get("fecha_desde"))
    hasta = fecha_iso(datos.get("fecha_hasta")) or desde
    tipo = datos.get("tipo") or EST_PERMISO
    if not dni:
        return None, "Debe indicar el empleado"
    if not desde:
        return None, "La fecha de inicio no es valida"
    if hasta < desde:
        desde, hasta = hasta, desde
    campos = (dni, desde, hasta, tipo, datos.get("motivo", ""), bool(datos.get("con_goce", True)),
              a_float(datos["horas"]) if datos.get("horas") not in (None, "") else None,
              datos.get("aprobado_por", ""), datos.get("archivo", ""),
              datetime.now().strftime("%d/%m/%Y %H:%M"))
    if datos.get("id"):
        ok, error = _ejecutar(
            "UPDATE nom_incidencias SET dni=%s, fecha_desde=%s, fecha_hasta=%s, tipo=%s, motivo=%s, "
            "con_goce=%s, horas=%s, aprobado_por=%s, archivo=%s WHERE id=%s",
            campos + (int(datos["id"]),))
        if not ok:
            return None, error
        registrar_auditoria(usuario, "Nomina", "Edito incidencia de %s" % dni)
        return int(datos["id"]), ""
    nuevo_id, error = _ejecutar_retorna_id(
        "INSERT INTO nom_incidencias (dni, fecha_desde, fecha_hasta, tipo, motivo, con_goce, horas, "
        "aprobado_por, archivo, creado) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id", campos)
    if nuevo_id:
        registrar_auditoria(usuario, "Nomina", "Registro %s para %s (%s a %s)" % (tipo, dni, desde, hasta))
    return nuevo_id, error


def eliminar_incidencia(incidencia_id, usuario="sistema"):
    ok, error = _ejecutar("DELETE FROM nom_incidencias WHERE id = %s", (int(incidencia_id),))
    if ok:
        registrar_auditoria(usuario, "Nomina", "Elimino incidencia id %s" % incidencia_id)
    return ok, error


def listar_horas_extra(dni=None, desde=None, hasta=None):
    columnas = ["id", "dni", "fecha", "minutos", "tipo", "origen", "aprobado", "observacion", "creado"]
    sql = ("SELECT id, dni, fecha, minutos, tipo, origen, aprobado, observacion, creado "
           "FROM nom_horas_extra WHERE 1=1 ")
    params = []
    if dni:
        sql += "AND dni = %s "
        params.append(str(dni))
    if desde:
        sql += "AND fecha >= %s "
        params.append(fecha_iso(desde))
    if hasta:
        sql += "AND fecha <= %s "
        params.append(fecha_iso(hasta))
    sql += "ORDER BY fecha DESC, dni"
    return [dict(zip(columnas, fila)) for fila in (_consultar(sql, tuple(params)) or [])]


def guardar_hora_extra(datos, usuario="sistema"):
    dni = str(datos.get("dni") or "").strip()
    fecha = fecha_iso(datos.get("fecha"))
    minutos = a_int(datos.get("minutos"))
    if not dni or not fecha:
        return False, "Empleado y fecha son obligatorios"
    if minutos <= 0:
        return False, "Los minutos deben ser mayores a cero"
    ok, error = _ejecutar(
        "INSERT INTO nom_horas_extra (dni, fecha, minutos, tipo, origen, aprobado, observacion, creado) "
        "VALUES (%s,%s,%s,%s,%s,%s,%s,%s) "
        "ON CONFLICT (dni, fecha) DO UPDATE SET minutos=EXCLUDED.minutos, tipo=EXCLUDED.tipo, "
        "origen=EXCLUDED.origen, aprobado=EXCLUDED.aprobado, observacion=EXCLUDED.observacion",
        (dni, fecha, minutos, datos.get("tipo", "25"), datos.get("origen", "MANUAL"),
         bool(datos.get("aprobado", True)), datos.get("observacion", ""),
         datetime.now().strftime("%d/%m/%Y %H:%M")))
    if ok:
        registrar_auditoria(usuario, "Nomina", "Registro horas extra de %s (%s min)" % (dni, minutos))
    return ok, error


def eliminar_hora_extra(hora_id, usuario="sistema"):
    return _ejecutar("DELETE FROM nom_horas_extra WHERE id = %s", (int(hora_id),))


# =========================================================================
# MOTOR DE ASISTENCIA
# =========================================================================
def _cargar_contexto(desde, hasta):
    """Carga en memoria todo lo necesario para calcular la asistencia del rango."""
    contexto = {
        "empleados": {str(e["dni"]): e for e in listar_empleados()},
        "turnos": {int(t["id"]): t for t in listar_turnos()},
        "horarios": {},
        "asignaciones": {},
        "calendario": {},
        "incidencias": {},
        "marcaciones": {},
        "mapeos": listar_mapeos(),
    }
    for horario in listar_horarios():
        completo = obtener_horario(horario["id"])
        if completo:
            contexto["horarios"][int(horario["id"])] = completo
    for asignacion in listar_asignaciones():
        if not asignacion["activo"]:
            continue
        contexto["asignaciones"].setdefault(str(asignacion["dni"]), []).append(asignacion)
    for dia in listar_calendario():
        contexto["calendario"][dia["fecha"]] = dia
    for incidencia in listar_incidencias(desde=desde, hasta=hasta):
        for fecha in rango_fechas(max(incidencia["fecha_desde"], desde),
                                  min(incidencia["fecha_hasta"], hasta)):
            contexto["incidencias"].setdefault(str(incidencia["dni"]), {})[fecha.strftime("%Y-%m-%d")] = incidencia
    # Marcaciones del rango (con un dia extra para turnos que cruzan medianoche)
    dias = rango_fechas(desde, hasta)
    if dias:
        inicio = (dias[0] - timedelta(days=1)).strftime("%Y-%m-%d")
        fin = (dias[-1] + timedelta(days=1)).strftime("%Y-%m-%d")
        filas = _consultar(
            "SELECT dni, codigo_reloj, nombre_reloj, fecha, hora FROM nom_marcaciones "
            "WHERE fecha BETWEEN %s AND %s ORDER BY dni, fecha, hora", (inicio, fin)) or []
        for dni, codigo, nombre, fecha, hora in filas:
            clave_dni = str(dni or "") or ("REL:" + str(codigo))
            contexto["marcaciones"].setdefault(clave_dni, {}).setdefault(fecha, []).append(hora)
    return contexto


def resolver_turno_del_dia(dni, fecha, contexto):
    """Devuelve el turno que corresponde a un empleado en una fecha (o None = descanso)."""
    fecha_texto = fecha.strftime("%Y-%m-%d") if hasattr(fecha, "strftime") else str(fecha)
    dia = dia_semana(fecha_texto)
    dni = str(dni)
    # 1) Asignacion especifica que cubre la fecha (la mas reciente gana)
    candidatas = [a for a in contexto["asignaciones"].get(dni, [])
                  if a["fecha_desde"] <= fecha_texto and (not a["fecha_hasta"] or a["fecha_hasta"] >= fecha_texto)]
    candidatas.sort(key=lambda a: a["fecha_desde"], reverse=True)
    for asignacion in candidatas:
        if asignacion.get("turno_id"):
            return contexto["turnos"].get(int(asignacion["turno_id"])), "ASIGNACION"
        if asignacion.get("horario_id"):
            horario = contexto["horarios"].get(int(asignacion["horario_id"]))
            if horario:
                turno_id = horario["detalle"].get(dia)
                return (contexto["turnos"].get(int(turno_id)) if turno_id else None), "HORARIO"
    # 2) Horario fijo del empleado
    empleado = contexto["empleados"].get(dni)
    if empleado:
        if empleado.get("horario_id"):
            horario = contexto["horarios"].get(int(empleado["horario_id"]))
            if horario:
                turno_id = horario["detalle"].get(dia)
                return (contexto["turnos"].get(int(turno_id)) if turno_id else None), "HORARIO EMPLEADO"
        if empleado.get("turno_id"):
            turno = contexto["turnos"].get(int(empleado["turno_id"]))
            if turno:
                return turno, "TURNO EMPLEADO"
    return None, "SIN ASIGNACION"


def _turno_cruza_medianoche(turno):
    """True si el turno termina al dia siguiente (por ejemplo 22:00 a 06:00)."""
    if not turno:
        return False
    if turno.get("cruza_medianoche"):
        return True
    entrada = minutos_de_hora(turno.get("hora_entrada")) or 0
    salida = minutos_de_hora(turno.get("hora_salida")) or 0
    return salida <= entrada


def _marcas_del_dia(dni, fecha, fecha_texto, turno, contexto, parametros):
    """Marcas utiles del dia, normalizadas a minutos desde la medianoche.

    - Los turnos que cruzan medianoche suman las marcas del dia siguiente hasta la
      hora de salida (por eso se les agregan 1440 minutos: asi quedan al final).
    - En los turnos nocturnos se descartan las marcas del propio dia anteriores a
      la ventana del turno (pertenecen al turno del dia anterior).
    - Se descartan las lecturas repetidas: los relojes biometricos suelen grabar
      dos o tres veces la misma marcacion (reintentos de huella o de rostro). Se
      comparan las marcas crudas entre si, de modo que una rafaga de lecturas
      seguidas (por ejemplo 05:46, 05:47 y 05:48) cuenta como UNA sola marcacion.
      Sin esta depuracion, un dia con una unica entrada pareceria tener entrada y
      salida, y se pagaria un dia completo que nadie trabajo.
    """
    marcas_dni = contexto["marcaciones"].get(str(dni), {})
    crudas = []
    for hora in marcas_dni.get(fecha_texto) or []:
        minutos = minutos_de_hora(hora)
        if minutos is not None:
            crudas.append(minutos)
    if _turno_cruza_medianoche(turno):
        entrada_teorica = minutos_de_hora(turno.get("hora_entrada")) or 0
        # La ventana del turno no puede empezar mas de 4 horas antes de la entrada
        crudas = [minutos for minutos in crudas if minutos >= max(0, entrada_teorica - 240)]
        siguiente = (fecha + timedelta(days=1)).strftime("%Y-%m-%d")
        salida_teorica = (minutos_de_hora(turno.get("hora_salida")) or 0) + 24 * 60
        for hora in marcas_dni.get(siguiente) or []:
            minutos = minutos_de_hora(hora)
            if minutos is not None and (minutos + 24 * 60) <= salida_teorica + 180:
                crudas.append(minutos + 24 * 60)
    crudas.sort()
    minimo = max(0, a_int(parametros.get("minutos_minimos_entre_marcas"), 2))
    depuradas = []
    anterior = None
    for minutos in crudas:
        if anterior is not None and (minutos - anterior) < minimo:
            # Lectura repetida de la misma rafaga: se descarta pero sigue contando
            # como referencia para las que vienen inmediatamente despues.
            anterior = minutos
            continue
        depuradas.append(minutos)
        anterior = minutos
    return depuradas


def calcular_dia(dni, fecha, contexto, parametros=None):
    """Calcula la asistencia de un empleado en un dia. Devuelve un diccionario."""
    parametros = parametros or obtener_parametros()
    fecha_texto = fecha.strftime("%Y-%m-%d")
    turno, origen_turno = resolver_turno_del_dia(dni, fecha, contexto)
    dia_calendario = contexto["calendario"].get(fecha_texto)
    incidencia = contexto["incidencias"].get(str(dni), {}).get(fecha_texto)

    # Lista de minutos (ya ordenada y depurada) de las marcas utiles del dia
    marcas = _marcas_del_dia(dni, fecha, fecha_texto, turno, contexto, parametros)

    resultado = {
        "dni": str(dni), "fecha": fecha_texto, "turno_id": turno["id"] if turno else None,
        "hora_entrada": None, "hora_salida": None, "primera_marca": None, "ultima_marca": None,
        "n_marcas": len(marcas), "minutos_tardanza": 0, "minutos_anticipo": 0,
        "minutos_trabajados": 0, "minutos_extra": 0, "minutos_falta": 0,
        "estado": EST_PENDIENTE, "manual": False,
        "incidencia_id": incidencia["id"] if incidencia else None,
        "observacion": origen_turno if turno else (dia_calendario["descripcion"] if dia_calendario else ""),
    }

    es_descanso = turno is None
    if turno:
        dias_turno = parsear_dias_semana(turno.get("dias_semana"))
        if dias_turno and dia_semana(fecha_texto) not in dias_turno:
            es_descanso = True
    tipo_calendario = (dia_calendario or {}).get("tipo", "")
    es_feriado = tipo_calendario == "FERIADO"
    if tipo_calendario == "LABORABLE":
        es_feriado = False
        es_descanso = False if turno else es_descanso

    if marcas:
        resultado["primera_marca"] = hora_de_minutos(marcas[0])
        resultado["ultima_marca"] = hora_de_minutos(marcas[-1])
        resultado["hora_entrada"] = hora_de_minutos(marcas[0])
        resultado["hora_salida"] = hora_de_minutos(marcas[-1]) if len(marcas) > 1 else None

    # --- Incidencia registrada: el dia queda justificado (no genera descuento) ---
    if incidencia:
        resultado["estado"] = incidencia["tipo"]
        if len(marcas) >= 2:
            resultado["minutos_trabajados"] = marcas[-1] - marcas[0]
        return resultado

    tolerancia = a_int(turno["tolerancia_min"], a_int(parametros.get("tolerancia_min"), 10)) if turno else 0
    refrigerio = a_int(turno["refrigerio_min"], a_int(parametros.get("refrigerio_min"), 45)) if turno else 0

    # --- Sin turno asignado (descanso) ---
    if es_descanso:
        if len(marcas) >= 2:
            trabajados = max(0, (marcas[-1] - marcas[0]) - refrigerio)
            resultado["minutos_trabajados"] = trabajados
            resultado["minutos_extra"] = trabajados
            resultado["estado"] = EST_FERIADO_TRAB if es_feriado else EST_DESCANSO_TRAB
        else:
            resultado["estado"] = EST_FERIADO if es_feriado else EST_DESCANSO
        return resultado

    # --- Turno programado ---
    if not marcas:
        resultado["estado"] = EST_FALTA
        jornada_min = int(round(a_float(turno.get("horas_jornada"), 8) * 60)) or 480
        resultado["minutos_falta"] = jornada_min
        return resultado

    entrada = (minutos_de_hora(turno["hora_entrada"]) or 0)
    salida_teorica = (minutos_de_hora(turno["hora_salida"]) or 0)
    if salida_teorica <= entrada:
        salida_teorica += 24 * 60
    marca_entrada, marca_salida = marcas[0], marcas[-1]

    tardanza = max(0, marca_entrada - entrada - tolerancia)
    resultado["minutos_tardanza"] = tardanza

    if len(marcas) == 1:
        resultado["estado"] = EST_INCOMPLETO
        resultado["minutos_falta"] = tardanza
        resultado["observacion"] = "Falta una marcacion (solo se registro la entrada o la salida)"
        return resultado

    anticipo = max(0, salida_teorica - marca_salida)
    resultado["minutos_anticipo"] = anticipo
    trabajados = max(0, marca_salida - marca_entrada - refrigerio)
    resultado["minutos_trabajados"] = trabajados

    # Horas extra: tiempo posterior a la hora de salida, redondeado a bloques
    extra_bruto = max(0, marca_salida - salida_teorica)
    bloque = max(1, a_int(parametros.get("he_bloque_min"), 15))
    minimo = a_int(parametros.get("he_minimo_min"), 15)
    tope = a_int(parametros.get("he_tope_dia_min"), 300)
    if extra_bruto >= minimo:
        extra = (extra_bruto // bloque) * bloque
        resultado["minutos_extra"] = min(extra, tope)
    descontar_anticipo = normalizar_texto(parametros.get("descontar_anticipo", "SI")) == "SI"
    resultado["minutos_falta"] = (tardanza + anticipo) if descontar_anticipo else tardanza
    resultado["estado"] = EST_PUNTUAL if tardanza == 0 else EST_TARDANZA
    return resultado


def recalcular_asistencia(desde, hasta, dnis=None, usuario="sistema", progreso=None):
    """Recalcula y guarda la asistencia diaria del rango indicado."""
    ok_esquema, error_esquema = inicializar_esquema_nomina()
    if not ok_esquema:
        return {"ok": False, "error": error_esquema}
    dias = rango_fechas(desde, hasta)
    if not dias:
        return {"ok": False, "error": "Rango de fechas invalido"}
    parametros = obtener_parametros()
    contexto = _cargar_contexto(dias[0], dias[-1])
    universo = set(str(d) for d in (dnis or []))
    if not universo:
        universo = {str(e["dni"]) for e in contexto["empleados"].values()
                    if normalizar_texto(e.get("estado")) not in ("CESADO", "INACTIVO")}
        universo |= set(contexto["marcaciones"].keys())
    universo = {d for d in universo if not d.startswith("REL:")}
    manuales = set()
    for fila in _consultar("SELECT dni, fecha FROM nom_asistencia WHERE manual = TRUE "
                           "AND fecha BETWEEN %s AND %s",
                           (dias[0].strftime("%Y-%m-%d"), dias[-1].strftime("%Y-%m-%d"))) or []:
        manuales.add((str(fila[0]), fila[1]))

    ahora = datetime.now().strftime("%d/%m/%Y %H:%M")
    filas_guardar = []
    resumen = {"dias": 0, "puntuales": 0, "tardanzas": 0, "faltas": 0, "descansos": 0,
               "feriados": 0, "incompletos": 0, "justificados": 0, "trabajados_descanso": 0,
               "minutos_tardanza": 0, "minutos_extra": 0, "minutos_trabajados": 0,
               "minutos_descanso_trabajado": 0,
               "empleados": len(universo), "periodo": "%s a %s" % (dias[0], dias[-1])}
    procesados = 0
    for dni in sorted(universo):
        for fecha in dias:
            if (dni, fecha.strftime("%Y-%m-%d")) in manuales:
                continue
            calculo = calcular_dia(dni, fecha, contexto, parametros)
            filas_guardar.append((
                calculo["dni"], calculo["fecha"], calculo["turno_id"], calculo["hora_entrada"],
                calculo["hora_salida"], calculo["primera_marca"], calculo["ultima_marca"],
                calculo["n_marcas"], calculo["minutos_tardanza"], calculo["minutos_anticipo"],
                calculo["minutos_trabajados"], calculo["minutos_extra"], calculo["minutos_falta"],
                calculo["estado"], False, calculo["incidencia_id"], calculo["observacion"], ahora,
            ))
            estado = calculo["estado"]
            resumen["dias"] += 1
            resumen["minutos_tardanza"] += calculo["minutos_tardanza"]
            # El trabajo en dia de descanso se informa aparte: se paga como dia con
            # recargo, no como hora extra (si se sumara aqui se contaria dos veces).
            if estado in (EST_DESCANSO_TRAB, EST_FERIADO_TRAB):
                resumen["minutos_descanso_trabajado"] += calculo["minutos_extra"]
            else:
                resumen["minutos_extra"] += calculo["minutos_extra"]
            resumen["minutos_trabajados"] += calculo["minutos_trabajados"]
            if estado == EST_PUNTUAL:
                resumen["puntuales"] += 1
            elif estado == EST_TARDANZA:
                resumen["tardanzas"] += 1
            elif estado == EST_FALTA:
                resumen["faltas"] += 1
            elif estado == EST_INCOMPLETO:
                resumen["incompletos"] += 1
            elif estado in (EST_DESCANSO_TRAB, EST_FERIADO_TRAB):
                resumen["trabajados_descanso"] += 1
            elif estado == EST_DESCANSO:
                resumen["descansos"] += 1
            elif estado == EST_FERIADO:
                resumen["feriados"] += 1
            else:
                resumen["justificados"] += 1
            procesados += 1
            if progreso and procesados % 100 == 0:
                try:
                    progreso(procesados, len(universo) * len(dias))
                except Exception:
                    pass

    ok_lote, _, error_lote = _ejecutar_lote(
        "INSERT INTO nom_asistencia (dni, fecha, turno_id, hora_entrada, hora_salida, primera_marca, "
        "ultima_marca, n_marcas, minutos_tardanza, minutos_anticipo, minutos_trabajados, minutos_extra, "
        "minutos_falta, estado, manual, incidencia_id, observacion, actualizado) "
        "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) "
        "ON CONFLICT (dni, fecha) DO UPDATE SET turno_id=EXCLUDED.turno_id, "
        "hora_entrada=EXCLUDED.hora_entrada, hora_salida=EXCLUDED.hora_salida, "
        "primera_marca=EXCLUDED.primera_marca, ultima_marca=EXCLUDED.ultima_marca, "
        "n_marcas=EXCLUDED.n_marcas, minutos_tardanza=EXCLUDED.minutos_tardanza, "
        "minutos_anticipo=EXCLUDED.minutos_anticipo, minutos_trabajados=EXCLUDED.minutos_trabajados, "
        "minutos_extra=EXCLUDED.minutos_extra, minutos_falta=EXCLUDED.minutos_falta, "
        "estado=EXCLUDED.estado, incidencia_id=EXCLUDED.incidencia_id, "
        "observacion=EXCLUDED.observacion, actualizado=EXCLUDED.actualizado", filas_guardar)
    if not ok_lote:
        return {"ok": False, "error": error_lote}
    _sincronizar_horas_extra(dni_list=sorted(universo), desde=dias[0], hasta=dias[-1], usuario=usuario)
    resumen["ok"] = True
    registrar_auditoria(usuario, "Nomina", "Calculo asistencia %s a %s (%d dias)" %
                        (dias[0].strftime("%d/%m/%Y"), dias[-1].strftime("%d/%m/%Y"), resumen["dias"]))
    return resumen


def _sincronizar_horas_extra(dni_list, desde, hasta, usuario="sistema"):
    """Traslada los minutos extra calculados a la tabla de horas extra (origen AUTO).

    Importante: el trabajo en dia de descanso o feriado NO se registra como hora extra,
    porque en la planilla ya se paga como dia completo con recargo (concepto "Trabajo en
    Descanso / Feriado"). Evita pagar dos veces el mismo tiempo.

    Las filas automaticas del rango se reemplazan completas; las registradas a mano
    (origen MANUAL) se conservan intactas.
    """
    contexto = _cargar_contexto(desde, hasta)
    parametros = obtener_parametros()
    umbral25 = a_float(parametros.get("he_desde_hora"), 2) * 60
    registros = []
    for dni in dni_list:
        for fecha in rango_fechas(desde, hasta):
            calculo = calcular_dia(dni, fecha, contexto, parametros)
            if calculo["minutos_extra"] <= 0:
                continue
            if calculo["estado"] in (EST_DESCANSO_TRAB, EST_FERIADO_TRAB):
                continue
            tipo = "35" if calculo["minutos_extra"] > umbral25 else "25"
            if turno_nocturno(contexto, dni, fecha):
                tipo = "NOCTURNA_35" if tipo == "35" else "NOCTURNA_25"
            registros.append((dni, calculo["fecha"], calculo["minutos_extra"], tipo, "AUTO", True,
                              "Calculado automaticamente", datetime.now().strftime("%d/%m/%Y %H:%M")))
    conn = conectar_db(silencioso=True)
    if not conn:
        return 0
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                "DELETE FROM nom_horas_extra WHERE origen = 'AUTO' AND fecha BETWEEN %s AND %s "
                "AND dni = ANY(%s)",
                (desde.strftime("%Y-%m-%d"), hasta.strftime("%Y-%m-%d"), [str(d) for d in dni_list]))
            if registros:
                cursor.executemany(
                    "INSERT INTO nom_horas_extra (dni, fecha, minutos, tipo, origen, aprobado, "
                    "observacion, creado) VALUES (%s,%s,%s,%s,%s,%s,%s,%s) "
                    "ON CONFLICT (dni, fecha) DO UPDATE SET minutos = EXCLUDED.minutos, "
                    "tipo = EXCLUDED.tipo, observacion = EXCLUDED.observacion "
                    "WHERE nom_horas_extra.origen = 'AUTO'", registros)
        conn.commit()
    except Exception as e:
        print("[Nomina] Error sincronizando horas extra:", e)
        try:
            conn.rollback()
        except Exception:
            pass
        return 0
    finally:
        liberar_conexion(conn)
    return len(registros)


def turno_nocturno(contexto, dni, fecha):
    turno, _ = resolver_turno_del_dia(dni, fecha, contexto)
    if not turno:
        return False
    if turno.get("aplica_nocturno"):
        return True
    entrada = minutos_de_hora(turno["hora_entrada"]) or 0
    salida = minutos_de_hora(turno["hora_salida"]) or 0
    return salida <= entrada or entrada >= 19 * 60


# =========================================================================
# CONSULTAS DE ASISTENCIA PARA LA INTERFAZ
# =========================================================================
def listar_asistencia(desde, hasta, dnis=None, estados=None):
    columnas = ["id", "dni", "fecha", "turno_id", "hora_entrada", "hora_salida", "primera_marca",
                "ultima_marca", "n_marcas", "minutos_tardanza", "minutos_anticipo",
                "minutos_trabajados", "minutos_extra", "minutos_falta", "estado", "manual",
                "observacion"]
    sql = ("SELECT id, dni, fecha, turno_id, hora_entrada, hora_salida, primera_marca, ultima_marca, "
           "n_marcas, minutos_tardanza, minutos_anticipo, minutos_trabajados, minutos_extra, "
           "minutos_falta, estado, manual, observacion FROM nom_asistencia "
           "WHERE fecha BETWEEN %s AND %s ")
    params = [fecha_iso(desde), fecha_iso(hasta)]
    if dnis:
        sql += "AND dni = ANY(%s) "
        params.append(list(dnis))
    if estados:
        sql += "AND estado = ANY(%s) "
        params.append(list(estados))
    sql += "ORDER BY fecha DESC, dni"
    return [dict(zip(columnas, fila)) for fila in (_consultar(sql, tuple(params)) or [])]


def resumen_asistencia(desde, hasta, dnis=None):
    """Totales por empleado en el rango (para reportes y planilla)."""
    sql = ("SELECT dni, COUNT(*), "
           "SUM(CASE WHEN estado = 'PUNTUAL' THEN 1 ELSE 0 END), "
           "SUM(CASE WHEN estado = 'TARDANZA' THEN 1 ELSE 0 END), "
           "SUM(CASE WHEN estado = 'FALTA' THEN 1 ELSE 0 END), "
           "SUM(CASE WHEN estado = 'INCOMPLETO' THEN 1 ELSE 0 END), "
           "SUM(CASE WHEN estado IN ('DESCANSO','FERIADO') THEN 1 ELSE 0 END), "
           "SUM(CASE WHEN estado IN ('DESCANSO TRABAJADO','FERIADO TRABAJADO') THEN 1 ELSE 0 END), "
           "COALESCE(SUM(minutos_tardanza),0), COALESCE(SUM(minutos_extra),0), "
           "COALESCE(SUM(minutos_trabajados),0), COALESCE(SUM(minutos_falta),0), "
           "SUM(CASE WHEN estado NOT IN ('DESCANSO','FERIADO','PENDIENTE') THEN 1 ELSE 0 END) "
           "FROM nom_asistencia WHERE fecha BETWEEN %s AND %s ")
    params = [fecha_iso(desde), fecha_iso(hasta)]
    if dnis:
        sql += "AND dni = ANY(%s) "
        params.append(list(dnis))
    sql += "GROUP BY dni ORDER BY dni"
    columnas = ["dni", "dias", "puntuales", "tardanzas", "faltas", "incompletos", "descansos",
                "descansos_trabajados", "minutos_tardanza", "minutos_extra", "minutos_trabajados",
                "minutos_falta", "dias_laborables"]
    resultado = {}
    for fila in _consultar(sql, tuple(params)) or []:
        datos = dict(zip(columnas, fila))
        datos["dni"] = str(datos["dni"])
        for clave in ("dias", "puntuales", "tardanzas", "faltas", "incompletos", "descansos",
                      "descansos_trabajados", "minutos_tardanza", "minutos_extra",
                      "minutos_trabajados", "minutos_falta", "dias_laborables"):
            datos[clave] = int(datos[clave] or 0)
        resultado[datos["dni"]] = datos
    return resultado


def matriz_asistencia(periodo, dnis=None):
    """Devuelve las filas de la grilla mensual de asistencia.

    Estructura: {'dias': [date, ...], 'filas': [{'empleado':..., 'dni':..., 'celdas': {fecha: registro}}]}
    """
    dias = dias_del_periodo(periodo)
    if not dias:
        return {"dias": [], "filas": []}
    registros = listar_asistencia(dias[0], dias[-1], dnis=dnis)
    por_dni = {}
    for registro in registros:
        por_dni.setdefault(str(registro["dni"]), {})[registro["fecha"]] = registro
    empleados = listar_empleados()
    if dnis:
        empleados = [e for e in empleados if str(e["dni"]) in set(str(d) for d in dnis)]
    filas = []
    for empleado in empleados:
        dni = str(empleado["dni"])
        celdas = {}
        for dia in dias:
            clave = dia.strftime("%Y-%m-%d")
            celdas[clave] = por_dni.get(dni, {}).get(clave)
        filas.append({"dni": dni, "empleado": empleado["nombre"], "celdas": celdas,
                      "empleado_datos": empleado})
    return {"dias": dias, "filas": filas}


def registrar_asistencia_manual(dni, fecha, datos, usuario="sistema"):
    """Corrige manualmente un dia de asistencia (queda marcado como manual)."""
    dni = str(dni or "").strip()
    fecha_texto = fecha_iso(fecha)
    if not dni or not fecha_texto:
        return False, "Empleado y fecha son obligatorios"
    hora_entrada = parse_hora(datos.get("hora_entrada"))
    hora_salida = parse_hora(datos.get("hora_salida"))
    minutos_trabajados = 0
    if hora_entrada and hora_salida:
        entrada, salida = minutos_de_hora(hora_entrada), minutos_de_hora(hora_salida)
        if salida < entrada:
            salida += 24 * 60
        minutos_trabajados = max(0, salida - entrada)
    ok, error = _ejecutar(
        "INSERT INTO nom_asistencia (dni, fecha, turno_id, hora_entrada, hora_salida, primera_marca, "
        "ultima_marca, n_marcas, minutos_tardanza, minutos_anticipo, minutos_trabajados, minutos_extra, "
        "minutos_falta, estado, manual, observacion, actualizado) "
        "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,TRUE,%s,%s) "
        "ON CONFLICT (dni, fecha) DO UPDATE SET hora_entrada=EXCLUDED.hora_entrada, "
        "hora_salida=EXCLUDED.hora_salida, minutos_trabajados=EXCLUDED.minutos_trabajados, "
        "minutos_tardanza=EXCLUDED.minutos_tardanza, minutos_anticipo=EXCLUDED.minutos_anticipo, "
        "minutos_extra=EXCLUDED.minutos_extra, minutos_falta=EXCLUDED.minutos_falta, "
        "estado=EXCLUDED.estado, manual=TRUE, observacion=EXCLUDED.observacion, "
        "actualizado=EXCLUDED.actualizado",
        (dni, fecha_texto, int(datos["turno_id"]) if datos.get("turno_id") else None,
         hora_entrada, hora_salida, hora_entrada, hora_salida,
         2 if (hora_entrada and hora_salida) else (1 if hora_entrada else 0),
         a_int(datos.get("minutos_tardanza")), a_int(datos.get("minutos_anticipo")),
         minutos_trabajados, a_int(datos.get("minutos_extra")), a_int(datos.get("minutos_falta")),
         datos.get("estado", EST_JUSTIFICADO), datos.get("observacion", ""),
         datetime.now().strftime("%d/%m/%Y %H:%M")))
    if ok:
        registrar_auditoria(usuario, "Nomina", "Corrigio asistencia de %s el %s" % (dni, fecha_texto))
    return ok, error


def limpiar_asistencia_periodo(periodo, usuario="sistema"):
    dias = dias_del_periodo(periodo)
    if not dias:
        return False, "Periodo invalido"
    return _ejecutar("DELETE FROM nom_asistencia WHERE fecha BETWEEN %s AND %s",
                     (dias[0].strftime("%Y-%m-%d"), dias[-1].strftime("%Y-%m-%d")))


def detalle_marcaciones_dia(dni, fecha):
    """Todas las marcas crudas de un empleado en un dia (para el detalle del dia)."""
    return _consultar(
        "SELECT hora, tipo_pase, metodo, fuente FROM nom_marcaciones WHERE dni = %s AND fecha = %s "
        "ORDER BY hora", (str(dni), fecha_iso(fecha))) or []


# =========================================================================
# INDICADORES (DASHBOARD)
# =========================================================================
def kpis_dashboard(periodo=None):
    """Indicadores generales del modulo para el panel principal."""
    hoy = date.today()
    periodo = periodo or hoy.strftime("%Y-%m")
    dias = dias_del_periodo(periodo)
    indicadores = {
        "periodo": periodo,
        "empleados": 0, "empleados_activos": 0, "marcaciones": 0, "turnos": 0, "horarios": 0,
        "sin_mapeo": 0, "dias_calculados": 0, "puntuales": 0, "tardanzas": 0, "faltas": 0,
        "incompletos": 0, "minutos_tardanza": 0, "minutos_extra": 0, "horas_extra": 0.0,
        "feriados_mes": 0, "incidencias": 0, "ultima_importacion": None, "planilla_estado": "SIN CALCULAR",
        "planilla_neto": 0.0, "asistencia_hoy": 0, "sin_marcar_hoy": 0,
    }
    try:
        empleados = listar_empleados()
        indicadores["empleados"] = len(empleados)
        indicadores["empleados_activos"] = len([e for e in empleados
                                                if normalizar_texto(e.get("estado")) not in ("CESADO", "INACTIVO")])
        indicadores["turnos"] = len(_consultar("SELECT id FROM nom_turnos WHERE activo = TRUE") or [])
        indicadores["horarios"] = len(_consultar("SELECT id FROM nom_horarios WHERE activo = TRUE") or [])
        indicadores["marcaciones"] = int((_consultar("SELECT COUNT(*) FROM nom_marcaciones",
                                                     uno=True) or [0])[0] or 0)
        indicadores["sin_mapeo"] = len([p for p in personas_del_reloj() if not p["dni"]])
        if dias:
            resumen = resumen_asistencia(dias[0], dias[-1])
            for datos in resumen.values():
                indicadores["dias_calculados"] += datos["dias"]
                indicadores["puntuales"] += datos["puntuales"]
                indicadores["tardanzas"] += datos["tardanzas"]
                indicadores["faltas"] += datos["faltas"]
                indicadores["incompletos"] += datos["incompletos"]
                indicadores["minutos_tardanza"] += datos["minutos_tardanza"]
            # Las horas extra se toman de la tabla de horas extra (lo que realmente se
            # paga), no de la asistencia, para no contar el trabajo en dia de descanso.
            fila_extra = _consultar(
                "SELECT COALESCE(SUM(minutos), 0) FROM nom_horas_extra "
                "WHERE aprobado = TRUE AND fecha BETWEEN %s AND %s",
                (dias[0].strftime("%Y-%m-%d"), dias[-1].strftime("%Y-%m-%d")), uno=True)
            indicadores["minutos_extra"] = a_int(fila_extra[0]) if fila_extra else 0
            indicadores["horas_extra"] = round(indicadores["minutos_extra"] / 60.0, 2)
        indicadores["feriados_mes"] = len([d for d in listar_calendario()
                                           if d["fecha"].startswith(periodo)])
        indicadores["incidencias"] = len(listar_incidencias(desde=dias[0] if dias else None,
                                                            hasta=dias[-1] if dias else None))
        importaciones = listar_importaciones(limite=1)
        indicadores["ultima_importacion"] = importaciones[0] if importaciones else None
        planilla = _consultar("SELECT estado, total_neto FROM nom_planilla WHERE periodo = %s",
                             (periodo,), uno=True)
        if planilla:
            indicadores["planilla_estado"] = planilla[0] or "BORRADOR"
            indicadores["planilla_neto"] = a_float(planilla[1])
        hoy_texto = hoy.strftime("%Y-%m-%d")
        indicadores["asistencia_hoy"] = int((_consultar(
            "SELECT COUNT(DISTINCT dni) FROM nom_marcaciones WHERE fecha = %s", (hoy_texto,),
            uno=True) or [0])[0] or 0)
        indicadores["sin_marcar_hoy"] = max(0, indicadores["empleados_activos"] - indicadores["asistencia_hoy"])
    except Exception as e:
        print("[Nomina] Error calculando indicadores:", e)
    return indicadores


# =========================================================================
# ASISTENTE PASO A PASO
# =========================================================================
FASES_ASISTENTE = ["Puesta en marcha", "Cierre del mes"]


def _crear_paso(clave, fase, titulo, estado, detalle, ayuda, destino, critico=False, accion=None):
    """Arma la definicion de un paso del asistente."""
    return {"clave": clave, "fase": fase, "titulo": titulo, "estado": estado, "detalle": detalle,
            "ayuda": ayuda, "destino": destino or ("", ""), "critico": bool(critico),
            "accion": accion}


def _nombres_cortos(lista, limite=3):
    """'Ana, Luis y 2 mas' para los mensajes del asistente."""
    nombres = [str(e.get("nombre") or e.get("dni") or "?").split()[0] for e in lista]
    if len(nombres) <= limite:
        return ", ".join(nombres)
    return "%s y %d mas" % (", ".join(nombres[:limite]), len(nombres) - limite)


def fecha_para_mostrar(valor):
    """'2026-09-09' -> '09/09/2026' (para los textos del asistente)."""
    fecha = parse_fecha(valor)
    return fecha.strftime("%d/%m/%Y") if fecha else str(valor or "")


def formatear_soles(valor):
    """'S/ 5,462.76' para los textos del asistente."""
    return "S/ %s" % format(a_float(valor), ",.2f")


def estado_configuracion(periodo=None):
    """Revisa el estado real del modulo para el asistente paso a paso.

    Devuelve la lista de pasos (cada uno con su estado, el detalle de lo que falta, una
    explicacion y el lugar de la interfaz donde se resuelve), mas un resumen con el avance
    y el primer paso pendiente.
    """
    periodo = periodo or date.today().strftime("%Y-%m")
    dias = dias_del_periodo(periodo)
    anio = int(str(periodo).split("-")[0])
    pasos = []

    # ---------------------------------------------------------------- padron
    empleados = listar_empleados()
    activos = [e for e in empleados if normalizar_texto(e.get("estado")) not in ("CESADO", "INACTIVO")]
    sin_ficha = [e for e in empleados if not e.get("en_nomina")]
    if not empleados:
        pasos.append(_crear_paso(
            "empleados", FASES_ASISTENTE[0], "Padron de empleados", "PENDIENTE",
            "Todavia no hay ningun empleado en el modulo.",
            "Registre a su personal en el modulo Control de Choferes (menu Modulos Operativos). "
            "Cuando esten creados, vuelva aqui y pulse Sincronizar padron.",
            ("Personal", "Empleados"), True, "sincronizar"))
    elif sin_ficha:
        pasos.append(_crear_paso(
            "empleados", FASES_ASISTENTE[0], "Padron de empleados", "ATENCION",
            "%d de %d empleados no tienen ficha de nomina." % (len(sin_ficha), len(empleados)),
            "El modulo toma los empleados de la tabla de Choferes, pero necesita crear su ficha "
            "de nomina para poder calcular sueldos. Pulse Ejecutar para crearlas.",
            ("Personal", "Empleados"), True, "sincronizar"))
    else:
        pasos.append(_crear_paso(
            "empleados", FASES_ASISTENTE[0], "Padron de empleados", "OK",
            "%d empleados, %d activos, todos con ficha de nomina." % (len(empleados), len(activos)),
            "Si ingresa personal nuevo, registrelo en Choferes y vuelva a sincronizar el padron.",
            ("Personal", "Empleados"), True))

    # ---------------------------------------------------------------- turnos
    turnos = listar_turnos(solo_activos=True)
    if not turnos:
        pasos.append(_crear_paso(
            "turnos", FASES_ASISTENTE[0], "Turnos de trabajo", "PENDIENTE",
            "No hay ningun turno activo.",
            "El turno define la hora de entrada, la de salida, la tolerancia y los dias que se "
            "trabajan: es lo que permite saber si alguien llego tarde o falto. Cree al menos uno "
            "con los horarios reales de su empresa.",
            ("Personal", "Turnos"), True))
    else:
        nombres = ", ".join(t["nombre"].split("(")[0].strip() for t in turnos[:3])
        pasos.append(_crear_paso(
            "turnos", FASES_ASISTENTE[0], "Turnos de trabajo", "OK",
            "%d turno(s) activo(s): %s%s" % (len(turnos), nombres, "..." if len(turnos) > 3 else ""),
            "Revise que la tolerancia y los minutos de refrigerio correspondan a su politica interna.",
            ("Personal", "Turnos"), True))

    # --------------------------------------------------------------- horario
    horarios = listar_horarios(solo_activos=True)
    if not horarios:
        pasos.append(_crear_paso(
            "horarios", FASES_ASISTENTE[0], "Horarios semanales", "PENDIENTE",
            "No hay horarios creados.",
            "El horario es la plantilla de la semana: que turno le toca a cada dia. Se arma una "
            "vez y se reutiliza en todos los empleados que trabajan igual.",
            ("Personal", "Horarios"), True))
    else:
        pasos.append(_crear_paso(
            "horarios", FASES_ASISTENTE[0], "Horarios semanales", "OK",
            "%d horario(s) activo(s): %s" % (len(horarios), ", ".join(h["nombre"] for h in horarios[:3])),
            "Puede duplicar un horario para crear variantes, por ejemplo cambiando solo el sabado.",
            ("Personal", "Horarios"), True))

    # ----------------------------------------------------- asignar el horario
    sin_horario = [e for e in activos if not e.get("horario_id") and not e.get("turno_id")]
    if not activos:
        pasos.append(_crear_paso(
            "asignacion", FASES_ASISTENTE[0], "Asignar horario a cada empleado", "PENDIENTE",
            "No hay empleados activos.", "Primero complete el paso del padron de empleados.",
            ("Personal", "Empleados"), True))
    elif sin_horario:
        pasos.append(_crear_paso(
            "asignacion", FASES_ASISTENTE[0], "Asignar horario a cada empleado", "ATENCION",
            "%d empleado(s) sin turno ni horario: %s." % (len(sin_horario), _nombres_cortos(sin_horario)),
            "Sin horario asignado, TODOS sus dias se calculan como descanso: no se detectan "
            "tardanzas ni faltas y la planilla les sale sin descuentos. Seleccione a cada uno en la "
            "lista y asignele su horario, o use la asignacion masiva si todos trabajan igual.",
            ("Personal", "Empleados"), True))
    else:
        pasos.append(_crear_paso(
            "asignacion", FASES_ASISTENTE[0], "Asignar horario a cada empleado", "OK",
            "Los %d empleados activos ya tienen horario o turno asignado." % len(activos),
            "Para turnos rotativos use Asignaciones por rango de fechas.",
            ("Personal", "Empleados"), True))

    # ---------------------------------------------------------------- sueldos
    sin_sueldo = [e for e in activos if a_float(e.get("sueldo_basico")) <= 0]
    afp_sin_comision = [e for e in activos if normalizar_texto(e.get("sistema_pension")) == "AFP"
                        and a_float(e.get("afp_comision_pct")) <= 0]
    if not activos:
        pasos.append(_crear_paso(
            "sueldos", FASES_ASISTENTE[0], "Sueldos y datos de pago", "PENDIENTE",
            "No hay empleados activos.", "Primero complete el paso del padron de empleados.",
            ("Personal", "Empleados"), True))
    elif sin_sueldo:
        pasos.append(_crear_paso(
            "sueldos", FASES_ASISTENTE[0], "Sueldos y datos de pago", "PENDIENTE",
            "%d empleado(s) sin sueldo basico: %s." % (len(sin_sueldo), _nombres_cortos(sin_sueldo)),
            "Sin sueldo, la planilla se calcula en cero. Complete tambien el sistema de pension "
            "(ONP o AFP); si es AFP, indique el nombre y el porcentaje de comision mas prima.",
            ("Personal", "Empleados"), True))
    elif afp_sin_comision:
        pasos.append(_crear_paso(
            "sueldos", FASES_ASISTENTE[0], "Sueldos y datos de pago", "ATENCION",
            "%d empleado(s) en AFP sin comision registrada: %s."
            % (len(afp_sin_comision), _nombres_cortos(afp_sin_comision)),
            "Si el porcentaje de comision queda en cero, se descontara solo el 10% de aporte "
            "obligatorio y faltara la comision y la prima del seguro.",
            ("Personal", "Empleados"), True))
    else:
        pasos.append(_crear_paso(
            "sueldos", FASES_ASISTENTE[0], "Sueldos y datos de pago", "OK",
            "Los %d empleados activos tienen sueldo y regimen de pension registrados." % len(activos),
            "Revise la asignacion familiar y los adelantos o retenciones judiciales si corresponden.",
            ("Personal", "Empleados"), True))

    # ------------------------------------------------------------- parametros
    parametros = obtener_parametros()
    _limpios, errores_param, avisos_param = validar_parametros(
        {clave: parametros.get(clave) for clave in PARAMETROS_DEFECTO})
    if errores_param:
        pasos.append(_crear_paso(
            "parametros", FASES_ASISTENTE[0], "Parametros de calculo", "ATENCION",
            "Hay %d parametro(s) con valor invalido." % len(errores_param),
            "Un parametro mal escrito o vacio puede dejar el calculo en cero. Pulse Ejecutar para "
            "restaurar los valores danados a su valor por defecto, o edítelos a mano.",
            ("Configuracion", "Parametros de calculo"), False, "reparar"))
    elif avisos_param:
        pasos.append(_crear_paso(
            "parametros", FASES_ASISTENTE[0], "Parametros de calculo", "ATENCION",
            avisos_param[0],
            "Revise la configuracion antes de calcular la planilla. Por ejemplo, con el tope diario "
            "de horas extra en 0 no se reconoceria ninguna hora extra.",
            ("Configuracion", "Parametros de calculo"), False))
    else:
        pasos.append(_crear_paso(
            "parametros", FASES_ASISTENTE[0], "Parametros de calculo", "OK",
            "Los %d parametros tienen valores validos." % len(PARAMETROS_DEFECTO),
            "Confirme que la RMV y la UIT correspondan al ano en curso, y que los porcentajes de "
            "ONP, ESSALUD y horas extra sean los que usa su empresa.",
            ("Configuracion", "Parametros de calculo"), False))

    # --------------------------------------------------------------- feriados
    feriados = listar_calendario(anio)
    if not feriados:
        pasos.append(_crear_paso(
            "feriados", FASES_ASISTENTE[0], "Feriados del ano", "PENDIENTE",
            "No hay dias registrados en el calendario de %d." % anio,
            "Los feriados evitan que un dia festivo se cuente como falta y permiten pagar con "
            "recargo el trabajo en esos dias. Pulse Ejecutar para cargar los feriados nacionales "
            "del Peru del ano.",
            ("Configuracion", "Feriados y calendario"), False, "feriados"))
    else:
        pasos.append(_crear_paso(
            "feriados", FASES_ASISTENTE[0], "Feriados del ano", "OK",
            "%d dia(s) registrados en el calendario de %d." % (len(feriados), anio),
            "Agregue los dias no laborables propios de la empresa si los tuviera.",
            ("Configuracion", "Feriados y calendario"), False))

    # ----------------------------------------------------- vincular el reloj
    personas = personas_del_reloj()
    sin_vincular = [p for p in personas if not p.get("dni")]
    if not personas:
        pasos.append(_crear_paso(
            "vinculacion", FASES_ASISTENTE[1], "Vincular el reloj con los DNI", "PENDIENTE",
            "Todavia no se ha importado ningun archivo del reloj.",
            "Cuando importe su primer export, aqui vera si todas las personas del reloj quedaron "
            "reconocidas. Mientras un codigo no este vinculado a un DNI, sus marcaciones no entran "
            "en el calculo.",
            ("Marcaciones", "Personas del reloj"), False))
    elif sin_vincular:
        pasos.append(_crear_paso(
            "vinculacion", FASES_ASISTENTE[1], "Vincular el reloj con los DNI", "PENDIENTE",
            "%d de %d personas del reloj sin DNI: %s."
            % (len(sin_vincular), len(personas),
               ", ".join((p.get("nombre_reloj") or p["codigo_reloj"])[:24] for p in sin_vincular[:3])),
            "Los equipos guardan un codigo interno que no siempre es el DNI. Vincule cada codigo "
            "con su empleado (hay un boton para vincular automaticamente los que coinciden).",
            ("Marcaciones", "Personas del reloj"), False))
    else:
        pasos.append(_crear_paso(
            "vinculacion", FASES_ASISTENTE[1], "Vincular el reloj con los DNI", "OK",
            "Las %d personas del reloj estan vinculadas a un DNI." % len(personas),
            "Si aparece personal nuevo en el reloj, vuelva a esta pantalla despues de importar.",
            ("Marcaciones", "Personas del reloj"), False))

    # ------------------------------------------------- cierre del mes: marcas
    total_marcas = 0
    ultima_marca = ""
    if dias:
        fila = _consultar("SELECT COUNT(*) FROM nom_marcaciones WHERE fecha BETWEEN %s AND %s",
                          (dias[0].strftime("%Y-%m-%d"), dias[-1].strftime("%Y-%m-%d")), uno=True)
        total_marcas = a_int(fila[0]) if fila else 0
    fila_ultima = _consultar("SELECT MAX(fecha) FROM nom_marcaciones", uno=True)
    ultima_marca = (fila_ultima[0] or "") if fila_ultima else ""
    if not total_marcas:
        detalle = "No hay marcaciones del reloj en %s." % nombre_mes(periodo)
        if ultima_marca:
            detalle += " Las ultimas cargadas son del %s." % fecha_para_mostrar(ultima_marca)
        pasos.append(_crear_paso(
            "marcaciones", FASES_ASISTENTE[1], "Importar las marcaciones del mes", "PENDIENTE",
            detalle,
            "Exporte el periodo desde el reloj biometrico y carguelo aqui. Sin marcaciones, todos "
            "los dias del mes saldrian como falta.",
            ("Marcaciones", "Importar Excel del reloj"), True))
    elif dias and ultima_marca and ultima_marca < dias[-1].strftime("%Y-%m-%d"):
        pasos.append(_crear_paso(
            "marcaciones", FASES_ASISTENTE[1], "Importar las marcaciones del mes", "ATENCION",
            "Hay %s marcas en %s, pero el archivo cargado llega solo hasta el %s."
            % (format(total_marcas, ","), nombre_mes(periodo), fecha_para_mostrar(ultima_marca)),
            "Si recalcula el mes completo, los dias posteriores al %s se contaran como falta. "
            "Importe el export mas reciente antes de continuar." % fecha_para_mostrar(ultima_marca),
            ("Marcaciones", "Importar Excel del reloj"), True))
    else:
        pasos.append(_crear_paso(
            "marcaciones", FASES_ASISTENTE[1], "Importar las marcaciones del mes", "OK",
            "%s marcas cargadas en %s." % (format(total_marcas, ","), nombre_mes(periodo)),
            "Puede volver a importar el mismo archivo cuando quiera: las marcas repetidas se "
            "descartan solas.",
            ("Marcaciones", "Importar Excel del reloj"), True))

    # --------------------------------------------- cierre del mes: asistencia
    dias_calculados = 0
    if dias:
        fila = _consultar("SELECT COUNT(*) FROM nom_asistencia WHERE fecha BETWEEN %s AND %s",
                          (dias[0].strftime("%Y-%m-%d"), dias[-1].strftime("%Y-%m-%d")), uno=True)
        dias_calculados = a_int(fila[0]) if fila else 0
    if not dias_calculados:
        pasos.append(_crear_paso(
            "asistencia", FASES_ASISTENTE[1], "Recalcular la asistencia", "PENDIENTE",
            "La asistencia de %s todavia no se ha calculado." % nombre_mes(periodo),
            "El recalculo cruza la primera y la ultima marcacion de cada dia con el turno asignado "
            "y calcula tardanzas, faltas y horas extra. Pulse Ejecutar para hacerlo ahora.",
            ("Asistencia", "Matriz mensual"), True, "recalcular"))
    else:
        resumen = resumen_asistencia(dias[0], dias[-1]) if dias else {}
        totales = {"puntuales": 0, "tardanzas": 0, "faltas": 0, "incompletos": 0}
        for datos in resumen.values():
            for clave in totales:
                totales[clave] += a_int(datos.get(clave))
        pasos.append(_crear_paso(
            "asistencia", FASES_ASISTENTE[1], "Recalcular la asistencia", "OK",
            "%d dias calculados: %d puntuales, %d tardanzas, %d faltas, %d incompletos."
            % (dias_calculados, totales["puntuales"], totales["tardanzas"], totales["faltas"],
               totales["incompletos"]),
            "Revise la grilla del mes en la pestana Asistencia y registre las incidencias "
            "(vacaciones, permisos, descansos medicos) antes de calcular la planilla. Cada vez que "
            "cambie un turno o una incidencia, vuelva a recalcular.",
            ("Asistencia", "Matriz mensual"), True, "recalcular"))

    # ------------------------------------------------ cierre del mes: planilla
    planilla = None
    try:
        import nomina_planilla as _planilla
        planilla = _planilla.obtener_planilla(periodo)
    except Exception as e:
        print("[Nomina] Asistente: no se pudo leer la planilla:", e)
    if not planilla:
        pasos.append(_crear_paso(
            "planilla", FASES_ASISTENTE[1], "Calcular la planilla", "PENDIENTE",
            "La planilla de %s aun no se ha calculado." % nombre_mes(periodo),
            "El calculo genera los haberes, los descuentos por tardanzas y faltas, los aportes y el "
            "neto a pagar de cada empleado. Pulse Ejecutar para calcularla con la asistencia ya "
            "revisada.",
            ("Planilla", "Resumen del periodo"), True, "planilla"))
    elif planilla["estado"] in ("BORRADOR", "CALCULADA"):
        pasos.append(_crear_paso(
            "planilla", FASES_ASISTENTE[1], "Calcular la planilla", "ATENCION",
            "La planilla de %s esta %s, con un neto de %s."
            % (nombre_mes(periodo), planilla["estado"], formatear_soles(planilla["total_neto"])),
            "Revise las boletas de cada empleado, exporte la planilla a Excel y despues cambie el "
            "estado a APROBADA y luego PAGADA para dejar constancia del cierre.",
            ("Planilla", "Resumen del periodo"), True))
    else:
        pasos.append(_crear_paso(
            "planilla", FASES_ASISTENTE[1], "Calcular la planilla", "OK",
            "Planilla de %s en estado %s, neto %s."
            % (nombre_mes(periodo), planilla["estado"], formatear_soles(planilla["total_neto"])),
            "Exporte la planilla y las boletas a Excel y archivelas como respaldo del mes.",
            ("Planilla", "Resumen del periodo"), True))

    # ------------------------------------------------------------------ orden final
    # El orden de los pasos sigue el trabajo real: primero se arma la puesta en marcha y,
    # en el ciclo mensual, se importan las marcaciones ANTES de vincular las personas
    # nuevas del reloj (sin marcaciones no hay nada que vincular).
    orden_pasos = ["empleados", "turnos", "horarios", "asignacion", "sueldos", "parametros",
                   "feriados", "marcaciones", "vinculacion", "asistencia", "planilla"]
    pasos.sort(key=lambda paso: orden_pasos.index(paso["clave"]) if paso["clave"] in orden_pasos else 99)

    # ------------------------------------------------------------------ resumen
    pendientes = [p for p in pasos if p["estado"] == "PENDIENTE"]
    atenciones = [p for p in pasos if p["estado"] == "ATENCION"]
    listos = [p for p in pasos if p["estado"] == "OK"]
    criticos = [p for p in pasos if p["critico"]]
    criticos_ok = [p for p in criticos if p["estado"] == "OK"]
    siguiente = None
    for paso in pasos:
        if paso["estado"] != "OK":
            siguiente = paso
            break
    return {
        "periodo": periodo,
        "pasos": pasos,
        "pendientes": pendientes,
        "atenciones": atenciones,
        "listos": listos,
        "avance": int(round(100.0 * len(listos) / len(pasos))) if pasos else 100,
        "criticos_ok": len(criticos_ok),
        "criticos_total": len(criticos),
        "siguiente": siguiente,
        "listo": len(criticos_ok) == len(criticos),
    }


# =========================================================================
# EXPORTACION A EXCEL
# =========================================================================
def exportar_excel(ruta, hojas):
    """Exporta un diccionario {nombre_hoja: (columnas, filas)} a un archivo .xlsx."""
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    except Exception:
        return False, "Falta la libreria openpyxl. Ejecute: pip install openpyxl"
    try:
        libro = openpyxl.Workbook()
        libro.remove(libro.active)
        borde = Border(*[Side(style="thin", color="D0D0D0")] * 4)
        for nombre_hoja, contenido in hojas.items():
            columnas, filas = contenido[0], contenido[1]
            hoja = libro.create_sheet(title=str(nombre_hoja)[:31])
            hoja.append(list(columnas))
            for celda in hoja[1]:
                celda.font = Font(bold=True, color="FFFFFF")
                celda.fill = PatternFill("solid", fgColor="1F538D")
                celda.alignment = Alignment(horizontal="center", vertical="center")
                celda.border = borde
            for fila in filas:
                hoja.append(list(fila))
            for indice, titulo in enumerate(columnas, start=1):
                ancho = max(12, min(42, len(str(titulo)) + 4))
                for fila in filas[:200]:
                    try:
                        ancho = max(ancho, min(42, len(str(fila[indice - 1])) + 2))
                    except Exception:
                        pass
                hoja.column_dimensions[openpyxl.utils.get_column_letter(indice)].width = ancho
            hoja.freeze_panes = "A2"
        libro.save(ruta)
        return True, ""
    except Exception as e:
        return False, str(e)


def exportar_asistencia_excel(ruta, periodo):
    """Genera el reporte de asistencia mensual en Excel."""
    matriz = matriz_asistencia(periodo)
    if not matriz["dias"]:
        return False, "Periodo invalido"
    encabezados = ["DNI", "Empleado"] + [d.strftime("%d") for d in matriz["dias"]] + \
                  ["Puntuales", "Tardanzas", "Faltas", "Incompletos", "Horas trabajadas", "Horas extra"]
    filas = []
    for fila in matriz["filas"]:
        valores = [fila["dni"], fila["empleado"]]
        puntuales = tardanzas = faltas = incompletos = 0
        minutos_trabajados = minutos_extra = 0
        for dia in matriz["dias"]:
            registro = fila["celdas"].get(dia.strftime("%Y-%m-%d"))
            if not registro:
                valores.append("")
                continue
            estado = registro["estado"]
            marca = ""
            if registro.get("hora_entrada"):
                marca = registro["hora_entrada"]
                if registro.get("hora_salida"):
                    marca += "-" + registro["hora_salida"]
            if estado == EST_PUNTUAL:
                puntuales += 1
            elif estado == EST_TARDANZA:
                tardanzas += 1
            elif estado == EST_FALTA:
                faltas += 1
            elif estado == EST_INCOMPLETO:
                incompletos += 1
            minutos_trabajados += int(registro.get("minutos_trabajados") or 0)
            minutos_extra += int(registro.get("minutos_extra") or 0)
            valores.append("%s %s" % (estado[:3], marca) if estado in ESTADOS_TRABAJADOS else estado[:6])
        valores += [puntuales, tardanzas, faltas, incompletos,
                    round(minutos_trabajados / 60.0, 2), round(minutos_extra / 60.0, 2)]
        filas.append(valores)
    return exportar_excel(ruta, {"Asistencia": (encabezados, filas)})
