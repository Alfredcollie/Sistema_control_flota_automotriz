# -*- coding: utf-8 -*-
"""
=================================================================
CALCULO_COBRANZA.PY — CÁLCULO DE COBRANZA QUINCENAL POR CLIENTE
=================================================================
Funcionalidades:
1. Selecciona un cliente de la base de datos y muestra su plan de
   cobro (Por Hora / Por Punto o Viaje).
2. Unidades asignadas al cliente (desde la Flota Automotriz): cada
   unidad maneja sus propios precios (Día Normal, Domingo, Feriado)
   y su cantidad de horas al día (jornada).
3. Selecciona mes, año y quincena (1ª: del 1 al 15 / 2ª: del 16 al fin).
4. Busca en internet los feriados oficiales de Perú del año (API pública
   de feriados Nager.Date y respaldo Google Calendar ICS); si no hay
   internet usa la lista local de feriados peruanos (Ley 31822 + Semana
   Santa móvil) y guarda caché local para próximas consultas.
5. Clasifica cada día de la quincena en 3 categorías: Día Normal (Lunes a
   Sábado), Domingo y Feriado (los feriados se pueden marcar/desmarcar
   manualmente).
6. Cálculo con desglose POR UNIDAD: cada unidad aporta su base
   (días × sus precios) menos sus deducciones (horas/minutos de servicio
   no prestado, o viajes según el plan). El total del cliente es la suma
   de todas sus unidades.
   - Plan Por Hora: cobro diario = precio por hora × horas del día;
     deducción = horas ausentes × precio por hora.
   - Plan Por Punto o Viaje: el precio por viaje depende de la distancia
     (tabla Distancia → Precio normal/domingo/feriado); deducción = viajes
     no realizados × precio según distancia y tipo de día.
   - Las deducciones se registran como ASIENTOS: cada asiento lleva su
     FECHA (el tipo de día se detecta solo), las horas/minutos o la
     cantidad de viajes no realizados y un motivo. Se pueden registrar
     varios asientos por unidad o por rango de distancia y se detallan
     uno por uno en el PDF.
   - Además de las deducciones, cada unidad puede registrar ASIENTOS DE
     HORAS EXTRAS con fecha: se valorizan al precio por hora de la unidad
     según el tipo de día (Normal / Domingo / Feriado) y SUMAN al cobro:
     subtotal = base − deducciones + horas extras.
7. Guarda la quincena en la base de datos (con detalle por unidad y por
   asiento de deducción) y permite editar/eliminar los registros desde
   una ventana aparte.
8. Genera un PDF con nombre del cliente, cuadro explicativo del cobro,
   detalle por unidad y el detalle de los asientos de deducción.
"""
import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk
import calendar
import os
import sys
import json
import ssl
import urllib.request
import threading
from datetime import datetime, date, timedelta

# 🚀 HERRAMIENTAS CORPORATIVAS
from conexion import conectar_db, registrar_auditoria, liberar_conexion
from app_paths import CONFIG_FILE, ruta_para_guardar

try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.utils import ImageReader
    REPORTLAB_DISPONIBLE = True
except ImportError:
    REPORTLAB_DISPONIBLE = False

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

PLANES = ["Por Hora", "Por Punto o Viaje"]
CATEGORIAS = [
    ("normal", "Días Normales (Lun–Sáb)"),
    ("domingo", "Domingos"),
    ("feriado", "Feriados"),
]
# Etiquetas cortas para los asientos de deducción (fecha → tipo de día)
ETIQUETAS_CORTAS = {"normal": "Normal", "domingo": "Domingo", "feriado": "Feriado"}
ETIQUETAS_CORTAS_A_CLAVE = {v: k for k, v in ETIQUETAS_CORTAS.items()}
NOMBRES_MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                 "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
CARPETA_PDF = "cobranzas_generadas"
ARCHIVO_CACHE_FERIADOS = "feriados_peru.json"


# =========================================================
# 🚀 UTILIDADES MULTIPLATAFORMA
# =========================================================
def abrir_documento(ruta):
    try:
        # Resuelve rutas guardadas en otro equipo/SO (Mac <-> Windows)
        try:
            from app_paths import resolver_ruta_archivo
            ruta = resolver_ruta_archivo(ruta) or ruta
        except Exception:
            pass
        ruta_abs = os.path.abspath(ruta)
        if sys.platform == "win32":
            os.startfile(ruta_abs)
        elif sys.platform == "darwin":
            import subprocess
            subprocess.call(["open", ruta_abs])
        else:
            import subprocess
            subprocess.call(["xdg-open", ruta_abs])
    except Exception as e:
        messagebox.showerror(
            "No se pudo abrir el PDF",
            f"No se pudo abrir el archivo.\n\nArchivo: {os.path.abspath(ruta)}\n\n"
            f"Puedes abrirlo manualmente desde esa ubicación.\nDetalle: {e}")


def _archivo_bloqueado(ruta):
    """Devuelve True si el archivo existe y está bloqueado (abierto por otro programa)."""
    try:
        if not os.path.exists(ruta):
            return False
        with open(ruta, "r+b"):
            pass
        return False
    except PermissionError:
        return True
    except Exception:
        return False


def _cargar_config_local():
    try:
        if os.path.exists(str(CONFIG_FILE)):
            with open(str(CONFIG_FILE), "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def obtener_ruta_logo():
    return str(_cargar_config_local().get("ruta_logo_cotizacion", "") or "")


def _obtener_logo_pdf():
    """Devuelve la ruta del logo para el PDF:
    primero el configurado, luego Logo.png / logo.png de la carpeta del programa."""
    ruta = obtener_ruta_logo()
    if ruta and os.path.exists(ruta):
        return ruta
    base = os.path.dirname(os.path.abspath(__file__))
    for nombre in ("Logo.png", "logo.png", "Logo_Collie_Software.png"):
        for carpeta in (base, os.getcwd()):
            ruta = os.path.join(carpeta, nombre)
            if os.path.exists(ruta):
                return ruta
    return ""


def _carpeta_pdf():
    """Carpeta donde se guardan los PDFs de cobranza.

    Usa la ruta configurada en 'Configuración del Sistema' (ruta_drive) y crea
    dentro la subcarpeta 'cobranzas_generadas'. Compatible con Windows y macOS:
    en sistemas que no sean Windows se ignoran rutas con prefijo de unidad (C:)
    y se normalizan los separadores (igual que normalizar_ruta_local de
    control_general.py). Si la ruta configurada no existe o no es válida, usa
    la carpeta local 'cobranzas_generadas' como respaldo.
    """
    try:
        from politica_almacenamiento import estado_almacenamiento
        if not estado_almacenamiento().get("autorizado"):
            # Equipo NO autorizado: no se guarda nada (ni local ni nube)
            return ""
    except Exception:
        pass
    try:
        config = _cargar_config_local()
        ruta_drive = str(config.get("ruta_drive", "") or "").strip()
        if ruta_drive:
            ruta = ruta_drive
            if sys.platform != "win32":
                if len(ruta) >= 2 and ruta[1] == ":" and ruta[0].isalpha():
                    ruta = ""  # ruta de Windows en Mac/Linux: se ignora
                else:
                    ruta = ruta.replace("\\", "/")
            ruta = os.path.expanduser(ruta)
            if ruta and os.path.isdir(ruta):
                # 1) Subcarpeta organizada
                destino = os.path.join(ruta, "cobranzas_generadas")
                try:
                    os.makedirs(destino, exist_ok=True)
                    return destino
                except Exception:
                    pass
                # 2) Directamente en la carpeta configurada (si no se pudo crear la subcarpeta)
                try:
                    prueba = os.path.join(ruta, ".cobranza_prueba_escritura")
                    with open(prueba, "w", encoding="utf-8") as f:
                        f.write("x")
                    os.remove(prueba)
                    return ruta
                except Exception:
                    pass
    except Exception:
        pass
    # Respaldo local: ruta ABSOLUTA junto al programa (no depende del directorio de trabajo).
    # En macOS (apps .app) el directorio de trabajo puede ser "/" o no escribible, por eso
    # una ruta relativa como "cobranzas_generadas" causaría Errno 2 al guardar el PDF.
    # Respaldo junto al programa: SOLO en modo monousuario (local).
    try:
        from politica_almacenamiento import permitir_respaldo_local
        if not permitir_respaldo_local():
            return ""
    except Exception:
        pass
    carpeta_local = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cobranzas_generadas")
    try:
        os.makedirs(carpeta_local, exist_ok=True)
    except Exception:
        pass
    return carpeta_local


def formatear_moneda(valor):
    config = _cargar_config_local()
    simbolo = config.get("simbolo_moneda", "S/.")
    formato = config.get("formato_numero", "1,000.00")
    try:
        valor = float(valor)
    except Exception:
        valor = 0.0
    if formato == "1.000,00":
        str_val = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    else:
        str_val = f"{valor:,.2f}"
    return f"{simbolo} {str_val}"


def _recortar_texto(c, texto, ancho_max, fuente="Helvetica", tam=10):
    """Recorta el texto con '...' para que NUNCA se salga del ancho indicado."""
    t = str(texto if texto is not None else "")
    if c.stringWidth(t, fuente, tam) <= ancho_max:
        return t
    while t and c.stringWidth(t + "...", fuente, tam) > ancho_max:
        t = t[:-1]
    return (t + "...") if t else ""


def _envolver_texto(c, texto, ancho_max, max_lineas=2, fuente="Helvetica", tam=10):
    """Parte el texto en líneas que quepan en 'ancho_max' (máximo 'max_lineas').

    Cada línea devuelta cumple el ancho máximo, así el contenido nunca se sale
    de la hoja; si sobra texto, la última línea termina en '...'.
    """
    lineas, actual = [], ""
    for palabra in str(texto if texto is not None else "").split():
        prueba = f"{actual} {palabra}".strip()
        if actual and c.stringWidth(prueba, fuente, tam) > ancho_max:
            lineas.append(actual)
            actual = palabra
        else:
            actual = prueba
    if actual:
        lineas.append(actual)
    if len(lineas) > max_lineas:
        lineas = lineas[:max_lineas]
        cola = lineas[-1]
        while cola and c.stringWidth(cola + "...", fuente, tam) > ancho_max:
            cola = cola[:-1]
        lineas[-1] = cola + "..."
    return [_recortar_texto(c, linea, ancho_max, fuente, tam) for linea in lineas]


def _ruta_cache_feriados():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), ARCHIVO_CACHE_FERIADOS)


def _leer_cache_feriados():
    try:
        with open(_ruta_cache_feriados(), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _guardar_cache_feriados(anio, feriados):
    try:
        data = _leer_cache_feriados()
        data[str(anio)] = {d.isoformat(): n for d, n in feriados.items()}
        with open(_ruta_cache_feriados(), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
    except Exception:
        pass


# =========================================================
# 🚀 MOTOR DE FERIADOS DE PERÚ (Internet + Lista Local + Caché)
# =========================================================
def _easter(year):
    """Algoritmo de Computus (Meeus) para calcular el Domingo de Pascua."""
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)


def feriados_fijos_peru(anio):
    """Feriados oficiales de Perú (Ley 31822) + Semana Santa móvil."""
    e = _easter(anio)
    return {
        date(anio, 1, 1): "Año Nuevo",
        e - timedelta(days=3): "Jueves Santo",
        e - timedelta(days=2): "Viernes Santo",
        date(anio, 5, 1): "Día del Trabajo",
        date(anio, 6, 29): "San Pedro y San Pablo",
        date(anio, 7, 28): "Fiestas Patrias",
        date(anio, 7, 29): "Fiestas Patrias",
        date(anio, 8, 30): "Santa Rosa de Lima",
        date(anio, 10, 8): "Combate de Angamos",
        date(anio, 11, 1): "Todos los Santos",
        date(anio, 12, 8): "Inmaculada Concepción",
        date(anio, 12, 25): "Navidad",
    }


def _parsear_fecha_feriado(valor):
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y%m%d"):
        try:
            return datetime.strptime(str(valor).strip(), fmt).date()
        except Exception:
            continue
    return None


def _descargar_feriados_internet(anio):
    """Intenta descargar los feriados de Perú desde internet (sin API key)."""
    resultados = {}

    # 1) Nager.Date API (v3 con respaldo v2) — datos públicos sin clave
    for url in (
        f"https://date.nager.at/api/v3/PublicHolidays/{anio}/PE",
        f"https://date.nager.at/api/v2/PublicHolidays/{anio}/PE",
    ):
        try:
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (ControlFlota)"})
            with urllib.request.urlopen(req, context=ctx, timeout=12) as resp:
                datos = json.loads(resp.read().decode("utf-8", "replace"))
            for h in datos:
                dt = _parsear_fecha_feriado(h.get("date") or h.get("fecha"))
                if dt and dt.year == anio:
                    resultados[dt] = str(h.get("localName") or h.get("name") or "Feriado")
            if resultados:
                return resultados
        except Exception:
            continue

    # 2) Google Calendar ICS de feriados de Perú (respaldo)
    try:
        url_ics = ("https://calendar.google.com/calendar/ical/"
                   "en.peruvian%23holiday%40group.v.calendar.google.com/public/basic.ics")
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        req = urllib.request.Request(url_ics, headers={"User-Agent": "Mozilla/5.0 (ControlFlota)"})
        with urllib.request.urlopen(req, context=ctx, timeout=12) as resp:
            texto = resp.read().decode("utf-8", "replace")

        lineas = texto.splitlines()
        bloque = None
        for linea in lineas:
            if linea.startswith("BEGIN:VEVENT"):
                bloque = []
            elif linea.startswith("END:VEVENT"):
                if bloque:
                    _procesar_vevent(bloque, resultados, anio)
                bloque = None
            elif bloque is not None:
                bloque.append(linea)
    except Exception:
        pass

    return resultados


def _procesar_vevent(bloque, resultados, anio):
    campos = {}
    prev = ""
    for lin in bloque:
        if lin.startswith(" ") and prev:
            campos[prev] += lin.strip()
            continue
        if ":" in lin:
            clave, _, valor = lin.partition(":")
            campos[clave] = valor
            prev = clave
    fecha_raw = (campos.get("DTSTART") or "").replace("DTSTART", "").strip()
    fecha_raw = fecha_raw.replace("VALUE=DATE:", "").replace(":", "").strip()
    dt = _parsear_fecha_feriado(fecha_raw)
    if dt and dt.year == anio:
        resultados[dt] = str(campos.get("SUMMARY") or "Feriado")


def feriados_locales(anio):
    """Feriados locales (caché local si existe; si no, lista fija de Perú)."""
    cache = _leer_cache_feriados().get(str(anio))
    if cache:
        try:
            return {datetime.strptime(k, "%Y-%m-%d").date(): v for k, v in cache.items()}
        except Exception:
            pass
    return feriados_fijos_peru(anio)


# =========================================================
# 🚀 CLASIFICACIÓN DE DÍAS DE LA QUINCENA
# =========================================================
def clasificar_dia(fecha, feriados):
    if fecha in feriados:
        return "feriado"
    if fecha.weekday() == 6:  # Domingo
        return "domingo"
    return "normal"  # Lunes a sábado


def dias_de_quincena(anio, mes, quincena, feriados):
    """Devuelve la lista de días [fecha, categoria, nombre_feriado] de la quincena."""
    dias = []
    ultimo = calendar.monthrange(anio, mes)[1]
    inicio, fin = (1, 15) if quincena == 1 else (16, ultimo)
    for d in range(inicio, fin + 1):
        f = date(anio, mes, d)
        cat = clasificar_dia(f, feriados)
        nombre = feriados.get(f, "") if cat == "feriado" else ""
        dias.append({"fecha": f, "categoria": cat, "feriado_nombre": nombre})
    return dias


def conteos_quincena(dias):
    c = {"normal": 0, "domingo": 0, "feriado": 0}
    for d in dias:
        c[d["categoria"]] = c.get(d["categoria"], 0) + 1
    return c


# =========================================================
# 🚀 ASIENTOS DE DEDUCCIÓN (una línea con fecha y motivo)
# =========================================================
def nuevo_asiento_deduccion(fecha=None, categoria="normal", horas=0.0, minutos=0.0,
                            cantidad=0.0, motivo="", desde=None, hasta=None, precio=0.0):
    """Crea un asiento de deducción: una línea con fecha, concepto y valor.

    - Plan Por Hora: 'horas' + 'minutos' de servicio no prestado.
    - Plan Por Punto o Viaje: 'cantidad' de viajes/puntos no realizados.
    """
    cat = categoria if categoria in dict(CATEGORIAS) else "normal"
    return {
        "fecha": fecha if isinstance(fecha, date) else None,
        "categoria": cat,
        "horas": float(horas or 0.0),
        "minutos": float(minutos or 0.0),
        "cantidad": float(cantidad or 0.0),
        "motivo": str(motivo or ""),
        "desde": desde,
        "hasta": hasta,
        "precio": float(precio or 0.0),
    }


def horas_de_asiento(asiento):
    """Horas totales de un asiento de deducción (horas + minutos ÷ 60)."""
    try:
        return float(asiento.get("horas") or 0.0) + float(asiento.get("minutos") or 0.0) / 60.0
    except Exception:
        return 0.0


def ded_agrupada_desde_asientos(asientos):
    """Agrupa los asientos por categoría → {categoría: [horas, minutos]}.

    Es el formato que usa el motor de cálculo (y las columnas ded_*_h de la
    base de datos), así los asientos con fecha siguen alimentando el mismo
    cálculo de siempre.
    """
    res = {clave: [0.0, 0.0] for clave, _ in CATEGORIAS}
    for a in asientos or []:
        cat = a.get("categoria") if a.get("categoria") in res else "normal"
        res[cat][0] += float(a.get("horas") or 0.0)
        res[cat][1] += float(a.get("minutos") or 0.0)
    # Los minutos que completan una hora pasan al contador de horas
    for cat, (h, m) in list(res.items()):
        if m >= 60:
            res[cat] = [h + float(int(m // 60)), m % 60]
    return res


def etiqueta_fecha_hora(fecha):
    """'dd/mm/aaaa' del asiento ('—' si el asiento no tiene fecha)."""
    try:
        return fecha.strftime("%d/%m/%Y")
    except Exception:
        return "—"


# =========================================================
# 🚀 CLASE PRINCIPAL: CÁLCULO DE COBRANZA
# =========================================================
class CalculoCobranzaApp:
    def __init__(self, parent):
        self.parent = parent
        self.usuario_activo = "Desconocido"

        self.cliente_id = None
        self.cliente_nombre = ""
        self.cliente_ruc = ""
        self.cliente_comercial = ""
        self.cliente_contacto = ""
        self.cliente_telefono = ""
        self.cliente_direccion = ""
        self.plan_cobro = "Por Hora"

        self.dias = []          # lista de dicts de la quincena cargada
        self.feriados = {}      # {date: nombre}
        self.unidades = []      # unidades asignadas del cliente
        self.viajes_registrados = []  # viajes agregados con el registro rápido
        self.deducciones_registradas = []  # asientos de deducción (fecha + motivo) del plan por viajes
        self.registro_editando = None   # id de cobranza_quincenas en edición
        self._buscando_feriados = False

        # 🚀 ESQUEMA EN SEGUNDO PLANO (DAEMON)
        threading.Thread(target=self._inicializar_db, daemon=True).start()

        self.crear_interfaz()
        self.cargar_clientes()

    # ---------- BASE DE DATOS ----------
    def _inicializar_db(self):
        conn = conectar_db(silencioso=True)
        if not conn:
            return
        try:
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS cobranza_quincenas (
                    id SERIAL PRIMARY KEY,
                    id_cliente INTEGER NOT NULL,
                    cliente_nombre VARCHAR(255),
                    cliente_ruc VARCHAR(11),
                    anio INTEGER NOT NULL,
                    mes INTEGER NOT NULL,
                    quincena INTEGER NOT NULL,
                    plan_cobro VARCHAR(30) DEFAULT 'Por Hora',
                    precio_normal NUMERIC(12,2) DEFAULT 0,
                    precio_domingo NUMERIC(12,2) DEFAULT 0,
                    precio_feriado NUMERIC(12,2) DEFAULT 0,
                    cant_lunvie INTEGER DEFAULT 0,
                    cant_sabado INTEGER DEFAULT 0,
                    cant_domingo INTEGER DEFAULT 0,
                    cant_feriado INTEGER DEFAULT 0,
                    ded_normal_h NUMERIC(12,3) DEFAULT 0,
                    ded_sabado_h NUMERIC(12,3) DEFAULT 0,
                    ded_domingo_h NUMERIC(12,3) DEFAULT 0,
                    ded_feriado_h NUMERIC(12,3) DEFAULT 0,
                    monto_lunvie NUMERIC(12,2) DEFAULT 0,
                    monto_sabado NUMERIC(12,2) DEFAULT 0,
                    monto_domingo NUMERIC(12,2) DEFAULT 0,
                    monto_feriado NUMERIC(12,2) DEFAULT 0,
                    monto_base NUMERIC(12,2) DEFAULT 0,
                    monto_deducciones NUMERIC(12,2) DEFAULT 0,
                    total NUMERIC(12,2) DEFAULT 0,
                    notas TEXT,
                    pdf_ruta TEXT,
                    fecha_registro VARCHAR(30),
                    facturado BOOLEAN DEFAULT FALSE,
                    factura_referencia VARCHAR(100) DEFAULT ''
                )
            ''')
            cursor.execute("ALTER TABLE cobranza_quincenas ADD COLUMN IF NOT EXISTS facturado BOOLEAN DEFAULT FALSE")
            cursor.execute("ALTER TABLE cobranza_quincenas ADD COLUMN IF NOT EXISTS factura_referencia VARCHAR(100) DEFAULT ''")
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS cobranza_detalle_dias (
                    id SERIAL PRIMARY KEY,
                    id_cobranza INTEGER NOT NULL,
                    fecha VARCHAR(10),
                    categoria VARCHAR(20),
                    feriado_nombre VARCHAR(120) DEFAULT ''
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS clientes_unidades (
                    id SERIAL PRIMARY KEY,
                    id_cliente INTEGER NOT NULL,
                    id_vehiculo INTEGER,
                    unidad VARCHAR(120),
                    precio_normal NUMERIC(12,2) DEFAULT 0,
                    precio_domingo NUMERIC(12,2) DEFAULT 0,
                    precio_feriado NUMERIC(12,2) DEFAULT 0,
                    horas_dia NUMERIC(10,2) DEFAULT 8.00,
                    activo BOOLEAN DEFAULT TRUE
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS cobranza_quincena_unidades (
                    id SERIAL PRIMARY KEY,
                    id_cobranza INTEGER NOT NULL,
                    unidad VARCHAR(120),
                    precio_normal NUMERIC(12,2) DEFAULT 0,
                    precio_domingo NUMERIC(12,2) DEFAULT 0,
                    precio_feriado NUMERIC(12,2) DEFAULT 0,
                    horas_dia NUMERIC(10,2) DEFAULT 0,
                    cant_normal INTEGER DEFAULT 0,
                    cant_domingo INTEGER DEFAULT 0,
                    cant_feriado INTEGER DEFAULT 0,
                    ded_normal_h NUMERIC(12,3) DEFAULT 0,
                    ded_domingo_h NUMERIC(12,3) DEFAULT 0,
                    ded_feriado_h NUMERIC(12,3) DEFAULT 0,
                    monto_normal NUMERIC(12,2) DEFAULT 0,
                    monto_domingo NUMERIC(12,2) DEFAULT 0,
                    monto_feriado NUMERIC(12,2) DEFAULT 0,
                    monto_deducciones NUMERIC(12,2) DEFAULT 0,
                    subtotal NUMERIC(12,2) DEFAULT 0
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS precios_viaje_distancia (
                    id SERIAL PRIMARY KEY,
                    id_cliente INTEGER NOT NULL,
                    distancia_desde NUMERIC(10,2) NOT NULL,
                    distancia_hasta NUMERIC(10,2) NOT NULL,
                    precio_normal NUMERIC(12,2) DEFAULT 0,
                    precio_domingo NUMERIC(12,2) DEFAULT 0,
                    precio_feriado NUMERIC(12,2) DEFAULT 0
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS cobranza_quincena_viajes (
                    id SERIAL PRIMARY KEY,
                    id_cobranza INTEGER NOT NULL,
                    distancia_desde NUMERIC(10,2) DEFAULT 0,
                    distancia_hasta NUMERIC(10,2) DEFAULT 0,
                    precio_normal NUMERIC(12,2) DEFAULT 0,
                    precio_domingo NUMERIC(12,2) DEFAULT 0,
                    precio_feriado NUMERIC(12,2) DEFAULT 0,
                    viajes_normal INTEGER DEFAULT 0,
                    viajes_domingo INTEGER DEFAULT 0,
                    viajes_feriado INTEGER DEFAULT 0,
                    ded_normal INTEGER DEFAULT 0,
                    ded_domingo INTEGER DEFAULT 0,
                    ded_feriado INTEGER DEFAULT 0,
                    monto_base NUMERIC(12,2) DEFAULT 0,
                    monto_deducciones NUMERIC(12,2) DEFAULT 0,
                    subtotal NUMERIC(12,2) DEFAULT 0
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS cobranza_viajes_detalle (
                    id SERIAL PRIMARY KEY,
                    id_cobranza INTEGER NOT NULL,
                    fecha VARCHAR(10),
                    vehiculo VARCHAR(120),
                    categoria VARCHAR(20),
                    distancia_desde NUMERIC(10,2) DEFAULT 0,
                    distancia_hasta NUMERIC(10,2) DEFAULT 0,
                    precio NUMERIC(12,2) DEFAULT 0,
                    monto NUMERIC(12,2) DEFAULT 0
                )
            ''')
            # Asientos con fecha: deducciones (tipo DEDUCCION) y horas extras
            # (tipo EXTRA), uno por línea (unidad o rango de distancia)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS cobranza_deducciones_detalle (
                    id SERIAL PRIMARY KEY,
                    id_cobranza INTEGER NOT NULL,
                    plan VARCHAR(30) DEFAULT 'Por Hora',
                    unidad VARCHAR(150) DEFAULT '',
                    fecha VARCHAR(10),
                    categoria VARCHAR(20) DEFAULT 'normal',
                    horas NUMERIC(12,3) DEFAULT 0,
                    minutos NUMERIC(12,3) DEFAULT 0,
                    cantidad NUMERIC(12,2) DEFAULT 0,
                    distancia_desde NUMERIC(10,2) DEFAULT 0,
                    distancia_hasta NUMERIC(10,2) DEFAULT 0,
                    precio NUMERIC(12,2) DEFAULT 0,
                    motivo VARCHAR(200) DEFAULT '',
                    monto NUMERIC(12,2) DEFAULT 0,
                    tipo VARCHAR(20) DEFAULT 'DEDUCCION'
                )
            ''')
            cursor.execute("ALTER TABLE cobranza_deducciones_detalle "
                           "ADD COLUMN IF NOT EXISTS tipo VARCHAR(20) DEFAULT 'DEDUCCION'")
            cursor.execute("ALTER TABLE clientes ADD COLUMN IF NOT EXISTS plan_cobro VARCHAR(30) DEFAULT 'Por Hora'")
            # Evita duplicar el mismo vehículo en las unidades de un cliente (sin romper la transacción)
            cursor.execute("""
                DO $$
                BEGIN
                    IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'uq_clientes_unidades_vehiculo') THEN
                        ALTER TABLE clientes_unidades ADD CONSTRAINT uq_clientes_unidades_vehiculo UNIQUE (id_cliente, id_vehiculo);
                    END IF;
                END $$;
            """)
            conn.commit()
            cursor.close()
        except Exception as e:
            print(f"[Schema Warning Cobranza] {e}")
        finally:
            liberar_conexion(conn)

    # ---------- INTERFAZ ----------
    def crear_interfaz(self):
        familia_fuente = "Helvetica" if sys.platform == "darwin" else "Arial"

        self.scroll = ctk.CTkScrollableFrame(self.parent, fg_color="transparent")
        self.scroll.pack(fill="both", expand=True, padx=12, pady=12)

        # Mensaje flash (aparece abajo al guardar)
        self.frame_flash = ctk.CTkFrame(self.scroll, fg_color="#27ae60", corner_radius=8)
        self.lbl_flash = ctk.CTkLabel(self.frame_flash, text="✅ Quincena guardada correctamente",
                                      font=(familia_fuente, 14, "bold"), text_color="white")
        self.lbl_flash.pack(padx=25, pady=10)
        self.frame_flash.pack_forget()

        # ---- Encabezado ----
        f_header = ctk.CTkFrame(self.scroll, fg_color="transparent")
        f_header.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(f_header, text="💰 CÁLCULO DE COBRANZA QUINCENAL",
                     font=(familia_fuente, 18, "bold"), text_color="#1f538d").pack(side="left")
        btn_registros = ctk.CTkButton(f_header, text="📋 Ver Registros de Quincenas", width=210,
                                      fg_color="#34495e", hover_color="#2c3e50",
                                      font=(familia_fuente, 12, "bold"), command=self.abrir_registros)
        btn_registros.pack(side="right")

        # ---- 1. Cliente ----
        f_cli = ctk.CTkFrame(self.scroll, corner_radius=10, border_width=1, border_color="#e0e0e0")
        f_cli.pack(fill="x", padx=5, pady=5, ipady=6)
        f_cli.columnconfigure(1, weight=1)

        ctk.CTkLabel(f_cli, text="👤 Cliente:", font=(familia_fuente, 12, "bold")).grid(
            row=0, column=0, sticky="w", padx=(15, 5), pady=8)
        self.combo_cliente = ctk.CTkComboBox(f_cli, values=["— Seleccione un cliente —"],
                                             font=(familia_fuente, 12), state="readonly",
                                             command=self.on_cliente_seleccionado)
        self.combo_cliente.grid(row=0, column=1, sticky="ew", padx=5, pady=8)
        btn_recargar = ctk.CTkButton(f_cli, text="🔄 Recargar", width=95,
                                     font=(familia_fuente, 11, "bold"), fg_color="#7f8c8d",
                                     hover_color="#606b6b", command=self.cargar_clientes)
        btn_recargar.grid(row=0, column=2, sticky="e", padx=(5, 5), pady=8)
        self.btn_unidades = ctk.CTkButton(f_cli, text="🚗 Unidades del Cliente", width=190,
                                        font=(familia_fuente, 11, "bold"), fg_color="#e67e22",
                                        hover_color="#d35400", command=self.abrir_ventana_unidades)
        self.btn_unidades.grid(row=0, column=3, sticky="e", padx=(5, 5), pady=8)
        self.btn_precios_dist = ctk.CTkButton(f_cli, text="📏 Precios por Distancia", width=190,
                                              font=(familia_fuente, 11, "bold"), fg_color="#8e44ad",
                                              hover_color="#6c3483", command=self.abrir_ventana_precios_distancia)
        self.btn_precios_dist.grid(row=0, column=4, sticky="e", padx=(5, 15), pady=8)

        ctk.CTkLabel(f_cli, text="🧾 Plan de Cobro:", font=(familia_fuente, 12, "bold")).grid(
            row=1, column=0, sticky="w", padx=(15, 5), pady=8)
        self.lbl_plan = ctk.CTkLabel(f_cli, text="Por Hora", font=(familia_fuente, 13, "bold"),
                                     text_color="#d35400")
        self.lbl_plan.grid(row=1, column=1, sticky="w", padx=5, pady=8)
        self.lbl_cliente_info = ctk.CTkLabel(f_cli, text="", font=(familia_fuente, 10),
                                             text_color="gray")
        self.lbl_cliente_info.grid(row=1, column=2, columnspan=3, sticky="e", padx=(5, 15), pady=8)

        # ---- 2. Periodo ----
        f_per = ctk.CTkFrame(self.scroll, corner_radius=10, border_width=1, border_color="#e0e0e0")
        f_per.pack(fill="x", padx=5, pady=5, ipady=6)
        f_per.columnconfigure(1, weight=1)

        ctk.CTkLabel(f_per, text="📅 Mes:", font=(familia_fuente, 12, "bold")).grid(
            row=0, column=0, sticky="w", padx=(15, 5), pady=8)
        self.combo_mes = ctk.CTkOptionMenu(f_per, values=[f"{i:02d} - {NOMBRES_MESES[i-1]}" for i in range(1, 13)],
                                           font=(familia_fuente, 12))
        hoy = datetime.now()
        self.combo_mes.set(f"{hoy.month:02d} - {NOMBRES_MESES[hoy.month-1]}")
        self.combo_mes.grid(row=0, column=1, sticky="ew", padx=5, pady=8)

        ctk.CTkLabel(f_per, text="Año:", font=(familia_fuente, 12, "bold")).grid(
            row=0, column=2, sticky="w", padx=(15, 5), pady=8)
        anios = [str(a) for a in range(hoy.year - 1, hoy.year + 3)]
        self.combo_anio = ctk.CTkOptionMenu(f_per, values=anios, font=(familia_fuente, 12))
        self.combo_anio.set(str(hoy.year))
        self.combo_anio.grid(row=0, column=3, sticky="ew", padx=5, pady=8)

        ctk.CTkLabel(f_per, text="Quincena:", font=(familia_fuente, 12, "bold")).grid(
            row=0, column=4, sticky="w", padx=(15, 5), pady=8)
        self.combo_quincena = ctk.CTkOptionMenu(f_per, values=["Primera (1 al 15)", "Segunda (16 al fin)"],
                                                font=(familia_fuente, 12))
        self.combo_quincena.grid(row=0, column=5, sticky="ew", padx=5, pady=8)

        btn_buscar = ctk.CTkButton(f_per, text="🔎 Buscar días de la quincena (internet)", width=260,
                                   font=(familia_fuente, 12, "bold"), fg_color="#1f538d",
                                   hover_color="#163b65", command=self.buscar_dias_quincena)
        btn_buscar.grid(row=0, column=6, sticky="e", padx=(10, 15), pady=8)

        # ---- 3. Días de la quincena ----
        f_dias = ctk.CTkFrame(self.scroll, corner_radius=10, border_width=1, border_color="#e0e0e0")
        f_dias.pack(fill="x", padx=5, pady=5, ipady=6)

        f_dias_tool = ctk.CTkFrame(f_dias, fg_color="transparent")
        f_dias_tool.pack(fill="x", padx=10, pady=(8, 2))
        ctk.CTkLabel(f_dias_tool, text="📆 Días de la quincena (seleccione una fila para marcar/desmarcar feriado):",
                     font=(familia_fuente, 12, "bold")).pack(side="left")
        btn_fer = ctk.CTkButton(f_dias_tool, text="🎌 Marcar/Desmarcar Feriado", width=180,
                                font=(familia_fuente, 11, "bold"), fg_color="#e67e22",
                                hover_color="#d35400", command=self.toggle_feriado)
        btn_fer.pack(side="right", padx=(5, 0))
        btn_rest = ctk.CTkButton(f_dias_tool, text="♻️ Restablecer (internet)", width=170,
                                 font=(familia_fuente, 11, "bold"), fg_color="#8e44ad",
                                 hover_color="#6c3483", command=self.restablecer_feriados)
        btn_rest.pack(side="right", padx=(5, 5))

        columnas = ("fecha", "dia", "categoria", "feriado")
        style = ttk.Style()
        if sys.platform == "darwin":
            style.theme_use("clam")
        bg_blanco, fg_negro, bg_seleccion = "#ffffff", "#000000", "#1f538d"
        style.configure("Treeview", background=bg_blanco, foreground=fg_negro,
                        fieldbackground=bg_blanco, rowheight=26, font=(familia_fuente, 10),
                        bordercolor="#e0e0e0", borderwidth=1)
        style.map("Treeview", background=[("selected", bg_seleccion)],
                  foreground=[("selected", "#ffffff")])
        style.configure("Treeview.Heading", background="#f0f0f0", foreground=fg_negro,
                        font=(familia_fuente, 10, "bold"), bordercolor="#e0e0e0", borderwidth=1, relief="flat")

        self.tabla_dias = ttk.Treeview(f_dias, columns=columnas, show="headings",
                                       selectmode="browse", style="Treeview")
        self.tabla_dias.heading("fecha", text="Fecha", anchor="center")
        self.tabla_dias.heading("dia", text="Día", anchor="center")
        self.tabla_dias.heading("categoria", text="Categoría", anchor="center")
        self.tabla_dias.heading("feriado", text="Feriado", anchor="center")
        self.tabla_dias.column("fecha", width=90, anchor="center")
        self.tabla_dias.column("dia", width=120, anchor="center")
        self.tabla_dias.column("categoria", width=120, anchor="center")
        self.tabla_dias.column("feriado", width=220, anchor="w")
        self.tabla_dias.pack(fill="x", padx=10, pady=(2, 5))

        self.lbl_conteo = ctk.CTkLabel(f_dias, text="", font=(familia_fuente, 11, "bold"),
                                       text_color="#1f538d")
        self.lbl_conteo.pack(anchor="w", padx=15, pady=(0, 6))
        self.lbl_origen_feriados = ctk.CTkLabel(f_dias, text="", font=(familia_fuente, 10),
                                                text_color="gray")
        self.lbl_origen_feriados.pack(anchor="w", padx=15, pady=(0, 6))

        # ---- 4. Unidades asignadas ----

        # ---- 4b. VIAJES DE LA QUINCENA (plan Por Punto o Viaje) ----
        self.f_via = ctk.CTkFrame(self.scroll, corner_radius=10, border_width=1, border_color="#e0e0e0")
        self.f_via.pack(fill="x", padx=5, pady=5, ipady=6)
        self.f_via.pack_forget()  # se muestra solo en plan Por Punto o Viaje

        f_via_tool = ctk.CTkFrame(self.f_via, fg_color="transparent")
        f_via_tool.pack(fill="x", padx=10, pady=(8, 2))
        ctk.CTkLabel(f_via_tool, text="🛣️ VIAJES DE LA QUINCENA (precio según distancia)",
                     font=(familia_fuente, 13, "bold"), text_color="#1f538d").pack(side="left")
        btn_ref_via = ctk.CTkButton(f_via_tool, text="↻ Refrescar", width=110,
                                    font=(familia_fuente, 11, "bold"), fg_color="#7f8c8d",
                                    hover_color="#606b6b", command=self.cargar_rangos_viaje)
        btn_ref_via.pack(side="right", padx=(5, 5))
        self.lbl_sin_viajes = ctk.CTkLabel(self.f_via, text="", font=(familia_fuente, 11, "bold"),
                                           text_color="#c0392b")
        self.lbl_sin_viajes.pack(anchor="w", padx=15, pady=(2, 2))
        self.frame_viajes_grid = ctk.CTkFrame(self.f_via, fg_color="transparent")
        self.frame_viajes_grid.pack(fill="x", padx=10, pady=(2, 4))

        self.lbl_total_viajes = ctk.CTkLabel(self.f_via, text="TOTAL VIAJES: S/ 0,00",
                                             font=(familia_fuente, 14, "bold"), text_color="#27ae60")
        self.lbl_total_viajes.pack(anchor="w", padx=15, pady=(0, 6))

        # ---- Vehículo que realizó el viaje ----
        f_veh = ctk.CTkFrame(self.f_via, corner_radius=8, border_width=1, border_color="#e0e0e0")
        f_veh.pack(fill="x", padx=10, pady=(0, 8))
        ctk.CTkLabel(f_veh, text="🚗 VEHÍCULO QUE REALIZÓ EL VIAJE",
                     font=(familia_fuente, 11, "bold"), text_color="#1f538d").pack(side="left", padx=10, pady=8)
        self.combo_vehiculo_viaje = ctk.CTkOptionMenu(f_veh, values=["— Seleccione vehículo —"], width=300,
                                                      font=(familia_fuente, 11))
        self.combo_vehiculo_viaje.pack(side="left", padx=5, pady=8)

        # ---- Registro rápido de viajes (se suma automáticamente a la tabla de arriba) ----
        f_reg = ctk.CTkFrame(self.f_via, corner_radius=8, border_width=1, border_color="#e0e0e0")
        f_reg.pack(fill="x", padx=10, pady=(0, 8))
        ctk.CTkLabel(f_reg, text="➕ REGISTRO RÁPIDO DE VIAJES (se suma y calcula automáticamente)",
                     font=(familia_fuente, 11, "bold"), text_color="#1f538d").grid(
            row=0, column=0, columnspan=6, sticky="w", padx=10, pady=(8, 4))

        ctk.CTkLabel(f_reg, text="Fecha:", font=(familia_fuente, 11, "bold")).grid(
            row=1, column=0, sticky="w", padx=(10, 5), pady=6)
        self.ent_fecha_viaje = ctk.CTkEntry(f_reg, width=110, placeholder_text="dd/mm/aaaa")
        self.ent_fecha_viaje.grid(row=1, column=1, sticky="w", padx=5, pady=6)
        btn_cal = ctk.CTkButton(f_reg, text="📅", width=42, font=(familia_fuente, 12),
                                command=self.abrir_calendario_viaje)
        btn_cal.grid(row=1, column=2, sticky="w", padx=2, pady=6)
        self.ent_fecha_viaje.bind("<KeyRelease>", lambda e: self._al_seleccionar_fecha_viaje())

        ctk.CTkLabel(f_reg, text="Tipo de viaje:", font=(familia_fuente, 11, "bold")).grid(
            row=1, column=3, sticky="w", padx=(15, 5), pady=6)
        self.combo_tipo_viaje = ctk.CTkOptionMenu(f_reg, values=["— Seleccione distancia —"], width=160,
                                                  font=(familia_fuente, 11))
        self.combo_tipo_viaje.grid(row=1, column=4, sticky="w", padx=5, pady=6)

        btn_add_viaje = ctk.CTkButton(f_reg, text="➕ Agregar viaje", width=140, font=(familia_fuente, 11, "bold"),
                                      fg_color="#27ae60", hover_color="#1e8449", command=self.agregar_viaje_registrado)
        btn_add_viaje.grid(row=1, column=5, sticky="e", padx=10, pady=6)

        self.lbl_dia_detectado = ctk.CTkLabel(f_reg, text="", font=(familia_fuente, 10, "bold"),
                                              text_color="#1f538d")
        self.lbl_dia_detectado.grid(row=2, column=0, columnspan=3, sticky="w", padx=10, pady=(0, 4))

        f_tab_reg = ctk.CTkFrame(f_reg, fg_color="transparent")
        f_tab_reg.grid(row=3, column=0, columnspan=6, sticky="ew", padx=10, pady=(4, 4))
        f_tab_reg.columnconfigure(0, weight=1)

        self.tabla_viajes_reg = ttk.Treeview(f_tab_reg, columns=("fecha", "dia", "tipo", "vehiculo"), show="headings",
                                             selectmode="browse", style="Treeview", height=3)
        self.tabla_viajes_reg.heading("fecha", text="Fecha", anchor="center")
        self.tabla_viajes_reg.heading("dia", text="Día", anchor="center")
        self.tabla_viajes_reg.heading("tipo", text="Tipo (distancia)", anchor="center")
        self.tabla_viajes_reg.heading("vehiculo", text="Vehículo", anchor="center")
        self.tabla_viajes_reg.column("fecha", width=95, anchor="center")
        self.tabla_viajes_reg.column("dia", width=130, anchor="center")
        self.tabla_viajes_reg.column("tipo", width=160, anchor="center")
        self.tabla_viajes_reg.column("vehiculo", width=220, anchor="w")
        self.tabla_viajes_reg.grid(row=0, column=0, sticky="nsew")
        scr_reg = ttk.Scrollbar(f_tab_reg, orient="vertical", command=self.tabla_viajes_reg.yview)
        self.tabla_viajes_reg.configure(yscrollcommand=scr_reg.set)
        scr_reg.grid(row=0, column=1, sticky="ns")
        # La rueda del mouse sobre esta lista solo desplaza la lista (no la ventana)
        self.tabla_viajes_reg.bind("<MouseWheel>", self._scroll_solo_tabla_viajes)
        self.tabla_viajes_reg.bind("<Button-4>", self._scroll_solo_tabla_viajes)
        self.tabla_viajes_reg.bind("<Button-5>", self._scroll_solo_tabla_viajes)

        btn_quitar_viaje = ctk.CTkButton(f_reg, text="➖ Quitar seleccionado", width=190,
                                         font=(familia_fuente, 11, "bold"), fg_color="#e74c3c",
                                         hover_color="#c0392b", command=self.quitar_viaje_registrado)
        btn_quitar_viaje.grid(row=4, column=0, columnspan=6, sticky="w", padx=10, pady=(0, 8))

        # ---- Registro de DEDUCCIONES con fecha (viajes / puntos no realizados) ----
        f_ded_reg = ctk.CTkFrame(self.f_via, corner_radius=8, border_width=1, border_color="#e0e0e0")
        f_ded_reg.pack(fill="x", padx=10, pady=(0, 8))
        f_ded_reg.columnconfigure(1, weight=1)
        ctk.CTkLabel(f_ded_reg, text="➖ DEDUCCIONES CON FECHA (viajes / puntos no realizados)",
                     font=(familia_fuente, 11, "bold"), text_color="#c0392b").grid(
            row=0, column=0, columnspan=8, sticky="w", padx=10, pady=(8, 4))

        ctk.CTkLabel(f_ded_reg, text="Fecha:", font=(familia_fuente, 11, "bold")).grid(
            row=1, column=0, sticky="w", padx=(10, 5), pady=6)
        self.ent_fecha_ded = ctk.CTkEntry(f_ded_reg, width=110, placeholder_text="dd/mm/aaaa")
        self.ent_fecha_ded.grid(row=1, column=1, sticky="w", padx=5, pady=6)
        btn_cal_ded = ctk.CTkButton(f_ded_reg, text="📅", width=42, font=(familia_fuente, 12),
                                    command=self.abrir_calendario_ded)
        btn_cal_ded.grid(row=1, column=2, sticky="w", padx=2, pady=6)
        self.ent_fecha_ded.bind("<KeyRelease>", lambda e: self._al_seleccionar_fecha_ded())

        ctk.CTkLabel(f_ded_reg, text="Tipo de viaje:", font=(familia_fuente, 11, "bold")).grid(
            row=1, column=3, sticky="w", padx=(15, 5), pady=6)
        self.combo_tipo_ded = ctk.CTkOptionMenu(f_ded_reg, values=["— Seleccione distancia —"], width=160,
                                                font=(familia_fuente, 11))
        self.combo_tipo_ded.grid(row=1, column=4, sticky="w", padx=5, pady=6)

        ctk.CTkLabel(f_ded_reg, text="Cantidad:", font=(familia_fuente, 11, "bold")).grid(
            row=1, column=5, sticky="w", padx=(15, 5), pady=6)
        self.ent_cant_ded = ctk.CTkEntry(f_ded_reg, width=70, justify="center", placeholder_text="1")
        self.ent_cant_ded.grid(row=1, column=6, sticky="w", padx=5, pady=6)

        ctk.CTkLabel(f_ded_reg, text="Motivo:", font=(familia_fuente, 11, "bold")).grid(
            row=2, column=0, sticky="w", padx=(10, 5), pady=6)
        self.ent_motivo_ded = ctk.CTkEntry(f_ded_reg, placeholder_text="Ej.: unidad en mantenimiento / servicio no prestado")
        self.ent_motivo_ded.grid(row=2, column=1, columnspan=5, sticky="ew", padx=5, pady=6)
        btn_add_ded = ctk.CTkButton(f_ded_reg, text="➕ Agregar deducción", width=180,
                                    font=(familia_fuente, 11, "bold"), fg_color="#e67e22",
                                    hover_color="#d35400", command=self.agregar_deduccion_registrada)
        btn_add_ded.grid(row=2, column=6, sticky="e", padx=10, pady=6)

        self.lbl_dia_detectado_ded = ctk.CTkLabel(f_ded_reg, text="", font=(familia_fuente, 10, "bold"),
                                                  text_color="#1f538d")
        self.lbl_dia_detectado_ded.grid(row=3, column=0, columnspan=4, sticky="w", padx=10, pady=(0, 4))

        f_tab_ded = ctk.CTkFrame(f_ded_reg, fg_color="transparent")
        f_tab_ded.grid(row=4, column=0, columnspan=8, sticky="ew", padx=10, pady=(4, 4))
        f_tab_ded.columnconfigure(0, weight=1)

        self.tabla_ded_reg = ttk.Treeview(f_tab_ded, columns=("fecha", "dia", "tipo", "cant", "motivo"),
                                          show="headings", selectmode="browse", style="Treeview", height=3)
        self.tabla_ded_reg.heading("fecha", text="Fecha", anchor="center")
        self.tabla_ded_reg.heading("dia", text="Día", anchor="center")
        self.tabla_ded_reg.heading("tipo", text="Tipo (distancia)", anchor="center")
        self.tabla_ded_reg.heading("cant", text="Cant.", anchor="center")
        self.tabla_ded_reg.heading("motivo", text="Motivo", anchor="center")
        self.tabla_ded_reg.column("fecha", width=95, anchor="center")
        self.tabla_ded_reg.column("dia", width=120, anchor="center")
        self.tabla_ded_reg.column("tipo", width=150, anchor="center")
        self.tabla_ded_reg.column("cant", width=55, anchor="center")
        self.tabla_ded_reg.column("motivo", width=330, anchor="w")
        self.tabla_ded_reg.grid(row=0, column=0, sticky="nsew")
        scr_ded = ttk.Scrollbar(f_tab_ded, orient="vertical", command=self.tabla_ded_reg.yview)
        self.tabla_ded_reg.configure(yscrollcommand=scr_ded.set)
        scr_ded.grid(row=0, column=1, sticky="ns")
        self.tabla_ded_reg.bind("<MouseWheel>", self._scroll_solo_tabla_ded)
        self.tabla_ded_reg.bind("<Button-4>", self._scroll_solo_tabla_ded)
        self.tabla_ded_reg.bind("<Button-5>", self._scroll_solo_tabla_ded)

        btn_quitar_ded = ctk.CTkButton(f_ded_reg, text="➖ Quitar seleccionada", width=190,
                                       font=(familia_fuente, 11, "bold"), fg_color="#e74c3c",
                                       hover_color="#c0392b", command=self.quitar_deduccion_registrada)
        btn_quitar_ded.grid(row=5, column=0, columnspan=3, sticky="w", padx=10, pady=(0, 8))
        self.lbl_total_ded = ctk.CTkLabel(f_ded_reg, text="TOTAL DEDUCCIONES REGISTRADAS: S/ 0,00",
                                          font=(familia_fuente, 12, "bold"), text_color="#c0392b")
        self.lbl_total_ded.grid(row=5, column=4, columnspan=3, sticky="e", padx=10, pady=(0, 8))

        self.f_uni = ctk.CTkFrame(self.scroll, corner_radius=10, border_width=1, border_color="#e0e0e0")
        self.f_uni.pack(fill="x", padx=5, pady=5, ipady=6)

        f_uni_tool = ctk.CTkFrame(self.f_uni, fg_color="transparent")
        f_uni_tool.pack(fill="x", padx=10, pady=(8, 2))
        ctk.CTkLabel(f_uni_tool, text="🚗 UNIDADES ASIGNADAS (cada una con sus precios y horas/día)",
                     font=(familia_fuente, 13, "bold"), text_color="#1f538d").pack(side="left")
        btn_edit_uni = ctk.CTkButton(f_uni_tool, text="✏️ Editar Precios / Deducciones", width=230,
                                     font=(familia_fuente, 11, "bold"), fg_color="#34495e",
                                     hover_color="#2c3e50", command=self.editar_unidad_seleccionada)
        btn_edit_uni.pack(side="right", padx=(5, 0))
        btn_add_uni = ctk.CTkButton(f_uni_tool, text="➕ Agregar de la Flota", width=170,
                                    font=(familia_fuente, 11, "bold"), fg_color="#27ae60",
                                    hover_color="#1e8449", command=self.agregar_unidad_flota)
        btn_add_uni.pack(side="right", padx=(5, 5))

        columnas_u = ("unidad", "horas", "pnormal", "pdom", "pfer", "subtotal")
        self.tabla_unidades = ttk.Treeview(self.f_uni, columns=columnas_u, show="headings",
                                           selectmode="browse", style="Treeview")
        self.tabla_unidades.heading("unidad", text="Unidad", anchor="center")
        self.tabla_unidades.heading("horas", text="Horas/Día", anchor="center")
        self.tabla_unidades.heading("pnormal", text="P. Normal", anchor="center")
        self.tabla_unidades.heading("pdom", text="P. Domingo", anchor="center")
        self.tabla_unidades.heading("pfer", text="P. Feriado", anchor="center")
        self.tabla_unidades.heading("subtotal", text="Subtotal", anchor="center")
        self.tabla_unidades.column("unidad", width=280, anchor="w")
        self.tabla_unidades.column("horas", width=80, anchor="center")
        self.tabla_unidades.column("pnormal", width=90, anchor="center")
        self.tabla_unidades.column("pdom", width=90, anchor="center")
        self.tabla_unidades.column("pfer", width=90, anchor="center")
        self.tabla_unidades.column("subtotal", width=110, anchor="center")
        self.tabla_unidades.pack(fill="x", padx=10, pady=(2, 5))
        self.tabla_unidades.bind("<Double-1>", lambda e: self.editar_unidad_seleccionada())

        self.lbl_sin_unidades = ctk.CTkLabel(self.f_uni, text="", font=(familia_fuente, 11, "bold"),
                                             text_color="#c0392b")
        self.lbl_sin_unidades.pack(anchor="w", padx=15, pady=(0, 6))

        # ---- 5. Resumen ----
        f_res = ctk.CTkFrame(self.scroll, corner_radius=10, border_width=1, border_color="#e0e0e0")
        f_res.pack(fill="x", padx=5, pady=5, ipady=6)

        self.lbl_resumen = ctk.CTkLabel(f_res, text="", font=(familia_fuente, 12), justify="left")
        self.lbl_resumen.pack(anchor="w", padx=15, pady=8)

        self.lbl_total = ctk.CTkLabel(f_res, text="TOTAL A COBRAR: S/ 0.00",
                                      font=(familia_fuente, 17, "bold"), text_color="#27ae60")
        self.lbl_total.pack(anchor="w", padx=15, pady=(0, 10))

        # ---- 6. Notas y acciones ----
        f_act = ctk.CTkFrame(self.scroll, corner_radius=10, border_width=1, border_color="#e0e0e0")
        f_act.pack(fill="x", padx=5, pady=5, ipady=8)
        f_act.columnconfigure(1, weight=1)

        ctk.CTkLabel(f_act, text="📝 Notas:", font=(familia_fuente, 12, "bold")).grid(
            row=0, column=0, sticky="w", padx=(15, 5), pady=8)
        self.ent_notas = ctk.CTkEntry(f_act, placeholder_text="Comentarios adicionales del cobro...")
        self.ent_notas.grid(row=0, column=1, sticky="ew", padx=5, pady=8)

        f_btns = ctk.CTkFrame(f_act, fg_color="transparent")
        f_btns.grid(row=1, column=0, columnspan=2, sticky="ew", padx=15, pady=6)
        btn_guardar = ctk.CTkButton(f_btns, text="💾 Guardar Quincena", width=200, height=40,
                                    font=(familia_fuente, 13, "bold"), fg_color="#27ae60",
                                    hover_color="#1e8449", command=self.guardar_quincena)
        btn_guardar.pack(side="left", padx=5)
        btn_pdf = ctk.CTkButton(f_btns, text="📄 Generar PDF", width=180, height=40,
                                font=(familia_fuente, 13, "bold"), fg_color="#c0392b",
                                hover_color="#922b21", command=self.generar_pdf)
        btn_pdf.pack(side="left", padx=5)
        btn_limp = ctk.CTkButton(f_btns, text="🧹 Limpiar Formulario", width=180, height=40,
                                 font=(familia_fuente, 13, "bold"), fg_color="#7f8c8d",
                                 hover_color="#606b6b", command=self.limpiar_formulario)
        btn_limp.pack(side="left", padx=5)

        self.aplicar_plan(self.plan_cobro)

    # ---------- CLIENTES ----------
    def cargar_clientes(self):
        self.clientes_lista = []
        valores = ["— Seleccione un cliente —"]
        conn = conectar_db(silencioso=True)
        if not conn:
            messagebox.showwarning("Modo Lectura", "Sin conexión a la base de datos.\nNo se pueden cargar los clientes.")
            try:
                self.combo_cliente.configure(values=valores)
            except Exception:
                pass
            return
        try:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, ruc, nombre_empresa, COALESCE(razon_comercial, ''),
                       COALESCE(plan_cobro, 'Por Hora'), COALESCE(persona_contacto, ''),
                       COALESCE(telefono, ''), COALESCE(direccion_fiscal, '')
                FROM clientes ORDER BY nombre_empresa ASC
            ''')
            for row in cursor.fetchall():
                self.clientes_lista.append({
                    "id": row[0], "ruc": str(row[1]), "nombre": str(row[2]),
                    "comercial": str(row[3]), "plan": str(row[4]) if row[4] in PLANES else "Por Hora",
                    "contacto": str(row[5]), "telefono": str(row[6]), "direccion": str(row[7]),
                })
                valores.append(f"{row[1]} | {row[2]}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar los clientes:\n{e}")
        finally:
            liberar_conexion(conn)
        try:
            self.combo_cliente.configure(values=valores)
        except Exception:
            pass

    def on_cliente_seleccionado(self, seleccion):
        if not seleccion or seleccion.startswith("—"):
            self.cliente_id = None
            self.lbl_plan.configure(text="Por Hora")
            self.aplicar_plan("Por Hora")
            self.unidades = []
            self.pintar_unidades()
            self.rangos_viaje = []
            self.deducciones_registradas = []
            self.pintar_viajes()
            self.recalcular()
            return
        for c in self.clientes_lista:
            if f"{c['ruc']} | {c['nombre']}" == seleccion:
                self.cliente_id = c["id"]
                self.cliente_nombre = c["nombre"]
                self.cliente_ruc = c["ruc"]
                self.cliente_comercial = c["comercial"]
                self.cliente_contacto = c["contacto"]
                self.cliente_telefono = c["telefono"]
                self.cliente_direccion = c["direccion"]
                self.plan_cobro = c["plan"]
                self.lbl_plan.configure(text=self.plan_cobro)
                self.lbl_cliente_info.configure(text=f"{self.cliente_ruc} | {self.cliente_nombre}")
                self.aplicar_plan(self.plan_cobro)
                # Carga unidades y tabla de distancias del cliente según su plan
                self.cargar_rangos_viaje()
                self.cargar_unidades_cliente()
                return

    def aplicar_plan(self, plan):
        """Muestra la sección y botones según el plan de cobro:
        Por Hora → unidades; Por Punto o Viaje → tabla de distancias."""
        try:
            f_dias = self.tabla_dias.master
            if plan == "Por Punto o Viaje":
                try:
                    self.f_uni.pack_forget()
                    self.btn_unidades.grid_remove()
                except Exception:
                    pass
                try:
                    self.btn_precios_dist.grid()
                except Exception:
                    pass
                self.f_via.pack(fill="x", padx=5, pady=5, ipady=6, after=f_dias)
            else:
                try:
                    self.f_via.pack_forget()
                    self.btn_precios_dist.grid_remove()
                except Exception:
                    pass
                try:
                    self.btn_unidades.grid()
                except Exception:
                    pass
                self.f_uni.pack(fill="x", padx=5, pady=5, ipady=6, after=f_dias)
        except Exception:
            pass

    # ---------- UNIDADES DEL CLIENTE ----------
    def cargar_unidades_cliente(self):
        """Carga las unidades asignadas del cliente desde clientes_unidades."""
        self.unidades = []
        if not self.cliente_id:
            self.pintar_unidades()
            self.recalcular()
            return
        conn = conectar_db(silencioso=True)
        if not conn:
            self.pintar_unidades()
            self.recalcular()
            return
        try:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, id_vehiculo, unidad, precio_normal, precio_domingo, precio_feriado, horas_dia
                FROM clientes_unidades
                WHERE id_cliente = %s ORDER BY id ASC
            ''', (self.cliente_id,))
            for r in cursor.fetchall():
                self.unidades.append({
                    "id": r[0], "id_vehiculo": r[1], "unidad": str(r[2] or ""),
                    "precio_normal": float(r[3] or 0), "precio_domingo": float(r[4] or 0),
                    "precio_feriado": float(r[5] or 0), "horas_dia": float(r[6] or 0),
                    "ded": {"normal": [0.0, 0.0], "domingo": [0.0, 0.0], "feriado": [0.0, 0.0]},
                    "ded_asientos": [],
                    "extras_asientos": [],
                })
            cursor.close()
        except Exception as e:
            print("[Unidades Error]", e)
        finally:
            liberar_conexion(conn)
        self.pintar_unidades()
        self.recalcular()

    def _persistir_unidades_cliente(self):
        """Guarda las unidades actuales (vehículos, horas y precios) en la configuración
        del cliente para que el próximo cálculo las recuerde."""
        if not self.cliente_id or self.registro_editando is not None:
            return  # no sobrescribir la configuración al editar un registro histórico
        conn = conectar_db()
        if not conn:
            return
        try:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM clientes_unidades WHERE id_cliente=%s", (self.cliente_id,))
            for u in self.unidades:
                cursor.execute('''
                    INSERT INTO clientes_unidades
                    (id_cliente, id_vehiculo, unidad, precio_normal, precio_domingo, precio_feriado, horas_dia)
                    VALUES (%s,%s,%s,%s,%s,%s,%s)
                ''', (self.cliente_id, u.get("id_vehiculo"), u["unidad"],
                      u["precio_normal"], u["precio_domingo"], u["precio_feriado"], u["horas_dia"]))
            conn.commit()
        except Exception as e:
            print("[Persistir Unidades Error]", e)
        finally:
            liberar_conexion(conn)

    def pintar_unidades(self):
        for item in self.tabla_unidades.get_children():
            self.tabla_unidades.delete(item)
        if not self.unidades:
            self.lbl_sin_unidades.configure(
                text="⚠️ Este cliente no tiene unidades asignadas. Usa '🚗 Unidades del Cliente' para asignarlas.")
        else:
            self.lbl_sin_unidades.configure(text="")
        for u in self.unidades:
            subt = u.get("_subtotal", 0)
            self.tabla_unidades.insert("", tk.END, values=(
                u["unidad"],
                f"{float(u['horas_dia'] or 0):g}",
                f"{float(u['precio_normal'] or 0):g}",
                f"{float(u['precio_domingo'] or 0):g}",
                f"{float(u['precio_feriado'] or 0):g}",
                formatear_moneda(subt),
            ))

    def abrir_ventana_unidades(self):
        if not self.cliente_id:
            messagebox.showwarning("Falta Cliente", "Primero seleccione un cliente de la base de datos.")
            return
        VentanaUnidadesCliente(self.parent, self, self.cliente_id, self.cliente_nombre)

    def agregar_unidad_flota(self):
        if not self.cliente_id:
            messagebox.showwarning("Falta Cliente", "Primero seleccione un cliente.")
            return
        excluir = [u.get("id_vehiculo") for u in self.unidades if u.get("id_vehiculo")]
        dlg = DialogoSeleccionVehiculo(self.parent, self.cliente_id, excluir_ids=excluir)
        if dlg.result:
            agregados = 0
            for v in dlg.result:
                if any(u.get("id_vehiculo") == v["id_vehiculo"] for u in self.unidades):
                    continue  # ya está asignado
                self.unidades.append({
                    "id": None, "id_vehiculo": v["id_vehiculo"], "unidad": v["unidad"],
                    "precio_normal": 0.0, "precio_domingo": 0.0, "precio_feriado": 0.0,
                    "horas_dia": 8.0,
                    "ded": {"normal": [0.0, 0.0], "domingo": [0.0, 0.0], "feriado": [0.0, 0.0]},
                    "ded_asientos": [],
                    "extras_asientos": [],
                })
                agregados += 1
            self.pintar_unidades()
            self.recalcular()
            self._persistir_unidades_cliente()  # guarda los vehículos para el próximo cálculo
            if agregados > 1:
                messagebox.showinfo("Unidades agregadas", f"{agregados} vehículos agregados al cálculo.")

    def editar_unidad_seleccionada(self):
        sel = self.tabla_unidades.selection()
        if not sel:
            messagebox.showinfo("Seleccione una unidad", "Seleccione una unidad de la tabla para editar sus precios, horas/día y deducciones.")
            return
        idx = self.tabla_unidades.index(sel[0])
        u = self.unidades[idx]
        dlg = DialogoUnidad(self.parent, u, self.plan_cobro, con_deducciones=True, app=self)
        if dlg.result is not None:
            self.unidades[idx] = dlg.result
            self.pintar_unidades()
            self.recalcular()
            self._persistir_unidades_cliente()  # guarda horas y precios para el próximo cálculo

    # ---------- VIAJES POR DISTANCIA (plan Por Punto o Viaje) ----------
    def abrir_ventana_precios_distancia(self):
        if not self.cliente_id:
            messagebox.showwarning("Falta Cliente", "Primero seleccione un cliente de la base de datos.")
            return
        VentanaPreciosDistancia(self.parent, self, self.cliente_id, self.cliente_nombre)

    def cargar_rangos_viaje(self):
        """Carga la tabla de precios por distancia del cliente."""
        self.rangos_viaje = []
        self.viajes_registrados = []
        self.deducciones_registradas = []
        self._cargar_vehiculos_combo()
        if not self.cliente_id:
            self.pintar_viajes()
            return
        conn = conectar_db(silencioso=True)
        if not conn:
            self.pintar_viajes()
            return
        try:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, distancia_desde, distancia_hasta, precio_normal, precio_domingo, precio_feriado
                FROM precios_viaje_distancia
                WHERE id_cliente = %s ORDER BY distancia_desde ASC
            ''', (self.cliente_id,))
            for r in cursor.fetchall():
                self.rangos_viaje.append({
                    "id": r[0], "distancia_desde": float(r[1]), "distancia_hasta": float(r[2]),
                    "precio_normal": float(r[3] or 0), "precio_domingo": float(r[4] or 0),
                    "precio_feriado": float(r[5] or 0),
                    "viajes": {"normal": 0, "domingo": 0, "feriado": 0},
                    "ded": {"normal": 0, "domingo": 0, "feriado": 0},
                })
            cursor.close()
        except Exception as e:
            print("[Rangos Viaje Error]", e)
        finally:
            liberar_conexion(conn)
        self._pintar_deducciones_registradas()
        self.pintar_viajes()
        self.recalcular()

    def pintar_viajes(self):
        familia_fuente = "Helvetica" if sys.platform == "darwin" else "Arial"
        for w in self.frame_viajes_grid.winfo_children():
            w.destroy()
        self.ent_viajes = []
        try:
            if hasattr(self, "combo_tipo_viaje"):
                valores = [f"{rg['distancia_desde']:g} - {rg['distancia_hasta']:g} km" for rg in self.rangos_viaje]
                self.combo_tipo_viaje.configure(values=valores or ["— Seleccione distancia —"])
                self.combo_tipo_viaje.set("— Seleccione distancia —")
                if hasattr(self, "combo_tipo_ded"):
                    self.combo_tipo_ded.configure(values=valores or ["— Seleccione distancia —"])
                    self.combo_tipo_ded.set("— Seleccione distancia —")
        except Exception:
            pass
        if not self.rangos_viaje:
            self.lbl_sin_viajes.configure(
                text="⚠️ Este cliente no tiene tabla de precios por distancia. Usa '📏 Precios por Distancia'.")
            return
        self.lbl_sin_viajes.configure(text="")
        cabeceras = ["Distancia", "P. Normal", "P. Domingo", "P. Feriado",
                     "Viajes\nNormales", "Viajes\nDomingos", "Viajes\nFeriados",
                     "No Real.\nNormales", "No Real.\nDomingos", "No Real.\nFeriados",
                     "Monto"]
        for j, txt in enumerate(cabeceras):
            ctk.CTkLabel(self.frame_viajes_grid, text=txt, font=(familia_fuente, 9, "bold"),
                         text_color="#7f8c8d").grid(row=0, column=j, padx=3, pady=2)
        self.lbl_monto_viaje = []
        for i, rg in enumerate(self.rangos_viaje):
            row = i + 1
            ctk.CTkLabel(self.frame_viajes_grid,
                         text=f"{rg['distancia_desde']:g} - {rg['distancia_hasta']:g} km",
                         font=(familia_fuente, 10, "bold")).grid(row=row, column=0, sticky="w", padx=3, pady=3)
            ctk.CTkLabel(self.frame_viajes_grid, text=formatear_moneda(rg['precio_normal']),
                         font=(familia_fuente, 10)).grid(row=row, column=1, padx=3)
            ctk.CTkLabel(self.frame_viajes_grid, text=formatear_moneda(rg['precio_domingo']),
                         font=(familia_fuente, 10)).grid(row=row, column=2, padx=3)
            ctk.CTkLabel(self.frame_viajes_grid, text=formatear_moneda(rg['precio_feriado']),
                         font=(familia_fuente, 10)).grid(row=row, column=3, padx=3)
            entradas = {"viajes": {}, "ded": {}}
            for col, clave in zip((4, 5, 6), ("normal", "domingo", "feriado")):
                e = ctk.CTkEntry(self.frame_viajes_grid, width=62, justify="center")
                e.grid(row=row, column=col, padx=2, pady=2)
                e.insert(0, str(int(rg["viajes"][clave])))
                entradas["viajes"][clave] = e
            for col, clave in zip((7, 8, 9), ("normal", "domingo", "feriado")):
                e = ctk.CTkEntry(self.frame_viajes_grid, width=62, justify="center")
                e.grid(row=row, column=col, padx=2, pady=2)
                e.insert(0, str(int(rg["ded"][clave])))
                entradas["ded"][clave] = e
            lbl_m = ctk.CTkLabel(self.frame_viajes_grid, text=formatear_moneda(0),
                                 font=(familia_fuente, 10, "bold"), text_color="#1f538d")
            lbl_m.grid(row=row, column=10, padx=3, pady=2)
            self.lbl_monto_viaje.append(lbl_m)
            self.ent_viajes.append(entradas)

    def _leer_int_entry(self, entry, nombre):
        try:
            v = int(float(str(entry.get()).strip().replace(",", "") or 0))
            if v < 0:
                raise ValueError
            return v
        except ValueError:
            raise ValueError(f"'{nombre}' debe ser un número entero mayor o igual a 0.")

    def _leer_viajes_desde_entradas(self):
        """Lee las cantidades de viajes de los campos y las aplica a self.rangos_viaje."""
        for i, rg in enumerate(self.rangos_viaje):
            if i < len(self.ent_viajes):
                for cat in ("normal", "domingo", "feriado"):
                    rg["viajes"][cat] = self._leer_int_entry(self.ent_viajes[i]["viajes"][cat], f"Viajes {cat}")
                    rg["ded"][cat] = self._leer_int_entry(self.ent_viajes[i]["ded"][cat], f"No realizados {cat}")

    # ---------- REGISTRO RÁPIDO DE VIAJES (plan Por Punto o Viaje) ----------
    def _cargar_vehiculos_combo(self):
        """Carga los vehículos de la flota en el selector de vehículo del viaje."""
        try:
            valores = ["— Seleccione vehículo —"]
            conn = conectar_db(silencioso=True)
            if conn:
                try:
                    cursor = conn.cursor()
                    cursor.execute("SELECT placa, marca, modelo, anio, color FROM flota_vehiculos ORDER BY placa")
                    for r in cursor.fetchall():
                        desc = " ".join(str(x) for x in (r[1], r[2], r[3], r[4]) if x and str(x).strip())
                        valores.append(f"{r[0]} — {desc}".strip())
                    cursor.close()
                finally:
                    liberar_conexion(conn)
            if hasattr(self, "combo_vehiculo_viaje"):
                self.combo_vehiculo_viaje.configure(values=valores)
                self.combo_vehiculo_viaje.set("— Seleccione vehículo —")
        except Exception as e:
            print("[Vehiculos Combo Error]", e)

    def abrir_calendario_viaje(self):
        # Abre el calendario en el mes/año seleccionado y resalta la quincena
        anio, mes, _ = self.periodo_actual()
        dias_q = {}
        for d in self.dias:
            if d["fecha"].year == anio and d["fecha"].month == mes:
                dias_q[d["fecha"].day] = d["categoria"]
        CalendarioPopup(self.parent, self.ent_fecha_viaje, self._al_seleccionar_fecha_viaje,
                        anio=anio, mes=mes, dias_resaltados=dias_q)

    def _fecha_desde_entry(self):
        try:
            return datetime.strptime(self.ent_fecha_viaje.get().strip(), "%d/%m/%Y").date()
        except Exception:
            return None

    def _categoria_para_fecha(self, fecha):
        # Usa la clasificación de la quincena (respeta feriados marcados manualmente)
        for d in self.dias:
            if d["fecha"] == fecha:
                return d["categoria"]
        return clasificar_dia(fecha, self.feriados)

    def _al_seleccionar_fecha_viaje(self):
        fecha = self._fecha_desde_entry()
        if fecha:
            cat = self._categoria_para_fecha(fecha)
            self.lbl_dia_detectado.configure(text=f"Día detectado: {dict(CATEGORIAS).get(cat, '')}")
        else:
            self.lbl_dia_detectado.configure(text="")

    def _pintar_viajes_registrados(self):
        for item in self.tabla_viajes_reg.get_children():
            self.tabla_viajes_reg.delete(item)
        for t in self.viajes_registrados:
            etiqueta = dict(CATEGORIAS)[t["categoria"]]
            rango = self.rangos_viaje[t["rango_idx"]]
            self.tabla_viajes_reg.insert("", tk.END, values=(
                t["fecha"].strftime("%d/%m/%Y"),
                etiqueta,
                f"{rango['distancia_desde']:g} - {rango['distancia_hasta']:g} km",
                t.get("vehiculo", ""),
            ))

    def _scroll_solo_tabla_viajes(self, event):
        """Desplaza solo la lista de viajes registrados (la rueda no mueve la ventana)."""
        try:
            if event.num == 4:
                self.tabla_viajes_reg.yview_scroll(-3, "units")
            elif event.num == 5:
                self.tabla_viajes_reg.yview_scroll(3, "units")
            else:
                delta = int(-1 * (event.delta / 120))
                self.tabla_viajes_reg.yview_scroll(delta, "units")
        except Exception:
            pass
        return "break"

    def agregar_viaje_registrado(self):
        if not self.dias:
            messagebox.showwarning("Sin quincena", "Primero busque los días de la quincena.")
            return
        if not self.rangos_viaje:
            messagebox.showwarning("Sin tabla de distancias",
                                   "Este cliente no tiene precios por distancia. Usa '📏 Precios por Distancia'.")
            return
        fecha = self._fecha_desde_entry()
        if not fecha:
            messagebox.showwarning("Fecha inválida", "Ingrese una fecha válida (dd/mm/aaaa) o use el botón 📅.")
            return
        if not any(d["fecha"] == fecha for d in self.dias):
            messagebox.showwarning("Fuera de la quincena",
                                   f"La fecha {fecha.strftime('%d/%m/%Y')} no pertenece a la quincena seleccionada.")
            return
        tipo = self.combo_tipo_viaje.get()
        if tipo.startswith("—"):
            messagebox.showwarning("Falta tipo de viaje", "Seleccione el tipo de viaje (rango de distancia).")
            return
        vehiculo = self.combo_vehiculo_viaje.get()
        if not vehiculo or vehiculo.startswith("—"):
            messagebox.showwarning("Falta vehículo", "Seleccione el vehículo que realizó el viaje.")
            return
        valores = list(self.combo_tipo_viaje.cget("values"))
        if tipo not in valores:
            return
        idx = valores.index(tipo)
        if idx >= len(self.rangos_viaje):
            return
        cat = self._categoria_para_fecha(fecha)
        rg = self.rangos_viaje[idx]
        precio = rg.get(f"precio_{cat}", 0)
        self.rangos_viaje[idx]["viajes"][cat] += 1
        self.viajes_registrados.append({
            "fecha": fecha, "categoria": cat, "rango_idx": idx,
            "vehiculo": vehiculo, "precio": float(precio or 0),
            "monto": float(precio or 0),
            "desde": rg["distancia_desde"], "hasta": rg["distancia_hasta"],
        })
        self._pintar_viajes_registrados()
        self.pintar_viajes()   # actualiza la tabla de viajes
        self.recalcular()      # calcula en caliente los montos y el total
        self.ent_fecha_viaje.delete(0, tk.END)
        self.lbl_dia_detectado.configure(text="")

    def quitar_viaje_registrado(self):
        sel = self.tabla_viajes_reg.selection()
        if not sel:
            messagebox.showinfo("Seleccione un viaje", "Seleccione un viaje de la lista para quitarlo.")
            return
        idx = self.tabla_viajes_reg.index(sel[0])
        t = self.viajes_registrados[idx]
        self.rangos_viaje[t["rango_idx"]]["viajes"][t["categoria"]] = max(
            0, self.rangos_viaje[t["rango_idx"]]["viajes"][t["categoria"]] - 1)
        self.viajes_registrados.pop(idx)
        self._pintar_viajes_registrados()
        self.pintar_viajes()   # actualiza la tabla de viajes
        self.recalcular()      # recalcula en caliente los montos y el total

    # ---------- DEDUCCIONES CON FECHA (plan Por Punto o Viaje) ----------
    def abrir_calendario_ded(self):
        """Abre el calendario de la quincena para elegir la fecha de la deducción."""
        anio, mes, _ = self.periodo_actual()
        dias_q = {}
        for d in self.dias:
            if d["fecha"].year == anio and d["fecha"].month == mes:
                dias_q[d["fecha"].day] = d["categoria"]
        CalendarioPopup(self.parent, self.ent_fecha_ded, self._al_seleccionar_fecha_ded,
                        anio=anio, mes=mes, dias_resaltados=dias_q)

    def _fecha_desde_entry_ded(self):
        try:
            return datetime.strptime(self.ent_fecha_ded.get().strip(), "%d/%m/%Y").date()
        except Exception:
            return None

    def _al_seleccionar_fecha_ded(self):
        fecha = self._fecha_desde_entry_ded()
        if fecha:
            cat = self._categoria_para_fecha(fecha)
            self.lbl_dia_detectado_ded.configure(text=f"Día detectado: {dict(CATEGORIAS).get(cat, '')}")
        else:
            self.lbl_dia_detectado_ded.configure(text="")

    def _pintar_deducciones_registradas(self):
        for item in self.tabla_ded_reg.get_children():
            self.tabla_ded_reg.delete(item)
        total = 0.0
        for t in self.deducciones_registradas:
            total += float(t.get("monto") or 0)
            self.tabla_ded_reg.insert("", tk.END, values=(
                etiqueta_fecha_hora(t.get("fecha")),
                dict(CATEGORIAS).get(t.get("categoria"), ""),
                f"{float(t.get('desde') or 0):g} - {float(t.get('hasta') or 0):g} km",
                f"{float(t.get('cantidad') or 0):g}",
                str(t.get("motivo") or ""),
            ))
        try:
            self.lbl_total_ded.configure(
                text=f"TOTAL DEDUCCIONES REGISTRADAS: {formatear_moneda(total)}")
        except Exception:
            pass

    def _scroll_solo_tabla_ded(self, event):
        """Desplaza solo la lista de deducciones registradas (la rueda no mueve la ventana)."""
        try:
            if event.num == 4:
                self.tabla_ded_reg.yview_scroll(-3, "units")
            elif event.num == 5:
                self.tabla_ded_reg.yview_scroll(3, "units")
            else:
                delta = int(-1 * (event.delta / 120))
                self.tabla_ded_reg.yview_scroll(delta, "units")
        except Exception:
            pass
        return "break"

    def agregar_deduccion_registrada(self):
        """Registra un asiento de deducción con fecha (viajes/puntos no realizados)."""
        if not self.dias:
            messagebox.showwarning("Sin quincena", "Primero busque los días de la quincena.")
            return
        if not self.rangos_viaje:
            messagebox.showwarning("Sin tabla de distancias",
                                   "Este cliente no tiene precios por distancia. Usa '📏 Precios por Distancia'.")
            return
        fecha = self._fecha_desde_entry_ded()
        if not fecha:
            messagebox.showwarning("Fecha inválida", "Ingrese una fecha válida (dd/mm/aaaa) o use el botón 📅.")
            return
        if not any(d["fecha"] == fecha for d in self.dias):
            messagebox.showwarning("Fuera de la quincena",
                                   f"La fecha {fecha.strftime('%d/%m/%Y')} no pertenece a la quincena seleccionada.")
            return
        tipo = self.combo_tipo_ded.get()
        if not tipo or tipo.startswith("—"):
            messagebox.showwarning("Falta tipo de viaje", "Seleccione el rango de distancia de la deducción.")
            return
        valores = list(self.combo_tipo_ded.cget("values"))
        if tipo not in valores:
            return
        idx = valores.index(tipo)
        if idx >= len(self.rangos_viaje):
            return
        try:
            cant = self._leer_int_entry(self.ent_cant_ded, "Cantidad no realizada")
        except ValueError as e:
            messagebox.showwarning("Dato inválido", str(e))
            return
        if cant <= 0:
            messagebox.showwarning("Dato inválido",
                                   "La cantidad de viajes/puntos no realizados debe ser mayor a 0.")
            return
        cat = self._categoria_para_fecha(fecha)
        rg = self.rangos_viaje[idx]
        precio = float(rg.get(f"precio_{cat}", 0) or 0)
        self.deducciones_registradas.append({
            "fecha": fecha, "categoria": cat, "rango_idx": idx,
            "cantidad": cant, "motivo": self.ent_motivo_ded.get().strip(),
            "precio": precio, "monto": precio * cant,
            "desde": rg["distancia_desde"], "hasta": rg["distancia_hasta"],
        })
        self._pintar_deducciones_registradas()
        self.recalcular()   # recalcula en caliente los montos y el total
        self.ent_fecha_ded.delete(0, tk.END)
        self.ent_cant_ded.delete(0, tk.END)
        self.ent_motivo_ded.delete(0, tk.END)
        self.lbl_dia_detectado_ded.configure(text="")

    def quitar_deduccion_registrada(self):
        sel = self.tabla_ded_reg.selection()
        if not sel:
            messagebox.showinfo("Seleccione una deducción", "Seleccione una deducción de la lista para quitarla.")
            return
        idx = self.tabla_ded_reg.index(sel[0])
        if 0 <= idx < len(self.deducciones_registradas):
            self.deducciones_registradas.pop(idx)
        self._pintar_deducciones_registradas()
        self.recalcular()

    # ---------- DÍAS DE LA QUINCENA ----------
    def periodo_actual(self):
        mes = int(self.combo_mes.get().split(" - ")[0])
        anio = int(self.combo_anio.get())
        quincena = 1 if self.combo_quincena.get().startswith("Primera") else 2
        return anio, mes, quincena

    def buscar_dias_quincena(self):
        if not self.cliente_id:
            messagebox.showwarning("Falta Cliente", "Primero seleccione un cliente de la base de datos.")
            return
        anio, mes, quincena = self.periodo_actual()
        self.registro_editando = None
        self.feriados = feriados_locales(anio)
        if _leer_cache_feriados().get(str(anio)):
            self.lbl_origen_feriados.configure(text="Feriados: caché local de Perú (consultando internet…).")
        else:
            self.lbl_origen_feriados.configure(text="Feriados: lista local de Perú (consultando internet…).")
        self.construir_dias()
        if not self._buscando_feriados:
            self._buscando_feriados = True
            threading.Thread(target=self._hilo_feriados, args=(anio,), daemon=True).start()

    def _hilo_feriados(self, anio):
        try:
            fer = _descargar_feriados_internet(anio)
            if fer:
                _guardar_cache_feriados(anio, fer)
        except Exception:
            fer = {}
        finally:
            self._buscando_feriados = False
        try:
            self.parent.after(0, lambda: self._aplicar_feriados_remotos(anio, fer))
        except Exception:
            pass

    def _aplicar_feriados_remotos(self, anio, fer):
        if not fer:
            self.lbl_origen_feriados.configure(text="Feriados: lista local de Perú (sin internet o sin respuesta).")
            return
        anio_act, _, _ = self.periodo_actual()
        if anio_act != anio:
            return
        self.feriados = fer
        self.lbl_origen_feriados.configure(text="Feriados: descargados de internet (actualizados).")
        self.construir_dias()

    def construir_dias(self):
        """Reconstruye la lista de días desde la lista de feriados y la pinta."""
        anio, mes, quincena = self.periodo_actual()
        self.dias = dias_de_quincena(anio, mes, quincena, self.feriados)
        self.pintar_dias()

    def pintar_dias(self):
        """Pinta la tabla de días tal como está en self.dias (sin reconstruirla)."""
        for item in self.tabla_dias.get_children():
            self.tabla_dias.delete(item)
        for d in self.dias:
            etiqueta_cat = dict(CATEGORIAS)[d["categoria"]]
            nombre = d["feriado_nombre"] if d["categoria"] == "feriado" else ""
            self.tabla_dias.insert("", tk.END, values=(
                d["fecha"].strftime("%d/%m/%Y"),
                DIAS_SEMANA[d["fecha"].weekday()],
                etiqueta_cat,
                nombre,
            ))
        self.recalcular()

    def toggle_feriado(self):
        sel = self.tabla_dias.selection()
        if not sel:
            messagebox.showinfo("Seleccione una fila", "Seleccione un día de la quincena en la tabla para marcar/desmarcar feriado.")
            return
        idx = self.tabla_dias.index(sel[0])
        d = self.dias[idx]
        if d["categoria"] == "feriado":
            nueva = clasificar_dia(d["fecha"], {})
            d["categoria"] = nueva
            d["feriado_nombre"] = ""
        else:
            d["categoria"] = "feriado"
            d["feriado_nombre"] = "Marcado manualmente"
        self.pintar_dias()  # conserva la marca manual (no reconstruye desde feriados)

    def restablecer_feriados(self):
        if not self.dias:
            messagebox.showinfo("Sin quincena", "Primero busque los días de la quincena.")
            return
        anio, _, _ = self.periodo_actual()
        self.feriados = feriados_locales(anio)
        if _leer_cache_feriados().get(str(anio)):
            self.lbl_origen_feriados.configure(text="Feriados: caché local de Perú (consultando internet…).")
        else:
            self.lbl_origen_feriados.configure(text="Feriados: lista local de Perú (consultando internet…).")
        self.construir_dias()
        if not self._buscando_feriados:
            self._buscando_feriados = True
            threading.Thread(target=self._hilo_feriados, args=(anio,), daemon=True).start()

    # ---------- CÁLCULO ----------
    def _leer_float(self, valor, nombre):
        try:
            v = float(str(valor).strip().replace(",", "") or 0)
            if v < 0:
                raise ValueError
            return v
        except ValueError:
            raise ValueError(f"'{nombre}' debe ser un número válido (mayor o igual a 0).")

    def recopilar_calculo(self):
        """Calcula el desglose (por unidad o por viajes) y devuelve montos y total."""
        if not self.dias:
            raise ValueError("Primero busque los días de la quincena.")
        c = conteos_quincena(self.dias)
        if self.plan_cobro == "Por Punto o Viaje":
            return self._recopilar_viajes(c)
        # Por Hora: desglose por unidad
        if not self.unidades:
            raise ValueError("Este cliente no tiene unidades asignadas. Usa '🚗 Unidades del Cliente'.")

        unidades_calc = []
        monto_base_total = 0.0
        monto_ded_total = 0.0
        monto_extras_total = 0.0
        horas_extras_total = 0.0
        total_general = 0.0
        ded_global = {"normal": 0.0, "domingo": 0.0, "feriado": 0.0}
        monto_cat_global = {"normal": 0.0, "domingo": 0.0, "feriado": 0.0}
        # Asientos con fecha (deducciones y horas extras) para el detalle del PDF
        deducciones_detalle = []

        for u in self.unidades:
            pn = self._leer_float(u["precio_normal"], "Precio Día Normal")
            pd = self._leer_float(u["precio_domingo"], "Precio Domingo")
            pf = self._leer_float(u["precio_feriado"], "Precio Feriado")
            horas_dia = self._leer_float(u["horas_dia"], "Horas al día")
            if horas_dia <= 0:
                horas_dia = 24.0  # respaldo si no se definió jornada
            precio_cat = {"normal": pn, "domingo": pd, "feriado": pf}

            base_u = 0.0
            ded_u = 0.0
            ded_horas_u = {"normal": 0.0, "domingo": 0.0, "feriado": 0.0}
            monto_cat_u = {"normal": 0.0, "domingo": 0.0, "feriado": 0.0}

            for clave, _ in CATEGORIAS:
                cant = c[clave]
                precio = precio_cat[clave]
                base_cat = cant * precio * horas_dia  # Por Hora: precio × horas del día
                monto_cat_u[clave] = base_cat
                monto_cat_global[clave] += base_cat
                base_u += base_cat

                h_ded, m_ded = u["ded"].get(clave, [0.0, 0.0])
                horas = self._leer_float(h_ded, f"Deducción horas {clave}") + self._leer_float(m_ded, f"Deducción minutos {clave}") / 60.0
                ded_horas_u[clave] = horas
                ded_global[clave] += horas
                monto_ded_cat = horas * precio  # Por Hora: horas ausentes × precio por hora
                ded_u += monto_ded_cat

            # Asientos de deducción con fecha de esta unidad (detalle para el PDF)
            for a in (u.get("ded_asientos") or []):
                horas_a = horas_de_asiento(a)
                if horas_a <= 0:
                    continue
                cat_a = a.get("categoria") if a.get("categoria") in precio_cat else "normal"
                deducciones_detalle.append({
                    "tipo": "DEDUCCION",
                    "plan": "Por Hora", "unidad": u["unidad"],
                    "fecha": a.get("fecha"), "categoria": cat_a,
                    "horas": float(a.get("horas") or 0), "minutos": float(a.get("minutos") or 0),
                    "cantidad": 0.0, "desde": None, "hasta": None,
                    "precio": precio_cat[cat_a], "motivo": str(a.get("motivo") or ""),
                    "monto": horas_a * precio_cat[cat_a],
                })

            # Asientos de HORAS EXTRAS con fecha: SUMAN al cobro, valorizadas al
            # precio por hora de la unidad según el tipo de día del asiento.
            extras_u = 0.0
            horas_extras_u = 0.0
            for a in (u.get("extras_asientos") or []):
                horas_a = horas_de_asiento(a)
                if horas_a <= 0:
                    continue
                cat_a = a.get("categoria") if a.get("categoria") in precio_cat else "normal"
                monto_a = horas_a * precio_cat[cat_a]
                horas_extras_u += horas_a
                extras_u += monto_a
                deducciones_detalle.append({
                    "tipo": "EXTRA",
                    "plan": "Por Hora", "unidad": u["unidad"],
                    "fecha": a.get("fecha"), "categoria": cat_a,
                    "horas": float(a.get("horas") or 0), "minutos": float(a.get("minutos") or 0),
                    "cantidad": 0.0, "desde": None, "hasta": None,
                    "precio": precio_cat[cat_a], "motivo": str(a.get("motivo") or ""),
                    "monto": monto_a,
                })

            subtotal_u = base_u - ded_u + extras_u
            if subtotal_u < 0:
                subtotal_u = 0.0
            monto_ded_total += ded_u
            monto_extras_total += extras_u
            horas_extras_total += horas_extras_u
            monto_base_total += base_u
            total_general += subtotal_u

            unidades_calc.append({
                "unidad": u["unidad"],
                "horas_dia": horas_dia if horas_dia != 24.0 or float(u["horas_dia"] or 0) > 0 else 0.0,
                "precio_normal": pn, "precio_domingo": pd, "precio_feriado": pf,
                "monto_normal": monto_cat_u["normal"],
                "monto_domingo": monto_cat_u["domingo"],
                "monto_feriado": monto_cat_u["feriado"],
                "ded_normal_h": ded_horas_u["normal"],
                "ded_domingo_h": ded_horas_u["domingo"],
                "ded_feriado_h": ded_horas_u["feriado"],
                "monto_deducciones": ded_u,
                "horas_extras": horas_extras_u,
                "monto_extras": extras_u,
                "subtotal": subtotal_u,
            })

        if total_general < 0:
            total_general = 0.0

        return {
            "conteos": c,
            "unidades": unidades_calc,
            "monto_base": monto_base_total,
            "monto_deducciones": monto_ded_total,
            "horas_extras": horas_extras_total,
            "monto_extras": monto_extras_total,
            "monto_normal": monto_cat_global["normal"],
            "monto_domingo": monto_cat_global["domingo"],
            "monto_feriado": monto_cat_global["feriado"],
            "ded_normal_h": ded_global["normal"],
            "ded_domingo_h": ded_global["domingo"],
            "ded_feriado_h": ded_global["feriado"],
            "deducciones_detalle": deducciones_detalle,
            "total": total_general,
        }

    def _recopilar_viajes(self, c):
        """Calcula el cobro por viajes según la distancia y el tipo de día."""
        if not self.rangos_viaje:
            raise ValueError("Este cliente no tiene tabla de precios por distancia. Usa '📏 Precios por Distancia'.")
        self._leer_viajes_desde_entradas()
        # Asientos de deducción registrados con FECHA: se suman a los "no realizados"
        # escritos en el cuadro, guardando su detalle aparte para el PDF.
        ded_reg = {}
        for t in self.deducciones_registradas:
            idx = t.get("rango_idx")
            if idx is None or not (0 <= idx < len(self.rangos_viaje)):
                continue
            cat = t.get("categoria") if t.get("categoria") in ("normal", "domingo", "feriado") else "normal"
            ded_reg.setdefault(idx, {"normal": 0, "domingo": 0, "feriado": 0})
            ded_reg[idx][cat] += int(t.get("cantidad") or 0)

        monto_base = 0.0
        monto_ded = 0.0
        total = 0.0
        monto_cat = {"normal": 0.0, "domingo": 0.0, "feriado": 0.0}
        detalle = []
        deducciones_detalle = []
        for i, rg in enumerate(self.rangos_viaje):
            pn = self._leer_float(rg["precio_normal"], "Precio Normal")
            pd = self._leer_float(rg["precio_domingo"], "Precio Domingo")
            pf = self._leer_float(rg["precio_feriado"], "Precio Feriado")
            extra = ded_reg.get(i, {})
            nr_normal = int(rg["ded"]["normal"] or 0) + int(extra.get("normal", 0))
            nr_domingo = int(rg["ded"]["domingo"] or 0) + int(extra.get("domingo", 0))
            nr_feriado = int(rg["ded"]["feriado"] or 0) + int(extra.get("feriado", 0))
            base_r = rg["viajes"]["normal"] * pn + rg["viajes"]["domingo"] * pd + rg["viajes"]["feriado"] * pf
            ded_r = nr_normal * pn + nr_domingo * pd + nr_feriado * pf
            monto_cat["normal"] += (rg["viajes"]["normal"] - nr_normal) * pn
            monto_cat["domingo"] += (rg["viajes"]["domingo"] - nr_domingo) * pd
            monto_cat["feriado"] += (rg["viajes"]["feriado"] - nr_feriado) * pf
            sub = max(0.0, base_r - ded_r)
            monto_base += base_r
            monto_ded += ded_r
            total += sub
            # OJO: rg["ded"] conserva solo lo escrito en el cuadro; los asientos con
            # fecha se suman aquí (no se escriben de vuelta para no duplicarlos al
            # repintar el cuadro). En el registro guardado sí van ya sumados.
            detalle.append({
                "desde": rg["distancia_desde"], "hasta": rg["distancia_hasta"],
                "distancia": f"{rg['distancia_desde']:g} - {rg['distancia_hasta']:g} km",
                "precio_normal": pn, "precio_domingo": pd, "precio_feriado": pf,
                "viajes_normal": rg["viajes"]["normal"], "viajes_domingo": rg["viajes"]["domingo"],
                "viajes_feriado": rg["viajes"]["feriado"],
                "ded_normal": nr_normal, "ded_domingo": nr_domingo, "ded_feriado": nr_feriado,
                "monto_base": base_r, "monto_deducciones": ded_r, "subtotal": sub,
            })

        # Detalle de los asientos de deducción con fecha (para el PDF)
        for t in self.deducciones_registradas:
            cant = float(t.get("cantidad") or 0)
            if cant <= 0:
                continue
            cat = t.get("categoria") if t.get("categoria") in ("normal", "domingo", "feriado") else "normal"
            precio_cat = {"normal": 0.0, "domingo": 0.0, "feriado": 0.0}
            idx = t.get("rango_idx")
            if idx is not None and 0 <= idx < len(detalle):
                precio_cat = {"normal": detalle[idx]["precio_normal"],
                              "domingo": detalle[idx]["precio_domingo"],
                              "feriado": detalle[idx]["precio_feriado"]}
            deducciones_detalle.append({
                "tipo": "DEDUCCION",
                "plan": "Por Punto o Viaje",
                "unidad": f"{float(t.get('desde') or 0):g} - {float(t.get('hasta') or 0):g} km",
                "fecha": t.get("fecha"), "categoria": cat,
                "horas": 0.0, "minutos": 0.0, "cantidad": cant,
                "desde": t.get("desde"), "hasta": t.get("hasta"),
                "precio": precio_cat[cat], "motivo": str(t.get("motivo") or ""),
                "monto": precio_cat[cat] * cant,
            })

        return {
            "conteos": c,
            "viajes": detalle,
            "monto_base": monto_base,
            "monto_deducciones": monto_ded,
            "horas_extras": 0.0,
            "monto_extras": 0.0,
            "monto_normal": monto_cat["normal"],
            "monto_domingo": monto_cat["domingo"],
            "monto_feriado": monto_cat["feriado"],
            "ded_normal_h": 0.0, "ded_domingo_h": 0.0, "ded_feriado_h": 0.0,
            "deducciones_detalle": deducciones_detalle,
            "total": total,
        }

    def recalcular(self):
        if not self.dias:
            self.lbl_resumen.configure(text="Busque los días de la quincena para ver el resumen.")
            self.lbl_total.configure(text="TOTAL A COBRAR: S/ 0.00")
            self.lbl_conteo.configure(text="")
            return
        try:
            r = self.recopilar_calculo()
        except ValueError as e:
            self.lbl_resumen.configure(text=f"⚠️ {e}")
            self.lbl_total.configure(text="TOTAL A COBRAR: S/ 0.00")
            # El conteo de días sigue siendo válido aunque falten unidades/precios
            c = conteos_quincena(self.dias)
            self.lbl_conteo.configure(
                text=(f"Días Normales (Lun–Sáb): {c['normal']}  |  Domingos: {c['domingo']}  |  "
                      f"Feriados: {c['feriado']}  |  Total días: {len(self.dias)}"))
            return
        c = r["conteos"]
        self.lbl_conteo.configure(
            text=(f"Días Normales (Lun–Sáb): {c['normal']}  |  Domingos: {c['domingo']}  |  "
                  f"Feriados: {c['feriado']}  |  Total días: {len(self.dias)}"))

        lineas = []
        if "viajes" in r:
            # Plan Por Punto o Viaje: resumen por rango de distancia
            for i, vj in enumerate(r["viajes"]):
                lineas.append(
                    f"🛣️ {vj['distancia']}: base {formatear_moneda(vj['monto_base'])}"
                    f" − no realizados {formatear_moneda(vj['monto_deducciones'])} = {formatear_moneda(vj['subtotal'])}")
                # Actualiza el monto por rango y el total en la tabla de viajes
                if hasattr(self, "lbl_monto_viaje") and i < len(self.lbl_monto_viaje):
                    self.lbl_monto_viaje[i].configure(text=formatear_moneda(vj["subtotal"]))
            try:
                self.lbl_total_viajes.configure(text=f"TOTAL VIAJES: {formatear_moneda(r['total'])}")
            except Exception:
                pass
        else:
            # Por Hora: resumen por unidad
            for i, uc in enumerate(r["unidades"]):
                if i < len(self.unidades):
                    self.unidades[i]["_subtotal"] = uc["subtotal"]
                texto_extra = ""
                if uc.get("monto_extras"):
                    texto_extra = f" + horas extras {formatear_moneda(uc['monto_extras'])}"
                lineas.append(
                    f"🚗 {uc['unidad']}: base {formatear_moneda(uc['monto_normal'] + uc['monto_domingo'] + uc['monto_feriado'])}"
                    f" − deducciones {formatear_moneda(uc['monto_deducciones'])}{texto_extra}"
                    f" = {formatear_moneda(uc['subtotal'])}")
            self.pintar_unidades()  # refresca la tabla con los subtotales
        lineas.append(f"SUBTOTAL (base): {formatear_moneda(r['monto_base'])}")
        if r.get("monto_extras"):
            lineas.append(f"HORAS EXTRAS: + {formatear_moneda(r['monto_extras'])}"
                          f"  ({r.get('horas_extras', 0):g} h)")
        if r["monto_deducciones"] > 0:
            lineas.append(f"DEDUCCIONES: - {formatear_moneda(r['monto_deducciones'])}")
            if r["monto_deducciones"] >= r["monto_base"]:
                lineas.append("⚠️ Las deducciones superan el subtotal: el total se muestra en 0.")
        self.lbl_resumen.configure(text="\n".join(lineas))
        self.lbl_total.configure(text=f"TOTAL A COBRAR: {formatear_moneda(r['total'])}")

    # ---------- GUARDAR ----------
    def guardar_quincena(self):
        if not self.cliente_id:
            messagebox.showwarning("Falta Cliente", "Primero seleccione un cliente.")
            return False
        if not self.dias:
            messagebox.showwarning("Sin Quincena", "Primero presione 'Buscar días de la quincena'.")
            return False
        try:
            r = self.recopilar_calculo()
        except ValueError as e:
            messagebox.showerror("Datos Inválidos", str(e))
            return False

        es_nuevo = self.registro_editando is None
        es_viaje = self.plan_cobro == "Por Punto o Viaje"

        conn = conectar_db()
        if not conn:
            messagebox.showwarning("Modo Lectura", "Estás sin conexión a internet.\nNo se puede guardar la quincena.")
            return False

        anio, mes, quincena = self.periodo_actual()
        notas = self.ent_notas.get().strip()
        fecha_reg = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        try:
            cursor = conn.cursor()
            id_cob = self.registro_editando

            if id_cob is None:
                cursor.execute(
                    "SELECT id FROM cobranza_quincenas WHERE id_cliente=%s AND anio=%s AND mes=%s AND quincena=%s",
                    (self.cliente_id, anio, mes, quincena))
                existe = cursor.fetchone()
                if existe:
                    if not messagebox.askyesno(
                            "Registro Existente",
                            f"Ya existe un cálculo de cobranza para este cliente y periodo.\n\n¿Desea actualizar el registro existente?"):
                        return False
                    id_cob = existe[0]

            if id_cob is None:
                cursor.execute('''
                    INSERT INTO cobranza_quincenas
                    (id_cliente, cliente_nombre, cliente_ruc, anio, mes, quincena, plan_cobro,
                     precio_normal, precio_domingo, precio_feriado,
                     cant_lunvie, cant_sabado, cant_domingo, cant_feriado,
                     ded_normal_h, ded_sabado_h, ded_domingo_h, ded_feriado_h,
                     monto_lunvie, monto_sabado, monto_domingo, monto_feriado,
                     monto_base, monto_deducciones, total, notas, pdf_ruta, fecha_registro)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,NULL,%s)
                    RETURNING id
                ''', (self.cliente_id, self.cliente_nombre, self.cliente_ruc, anio, mes, quincena, self.plan_cobro,
                      0, 0, 0,
                      r["conteos"]["normal"], 0, r["conteos"]["domingo"], r["conteos"]["feriado"],
                      r["ded_normal_h"], 0, r["ded_domingo_h"], r["ded_feriado_h"],
                      r["monto_normal"], 0, r["monto_domingo"], r["monto_feriado"],
                      r["monto_base"], r["monto_deducciones"], r["total"], notas, fecha_reg))
                id_cob = cursor.fetchone()[0]
            else:
                cursor.execute('''
                    UPDATE cobranza_quincenas SET
                     id_cliente=%s, cliente_nombre=%s, cliente_ruc=%s, anio=%s, mes=%s, quincena=%s, plan_cobro=%s,
                     precio_normal=%s, precio_domingo=%s, precio_feriado=%s,
                     cant_lunvie=%s, cant_sabado=%s, cant_domingo=%s, cant_feriado=%s,
                     ded_normal_h=%s, ded_sabado_h=%s, ded_domingo_h=%s, ded_feriado_h=%s,
                     monto_lunvie=%s, monto_sabado=%s, monto_domingo=%s, monto_feriado=%s,
                     monto_base=%s, monto_deducciones=%s, total=%s, notas=%s, fecha_registro=%s
                    WHERE id=%s
                ''', (self.cliente_id, self.cliente_nombre, self.cliente_ruc, anio, mes, quincena, self.plan_cobro,
                      0, 0, 0,
                      r["conteos"]["normal"], 0, r["conteos"]["domingo"], r["conteos"]["feriado"],
                      r["ded_normal_h"], 0, r["ded_domingo_h"], r["ded_feriado_h"],
                      r["monto_normal"], 0, r["monto_domingo"], r["monto_feriado"],
                      r["monto_base"], r["monto_deducciones"], r["total"], notas, fecha_reg, id_cob))

            # Detalle día por día (reemplazo)
            cursor.execute("DELETE FROM cobranza_detalle_dias WHERE id_cobranza=%s", (id_cob,))
            for d in self.dias:
                cursor.execute(
                    "INSERT INTO cobranza_detalle_dias (id_cobranza, fecha, categoria, feriado_nombre) VALUES (%s,%s,%s,%s)",
                    (id_cob, d["fecha"].strftime("%d/%m/%Y"), d["categoria"], d["feriado_nombre"]))

            # Detalle por unidad (solo plan Por Hora) (reemplazo)
            cursor.execute("DELETE FROM cobranza_quincena_unidades WHERE id_cobranza=%s", (id_cob,))
            if not es_viaje:
                for uc in r["unidades"]:
                    cursor.execute('''
                        INSERT INTO cobranza_quincena_unidades
                        (id_cobranza, unidad, precio_normal, precio_domingo, precio_feriado, horas_dia,
                         cant_normal, cant_domingo, cant_feriado,
                         ded_normal_h, ded_domingo_h, ded_feriado_h,
                         monto_normal, monto_domingo, monto_feriado,
                         monto_deducciones, subtotal)
                        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                    ''', (id_cob, uc["unidad"], uc["precio_normal"], uc["precio_domingo"], uc["precio_feriado"],
                          uc["horas_dia"],
                          r["conteos"]["normal"], r["conteos"]["domingo"], r["conteos"]["feriado"],
                          uc["ded_normal_h"], uc["ded_domingo_h"], uc["ded_feriado_h"],
                          uc["monto_normal"], uc["monto_domingo"], uc["monto_feriado"],
                          uc["monto_deducciones"], uc["subtotal"]))

            # Detalle de viajes por distancia (solo plan Por Punto o Viaje) (reemplazo)
            cursor.execute("DELETE FROM cobranza_quincena_viajes WHERE id_cobranza=%s", (id_cob,))
            if es_viaje:
                for vj in r["viajes"]:
                    cursor.execute('''
                        INSERT INTO cobranza_quincena_viajes
                        (id_cobranza, distancia_desde, distancia_hasta, precio_normal, precio_domingo, precio_feriado,
                         viajes_normal, viajes_domingo, viajes_feriado,
                         ded_normal, ded_domingo, ded_feriado,
                         monto_base, monto_deducciones, subtotal)
                        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                    ''', (id_cob,
                          vj.get("desde", 0), vj.get("hasta", 0),
                          vj["precio_normal"], vj["precio_domingo"], vj["precio_feriado"],
                          vj["viajes_normal"], vj["viajes_domingo"], vj["viajes_feriado"],
                          vj["ded_normal"], vj["ded_domingo"], vj["ded_feriado"],
                          vj["monto_base"], vj["monto_deducciones"], vj["subtotal"]))

                # Detalle de viajes individuales (fecha, vehículo, tipo) — reemplazo
                cursor.execute("DELETE FROM cobranza_viajes_detalle WHERE id_cobranza=%s", (id_cob,))
                for t in self.viajes_registrados:
                    cursor.execute('''
                        INSERT INTO cobranza_viajes_detalle
                        (id_cobranza, fecha, vehiculo, categoria, distancia_desde, distancia_hasta, precio, monto)
                        VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
                    ''', (id_cob, t["fecha"].strftime("%d/%m/%Y"), t.get("vehiculo", ""),
                          t["categoria"], t.get("desde", 0), t.get("hasta", 0),
                          t.get("precio", 0), t.get("monto", 0)))

            # Asientos de deducción con fecha (unidad o rango de distancia) — reemplazo
            cursor.execute("DELETE FROM cobranza_deducciones_detalle WHERE id_cobranza=%s", (id_cob,))
            for d in r.get("deducciones_detalle", []):
                fecha_d = d.get("fecha")
                cursor.execute('''
                    INSERT INTO cobranza_deducciones_detalle
                    (id_cobranza, plan, unidad, fecha, categoria, horas, minutos, cantidad,
                     distancia_desde, distancia_hasta, precio, motivo, monto, tipo)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                ''', (id_cob, str(d.get("plan") or self.plan_cobro), str(d.get("unidad") or "")[:150],
                      fecha_d.strftime("%d/%m/%Y") if isinstance(fecha_d, date) else None,
                      d.get("categoria") or "normal", d.get("horas") or 0, d.get("minutos") or 0,
                      d.get("cantidad") or 0, d.get("desde") or 0, d.get("hasta") or 0,
                      d.get("precio") or 0, str(d.get("motivo") or "")[:200], d.get("monto") or 0,
                      d.get("tipo") or "DEDUCCION"))

            # Actualiza los precios/horas usados en la configuración de la unidad
            # (solo plan Por Hora; así se recuerdan los últimos precios al elegir al cliente).
            if es_nuevo and not es_viaje:
                for i, u in enumerate(self.unidades):
                    pn = r["unidades"][i]["precio_normal"]
                    pd = r["unidades"][i]["precio_domingo"]
                    pf = r["unidades"][i]["precio_feriado"]
                    hd = self._leer_float(u["horas_dia"], "Horas al día")
                    if u.get("id"):
                        cursor.execute('''
                            UPDATE clientes_unidades
                            SET precio_normal=%s, precio_domingo=%s, precio_feriado=%s, horas_dia=%s
                            WHERE id=%s
                        ''', (pn, pd, pf, hd, u["id"]))
                    else:
                        cursor.execute('''
                            INSERT INTO clientes_unidades
                            (id_cliente, id_vehiculo, unidad, precio_normal, precio_domingo, precio_feriado, horas_dia)
                            VALUES (%s,%s,%s,%s,%s,%s,%s)
                        ''', (self.cliente_id, u.get("id_vehiculo"), u["unidad"], pn, pd, pf, hd))

            conn.commit()
            self.registro_editando = id_cob
            registrar_auditoria(self.usuario_activo, "Cobranza",
                                f"Guardó cálculo de cobranza {mes:02d}/{anio} Q{quincena} de {self.cliente_nombre} (total {formatear_moneda(r['total'])})")
        except Exception as e:
            conn.rollback()
            messagebox.showerror("Error SQL", f"No se pudo guardar la quincena:\n{e}")
            return False
        finally:
            liberar_conexion(conn)

        # Mensaje flash
        try:
            self.frame_flash.place(relx=0.5, rely=0.95, anchor="s")
            self.frame_flash.lift()
            self.parent.update_idletasks()
            self.parent.after(1800, lambda: self.frame_flash.place_forget())
        except Exception:
            pass
        return True

    # ---------- PDF ----------
    def generar_pdf(self):
        if not REPORTLAB_DISPONIBLE:
            messagebox.showerror("Falta ReportLab", "No se pudo importar reportlab.\nInstale las dependencias con: python instalar_dependencias.py")
            return
        if self.registro_editando is None:
            if not messagebox.askyesno("Guardar primero",
                                       "Para generar el PDF primero se debe guardar la quincena.\n\n¿Desea guardarla ahora?"):
                return
            if not self.guardar_quincena():
                return
        try:
            ruta = self._generar_pdf_desde_registro(self.registro_editando)
        except Exception as e:
            messagebox.showerror("Error al generar el PDF", f"Ocurrió un error inesperado:\n{e}")
            return
        if ruta:
            abrir_documento(ruta)

    def _generar_pdf_desde_registro(self, id_cob):
        conn = conectar_db(silencioso=True)
        if not conn:
            messagebox.showerror("Sin Conexión", "No se pudo conectar a la base de datos para generar el PDF.")
            return None
        try:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, id_cliente, cliente_nombre, cliente_ruc, anio, mes, quincena, plan_cobro,
                       cant_lunvie, cant_sabado, cant_domingo, cant_feriado,
                       monto_base, monto_deducciones, total, notas, fecha_registro
                FROM cobranza_quincenas WHERE id=%s
            ''', (id_cob,))
            row = cursor.fetchone()
            if not row:
                messagebox.showerror("No Encontrado", "El registro de cobranza no existe.")
                return None
            (id_rec, id_cli, nom_cli, ruc_cli, anio, mes, quincena, plan,
             c_lunvie, c_sab, c_dom, c_fer,
             m_base, m_ded, total, notas, fecha_reg) = row

            cursor.execute("SELECT fecha, categoria, feriado_nombre FROM cobranza_detalle_dias WHERE id_cobranza=%s ORDER BY fecha", (id_cob,))
            detalle = cursor.fetchall()

            cursor.execute('''
                SELECT unidad, precio_normal, precio_domingo, precio_feriado, horas_dia,
                       cant_normal, cant_domingo, cant_feriado,
                       ded_normal_h, ded_domingo_h, ded_feriado_h,
                       monto_normal, monto_domingo, monto_feriado,
                       monto_deducciones, subtotal
                FROM cobranza_quincena_unidades WHERE id_cobranza=%s ORDER BY id
            ''', (id_cob,))
            unidades = cursor.fetchall()

            cursor.execute('''
                SELECT distancia_desde, distancia_hasta, precio_normal, precio_domingo, precio_feriado,
                       viajes_normal, viajes_domingo, viajes_feriado,
                       ded_normal, ded_domingo, ded_feriado,
                       monto_base, monto_deducciones, subtotal
                FROM cobranza_quincena_viajes WHERE id_cobranza=%s ORDER BY id
            ''', (id_cob,))
            viajes_pdf = cursor.fetchall()

            # Asientos con fecha (deducciones y horas extras) — detalle uno por uno
            try:
                cursor.execute('''
                    SELECT plan, unidad, fecha, categoria, horas, minutos, cantidad,
                           distancia_desde, distancia_hasta, precio, motivo, monto, tipo
                    FROM cobranza_deducciones_detalle WHERE id_cobranza=%s ORDER BY id
                ''', (id_cob,))
                asientos_pdf = cursor.fetchall()
            except Exception:
                try:
                    # Base de datos anterior sin la columna 'tipo'
                    cursor.execute('''
                        SELECT plan, unidad, fecha, categoria, horas, minutos, cantidad,
                               distancia_desde, distancia_hasta, precio, motivo, monto
                        FROM cobranza_deducciones_detalle WHERE id_cobranza=%s ORDER BY id
                    ''', (id_cob,))
                    asientos_pdf = [tuple(r) + ("DEDUCCION",) for r in cursor.fetchall()]
                except Exception:
                    asientos_pdf = []

            # Datos del cliente (actualizados)
            cliente_info = {"comercial": "", "contacto": "", "telefono": "", "direccion": ""}
            try:
                cursor.execute('''
                    SELECT COALESCE(razon_comercial,''), COALESCE(persona_contacto,''),
                           COALESCE(telefono,''), COALESCE(direccion_fiscal,'')
                    FROM clientes WHERE id=%s
                ''', (id_cli,))
                c_i = cursor.fetchone()
                if c_i:
                    cliente_info = {"comercial": c_i[0], "contacto": c_i[1], "telefono": c_i[2], "direccion": c_i[3]}
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el registro:\n{e}")
            return None
        finally:
            liberar_conexion(conn)

        # Deducciones y horas extras por separado (cada una con su cuadro en el PDF)
        def _es_extra(fila):
            try:
                return str(fila[12] or "DEDUCCION").upper() == "EXTRA"
            except Exception:
                return False

        deducciones_pdf = [f for f in asientos_pdf if not _es_extra(f)]
        extras_pdf = [f for f in asientos_pdf if _es_extra(f)]
        # Horas y monto de horas extras por unidad (para el cuadro de cobro)
        extras_por_unidad_pdf = {}
        for fila in extras_pdf:
            nombre_u = str(fila[1] or "")
            horas_u, monto_u = extras_por_unidad_pdf.get(nombre_u, (0.0, 0.0))
            extras_por_unidad_pdf[nombre_u] = (
                horas_u + horas_de_asiento({"horas": fila[4], "minutos": fila[5]}),
                monto_u + float(fila[11] or 0))
        m_extras_pdf = sum(v[1] for v in extras_por_unidad_pdf.values())

        config = _cargar_config_local()
        ruc_empresa = config.get("ruc_empresa", "")
        razon_empresa = config.get("razon_social_empresa", "")
        simbolo = config.get("simbolo_moneda", "S/.")

        # Carpeta de destino: la configurada en 'Configuración del Sistema'
        # (ruta_drive + cobranzas_generadas), con respaldo local si no existe.
        carpeta_pdf = _carpeta_pdf()
        if not carpeta_pdf:
            try:
                from politica_almacenamiento import advertir
                advertir()
            except Exception:
                messagebox.showwarning("Almacenamiento bloqueado", "Este equipo no está autorizado a guardar archivos.")
            return
        # Si el PDF del mismo periodo ya está abierto en el visor (bloqueado),
        # se guarda una copia nueva con marca de tiempo para poder abrirla.
        nombre_base = os.path.join(
            carpeta_pdf,
            f"Cobranza_{ruc_cli}_{anio:04d}-{mes:02d}_Q{quincena}.pdf")
        if _archivo_bloqueado(nombre_base):
            nombre_archivo = os.path.join(
                carpeta_pdf,
                f"Cobranza_{ruc_cli}_{anio:04d}-{mes:02d}_Q{quincena}_{datetime.now().strftime('%H%M%S')}.pdf")
            copia_generada = True
        else:
            nombre_archivo = nombre_base
            copia_generada = False

        # Red de seguridad: asegurar que la carpeta del PDF existe antes de guardar
        try:
            os.makedirs(os.path.dirname(nombre_archivo), exist_ok=True)
        except Exception:
            pass

        c = canvas.Canvas(nombre_archivo, pagesize=letter)
        ancho_pag = 612
        y = 760.0

        # ---- Encabezado con logo ----
        ruta_logo = _obtener_logo_pdf()
        if ruta_logo and os.path.exists(ruta_logo):
            try:
                img = ImageReader(ruta_logo)
                w_orig, h_orig = img.getSize()
                alto = 70.0
                ancho = alto * (float(w_orig) / float(h_orig))
                c.drawImage(img, 40, y - alto, width=min(ancho, 160), height=alto)
                c.setFont("Helvetica-Bold", 15)
                c.setFillColorRGB(0.12, 0.32, 0.55)
                c.drawString(220, y - 25, "CÁLCULO DE COBRANZA QUINCENAL")
            except Exception:
                c.setFont("Helvetica-Bold", 15)
                c.setFillColorRGB(0.12, 0.32, 0.55)
                c.drawString(40, y - 25, "CÁLCULO DE COBRANZA QUINCENAL")
        else:
            c.setFont("Helvetica-Bold", 15)
            c.setFillColorRGB(0.12, 0.32, 0.55)
            c.drawString(40, y - 25, "CÁLCULO DE COBRANZA QUINCENAL")

        c.setFont("Helvetica-Bold", 10)
        c.setFillColorRGB(0, 0, 0)
        if razon_empresa:
            c.drawString(220, y - 42, f"{razon_empresa}")
        if ruc_empresa:
            c.drawRightString(572, y - 42, f"RUC: {ruc_empresa}")
        c.setFont("Helvetica", 9)
        c.drawRightString(572, y - 57, f"Emisión: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        c.drawRightString(572, y - 69, f"N° Registro: {id_rec:04d}")
        y -= 80.0

        c.setLineWidth(1)
        c.line(40, y, 572, y)
        y -= 18.0

        # ---- Datos del cliente (nada se sale del margen de la hoja) ----
        MARGEN_DER = 572.0
        ANCHO_TOTAL = MARGEN_DER - 40.0     # 532 pt de ancho útil
        ANCHO_IZQ = 250.0
        ANCHO_DER = 265.0

        c.setFont("Helvetica-Bold", 11)
        c.drawString(40, y, "CLIENTE")
        c.setFont("Helvetica", 10)
        y_cli = y - 14.0

        # Razón Social: ancho completo, hasta 2 líneas
        for linea in _envolver_texto(c, f"Razón Social: {nom_cli}", ANCHO_TOTAL, max_lineas=2):
            c.drawString(40, y_cli, linea)
            y_cli -= 13.0

        # Razón Comercial: ancho completo, 1 línea
        if cliente_info["comercial"]:
            c.drawString(40, y_cli,
                         _envolver_texto(c, f"Razón Comercial: {cliente_info['comercial']}",
                                         ANCHO_TOTAL, max_lineas=1)[0])
            y_cli -= 13.0

        # Contacto (izquierda) y Teléfono (derecha) en la misma fila
        fila_usada = False
        if cliente_info["contacto"]:
            c.drawString(40, y_cli, _envolver_texto(
                c, f"Contacto: {cliente_info['contacto']}", ANCHO_IZQ, max_lineas=1)[0])
            fila_usada = True
        if cliente_info["telefono"]:
            c.drawRightString(MARGEN_DER, y_cli, _envolver_texto(
                c, f"Teléfono: {cliente_info['telefono']}", ANCHO_DER, max_lineas=1)[0])
            fila_usada = True
        if fila_usada:
            y_cli -= 13.0

        # RUC (izquierda) y DIRECCIÓN (derecha, hasta 2 líneas)
        c.drawString(40, y_cli, f"RUC: {ruc_cli}")
        lineas_dire = []
        if cliente_info["direccion"]:
            lineas_dire = _envolver_texto(c, f"Dirección: {cliente_info['direccion']}",
                                          ANCHO_DER, max_lineas=2)
            c.drawRightString(MARGEN_DER, y_cli, lineas_dire[0])
            if len(lineas_dire) > 1:
                c.drawRightString(MARGEN_DER, y_cli - 13.0, lineas_dire[1])
        y_cli -= 13.0 + (13.0 if len(lineas_dire) > 1 else 0.0)

        # ---- Periodo (debajo del bloque del cliente) ----
        y_cli -= 4.0
        c.setFont("Helvetica-Bold", 11)
        c.drawString(40, y_cli, "PERIODO")
        c.setFont("Helvetica", 10)
        periodo_txt = f"{NOMBRES_MESES[mes-1]} {anio} — Quincena {'1ª (1 al 15)' if quincena == 1 else '2ª (16 al fin)'}"
        c.drawString(40, y_cli - 13.0, _recortar_texto(c, periodo_txt, ANCHO_TOTAL, tam=10))
        c.drawString(40, y_cli - 26.0, f"Plan de Cobro: {plan}")
        y = y_cli - 40.0

        c.line(40, y, 572, y)
        y -= 16.0

        # ---- Días de la quincena: solo el resumen total ----
        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica-Bold", 11)
        c.drawString(40, y, "DETALLE DE DÍAS DE LA QUINCENA")
        y -= 16.0
        c.setFont("Helvetica-Bold", 10)
        c.drawString(40, y, f"Total días: {len(detalle)}  |  Días Normales (Lun–Sáb): {c_lunvie}  |  Domingos: {c_dom}  |  Feriados: {c_fer}")
        y -= 24.0

        if viajes_pdf:
            # ---- Detalle de viajes por distancia ----
            c.setFont("Helvetica-Bold", 11)
            c.drawString(40, y, "DETALLE DE VIAJES POR DISTANCIA")
            y -= 16.0
            if y < 90:
                c.showPage()
                y = 750.0
            # Tabla A: viajes realizados
            cols_a = [80, 40, 40, 40, 45, 45, 45, 80]
            xs_a = [40]
            for w in cols_a:
                xs_a.append(xs_a[-1] + w)
            ancho_a = sum(cols_a)
            c.setFillColorRGB(0.9, 0.93, 0.97)
            c.rect(40, y - 14, ancho_a, 14, stroke=0, fill=1)
            c.setFillColorRGB(0, 0, 0)
            c.setFont("Helvetica-Bold", 7)
            cab_a = ["DISTANCIA", "P.NORM", "P.DOM", "P.FER", "V.NORM", "V.DOM", "V.FER", "MONTO"]
            for j, (w, txt) in enumerate(zip(cols_a, cab_a)):
                c.drawCentredString(xs_a[j] + w / 2, y - 10, txt)
            y -= 16.0
            c.setFont("Helvetica", 8)
            for (desde, hasta, pn, pd, pf, vn, vd, vf, dn, dd, df, mb2, md2, sub) in viajes_pdf:
                if y < 90:
                    c.showPage()
                    y = 750.0
                c.setFillColorRGB(0, 0, 0)
                c.drawString(xs_a[0] + 3, y - 10, f"{float(desde or 0):g}-{float(hasta or 0):g}")
                c.drawCentredString(xs_a[1] + cols_a[1] / 2, y - 10, f"{float(pn or 0):g}")
                c.drawCentredString(xs_a[2] + cols_a[2] / 2, y - 10, f"{float(pd or 0):g}")
                c.drawCentredString(xs_a[3] + cols_a[3] / 2, y - 10, f"{float(pf or 0):g}")
                c.drawCentredString(xs_a[4] + cols_a[4] / 2, y - 10, str(int(vn or 0)))
                c.drawCentredString(xs_a[5] + cols_a[5] / 2, y - 10, str(int(vd or 0)))
                c.drawCentredString(xs_a[6] + cols_a[6] / 2, y - 10, str(int(vf or 0)))
                c.drawCentredString(xs_a[7] + cols_a[7] / 2, y - 10, f"{simbolo} {float(mb2 or 0):,.2f}")
                c.setStrokeColorRGB(0.85, 0.85, 0.85)
                c.line(40, y - 13, 40 + ancho_a, y - 13)
                y -= 15.0
            c.setStrokeColorRGB(0, 0, 0)
            c.setLineWidth(1)
            c.line(40, y + 1, 40 + ancho_a, y + 1)
            # Tabla B: viajes no realizados (deducción) en la página siguiente
            c.showPage()
            y = 750.0
            cols_b = [120, 60, 60, 60, 100]
            xs_b = [40]
            for w in cols_b:
                xs_b.append(xs_b[-1] + w)
            ancho_b = sum(cols_b)
            c.setFont("Helvetica-Bold", 9)
            c.drawString(40, y, "VIAJES NO REALIZADOS (DEDUCCIÓN)")
            y -= 14.0
            c.setFillColorRGB(0.9, 0.93, 0.97)
            c.rect(40, y - 14, ancho_b, 14, stroke=0, fill=1)
            c.setFillColorRGB(0, 0, 0)
            c.setFont("Helvetica-Bold", 7)
            cab_b = ["DISTANCIA", "NR.NORM", "NR.DOM", "NR.FER", "MONTO"]
            for j, (w, txt) in enumerate(zip(cols_b, cab_b)):
                c.drawCentredString(xs_b[j] + w / 2, y - 10, txt)
            y -= 16.0
            c.setFont("Helvetica", 8)
            for (desde, hasta, pn, pd, pf, vn, vd, vf, dn, dd, df, mb2, md2, sub) in viajes_pdf:
                if y < 90:
                    c.showPage()
                    y = 750.0
                c.setFillColorRGB(0, 0, 0)
                c.drawString(xs_b[0] + 3, y - 10, f"{float(desde or 0):g}-{float(hasta or 0):g}")
                c.drawCentredString(xs_b[1] + cols_b[1] / 2, y - 10, str(int(dn or 0)))
                c.drawCentredString(xs_b[2] + cols_b[2] / 2, y - 10, str(int(dd or 0)))
                c.drawCentredString(xs_b[3] + cols_b[3] / 2, y - 10, str(int(df or 0)))
                c.drawCentredString(xs_b[4] + cols_b[4] / 2, y - 10, f"{simbolo} {float(md2 or 0):,.2f}")
                c.setStrokeColorRGB(0.85, 0.85, 0.85)
                c.line(40, y - 13, 40 + ancho_b, y - 13)
                y -= 15.0
            c.setStrokeColorRGB(0, 0, 0)
            c.setLineWidth(1)
            c.line(40, y + 1, 40 + ancho_b, y + 1)
            y -= 8.0
            c.setFont("Helvetica-Bold", 10)
            c.drawString(42, y - 10, "SUBTOTAL (BASE)")
            c.drawRightString(40 + ancho_a - 2, y - 10, f"{simbolo} {float(m_base or 0):,.2f}")
            y -= 15.0
            c.drawString(42, y - 10, "TOTAL NO REALIZADOS (DEDUCCIÓN)")
            c.drawRightString(40 + ancho_a - 2, y - 10, f"- {simbolo} {float(m_ded or 0):,.2f}")
            y -= 26.0
        else:
            # ---- Cuadro de cobro POR UNIDAD ----
            c.setFont("Helvetica-Bold", 11)
            c.drawString(40, y, "CUADRO DE COBRO POR UNIDAD")
            y -= 16.0
            if y < 90:
                c.showPage()
                y = 750.0
            # Cuadro limpio: solo precios y base por unidad (sin deducciones ni horas extras)
            cols_u = [150, 55, 62, 62, 62, 100]
            xs_u = [40]
            for w in cols_u:
                xs_u.append(xs_u[-1] + w)
            ancho_tabla = sum(cols_u)
            cab_u = ["UNIDAD", "HORAS/DÍA", "P. NORMAL", "P. DOMINGO", "P. FERIADO", "SUBTOTAL"]

            def _cabecera_unidades():
                c.setFillColorRGB(0.9, 0.93, 0.97)
                c.rect(40, y - 14, ancho_tabla, 14, stroke=0, fill=1)
                c.setFillColorRGB(0, 0, 0)
                c.setFont("Helvetica-Bold", 8)
                for j, (w, txt) in enumerate(zip(cols_u, cab_u)):
                    c.drawCentredString(xs_u[j] + w / 2, y - 10, txt)
                c.setFont("Helvetica", 9)

            _cabecera_unidades()
            y -= 16.0
            for (uni, pn, pd, pf, hd, cn, cdom, cfer,
                 dn, dd, df, mn, md, mf, mded, sub) in unidades:
                if y < 90:
                    c.showPage()
                    y = 750.0
                    _cabecera_unidades()
                    y -= 16.0
                # Base de la unidad = días normales + domingos + feriados (sin ajustes)
                base_u = float(mn or 0) + float(md or 0) + float(mf or 0)
                c.setFillColorRGB(0, 0, 0)
                c.drawString(xs_u[0] + 4, y - 10, str(uni)[:36])
                c.drawCentredString(xs_u[1] + cols_u[1] / 2, y - 10, f"{float(hd or 0):g}")
                c.drawCentredString(xs_u[2] + cols_u[2] / 2, y - 10, f"{simbolo} {float(pn or 0):,.2f}")
                c.drawCentredString(xs_u[3] + cols_u[3] / 2, y - 10, f"{simbolo} {float(pd or 0):,.2f}")
                c.drawCentredString(xs_u[4] + cols_u[4] / 2, y - 10, f"{simbolo} {float(pf or 0):,.2f}")
                c.drawCentredString(xs_u[5] + cols_u[5] / 2, y - 10, f"{simbolo} {base_u:,.2f}")
                c.setStrokeColorRGB(0.85, 0.85, 0.85)
                c.line(40, y - 13, 40 + ancho_tabla, y - 13)
                y -= 15.0
            c.setStrokeColorRGB(0, 0, 0)
            c.setLineWidth(1)
            c.line(40, y + 1, 40 + ancho_tabla, y + 1)
            c.setFont("Helvetica-Bold", 10)
            c.drawString(xs_u[0] + 4, y - 10, "SUBTOTAL (BASE)")
            c.drawRightString(40 + ancho_tabla, y - 10, f"{simbolo} {float(m_base or 0):,.2f}")
            y -= 26.0

        # ---- Asientos con fecha: deducciones y horas extras (uno por uno) ----
        es_viaje_pdf = (plan == "Por Punto o Viaje")

        def _tabla_asientos(titulo, filas, etiqueta_total, signo, monto_declarado=None):
            """Dibuja una tabla de asientos con fecha; devuelve la suma de montos."""
            nonlocal y
            if not filas:
                return 0.0
            if y < 200:
                c.showPage()
                y = 750.0
            cols_dd = [62, 150, 68, 48, 120, 84]
            xs_dd = [40]
            for w in cols_dd:
                xs_dd.append(xs_dd[-1] + w)
            ancho_dd = sum(cols_dd)
            cab_dd = ["FECHA", "UNIDAD / DISTANCIA", "TIPO",
                      "CANT." if es_viaje_pdf else "HORAS", "MOTIVO", "MONTO"]

            def _cabecera():
                nonlocal y
                c.setFillColorRGB(0.9, 0.93, 0.97)
                c.rect(40, y - 14, ancho_dd, 14, stroke=0, fill=1)
                c.setFillColorRGB(0, 0, 0)
                c.setFont("Helvetica-Bold", 7.5)
                for j, (w, txt) in enumerate(zip(cols_dd, cab_dd)):
                    c.drawCentredString(xs_dd[j] + w / 2, y - 10, txt)
                y -= 16.0
                c.setFont("Helvetica", 8)

            c.setFillColorRGB(0, 0, 0)
            c.setFont("Helvetica-Bold", 11)
            c.drawString(40, y, titulo)
            y -= 16.0
            _cabecera()
            total_asientos = 0.0
            for (plan_d, unidad_d, fecha_d, cat_d, horas_d, min_d, cant_d,
                 desde_d, hasta_d, precio_d, motivo_d, monto_d, tipo_d) in filas:
                if y < 90:
                    c.showPage()
                    y = 750.0
                    _cabecera()
                try:
                    cant_v = float(cant_d or 0) if es_viaje_pdf else horas_de_asiento(
                        {"horas": horas_d, "minutos": min_d})
                except Exception:
                    cant_v = 0.0
                monto_v = float(monto_d or 0)
                total_asientos += monto_v
                c.setFillColorRGB(0, 0, 0)
                c.drawCentredString(xs_dd[0] + cols_dd[0] / 2, y - 10, str(fecha_d or "—"))
                c.drawString(xs_dd[1] + 3, y - 10, str(unidad_d or "")[:34])
                c.drawCentredString(xs_dd[2] + cols_dd[2] / 2, y - 10,
                                    str(ETIQUETAS_CORTAS.get(cat_d, cat_d) or "")[:14])
                c.drawCentredString(xs_dd[3] + cols_dd[3] / 2, y - 10, f"{cant_v:g}")
                c.drawString(xs_dd[4] + 3, y - 10, str(motivo_d or "")[:30])
                c.drawCentredString(xs_dd[5] + cols_dd[5] / 2, y - 10, f"{simbolo} {monto_v:,.2f}")
                c.setStrokeColorRGB(0.85, 0.85, 0.85)
                c.line(40, y - 13, 40 + ancho_dd, y - 13)
                y -= 15.0
            c.setStrokeColorRGB(0, 0, 0)
            c.setLineWidth(1)
            c.line(40, y + 1, 40 + ancho_dd, y + 1)
            c.setFont("Helvetica-Bold", 10)
            c.drawString(42, y - 12, etiqueta_total)
            c.drawRightString(40 + ancho_dd, y - 12, f"{signo}{simbolo} {total_asientos:,.2f}")
            y -= 16.0
            if monto_declarado is not None:
                diferencia_ded = float(monto_declarado) - total_asientos
                if abs(diferencia_ded) > 0.009:
                    c.setFont("Helvetica", 8)
                    c.setFillColorRGB(0.4, 0.4, 0.4)
                    c.drawString(42, y - 10, "Además hay deducciones registradas sin asiento detallado: "
                                             f"{simbolo} {diferencia_ded:,.2f}")
                    c.setFillColorRGB(0, 0, 0)
                    y -= 12.0
            y -= 20.0
            return total_asientos

        _tabla_asientos("DETALLE DE DEDUCCIONES REGISTRADAS (ASIENTOS CON FECHA)",
                        deducciones_pdf, "TOTAL ASIENTOS DE DEDUCCIÓN", "- ", float(m_ded or 0))
        _tabla_asientos("DETALLE DE HORAS EXTRAS (ASIENTOS CON FECHA)",
                        extras_pdf, "TOTAL HORAS EXTRAS", "+ ")

        # ---- Total ----
        c.setFillColorRGB(0.9, 0.97, 0.93)
        c.rect(40, y - 20, 420, 20, stroke=0, fill=1)
        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica-Bold", 13)
        c.drawString(48, y - 14, "TOTAL A COBRAR")
        c.drawRightString(460, y - 14, f"{simbolo} {float(total or 0):,.2f}")
        y -= 34.0

        if notas:
            c.setFont("Helvetica-Bold", 10)
            c.drawString(40, y, "NOTAS:")
            c.setFont("Helvetica", 9)
            for linea in str(notas).split("\n"):
                c.drawString(80, y, linea[:95])
                y -= 12.0
            y -= 10.0

        # (Sin firmas: el documento ya no lleva líneas de firma)

        c.setFont("Helvetica", 8)
        c.setFillColorRGB(0.4, 0.4, 0.4)
        c.drawCentredString(306, 40, _recortar_texto(
            c, f"Documento generado por el Sistema de Control de Flota Automotriz — {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            520, "Helvetica", 8))
        c.drawCentredString(306, 30, _recortar_texto(
            c, f"Registro N° {id_rec:04d} — {periodo_txt} — Cliente: {nom_cli}", 520, "Helvetica", 8))

        try:
            c.save()
        except PermissionError:
            messagebox.showerror(
                "No se pudo guardar el PDF",
                f"No se pudo escribir el archivo PDF.\n\n"
                f"Puede estar abierto en el visor o la carpeta no permite escritura.\n"
                f"Ciérralo en el visor y verifica la carpeta, e inténtalo de nuevo.\n\n"
                f"Archivo: {nombre_archivo}")
            return None
        except Exception as e:
            messagebox.showerror("Error al generar el PDF", f"Ocurrió un error inesperado:\n{e}")
            return None

        if copia_generada:
            messagebox.showinfo(
                "PDF generado",
                f"El PDF anterior estaba abierto en el visor, por lo que se guardó una copia nueva:\n\n{os.path.basename(nombre_archivo)}")

        # Actualizar ruta del PDF en el registro
        conn2 = conectar_db(silencioso=True)
        if conn2:
            try:
                cur2 = conn2.cursor()
                cur2.execute("UPDATE cobranza_quincenas SET pdf_ruta=%s WHERE id=%s", (ruta_para_guardar(nombre_archivo), id_cob))
                conn2.commit()
                cur2.close()
            except Exception:
                pass
            finally:
                liberar_conexion(conn2)
        return nombre_archivo

    # ---------- REGISTROS ----------
    def abrir_registros(self):
        VentanaRegistrosCobranza(self.parent, self)

    def cargar_registro(self, id_cob):
        """Carga un registro guardado de vuelta en el formulario (editar)."""
        conn = conectar_db(silencioso=True)
        if not conn:
            messagebox.showerror("Sin Conexión", "No se pudo conectar a la base de datos.")
            return
        try:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, id_cliente, cliente_nombre, cliente_ruc, anio, mes, quincena, plan_cobro,
                       precio_normal, precio_domingo, precio_feriado,
                       cant_lunvie, cant_sabado, cant_domingo, cant_feriado,
                       ded_normal_h, ded_sabado_h, ded_domingo_h, ded_feriado_h,
                       monto_lunvie, monto_domingo, monto_feriado,
                       notas
                FROM cobranza_quincenas WHERE id=%s
            ''', (id_cob,))
            row = cursor.fetchone()
            if not row:
                messagebox.showerror("No Encontrado", "El registro no existe.")
                return
            cursor.execute("SELECT fecha, categoria, feriado_nombre FROM cobranza_detalle_dias WHERE id_cobranza=%s ORDER BY fecha", (id_cob,))
            detalle = cursor.fetchall()
            cursor.execute('''
                SELECT unidad, precio_normal, precio_domingo, precio_feriado, horas_dia,
                       ded_normal_h, ded_domingo_h, ded_feriado_h
                FROM cobranza_quincena_unidades WHERE id_cobranza=%s ORDER BY id
            ''', (id_cob,))
            unidades_rows = cursor.fetchall()
            cursor.execute('''
                SELECT distancia_desde, distancia_hasta, precio_normal, precio_domingo, precio_feriado,
                       viajes_normal, viajes_domingo, viajes_feriado,
                       ded_normal, ded_domingo, ded_feriado
                FROM cobranza_quincena_viajes WHERE id_cobranza=%s ORDER BY id
            ''', (id_cob,))
            viajes_rows = cursor.fetchall()
            cursor.execute('''
                SELECT fecha, vehiculo, categoria, distancia_desde, distancia_hasta, precio, monto
                FROM cobranza_viajes_detalle WHERE id_cobranza=%s ORDER BY id
            ''', (id_cob,))
            viajes_detalle_rows = cursor.fetchall()
            # Asientos de deducción con fecha (horas/minutos por unidad o cantidad por rango)
            try:
                cursor.execute('''
                    SELECT plan, unidad, fecha, categoria, horas, minutos, cantidad,
                           distancia_desde, distancia_hasta, precio, motivo, monto, tipo
                    FROM cobranza_deducciones_detalle WHERE id_cobranza=%s ORDER BY id
                ''', (id_cob,))
                deducciones_rows = cursor.fetchall()
            except Exception:
                deducciones_rows = []
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el registro:\n{e}")
            return
        finally:
            liberar_conexion(conn)

        (id_rec, id_cli, nom_cli, ruc_cli, anio, mes, quincena, plan,
         p_normal, p_domingo, p_feriado, c_lunvie, c_sab, c_dom, c_fer,
         d_normal, d_sab, d_dom, d_fer,
         m_lunvie, m_domingo, m_feriado, notas) = row

        # Cliente
        self.cliente_id = id_cli
        self.cliente_nombre = nom_cli
        self.cliente_ruc = ruc_cli
        self.plan_cobro = plan if plan in PLANES else "Por Hora"
        try:
            seleccion = f"{ruc_cli} | {nom_cli}"
            if seleccion in self.combo_cliente.cget("values"):
                self.combo_cliente.set(seleccion)
            else:
                self.combo_cliente.configure(values=list(self.combo_cliente.cget("values")) + [seleccion])
                self.combo_cliente.set(seleccion)
            self.on_cliente_seleccionado(seleccion)
        except Exception:
            pass
        # Asegurar etiquetas del plan aunque el cliente no esté en la lista local
        self.lbl_plan.configure(text=self.plan_cobro)
        self.lbl_cliente_info.configure(text=f"{self.cliente_ruc} | {self.cliente_nombre}")
        self.aplicar_plan(self.plan_cobro)

        # Periodo
        try:
            self.combo_mes.set(f"{mes:02d} - {NOMBRES_MESES[mes-1]}")
            self.combo_anio.set(str(anio))
            self.combo_quincena.set("Primera (1 al 15)" if quincena == 1 else "Segunda (16 al fin)")
        except Exception:
            pass

        # Días desde el detalle guardado
        self.dias = []
        for fecha_s, categoria, fer_nombre in detalle:
            try:
                dt = datetime.strptime(fecha_s, "%d/%m/%Y").date()
            except Exception:
                continue
            self.dias.append({"fecha": dt, "categoria": categoria if categoria in dict(CATEGORIAS) else "normal",
                              "feriado_nombre": fer_nombre or ""})
        for item in self.tabla_dias.get_children():
            self.tabla_dias.delete(item)
        for d in self.dias:
            etiqueta_cat = dict(CATEGORIAS)[d["categoria"]]
            nombre = d["feriado_nombre"] if d["categoria"] == "feriado" else ""
            self.tabla_dias.insert("", tk.END, values=(
                d["fecha"].strftime("%d/%m/%Y"),
                DIAS_SEMANA[d["fecha"].weekday()],
                etiqueta_cat,
                nombre,
            ))

        # Detalle según el plan: unidades (Por Hora) o viajes por distancia (Por Punto/Viaje)
        self.unidades = []
        self.rangos_viaje = []
        self.viajes_registrados = []
        self.deducciones_registradas = []
        if self.plan_cobro == "Por Punto o Viaje":
            for (desde, hasta, pn, pd, pf, vn, vd, vf, dn, dd, df) in viajes_rows:
                self.rangos_viaje.append({
                    "id": None, "distancia_desde": float(desde or 0), "distancia_hasta": float(hasta or 0),
                    "precio_normal": float(pn or 0), "precio_domingo": float(pd or 0),
                    "precio_feriado": float(pf or 0),
                    "viajes": {"normal": int(vn or 0), "domingo": int(vd or 0), "feriado": int(vf or 0)},
                    "ded": {"normal": int(dn or 0), "domingo": int(dd or 0), "feriado": int(df or 0)},
                })
            # Recupera los asientos de deducción con fecha (cantidad no realizada por rango)
            for (plan_d, unidad_d, fecha_s, cat_d, horas_d, min_d, cant_d,
                 desde_d, hasta_d, precio_d, motivo_d, monto_d, tipo_d) in deducciones_rows:
                if str(tipo_d or "DEDUCCION").upper() == "EXTRA":
                    continue  # el plan por viajes no usa horas extras
                idx = next((i for i, rg in enumerate(self.rangos_viaje)
                            if abs(rg["distancia_desde"] - float(desde_d or 0)) < 0.001
                            and abs(rg["distancia_hasta"] - float(hasta_d or 0)) < 0.001), None)
                if idx is None:
                    continue
                try:
                    fecha_dt = datetime.strptime(str(fecha_s), "%d/%m/%Y").date() if fecha_s else None
                except Exception:
                    fecha_dt = None
                cat = cat_d if cat_d in dict(CATEGORIAS) else "normal"
                cant = int(float(cant_d or 0))
                asiento = nuevo_asiento_deduccion(fecha_dt, cat, 0.0, 0.0, cant, motivo_d,
                                                  desde_d, hasta_d, precio_d)
                # Las columnas del registro guardan cuadro + asientos: se descuentan aquí
                # porque los asientos se vuelven a sumar al recalcular.
                self.rangos_viaje[idx]["ded"][cat] = max(
                    0, int(self.rangos_viaje[idx]["ded"].get(cat, 0) or 0) - cant)
                self.deducciones_registradas.append({
                    "fecha": fecha_dt, "categoria": cat, "rango_idx": idx,
                    "cantidad": cant, "motivo": str(motivo_d or ""),
                    "precio": float(precio_d or 0), "monto": float(monto_d or 0),
                    "desde": float(desde_d or 0), "hasta": float(hasta_d or 0),
                })

            # Recupera los viajes individuales registrados (fecha, vehículo, tipo)
            for (fecha_s, vehiculo, categoria, desde, hasta, precio, monto) in viajes_detalle_rows:
                rango_idx = next((i for i, rg in enumerate(self.rangos_viaje)
                                  if abs(rg["distancia_desde"] - float(desde or 0)) < 0.001
                                  and abs(rg["distancia_hasta"] - float(hasta or 0)) < 0.001), None)
                if rango_idx is None:
                    continue
                try:
                    dt = datetime.strptime(fecha_s, "%d/%m/%Y").date()
                except Exception:
                    continue
                self.viajes_registrados.append({
                    "fecha": dt, "categoria": categoria if categoria in dict(CATEGORIAS) else "normal",
                    "rango_idx": rango_idx, "vehiculo": str(vehiculo or ""),
                    "precio": float(precio or 0), "monto": float(monto or 0),
                    "desde": float(desde or 0), "hasta": float(hasta or 0),
                })
            self.pintar_viajes()
            self._pintar_viajes_registrados()
            self._pintar_deducciones_registradas()
        else:
            if unidades_rows:
                for (uni, pn, pd, pf, hd, dn, dd, df) in unidades_rows:
                    self.unidades.append({
                        "id": None, "id_vehiculo": None, "unidad": str(uni or ""),
                        "precio_normal": float(pn or 0), "precio_domingo": float(pd or 0),
                        "precio_feriado": float(pf or 0), "horas_dia": float(hd or 0),
                        "ded": {"normal": [float(dn or 0), 0.0], "domingo": [float(dd or 0), 0.0],
                                "feriado": [float(df or 0), 0.0]},
                    })
            else:
                # Registro sin detalle por unidad: usar las unidades actuales configuradas
                # del cliente (vehículos, horas y precios) para poder visualizarlo/recargarlo.
                self.cargar_unidades_cliente()
                if not self.unidades:
                    # Respaldo: una unidad "Cliente" reconstruida con los montos del registro
                    # (precio efectivo diario = monto ÷ días de la categoría; horas/día 1 para
                    #  que el cálculo actual reproduzca el mismo total).
                    def _precio_efectivo(monto, cant):
                        try:
                            return float(monto or 0) / float(cant) if float(cant or 0) > 0 else 0.0
                        except Exception:
                            return 0.0
                    self.unidades.append({
                        "id": None, "id_vehiculo": None, "unidad": "Cliente",
                        "precio_normal": _precio_efectivo(m_lunvie, c_lunvie),
                        "precio_domingo": _precio_efectivo(m_domingo, c_dom),
                        "precio_feriado": _precio_efectivo(m_feriado, c_fer),
                        "horas_dia": 1.0,
                        "ded": {"normal": [float(d_normal or 0), 0.0], "domingo": [float(d_dom or 0), 0.0],
                                "feriado": [float(d_fer or 0), 0.0]},
                    })
            # Asientos con fecha de cada unidad: deducciones y horas extras
            ded_por_unidad = {}
            extras_por_unidad = {}
            for (plan_d, unidad_d, fecha_s, cat_d, horas_d, min_d, cant_d,
                 desde_d, hasta_d, precio_d, motivo_d, monto_d, tipo_d) in deducciones_rows:
                try:
                    fecha_dt = datetime.strptime(str(fecha_s), "%d/%m/%Y").date() if fecha_s else None
                except Exception:
                    fecha_dt = None
                asiento = nuevo_asiento_deduccion(fecha_dt, cat_d, horas_d, min_d, 0.0, motivo_d)
                if str(tipo_d or "DEDUCCION").upper() == "EXTRA":
                    extras_por_unidad.setdefault(str(unidad_d or ""), []).append(asiento)
                else:
                    ded_por_unidad.setdefault(str(unidad_d or ""), []).append(asiento)
            for u in self.unidades:
                nombre_u = str(u.get("unidad") or "")
                u.setdefault("ded_asientos", [])
                u.setdefault("extras_asientos", [])
                asientos_u = ded_por_unidad.get(nombre_u)
                if asientos_u:
                    u["ded_asientos"] = asientos_u
                    # Las horas guardadas ya incluyen los asientos: se recalculan desde
                    # el detalle para no contarlas dos veces.
                    u["ded"] = ded_agrupada_desde_asientos(asientos_u)
                if extras_por_unidad.get(nombre_u):
                    u["extras_asientos"] = extras_por_unidad[nombre_u]
            self.pintar_unidades()

        self.ent_notas.delete(0, tk.END)
        self.ent_notas.insert(0, str(notas or ""))

        self.registro_editando = id_rec
        self.recalcular()

    # ---------- LIMPIAR ----------
    def limpiar_formulario(self):
        self.cliente_id = None
        self.registro_editando = None
        self.dias = []
        self.feriados = {}
        self.unidades = []
        self.rangos_viaje = []
        self.viajes_registrados = []
        self.deducciones_registradas = []
        try:
            self.combo_cliente.set("— Seleccione un cliente —")
        except Exception:
            pass
        self.lbl_plan.configure(text="Por Hora")
        self.lbl_cliente_info.configure(text="")
        self.lbl_conteo.configure(text="")
        self.lbl_origen_feriados.configure(text="")
        self.lbl_sin_unidades.configure(text="")
        self.lbl_sin_viajes.configure(text="")
        self.lbl_dia_detectado.configure(text="")
        for item in self.tabla_dias.get_children():
            self.tabla_dias.delete(item)
        for item in self.tabla_unidades.get_children():
            self.tabla_unidades.delete(item)
        for item in self.tabla_viajes_reg.get_children():
            self.tabla_viajes_reg.delete(item)
        for item in self.tabla_ded_reg.get_children():
            self.tabla_ded_reg.delete(item)
        for w in self.frame_viajes_grid.winfo_children():
            w.destroy()
        self.ent_viajes = []
        self.ent_fecha_viaje.delete(0, tk.END)
        self.ent_fecha_ded.delete(0, tk.END)
        self.ent_cant_ded.delete(0, tk.END)
        self.ent_motivo_ded.delete(0, tk.END)
        self.lbl_dia_detectado_ded.configure(text="")
        self.lbl_total_ded.configure(text="TOTAL DEDUCCIONES REGISTRADAS: S/ 0,00")
        try:
            self.combo_vehiculo_viaje.set("— Seleccione vehículo —")
            self.lbl_total_viajes.configure(text="TOTAL VIAJES: S/ 0,00")
        except Exception:
            pass
        self.ent_notas.delete(0, tk.END)
        self.aplicar_plan("Por Hora")
        self.recalcular()


# =========================================================
# 🚀 DIÁLOGO: EDITAR UNIDAD (precios, horas/día y deducciones)
# =========================================================
class DialogoUnidad(ctk.CTkToplevel):
    """Permite editar precios, horas/día y (opcionalmente) deducciones de una unidad."""

    def __init__(self, parent, unidad, plan, con_deducciones=True, app=None):
        super().__init__(parent)
        self.result = None
        self.plan = plan
        self.con_deducciones = con_deducciones
        self.app = app          # aplicación principal (días de la quincena → tipo de día)
        self.filas_ded = []     # asientos de deducción (una fila por asiento)
        self.filas_extras = []  # asientos de horas extras (una fila por asiento)
        familia_fuente = "Helvetica" if sys.platform == "darwin" else "Arial"
        self.familia_fuente = familia_fuente

        self.title("Editar Unidad" if con_deducciones else "Configurar Unidad")
        ancho = 780 if con_deducciones else 440
        alto = 760 if con_deducciones else 380
        self.geometry(f"{ancho}x{alto}")
        self.resizable(False, False)
        self.transient(parent)
        try:
            self.grab_set()
        except Exception:
            pass
        self.update_idletasks()
        try:
            x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (ancho // 2)
            y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (alto // 2)
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

        scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(scroll, text=f"🚗 {unidad['unidad']}", font=(familia_fuente, 15, "bold"),
                     text_color="#1f538d").pack(anchor="w", pady=(0, 10))

        # ---- Datos generales ----
        f_datos = ctk.CTkFrame(scroll, corner_radius=10, border_width=1, border_color="#e0e0e0")
        f_datos.pack(fill="x", pady=5, ipady=6)
        f_datos.columnconfigure(1, weight=1)

        ctk.CTkLabel(f_datos, text="Horas al día:", font=(familia_fuente, 12, "bold")).grid(
            row=0, column=0, sticky="w", padx=(15, 5), pady=8)
        self.ent_horas = ctk.CTkEntry(f_datos, width=140, placeholder_text="8")
        self.ent_horas.grid(row=0, column=1, sticky="w", padx=5, pady=8)
        self.ent_horas.insert(0, f"{float(unidad.get('horas_dia') or 0):g}")

        ctk.CTkLabel(f_datos, text="Precio Día Normal (Lun–Sáb):", font=(familia_fuente, 12, "bold")).grid(
            row=1, column=0, sticky="w", padx=(15, 5), pady=8)
        self.ent_pn = ctk.CTkEntry(f_datos, width=140, placeholder_text="0.00")
        self.ent_pn.grid(row=1, column=1, sticky="w", padx=5, pady=8)
        self.ent_pn.insert(0, f"{float(unidad.get('precio_normal') or 0):g}")

        ctk.CTkLabel(f_datos, text="Precio Domingo:", font=(familia_fuente, 12, "bold")).grid(
            row=2, column=0, sticky="w", padx=(15, 5), pady=8)
        self.ent_pd = ctk.CTkEntry(f_datos, width=140, placeholder_text="0.00")
        self.ent_pd.grid(row=2, column=1, sticky="w", padx=5, pady=8)
        self.ent_pd.insert(0, f"{float(unidad.get('precio_domingo') or 0):g}")

        ctk.CTkLabel(f_datos, text="Precio Feriado:", font=(familia_fuente, 12, "bold")).grid(
            row=3, column=0, sticky="w", padx=(15, 5), pady=8)
        self.ent_pf = ctk.CTkEntry(f_datos, width=140, placeholder_text="0.00")
        self.ent_pf.grid(row=3, column=1, sticky="w", padx=5, pady=8)
        self.ent_pf.insert(0, f"{float(unidad.get('precio_feriado') or 0):g}")

        if plan == "Por Hora":
            hint = "Cobro diario = precio por hora × horas del día. Deducción: horas ausentes × precio por hora."
        else:
            hint = "Deducción: viajes / puntos no realizados × precio por viaje."
        ctk.CTkLabel(f_datos, text=hint, font=(familia_fuente, 10), text_color="gray").grid(
            row=4, column=0, columnspan=2, sticky="w", padx=15, pady=(0, 8))

        # ---- Asientos con fecha: DEDUCCIONES y HORAS EXTRAS ----
        if con_deducciones:
            f_asi = ctk.CTkFrame(scroll, corner_radius=10, border_width=1, border_color="#e0e0e0")
            f_asi.pack(fill="x", pady=5, ipady=6)

            def _crear_seccion(titulo, ayuda, color_total, destino):
                """Crea un cuadro de asientos (fecha, tipo de día, horas, minutos, motivo)."""
                ctk.CTkLabel(f_asi, text=titulo, font=(familia_fuente, 13, "bold"),
                             text_color="#1f538d").pack(anchor="w", padx=15, pady=(8, 2))
                ctk.CTkLabel(f_asi, text=ayuda, font=(familia_fuente, 10), text_color="gray",
                             justify="left", wraplength=700).pack(anchor="w", padx=15, pady=(0, 6))

                f_cab = ctk.CTkFrame(f_asi, fg_color="transparent")
                f_cab.pack(fill="x", padx=12)
                for col, (txt, w) in enumerate((("Fecha", 96), ("", 34), ("Tipo de día", 118),
                                                ("Horas", 66), ("Minutos", 66), ("Motivo", 220), ("", 32))):
                    ctk.CTkLabel(f_cab, text=txt, width=w, anchor="w",
                                 font=(familia_fuente, 10, "bold"),
                                 text_color="#7f8c8d").grid(row=0, column=col, padx=3, sticky="w")

                contenedor = ctk.CTkFrame(f_asi, fg_color="transparent")
                contenedor.pack(fill="x", padx=12, pady=(2, 4))

                f_acc = ctk.CTkFrame(f_asi, fg_color="transparent")
                f_acc.pack(fill="x", padx=12, pady=(2, 4))
                ctk.CTkButton(f_acc, text="➕ Agregar asiento", width=170,
                              font=(familia_fuente, 11, "bold"), fg_color="#e67e22",
                              hover_color="#d35400",
                              command=lambda d=destino: self._agregar_fila_asiento(d)).pack(
                    side="left", padx=(0, 8))
                ctk.CTkButton(f_acc, text="🧹 Vaciar asientos", width=150,
                              font=(familia_fuente, 11, "bold"), fg_color="#7f8c8d",
                              hover_color="#606b6b",
                              command=lambda d=destino: self._vaciar_asientos(d)).pack(side="left")

                lbl = ctk.CTkLabel(f_asi, text="", font=(familia_fuente, 11, "bold"),
                                   text_color=color_total, justify="left")
                lbl.pack(anchor="w", padx=15, pady=(2, 8))
                ctk.CTkFrame(f_asi, height=1, fg_color="#dcdcdc").pack(fill="x", padx=15, pady=(0, 4))
                return contenedor, lbl

            self.frame_asientos, self.lbl_total_ded = _crear_seccion(
                "➖ DEDUCCIONES DE LA QUINCENA (asientos con fecha)",
                ("Registre uno o varios asientos: fecha, horas y minutos no prestados y el motivo.\n"
                 "El tipo de día (Normal / Domingo / Feriado) se detecta solo con la fecha; puede cambiarlo.\n"
                 "Estos asientos se RESTAN del cobro y se detallan en el PDF."),
                "#c0392b", "ded")

            self.frame_extras, self.lbl_total_extras = _crear_seccion(
                "⏱️ HORAS EXTRAS DE LA QUINCENA (asientos con fecha)",
                ("Registre las horas trabajadas de más: fecha, horas/minutos extras y el motivo.\n"
                 "Se valorizan al precio por hora de la unidad según el tipo de día de la fecha\n"
                 "y SUMAN al cobro: subtotal = base − deducciones + horas extras."),
                "#27ae60", "ext")

            # Asientos guardados; si el registro es antiguo (sin detalle) se muestra el resumen.
            asientos_ini = [dict(a) for a in (unidad.get("ded_asientos") or [])]
            if not asientos_ini:
                for clave, _etq in CATEGORIAS:
                    h, m = (unidad.get("ded") or {}).get(clave, [0.0, 0.0])
                    if float(h or 0) or float(m or 0):
                        asientos_ini.append(nuevo_asiento_deduccion(
                            None, clave, h, m, motivo="Sin detalle (registro anterior)"))
            for a in asientos_ini:
                self._agregar_fila_asiento("ded", a)
            for a in (unidad.get("extras_asientos") or []):
                self._agregar_fila_asiento("ext", a)
            self._actualizar_total_asientos("ded")
            self._actualizar_total_asientos("ext")

        # ---- Acciones ----
        f_acc = ctk.CTkFrame(scroll, fg_color="transparent")
        f_acc.pack(fill="x", pady=10)
        btn_ok = ctk.CTkButton(f_acc, text="💾 Guardar", width=150, height=38,
                               font=(familia_fuente, 13, "bold"), fg_color="#27ae60",
                               hover_color="#1e8449", command=lambda: self._aceptar(unidad))
        btn_ok.pack(side="left", padx=5)
        btn_cancel = ctk.CTkButton(f_acc, text="✖ Cancelar", width=120, height=38,
                                   font=(familia_fuente, 13, "bold"), fg_color="#7f8c8d",
                                   hover_color="#606b6b", command=self.destroy)
        btn_cancel.pack(side="left", padx=5)

        self.wait_window()

    def _leer_num(self, entry, nombre):
        try:
            v = float(str(entry.get()).strip().replace(",", "") or 0)
            if v < 0:
                raise ValueError
            return v
        except ValueError:
            raise ValueError(f"'{nombre}' debe ser un número válido (mayor o igual a 0).")

    # ---------- Asientos con fecha (deducciones y horas extras) ----------
    def _categoria_para_fecha(self, fecha):
        """Tipo de día (normal/domingo/feriado) según la quincena cargada."""
        for d in (getattr(self.app, "dias", None) or []):
            if d.get("fecha") == fecha:
                return d.get("categoria", "normal")
        try:
            return clasificar_dia(fecha, getattr(self.app, "feriados", None) or {})
        except Exception:
            return "normal"

    def _sincronizar_tipo(self, item):
        """Ajusta el tipo de día del asiento según la fecha escrita."""
        try:
            fecha = datetime.strptime(item["fecha"].get().strip(), "%d/%m/%Y").date()
        except Exception:
            return
        try:
            item["tipo"].set(ETIQUETAS_CORTAS.get(self._categoria_para_fecha(fecha), "Normal"))
        except Exception:
            pass

    def _abrir_calendario_asiento(self, item):
        """Calendario de la quincena para elegir la fecha del asiento."""
        anio = mes = None
        dias_q = {}
        dias = getattr(self.app, "dias", None) or []
        if dias:
            anio, mes = dias[0]["fecha"].year, dias[0]["fecha"].month
            for d in dias:
                if d["fecha"].year == anio and d["fecha"].month == mes:
                    dias_q[d["fecha"].day] = d["categoria"]
        try:
            CalendarioPopup(self, item["fecha"], lambda: self._sincronizar_tipo(item),
                            anio=anio, mes=mes, dias_resaltados=dias_q)
        except Exception as e:
            print("[Calendario Asiento Error]", e)

    def _widgets_asiento(self, destino):
        """(contenedor, lista de filas, etiqueta de totales) del cuadro indicado."""
        if destino == "ext":
            return self.frame_extras, self.filas_extras, self.lbl_total_extras
        return self.frame_asientos, self.filas_ded, self.lbl_total_ded

    def _agregar_fila_asiento(self, destino="ded", asiento=None):
        """Agrega una fila (asiento con fecha) al cuadro de deducciones o de horas extras."""
        asiento = asiento or {}
        contenedor, lista, _lbl = self._widgets_asiento(destino)
        fam = self.familia_fuente
        fila = ctk.CTkFrame(contenedor, fg_color="transparent")
        fila.pack(fill="x", pady=1)
        item = {"frame": fila, "destino": destino}

        e_fecha = ctk.CTkEntry(fila, width=96, placeholder_text="dd/mm/aaaa")
        e_fecha.grid(row=0, column=0, padx=3, sticky="w")
        item["fecha"] = e_fecha

        ctk.CTkButton(fila, text="📅", width=30, font=(fam, 12),
                      command=lambda it=item: self._abrir_calendario_asiento(it)).grid(
            row=0, column=1, padx=3)

        combo = ctk.CTkOptionMenu(fila, values=list(ETIQUETAS_CORTAS.values()), width=118,
                                  font=(fam, 11),
                                  command=lambda _v, d=destino: self._actualizar_total_asientos(d))
        combo.grid(row=0, column=2, padx=3, sticky="w")
        item["tipo"] = combo

        e_h = ctk.CTkEntry(fila, width=66, justify="center", placeholder_text="0")
        e_h.grid(row=0, column=3, padx=3)
        item["horas"] = e_h
        e_m = ctk.CTkEntry(fila, width=66, justify="center", placeholder_text="0")
        e_m.grid(row=0, column=4, padx=3)
        item["min"] = e_m
        e_mot = ctk.CTkEntry(fila, width=220, placeholder_text="Motivo / detalle")
        e_mot.grid(row=0, column=5, padx=3)
        item["motivo"] = e_mot

        ctk.CTkButton(fila, text="🗑", width=32, font=(fam, 12), fg_color="#e74c3c",
                      hover_color="#c0392b",
                      command=lambda it=item: self._quitar_fila_asiento(it)).grid(row=0, column=6, padx=3)

        cat = asiento.get("categoria") if asiento.get("categoria") in ETIQUETAS_CORTAS else "normal"
        combo.set(ETIQUETAS_CORTAS.get(cat, "Normal"))
        if asiento.get("fecha"):
            try:
                e_fecha.insert(0, asiento["fecha"].strftime("%d/%m/%Y"))
            except Exception:
                pass
        horas_ini = float(asiento.get("horas") or 0)
        min_ini = float(asiento.get("minutos") or 0)
        if horas_ini:
            e_h.insert(0, f"{horas_ini:g}")
        if min_ini:
            e_m.insert(0, f"{min_ini:g}")
        e_mot.insert(0, str(asiento.get("motivo") or ""))

        for w in (e_h, e_m):
            w.bind("<KeyRelease>", lambda e, d=destino: self._actualizar_total_asientos(d))
        e_fecha.bind("<KeyRelease>", lambda e, it=item: self._sincronizar_tipo(it))
        e_fecha.bind("<FocusOut>", lambda e, it=item: self._sincronizar_tipo(it))

        lista.append(item)
        self._actualizar_total_asientos(destino)

    def _quitar_fila_asiento(self, item):
        destino = item.get("destino", "ded")
        try:
            item["frame"].destroy()
        except Exception:
            pass
        _c, lista, _lbl = self._widgets_asiento(destino)
        if item in lista:
            lista.remove(item)
        self._actualizar_total_asientos(destino)

    def _vaciar_asientos(self, destino="ded"):
        """Quita todas las filas del cuadro indicado ('ded' o 'ext')."""
        _c, lista, _lbl = self._widgets_asiento(destino)
        for item in list(lista):
            self._quitar_fila_asiento(item)

    def _leer_asientos_filas(self, destino="ded", validar=False):
        """Lee un cuadro de asientos y devuelve la lista de asientos."""
        _c, lista, _lbl = self._widgets_asiento(destino)
        nombre = "Horas extras" if destino == "ext" else "Deducción"
        asientos = []
        for i, item in enumerate(lista, start=1):
            texto = item["fecha"].get().strip()
            fecha = None
            if texto:
                try:
                    fecha = datetime.strptime(texto, "%d/%m/%Y").date()
                except Exception:
                    if validar:
                        raise ValueError(f"{nombre} {i}: la fecha '{texto}' no es válida (use dd/mm/aaaa).")
            cat = ETIQUETAS_CORTAS_A_CLAVE.get(item["tipo"].get(), "normal")
            horas = self._leer_num(item["horas"], f"{nombre} {i}: horas")
            minutos = self._leer_num(item["min"], f"{nombre} {i}: minutos")
            asientos.append(nuevo_asiento_deduccion(fecha, cat, horas, minutos, 0.0,
                                                    item["motivo"].get().strip()))
        return asientos

    def _actualizar_total_asientos(self, destino="ded"):
        """Resumen de horas y monto de los asientos (deducido o a sumar por horas extras)."""
        try:
            asientos = self._leer_asientos_filas(destino)
            precios = {"normal": self._leer_num(self.ent_pn, "Precio Día Normal"),
                       "domingo": self._leer_num(self.ent_pd, "Precio Domingo"),
                       "feriado": self._leer_num(self.ent_pf, "Precio Feriado")}
        except Exception:
            return
        agrupado = ded_agrupada_desde_asientos(asientos)
        monto = 0.0
        partes = []
        for clave, _etq in CATEGORIAS:
            h, m = agrupado[clave]
            horas_cat = h + m / 60.0
            subtotal = horas_cat * precios[clave]
            monto += subtotal
            partes.append(f"{ETIQUETAS_CORTAS[clave]}: {horas_cat:g} h ({formatear_moneda(subtotal)})")
        resumen = f"Asientos: {len(asientos)}  |  " + "  ·  ".join(partes)
        if destino == "ext":
            texto = resumen + f"\nMONTO A SUMAR (HORAS EXTRAS): {formatear_moneda(monto)}"
            etiqueta = self.lbl_total_extras
        else:
            texto = resumen + f"\nMONTO DEDUCIDO: {formatear_moneda(monto)}"
            etiqueta = self.lbl_total_ded
        try:
            etiqueta.configure(text=texto)
        except Exception:
            pass

    def _aceptar(self, unidad):
        try:
            horas_dia = self._leer_num(self.ent_horas, "Horas al día")
            pn = self._leer_num(self.ent_pn, "Precio Día Normal")
            pd = self._leer_num(self.ent_pd, "Precio Domingo")
            pf = self._leer_num(self.ent_pf, "Precio Feriado")
            asientos = []
            extras = []
            ded = {}
            if self.con_deducciones:
                asientos = self._leer_asientos_filas("ded", validar=True)
                # Las horas por categoría salen de la suma de los asientos con fecha
                ded = ded_agrupada_desde_asientos(asientos)
                extras = self._leer_asientos_filas("ext", validar=True)
        except ValueError as e:
            messagebox.showerror("Datos Inválidos", str(e), parent=self)
            return
        resultado = dict(unidad)
        resultado["horas_dia"] = horas_dia
        resultado["precio_normal"] = pn
        resultado["precio_domingo"] = pd
        resultado["precio_feriado"] = pf
        if self.con_deducciones:
            resultado["ded"] = ded
            resultado["ded_asientos"] = asientos
            resultado["extras_asientos"] = extras
        self.result = resultado
        self.destroy()


# =========================================================
# 🚀 DIÁLOGO: SELECCIONAR VEHÍCULO DE LA FLOTA
# =========================================================
class DialogoSeleccionVehiculo(ctk.CTkToplevel):
    """Lista los vehículos de la flota no asignados al cliente para agregarlos."""

    def __init__(self, parent, id_cliente, excluir_ids=None):
        super().__init__(parent)
        self.result = None
        self.id_cliente = id_cliente
        self.excluir_ids = set(excluir_ids or [])
        familia_fuente = "Helvetica" if sys.platform == "darwin" else "Arial"

        self.title("Agregar Vehículo de la Flota")
        self.geometry("700x560")
        self.resizable(False, False)
        self.transient(parent)
        try:
            self.grab_set()
        except Exception:
            pass
        self.update_idletasks()
        try:
            x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (700 // 2)
            y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (560 // 2)
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

        ctk.CTkLabel(self, text="🚗 SELECCIONAR VEHÍCULOS DE LA FLOTA",
                     font=(familia_fuente, 15, "bold"), text_color="#1f538d").pack(pady=(12, 4))
        ctk.CTkLabel(self, text="Solo se muestran los vehículos aún no asignados a este cliente.\n"
                                "Mantén presionada la tecla Ctrl (o Shift) para seleccionar varios y agregarlos en bloque.",
                     font=(familia_fuente, 10), text_color="gray", justify="center").pack()

        f_tbl = ctk.CTkFrame(self, corner_radius=10)
        f_tbl.pack(fill="both", expand=True, padx=15, pady=8)

        columnas = ("placa", "vehiculo", "estado")
        style = ttk.Style()
        if sys.platform == "darwin":
            style.theme_use("clam")
        style.configure("Treeview", background="#ffffff", foreground="#000000",
                        fieldbackground="#ffffff", rowheight=26, font=(familia_fuente, 10),
                        bordercolor="#e0e0e0", borderwidth=1)
        style.map("Treeview", background=[("selected", "#1f538d")], foreground=[("selected", "#ffffff")])
        style.configure("Treeview.Heading", background="#f0f0f0", foreground="#000000",
                        font=(familia_fuente, 10, "bold"), bordercolor="#e0e0e0", borderwidth=1, relief="flat")

        self.tabla = ttk.Treeview(f_tbl, columns=columnas, show="headings", selectmode="extended",
                                  style="Treeview")
        self.tabla.heading("placa", text="Placa", anchor="center")
        self.tabla.heading("vehiculo", text="Vehículo", anchor="center")
        self.tabla.heading("estado", text="Estado", anchor="center")
        self.tabla.column("placa", width=90, anchor="center")
        self.tabla.column("vehiculo", width=380, anchor="w")
        self.tabla.column("estado", width=100, anchor="center")
        self.tabla.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        scr = ttk.Scrollbar(f_tbl, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scr.set)
        scr.pack(side="right", fill="y", pady=10, padx=(0, 10))
        self.tabla.bind("<Double-1>", lambda e: self._agregar())

        f_btns = ctk.CTkFrame(self, fg_color="transparent")
        f_btns.pack(fill="x", padx=15, pady=10)
        btn_add = ctk.CTkButton(f_btns, text="➕ Agregar seleccionados", width=230, height=40,
                                font=(familia_fuente, 13, "bold"),
                                fg_color="#27ae60", hover_color="#1e8449", command=self._agregar)
        btn_add.pack(side="left", padx=5)
        btn_cancel = ctk.CTkButton(f_btns, text="✖ Cerrar", width=120, height=40,
                                   font=(familia_fuente, 13, "bold"),
                                   fg_color="#7f8c8d", hover_color="#606b6b", command=self.destroy)
        btn_cancel.pack(side="right", padx=5)

        self.lbl_sel = ctk.CTkLabel(f_btns, text="", font=(familia_fuente, 11, "bold"),
                                    text_color="#1f538d")
        self.lbl_sel.pack(side="right", padx=10)

        self.cargar_vehiculos()
        self.tabla.bind("<<TreeviewSelect>>", lambda e: self._actualizar_contador())
        self._actualizar_contador()
        self.wait_window()

    def cargar_vehiculos(self):
        conn = conectar_db(silencioso=True)
        if not conn:
            messagebox.showwarning("Modo Lectura", "Sin conexión a la base de datos.", parent=self)
            return
        try:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT v.id, v.placa, v.marca, v.modelo, v.anio, v.color, v.estado
                FROM flota_vehiculos v
                WHERE v.id NOT IN (
                    SELECT id_vehiculo FROM clientes_unidades
                    WHERE id_cliente = %s AND id_vehiculo IS NOT NULL
                )
                ORDER BY v.placa ASC
            ''', (self.id_cliente,))
            for r in cursor.fetchall():
                if r[0] in self.excluir_ids:
                    continue  # ya está en la lista actual (sin guardar)
                desc = " ".join(str(x) for x in (r[2], r[3], r[4], r[5]) if x and str(x).strip())
                self.tabla.insert("", tk.END, values=(r[1], desc, r[6]))
                self._filas = getattr(self, "_filas", [])
                self._filas.append({"id_vehiculo": r[0], "placa": str(r[1]), "desc": desc})
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar los vehículos:\n{e}", parent=self)
        finally:
            liberar_conexion(conn)

    def _actualizar_contador(self):
        try:
            n = len(self.tabla.selection())
            self.lbl_sel.configure(text=f"Seleccionados: {n}" if n else "Selecciona vehículos")
        except Exception:
            pass

    def _agregar(self):
        sel = self.tabla.selection()
        if not sel:
            messagebox.showinfo("Seleccione vehículos", "Seleccione uno o varios vehículos de la lista.", parent=self)
            return
        filas = getattr(self, "_filas", [])
        resultados = []
        for iid in sel:
            idx = self.tabla.index(iid)
            if idx < len(filas):
                f = filas[idx]
                resultados.append({
                    "id_vehiculo": f["id_vehiculo"],
                    "unidad": f"{f['placa']} — {f['desc']}".strip(),
                })
        if resultados:
            self.result = resultados  # lista de vehículos seleccionados
        self.destroy()


# =========================================================
# 🚀 VENTANA: UNIDADES ASIGNADAS DEL CLIENTE
# =========================================================
class VentanaUnidadesCliente(ctk.CTkToplevel):
    """Asigna vehículos de la flota al cliente y configura sus precios y horas/día."""

    def __init__(self, parent, app, id_cliente, nombre_cliente):
        super().__init__(parent)
        self.app = app
        self.id_cliente = id_cliente
        self.nombre_cliente = nombre_cliente
        self.filas = []
        familia_fuente = "Helvetica" if sys.platform == "darwin" else "Arial"

        self.title(f"Unidades Asignadas — {nombre_cliente}")
        self.geometry("760x480")
        self.transient(parent)
        try:
            self.grab_set()
        except Exception:
            pass
        self.update_idletasks()
        try:
            x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (760 // 2)
            y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (480 // 2)
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

        f_top = ctk.CTkFrame(self, fg_color="transparent")
        f_top.pack(fill="x", padx=15, pady=(12, 6))
        ctk.CTkLabel(f_top, text=f"🚗 UNIDADES ASIGNADAS — {nombre_cliente}",
                     font=(familia_fuente, 15, "bold"), text_color="#1f538d").pack(side="left")

        f_tbl = ctk.CTkFrame(self, corner_radius=10)
        f_tbl.pack(fill="both", expand=True, padx=15, pady=6)

        columnas = ("unidad", "horas", "pnormal", "pdom", "pfer")
        style = ttk.Style()
        if sys.platform == "darwin":
            style.theme_use("clam")
        style.configure("Treeview", background="#ffffff", foreground="#000000",
                        fieldbackground="#ffffff", rowheight=26, font=(familia_fuente, 10),
                        bordercolor="#e0e0e0", borderwidth=1)
        style.map("Treeview", background=[("selected", "#1f538d")], foreground=[("selected", "#ffffff")])
        style.configure("Treeview.Heading", background="#f0f0f0", foreground="#000000",
                        font=(familia_fuente, 10, "bold"), bordercolor="#e0e0e0", borderwidth=1, relief="flat")

        self.tabla = ttk.Treeview(f_tbl, columns=columnas, show="headings", selectmode="browse",
                                  style="Treeview")
        self.tabla.heading("unidad", text="Unidad", anchor="center")
        self.tabla.heading("horas", text="Horas/Día", anchor="center")
        self.tabla.heading("pnormal", text="P. Normal", anchor="center")
        self.tabla.heading("pdom", text="P. Domingo", anchor="center")
        self.tabla.heading("pfer", text="P. Feriado", anchor="center")
        self.tabla.column("unidad", width=280, anchor="w")
        self.tabla.column("horas", width=80, anchor="center")
        self.tabla.column("pnormal", width=90, anchor="center")
        self.tabla.column("pdom", width=90, anchor="center")
        self.tabla.column("pfer", width=90, anchor="center")
        self.tabla.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        scr = ttk.Scrollbar(f_tbl, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scr.set)
        scr.pack(side="right", fill="y", pady=10, padx=(0, 10))
        self.tabla.bind("<Double-1>", lambda e: self.editar())

        f_btns = ctk.CTkFrame(self, fg_color="transparent")
        f_btns.pack(fill="x", padx=15, pady=8)
        btn_add = ctk.CTkButton(f_btns, text="➕ Agregar de la Flota", width=170, font=(familia_fuente, 12, "bold"),
                                fg_color="#27ae60", hover_color="#1e8449", command=self.agregar)
        btn_add.pack(side="left", padx=5)
        btn_edit = ctk.CTkButton(f_btns, text="✏️ Editar", width=110, font=(familia_fuente, 12, "bold"),
                                 fg_color="#34495e", hover_color="#2c3e50", command=self.editar)
        btn_edit.pack(side="left", padx=5)
        btn_del = ctk.CTkButton(f_btns, text="➖ Quitar", width=110, font=(familia_fuente, 12, "bold"),
                                fg_color="#e74c3c", hover_color="#c0392b", command=self.quitar)
        btn_del.pack(side="left", padx=5)
        btn_save = ctk.CTkButton(f_btns, text="💾 Guardar", width=140, font=(familia_fuente, 12, "bold"),
                                 fg_color="#1f538d", hover_color="#163b65", command=self.guardar)
        btn_save.pack(side="right", padx=5)
        btn_cerrar = ctk.CTkButton(f_btns, text="✖ Cerrar", width=100, font=(familia_fuente, 12, "bold"),
                                   fg_color="#7f8c8d", hover_color="#606b6b", command=self.destroy)
        btn_cerrar.pack(side="right", padx=5)

        self.cargar_filas()

    def cargar_filas(self):
        self.filas = []
        conn = conectar_db(silencioso=True)
        if conn:
            try:
                cursor = conn.cursor()
                cursor.execute('''
                    SELECT id, id_vehiculo, unidad, precio_normal, precio_domingo, precio_feriado, horas_dia
                    FROM clientes_unidades WHERE id_cliente=%s ORDER BY id
                ''', (self.id_cliente,))
                for r in cursor.fetchall():
                    self.filas.append({
                        "id": r[0], "id_vehiculo": r[1], "unidad": str(r[2] or ""),
                        "precio_normal": float(r[3] or 0), "precio_domingo": float(r[4] or 0),
                        "precio_feriado": float(r[5] or 0), "horas_dia": float(r[6] or 0),
                    })
                cursor.close()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudieron cargar las unidades:\n{e}", parent=self)
            finally:
                liberar_conexion(conn)
        self.pintar()

    def pintar(self):
        for item in self.tabla.get_children():
            self.tabla.delete(item)
        for f in self.filas:
            self.tabla.insert("", tk.END, values=(
                f["unidad"],
                f"{float(f['horas_dia'] or 0):g}",
                f"{float(f['precio_normal'] or 0):g}",
                f"{float(f['precio_domingo'] or 0):g}",
                f"{float(f['precio_feriado'] or 0):g}",
            ))

    def agregar(self):
        excluir = [f.get("id_vehiculo") for f in self.filas if f.get("id_vehiculo")]
        dlg = DialogoSeleccionVehiculo(self, self.id_cliente, excluir_ids=excluir)
        if dlg.result:
            agregados = 0
            for v in dlg.result:
                if any(f.get("id_vehiculo") == v["id_vehiculo"] for f in self.filas):
                    continue  # ya está en la lista
                self.filas.append({
                    "id": None, "id_vehiculo": v["id_vehiculo"], "unidad": v["unidad"],
                    "precio_normal": 0.0, "precio_domingo": 0.0, "precio_feriado": 0.0, "horas_dia": 8.0,
                })
                agregados += 1
            self.pintar()
            if agregados > 1:
                messagebox.showinfo("Unidades agregadas", f"{agregados} vehículos agregados.", parent=self)

    def editar(self):
        sel = self.tabla.selection()
        if not sel:
            messagebox.showinfo("Seleccione una unidad", "Seleccione una unidad de la lista.", parent=self)
            return
        idx = self.tabla.index(sel[0])
        dlg = DialogoUnidad(self, self.filas[idx], "Por Hora", con_deducciones=False)
        if dlg.result is not None:
            self.filas[idx] = dlg.result
            self.pintar()

    def quitar(self):
        sel = self.tabla.selection()
        if not sel:
            messagebox.showinfo("Seleccione una unidad", "Seleccione una unidad de la lista.", parent=self)
            return
        idx = self.tabla.index(sel[0])
        if messagebox.askyesno("Quitar unidad", f"¿Quitar la unidad '{self.filas[idx]['unidad']}' de este cliente?", parent=self):
            self.filas.pop(idx)
            self.pintar()

    def guardar(self):
        conn = conectar_db()
        if not conn:
            messagebox.showwarning("Modo Lectura", "Sin conexión. No se puede guardar.", parent=self)
            return
        try:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM clientes_unidades WHERE id_cliente=%s", (self.id_cliente,))
            vistos = set()
            for f in self.filas:
                idv = f.get("id_vehiculo")
                if idv is not None:
                    if idv in vistos:
                        continue  # evita duplicar el mismo vehículo
                    vistos.add(idv)
                cursor.execute('''
                    INSERT INTO clientes_unidades
                    (id_cliente, id_vehiculo, unidad, precio_normal, precio_domingo, precio_feriado, horas_dia)
                    VALUES (%s,%s,%s,%s,%s,%s,%s)
                ''', (self.id_cliente, idv, f["unidad"],
                      f["precio_normal"], f["precio_domingo"], f["precio_feriado"], f["horas_dia"]))
            conn.commit()
            registrar_auditoria(self.app.usuario_activo, "Cobranza",
                                f"Configuró las unidades asignadas del cliente '{self.nombre_cliente}'")
            messagebox.showinfo("Guardado", "Unidades del cliente guardadas correctamente.", parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar:\n{e}", parent=self)
        finally:
            liberar_conexion(conn)
        self.app.cargar_unidades_cliente()
        self.destroy()




# =========================================================
# 🚀 DIÁLOGO: AGREGAR/EDITAR RANGO DE DISTANCIA
# =========================================================
class DialogoRangoDistancia(ctk.CTkToplevel):
    def __init__(self, parent, datos=None):
        super().__init__(parent)
        self.result = None
        familia_fuente = "Helvetica" if sys.platform == "darwin" else "Arial"
        self.title("Rango de Distancia" if not datos else "Editar Rango de Distancia")
        self.geometry("380x330")
        self.resizable(False, False)
        self.transient(parent)
        try:
            self.grab_set()
        except Exception:
            pass
        self.update_idletasks()
        try:
            x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (380 // 2)
            y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (330 // 2)
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

        f = ctk.CTkFrame(self, corner_radius=10, border_width=1, border_color="#e0e0e0")
        f.pack(fill="both", expand=True, padx=15, pady=15, ipady=10)
        f.columnconfigure(1, weight=1)

        ctk.CTkLabel(f, text="Distancia desde (km):", font=(familia_fuente, 12, "bold")).grid(
            row=0, column=0, sticky="w", padx=(15, 5), pady=8)
        self.ent_desde = ctk.CTkEntry(f, width=120)
        self.ent_desde.grid(row=0, column=1, sticky="w", padx=5, pady=8)

        ctk.CTkLabel(f, text="Distancia hasta (km):", font=(familia_fuente, 12, "bold")).grid(
            row=1, column=0, sticky="w", padx=(15, 5), pady=8)
        self.ent_hasta = ctk.CTkEntry(f, width=120)
        self.ent_hasta.grid(row=1, column=1, sticky="w", padx=5, pady=8)

        ctk.CTkLabel(f, text="Precio día normal:", font=(familia_fuente, 12, "bold")).grid(
            row=2, column=0, sticky="w", padx=(15, 5), pady=8)
        self.ent_pn = ctk.CTkEntry(f, width=120, placeholder_text="0.00")
        self.ent_pn.grid(row=2, column=1, sticky="w", padx=5, pady=8)

        ctk.CTkLabel(f, text="Precio domingo:", font=(familia_fuente, 12, "bold")).grid(
            row=3, column=0, sticky="w", padx=(15, 5), pady=8)
        self.ent_pd = ctk.CTkEntry(f, width=120, placeholder_text="0.00")
        self.ent_pd.grid(row=3, column=1, sticky="w", padx=5, pady=8)

        ctk.CTkLabel(f, text="Precio feriado:", font=(familia_fuente, 12, "bold")).grid(
            row=4, column=0, sticky="w", padx=(15, 5), pady=8)
        self.ent_pf = ctk.CTkEntry(f, width=120, placeholder_text="0.00")
        self.ent_pf.grid(row=4, column=1, sticky="w", padx=5, pady=8)

        f_acc = ctk.CTkFrame(f, fg_color="transparent")
        f_acc.grid(row=5, column=0, columnspan=2, pady=12)
        btn_ok = ctk.CTkButton(f_acc, text="💾 Guardar", width=130, font=(familia_fuente, 12, "bold"),
                               fg_color="#27ae60", hover_color="#1e8449", command=self._aceptar)
        btn_ok.pack(side="left", padx=5)
        btn_cancel = ctk.CTkButton(f_acc, text="✖ Cancelar", width=110, font=(familia_fuente, 12, "bold"),
                                   fg_color="#7f8c8d", hover_color="#606b6b", command=self.destroy)
        btn_cancel.pack(side="left", padx=5)

        if datos:
            self.ent_desde.insert(0, f"{float(datos.get('distancia_desde') or 0):g}")
            self.ent_hasta.insert(0, f"{float(datos.get('distancia_hasta') or 0):g}")
            self.ent_pn.insert(0, f"{float(datos.get('precio_normal') or 0):g}")
            self.ent_pd.insert(0, f"{float(datos.get('precio_domingo') or 0):g}")
            self.ent_pf.insert(0, f"{float(datos.get('precio_feriado') or 0):g}")

        self.wait_window()

    def _leer_num(self, entry, nombre):
        try:
            v = float(str(entry.get()).strip().replace(",", "") or 0)
            if v < 0:
                raise ValueError
            return v
        except ValueError:
            raise ValueError(f"'{nombre}' debe ser un número válido (mayor o igual a 0).")

    def _aceptar(self):
        try:
            desde = self._leer_num(self.ent_desde, "Distancia desde")
            hasta = self._leer_num(self.ent_hasta, "Distancia hasta")
            if hasta < desde:
                raise ValueError("La distancia 'hasta' debe ser mayor o igual que 'desde'.")
            pn = self._leer_num(self.ent_pn, "Precio día normal")
            pd = self._leer_num(self.ent_pd, "Precio domingo")
            pf = self._leer_num(self.ent_pf, "Precio feriado")
        except ValueError as e:
            messagebox.showerror("Datos Inválidos", str(e), parent=self)
            return
        self.result = {"distancia_desde": desde, "distancia_hasta": hasta,
                       "precio_normal": pn, "precio_domingo": pd, "precio_feriado": pf}
        self.destroy()


# =========================================================
# 🚀 VENTANA: TABLA DE PRECIOS POR DISTANCIA DEL CLIENTE
# =========================================================
class VentanaPreciosDistancia(ctk.CTkToplevel):
    def __init__(self, parent, app, id_cliente, nombre_cliente):
        super().__init__(parent)
        self.app = app
        self.id_cliente = id_cliente
        self.nombre_cliente = nombre_cliente
        self.filas = []
        familia_fuente = "Helvetica" if sys.platform == "darwin" else "Arial"

        self.title(f"Precios por Distancia — {nombre_cliente}")
        self.geometry("680x460")
        self.transient(parent)
        try:
            self.grab_set()
        except Exception:
            pass
        self.update_idletasks()
        try:
            x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (680 // 2)
            y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (460 // 2)
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

        f_top = ctk.CTkFrame(self, fg_color="transparent")
        f_top.pack(fill="x", padx=15, pady=(12, 6))
        ctk.CTkLabel(f_top, text=f"📏 PRECIOS POR DISTANCIA — {nombre_cliente}",
                     font=(familia_fuente, 15, "bold"), text_color="#1f538d").pack(side="left")

        f_tbl = ctk.CTkFrame(self, corner_radius=10)
        f_tbl.pack(fill="both", expand=True, padx=15, pady=6)

        columnas = ("distancia", "pnormal", "pdom", "pfer")
        style = ttk.Style()
        if sys.platform == "darwin":
            style.theme_use("clam")
        style.configure("Treeview", background="#ffffff", foreground="#000000",
                        fieldbackground="#ffffff", rowheight=26, font=(familia_fuente, 10),
                        bordercolor="#e0e0e0", borderwidth=1)
        style.map("Treeview", background=[("selected", "#1f538d")], foreground=[("selected", "#ffffff")])
        style.configure("Treeview.Heading", background="#f0f0f0", foreground="#000000",
                        font=(familia_fuente, 10, "bold"), bordercolor="#e0e0e0", borderwidth=1, relief="flat")

        self.tabla = ttk.Treeview(f_tbl, columns=columnas, show="headings", selectmode="browse",
                                  style="Treeview")
        self.tabla.heading("distancia", text="Distancia (km)", anchor="center")
        self.tabla.heading("pnormal", text="P. Normal", anchor="center")
        self.tabla.heading("pdom", text="P. Domingo", anchor="center")
        self.tabla.heading("pfer", text="P. Feriado", anchor="center")
        self.tabla.column("distancia", width=220, anchor="center")
        self.tabla.column("pnormal", width=120, anchor="center")
        self.tabla.column("pdom", width=120, anchor="center")
        self.tabla.column("pfer", width=120, anchor="center")
        self.tabla.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        scr = ttk.Scrollbar(f_tbl, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scr.set)
        scr.pack(side="right", fill="y", pady=10, padx=(0, 10))
        self.tabla.bind("<Double-1>", lambda e: self.editar())

        f_btns = ctk.CTkFrame(self, fg_color="transparent")
        f_btns.pack(fill="x", padx=15, pady=8)
        btn_add = ctk.CTkButton(f_btns, text="➕ Agregar Rango", width=140, font=(familia_fuente, 12, "bold"),
                                fg_color="#27ae60", hover_color="#1e8449", command=self.agregar)
        btn_add.pack(side="left", padx=5)
        btn_edit = ctk.CTkButton(f_btns, text="✏️ Editar", width=110, font=(familia_fuente, 12, "bold"),
                                 fg_color="#34495e", hover_color="#2c3e50", command=self.editar)
        btn_edit.pack(side="left", padx=5)
        btn_del = ctk.CTkButton(f_btns, text="➖ Quitar", width=110, font=(familia_fuente, 12, "bold"),
                                fg_color="#e74c3c", hover_color="#c0392b", command=self.quitar)
        btn_del.pack(side="left", padx=5)
        btn_save = ctk.CTkButton(f_btns, text="💾 Guardar", width=130, font=(familia_fuente, 12, "bold"),
                                 fg_color="#1f538d", hover_color="#163b65", command=self.guardar)
        btn_save.pack(side="right", padx=5)
        btn_cerrar = ctk.CTkButton(f_btns, text="✖ Cerrar", width=100, font=(familia_fuente, 12, "bold"),
                                   fg_color="#7f8c8d", hover_color="#606b6b", command=self.destroy)
        btn_cerrar.pack(side="right", padx=5)

        self.cargar_filas()

    def cargar_filas(self):
        self.filas = []
        conn = conectar_db(silencioso=True)
        if conn:
            try:
                cursor = conn.cursor()
                cursor.execute('''
                    SELECT id, distancia_desde, distancia_hasta, precio_normal, precio_domingo, precio_feriado
                    FROM precios_viaje_distancia WHERE id_cliente=%s ORDER BY distancia_desde
                ''', (self.id_cliente,))
                for r in cursor.fetchall():
                    self.filas.append({
                        "id": r[0], "distancia_desde": float(r[1]), "distancia_hasta": float(r[2]),
                        "precio_normal": float(r[3] or 0), "precio_domingo": float(r[4] or 0),
                        "precio_feriado": float(r[5] or 0),
                    })
                cursor.close()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudieron cargar los rangos:\n{e}", parent=self)
            finally:
                liberar_conexion(conn)
        self.pintar()

    def pintar(self):
        for item in self.tabla.get_children():
            self.tabla.delete(item)
        for f in self.filas:
            self.tabla.insert("", tk.END, values=(
                f"{float(f['distancia_desde']):g} - {float(f['distancia_hasta']):g}",
                f"{float(f['precio_normal']):g}", f"{float(f['precio_domingo']):g}",
                f"{float(f['precio_feriado']):g}"))

    def agregar(self):
        dlg = DialogoRangoDistancia(self)
        if dlg.result:
            self.filas.append({"id": None, **dlg.result})
            self.pintar()

    def editar(self):
        sel = self.tabla.selection()
        if not sel:
            messagebox.showinfo("Seleccione un rango", "Seleccione un rango de la lista.", parent=self)
            return
        idx = self.tabla.index(sel[0])
        dlg = DialogoRangoDistancia(self, self.filas[idx])
        if dlg.result is not None:
            self.filas[idx] = {"id": self.filas[idx]["id"], **dlg.result}
            self.pintar()

    def quitar(self):
        sel = self.tabla.selection()
        if not sel:
            messagebox.showinfo("Seleccione un rango", "Seleccione un rango de la lista.", parent=self)
            return
        idx = self.tabla.index(sel[0])
        if messagebox.askyesno("Quitar rango", f"¿Quitar el rango de {self.filas[idx]['distancia_desde']:g} a {self.filas[idx]['distancia_hasta']:g} km?", parent=self):
            self.filas.pop(idx)
            self.pintar()

    def guardar(self):
        conn = conectar_db()
        if not conn:
            messagebox.showwarning("Modo Lectura", "Sin conexión. No se puede guardar.", parent=self)
            return
        try:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM precios_viaje_distancia WHERE id_cliente=%s", (self.id_cliente,))
            for f in self.filas:
                cursor.execute('''
                    INSERT INTO precios_viaje_distancia
                    (id_cliente, distancia_desde, distancia_hasta, precio_normal, precio_domingo, precio_feriado)
                    VALUES (%s,%s,%s,%s,%s,%s)
                ''', (self.id_cliente, f["distancia_desde"], f["distancia_hasta"],
                      f["precio_normal"], f["precio_domingo"], f["precio_feriado"]))
            conn.commit()
            registrar_auditoria(self.app.usuario_activo, "Cobranza",
                                f"Configuró la tabla de precios por distancia del cliente '{self.nombre_cliente}'")
            messagebox.showinfo("Guardado", "Tabla de precios por distancia guardada.", parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar:\n{e}", parent=self)
        finally:
            liberar_conexion(conn)
        self.app.cargar_rangos_viaje()
        self.destroy()


# =========================================================
# 🚀 CALENDARIO AMIGABLE PARA ELEGIR FECHA
# =========================================================
class CalendarioPopup(ctk.CTkToplevel):
    def __init__(self, parent, entry, on_change=None, anio=None, mes=None, dias_resaltados=None):
        super().__init__(parent)
        self.entry = entry
        self.on_change = on_change
        familia_fuente = "Helvetica" if sys.platform == "darwin" else "Arial"
        self.title("Calendario")
        self.geometry("340x420")
        self.resizable(False, False)
        self.transient(parent)
        try:
            self.grab_set()
        except Exception:
            pass
        self.update_idletasks()
        try:
            x = parent.winfo_rootx() + (parent.winfo_width() // 2) - 170
            y = parent.winfo_rooty() + (parent.winfo_height() // 2) - 210
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

        hoy = datetime.now()
        # Abre en el mes/año seleccionado en el formulario (si se indican)
        self.anio = anio if anio else hoy.year
        self.mes = mes if mes else hoy.month
        self.dias_resaltados = dias_resaltados or {}  # {dia: categoria}

        f_nav = ctk.CTkFrame(self, fg_color="transparent")
        f_nav.pack(fill="x", padx=10, pady=(10, 2))
        ctk.CTkButton(f_nav, text="◀", width=38, command=self._mes_anterior).pack(side="left", padx=2)
        self.lbl_mes = ctk.CTkLabel(f_nav, text="", font=(familia_fuente, 14, "bold"), text_color="#1f538d")
        self.lbl_mes.pack(side="left", expand=True)
        ctk.CTkButton(f_nav, text="▶", width=38, command=self._mes_siguiente).pack(side="left", padx=2)

        self.frame_dias = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_dias.pack(fill="both", expand=True, padx=10, pady=4)
        for i, nombre in enumerate(["Lu", "Ma", "Mi", "Ju", "Vi", "Sá", "Do"]):
            ctk.CTkLabel(self.frame_dias, text=nombre, font=(familia_fuente, 9, "bold"),
                         text_color="#7f8c8d").grid(row=0, column=i, padx=1, pady=2)

        # Leyenda de la quincena
        f_ley = ctk.CTkFrame(self, fg_color="transparent")
        f_ley.pack(fill="x", padx=12, pady=(0, 2))
        ctk.CTkLabel(f_ley, text="Quincena:", font=(familia_fuente, 9, "bold"),
                     text_color="#7f8c8d").pack(side="left", padx=(0, 4))
        for color, texto in (("#d0e4f7", "Lun–Sáb"), ("#fdebd0", "Domingo"), ("#f8d7da", "Feriado")):
            ctk.CTkLabel(f_ley, text="  ", width=20, height=14, fg_color=color,
                         corner_radius=2).pack(side="left", padx=(2, 2))
            ctk.CTkLabel(f_ley, text=texto, font=(familia_fuente, 9),
                         text_color="#555555").pack(side="left", padx=(0, 6))

        btn_hoy = ctk.CTkButton(self, text="Hoy", width=90, font=(familia_fuente, 11, "bold"),
                                fg_color="#34495e", hover_color="#2c3e50", command=self._ir_hoy)
        btn_hoy.pack(pady=(2, 10))
        self._pintar_mes()

    def _pintar_mes(self):
        for w in self.frame_dias.grid_slaves():
            try:
                if int(w.grid_info()["row"]) > 0:
                    w.destroy()
            except Exception:
                pass
        self.lbl_mes.configure(text=f"{NOMBRES_MESES[self.mes-1]} {self.anio}")
        primer = date(self.anio, self.mes, 1)
        ultimo = calendar.monthrange(self.anio, self.mes)[1]
        col_inicio = primer.weekday()  # 0 = lunes
        fila = 1
        for dia in range(1, ultimo + 1):
            col = (col_inicio + dia - 1) % 7
            if dia > 1 and col == 0:
                fila += 1
            # Resalta los días de la quincena según su tipo de día
            cat = self.dias_resaltados.get(dia)
            if cat == "feriado":
                fg_color, hover_color = "#f8d7da", "#e6b8ba"
            elif cat == "domingo":
                fg_color, hover_color = "#fdebd0", "#f2d3a3"
            elif cat == "normal":
                fg_color, hover_color = "#d0e4f7", "#b5d2ee"
            else:
                fg_color, hover_color = "#ffffff", "#d0e4f7"
            btn = ctk.CTkButton(self.frame_dias, text=str(dia), width=36, height=30,
                                font=("Helvetica" if sys.platform == "darwin" else "Arial", 10),
                                fg_color=fg_color, text_color="#000000", hover_color=hover_color,
                                command=lambda d=dia: self._elegir(d))
            btn.grid(row=fila, column=col, padx=1, pady=1)

    def _mes_anterior(self):
        if self.mes == 1:
            self.mes = 12
            self.anio -= 1
        else:
            self.mes -= 1
        self._pintar_mes()

    def _mes_siguiente(self):
        if self.mes == 12:
            self.mes = 1
            self.anio += 1
        else:
            self.mes += 1
        self._pintar_mes()

    def _ir_hoy(self):
        hoy = datetime.now()
        self.anio, self.mes = hoy.year, hoy.month
        self._pintar_mes()
        if hoy.day in self.dias_resaltados:
            self._elegir(hoy.day)

    def _elegir(self, dia):
        try:
            self.entry.delete(0, tk.END)
            self.entry.insert(0, date(self.anio, self.mes, dia).strftime("%d/%m/%Y"))
        except Exception:
            pass
        if self.on_change:
            try:
                self.on_change()
            except Exception:
                pass
        self.destroy()


# =========================================================
# 🚀 VENTANA DE REGISTROS DE QUINCENAS (Editar / Eliminar / PDF)
# =========================================================
class VentanaRegistrosCobranza(ctk.CTkToplevel):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.title("Registros de Cobranza Guardados")
        self.geometry("980x520")
        self.transient(parent)
        try:
            self.grab_set()
        except Exception:
            pass
        self.update_idletasks()
        try:
            x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (980 // 2)
            y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (520 // 2)
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

        familia_fuente = "Helvetica" if sys.platform == "darwin" else "Arial"

        f_top = ctk.CTkFrame(self, fg_color="transparent")
        f_top.pack(fill="x", padx=15, pady=(12, 8))
        ctk.CTkLabel(f_top, text="📋 REGISTROS DE COBRANZA QUINCENAL",
                     font=(familia_fuente, 16, "bold"), text_color="#1f538d").pack(side="left")
        btn_ref = ctk.CTkButton(f_top, text="🔄 Actualizar", width=110,
                                font=(familia_fuente, 11, "bold"), fg_color="#7f8c8d",
                                hover_color="#606b6b", command=self.cargar_registros)
        btn_ref.pack(side="right")

        f_tbl = ctk.CTkFrame(self, corner_radius=10)
        f_tbl.pack(fill="both", expand=True, padx=15, pady=5)

        columnas = ("id", "cliente", "ruc", "periodo", "quincena", "plan", "total", "registro")
        style = ttk.Style()
        if sys.platform == "darwin":
            style.theme_use("clam")
        style.configure("Treeview", background="#ffffff", foreground="#000000",
                        fieldbackground="#ffffff", rowheight=26, font=(familia_fuente, 10),
                        bordercolor="#e0e0e0", borderwidth=1)
        style.map("Treeview", background=[("selected", "#1f538d")],
                  foreground=[("selected", "#ffffff")])
        style.configure("Treeview.Heading", background="#f0f0f0", foreground="#000000",
                        font=(familia_fuente, 10, "bold"), bordercolor="#e0e0e0", borderwidth=1, relief="flat")

        self.tabla = ttk.Treeview(f_tbl, columns=columnas, show="headings", selectmode="browse",
                                  style="Treeview")
        self.tabla.heading("id", text="N°", anchor="center")
        self.tabla.heading("cliente", text="Cliente", anchor="center")
        self.tabla.heading("ruc", text="RUC", anchor="center")
        self.tabla.heading("periodo", text="Mes / Año", anchor="center")
        self.tabla.heading("quincena", text="Quincena", anchor="center")
        self.tabla.heading("plan", text="Plan", anchor="center")
        self.tabla.heading("total", text="Total", anchor="center")
        self.tabla.heading("registro", text="Registrado", anchor="center")
        self.tabla.column("id", width=45, anchor="center")
        self.tabla.column("cliente", width=240, anchor="w")
        self.tabla.column("ruc", width=110, anchor="center")
        self.tabla.column("periodo", width=110, anchor="center")
        self.tabla.column("quincena", width=90, anchor="center")
        self.tabla.column("plan", width=110, anchor="center")
        self.tabla.column("total", width=110, anchor="center")
        self.tabla.column("registro", width=140, anchor="center")
        self.tabla.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        scr = ttk.Scrollbar(f_tbl, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scr.set)
        scr.pack(side="right", fill="y", pady=10, padx=(0, 10))
        self.tabla.bind("<Double-1>", lambda e: self.editar_registro())

        f_btns = ctk.CTkFrame(self, fg_color="transparent")
        f_btns.pack(fill="x", padx=15, pady=10)
        btn_edit = ctk.CTkButton(f_btns, text="✏️ Editar", width=140, font=(familia_fuente, 12, "bold"),
                                 fg_color="#34495e", hover_color="#2c3e50", command=self.editar_registro)
        btn_edit.pack(side="left", padx=5)
        btn_del = ctk.CTkButton(f_btns, text="❌ Eliminar", width=140, font=(familia_fuente, 12, "bold"),
                                fg_color="#e74c3c", hover_color="#c0392b", command=self.eliminar_registro)
        btn_del.pack(side="left", padx=5)
        btn_pdf = ctk.CTkButton(f_btns, text="📄 Ver / Generar PDF", width=170, font=(familia_fuente, 12, "bold"),
                                fg_color="#c0392b", hover_color="#922b21", command=self.ver_pdf)
        btn_pdf.pack(side="left", padx=5)
        btn_cerrar = ctk.CTkButton(f_btns, text="✖ Cerrar", width=110, font=(familia_fuente, 12, "bold"),
                                   fg_color="#7f8c8d", hover_color="#606b6b", command=self.destroy)
        btn_cerrar.pack(side="right", padx=5)

        self.cargar_registros()

    def cargar_registros(self):
        for item in self.tabla.get_children():
            self.tabla.delete(item)
        conn = conectar_db(silencioso=True)
        if not conn:
            messagebox.showwarning("Modo Lectura", "Sin conexión a la base de datos.", parent=self)
            return
        try:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT cc.id, cc.cliente_nombre, cc.cliente_ruc, cc.anio, cc.mes, cc.quincena,
                       cc.plan_cobro, cc.total, cc.fecha_registro, cc.pdf_ruta
                FROM cobranza_quincenas cc ORDER BY cc.id DESC
            ''')
            for r in cursor.fetchall():
                (id_rec, nom, ruc, anio, mes, quincena, plan, total, fecha_reg, pdf) = r
                periodo = f"{NOMBRES_MESES[mes-1]} {anio}" if mes else ""
                q_txt = "1ª (1-15)" if quincena == 1 else "2ª (16-fin)"
                self.tabla.insert("", tk.END, values=(
                    id_rec, nom, ruc, periodo, q_txt, plan, formatear_moneda(total),
                    fecha_reg or "", "✔ PDF" if pdf else "—"
                ))
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar los registros:\n{e}", parent=self)
        finally:
            liberar_conexion(conn)

    def _id_seleccionado(self):
        sel = self.tabla.selection()
        if not sel:
            messagebox.showinfo("Seleccione un registro", "Seleccione una quincena de la lista.", parent=self)
            return None
        return int(self.tabla.item(sel[0], "values")[0])

    def editar_registro(self):
        id_rec = self._id_seleccionado()
        if id_rec is None:
            return
        self.destroy()
        self.app.cargar_registro(id_rec)

    def eliminar_registro(self):
        id_rec = self._id_seleccionado()
        if id_rec is None:
            return
        valores = self.tabla.item(self.tabla.selection()[0], "values")
        if not messagebox.askyesno("Confirmar Eliminación",
                                   f"¿Eliminar la quincena N° {id_rec} del cliente '{valores[1]}' ({valores[3]})?", parent=self):
            return
        conn = conectar_db()
        if not conn:
            messagebox.showwarning("Modo Lectura", "Sin conexión. No se puede eliminar.", parent=self)
            return
        pdf_ruta = ""
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT pdf_ruta FROM cobranza_quincenas WHERE id=%s", (id_rec,))
            fila = cursor.fetchone()
            if fila:
                pdf_ruta = fila[0] or ""
            cursor.execute("DELETE FROM cobranza_detalle_dias WHERE id_cobranza=%s", (id_rec,))
            cursor.execute("DELETE FROM cobranza_quincena_unidades WHERE id_cobranza=%s", (id_rec,))
            cursor.execute("DELETE FROM cobranza_quincena_viajes WHERE id_cobranza=%s", (id_rec,))
            cursor.execute("DELETE FROM cobranza_viajes_detalle WHERE id_cobranza=%s", (id_rec,))
            cursor.execute("DELETE FROM cobranza_deducciones_detalle WHERE id_cobranza=%s", (id_rec,))
            cursor.execute("DELETE FROM cobranza_quincenas WHERE id=%s", (id_rec,))
            conn.commit()
            registrar_auditoria(self.app.usuario_activo, "Cobranza", f"Eliminó el cálculo de cobranza N° {id_rec}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo eliminar:\n{e}", parent=self)
        finally:
            liberar_conexion(conn)

        from app_paths import resolver_ruta_archivo, eliminar_archivo
        pdf_ruta_real = resolver_ruta_archivo(pdf_ruta)
        if pdf_ruta_real:
            if messagebox.askyesno("Archivo PDF", "¿Eliminar también el archivo PDF del disco?", parent=self):
                try:
                    eliminar_archivo(pdf_ruta)
                except Exception:
                    pass
        self.cargar_registros()

    def ver_pdf(self):
        id_rec = self._id_seleccionado()
        if id_rec is None:
            return
        if not REPORTLAB_DISPONIBLE:
            messagebox.showerror("Falta ReportLab", "Instale las dependencias con: python instalar_dependencias.py", parent=self)
            return
        try:
            ruta = self.app._generar_pdf_desde_registro(id_rec)
        except Exception as e:
            messagebox.showerror("Error al generar el PDF", f"Ocurrió un error inesperado:\n{e}", parent=self)
            return
        if ruta:
            abrir_documento(ruta)


if __name__ == "__main__":
    root = ctk.CTk()
    root.geometry("1000x760")
    app = CalculoCobranzaApp(root)
    root.mainloop()