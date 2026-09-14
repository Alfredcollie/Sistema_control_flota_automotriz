# -*- coding: utf-8 -*-
"""
MODULO_BANCO.PY - MODULO DE BANCO (SALDOS Y CONCILIACION BANCARIA)
- Saldo de cada cuenta bancaria configurada en "Configuracion General":
    Saldo = Saldo Inicial + Ingresos Cobrados - Egresos Pagados.
- Conciliacion bancaria: subir el PDF del estado de cuenta del banco,
  extraer sus movimientos, compararlos contra los movimientos registrados
  en el sistema (cobros y pagos) y generar un reporte en pantalla.
- Permite conciliar, corregir montos/fechas y agregar movimientos faltantes
  (tanto ajustes manuales como cobros/pagos reales en el sistema).
- Registro unificado de pagos ("Registrar Pago / Compra Cruzada"): al marcar el
  check "Es compra cruzada" el pago a un tercero se registra tambien como factura
  pagada por tercero en el modulo de Compras, adjuntando el PDF del soporte.
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import customtkinter as ctk
import os
import re
import json
import shutil
import calendar
import sys
import subprocess
import threading
from datetime import datetime

from conexion import conectar_db, registrar_auditoria, liberar_conexion
from dialogos_seguros import (seleccionar_archivo_dialogo, seleccionar_archivos_dialogo,
                              guardar_archivo_dialogo)
from app_paths import CONFIG_FILE, ruta_para_guardar
from config_nube import cargar_bancos

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    from pypdf import PdfReader as _PypdfReader
except ImportError:
    try:
        from PyPDF2 import PdfReader as _PypdfReader
    except ImportError:
        _PypdfReader = None

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")


# =========================================================
# CONFIGURACION Y UTILIDADES
# =========================================================
def cargar_config():
    config = {
        "simbolo_moneda": "S/.",
        "formato_numero": "1,000.00",
        "formato_fecha": "DD/MM/AAAA",
        "cuentas_bancarias": [],
    }
    try:
        if os.path.exists(str(CONFIG_FILE)):
            with open(str(CONFIG_FILE), "r", encoding="utf-8") as f:
                config.update(json.load(f))
    except Exception:
        pass
    return config


CONFIG = cargar_config()


def obtener_ruta_base():
    """Carpeta de archivos del programa (misma lógica que Compras/Ventas).

    Devuelve "" si el equipo NO está autorizado (equipo secundario sin la cuenta
    Rclone del principal): así ningún flujo del Banco guarda archivos.
    """
    try:
        from politica_almacenamiento import estado_almacenamiento, ruta_base_autorizada
        if not estado_almacenamiento().get("autorizado"):
            return ""
        ruta = ruta_base_autorizada(mostrar_alerta=False)
        if ruta and os.path.isdir(ruta):
            return ruta
    except Exception:
        pass
    ruta = CONFIG.get("ruta_drive", "").strip()
    if ruta:
        ruta = os.path.expanduser(ruta)
        if os.path.isdir(ruta):
            return ruta
    # Respaldo junto al programa: SOLO en modo monousuario (local)
    try:
        from politica_almacenamiento import permitir_respaldo_local
        if not permitir_respaldo_local():
            return ""
    except Exception:
        pass
    return os.path.dirname(os.path.abspath(__file__))


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
            subprocess.call(["open", ruta_abs])
        else:
            subprocess.call(["xdg-open", ruta_abs])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")


CATEGORIAS_GASTOS_DEFAULT = {
    "Gastos Fijos": ["Alquiler", "Planilla / Sueldos", "Servicios (Luz, Agua, Internet)", "Seguros", "Préstamos / Cuotas"],
    "Gastos Operativos": ["Combustible", "Mantenimiento", "Repuestos", "Peajes", "Otros operativos"],
    "Ingresos": ["Cobro de cliente", "Otro ingreso"],
    "Otros": ["Varios"],
}

COMISION_INTERBANCARIA_DEFAULT = "4.80"


def leer_config_disco():
    cfg = {}
    try:
        if os.path.exists(str(CONFIG_FILE)):
            with open(str(CONFIG_FILE), "r", encoding="utf-8") as f:
                cfg = json.load(f)
    except Exception:
        pass
    return cfg


def guardar_config_disco(cfg):
    try:
        with open(str(CONFIG_FILE), "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=4)
        return True
    except Exception:
        return False


def cargar_categorias_gastos():
    cfg = leer_config_disco()
    if "categorias_gastos" in cfg and isinstance(cfg["categorias_gastos"], dict):
        cats = cfg["categorias_gastos"]
        return {k: (v if isinstance(v, list) else []) for k, v in cats.items()}
    return dict(CATEGORIAS_GASTOS_DEFAULT)


def guardar_categorias_gastos(cats):
    cfg = leer_config_disco()
    cfg["categorias_gastos"] = cats
    return guardar_config_disco(cfg)


def cargar_comision_interbancaria():
    v = leer_config_disco().get("comision_interbancaria", COMISION_INTERBANCARIA_DEFAULT)
    try:
        return float(str(v).replace(",", "."))
    except Exception:
        return 4.8


def guardar_comision_interbancaria(valor):
    cfg = leer_config_disco()
    cfg["comision_interbancaria"] = valor
    return guardar_config_disco(cfg)


def cargar_placas_flota():
    """Placas de vehículos registrados en Flota Automotriz."""
    conn = conectar_db(silencioso=True)
    if not conn:
        return []
    placas = []
    try:
        with conn.cursor() as c:
            c.execute("SELECT placa FROM flota_vehiculos WHERE placa IS NOT NULL AND TRIM(placa) != '' ORDER BY placa")
            placas = [r[0] for r in c.fetchall()]
    except Exception:
        pass
    finally:
        liberar_conexion(conn)
    return placas


def es_categoria_planilla(texto):
    """True si el texto de categoría corresponde a planilla / sueldos."""
    t = (texto or "").lower()
    return "planilla" in t or "sueldo" in t


def cargar_choferes():
    """Choferes registrados en el módulo de Choferes (activos primero)."""
    conn = conectar_db(silencioso=True)
    if not conn:
        return []
    try:
        with conn.cursor() as c:
            c.execute("""
                SELECT nombres FROM choferes
                WHERE nombres IS NOT NULL AND TRIM(nombres) != ''
                ORDER BY (COALESCE(estado, '') = 'Activo') DESC, nombres ASC
            """)
            return [str(r[0]) for r in c.fetchall()]
    except Exception:
        return []
    finally:
        liberar_conexion(conn)


def cargar_proveedores():
    """Proveedores registrados en el módulo de Proveedores."""
    conn = conectar_db(silencioso=True)
    if not conn:
        return []
    try:
        with conn.cursor() as c:
            c.execute("SELECT nombre FROM proveedores WHERE nombre IS NOT NULL AND TRIM(nombre) != '' ORDER BY nombre ASC")
            return [str(r[0]) for r in c.fetchall()]
    except Exception:
        return []
    finally:
        liberar_conexion(conn)


def formatear_numero_entrada(texto):
    """Formatea un monto con separador de miles y decimal según la configuración,
    sin símbolo de moneda, para mostrarlo mientras el usuario escribe."""
    formato = CONFIG.get("formato_numero", "1,000.00")
    es_latino = (formato == "1.000,00")
    sep_dec = "," if es_latino else "."
    s = (texto or "").strip()
    if not s:
        return ""
    m = re.search(r"[.,](\d{1,2})$", s)
    termina_sep = bool(re.search(r"[.,]$", s))
    if m:
        parte_dec = m.group(1)
        base = s[:m.start()]
    elif termina_sep:
        parte_dec = None
        base = s[:-1]
    else:
        parte_dec = None
        base = s
    digitos = re.sub(r"\D", "", base)
    if not digitos and parte_dec is None:
        return ""
    entero = int(digitos) if digitos else 0
    ent_fmt = f"{entero:,}"
    if es_latino:
        ent_fmt = ent_fmt.replace(",", ".")
    if m:
        return f"{ent_fmt}{sep_dec}{parte_dec}"
    if termina_sep:
        return f"{ent_fmt}{sep_dec}"
    return ent_fmt


def normalizar_monto(v):
    """Convierte cualquier texto de monto a float, aceptando ambos formatos
    (1,234.56 y 1.234,56) y simbolos de moneda."""
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v or "").strip()
    if not s:
        return 0.0
    s = re.sub(r"[^\d.,+-]", "", s)
    if not s:
        return 0.0
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        if re.fullmatch(r"[+-]?\d{1,3}(,\d{3})+", s):
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        dig = re.sub(r"[^\d.]", "", s)
        try:
            return float(dig)
        except ValueError:
            return 0.0


def monto_desde_texto(texto):
    """Extrae el primer monto que aparezca en un texto (ej. 'Comisión S/ 12.00')."""
    coincidencia = re.search(r"\d+(?:[.,]\d+)?", str(texto or ""))
    return normalizar_monto(coincidencia.group(0)) if coincidencia else 0.0


def formatear_monto(valor):
    simbolo = CONFIG.get("simbolo_moneda", "S/.")
    formato = CONFIG.get("formato_numero", "1,000.00")
    try:
        valor = float(valor)
    except Exception:
        valor = 0.0
    signo = "-" if valor < 0 else ""
    valor = abs(valor)
    if formato == "1.000,00":
        txt = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    else:
        txt = f"{valor:,.2f}"
    return f"{signo}{simbolo} {txt}"


def normalizar_fecha(s):
    s = (s or "").strip()
    if not s:
        return ""
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y", "%Y%m%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except Exception:
            continue
    return s


def construir_etiqueta_banco(b):
    return f"{b.get('banco', '')} - {b.get('cuenta', '')}".strip(" -")


def _norm_texto(s):
    return re.sub(r"\s+", " ", (s or "")).strip().lower()


def banco_matchea(banco, texto_cuenta):
    """Determina si un texto de cuenta (cuenta_destino / cuenta_origen)
    corresponde al banco indicado."""
    t = _norm_texto(texto_cuenta)
    if not t:
        return False
    nombre = _norm_texto(banco.get("banco", ""))
    cuenta = _norm_texto(banco.get("cuenta", ""))
    etiqueta = _norm_texto(construir_etiqueta_banco(banco))
    if t == etiqueta and etiqueta:
        return True
    if nombre and nombre in t:
        return True
    if cuenta:
        digitos_cuenta = re.sub(r"\D", "", cuenta)
        digitos_t = re.sub(r"\D", "", t)
        if len(digitos_cuenta) >= 4 and digitos_cuenta in digitos_t:
            return True
    return False


def aplicar_estilo_treeview():
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("Treeview", background="#ffffff", foreground="#000000",
                    fieldbackground="#ffffff", bordercolor="#e0e0e0",
                    borderwidth=1, rowheight=26, font=("Arial", 10))
    style.map("Treeview", background=[("selected", "#1f538d")],
              foreground=[("selected", "#ffffff")])
    style.configure("Treeview.Heading", background="#f0f0f0", foreground="#000000",
                    relief="flat", font=("Arial", 10, "bold"),
                    bordercolor="#e0e0e0", borderwidth=1)


# =========================================================
# EXTRACCION Y ANALISIS DE PDF (ESTADO DE CUENTA)
# =========================================================
def extraer_texto_pdf(ruta):
    texto = ""
    if pdfplumber is not None:
        try:
            with pdfplumber.open(ruta) as pdf:
                partes = [p.extract_text() or "" for p in pdf.pages]
            texto = "\n".join(partes).strip()
        except Exception:
            texto = ""
    if not texto and fitz is not None:
        try:
            doc = fitz.open(ruta)
            partes = [page.get_text() or "" for page in doc]
            doc.close()
            texto = "\n".join(partes).strip()
        except Exception:
            texto = ""
    if not texto and _PypdfReader is not None:
        try:
            reader = _PypdfReader(ruta)
            partes = [(p.extract_text() or "") for p in reader.pages]
            texto = "\n".join(partes).strip()
        except Exception:
            texto = ""
    return texto


FECHA_RE = re.compile(r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b")
MONTO_RE = re.compile(r"[-+]?\s?\d{1,3}(?:[.,]\d{3})+(?:[.,]\d{1,2})?|[-+]?\s?\d+(?:[.,]\d{1,2})?")

PALABRAS_ABONO = ["abono", "abon", "deposito", "depósito", "credito", "crédito", "ingreso",
                  "entrada", "transferencia recibida", "recibido", "cobro", "cobranza", "dep"]
PALABRAS_CARGO = ["cargo", "retiro", "debito", "débito", "pago", "egreso", "salida",
                  "transferencia enviada", "compra", "ret", "cargos"]


def _signo_por_palabras(linea):
    low = linea.lower()
    for p in PALABRAS_ABONO:
        if p in low:
            return 1
    for p in PALABRAS_CARGO:
        if p in low:
            return -1
    return None


def _seleccionar_monto(linea, montos):
    """montos: lista de (texto, posicion). Devuelve (valor, token) del movimiento."""
    if not montos:
        return None
    if len(montos) >= 2:
        candidatos = montos[:-1]
        vals = []
        for tok, pos in candidatos:
            v = normalizar_monto(tok)
            vals.append((v, tok))
        no_cero = [x for x in vals if abs(x[0]) > 0.0009]
        if no_cero:
            return max(no_cero, key=lambda x: abs(x[0]))
        return vals[-1]
    tok, pos = montos[0]
    return normalizar_monto(tok), tok


def parsear_movimientos(texto):
    """Heuristica para extraer movimientos de un estado de cuenta."""
    movimientos = []
    if not texto:
        return movimientos
    for linea in texto.splitlines():
        linea = linea.strip()
        if not linea:
            continue
        low = linea.lower()
        # Encabezados / rangos de periodo no son transacciones
        if any(k in low for k in ("periodo", "period", "estado de cuenta", "saldo inicial")):
            continue
        fechas = FECHA_RE.findall(linea)
        if len(fechas) >= 2:
            continue
        mf = FECHA_RE.search(linea)
        if not mf:
            continue
        fecha = mf.group(1)
        resto = linea[:mf.start()] + " " + linea[mf.end():]
        montos = [(m.group(0), m.start()) for m in MONTO_RE.finditer(resto)]
        if not montos:
            continue
        seleccion = _seleccionar_monto(resto, montos)
        if seleccion is None:
            continue
        valor, token = seleccion
        signo_palabras = _signo_por_palabras(linea)
        signo_explicito = token.strip().startswith("+") or token.strip().startswith("-")
        if not signo_explicito and signo_palabras is not None:
            valor = signo_palabras * abs(valor)
        descripcion = re.sub(r"\s+", " ", resto).strip()
        descripcion = MONTO_RE.sub("", descripcion)
        descripcion = re.sub(r"\s+", " ", descripcion).strip(" -|")
        movimientos.append({
            "fecha": normalizar_fecha(fecha) or fecha,
            "descripcion": descripcion or "Movimiento sin descripción",
            "monto": valor,
        })
    return movimientos


def detectar_saldo_final(texto):
    """Intenta detectar el saldo final del estado de cuenta."""
    if not texto:
        return None
    lineas_saldo = [ln for ln in texto.splitlines() if "saldo" in ln.lower()]
    if not lineas_saldo:
        return None
    ultima = lineas_saldo[-1]
    montos = [(m.group(0), m.start()) for m in MONTO_RE.finditer(ultima)]
    if not montos:
        return None
    return normalizar_monto(montos[-1][0])


def extraer_datos_factura_pdf(ruta):
    """Extrae datos básicos de una factura de compra (SUNAT) desde un PDF."""
    datos = {
        "tipo_documento": "FACTURA",
        "numero_documento": "",
        "fecha": datetime.now().strftime("%d/%m/%Y"),
        "proveedor": "",
        "ruc": "",
        "total": 0.0,
    }
    texto = extraer_texto_pdf(ruta)
    if not texto:
        return datos
    if re.search(r"BOLETA\s+DE\s+VENTA", texto, re.IGNORECASE):
        datos["tipo_documento"] = "BOLETA"
    elif re.search(r"RECIBO\s+POR\s+HONORARIOS", texto, re.IGNORECASE):
        datos["tipo_documento"] = "RECIBO"
    m = re.search(r"([EFB][0-9A-Z]{3}\s*-\s*\d+)", texto)
    if m:
        datos["numero_documento"] = m.group(1).replace(" ", "")
    m = re.search(r"Fecha de Emisi[oó]n\s*[:\-]?\s*(\d{2})[/\-.](\d{2})[/\-.](\d{4})", texto, re.IGNORECASE)
    if not m:
        m = re.search(r"(\d{2})[/\-.](\d{2})[/\-.](\d{4})", texto)
    if m:
        datos["fecha"] = f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
    rucs = re.findall(r"(?:RUC|R\.U\.C\.)\s*[:\-]?\s*(\d{11})", texto, re.IGNORECASE)
    if rucs:
        datos["ruc"] = rucs[0]
    lineas = [re.sub(r"\s+", " ", l).strip() for l in texto.splitlines() if l.strip()]
    for l in lineas[:12]:
        if re.search(r"S\.A\.C\.|S\.A\.|E\.I\.R\.L\.|S\.R\.L\.|SAC|SRL|S\.C\.R\.L", l, re.IGNORECASE):
            datos["proveedor"] = l
            break
    if not datos["proveedor"]:
        for l in lineas[:8]:
            if re.search(r"R\.?U\.?C|DIRECCI|TEL[ÉE]FONO|CIUDAD|SE[ÑN]OR", l, re.IGNORECASE) or re.search(r"\d{11}", l):
                continue
            if len(l) > 4:
                datos["proveedor"] = l
                break
    m = re.search(r"(?:IMPORTE\s*TOTAL|TOTAL\s*A\s*PAGAR|Total Neto Recibido|TOTAL)[\s:S/\|]+([\d,\.]+)", texto, re.IGNORECASE)
    if m:
        datos["total"] = normalizar_monto(m.group(1))
    return datos


NOMBRES_MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                 "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]


def construir_valores_mes():
    hoy = datetime.now()
    valores = ["Todos los meses"]
    for anio in (hoy.year - 1, hoy.year, hoy.year + 1):
        for nombre in NOMBRES_MESES:
            valores.append(f"{nombre} {anio}")
    return valores


def mes_actual_etiqueta():
    hoy = datetime.now()
    return f"{NOMBRES_MESES[hoy.month - 1]} {hoy.year}"


def mes_etiqueta_a_key(etiqueta):
    if not etiqueta or etiqueta == "Todos los meses":
        return None
    partes = (etiqueta or "").strip().split()
    if len(partes) < 2 or partes[0] not in NOMBRES_MESES:
        return None
    mes = NOMBRES_MESES.index(partes[0]) + 1
    anio = partes[-1]
    return f"{anio}-{mes:02d}"


# =========================================================
# CLASE PRINCIPAL
# =========================================================
def centrar_ventana(ventana, parent, ancho, alto):
    """Coloca una ventana secundaria centrada sobre su ventana padre."""
    ventana.update_idletasks()
    try:
        if parent and parent.winfo_ismapped():
            x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (ancho // 2)
            y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (alto // 2)
        else:
            x = (ventana.winfo_screenwidth() // 2) - (ancho // 2)
            y = (ventana.winfo_screenheight() // 2) - (alto // 2)
    except Exception:
        x = (ventana.winfo_screenwidth() // 2) - (ancho // 2)
        y = (ventana.winfo_screenheight() // 2) - (alto // 2)
    ventana.geometry(f"{ancho}x{alto}+{max(10, x)}+{max(35, y)}")


class CalendarioNativo(ctk.CTkToplevel):
    """Calendario emergente para escribir la fecha en un campo de texto.

    Se abre en el mes de la fecha que ya tenga el campo (o en el mes actual),
    resalta el día de hoy y trae un botón "Hoy" para ir a la fecha de hoy."""

    def __init__(self, parent, target_entry):
        super().__init__(parent)
        self.target_entry = target_entry
        self.title("Seleccionar Fecha")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        hoy = datetime.now()
        self.current_year = hoy.year
        self.current_month = hoy.month
        self.dia_marcado = None

        # Si el campo ya trae una fecha válida, se abre en ese mes y se resalta ese día
        texto = ""
        try:
            texto = (target_entry.get() or "").strip()
        except Exception:
            texto = ""
        m = re.match(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})$", texto)
        if m:
            d, mes, anio = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if anio < 100:
                anio += 2000
            if 1 <= mes <= 12 and 1 <= d <= 31:
                self.current_month, self.current_year = mes, anio
                self.dia_marcado = d

        centrar_ventana(self, parent, 330, 350)

        header = ctk.CTkFrame(self, fg_color="#1f538d", corner_radius=0)
        header.pack(fill="x")

        ctk.CTkButton(header, text="<", width=28, fg_color="transparent", text_color="white",
                      hover_color="#163b65", font=("Arial", 14, "bold"),
                      command=self.mes_anterior).pack(side="left", padx=5, pady=10)

        self.cmb_mes = ctk.CTkComboBox(header, values=NOMBRES_MESES, width=105,
                                       command=self.cambiar_mes, state="readonly")
        self.cmb_mes.pack(side="left", padx=2, pady=10)

        anios = [str(y) for y in range(hoy.year - 80, hoy.year + 20)]
        self.cmb_anio = ctk.CTkComboBox(header, values=anios, width=80,
                                        command=self.cambiar_anio, state="readonly")
        self.cmb_anio.pack(side="left", padx=2, pady=10)

        ctk.CTkButton(header, text=">", width=28, fg_color="transparent", text_color="white",
                      hover_color="#163b65", font=("Arial", 14, "bold"),
                      command=self.mes_siguiente).pack(side="right", padx=5, pady=10)

        self.days_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.days_frame.pack(fill="both", expand=True, padx=10, pady=(10, 4))

        for i, dia in enumerate(["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]):
            ctk.CTkLabel(self.days_frame, text=dia, font=("Arial", 11, "bold"),
                         text_color="#1f538d").grid(row=0, column=i, padx=5, pady=5)

        f_pie = ctk.CTkFrame(self, fg_color="transparent")
        f_pie.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkButton(f_pie, text="📅 Hoy", width=90, font=("Arial", 11, "bold"),
                      fg_color="#27ae60", hover_color="#1e8449",
                      command=self.ir_a_hoy).pack(side="left")
        ctk.CTkButton(f_pie, text="✖ Cerrar", width=90, font=("Arial", 11),
                      fg_color="#7f8c8d", hover_color="#606b6b",
                      command=self.destroy).pack(side="right")

        self.dibujar_mes()

    # ------------------------------------------------------------------
    def dibujar_mes(self):
        self.cmb_mes.set(NOMBRES_MESES[self.current_month - 1])
        self.cmb_anio.set(str(self.current_year))

        for widget in self.days_frame.winfo_children():
            try:
                if int(widget.grid_info()["row"]) > 0:
                    widget.destroy()
            except Exception:
                pass

        hoy = datetime.now()
        for fila, semana in enumerate(calendar.monthcalendar(self.current_year, self.current_month), start=1):
            for col, dia in enumerate(semana):
                if dia == 0:
                    continue
                es_hoy = (dia == hoy.day and self.current_month == hoy.month
                          and self.current_year == hoy.year)
                es_marcado = (self.dia_marcado == dia)
                if es_marcado:
                    color, texto_color = "#1f538d", "white"
                elif es_hoy:
                    color, texto_color = "#d4edda", "#155724"
                else:
                    color, texto_color = "transparent", "black"
                btn = ctk.CTkButton(self.days_frame, text=str(dia), width=30, height=30,
                                    fg_color=color, text_color=texto_color,
                                    hover_color="#e0e0e0", font=("Arial", 11),
                                    command=lambda d=dia: self.seleccionar(d))
                btn.grid(row=fila, column=col, padx=3, pady=2)

    def cambiar_mes(self, eleccion):
        try:
            self.current_month = NOMBRES_MESES.index(eleccion) + 1
        except ValueError:
            return
        self.dia_marcado = None
        self.dibujar_mes()

    def cambiar_anio(self, eleccion):
        try:
            self.current_year = int(eleccion)
        except (TypeError, ValueError):
            return
        self.dia_marcado = None
        self.dibujar_mes()

    def mes_anterior(self):
        self.current_month -= 1
        if self.current_month < 1:
            self.current_month, self.current_year = 12, self.current_year - 1
        self.dia_marcado = None
        self.dibujar_mes()

    def mes_siguiente(self):
        self.current_month += 1
        if self.current_month > 12:
            self.current_month, self.current_year = 1, self.current_year + 1
        self.dia_marcado = None
        self.dibujar_mes()

    def ir_a_hoy(self):
        hoy = datetime.now()
        self.current_year, self.current_month, self.dia_marcado = hoy.year, hoy.month, hoy.day
        self.dibujar_mes()

    def seleccionar(self, dia):
        try:
            self.target_entry.delete(0, tk.END)
            self.target_entry.insert(0, f"{dia:02d}/{self.current_month:02d}/{self.current_year}")
        except Exception:
            pass
        self.destroy()


class ModuloBancoApp:
    def __init__(self, parent_frame, usuario_activo=""):
        self.parent_frame = parent_frame
        self.usuario_activo = usuario_activo or "Desconocido"
        self.config = cargar_config()
        self.bancos = cargar_bancos()
        self.movimientos_pdf = []
        self.texto_pdf = ""
        self.filas_conciliacion = []
        self.saldo_final_estado = None
        self.banco_conciliado = None
        aplicar_estilo_treeview()
        self.inicializar_db()

        self.frame_main = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.frame_main.pack(fill="both", expand=True, padx=15, pady=15)

        header = ctk.CTkFrame(self.frame_main, fg_color="transparent")
        header.pack(fill="x")
        ctk.CTkLabel(header, text="🏦 MÓDULO DE BANCO (SALDOS Y CONCILIACIÓN BANCARIA)",
                     font=("Arial", 18, "bold"), text_color="#1f538d").pack(side="left")

        self.tabview = ctk.CTkTabview(self.frame_main, segmented_button_selected_color="#1f538d")
        self.tabview.pack(fill="both", expand=True, pady=(10, 0))
        self.tab_saldos = self.tabview.add(" 💵 Saldo de Bancos ")
        self.tab_conciliacion = self.tabview.add(" 🧾 Conciliación Bancaria ")
        self.tab_transferencias = self.tabview.add(" 🔄 Transferencias ")

        self.construir_tab_saldos()
        self.construir_tab_conciliacion()
        self.construir_tab_transferencias()

    # -----------------------------------------------------
    # BASE DE DATOS DE CONCILIACION
    # -----------------------------------------------------
    def inicializar_db(self):
        conn = conectar_db(silencioso=True)
        if not conn:
            return
        try:
            with conn.cursor() as c:
                c.execute("""
                    CREATE TABLE IF NOT EXISTS conciliacion_bancaria (
                        id SERIAL PRIMARY KEY,
                        banco VARCHAR(255),
                        cuenta VARCHAR(255),
                        fecha VARCHAR(20),
                        descripcion TEXT,
                        monto NUMERIC,
                        tipo VARCHAR(20),
                        origen VARCHAR(30),
                        id_movimiento INTEGER DEFAULT 0,
                        estado VARCHAR(20) DEFAULT 'pendiente',
                        fecha_conciliacion VARCHAR(20) DEFAULT ''
                    )
                """)
                c.execute("""
                    CREATE TABLE IF NOT EXISTS transferencias_bancarias (
                        id SERIAL PRIMARY KEY,
                        banco_origen VARCHAR(255),
                        cuenta_origen VARCHAR(255),
                        banco_destino VARCHAR(255),
                        cuenta_destino VARCHAR(255),
                        fecha VARCHAR(20),
                        descripcion TEXT,
                        monto NUMERIC
                    )
                """)
                c.execute("""
                    CREATE TABLE IF NOT EXISTS facturas_recibidas (
                        id SERIAL PRIMARY KEY, tipo_documento VARCHAR(100), fecha VARCHAR(50),
                        proveedor VARCHAR(255), descripcion TEXT, evento_asociado VARCHAR(255),
                        subtotal NUMERIC, impuesto NUMERIC, total NUMERIC, archivo_ruta TEXT,
                        dias_credito INTEGER DEFAULT 0, det_porcentaje NUMERIC DEFAULT 0,
                        det_monto NUMERIC DEFAULT 0, numero_documento VARCHAR(100) DEFAULT '',
                        categoria VARCHAR(255) DEFAULT 'GENERAL / NO ASIGNADO'
                    )
                """)
                conn.commit()
                for col_sql in (
                    "ALTER TABLE facturas_recibidas ADD COLUMN IF NOT EXISTS ruc VARCHAR(50)",
                    "ALTER TABLE facturas_recibidas ADD COLUMN IF NOT EXISTS pagado_por_tercero VARCHAR(255) DEFAULT ''",
                    "ALTER TABLE facturas_recibidas ADD COLUMN IF NOT EXISTS soporte_pago_tercero TEXT DEFAULT ''",
                    "ALTER TABLE facturas_recibidas ADD COLUMN IF NOT EXISTS es_compra_cruzada BOOLEAN DEFAULT FALSE",
                    # Datos del pago manual, para poder reabrirlo y editarlo completo
                    "ALTER TABLE conciliacion_bancaria ADD COLUMN IF NOT EXISTS categoria VARCHAR(120) DEFAULT ''",
                    "ALTER TABLE conciliacion_bancaria ADD COLUMN IF NOT EXISTS subcategoria VARCHAR(120) DEFAULT ''",
                    "ALTER TABLE conciliacion_bancaria ADD COLUMN IF NOT EXISTS placa VARCHAR(80) DEFAULT ''",
                    "ALTER TABLE conciliacion_bancaria ADD COLUMN IF NOT EXISTS chofer VARCHAR(150) DEFAULT ''",
                    "ALTER TABLE conciliacion_bancaria ADD COLUMN IF NOT EXISTS comision NUMERIC DEFAULT 0",
                    "ALTER TABLE conciliacion_bancaria ADD COLUMN IF NOT EXISTS interbancario BOOLEAN DEFAULT FALSE",
                    "ALTER TABLE conciliacion_bancaria ADD COLUMN IF NOT EXISTS es_cruzada BOOLEAN DEFAULT FALSE",
                ):
                    try:
                        c.execute(col_sql)
                        conn.commit()
                    except Exception:
                        conn.rollback()
        except Exception:
            pass
        finally:
            liberar_conexion(conn)

    # -----------------------------------------------------
    # TAB: SALDO DE BANCOS
    # -----------------------------------------------------
    def construir_tab_saldos(self):
        f_top = ctk.CTkFrame(self.tab_saldos, fg_color="transparent")
        f_top.pack(fill="x", pady=(5, 10))
        ctk.CTkLabel(f_top, text="Saldo calculado como: Saldo Inicial + Ingresos Cobrados − Egresos Pagados.",
                     font=("Arial", 12, "italic"), text_color="gray").pack(side="left")
        ctk.CTkButton(f_top, text="🔄 Actualizar", width=130, font=("Arial", 12, "bold"),
                      fg_color="#1f538d", hover_color="#163b65",
                      command=self.refrescar_saldos).pack(side="right")

        self.scroll_saldos = ctk.CTkScrollableFrame(self.tab_saldos, fg_color="transparent")
        self.scroll_saldos.pack(fill="both", expand=True)
        self.refrescar_saldos()

    def refrescar_saldos(self):
        for w in self.scroll_saldos.winfo_children():
            w.destroy()

        if not self.bancos:
            f_aviso = ctk.CTkFrame(self.scroll_saldos, corner_radius=10, fg_color="#fff3cd",
                                   border_width=1, border_color="#ffeeba")
            f_aviso.pack(fill="x", pady=10, padx=5)
            ctk.CTkLabel(f_aviso, text="⚠️ No hay cuentas bancarias configuradas.\n"
                                       "Agréguelas en: Ajustes de Sistema → ⚙️ Configuración General → Cuentas Bancarias.",
                         font=("Arial", 13), text_color="#856404", justify="left").pack(padx=15, pady=15)
            return

        total_general = 0.0
        for idx, banco in enumerate(self.bancos):
            resumen = self.calcular_resumen(banco)
            total_general += resumen["saldo_actual"]
            self._crear_tarjeta_banco(banco, resumen, idx)

        f_total = ctk.CTkFrame(self.scroll_saldos, corner_radius=10, fg_color="#1f538d")
        f_total.pack(fill="x", pady=(15, 5), padx=5)
        ctk.CTkLabel(f_total, text="TOTAL DISPONIBLE EN BANCOS", font=("Arial", 14, "bold"),
                     text_color="white").pack(side="left", padx=15, pady=12)
        ctk.CTkLabel(f_total, text=formatear_monto(total_general), font=("Arial", 16, "bold"),
                     text_color="white").pack(side="right", padx=15, pady=12)

    def _crear_tarjeta_banco(self, banco, resumen, idx):
        etiqueta = construir_etiqueta_banco(banco) or f"Banco {idx + 1}"
        card = ctk.CTkFrame(self.scroll_saldos, corner_radius=10, fg_color="#f8f9fa",
                            border_width=1, border_color="#e0e0e0")
        card.pack(fill="x", pady=6, padx=5)

        fila_titulo = ctk.CTkFrame(card, fg_color="transparent")
        fila_titulo.pack(fill="x", padx=12, pady=(10, 4))
        ctk.CTkLabel(fila_titulo, text=f"🏦 {etiqueta}", font=("Arial", 15, "bold"),
                     text_color="#1f538d").pack(side="left")
        ctk.CTkLabel(fila_titulo, text=resumen["saldo_actual_txt"], font=("Arial", 16, "bold"),
                     text_color="#27ae60" if resumen["saldo_actual"] >= 0 else "#c0392b").pack(side="right")

        fila_det = ctk.CTkFrame(card, fg_color="transparent")
        fila_det.pack(fill="x", padx=12, pady=(0, 6))
        ctk.CTkLabel(fila_det, text=f"Saldo inicial: {formatear_monto(resumen['saldo_inicial'])}",
                     font=("Arial", 11), text_color="gray").pack(side="left", padx=(0, 15))
        ctk.CTkLabel(fila_det, text=f"+ Ingresos: {formatear_monto(resumen['ingresos'])}",
                     font=("Arial", 11), text_color="#27ae60").pack(side="left", padx=(0, 15))
        ctk.CTkLabel(fila_det, text=f"− Egresos: {formatear_monto(resumen['egresos'])}",
                     font=("Arial", 11), text_color="#c0392b").pack(side="left", padx=(0, 15))
        ctk.CTkLabel(fila_det, text=f"({resumen['n_movimientos']} movimientos)",
                     font=("Arial", 10, "italic"), text_color="#7f8c8d").pack(side="left")

        fila_btn = ctk.CTkFrame(card, fg_color="transparent")
        fila_btn.pack(fill="x", padx=12, pady=(0, 10))
        ctk.CTkButton(fila_btn, text="👁 Ver Movimientos", width=160, font=("Arial", 11, "bold"),
                      fg_color="#34495e", hover_color="#2c3e50",
                      command=lambda b=banco: self.ver_movimientos(b)).pack(side="left", padx=(0, 8))
        ctk.CTkButton(fila_btn, text="🧾 Conciliar este Banco", width=180, font=("Arial", 11, "bold"),
                      fg_color="#1f538d", hover_color="#163b65",
                      command=lambda b=banco: self.ir_a_conciliacion(b)).pack(side="left")

    def calcular_resumen(self, banco):
        movs = self.cargar_movimientos_sistema(banco)
        ingresos = sum(m["monto"] for m in movs if m["monto"] > 0)
        egresos = sum(-m["monto"] for m in movs if m["monto"] < 0)
        saldo_inicial = normalizar_monto(banco.get("saldo_inicial", ""))
        saldo = saldo_inicial + ingresos - egresos
        return {
            "saldo_inicial": saldo_inicial,
            "ingresos": ingresos,
            "egresos": egresos,
            "saldo_actual": saldo,
            "saldo_actual_txt": formatear_monto(saldo),
            "n_movimientos": len(movs),
            "movimientos": movs,
        }

    def cargar_movimientos_sistema(self, banco, mes=None):
        """Cobros (ingresos), pagos (egresos) y transferencias registrados para
        este banco. Si se pasa 'mes' (formato 'YYYY-MM'), filtra solo ese mes."""
        conn = conectar_db(silencioso=True)
        if not conn:
            return []
        movs = []
        try:
            with conn.cursor() as c:
                try:
                    c.execute("""
                        SELECT p.id, p.fecha_pago, p.cliente_nombre, p.monto_pagado,
                               COALESCE(p.cuenta_destino, ''), p.id_factura,
                               COALESCE(f.numero_documento, ''), COALESCE(p.codigo_cotizacion, '')
                        FROM pagos_clientes p
                        LEFT JOIN facturas_emitidas f ON p.id_factura = f.id
                    """)
                    for idp, fecha, nombre, monto, cuenta, idf, num_doc, ref in c.fetchall():
                        if banco_matchea(banco, cuenta):
                            doc = (num_doc or ref or "").strip()
                            desc = f"Cobro {nombre or ''}".replace("  ", " ").strip()
                            movs.append({
                                "id": idp, "fecha": fecha or "", "descripcion": desc,
                                "monto": float(monto or 0), "tipo": "ingreso",
                                "origen": "sistema", "tabla": "pagos_clientes", "documento": doc,
                            })
                except Exception:
                    conn.rollback()
                try:
                    c.execute("""
                        SELECT p.id, p.fecha_pago, p.proveedor_nombre, p.monto_pagado,
                               COALESCE(p.cuenta_origen, ''), p.id_factura,
                               COALESCE(f.numero_documento, ''), COALESCE(p.codigo_cotizacion, '')
                        FROM pagos_comprobantes p
                        LEFT JOIN facturas_recibidas f ON p.id_factura = f.id
                    """)
                    for idp, fecha, nombre, monto, cuenta, idf, num_doc, ref in c.fetchall():
                        if banco_matchea(banco, cuenta):
                            doc = (num_doc or ref or "").strip()
                            desc = f"Pago {nombre or ''}".replace("  ", " ").strip()
                            movs.append({
                                "id": idp, "fecha": fecha or "", "descripcion": desc,
                                "monto": -float(monto or 0), "tipo": "egreso",
                                "origen": "sistema", "tabla": "pagos_comprobantes", "documento": doc,
                            })
                except Exception:
                    conn.rollback()
                try:
                    c.execute("""
                        SELECT id, banco_origen, cuenta_origen, banco_destino, cuenta_destino,
                               fecha, descripcion, monto
                        FROM transferencias_bancarias
                    """)
                    for idt, bo, co, bd, cd, fecha, desc, monto in c.fetchall():
                        monto_v = float(monto or 0)
                        et_origen = construir_etiqueta_banco({"banco": bo, "cuenta": co})
                        et_destino = construir_etiqueta_banco({"banco": bd, "cuenta": cd})
                        if banco_matchea(banco, et_origen):
                            desc_eg = f"Transferencia a {bd or 'otra cuenta'}"
                            if desc:
                                desc_eg += f" ({desc})"
                            movs.append({
                                "id": idt, "fecha": fecha or "", "descripcion": desc_eg,
                                "monto": -monto_v, "tipo": "transferencia",
                                "origen": "sistema", "tabla": "transferencias_bancarias", "documento": "",
                            })
                        if banco_matchea(banco, et_destino):
                            desc_in = f"Transferencia de {bo or 'otra cuenta'}"
                            if desc:
                                desc_in += f" ({desc})"
                            movs.append({
                                "id": idt, "fecha": fecha or "", "descripcion": desc_in,
                                "monto": monto_v, "tipo": "transferencia",
                                "origen": "sistema", "tabla": "transferencias_bancarias", "documento": "",
                            })
                except Exception:
                    conn.rollback()
        except Exception:
            pass
        finally:
            liberar_conexion(conn)
        if mes:
            movs = [m for m in movs if normalizar_fecha(m["fecha"]).startswith(mes)]
        movs.sort(key=lambda m: normalizar_fecha(m["fecha"]))
        return movs

    def ver_movimientos(self, banco):
        resumen = self.calcular_resumen(banco)
        movs = resumen["movimientos"]
        v = ctk.CTkToplevel(self.parent_frame)
        v.title(f"Movimientos - {construir_etiqueta_banco(banco)}")
        v.geometry("760x480")
        v.transient(self.parent_frame)

        ctk.CTkLabel(v, text=f"Movimientos de {construir_etiqueta_banco(banco)}",
                     font=("Arial", 15, "bold"), text_color="#1f538d").pack(pady=(15, 5))

        f_tabla = ctk.CTkFrame(v, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True, padx=15, pady=10)
        columnas = ("fecha", "documento", "descripcion", "tipo", "monto")
        tabla = ttk.Treeview(f_tabla, columns=columnas, show="headings", height=14)
        tabla.heading("fecha", text="Fecha")
        tabla.heading("documento", text="N° Documento")
        tabla.heading("descripcion", text="Descripción")
        tabla.heading("tipo", text="Tipo")
        tabla.heading("monto", text="Monto")
        tabla.column("fecha", width=95, anchor="center")
        tabla.column("documento", width=150, anchor="center")
        tabla.column("descripcion", width=260, anchor="w")
        tabla.column("tipo", width=110, anchor="center")
        tabla.column("monto", width=130, anchor="e")
        vsb = ttk.Scrollbar(f_tabla, orient="vertical", command=tabla.yview)
        tabla.configure(yscrollcommand=vsb.set)
        tabla.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        movs = sorted(movs, key=lambda m: (normalizar_fecha(m["fecha"]), m["monto"]))
        for m in movs:
            if m.get("tipo") == "transferencia":
                tipo_txt = "Transferencia (Ingreso)" if m["monto"] > 0 else "Transferencia (Egreso)"
            else:
                tipo_txt = "Ingreso (Cobro)" if m["monto"] > 0 else "Egreso (Pago)"
            tabla.insert("", tk.END, values=(m["fecha"], m.get("documento", ""), m["descripcion"],
                                             tipo_txt, formatear_monto(m["monto"])))
        ctk.CTkButton(v, text="Cerrar", width=120, fg_color="#34495e",
                      command=v.destroy).pack(pady=(0, 15))

    # -----------------------------------------------------
    # TAB: TRANSFERENCIAS ENTRE CUENTAS
    # -----------------------------------------------------
    def construir_tab_transferencias(self):
        f_form = ctk.CTkFrame(self.tab_transferencias, fg_color="#f8f9fa",
                              border_width=1, border_color="#e0e0e0", corner_radius=8)
        f_form.pack(fill="x", padx=5, pady=(5, 10))

        etiquetas = [construir_etiqueta_banco(b) for b in self.bancos]
        if not etiquetas:
            etiquetas = ["(Sin bancos configurados)"]

        r1 = ctk.CTkFrame(f_form, fg_color="transparent"); r1.pack(fill="x", padx=12, pady=(12, 4))
        ctk.CTkLabel(r1, text="Banco Origen:", width=130, anchor="w", font=("Arial", 12, "bold")).pack(side="left")
        self.cmb_tx_origen = ctk.CTkComboBox(r1, values=etiquetas, width=340, state="readonly")
        self.cmb_tx_origen.pack(side="left", padx=6)
        if etiquetas:
            self.cmb_tx_origen.set(etiquetas[0])

        r2 = ctk.CTkFrame(f_form, fg_color="transparent"); r2.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(r2, text="Banco Destino:", width=130, anchor="w", font=("Arial", 12, "bold")).pack(side="left")
        self.cmb_tx_destino = ctk.CTkComboBox(r2, values=etiquetas, width=340, state="readonly")
        self.cmb_tx_destino.pack(side="left", padx=6)
        if len(etiquetas) > 1:
            self.cmb_tx_destino.set(etiquetas[1])
        elif etiquetas:
            self.cmb_tx_destino.set(etiquetas[0])

        r3 = ctk.CTkFrame(f_form, fg_color="transparent"); r3.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(r3, text="Fecha:", width=130, anchor="w", font=("Arial", 12, "bold")).pack(side="left")
        self.ent_tx_fecha = ctk.CTkEntry(r3, width=170)
        self.ent_tx_fecha.pack(side="left", padx=6)
        self.ent_tx_fecha.insert(0, datetime.now().strftime("%d/%m/%Y"))
        ctk.CTkLabel(r3, text="Monto:", width=90, anchor="w", font=("Arial", 12, "bold")).pack(side="left", padx=(16, 0))
        self.ent_tx_monto = ctk.CTkEntry(r3, width=170)
        self.ent_tx_monto.pack(side="left", padx=6)

        r4 = ctk.CTkFrame(f_form, fg_color="transparent"); r4.pack(fill="x", padx=12, pady=(4, 4))
        ctk.CTkLabel(r4, text="Descripción:", width=130, anchor="w", font=("Arial", 12, "bold")).pack(side="left")
        self.ent_tx_desc = ctk.CTkEntry(r4)
        self.ent_tx_desc.pack(side="left", fill="x", expand=True, padx=6)

        r5 = ctk.CTkFrame(f_form, fg_color="transparent"); r5.pack(fill="x", padx=12, pady=(4, 12))
        ctk.CTkButton(r5, text="💸 Realizar Transferencia", width=220, height=38,
                      font=("Arial", 13, "bold"), fg_color="#27ae60", hover_color="#1e8449",
                      command=self.realizar_transferencia).pack(side="left", padx=(0, 8))
        ctk.CTkButton(r5, text="✏️ Editar", width=120, height=38,
                      font=("Arial", 12, "bold"), fg_color="#2980b9", hover_color="#1f618d",
                      command=self.editar_transferencia).pack(side="left", padx=(0, 8))
        ctk.CTkButton(r5, text="🗑️ Eliminar", width=120, height=38,
                      font=("Arial", 12, "bold"), fg_color="#e74c3c", hover_color="#c0392b",
                      command=self.eliminar_transferencia).pack(side="left")

        ctk.CTkLabel(self.tab_transferencias, text="Historial de transferencias:",
                     font=("Arial", 13, "bold"), text_color="#1f538d").pack(anchor="w", padx=5, pady=(0, 4))
        f_tabla = ctk.CTkFrame(self.tab_transferencias, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True, padx=5, pady=(0, 5))
        columnas = ("fecha", "origen", "destino", "descripcion", "monto")
        self.tabla_tx = ttk.Treeview(f_tabla, columns=columnas, show="headings", selectmode="extended")
        for col, txt, w, anc in (("fecha", "Fecha", 100, "center"),
                                 ("origen", "Banco Origen", 200, "center"),
                                 ("destino", "Banco Destino", 200, "center"),
                                 ("descripcion", "Descripción", 260, "w"),
                                 ("monto", "Monto", 130, "e")):
            self.tabla_tx.heading(col, text=txt)
            self.tabla_tx.column(col, width=w, anchor=anc)
        vsb = ttk.Scrollbar(f_tabla, orient="vertical", command=self.tabla_tx.yview)
        self.tabla_tx.configure(yscrollcommand=vsb.set)
        self.tabla_tx.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.refrescar_transferencias()

    def realizar_transferencia(self):
        etiquetas = [construir_etiqueta_banco(b) for b in self.bancos]
        origen = self.cmb_tx_origen.get()
        destino = self.cmb_tx_destino.get()
        if origen not in etiquetas or destino not in etiquetas:
            messagebox.showwarning("Transferencia", "Seleccione bancos válidos.", parent=self.parent_frame)
            return
        if origen == destino:
            messagebox.showwarning("Transferencia", "El banco origen y destino deben ser diferentes.", parent=self.parent_frame)
            return
        monto = normalizar_monto(self.ent_tx_monto.get())
        if monto <= 0:
            messagebox.showerror("Error", "Ingrese un monto mayor a 0.", parent=self.parent_frame)
            return
        fecha = self.ent_tx_fecha.get().strip() or datetime.now().strftime("%d/%m/%Y")
        desc = self.ent_tx_desc.get().strip() or "Transferencia entre cuentas"

        bo = self.bancos[etiquetas.index(origen)]
        bd = self.bancos[etiquetas.index(destino)]

        conn = conectar_db(silencioso=True)
        if not conn:
            messagebox.showerror("Error", "Sin conexión a la base de datos.", parent=self.parent_frame)
            return
        try:
            with conn.cursor() as c:
                c.execute("""
                    INSERT INTO transferencias_bancarias
                    (banco_origen, cuenta_origen, banco_destino, cuenta_destino, fecha, descripcion, monto)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                """, (bo.get("banco", ""), bo.get("cuenta", ""), bd.get("banco", ""), bd.get("cuenta", ""),
                      fecha, desc, monto))
                conn.commit()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo registrar la transferencia:\n{e}", parent=self.parent_frame)
            return
        finally:
            liberar_conexion(conn)

        registrar_auditoria(self.usuario_activo, "Banco",
                            f"Transferencia {formatear_monto(monto)} de {origen} a {destino}")
        self.ent_tx_monto.delete(0, tk.END)
        self.ent_tx_desc.delete(0, tk.END)
        self.refrescar_transferencias()
        self.refrescar_saldos()
        messagebox.showinfo("Éxito", f"Transferencia de {formatear_monto(monto)} registrada:\n"
                                     f"• Egreso en {origen}\n• Ingreso en {destino}", parent=self.parent_frame)

    def cargar_transferencias(self):
        conn = conectar_db(silencioso=True)
        if not conn:
            return []
        rows = []
        try:
            with conn.cursor() as c:
                c.execute("""
                    SELECT id, banco_origen, cuenta_origen, banco_destino, cuenta_destino,
                           fecha, descripcion, monto
                    FROM transferencias_bancarias
                    ORDER BY id DESC
                """)
                for idt, bo, co, bd, cd, fecha, desc, monto in c.fetchall():
                    rows.append({
                        "id": idt,
                        "origen": construir_etiqueta_banco({"banco": bo, "cuenta": co}),
                        "destino": construir_etiqueta_banco({"banco": bd, "cuenta": cd}),
                        "fecha": fecha or "", "descripcion": desc or "", "monto": float(monto or 0),
                    })
        except Exception:
            pass
        finally:
            liberar_conexion(conn)
        return rows

    def refrescar_transferencias(self):
        self.tabla_tx.delete(*self.tabla_tx.get_children())
        for t in self.cargar_transferencias():
            self.tabla_tx.insert("", tk.END, iid=str(t["id"]),
                                 values=(t["fecha"], t["origen"], t["destino"],
                                         t["descripcion"], formatear_monto(t["monto"])))

    def editar_transferencia(self):
        sel = self.tabla_tx.selection()
        if len(sel) != 1:
            messagebox.showinfo("Editar", "Seleccione exactamente una transferencia para editar.", parent=self.parent_frame)
            return
        id_tx = int(sel[0])
        tx = None
        for t in self.cargar_transferencias():
            if t["id"] == id_tx:
                tx = t
                break
        if not tx:
            messagebox.showerror("Error", "No se encontró la transferencia.", parent=self.parent_frame)
            return

        etiquetas = [construir_etiqueta_banco(b) for b in self.bancos]
        if not etiquetas:
            etiquetas = ["(Sin bancos configurados)"]

        v = ctk.CTkToplevel(self.parent_frame)
        v.title("Editar Transferencia")
        v.geometry("460x430")
        v.transient(self.parent_frame)
        v.grab_set()

        ctk.CTkLabel(v, text="✏️ Editar transferencia", font=("Arial", 15, "bold"),
                     text_color="#1f538d").pack(pady=(15, 5))
        f = ctk.CTkFrame(v, fg_color="transparent")
        f.pack(fill="x", padx=20)

        ctk.CTkLabel(f, text="Banco Origen:", font=("Arial", 11, "bold")).pack(anchor="w")
        cmb_origen = ctk.CTkComboBox(f, values=etiquetas, width=300, state="readonly")
        cmb_origen.pack(fill="x", pady=(0, 8))
        cmb_origen.set(tx["origen"] if tx["origen"] in etiquetas else (etiquetas[0] if etiquetas else ""))

        ctk.CTkLabel(f, text="Banco Destino:", font=("Arial", 11, "bold")).pack(anchor="w")
        cmb_destino = ctk.CTkComboBox(f, values=etiquetas, width=300, state="readonly")
        cmb_destino.pack(fill="x", pady=(0, 8))
        cmb_destino.set(tx["destino"] if tx["destino"] in etiquetas else (etiquetas[1] if len(etiquetas) > 1 else (etiquetas[0] if etiquetas else "")))

        ctk.CTkLabel(f, text="Fecha:", font=("Arial", 11, "bold")).pack(anchor="w")
        f_fecha = ctk.CTkFrame(f, fg_color="transparent")
        f_fecha.pack(fill="x", pady=(0, 8))
        ent_fecha = ctk.CTkEntry(f_fecha)
        ent_fecha.pack(side="left", fill="x", expand=True)
        ent_fecha.insert(0, tx["fecha"])
        ctk.CTkButton(f_fecha, text="📅", width=42, font=("Arial", 13, "bold"),
                      fg_color="#1f538d", hover_color="#163b65",
                      command=lambda: CalendarioNativo(v, ent_fecha)).pack(side="left", padx=(6, 0))

        ctk.CTkLabel(f, text="Descripción:", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_desc = ctk.CTkEntry(f)
        ent_desc.pack(fill="x", pady=(0, 8)); ent_desc.insert(0, tx["descripcion"])

        ctk.CTkLabel(f, text="Monto:", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_monto = ctk.CTkEntry(f)
        ent_monto.pack(fill="x", pady=(0, 8)); ent_monto.insert(0, f"{tx['monto']:.2f}")

        def guardar():
            etiquetas_all = [construir_etiqueta_banco(b) for b in self.bancos]
            origen = cmb_origen.get()
            destino = cmb_destino.get()
            if origen not in etiquetas_all or destino not in etiquetas_all:
                messagebox.showwarning("Transferencia", "Seleccione bancos válidos.", parent=v)
                return
            if origen == destino:
                messagebox.showwarning("Transferencia", "El banco origen y destino deben ser diferentes.", parent=v)
                return
            monto = normalizar_monto(ent_monto.get())
            if monto <= 0:
                messagebox.showerror("Error", "Ingrese un monto mayor a 0.", parent=v)
                return
            fecha = ent_fecha.get().strip() or datetime.now().strftime("%d/%m/%Y")
            desc = ent_desc.get().strip() or "Transferencia entre cuentas"
            bo = self.bancos[etiquetas_all.index(origen)]
            bd = self.bancos[etiquetas_all.index(destino)]

            conn = conectar_db(silencioso=True)
            if not conn:
                messagebox.showerror("Error", "Sin conexión a la base de datos.", parent=v)
                return
            try:
                with conn.cursor() as c:
                    c.execute("""
                        UPDATE transferencias_bancarias
                        SET banco_origen=%s, cuenta_origen=%s, banco_destino=%s, cuenta_destino=%s,
                            fecha=%s, descripcion=%s, monto=%s
                        WHERE id=%s
                    """, (bo.get("banco", ""), bo.get("cuenta", ""), bd.get("banco", ""), bd.get("cuenta", ""),
                          fecha, desc, monto, id_tx))
                    conn.commit()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo actualizar:\n{e}", parent=v)
                return
            finally:
                liberar_conexion(conn)

            registrar_auditoria(self.usuario_activo, "Banco", f"Editó transferencia #{id_tx}")
            v.destroy()
            self.refrescar_transferencias()
            self.refrescar_saldos()
            messagebox.showinfo("Éxito", "Transferencia actualizada correctamente.", parent=self.parent_frame)

        ctk.CTkButton(v, text="✅ Guardar Cambios", width=160, fg_color="#27ae60", command=guardar).pack(pady=10)

    def eliminar_transferencia(self):
        sel = self.tabla_tx.selection()
        if not sel:
            messagebox.showinfo("Eliminar", "Seleccione una o más transferencias para eliminar.", parent=self.parent_frame)
            return
        if not messagebox.askyesno("Confirmar", f"¿Eliminar {len(sel)} transferencia(s) seleccionada(s)?", parent=self.parent_frame):
            return
        ids = [int(i) for i in sel]
        conn = conectar_db(silencioso=True)
        if not conn:
            messagebox.showerror("Error", "Sin conexión a la base de datos.", parent=self.parent_frame)
            return
        try:
            with conn.cursor() as c:
                c.execute("DELETE FROM transferencias_bancarias WHERE id = ANY(%s)", (ids,))
                conn.commit()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo eliminar:\n{e}", parent=self.parent_frame)
            return
        finally:
            liberar_conexion(conn)
        registrar_auditoria(self.usuario_activo, "Banco", f"Eliminó {len(ids)} transferencia(s)")
        self.refrescar_transferencias()
        self.refrescar_saldos()
        messagebox.showinfo("Éxito", "Transferencia(s) eliminada(s) correctamente.", parent=self.parent_frame)

    # -----------------------------------------------------
    # TAB: CONCILIACION BANCARIA
    # -----------------------------------------------------
    def construir_tab_conciliacion(self):
        f_sel = ctk.CTkFrame(self.tab_conciliacion, fg_color="#f8f9fa",
                             border_width=1, border_color="#e0e0e0", corner_radius=8)
        f_sel.pack(fill="x", padx=5, pady=(5, 10))

        ctk.CTkLabel(f_sel, text="Banco a conciliar:", font=("Arial", 12, "bold")).pack(side="left", padx=(12, 6), pady=10)
        etiquetas = [construir_etiqueta_banco(b) for b in self.bancos]
        if not etiquetas:
            etiquetas = ["(Sin bancos configurados)"]
        self.cmb_banco = ctk.CTkComboBox(f_sel, values=etiquetas, width=320, state="readonly")
        self.cmb_banco.pack(side="left", padx=6, pady=10)
        if etiquetas:
            self.cmb_banco.set(etiquetas[0])

        ctk.CTkLabel(f_sel, text="Mes a conciliar:", font=("Arial", 12, "bold")).pack(side="left", padx=(20, 6), pady=10)
        valores_mes = construir_valores_mes()
        self.cmb_mes = ctk.CTkComboBox(f_sel, values=valores_mes, width=160, state="readonly")
        self.cmb_mes.pack(side="left", padx=6, pady=10)
        self.cmb_mes.set(mes_actual_etiqueta())

        f_btns = ctk.CTkFrame(self.tab_conciliacion, fg_color="transparent")
        f_btns.pack(fill="x", padx=5, pady=(0, 8))
        ctk.CTkButton(f_btns, text="📄 Subir PDF Estado de Cuenta", font=("Arial", 12, "bold"),
                      fg_color="#e67e22", hover_color="#ca6f1e",
                      command=self.subir_pdf).pack(side="left", padx=4)
        ctk.CTkButton(f_btns, text="🔄 Cargar Movimientos del Sistema", font=("Arial", 12, "bold"),
                      fg_color="#1f538d", hover_color="#163b65",
                      command=self.generar_reporte).pack(side="left", padx=4)
        ctk.CTkButton(f_btns, text="✅ Conciliar Seleccionados", font=("Arial", 12, "bold"),
                      fg_color="#27ae60", hover_color="#1e8449",
                      command=self.conciliar_seleccionados).pack(side="left", padx=4)
        ctk.CTkButton(f_btns, text="➕ Registrar Pago / Compra Cruzada", font=("Arial", 12, "bold"),
                      fg_color="#8e44ad", hover_color="#703688",
                      command=self.agregar_movimiento_manual).pack(side="left", padx=4)
        ctk.CTkButton(f_btns, text="✏️ Editar / Corregir", font=("Arial", 12, "bold"),
                      fg_color="#34495e", hover_color="#2c3e50",
                      command=self.corregir_movimiento).pack(side="left", padx=4)
        ctk.CTkButton(f_btns, text="🗑️ Eliminar", font=("Arial", 12, "bold"),
                      fg_color="#c0392b", hover_color="#96281b",
                      command=self.eliminar_movimiento_conciliacion).pack(side="left", padx=4)
        ctk.CTkButton(f_btns, text="💾 Exportar", font=("Arial", 12, "bold"),
                      fg_color="#7f8c8d", hover_color="#606b6b",
                      command=self.exportar_reporte).pack(side="left", padx=4)

        self.f_resumen = ctk.CTkFrame(self.tab_conciliacion, fg_color="#eaf2f8",
                                      border_width=1, border_color="#d4e6f1", corner_radius=8)
        self.f_resumen.pack(fill="x", padx=5, pady=(0, 8))
        self.lbl_resumen = ctk.CTkLabel(self.f_resumen, text="Seleccione un banco, suba el PDF del estado de cuenta "
                                                             "y cargue los movimientos del sistema.",
                                        font=("Arial", 12), text_color="#1f538d", justify="left")
        self.lbl_resumen.pack(anchor="w", padx=12, pady=10)

        f_tabla = ctk.CTkFrame(self.tab_conciliacion, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True, padx=5, pady=(0, 5))
        columnas = ("estado", "fecha", "documento", "descripcion", "tipo", "monto", "origen")
        self.tabla_conc = ttk.Treeview(f_tabla, columns=columnas, show="headings",
                                       selectmode="extended")
        for col, txt, w, anc in (("estado", "Estado", 140, "center"),
                                 ("fecha", "Fecha", 90, "center"),
                                 ("documento", "N° Documento", 140, "center"),
                                 ("descripcion", "Descripción", 250, "w"),
                                 ("tipo", "Tipo", 100, "center"),
                                 ("monto", "Monto", 110, "e"),
                                 ("origen", "Origen", 120, "center")):
            self.tabla_conc.heading(col, text=txt)
            self.tabla_conc.column(col, width=w, anchor=anc)
        vsb = ttk.Scrollbar(f_tabla, orient="vertical", command=self.tabla_conc.yview)
        self.tabla_conc.configure(yscrollcommand=vsb.set)
        self.tabla_conc.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tabla_conc.bind("<Delete>", lambda _e: self.eliminar_movimiento_conciliacion())
        self.tabla_conc.tag_configure("conciliado", background="#d5f5e3")
        self.tabla_conc.tag_configure("diferencia", background="#fadbd8")
        self.tabla_conc.tag_configure("estado_cuenta", background="#d6eaf8")

        f_pie = ctk.CTkFrame(self.tab_conciliacion, fg_color="transparent")
        f_pie.pack(fill="x", padx=5)
        ctk.CTkButton(f_pie, text="👁 Ver Texto Extraído del PDF", width=220, font=("Arial", 11),
                      fg_color="#7f8c8d", hover_color="#606b6b",
                      command=self.mostrar_texto_pdf).pack(side="left")

    def banco_seleccionado(self):
        etiquetas = [construir_etiqueta_banco(b) for b in self.bancos]
        sel = self.cmb_banco.get()
        if sel in etiquetas:
            return self.bancos[etiquetas.index(sel)]
        if self.bancos:
            return self.bancos[0]
        return None

    def ir_a_conciliacion(self, banco):
        etiquetas = [construir_etiqueta_banco(b) for b in self.bancos]
        et = construir_etiqueta_banco(banco)
        if et in etiquetas:
            self.cmb_banco.set(et)
        self.tabview.set(" 🧾 Conciliación Bancaria ")

    def subir_pdf(self):
        ruta = seleccionar_archivo_dialogo(
            titulo="Seleccionar Estado de Cuenta (PDF)",
            tipos=[("Archivos PDF", "*.pdf")])
        if not ruta:
            return
        self.lbl_resumen.configure(text="⏳ Extrayendo texto del PDF...")
        self.parent_frame.update_idletasks()

        def tarea():
            texto = extraer_texto_pdf(ruta)
            if not texto.strip():
                self.lbl_resumen.configure(text="❌ No se pudo extraer texto del PDF. "
                                                "Verifique que no sea un PDF escaneado (imagen).")
                return
            movs = parsear_movimientos(texto)
            self.texto_pdf = texto
            self.movimientos_pdf = movs
            self.saldo_final_estado = detectar_saldo_final(texto)
            self.lbl_resumen.configure(text=f"✅ PDF procesado: {len(movs)} movimientos detectados.")
            self.generar_reporte()

        threading.Thread(target=tarea, daemon=True).start()

    def generar_reporte(self):
        banco = self.banco_seleccionado()
        if not banco:
            self.lbl_resumen.configure(text="⚠️ No hay banco seleccionado.")
            return

        mes_key = mes_etiqueta_a_key(self.cmb_mes.get())
        sys_movs = self.cargar_movimientos_sistema(banco, mes=mes_key)
        manuales = self.cargar_manuales(banco)
        bank_movs = [b for b in self.movimientos_pdf
                     if not mes_key or normalizar_fecha(b["fecha"]).startswith(mes_key)]

        filas = []
        usados = set()

        for s in sys_movs:
            s_fecha = normalizar_fecha(s["fecha"])
            idx_match = None
            for i, b in enumerate(bank_movs):
                if i in usados:
                    continue
                if s_fecha and normalizar_fecha(b["fecha"]) == s_fecha and abs(b["monto"] - s["monto"]) <= 0.02:
                    idx_match = i
                    break
            if idx_match is None:
                cands = [i for i, b in enumerate(bank_movs)
                         if i not in usados and abs(b["monto"] - s["monto"]) <= 0.02]
                if len(cands) == 1:
                    idx_match = cands[0]
            if idx_match is not None:
                usados.add(idx_match)
                b = bank_movs[idx_match]
                filas.append({**s, "estado": "conciliado"})
                filas.append({**b, "id": 0, "tipo": "ingreso" if b["monto"] > 0 else "egreso",
                              "origen": "estado_cuenta", "tabla": "", "estado": "conciliado"})
            else:
                filas.append({**s, "estado": "diferencia"})

        for i, b in enumerate(bank_movs):
            if i not in usados:
                filas.append({**b, "id": 0, "tipo": "ingreso" if b["monto"] > 0 else "egreso",
                              "origen": "estado_cuenta", "tabla": "", "estado": "diferencia"})

        for m in manuales:
            filas.append({**m, "estado": m.get("estado", "pendiente"), "origen": "manual"})

        filas.sort(key=lambda f: (normalizar_fecha(f["fecha"]), f["monto"]))
        self.filas_conciliacion = filas
        self.refrescar_tree()

        saldo_inicial = normalizar_monto(banco.get("saldo_inicial", ""))
        ingresos = sum(m["monto"] for m in sys_movs if m["monto"] > 0)
        egresos = sum(-m["monto"] for m in sys_movs if m["monto"] < 0)
        neto_sistema = ingresos - egresos
        neto_pdf = sum(b["monto"] for b in bank_movs)
        diferencia_mes = neto_sistema - neto_pdf

        if mes_key:
            etiqueta_mes = self.cmb_mes.get()
            txt = (f"Banco: {construir_etiqueta_banco(banco)}   |   Mes: {etiqueta_mes}   |   "
                   f"Sistema (neto): {formatear_monto(neto_sistema)}   |   "
                   f"Estado de cuenta (neto): {formatear_monto(neto_pdf)}   |   "
                   f"Diferencia: {formatear_monto(diferencia_mes)}")
        else:
            saldo_sistema = saldo_inicial + neto_sistema
            if self.saldo_final_estado is not None:
                diferencia = saldo_sistema - self.saldo_final_estado
                txt = (f"Banco: {construir_etiqueta_banco(banco)}   |   "
                       f"Saldo Sistema: {formatear_monto(saldo_sistema)}   |   "
                       f"Saldo Estado de Cuenta: {formatear_monto(self.saldo_final_estado)}   |   "
                       f"Diferencia: {formatear_monto(diferencia)}")
            else:
                txt = (f"Banco: {construir_etiqueta_banco(banco)}   |   "
                       f"Saldo Sistema: {formatear_monto(saldo_sistema)}   |   "
                       f"Movimientos PDF (neto): {formatear_monto(neto_pdf)}   |   "
                       f"(No se detectó saldo final en el PDF)")
        self.lbl_resumen.configure(text=txt)

    def cargar_manuales(self, banco):
        conn = conectar_db(silencioso=True)
        if not conn:
            return []
        movs = []
        try:
            with conn.cursor() as c:
                try:
                    c.execute("""
                        SELECT id, fecha, descripcion, monto, tipo, estado,
                               COALESCE(categoria, ''), COALESCE(subcategoria, ''),
                               COALESCE(placa, ''), COALESCE(chofer, ''),
                               COALESCE(comision, 0), COALESCE(interbancario, FALSE),
                               COALESCE(es_cruzada, FALSE)
                        FROM conciliacion_bancaria
                        WHERE origen = 'manual' AND banco = %s
                        ORDER BY id
                    """, (banco.get("banco", ""),))
                    filas = c.fetchall()
                except Exception:
                    # Base de datos anterior sin las columnas de detalle
                    conn.rollback()
                    c.execute("""
                        SELECT id, fecha, descripcion, monto, tipo, estado
                        FROM conciliacion_bancaria
                        WHERE origen = 'manual' AND banco = %s
                        ORDER BY id
                    """, (banco.get("banco", ""),))
                    filas = [tuple(f) + ("", "", "", "", 0, False, False) for f in c.fetchall()]

                for (idp, fecha, desc, monto, tipo, estado, categoria, subcategoria,
                     placa, chofer, comision, interbancario, es_cruzada) in filas:
                    movs.append({
                        "id": idp, "fecha": fecha or "", "descripcion": desc or "",
                        "monto": float(monto or 0), "tipo": tipo or "ingreso",
                        "tabla": "conciliacion", "estado": estado or "pendiente",
                        "categoria": categoria or "", "subcategoria": subcategoria or "",
                        "placa": placa or "", "chofer": chofer or "",
                        "comision": float(comision or 0), "interbancario": bool(interbancario),
                        "es_cruzada": bool(es_cruzada),
                    })
        except Exception:
            pass
        finally:
            liberar_conexion(conn)
        return movs

    def refrescar_tree(self):
        self.tabla_conc.delete(*self.tabla_conc.get_children())
        for i, f in enumerate(self.filas_conciliacion):
            estado = f.get("estado", "pendiente")
            if estado == "conciliado":
                estado_txt = "✅ Conciliado"
                tag = "conciliado"
            elif f.get("origen") == "estado_cuenta":
                estado_txt = "❌ Sin conciliar (Banco)"
                tag = "diferencia"
            elif estado == "diferencia":
                estado_txt = "⚠️ Sin conciliar (Sistema)"
                tag = "diferencia"
            else:
                estado_txt = "📝 Manual / Pendiente"
                tag = "estado_cuenta"
            if f.get("tipo") == "transferencia":
                tipo_txt = "Transferencia (Ingreso)" if f["monto"] >= 0 else "Transferencia (Egreso)"
            else:
                tipo_txt = "Ingreso" if f["monto"] >= 0 else "Egreso"
            origen_txt = {"sistema": "Sistema", "estado_cuenta": "Estado de Cuenta",
                          "manual": "Manual"}.get(f.get("origen"), f.get("origen", ""))
            self.tabla_conc.insert("", tk.END, iid=str(i),
                                   values=(estado_txt, f["fecha"], f.get("documento", ""),
                                           f["descripcion"], tipo_txt, formatear_monto(f["monto"]),
                                           origen_txt),
                                   tags=(tag,))

    # -----------------------------------------------------
    # ACCIONES DE CONCILIACION
    # -----------------------------------------------------
    def _filas_seleccionadas(self):
        sel = self.tabla_conc.selection()
        return [self.filas_conciliacion[int(i)] for i in sel if i.isdigit()]

    def conciliar_seleccionados(self):
        filas = self._filas_seleccionadas()
        if not filas:
            messagebox.showinfo("Conciliación", "Seleccione uno o más movimientos para conciliar.",
                                parent=self.parent_frame)
            return
        banco = self.banco_seleccionado()
        for f in filas:
            f["estado"] = "conciliado"
            self._persistir_conciliado(f, banco)
        self.refrescar_tree()
        registrar_auditoria(self.usuario_activo, "Banco",
                            f"Concilió {len(filas)} movimiento(s) en {construir_etiqueta_banco(banco) if banco else 'banco'}")
        messagebox.showinfo("Conciliación", f"{len(filas)} movimiento(s) marcado(s) como conciliado(s).",
                            parent=self.parent_frame)

    def _persistir_conciliado(self, f, banco):
        if not banco:
            return
        conn = conectar_db(silencioso=True)
        if not conn:
            return
        try:
            with conn.cursor() as c:
                c.execute("""
                    SELECT COUNT(*) FROM conciliacion_bancaria
                    WHERE origen = %s AND id_movimiento = %s AND banco = %s
                """, (f.get("origen", ""), f.get("id", 0), banco.get("banco", "")))
                if c.fetchone()[0] == 0:
                    c.execute("""
                        INSERT INTO conciliacion_bancaria
                        (banco, cuenta, fecha, descripcion, monto, tipo, origen, id_movimiento, estado, fecha_conciliacion)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, 'conciliado', %s)
                    """, (banco.get("banco", ""), banco.get("cuenta", ""), f.get("fecha", ""),
                          f.get("descripcion", ""), f.get("monto", 0), f.get("tipo", "ingreso"),
                          f.get("origen", ""), f.get("id", 0), datetime.now().strftime("%Y-%m-%d")))
                    conn.commit()
        except Exception:
            pass
        finally:
            liberar_conexion(conn)

    def _datos_iniciales_pago(self, registro, cats, principales):
        """Datos de un pago ya registrado, para abrir la ventana en modo edición.

        Se toman de las columnas de detalle del pago y, cuando el registro es
        antiguo y no las tiene, se deducen de su descripción.
        """
        datos = {
            "fecha": str(registro.get("fecha") or ""),
            "monto": abs(float(registro.get("monto") or 0)),
            "principal": str(registro.get("categoria") or ""),
            "sub": str(registro.get("subcategoria") or ""),
            "placa": str(registro.get("placa") or ""),
            "chofer": str(registro.get("chofer") or ""),
            "comision": float(registro.get("comision") or 0),
            "interbancario": bool(registro.get("interbancario")),
            "es_cruzada": bool(registro.get("es_cruzada")),
            "descripcion": "",
        }
        libres = []
        for parte in [p.strip() for p in str(registro.get("descripcion") or "").split(" - ") if p.strip()]:
            bajo = parte.lower()
            if bajo.startswith("placa "):
                datos["placa"] = datos["placa"] or parte[6:].strip()
            elif bajo.startswith("chofer "):
                datos["chofer"] = datos["chofer"] or parte[7:].strip()
            elif bajo.startswith("comisión interbancaria") or bajo.startswith("comision interbancaria"):
                datos["interbancario"] = True
                if not datos["comision"]:
                    datos["comision"] = monto_desde_texto(parte)
            elif parte == datos["principal"] or (datos["sub"] and parte == datos["sub"]):
                pass                      # ya se recuperó de las columnas del pago
            elif not datos["principal"] and parte in principales:
                datos["principal"] = parte
            elif datos["principal"] and not datos["sub"] and parte in (cats.get(datos["principal"]) or []):
                datos["sub"] = parte
            else:
                libres.append(parte)
        datos["descripcion"] = " - ".join(libres)
        return datos

    def agregar_movimiento_manual(self, registro=None):
        """Ventana de Registrar Pago / Compra Cruzada.

        Si se recibe 'registro' (una fila de la conciliación de origen manual)
        se abre ESTA MISMA ventana en modo edición, con todos los datos del
        pago cargados, y al guardar se actualiza ese registro en lugar de
        crear uno nuevo.
        """
        banco = self.banco_seleccionado()
        if not banco:
            messagebox.showwarning("Banco", "Seleccione un banco.", parent=self.parent_frame)
            return

        cats = cargar_categorias_gastos()
        principales = list(cats.keys()) or ["Gastos Operativos"]
        comision = cargar_comision_interbancaria()
        en_edicion = bool(registro)
        datos_ini = (self._datos_iniciales_pago(registro, cats, principales) if en_edicion else {})
        id_edicion = registro.get("id") if en_edicion else None

        v = ctk.CTkToplevel(self.parent_frame)
        v.title("Editar Pago (Conciliación)" if en_edicion else "Registrar Pago (Conciliación)")
        v.geometry("580x700")
        v.minsize(520, 480)
        v.transient(self.parent_frame)
        v.grab_set()

        ctk.CTkLabel(v, text=("✏️ Editar Pago / Compra Cruzada" if en_edicion
                              else "➕ Registrar Pago / Compra Cruzada"),
                     font=("Arial", 15, "bold"), text_color="#1f538d").pack(pady=(15, 5))

        # Pie fijo: el botón Guardar siempre visible
        f_pie = ctk.CTkFrame(v, fg_color="transparent")
        f_pie.pack(side="bottom", fill="x")

        # Área desplazable (scroll) con todos los campos
        f = ctk.CTkScrollableFrame(v, fg_color="transparent")
        f.pack(fill="both", expand=True, padx=6, pady=(0, 4))

        principal_ini = str(datos_ini.get("principal") or "")
        if principal_ini and principal_ini not in principales:
            principales = principales + [principal_ini]

        ctk.CTkLabel(f, text="Categoría Principal:", font=("Arial", 11, "bold")).pack(anchor="w")
        cmb_principal = ctk.CTkComboBox(f, values=principales, width=300, state="readonly")
        cmb_principal.pack(fill="x", pady=(0, 8))
        cmb_principal.set(principal_ini or principales[0])

        sub_ini = str(datos_ini.get("sub") or "")
        ctk.CTkLabel(f, text="Categoría:", font=("Arial", 11, "bold")).pack(anchor="w")
        subs_inicial = list(cats.get(cmb_principal.get(), []) or ["(Sin categoría)"])
        if sub_ini and sub_ini not in subs_inicial:
            subs_inicial.append(sub_ini)
        cmb_sub = ctk.CTkComboBox(f, values=subs_inicial, width=300, state="readonly")
        cmb_sub.pack(fill="x", pady=(0, 8))
        cmb_sub.set(sub_ini or subs_inicial[0])

        # Tercer desplegable: placa (solo para Gastos Operativos)
        placas = cargar_placas_flota()
        lbl_placa = ctk.CTkLabel(f, text="Placa / Vehículo:", font=("Arial", 11, "bold"))
        cmb_placa = ctk.CTkComboBox(f, values=placas or ["(Sin placas)"], width=300, state="readonly")
        if placas:
            cmb_placa.set(placas[0])

        # Cuarto desplegable: chofer (solo para categorías de Planilla / Sueldos)
        choferes = cargar_choferes()
        lbl_chofer = ctk.CTkLabel(f, text="Chofer (planilla):", font=("Arial", 11, "bold"))
        cmb_chofer = ctk.CTkComboBox(f, values=["Ninguno"] + choferes, width=300, state="readonly")
        cmb_chofer.set("Ninguno")

        def actualizar_chofer():
            if es_categoria_planilla(f"{cmb_principal.get()} {cmb_sub.get()}"):
                lbl_chofer.pack(anchor="w", pady=(0, 2), before=btn_gestion)
                cmb_chofer.pack(fill="x", pady=(0, 8), before=btn_gestion)
            else:
                lbl_chofer.pack_forget()
                cmb_chofer.pack_forget()

        def on_principal(_=None):
            subs = cats.get(cmb_principal.get(), []) or ["(Sin categoría)"]
            cmb_sub.configure(values=subs)
            cmb_sub.set(subs[0])
            if cmb_principal.get() == "Gastos Operativos":
                lbl_placa.pack(anchor="w", pady=(0, 2), before=btn_gestion)
                cmb_placa.pack(fill="x", pady=(0, 8), before=btn_gestion)
            else:
                lbl_placa.pack_forget()
                cmb_placa.pack_forget()
            actualizar_chofer()
        cmb_principal.configure(command=on_principal)
        cmb_sub.configure(command=lambda _=None: actualizar_chofer())

        def abrir_gestion():
            self.gestionar_categorias()
            nonlocal cats, principales
            cats = cargar_categorias_gastos()
            principales = list(cats.keys()) or ["Gastos Operativos"]
            cmb_principal.configure(values=principales)
            cmb_principal.set(principales[0])
            on_principal()
        btn_gestion = ctk.CTkButton(f, text="⚙️ Gestionar Categorías", height=26, font=("Arial", 11),
                                    fg_color="#8e44ad", hover_color="#703688", command=abrir_gestion)
        btn_gestion.pack(fill="x", pady=(0, 8))
        on_principal()

        # En edición se restauran los valores guardados (on_principal los reinicia)
        if sub_ini:
            valores_sub = list(cmb_sub.cget("values"))
            if sub_ini not in valores_sub:
                cmb_sub.configure(values=valores_sub + [sub_ini])
            cmb_sub.set(sub_ini)
        placa_ini = str(datos_ini.get("placa") or "")
        if placa_ini and placa_ini != "(Sin placas)":
            valores_placa = list(cmb_placa.cget("values"))
            if placa_ini not in valores_placa:
                valores_placa.append(placa_ini)          # placa que ya no está en la flota
                cmb_placa.configure(values=valores_placa)
            cmb_placa.set(placa_ini)
        chofer_ini = str(datos_ini.get("chofer") or "")
        if chofer_ini:
            valores_chofer = list(cmb_chofer.cget("values"))
            if chofer_ini not in valores_chofer:
                valores_chofer.append(chofer_ini)
                cmb_chofer.configure(values=valores_chofer)
            cmb_chofer.set(chofer_ini)
        actualizar_chofer()

        ctk.CTkLabel(f, text="Fecha (DD/MM/AAAA):", font=("Arial", 11, "bold")).pack(anchor="w")
        f_fecha = ctk.CTkFrame(f, fg_color="transparent")
        f_fecha.pack(fill="x", pady=(0, 8))
        ent_fecha = ctk.CTkEntry(f_fecha)
        ent_fecha.pack(side="left", fill="x", expand=True)
        ent_fecha.insert(0, str(datos_ini.get("fecha") or datetime.now().strftime("%d/%m/%Y")))
        ctk.CTkButton(f_fecha, text="📅", width=42, font=("Arial", 13, "bold"),
                      fg_color="#1f538d", hover_color="#163b65",
                      command=lambda: CalendarioNativo(v, ent_fecha)).pack(side="left", padx=(6, 0))

        ctk.CTkLabel(f, text="Descripción:", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_desc = ctk.CTkEntry(f)
        ent_desc.pack(fill="x", pady=(0, 8))
        ent_desc.insert(0, str(datos_ini.get("descripcion") or ""))

        ctk.CTkLabel(f, text="Monto:", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_monto = ctk.CTkEntry(f)
        ent_monto.pack(fill="x", pady=(0, 8))
        if datos_ini.get("monto"):
            ent_monto.insert(0, formatear_numero_entrada(f"{float(datos_ini['monto']):.2f}"))
        ent_monto.bind("<KeyRelease>", lambda e: self._formatear_entrada_monto(ent_monto))
        ent_monto.bind("<FocusOut>", lambda e: self._formatear_entrada_monto(ent_monto))

        var_inter = tk.BooleanVar(value=bool(datos_ini.get("interbancario")))
        chk_inter = ctk.CTkCheckBox(f, text="Pago Interbancario", variable=var_inter)
        chk_inter.pack(anchor="w", pady=(0, 4))
        f_com = ctk.CTkFrame(f, fg_color="transparent")
        f_com.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(f_com, text="Comisión (S/.):", font=("Arial", 11, "bold")).pack(side="left")
        ent_comision = ctk.CTkEntry(f_com, width=90)
        ent_comision.pack(side="left", padx=6)
        comision_ini = datos_ini.get("comision")
        ent_comision.insert(0, f"{(float(comision_ini) if comision_ini else comision):.2f}")
        lbl_total = ctk.CTkLabel(f_com, text=f"Total: {formatear_monto(0)}", font=("Arial", 12, "bold"),
                                 text_color="#1f538d")
        lbl_total.pack(side="right")

        def actualizar_total(*_):
            base = normalizar_monto(ent_monto.get())
            com = normalizar_monto(ent_comision.get())
            total = base + (com if var_inter.get() else 0)
            lbl_total.configure(text=f"Total: {formatear_monto(total)}")
        chk_inter.configure(command=actualizar_total)
        ent_monto.bind("<KeyRelease>", actualizar_total)
        ent_comision.bind("<KeyRelease>", actualizar_total)

        # ------------------------------------------------------------------
        # UNIFICACIÓN CON COMPRAS CRUZADAS (pago a tercero)
        # Al marcar el check se habilitan los datos del tercero y el PDF del soporte
        # ------------------------------------------------------------------
        var_cruzada = tk.BooleanVar(value=False)
        chk_cruzada = ctk.CTkCheckBox(f, text="🔁 Es compra cruzada (factura pagada por un tercero)",
                                      font=("Arial", 11, "bold"), variable=var_cruzada)
        if not en_edicion:
            chk_cruzada.pack(anchor="w", pady=(4, 4))
        elif datos_ini.get("es_cruzada"):
            # La compra cruzada ya quedó registrada en Compras: al editar no se repite
            ctk.CTkLabel(f, text=("🔁 Este pago se registró como COMPRA CRUZADA en el módulo de Compras.\n"
                                  "Aquí solo se modifican los datos del movimiento bancario."),
                         font=("Arial", 11, "italic"), text_color="#7d3c98",
                         justify="left").pack(anchor="w", pady=(6, 4))

        estado_cruzada = {"soporte": "", "factura": ""}
        f_cruzada = ctk.CTkFrame(f, fg_color="#f4ecf7", corner_radius=8,
                                 border_width=1, border_color="#d7bde2")

        ctk.CTkLabel(f_cruzada, text="Proveedor de la factura:", font=("Arial", 11, "bold")).pack(
            anchor="w", padx=10, pady=(8, 0))
        cmb_cruz_prov = ctk.CTkComboBox(f_cruzada, values=cargar_proveedores() or ["(Sin proveedores)"],
                                        width=300)
        cmb_cruz_prov.pack(fill="x", padx=10, pady=(0, 6))
        cmb_cruz_prov.set("")

        ctk.CTkLabel(f_cruzada, text="N° Documento / Factura (opcional):", font=("Arial", 11, "bold")).pack(
            anchor="w", padx=10)
        ent_cruz_nro = ctk.CTkEntry(f_cruzada)
        ent_cruz_nro.pack(fill="x", padx=10, pady=(0, 6))

        ctk.CTkLabel(f_cruzada, text="Pagado a (tercero):", font=("Arial", 11, "bold")).pack(
            anchor="w", padx=10)
        ent_cruz_tercero = ctk.CTkEntry(f_cruzada)
        ent_cruz_tercero.pack(fill="x", padx=10, pady=(0, 6))

        lbl_cruz_soporte = ctk.CTkLabel(f_cruzada, text="📄 Factura de compra: sin cargar\n🧾 Soporte del pago: sin cargar",
                                        font=("Arial", 11, "italic"), text_color="#7d3c98",
                                        wraplength=400, justify="left")
        lbl_cruz_soporte.pack(anchor="w", padx=10, pady=(0, 4))

        def actualizar_lbl_docs():
            fact = os.path.basename(estado_cruzada["factura"]) if estado_cruzada["factura"] else "sin cargar"
            sop = os.path.basename(estado_cruzada["soporte"]) if estado_cruzada["soporte"] else "sin cargar"
            lbl_cruz_soporte.configure(text=f"📄 Factura de compra: {fact}\n🧾 Soporte del pago: {sop}")

        def cargar_documentos_cruzada():
            """Un solo botón: carga la factura de compra y/o el comprobante del pago a tercero."""
            rutas = seleccionar_archivos_dialogo(
                titulo="Cargar Factura de Compra y/o Comprobante de Pago (puede elegir ambos)",
                tipos=[("Archivos PDF", "*.pdf"), ("Imágenes", "*.png;*.jpg;*.jpeg")])
            if not rutas:
                return
            for ruta in rutas:
                datos = {}
                if ruta.lower().endswith(".pdf"):
                    try:
                        datos = extraer_datos_factura_pdf(ruta) or {}
                    except Exception:
                        datos = {}
                # Se reconoce como factura de compra si el PDF trae N° de documento o RUC SUNAT
                es_factura = bool(datos.get("numero_documento") or datos.get("ruc"))
                if es_factura and not estado_cruzada["factura"]:
                    estado_cruzada["factura"] = ruta
                    if datos.get("proveedor") and not cmb_cruz_prov.get().strip():
                        cmb_cruz_prov.set(datos["proveedor"])
                    if datos.get("numero_documento"):
                        ent_cruz_nro.delete(0, tk.END)
                        ent_cruz_nro.insert(0, datos["numero_documento"])
                    if datos.get("fecha"):
                        ent_fecha.delete(0, tk.END)
                        ent_fecha.insert(0, datos["fecha"])
                    if datos.get("total") and not ent_monto.get().strip():
                        ent_monto.insert(0, formatear_numero_entrada(f"{datos['total']:.2f}"))
                        actualizar_total()
                else:
                    estado_cruzada["soporte"] = ruta
            actualizar_lbl_docs()

        def intercambiar_documentos():
            """Si el reconocimiento automático los asignó al revés, se intercambian."""
            estado_cruzada["factura"], estado_cruzada["soporte"] = \
                estado_cruzada["soporte"], estado_cruzada["factura"]
            actualizar_lbl_docs()

        def ver_documento_cruzada(cual):
            ruta = estado_cruzada.get(cual, "")
            if ruta and os.path.isfile(ruta):
                abrir_documento(ruta)
            else:
                messagebox.showinfo("Documentos",
                                    "Todavía no ha cargado " +
                                    ("la factura de compra." if cual == "factura"
                                     else "el comprobante del pago."),
                                    parent=v)

        f_btn_cruz = ctk.CTkFrame(f_cruzada, fg_color="transparent")
        f_btn_cruz.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkButton(f_btn_cruz, text="🧾 Cargar Factura de Compra y Comprobante de Pago", height=30,
                      font=("Arial", 11, "bold"), fg_color="#2980b9", hover_color="#1f618d",
                      command=cargar_documentos_cruzada).pack(fill="x")
        f_ver_cruz = ctk.CTkFrame(f_cruzada, fg_color="transparent")
        f_ver_cruz.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkButton(f_ver_cruz, text="👁 Ver factura", width=110, height=26, font=("Arial", 11),
                      fg_color="#34495e", hover_color="#2c3e50",
                      command=lambda: ver_documento_cruzada("factura")).pack(side="left")
        ctk.CTkButton(f_ver_cruz, text="👁 Ver comprobante", width=140, height=26, font=("Arial", 11),
                      fg_color="#34495e", hover_color="#2c3e50",
                      command=lambda: ver_documento_cruzada("soporte")).pack(side="left", padx=6)
        ctk.CTkButton(f_ver_cruz, text="🔄 Intercambiar", width=120, height=26, font=("Arial", 11),
                      fg_color="#8e44ad", hover_color="#703688",
                      command=intercambiar_documentos).pack(side="left")

        def on_cruzada():
            if var_cruzada.get():
                f_cruzada.pack(fill="x", pady=(2, 8))
            else:
                f_cruzada.pack_forget()
        chk_cruzada.configure(command=on_cruzada)

        def guardar():
            base = normalizar_monto(ent_monto.get())
            if base <= 0:
                messagebox.showerror("Error", "El monto debe ser mayor a 0.", parent=v)
                return
            com = normalizar_monto(ent_comision.get())
            inter = var_inter.get()
            total = base + (com if inter else 0)

            guardar_comision_interbancaria(f"{com:.2f}")

            principal = cmb_principal.get()
            sub = cmb_sub.get()
            if sub == "(Sin categoría)":
                sub = ""
            placa = ""
            if principal == "Gastos Operativos":
                pv = cmb_placa.get()
                if pv and pv != "(Sin placas)":
                    placa = pv
            chofer = ""
            if es_categoria_planilla(f"{principal} {sub}"):
                chv = cmb_chofer.get().strip()
                if chv and chv != "Ninguno":
                    chofer = chv
            fecha = ent_fecha.get().strip() or datetime.now().strftime("%d/%m/%Y")
            desc_manual = ent_desc.get().strip()

            tipo = "ingreso" if principal == "Ingresos" else "egreso"
            monto_final = total if tipo == "ingreso" else -total

            partes = []
            if principal:
                partes.append(principal)
            if sub:
                partes.append(sub)
            if placa:
                partes.append(f"Placa {placa}")
            if chofer:
                partes.append(f"Chofer {chofer}")
            if desc_manual:
                partes.append(desc_manual)
            if inter:
                partes.append(f"Comisión interbancaria {formatear_monto(com)}")
            desc = " - ".join(p for p in partes if p) or "Ajuste manual de conciliación"

            es_cruzada = bool(var_cruzada.get())
            if es_cruzada:
                if not cmb_cruz_prov.get().strip():
                    messagebox.showwarning("Compra cruzada", "Ingrese el proveedor de la factura.",
                                           parent=v)
                    return
                if not ent_cruz_tercero.get().strip():
                    messagebox.showwarning("Compra cruzada", "Ingrese quién pagó (tercero).", parent=v)
                    return
                faltantes = []
                if not (estado_cruzada["factura"] and os.path.isfile(estado_cruzada["factura"])):
                    faltantes.append("la factura de compra")
                if not (estado_cruzada["soporte"] and os.path.isfile(estado_cruzada["soporte"])):
                    faltantes.append("el comprobante del pago a tercero")
                if faltantes:
                    if not messagebox.askyesno(
                            "Compra cruzada",
                            "No se ha cargado " + " ni ".join(faltantes) + ".\n"
                            "¿Desea registrar la compra cruzada sin " +
                            ("esos documentos" if len(faltantes) > 1 else "ese documento") + "?",
                            parent=v):
                        return

            conn = conectar_db(silencioso=True)
            if conn:
                try:
                    with conn.cursor() as c:
                        if en_edicion:
                            c.execute("""
                                UPDATE conciliacion_bancaria
                                SET fecha=%s, descripcion=%s, monto=%s, tipo=%s,
                                    categoria=%s, subcategoria=%s, placa=%s, chofer=%s,
                                    comision=%s, interbancario=%s
                                WHERE id=%s
                            """, (fecha, desc, monto_final, tipo, principal, sub, placa, chofer,
                                  com, bool(inter), id_edicion))
                        else:
                            c.execute("""
                                INSERT INTO conciliacion_bancaria
                                (banco, cuenta, fecha, descripcion, monto, tipo, origen, id_movimiento, estado,
                                 categoria, subcategoria, placa, chofer, comision, interbancario, es_cruzada)
                                VALUES (%s, %s, %s, %s, %s, %s, 'manual', 0, 'pendiente',
                                        %s, %s, %s, %s, %s, %s, %s)
                            """, (banco.get("banco", ""), banco.get("cuenta", ""), fecha, desc,
                                  monto_final, tipo, principal, sub, placa, chofer, com,
                                  bool(inter), es_cruzada))
                        conn.commit()
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo guardar:\n{e}", parent=v)
                    liberar_conexion(conn)
                    return
                finally:
                    liberar_conexion(conn)
            registrar_auditoria(self.usuario_activo, "Banco",
                                (f"Editó el pago {formatear_monto(monto_final)} en {construir_etiqueta_banco(banco)}"
                                 if en_edicion else
                                 f"Agregó movimiento manual {formatear_monto(monto_final)} en {construir_etiqueta_banco(banco)}"))

            if es_cruzada:
                ok, estado_txt = self.registrar_compra_cruzada_desde_pago({
                    "proveedor": cmb_cruz_prov.get().strip(),
                    "nro": ent_cruz_nro.get().strip(),
                    "fecha": fecha,
                    "total": total,
                    "tercero": ent_cruz_tercero.get().strip(),
                    "categoria": f"{principal} - {sub}".strip(" -") if sub else principal,
                    "descripcion": desc,
                    "soporte_origen": estado_cruzada["soporte"],
                    "factura_origen": estado_cruzada["factura"],
                }, v)
                if not ok:
                    messagebox.showwarning(
                        "Compra cruzada",
                        "El movimiento del banco quedó registrado, pero la compra cruzada no se pudo "
                        "guardar en Compras. Vuelva a intentar el registro del pago.",
                        parent=v)
                    self.generar_reporte()
                    return
                messagebox.showinfo(
                    "Compra cruzada",
                    f"Compra cruzada {estado_txt} en el módulo de Compras "
                    f"(proveedor {cmb_cruz_prov.get().strip()}, pagada por "
                    f"{ent_cruz_tercero.get().strip()}).",
                    parent=v)

            if en_edicion:
                messagebox.showinfo("Pago actualizado",
                                    "Los datos del pago se actualizaron correctamente.", parent=v)
            v.destroy()
            self.generar_reporte()

        ctk.CTkButton(f_pie, text=("💾 Guardar Cambios" if en_edicion else "✅ Guardar"),
                      width=160, height=36, font=("Arial", 13, "bold"),
                      fg_color="#27ae60", hover_color="#1e8449",
                      command=guardar).pack(pady=10)
        ctk.CTkButton(f_pie, text="✖ Cancelar", width=110, height=36, font=("Arial", 12),
                      fg_color="#7f8c8d", hover_color="#606b6b",
                      command=v.destroy).pack(pady=(0, 10))

    def _formatear_entrada_monto(self, entry):
        """Da formato con separador de miles al escribir un monto en un campo de texto."""
        try:
            txt = entry.get()
        except Exception:
            return
        fmt = formatear_numero_entrada(txt)
        if fmt != txt:
            entry.delete(0, tk.END)
            entry.insert(0, fmt)

    def registrar_compra_cruzada_desde_pago(self, datos, parent=None):
        """Registra (o actualiza, sin duplicar) una compra cruzada en el módulo de Compras
        a partir del pago registrado en la pestaña de Conciliación.

        Devuelve (ok, estado) donde estado es 'creada' o 'actualizada'."""
        parent = parent or self.parent_frame
        proveedor = (datos.get("proveedor") or "").strip()
        tercero = (datos.get("tercero") or "").strip()
        if not proveedor:
            messagebox.showwarning("Compra cruzada", "Ingrese el proveedor de la factura.", parent=parent)
            return False, ""
        if not tercero:
            messagebox.showwarning("Compra cruzada", "Ingrese quién pagó (tercero).", parent=parent)
            return False, ""

        nro = (datos.get("nro") or "").strip().replace(" ", "")
        fecha = datos.get("fecha") or datetime.now().strftime("%d/%m/%Y")
        total = float(datos.get("total") or 0)
        categoria = datos.get("categoria") or "Compra Cruzada"
        descripcion = (datos.get("descripcion") or "").strip() or "Compra cruzada pagada por tercero"

        # Copiar los PDF (factura de compra y soporte del pago a tercero) a las carpetas autorizadas
        origen_soporte = datos.get("soporte_origen") or ""
        origen_factura = datos.get("factura_origen") or ""
        ruta_soporte = ""
        ruta_factura = ""

        if (origen_soporte and os.path.isfile(origen_soporte)) or            (origen_factura and os.path.isfile(origen_factura)):
            base = obtener_ruta_base()
            if not base:
                try:
                    from politica_almacenamiento import advertir
                    advertir(parent)
                except Exception:
                    messagebox.showwarning("Almacenamiento bloqueado",
                                           "Este equipo no está autorizado a guardar archivos.",
                                           parent=parent)
                return False, ""
            marca = datetime.now().strftime('%Y%m%d%H%M%S')
            if origen_factura and os.path.isfile(origen_factura):
                try:
                    carpeta = os.path.normpath(os.path.join(base, "facturas_recibidas"))
                    if not os.path.exists(carpeta):
                        os.makedirs(carpeta)
                    prov_limpio = re.sub(r'[\\/*?:"<>|]', '-', proveedor).replace(" ", "_")[:40]
                    ext = os.path.splitext(origen_factura)[1] or ".pdf"
                    ruta_factura = os.path.normpath(os.path.join(
                        carpeta, f"Cruzada_{marca}_{prov_limpio}{ext}"))
                    shutil.copy2(origen_factura, ruta_factura)
                except Exception as e:
                    messagebox.showerror("Error", f"Fallo al guardar la factura de compra:\n{e}",
                                         parent=parent)
                    return False, ""
            if origen_soporte and os.path.isfile(origen_soporte):
                try:
                    carpeta = os.path.normpath(os.path.join(base, "soportes_pagos_terceros"))
                    if not os.path.exists(carpeta):
                        os.makedirs(carpeta)
                    ext = os.path.splitext(origen_soporte)[1] or ".pdf"
                    ruta_soporte = os.path.normpath(os.path.join(
                        carpeta, f"Soporte_{marca}_{nro or 'sin_doc'}{ext}"))
                    shutil.copy2(origen_soporte, ruta_soporte)
                except Exception as e:
                    messagebox.showerror("Error", f"Fallo al guardar el soporte del pago a tercero:\n{e}",
                                         parent=parent)
                    return False, ""

        conn = conectar_db(silencioso=True)
        if not conn:
            messagebox.showerror("Error", "Sin conexión a la base de datos.", parent=parent)
            return False, ""

        actualizada = False
        try:
            with conn.cursor() as c:
                # Evitar duplicar: ¿la factura ya está registrada en Compras?
                existente = None
                if nro:
                    c.execute("""SELECT id, COALESCE(archivo_ruta, '') FROM facturas_recibidas
                                 WHERE REPLACE(numero_documento, ' ', '') = %s ORDER BY id LIMIT 1""", (nro,))
                    existente = c.fetchone()
                    if not existente and proveedor:
                        c.execute("""SELECT id, COALESCE(archivo_ruta, '') FROM facturas_recibidas
                                     WHERE REPLACE(numero_documento, ' ', '') = %s AND proveedor ILIKE %s
                                     ORDER BY id LIMIT 1""", (nro, proveedor))
                        existente = c.fetchone()

                if existente:
                    sets = ["es_compra_cruzada = TRUE", "categoria = %s", "pagado_por_tercero = %s"]
                    params = [categoria, tercero]
                    if ruta_soporte:
                        sets.append("soporte_pago_tercero = %s")
                        params.append(ruta_para_guardar(ruta_soporte))
                    # La factura ya adjunta no se reemplaza (igual que en el flujo anterior)
                    if ruta_factura and not existente[1]:
                        sets.append("archivo_ruta = %s")
                        params.append(ruta_para_guardar(ruta_factura))
                    params.append(existente[0])
                    c.execute("UPDATE facturas_recibidas SET " + ", ".join(sets) + " WHERE id = %s",
                              tuple(params))
                    actualizada = True
                else:
                    c.execute("""
                        INSERT INTO facturas_recibidas
                        (tipo_documento, numero_documento, fecha, proveedor, descripcion, evento_asociado,
                         subtotal, impuesto, total, archivo_ruta, dias_credito, det_porcentaje, det_monto,
                         categoria, ruc, pagado_por_tercero, soporte_pago_tercero, es_compra_cruzada)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """, ("FACTURA", nro, fecha, proveedor, descripcion, "",
                            total, 0, total, ruta_para_guardar(ruta_factura), 0, 0, 0, categoria, "", tercero,
                            ruta_para_guardar(ruta_soporte), True))
                conn.commit()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo registrar la compra cruzada:\n{e}", parent=parent)
            return False, ""
        finally:
            liberar_conexion(conn)

        registrar_auditoria(self.usuario_activo, "Banco",
                            f"Compra cruzada {nro or ''} de {proveedor} pagada a {tercero} "
                            f"(registrada desde Registrar Pago)")
        return True, ("actualizada" if actualizada else "creada")

    def gestionar_categorias(self):
        cats = cargar_categorias_gastos()
        v = ctk.CTkToplevel(self.parent_frame)
        v.title("Gestionar Categorías de Gastos")
        v.geometry("640x500")
        v.transient(self.parent_frame)
        v.grab_set()

        ctk.CTkLabel(v, text="⚙️ Categorías (principales y subcategorías)", font=("Arial", 14, "bold"),
                     text_color="#1f538d").pack(pady=(15, 5))

        f_tabla = ctk.CTkFrame(v, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True, padx=15, pady=10)
        columnas = ("principal", "categoria")
        tree = ttk.Treeview(f_tabla, columns=columnas, show="headings", selectmode="browse")
        tree.heading("principal", text="Principal")
        tree.heading("categoria", text="Categoría")
        tree.column("principal", width=220, anchor="w")
        tree.column("categoria", width=320, anchor="w")
        vsb = ttk.Scrollbar(f_tabla, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        def cargar_tree():
            tree.delete(*tree.get_children())
            for p, subs in cats.items():
                for s in subs:
                    tree.insert("", tk.END, values=(p, s))

        cargar_tree()

        def seleccion():
            sel = tree.selection()
            if not sel:
                return None, None
            vals = tree.item(sel[0], "values")
            return vals[0], vals[1]

        def add_principal():
            nombre = simpledialog.askstring("Nueva principal", "Nombre de la categoría principal:", parent=v)
            if nombre:
                nombre = nombre.strip()
                if nombre and nombre not in cats:
                    cats[nombre] = ["Varios"]
                    cargar_tree()

        def add_sub():
            p, s = seleccion()
            if not p:
                messagebox.showinfo("Aviso", "Seleccione una fila para saber a qué principal agregar.", parent=v)
                return
            nombre = simpledialog.askstring("Nueva categoría", f"Categoría para '{p}':", parent=v)
            if nombre:
                nombre = nombre.strip()
                if nombre and nombre not in cats.get(p, []):
                    cats.setdefault(p, []).append(nombre)
                    cargar_tree()

        def edit():
            p, s = seleccion()
            if not p:
                messagebox.showinfo("Aviso", "Seleccione una fila para editar.", parent=v)
                return
            if s:
                nuevo = simpledialog.askstring("Editar categoría", f"Editar '{s}' en '{p}':", parent=v, initialvalue=s)
                if nuevo:
                    nuevo = nuevo.strip()
                    if nuevo and nuevo != s:
                        lista = cats.get(p, [])
                        if s in lista:
                            lista[lista.index(s)] = nuevo
                        cargar_tree()
            else:
                nuevo = simpledialog.askstring("Editar principal", f"Editar '{p}':", parent=v, initialvalue=p)
                if nuevo:
                    nuevo = nuevo.strip()
                    if nuevo and nuevo != p:
                        cats[nuevo] = cats.pop(p, [])
                        cargar_tree()

        def delete():
            p, s = seleccion()
            if not p:
                messagebox.showinfo("Aviso", "Seleccione una fila para eliminar.", parent=v)
                return
            if s:
                if messagebox.askyesno("Eliminar", f"¿Eliminar la categoría '{s}' de '{p}'?", parent=v):
                    lista = cats.get(p, [])
                    if s in lista:
                        lista.remove(s)
                    cargar_tree()
            else:
                if messagebox.askyesno("Eliminar", f"¿Eliminar la principal '{p}' y todas sus categorías?", parent=v):
                    cats.pop(p, None)
                    cargar_tree()

        def guardar():
            for p in list(cats.keys()):
                if not cats.get(p):
                    cats[p] = ["Varios"]
            if guardar_categorias_gastos(cats):
                v.destroy()
                messagebox.showinfo("Éxito", "Categorías guardadas.", parent=self.parent_frame)
            else:
                messagebox.showerror("Error", "No se pudieron guardar las categorías.", parent=v)

        f_btns = ctk.CTkFrame(v, fg_color="transparent"); f_btns.pack(fill="x", padx=15, pady=(0, 15))
        ctk.CTkButton(f_btns, text="➕ Principal", width=120, command=add_principal).pack(side="left", padx=4)
        ctk.CTkButton(f_btns, text="➕ Categoría", width=120, command=add_sub).pack(side="left", padx=4)
        ctk.CTkButton(f_btns, text="✏️ Editar", width=100, command=edit).pack(side="left", padx=4)
        ctk.CTkButton(f_btns, text="🗑️ Eliminar", width=100, command=delete).pack(side="left", padx=4)
        ctk.CTkButton(f_btns, text="💾 Guardar y Cerrar", width=150, fg_color="#27ae60", command=guardar).pack(side="right", padx=4)

    def corregir_movimiento(self):
        filas = self._filas_seleccionadas()
        if len(filas) != 1:
            messagebox.showinfo("Corregir", "Seleccione exactamente un movimiento para corregir.",
                                parent=self.parent_frame)
            return
        f = filas[0]

        # Los pagos registrados desde este módulo ("➕ Registrar Pago / Compra
        # Cruzada") se editan en ESA MISMA ventana, con todos los datos cargados.
        if f.get("origen") == "manual" and f.get("id"):
            self.agregar_movimiento_manual(registro=f)
            return

        v = ctk.CTkToplevel(self.parent_frame)
        v.title("Corregir Movimiento")
        v.geometry("460x400")
        v.transient(self.parent_frame)
        v.grab_set()

        ctk.CTkLabel(v, text=f"✏️ Corregir: {f['descripcion'][:40]}", font=("Arial", 13, "bold"),
                     text_color="#1f538d").pack(pady=(15, 5))
        frm = ctk.CTkFrame(v, fg_color="transparent")
        frm.pack(fill="x", padx=20)

        ctk.CTkLabel(frm, text="Fecha:", font=("Arial", 11, "bold")).pack(anchor="w")
        f_fecha = ctk.CTkFrame(frm, fg_color="transparent")
        f_fecha.pack(fill="x", pady=(0, 8))
        ent_fecha = ctk.CTkEntry(f_fecha)
        ent_fecha.pack(side="left", fill="x", expand=True)
        ent_fecha.insert(0, f["fecha"])
        ctk.CTkButton(f_fecha, text="📅", width=42, font=("Arial", 13, "bold"),
                      fg_color="#1f538d", hover_color="#163b65",
                      command=lambda: CalendarioNativo(v, ent_fecha)).pack(side="left", padx=(6, 0))

        ctk.CTkLabel(frm, text="Descripción:", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_desc = ctk.CTkEntry(frm)
        ent_desc.pack(fill="x", pady=(0, 8)); ent_desc.insert(0, f["descripcion"])

        ctk.CTkLabel(frm, text="Monto (negativo = egreso):", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_monto = ctk.CTkEntry(frm)
        ent_monto.pack(fill="x", pady=(0, 8)); ent_monto.insert(0, f"{f['monto']:.2f}")

        def guardar():
            try:
                monto = float(ent_monto.get().strip())
            except ValueError:
                messagebox.showerror("Error", "Monto inválido.", parent=v)
                return
            fecha = ent_fecha.get().strip()
            desc = ent_desc.get().strip()
            origen = f.get("origen", "")

            if origen == "sistema":
                tabla = f.get("tabla", "")
                idp = f.get("id", 0)
                conn = conectar_db(silencioso=True)
                if conn and idp:
                    try:
                        with conn.cursor() as c:
                            if tabla == "pagos_clientes":
                                c.execute("UPDATE pagos_clientes SET monto_pagado=%s, fecha_pago=%s WHERE id=%s",
                                          (abs(monto), fecha, idp))
                            elif tabla == "pagos_comprobantes":
                                c.execute("UPDATE pagos_comprobantes SET monto_pagado=%s, fecha_pago=%s WHERE id=%s",
                                          (abs(monto), fecha, idp))
                            conn.commit()
                    except Exception as e:
                        messagebox.showerror("Error", f"No se pudo corregir:\n{e}", parent=v)
                        liberar_conexion(conn)
                        return
                    finally:
                        liberar_conexion(conn)
            elif origen == "manual" and f.get("id"):
                conn = conectar_db(silencioso=True)
                if conn:
                    try:
                        with conn.cursor() as c:
                            c.execute("UPDATE conciliacion_bancaria SET monto=%s, fecha=%s, descripcion=%s WHERE id=%s",
                                      (monto, fecha, desc, f["id"]))
                            conn.commit()
                    except Exception as e:
                        messagebox.showerror("Error", f"No se pudo corregir:\n{e}", parent=v)
                        liberar_conexion(conn)
                        return
                    finally:
                        liberar_conexion(conn)

            f["monto"] = monto
            f["fecha"] = fecha
            f["descripcion"] = desc
            registrar_auditoria(self.usuario_activo, "Banco", f"Corrigió movimiento: {desc[:60]}")
            v.destroy()
            self.refrescar_tree()

        ctk.CTkButton(v, text="✅ Guardar", width=140, fg_color="#27ae60", command=guardar).pack(pady=10)

    def eliminar_movimiento_conciliacion(self):
        sel = self.tabla_conc.selection()
        if not sel:
            messagebox.showinfo("Eliminar", "Seleccione uno o más movimientos para eliminar.",
                                parent=self.parent_frame)
            return
        indices = sorted({int(i) for i in sel if str(i).isdigit()})
        filas = [self.filas_conciliacion[i] for i in indices
                 if 0 <= i < len(self.filas_conciliacion)]
        if not filas:
            return

        if not messagebox.askyesno(
                "Confirmar eliminación",
                f"¿Eliminar {len(filas)} movimiento(s) de la conciliación?\n\n"
                "• Los movimientos manuales se borran de la base de datos.\n"
                "• Los movimientos del sistema / estado de cuenta pierden su marca de conciliado "
                "y vuelven a quedar pendientes.\n"
                "• Las líneas leídas del PDF desaparecen de la vista hasta que vuelva a cargar el PDF.",
                parent=self.parent_frame):
            return

        banco = self.banco_seleccionado()
        nombre_banco = banco.get("banco", "") if banco else ""
        borrados_db = 0
        errores = []

        conn = conectar_db(silencioso=True)
        try:
            if conn:
                with conn.cursor() as c:
                    for f in filas:
                        origen = f.get("origen", "")
                        idp = f.get("id", 0) or 0
                        try:
                            if origen == "manual" and idp:
                                c.execute("DELETE FROM conciliacion_bancaria WHERE id = %s", (idp,))
                                borrados_db += max(c.rowcount, 0)
                            elif origen == "sistema" and idp:
                                c.execute("""DELETE FROM conciliacion_bancaria
                                             WHERE origen = 'sistema' AND id_movimiento = %s AND banco = %s""",
                                          (idp, nombre_banco))
                                borrados_db += max(c.rowcount, 0)
                            elif origen == "estado_cuenta":
                                c.execute("""DELETE FROM conciliacion_bancaria
                                             WHERE origen = 'estado_cuenta' AND banco = %s
                                               AND fecha = %s AND monto = %s AND descripcion = %s""",
                                          (nombre_banco, f.get("fecha", ""),
                                           f.get("monto", 0), f.get("descripcion", "")))
                                borrados_db += max(c.rowcount, 0)
                        except Exception as e:
                            errores.append(str(e))
                    conn.commit()
        except Exception as e:
            errores.append(str(e))
        finally:
            if conn:
                liberar_conexion(conn)

        excluir = set(indices)
        self.filas_conciliacion = [f for i, f in enumerate(self.filas_conciliacion)
                                   if i not in excluir]
        self.refrescar_tree()

        registrar_auditoria(self.usuario_activo, "Banco",
                            f"Eliminó {len(filas)} movimiento(s) de la conciliación "
                            f"en {construir_etiqueta_banco(banco) if banco else 'banco'}")

        msg = f"{len(filas)} movimiento(s) eliminado(s) de la conciliación."
        if borrados_db:
            msg += f"\n{borrados_db} registro(s) eliminado(s) de la base de datos."
        if errores:
            msg += f"\n⚠️ Algunos registros no se pudieron borrar: {errores[0]}"
        messagebox.showinfo("Eliminar", msg, parent=self.parent_frame)

    def exportar_reporte(self):
        if not self.filas_conciliacion:
            messagebox.showinfo("Exportar", "No hay movimientos para exportar.", parent=self.parent_frame)
            return
        ruta = guardar_archivo_dialogo(
            titulo="Guardar Reporte de Conciliación",
            defaultextension=".txt",
            tipos=[("Archivo de texto", "*.txt"), ("CSV", "*.csv")])
        if not ruta:
            return
        try:
            es_csv = ruta.lower().endswith(".csv")
            sep = ";" if es_csv else "\t"
            with open(ruta, "w", encoding="utf-8-sig") as out:
                cab = ["Estado", "Fecha", "Descripción", "Tipo", "Monto", "Origen"]
                out.write(sep.join(cab) + "\n")
                for f in self.filas_conciliacion:
                    tipo = "Ingreso" if f["monto"] >= 0 else "Egreso"
                    estado = f.get("estado", "pendiente")
                    out.write(sep.join([estado, f["fecha"], f["descripcion"],
                                        tipo, f"{f['monto']:.2f}", f.get("origen", "")]) + "\n")
            messagebox.showinfo("Exportar", f"Reporte guardado en:\n{ruta}", parent=self.parent_frame)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar:\n{e}", parent=self.parent_frame)

    def mostrar_texto_pdf(self):
        if not self.texto_pdf:
            messagebox.showinfo("Texto PDF", "Aún no se ha cargado ningún PDF.", parent=self.parent_frame)
            return
        v = ctk.CTkToplevel(self.parent_frame)
        v.title("Texto Extraído del Estado de Cuenta")
        v.geometry("820x560")
        v.transient(self.parent_frame)
        txt = tk.Text(v, wrap="word", font=("Consolas", 10))
        txt.pack(fill="both", expand=True, padx=10, pady=10)
        txt.insert("1.0", self.texto_pdf)
        txt.configure(state="disabled")
        ctk.CTkButton(v, text="Cerrar", width=120, fg_color="#34495e",
                      command=v.destroy).pack(pady=(0, 10))
