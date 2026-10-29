# -*- coding: utf-8 -*-
"""📋 Módulo de Amonestaciones y Suspensiones.

Permite registrar amonestaciones escritas y suspensiones temporales a los
choferes, con motivos configurables, historial con búsqueda y generación de
un PDF corporativo (logotipo, texto formal y firma del Gerente de Operaciones).
"""
import os
import sys
import subprocess
import json
import datetime
import tkinter as tk
from tkinter import messagebox, ttk

import customtkinter as ctk

from conexion import conectar_db, liberar_conexion, registrar_auditoria
from pathlib import Path

from app_paths import (DATA_DIR, BASE_DIR, ruta_drive_configurada,
                       ruta_para_guardar, resolver_ruta_archivo)

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas as _canvas
    _PDF_OK = True
except Exception:            # pragma: no cover
    _PDF_OK = False


def abrir_documento(ruta, titulo="Documento"):
    """🖼️ Muestra el documento en el VISOR INTERNO del programa.

    Antes se abría con el programa externo (Acrobat, Fotos…) y la ventana
    quedaba muchas veces detrás del sistema. Ahora se ve dentro del programa.
    """
    try:
        from visor_documentos import ver_documento_ventana
        return ver_documento_ventana(ruta, titulo)
    except Exception as e_visor:
        print("Aviso - visor interno:", e_visor)
    return abrir_documento_externo(ruta)


def abrir_documento_externo(ruta):
    """Abre un PDF/archivo con la aplicación del sistema."""
    ruta = str(ruta or "")
    if not ruta or not Path(ruta).is_file():
        raise RuntimeError("El archivo no existe: " + ruta)
    try:
        if os.name == "nt":
            os.startfile(ruta)
        elif sys.platform == "darwin":
            subprocess.call(["open", ruta])
        else:
            subprocess.call(["xdg-open", ruta])
    except Exception as e:
        raise RuntimeError(str(e))


MOTIVOS_BASE = (
    "Uniforme incompleto o en mal estado",
    "Llegada tarde / tardanza reiterada",
    "Falta injustificada",
    "Abandono del puesto de trabajo",
    "Uso indebido del vehículo de la empresa",
    "Incumplimiento de ruta o servicio asignado",
    "Falta de respeto al cliente o a un compañero",
    "No reportar incidentes o accidentes",
    "Uso del celular durante la conducción",
    "Documentación vencida (licencia, SOAT, revisión)",
    "No realizar el mantenimiento o limpieza del vehículo",
    "Consumo de alcohol o sustancias prohibidas",
)

ARCHIVO_CONFIG = Path(str(DATA_DIR)) / "amonestaciones.json"

TEXTO_AMONESTACION = """Por medio de la presente, la empresa deja constancia formal de la AMONESTACIÓN ESCRITA dirigida al colaborador antes mencionado, por el siguiente motivo:

{motivo}

Los hechos descritos contravienen las normas internas de conducta, seguridad y buen desempeño establecidas por la empresa, así como los principios de disciplina, respeto y responsabilidad que rigen nuestra relación laboral.

Se le exhorta a corregir de manera inmediata la conducta observada y a mantener el cumplimiento estricto de sus obligaciones, horarios, presentación personal, uso del uniforme y demás disposiciones internas de la empresa.

Se deja constancia que la reincidencia en este tipo de faltas dará lugar a la aplicación de medidas disciplinarias de mayor gravedad, incluyendo la suspensión temporal de labores o la terminación del vínculo laboral, conforme a la normativa vigente.

El colaborador firma la presente en señal de recepción, quedando copia de este documento en su legajo personal."""

HECHOS_SUGERIDOS = """El día __/__/____, el colaborador hizo abandono injustificado de su puesto de trabajo a partir de las __:__ horas, retirándose de las instalaciones de la empresa sin autorización de su jefatura directa ni justificación válida, y sin retornar a culminar su turno laboral.

Describa también, si corresponde: intentos de contacto, testigos, servicios afectados y cualquier antecedente relevante."""

DESCARGOS = """[Opcional para el debido proceso] Descargos del colaborador (opcional):"""

TEXTO_SUSPENSION = """Por medio de la presente, la empresa deja constancia formal de la SUSPENSIÓN TEMPORAL DE LABORES impuesta al colaborador antes mencionado, por el siguiente motivo:

{motivo}

La medida se adopta por {dias} día(s) calendario(s), con fecha de inicio el {inicio}, luego de haber evaluado los hechos ocurridos y de habérsele otorgado la oportunidad de presentar sus descargos, los cuales fueron valorados por la empresa.

Durante el período de suspensión, el colaborador no deberá asistir a su centro de labores ni realizar actividad alguna en nombre de la empresa, quedando suspendidos temporalmente los efectos del contrato de trabajo en lo referido a la prestación de servicios y al pago de la remuneración, conforme a ley.

Se deja constancia que la presente medida no constituye un despido y que, culminado el plazo indicado, el colaborador deberá reincorporarse a su puesto y horario habituales, manteniendo las condiciones laborales pactadas.

El colaborador firma la presente en señal de recepción, quedando copia de este documento en su legajo personal."""


# --------------------------------------------------------------------------
# Base de datos
# --------------------------------------------------------------------------
def asegurar_schema():
    """Crea las tablas del módulo y carga los motivos iniciales."""
    conn = conectar_db(silencioso=True)
    if not conn:
        return False
    try:
        with conn.cursor() as c:
            c.execute("""
                CREATE TABLE IF NOT EXISTS amonestaciones (
                    id SERIAL PRIMARY KEY,
                    tipo VARCHAR(30),
                    id_chofer INTEGER,
                    chofer VARCHAR(200),
                    dni VARCHAR(20),
                    motivo VARCHAR(250),
                    detalle TEXT,
                    fecha VARCHAR(20),
                    fecha_inicio VARCHAR(20),
                    dias INTEGER DEFAULT 0,
                    gerente VARCHAR(200),
                    archivo_ruta VARCHAR(600),
                    creado TIMESTAMP DEFAULT now()
                )
            """)
            c.execute("""
                CREATE TABLE IF NOT EXISTS motivos_amonestacion (
                    id SERIAL PRIMARY KEY,
                    motivo VARCHAR(250) UNIQUE,
                    activo BOOLEAN DEFAULT TRUE,
                    creado TIMESTAMP DEFAULT now()
                )
            """)
            # 🔢 Número de documento (AMO-2026-0001 / SUS-2026-0001)
            c.execute("ALTER TABLE amonestaciones ADD COLUMN IF NOT EXISTS numero_documento VARCHAR(30)")
            c.execute("CREATE UNIQUE INDEX IF NOT EXISTS ux_amonestaciones_numero "
                      "ON amonestaciones (numero_documento) WHERE numero_documento IS NOT NULL")
            c.execute("SELECT COUNT(*) FROM motivos_amonestacion")
            if not c.fetchone()[0]:
                for motivo in MOTIVOS_BASE:
                    c.execute(
                        "INSERT INTO motivos_amonestacion (motivo) VALUES (%s) ON CONFLICT DO NOTHING",
                        (motivo,))
        conn.commit()
        return True
    except Exception as e:
        print("[amonestaciones] schema:", e)
        try:
            conn.rollback()
        except Exception:
            pass
        return False
    finally:
        liberar_conexion(conn)


def cargar_motivos():
    """Lista de motivos activos (con respaldo si falla la base)."""
    conn = conectar_db(silencioso=True)
    motivos = []
    if conn:
        try:
            with conn.cursor() as c:
                c.execute("SELECT motivo FROM motivos_amonestacion WHERE activo ORDER BY motivo")
                motivos = [str(r[0]) for r in c.fetchall() if r[0]]
        except Exception as e:
            print("[amonestaciones] motivos:", e)
        finally:
            liberar_conexion(conn)
    return motivos or list(MOTIVOS_BASE)


def guardar_motivo(motivo):
    """Agrega un motivo nuevo. Devuelve (ok, mensaje)."""
    motivo = str(motivo or "").strip()
    if not motivo:
        return False, "Escriba el motivo."
    conn = conectar_db(silencioso=True)
    if not conn:
        return False, "No hay conexión con la base."
    try:
        with conn.cursor() as c:
            c.execute("SELECT 1 FROM motivos_amonestacion WHERE UPPER(TRIM(motivo)) = UPPER(TRIM(%s))",
                      (motivo,))
            if c.fetchone():
                return False, "Ese motivo ya existe."
            c.execute("INSERT INTO motivos_amonestacion (motivo) VALUES (%s)", (motivo,))
        conn.commit()
        return True, "Motivo agregado."
    except Exception as e:
        return False, str(e)
    finally:
        liberar_conexion(conn)


def _prefijo_documento(tipo):
    """AMO para amonestaciones, SUS para suspensiones (series independientes)."""
    return "SUS" if "suspens" in str(tipo or "").lower() else "AMO"


def siguiente_numero_documento(tipo):
    """Devuelve el próximo número correlativo del año para ese tipo.

    Formato: AMO-2026-0001 / SUS-2026-0001. Cada serie lleva su propio
    correlativo, así que nunca se repite un número entre amonestaciones y
    suspensiones. Se calcula desde la base (no es editable por el usuario).
    """
    prefijo = _prefijo_documento(tipo)
    anio = datetime.datetime.now().year
    base = "%s-%d-" % (prefijo, anio)
    conn = conectar_db(silencioso=True)
    if not conn:
        return base + "0001"
    try:
        with conn.cursor() as c:
            c.execute("SELECT numero_documento FROM amonestaciones "
                      "WHERE numero_documento LIKE %s ORDER BY numero_documento DESC LIMIT 1",
                      (base + "%",))
            fila = c.fetchone()
            ultimo = 0
            if fila and fila[0]:
                try:
                    ultimo = int(str(fila[0]).split("-")[-1])
                except Exception:
                    ultimo = 0
            return base + "%04d" % (ultimo + 1)
    except Exception as e:
        print("[amonestaciones] numeración:", e)
        return base + "0001"
    finally:
        liberar_conexion(conn)


def listar_amonestaciones(filtro=""):
    """Historial (más recientes primero), con búsqueda por texto."""
    conn = conectar_db(silencioso=True)
    if not conn:
        return []
    filas = []
    try:
        with conn.cursor() as c:
            sql = ("SELECT id, tipo, fecha, chofer, dni, motivo, gerente, COALESCE(archivo_ruta,''), "
                   "COALESCE(detalle,''), COALESCE(dias,0), COALESCE(fecha_inicio,''), "
                   "COALESCE(numero_documento,'') "
                   "FROM amonestaciones")
            params = ()
            if filtro:
                sql += (" WHERE chofer ILIKE %s OR motivo ILIKE %s OR tipo ILIKE %s OR dni ILIKE %s")
                patron = "%" + filtro + "%"
                params = (patron, patron, patron, patron)
            sql += " ORDER BY id DESC"
            c.execute(sql, params)
            filas = c.fetchall()
    except Exception as e:
        print("[amonestaciones] historial:", e)
    finally:
        liberar_conexion(conn)
    return filas


# --------------------------------------------------------------------------
# Configuración (nombre del gerente)
# --------------------------------------------------------------------------
def cargar_config():
    try:
        if ARCHIVO_CONFIG.is_file():
            with open(ARCHIVO_CONFIG, "r", encoding="utf-8") as f:
                return json.load(f) or {}
    except Exception:
        pass
    return {}


def guardar_config(datos):
    try:
        ARCHIVO_CONFIG.parent.mkdir(parents=True, exist_ok=True)
        with open(ARCHIVO_CONFIG, "w", encoding="utf-8") as f:
            json.dump(datos or {}, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("[amonestaciones] config:", e)


# --------------------------------------------------------------------------
# PDF corporativo
# --------------------------------------------------------------------------
def _ruta_logo():
    """Busca el logotipo en la carpeta del programa y junto a este módulo."""
    nombres = ("Logo.png", "logo.png", "LogoCotizacion.png", "logo_encabezado.png")
    carpetas = []
    try:
        carpetas.append(str(BASE_DIR))
    except Exception:
        pass
    carpetas.append(str(Path(__file__).resolve().parent))
    for carpeta in carpetas:
        for nombre in nombres:
            ruta = Path(carpeta) / nombre
            if ruta.is_file():
                return str(ruta)
    return ""


def _nombre_archivo_seguro(texto, maximo=30):
    """Nombre de archivo valido en Windows y macOS (sin caracteres prohibidos)."""
    base = str(texto or "chofer").strip()
    limpio = []
    for letra in base:
        if letra.isalnum() or letra in (" ", "_", "-"):
            limpio.append(letra)
        else:
            limpio.append("_")
    resultado = "".join(limpio).replace(" ", "_").strip("_.-")
    while "__" in resultado:
        resultado = resultado.replace("__", "_")
    return (resultado[:maximo] or "chofer")


def generar_pdf(datos, ruta_destino, gerente, empresa="BLACK RIDERS E.I.R.L.", ruc=""):
    """Genera el PDF de la amonestación o suspensión. Devuelve la ruta."""
    if not _PDF_OK:
        raise RuntimeError("Falta la librería reportlab.")
    Path(str(ruta_destino)).parent.mkdir(parents=True, exist_ok=True)
    ancho, alto = A4
    lienzo = _canvas.Canvas(ruta_destino, pagesize=A4)

    # ---------- encabezado con logo ----------
    margen = 46
    y = alto - 52
    logo = _ruta_logo()
    if logo:
        try:
            imagen = ImageReader(logo)
            lienzo.drawImage(imagen, margen, y - 46, width=92, height=58,
                             preserveAspectRatio=True, mask="auto")
        except Exception:
            pass
    lienzo.setFont("Helvetica-Bold", 13)
    lienzo.drawString(margen + 105, y - 8, empresa)
    lienzo.setFont("Helvetica", 9)
    if ruc:
        lienzo.drawString(margen + 105, y - 21, "RUC: " + str(ruc))
    lienzo.drawString(margen + 105, y - 32, "Gestión de Personas y Operaciones")
    y -= 62
    lienzo.setLineWidth(1.2)
    lienzo.line(margen, y, ancho - margen, y)
    # 🔢 Número del documento (arriba a la derecha)
    numero_doc = str(datos.get("numero") or "").strip()
    if numero_doc:
        lienzo.setFont("Helvetica-Bold", 10)
        lienzo.drawRightString(ancho - margen, y + 14, "N° " + numero_doc)
    y -= 34

    # ---------- título ----------
    tipo = str(datos.get("tipo") or "")
    titulo = "SUSPENSIÓN TEMPORAL DE LABORES" if "suspens" in tipo.lower() else "AMONESTACIÓN ESCRITA"
    lienzo.setFont("Helvetica-Bold", 14)
    lienzo.drawCentredString(ancho / 2, y, titulo)
    y -= 16
    lienzo.setFont("Helvetica", 9)
    lienzo.drawCentredString(ancho / 2, y, "Documento interno de carácter laboral")
    y -= 30

    # ---------- datos del colaborador ----------
    lienzo.setFont("Helvetica-Bold", 10)
    lienzo.drawString(margen, y, "DATOS DEL TRABAJADOR")
    y -= 16
    lienzo.setFont("Helvetica", 10)
    lienzo.drawString(margen, y, "Nombre: " + str(datos.get("chofer") or ""))
    y -= 15
    lienzo.drawString(margen, y, "DNI: " + (str(datos.get("dni") or "—")))
    y -= 15
    lienzo.drawString(margen, y, "Fecha del documento: " + str(datos.get("fecha") or ""))
    y -= 15
    lienzo.drawString(margen, y, "Motivo: " + str(datos.get("motivo") or ""))
    if "suspens" in tipo.lower():
        y -= 15
        lienzo.drawString(margen, y, "Días de suspensión: %s   |   Inicio: %s"
                          % (str(datos.get("dias") or ""), str(datos.get("fecha_inicio") or "")))
    y -= 26

    # ---------- cuerpo ----------
    detalle = str(datos.get("detalle") or "").strip()
    texto = (TEXTO_SUSPENSION if "suspens" in tipo.lower() else TEXTO_AMONESTACION).format(
        motivo=str(datos.get("motivo") or ""),
        dias=str(datos.get("dias") or ""),
        inicio=str(datos.get("fecha_inicio") or ""))
    if detalle:
        texto += chr(10) * 2 + "Observaciones: " + detalle
    lienzo.setFont("Helvetica", 10)
    for parrafo in str(texto).split(chr(10)):
        if not parrafo.strip():
            y -= 8
            continue
        for linea in _envolver(parrafo, 92):
            if y <= 150:
                lienzo.showPage()
                lienzo.setFont("Helvetica", 10)
                y = alto - 70
            lienzo.drawString(margen, y, linea)
            y -= 14

    # ---------- firma ----------
    y = max(y, 120)
    if y > 190:
        y = 190
    lienzo.setLineWidth(0.8)
    lienzo.line(margen, y, margen + 210, y)
    lienzo.setFont("Helvetica-Bold", 10)
    lienzo.drawString(margen, y - 14, str(gerente or ""))
    lienzo.setFont("Helvetica", 9)
    lienzo.drawString(margen, y - 26, "Gerente de Operaciones")
    lienzo.line(ancho - margen - 210, y, ancho - margen, y)
    lienzo.setFont("Helvetica-Bold", 10)
    lienzo.drawString(ancho - margen - 210, y - 14, str(datos.get("chofer") or ""))
    lienzo.setFont("Helvetica", 9)
    lienzo.drawString(ancho - margen - 210, y - 26, "Colaborador")

    lienzo.setFont("Helvetica-Oblique", 7.5)
    lienzo.drawCentredString(ancho / 2, 40,
                             "Documento generado por el Sistema de Control de Flota — copia para el legajo personal")
    lienzo.showPage()
    lienzo.save()
    return ruta_destino


def _envolver(texto, maximo):
    """Parte un texto en líneas de hasta 'maximo' caracteres."""
    palabras = str(texto).split()
    lineas, actual = [], ""
    for palabra in palabras:
        if len(actual) + len(palabra) + 1 <= maximo:
            actual = (actual + " " + palabra).strip()
        else:
            lineas.append(actual)
            actual = palabra
    if actual:
        lineas.append(actual)
    return lineas or [""]


# --------------------------------------------------------------------------
# Interfaz
# --------------------------------------------------------------------------
class AmonestacionesApp(ctk.CTkFrame):
    """Pantalla principal del módulo de Amonestaciones y Suspensiones."""

    def __init__(self, contenedor, usuario="Sistema"):
        super().__init__(contenedor, fg_color="transparent")
        self.pack(fill="both", expand=True)
        self.usuario = usuario
        self.config = cargar_config()
        asegurar_schema()
        self._construir()

    # ------------------------------------------------------------------
    def _construir(self):
        titulo = ctk.CTkFrame(self, fg_color="#1f538d", corner_radius=8)
        titulo.pack(fill="x", padx=10, pady=(10, 6))
        ctk.CTkLabel(titulo, text="📋 Amonestaciones y Suspensiones",
                     font=("Arial", 18, "bold"), text_color="white").pack(side="left", padx=14, pady=10)
        ctk.CTkButton(titulo, text="📚 Ver Historial", width=150, font=("Arial", 12, "bold"),
                      fg_color="#2c3e50", hover_color="#1b2631",
                      command=self.abrir_historial).pack(side="right", padx=10)

        cuerpo = ctk.CTkFrame(self, fg_color="transparent")
        cuerpo.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # ---------------- formulario ----------------
        self.f_form = ctk.CTkScrollableFrame(cuerpo, label_text="Nuevo documento",
                                             width=430, fg_color="#f8f9fa")
        self.f_form.pack(side="left", fill="y", padx=(0, 12))

        ctk.CTkLabel(self.f_form, text="Tipo de documento:",
                     font=("Arial", 11, "bold")).pack(anchor="w", padx=10, pady=(8, 2))
        self.combo_tipo = ctk.CTkComboBox(self.f_form,
                                          values=["Amonestación", "Suspensión"],
                                          state="readonly", command=self._al_cambiar_tipo)
        self.combo_tipo.pack(fill="x", padx=10, pady=(0, 8))
        self.combo_tipo.set("Amonestación")

        ctk.CTkLabel(self.f_form, text="Chofer:",
                     font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.combo_chofer = ctk.CTkComboBox(self.f_form, values=[], command=self._al_elegir_chofer)
        self.combo_chofer.pack(fill="x", padx=10, pady=(0, 8))

        fila_motivo = ctk.CTkFrame(self.f_form, fg_color="transparent")
        fila_motivo.pack(fill="x", padx=10)
        ctk.CTkLabel(fila_motivo, text="Motivo:",
                     font=("Arial", 11, "bold")).pack(side="left")
        self.combo_motivo = ctk.CTkComboBox(self.f_form, values=cargar_motivos())
        self.combo_motivo.pack(fill="x", padx=10, pady=(0, 4))
        ctk.CTkButton(self.f_form, text="➕ Crear nuevo motivo", height=26,
                      font=("Arial", 11, "bold"), fg_color="#8e44ad", hover_color="#703688",
                      command=self.crear_motivo).pack(fill="x", padx=10, pady=(0, 8))

        ctk.CTkLabel(self.f_form, text="Fecha del documento:",
                     font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.ent_fecha = ctk.CTkEntry(self.f_form)
        self.ent_fecha.pack(fill="x", padx=10, pady=(0, 8))
        self.ent_fecha.insert(0, datetime.date.today().strftime("%d/%m/%Y"))

        self.f_suspension = ctk.CTkFrame(self.f_form, fg_color="#fdf2e3", corner_radius=6,
                                          border_width=1, border_color="#e3c39a")
        ctk.CTkLabel(self.f_suspension, text="Días de suspensión:",
                     font=("Arial", 11, "bold")).pack(anchor="w", padx=8, pady=(6, 0))
        self.ent_dias = ctk.CTkEntry(self.f_suspension, border_width=0, fg_color="#ffffff")
        self.ent_dias.pack(fill="x", padx=8, pady=(0, 4))
        self.ent_dias.insert(0, "1")
        ctk.CTkLabel(self.f_suspension, text="Fecha de inicio de la suspensión:",
                     font=("Arial", 11, "bold")).pack(anchor="w", padx=8)
        self.ent_inicio = ctk.CTkEntry(self.f_suspension, border_width=0, fg_color="#ffffff")
        self.ent_inicio.pack(fill="x", padx=8, pady=(0, 8))
        self.ent_inicio.insert(0, datetime.date.today().strftime("%d/%m/%Y"))
        # (se muestra solo cuando el tipo es Suspensión)

        ctk.CTkLabel(self.f_form, text="Detalle / observaciones:",
                     font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        # ✏️ Cuadro con borde y aviso: se entiende que aquí se escribe
        self.marco_detalle = ctk.CTkFrame(self.f_form, fg_color="#ffffff", corner_radius=6,
                                          border_width=1, border_color="#9aa5b1")
        self.marco_detalle.pack(fill="x", padx=10, pady=(2, 8))
        ctk.CTkLabel(self.marco_detalle, text="✏️ Haga clic aquí y escriba las observaciones",
                     font=("Arial", 9), text_color="#6b7785").pack(anchor="w", padx=8, pady=(5, 0))
        # el cuadro exterior ya dibuja el borde: el de adentro va SIN borde
        self.txt_detalle = ctk.CTkTextbox(self.marco_detalle, height=82, border_width=0,
                                          corner_radius=0, fg_color="#ffffff")
        self.txt_detalle.pack(fill="x", padx=10, pady=(0, 8))
        self._placeholder_detalle = "Ej.: Se presentó sin corbata y con la camisa arrugada..."
        self.txt_detalle.insert("1.0", self._placeholder_detalle)
        self.txt_detalle.configure(text_color="#9aa5b1")
        self.txt_detalle.bind("<FocusIn>", lambda e: self._quitar_placeholder(), add="+")
        self.txt_detalle.bind("<Key>", lambda e: self._quitar_placeholder(), add="+")
        self.txt_detalle.bind("<FocusOut>", lambda e: self._restaurar_placeholder(), add="+")
        self._placeholder_activo = True

        ctk.CTkLabel(self.f_form, text="Gerente de Operaciones (firma):",
                     font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.ent_gerente = ctk.CTkEntry(self.f_form,
                                        placeholder_text="Nombre completo del gerente")
        self.ent_gerente.pack(fill="x", padx=10, pady=(0, 4))
        if self.config.get("gerente"):
            self.ent_gerente.insert(0, str(self.config["gerente"]))
        ctk.CTkLabel(self.f_form,
                     text="El nombre se recuerda para los próximos documentos.",
                     font=("Arial", 9), text_color="gray").pack(anchor="w", padx=10, pady=(0, 8))

        self.btn_pdf = ctk.CTkButton(self.f_form, text="📄 Generar PDF y Guardar",
                                     font=("Arial", 13, "bold"), height=38,
                                     fg_color="#27ae60", hover_color="#1e8449",
                                     command=self.generar)
        self.btn_pdf.pack(fill="x", padx=10, pady=(6, 4))
        ctk.CTkButton(self.f_form, text="🧹 Limpiar formulario", font=("Arial", 11),
                      fg_color="#7f8c8d", hover_color="#606b6b",
                      command=self.limpiar).pack(fill="x", padx=10, pady=(0, 12))

        # ---------------- vista previa del texto ----------------
        self.f_previa = ctk.CTkFrame(cuerpo, fg_color="#ffffff", corner_radius=8)
        self.f_previa.pack(side="left", fill="both", expand=True)
        ctk.CTkLabel(self.f_previa, text="Vista previa del documento",
                     font=("Arial", 14, "bold"), text_color="#1f538d").pack(anchor="w", padx=14, pady=(12, 4))
        self.txt_previa = ctk.CTkTextbox(self.f_previa, font=("Arial", 11), wrap="word")
        self.txt_previa.pack(fill="both", expand=True, padx=14, pady=(0, 12))

        self._cargar_choferes()
        self._al_cambiar_tipo("Amonestación")
        self._actualizar_previa()

    # ------------------------------------------------------------------
    def _cargar_choferes(self):
        conn = conectar_db(silencioso=True)
        valores = []
        if conn:
            try:
                with conn.cursor() as c:
                    c.execute("SELECT nombres FROM choferes ORDER BY nombres")
                    valores = [str(r[0]).strip() for r in c.fetchall() if r[0]]
            except Exception as e:
                print("[amonestaciones] choferes:", e)
            finally:
                liberar_conexion(conn)
        self.combo_chofer.configure(values=valores)
        if valores:
            self.combo_chofer.set(valores[0])
        self._al_elegir_chofer()

    def _dni_chofer(self, nombre):
        conn = conectar_db(silencioso=True)
        dni = ""
        if conn:
            try:
                with conn.cursor() as c:
                    c.execute("SELECT COALESCE(dni,'') FROM choferes WHERE nombres = %s LIMIT 1",
                              (nombre,))
                    fila = c.fetchone()
                    dni = str(fila[0]) if fila and fila[0] else ""
            except Exception:
                pass
            finally:
                liberar_conexion(conn)
        return dni

    def _al_elegir_chofer(self, valor=None):
        self._actualizar_previa()

    def _al_cambiar_tipo(self, valor=None):
        es_suspension = "suspens" in str(self.combo_tipo.get()).lower()
        if es_suspension:
            try:
                self.f_suspension.pack(fill="x", padx=10, pady=(0, 8),
                                       before=self.marco_detalle)   # mismo padre: f_form
            except Exception:
                self.f_suspension.pack(fill="x", padx=10, pady=(0, 8))
        else:
            self.f_suspension.pack_forget()
        self._actualizar_previa()

    def insertar_hechos(self):
        """Rellena el cuadro con el texto sugerido de la descripción de los hechos."""
        try:
            self._quitar_placeholder()
            self.txt_detalle.delete("1.0", "end")
            self.txt_detalle.insert("1.0", HECHOS_SUGERIDOS)
            self._placeholder_activo = False
            self._actualizar_previa()
        except Exception as e:
            print("[amonestaciones] hechos:", e)

    def _quitar_placeholder(self):
        """Borra el texto de ejemplo en cuanto el usuario hace clic/escribe."""
        try:
            if getattr(self, "_placeholder_activo", True):
                if self.txt_detalle.get("1.0", "end").strip() == self._placeholder_detalle:
                    self.txt_detalle.delete("1.0", "end")
                self.txt_detalle.configure(text_color=("#000000", "#ffffff"))
                self._placeholder_activo = False
        except Exception:
            pass

    def _restaurar_placeholder(self):
        """Vuelve a poner el texto de ejemplo si el cuadro quedó vacío."""
        try:
            if not self.txt_detalle.get("1.0", "end").strip():
                self.txt_detalle.insert("1.0", self._placeholder_detalle)
                self.txt_detalle.configure(text_color="#9aa5b1")
                self._placeholder_activo = True
        except Exception:
            pass

    def _detalle_valor(self):
        """Texto real de las observaciones (sin el ejemplo)."""
        try:
            valor = self.txt_detalle.get("1.0", "end").strip()
        except Exception:
            return ""
        if valor == getattr(self, "_placeholder_detalle", ""):
            return ""
        return valor

    def _datos_formulario(self):
        return {
            "tipo": self.combo_tipo.get(),
            "chofer": self.combo_chofer.get().strip(),
            "dni": self._dni_chofer(self.combo_chofer.get().strip()),
            "motivo": self.combo_motivo.get().strip(),
            "fecha": self.ent_fecha.get().strip(),
            "detalle": self._detalle_valor(),
            "dias": self.ent_dias.get().strip(),
            "fecha_inicio": self.ent_inicio.get().strip(),
            "gerente": self.ent_gerente.get().strip(),
        }

    def _actualizar_previa(self, *_a):
        datos = self._datos_formulario()
        if not hasattr(self, "txt_previa"):
            return
        es_susp = "suspens" in datos["tipo"].lower()
        cuerpo = (TEXTO_SUSPENSION if es_susp else TEXTO_AMONESTACION).format(
            motivo=datos["motivo"] or "(seleccione un motivo)",
            dias=datos["dias"] or "1",
            inicio=datos["fecha_inicio"] or datos["fecha"])
        texto = ("%s" % ("SUSPENSIÓN TEMPORAL DE LABORES" if es_susp else "AMONESTACIÓN ESCRITA"))
        texto += chr(10) + chr(10)
        texto += "Nombre: " + (datos["chofer"] or "—") + chr(10)
        texto += "DNI: " + (datos["dni"] or "—") + chr(10)
        texto += "Fecha: " + (datos["fecha"] or "—") + chr(10)
        texto += "Motivo: " + (datos["motivo"] or "—") + chr(10) + chr(10)
        texto += cuerpo
        if datos["detalle"]:
            texto += chr(10) + chr(10) + "Observaciones: " + datos["detalle"]
        texto += chr(10) + chr(10) + chr(10) + "_______________________________" + chr(10)
        texto += (datos["gerente"] or "(Gerente de Operaciones)") + chr(10) + "Gerente de Operaciones"
        self.txt_previa.configure(state="normal")
        self.txt_previa.delete("1.0", "end")
        self.txt_previa.insert("1.0", texto)
        self.txt_previa.configure(state="disabled")

    # ------------------------------------------------------------------
    def crear_motivo(self):
        ventana = ctk.CTkToplevel(self)
        ventana.title("Nuevo motivo")
        ventana.geometry("420x180")
        ventana.transient(self.winfo_toplevel())
        ventana.grab_set()
        ctk.CTkLabel(ventana, text="Escriba el nuevo motivo:",
                     font=("Arial", 12, "bold")).pack(anchor="w", padx=16, pady=(16, 6))
        entrada = ctk.CTkEntry(ventana, placeholder_text="Ej.: Uso indebido del celular")
        entrada.pack(fill="x", padx=16)
        entrada.focus_set()

        def guardar():
            ok, mensaje = guardar_motivo(entrada.get())
            if not ok:
                return messagebox.showwarning("Motivo", mensaje, parent=ventana)
            self.combo_motivo.configure(values=cargar_motivos())
            self.combo_motivo.set(entrada.get().strip())
            registrar_auditoria(self.usuario, "Amonestaciones",
                                "Creó el motivo: " + entrada.get().strip())
            ventana.destroy()

        ctk.CTkButton(ventana, text="Guardar", font=("Arial", 12, "bold"),
                      fg_color="#27ae60", hover_color="#1e8449",
                      command=guardar).pack(fill="x", padx=16, pady=14)

    # ------------------------------------------------------------------
    def _reservar_numero(self, tipo, chofer):
        """Reserva el próximo número de la serie (atómico). Devuelve (id, número)."""
        prefijo = _prefijo_documento(tipo)
        anio = datetime.datetime.now().year
        patron = "%s-%d-%%" % (prefijo, anio)
        conn = conectar_db(silencioso=True)
        if not conn:
            raise RuntimeError("sin conexión a la base")
        try:
            with conn.cursor() as c:
                c.execute("""
                    INSERT INTO amonestaciones (tipo, chofer, numero_documento)
                    SELECT %s, %s,
                           %s || '-' || %s || '-' ||
                           LPAD((COALESCE(MAX(CAST(SPLIT_PART(numero_documento, '-', 3) AS INTEGER)), 0) + 1)::text, 4, '0')
                    FROM amonestaciones
                    WHERE numero_documento LIKE %s
                    RETURNING id, numero_documento
                """, (tipo, chofer, prefijo, str(anio), patron))
                fila = c.fetchone()
            conn.commit()
            if not fila:
                raise RuntimeError("no se pudo reservar el número")
            return int(fila[0]), str(fila[1])
        except Exception:
            try:
                conn.rollback()
            except Exception:
                pass
            raise
        finally:
            liberar_conexion(conn)

    def _descartar_reserva(self, id_fila):
        """Borra la fila reservada si no se pudo generar el documento."""
        conn = conectar_db(silencioso=True)
        if not conn:
            return
        try:
            with conn.cursor() as c:
                c.execute("DELETE FROM amonestaciones WHERE id = %s AND COALESCE(archivo_ruta,'') = ''",
                          (id_fila,))
            conn.commit()
        except Exception as e:
            print("[amonestaciones] descartar reserva:", e)
        finally:
            liberar_conexion(conn)

    def _numero_para_guardar(self, tipo):
        """Número consecutivo del documento (se reintenta si otro equipo lo tomó)."""
        for _intento in range(6):
            numero = siguiente_numero_documento(tipo)
            conn = conectar_db(silencioso=True)
            if not conn:
                return numero
            try:
                with conn.cursor() as c:
                    c.execute("SELECT 1 FROM amonestaciones WHERE numero_documento = %s", (numero,))
                    if c.fetchone():
                        continue          # ya existe: se prueba el siguiente
                    return numero
            except Exception as e:
                print("[amonestaciones] verificar número:", e)
                return numero
            finally:
                liberar_conexion(conn)
        return siguiente_numero_documento(tipo)

    def generar(self):
        datos = self._datos_formulario()
        if not datos["chofer"]:
            return messagebox.showwarning("Amonestaciones", "Seleccione el chofer.")
        if not datos["motivo"]:
            return messagebox.showwarning("Amonestaciones", "Seleccione el motivo.")
        if not datos["gerente"]:
            return messagebox.showwarning("Amonestaciones",
                                          "Escriba el nombre del Gerente de Operaciones (va en la firma).")
        # 🔁 el nombre del gerente se recuerda para los próximos documentos
        self.config["gerente"] = datos["gerente"]
        guardar_config(self.config)

        # ☁️ Se guarda en la carpeta SINCRONIZADA (Google Drive) para que viaje a la nube
        base_nube = ""
        try:
            base_nube = ruta_drive_configurada()
        except Exception:
            base_nube = ""
        if base_nube and Path(str(base_nube)).is_dir():
            carpeta = Path(str(base_nube)) / "amonestaciones"
        else:
            carpeta = Path(str(DATA_DIR)) / "amonestaciones"
        carpeta.mkdir(parents=True, exist_ok=True)
        # 🔢 Se RESERVA el número antes de generar el PDF (atómico: nunca se repite)
        id_fila, datos["numero"] = None, ""
        try:
            id_fila, datos["numero"] = self._reservar_numero(datos["tipo"], datos["chofer"])
        except Exception as e_num:
            print("[amonestaciones] número:", e_num)

        marca = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        tipo_txt = "Suspension" if "suspens" in datos["tipo"].lower() else "Amonestacion"
        nombre_archivo = "%s_%s_%s.pdf" % (
            (datos.get("numero") or tipo_txt).replace("/", "-"),
            _nombre_archivo_seguro(datos["chofer"]), marca)
        ruta_pdf = str(carpeta / nombre_archivo)

        try:
            generar_pdf(datos, ruta_pdf, datos["gerente"])
        except Exception as e:
            if id_fila:
                self._descartar_reserva(id_fila)
            return messagebox.showerror("Amonestaciones", "No se pudo generar el PDF: " + str(e))
        try:
            ruta_bd = ruta_para_guardar(ruta_pdf, base_nube) if base_nube else ruta_pdf
        except Exception:
            ruta_bd = ruta_pdf

        conn = conectar_db(silencioso=True)
        if conn:
            try:
                with conn.cursor() as c:
                    if id_fila:
                        # se completa la fila ya reservada (con su número)
                        c.execute("""UPDATE amonestaciones SET
                                        id_chofer=%s, dni=%s, motivo=%s, detalle=%s, fecha=%s,
                                        fecha_inicio=%s, dias=%s, gerente=%s, archivo_ruta=%s
                                     WHERE id=%s""",
                                  (0, datos["dni"], datos["motivo"], datos["detalle"],
                                   datos["fecha"], datos["fecha_inicio"], int(datos["dias"] or 0),
                                   datos["gerente"], ruta_pdf, id_fila))
                    else:
                        c.execute("""
                            INSERT INTO amonestaciones
                                (tipo, id_chofer, chofer, dni, motivo, detalle, fecha, fecha_inicio,
                                 dias, gerente, archivo_ruta, numero_documento)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """, (datos["tipo"], 0, datos["chofer"], datos["dni"], datos["motivo"],
                              datos["detalle"], datos["fecha"], datos["fecha_inicio"],
                              int(datos["dias"] or 0), datos["gerente"], ruta_pdf,
                              datos.get("numero") or None))
                conn.commit()
            except Exception as e:
                print("[amonestaciones] guardar:", e)
            finally:
                liberar_conexion(conn)

        registrar_auditoria(self.usuario, "Amonestaciones",
                            "%s a %s por: %s" % (datos["tipo"], datos["chofer"], datos["motivo"]))
        try:
            abrir_documento(ruta_pdf)
        except Exception:
            pass
        messagebox.showinfo("Amonestaciones",
                            "Documento generado y guardado:" + chr(10) + ruta_pdf)

    def limpiar(self):
        self.txt_detalle.delete("1.0", "end")
        self.txt_detalle.insert("1.0", self._placeholder_detalle)
        self.txt_detalle.configure(text_color="#9aa5b1")
        self._placeholder_activo = True
        self.ent_dias.delete(0, "end")
        self.ent_dias.insert(0, "1")
        self.ent_fecha.delete(0, "end")
        self.ent_fecha.insert(0, datetime.date.today().strftime("%d/%m/%Y"))
        self.ent_inicio.delete(0, "end")
        self.ent_inicio.insert(0, datetime.date.today().strftime("%d/%m/%Y"))
        self._actualizar_previa()

    # ------------------------------------------------------------------
    def abrir_historial(self):
        HistorialAmonestaciones(self)


class HistorialAmonestaciones(ctk.CTkToplevel):
    """Historial de documentos con búsqueda."""

    def __init__(self, padre):
        super().__init__(padre)
        self.padre = padre
        self.title("Historial de Amonestaciones y Suspensiones")
        self.geometry("1000x560")
        self.transient(padre.winfo_toplevel())

        barra = ctk.CTkFrame(self, fg_color="transparent")
        barra.pack(fill="x", padx=12, pady=10)
        ctk.CTkLabel(barra, text="🔍 Buscar:", font=("Arial", 12, "bold")).pack(side="left")
        self.ent_buscar = ctk.CTkEntry(barra, placeholder_text="Chofer, DNI, motivo o tipo...")
        self.ent_buscar.pack(side="left", fill="x", expand=True, padx=8)
        self.ent_buscar.bind("<KeyRelease>", lambda e: self.cargar())
        ctk.CTkButton(barra, text="🔄 Actualizar", width=110,
                      command=self.cargar).pack(side="left")

        columnas = ("numero", "id", "tipo", "fecha", "chofer", "dni", "motivo", "gerente", "archivo")
        self.tabla = ttk.Treeview(self, columns=columnas, show="headings", height=18)
        titulos = {"numero": "N° Documento", "id": "ID", "tipo": "Tipo", "fecha": "Fecha",
                   "chofer": "Chofer", "dni": "DNI", "motivo": "Motivo",
                   "gerente": "Gerente", "archivo": "PDF"}
        anchos = {"numero": 130, "id": 45, "tipo": 120, "fecha": 90, "chofer": 190, "dni": 85,
                  "motivo": 240, "gerente": 160, "archivo": 55}
        for col in columnas:
            self.tabla.heading(col, text=titulos[col])
            self.tabla.column(col, width=anchos[col], anchor="w" if col in ("chofer", "motivo") else "center")
        self.tabla.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        self.tabla.bind("<Double-1>", lambda e: self.ver_pdf())

        botones = ctk.CTkFrame(self, fg_color="transparent")
        botones.pack(fill="x", pady=(0, 12))
        ctk.CTkButton(botones, text="📄 Ver PDF", font=("Arial", 12, "bold"),
                      fg_color="#1f538d", hover_color="#163b65",
                      command=self.ver_pdf).pack(side="left", padx=12)
        ctk.CTkButton(botones, text="✏️ Editar", font=("Arial", 12, "bold"),
                      fg_color="#e67e22", hover_color="#ca6f1e",
                      command=self.editar).pack(side="left", padx=6)
        ctk.CTkButton(botones, text="🗑️ Eliminar", font=("Arial", 12, "bold"),
                      fg_color="#c0392b", hover_color="#96281b",
                      command=self.eliminar).pack(side="left", padx=6)
        self.cargar()

    def cargar(self):
        for i in self.tabla.get_children():
            self.tabla.delete(i)
        for fila in listar_amonestaciones(self.ent_buscar.get().strip()):
            (id_, tipo, fecha, chofer, dni, motivo, gerente, archivo, detalle, dias, inicio,
             numero_doc) = fila
            self.tabla.insert("", tk.END, iid=str(id_), values=(
                numero_doc or "—", id_, tipo, fecha, chofer, dni, motivo, gerente,
                "Sí" if archivo else "—"))


    # ------------------------------------------------------------------
    def _fila_seleccionada(self):
        sel = self.tabla.selection()
        if not sel:
            messagebox.showinfo("Historial", "Seleccione un documento.", parent=self)
            return None
        return int(sel[0])

    def _leer_documento(self, id_doc):
        conn = conectar_db(silencioso=True)
        datos = None
        if conn:
            try:
                with conn.cursor() as c:
                    c.execute("""SELECT tipo, chofer, COALESCE(dni,''), motivo, COALESCE(detalle,''),
                                        COALESCE(fecha,''), COALESCE(fecha_inicio,''), COALESCE(dias,0),
                                        COALESCE(gerente,''), COALESCE(archivo_ruta,''),
                                        COALESCE(numero_documento,'')
                                 FROM amonestaciones WHERE id = %s""", (id_doc,))
                    fila = c.fetchone()
                    if fila:
                        datos = {"id": id_doc, "tipo": fila[0], "chofer": fila[1], "dni": fila[2],
                                 "motivo": fila[3], "detalle": fila[4], "fecha": fila[5],
                                 "fecha_inicio": fila[6], "dias": fila[7], "gerente": fila[8],
                                 "archivo": fila[9], "numero": fila[10]}
            except Exception as e:
                print("[amonestaciones] leer:", e)
            finally:
                liberar_conexion(conn)
        return datos

    def editar(self):
        id_doc = self._fila_seleccionada()
        if not id_doc:
            return
        datos = self._leer_documento(id_doc)
        if not datos:
            return messagebox.showwarning("Historial", "No se pudo leer el documento.", parent=self)

        ventana = ctk.CTkToplevel(self)
        ventana.title("Editar documento %s" % (datos.get("numero") or ("#%s" % id_doc)))
        ventana.geometry("520x620")
        ventana.transient(self)
        ventana.grab_set()
        marco = ctk.CTkScrollableFrame(ventana, width=480, height=460)
        marco.pack(fill="both", expand=True, padx=10, pady=10)

        entradas = {}
        def campo(etiqueta, valor, clave):
            ctk.CTkLabel(marco, text=etiqueta, font=("Arial", 11, "bold")).pack(anchor="w", pady=(6, 0))
            entrada = ctk.CTkEntry(marco)
            entrada.pack(fill="x")
            entrada.insert(0, "" if valor is None else str(valor))
            entradas[clave] = entrada

        campo("Colaborador:", datos["chofer"], "chofer")
        campo("DNI:", datos["dni"], "dni")
        campo("Fecha del documento:", datos["fecha"], "fecha")

        ctk.CTkLabel(marco, text="Motivo:", font=("Arial", 11, "bold")).pack(anchor="w", pady=(6, 0))
        combo_motivo = ctk.CTkComboBox(marco, values=cargar_motivos())
        combo_motivo.pack(fill="x")
        combo_motivo.set(datos["motivo"] or "")

        campo("Días de suspensión:", datos["dias"], "dias")
        campo("Inicio de la suspensión:", datos["fecha_inicio"], "fecha_inicio")
        campo("Gerente de Operaciones:", datos["gerente"], "gerente")

        ctk.CTkLabel(marco, text="Observaciones:", font=("Arial", 11, "bold")).pack(anchor="w", pady=(6, 0))
        txt = ctk.CTkTextbox(marco, height=80, border_width=1, border_color="#9aa5b1")
        txt.pack(fill="x")
        txt.insert("1.0", datos["detalle"] or "")
        marcar_pdf = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(marco, text="Volver a generar el PDF con estos datos",
                       variable=marcar_pdf).pack(anchor="w", pady=(10, 4))

        def guardar():
            nuevos = {c: e.get().strip() for c, e in entradas.items()}
            nuevos["motivo"] = combo_motivo.get().strip()
            nuevos["detalle"] = txt.get("1.0", "end").strip()
            conn = conectar_db(silencioso=True)
            if not conn:
                return
            try:
                with conn.cursor() as c:
                    c.execute("""UPDATE amonestaciones SET chofer=%s, dni=%s, fecha=%s, motivo=%s,
                                        detalle=%s, dias=%s, fecha_inicio=%s, gerente=%s
                                 WHERE id=%s""",
                              (nuevos["chofer"], nuevos["dni"], nuevos["fecha"], nuevos["motivo"],
                               nuevos["detalle"], int(nuevos["dias"] or 0), nuevos["fecha_inicio"],
                               nuevos["gerente"], id_doc))
                conn.commit()
            except Exception as e:
                return messagebox.showerror("Historial", "No se pudo guardar: " + str(e), parent=ventana)
            finally:
                liberar_conexion(conn)

            if marcar_pdf.get() and datos.get("archivo"):
                try:
                    destino = datos["archivo"]
                    try:
                        destino = resolver_ruta_archivo(destino) or destino
                    except Exception:
                        pass
                    pdf_datos = dict(nuevos)
                    pdf_datos["tipo"] = datos["tipo"]
                    # 🔢 El número no se edita: se conserva el original
                    pdf_datos["numero"] = datos.get("numero", "")
                    generar_pdf(pdf_datos, destino, nuevos["gerente"])
                except Exception as e:
                    messagebox.showwarning("Historial", "Se guardó, pero no se pudo regenerar el PDF: "
                                           + str(e), parent=ventana)
            registrar_auditoria(getattr(self.padre, "usuario", "Sistema"), "Amonestaciones",
                                "Editó el documento #%s" % id_doc)
            ventana.destroy()
            self.cargar()
            messagebox.showinfo("Historial", "Documento actualizado.", parent=self)

        ctk.CTkButton(marco, text="💾 Guardar cambios", font=("Arial", 12, "bold"),
                      fg_color="#27ae60", hover_color="#1e8449",
                      command=guardar).pack(fill="x", pady=(12, 8))

    def eliminar(self):
        id_doc = self._fila_seleccionada()
        if not id_doc:
            return
        datos = self._leer_documento(id_doc)
        if not messagebox.askyesno("Eliminar", "¿Eliminar el documento #%s de %s?"
                                   % (id_doc, (datos or {}).get("chofer", "")), parent=self):
            return
        borrar_pdf = messagebox.askyesno("Eliminar",
                                         "¿Borrar también el PDF del disco?", parent=self)
        conn = conectar_db(silencioso=True)
        if not conn:
            return
        try:
            with conn.cursor() as c:
                c.execute("DELETE FROM amonestaciones WHERE id = %s", (id_doc,))
            conn.commit()
        except Exception as e:
            return messagebox.showerror("Historial", "No se pudo eliminar: " + str(e), parent=self)
        finally:
            liberar_conexion(conn)
        if borrar_pdf and datos and datos.get("archivo"):
            try:
                Path(str(datos["archivo"])).unlink(missing_ok=True)
            except Exception as e:
                print("[amonestaciones] borrar pdf:", e)
        registrar_auditoria(getattr(self.padre, "usuario", "Sistema"), "Amonestaciones",
                            "Eliminó el documento #%s" % id_doc)
        self.cargar()

    def ver_pdf(self):
        sel = self.tabla.selection()
        if not sel:
            return messagebox.showinfo("Historial", "Seleccione un documento.", parent=self)
        conn = conectar_db(silencioso=True)
        ruta = ""
        if conn:
            try:
                with conn.cursor() as c:
                    c.execute("SELECT COALESCE(archivo_ruta,'') FROM amonestaciones WHERE id = %s",
                              (int(sel[0]),))
                    fila = c.fetchone()
                    ruta = str(fila[0]) if fila and fila[0] else ""
            except Exception:
                pass
            finally:
                liberar_conexion(conn)
        try:
            ruta = resolver_ruta_archivo(ruta) or ruta
        except Exception:
            pass
        if not ruta or not Path(ruta).is_file():
            return messagebox.showwarning("Historial", "El PDF no está disponible.", parent=self)
        try:
            abrir_documento(ruta)
        except Exception as e:
            messagebox.showerror("Historial", "No se pudo abrir el PDF: " + str(e), parent=self)
