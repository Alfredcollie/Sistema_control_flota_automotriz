# -*- coding: utf-8 -*-
"""
CONTRATOS_ACTAS.PY — Módulo "Contratos y Actas" (Flota)
================================================================
Pestañas: Contratos · Renovaciones · Actas de Entrega · Históricos

Regla de oro: al renovar SOLO se actualiza la fecha de vencimiento del chofer
(choferes.fecha_fin_contrato). La fecha de inicio original
(choferes.fecha_inicio_contrato) NUNCA se toca.
"""
import json
import os
import re
import sys
import threading
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import messagebox, ttk

import customtkinter as ctk

from conexion import conectar_db, liberar_conexion, registrar_auditoria
from app_paths import CONFIG_FILE, DATA_DIR

import plantillas_contratos as PL
import contratos_actas_pdf as pdf

FAMILIA = "Helvetica" if sys.platform == "darwin" else "Arial"
COLOR_AZUL = "#1f538d"
COLOR_VERDE = "#16a085"
COLOR_GRIS = "#34495e"

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
         "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

_SCHEMA_OK = False


# =========================================================
# CONFIGURACIÓN
# =========================================================
def cargar_configuracion():
    """Lee la configuración general (datos de la empresa, ruta, etc.)."""
    config = {}
    try:
        if os.path.exists(str(CONFIG_FILE)):
            with open(str(CONFIG_FILE), "r", encoding="utf-8") as archivo:
                config = json.load(archivo) or {}
    except Exception as error:
        print("Aviso - configuración:", error)
    return config


def datos_empresa(config=None):
    """Datos del empleador para los documentos."""
    config = config if isinstance(config, dict) else cargar_configuracion()
    return {
        "empresa": str(config.get("razon_social_empresa") or "").strip(),
        "ruc_empresa": re.sub(r"\D", "", str(config.get("ruc_empresa") or "")),
        "domicilio_empresa": str(config.get("domicilio_fiscal_empresa") or "").strip(),
    }


def falta_configuracion(config=None):
    """Devuelve la lista de datos de la empresa que faltan."""
    datos = datos_empresa(config)
    faltan = []
    if not datos["empresa"]:
        faltan.append("Razón social de la empresa")
    if not datos["ruc_empresa"]:
        faltan.append("RUC de la empresa")
    if not datos["domicilio_empresa"]:
        faltan.append("Domicilio fiscal de la empresa")
    return faltan


RUTA_BASE_FIJA = None      # ── ruta de la carpeta raíz de documentos


def ruta_base():
    """Carpeta "Contratos y Actas" dentro de la carpeta sincronizada con Drive."""
    global RUTA_BASE_FIJA
    if RUTA_BASE_FIJA:
        return RUTA_BASE_FIJA
    base = ""
    try:
        from modulo_compras import obtener_ruta_base_drive
        base = obtener_ruta_base_drive() or ""
    except Exception:
        base = ""
    if not base:
        try:
            base = str(DATA_DIR)
        except Exception:
            base = os.path.join(os.path.expanduser("~"), "Flota")
    carpeta = os.path.join(base, "Contratos y Actas")
    try:
        os.makedirs(carpeta, exist_ok=True)
    except Exception as error:
        print("Aviso - carpeta Contratos y Actas:", error)
    RUTA_BASE_FIJA = carpeta
    return carpeta



def movil_valido(texto):
    """Devuelve la placa solo si es un dato real (descarta 'Ninguno / Sin Asignar')."""
    valor = str(texto or "").strip()
    if not valor:
        return ""
    if valor.upper() in ("NINGUNO", "SIN ASIGNAR", "NINGUNO / SIN ASIGNAR", "-", "N/A",
                         "NINGUNO/SIN ASIGNAR", "SIN PLACA"):
        return ""
    if "SIN ASIGNAR" in valor.upper():
        return ""
    return valor


def carpeta_chofer(dni, nombre):
    """Subcarpeta por chofer (DNI - nombre)."""
    limpio = re.sub(r"[^\w\sáéíóúÁÉÍÓÚñÑ\-]", "", str(nombre or "")).strip()
    limpio = re.sub(r"\s+", " ", limpio)[:60]
    nombre_carpeta = (str(dni or "") + " - " + limpio).strip(" -")
    carpeta = os.path.join(ruta_base(), nombre_carpeta or "Sin datos")
    try:
        os.makedirs(carpeta, exist_ok=True)
    except Exception as error:
        print("Aviso - carpeta del chofer:", error)
    return carpeta


# =========================================================
# FECHAS
# =========================================================
def hoy_texto():
    return datetime.now().strftime("%d/%m/%Y")


def sumar_meses(fecha_texto, meses=1):
    """Suma meses a una fecha dd/mm/aaaa (ajustando el día si hace falta)."""
    try:
        fecha = datetime.strptime(str(fecha_texto).strip(), "%d/%m/%Y")
    except Exception:
        fecha = datetime.now()
    anio = fecha.year + (fecha.month - 1 + meses) // 12
    mes = (fecha.month - 1 + meses) % 12 + 1
    dia = fecha.day
    while True:
        try:
            return datetime(anio, mes, dia).strftime("%d/%m/%Y")
        except ValueError:
            dia -= 1
            if dia < 1:
                return datetime(anio, mes, 1).strftime("%d/%m/%Y")


def estado_contrato(fecha_fin, dias_aviso=30):
    """Devuelve (estado, días restantes)."""
    texto = str(fecha_fin or "").strip()
    if not texto:
        return "Sin contrato", None
    try:
        fin = datetime.strptime(texto, "%d/%m/%Y")
    except Exception:
        try:
            fin = datetime.strptime(texto, "%Y-%m-%d")
        except Exception:
            return "Sin contrato", None
    dias = (fin.date() - datetime.now().date()).days
    if dias < 0:
        return "Vencido", dias
    if dias <= int(dias_aviso or 30):
        return "Próximo a vencer", dias
    return "Vigente", dias


def numero_a_texto(monto):
    """1,230.00 como texto con dos decimales."""
    try:
        return "%.2f" % float(str(monto).replace(",", ""))
    except Exception:
        return "0.00"


# =========================================================
# BASE DE DATOS
# =========================================================
def crear_tablas():
    """Crea las tablas del módulo si no existen."""
    global _SCHEMA_OK
    if _SCHEMA_OK:
        return True
    conexion = conectar_db(silencioso=True)
    if not conexion:
        return False
    try:
        cursor = conexion.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS contratos (
                id SERIAL PRIMARY KEY,
                id_chofer INTEGER,
                dni VARCHAR(20),
                trabajador VARCHAR(200),
                cargo VARCHAR(160),
                remuneracion NUMERIC,
                fecha_inicio VARCHAR(20),
                fecha_fin VARCHAR(20),
                fecha_inicio_original VARCHAR(20),
                ciudad_firma VARCHAR(80),
                fecha_firma VARCHAR(20),
                tipo VARCHAR(30) DEFAULT 'Inicial',
                id_contrato_origen INTEGER,
                ruta_pdf TEXT,
                usuario VARCHAR(80),
                creado VARCHAR(30)
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS renovaciones (
                id SERIAL PRIMARY KEY,
                id_contrato INTEGER,
                id_chofer INTEGER,
                dni VARCHAR(20),
                trabajador VARCHAR(200),
                fecha_inicio_original VARCHAR(20),
                vencimiento_anterior VARCHAR(20),
                nuevo_vencimiento VARCHAR(20),
                fecha_renovacion VARCHAR(30),
                usuario VARCHAR(80),
                ruta_pdf TEXT,
                creado VARCHAR(30)
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS actas_entrega (
                id SERIAL PRIMARY KEY,
                id_chofer INTEGER,
                dni VARCHAR(20),
                trabajador VARCHAR(200),
                cargo VARCHAR(160),
                fecha_acta VARCHAR(20),
                ciudad_firma VARCHAR(80),
                bienes TEXT,
                celular_marca_modelo VARCHAR(120),
                celular_imei VARCHAR(80),
                vehiculo_marca VARCHAR(60),
                vehiculo_modelo VARCHAR(60),
                vehiculo_anio VARCHAR(10),
                vehiculo_placa VARCHAR(20),
                ruta_pdf TEXT,
                usuario VARCHAR(80),
                creado VARCHAR(30)
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS contratos_documentos (
                id SERIAL PRIMARY KEY,
                tipo VARCHAR(40),
                id_origen INTEGER,
                id_chofer INTEGER,
                dni VARCHAR(20),
                trabajador VARCHAR(200),
                descripcion VARCHAR(200),
                archivo TEXT,
                usuario VARCHAR(80),
                creado VARCHAR(30)
            )
        """)
        conexion.commit()
        _SCHEMA_OK = True
    except Exception as error:
        print("Aviso - tablas de Contratos y Actas:", error)
        try:
            conexion.rollback()
        except Exception:
            pass
        return False
    finally:
        liberar_conexion(conexion)
    return True


def inicializar_tablas_en_hilo():
    """Crea las tablas en segundo plano (no frena la interfaz)."""
    def tarea():
        crear_tablas()
    threading.Thread(target=tarea, daemon=True).start()


def buscar_choferes(texto="", incluir_inactivos=False):
    """Choferes que coincidan con el texto (nombre, DNI o móvil).

    Por defecto NO se incluyen los choferes con estado "Inactivo": en este
    módulo solo tienen sentido los choferes que están trabajando.
    """
    conexion = conectar_db(silencioso=True)
    if not conexion:
        return []
    try:
        cursor = conexion.cursor()
        patron = "%" + str(texto or "").strip() + "%"
        filtro_estado = "" if incluir_inactivos else \
            "AND COALESCE(estado, 'Activo') NOT IN ('Inactivo', 'INACTIVO', 'inactivo')"
        cursor.execute("""
            SELECT id, dni, nombres, direccion, movil_asignado,
                   fecha_inicio_contrato, fecha_fin_contrato, telefono, estado
            FROM choferes
            WHERE (COALESCE(nombres, '') ILIKE %s
               OR COALESCE(dni, '') ILIKE %s
               OR COALESCE(movil_asignado, '') ILIKE %s)
              """ + filtro_estado + """
            ORDER BY nombres
            LIMIT 200
        """, (patron, patron, patron))
        return cursor.fetchall()
    except Exception as error:
        print("Aviso - buscar choferes:", error)
        return []
    finally:
        liberar_conexion(conexion)


def datos_vehiculo(placa):
    """Marca, modelo y año del vehículo (tabla flota_vehiculos)."""
    if not str(placa or "").strip():
        return {}
    conexion = conectar_db(silencioso=True)
    if not conexion:
        return {}
    try:
        cursor = conexion.cursor()
        cursor.execute("""
            SELECT placa, marca, modelo, anio, color FROM flota_vehiculos
            WHERE TRIM(UPPER(placa)) = TRIM(UPPER(%s)) LIMIT 1
        """, (str(placa).strip(),))
        fila = cursor.fetchone()
        if not fila:
            return {}
        return {"placa": fila[0] or "", "marca": fila[1] or "", "modelo": fila[2] or "",
                "anio": fila[3] or "", "color": fila[4] or ""}
    except Exception as error:
        print("Aviso - datos del vehículo:", error)
        return {}
    finally:
        liberar_conexion(conexion)


def contratos_del_chofer(id_chofer):
    """Historial de contratos y renovaciones de un chofer."""
    conexion = conectar_db(silencioso=True)
    if not conexion:
        return []
    try:
        cursor = conexion.cursor()
        cursor.execute("""
            SELECT id, fecha_inicio, fecha_fin, fecha_inicio_original, tipo,
                   remuneracion, cargo, ruta_pdf, creado
            FROM contratos WHERE id_chofer = %s ORDER BY id DESC
        """, (id_chofer,))
        return cursor.fetchall()
    except Exception as error:
        print("Aviso - contratos del chofer:", error)
        return []
    finally:
        liberar_conexion(conexion)


def renovaciones_del_chofer(id_chofer):
    conexion = conectar_db(silencioso=True)
    if not conexion:
        return []
    try:
        cursor = conexion.cursor()
        cursor.execute("""
            SELECT id, vencimiento_anterior, nuevo_vencimiento, fecha_renovacion,
                   usuario, ruta_pdf
            FROM renovaciones WHERE id_chofer = %s ORDER BY id DESC
        """, (id_chofer,))
        return cursor.fetchall()
    except Exception as error:
        print("Aviso - renovaciones:", error)
        return []
    finally:
        liberar_conexion(conexion)



# =========================================================
# MÓDULO
# =========================================================
class ContratosActasApp:
    """Módulo "Contratos y Actas" (se abre dentro del programa de Flota)."""

    def __init__(self, parent_frame, usuario_activo="Desconocido"):
        self.parent_frame = parent_frame
        self.usuario_activo = usuario_activo
        self.usuario = usuario_activo or "Sistema"
        self.config = cargar_configuracion()
        self.empresa = datos_empresa(self.config)
        self.chofer = None
        self.contrato_actual = None
        self.vehiculo = {}

        inicializar_tablas_en_hilo()
        self.crear_interfaz()

    # -----------------------------------------------------
    # INTERFAZ
    # -----------------------------------------------------
    def crear_interfaz(self):
        self.frame_main = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.frame_main.pack(fill="both", expand=True, padx=12, pady=10)

        cabecera = ctk.CTkFrame(self.frame_main, fg_color="transparent")
        cabecera.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(cabecera, text="📄 CONTRATOS Y ACTAS", font=(FAMILIA, 18, "bold"),
                     text_color=COLOR_AZUL).pack(side="left")
        ctk.CTkLabel(cabecera, text="Contratos laborales, renovaciones y actas de entrega",
                     font=(FAMILIA, 11), text_color="#7f8c8d").pack(side="left", padx=12)
        ctk.CTkLabel(cabecera, text="📁 " + ruta_base(), font=(FAMILIA, 9),
                     text_color="#95a5a6").pack(side="right")
        self.incluir_inactivos = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(cabecera, text="Incluir inactivos", variable=self.incluir_inactivos,
                        font=(FAMILIA, 10), checkbox_width=16, checkbox_height=16,
                        command=self._recargar_listas).pack(side="right", padx=(0, 14))

        # aviso si falta configurar la empresa
        faltan = falta_configuracion(self.config)
        if faltan:
            aviso = ctk.CTkFrame(self.frame_main, fg_color="#fff3cd", corner_radius=6)
            aviso.pack(fill="x", pady=(0, 8))
            ctk.CTkLabel(aviso,
                         text="⚠️ Falta configurar: " + ", ".join(faltan) +
                              "  →  Ajustes ▸ Configuración General",
                         font=(FAMILIA, 11, "bold"), text_color="#856404").pack(padx=10, pady=6)

        self.pestanas = ctk.CTkTabview(self.frame_main, segmented_button_selected_color=COLOR_AZUL,
                                       segmented_button_selected_hover_color="#163b65")
        self.pestanas.pack(fill="both", expand=True)
        for nombre in ("Contratos", "Renovaciones", "Actas de Entrega", "Históricos"):
            self.pestanas.add(nombre)

        self._crear_tab_contratos(self.pestanas.tab("Contratos"))
        self._crear_tab_renovaciones(self.pestanas.tab("Renovaciones"))
        self._crear_tab_actas(self.pestanas.tab("Actas de Entrega"))
        self._crear_tab_historicos(self.pestanas.tab("Históricos"))

        # carga inicial de las pestañas que dependen de la primera
        try:
            self._cargar_choferes_en_renovaciones()
            self._cargar_choferes_en_actas()
        except Exception as error:
            print("Aviso - carga inicial:", error)

    # -----------------------------------------------------
    # UTILIDADES DE PANTALLA
    # -----------------------------------------------------


    def _ver_inactivos(self):
        """True si el usuario marcó 'Incluir inactivos'."""
        try:
            return bool(self.incluir_inactivos.get())
        except Exception:
            return False

    def _recargar_listas(self):
        """Vuelve a cargar las listas respetando la casilla de inactivos."""
        try:
            self._buscar_en_contratos()
        except Exception as error:
            print("Aviso - recargar contratos:", error)
        try:
            self._cargar_choferes_en_renovaciones()
        except Exception as error:
            print("Aviso - recargar renovaciones:", error)
        try:
            self._cargar_choferes_en_actas()
        except Exception as error:
            print("Aviso - recargar actas:", error)

    def _tabla(self, padre, columnas, anchos, alto=10):
        marco = ctk.CTkFrame(padre, fg_color="transparent")
        marco.pack(fill="both", expand=True)
        estilo = ttk.Style()
        estilo.configure("CT.Treeview", rowheight=26, font=(FAMILIA, 10))
        estilo.configure("CT.Treeview.Heading", font=(FAMILIA, 10, "bold"))
        tabla = ttk.Treeview(marco, columns=columnas, show="headings", height=alto,
                             style="CT.Treeview")
        for nombre, ancho in zip(columnas, anchos):
            tabla.heading(nombre, text=nombre)
            tabla.column(nombre, width=ancho, anchor="w")
        barra = ctk.CTkScrollbar(marco, orientation="vertical", command=tabla.yview)
        tabla.configure(yscrollcommand=barra.set)
        tabla.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        return tabla

    def _campo(self, padre, etiqueta, fila, columna=0, ancho=260, valor=""):
        ctk.CTkLabel(padre, text=etiqueta, font=(FAMILIA, 11, "bold"),
                     text_color="#2c3e50").grid(row=fila, column=columna, sticky="w",
                                                padx=(0, 6), pady=3)
        entrada = ctk.CTkEntry(padre, width=ancho, font=(FAMILIA, 11))
        entrada.grid(row=fila, column=columna + 1, sticky="w", pady=3)
        if valor:
            entrada.insert(0, str(valor))
        return entrada

    def _poner(self, entrada, valor):
        try:
            entrada.delete(0, tk.END)
            entrada.insert(0, str(valor or ""))
        except Exception as error:
            print("Aviso - campo:", error)

    def _elegir_chofer(self, fila_cruda):
        """Guarda los datos del chofer elegido y los muestra en pantalla."""
        if not fila_cruda:
            return
        (ident, dni, nombres, direccion, movil,
         inicio, fin, telefono, estado) = (list(fila_cruda) + [None] * 9)[:9]
        self.chofer = {
            "id": ident, "dni": str(dni or ""), "nombres": str(nombres or ""),
            "direccion": str(direccion or ""), "movil": str(movil or ""),
            "fecha_inicio_contrato": str(inicio or ""), "fecha_fin_contrato": str(fin or ""),
            "telefono": str(telefono or ""), "estado": str(estado or ""),
        }
        estado_txt, dias = estado_contrato(self.chofer["fecha_fin_contrato"])
        detalle = estado_txt
        if dias is not None:
            detalle += "  (" + ("venció hace %d día(s)" % abs(dias) if dias < 0
                                else "quedan %d día(s)" % dias) + ")"
        for etiqueta in (self.lbl_chofer_contrato, self.lbl_chofer_renovacion,
                         self.lbl_chofer_acta):
            try:
                etiqueta.configure(text="👤 " + self.chofer["nombres"] +
                                        "   ·   DNI " + self.chofer["dni"] +
                                        "   ·   " + detalle)
            except Exception:
                pass

    # =====================================================
    # PESTAÑA 1 · CONTRATOS
    # =====================================================
    def _crear_tab_contratos(self, padre):
        cuerpo = ctk.CTkFrame(padre, fg_color="transparent")
        cuerpo.pack(fill="both", expand=True)

        # ---- búsqueda del chofer
        busqueda = ctk.CTkFrame(cuerpo, fg_color="#f4f6f7", corner_radius=6)
        busqueda.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(busqueda, text="🔍 Buscar chofer:", font=(FAMILIA, 11, "bold")).pack(side="left", padx=8)
        self.ent_buscar_contrato = ctk.CTkEntry(busqueda, width=280,
                                               placeholder_text="Nombre, DNI o móvil…")
        self.ent_buscar_contrato.pack(side="left", padx=4, pady=6)
        self.ent_buscar_contrato.bind("<Return>", lambda _e: self._buscar_en_contratos())
        ctk.CTkButton(busqueda, text="Buscar", width=90, fg_color=COLOR_AZUL,
                      command=self._buscar_en_contratos).pack(side="left", padx=4)
        self.lbl_chofer_contrato = ctk.CTkLabel(busqueda, text="(ningún chofer elegido)",
                                                font=(FAMILIA, 11, "bold"), text_color="#c0392b")
        self.lbl_chofer_contrato.pack(side="left", padx=12)

        zona = ctk.CTkFrame(cuerpo, fg_color="transparent")
        zona.pack(fill="both", expand=True)

        # ---- lista de choferes
        izquierda = ctk.CTkFrame(zona, fg_color="#ffffff", corner_radius=6)
        izquierda.pack(side="left", fill="both", expand=True, padx=(0, 6))
        ctk.CTkLabel(izquierda, text="Choferes", font=(FAMILIA, 12, "bold"),
                     text_color=COLOR_AZUL).pack(anchor="w", padx=8, pady=(6, 2))
        self.tabla_choferes = self._tabla(izquierda,
                                          ("DNI", "Nombres", "Móvil", "Estado",
                                           "Contrato", "Vence contrato"),
                                          (78, 190, 68, 66, 90, 96), alto=12)
        self.tabla_choferes.bind("<<TreeviewSelect>>", self._on_chofer_contrato)

        # ---- datos del contrato
        derecha = ctk.CTkFrame(zona, fg_color="#ffffff", corner_radius=6)
        derecha.pack(side="left", fill="both", expand=True)
        ctk.CTkLabel(derecha, text="Datos del contrato", font=(FAMILIA, 12, "bold"),
                     text_color=COLOR_AZUL).pack(anchor="w", padx=10, pady=(6, 2))

        formulario = ctk.CTkFrame(derecha, fg_color="transparent")
        formulario.pack(fill="x", padx=10)
        self.ent_trabajador = self._campo(formulario, "Trabajador:", 0, valor="")
        self.ent_dni = self._campo(formulario, "DNI:", 1)
        self.ent_direccion = self._campo(formulario, "Domicilio:", 2)
        self.ent_cargo = self._campo(formulario, "Cargo:", 3,
                                     valor="CONDUCTOR DE TRANSPORTE DE CARGA")
        self.ent_remuneracion = self._campo(formulario, "Remuneración (S/):", 4, valor="1230.00")
        self.ent_inicio = self._campo(formulario, "Fecha de inicio:", 5, valor=hoy_texto())
        self.ent_fin = self._campo(formulario, "Fecha de vencimiento:", 6,
                                   valor=sumar_meses(hoy_texto(), 1))
        self.ent_ciudad_contrato = self._campo(formulario, "Ciudad de firma:", 7, valor="Lima")
        self.ent_firma_contrato = self._campo(formulario, "Fecha de firma:", 8, valor=hoy_texto())
        self.lbl_estado_contrato = ctk.CTkLabel(formulario, text="", font=(FAMILIA, 11, "bold"))
        self.lbl_estado_contrato.grid(row=9, column=0, columnspan=2, sticky="w", pady=(6, 0))

        botones = ctk.CTkFrame(derecha, fg_color="transparent")
        botones.pack(fill="x", padx=10, pady=10)
        ctk.CTkButton(botones, text="👁 Vista previa", width=130, fg_color=COLOR_GRIS,
                      command=self._vista_previa_contrato).pack(side="left")
        ctk.CTkButton(botones, text="📄 Generar contrato (PDF)", width=210, fg_color=COLOR_VERDE,
                      hover_color="#11806a",
                      command=self._generar_contrato).pack(side="left", padx=8)

        self._buscar_en_contratos()

    def _buscar_en_contratos(self):
        """Llena la lista de choferes y guarda los datos en memoria."""
        try:
            texto = self.ent_buscar_contrato.get().strip()
        except Exception:
            texto = ""
        for fila in self.tabla_choferes.get_children():
            self.tabla_choferes.delete(fila)
        self.choferes_cache = {}
        for registro in buscar_choferes(texto, self._ver_inactivos()):
            self.choferes_cache[str(registro[0])] = registro
            estado_txt, _dias = estado_contrato(registro[6])
            self.tabla_choferes.insert("", tk.END, iid=str(registro[0]),
                                       values=(registro[1] or "", registro[2] or "",
                                               registro[4] or "", registro[8] or "",
                                               estado_txt, registro[6] or "—"))
        # las otras pestañas pueden no existir todavía al construir la primera
        if hasattr(self, "tabla_renov"):
            self._cargar_choferes_en_renovaciones()
        if hasattr(self, "tabla_actas_choferes"):
            self._cargar_choferes_en_actas()

    def _on_chofer_contrato(self, _evento=None):
        seleccion = self.tabla_choferes.selection()
        if not seleccion:
            return
        registro = (getattr(self, "choferes_cache", {}) or {}).get(seleccion[0])
        if not registro:
            return
        self._elegir_chofer(registro)
        inicio = self.chofer["fecha_inicio_contrato"] or hoy_texto()
        self._poner(self.ent_trabajador, self.chofer["nombres"])
        self._poner(self.ent_dni, self.chofer["dni"])
        self._poner(self.ent_direccion, self.chofer["direccion"])
        self._poner(self.ent_inicio, inicio)
        if self.chofer["fecha_fin_contrato"]:
            self._poner(self.ent_fin, self.chofer["fecha_fin_contrato"])
        else:
            self._poner(self.ent_fin, sumar_meses(inicio, 1))
        estado_txt, _dias = estado_contrato(self.chofer["fecha_fin_contrato"])
        self.lbl_estado_contrato.configure(text="Estado actual: " + estado_txt,
                                           text_color=COLOR_AZUL)

    def _datos_documento(self):
        """Junta los datos del formulario en un diccionario."""
        remuneracion = numero_a_texto(self.ent_remuneracion.get())
        valores = dict(self.empresa)
        valores.update({
            "trabajador": self.ent_trabajador.get().strip(),
            "dni": self.ent_dni.get().strip(),
            "direccion_trabajador": self.ent_direccion.get().strip(),
            "cargo": self.ent_cargo.get().strip(),
            "remuneracion_texto": remuneracion.replace(".", ",") if "," not in remuneracion else remuneracion,
            "remuneracion_letras": pdf.remuneracion_en_letras(remuneracion),
            "fecha_inicio": self.ent_inicio.get().strip(),
            "fecha_fin": self.ent_fin.get().strip(),
            "ciudad_firma": self.ent_ciudad_contrato.get().strip(),
            "fecha_firma": self.ent_firma_contrato.get().strip(),
        })
        return valores

    def _validar_contrato(self, valores, es_renovacion=False):
        """Comprueba lo mínimo antes de generar."""
        if not valores.get("empresa") or not valores.get("ruc_empresa"):
            return "Faltan los datos de la empresa (Ajustes ▸ Configuración General)."
        if not valores.get("trabajador") or not valores.get("dni"):
            return "Elija un chofer de la lista."
        if not valores.get("cargo"):
            return "Escriba el cargo del trabajador."
        if not valores.get("fecha_inicio") or not valores.get("fecha_fin"):
            return "Complete las fechas de inicio y vencimiento."
        for etiqueta, fecha in (("inicio", valores.get("fecha_inicio")),
                                ("vencimiento", valores.get("fecha_fin"))):
            try:
                datetime.strptime(fecha, "%d/%m/%Y")
            except Exception:
                return "La fecha de " + etiqueta + " debe ser DD/MM/AAAA."
        if not valores.get("ciudad_firma"):
            return "Escriba la ciudad de firma."
        return ""

    def _vista_previa_contrato(self):
        valores = self._datos_documento()
        error = self._validar_contrato(valores)
        if error:
            return messagebox.showwarning("Contratos y Actas", error, parent=self.parent_frame)
        estado_txt, dias = estado_contrato(valores.get("fecha_fin"))
        texto = (
            "CONTRATO DE NATURALEZA TEMPORAL – LEY MYPE\n\n"
            "Empleador : " + valores["empresa"] + "  (RUC " + valores["ruc_empresa"] + ")\n"
            "Domicilio : " + valores["domicilio_empresa"] + "\n\n"
            "Trabajador: " + valores["trabajador"] + "\n"
            "DNI       : " + valores["dni"] + "\n"
            "Domicilio : " + valores["direccion_trabajador"] + "\n"
            "Cargo     : " + valores["cargo"] + "\n"
            "Remuneración: S/ " + valores["remuneracion_texto"] +
            "  (" + valores["remuneracion_letras"] + ")\n\n"
            "Vigencia  : " + valores["fecha_inicio"] + "  →  " + valores["fecha_fin"] + "\n"
            "Estado    : " + estado_txt + "\n"
            "Firma     : " + valores["ciudad_firma"] + ", " + valores["fecha_firma"] + "\n"
        )
        ventana = ctk.CTkToplevel(self.parent_frame)
        ventana.title("Vista previa del contrato")
        ventana.geometry("720x560")
        ventana.transient(self.parent_frame.winfo_toplevel() if hasattr(self.parent_frame, "winfo_toplevel") else None)
        ctk.CTkLabel(ventana, text="Vista previa del contrato", font=(FAMILIA, 15, "bold"),
                     text_color=COLOR_AZUL).pack(pady=8)
        caja = ctk.CTkTextbox(ventana, font=(FAMILIA, 11), wrap="word")
        caja.pack(fill="both", expand=True, padx=14, pady=6)
        caja.insert("1.0", texto)
        caja.configure(state="disabled")
        ctk.CTkButton(ventana, text="Cerrar", width=120, fg_color=COLOR_GRIS,
                      command=ventana.destroy).pack(pady=8)
        try:
            ventana.grab_set()
        except Exception:
            pass
        return ventana

    def _ruta_destino(self, dni, nombre, tipo):
        marca = datetime.now().strftime("%Y%m%d_%H%M%S")
        carpeta = carpeta_chofer(dni, nombre)
        nombre_archivo = "%s_%s.pdf" % (tipo, marca)
        return os.path.join(carpeta, nombre_archivo)

    def _generar_contrato(self):
        valores = self._datos_documento()
        error = self._validar_contrato(valores)
        if error:
            return messagebox.showwarning("Contratos y Actas", error, parent=self.parent_frame)

        if not messagebox.askyesno("Generar contrato",
                                   "¿Generar el contrato de " + valores["trabajador"] + "?\n\n" +
                                   "Vigencia: " + valores["fecha_inicio"] + " → " + valores["fecha_fin"],
                                   parent=self.parent_frame):
            return
        try:
            ruta = self._ruta_destino(valores["dni"], valores["trabajador"], "Contrato")
            pdf.generar_contrato(valores, ruta, renovacion=False)
        except Exception as error_pdf:
            return messagebox.showerror("Contratos y Actas",
                                        "No se pudo generar el PDF:\n" + str(error_pdf),
                                        parent=self.parent_frame)

        self._registrar_contrato(valores, ruta, "Inicial", None)
        messagebox.showinfo("Contratos y Actas", "Contrato generado:\n" + os.path.basename(ruta),
                            parent=self.parent_frame)
        self._abrir_pdf(ruta)
        self._cargar_historicos()
        self._actualizar_estado_chofer(valores)
        return ruta

    def _registrar_contrato(self, valores, ruta, tipo, id_origen):
        conexion = conectar_db(silencioso=True)
        if not conexion:
            return None
        identificador = None
        try:
            cursor = conexion.cursor()
            cursor.execute("""
                INSERT INTO contratos (id_chofer, dni, trabajador, cargo, remuneracion,
                                       fecha_inicio, fecha_fin, fecha_inicio_original,
                                       ciudad_firma, fecha_firma, tipo, id_contrato_origen,
                                       ruta_pdf, usuario, creado)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                RETURNING id
            """, (self.chofer["id"] if self.chofer else None, valores.get("dni"),
                  valores.get("trabajador"), valores.get("cargo"),
                  float(numero_a_texto(valores.get("remuneracion_texto"))),
                  valores.get("fecha_inicio"), valores.get("fecha_fin"),
                  valores.get("fecha_inicio_original") or valores.get("fecha_inicio"),
                  valores.get("ciudad_firma"), valores.get("fecha_firma"), tipo, id_origen,
                  ruta, self.usuario, datetime.now().strftime("%d/%m/%Y %H:%M:%S")))
            identificador = cursor.fetchone()[0]
            conexion.commit()
            self._registrar_documento(conexion, tipo, identificador, valores, ruta)
        except Exception as error:
            print("Aviso - registrar contrato:", error)
            try:
                conexion.rollback()
            except Exception:
                pass
        finally:
            liberar_conexion(conexion)
        return identificador

    def _registrar_documento(self, conexion, tipo, id_origen, valores, ruta):
        try:
            cursor = conexion.cursor()
            cursor.execute("""
                INSERT INTO contratos_documentos
                    (tipo, id_origen, id_chofer, dni, trabajador, descripcion, archivo,
                     usuario, creado)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (tipo, id_origen, self.chofer["id"] if self.chofer else None,
                  valores.get("dni"), valores.get("trabajador"),
                  tipo + " · " + str(valores.get("fecha_inicio") or "") + " al " +
                  str(valores.get("fecha_fin") or ""),
                  ruta, self.usuario, datetime.now().strftime("%d/%m/%Y %H:%M:%S")))
            conexion.commit()
        except Exception as error:
            print("Aviso - registrar documento:", error)
            try:
                conexion.rollback()
            except Exception:
                pass

    def _actualizar_estado_chofer(self, valores):
        """Al generar el contrato inicial se guardan las fechas en la ficha.

        La fecha de inicio SOLO se escribe si el chofer todavía no la tenía.
        """
        if not self.chofer or not self.chofer.get("id"):
            return
        conexion = conectar_db(silencioso=True)
        if not conexion:
            return
        try:
            cursor = conexion.cursor()
            if self.chofer.get("fecha_inicio_contrato"):
                cursor.execute("""
                    UPDATE choferes SET fecha_fin_contrato = %s WHERE id = %s
                """, (valores.get("fecha_fin"), self.chofer["id"]))
            else:
                cursor.execute("""
                    UPDATE choferes SET fecha_inicio_contrato = %s, fecha_fin_contrato = %s
                    WHERE id = %s
                """, (valores.get("fecha_inicio"), valores.get("fecha_fin"), self.chofer["id"]))
            conexion.commit()
            self.chofer["fecha_inicio_contrato"] = (self.chofer.get("fecha_inicio_contrato")
                                                    or valores.get("fecha_inicio"))
            self.chofer["fecha_fin_contrato"] = valores.get("fecha_fin")
        except Exception as error:
            print("Aviso - actualizar ficha del chofer:", error)
            try:
                conexion.rollback()
            except Exception:
                pass
        finally:
            liberar_conexion(conexion)

    def _abrir_pdf(self, ruta):
        try:
            from modulo_compras import abrir_documento
            abrir_documento(ruta)
        except Exception as error:
            print("Aviso - abrir PDF:", error)


    # =====================================================
    # PESTAÑA 2 · RENOVACIONES
    # =====================================================
    def _crear_tab_renovaciones(self, padre):
        cuerpo = ctk.CTkFrame(padre, fg_color="transparent")
        cuerpo.pack(fill="both", expand=True)

        busqueda = ctk.CTkFrame(cuerpo, fg_color="#f4f6f7", corner_radius=6)
        busqueda.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(busqueda, text="🔍 Chofer:", font=(FAMILIA, 11, "bold")).pack(side="left", padx=8)
        self.ent_buscar_renov = ctk.CTkEntry(busqueda, width=280,
                                             placeholder_text="Nombre, DNI o móvil…")
        self.ent_buscar_renov.pack(side="left", padx=4, pady=6)
        self.ent_buscar_renov.bind("<Return>", lambda _e: self._cargar_choferes_en_renovaciones())
        ctk.CTkButton(busqueda, text="Buscar", width=90, fg_color=COLOR_AZUL,
                      command=self._cargar_choferes_en_renovaciones).pack(side="left", padx=4)
        self.lbl_chofer_renovacion = ctk.CTkLabel(busqueda, text="(ningún chofer elegido)",
                                                  font=(FAMILIA, 11, "bold"), text_color="#c0392b")
        self.lbl_chofer_renovacion.pack(side="left", padx=12)

        zona = ctk.CTkFrame(cuerpo, fg_color="transparent")
        zona.pack(fill="both", expand=True)

        izquierda = ctk.CTkFrame(zona, fg_color="#ffffff", corner_radius=6)
        izquierda.pack(side="left", fill="both", expand=True, padx=(0, 6))
        ctk.CTkLabel(izquierda, text="Choferes", font=(FAMILIA, 12, "bold"),
                     text_color=COLOR_AZUL).pack(anchor="w", padx=8, pady=(6, 2))
        self.tabla_renov = self._tabla(izquierda,
                                       ("DNI", "Nombres", "Estado", "Contrato", "Vence contrato"),
                                       (80, 190, 70, 90, 100), alto=12)
        self.tabla_renov.bind("<<TreeviewSelect>>", self._on_chofer_renovacion)

        derecha = ctk.CTkFrame(zona, fg_color="#ffffff", corner_radius=6)
        derecha.pack(side="left", fill="both", expand=True)
        ctk.CTkLabel(derecha, text="Situación del contrato", font=(FAMILIA, 12, "bold"),
                     text_color=COLOR_AZUL).pack(anchor="w", padx=10, pady=(6, 2))
        self.lbl_situacion = ctk.CTkLabel(derecha, text="", font=(FAMILIA, 11),
                                          justify="left", anchor="w")
        self.lbl_situacion.pack(fill="x", padx=10, pady=(2, 6))

        ctk.CTkLabel(derecha, text="Historial de contratos y renovaciones",
                     font=(FAMILIA, 11, "bold"), text_color="#2c3e50").pack(anchor="w", padx=10)
        self.tabla_historial = self._tabla(derecha, ("Tipo", "Desde", "Hasta", "Registrado"),
                                           (110, 100, 100, 150), alto=7)
        self.tabla_historial.pack_configure(pady=4)

        botones = ctk.CTkFrame(derecha, fg_color="transparent")
        botones.pack(fill="x", padx=10, pady=8)
        ctk.CTkButton(botones, text="🔄 Renovar contrato", width=190, fg_color=COLOR_VERDE,
                      hover_color="#11806a",
                      command=self._renovar_contrato).pack(side="left")
        ctk.CTkLabel(botones, text="Solo se actualiza la fecha de vencimiento",
                     font=(FAMILIA, 10, "italic"), text_color="#7f8c8d").pack(side="left", padx=10)

    def _cargar_choferes_en_renovaciones(self):
        try:
            texto = self.ent_buscar_renov.get().strip()
        except Exception:
            texto = ""
        for fila in self.tabla_renov.get_children():
            self.tabla_renov.delete(fila)
        for registro in buscar_choferes(texto, self._ver_inactivos()):
            estado_txt, _dias = estado_contrato(registro[6])
            self.tabla_renov.insert("", tk.END, iid=str(registro[0]),
                                    values=(registro[1] or "", registro[2] or "",
                                            registro[8] or "", estado_txt,
                                            registro[6] or "—"))

    def _on_chofer_renovacion(self, _evento=None):
        seleccion = self.tabla_renov.selection()
        if not seleccion:
            return
        identificador = seleccion[0]
        for registro in buscar_choferes("", True):
            if str(registro[0]) == str(identificador):
                self._elegir_chofer(registro)
                break
        if not self.chofer:
            return
        inicio_original = self.chofer["fecha_inicio_contrato"] or "(sin registrar)"
        vencimiento = self.chofer["fecha_fin_contrato"] or "(sin registrar)"
        estado_txt, dias = estado_contrato(self.chofer["fecha_fin_contrato"])
        extra = ""
        if dias is not None:
            extra = ("  ·  venció hace %d día(s)" % abs(dias) if dias < 0
                     else "  ·  quedan %d día(s)" % dias)
        self.lbl_situacion.configure(
            text="Trabajador: " + self.chofer["nombres"] + "\n"
                 "DNI: " + self.chofer["dni"] + "\n"
                 "Fecha de inicio ORIGINAL: " + inicio_original + "\n"
                 "Vencimiento actual: " + vencimiento + "\n"
                 "Estado: " + estado_txt + extra)

        for fila in self.tabla_historial.get_children():
            self.tabla_historial.delete(fila)
        for contrato in contratos_del_chofer(self.chofer["id"]):
            self.tabla_historial.insert("", tk.END,
                                        values=(contrato[4] or "Inicial",
                                                contrato[1] or "", contrato[2] or "",
                                                contrato[8] or ""))
        for renovacion in renovaciones_del_chofer(self.chofer["id"]):
            self.tabla_historial.insert("", tk.END,
                                        values=("Renovación",
                                                renovacion[1] or "", renovacion[2] or "",
                                                renovacion[3] or ""))

    def _renovar_contrato(self):
        if not self.chofer:
            return messagebox.showwarning("Renovaciones", "Elija primero un chofer.",
                                          parent=self.parent_frame)
        if not self.chofer.get("fecha_inicio_contrato"):
            return messagebox.showwarning(
                "Renovaciones",
                "Este chofer no tiene contrato registrado.\nGenere primero el contrato inicial.",
                parent=self.parent_frame)

        anterior = self.chofer.get("fecha_fin_contrato") or ""
        sugerida = sumar_meses(anterior or hoy_texto(), 1)

        ventana = ctk.CTkToplevel(self.parent_frame)
        ventana.title("Renovar contrato")
        ventana.geometry("480x330")
        ventana.grab_set()
        ctk.CTkLabel(ventana, text="🔄 Renovar contrato", font=(FAMILIA, 15, "bold"),
                     text_color=COLOR_AZUL).pack(pady=10)
        ctk.CTkLabel(ventana, text=self.chofer["nombres"] + " · DNI " + self.chofer["dni"],
                     font=(FAMILIA, 11, "bold")).pack()
        ctk.CTkLabel(ventana, text="Inicio original (no cambia): " +
                                   self.chofer["fecha_inicio_contrato"],
                     font=(FAMILIA, 11), text_color="#7f8c8d").pack(pady=(4, 0))
        ctk.CTkLabel(ventana, text="Vencimiento actual: " + (anterior or "(ninguno)"),
                     font=(FAMILIA, 11), text_color="#7f8c8d").pack()
        ctk.CTkLabel(ventana, text="Nuevo vencimiento (DD/MM/AAAA):",
                     font=(FAMILIA, 11, "bold")).pack(pady=(12, 2))
        entrada = ctk.CTkEntry(ventana, width=170, font=(FAMILIA, 12))
        entrada.pack()
        entrada.insert(0, sugerida)
        ctk.CTkLabel(ventana, text="Ciudad de firma:", font=(FAMILIA, 11, "bold")).pack(pady=(10, 2))
        ciudad = ctk.CTkEntry(ventana, width=170, font=(FAMILIA, 12))
        ciudad.pack()
        ciudad.insert(0, "Lima")

        def aceptar():
            nuevo = entrada.get().strip()
            try:
                fecha_nueva = datetime.strptime(nuevo, "%d/%m/%Y")
            except Exception:
                return messagebox.showwarning("Renovaciones", "Escriba la fecha como DD/MM/AAAA.",
                                              parent=ventana)
            try:
                inicio = datetime.strptime(self.chofer["fecha_inicio_contrato"], "%d/%m/%Y")
            except Exception:
                inicio = None
            if inicio and fecha_nueva <= inicio:
                return messagebox.showwarning(
                    "Renovaciones",
                    "El nuevo vencimiento no puede ser anterior ni igual a la fecha de inicio original.",
                    parent=ventana)
            if anterior:
                try:
                    if fecha_nueva <= datetime.strptime(anterior, "%d/%m/%Y"):
                        if not messagebox.askyesno(
                                "Renovaciones",
                                "El nuevo vencimiento NO es posterior al actual (" + anterior + ").\n"
                                "¿Desea continuar de todos modos?", parent=ventana):
                            return
                except Exception:
                    pass
            if not messagebox.askyesno(
                    "Confirmar renovación",
                    "¿Renovar el contrato de " + self.chofer["nombres"] + "?\n\n"
                    "Vencimiento: " + (anterior or "-") + "  →  " + nuevo + "\n"
                    "La fecha de inicio original (" + self.chofer["fecha_inicio_contrato"] +
                    ") NO se modifica.", parent=ventana):
                return
            ventana.destroy()
            self._ejecutar_renovacion(nuevo, ciudad.get().strip() or "Lima")

        ctk.CTkButton(ventana, text="✔ Renovar", width=140, fg_color=COLOR_VERDE,
                      hover_color="#11806a", command=aceptar).pack(pady=14)

    def _ejecutar_renovacion(self, nuevo_vencimiento, ciudad):
        inicio_original = self.chofer["fecha_inicio_contrato"]
        anterior = self.chofer.get("fecha_fin_contrato") or ""
        valores = dict(self.empresa)
        valores.update({
            "trabajador": self.chofer["nombres"],
            "dni": self.chofer["dni"],
            "direccion_trabajador": self.chofer["direccion"],
            "cargo": "CONDUCTOR DE TRANSPORTE DE CARGA",
            "remuneracion_texto": "1,230.00",
            "remuneracion_letras": pdf.remuneracion_en_letras(1230),
            "fecha_inicio": anterior or sumar_meses(nuevo_vencimiento, -1),
            "fecha_fin": nuevo_vencimiento,
            "fecha_inicio_original": inicio_original,
            "vencimiento_anterior": anterior,
            "ciudad_firma": ciudad,
            "fecha_firma": hoy_texto(),
        })
        if not valores.get("empresa") or not valores.get("ruc_empresa"):
            return messagebox.showerror(
                "Renovaciones", "Faltan los datos de la empresa en Configuración General.",
                parent=self.parent_frame)
        try:
            ruta = self._ruta_destino(self.chofer["dni"], self.chofer["nombres"], "Renovacion")
            pdf.generar_contrato(valores, ruta, renovacion=True)
        except Exception as error:
            return messagebox.showerror("Renovaciones",
                                        "No se pudo generar el PDF:\n" + str(error),
                                        parent=self.parent_frame)

        id_contrato = self._registrar_contrato(valores, ruta, "Renovación", None)
        self._registrar_renovacion(valores, ruta, id_contrato, anterior, nuevo_vencimiento)
        self._actualizar_vencimiento_chofer(nuevo_vencimiento)
        messagebox.showinfo("Renovaciones",
                            "Contrato de renovación generado y registrado.\n" +
                            os.path.basename(ruta), parent=self.parent_frame)
        self._abrir_pdf(ruta)
        self._on_chofer_renovacion()
        self._cargar_historicos()

    def _registrar_renovacion(self, valores, ruta, id_contrato, anterior, nuevo):
        conexion = conectar_db(silencioso=True)
        if not conexion:
            return
        try:
            cursor = conexion.cursor()
            cursor.execute("""
                INSERT INTO renovaciones (id_contrato, id_chofer, dni, trabajador,
                                          fecha_inicio_original, vencimiento_anterior,
                                          nuevo_vencimiento, fecha_renovacion, usuario,
                                          ruta_pdf, creado)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (id_contrato, self.chofer["id"], self.chofer["dni"],
                  self.chofer["nombres"], valores.get("fecha_inicio_original"), anterior,
                  nuevo, hoy_texto(), self.usuario, ruta,
                  datetime.now().strftime("%d/%m/%Y %H:%M:%S")))
            conexion.commit()
            try:
                registrar_auditoria(self.usuario, "Contratos y Actas",
                                    "Renovó el contrato de " + str(self.chofer["nombres"]) +
                                    " al " + str(nuevo))
            except Exception:
                pass
        except Exception as error:
            print("Aviso - registrar renovación:", error)
            try:
                conexion.rollback()
            except Exception:
                pass
        finally:
            liberar_conexion(conexion)

    def _actualizar_vencimiento_chofer(self, nuevo_vencimiento):
        """ÚNICO campo que cambia en la ficha: fecha_fin_contrato."""
        conexion = conectar_db(silencioso=True)
        if not conexion:
            return
        try:
            cursor = conexion.cursor()
            cursor.execute("UPDATE choferes SET fecha_fin_contrato = %s WHERE id = %s",
                           (nuevo_vencimiento, self.chofer["id"]))
            conexion.commit()
            self.chofer["fecha_fin_contrato"] = nuevo_vencimiento
        except Exception as error:
            print("Aviso - actualizar vencimiento:", error)
            try:
                conexion.rollback()
            except Exception:
                pass
        finally:
            liberar_conexion(conexion)


    # =====================================================
    # PESTAÑA 3 · ACTAS DE ENTREGA
    # =====================================================
    def _crear_tab_actas(self, padre):
        cuerpo = ctk.CTkFrame(padre, fg_color="transparent")
        cuerpo.pack(fill="both", expand=True)

        busqueda = ctk.CTkFrame(cuerpo, fg_color="#f4f6f7", corner_radius=6)
        busqueda.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(busqueda, text="🔍 Chofer:", font=(FAMILIA, 11, "bold")).pack(side="left", padx=8)
        self.ent_buscar_acta = ctk.CTkEntry(busqueda, width=260,
                                            placeholder_text="Nombre, DNI o móvil…")
        self.ent_buscar_acta.pack(side="left", padx=4, pady=6)
        self.ent_buscar_acta.bind("<Return>", lambda _e: self._cargar_choferes_en_actas())
        ctk.CTkButton(busqueda, text="Buscar", width=90, fg_color=COLOR_AZUL,
                      command=self._cargar_choferes_en_actas).pack(side="left", padx=4)
        self.lbl_chofer_acta = ctk.CTkLabel(busqueda, text="(ningún chofer elegido)",
                                            font=(FAMILIA, 11, "bold"), text_color="#c0392b")
        self.lbl_chofer_acta.pack(side="left", padx=12)

        zona = ctk.CTkFrame(cuerpo, fg_color="transparent")
        zona.pack(fill="both", expand=True)

        izquierda = ctk.CTkFrame(zona, fg_color="#ffffff", corner_radius=6)
        izquierda.pack(side="left", fill="both", expand=True, padx=(0, 6))
        ctk.CTkLabel(izquierda, text="Choferes", font=(FAMILIA, 12, "bold"),
                     text_color=COLOR_AZUL).pack(anchor="w", padx=8, pady=(6, 2))
        self.tabla_actas_choferes = self._tabla(izquierda,
                                                ("DNI", "Nombres", "Móvil", "Estado"),
                                                (85, 200, 80, 75), alto=12)
        self.tabla_actas_choferes.bind("<<TreeviewSelect>>", self._on_chofer_acta)

        derecha = ctk.CTkFrame(zona, fg_color="#ffffff", corner_radius=6)
        derecha.pack(side="left", fill="both", expand=True)
        ctk.CTkLabel(derecha, text="Datos del acta", font=(FAMILIA, 12, "bold"),
                     text_color=COLOR_AZUL).pack(anchor="w", padx=10, pady=(6, 2))

        formulario = ctk.CTkFrame(derecha, fg_color="transparent")
        formulario.pack(fill="x", padx=10)
        self.ent_cargo_acta = self._campo(formulario, "Cargo:", 0,
                                          valor="CONDUCTOR DE TRANSPORTE DE CARGA")
        self.ent_placa = self._campo(formulario, "Placa asignada:", 1)
        ctk.CTkButton(formulario, text="↻", width=30, fg_color=COLOR_GRIS,
                      command=self._traer_vehiculo_del_chofer).grid(row=1, column=2, padx=4)
        self.lbl_vehiculo = ctk.CTkLabel(formulario, text="(vehículo sin datos)",
                                         font=(FAMILIA, 11), text_color="#7f8c8d")
        self.lbl_vehiculo.grid(row=2, column=0, columnspan=3, sticky="w", pady=(2, 4))
        self.ent_celular = self._campo(formulario, "Celular (marca/modelo):", 3,
                                       valor="Samsung Galaxy A04E")
        self.ent_imei = self._campo(formulario, "IMEI / serie:", 4)
        self.ent_ciudad_acta = self._campo(formulario, "Ciudad de firma:", 5, valor="Lima")
        self.ent_fecha_acta = self._campo(formulario, "Fecha del acta:", 6, valor=hoy_texto())

        ctk.CTkLabel(derecha, text="Bienes y equipos entregados (uno por línea):",
                     font=(FAMILIA, 11, "bold"), text_color="#2c3e50").pack(anchor="w",
                                                                           padx=10, pady=(8, 2))
        self.txt_bienes = ctk.CTkTextbox(derecha, height=110, font=(FAMILIA, 11))
        self.txt_bienes.pack(fill="x", padx=10)
        self.txt_bienes.insert("1.0", "\n".join(PL.ACTA_BIENES_POR_DEFECTO))

        botones = ctk.CTkFrame(derecha, fg_color="transparent")
        botones.pack(fill="x", padx=10, pady=8)
        ctk.CTkButton(botones, text="📄 Generar acta (PDF)", width=190, fg_color=COLOR_VERDE,
                      hover_color="#11806a",
                      command=self._generar_acta).pack(side="left")
        ctk.CTkLabel(botones, text="El acta no modifica el móvil asignado del chofer",
                     font=(FAMILIA, 10, "italic"), text_color="#7f8c8d").pack(side="left", padx=10)

    def _cargar_choferes_en_actas(self):
        try:
            texto = self.ent_buscar_acta.get().strip()
        except Exception:
            texto = ""
        for fila in self.tabla_actas_choferes.get_children():
            self.tabla_actas_choferes.delete(fila)
        for registro in buscar_choferes(texto, self._ver_inactivos()):
            self.tabla_actas_choferes.insert("", tk.END, iid=str(registro[0]),
                                             values=(registro[1] or "", registro[2] or "",
                                                     registro[4] or "", registro[8] or ""))

    def _on_chofer_acta(self, _evento=None):
        seleccion = self.tabla_actas_choferes.selection()
        if not seleccion:
            return
        identificador = seleccion[0]
        for registro in buscar_choferes("", True):
            if str(registro[0]) == str(identificador):
                self._elegir_chofer(registro)
                break
        if not self.chofer:
            return
        self._poner(self.ent_placa, movil_valido(self.chofer.get("movil")))
        self._traer_vehiculo_del_chofer()

    def _traer_vehiculo_del_chofer(self):
        """Completa marca/modelo/año desde flota_vehiculos (solo lectura)."""
        placa = movil_valido(self.ent_placa.get()) or movil_valido((self.chofer or {}).get("movil"))
        self._poner(self.ent_placa, placa)
        self.vehiculo = datos_vehiculo(placa)
        if not self.vehiculo:
            self.lbl_vehiculo.configure(text="(no se encontró el vehículo " + (placa or "-") + ")",
                                        text_color="#c0392b")
            return
        self.lbl_vehiculo.configure(
            text="Vehículo: " + str(self.vehiculo.get("marca", "")) + " " +
                 str(self.vehiculo.get("modelo", "")) + " · Año " +
                 str(self.vehiculo.get("anio", "")) + " · Placa " +
                 str(self.vehiculo.get("placa", "")),
            text_color="#2c3e50")

    def _bienes_del_acta(self):
        texto = self.txt_bienes.get("1.0", "end-1c")
        return [linea.strip() for linea in texto.splitlines() if linea.strip()]

    def _generar_acta(self):
        if not self.chofer:
            return messagebox.showwarning("Actas de Entrega", "Elija primero un chofer.",
                                          parent=self.parent_frame)
        valores = dict(self.empresa)
        valores.update({
            "trabajador": self.chofer["nombres"],
            "dni": self.chofer["dni"],
            "direccion_trabajador": self.chofer["direccion"],
            "cargo": self.ent_cargo_acta.get().strip(),
            "ciudad_firma": self.ent_ciudad_acta.get().strip() or "Lima",
            "fecha_firma": self.ent_fecha_acta.get().strip() or hoy_texto(),
            "celular_marca_modelo": self.ent_celular.get().strip(),
            "vehiculo_marca": self.vehiculo.get("marca", ""),
            "vehiculo_modelo": self.vehiculo.get("modelo", ""),
            "vehiculo_anio": self.vehiculo.get("anio", ""),
            "vehiculo_placa": (self.vehiculo.get("placa", "")
                               or movil_valido(self.ent_placa.get())),
            "bienes": self._bienes_del_acta(),
        })
        if not valores.get("empresa") or not valores.get("ruc_empresa"):
            return messagebox.showerror("Actas de Entrega",
                                        "Faltan los datos de la empresa en Configuración General.",
                                        parent=self.parent_frame)
        if not valores.get("bienes"):
            return messagebox.showwarning("Actas de Entrega",
                                          "Escriba al menos un bien entregado.",
                                          parent=self.parent_frame)
        if not messagebox.askyesno("Generar acta",
                                   "¿Generar el acta de entrega de " +
                                   valores["trabajador"] + "?", parent=self.parent_frame):
            return
        try:
            ruta = self._ruta_destino(valores["dni"], valores["trabajador"], "Acta")
            pdf.generar_acta(valores, ruta)
        except Exception as error:
            return messagebox.showerror("Actas de Entrega",
                                        "No se pudo generar el PDF:\n" + str(error),
                                        parent=self.parent_frame)
        self._registrar_acta(valores, ruta)
        messagebox.showinfo("Actas de Entrega",
                            "Acta generada:\n" + os.path.basename(ruta),
                            parent=self.parent_frame)
        self._abrir_pdf(ruta)
        self._cargar_historicos()

    def _registrar_acta(self, valores, ruta):
        conexion = conectar_db(silencioso=True)
        if not conexion:
            return
        try:
            cursor = conexion.cursor()
            cursor.execute("""
                INSERT INTO actas_entrega (id_chofer, dni, trabajador, cargo, fecha_acta,
                                           ciudad_firma, bienes, celular_marca_modelo,
                                           celular_imei, vehiculo_marca, vehiculo_modelo,
                                           vehiculo_anio, vehiculo_placa, ruta_pdf,
                                           usuario, creado)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                RETURNING id
            """, (self.chofer["id"], valores.get("dni"), valores.get("trabajador"),
                  valores.get("cargo"), valores.get("fecha_firma"), valores.get("ciudad_firma"),
                  json.dumps(valores.get("bienes"), ensure_ascii=False),
                  valores.get("celular_marca_modelo"), self.ent_imei.get().strip(),
                  valores.get("vehiculo_marca"), valores.get("vehiculo_modelo"),
                  valores.get("vehiculo_anio"), valores.get("vehiculo_placa"), ruta,
                  self.usuario, datetime.now().strftime("%d/%m/%Y %H:%M:%S")))
            identificador = cursor.fetchone()[0]
            conexion.commit()
            self._registrar_documento(conexion, "Acta", identificador, valores, ruta)
            try:
                registrar_auditoria(self.usuario, "Contratos y Actas",
                                    "Generó el acta de entrega de " + str(valores.get("trabajador")))
            except Exception:
                pass
        except Exception as error:
            print("Aviso - registrar acta:", error)
            try:
                conexion.rollback()
            except Exception:
                pass
        finally:
            liberar_conexion(conexion)

    # =====================================================
    # PESTAÑA 4 · HISTÓRICOS
    # =====================================================
    def _crear_tab_historicos(self, padre):
        cuerpo = ctk.CTkFrame(padre, fg_color="transparent")
        cuerpo.pack(fill="both", expand=True)

        barra = ctk.CTkFrame(cuerpo, fg_color="#f4f6f7", corner_radius=6)
        barra.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(barra, text="🔍 Filtrar:", font=(FAMILIA, 11, "bold")).pack(side="left", padx=8)
        self.ent_buscar_historico = ctk.CTkEntry(barra, width=280,
                                                  placeholder_text="Chofer, DNI o tipo…")
        self.ent_buscar_historico.pack(side="left", padx=4, pady=6)
        self.ent_buscar_historico.bind("<Return>", lambda _e: self._cargar_historicos())
        ctk.CTkButton(barra, text="Filtrar", width=90, fg_color=COLOR_AZUL,
                      command=self._cargar_historicos).pack(side="left", padx=4)
        ctk.CTkButton(barra, text="🔄 Actualizar", width=110, fg_color=COLOR_VERDE,
                      hover_color="#11806a",
                      command=self._cargar_historicos).pack(side="left", padx=4)
        ctk.CTkButton(barra, text="📂 Abrir carpeta", width=130, fg_color=COLOR_GRIS,
                      command=self._abrir_carpeta_base).pack(side="right", padx=8)

        self.tabla_historicos = self._tabla(cuerpo,
                                            ("Fecha", "Tipo", "Trabajador", "DNI",
                                             "Descripción", "Usuario"),
                                            (120, 100, 200, 80, 240, 100), alto=16)
        self.tabla_historicos.bind("<Double-1>", lambda _e: self._abrir_historico_seleccionado())

        botones = ctk.CTkFrame(cuerpo, fg_color="transparent")
        botones.pack(fill="x", pady=6)
        ctk.CTkButton(botones, text="📄 Abrir PDF", width=130, fg_color=COLOR_AZUL,
                      command=self._abrir_historico_seleccionado).pack(side="left")
        self.lbl_total_historicos = ctk.CTkLabel(botones, text="", font=(FAMILIA, 10, "italic"),
                                                 text_color="#7f8c8d")
        self.lbl_total_historicos.pack(side="left", padx=12)

        self._cargar_historicos()

    def _cargar_historicos(self):
        try:
            filtro = self.ent_buscar_historico.get().strip()
        except Exception:
            filtro = ""
        for fila in self.tabla_historicos.get_children():
            self.tabla_historicos.delete(fila)
        conexion = conectar_db(silencioso=True)
        if not conexion:
            return
        documentos = []
        try:
            cursor = conexion.cursor()
            patron = "%" + filtro + "%"
            cursor.execute("""
                SELECT id, tipo, trabajador, dni, descripcion, usuario, archivo, creado
                FROM contratos_documentos
                WHERE COALESCE(trabajador, '') ILIKE %s
                   OR COALESCE(dni, '') ILIKE %s
                   OR COALESCE(tipo, '') ILIKE %s
                ORDER BY id DESC LIMIT 500
            """, (patron, patron, patron))
            documentos = cursor.fetchall()
        except Exception as error:
            print("Aviso - históricos:", error)
        finally:
            liberar_conexion(conexion)

        self._historicos = {}
        for fila in documentos:
            self._historicos[str(fila[0])] = fila[6]
            self.tabla_historicos.insert("", tk.END, iid=str(fila[0]),
                                         values=(fila[7] or "", fila[1] or "",
                                                 fila[2] or "", fila[3] or "",
                                                 fila[4] or "", fila[5] or ""))
        try:
            self.lbl_total_historicos.configure(text=str(len(documentos)) + " documento(s)")
        except Exception:
            pass

    def _abrir_historico_seleccionado(self):
        seleccion = self.tabla_historicos.selection()
        if not seleccion:
            return
        ruta = (getattr(self, "_historicos", {}) or {}).get(seleccion[0])
        if ruta and os.path.exists(ruta):
            self._abrir_pdf(ruta)
        else:
            messagebox.showwarning("Históricos", "El archivo no se encuentra en disco.",
                                   parent=self.parent_frame)

    def _abrir_carpeta_base(self):
        carpeta = ruta_base()
        try:
            if sys.platform == "win32":
                os.startfile(carpeta)
            elif sys.platform == "darwin":
                os.system("open '" + carpeta + "'")
            else:
                os.system("xdg-open '" + carpeta + "'")
        except Exception as error:
            messagebox.showerror("Históricos", "No se pudo abrir la carpeta:\n" + str(error),
                                 parent=self.parent_frame)
