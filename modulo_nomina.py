# -*- coding: utf-8 -*-
"""
MODULO_NOMINA.PY - MÓDULO DE NÓMINA Y CONTROL DE ASISTENCIA
============================================================
Interfaz del módulo de nómina del Sistema de Control de Flota Automotriz.

- Carga el Excel del reloj biométrico (captahuella) con detección automática
  de formato (Hikvision, ZKTeco, genérico) y asistente de mapeo de columnas.
- Cruza la primera y la última marcación del día con los turnos y horarios
  asignados a cada empleado y calcula tardanzas, faltas, horas extra,
  descansos, feriados e incidencias.
- Liquida la planilla del período (haberes, descuentos, aportes del
  empleador), emite boletas y reportes exportables a Excel.

Los datos se guardan en las tablas nom_* de la base de datos y el padrón de
empleados se toma de la tabla 'choferes' existente (datos de nómina en la
tabla satélite 'nom_empleado').

Autor: Collie Software
"""
import os
import re
import sys
import json
import calendar
import functools
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, date, timedelta

import customtkinter as ctk

from conexion import registrar_auditoria
from dialogos_seguros import seleccionar_archivo_dialogo, guardar_archivo_dialogo
from tareas_seguras import ejecutar_en_hilo

import nomina_core as nc
import nomina_planilla as npl

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

# =========================================================
# COLORES Y CONSTANTES DE ESTILO
# =========================================================
COLOR_PRIMARIO = "#1f538d"
COLOR_HOVER = "#163b65"
COLOR_OK = "#27ae60"
COLOR_ALERTA = "#e67e22"
COLOR_ERROR = "#c0392b"
COLOR_GRIS = "#7f8c8d"
COLOR_FONDO_SUAVE = "#f8f9fa"

# Códigos cortos y colores por estado (grilla mensual)
CODIGO_ESTADO = {
    nc.EST_PUNTUAL: "P",
    nc.EST_TARDANZA: "T",
    nc.EST_FALTA: "F",
    nc.EST_INCOMPLETO: "I",
    nc.EST_DESCANSO: "D",
    nc.EST_DESCANSO_TRAB: "DT",
    nc.EST_FERIADO: "FE",
    nc.EST_FERIADO_TRAB: "FT",
    nc.EST_VACACIONES: "V",
    nc.EST_PERMISO: "PE",
    nc.EST_LICENCIA: "LI",
    nc.EST_DESCANSO_MEDICO: "DM",
    nc.EST_COMISION: "CO",
    nc.EST_FALTA_JUST: "FJ",
    nc.EST_JUSTIFICADO: "JU",
    nc.EST_PENDIENTE: "-",
}

ETIQUETA_CAMPO_MAPEO = [
    ("codigo", "Código / DNI del reloj"),
    ("nombre", "Nombres"),
    ("apellido", "Apellidos"),
    ("nombre_completo", "Nombre completo (respaldo)"),
    ("fecha", "Fecha"),
    ("hora", "Hora"),
    ("tipo_pase", "Tipo de pase (entrada/salida)"),
    ("metodo", "Método de verificación"),
    ("departamento", "Departamento / Área"),
]

SIN_COLUMNA = "— (ninguna) —"


# =========================================================
# FUNCIONES AUXILIARES
# =========================================================
def abrir_documento(ruta):
    """Abre un archivo con la aplicación predeterminada del sistema."""
    try:
        if sys.platform == "win32":
            os.startfile(ruta)
        elif sys.platform == "darwin":
            subprocess.call(["open", ruta])
        else:
            subprocess.call(["xdg-open", ruta])
    except Exception as e:
        messagebox.showerror("Error", "No se pudo abrir el archivo:\n%s" % e)


def aplicar_estilo_treeview():
    """Aplica el estilo de tablas usado por el resto del sistema."""
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("Treeview", background="#ffffff", foreground="#000000",
                    fieldbackground="#ffffff", bordercolor="#e0e0e0",
                    borderwidth=1, rowheight=24, font=("Arial", 10))
    style.map("Treeview", background=[("selected", COLOR_PRIMARIO)],
              foreground=[("selected", "#ffffff")])
    style.configure("Treeview.Heading", background="#f0f0f0", foreground="#000000",
                    relief="flat", font=("Arial", 10, "bold"),
                    bordercolor="#e0e0e0", borderwidth=1)


def formatear_monto(valor):
    """S/ 1,234.56 (respeta la configuración regional del sistema)."""
    try:
        numero = float(valor or 0)
    except (TypeError, ValueError):
        numero = 0.0
    return "S/ %s" % format(numero, ",.2f")


def formatear_numero(valor, decimales=2):
    try:
        return format(float(valor or 0), ",.%df" % decimales)
    except (TypeError, ValueError):
        return "0"


def parse_num(texto, por_defecto=0.0):
    """Lee un número escrito por el usuario.

    Acepta las formas habituales: '1130', '1,130.50', '1.130,50' y '1,130'.
    La coma seguida de exactamente tres cifras se interpreta como separador de miles
    ('1,130' = 1130), porque así se muestran los montos en el sistema; en cualquier
    otro caso es decimal ('2,5' = 2.5). Si antes se escribía '1,130' el sistema
    entendía S/ 1.13.
    """
    if texto is None:
        return por_defecto
    limpio = str(texto).strip().replace("S/", "").replace(" ", "")
    if not limpio:
        return por_defecto
    if "," in limpio and "." in limpio:
        if limpio.rfind(",") > limpio.rfind("."):
            limpio = limpio.replace(".", "").replace(",", ".")
        else:
            limpio = limpio.replace(",", "")
    elif "," in limpio:
        partes = limpio.split(",")
        if partes[0] != "0" and all(len(grupo) == 3 for grupo in partes[1:]):
            limpio = limpio.replace(",", "")
        else:
            limpio = limpio.replace(",", ".")
    try:
        return float(limpio)
    except ValueError:
        return por_defecto


def parse_fecha_texto(texto):
    """Lee una fecha escrita como DD/MM/AAAA o AAAA-MM-DD y la devuelve en ISO."""
    if not texto:
        return None
    limpio = str(texto).strip()
    for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(limpio, formato).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def fecha_para_entry(fecha_iso):
    """'2026-09-09' -> '09/09/2026' (formato que se muestra en los campos)."""
    if not fecha_iso:
        return ""
    try:
        return datetime.strptime(str(fecha_iso)[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return str(fecha_iso)


def periodo_para_entry(periodo_iso):
    """'2026-09' -> '09/2026'."""
    try:
        anio, mes = str(periodo_iso).split("-")[:2]
        return "%s/%s" % (mes, anio)
    except Exception:
        return str(periodo_iso or "")


def periodo_desde_entry(texto):
    """'09/2026', '9-2026' o '2026-09' -> '2026-09'."""
    if not texto:
        return None
    limpio = str(texto).strip().replace(" ", "")
    coincidencia = re.match(r"^(\d{1,2})[/\-.](\d{4})$", limpio)
    if coincidencia:
        mes, anio = int(coincidencia.group(1)), int(coincidencia.group(2))
        if 1 <= mes <= 12:
            return "%04d-%02d" % (anio, mes)
    coincidencia = re.match(r"^(\d{4})[/\-.](\d{1,2})$", limpio)
    if coincidencia:
        anio, mes = int(coincidencia.group(1)), int(coincidencia.group(2))
        if 1 <= mes <= 12:
            return "%04d-%02d" % (anio, mes)
    return None


def sumar_meses(periodo, cantidad):
    """Suma (o resta) meses a un periodo 'AAAA-MM'."""
    try:
        anio, mes = [int(x) for x in str(periodo).split("-")[:2]]
    except Exception:
        return periodo
    total = (anio * 12 + (mes - 1)) + int(cantidad)
    return "%04d-%02d" % (total // 12, (total % 12) + 1)


def crear_tabla(padre, columnas, alto=10):
    """Crea un Treeview con scroll vertical. 'columnas' = [(clave, título, ancho, ancla)]."""
    contenedor = ctk.CTkFrame(padre, fg_color="transparent")
    contenedor.pack(fill="both", expand=True)
    claves = [c[0] for c in columnas]
    tabla = ttk.Treeview(contenedor, columns=claves, show="headings", height=alto, style="Treeview")
    for clave, titulo, ancho, ancla in columnas:
        tabla.heading(clave, text=titulo)
        tabla.column(clave, width=ancho, anchor=ancla, stretch=False)
    scroll_v = ttk.Scrollbar(contenedor, orient="vertical", command=tabla.yview)
    scroll_h = ttk.Scrollbar(contenedor, orient="horizontal", command=tabla.xview)
    tabla.configure(yscrollcommand=scroll_v.set, xscrollcommand=scroll_h.set)
    tabla.grid(row=0, column=0, sticky="nsew")
    scroll_v.grid(row=0, column=1, sticky="ns")
    scroll_h.grid(row=1, column=0, sticky="ew")
    contenedor.grid_rowconfigure(0, weight=1)
    contenedor.grid_columnconfigure(0, weight=1)
    return tabla


def limpiar_tabla(tabla):
    for item in tabla.get_children():
        tabla.delete(item)


class CalendarioNativo(ctk.CTkToplevel):
    """Calendario emergente que escribe la fecha elegida en un campo de texto."""

    def __init__(self, parent, target_entry):
        super().__init__(parent)
        self.target_entry = target_entry
        self.title("Seleccionar fecha")
        self.geometry("300x330")
        self.resizable(False, False)
        try:
            self.transient(parent)
            self.grab_set()
        except Exception:
            pass
        hoy = datetime.now()
        self.anio = hoy.year
        self.mes = hoy.month
        self.contenedor = ctk.CTkFrame(self, fg_color="transparent")
        self.contenedor.pack(fill="both", expand=True, padx=8, pady=8)
        self.dibujar()
        self.lift()
        self.focus_force()

    def dibujar(self):
        for widget in self.contenedor.winfo_children():
            widget.destroy()
        cabecera = ctk.CTkFrame(self.contenedor, fg_color="transparent")
        cabecera.pack(fill="x", pady=(0, 6))
        ctk.CTkButton(cabecera, text="◀", width=36, fg_color=COLOR_PRIMARIO, hover_color=COLOR_HOVER,
                      command=lambda: self.cambiar_mes(-1)).pack(side="left")
        nombres = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
                   "Setiembre", "Octubre", "Noviembre", "Diciembre"]
        ctk.CTkLabel(cabecera, text="%s %d" % (nombres[self.mes], self.anio),
                     font=("Arial", 13, "bold")).pack(side="left", expand=True)
        ctk.CTkButton(cabecera, text="▶", width=36, fg_color=COLOR_PRIMARIO, hover_color=COLOR_HOVER,
                      command=lambda: self.cambiar_mes(1)).pack(side="right")
        grid = ctk.CTkFrame(self.contenedor, fg_color="transparent")
        grid.pack(fill="both", expand=True)
        for indice, dia in enumerate(["Lu", "Ma", "Mi", "Ju", "Vi", "Sa", "Do"]):
            ctk.CTkLabel(grid, text=dia, font=("Arial", 10, "bold"), width=36,
                         text_color=COLOR_GRIS).grid(row=0, column=indice, padx=1, pady=1)
        primer_dia = date(self.anio, self.mes, 1)
        desplazamiento = primer_dia.weekday()
        total_dias = calendar.monthrange(self.anio, self.mes)[1]
        fila, columna = 1, desplazamiento
        for dia in range(1, total_dias + 1):
            boton = ctk.CTkButton(grid, text=str(dia), width=36, height=26, font=("Arial", 10),
                                  fg_color="#e8eef6", text_color="#1a1a1a", hover_color="#c9d9ec",
                                  command=lambda d=dia: self.seleccionar(d))
            boton.grid(row=fila, column=columna, padx=1, pady=1)
            columna += 1
            if columna > 6:
                columna = 0
                fila += 1
        ctk.CTkButton(self.contenedor, text="Hoy", fg_color=COLOR_OK, hover_color="#1e8449",
                      command=self.hoy).pack(fill="x", pady=(8, 0))

    def cambiar_mes(self, cantidad):
        total = (self.anio * 12 + (self.mes - 1)) + cantidad
        self.anio, self.mes = total // 12, (total % 12) + 1
        self.dibujar()

    def hoy(self):
        fecha = datetime.now()
        self.anio, self.mes = fecha.year, fecha.month
        self.seleccionar(fecha.day)

    def seleccionar(self, dia):
        try:
            self.target_entry.delete(0, tk.END)
            self.target_entry.insert(0, "%02d/%02d/%04d" % (dia, self.mes, self.anio))
        except Exception:
            pass
        self.destroy()


def ruta_instructivo():
    """Ubica el instructivo en PDF (funciona en desarrollo y en la app compilada)."""
    nombre = "INSTRUCTIVO_MODULO_NOMINA.pdf"
    candidatas = []
    try:
        candidatas.append(os.path.join(sys._MEIPASS, nombre))
    except Exception:
        pass
    candidatas.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), nombre))
    try:
        candidatas.append(os.path.join(os.path.dirname(sys.executable), nombre))
    except Exception:
        pass
    try:
        candidatas.append(os.path.join(os.getcwd(), nombre))
    except Exception:
        pass
    for candidata in candidatas:
        try:
            if os.path.exists(candidata):
                return candidata
        except Exception:
            pass
    return ""


class VisorInstructivo(ctk.CTkToplevel):
    """Muestra el instructivo en PDF dentro del sistema, página por página.

    Si el sistema no puede dibujar el PDF (falta la librería), el módulo abre el
    archivo con el visor de PDF de Windows/Mac, que es el comportamiento de respaldo.
    """

    def __init__(self, ventana, ruta):
        super().__init__(ventana)
        import fitz  # PyMuPDF: si no está, el módulo abre el PDF por fuera
        from PIL import Image
        self._fitz = fitz
        self._Image = Image
        self.ruta = ruta
        self.documento = fitz.open(ruta)
        self.total = self.documento.page_count
        self.numero = 0
        self.zoom = 0.0          # 0 = ajustar al ancho de la ventana
        self._imagen = None
        # Búsqueda rápida dentro del instructivo
        self.consulta = ""
        self.coincidencias = []  # [(numero_de_pagina, rectangulo), ...]
        self.indice = 0

        self.title("❓ Ayuda — Instructivo del Módulo de Nómina y Asistencia")
        ancho, alto = 980, 780
        try:
            x = ventana.winfo_rootx() + max(0, (ventana.winfo_width() - ancho) // 2)
            y = ventana.winfo_rooty() + max(0, (ventana.winfo_height() - alto) // 3)
            self.geometry("%dx%d+%d+%d" % (ancho, alto, x, y))
        except Exception:
            self.geometry("%dx%d" % (ancho, alto))
        self.minsize(700, 520)

        cabecera = ctk.CTkFrame(self, fg_color=COLOR_PRIMARIO, corner_radius=0)
        cabecera.pack(fill="x")
        ctk.CTkLabel(cabecera, text="❓ INSTRUCTIVO DEL MÓDULO DE NÓMINA Y ASISTENCIA",
                     font=("Arial", 15, "bold"), text_color="white").pack(side="left", padx=16, pady=10)
        ctk.CTkButton(cabecera, text="🔍 Abrir con el visor del sistema", width=230, height=28,
                      font=("Arial", 11), fg_color="#16a085", hover_color="#11806a",
                      command=lambda: abrir_documento(self.ruta)).pack(side="right", padx=10, pady=8)

        barra = ctk.CTkFrame(self, fg_color="transparent")
        barra.pack(fill="x", padx=12, pady=(8, 4))
        ctk.CTkButton(barra, text="◀ Anterior", width=110, fg_color="#34495e", hover_color="#2c3e50",
                      command=lambda: self.mostrar_pagina(self.numero - 1)).pack(side="left")
        ctk.CTkButton(barra, text="Siguiente ▶", width=110, fg_color="#34495e", hover_color="#2c3e50",
                      command=lambda: self.mostrar_pagina(self.numero + 1)).pack(side="left", padx=6)
        ctk.CTkLabel(barra, text="Ir a la página:", font=("Arial", 11, "bold")).pack(side="left", padx=(14, 4))
        self.ent_pagina = ctk.CTkEntry(barra, width=60, justify="center")
        self.ent_pagina.pack(side="left")
        self.ent_pagina.bind("<Return>", lambda _e: self.ir_a_pagina())
        ctk.CTkButton(barra, text="Ir", width=50, fg_color=COLOR_PRIMARIO, hover_color=COLOR_HOVER,
                      command=self.ir_a_pagina).pack(side="left", padx=4)
        ctk.CTkButton(barra, text="➖", width=40, fg_color="#7f8c8d", hover_color="#606b6b",
                      command=lambda: self.cambiar_zoom(-0.2)).pack(side="right", padx=2)
        self.lbl_zoom = ctk.CTkLabel(barra, text="100%", font=("Arial", 11, "bold"), width=60)
        self.lbl_zoom.pack(side="right")
        ctk.CTkButton(barra, text="➕", width=40, fg_color="#7f8c8d", hover_color="#606b6b",
                      command=lambda: self.cambiar_zoom(0.2)).pack(side="right", padx=2)
        ctk.CTkButton(barra, text="Ajustar al ancho", width=140, fg_color="#7f8c8d",
                      hover_color="#606b6b", command=self.ajustar_ancho).pack(side="right", padx=8)

        # Barra de búsqueda rápida
        busqueda = ctk.CTkFrame(self, fg_color="#f4f7fa", corner_radius=8, border_width=1,
                                border_color="#dbe3ea")
        busqueda.pack(fill="x", padx=12, pady=(0, 4))
        ctk.CTkLabel(busqueda, text="🔍 Buscar:", font=("Arial", 11, "bold")).pack(side="left",
                                                                                  padx=(10, 4), pady=7)
        self.ent_buscar = ctk.CTkEntry(busqueda, width=240,
                                       placeholder_text="turno, tardanza, horas extra, sueldo...")
        self.ent_buscar.pack(side="left", pady=7)
        self.ent_buscar.bind("<Return>", lambda _e: self.buscar())
        ctk.CTkButton(busqueda, text="Buscar", width=80, height=26, fg_color=COLOR_PRIMARIO,
                      hover_color=COLOR_HOVER, command=self.buscar).pack(side="left", padx=6)
        ctk.CTkButton(busqueda, text="◀", width=38, height=26, fg_color="#34495e",
                      hover_color="#2c3e50", command=lambda: self.ir_a_coincidencia(-1)).pack(side="left")
        ctk.CTkButton(busqueda, text="▶", width=38, height=26, fg_color="#34495e",
                      hover_color="#2c3e50", command=lambda: self.ir_a_coincidencia(1)).pack(side="left", padx=(2, 6))
        ctk.CTkButton(busqueda, text="Limpiar", width=80, height=26, fg_color=COLOR_GRIS,
                      hover_color="#606b6b", command=self.limpiar_busqueda).pack(side="left")
        self.lbl_busqueda = ctk.CTkLabel(busqueda, text="Escriba una palabra y pulse Buscar (o Enter)",
                                         font=("Arial", 11, "bold"), text_color=COLOR_PRIMARIO)
        self.lbl_busqueda.pack(side="left", padx=10)

        self.area = ctk.CTkScrollableFrame(self, fg_color="#e9edf2", corner_radius=8)
        self.area.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        self.lbl_pagina = ctk.CTkLabel(self.area, text="Cargando...", font=("Arial", 12))
        self.lbl_pagina.pack(padx=10, pady=10)

        pie = ctk.CTkFrame(self, fg_color="transparent")
        pie.pack(fill="x", padx=12, pady=(0, 10))
        self.lbl_estado = ctk.CTkLabel(pie, text="", font=("Arial", 11, "bold"), text_color=COLOR_PRIMARIO)
        self.lbl_estado.pack(side="left")
        ctk.CTkLabel(pie, text="También puede usar las flechas del teclado (← →) y Esc para cerrar",
                     font=("Arial", 10), text_color=COLOR_GRIS).pack(side="right")

        # Atajos del visor (Esc cierra la ayuda sin afectar al resto del sistema)
        try:
            self.bind("<Left>", lambda _e: self.mostrar_pagina(self.numero - 1))
            self.bind("<Right>", lambda _e: self.mostrar_pagina(self.numero + 1))
            self.bind("<Prior>", lambda _e: self.mostrar_pagina(self.numero - 1))
            self.bind("<Next>", lambda _e: self.mostrar_pagina(self.numero + 1))
            self.bind("<Home>", lambda _e: self.mostrar_pagina(0))
            self.bind("<End>", lambda _e: self.mostrar_pagina(self.total - 1))
            self.bind("<plus>", lambda _e: self.cambiar_zoom(0.2))
            self.bind("<minus>", lambda _e: self.cambiar_zoom(-0.2))
            self.bind("<Escape>", lambda _e: self.destroy())
            self.bind("<Control-f>", lambda _e: self.ent_buscar.focus_set())
            self.bind("<Control-F>", lambda _e: self.ent_buscar.focus_set())
        except Exception:
            pass
        self.lift()
        self.focus_force()
        # La primera página se dibuja cuando la ventana ya tomó su tamaño definitivo
        # (así el "ajustar al ancho" usa el ancho real y no el inicial de 1 píxel).
        self.after(120, lambda: self.mostrar_pagina(0))

    # ---------- BÚSQUEDA RÁPIDA ----------
    def coincidencias_en_pagina(self, numero):
        """Rectángulos donde aparece la palabra buscada en una página.

        Compara sin distinguir mayúsculas ni acentos (escribir "nomina" encuentra "nómina").
        Si la búsqueda tiene varias palabras, deben aparecer juntas en ese orden.
        """
        if not self.consulta:
            return []
        try:
            palabras = self.documento.load_page(numero).get_text("words")
        except Exception:
            return []
        if not palabras:
            return []
        partes = nc.normalizar_texto(self.consulta).lower().split()
        if not partes:
            return []
        encontrados = []
        for inicio in range(len(palabras) - len(partes) + 1):
            ventana = [nc.normalizar_texto(palabras[inicio + salto][4]).lower()
                       for salto in range(len(partes))]
            if all(partes[salto] in ventana[salto] for salto in range(len(partes))):
                union = None
                for salto in range(len(partes)):
                    caja = self._fitz.Rect(palabras[inicio + salto][:4])
                    union = caja if union is None else (union | caja)
                encontrados.append(union)
        return encontrados

    def buscar(self):
        """Busca el texto en todo el instructivo y salta a la primera coincidencia."""
        termino = self.ent_buscar.get().strip()
        self.consulta = termino
        self.coincidencias = []
        self.indice = 0
        if not termino:
            self.lbl_busqueda.configure(text="Escriba una palabra para buscar", text_color=COLOR_ALERTA)
            self.mostrar_pagina(self.numero)
            return
        # La búsqueda recorre todo el instructivo: se avisa para que no parezca colgado
        try:
            self.lbl_busqueda.configure(text="⏳ Buscando «%s» en %d páginas..." % (termino, self.total),
                                        text_color=COLOR_ALERTA)
            self.update_idletasks()
        except Exception:
            pass
        for numero in range(self.total):
            for caja in self.coincidencias_en_pagina(numero):
                self.coincidencias.append((numero, caja))
        if not self.coincidencias:
            self.lbl_busqueda.configure(text="Sin coincidencias para «%s»" % termino,
                                        text_color=COLOR_ERROR)
            self.mostrar_pagina(self.numero)
            return
        self.mostrar_pagina(self.coincidencias[0][0])
        self.avisar_coincidencias()

    def avisar_coincidencias(self):
        """Muestra en qué coincidencia va el usuario."""
        if not self.coincidencias:
            return
        pagina = self.coincidencias[self.indice][0]
        en_pagina = [i for i, (numero, _caja) in enumerate(self.coincidencias) if numero == pagina]
        self.lbl_busqueda.configure(
            text="«%s»: coincidencia %d de %d  ·  página %d  ·  %d en esta página"
                 % (self.consulta, self.indice + 1, len(self.coincidencias), pagina + 1, len(en_pagina)),
            text_color=COLOR_PRIMARIO)

    def ir_a_coincidencia(self, paso):
        """Pasa a la coincidencia anterior o siguiente (recorre todo el instructivo)."""
        if not self.coincidencias:
            return self.buscar()
        self.indice = (self.indice + paso) % len(self.coincidencias)
        self.mostrar_pagina(self.coincidencias[self.indice][0])
        self.avisar_coincidencias()

    def limpiar_busqueda(self):
        """Quita el resaltado y limpia el buscador."""
        self.consulta = ""
        self.coincidencias = []
        self.indice = 0
        try:
            self.ent_buscar.delete(0, tk.END)
        except Exception:
            pass
        self.lbl_busqueda.configure(text="Escriba una palabra y pulse Buscar (o Enter)",
                                    text_color=COLOR_PRIMARIO)
        self.mostrar_pagina(self.numero)

    def ancho_visible(self):
        try:
            return max(420, self.area.winfo_width() - 24)
        except Exception:
            return 800

    def escala(self):
        if self.zoom:
            return self.zoom
        try:
            pagina = self.documento.load_page(self.numero)
            return max(0.4, min(3.0, self.ancho_visible() / float(pagina.rect.width)))
        except Exception:
            return 1.0

    def resaltar_coincidencias(self, imagen, escala):
        """Pinta de amarillo las coincidencias de la búsqueda (la actual, en naranja)."""
        try:
            cajas = self.coincidencias_en_pagina(self.numero)
        except Exception:
            return imagen
        if not cajas:
            return imagen
        actual = None
        if 0 <= self.indice < len(self.coincidencias):
            pagina_actual, caja_actual = self.coincidencias[self.indice]
            if pagina_actual == self.numero:
                actual = caja_actual
        try:
            from PIL import ImageDraw
            base = imagen.convert("RGBA")
            capa = self._Image.new("RGBA", base.size, (0, 0, 0, 0))
            dibujo = ImageDraw.Draw(capa)
            for caja in cajas:
                es_actual = (actual is not None and abs(caja.x0 - actual.x0) < 0.5
                             and abs(caja.y0 - actual.y0) < 0.5)
                relleno = (255, 165, 0, 115) if es_actual else (255, 214, 0, 90)
                borde = (200, 110, 0, 255) if es_actual else (225, 170, 0, 255)
                dibujo.rectangle([caja.x0 * escala, caja.y0 * escala,
                                  caja.x1 * escala, caja.y1 * escala],
                                 fill=relleno, outline=borde, width=2)
            return self._Image.alpha_composite(base, capa).convert("RGB")
        except Exception as e:
            print("[Nómina] No se pudo resaltar la búsqueda:", e)
            return imagen

    def mostrar_pagina(self, numero):
        """Dibuja la página indicada del instructivo."""
        if not self.documento or self.total == 0:
            return
        self.numero = max(0, min(self.total - 1, int(numero)))
        try:
            pagina = self.documento.load_page(self.numero)
            escala = self.escala()
            mapa = pagina.get_pixmap(matrix=self._fitz.Matrix(escala, escala), alpha=False)
            imagen = self._Image.frombytes("RGB", (mapa.width, mapa.height), mapa.samples)
            if self.consulta:
                imagen = self.resaltar_coincidencias(imagen, escala)
            self._imagen = ctk.CTkImage(light_image=imagen, dark_image=imagen,
                                        size=(mapa.width, mapa.height))
            self.lbl_pagina.configure(image=self._imagen, text="")
        except Exception as e:
            self.lbl_pagina.configure(image=None, text="No se pudo dibujar la página: %s" % e)
        try:
            self.ent_pagina.delete(0, tk.END)
            self.ent_pagina.insert(0, str(self.numero + 1))
            self.lbl_estado.configure(text="Página %d de %d" % (self.numero + 1, self.total))
            self.lbl_zoom.configure(text="%d%%" % int(round(self.escala() * 100)))
        except Exception:
            pass
        try:
            self.area._parent_canvas.yview_moveto(0.0)
        except Exception:
            pass

    def ir_a_pagina(self):
        try:
            self.mostrar_pagina(int(self.ent_pagina.get()) - 1)
        except Exception:
            self.mostrar_pagina(self.numero)

    def cambiar_zoom(self, delta):
        actual = self.escala() + delta
        self.zoom = max(0.4, min(3.0, actual))
        self.mostrar_pagina(self.numero)

    def ajustar_ancho(self):
        self.zoom = 0.0
        self.mostrar_pagina(self.numero)


# =========================================================
# LETRERO CENTRADO: "LEYENDO BASE DE DATOS / CALCULANDO"
# =========================================================
_LETRERO_PROFUNDIDAD = 0


class LetreroCarga:
    """Aviso centrado en la pantalla mientras el sistema trabaja.

    Es el mismo letrero del módulo de Cálculo de Cobranza: un recuadro blanco con el
    icono de reloj, el mensaje y una barra de progreso indefinida, colocado en el
    centro de la ventana mientras se lee la base de datos o se hacen cálculos.

    Uso:
        letrero = LetreroCarga(self.frame_main, "Leyendo base de datos...")
        try:
            ...trabajo pesado...
        finally:
            letrero.cerrar()

    También se puede usar el decorador @con_letrero("mensaje") sobre un método.
    """

    def __init__(self, widget, mensaje="Procesando...", detalle=""):
        global _LETRERO_PROFUNDIDAD
        self._raiz = None
        self._frame = None
        self._barra = None
        self._lbl = None
        self._activo = False
        # Si ya hay un letrero visible (llamada anidada) no se duplica
        self._contado = True
        self._anidado = _LETRERO_PROFUNDIDAD > 0
        _LETRERO_PROFUNDIDAD += 1
        if self._anidado:
            return
        try:
            self._raiz = widget.winfo_toplevel()
        except Exception:
            try:
                self._raiz = widget
            except Exception:
                self._raiz = None
        if self._raiz is None:
            return
        try:
            familia = "Helvetica" if sys.platform == "darwin" else "Arial"
            self._frame = ctk.CTkFrame(self._raiz, corner_radius=14, border_width=2,
                                       border_color=COLOR_PRIMARIO, fg_color="#ffffff")
            ctk.CTkLabel(self._frame, text="⏳", font=(familia, 26)).pack(pady=(16, 0))
            self._lbl = ctk.CTkLabel(self._frame, text=mensaje, font=(familia, 14, "bold"),
                                     text_color=COLOR_PRIMARIO, wraplength=320, justify="center")
            self._lbl.pack(padx=34, pady=(4, 2))
            self._detalle = ctk.CTkLabel(self._frame, text=detalle, font=(familia, 10),
                                         text_color=COLOR_GRIS, wraplength=320, justify="center")
            if detalle:
                self._detalle.pack(padx=34, pady=(0, 4))
            self._barra = ctk.CTkProgressBar(self._frame, mode="indeterminate", width=250, height=8,
                                             progress_color=COLOR_PRIMARIO)
            self._barra.pack(padx=34, pady=(6, 16))
            self._barra.start()
            self._frame.place(relx=0.5, rely=0.5, anchor="center")
            self._frame.lift()
            self._activo = True
            self._pintar()
        except Exception:
            self._activo = False

    def _pintar(self):
        """Fuerza el dibujado inmediato (para que se vea antes del trabajo pesado)."""
        try:
            self._raiz.update_idletasks()
        except Exception:
            pass

    def cambiar(self, mensaje, detalle=None):
        """Cambia el texto del letrero mientras sigue visible."""
        try:
            self._lbl.configure(text=mensaje)
            if detalle is not None:
                self._detalle.configure(text=detalle)
                if detalle and not self._detalle.winfo_ismapped():
                    self._detalle.pack(padx=34, pady=(0, 4), before=self._barra)
            self._pintar()
        except Exception:
            pass

    def cerrar(self):
        """Quita el letrero de la pantalla."""
        global _LETRERO_PROFUNDIDAD
        if getattr(self, "_contado", False):
            self._contado = False
            # Se descuenta SIEMPRE: cada instancia crea (suma) y cierra (resta) una sola vez.
            # Antes había un tope en cero y, si un letrero anidado cerraba después que el
            # principal, su resta se perdía y el letrero dejaba de aparecer para siempre.
            _LETRERO_PROFUNDIDAD -= 1
        if not self._activo:
            return
        try:
            if self._barra is not None:
                self._barra.stop()
        except Exception:
            pass
        try:
            if self._frame is not None:
                self._frame.destroy()
        except Exception:
            pass
        self._activo = False
        try:
            self._raiz.update_idletasks()
        except Exception:
            pass


def _padre_para_letrero(objeto):
    """Marco donde colocar el letrero: el del módulo, o la propia ventana si es un diálogo."""
    try:
        if hasattr(objeto, "winfo_toplevel") and objeto.winfo_exists():
            return objeto
    except Exception:
        pass
    for atributo in ("frame_main", "parent_frame", "parent"):
        try:
            valor = getattr(objeto, atributo, None)
            if valor is not None:
                return valor
        except Exception:
            pass
    try:
        app = getattr(objeto, "app", None)
        if app is not None:
            return getattr(app, "frame_main", None)
    except Exception:
        pass
    return None


def con_letrero(mensaje, detalle=""):
    """Decorador: muestra el letrero centrado mientras el método trabaja."""
    def decorador(func):
        @functools.wraps(func)
        def envoltura(self, *args, **kwargs):
            letrero = None
            try:
                padre = _padre_para_letrero(self)
                if padre is not None:
                    letrero = LetreroCarga(padre, mensaje, detalle)
            except Exception:
                letrero = None
            try:
                return func(self, *args, **kwargs)
            finally:
                if letrero is not None:
                    try:
                        letrero.cerrar()
                    except Exception:
                        pass
        return envoltura
    return decorador


class AsistenteNomina(ctk.CTkToplevel):
    """Asistente paso a paso: revisa el estado del módulo y guía al operario.

    Muestra una lista de pasos con su estado real (leído de la base de datos), explica
    qué falta y por qué importa, permite saltar a la pantalla donde se resuelve y, en
    los casos automáticos, ejecutar la tarea desde aquí mismo.
    """

    ICONOS = {"OK": "✅", "ATENCION": "⚠️", "PENDIENTE": "⛔"}
    COLORES = {"OK": COLOR_OK, "ATENCION": COLOR_ALERTA, "PENDIENTE": COLOR_ERROR}

    def __init__(self, app):
        super().__init__(app.parent_frame.winfo_toplevel())
        self.app = app
        self.title("🧭 Asistente de configuración de nómina")
        ancho, alto = 1010, 700
        ventana = app.parent_frame.winfo_toplevel()
        try:
            x = ventana.winfo_rootx() + max(0, (ventana.winfo_width() - ancho) // 2)
            y = ventana.winfo_rooty() + max(0, (ventana.winfo_height() - alto) // 3)
            self.geometry("%dx%d+%d+%d" % (ancho, alto, x, y))
        except Exception:
            self.geometry("%dx%d" % (ancho, alto))
        self.minsize(880, 560)
        try:
            self.transient(ventana)
            self.grab_set()
        except Exception:
            pass
        self.datos = None
        self.paso_actual = 0
        self._construir()
        self.cargar_estado()
        self.lift()
        self.focus_force()

    # ------------------------------------------------------------- interfaz
    def _construir(self):
        cabecera = ctk.CTkFrame(self, fg_color=COLOR_PRIMARIO, corner_radius=0)
        cabecera.pack(fill="x")
        ctk.CTkLabel(cabecera, text="🧭 ASISTENTE PASO A PASO", font=("Arial", 18, "bold"),
                     text_color="white").pack(anchor="w", padx=18, pady=(12, 0))
        ctk.CTkLabel(cabecera, text="Le indico qué falta para dejar el módulo listo y para cerrar "
                                    "la planilla del mes, en el orden correcto.",
                     font=("Arial", 11), text_color="#d6e6f5").pack(anchor="w", padx=18, pady=(0, 10))

        progreso = ctk.CTkFrame(self, fg_color="transparent")
        progreso.pack(fill="x", padx=18, pady=(10, 4))
        self.barra = ctk.CTkProgressBar(progreso, height=14, progress_color=COLOR_OK)
        self.barra.pack(fill="x")
        self.barra.set(0)
        self.lbl_progreso = ctk.CTkLabel(progreso, text="Revisando...", font=("Arial", 11, "bold"),
                                         text_color=COLOR_PRIMARIO)
        self.lbl_progreso.pack(anchor="w", pady=(4, 0))

        cuerpo = ctk.CTkFrame(self, fg_color="transparent")
        cuerpo.pack(fill="both", expand=True, padx=18, pady=6)

        izquierda = ctk.CTkFrame(cuerpo, width=330, fg_color=COLOR_FONDO_SUAVE, corner_radius=10,
                                 border_width=1, border_color="#e0e0e0")
        izquierda.pack(side="left", fill="y")
        izquierda.pack_propagate(False)
        ctk.CTkLabel(izquierda, text="PASOS", font=("Arial", 11, "bold"),
                     text_color=COLOR_GRIS).pack(anchor="w", padx=12, pady=(10, 4))
        self.lista_pasos = ctk.CTkScrollableFrame(izquierda, fg_color="transparent", width=300)
        self.lista_pasos.pack(fill="both", expand=True, padx=6, pady=(0, 8))

        self.panel_detalle = ctk.CTkScrollableFrame(cuerpo, fg_color="#ffffff", corner_radius=10,
                                                    border_width=1, border_color="#e0e0e0")
        self.panel_detalle.pack(side="left", fill="both", expand=True, padx=(12, 0))

        pie = ctk.CTkFrame(self, fg_color="transparent")
        pie.pack(fill="x", padx=18, pady=(4, 14))
        ctk.CTkButton(pie, text="◀ Paso anterior", width=150, fg_color="#34495e", hover_color="#2c3e50",
                      command=lambda: self.mover(-1)).pack(side="left")
        ctk.CTkButton(pie, text="Paso siguiente ▶", width=150, fg_color="#34495e", hover_color="#2c3e50",
                      command=lambda: self.mover(1)).pack(side="left", padx=6)
        ctk.CTkButton(pie, text="🔄 Verificar de nuevo", width=180, fg_color=COLOR_PRIMARIO,
                      hover_color=COLOR_HOVER, command=self.cargar_estado).pack(side="left", padx=6)
        ctk.CTkButton(pie, text="Cerrar", width=110, fg_color=COLOR_GRIS, hover_color="#606b6b",
                      command=self.destroy).pack(side="right")

    # ------------------------------------------------------------- estado
    def cargar_estado(self):
        """Lee el estado real del módulo (en segundo plano) y repinta el asistente."""
        self.lbl_progreso.configure(text="⏳ Revisando el estado del módulo...", text_color=COLOR_ALERTA)
        periodo = self.app.periodo_actual

        def tarea(estado):
            estado["datos"] = nc.estado_configuracion(periodo)

        def terminar(estado):
            if estado.get("error") or not estado.get("datos"):
                self.lbl_progreso.configure(text="❌ No se pudo revisar el estado: %s"
                                                 % estado.get("error"), text_color=COLOR_ERROR)
                return
            self.datos = estado["datos"]
            self.pintar()

        ejecutar_en_hilo(self, tarea, al_terminar=terminar)

    def pintar(self):
        datos = self.datos or {}
        pasos = datos.get("pasos") or []
        self.barra.set((datos.get("avance") or 0) / 100.0)
        self.barra.configure(progress_color=COLOR_OK if datos.get("listo") else COLOR_ALERTA)
        self.lbl_progreso.configure(
            text="%d%% completado · %d de %d pasos críticos listos · %d pendiente(s) y %d aviso(s)"
                 % (datos.get("avance", 0), datos.get("criticos_ok", 0), datos.get("criticos_total", 0),
                    len(datos.get("pendientes") or []), len(datos.get("atenciones") or [])),
            text_color=COLOR_OK if datos.get("listo") else COLOR_PRIMARIO)

        for widget in self.lista_pasos.winfo_children():
            widget.destroy()
        for indice, paso in enumerate(pasos):
            icono = self.ICONOS.get(paso["estado"], "•")
            boton = ctk.CTkButton(
                self.lista_pasos, text="%s  %d. %s" % (icono, indice + 1, paso["titulo"]),
                anchor="w", width=292, height=34, font=("Arial", 11),
                fg_color=self.COLORES.get(paso["estado"], COLOR_GRIS) if indice == self.paso_actual
                else "transparent",
                hover_color="#dfe6ec",
                text_color="white" if indice == self.paso_actual else "#2c3e50",
                command=lambda i=indice: self.seleccionar(i))
            boton.pack(fill="x", pady=2)

        siguiente = datos.get("siguiente")
        if siguiente and self.paso_actual >= len(pasos):
            self.paso_actual = 0
        if self.paso_actual >= len(pasos):
            self.paso_actual = max(0, len(pasos) - 1)
        self.pintar_detalle()

    def seleccionar(self, indice):
        self.paso_actual = indice
        self.pintar()

    def mover(self, cantidad):
        pasos = (self.datos or {}).get("pasos") or []
        if not pasos:
            return
        self.paso_actual = max(0, min(len(pasos) - 1, self.paso_actual + cantidad))
        self.pintar()

    # ------------------------------------------------------------- detalle
    def pintar_detalle(self):
        for widget in self.panel_detalle.winfo_children():
            widget.destroy()
        pasos = (self.datos or {}).get("pasos") or []
        if not pasos:
            ctk.CTkLabel(self.panel_detalle, text="Sin información.", font=("Arial", 12)).pack(pady=30)
            return
        paso = pasos[self.paso_actual]
        color = self.COLORES.get(paso["estado"], COLOR_GRIS)
        etiquetas = {"OK": "LISTO", "ATENCION": "REVISAR", "PENDIENTE": "PENDIENTE"}
        fase = paso.get("fase") or ""

        ctk.CTkLabel(self.panel_detalle, text="%s · %s" % (fase, "Paso crítico" if paso["critico"]
                                                           else "Recomendado"),
                     font=("Arial", 10, "bold"), text_color=COLOR_GRIS).pack(anchor="w", padx=16,
                                                                            pady=(14, 0))
        ctk.CTkLabel(self.panel_detalle, text="%s %s" % (self.ICONOS.get(paso["estado"], ""),
                                                         paso["titulo"]),
                     font=("Arial", 17, "bold"), text_color=color).pack(anchor="w", padx=16, pady=(2, 4))
        insignia = ctk.CTkFrame(self.panel_detalle, fg_color=color, corner_radius=6)
        insignia.pack(anchor="w", padx=16, pady=(0, 8))
        ctk.CTkLabel(insignia, text=etiquetas.get(paso["estado"], paso["estado"]), font=("Arial", 10, "bold"),
                     text_color="white").pack(padx=10, pady=2)
        ctk.CTkLabel(self.panel_detalle, text=paso["detalle"], font=("Arial", 12), justify="left",
                     anchor="w", wraplength=560).pack(anchor="w", padx=16, pady=(0, 10))

        caja = ctk.CTkFrame(self.panel_detalle, fg_color="#eaf2fb", corner_radius=8)
        caja.pack(fill="x", padx=16, pady=(0, 10))
        ctk.CTkLabel(caja, text="💡 Qué hacer", font=("Arial", 11, "bold"),
                     text_color=COLOR_PRIMARIO).pack(anchor="w", padx=12, pady=(8, 2))
        ctk.CTkLabel(caja, text=paso["ayuda"], font=("Arial", 11), justify="left", anchor="w",
                     wraplength=540).pack(anchor="w", padx=12, pady=(0, 10))

        botones = ctk.CTkFrame(self.panel_detalle, fg_color="transparent")
        botones.pack(fill="x", padx=16, pady=(4, 16))
        pestana, subpestana = paso.get("destino") or ("", "")
        if pestana:
            ctk.CTkButton(botones, text="➡️ Ir a esa pantalla", width=200, fg_color=COLOR_PRIMARIO,
                          hover_color=COLOR_HOVER,
                          command=lambda p=pestana, s=subpestana: self.ir_y_cerrar(p, s)
                          ).pack(side="left")
        accion = paso.get("accion")
        if accion and paso["estado"] != "OK":
            etiqueta = self.ETIQUETAS_ACCION.get(accion)
            if etiqueta:
                ctk.CTkButton(botones, text="⚡ %s" % etiqueta, width=280, fg_color=COLOR_OK,
                              hover_color="#1e8449",
                              command=lambda a=accion: self.ejecutar_accion(a)).pack(side="left", padx=8)
        if paso["estado"] == "OK":
            ctk.CTkLabel(botones, text="Este paso ya está resuelto.", font=("Arial", 11, "italic"),
                         text_color=COLOR_OK).pack(side="left", padx=6)

    def ir_y_cerrar(self, pestana, subpestana):
        self.app.ir_a_pestana(pestana, subpestana)
        self.destroy()

    # ------------------------------------------------------------- acciones
    ETIQUETAS_ACCION = {
        "sincronizar": "Crear las fichas que faltan",
        "feriados": "Cargar los feriados del año",
        "reparar": "Reparar los parámetros dañados",
        "recalcular": "Recalcular la asistencia del mes",
        "planilla": "Calcular la planilla del mes",
    }

    def ejecutar_accion(self, accion):
        """Ejecuta la tarea del paso actual, pidiendo confirmación cuando modifica datos."""
        try:
            if accion == "sincronizar":
                if not messagebox.askyesno("Confirmar", "¿Crear las fichas de nómina que faltan para "
                                                        "los empleados registrados en Choferes?",
                                           parent=self):
                    return
                cantidad = nc.sincronizar_empleados(self.app.usuario_activo)
                self.app.recargar_catalogos()
                self.after(1500, self.cargar_estado)
                messagebox.showinfo("Listo", "Padrón sincronizado.", parent=self)
            elif accion == "feriados":
                anio = int(str(self.app.periodo_actual).split("-")[0])
                if not messagebox.askyesno("Confirmar", "¿Cargar los feriados nacionales del Perú "
                                                        "del año %d en el calendario?" % anio, parent=self):
                    return
                total = nc.sembrar_feriados_peru(anio, self.app.usuario_activo)
                self.after(1200, self.cargar_estado)
                messagebox.showinfo("Listo", "Se registraron %d feriados del año %d." % (total, anio),
                                    parent=self)
            elif accion == "reparar":
                detalle = nc.reparar_parametros(self.app.usuario_activo)
                self.app.parametros = nc.obtener_parametros()
                self.app.cargar_parametros()
                self.app.cargar_panel()
                self.cargar_estado()
                if detalle:
                    texto = "\n".join("• %s: %s  →  %s" % (c, repr(a), n) for c, a, n in detalle)
                    messagebox.showinfo("Parámetros reparados",
                                        "Se restauraron %d valores por defecto:\n\n%s" % (len(detalle), texto),
                                        parent=self)
                else:
                    messagebox.showinfo("Parámetros", "No hubo nada que reparar.", parent=self)
            elif accion == "recalcular":
                self.destroy()
                self.app.ir_a_pestana("Asistencia", "Matriz mensual")
                self.app.recalcular_periodo_actual()
            elif accion == "planilla":
                self.destroy()
                self.app.ir_a_pestana("Planilla", "Resumen del período")
                self.app.calcular_planilla()
        except Exception as e:
            messagebox.showerror("Error", "No se pudo ejecutar la acción:\n%s" % e, parent=self)


# =========================================================
# CARGA MASIVA DE SUELDOS Y GARANTIAS
# =========================================================
class DialogoSueldos(ctk.CTkToplevel):
    """Carga el sueldo, la garantia y la jornada de varios empleados a la vez.

    Evita tener que abrir la ficha de cada chofer: se marcan los empleados de la
    lista, se escriben los montos y se aplican todos de una sola vez. Los campos
    que se dejan en blanco NO se modifican.
    """

    def __init__(self, parent, app, empleados, usuario_activo=""):
        super().__init__(parent)
        self.app = app
        self.usuario = usuario_activo
        self.empleados = [e for e in (empleados or [])
                          if nc.normalizar_texto(e.get("estado")) != "CESADO"]
        self.variables = {}
        self.title("💵 Sueldos y garantías del personal")
        self.geometry("900x580")
        self.minsize(820, 500)
        try:
            self.transient(parent)
            self.grab_set()
        except Exception:
            pass
        self._construir()
        self.lift()
        self.focus_force()

    # ------------------------------------------------------------------ dibujo
    def _construir(self):
        ctk.CTkLabel(self, text="💵 SUELDOS Y GARANTÍAS DEL PERSONAL", font=("Arial", 15, "bold"),
                     text_color=COLOR_PRIMARIO).pack(anchor="w", padx=14, pady=(12, 2))
        ctk.CTkLabel(self, text="Marque los empleados, escriba los montos y pulse APLICAR. "
                                "Los campos que deje en blanco se conservan tal como están.",
                     font=("Arial", 11), text_color=COLOR_GRIS,
                     justify="left").pack(anchor="w", padx=14)

        cuerpo = ctk.CTkFrame(self, fg_color="transparent")
        cuerpo.pack(fill="both", expand=True, padx=12, pady=8)

        # ---- lista de empleados (izquierda) ----
        izquierda = ctk.CTkFrame(cuerpo, fg_color=COLOR_FONDO_SUAVE, corner_radius=10,
                                 border_width=1, border_color="#e0e0e0")
        izquierda.pack(side="left", fill="both", expand=True)
        ctk.CTkLabel(izquierda, text="👥 ¿A quiénes?", font=("Arial", 12, "bold"),
                     text_color=COLOR_PRIMARIO).pack(anchor="w", padx=10, pady=(8, 2))
        botones_lista = ctk.CTkFrame(izquierda, fg_color="transparent")
        botones_lista.pack(fill="x", padx=10)
        ctk.CTkButton(botones_lista, text="Marcar todos", width=120, fg_color=COLOR_PRIMARIO,
                      hover_color=COLOR_HOVER,
                      command=lambda: self._marcar(True)).pack(side="left", padx=(0, 6))
        ctk.CTkButton(botones_lista, text="Ninguno", width=90, fg_color=COLOR_GRIS,
                      hover_color="#606b6b",
                      command=lambda: self._marcar(False)).pack(side="left")
        self.lbl_seleccion = ctk.CTkLabel(botones_lista, text="", font=("Arial", 10),
                                          text_color=COLOR_GRIS)
        self.lbl_seleccion.pack(side="right")
        lista = ctk.CTkScrollableFrame(izquierda, fg_color="transparent")
        lista.pack(fill="both", expand=True, padx=6, pady=6)
        if not self.empleados:
            ctk.CTkLabel(lista, text="No hay empleados activos en el padrón.",
                         font=("Arial", 11)).pack(anchor="w", padx=6, pady=6)
        for empleado in self.empleados:
            dni = str(empleado["dni"])
            variable = tk.BooleanVar(value=True)
            self.variables[dni] = variable
            actual = nc.a_float(empleado.get("sueldo_basico"))
            garantia = nc.a_float(empleado.get("garantia_mensual"))
            detalle = "S/ %s" % formatear_numero(actual, 2)
            if garantia > 0:
                detalle += " · garantía S/ %s" % formatear_numero(garantia, 2)
            ctk.CTkCheckBox(lista, text="%s  (%s)  ·  hoy: %s"
                                       % (empleado.get("nombre") or dni, dni, detalle),
                            variable=variable, font=("Arial", 11),
                            command=self._actualizar_contador).pack(anchor="w", padx=6, pady=3)

        # ---- montos (derecha) ----
        derecha = ctk.CTkFrame(cuerpo, fg_color=COLOR_FONDO_SUAVE, corner_radius=10,
                               border_width=1, border_color="#e0e0e0", width=340)
        derecha.pack(side="right", fill="y", padx=(10, 0))
        derecha.pack_propagate(False)
        ctk.CTkLabel(derecha, text="💵 ¿Cuánto?", font=("Arial", 12, "bold"),
                     text_color=COLOR_PRIMARIO).pack(anchor="w", padx=12, pady=(10, 4))

        def campo(etiqueta, ayuda=""):
            ctk.CTkLabel(derecha, text=etiqueta, font=("Arial", 11, "bold"),
                         anchor="w").pack(anchor="w", padx=12, pady=(8, 0))
            entrada = ctk.CTkEntry(derecha, width=300)
            entrada.pack(anchor="w", padx=12)
            if ayuda:
                ctk.CTkLabel(derecha, text=ayuda, font=("Arial", 9), text_color=COLOR_GRIS,
                             justify="left", wraplength=290).pack(anchor="w", padx=12)
            return entrada

        self.ent_sueldo = campo("Sueldo básico (S/)",
                                "Sueldo mensual por las 8 horas de ley. La RMV vigente es S/ %s."
                                % formatear_numero(nc.a_float(self.app.parametros.get("rmv"), 1130), 2))
        self.ent_garantia = campo("Garantía mensual (S/)",
                                  "Monto mínimo garantizado. Sugerencia: turno de 12 h → 2,000 · "
                                  "turno de 14 h → 2,300. Deje 0 si no tiene garantía.")
        self.ent_jornada = campo("Jornada (horas)",
                                 "Horas de la jornada legal. Normalmente 8.")

        ctk.CTkLabel(derecha, text="Sistema de pensión", font=("Arial", 11, "bold"),
                     anchor="w").pack(anchor="w", padx=12, pady=(10, 0))
        self.cmb_pension = ctk.CTkComboBox(derecha, values=["(No cambiar)", "ONP", "AFP", "NINGUNO"],
                                           width=300, state="readonly")
        self.cmb_pension.set("(No cambiar)")
        self.cmb_pension.pack(anchor="w", padx=12)
        self.var_asignacion = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(derecha, text="Marcar asignación familiar", variable=self.var_asignacion,
                        font=("Arial", 11)).pack(anchor="w", padx=12, pady=(10, 2))
        botones_rapidos = ctk.CTkFrame(derecha, fg_color="transparent")
        botones_rapidos.pack(fill="x", padx=12, pady=(6, 0))
        ctk.CTkButton(botones_rapidos, text="Sueldo = RMV", width=140, fg_color="#34495e",
                      hover_color="#2c3e50", command=self._poner_rmv).pack(side="left", padx=(0, 6))
        ctk.CTkButton(botones_rapidos, text="Jornada = 8 h", width=140, fg_color="#34495e",
                      hover_color="#2c3e50",
                      command=lambda: self._escribir(self.ent_jornada, "8")).pack(side="left")

        # ---- acciones ----
        acciones = ctk.CTkFrame(self, fg_color="transparent")
        acciones.pack(fill="x", padx=14, pady=(0, 12))
        ctk.CTkButton(acciones, text="💾 APLICAR A LOS SELECCIONADOS", width=300, height=36,
                      font=("Arial", 12, "bold"), fg_color=COLOR_OK, hover_color="#1e8449",
                      command=self.aplicar).pack(side="left")
        ctk.CTkButton(acciones, text="Cerrar", width=110, height=36, fg_color=COLOR_GRIS,
                      hover_color="#606b6b", command=self.destroy).pack(side="right")
        ctk.CTkLabel(acciones, text="Recuerde recalcular la planilla después de cambiar los sueldos.",
                     font=("Arial", 10), text_color=COLOR_GRIS).pack(side="left", padx=12)
        self._actualizar_contador()

    # ---------------------------------------------------------------- ayudas
    def _escribir(self, entrada, valor):
        entrada.delete(0, tk.END)
        entrada.insert(0, valor)

    def _poner_rmv(self):
        """Escribe la RMV vigente sin separador de miles (1130, no 1,130)."""
        self._escribir(self.ent_sueldo, "%g" % (nc.a_float(self.app.parametros.get("rmv"), 1130) or 1130))

    def _marcar(self, valor):
        for variable in self.variables.values():
            variable.set(valor)
        self._actualizar_contador()

    def _actualizar_contador(self):
        try:
            marcados = len([1 for v in self.variables.values() if v.get()])
            self.lbl_seleccion.configure(text="%d de %d marcados"
                                              % (marcados, len(self.variables)))
        except Exception:
            pass

    def _leer(self, entrada, etiqueta):
        """Devuelve None si el campo quedó vacío; lanza ValueError si está mal escrito."""
        texto = entrada.get().strip()
        if not texto:
            return None
        numero = parse_num(texto, None)
        if numero is None:
            raise ValueError("«%s» no es un número válido: %s" % (etiqueta, texto))
        return numero

    # --------------------------------------------------------------- aplicar
    def aplicar(self):
        dnis = [dni for dni, variable in self.variables.items() if variable.get()]
        if not dnis:
            return messagebox.showwarning("Atención", "Marque al menos un empleado de la lista.",
                                          parent=self)
        try:
            sueldo = self._leer(self.ent_sueldo, "Sueldo básico")
            garantia = self._leer(self.ent_garantia, "Garantía mensual")
            jornada = self._leer(self.ent_jornada, "Jornada")
        except ValueError as e:
            return messagebox.showerror("Revisar los datos", str(e), parent=self)
        if sueldo is None and garantia is None and jornada is None \
                and self.cmb_pension.get() == "(No cambiar)" and not self.var_asignacion.get():
            return messagebox.showwarning("Atención",
                                          "Escriba al menos un monto (sueldo, garantía o jornada).",
                                          parent=self)
        rmv = nc.a_float(self.app.parametros.get("rmv"), 1130) or 1130
        if sueldo is not None and 0 < sueldo < rmv * 0.5:
            if not messagebox.askyesno(
                    "Revise el sueldo",
                    "El sueldo escrito es %s, muy por debajo de la RMV (%s).\n\n"
                    "¿Está seguro de que no le falta un cero?" % (formatear_monto(sueldo),
                                                                 formatear_monto(rmv)), parent=self):
                return
        pension = self.cmb_pension.get()
        pension = None if pension == "(No cambiar)" else pension
        asignacion = True if self.var_asignacion.get() else None
        if not messagebox.askyesno(
                "Confirmar",
                "¿Aplicar estos valores a %d empleado(s)?\n\nSueldo: %s\nGarantía: %s\n"
                "Jornada: %s\nPensión: %s\n\nLos campos en blanco no se modifican."
                % (len(dnis),
                   "sin cambio" if sueldo is None else formatear_monto(sueldo),
                   "sin cambio" if garantia is None else formatear_monto(garantia),
                   "sin cambio" if jornada is None else ("%.2f h" % jornada),
                   pension or "sin cambio"), parent=self):
            return
        ok, mensaje = nc.guardar_sueldos_lote(dnis, sueldo_basico=sueldo, garantia_mensual=garantia,
                                              jornada_horas=jornada, sistema_pension=pension,
                                              asignacion_familiar=asignacion, usuario=self.usuario)
        if not ok:
            return messagebox.showerror("Error", mensaje, parent=self)
        messagebox.showinfo("Sueldos actualizados", mensaje + "\n\nRecalcule la planilla del mes "
                                                            "para ver el resultado.", parent=self)
        try:
            self.app.recargar_catalogos()
        except Exception as e:
            print("[Nomina] No se pudo refrescar el padron:", e)
        self.destroy()


# =========================================================
# MÓDULO PRINCIPAL
# =========================================================
class ModuloNominaApp:
    """Módulo de nómina y control de asistencia."""

    def __init__(self, parent_frame, usuario_activo=""):
        self.parent_frame = parent_frame
        self.usuario_activo = usuario_activo or "Desconocido"
        aplicar_estilo_treeview()

        # Cachés en memoria (se refrescan en segundo plano)
        self.empleados = []
        self.turnos = []
        self.horarios = []
        self.parametros = {clave: valor for clave, (valor, _desc) in nc.PARAMETROS_DEFECTO.items()}
        self.periodo_actual = datetime.now().strftime("%Y-%m")
        self.archivo_actual = ""
        self.lectura_actual = None
        self.combos_mapeo = {}

        self.frame_main = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.frame_main.pack(fill="both", expand=True, padx=10, pady=10)
        # Armar la pantalla toma unos segundos: se avisa con el mismo letrero centrado
        letrero_arranque = LetreroCarga(self.frame_main, "Preparando la pantalla...",
                                        "Armando las pestañas del módulo")

        cabecera = ctk.CTkFrame(self.frame_main, fg_color="transparent")
        cabecera.pack(fill="x")
        ctk.CTkLabel(cabecera, text="🧾 MÓDULO DE NÓMINA Y ASISTENCIA", font=("Arial", 18, "bold"),
                     text_color=COLOR_PRIMARIO).pack(side="left")
        # Pantalla completa (F11): la matriz de asistencia y las tablas anchas se leen mejor
        self.btn_pantalla = ctk.CTkButton(cabecera, text="⛶ Pantalla completa", width=190, height=30,
                                          font=("Arial", 11, "bold"), fg_color="#34495e",
                                          hover_color="#2c3e50", command=self.alternar_pantalla_completa)
        # La marca permite que la barra lateral y este botón se sincronicen entre sí
        self.btn_pantalla._boton_pantalla_completa = True
        self.btn_pantalla.pack(side="right", padx=(10, 0))
        # Ayuda: abre el instructivo del módulo en PDF
        self.btn_ayuda = ctk.CTkButton(cabecera, text="❓ Ayuda", width=110, height=30,
                                       font=("Arial", 11, "bold"), fg_color="#16a085",
                                       hover_color="#11806a", command=self.abrir_ayuda)
        self.btn_ayuda.pack(side="right", padx=(8, 0))
        self.lbl_estado = ctk.CTkLabel(cabecera, text="Iniciando...", font=("Arial", 11),
                                       text_color=COLOR_GRIS)
        self.lbl_estado.pack(side="right")

        self.tabview = ctk.CTkTabview(self.frame_main, segmented_button_selected_color=COLOR_PRIMARIO,
                                      segmented_button_selected_hover_color=COLOR_HOVER)
        self.tabview.pack(fill="both", expand=True, pady=(8, 0))
        # Las pestañas siguen EL MISMO ORDEN que los pasos del asistente, para que el
        # operario avance de izquierda a derecha sin tener que pensar en qué toca:
        #   Personal (pasos 1 a 5) -> Configuración (6 y 7) -> Marcaciones (8 y 9)
        #   -> Asistencia (10) -> Planilla (11) -> Reportes (consultas)
        self.tab_panel = self.tabview.add(" 📊 Panel ")
        self.tab_personal = self.tabview.add(" 👥 Personal ")
        self.tab_config = self.tabview.add(" ⚙️ Configuración ")
        self.tab_marcaciones = self.tabview.add(" 📥 Marcaciones ")
        self.tab_asistencia = self.tabview.add(" 🗓️ Asistencia ")
        self.tab_planilla = self.tabview.add(" 💵 Planilla ")
        self.tab_reportes = self.tabview.add(" 📈 Reportes ")

        try:
            for nombre, constructor in (
                    ("Panel", self.construir_tab_panel),
                    ("Personal: empleados, turnos y horarios", self.construir_tab_personal),
                    ("Configuración: parámetros y feriados", self.construir_tab_configuracion),
                    ("Marcaciones del reloj", self.construir_tab_marcaciones),
                    ("Asistencia", self.construir_tab_asistencia),
                    ("Planilla", self.construir_tab_planilla),
                    ("Reportes", self.construir_tab_reportes)):
                letrero_arranque.cambiar("Preparando la pantalla...", nombre)
                constructor()
        finally:
            letrero_arranque.cerrar()

        self.actualizar_boton_pantalla()
        self.cargar_datos_iniciales()

    # =====================================================
    # UTILIDADES INTERNAS
    # =====================================================
    def aviso(self, texto, color=COLOR_OK):
        """Muestra un mensaje breve en la barra superior del módulo."""
        try:
            self.lbl_estado.configure(text=texto, text_color=color)
        except Exception:
            pass

    # ---------- PANTALLA COMPLETA ----------
    def ventana_principal(self):
        return self.parent_frame.winfo_toplevel()

    def pantalla_completa_activa(self):
        """El estado vive en la ventana, así se comparte con el botón de la barra lateral."""
        try:
            return bool(getattr(self.ventana_principal(), "_pantalla_completa_activa", False))
        except Exception:
            return False

    def actualizar_boton_pantalla(self):
        """Alinea el botón del módulo con el de la barra lateral (mismo estado)."""
        activa = self.pantalla_completa_activa()
        texto = "🗗 Salir de pantalla completa" if activa else "⛶ Pantalla completa"
        color = COLOR_ALERTA if activa else "#34495e"
        hover = "#b9640f" if activa else "#2c3e50"
        # Se recorren los widgets de la ventana para actualizar también el botón de la
        # barra lateral, así los dos nunca muestran estados distintos.
        pendientes = [self.ventana_principal()]
        while pendientes:
            widget = pendientes.pop()
            try:
                if getattr(widget, "_boton_pantalla_completa", False):
                    widget.configure(text=texto, fg_color=color, hover_color=hover)
                pendientes.extend(widget.winfo_children())
            except Exception:
                pass

    def alternar_pantalla_completa(self, evento=None):
        """Alterna entre pantalla completa y ventana maximizada (también con F11 o Esc)."""
        ventana = self.ventana_principal()
        activa = self.pantalla_completa_activa()
        try:
            if activa:
                try:
                    ventana.attributes("-fullscreen", False)
                except Exception:
                    pass
                geometria = getattr(ventana, "_geometria_previa", "")
                if geometria:
                    try:
                        ventana.geometry(geometria)
                    except Exception:
                        pass
                try:  # se vuelve al estado normal de trabajo: maximizada
                    if sys.platform == "win32":
                        ventana.state("zoomed")
                    elif sys.platform == "darwin":
                        ventana.geometry("%dx%d+0+0" % (ventana.winfo_screenwidth(),
                                                        ventana.winfo_screenheight() - 30))
                    else:
                        ventana.attributes("-zoomed", True)
                except Exception:
                    pass
                ventana._pantalla_completa_activa = False
                self.aviso("🖥️ Ventana normal restaurada", COLOR_GRIS)
            else:
                try:
                    ventana._geometria_previa = ventana.geometry()
                except Exception:
                    ventana._geometria_previa = ""
                try:
                    ventana.attributes("-fullscreen", True)
                except Exception:
                    # Respaldo si el sistema no acepta -fullscreen
                    ventana.geometry("%dx%d+0+0" % (ventana.winfo_screenwidth(),
                                                    ventana.winfo_screenheight()))
                ventana._pantalla_completa_activa = True
                self.aviso("🖥️ Pantalla completa activada (F11 o Esc para salir)", COLOR_OK)
        except Exception as e:
            print("[Nómina] Pantalla completa:", e)
            return
        self.actualizar_boton_pantalla()
        # Ocultar los botones del sistema para aprovechar todo el ancho de la pantalla
        self.alternar_barra_lateral_sistema(self.pantalla_completa_activa())

    def salir_pantalla_completa(self, evento=None):
        """Sale de pantalla completa con Esc (solo actúa si estaba activa)."""
        if self.pantalla_completa_activa():
            self.alternar_pantalla_completa()

    def alternar_barra_lateral_sistema(self, ocultar):
        """Pide al sistema principal que oculte o muestre la barra del menú lateral.

        En pantalla completa se quitan los botones del sistema para que el módulo use
        todo el ancho; el sistema deja un botón flotante para poder salir.
        """
        try:
            accion = getattr(self.ventana_principal(), "_alternar_barra_lateral", None)
            if callable(accion):
                accion(bool(ocultar))
        except Exception as e:
            print("[Nómina] No se pudo cambiar la barra lateral:", e)

    # ---------- AYUDA (INSTRUCTIVO EN PDF) ----------
    def abrir_ayuda(self):
        """Abre el instructivo del módulo (INSTRUCTIVO_MODULO_NOMINA.pdf).

        Si el sistema puede dibujar PDF se muestra dentro del propio sistema; si no,
        se abre con el visor de PDF de Windows o Mac.
        """
        ruta = ruta_instructivo()
        if not ruta:
            return messagebox.showwarning(
                "Instructivo no encontrado",
                "No se encontró el archivo INSTRUCTIVO_MODULO_NOMINA.pdf.\n\n"
                "Debe estar en la misma carpeta del sistema.\n"
                "Si hace falta, genérelo con:   python generar_instructivo_pdf.py")
        # Si la ayuda ya está abierta, se trae al frente en lugar de duplicarla
        try:
            if getattr(self, "_visor", None) is not None and self._visor.winfo_exists():
                self._visor.lift()
                self._visor.focus_force()
                return
        except Exception:
            pass
        try:
            import fitz  # noqa: F401  (PyMuPDF: permite dibujar el PDF dentro del sistema)
            self._visor = VisorInstructivo(self.parent_frame.winfo_toplevel(), ruta)
        except Exception as e:
            print("[Nómina] El instructivo se abrirá con el visor del sistema:", e)
            abrir_documento(ruta)

    # ---------- ASISTENTE PASO A PASO ----------
    def abrir_asistente(self):
        """Abre el asistente que revisa el estado del módulo y guía al operario."""
        try:
            if getattr(self, "_asistente", None) is not None and self._asistente.winfo_exists():
                self._asistente.lift()
                self._asistente.focus_force()
                self._asistente.cargar_estado()
                return
        except Exception:
            pass
        try:
            self._asistente = AsistenteNomina(self)
        except Exception as e:
            print("[Nómina] Asistente:", e)
            messagebox.showerror("Error", "No se pudo abrir el asistente:\n%s" % e)

    def ir_a_pestana(self, pestana, subpestana=None):
        """Lleva al usuario a la pestaña (y vista interna) indicada, sin depender de acentos."""
        destinos = {
            "panel": (" 📊 Panel ", None),
            "marcaciones": (" 📥 Marcaciones ", "marcaciones"),
            "asistencia": (" 🗓️ Asistencia ", "asistencia"),
            "personal": (" 👥 Personal ", "personal"),
            "planilla": (" 💵 Planilla ", "planilla"),
            "reportes": (" 📈 Reportes ", None),
            "configuracion": (" ⚙️ Configuración ", "configuracion"),
        }
        destino = destinos.get(nc.normalizar_texto(pestana).lower())
        if not destino:
            return False
        titulo, grupo = destino
        try:
            self.tabview.set(titulo)
        except Exception as e:
            print("[Nómina] No se pudo abrir la pestaña %s: %s" % (titulo, e))
            return False
        if subpestana and grupo:
            widget = {"marcaciones": self.sub_marcaciones, "asistencia": self.sub_asistencia,
                      "personal": self.sub_personal, "planilla": self.sub_planilla,
                      "configuracion": self.sub_config}.get(grupo)
            if widget is not None:
                try:
                    valores = list(widget.cget("values"))
                except Exception:
                    valores = list(getattr(widget, "_values", []) or [])
                objetivo = nc.normalizar_texto(subpestana).lower()
                for valor in valores:
                    if nc.normalizar_texto(valor).lower() == objetivo:
                        widget.set(valor)
                        self._cambiar_vista_interna(titulo, valor)
                        break
        return True

    def _cambiar_vista_interna(self, titulo_pestana, valor):
        """Ejecuta el cambio de vista interna de la pestaña activa."""
        try:
            if titulo_pestana == " 📥 Marcaciones ":
                self.cambiar_sub_marcaciones(valor)
            elif titulo_pestana == " 🗓️ Asistencia ":
                self.cambiar_sub_asistencia(valor)
            elif titulo_pestana == " 👥 Personal ":
                self.cambiar_sub_personal(valor)
            elif titulo_pestana == " 💵 Planilla ":
                self.cambiar_sub_planilla(valor)
            elif titulo_pestana == " ⚙️ Configuración ":
                self.cambiar_sub_config(valor)
        except Exception as e:
            print("[Nómina] No se pudo cambiar la vista interna:", e)

    def opciones_empleados(self, incluir_todos=False):
        """Lista 'DNI | NOMBRE' para los combos."""
        opciones = []
        for empleado in self.empleados:
            opciones.append("%s | %s" % (empleado["dni"], empleado["nombre"]))
        if incluir_todos:
            opciones.insert(0, "TODOS | -- Todos los empleados --")
        return opciones or ["(sin empleados)"]

    def dni_de_opcion(self, texto):
        if not texto:
            return None
        return str(texto).split("|")[0].strip()

    def nombre_empleado(self, dni):
        for empleado in self.empleados:
            if str(empleado["dni"]) == str(dni):
                return empleado["nombre"]
        return str(dni)

    def abrir_calendario(self, entry_objetivo):
        CalendarioNativo(self.parent_frame.winfo_toplevel(), entry_objetivo)

    def cargar_datos_iniciales(self):
        """Carga catálogos en segundo plano para que el módulo abra al instante."""
        self.aviso("⏳ Cargando catálogos...", COLOR_ALERTA)
        letrero = LetreroCarga(self.frame_main, "Leyendo la base de datos...",
                               "Preparando el módulo de nómina")

        def tarea(estado):
            estado["esquema"] = nc.inicializar_esquema_nomina()
            estado["empleados"] = nc.listar_empleados()
            estado["turnos"] = nc.listar_turnos()
            estado["horarios"] = nc.listar_horarios()
            estado["parametros"] = nc.obtener_parametros()

        def terminar(estado):
            # El letrero se cierra en el bloque finally: así nunca queda pegado en pantalla,
            # ni siquiera si alguna lectura falla. Mientras está visible, las lecturas
            # internas (turnos, horarios, feriados...) no abren otro letrero ni parpadean.
            try:
                if estado.get("error"):
                    self.aviso("❌ Error cargando catálogos: %s" % estado["error"], COLOR_ERROR)
                    messagebox.showerror("Error", "No se pudo inicializar el módulo de nómina:\n%s"
                                         % estado["error"])
                    return
                self.empleados = estado.get("empleados") or []
                self.turnos = estado.get("turnos") or []
                self.horarios = estado.get("horarios") or []
                self.parametros = estado.get("parametros") or self.parametros
                self.refrescar_combos_globales()
                self.cargar_panel()
                self.cargar_historial_importaciones()
                self.cargar_personas_reloj()
                self.cargar_mapeos()
                self.cargar_turnos()
                self.cargar_horarios()
                self.cargar_parametros()
                self.cargar_feriados()
                self.aviso("✅ Módulo listo. Empleados: %d" % len(self.empleados), COLOR_OK)
            finally:
                letrero.cerrar()

        ejecutar_en_hilo(self.parent_frame, tarea, al_terminar=terminar)

    def refrescar_combos_globales(self):
        """Actualiza los combos que dependen del padrón de empleados."""
        opciones = self.opciones_empleados()
        for combo in getattr(self, "_combos_empleados", []):
            try:
                actual = combo.get()
                combo.configure(values=opciones)
                if actual not in opciones:
                    combo.set(opciones[0] if opciones else "")
            except Exception:
                pass
        turnos_opciones = ["(Descanso)"] + ["%s | %s (%s-%s)" % (t["id"], t["nombre"], t["hora_entrada"],
                                                                t["hora_salida"]) for t in self.turnos]
        horarios_opciones = ["(Ninguno)"] + ["%s | %s" % (h["id"], h["nombre"]) for h in self.horarios]
        for combo in getattr(self, "_combos_turnos", []):
            try:
                actual = combo.get()
                combo.configure(values=turnos_opciones)
                if actual not in turnos_opciones:
                    combo.set(turnos_opciones[0])
            except Exception:
                pass
        for combo in getattr(self, "_combos_horarios", []):
            try:
                actual = combo.get()
                combo.configure(values=horarios_opciones)
                if actual not in horarios_opciones:
                    combo.set(horarios_opciones[0])
            except Exception:
                pass

    def registrar_combo_empleado(self, combo):
        if not hasattr(self, "_combos_empleados"):
            self._combos_empleados = []
        self._combos_empleados.append(combo)
        return combo

    def registrar_combo_turno(self, combo):
        if not hasattr(self, "_combos_turnos"):
            self._combos_turnos = []
        self._combos_turnos.append(combo)
        return combo

    def registrar_combo_horario(self, combo):
        if not hasattr(self, "_combos_horarios"):
            self._combos_horarios = []
        self._combos_horarios.append(combo)
        return combo

    # =====================================================
    # PESTAÑA: PANEL
    # =====================================================
    def construir_tab_panel(self):
        tab = self.tab_panel
        self.frame_kpis = ctk.CTkFrame(tab, fg_color="transparent")
        self.frame_kpis.pack(fill="x", pady=(10, 5))

        contenedor_acciones = ctk.CTkFrame(tab, fg_color=COLOR_FONDO_SUAVE, corner_radius=10,
                                          border_width=1, border_color="#e0e0e0")
        contenedor_acciones.pack(fill="x", pady=10)
        ctk.CTkLabel(contenedor_acciones, text="⚡ ACCIONES RÁPIDAS", font=("Arial", 13, "bold"),
                     text_color=COLOR_PRIMARIO).pack(anchor="w", padx=15, pady=(10, 5))
        fila = ctk.CTkFrame(contenedor_acciones, fg_color="transparent")
        fila.pack(fill="x", padx=15, pady=(0, 12))
        ctk.CTkButton(fila, text="📥 Cargar marcaciones", width=175, fg_color=COLOR_PRIMARIO,
                      hover_color=COLOR_HOVER,
                      command=lambda: self.tabview.set(" 📥 Marcaciones ")).pack(side="left", padx=3)
        ctk.CTkButton(fila, text="🔄 Recalcular asistencia", width=175, fg_color=COLOR_ALERTA,
                      hover_color="#b9640f",
                      command=self.recalcular_periodo_actual).pack(side="left", padx=3)
        ctk.CTkButton(fila, text="🧮 Calcular planilla", width=155, fg_color=COLOR_OK,
                      hover_color="#1e8449",
                      command=lambda: self.tabview.set(" 💵 Planilla ")).pack(side="left", padx=3)
        ctk.CTkButton(fila, text="📈 Ver reportes", width=135, fg_color="#8e44ad",
                      hover_color="#6c3483",
                      command=lambda: self.tabview.set(" 📈 Reportes ")).pack(side="left", padx=3)
        # El asistente va en su propia fila para que siempre se vea, incluso en ventanas angostas
        fila_asistente = ctk.CTkFrame(contenedor_acciones, fg_color="transparent")
        fila_asistente.pack(fill="x", padx=15, pady=(0, 12))
        ctk.CTkButton(fila_asistente, text="🧭 ASISTENTE PASO A PASO", width=250, height=34,
                      font=("Arial", 13, "bold"), fg_color="#16a085", hover_color="#11806a",
                      command=self.abrir_asistente).pack(side="left")
        ctk.CTkLabel(fila_asistente, text="Le indica qué falta y lo lleva a cada pantalla, "
                                          "paso a paso (configuración y cierre del mes)",
                     font=("Arial", 10), text_color=COLOR_GRIS).pack(side="left", padx=12)
        ctk.CTkLabel(contenedor_acciones,
                     text="Orden de trabajo (es el mismo orden de las pestañas):   "
                          "1) Personal   →   2) Configuración   →   3) Marcaciones   →   "
                          "4) Asistencia   →   5) Planilla   →   6) Reportes",
                     font=("Arial", 10), text_color=COLOR_GRIS).pack(anchor="w", padx=15, pady=(0, 10))

        self.frame_alertas = ctk.CTkFrame(tab, fg_color="#fff3cd", corner_radius=10, border_width=1,
                                         border_color="#ffeeba")
        self.frame_alertas.pack(fill="x", pady=(0, 10))
        self.lbl_alertas = ctk.CTkLabel(self.frame_alertas, text="", font=("Arial", 12),
                                        text_color="#856404", justify="left", anchor="w")
        self.lbl_alertas.pack(fill="x", padx=15, pady=12)

        contenedor_ult = ctk.CTkFrame(tab, fg_color="transparent")
        contenedor_ult.pack(fill="both", expand=True)
        ctk.CTkLabel(contenedor_ult, text="📜 ÚLTIMAS IMPORTACIONES DEL RELOJ", font=("Arial", 13, "bold"),
                     text_color=COLOR_PRIMARIO).pack(anchor="w", pady=(5, 4))
        self.tabla_panel_import = crear_tabla(contenedor_ult, [
            ("archivo", "Archivo", 300, "w"),
            ("formato", "Formato", 190, "w"),
            ("total", "Filas", 70, "center"),
            ("nuevas", "Nuevas", 70, "center"),
            ("duplicadas", "Duplicadas", 80, "center"),
            ("sin_empleado", "Sin vincular", 90, "center"),
            ("rango", "Rango de fechas", 190, "center"),
            ("usuario", "Usuario", 120, "w"),
            ("fecha", "Importado", 140, "center"),
        ], alto=8)

    def cargar_panel(self):
        """Lee los indicadores en segundo plano y luego pinta el panel.

        El panel hace varias consultas a la base de datos (cada una tarda unas décimas por
        la conexión a la nube). Antes se hacía en el hilo principal y la ventana se congelaba
        unos segundos; ahora se lee aparte, con el letrero centrado, y la ventana sigue viva.
        """
        if getattr(self, "_panel_cargando", False):
            return
        self._panel_cargando = True
        letrero = LetreroCarga(self.frame_main, "Leyendo la base de datos...",
                               "Indicadores de %s" % nc.nombre_mes(self.periodo_actual))
        periodo = self.periodo_actual

        def tarea(estado):
            estado["indicadores"] = nc.kpis_dashboard(periodo)
            try:
                estado["asistente"] = nc.estado_configuracion(periodo)
            except Exception as e:
                print("[Nómina] No se pudo evaluar el asistente:", e)
                estado["asistente"] = None
            estado["importaciones"] = nc.listar_importaciones(limite=10)

        def terminar(estado):
            self._panel_cargando = False
            letrero.cerrar()
            if estado.get("error"):
                print("[Nómina] Error cargando el panel:", estado["error"])
                return
            self.pintar_panel(estado.get("indicadores") or {}, estado.get("asistente"),
                              estado.get("importaciones") or [])

        ejecutar_en_hilo(self.parent_frame, tarea, al_terminar=terminar)

    def pintar_panel(self, indicadores, estado_asistente, importaciones):
        """Dibuja los indicadores, los avisos y las últimas importaciones del panel."""
        for widget in self.frame_kpis.winfo_children():
            widget.destroy()

        tarjetas = [
            ("👥 Empleados activos", str(indicadores.get("empleados_activos", 0)), COLOR_PRIMARIO),
            ("📥 Marcaciones cargadas", formatear_numero(indicadores.get("marcaciones", 0), 0), "#16a085"),
            ("⏰ Tardanzas del mes", str(indicadores.get("tardanzas", 0)), COLOR_ALERTA),
            ("🚫 Faltas del mes", str(indicadores.get("faltas", 0)), COLOR_ERROR),
            ("⚡ Horas extra del mes", "%s h" % formatear_numero(indicadores.get("horas_extra", 0), 2),
             "#8e44ad"),
            ("🎁 Horas a bono", "%s h" % formatear_numero(indicadores.get("horas_excedente", 0), 2),
             "#d35400"),
            ("💵 Neto planilla", formatear_monto(indicadores.get("planilla_neto", 0)), COLOR_OK),
        ]
        for indice, (titulo, valor, color) in enumerate(tarjetas):
            tarjeta = ctk.CTkFrame(self.frame_kpis, fg_color="#ffffff", corner_radius=10,
                                   border_width=1, border_color="#e0e0e0")
            tarjeta.pack(side="left", fill="both", expand=True, padx=5, pady=5)
            ctk.CTkLabel(tarjeta, text=titulo, font=("Arial", 11), text_color=COLOR_GRIS).pack(pady=(12, 2))
            ctk.CTkLabel(tarjeta, text=valor, font=("Arial", 20, "bold"),
                         text_color=color).pack(pady=(0, 12))

        alertas = []
        if indicadores.get("sin_mapeo"):
            alertas.append("⚠️ %d persona(s) del reloj sin vincular a un DNI. Revise «Marcaciones → "
                           "Personas del reloj»." % indicadores["sin_mapeo"])
        sin_horario = [e for e in self.empleados
                       if not e.get("horario_id") and not e.get("turno_id")
                       and nc.normalizar_texto(e.get("estado")) != "CESADO"]
        if sin_horario:
            alertas.append("⚠️ %d empleado(s) sin turno ni horario asignado (se calcularán como "
                           "descanso): %s" % (len(sin_horario),
                                              ", ".join(e["nombre"].split()[0] for e in sin_horario[:6])))
        sin_sueldo = [e for e in self.empleados
                      if nc.a_float(e.get("sueldo_basico")) <= 0
                      and nc.normalizar_texto(e.get("estado")) != "CESADO"]
        if sin_sueldo:
            alertas.append("💵 %d empleado(s) sin sueldo registrado: la planilla les saldrá en cero. "
                           "Cárguelos de una sola vez en «Personal → Empleados → 💵 Sueldos y "
                           "garantías»." % len(sin_sueldo))
        if not self.turnos:
            alertas.append("⚠️ No hay turnos configurados. Créelos en «Personal → Turnos».")
        # Revisión de parámetros: evita que un 0 accidental deje el cálculo en cero
        _limpios, errores_param, avisos_param = nc.validar_parametros(
            {clave: self.parametros.get(clave) for clave in nc.PARAMETROS_DEFECTO})
        if errores_param:
            alertas.append("⛔ %d parámetro(s) con valor inválido. Vaya a «Configuración → Parámetros de "
                           "cálculo → Reparar valores inválidos»." % len(errores_param))
        for aviso in avisos_param[:3]:
            alertas.append("⚠️ Revise los parámetros: %s" % aviso)
        indicadores_mes = indicadores.get("dias_calculados", 0)
        if not indicadores_mes:
            alertas.append("ℹ️ Aún no se ha calculado la asistencia de %s. Use «Recalcular asistencia»."
                           % nc.nombre_mes(self.periodo_actual))
        ultima = indicadores.get("ultima_importacion")
        if ultima:
            alertas.append("📥 Última importación: %s el %s (%s nuevas marcas)."
                           % (ultima["archivo"], ultima["fecha"], ultima["nuevas"]))
        # Resumen del asistente: indica cuántos pasos faltan y cuál toca ahora
        try:
            self.estado_asistente = estado_asistente or {}
            pendientes = (self.estado_asistente.get("pendientes") or [])
            atenciones = (self.estado_asistente.get("atenciones") or [])
            if pendientes:
                alertas.append("🧭 Asistente: %d paso(s) pendiente(s). El siguiente es «%s». "
                               "Use el botón «ASISTENTE PASO A PASO»."
                               % (len(pendientes), pendientes[0]["titulo"]))
            elif atenciones:
                alertas.append("🧭 Asistente: lo crítico está listo; conviene revisar «%s»."
                               % atenciones[0]["titulo"])
            else:
                alertas.append("✅ Asistente: %d%% completado, sin pendientes."
                               % self.estado_asistente.get("avance", 100))
        except Exception as e:
            print("[Nómina] No se pudo evaluar el asistente:", e)
        self.lbl_alertas.configure(text="\n".join(alertas) if alertas else
                                   "✅ Todo en orden: padrón, turnos y marcaciones sincronizados.")

        limpiar_tabla(self.tabla_panel_import)
        for importacion in importaciones:
            self.tabla_panel_import.insert("", tk.END, values=(
                importacion["archivo"], importacion["formato"], importacion["total_filas"],
                importacion["nuevas"], importacion["duplicadas"], importacion["sin_empleado"],
                "%s → %s" % (importacion["fecha_min"] or "-", importacion["fecha_max"] or "-"),
                importacion["usuario"], importacion["fecha"]))

    # =====================================================
    # PESTAÑA: MARCACIONES
    # =====================================================
    def construir_tab_marcaciones(self):
        tab = self.tab_marcaciones
        self.sub_marcaciones = ctk.CTkSegmentedButton(
            tab, values=["Importar Excel del reloj", "Personas del reloj", "Historial"],
            selected_color=COLOR_PRIMARIO, selected_hover_color=COLOR_HOVER,
            command=self.cambiar_sub_marcaciones)
        self.sub_marcaciones.pack(fill="x", padx=10, pady=(10, 5))
        self.sub_marcaciones.set("Importar Excel del reloj")

        self.frame_marc_importar = ctk.CTkFrame(tab, fg_color="transparent")
        self.frame_marc_personas = ctk.CTkFrame(tab, fg_color="transparent")
        self.frame_marc_historial = ctk.CTkFrame(tab, fg_color="transparent")
        self._construir_importador(self.frame_marc_importar)
        self._construir_personas_reloj(self.frame_marc_personas)
        self._construir_historial(self.frame_marc_historial)
        self.cambiar_sub_marcaciones("Importar Excel del reloj")

    def cambiar_sub_marcaciones(self, valor):
        for frame in (self.frame_marc_importar, self.frame_marc_personas, self.frame_marc_historial):
            frame.pack_forget()
        if valor == "Importar Excel del reloj":
            self.frame_marc_importar.pack(fill="both", expand=True)
        elif valor == "Personas del reloj":
            self.frame_marc_personas.pack(fill="both", expand=True)
        else:
            self.frame_marc_historial.pack(fill="both", expand=True)

    def _construir_importador(self, padre):
        contenedor = ctk.CTkFrame(padre, fg_color=COLOR_FONDO_SUAVE, corner_radius=10,
                                  border_width=1, border_color="#e0e0e0")
        contenedor.pack(fill="x", pady=(8, 6))
        ctk.CTkLabel(contenedor, text="1️⃣ Seleccione el archivo exportado por el captahuella",
                     font=("Arial", 13, "bold"), text_color=COLOR_PRIMARIO).pack(anchor="w", padx=15,
                                                                                 pady=(10, 4))
        fila = ctk.CTkFrame(contenedor, fg_color="transparent")
        fila.pack(fill="x", padx=15, pady=(0, 12))
        self.ent_archivo = ctk.CTkEntry(fila, width=520, placeholder_text="Ruta del archivo .xlsx / .csv")
        self.ent_archivo.pack(side="left", padx=(0, 6))
        ctk.CTkButton(fila, text="📂 Seleccionar archivo", width=170, fg_color=COLOR_PRIMARIO,
                      hover_color=COLOR_HOVER, command=self.seleccionar_archivo_reloj).pack(side="left", padx=4)
        ctk.CTkButton(fila, text="🔎 Analizar", width=120, fg_color="#16a085", hover_color="#11806a",
                      command=self.analizar_archivo).pack(side="left", padx=4)

        self.lbl_analisis = ctk.CTkLabel(contenedor, text="Sin archivo analizado.", font=("Arial", 11),
                                         text_color=COLOR_GRIS, justify="left", anchor="w")
        self.lbl_analisis.pack(fill="x", padx=15, pady=(0, 10))

        contenedor_mapeo = ctk.CTkFrame(padre, fg_color="#ffffff", corner_radius=10, border_width=1,
                                        border_color="#e0e0e0")
        contenedor_mapeo.pack(fill="x", pady=6)
        ctk.CTkLabel(contenedor_mapeo, text="2️⃣ Verifique las columnas detectadas (puede corregirlas)",
                     font=("Arial", 13, "bold"), text_color=COLOR_PRIMARIO).pack(anchor="w", padx=15,
                                                                                 pady=(10, 4))
        self.frame_combos_mapeo = ctk.CTkFrame(contenedor_mapeo, fg_color="transparent")
        self.frame_combos_mapeo.pack(fill="x", padx=15, pady=(0, 10))
        # Dos columnas (no tres): asi las casillas caben incluso en ventanas angostas
        columnas_mapeo = 2
        for indice, (campo, etiqueta) in enumerate(ETIQUETA_CAMPO_MAPEO):
            columna = indice % columnas_mapeo
            fila_g = indice // columnas_mapeo
            celda = ctk.CTkFrame(self.frame_combos_mapeo, fg_color="transparent")
            celda.grid(row=fila_g, column=columna, sticky="w", padx=6, pady=3)
            ctk.CTkLabel(celda, text=etiqueta + ":", font=("Arial", 10, "bold"), width=150,
                         anchor="w").pack(side="left")
            combo = ctk.CTkComboBox(celda, values=[SIN_COLUMNA], width=210, state="readonly",
                                    font=("Arial", 10))
            combo.pack(side="left")
            combo.set(SIN_COLUMNA)
            self.combos_mapeo[campo] = combo

        contenedor_prev = ctk.CTkFrame(padre, fg_color="transparent")
        contenedor_prev.pack(fill="both", expand=True, pady=6)
        ctk.CTkLabel(contenedor_prev, text="3️⃣ Vista previa (primeras 150 filas válidas)",
                     font=("Arial", 13, "bold"), text_color=COLOR_PRIMARIO).pack(anchor="w")
        self.tabla_preview = crear_tabla(contenedor_prev, [
            ("codigo", "Código", 110, "w"),
            ("nombre", "Nombre en el reloj", 260, "w"),
            ("fecha", "Fecha", 100, "center"),
            ("hora", "Hora", 80, "center"),
            ("tipo", "Tipo de pase", 140, "w"),
            ("metodo", "Método", 120, "w"),
            ("vinculo", "Vinculación", 220, "w"),
        ], alto=9)

        barra = ctk.CTkFrame(padre, fg_color="transparent")
        barra.pack(fill="x", pady=(6, 2))
        self.btn_importar = ctk.CTkButton(barra, text="✅ IMPORTAR MARCACIONES A LA BASE DE DATOS",
                                         width=420, height=38, font=("Arial", 13, "bold"),
                                         fg_color=COLOR_OK, hover_color="#1e8449",
                                         command=self.importar_marcaciones)
        self.btn_importar.pack(side="left")
        ctk.CTkButton(barra, text="🗑️ Eliminar marcaciones de un rango", width=250, height=38,
                      fg_color=COLOR_ERROR, hover_color="#922b21",
                      command=self.eliminar_marcaciones_rango).pack(side="left", padx=8)
        self.lbl_progreso = ctk.CTkLabel(barra, text="", font=("Arial", 11), text_color=COLOR_GRIS)
        self.lbl_progreso.pack(side="left", padx=10)
        # La barra se ancla al fondo para que el boton de importar nunca quede fuera de vista
        contenedor_prev.pack_forget()
        barra.pack_forget()
        barra.pack(side="bottom", fill="x", pady=(6, 2))
        contenedor_prev.pack(side="top", fill="both", expand=True, pady=6)

    def _construir_personas_reloj(self, padre):
        ctk.CTkLabel(padre, text="👤 PERSONAS DETECTADAS EN EL RELOJ", font=("Arial", 14, "bold"),
                     text_color=COLOR_PRIMARIO).pack(anchor="w", pady=(10, 2))
        ctk.CTkLabel(padre, text="Vincule cada código del reloj con el DNI del empleado. Mientras no exista "
                                "vínculo, sus marcaciones no entran en la asistencia.",
                     font=("Arial", 11), text_color=COLOR_GRIS).pack(anchor="w", pady=(0, 6))
        self.tabla_personas = crear_tabla(padre, [
            ("codigo", "Código en el reloj", 140, "w"),
            ("nombre", "Nombre registrado en el reloj", 300, "w"),
            ("marcas", "Marcas", 70, "center"),
            ("rango", "Rango de fechas", 190, "center"),
            ("sugerencia", "Sugerencia del sistema", 210, "w"),
            ("estado", "Estado", 130, "w"),
        ], alto=9)
        self.tabla_personas.bind("<<TreeviewSelect>>", self.al_seleccionar_persona)

        contenedor = ctk.CTkFrame(padre, fg_color=COLOR_FONDO_SUAVE, corner_radius=10, border_width=1,
                                  border_color="#e0e0e0")
        contenedor.pack(fill="x", pady=8)
        ctk.CTkLabel(contenedor, text="🔗 Vincular la persona seleccionada", font=("Arial", 12, "bold"),
                     text_color=COLOR_PRIMARIO).pack(anchor="w", padx=15, pady=(10, 4))
        fila = ctk.CTkFrame(contenedor, fg_color="transparent")
        fila.pack(fill="x", padx=15, pady=(0, 12))
        ctk.CTkLabel(fila, text="Empleado (DNI):", font=("Arial", 11, "bold")).pack(side="left")
        self.cmb_persona_empleado = ctk.CTkComboBox(fila, values=self.opciones_empleados(), width=380,
                                                    state="readonly")
        self.cmb_persona_empleado.pack(side="left", padx=8)
        self.registrar_combo_empleado(self.cmb_persona_empleado)
        self.ent_codigo_reloj = ctk.CTkEntry(fila, width=150, placeholder_text="Código del reloj")
        self.ent_codigo_reloj.pack(side="left", padx=8)
        ctk.CTkButton(fila, text="🔗 Vincular", width=120, fg_color=COLOR_OK, hover_color="#1e8449",
                      command=self.vincular_persona).pack(side="left", padx=4)
        ctk.CTkButton(fila, text="🗑️ Quitar vínculo", width=140, fg_color=COLOR_ERROR,
                      hover_color="#922b21", command=self.quitar_vinculo).pack(side="left", padx=4)
        ctk.CTkButton(fila, text="🤖 Vincular automáticamente", width=190, fg_color="#8e44ad",
                      hover_color="#6c3483",
                      command=self.vincular_automatico).pack(side="left", padx=4)

        ctk.CTkLabel(padre, text="📋 VÍNCULOS GUARDADOS", font=("Arial", 13, "bold"),
                     text_color=COLOR_PRIMARIO).pack(anchor="w", pady=(6, 2))
        self.tabla_mapeos = crear_tabla(padre, [
            ("codigo", "Código del reloj", 160, "w"),
            ("dni", "DNI vinculado", 140, "w"),
            ("empleado", "Empleado", 300, "w"),
            ("nombre_reloj", "Nombre en el reloj", 300, "w"),
        ], alto=6)

    def _construir_historial(self, padre):
        ctk.CTkLabel(padre, text="📜 HISTORIAL DE IMPORTACIONES", font=("Arial", 14, "bold"),
                     text_color=COLOR_PRIMARIO).pack(anchor="w", pady=(10, 4))
        self.tabla_historial = crear_tabla(padre, [
            ("id", "ID", 55, "center"),
            ("archivo", "Archivo", 300, "w"),
            ("formato", "Formato", 190, "w"),
            ("total", "Filas leídas", 90, "center"),
            ("nuevas", "Nuevas", 80, "center"),
            ("duplicadas", "Duplicadas", 90, "center"),
            ("sin_empleado", "Sin vincular", 90, "center"),
            ("rango", "Rango de marcaciones", 200, "center"),
            ("usuario", "Usuario", 130, "w"),
            ("fecha", "Fecha de importación", 150, "center"),
        ], alto=14)

    def seleccionar_archivo_reloj(self):
        ruta = seleccionar_archivo_dialogo("Seleccionar el archivo del reloj biométrico",
                                          [("Excel / CSV", "*.xlsx *.xls *.csv"), ("Todos", "*.*")])
        if not ruta:
            return
        self.ent_archivo.delete(0, tk.END)
        self.ent_archivo.insert(0, ruta)
        self.analizar_archivo()

    def analizar_archivo(self):
        ruta = self.ent_archivo.get().strip()
        if not ruta or not os.path.exists(ruta):
            return messagebox.showwarning("Atención", "Seleccione primero un archivo válido.")
        self.aviso("⏳ Analizando archivo...", COLOR_ALERTA)
        letrero = LetreroCarga(self.frame_main, "Leyendo el archivo del reloj...",
                               os.path.basename(ruta))

        def tarea(estado):
            estado["lectura"] = nc.leer_archivo_marcaciones(ruta)

        def terminar(estado):
            letrero.cerrar()
            if estado.get("error"):
                self.aviso("❌ %s" % estado["error"], COLOR_ERROR)
                return messagebox.showerror("Error", "No se pudo leer el archivo:\n%s" % estado["error"])
            lectura = estado.get("lectura") or {}
            if not lectura.get("ok"):
                self.aviso("❌ %s" % lectura.get("error"), COLOR_ERROR)
                return messagebox.showerror("Error", lectura.get("error") or "Archivo no reconocido.")
            self.lectura_actual = lectura
            self.archivo_actual = ruta
            encabezados = lectura["encabezados"]
            opciones = [SIN_COLUMNA] + [("%d - %s" % (i, (t or "(vacío)")[:40]))
                                        for i, t in enumerate(encabezados)]
            for campo, combo in self.combos_mapeo.items():
                combo.configure(values=opciones)
                indice = lectura["mapeo"].get(campo)
                if indice is not None and indice < len(encabezados):
                    combo.set("%d - %s" % (indice, (encabezados[indice] or "(vacío)")[:40]))
                else:
                    combo.set(SIN_COLUMNA)
            validas = [f for f in lectura["filas"] if f["valida"]]
            self.lbl_analisis.configure(
                text="📄 %s\nFormato detectado: %s   |   Encabezado en la fila %d   |   "
                     "Filas leídas: %d   |   Filas válidas: %d   |   Hojas: %s"
                     % (os.path.basename(ruta), nc.PERFILES_FORMATO.get(lectura["formato"],
                                                                       lectura["formato"]),
                        lectura["fila_encabezado"] + 1, lectura["total"], len(validas),
                        ", ".join(lectura.get("hojas") or ["-"])),
                text_color="#1a1a1a")
            self.pintar_preview(validas[:150])
            self.aviso("✅ Archivo analizado: %d filas válidas de %d" % (len(validas), lectura["total"]),
                       COLOR_OK)
            if not validas:
                messagebox.showwarning("Sin datos", "El archivo no contiene filas válidas de marcación. "
                                                    "Revise el mapeo de columnas.")

        ejecutar_en_hilo(self.parent_frame, tarea, al_terminar=terminar)

    def mapeo_desde_combos(self):
        """Construye el diccionario {campo: índice} según los combos de la interfaz."""
        mapeo = {}
        for campo, combo in self.combos_mapeo.items():
            texto = combo.get()
            if texto and texto != SIN_COLUMNA:
                try:
                    mapeo[campo] = int(texto.split(" - ")[0])
                except (ValueError, IndexError):
                    pass
        return mapeo

    def pintar_preview(self, filas):
        limpiar_tabla(self.tabla_preview)
        mapeos = nc.listar_mapeos()
        empleados = self.empleados
        for fila in filas:
            vinculo = "—"
            if fila["codigo"] in mapeos:
                dni = mapeos[fila["codigo"]]["dni"]
                vinculo = "✅ %s %s" % (dni, self.nombre_empleado(dni)[:24])
            else:
                dni, metodo, _ = nc.sugerir_dni(fila["codigo"], fila["nombre_reloj"], empleados, mapeos)
                vinculo = ("🔶 %s (%s)" % (dni, metodo)) if dni else "❌ Sin vincular"
            self.tabla_preview.insert("", tk.END, values=(
                fila["codigo"], fila["nombre_reloj"], fila["fecha"], fila["hora"],
                fila["tipo_pase"], fila["metodo"], vinculo))

    def importar_marcaciones(self):
        if not self.archivo_actual:
            return messagebox.showwarning("Atención", "Primero seleccione y analice un archivo.")
        if not messagebox.askyesno("Confirmar importación",
                                   "¿Importar las marcaciones del archivo seleccionado?\n\n"
                                   "Las marcas repetidas se descartan automáticamente."):
            return
        mapeo = self.mapeo_desde_combos()
        if "fecha" not in mapeo or ("hora" not in mapeo and mapeo.get("fecha") is None):
            return messagebox.showwarning("Mapeo incompleto",
                                          "Debe indicar al menos la columna de Fecha y la de Hora.")
        archivo = self.archivo_actual
        self.btn_importar.configure(state="disabled", text="⏳ IMPORTANDO...")
        self.aviso("⏳ Importando marcaciones...", COLOR_ALERTA)
        letrero = LetreroCarga(self.frame_main, "Importando las marcaciones...",
                               os.path.basename(archivo))

        def progreso(actual, total):
            try:
                self.lbl_progreso.configure(text="Procesando %d de %d..." % (actual, total))
            except Exception:
                pass
            letrero.cambiar("Importando las marcaciones...", "Fila %d de %d" % (actual, total))

        def tarea(estado):
            resultado, _ = nc.importar_marcaciones(
                archivo, mapeo=mapeo, fila_encabezado=self.lectura_actual["fila_encabezado"],
                usuario=self.usuario_activo, progreso=progreso)
            estado["resultado"] = resultado

        def terminar(estado):
            letrero.cerrar()
            self.btn_importar.configure(state="normal", text="✅ IMPORTAR MARCACIONES A LA BASE DE DATOS")
            self.lbl_progreso.configure(text="")
            if estado.get("error"):
                self.aviso("❌ Error al importar", COLOR_ERROR)
                return messagebox.showerror("Error", "No se pudo importar:\n%s" % estado["error"])
            resultado = estado.get("resultado") or {}
            if not resultado.get("ok"):
                self.aviso("❌ %s" % resultado.get("error"), COLOR_ERROR)
                return messagebox.showerror("Error", resultado.get("error") or "Importación fallida.")
            self.aviso("✅ %d marcas nuevas importadas" % resultado["nuevas"], COLOR_OK)
            mensaje = ("Importación completada\n\n"
                       "Archivo: %s\n"
                       "Formato: %s\n"
                       "Filas leídas: %d\n"
                       "Marcas nuevas: %d\n"
                       "Marcas duplicadas (descartadas): %d\n"
                       "Filas inválidas: %d\n"
                       "Período: %s → %s\n"
                       "Personas detectadas: %d"
                       % (resultado["archivo"], resultado["formato"], resultado["total_filas"],
                          resultado["nuevas"], resultado["duplicadas"], resultado["invalidas"],
                          resultado["fecha_min"], resultado["fecha_max"], len(resultado["personas"])))
            if resultado.get("sin_empleado_codigos"):
                mensaje += ("\n\n⚠️ %d persona(s) sin vincular a un DNI. Use «Personas del reloj» "
                            "para vincularlas y luego recalcule la asistencia."
                            % len(resultado["sin_empleado_codigos"]))
            messagebox.showinfo("Importación completada", mensaje)
            self.cargar_historial_importaciones()
            self.cargar_personas_reloj()
            self.cargar_mapeos()
            self.cargar_panel()

        ejecutar_en_hilo(self.parent_frame, tarea, al_terminar=terminar)

    def eliminar_marcaciones_rango(self):
        from tkinter import simpledialog
        desde = simpledialog.askstring("Eliminar marcaciones", "Fecha inicial (DD/MM/AAAA):",
                                       parent=self.parent_frame)
        if not desde:
            return
        hasta = simpledialog.askstring("Eliminar marcaciones", "Fecha final (DD/MM/AAAA):",
                                       parent=self.parent_frame)
        if not hasta:
            return
        if not messagebox.askyesno("Confirmar", "¿Eliminar TODAS las marcaciones entre %s y %s?\n\n"
                                                "Esta acción no se puede deshacer." % (desde, hasta)):
            return
        ok, mensaje = nc.eliminar_marcaciones_periodo(desde, hasta, self.usuario_activo)
        if ok:
            messagebox.showinfo("Listo", mensaje)
            self.cargar_historial_importaciones()
            self.cargar_personas_reloj()
            self.cargar_panel()
        else:
            messagebox.showerror("Error", mensaje)

    def cargar_personas_reloj(self):
        letrero = LetreroCarga(self.frame_main, "Leyendo la base de datos...",
                               "Personas detectadas en el reloj")

        def tarea(estado):
            estado["personas"] = nc.personas_del_reloj()
            estado["mapeos"] = nc.listar_mapeos()

        def terminar(estado):
            letrero.cerrar()
            if estado.get("error"):
                return
            personas = estado.get("personas") or []
            mapeos = estado.get("mapeos") or {}
            limpiar_tabla(self.tabla_personas)
            for persona in personas:
                codigo = persona["codigo_reloj"]
                if codigo in mapeos:
                    sugerencia = "Vinculado: %s" % self.nombre_empleado(mapeos[codigo]["dni"])
                    estado_texto = "✅ Vinculado"
                else:
                    dni, metodo, _ = nc.sugerir_dni(codigo, persona["nombre_reloj"], self.empleados, mapeos)
                    sugerencia = ("%s (%s)" % (dni, metodo)) if dni else "—"
                    estado_texto = "🔶 Por vincular" if dni else "❌ Desconocido"
                self.tabla_personas.insert("", tk.END, values=(
                    codigo, persona["nombre_reloj"], persona["marcas"],
                    "%s → %s" % (persona["fecha_min"], persona["fecha_max"]), sugerencia, estado_texto))

        ejecutar_en_hilo(self.parent_frame, tarea, al_terminar=terminar)

    def cargar_mapeos(self):
        limpiar_tabla(self.tabla_mapeos)
        for codigo, datos in sorted(nc.listar_mapeos().items()):
            self.tabla_mapeos.insert("", tk.END, values=(
                codigo, datos["dni"], self.nombre_empleado(datos["dni"]), datos.get("nombre_reloj") or ""))

    def al_seleccionar_persona(self, evento=None):
        seleccion = self.tabla_personas.selection()
        if not seleccion:
            return
        valores = self.tabla_personas.item(seleccion[0], "values")
        codigo = valores[0]
        self.ent_codigo_reloj.delete(0, tk.END)
        self.ent_codigo_reloj.insert(0, codigo)
        mapeos = nc.listar_mapeos()
        if codigo in mapeos:
            objetivo = mapeos[codigo]["dni"]
            for opcion in self.cmb_persona_empleado.cget("values"):
                if opcion.split("|")[0].strip() == str(objetivo):
                    self.cmb_persona_empleado.set(opcion)
                    break
        else:
            dni, _, _ = nc.sugerir_dni(codigo, valores[1], self.empleados, mapeos)
            if dni:
                for opcion in self.cmb_persona_empleado.cget("values"):
                    if opcion.split("|")[0].strip() == str(dni):
                        self.cmb_persona_empleado.set(opcion)
                        break

    def vincular_persona(self):
        codigo = self.ent_codigo_reloj.get().strip()
        if not codigo:
            return messagebox.showwarning("Atención", "Indique el código del reloj (o seleccione una fila).")
        dni = self.dni_de_opcion(self.cmb_persona_empleado.get())
        seleccion = self.tabla_personas.selection()
        nombre_reloj = self.tabla_personas.item(seleccion[0], "values")[1] if seleccion else ""
        ok, error = nc.guardar_mapeo(codigo, dni, nombre_reloj, self.usuario_activo)
        if not ok:
            return messagebox.showerror("Error", error)
        self.aviso("✅ Código %s vinculado con el DNI %s" % (codigo, dni), COLOR_OK)
        self.cargar_personas_reloj()
        self.cargar_mapeos()

    def quitar_vinculo(self):
        codigo = self.ent_codigo_reloj.get().strip()
        if not codigo:
            return messagebox.showwarning("Atención", "Seleccione la persona o escriba el código.")
        if not messagebox.askyesno("Confirmar", "¿Quitar el vínculo del código %s?" % codigo):
            return
        nc.eliminar_mapeo(codigo, self.usuario_activo)
        self.cargar_personas_reloj()
        self.cargar_mapeos()

    def vincular_automatico(self):
        """Vincula automáticamente las personas cuya coincidencia sea confiable."""
        mapeos = nc.listar_mapeos()
        vinculados = 0
        for persona in nc.personas_del_reloj():
            codigo = persona["codigo_reloj"]
            if codigo in mapeos:
                continue
            dni, metodo, confianza = nc.sugerir_dni(codigo, persona["nombre_reloj"], self.empleados, mapeos)
            if dni and confianza >= 0.6:
                ok, _ = nc.guardar_mapeo(codigo, dni, persona["nombre_reloj"], self.usuario_activo)
                if ok:
                    vinculados += 1
        self.cargar_personas_reloj()
        self.cargar_mapeos()
        messagebox.showinfo("Vinculación automática",
                            "Se vincularon %d persona(s) del reloj con su DNI.\n\n"
                            "Revise las que quedaron pendientes y vincúlelas manualmente." % vinculados)

    @con_letrero("Leyendo el historial de importaciones...", "")
    def cargar_historial_importaciones(self):
        limpiar_tabla(self.tabla_historial)
        for importacion in nc.listar_importaciones(limite=200):
            self.tabla_historial.insert("", tk.END, values=(
                importacion["id"], importacion["archivo"], importacion["formato"],
                importacion["total_filas"], importacion["nuevas"], importacion["duplicadas"],
                importacion["sin_empleado"],
                "%s → %s" % (importacion["fecha_min"] or "-", importacion["fecha_max"] or "-"),
                importacion["usuario"], importacion["fecha"]))

    # =====================================================
    # PESTAÑA: ASISTENCIA
    # =====================================================
    def construir_tab_asistencia(self):
        tab = self.tab_asistencia
        barra = ctk.CTkFrame(tab, fg_color="transparent")
        barra.pack(fill="x", padx=10, pady=(10, 4))
        ctk.CTkLabel(barra, text="Período:", font=("Arial", 12, "bold")).pack(side="left")
        ctk.CTkButton(barra, text="◀", width=36, fg_color=COLOR_PRIMARIO, hover_color=COLOR_HOVER,
                      command=lambda: self.cambiar_periodo(-1)).pack(side="left", padx=4)
        self.ent_periodo = ctk.CTkEntry(barra, width=90, justify="center")
        self.ent_periodo.insert(0, periodo_para_entry(self.periodo_actual))
        self.ent_periodo.pack(side="left")
        ctk.CTkButton(barra, text="▶", width=36, fg_color=COLOR_PRIMARIO, hover_color=COLOR_HOVER,
                      command=lambda: self.cambiar_periodo(1)).pack(side="left", padx=4)
        ctk.CTkButton(barra, text="🔄 Ver período", width=130, fg_color="#34495e", hover_color="#2c3e50",
                      command=self.cargar_matriz_asistencia).pack(side="left", padx=8)
        ctk.CTkButton(barra, text="🧮 RECALCULAR ASISTENCIA DEL MES", width=290, height=32,
                      font=("Arial", 12, "bold"), fg_color=COLOR_ALERTA, hover_color="#b9640f",
                      command=self.recalcular_periodo_actual).pack(side="left", padx=8)
        ctk.CTkButton(barra, text="📊 Exportar a Excel", width=160, fg_color=COLOR_OK,
                      hover_color="#1e8449", command=self.exportar_asistencia).pack(side="left", padx=4)
        self.lbl_resumen_asistencia = ctk.CTkLabel(barra, text="", font=("Arial", 11),
                                                   text_color=COLOR_GRIS)
        self.lbl_resumen_asistencia.pack(side="left", padx=12)

        self.sub_asistencia = ctk.CTkSegmentedButton(
            tab, values=["Matriz mensual", "Detalle por día", "Incidencias", "Horas extra"],
            selected_color=COLOR_PRIMARIO, selected_hover_color=COLOR_HOVER,
            command=self.cambiar_sub_asistencia)
        self.sub_asistencia.pack(fill="x", padx=10, pady=(4, 6))
        self.frame_matriz = ctk.CTkFrame(tab, fg_color="transparent")
        self.frame_detalle_dia = ctk.CTkFrame(tab, fg_color="transparent")
        self.frame_incidencias = ctk.CTkFrame(tab, fg_color="transparent")
        self.frame_horas_extra = ctk.CTkFrame(tab, fg_color="transparent")
        self._construir_matriz(self.frame_matriz)
        self._construir_detalle_dia(self.frame_detalle_dia)
        self._construir_incidencias(self.frame_incidencias)
        self._construir_horas_extra(self.frame_horas_extra)
        self.sub_asistencia.set("Matriz mensual")
        self.cambiar_sub_asistencia("Matriz mensual")

    def cambiar_sub_asistencia(self, valor):
        for frame in (self.frame_matriz, self.frame_detalle_dia, self.frame_incidencias,
                      self.frame_horas_extra):
            frame.pack_forget()
        mapa = {"Matriz mensual": self.frame_matriz, "Detalle por día": self.frame_detalle_dia,
                "Incidencias": self.frame_incidencias, "Horas extra": self.frame_horas_extra}
        frame = mapa.get(valor, self.frame_matriz)
        frame.pack(fill="both", expand=True)
        if valor == "Matriz mensual":
            self.cargar_matriz_asistencia()
        elif valor == "Detalle por día":
            self.cargar_detalle_dia()
        elif valor == "Incidencias":
            self.cargar_incidencias()
        elif valor == "Horas extra":
            self.cargar_horas_extra()

    def _construir_matriz(self, padre):
        contenedor_leyenda = ctk.CTkFrame(padre, fg_color="transparent")
        contenedor_leyenda.pack(fill="x", pady=(2, 4))
        ctk.CTkLabel(contenedor_leyenda, text="Leyenda:", font=("Arial", 10, "bold"),
                     text_color=COLOR_GRIS).pack(side="left", padx=(4, 6))
        for estado in (nc.EST_PUNTUAL, nc.EST_TARDANZA, nc.EST_FALTA, nc.EST_INCOMPLETO,
                       nc.EST_DESCANSO, nc.EST_FERIADO, nc.EST_VACACIONES, nc.EST_PERMISO):
            ctk.CTkLabel(contenedor_leyenda, text="%s = %s" % (CODIGO_ESTADO.get(estado, "?"), estado),
                         font=("Arial", 10), text_color=nc.COLORES_ESTADO.get(estado, COLOR_GRIS)
                         ).pack(side="left", padx=6)
        self.contenedor_matriz = ctk.CTkFrame(padre, fg_color="transparent")
        self.contenedor_matriz.pack(fill="both", expand=True)

    def _construir_detalle_dia(self, padre):
        barra = ctk.CTkFrame(padre, fg_color="transparent")
        barra.pack(fill="x", pady=(4, 4))
        ctk.CTkLabel(barra, text="Empleado:", font=("Arial", 11, "bold")).pack(side="left")
        self.cmb_detalle_empleado = ctk.CTkComboBox(barra, values=self.opciones_empleados(), width=340,
                                                    state="readonly",
                                                    command=lambda _v: self.cargar_detalle_dia())
        self.cmb_detalle_empleado.pack(side="left", padx=6)
        self.registrar_combo_empleado(self.cmb_detalle_empleado)
        ctk.CTkLabel(barra, text="Estado:", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 2))
        self.cmb_detalle_estado = ctk.CTkComboBox(barra, values=["Todos"] + sorted(
            set(CODIGO_ESTADO.keys())), width=180, state="readonly",
            command=lambda _v: self.cargar_detalle_dia())
        self.cmb_detalle_estado.set("Todos")
        self.cmb_detalle_estado.pack(side="left")
        ctk.CTkButton(barra, text="✏️ Corregir el día seleccionado", width=230, fg_color="#8e44ad",
                      hover_color="#6c3483",
                      command=self.corregir_dia_seleccionado).pack(side="left", padx=10)
        ctk.CTkLabel(barra, text="Doble clic en una fila para ver las marcas del reloj",
                     font=("Arial", 10), text_color=COLOR_GRIS).pack(side="left", padx=6)
        self.tabla_detalle = crear_tabla(padre, [
            ("fecha", "Fecha", 100, "center"),
            ("dia", "Día", 95, "w"),
            ("entrada", "Entrada", 85, "center"),
            ("salida", "Salida", 85, "center"),
            ("marcas", "N° marcas", 80, "center"),
            ("tardanza", "Tardanza", 90, "center"),
            ("anticipo", "Salida anticip.", 105, "center"),
            ("trabajado", "Tiempo trabajado", 125, "center"),
            ("extra", "Horas extra (tope legal)", 140, "center"),
            ("exceso", "Exceso a bono", 110, "center"),
            ("estado", "Estado", 160, "w"),
            ("observacion", "Observación", 300, "w"),
        ], alto=16)
        self.tabla_detalle.bind("<Double-1>", self.ver_marcas_del_dia)

    def _construir_incidencias(self, padre):
        ctk.CTkLabel(padre, text="📝 INCIDENCIAS: VACACIONES, PERMISOS, LICENCIAS Y DESCANSOS MÉDICOS",
                     font=("Arial", 13, "bold"), text_color=COLOR_PRIMARIO).pack(anchor="w", pady=(4, 2))
        ctk.CTkLabel(padre, text="Los días cubiertos por una incidencia no generan descuento ni falta.",
                     font=("Arial", 10), text_color=COLOR_GRIS).pack(anchor="w", pady=(0, 4))
        formulario = ctk.CTkFrame(padre, fg_color=COLOR_FONDO_SUAVE, corner_radius=10, border_width=1,
                                  border_color="#e0e0e0")
        formulario.pack(fill="x", pady=4)
        fila1 = ctk.CTkFrame(formulario, fg_color="transparent")
        fila1.pack(fill="x", padx=12, pady=(10, 4))
        ctk.CTkLabel(fila1, text="Empleado:", font=("Arial", 11, "bold")).pack(side="left")
        self.cmb_inc_empleado = ctk.CTkComboBox(fila1, values=self.opciones_empleados(), width=300,
                                                state="readonly")
        self.cmb_inc_empleado.pack(side="left", padx=6)
        self.registrar_combo_empleado(self.cmb_inc_empleado)
        ctk.CTkLabel(fila1, text="Tipo:", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 2))
        self.cmb_inc_tipo = ctk.CTkComboBox(fila1, values=nc.TIPOS_INCIDENCIA, width=180,
                                            state="readonly")
        self.cmb_inc_tipo.set(nc.EST_PERMISO)
        self.cmb_inc_tipo.pack(side="left")
        ctk.CTkLabel(fila1, text="Desde:", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 2))
        self.ent_inc_desde = ctk.CTkEntry(fila1, width=100, placeholder_text="DD/MM/AAAA")
        self.ent_inc_desde.pack(side="left")
        ctk.CTkButton(fila1, text="📅", width=32, fg_color=COLOR_PRIMARIO, hover_color=COLOR_HOVER,
                      command=lambda: self.abrir_calendario(self.ent_inc_desde)).pack(side="left", padx=2)
        ctk.CTkLabel(fila1, text="Hasta:", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 2))
        self.ent_inc_hasta = ctk.CTkEntry(fila1, width=100, placeholder_text="DD/MM/AAAA")
        self.ent_inc_hasta.pack(side="left")
        ctk.CTkButton(fila1, text="📅", width=32, fg_color=COLOR_PRIMARIO, hover_color=COLOR_HOVER,
                      command=lambda: self.abrir_calendario(self.ent_inc_hasta)).pack(side="left", padx=2)

        fila2 = ctk.CTkFrame(formulario, fg_color="transparent")
        fila2.pack(fill="x", padx=12, pady=(2, 4))
        ctk.CTkLabel(fila2, text="Motivo:", font=("Arial", 11, "bold")).pack(side="left")
        self.ent_inc_motivo = ctk.CTkEntry(fila2, width=420, placeholder_text="Detalle o sustento")
        self.ent_inc_motivo.pack(side="left", padx=6)
        self.var_inc_goce = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(fila2, text="Con goce de haber", variable=self.var_inc_goce,
                        font=("Arial", 11)).pack(side="left", padx=10)
        ctk.CTkLabel(fila2, text="Horas (opcional):", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 2))
        self.ent_inc_horas = ctk.CTkEntry(fila2, width=70)
        self.ent_inc_horas.pack(side="left")
        ctk.CTkLabel(fila2, text="Aprobado por:", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 2))
        self.ent_inc_aprobado = ctk.CTkEntry(fila2, width=150)
        self.ent_inc_aprobado.pack(side="left")

        fila3 = ctk.CTkFrame(formulario, fg_color="transparent")
        fila3.pack(fill="x", padx=12, pady=(2, 10))
        ctk.CTkButton(fila3, text="💾 Guardar incidencia", width=180, fg_color=COLOR_OK,
                      hover_color="#1e8449", command=self.guardar_incidencia).pack(side="left")
        ctk.CTkButton(fila3, text="🧹 Limpiar", width=110, fg_color=COLOR_GRIS, hover_color="#606b6b",
                      command=self.limpiar_formulario_incidencia).pack(side="left", padx=6)
        ctk.CTkButton(fila3, text="🗑️ Eliminar seleccionada", width=200, fg_color=COLOR_ERROR,
                      hover_color="#922b21", command=self.eliminar_incidencia).pack(side="left", padx=6)
        self.incidencia_id = None

        self.tabla_incidencias = crear_tabla(padre, [
            ("id", "ID", 55, "center"),
            ("dni", "DNI", 100, "w"),
            ("empleado", "Empleado", 260, "w"),
            ("tipo", "Tipo", 150, "w"),
            ("desde", "Desde", 100, "center"),
            ("hasta", "Hasta", 100, "center"),
            ("dias", "Días", 60, "center"),
            ("goce", "Con goce", 80, "center"),
            ("motivo", "Motivo", 300, "w"),
            ("aprobado", "Aprobado por", 150, "w"),
        ], alto=10)
        self.tabla_incidencias.bind("<<TreeviewSelect>>", self.al_seleccionar_incidencia)

    def _construir_horas_extra(self, padre):
        ctk.CTkLabel(padre, text="⚡ HORAS EXTRA", font=("Arial", 13, "bold"),
                     text_color=COLOR_PRIMARIO).pack(anchor="w", pady=(4, 2))
        ctk.CTkLabel(padre, text="Las horas extra se calculan automáticamente al recalcular la asistencia "
                                "(tiempo posterior a la hora de salida del turno). Aquí puede agregarlas "
                                "o corregirlas manualmente.",
                     font=("Arial", 10), text_color=COLOR_GRIS, justify="left").pack(anchor="w", pady=(0, 4))
        formulario = ctk.CTkFrame(padre, fg_color=COLOR_FONDO_SUAVE, corner_radius=10, border_width=1,
                                  border_color="#e0e0e0")
        formulario.pack(fill="x", pady=4)
        fila = ctk.CTkFrame(formulario, fg_color="transparent")
        fila.pack(fill="x", padx=12, pady=10)
        ctk.CTkLabel(fila, text="Empleado:", font=("Arial", 11, "bold")).pack(side="left")
        self.cmb_he_empleado = ctk.CTkComboBox(fila, values=self.opciones_empleados(), width=300,
                                               state="readonly")
        self.cmb_he_empleado.pack(side="left", padx=6)
        self.registrar_combo_empleado(self.cmb_he_empleado)
        ctk.CTkLabel(fila, text="Fecha:", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 2))
        self.ent_he_fecha = ctk.CTkEntry(fila, width=100, placeholder_text="DD/MM/AAAA")
        self.ent_he_fecha.pack(side="left")
        ctk.CTkButton(fila, text="📅", width=32, fg_color=COLOR_PRIMARIO, hover_color=COLOR_HOVER,
                      command=lambda: self.abrir_calendario(self.ent_he_fecha)).pack(side="left", padx=2)
        ctk.CTkLabel(fila, text="Minutos:", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 2))
        self.ent_he_minutos = ctk.CTkEntry(fila, width=70)
        self.ent_he_minutos.pack(side="left")
        ctk.CTkLabel(fila, text="Tipo:", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 2))
        self.cmb_he_tipo = ctk.CTkComboBox(fila, values=["25", "35", "NOCTURNA_25", "NOCTURNA_35"],
                                           width=140, state="readonly")
        self.cmb_he_tipo.set("25")
        self.cmb_he_tipo.pack(side="left")
        self.var_he_aprobado = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(fila, text="Aprobado", variable=self.var_he_aprobado,
                        font=("Arial", 11)).pack(side="left", padx=10)
        ctk.CTkButton(fila, text="💾 Guardar", width=110, fg_color=COLOR_OK, hover_color="#1e8449",
                      command=self.guardar_hora_extra).pack(side="left", padx=4)
        ctk.CTkButton(fila, text="🗑️ Eliminar", width=110, fg_color=COLOR_ERROR, hover_color="#922b21",
                      command=self.eliminar_hora_extra).pack(side="left", padx=4)

        self.tabla_horas_extra = crear_tabla(padre, [
            ("id", "ID", 55, "center"),
            ("dni", "DNI", 100, "w"),
            ("empleado", "Empleado", 260, "w"),
            ("fecha", "Fecha", 100, "center"),
            ("minutos", "Minutos", 80, "center"),
            ("horas", "Horas", 80, "center"),
            ("tipo", "Tipo", 130, "w"),
            ("origen", "Origen", 90, "center"),
            ("aprobado", "Aprobado", 90, "center"),
            ("observacion", "Observación", 300, "w"),
        ], alto=12)

    def cambiar_periodo(self, cantidad):
        periodo = periodo_desde_entry(self.ent_periodo.get()) or self.periodo_actual
        nuevo = sumar_meses(periodo, cantidad)
        self.ent_periodo.delete(0, tk.END)
        self.ent_periodo.insert(0, periodo_para_entry(nuevo))
        self.cargar_matriz_asistencia()

    def periodo_seleccionado(self):
        periodo = periodo_desde_entry(self.ent_periodo.get())
        if not periodo:
            messagebox.showwarning("Atención", "Escriba el período con el formato MM/AAAA.")
            return None
        return periodo

    def cargar_matriz_asistencia(self):
        periodo = self.periodo_seleccionado()
        if not periodo:
            return
        self.periodo_actual = periodo
        self.aviso("⏳ Cargando asistencia de %s..." % nc.nombre_mes(periodo), COLOR_ALERTA)
        letrero = LetreroCarga(self.frame_main, "Leyendo la base de datos...",
                               "Asistencia de %s" % nc.nombre_mes(periodo))

        def tarea(estado):
            estado["matriz"] = nc.matriz_asistencia(periodo)
            estado["resumen"] = nc.resumen_asistencia(*self._extremos(periodo))

        def terminar(estado):
            letrero.cerrar()
            if estado.get("error"):
                return self.aviso("❌ Error cargando asistencia: %s" % estado["error"], COLOR_ERROR)
            self.matriz_actual = estado.get("matriz") or {"dias": [], "filas": []}
            self.pintar_matriz()
            resumen = estado.get("resumen") or {}
            total_tardanzas = sum(d["tardanzas"] for d in resumen.values())
            total_faltas = sum(d["faltas"] for d in resumen.values())
            total_extra = sum(d["minutos_extra"] for d in resumen.values()) / 60.0
            self.lbl_resumen_asistencia.configure(
                text="%s → tardanzas: %d | faltas: %d | horas extra: %.1f h"
                     % (nc.nombre_mes(periodo), total_tardanzas, total_faltas, total_extra))
            self.aviso("✅ Asistencia de %s cargada" % nc.nombre_mes(periodo), COLOR_OK)
            self.cargar_panel()

        ejecutar_en_hilo(self.parent_frame, tarea, al_terminar=terminar)

    def _aviso_rango_marcaciones(self, periodo):
        """Avisa si el período que se va a calcular supera la última marcación cargada.

        Sin este aviso es fácil recalcular un mes que todavía no se ha exportado del
        reloj y ver el mes entero como faltas.
        """
        dias = nc.dias_del_periodo(periodo)
        if not dias:
            return ""
        fila = nc._consultar("SELECT MAX(fecha) FROM nom_marcaciones", uno=True)
        ultima = (fila[0] if fila else None) or ""
        if not ultima:
            return "\n⚠️ Todavía no hay marcaciones cargadas: todo el período saldría como falta.\n"
        if ultima == "None":
            return ""
        if dias[-1].strftime("%Y-%m-%d") > ultima:
            return ("\n⚠️ ATENCIÓN: las marcaciones cargadas llegan solo hasta el %s.\n"
                    "   Los días posteriores de este período se calcularán como FALTA.\n"
                    % fecha_para_entry(ultima))
        return ""

    @staticmethod
    def _extremos(periodo):
        dias = nc.dias_del_periodo(periodo)
        return (dias[0], dias[-1]) if dias else (None, None)

    def pintar_matriz(self):
        """Dibuja la grilla mensual (una columna por día)."""
        for widget in self.contenedor_matriz.winfo_children():
            widget.destroy()
        dias = self.matriz_actual.get("dias") or []
        if not dias:
            ctk.CTkLabel(self.contenedor_matriz, text="Período sin días.", font=("Arial", 12)).pack()
            return
        columnas = [("dni", "DNI", 95, "w"), ("empleado", "Empleado", 250, "w")]
        for dia in dias:
            columnas.append((dia.strftime("%Y-%m-%d"),
                             "%d %s" % (dia.day, nc.DIAS_CORTOS[dia.weekday()]), 48, "center"))
        columnas += [("puntuales", "P", 45, "center"), ("tardanzas", "T", 45, "center"),
                     ("faltas", "F", 45, "center"), ("extra", "Extra", 70, "center")]
        tabla = crear_tabla(self.contenedor_matriz, columnas, alto=22)
        self.tabla_matriz = tabla
        for clave, color in ((nc.EST_PUNTUAL, "#eafaf1"), (nc.EST_TARDANZA, "#fef5e7"),
                             (nc.EST_FALTA, "#fdedec"), (nc.EST_INCOMPLETO, "#f4ecf7"),
                             (nc.EST_DESCANSO, "#f2f4f4"), (nc.EST_FERIADO, "#eaf2f8"),
                             (nc.EST_VACACIONES, "#e8f4fd"), (nc.EST_PERMISO, "#fef9e7")):
            tabla.tag_configure(clave, background=color)
        for fila in self.matriz_actual.get("filas") or []:
            valores = [fila["dni"], fila["empleado"]]
            cuenta = {"P": 0, "T": 0, "F": 0}
            minutos_extra = 0
            etiqueta_predominante = None
            for dia in dias:
                registro = fila["celdas"].get(dia.strftime("%Y-%m-%d"))
                if not registro:
                    valores.append("")
                    continue
                estado = registro.get("estado") or nc.EST_PENDIENTE
                codigo = CODIGO_ESTADO.get(estado, "?")
                if codigo in cuenta:
                    cuenta[codigo] += 1
                # El dia de descanso o feriado trabajado se paga como dia con recargo,
                # no como hora extra: no se suma en esta columna.
                if estado not in (nc.EST_DESCANSO_TRAB, nc.EST_FERIADO_TRAB):
                    minutos_extra += nc.a_int(registro.get("minutos_extra"))
                if estado == nc.EST_FALTA and etiqueta_predominante != nc.EST_FALTA:
                    etiqueta_predominante = nc.EST_FALTA
                elif estado == nc.EST_TARDANZA and etiqueta_predominante is None:
                    etiqueta_predominante = nc.EST_TARDANZA
                elif estado == nc.EST_INCOMPLETO and etiqueta_predominante is None:
                    etiqueta_predominante = nc.EST_INCOMPLETO
                elif estado == nc.EST_PUNTUAL and etiqueta_predominante is None:
                    etiqueta_predominante = nc.EST_PUNTUAL
                marca = codigo
                if registro.get("hora_entrada"):
                    marca = "%s %s" % (codigo, registro["hora_entrada"])
                valores.append(marca)
            valores += [cuenta["P"], cuenta["T"], cuenta["F"],
                        "%.1f h" % (minutos_extra / 60.0) if minutos_extra else ""]
            tabla.insert("", tk.END, values=valores,
                         tags=(etiqueta_predominante or nc.EST_DESCANSO,))
        self.tabla_matriz.bind("<<TreeviewSelect>>", self.al_seleccionar_fila_matriz)

    def al_seleccionar_fila_matriz(self, evento=None):
        seleccion = self.tabla_matriz.selection()
        if not seleccion:
            return
        dni = self.tabla_matriz.item(seleccion[0], "values")[0]
        for opcion in self.cmb_detalle_empleado.cget("values"):
            if opcion.split("|")[0].strip() == str(dni):
                self.cmb_detalle_empleado.set(opcion)
                break

    def recalcular_periodo_actual(self):
        periodo = self.periodo_seleccionado() if hasattr(self, "ent_periodo") else self.periodo_actual
        if not periodo:
            return
        aviso_rango = self._aviso_rango_marcaciones(periodo)
        if not messagebox.askyesno(
                "Recalcular asistencia",
                "Se recalculará la asistencia de %s para todos los empleados.\n\n"
                "• Se toman la PRIMERA y la ÚLTIMA marcación de cada día.\n"
                "• Se cruzan con el turno/horario asignado a cada empleado.\n"
                "• Las lecturas repetidas del mismo minuto cuentan como una sola.\n"
                "• Las correcciones manuales NO se sobrescriben.\n%s\n"
                "¿Continuar?" % (nc.nombre_mes(periodo), aviso_rango)):
            return
        self.aviso("⏳ Calculando asistencia de %s..." % nc.nombre_mes(periodo), COLOR_ALERTA)
        letrero = LetreroCarga(self.frame_main, "Calculando la asistencia...",
                               "Cruzando las marcaciones con los turnos de %s" % nc.nombre_mes(periodo))
        dias = nc.dias_del_periodo(periodo)

        def progreso(actual, total):
            try:
                self.lbl_resumen_asistencia.configure(text="Calculando %d de %d..." % (actual, total))
            except Exception:
                pass
            letrero.cambiar("Calculando la asistencia...", "Avance: %d de %d" % (actual, total))

        def tarea(estado):
            estado["resultado"] = nc.recalcular_asistencia(dias[0], dias[-1], usuario=self.usuario_activo,
                                                          progreso=progreso)

        def terminar(estado):
            letrero.cerrar()
            if estado.get("error"):
                self.aviso("❌ Error al calcular: %s" % estado["error"], COLOR_ERROR)
                return messagebox.showerror("Error", "No se pudo calcular la asistencia:\n%s"
                                            % estado["error"])
            resultado = estado.get("resultado") or {}
            if not resultado.get("ok"):
                return messagebox.showerror("Error", resultado.get("error") or "Cálculo fallido.")
            messagebox.showinfo(
                "Asistencia calculada",
                "Período: %s\n\n"
                "Días procesados: %d\n"
                "Puntuales: %d\n"
                "Tardanzas: %d\n"
                "Faltas: %d\n"
                "Incompletos (una sola marca): %d\n"
                "Descansos: %d   |   Feriados: %d\n"
                "Días justificados por incidencia: %d\n"
                "Trabajos en descanso/feriado: %d\n\n"
                "Minutos de tardanza: %d   |   Horas extra: %.2f h"
                % (resultado["periodo"], resultado["dias"], resultado["puntuales"],
                   resultado["tardanzas"], resultado["faltas"], resultado["incompletos"],
                   resultado["descansos"], resultado["feriados"], resultado["justificados"],
                   resultado["trabajados_descanso"], resultado["minutos_tardanza"],
                   resultado["minutos_extra"] / 60.0))
            self.cargar_matriz_asistencia()
            self.cargar_horas_extra()
            self.cargar_panel()

        ejecutar_en_hilo(self.parent_frame, tarea, al_terminar=terminar)

    def cargar_detalle_dia(self):
        periodo = self.periodo_seleccionado() if hasattr(self, "ent_periodo") else self.periodo_actual
        if not periodo:
            return
        dias = nc.dias_del_periodo(periodo)
        if not dias:
            return
        dni = self.dni_de_opcion(self.cmb_detalle_empleado.get())
        if not dni:
            return
        estados = None
        filtro = self.cmb_detalle_estado.get()
        if filtro and filtro != "Todos":
            estados = [filtro]
        registros = nc.listar_asistencia(dias[0], dias[-1], dnis=[dni], estados=estados)
        limpiar_tabla(self.tabla_detalle)
        por_fecha = {r["fecha"]: r for r in registros}
        for dia in dias:
            clave = dia.strftime("%Y-%m-%d")
            registro = por_fecha.get(clave)
            if not registro:
                if filtro and filtro != "Todos":
                    continue
                self.tabla_detalle.insert("", tk.END, values=(
                    clave, nc.DIAS_SEMANA[dia.weekday()], "", "", "", "", "", "", "", "", "SIN CALCULAR",
                    ""), tags=(nc.EST_PENDIENTE,))
                continue
            estado = registro.get("estado") or nc.EST_PENDIENTE
            self.tabla_detalle.insert("", tk.END, values=(
                registro["fecha"], nc.DIAS_SEMANA[dia.weekday()],
                registro.get("hora_entrada") or "", registro.get("hora_salida") or "",
                registro.get("n_marcas") or 0,
                nc.formato_duracion(registro.get("minutos_tardanza")) if registro.get("minutos_tardanza") else "",
                nc.formato_duracion(registro.get("minutos_anticipo")) if registro.get("minutos_anticipo") else "",
                nc.formato_duracion(registro.get("minutos_trabajados")),
                ("dia con recargo" if estado in (nc.EST_DESCANSO_TRAB, nc.EST_FERIADO_TRAB)
                 else (nc.formato_duracion(registro.get("minutos_extra"))
                       if registro.get("minutos_extra") else "")),
                ("dia con recargo" if estado in (nc.EST_DESCANSO_TRAB, nc.EST_FERIADO_TRAB)
                 else (nc.formato_duracion(registro.get("minutos_excedente"))
                       if registro.get("minutos_excedente") else "")),
                estado, registro.get("observacion") or ""), tags=(estado,))
        for estado, color in nc.COLORES_ESTADO.items():
            self.tabla_detalle.tag_configure(estado, foreground=color)

    def ver_marcas_del_dia(self, evento=None):
        seleccion = self.tabla_detalle.selection()
        if not seleccion:
            return
        valores = self.tabla_detalle.item(seleccion[0], "values")
        fecha, dia = valores[0], valores[1]
        dni = self.dni_de_opcion(self.cmb_detalle_empleado.get())
        marcas = nc.detalle_marcaciones_dia(dni, fecha)
        texto = "Empleado: %s\nDNI: %s\nFecha: %s (%s)\n\n" % (self.nombre_empleado(dni), dni, fecha, dia)
        if not marcas:
            texto += "No hay marcaciones del reloj registradas ese día."
        else:
            texto += "Marcaciones del reloj (%d):\n" % len(marcas)
            for hora, tipo, metodo, fuente in marcas:
                texto += "   • %s   %s   %s\n" % (hora, tipo or "", metodo or "")
        messagebox.showinfo("Marcaciones del día", texto)

    def corregir_dia_seleccionado(self):
        seleccion = self.tabla_detalle.selection()
        if not seleccion:
            return messagebox.showwarning("Atención", "Seleccione primero un día en la lista.")
        valores = self.tabla_detalle.item(seleccion[0], "values")
        fecha = valores[0]
        dni = self.dni_de_opcion(self.cmb_detalle_empleado.get())
        ventana = ctk.CTkToplevel(self.parent_frame)
        ventana.title("Corregir asistencia del día")
        ventana.geometry("520x430")
        ventana.transient(self.parent_frame.winfo_toplevel())
        ventana.grab_set()
        ctk.CTkLabel(ventana, text="✏️ CORRECCIÓN MANUAL DE ASISTENCIA", font=("Arial", 14, "bold"),
                     text_color=COLOR_PRIMARIO).pack(pady=(14, 2))
        ctk.CTkLabel(ventana, text="%s — %s" % (self.nombre_empleado(dni), fecha_para_entry(fecha)),
                     font=("Arial", 12)).pack(pady=(0, 10))
        formulario = ctk.CTkFrame(ventana, fg_color="transparent")
        formulario.pack(fill="x", padx=20)
        campos = {}
        for indice, (clave, etiqueta, valor) in enumerate([
                ("hora_entrada", "Hora de entrada (HH:MM)", valores[2]),
                ("hora_salida", "Hora de salida (HH:MM)", valores[3]),
                ("minutos_tardanza", "Minutos de tardanza", ""),
                ("minutos_anticipo", "Minutos de salida anticipada", ""),
                ("minutos_extra", "Minutos de hora extra", ""),
                ("minutos_falta", "Minutos no trabajados", "")]):
            ctk.CTkLabel(formulario, text=etiqueta + ":", font=("Arial", 11, "bold"),
                         anchor="w", width=240).grid(row=indice, column=0, sticky="w", pady=4)
            entrada = ctk.CTkEntry(formulario, width=180)
            entrada.grid(row=indice, column=1, sticky="w", pady=4)
            if valor:
                entrada.insert(0, str(valor))
            campos[clave] = entrada
        ctk.CTkLabel(formulario, text="Estado:", font=("Arial", 11, "bold"), anchor="w").grid(
            row=6, column=0, sticky="w", pady=4)
        cmb_estado = ctk.CTkComboBox(formulario, values=[nc.EST_JUSTIFICADO, nc.EST_PUNTUAL,
                                                        nc.EST_TARDANZA, nc.EST_FALTA,
                                                        nc.EST_DESCANSO, nc.EST_FERIADO,
                                                        nc.EST_PERMISO, nc.EST_VACACIONES],
                                     width=180, state="readonly")
        cmb_estado.set(valores[9] if valores[9] else nc.EST_JUSTIFICADO)
        cmb_estado.grid(row=6, column=1, sticky="w", pady=4)
        ctk.CTkLabel(formulario, text="Observación:", font=("Arial", 11, "bold"), anchor="w").grid(
            row=7, column=0, sticky="w", pady=4)
        ent_observacion = ctk.CTkEntry(formulario, width=250)
        ent_observacion.grid(row=7, column=1, sticky="w", pady=4)
        ent_observacion.insert(0, valores[10] if len(valores) > 10 else "")

        def guardar():
            datos = {clave: entrada.get().strip() for clave, entrada in campos.items()}
            datos["estado"] = cmb_estado.get()
            datos["observacion"] = ent_observacion.get().strip()
            ok, error = nc.registrar_asistencia_manual(dni, fecha, datos, self.usuario_activo)
            if not ok:
                return messagebox.showerror("Error", error, parent=ventana)
            ventana.destroy()
            self.aviso("✅ Día corregido manualmente", COLOR_OK)
            self.cargar_detalle_dia()
            self.cargar_matriz_asistencia()

        ctk.CTkButton(ventana, text="💾 Guardar corrección", width=220, height=36, fg_color=COLOR_OK,
                      hover_color="#1e8449", command=guardar).pack(pady=16)

    @con_letrero("Leyendo las incidencias...", "")
    def cargar_incidencias(self):
        dias = nc.dias_del_periodo(self.periodo_actual)
        incidencias = nc.listar_incidencias(desde=dias[0] if dias else None,
                                            hasta=dias[-1] if dias else None)
        limpiar_tabla(self.tabla_incidencias)
        for incidencia in incidencias:
            desde, hasta = incidencia["fecha_desde"], incidencia["fecha_hasta"]
            dias_cubiertos = len(nc.rango_fechas(desde, hasta))
            self.tabla_incidencias.insert("", tk.END, values=(
                incidencia["id"], incidencia["dni"], self.nombre_empleado(incidencia["dni"]),
                incidencia["tipo"], desde, hasta, dias_cubiertos,
                "SÍ" if incidencia.get("con_goce") else "NO", incidencia.get("motivo") or "",
                incidencia.get("aprobado_por") or ""))

    def al_seleccionar_incidencia(self, evento=None):
        seleccion = self.tabla_incidencias.selection()
        if not seleccion:
            return
        valores = self.tabla_incidencias.item(seleccion[0], "values")
        self.incidencia_id = valores[0]
        for opcion in self.cmb_inc_empleado.cget("values"):
            if opcion.split("|")[0].strip() == str(valores[1]):
                self.cmb_inc_empleado.set(opcion)
                break
        self.cmb_inc_tipo.set(valores[3])
        self.ent_inc_desde.delete(0, tk.END)
        self.ent_inc_desde.insert(0, fecha_para_entry(valores[4]))
        self.ent_inc_hasta.delete(0, tk.END)
        self.ent_inc_hasta.insert(0, fecha_para_entry(valores[5]))
        self.var_inc_goce.set(valores[7] == "SÍ")
        self.ent_inc_motivo.delete(0, tk.END)
        self.ent_inc_motivo.insert(0, valores[8])
        self.ent_inc_aprobado.delete(0, tk.END)
        self.ent_inc_aprobado.insert(0, valores[9])

    def limpiar_formulario_incidencia(self):
        self.incidencia_id = None
        self.ent_inc_desde.delete(0, tk.END)
        self.ent_inc_hasta.delete(0, tk.END)
        self.ent_inc_motivo.delete(0, tk.END)
        self.ent_inc_horas.delete(0, tk.END)
        self.ent_inc_aprobado.delete(0, tk.END)
        self.var_inc_goce.set(True)

    def guardar_incidencia(self):
        datos = {
            "id": self.incidencia_id,
            "dni": self.dni_de_opcion(self.cmb_inc_empleado.get()),
            "fecha_desde": parse_fecha_texto(self.ent_inc_desde.get()),
            "fecha_hasta": parse_fecha_texto(self.ent_inc_hasta.get() or self.ent_inc_desde.get()),
            "tipo": self.cmb_inc_tipo.get(),
            "motivo": self.ent_inc_motivo.get().strip(),
            "con_goce": self.var_inc_goce.get(),
            "horas": self.ent_inc_horas.get().strip() or None,
            "aprobado_por": self.ent_inc_aprobado.get().strip(),
        }
        if not datos["dni"]:
            return messagebox.showwarning("Atención", "Seleccione el empleado.")
        nuevo_id, error = nc.guardar_incidencia(datos, self.usuario_activo)
        if not nuevo_id:
            return messagebox.showerror("Error", error)
        self.limpiar_formulario_incidencia()
        self.cargar_incidencias()
        self.aviso("✅ Incidencia guardada. Recalcule la asistencia para aplicarla.", COLOR_OK)
        if messagebox.askyesno("Incidencia guardada",
                               "¿Recalcular ahora la asistencia del mes para aplicar la incidencia?"):
            self.recalcular_periodo_actual()

    def eliminar_incidencia(self):
        if not self.incidencia_id:
            return messagebox.showwarning("Atención", "Seleccione una incidencia de la lista.")
        if not messagebox.askyesno("Confirmar", "¿Eliminar la incidencia seleccionada?"):
            return
        nc.eliminar_incidencia(self.incidencia_id, self.usuario_activo)
        self.limpiar_formulario_incidencia()
        self.cargar_incidencias()

    @con_letrero("Leyendo las horas extra...", "")
    def cargar_horas_extra(self):
        dias = nc.dias_del_periodo(self.periodo_actual)
        registros = nc.listar_horas_extra(desde=dias[0] if dias else None, hasta=dias[-1] if dias else None)
        limpiar_tabla(self.tabla_horas_extra)
        for registro in registros:
            self.tabla_horas_extra.insert("", tk.END, values=(
                registro["id"], registro["dni"], self.nombre_empleado(registro["dni"]),
                registro["fecha"], registro["minutos"],
                "%.2f" % (nc.a_int(registro["minutos"]) / 60.0),
                nc.ETIQUETA_HORA_EXTRA.get(registro["tipo"], registro["tipo"]),
                registro.get("origen") or "", "SÍ" if registro.get("aprobado") else "NO",
                registro.get("observacion") or ""))

    def guardar_hora_extra(self):
        datos = {
            "dni": self.dni_de_opcion(self.cmb_he_empleado.get()),
            "fecha": parse_fecha_texto(self.ent_he_fecha.get()),
            "minutos": parse_num(self.ent_he_minutos.get()),
            "tipo": self.cmb_he_tipo.get(),
            "aprobado": self.var_he_aprobado.get(),
            "origen": "MANUAL",
            "observacion": "Registrado manualmente",
        }
        if not datos["dni"] or not datos["fecha"]:
            return messagebox.showwarning("Atención", "Indique el empleado y la fecha.")
        ok, error = nc.guardar_hora_extra(datos, self.usuario_activo)
        if not ok:
            return messagebox.showerror("Error", error)
        self.ent_he_minutos.delete(0, tk.END)
        self.cargar_horas_extra()
        self.aviso("✅ Hora extra registrada", COLOR_OK)

    def eliminar_hora_extra(self):
        seleccion = self.tabla_horas_extra.selection()
        if not seleccion:
            return messagebox.showwarning("Atención", "Seleccione una hora extra de la lista.")
        identificador = self.tabla_horas_extra.item(seleccion[0], "values")[0]
        if not messagebox.askyesno("Confirmar", "¿Eliminar la hora extra seleccionada?"):
            return
        nc.eliminar_hora_extra(identificador, self.usuario_activo)
        self.cargar_horas_extra()

    def exportar_asistencia(self):
        periodo = self.periodo_seleccionado() if hasattr(self, "ent_periodo") else self.periodo_actual
        if not periodo:
            return
        ruta = guardar_archivo_dialogo("Exportar asistencia del período", ".xlsx",
                                       "Asistencia_%s.xlsx" % periodo,
                                       [("Archivos Excel", "*.xlsx")])
        if not ruta:
            return
        ok, error = nc.exportar_asistencia_excel(ruta, periodo)
        if not ok:
            return messagebox.showerror("Error", "No se pudo exportar:\n%s" % error)
        registrar_auditoria(self.usuario_activo, "Nómina", "Exportó la asistencia de %s" % periodo)
        if messagebox.askyesno("Exportación lista", "Reporte guardado en:\n%s\n\n¿Desea abrirlo?" % ruta):
            abrir_documento(ruta)

    # =====================================================
    # PESTAÑA: PERSONAL (EMPLEADOS, TURNOS, HORARIOS Y ASIGNACIONES)
    # =====================================================
    def construir_tab_personal(self):
        tab = self.tab_personal
        self.sub_personal = ctk.CTkSegmentedButton(
            tab, values=["Empleados", "Turnos", "Horarios", "Asignaciones"],
            selected_color=COLOR_PRIMARIO, selected_hover_color=COLOR_HOVER,
            command=self.cambiar_sub_personal)
        self.sub_personal.pack(fill="x", padx=10, pady=(10, 6))
        self.frame_empleados = ctk.CTkFrame(tab, fg_color="transparent")
        self.frame_turnos = ctk.CTkFrame(tab, fg_color="transparent")
        self.frame_horarios = ctk.CTkFrame(tab, fg_color="transparent")
        self.frame_asignaciones = ctk.CTkFrame(tab, fg_color="transparent")
        self._construir_empleados(self.frame_empleados)
        self._construir_turnos(self.frame_turnos)
        self._construir_horarios(self.frame_horarios)
        self._construir_asignaciones(self.frame_asignaciones)
        self.sub_personal.set("Empleados")
        self.cambiar_sub_personal("Empleados")

    def cambiar_sub_personal(self, valor):
        for frame in (self.frame_empleados, self.frame_turnos, self.frame_horarios,
                      self.frame_asignaciones):
            frame.pack_forget()
        mapa = {"Empleados": self.frame_empleados, "Turnos": self.frame_turnos,
                "Horarios": self.frame_horarios, "Asignaciones": self.frame_asignaciones}
        frame = mapa.get(valor, self.frame_empleados)
        frame.pack(fill="both", expand=True)
        if valor == "Empleados":
            self.cargar_empleados()
        elif valor == "Turnos":
            self.cargar_turnos()
        elif valor == "Horarios":
            self.cargar_horarios()
        else:
            self.cargar_asignaciones()

    # ---------- EMPLEADOS ----------
    def _construir_empleados(self, padre):
        barra = ctk.CTkFrame(padre, fg_color="transparent")
        barra.pack(fill="x", pady=(2, 4))
        ctk.CTkLabel(barra, text="👥 PADRÓN DE EMPLEADOS (tomado de la tabla «Choferes»)",
                     font=("Arial", 13, "bold"), text_color=COLOR_PRIMARIO).pack(side="left")
        ctk.CTkButton(barra, text="🔄 Sincronizar desde Choferes", width=210, fg_color="#34495e",
                      hover_color="#2c3e50", command=self.sincronizar_empleados).pack(side="right", padx=4)
        ctk.CTkButton(barra, text="💵 Sueldos y garantías", width=190, fg_color="#8e44ad",
                      hover_color="#6c3483", font=("Arial", 11, "bold"),
                      command=self.abrir_sueldos).pack(side="right", padx=4)
        contenedor = ctk.CTkFrame(padre, fg_color="transparent")
        contenedor.pack(fill="both", expand=True)
        self.tabla_empleados = crear_tabla(contenedor, [
            ("dni", "DNI", 100, "w"),
            ("nombre", "Empleado", 250, "w"),
            ("cargo", "Cargo", 140, "w"),
            ("sueldo", "Sueldo básico", 110, "e"),
            ("jornada", "Jornada", 80, "center"),
            ("pension", "Pensión", 100, "center"),
            ("asignacion", "Asig. fam.", 80, "center"),
            ("horario", "Horario", 180, "w"),
            ("turno", "Turno", 190, "w"),
            ("estado", "Estado", 100, "center"),
        ], alto=12)
        self.tabla_empleados.bind("<<TreeviewSelect>>", self.al_seleccionar_empleado)

        # La ficha usa altura fija con desplazamiento: así la tabla conserva su espacio y
        # TODOS los campos y botones quedan alcanzables aunque la ventana sea baja.
        formulario = ctk.CTkScrollableFrame(padre, fg_color=COLOR_FONDO_SUAVE, corner_radius=10,
                                            border_width=1, border_color="#e0e0e0", height=250)
        formulario.pack(fill="x", pady=6)
        # Se respeta la altura indicada (si no, la ficha crece con su contenido y deja
        # a la tabla sin espacio). El contenido que no cabe se alcanza desplazando.
        try:
            formulario.pack_propagate(False)
        except Exception:
            pass
        # Encabezado en una sola línea (antes ocupaba dos y robaba altura a la tabla)
        cabecera_ficha = ctk.CTkFrame(formulario, fg_color="transparent")
        cabecera_ficha.pack(fill="x", padx=12, pady=(2, 2))
        ctk.CTkLabel(cabecera_ficha, text="📝 FICHA DE NÓMINA —", font=("Arial", 12, "bold"),
                     text_color=COLOR_PRIMARIO).pack(side="left")
        self.lbl_empleado_sel = ctk.CTkLabel(cabecera_ficha, text="Seleccione un empleado en la lista",
                                             font=("Arial", 11), text_color=COLOR_GRIS)
        self.lbl_empleado_sel.pack(side="left", padx=6)
        # Guardar y limpiar en el ENCABEZADO: quedan visibles aunque la ficha se desplace
        ctk.CTkButton(cabecera_ficha, text="🧹 Limpiar", width=100, fg_color=COLOR_GRIS,
                      hover_color="#606b6b", command=self.limpiar_ficha_empleado).pack(side="right", padx=4)
        ctk.CTkButton(cabecera_ficha, text="💾 GUARDAR FICHA", width=185, height=28,
                      font=("Arial", 11, "bold"), fg_color=COLOR_OK, hover_color="#1e8449",
                      command=self.guardar_ficha_empleado).pack(side="right", padx=4)
        grid = ctk.CTkFrame(formulario, fg_color="transparent")
        grid.pack(fill="x", padx=12, pady=6)
        self.campos_empleado = {}
        # Etiquetas cortas a propósito: con las largas, la tercera columna se salía 18 px
        # del ancho de la ventana (los campos de la derecha quedaban cortados).
        definicion = [
            ("sueldo_basico", "Sueldo S/", 130),
            ("jornada_horas", "Jornada h", 130),
            ("cargo", "Cargo", 130),
            ("fecha_ingreso", "Ingreso", 130),
            ("banco_haberes", "Banco", 130),
            ("cuenta_haberes", "Cuenta / CCI", 130),
            ("essalud_codigo", "ESSALUD", 130),
            ("afp_nombre", "AFP", 130),
            ("afp_comision_pct", "Comisión AFP %", 130),
            ("cuspp", "CUSPP", 130),
            ("otros_ingresos", "Otros ingresos", 130),
            ("otros_descuentos", "Otros desc. S/", 130),
            ("adelanto_mensual", "Adelanto S/", 130),
            ("retencion_judicial_pct", "Retención jud. %", 130),
            ("garantia_mensual", "Garantía S/", 130),
        ]
        for indice, (clave, etiqueta, ancho) in enumerate(definicion):
            columna = (indice % 3) * 2
            fila = indice // 3
            ctk.CTkLabel(grid, text=etiqueta + ":", font=("Arial", 10, "bold"), anchor="w",
                         width=100).grid(row=fila, column=columna, sticky="w", padx=3, pady=2)
            entrada = ctk.CTkEntry(grid, width=ancho)
            entrada.grid(row=fila, column=columna + 1, sticky="w", padx=3, pady=2)
            self.campos_empleado[clave] = entrada
        fecha_grid = self.campos_empleado["fecha_ingreso"]
        ctk.CTkButton(grid, text="📅", width=30, fg_color=COLOR_PRIMARIO, hover_color=COLOR_HOVER,
                      command=lambda: self.abrir_calendario(self.campos_empleado["fecha_ingreso"])
                      ).grid(row=0, column=3, sticky="w", padx=(0, 4))

        # Dos filas de combos: en una sola se salían del ancho de la ventana
        combos = ctk.CTkFrame(formulario, fg_color="transparent")
        combos.pack(fill="x", padx=12, pady=(2, 2))
        ctk.CTkLabel(combos, text="Régimen:", font=("Arial", 10, "bold")).pack(side="left")
        self.cmb_emp_regimen = ctk.CTkComboBox(combos, values=["GENERAL", "MYPE_PEQUENA", "MYPE_MICRO"],
                                               width=140, state="readonly")
        self.cmb_emp_regimen.set("GENERAL")
        self.cmb_emp_regimen.pack(side="left", padx=(2, 10))
        ctk.CTkLabel(combos, text="Sistema de pensión:", font=("Arial", 10, "bold")).pack(side="left")
        self.cmb_emp_pension = ctk.CTkComboBox(combos, values=["ONP", "AFP", "NINGUNO"], width=110,
                                               state="readonly")
        self.cmb_emp_pension.set("ONP")
        self.cmb_emp_pension.pack(side="left", padx=(2, 10))
        ctk.CTkLabel(combos, text="Estado:", font=("Arial", 10, "bold")).pack(side="left")
        self.cmb_emp_estado = ctk.CTkComboBox(combos, values=["ACTIVO", "CESADO", "SUSPENDIDO"],
                                              width=120, state="readonly")
        self.cmb_emp_estado.set("ACTIVO")
        self.cmb_emp_estado.pack(side="left", padx=(2, 10))

        asignacion = ctk.CTkFrame(formulario, fg_color="transparent")
        asignacion.pack(fill="x", padx=12, pady=(0, 2))
        ctk.CTkLabel(asignacion, text="Horario asignado:", font=("Arial", 10, "bold")).pack(side="left")
        self.cmb_emp_horario = ctk.CTkComboBox(asignacion, values=["(Ninguno)"], width=250,
                                               state="readonly")
        self.cmb_emp_horario.pack(side="left", padx=(2, 12))
        self.registrar_combo_horario(self.cmb_emp_horario)
        ctk.CTkLabel(asignacion, text="Turno asignado:", font=("Arial", 10, "bold")).pack(side="left")
        self.cmb_emp_turno = ctk.CTkComboBox(asignacion, values=["(Descanso)"], width=250,
                                             state="readonly")
        self.cmb_emp_turno.pack(side="left", padx=(2, 10))
        self.registrar_combo_turno(self.cmb_emp_turno)

        checks = ctk.CTkFrame(formulario, fg_color="transparent")
        checks.pack(fill="x", padx=12, pady=(2, 4))
        self.var_emp_asignacion = tk.BooleanVar(value=False)
        self.var_emp_discapacidad = tk.BooleanVar(value=False)
        self.var_emp_confianza = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(checks, text="Asignación familiar (10% RMV)", variable=self.var_emp_asignacion,
                        font=("Arial", 11)).pack(side="left", padx=6)
        ctk.CTkCheckBox(checks, text="Discapacidad", variable=self.var_emp_discapacidad,
                        font=("Arial", 11)).pack(side="left", padx=6)
        ctk.CTkCheckBox(checks, text="Personal de confianza / dirección",
                        variable=self.var_emp_confianza, font=("Arial", 11)).pack(side="left", padx=6)

        # La asignación masiva va en su propia fila (antes se salía del ancho de la ventana)
        masiva = ctk.CTkFrame(formulario, fg_color="transparent")
        # Se coloca arriba (justo bajo el encabezado) porque es una acción, no un dato
        masiva.pack(fill="x", padx=12, pady=(0, 6), before=grid)
        ctk.CTkLabel(masiva, text="🎯 Asignación masiva →", font=("Arial", 11, "bold"),
                     text_color=COLOR_PRIMARIO).pack(side="left", padx=(0, 6))
        ctk.CTkLabel(masiva, text="Horario:", font=("Arial", 10, "bold")).pack(side="left")
        self.cmb_masivo_horario = ctk.CTkComboBox(masiva, values=["(Ninguno)"], width=190,
                                                  state="readonly")
        self.cmb_masivo_horario.pack(side="left", padx=(2, 8))
        self.registrar_combo_horario(self.cmb_masivo_horario)
        ctk.CTkLabel(masiva, text="Turno:", font=("Arial", 10, "bold")).pack(side="left")
        self.cmb_masivo_turno = ctk.CTkComboBox(masiva, values=["(Descanso)"], width=190,
                                                state="readonly")
        self.cmb_masivo_turno.pack(side="left", padx=(2, 8))
        self.registrar_combo_turno(self.cmb_masivo_turno)
        ctk.CTkButton(masiva, text="🎯 Aplicar a TODOS", width=170, fg_color="#8e44ad",
                      hover_color="#6c3483", command=self.asignacion_masiva).pack(side="left")
        # La ficha se ancla al fondo con su altura fija; la tabla ocupa el resto y se desplaza
        contenedor.pack_forget()
        formulario.pack_forget()
        formulario.pack(side="bottom", fill="x", pady=6)
        contenedor.pack(side="top", fill="both", expand=True)

    @con_letrero("Leyendo la base de datos...", "Padrón de empleados")
    def cargar_empleados(self):
        limpiar_tabla(self.tabla_empleados)
        horarios = {int(h["id"]): h["nombre"] for h in self.horarios}
        turnos = {int(t["id"]): t["nombre"] for t in self.turnos}
        for empleado in self.empleados:
            self.tabla_empleados.insert("", tk.END, values=(
                empleado["dni"], empleado["nombre"], empleado.get("cargo") or "",
                formatear_numero(empleado.get("sueldo_basico"), 2),
                formatear_numero(empleado.get("jornada_horas"), 1),
                empleado.get("sistema_pension") or "",
                "SÍ" if empleado.get("asignacion_familiar") else "NO",
                horarios.get(empleado.get("horario_id"), ""),
                turnos.get(empleado.get("turno_id"), ""),
                empleado.get("estado") or "ACTIVO"))

    def al_seleccionar_empleado(self, evento=None):
        seleccion = self.tabla_empleados.selection()
        if not seleccion:
            return
        dni = self.tabla_empleados.item(seleccion[0], "values")[0]
        empleado = None
        for registro in self.empleados:
            if str(registro["dni"]) == str(dni):
                empleado = registro
                break
        if not empleado:
            return
        self.empleado_seleccionado = empleado
        self.lbl_empleado_sel.configure(text="%s — DNI %s" % (empleado["nombre"], empleado["dni"]),
                                        text_color="#1a1a1a")
        for clave, entrada in self.campos_empleado.items():
            entrada.delete(0, tk.END)
            valor = empleado.get(clave)
            if clave in ("sueldo_basico", "jornada_horas", "otros_ingresos", "otros_descuentos",
                         "adelanto_mensual", "afp_comision_pct", "retencion_judicial_pct",
                         "garantia_mensual"):
                if nc.a_float(valor):
                    entrada.insert(0, formatear_numero(valor, 2))
            elif clave == "fecha_ingreso":
                entrada.insert(0, fecha_para_entry(valor) if valor else "")
            elif valor:
                entrada.insert(0, str(valor))
        self.cmb_emp_regimen.set(empleado.get("regimen") or "GENERAL")
        self.cmb_emp_pension.set(empleado.get("sistema_pension") or "ONP")
        self.cmb_emp_estado.set(empleado.get("estado") or "ACTIVO")
        self.var_emp_asignacion.set(bool(empleado.get("asignacion_familiar")))
        self.var_emp_discapacidad.set(bool(empleado.get("discapacidad")))
        self.var_emp_confianza.set(bool(empleado.get("confianza")))
        for combo, identificador in ((self.cmb_emp_horario, empleado.get("horario_id")),
                                     (self.cmb_masivo_horario, None)):
            if identificador is None:
                continue
            for opcion in combo.cget("values"):
                if opcion.split("|")[0].strip() == str(identificador):
                    combo.set(opcion)
                    break
        if not empleado.get("horario_id"):
            self.cmb_emp_horario.set("(Ninguno)")
        for combo, identificador in ((self.cmb_emp_turno, empleado.get("turno_id")),
                                     (self.cmb_masivo_turno, None)):
            if identificador is None:
                continue
            for opcion in combo.cget("values"):
                if opcion.split("|")[0].strip() == str(identificador):
                    combo.set(opcion)
                    break
        if not empleado.get("turno_id"):
            self.cmb_emp_turno.set("(Descanso)")

    def limpiar_ficha_empleado(self):
        self.empleado_seleccionado = None
        self.lbl_empleado_sel.configure(text="Seleccione un empleado en la lista", text_color=COLOR_GRIS)
        for entrada in self.campos_empleado.values():
            entrada.delete(0, tk.END)
        self.var_emp_asignacion.set(False)
        self.var_emp_discapacidad.set(False)
        self.var_emp_confianza.set(False)

    def guardar_ficha_empleado(self):
        empleado = getattr(self, "empleado_seleccionado", None)
        if not empleado:
            return messagebox.showwarning("Atención", "Seleccione un empleado de la lista.")
        datos = dict(empleado)
        datos["sueldo_basico"] = parse_num(self.campos_empleado["sueldo_basico"].get())
        datos["jornada_horas"] = parse_num(self.campos_empleado["jornada_horas"].get(), 8) or 8
        datos["cargo"] = self.campos_empleado["cargo"].get().strip()
        datos["fecha_ingreso"] = parse_fecha_texto(self.campos_empleado["fecha_ingreso"].get()) or ""
        datos["banco_haberes"] = self.campos_empleado["banco_haberes"].get().strip()
        datos["cuenta_haberes"] = self.campos_empleado["cuenta_haberes"].get().strip()
        datos["essalud_codigo"] = self.campos_empleado["essalud_codigo"].get().strip()
        datos["afp_nombre"] = self.campos_empleado["afp_nombre"].get().strip()
        datos["afp_comision_pct"] = parse_num(self.campos_empleado["afp_comision_pct"].get())
        datos["cuspp"] = self.campos_empleado["cuspp"].get().strip()
        datos["otros_ingresos"] = parse_num(self.campos_empleado["otros_ingresos"].get())
        datos["otros_descuentos"] = parse_num(self.campos_empleado["otros_descuentos"].get())
        datos["adelanto_mensual"] = parse_num(self.campos_empleado["adelanto_mensual"].get())
        datos["retencion_judicial_pct"] = parse_num(self.campos_empleado["retencion_judicial_pct"].get())
        datos["garantia_mensual"] = parse_num(self.campos_empleado["garantia_mensual"].get())
        datos["regimen"] = self.cmb_emp_regimen.get()
        datos["sistema_pension"] = self.cmb_emp_pension.get()
        datos["estado"] = self.cmb_emp_estado.get()
        datos["asignacion_familiar"] = self.var_emp_asignacion.get()
        datos["discapacidad"] = self.var_emp_discapacidad.get()
        datos["confianza"] = self.var_emp_confianza.get()
        horario = self.cmb_emp_horario.get()
        datos["horario_id"] = int(horario.split("|")[0].strip()) if horario != "(Ninguno)" else None
        turno = self.cmb_emp_turno.get()
        datos["turno_id"] = int(turno.split("|")[0].strip()) if turno != "(Descanso)" else None
        if datos["sistema_pension"] == "AFP" and not datos["afp_nombre"]:
            messagebox.showwarning("Atención", "Indique el nombre de la AFP o cambie el sistema de pensión.")
            return
        ok, error = nc.guardar_empleado_nomina(empleado["dni"], datos, self.usuario_activo)
        if not ok:
            return messagebox.showerror("Error", "No se pudo guardar la ficha:\n%s" % error)
        self.aviso("✅ Ficha de %s actualizada" % empleado["nombre"], COLOR_OK)
        self.recargar_catalogos()

    def asignacion_masiva(self):
        """Aplica el horario o turno elegido a todo el personal activo."""
        horario = self.cmb_masivo_horario.get()
        turno = self.cmb_masivo_turno.get()
        horario_id = int(horario.split("|")[0].strip()) if horario != "(Ninguno)" else None
        turno_id = int(turno.split("|")[0].strip()) if turno != "(Descanso)" else None
        if not horario_id and not turno_id:
            return messagebox.showwarning("Atención", "Elija un horario o un turno para asignar.")
        destino = [(horario or "(Ninguno)", turno or "(Descanso)")]
        activos = [e for e in self.empleados if nc.normalizar_texto(e.get("estado")) != "CESADO"]
        if not messagebox.askyesno(
                "Asignación masiva",
                "¿Asignar a los %d empleados activos?\n\nHorario: %s\nTurno: %s\n\n"
                "Esto reemplaza la asignación individual actual." % (len(activos), destino[0][0],
                                                                     destino[0][1])):
            return
        guardados = 0
        for empleado in activos:
            datos = dict(empleado)
            if horario_id:
                datos["horario_id"] = horario_id
            if turno_id:
                datos["turno_id"] = turno_id
            ok, _ = nc.guardar_empleado_nomina(empleado["dni"], datos, self.usuario_activo)
            guardados += 1 if ok else 0
        self.aviso("✅ Asignación masiva aplicada a %d empleados" % guardados, COLOR_OK)
        messagebox.showinfo("Asignación masiva", "Se actualizaron %d empleados.\n\n"
                                                "Recalcule la asistencia para aplicar los cambios." % guardados)
        self.recargar_catalogos()

    def abrir_sueldos(self):
        """Abre la pantalla de carga masiva de sueldos, garantías y jornada."""
        if not self.empleados:
            return messagebox.showwarning("Atención", "No hay empleados en el padrón. Use "
                                                      "«Sincronizar desde Choferes» primero.")
        DialogoSueldos(self.parent_frame, self, self.empleados, self.usuario_activo)

    def sincronizar_empleados(self):
        cantidad = nc.sincronizar_empleados(self.usuario_activo)
        self.recargar_catalogos()
        messagebox.showinfo("Padrón sincronizado",
                            "Se revisó la tabla de choferes y se crearon/actualizaron las fichas de nómina.")

    # ---------- TURNOS ----------
    def _construir_turnos(self, padre):
        barra = ctk.CTkFrame(padre, fg_color="transparent")
        barra.pack(fill="x", pady=(2, 4))
        ctk.CTkLabel(barra, text="🕒 TURNOS DE TRABAJO", font=("Arial", 13, "bold"),
                     text_color=COLOR_PRIMARIO).pack(side="left")
        ctk.CTkButton(barra, text="➕ Nuevo turno", width=150, fg_color=COLOR_OK, hover_color="#1e8449",
                      command=self.nuevo_turno).pack(side="right", padx=4)
        contenedor = ctk.CTkFrame(padre, fg_color="transparent")
        contenedor.pack(fill="both", expand=True)
        self.tabla_turnos = crear_tabla(contenedor, [
            ("id", "ID", 50, "center"),
            ("nombre", "Turno", 300, "w"),
            ("entrada", "Entrada", 80, "center"),
            ("salida", "Salida", 80, "center"),
            ("jornada", "Jornada", 80, "center"),
            ("tolerancia", "Tolerancia", 85, "center"),
            ("refrigerio", "Refrigerio", 85, "center"),
            ("dias", "Días", 220, "w"),
            ("nocturno", "Nocturno", 85, "center"),
            ("activo", "Activo", 70, "center"),
        ], alto=10)
        self.tabla_turnos.bind("<<TreeviewSelect>>", self.al_seleccionar_turno)

        formulario = ctk.CTkFrame(padre, fg_color=COLOR_FONDO_SUAVE, corner_radius=10, border_width=1,
                                  border_color="#e0e0e0")
        formulario.pack(fill="x", pady=6)
        ctk.CTkLabel(formulario, text="📝 DATOS DEL TURNO", font=("Arial", 12, "bold"),
                     text_color=COLOR_PRIMARIO).pack(anchor="w", padx=12, pady=(8, 2))
        fila1 = ctk.CTkFrame(formulario, fg_color="transparent")
        fila1.pack(fill="x", padx=12, pady=3)
        ctk.CTkLabel(fila1, text="Nombre:", font=("Arial", 10, "bold")).pack(side="left")
        self.ent_turno_nombre = ctk.CTkEntry(fila1, width=300)
        self.ent_turno_nombre.pack(side="left", padx=(2, 10))
        ctk.CTkLabel(fila1, text="Entrada:", font=("Arial", 10, "bold")).pack(side="left")
        self.ent_turno_entrada = ctk.CTkEntry(fila1, width=70)
        self.ent_turno_entrada.insert(0, "08:00")
        self.ent_turno_entrada.pack(side="left", padx=(2, 10))
        ctk.CTkLabel(fila1, text="Salida:", font=("Arial", 10, "bold")).pack(side="left")
        self.ent_turno_salida = ctk.CTkEntry(fila1, width=70)
        self.ent_turno_salida.insert(0, "17:00")
        self.ent_turno_salida.pack(side="left", padx=(2, 10))
        ctk.CTkLabel(fila1, text="Jornada (h):", font=("Arial", 10, "bold")).pack(side="left")
        self.ent_turno_jornada = ctk.CTkEntry(fila1, width=60)
        self.ent_turno_jornada.insert(0, "8")
        self.ent_turno_jornada.pack(side="left", padx=(2, 10))
        fila2 = ctk.CTkFrame(formulario, fg_color="transparent")
        fila2.pack(fill="x", padx=12, pady=3)
        ctk.CTkLabel(fila2, text="Tolerancia (min):", font=("Arial", 10, "bold")).pack(side="left")
        self.ent_turno_tolerancia = ctk.CTkEntry(fila2, width=60)
        self.ent_turno_tolerancia.insert(0, "10")
        self.ent_turno_tolerancia.pack(side="left", padx=(2, 10))
        ctk.CTkLabel(fila2, text="Refrigerio (min):", font=("Arial", 10, "bold")).pack(side="left")
        self.ent_turno_refrigerio = ctk.CTkEntry(fila2, width=60)
        self.ent_turno_refrigerio.insert(0, "45")
        self.ent_turno_refrigerio.pack(side="left", padx=(2, 10))
        ctk.CTkLabel(fila2, text="Color:", font=("Arial", 10, "bold")).pack(side="left")
        self.ent_turno_color = ctk.CTkEntry(fila2, width=90)
        self.ent_turno_color.insert(0, COLOR_PRIMARIO)
        self.ent_turno_color.pack(side="left", padx=(2, 10))
        self.var_turno_cruza = tk.BooleanVar(value=False)
        self.var_turno_nocturno = tk.BooleanVar(value=False)
        self.var_turno_activo = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(fila2, text="Cruza medianoche", variable=self.var_turno_cruza,
                        font=("Arial", 10)).pack(side="left", padx=4)
        ctk.CTkCheckBox(fila2, text="Nocturno", variable=self.var_turno_nocturno,
                        font=("Arial", 10)).pack(side="left", padx=4)
        ctk.CTkCheckBox(fila2, text="Activo", variable=self.var_turno_activo,
                        font=("Arial", 10)).pack(side="left", padx=4)
        fila3 = ctk.CTkFrame(formulario, fg_color="transparent")
        fila3.pack(fill="x", padx=12, pady=3)
        ctk.CTkLabel(fila3, text="Días que aplica:", font=("Arial", 10, "bold")).pack(side="left")
        self.vars_turno_dias = {}
        for indice, dia in enumerate(nc.DIAS_CORTOS, start=1):
            variable = tk.BooleanVar(value=indice <= 6)
            self.vars_turno_dias[indice] = variable
            ctk.CTkCheckBox(fila3, text=dia, variable=variable, font=("Arial", 10),
                            width=60).pack(side="left", padx=2)
        ctk.CTkLabel(fila3, text="Observación:", font=("Arial", 10, "bold")).pack(side="left", padx=(10, 2))
        self.ent_turno_observacion = ctk.CTkEntry(fila3, width=240)
        self.ent_turno_observacion.pack(side="left")
        acciones = ctk.CTkFrame(formulario, fg_color="transparent")
        acciones.pack(fill="x", padx=12, pady=(6, 10))
        ctk.CTkButton(acciones, text="💾 Guardar turno", width=160, fg_color=COLOR_OK,
                      hover_color="#1e8449", command=self.guardar_turno).pack(side="left")
        ctk.CTkButton(acciones, text="🗑️ Eliminar turno", width=160, fg_color=COLOR_ERROR,
                      hover_color="#922b21", command=self.eliminar_turno).pack(side="left", padx=6)
        # El formulario se ancla al fondo: sus botones siempre visibles
        contenedor.pack_forget()
        formulario.pack_forget()
        formulario.pack(side="bottom", fill="x", pady=6)
        contenedor.pack(side="top", fill="both", expand=True)
        self.turno_id = None

    def nuevo_turno(self):
        self.turno_id = None
        self.ent_turno_nombre.delete(0, tk.END)
        self.ent_turno_entrada.delete(0, tk.END)
        self.ent_turno_entrada.insert(0, "08:00")
        self.ent_turno_salida.delete(0, tk.END)
        self.ent_turno_salida.insert(0, "17:00")
        self.ent_turno_observacion.delete(0, tk.END)

    @con_letrero("Leyendo los turnos...", "")
    def cargar_turnos(self):
        limpiar_tabla(self.tabla_turnos)
        for turno in self.turnos:
            self.tabla_turnos.insert("", tk.END, values=(
                turno["id"], turno["nombre"], turno["hora_entrada"], turno["hora_salida"],
                formatear_numero(turno["horas_jornada"], 1), turno["tolerancia_min"],
                turno["refrigerio_min"], nc.dias_semana_texto(nc.parsear_dias_semana(turno["dias_semana"])),
                "SÍ" if turno.get("aplica_nocturno") else "NO", "SÍ" if turno.get("activo") else "NO"))

    def al_seleccionar_turno(self, evento=None):
        seleccion = self.tabla_turnos.selection()
        if not seleccion:
            return
        identificador = self.tabla_turnos.item(seleccion[0], "values")[0]
        for turno in self.turnos:
            if str(turno["id"]) == str(identificador):
                self.turno_id = turno["id"]
                self.ent_turno_nombre.delete(0, tk.END)
                self.ent_turno_nombre.insert(0, turno["nombre"])
                for entrada, valor in ((self.ent_turno_entrada, turno["hora_entrada"]),
                                       (self.ent_turno_salida, turno["hora_salida"]),
                                       (self.ent_turno_jornada, turno["horas_jornada"]),
                                       (self.ent_turno_tolerancia, turno["tolerancia_min"]),
                                       (self.ent_turno_refrigerio, turno["refrigerio_min"]),
                                       (self.ent_turno_color, turno["color"]),
                                       (self.ent_turno_observacion, turno["observacion"])):
                    entrada.delete(0, tk.END)
                    entrada.insert(0, "" if valor is None else str(valor))
                self.var_turno_cruza.set(bool(turno.get("cruza_medianoche")))
                self.var_turno_nocturno.set(bool(turno.get("aplica_nocturno")))
                self.var_turno_activo.set(bool(turno.get("activo")))
                dias = nc.parsear_dias_semana(turno.get("dias_semana"))
                for indice, variable in self.vars_turno_dias.items():
                    variable.set(indice in dias)
                break

    def guardar_turno(self):
        entrada = nc.parse_hora(self.ent_turno_entrada.get())
        salida = nc.parse_hora(self.ent_turno_salida.get())
        if not entrada or not salida:
            return messagebox.showwarning("Atención", "Escriba la hora de entrada y de salida (HH:MM).")
        dias = [indice for indice, variable in self.vars_turno_dias.items() if variable.get()]
        if not dias:
            return messagebox.showwarning("Atención", "Seleccione al menos un día de la semana.")
        datos = {
            "id": self.turno_id,
            "nombre": self.ent_turno_nombre.get().strip(),
            "hora_entrada": entrada,
            "hora_salida": salida,
            "jornada_horas": parse_num(self.ent_turno_jornada.get(), 8),
            "tolerancia_min": parse_num(self.ent_turno_tolerancia.get(), 10),
            "refrigerio_min": parse_num(self.ent_turno_refrigerio.get(), 45),
            "dias_semana": dias,
            "cruza_medianoche": self.var_turno_cruza.get(),
            "aplica_nocturno": self.var_turno_nocturno.get(),
            "color": self.ent_turno_color.get().strip() or COLOR_PRIMARIO,
            "activo": self.var_turno_activo.get(),
            "observacion": self.ent_turno_observacion.get().strip(),
        }
        nuevo_id, error = nc.guardar_turno(datos, self.usuario_activo)
        if not nuevo_id:
            return messagebox.showerror("Error", error)
        self.turno_id = nuevo_id
        self.aviso("✅ Turno guardado: %s" % datos["nombre"], COLOR_OK)
        self.recargar_catalogos()

    def eliminar_turno(self):
        if not self.turno_id:
            return messagebox.showwarning("Atención", "Seleccione un turno de la lista.")
        if not messagebox.askyesno("Confirmar", "¿Eliminar el turno seleccionado?\n\n"
                                                "Los empleados y horarios que lo usaban quedarán sin turno."):
            return
        nc.eliminar_turno(self.turno_id, self.usuario_activo)
        self.turno_id = None
        self.nuevo_turno()
        self.recargar_catalogos()

    # ---------- HORARIOS ----------
    def _construir_horarios(self, padre):
        barra = ctk.CTkFrame(padre, fg_color="transparent")
        barra.pack(fill="x", pady=(2, 4))
        ctk.CTkLabel(barra, text="📅 HORARIOS SEMANALES", font=("Arial", 13, "bold"),
                     text_color=COLOR_PRIMARIO).pack(side="left")
        ctk.CTkButton(barra, text="➕ Nuevo horario", width=160, fg_color=COLOR_OK, hover_color="#1e8449",
                      command=self.nuevo_horario).pack(side="right", padx=4)
        ctk.CTkButton(barra, text="📋 Duplicar", width=120, fg_color="#8e44ad", hover_color="#6c3483",
                      command=self.duplicar_horario).pack(side="right", padx=4)
        contenedor = ctk.CTkFrame(padre, fg_color="transparent")
        contenedor.pack(fill="both", expand=True)
        self.tabla_horarios = crear_tabla(contenedor, [
            ("id", "ID", 50, "center"),
            ("nombre", "Horario", 300, "w"),
            ("descripcion", "Descripción", 300, "w"),
            ("resumen", "Distribución semanal", 480, "w"),
            ("activo", "Activo", 70, "center"),
        ], alto=5)
        self.tabla_horarios.bind("<<TreeviewSelect>>", self.al_seleccionar_horario)

        formulario = ctk.CTkFrame(padre, fg_color=COLOR_FONDO_SUAVE, corner_radius=10, border_width=1,
                                  border_color="#e0e0e0")
        formulario.pack(fill="x", pady=6)
        ctk.CTkLabel(formulario, text="📝 EDITOR DEL HORARIO SEMANAL", font=("Arial", 12, "bold"),
                     text_color=COLOR_PRIMARIO).pack(anchor="w", padx=12, pady=(8, 2))
        fila = ctk.CTkFrame(formulario, fg_color="transparent")
        fila.pack(fill="x", padx=12, pady=3)
        ctk.CTkLabel(fila, text="Nombre:", font=("Arial", 10, "bold")).pack(side="left")
        self.ent_horario_nombre = ctk.CTkEntry(fila, width=280)
        self.ent_horario_nombre.pack(side="left", padx=(2, 10))
        ctk.CTkLabel(fila, text="Descripción:", font=("Arial", 10, "bold")).pack(side="left")
        self.ent_horario_descripcion = ctk.CTkEntry(fila, width=320)
        self.ent_horario_descripcion.pack(side="left", padx=(2, 10))
        self.var_horario_activo = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(fila, text="Activo", variable=self.var_horario_activo,
                        font=("Arial", 10)).pack(side="left", padx=4)

        self.grid_horario = ctk.CTkFrame(formulario, fg_color="transparent")
        self.grid_horario.pack(fill="x", padx=12, pady=2)
        self.combos_dias_horario = {}
        # Los siete días en DOS columnas: el formulario es casi la mitad de alto y los
        # botones de guardar y eliminar quedan siempre dentro de la pantalla.
        for indice, dia in enumerate(nc.DIAS_SEMANA, start=1):
            columna = 0 if indice <= 4 else 1
            fila_dia = (indice - 1) if indice <= 4 else (indice - 5)
            celda = ctk.CTkFrame(self.grid_horario, fg_color="transparent")
            celda.grid(row=fila_dia, column=columna, sticky="w", padx=(0, 10), pady=1)
            ctk.CTkLabel(celda, text=dia + ":", font=("Arial", 10, "bold"), width=78,
                         anchor="w").pack(side="left")
            combo = ctk.CTkComboBox(celda, values=["(Descanso)"], width=245, state="readonly")
            combo.pack(side="left", padx=4)
            self.combos_dias_horario[indice] = combo
            self.registrar_combo_turno(combo)
        acciones = ctk.CTkFrame(formulario, fg_color="transparent")
        acciones.pack(fill="x", padx=12, pady=(6, 10))
        ctk.CTkButton(acciones, text="💾 Guardar horario", width=170, fg_color=COLOR_OK,
                      hover_color="#1e8449", command=self.guardar_horario).pack(side="left")
        ctk.CTkButton(acciones, text="🗑️ Eliminar horario", width=170, fg_color=COLOR_ERROR,
                      hover_color="#922b21", command=self.eliminar_horario).pack(side="left", padx=6)
        # El editor se ancla al fondo de la pestaña: sus botones nunca quedan cortados,
        # la tabla es la que cede espacio (y tiene su propia barra de desplazamiento).
        contenedor.pack_forget()
        formulario.pack_forget()
        formulario.pack(side="bottom", fill="x", pady=6)
        contenedor.pack(side="top", fill="both", expand=True)
        self.horario_id = None

    def nuevo_horario(self):
        self.horario_id = None
        self.ent_horario_nombre.delete(0, tk.END)
        self.ent_horario_descripcion.delete(0, tk.END)
        self.var_horario_activo.set(True)
        for combo in self.combos_dias_horario.values():
            combo.set("(Descanso)")

    @con_letrero("Leyendo los horarios...", "")
    def cargar_horarios(self):
        limpiar_tabla(self.tabla_horarios)
        turnos = {int(t["id"]): t for t in self.turnos}
        for horario in self.horarios:
            completo = nc.obtener_horario(horario["id"]) or {"detalle": {}}
            resumen = []
            for indice in range(1, 8):
                turno_id = completo["detalle"].get(indice)
                if turno_id and int(turno_id) in turnos:
                    turno = turnos[int(turno_id)]
                    resumen.append("%s: %s" % (nc.DIAS_CORTOS[indice - 1], turno["nombre"].split("(")[0].strip()))
                else:
                    resumen.append("%s: Descanso" % nc.DIAS_CORTOS[indice - 1])
            self.tabla_horarios.insert("", tk.END, values=(
                horario["id"], horario["nombre"], horario.get("descripcion") or "",
                " | ".join(resumen), "SÍ" if horario.get("activo") else "NO"))

    def al_seleccionar_horario(self, evento=None):
        seleccion = self.tabla_horarios.selection()
        if not seleccion:
            return
        identificador = self.tabla_horarios.item(seleccion[0], "values")[0]
        completo = nc.obtener_horario(identificador)
        if not completo:
            return
        self.horario_id = completo["id"]
        self.ent_horario_nombre.delete(0, tk.END)
        self.ent_horario_nombre.insert(0, completo["nombre"])
        self.ent_horario_descripcion.delete(0, tk.END)
        self.ent_horario_descripcion.insert(0, completo.get("descripcion") or "")
        self.var_horario_activo.set(bool(completo.get("activo")))
        for indice, combo in self.combos_dias_horario.items():
            turno_id = completo["detalle"].get(indice)
            asignado = "(Descanso)"
            if turno_id:
                for opcion in combo.cget("values"):
                    if opcion.split("|")[0].strip() == str(turno_id):
                        asignado = opcion
                        break
            combo.set(asignado)

    def guardar_horario(self):
        detalle = {}
        for indice, combo in self.combos_dias_horario.items():
            texto = combo.get()
            detalle[indice] = int(texto.split("|")[0].strip()) if texto != "(Descanso)" else None
        datos = {"id": self.horario_id, "nombre": self.ent_horario_nombre.get().strip(),
                 "descripcion": self.ent_horario_descripcion.get().strip(),
                 "activo": self.var_horario_activo.get()}
        nuevo_id, error = nc.guardar_horario(datos, detalle, self.usuario_activo)
        if not nuevo_id:
            return messagebox.showerror("Error", error)
        self.horario_id = nuevo_id
        self.aviso("✅ Horario guardado: %s" % datos["nombre"], COLOR_OK)
        self.recargar_catalogos()

    def eliminar_horario(self):
        if not self.horario_id:
            return messagebox.showwarning("Atención", "Seleccione un horario de la lista.")
        if not messagebox.askyesno("Confirmar", "¿Eliminar el horario seleccionado?"):
            return
        nc.eliminar_horario(self.horario_id, self.usuario_activo)
        self.horario_id = None
        self.nuevo_horario()
        self.recargar_catalogos()

    def duplicar_horario(self):
        if not self.horario_id:
            return messagebox.showwarning("Atención", "Seleccione el horario que desea duplicar.")
        from tkinter import simpledialog
        nombre = simpledialog.askstring("Duplicar horario", "Nombre del nuevo horario:",
                                        initialvalue="Copia de %s" % (self.ent_horario_nombre.get() or ""),
                                        parent=self.parent_frame)
        if not nombre:
            return
        nuevo_id, error = nc.duplicar_horario(self.horario_id, nombre, self.usuario_activo)
        if not nuevo_id:
            return messagebox.showerror("Error", error)
        self.recargar_catalogos()
        self.aviso("✅ Horario duplicado", COLOR_OK)

    # ---------- ASIGNACIONES ----------
    def _construir_asignaciones(self, padre):
        ctk.CTkLabel(padre, text="🔀 ASIGNACIONES POR RANGO DE FECHAS", font=("Arial", 13, "bold"),
                     text_color=COLOR_PRIMARIO).pack(anchor="w", pady=(2, 2))
        ctk.CTkLabel(padre, text="Sirve para turnos rotativos o cambios temporales: la asignación por fecha "
                                "tiene prioridad sobre el horario fijo del empleado.",
                     font=("Arial", 10), text_color=COLOR_GRIS).pack(anchor="w", pady=(0, 4))
        formulario = ctk.CTkFrame(padre, fg_color=COLOR_FONDO_SUAVE, corner_radius=10, border_width=1,
                                  border_color="#e0e0e0")
        formulario.pack(fill="x", pady=4)
        fila = ctk.CTkFrame(formulario, fg_color="transparent")
        fila.pack(fill="x", padx=12, pady=8)
        ctk.CTkLabel(fila, text="Empleado:", font=("Arial", 10, "bold")).pack(side="left")
        self.cmb_asig_empleado = ctk.CTkComboBox(fila, values=self.opciones_empleados(), width=280,
                                                 state="readonly")
        self.cmb_asig_empleado.pack(side="left", padx=(2, 8))
        self.registrar_combo_empleado(self.cmb_asig_empleado)
        ctk.CTkLabel(fila, text="Horario:", font=("Arial", 10, "bold")).pack(side="left")
        self.cmb_asig_horario = ctk.CTkComboBox(fila, values=["(Ninguno)"], width=200, state="readonly")
        self.cmb_asig_horario.pack(side="left", padx=(2, 8))
        self.registrar_combo_horario(self.cmb_asig_horario)
        ctk.CTkLabel(fila, text="o Turno:", font=("Arial", 10, "bold")).pack(side="left")
        self.cmb_asig_turno = ctk.CTkComboBox(fila, values=["(Descanso)"], width=200, state="readonly")
        self.cmb_asig_turno.pack(side="left", padx=(2, 8))
        self.registrar_combo_turno(self.cmb_asig_turno)
        ctk.CTkLabel(fila, text="Desde:", font=("Arial", 10, "bold")).pack(side="left")
        self.ent_asig_desde = ctk.CTkEntry(fila, width=100, placeholder_text="DD/MM/AAAA")
        self.ent_asig_desde.pack(side="left", padx=2)
        ctk.CTkButton(fila, text="📅", width=30, fg_color=COLOR_PRIMARIO, hover_color=COLOR_HOVER,
                      command=lambda: self.abrir_calendario(self.ent_asig_desde)).pack(side="left", padx=2)
        ctk.CTkLabel(fila, text="Hasta:", font=("Arial", 10, "bold")).pack(side="left", padx=(8, 2))
        self.ent_asig_hasta = ctk.CTkEntry(fila, width=100, placeholder_text="(opcional)")
        self.ent_asig_hasta.pack(side="left", padx=2)
        ctk.CTkButton(fila, text="📅", width=30, fg_color=COLOR_PRIMARIO, hover_color=COLOR_HOVER,
                      command=lambda: self.abrir_calendario(self.ent_asig_hasta)).pack(side="left", padx=2)
        fila2 = ctk.CTkFrame(formulario, fg_color="transparent")
        fila2.pack(fill="x", padx=12, pady=(0, 10))
        ctk.CTkLabel(fila2, text="Observación:", font=("Arial", 10, "bold")).pack(side="left")
        self.ent_asig_observacion = ctk.CTkEntry(fila2, width=420)
        self.ent_asig_observacion.pack(side="left", padx=(2, 10))
        ctk.CTkButton(fila2, text="💾 Guardar asignación", width=180, fg_color=COLOR_OK,
                      hover_color="#1e8449", command=self.guardar_asignacion).pack(side="left")
        ctk.CTkButton(fila2, text="🗑️ Eliminar seleccionada", width=190, fg_color=COLOR_ERROR,
                      hover_color="#922b21", command=self.eliminar_asignacion).pack(side="left", padx=6)
        self.tabla_asignaciones = crear_tabla(padre, [
            ("id", "ID", 50, "center"),
            ("dni", "DNI", 100, "w"),
            ("empleado", "Empleado", 250, "w"),
            ("horario", "Horario", 220, "w"),
            ("turno", "Turno", 220, "w"),
            ("desde", "Desde", 100, "center"),
            ("hasta", "Hasta", 100, "center"),
            ("activo", "Activo", 70, "center"),
            ("observacion", "Observación", 280, "w"),
        ], alto=10)

    @con_letrero("Leyendo las asignaciones...", "")
    def cargar_asignaciones(self):
        limpiar_tabla(self.tabla_asignaciones)
        horarios = {int(h["id"]): h["nombre"] for h in self.horarios}
        turnos = {int(t["id"]): t["nombre"] for t in self.turnos}
        for asignacion in nc.listar_asignaciones():
            self.tabla_asignaciones.insert("", tk.END, values=(
                asignacion["id"], asignacion["dni"], self.nombre_empleado(asignacion["dni"]),
                horarios.get(asignacion.get("horario_id"), ""),
                turnos.get(asignacion.get("turno_id"), ""),
                asignacion["fecha_desde"], asignacion.get("fecha_hasta") or "(sin fin)",
                "SÍ" if asignacion.get("activo") else "NO",
                asignacion.get("observacion") or ""))

    def guardar_asignacion(self):
        horario = self.cmb_asig_horario.get()
        turno = self.cmb_asig_turno.get()
        datos = {
            "dni": self.dni_de_opcion(self.cmb_asig_empleado.get()),
            "horario_id": int(horario.split("|")[0].strip()) if horario != "(Ninguno)" else None,
            "turno_id": int(turno.split("|")[0].strip()) if turno != "(Descanso)" else None,
            "fecha_desde": parse_fecha_texto(self.ent_asig_desde.get()),
            "fecha_hasta": parse_fecha_texto(self.ent_asig_hasta.get()),
            "observacion": self.ent_asig_observacion.get().strip(),
            "activo": True,
        }
        nuevo_id, error = nc.guardar_asignacion(datos, self.usuario_activo)
        if not nuevo_id:
            return messagebox.showerror("Error", error)
        self.cargar_asignaciones()
        self.aviso("✅ Asignación guardada", COLOR_OK)

    def eliminar_asignacion(self):
        seleccion = self.tabla_asignaciones.selection()
        if not seleccion:
            return messagebox.showwarning("Atención", "Seleccione una asignación de la lista.")
        identificador = self.tabla_asignaciones.item(seleccion[0], "values")[0]
        if not messagebox.askyesno("Confirmar", "¿Eliminar la asignación seleccionada?"):
            return
        nc.eliminar_asignacion(identificador, self.usuario_activo)
        self.cargar_asignaciones()

    # =====================================================
    # PESTAÑA: PLANILLA
    # =====================================================
    def construir_tab_planilla(self):
        tab = self.tab_planilla
        barra = ctk.CTkFrame(tab, fg_color="transparent")
        barra.pack(fill="x", padx=10, pady=(10, 6))
        ctk.CTkLabel(barra, text="Período:", font=("Arial", 12, "bold")).pack(side="left")
        ctk.CTkButton(barra, text="◀", width=36, fg_color=COLOR_PRIMARIO, hover_color=COLOR_HOVER,
                      command=lambda: self.cambiar_periodo_planilla(-1)).pack(side="left", padx=4)
        self.ent_periodo_planilla = ctk.CTkEntry(barra, width=90, justify="center")
        self.ent_periodo_planilla.insert(0, periodo_para_entry(self.periodo_actual))
        self.ent_periodo_planilla.pack(side="left")
        ctk.CTkButton(barra, text="▶", width=36, fg_color=COLOR_PRIMARIO, hover_color=COLOR_HOVER,
                      command=lambda: self.cambiar_periodo_planilla(1)).pack(side="left", padx=4)
        ctk.CTkButton(barra, text="🔄 Ver período", width=130, fg_color="#34495e", hover_color="#2c3e50",
                      command=self.cargar_planilla).pack(side="left", padx=8)
        ctk.CTkButton(barra, text="🧮 CALCULAR / RECALCULAR PLANILLA", width=280, height=32,
                      font=("Arial", 12, "bold"), fg_color=COLOR_OK, hover_color="#1e8449",
                      command=self.calcular_planilla).pack(side="left", padx=8)
        self.lbl_estado_planilla = ctk.CTkLabel(barra, text="", font=("Arial", 12, "bold"),
                                                text_color=COLOR_PRIMARIO)
        self.lbl_estado_planilla.pack(side="left", padx=10)

        self.sub_planilla = ctk.CTkSegmentedButton(
            tab, values=["Resumen del período", "Boleta de pago", "Historial"],
            selected_color=COLOR_PRIMARIO, selected_hover_color=COLOR_HOVER,
            command=self.cambiar_sub_planilla)
        self.sub_planilla.pack(fill="x", padx=10, pady=(0, 6))
        self.frame_planilla_resumen = ctk.CTkFrame(tab, fg_color="transparent")
        self.frame_planilla_boleta = ctk.CTkFrame(tab, fg_color="transparent")
        self.frame_planilla_historial = ctk.CTkFrame(tab, fg_color="transparent")
        self._construir_planilla_resumen(self.frame_planilla_resumen)
        self._construir_planilla_boleta(self.frame_planilla_boleta)
        self._construir_planilla_historial(self.frame_planilla_historial)
        self.sub_planilla.set("Resumen del período")
        self.cambiar_sub_planilla("Resumen del período")

    def cambiar_sub_planilla(self, valor):
        for frame in (self.frame_planilla_resumen, self.frame_planilla_boleta,
                      self.frame_planilla_historial):
            frame.pack_forget()
        mapa = {"Resumen del período": self.frame_planilla_resumen, "Boleta de pago":
                self.frame_planilla_boleta, "Historial": self.frame_planilla_historial}
        frame = mapa.get(valor, self.frame_planilla_resumen)
        frame.pack(fill="both", expand=True)
        if valor == "Resumen del período":
            self.cargar_planilla()
        elif valor == "Boleta de pago":
            self.cargar_combo_boleta()
        else:
            self.cargar_historial_planillas()

    def _construir_planilla_resumen(self, padre):
        contenedor_totales = ctk.CTkFrame(padre, fg_color="transparent")
        contenedor_totales.pack(fill="x", pady=(2, 4))
        # Las etiquetas de totales se envuelven y los botones van en su propia fila:
        # el texto de totales es largo (~900 px) y antes aplastaba los botones de exportar.
        self.lbl_totales_planilla = ctk.CTkLabel(contenedor_totales, text="", font=("Arial", 13, "bold"),
                                                 text_color=COLOR_PRIMARIO, justify="left", anchor="w",
                                                 wraplength=740)
        self.lbl_totales_planilla.pack(anchor="w", fill="x")
        fila_exportar = ctk.CTkFrame(padre, fg_color="transparent")
        fila_exportar.pack(fill="x", pady=(0, 4))
        ctk.CTkButton(fila_exportar, text="📊 Exportar planilla a Excel", width=230,
                      fg_color=COLOR_OK, hover_color="#1e8449",
                      command=self.exportar_planilla).pack(side="left", padx=(0, 6))
        ctk.CTkButton(fila_exportar, text="📄 Exportar todas las boletas", width=230,
                      fg_color="#8e44ad", hover_color="#6c3483",
                      command=self.exportar_boletas).pack(side="left", padx=0)
        self.tabla_planilla = crear_tabla(padre, [
            ("dni", "DNI", 100, "w"),
            ("empleado", "Empleado", 260, "w"),
            ("dias", "Días lab.", 75, "center"),
            ("tardanzas", "Tardanzas", 80, "center"),
            ("faltas", "Faltas", 70, "center"),
            ("horas_extra", "Horas extra", 90, "center"),
            ("horas_excedente", "Horas a bono", 95, "center"),
            ("ingresos", "Total ingresos", 120, "e"),
            ("descuentos", "Total descuentos", 130, "e"),
            ("neto", "Neto a pagar", 120, "e"),
            ("aportes", "ESSALUD (empleador)", 150, "e"),
        ], alto=12)
        self.tabla_planilla.bind("<Double-1>", lambda _e: self.ir_a_boleta())

        acciones = ctk.CTkFrame(padre, fg_color=COLOR_FONDO_SUAVE, corner_radius=10, border_width=1,
                                border_color="#e0e0e0")
        acciones.pack(fill="x", pady=6)
        ctk.CTkLabel(acciones, text="Estado de la planilla:", font=("Arial", 11, "bold")).pack(side="left",
                                                                                              padx=12)
        self.cmb_estado_planilla = ctk.CTkComboBox(acciones, values=npl.ESTADOS_PLANILLA, width=160,
                                                   state="readonly")
        self.cmb_estado_planilla.set("CALCULADA")
        self.cmb_estado_planilla.pack(side="left", padx=6)
        ctk.CTkButton(acciones, text="💾 Cambiar estado", width=160, fg_color=COLOR_PRIMARIO,
                      hover_color=COLOR_HOVER, command=self.cambiar_estado_planilla).pack(side="left", padx=6)
        ctk.CTkButton(acciones, text="🗑️ Eliminar planilla del período", width=230, fg_color=COLOR_ERROR,
                      hover_color="#922b21", command=self.eliminar_planilla).pack(side="left", padx=6)
        ctk.CTkLabel(acciones, text="Doble clic en un empleado para ver su boleta", font=("Arial", 10),
                     text_color=COLOR_GRIS).pack(side="left", padx=10)
        # La barra de estado se ancla al fondo para que sus botones siempre se vean
        contenedor_tabla_planilla = self.tabla_planilla.master
        contenedor_tabla_planilla.pack_forget()
        acciones.pack_forget()
        acciones.pack(side="bottom", fill="x", pady=6)
        contenedor_tabla_planilla.pack(side="top", fill="both", expand=True)

    def _construir_planilla_boleta(self, padre):
        barra = ctk.CTkFrame(padre, fg_color="transparent")
        barra.pack(fill="x", pady=(2, 4))
        ctk.CTkLabel(barra, text="Empleado:", font=("Arial", 11, "bold")).pack(side="left")
        self.cmb_boleta_empleado = ctk.CTkComboBox(barra, values=self.opciones_empleados(), width=340,
                                                   state="readonly")
        self.cmb_boleta_empleado.pack(side="left", padx=6)
        self.registrar_combo_empleado(self.cmb_boleta_empleado)
        ctk.CTkButton(barra, text="🔎 Ver boleta", width=140, fg_color=COLOR_PRIMARIO,
                      hover_color=COLOR_HOVER, command=self.mostrar_boleta).pack(side="left", padx=6)
        ctk.CTkButton(barra, text="📄 Exportar esta boleta", width=190, fg_color=COLOR_OK,
                      hover_color="#1e8449", command=self.exportar_boleta_actual).pack(side="left", padx=6)
        self.frame_boleta = ctk.CTkScrollableFrame(padre, fg_color="#ffffff", corner_radius=10,
                                                   border_width=1, border_color="#e0e0e0")
        self.frame_boleta.pack(fill="both", expand=True, pady=4)

    def _construir_planilla_historial(self, padre):
        ctk.CTkLabel(padre, text="📚 HISTORIAL DE PLANILLAS", font=("Arial", 13, "bold"),
                     text_color=COLOR_PRIMARIO).pack(anchor="w", pady=(4, 4))
        self.tabla_historial_planilla = crear_tabla(padre, [
            ("periodo", "Período", 120, "center"),
            ("estado", "Estado", 120, "center"),
            ("empleados", "Empleados", 90, "center"),
            ("ingresos", "Ingresos", 130, "e"),
            ("descuentos", "Descuentos", 130, "e"),
            ("neto", "Neto", 130, "e"),
            ("aportes", "Aportes", 130, "e"),
            ("usuario", "Usuario", 130, "w"),
            ("fecha", "Fecha de cálculo", 150, "center"),
        ], alto=14)

    def cambiar_periodo_planilla(self, cantidad):
        periodo = periodo_desde_entry(self.ent_periodo_planilla.get()) or self.periodo_actual
        nuevo = sumar_meses(periodo, cantidad)
        self.ent_periodo_planilla.delete(0, tk.END)
        self.ent_periodo_planilla.insert(0, periodo_para_entry(nuevo))
        self.cargar_planilla()

    @con_letrero("Leyendo la planilla del período...", "Calculando los totales")
    def cargar_planilla(self):
        periodo = periodo_desde_entry(self.ent_periodo_planilla.get())
        if not periodo:
            return messagebox.showwarning("Atención", "Escriba el período con el formato MM/AAAA.")
        planilla = npl.obtener_planilla(periodo)
        limpiar_tabla(self.tabla_planilla)
        if not planilla:
            self.lbl_estado_planilla.configure(text="SIN CALCULAR", text_color=COLOR_ALERTA)
            self.lbl_totales_planilla.configure(text="La planilla de %s aún no se ha calculado."
                                                     % nc.nombre_mes(periodo))
            return
        self.lbl_estado_planilla.configure(text=planilla["estado"], text_color=COLOR_OK)
        for registro in npl.resumen_planilla(periodo):
            datos_asistencia = self._asistencia_empleado(periodo, registro["dni"])
            self.tabla_planilla.insert("", tk.END, values=(
                registro["dni"], registro["empleado"], datos_asistencia.get("dias_laborables", 0),
                datos_asistencia.get("tardanzas", 0), datos_asistencia.get("faltas", 0),
                "%.1f h" % (nc.a_int(datos_asistencia.get("minutos_extra")) / 60.0),
                "%.1f h" % (nc.a_int(datos_asistencia.get("minutos_excedente")) / 60.0),
                formatear_monto(registro["ingresos"]), formatear_monto(registro["descuentos"]),
                formatear_monto(registro["neto"]), formatear_monto(registro["aportes"])))
        self.lbl_totales_planilla.configure(
            text="%s  |  Ingresos: %s  |  Descuentos: %s  |  NETO A PAGAR: %s  |  Aportes: %s  |  "
                 "Empleados: %d" % (nc.nombre_mes(periodo), formatear_monto(planilla["total_ingresos"]),
                                    formatear_monto(planilla["total_descuentos"]),
                                    formatear_monto(planilla["total_neto"]),
                                    formatear_monto(planilla["total_aportes"]),
                                    planilla["n_empleados"]))
        self.sub_planilla.set("Resumen del período")

    def _asistencia_empleado(self, periodo, dni):
        clave = "%s|%s" % (periodo, dni)
        if not hasattr(self, "_cache_resumen_asistencia"):
            self._cache_resumen_asistencia = {}
        if clave not in self._cache_resumen_asistencia:
            dias = nc.dias_del_periodo(periodo)
            resumen = nc.resumen_asistencia(dias[0], dias[-1], dnis=[dni]) if dias else {}
            self._cache_resumen_asistencia[clave] = resumen.get(str(dni), {})
        return self._cache_resumen_asistencia[clave]

    def calcular_planilla(self):
        periodo = periodo_desde_entry(self.ent_periodo_planilla.get())
        if not periodo:
            return messagebox.showwarning("Atención", "Escriba el período con el formato MM/AAAA.")
        if not messagebox.askyesno(
                "Calcular planilla",
                "Se calculará la planilla de %s con la asistencia registrada.\n\n"
                "¿Desea continuar?" % nc.nombre_mes(periodo)):
            return
        self.aviso("⏳ Calculando planilla de %s..." % nc.nombre_mes(periodo), COLOR_ALERTA)
        letrero = LetreroCarga(self.frame_main, "Calculando la planilla...",
                               "Haberes, descuentos y aportes de %s" % nc.nombre_mes(periodo))
        self._cache_resumen_asistencia = {}

        def tarea(estado):
            estado["resultado"] = npl.calcular_planilla(periodo, usuario=self.usuario_activo)

        def terminar(estado):
            letrero.cerrar()
            if estado.get("error"):
                self.aviso("❌ Error al calcular la planilla", COLOR_ERROR)
                return messagebox.showerror("Error", "No se pudo calcular la planilla:\n%s" % estado["error"])
            resultado = estado.get("resultado") or {}
            if not resultado.get("ok"):
                self.aviso("❌ %s" % resultado.get("error"), COLOR_ERROR)
                return messagebox.showerror("Error", resultado.get("error"))
            messagebox.showinfo(
                "Planilla calculada",
                "Período: %s\n\n"
                "Empleados: %d\n"
                "Líneas de cálculo: %d\n\n"
                "Total ingresos: %s\n"
                "Total descuentos: %s\n"
                "NETO A PAGAR: %s\n"
                "Aportes del empleador: %s"
                % (nc.nombre_mes(periodo), resultado["empleados"], resultado["lineas"],
                   formatear_monto(resultado["total_ingresos"]),
                   formatear_monto(resultado["total_descuentos"]),
                   formatear_monto(resultado["total_neto"]),
                   formatear_monto(resultado["total_aportes"])))
            self.aviso("✅ Planilla de %s calculada" % nc.nombre_mes(periodo), COLOR_OK)
            self.cargar_planilla()
            self.cargar_panel()

        ejecutar_en_hilo(self.parent_frame, tarea, al_terminar=terminar)

    def cambiar_estado_planilla(self):
        periodo = periodo_desde_entry(self.ent_periodo_planilla.get())
        if not periodo:
            return
        estado = self.cmb_estado_planilla.get()
        ok, error = npl.cambiar_estado_planilla(periodo, estado, self.usuario_activo)
        if not ok:
            return messagebox.showerror("Error", error)
        self.aviso("✅ Planilla de %s marcada como %s" % (nc.nombre_mes(periodo), estado), COLOR_OK)
        self.cargar_planilla()

    def eliminar_planilla(self):
        periodo = periodo_desde_entry(self.ent_periodo_planilla.get())
        if not periodo:
            return
        if not messagebox.askyesno("Confirmar", "¿Eliminar la planilla calculada de %s?\n\n"
                                                "No se borra la asistencia ni las marcaciones."
                                                % nc.nombre_mes(periodo)):
            return
        ok, error = npl.eliminar_planilla(periodo, self.usuario_activo)
        if not ok:
            return messagebox.showerror("Error", error)
        self.cargar_planilla()
        self.cargar_historial_planillas()

    @con_letrero("Leyendo el historial de planillas...", "")
    def cargar_historial_planillas(self):
        limpiar_tabla(self.tabla_historial_planilla)
        for periodo in npl.listar_periodos():
            self.tabla_historial_planilla.insert("", tk.END, values=(
                nc.nombre_mes(periodo["periodo"]), periodo["estado"], periodo["n_empleados"],
                formatear_monto(periodo["total_ingresos"]), formatear_monto(periodo["total_descuentos"]),
                formatear_monto(periodo["total_neto"]), formatear_monto(periodo["total_aportes"]),
                periodo["usuario"], periodo["fecha_calculo"]))

    def ir_a_boleta(self):
        seleccion = self.tabla_planilla.selection()
        if not seleccion:
            return
        dni = self.tabla_planilla.item(seleccion[0], "values")[0]
        for opcion in self.cmb_boleta_empleado.cget("values"):
            if opcion.split("|")[0].strip() == str(dni):
                self.cmb_boleta_empleado.set(opcion)
                break
        self.sub_planilla.set("Boleta de pago")
        self.cambiar_sub_planilla("Boleta de pago")
        self.mostrar_boleta()

    def cargar_combo_boleta(self):
        """El combo de empleados ya está sincronizado; se refresca por si cambió el padrón."""
        self.refrescar_combos_globales()

    @con_letrero("Leyendo la boleta del empleado...", "")
    def mostrar_boleta(self):
        periodo = periodo_desde_entry(self.ent_periodo_planilla.get())
        dni = self.dni_de_opcion(self.cmb_boleta_empleado.get())
        if not periodo or not dni:
            return messagebox.showwarning("Atención", "Seleccione el empleado y verifique el período.")
        for widget in self.frame_boleta.winfo_children():
            widget.destroy()
        boleta = npl.boleta_empleado(periodo, dni)
        if not boleta["ingresos"] and not boleta["descuentos"]:
            ctk.CTkLabel(self.frame_boleta, text="No hay planilla calculada para %s.\n\n"
                                                "Calcule primero la planilla del período."
                         % nc.nombre_mes(periodo), font=("Arial", 13), text_color=COLOR_ALERTA
                         ).pack(pady=40)
            return
        empresa = "BLACK RIDERS E.I.R.L."
        try:
            with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "config_local.json"),
                      "r", encoding="utf-8") as manejador:
                configuracion = json.load(manejador)
            empresa = configuracion.get("razon_social_empresa") or empresa
        except Exception:
            pass
        encabezado = ctk.CTkFrame(self.frame_boleta, fg_color=COLOR_PRIMARIO, corner_radius=8)
        encabezado.pack(fill="x", padx=10, pady=(10, 6))
        ctk.CTkLabel(encabezado, text="BOLETA DE PAGO — %s" % boleta["periodo_nombre"],
                     font=("Arial", 16, "bold"), text_color="white").pack(pady=(10, 2))
        ctk.CTkLabel(encabezado, text=empresa, font=("Arial", 12), text_color="white").pack(pady=(0, 10))
        datos = ctk.CTkFrame(self.frame_boleta, fg_color="transparent")
        datos.pack(fill="x", padx=10)
        filas = [("Empleado", boleta["empleado"].get("nombre", "")),
                 ("DNI", boleta["dni"]),
                 ("Cargo", boleta["empleado"].get("cargo") or "-"),
                 ("Fecha de ingreso", fecha_para_entry(boleta["empleado"].get("fecha_ingreso")) or "-"),
                 ("Sueldo básico", formatear_monto(boleta["empleado"].get("sueldo_basico"))),
                 ("Valor día", formatear_monto(boleta["valor_dia"])),
                 ("Sistema de pensión", boleta["empleado"].get("sistema_pension") or "-")]
        asistencia = boleta.get("asistencia") or {}
        filas += [("Días laborables", asistencia.get("dias_laborables", 0)),
                  ("Puntuales / Tardanzas", "%s / %s" % (asistencia.get("puntuales", 0),
                                                         asistencia.get("tardanzas", 0))),
                  ("Faltas", asistencia.get("faltas", 0)),
                  ("Horas extra", "%.2f h" % (nc.a_int(asistencia.get("minutos_extra")) / 60.0))]
        garantia = boleta.get("garantia")
        if garantia:
            filas += [("Garantía mensual", formatear_monto(garantia["garantia"])),
                      ("Cálculo de ley del mes", formatear_monto(garantia["calculo_ley"])),
                      ("Bono de complemento", formatear_monto(garantia["complemento"])),
                      ("Total del mes", formatear_monto(garantia["total_mes"])),
                      ("Garantía medida sobre", garantia["sobre"])]
        for indice, (etiqueta, valor) in enumerate(filas):
            ctk.CTkLabel(datos, text=etiqueta + ":", font=("Arial", 11, "bold"), width=180,
                         anchor="w").grid(row=indice // 2, column=(indice % 2) * 2, sticky="w", padx=6, pady=2)
            ctk.CTkLabel(datos, text=str(valor), font=("Arial", 11), anchor="w").grid(
                row=indice // 2, column=(indice % 2) * 2 + 1, sticky="w", padx=6, pady=2)

        self._pintar_bloques_boleta(boleta)
        totales = ctk.CTkFrame(self.frame_boleta, fg_color="#eafaf1", corner_radius=8, border_width=1,
                               border_color="#a9dfbf")
        totales.pack(fill="x", padx=10, pady=(6, 14))
        ctk.CTkLabel(totales, text="NETO A PAGAR:   %s" % formatear_monto(boleta["neto_pagar"]),
                     font=("Arial", 18, "bold"), text_color=COLOR_OK).pack(pady=12)

    def _pintar_bloques_boleta(self, boleta):
        bloques = [("INGRESOS", boleta["ingresos"], COLOR_OK, boleta["total_ingresos"]),
                   ("DESCUENTOS", boleta["descuentos"], COLOR_ERROR, boleta["total_descuentos"]),
                   ("APORTES DEL EMPLEADOR", boleta["aportes"], COLOR_PRIMARIO, boleta["total_aportes"])]
        contenedor = ctk.CTkFrame(self.frame_boleta, fg_color="transparent")
        contenedor.pack(fill="x", padx=10, pady=6)
        for indice, (titulo, lineas, color, total) in enumerate(bloques):
            bloque = ctk.CTkFrame(contenedor, fg_color="#ffffff", corner_radius=8, border_width=1,
                                  border_color="#e0e0e0")
            bloque.grid(row=0, column=indice, sticky="nsew", padx=4)
            contenedor.grid_columnconfigure(indice, weight=1)
            ctk.CTkLabel(bloque, text=titulo, font=("Arial", 12, "bold"), text_color=color).pack(
                anchor="w", padx=10, pady=(8, 2))
            if not lineas:
                ctk.CTkLabel(bloque, text="(sin movimientos)", font=("Arial", 10),
                             text_color=COLOR_GRIS).pack(anchor="w", padx=10, pady=4)
            for linea in lineas:
                ctk.CTkLabel(bloque, text="%s" % linea["concepto"], font=("Arial", 10),
                             anchor="w", wraplength=210, justify="left").pack(anchor="w", padx=10)
                ctk.CTkLabel(bloque, text=formatear_monto(linea["monto"]), font=("Arial", 11, "bold"),
                             text_color=color, anchor="e").pack(anchor="e", padx=10)
                if linea.get("formula"):
                    ctk.CTkLabel(bloque, text=linea["formula"], font=("Arial", 8, "italic"),
                                 text_color=COLOR_GRIS, anchor="w", wraplength=210,
                                 justify="left").pack(anchor="w", padx=10, pady=(0, 2))
            ctk.CTkLabel(bloque, text="TOTAL: %s" % formatear_monto(total), font=("Arial", 11, "bold"),
                         text_color=color).pack(anchor="e", padx=10, pady=(4, 10))

    def exportar_planilla(self):
        periodo = periodo_desde_entry(self.ent_periodo_planilla.get())
        if not periodo:
            return
        ruta = guardar_archivo_dialogo("Exportar planilla a Excel", ".xlsx",
                                       "Planilla_%s.xlsx" % periodo, [("Archivos Excel", "*.xlsx")])
        if not ruta:
            return
        ok, error = npl.exportar_planilla_excel(ruta, periodo)
        if not ok:
            return messagebox.showerror("Error", "No se pudo exportar:\n%s" % error)
        registrar_auditoria(self.usuario_activo, "Nómina", "Exportó la planilla de %s" % periodo)
        if messagebox.askyesno("Exportación lista", "Reporte guardado en:\n%s\n\n¿Desea abrirlo?" % ruta):
            abrir_documento(ruta)

    def exportar_boletas(self):
        periodo = periodo_desde_entry(self.ent_periodo_planilla.get())
        if not periodo:
            return
        ruta = guardar_archivo_dialogo("Exportar boletas a Excel", ".xlsx",
                                       "Boletas_%s.xlsx" % periodo, [("Archivos Excel", "*.xlsx")])
        if not ruta:
            return
        ok, error = npl.exportar_boletas_excel(ruta, periodo)
        if not ok:
            return messagebox.showerror("Error", "No se pudo exportar:\n%s" % error)
        if messagebox.askyesno("Exportación lista", "Boletas guardadas en:\n%s\n\n¿Desea abrirlas?" % ruta):
            abrir_documento(ruta)

    def exportar_boleta_actual(self):
        periodo = periodo_desde_entry(self.ent_periodo_planilla.get())
        dni = self.dni_de_opcion(self.cmb_boleta_empleado.get())
        if not periodo or not dni:
            return
        ruta = guardar_archivo_dialogo("Exportar boleta de pago", ".xlsx",
                                       "Boleta_%s_%s.xlsx" % (dni, periodo),
                                       [("Archivos Excel", "*.xlsx")])
        if not ruta:
            return
        ok, error = npl.exportar_boletas_excel(ruta, periodo, dnis=[dni])
        if not ok:
            return messagebox.showerror("Error", "No se pudo exportar:\n%s" % error)
        if messagebox.askyesno("Exportación lista", "Boleta guardada en:\n%s\n\n¿Desea abrirla?" % ruta):
            abrir_documento(ruta)

    # =====================================================
    # PESTAÑA: REPORTES
    # =====================================================
    def construir_tab_reportes(self):
        tab = self.tab_reportes
        barra = ctk.CTkFrame(tab, fg_color="transparent")
        barra.pack(fill="x", padx=10, pady=(10, 6))
        ctk.CTkLabel(barra, text="Período:", font=("Arial", 12, "bold")).pack(side="left")
        self.ent_periodo_reporte = ctk.CTkEntry(barra, width=90, justify="center")
        self.ent_periodo_reporte.insert(0, periodo_para_entry(self.periodo_actual))
        self.ent_periodo_reporte.pack(side="left", padx=4)
        ctk.CTkLabel(barra, text="Desde:", font=("Arial", 11, "bold")).pack(side="left", padx=(12, 2))
        self.ent_rep_desde = ctk.CTkEntry(barra, width=100, placeholder_text="DD/MM/AAAA")
        self.ent_rep_desde.pack(side="left")
        ctk.CTkLabel(barra, text="Hasta:", font=("Arial", 11, "bold")).pack(side="left", padx=(8, 2))
        self.ent_rep_hasta = ctk.CTkEntry(barra, width=100, placeholder_text="DD/MM/AAAA")
        self.ent_rep_hasta.pack(side="left")
        ctk.CTkButton(barra, text="🔄 Generar reporte", width=170, fg_color=COLOR_PRIMARIO,
                      hover_color=COLOR_HOVER, command=self.generar_reporte).pack(side="left", padx=8)
        ctk.CTkButton(barra, text="📊 Exportar a Excel", width=170, fg_color=COLOR_OK,
                      hover_color="#1e8449", command=self.exportar_reporte).pack(side="left", padx=4)

        self.sub_reportes = ctk.CTkSegmentedButton(
            tab, values=["Puntualidad", "Tardanzas", "Faltas", "Horas extra", "Garantías",
                         "Marcaciones"],
            selected_color=COLOR_PRIMARIO, selected_hover_color=COLOR_HOVER,
            command=lambda _v: self.generar_reporte())
        self.sub_reportes.pack(fill="x", padx=10, pady=(0, 6))
        self.sub_reportes.set("Puntualidad")
        self.contenedor_reporte = ctk.CTkFrame(tab, fg_color="transparent")
        self.contenedor_reporte.pack(fill="both", expand=True)
        self.tabla_reporte = None
        self.datos_reporte = []

    def _columnas_reporte(self, nombre):
        if nombre == "Puntualidad":
            return [("dni", "DNI", 100, "w"), ("empleado", "Empleado", 250, "w"),
                    ("dias_laborables", "Días lab.", 80, "center"), ("puntuales", "Puntuales", 85, "center"),
                    ("tardanzas", "Tardanzas", 85, "center"), ("faltas", "Faltas", 75, "center"),
                    ("incompletos", "Incompletos", 90, "center"),
                    ("minutos_tardanza", "Min. tardanza", 100, "center"),
                    ("horas_extra", "Horas extra", 95, "center"),
                    ("horas_trabajadas", "Horas trabajadas", 120, "center"),
                    ("porcentaje_puntualidad", "% Puntualidad", 110, "center")]
        if nombre == "Tardanzas":
            return [("dni", "DNI", 100, "w"), ("empleado", "Empleado", 250, "w"),
                    ("fecha", "Fecha", 100, "center"), ("entrada", "Entrada", 85, "center"),
                    ("salida", "Salida", 85, "center"), ("minutos_tardanza", "Min. tardanza", 100, "center"),
                    ("minutos_anticipo", "Min. anticipo", 100, "center"),
                    ("estado", "Estado", 140, "w")]
        if nombre == "Faltas":
            return [("dni", "DNI", 100, "w"), ("empleado", "Empleado", 250, "w"),
                    ("fecha", "Fecha", 100, "center"), ("dia", "Día", 100, "w"),
                    ("observacion", "Observación", 420, "w")]
        if nombre == "Horas extra":
            return [("dni", "DNI", 100, "w"), ("empleado", "Empleado", 250, "w"),
                    ("fecha", "Fecha", 100, "center"), ("minutos", "Minutos", 85, "center"),
                    ("horas", "Horas", 80, "center"), ("tipo", "Tipo", 120, "center"),
                    ("origen", "Origen", 90, "center"), ("aprobado", "Aprobado", 90, "center"),
                    ("monto_estimado", "Monto estimado", 130, "e")]
        if nombre == "Garantías":
            return [("dni", "DNI", 100, "w"), ("empleado", "Empleado", 220, "w"),
                    ("garantia", "Garantía mensual", 120, "e"), ("calculo_ley", "Cálculo de ley", 120, "e"),
                    ("horas_excedente", "Horas a bono", 100, "center"),
                    ("bono_excedente", "Bono por horas adicionales", 160, "e"),
                    ("complemento", "Bono de complemento", 140, "e"),
                    ("total_mes", "Total del mes", 120, "e"), ("base_afecta", "Base afecta", 110, "e"),
                    ("neto_mes", "Neto del mes", 120, "e"),
                    ("llego_por_calculo", "¿Llegó por su cálculo?", 130, "center")]
        return [("dni", "DNI vinculado", 110, "w"), ("codigo_reloj", "Código del reloj", 130, "w"),
                ("nombre_reloj", "Nombre en el reloj", 260, "w"), ("fecha", "Fecha", 100, "center"),
                ("hora", "Hora", 80, "center"), ("tipo_pase", "Tipo de pase", 150, "w"),
                ("metodo", "Método", 130, "w"), ("fuente", "Archivo fuente", 260, "w")]

    def generar_reporte(self):
        nombre = self.sub_reportes.get()
        periodo = periodo_desde_entry(self.ent_periodo_reporte.get())
        dias = nc.dias_del_periodo(periodo) if periodo else []
        self.aviso("⏳ Generando reporte de %s..." % nombre, COLOR_ALERTA)
        letrero = LetreroCarga(self.frame_main, "Leyendo la base de datos...",
                               "Preparando el reporte de %s" % nombre.lower())

        def tarea(estado):
            if nombre == "Puntualidad":
                estado["datos"] = npl.reporte_puntualidad(periodo)
            elif nombre == "Tardanzas":
                estado["datos"] = npl.reporte_tardanzas(periodo)
            elif nombre == "Faltas":
                estado["datos"] = npl.reporte_faltas(periodo)
            elif nombre == "Horas extra":
                estado["datos"] = npl.reporte_horas_extra(periodo)
            elif nombre == "Garantías":
                estado["datos"] = npl.reporte_garantias(periodo)
            else:
                desde = parse_fecha_texto(self.ent_rep_desde.get()) or (dias[0].strftime("%Y-%m-%d")
                                                                       if dias else None)
                hasta = parse_fecha_texto(self.ent_rep_hasta.get()) or (dias[-1].strftime("%Y-%m-%d")
                                                                       if dias else None)
                estado["datos"] = npl.reporte_marcaciones(desde, hasta) if desde and hasta else []

        def terminar(estado):
            letrero.cerrar()
            if estado.get("error"):
                self.aviso("❌ Error generando el reporte", COLOR_ERROR)
                return
            for widget in self.contenedor_reporte.winfo_children():
                widget.destroy()
            columnas = self._columnas_reporte(nombre)
            self.tabla_reporte = crear_tabla(self.contenedor_reporte, columnas, alto=18)
            self.datos_reporte = estado.get("datos") or []
            claves = [c[0] for c in columnas]
            for registro in self.datos_reporte:
                valores = []
                for clave in claves:
                    valor = registro.get(clave, "")
                    if clave in ("monto_estimado",) or clave.startswith("total_"):
                        valores.append(formatear_monto(valor))
                    elif clave == "aprobado":
                        valores.append("SÍ" if valor else "NO")
                    elif clave in ("horas", "horas_extra", "horas_trabajadas"):
                        valores.append("%.2f" % nc.a_float(valor))
                    else:
                        valores.append("" if valor is None else str(valor))
                self.tabla_reporte.insert("", tk.END, values=valores)
            self.aviso("✅ %s: %d registro(s)" % (nombre, len(self.datos_reporte)), COLOR_OK)

        ejecutar_en_hilo(self.parent_frame, tarea, al_terminar=terminar)

    def exportar_reporte(self):
        nombre = self.sub_reportes.get()
        if not self.datos_reporte:
            return messagebox.showwarning("Sin datos", "Genere primero el reporte.")
        ruta = guardar_archivo_dialogo("Exportar reporte a Excel", ".xlsx",
                                       "Reporte_%s_%s.xlsx" % (nombre.replace(" ", "_"),
                                                               datetime.now().strftime("%Y%m%d")),
                                       [("Archivos Excel", "*.xlsx")])
        if not ruta:
            return
        columnas = [c[1] for c in self._columnas_reporte(nombre)]
        claves = [c[0] for c in self._columnas_reporte(nombre)]
        filas = [[registro.get(clave, "") for clave in claves] for registro in self.datos_reporte]
        ok, error = nc.exportar_excel(ruta, {nombre[:31]: (columnas, filas)})
        if not ok:
            return messagebox.showerror("Error", "No se pudo exportar:\n%s" % error)
        if messagebox.askyesno("Exportación lista", "Reporte guardado en:\n%s\n\n¿Desea abrirlo?" % ruta):
            abrir_documento(ruta)

    # =====================================================
    # PESTAÑA: CONFIGURACIÓN
    # =====================================================
    def construir_tab_configuracion(self):
        tab = self.tab_config
        self.sub_config = ctk.CTkSegmentedButton(
            tab, values=["Parámetros de cálculo", "Feriados y calendario", "Herramientas"],
            selected_color=COLOR_PRIMARIO, selected_hover_color=COLOR_HOVER,
            command=self.cambiar_sub_config)
        self.sub_config.pack(fill="x", padx=10, pady=(10, 6))
        self.frame_parametros = ctk.CTkFrame(tab, fg_color="transparent")
        self.frame_feriados = ctk.CTkFrame(tab, fg_color="transparent")
        self.frame_herramientas = ctk.CTkFrame(tab, fg_color="transparent")
        self._construir_parametros(self.frame_parametros)
        self._construir_feriados(self.frame_feriados)
        self._construir_herramientas(self.frame_herramientas)
        self.sub_config.set("Parámetros de cálculo")
        self.cambiar_sub_config("Parámetros de cálculo")

    def cambiar_sub_config(self, valor):
        for frame in (self.frame_parametros, self.frame_feriados, self.frame_herramientas):
            frame.pack_forget()
        mapa = {"Parámetros de cálculo": self.frame_parametros, "Feriados y calendario": self.frame_feriados,
                "Herramientas": self.frame_herramientas}
        frame = mapa.get(valor, self.frame_parametros)
        frame.pack(fill="both", expand=True)
        if valor == "Feriados y calendario":
            self.cargar_feriados()

    def _construir_parametros(self, padre):
        ctk.CTkLabel(padre, text="⚙️ PARÁMETROS DE CÁLCULO DE ASISTENCIA Y PLANILLA",
                     font=("Arial", 13, "bold"), text_color=COLOR_PRIMARIO).pack(anchor="w", pady=(4, 2))
        contenedor = ctk.CTkScrollableFrame(padre, fg_color=COLOR_FONDO_SUAVE, corner_radius=10,
                                            border_width=1, border_color="#e0e0e0")
        contenedor.pack(fill="both", expand=True, pady=4)
        self.entradas_parametros = {}
        grupos = {
            "⏱️ Jornada y marcaciones": ["tolerancia_min", "refrigerio_min", "jornada_horas", "dias_base_mes",
                                          "minutos_minimos_entre_marcas"],
            "⚡ Horas extra y recargos": ["he_25_pct", "he_35_pct", "he_desde_hora", "he_tope_dia_min",
                                          "he_minimo_min", "he_bloque_min", "recargo_nocturno_pct",
                                          "recargo_feriado_pct", "hora_inicio_nocturno", "hora_fin_nocturno",
                                          "pagar_descanso_trabajado"],
            "💰 Remuneraciones": ["rmv", "uit", "onp_pct", "essalud_pct", "asignacion_familiar_pct",
                                  "descontar_tardanza", "descontar_falta", "descontar_anticipo"],
            "🎯 Garantía de sueldo y horas sobre el tope legal": [
                "garantia_sobre", "garantia_prorratear", "bono_complemento_afecto",
                "nombre_bono_complemento", "bono_excedente_activo", "bono_excedente_afecto",
                "nombre_bono_excedente"],
            "🧾 Renta de quinta categoría (opcional)": ["calcular_renta_5ta", "deduccion_uit_5ta"],
        }
        for titulo, claves in grupos.items():
            ctk.CTkLabel(contenedor, text=titulo, font=("Arial", 12, "bold"),
                         text_color=COLOR_PRIMARIO).pack(anchor="w", padx=10, pady=(10, 2))
            for clave in claves:
                valor, descripcion = nc.PARAMETROS_DEFECTO.get(clave, ("", ""))
                fila = ctk.CTkFrame(contenedor, fg_color="transparent")
                fila.pack(fill="x", padx=10, pady=2)
                ctk.CTkLabel(fila, text=clave, font=("Arial", 10, "bold"), width=210, anchor="w").pack(
                    side="left")
                # Los parámetros de opciones fijas o de SI/NO se eligen de una lista, así no
                # se pueden escribir mal (y el motor ya no tiene que rechazarlos).
                if clave in nc.PARAMETROS_OPCIONES:
                    entrada = ctk.CTkComboBox(fila, values=list(nc.PARAMETROS_OPCIONES[clave]),
                                              width=110, state="readonly")
                elif clave in nc.PARAMETROS_SI_NO:
                    entrada = ctk.CTkComboBox(fila, values=["SI", "NO"], width=110, state="readonly")
                else:
                    entrada = ctk.CTkEntry(fila, width=110)
                entrada.pack(side="left", padx=6)
                ctk.CTkLabel(fila, text=descripcion, font=("Arial", 10), text_color=COLOR_GRIS,
                             anchor="w").pack(side="left", padx=6)
                self.entradas_parametros[clave] = entrada
        acciones = ctk.CTkFrame(padre, fg_color="transparent")
        acciones.pack(fill="x", pady=6)
        ctk.CTkButton(acciones, text="💾 Guardar parámetros", width=200, fg_color=COLOR_OK,
                      hover_color="#1e8449", command=self.guardar_parametros).pack(side="left")
        ctk.CTkButton(acciones, text="↩️ Restaurar valores por defecto", width=240, fg_color=COLOR_GRIS,
                      hover_color="#606b6b", command=self.restaurar_parametros).pack(side="left", padx=6)
        ctk.CTkButton(acciones, text="🩹 Reparar valores inválidos", width=230, fg_color="#16a085",
                      hover_color="#11806a", command=self.reparar_parametros).pack(side="left", padx=6)
        ctk.CTkButton(acciones, text="🔧 Recrear tablas del módulo", width=220, fg_color="#8e44ad",
                      hover_color="#6c3483", command=self.recrear_esquema).pack(side="left", padx=6)

    def cargar_parametros(self):
        for clave, entrada in getattr(self, "entradas_parametros", {}).items():
            valor = str(self.parametros.get(clave, ""))
            try:
                if entrada.__class__.__name__ == "CTkComboBox":
                    opciones = list(entrada.cget("values") or [""])
                    entrada.set(valor if valor in opciones else opciones[0])
                else:
                    entrada.delete(0, tk.END)
                    entrada.insert(0, valor)
            except Exception:
                pass

    def guardar_parametros(self):
        """Valida y guarda los parámetros. No guarda nada si algún campo es inválido."""
        datos = {}
        for clave, entrada in self.entradas_parametros.items():
            datos[clave] = entrada.get().strip()
        # La validación es doble: aquí se avisa al usuario y el motor vuelve a validar
        _limpios, errores, advertencias = nc.validar_parametros(datos)
        if errores:
            return messagebox.showerror(
                "Parámetros inválidos",
                "No se guardó nada. Corrija estos campos:\n\n- " + "\n- ".join(errores))
        if advertencias and not messagebox.askyesno(
                "Revise estos valores",
                "Los valores son válidos, pero probablemente no sean los que desea:\n\n- "
                + "\n- ".join(advertencias) + "\n\n¿Desea guardarlos igualmente?"):
            return
        ok, error = nc.guardar_parametros(datos, self.usuario_activo)
        if not ok:
            return messagebox.showerror("Error", error)
        self.parametros = nc.obtener_parametros()
        self.cargar_parametros()
        self.aviso("✅ Parámetros guardados", COLOR_OK)
        self.cargar_panel()
        messagebox.showinfo("Parámetros", "Los parámetros se guardaron correctamente.\n\n"
                                         "Recalcule la asistencia y la planilla para aplicarlos.")

    def reparar_parametros(self):
        """Restaura los parámetros guardados vacíos o con texto inválido."""
        detalle = nc.reparar_parametros(self.usuario_activo)
        self.parametros = nc.obtener_parametros()
        self.cargar_parametros()
        self.cargar_panel()
        if not detalle:
            return messagebox.showinfo("Parámetros",
                                       "Todos los parámetros tienen un valor válido. No hubo nada que reparar.")
        texto = "\n".join("• %s: %s  →  %s" % (clave, repr(anterior), nuevo)
                          for clave, anterior, nuevo in detalle)
        messagebox.showinfo("Parámetros reparados",
                            "Se restauraron %d valor(es) por defecto:\n\n%s" % (len(detalle), texto))

    def restaurar_parametros(self):
        if not messagebox.askyesno("Confirmar", "¿Restaurar todos los parámetros a sus valores por defecto?"):
            return
        datos = {clave: valor for clave, (valor, _) in nc.PARAMETROS_DEFECTO.items()}
        nc.guardar_parametros(datos, self.usuario_activo)
        self.parametros = nc.obtener_parametros()
        self.cargar_parametros()
        self.aviso("✅ Parámetros restaurados", COLOR_OK)

    def recrear_esquema(self):
        ok, error = nc.inicializar_esquema_nomina(forzar=True)
        if not ok:
            return messagebox.showerror("Error", "No se pudieron crear las tablas:\n%s" % error)
        messagebox.showinfo("Tablas verificadas", "Las tablas del módulo de nómina están listas en la "
                                                 "base de datos.")

    def _construir_feriados(self, padre):
        ctk.CTkLabel(padre, text="📅 FERIADOS Y DÍAS NO LABORABLES", font=("Arial", 13, "bold"),
                     text_color=COLOR_PRIMARIO).pack(anchor="w", pady=(4, 2))
        ctk.CTkLabel(padre, text="Los días marcados como FERIADO se pagan con recargo si el empleado "
                                "trabaja; los NO LABORABLES no generan falta.",
                     font=("Arial", 10), text_color=COLOR_GRIS).pack(anchor="w", pady=(0, 4))
        formulario = ctk.CTkFrame(padre, fg_color=COLOR_FONDO_SUAVE, corner_radius=10, border_width=1,
                                  border_color="#e0e0e0")
        formulario.pack(fill="x", pady=4)
        fila = ctk.CTkFrame(formulario, fg_color="transparent")
        fila.pack(fill="x", padx=12, pady=8)
        ctk.CTkLabel(fila, text="Fecha:", font=("Arial", 11, "bold")).pack(side="left")
        self.ent_feriado_fecha = ctk.CTkEntry(fila, width=110, placeholder_text="DD/MM/AAAA")
        self.ent_feriado_fecha.pack(side="left", padx=4)
        ctk.CTkButton(fila, text="📅", width=32, fg_color=COLOR_PRIMARIO, hover_color=COLOR_HOVER,
                      command=lambda: self.abrir_calendario(self.ent_feriado_fecha)).pack(side="left", padx=2)
        ctk.CTkLabel(fila, text="Tipo:", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 2))
        self.cmb_feriado_tipo = ctk.CTkComboBox(fila, values=["FERIADO", "NO_LABORABLE", "LABORABLE"],
                                                width=150, state="readonly")
        self.cmb_feriado_tipo.set("FERIADO")
        self.cmb_feriado_tipo.pack(side="left")
        ctk.CTkLabel(fila, text="Descripción:", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 2))
        self.ent_feriado_descripcion = ctk.CTkEntry(fila, width=230)
        self.ent_feriado_descripcion.pack(side="left")
        # Los botones van en su propia fila: antes se salian del ancho de la ventana
        acciones_feriado = ctk.CTkFrame(formulario, fg_color="transparent")
        acciones_feriado.pack(fill="x", padx=12, pady=(0, 10))
        ctk.CTkButton(acciones_feriado, text="💾 Guardar fecha", width=140, fg_color=COLOR_OK,
                      hover_color="#1e8449",
                      command=self.guardar_feriado).pack(side="left")
        ctk.CTkButton(acciones_feriado, text="🗑️ Eliminar fecha", width=140, fg_color=COLOR_ERROR,
                      hover_color="#922b21",
                      command=self.eliminar_feriado).pack(side="left", padx=6)
        ctk.CTkButton(acciones_feriado, text="🇵🇪 Cargar feriados del año", width=215,
                      fg_color="#8e44ad", hover_color="#6c3483",
                      command=self.cargar_feriados_peru).pack(side="left", padx=6)
        self.tabla_feriados = crear_tabla(padre, [
            ("fecha", "Fecha", 120, "center"),
            ("dia", "Día", 120, "w"),
            ("tipo", "Tipo", 140, "center"),
            ("descripcion", "Descripción", 420, "w"),
        ], alto=14)

    @con_letrero("Leyendo el calendario de feriados...", "")
    def cargar_feriados(self):
        limpiar_tabla(self.tabla_feriados)
        for dia in nc.listar_calendario():
            fecha = nc.parse_fecha(dia["fecha"])
            self.tabla_feriados.insert("", tk.END, values=(
                dia["fecha"], nc.DIAS_SEMANA[fecha.weekday()] if fecha else "", dia["tipo"],
                dia.get("descripcion") or ""))

    def guardar_feriado(self):
        fecha = parse_fecha_texto(self.ent_feriado_fecha.get())
        if not fecha:
            return messagebox.showwarning("Atención", "Escriba una fecha válida (DD/MM/AAAA).")
        ok, error = nc.guardar_dia_calendario(fecha, self.cmb_feriado_tipo.get(),
                                              self.ent_feriado_descripcion.get().strip(),
                                              self.usuario_activo)
        if not ok:
            return messagebox.showerror("Error", error)
        self.ent_feriado_fecha.delete(0, tk.END)
        self.ent_feriado_descripcion.delete(0, tk.END)
        self.cargar_feriados()
        self.aviso("✅ Fecha registrada en el calendario", COLOR_OK)

    def eliminar_feriado(self):
        seleccion = self.tabla_feriados.selection()
        if not seleccion:
            return messagebox.showwarning("Atención", "Seleccione una fecha de la lista.")
        fecha = self.tabla_feriados.item(seleccion[0], "values")[0]
        if not messagebox.askyesno("Confirmar", "¿Eliminar %s del calendario?" % fecha):
            return
        nc.eliminar_dia_calendario(fecha, self.usuario_activo)
        self.cargar_feriados()

    def cargar_feriados_peru(self):
        from tkinter import simpledialog
        anio = simpledialog.askstring("Feriados del Perú", "Año:",
                                      initialvalue=str(datetime.now().year),
                                      parent=self.parent_frame)
        if not anio:
            return
        try:
            total = nc.sembrar_feriados_peru(int(anio), self.usuario_activo)
        except ValueError:
            return messagebox.showwarning("Atención", "Escriba un año válido.")
        self.cargar_feriados()
        messagebox.showinfo("Feriados cargados", "Se registraron %d feriados nacionales del año %s."
                            % (total, anio))

    def _construir_herramientas(self, padre):
        ctk.CTkLabel(padre, text="🔧 HERRAMIENTAS DE MANTENIMIENTO", font=("Arial", 13, "bold"),
                     text_color=COLOR_PRIMARIO).pack(anchor="w", pady=(4, 6))
        # Con desplazamiento: asi todas las tarjetas son alcanzables en pantallas bajas
        contenedor_herr = ctk.CTkScrollableFrame(padre, fg_color="transparent")
        contenedor_herr.pack(fill="both", expand=True)
        opciones = [
            ("🔄 Sincronizar padrón desde Choferes",
             "Crea las fichas de nómina que falten para los choferes registrados.", self.sincronizar_empleados),
            ("🧹 Borrar la asistencia calculada del período",
             "Elimina el cálculo del mes mostrado en la pestaña Asistencia (no borra marcaciones).",
             self.limpiar_asistencia_periodo),
            ("🧮 Recalcular todo el año con marcaciones",
             "Recalcula la asistencia de todos los meses que tienen marcaciones cargadas.",
             self.recalcular_anio),
            ("🔧 Recrear/verificar las tablas del módulo",
             "Crea las tablas nom_* si no existen (no borra datos).", self.recrear_esquema),
        ]
        for titulo, descripcion, comando in opciones:
            tarjeta = ctk.CTkFrame(contenedor_herr, fg_color=COLOR_FONDO_SUAVE, corner_radius=10,
                                   border_width=1, border_color="#e0e0e0")
            tarjeta.pack(fill="x", pady=4)
            ctk.CTkLabel(tarjeta, text=titulo, font=("Arial", 12, "bold")).pack(anchor="w", padx=14,
                                                                               pady=(10, 0))
            ctk.CTkLabel(tarjeta, text=descripcion, font=("Arial", 10),
                         text_color=COLOR_GRIS).pack(anchor="w", padx=14)
            ctk.CTkButton(tarjeta, text="Ejecutar", width=140, fg_color=COLOR_PRIMARIO,
                          hover_color=COLOR_HOVER, command=comando).pack(anchor="w", padx=14, pady=(6, 10))

    def limpiar_asistencia_periodo(self):
        periodo = self.periodo_actual
        if not messagebox.askyesno("Confirmar", "¿Borrar la asistencia calculada de %s?\n\n"
                                                "Deberá recalcularla después." % nc.nombre_mes(periodo)):
            return
        ok, resultado = nc.limpiar_asistencia_periodo(periodo, self.usuario_activo)
        if not ok:
            return messagebox.showerror("Error", resultado)
        self.cargar_matriz_asistencia()
        messagebox.showinfo("Listo", "Se borró la asistencia calculada del período.")

    def recalcular_anio(self):
        periodos = sorted({p["periodo"] for p in npl.listar_periodos() if p["periodo"]})
        if not periodos:
            return messagebox.showwarning("Sin datos", "No hay marcaciones cargadas para calcular.")
        if not messagebox.askyesno("Confirmar", "¿Recalcular la asistencia de %d período(s)?\n\n%s"
                                                % (len(periodos), ", ".join(nc.nombre_mes(p) for p in periodos))):
            return
        self.aviso("⏳ Recalculando %d períodos..." % len(periodos), COLOR_ALERTA)
        letrero = LetreroCarga(self.frame_main, "Calculando la asistencia...",
                               "Recalculando %d período(s)" % len(periodos))

        def tarea(estado):
            resultados = []
            for periodo in periodos:
                dias = nc.dias_del_periodo(periodo)
                if dias:
                    resultados.append(nc.recalcular_asistencia(dias[0], dias[-1], usuario=self.usuario_activo))
            estado["resultados"] = resultados

        def terminar(estado):
            letrero.cerrar()
            if estado.get("error"):
                self.aviso("❌ Error en el recálculo", COLOR_ERROR)
                return messagebox.showerror("Error", str(estado["error"]))
            resultados = estado.get("resultados") or []
            dias = sum(r.get("dias", 0) for r in resultados)
            self.aviso("✅ %d períodos recalculados (%d días)" % (len(resultados), dias), COLOR_OK)
            self.cargar_matriz_asistencia()
            self.cargar_panel()
            messagebox.showinfo("Recálculo terminado",
                                "Se recalcularon %d período(s) con un total de %d días de asistencia."
                                % (len(resultados), dias))

        ejecutar_en_hilo(self.parent_frame, tarea, al_terminar=terminar)

    def recargar_catalogos(self):
        """Vuelve a leer los catálogos y refresca todas las vistas dependientes."""
        self.aviso("⏳ Actualizando...", COLOR_ALERTA)
        letrero = LetreroCarga(self.frame_main, "Leyendo la base de datos...",
                               "Actualizando el padrón, turnos y horarios")

        def tarea(estado):
            estado["empleados"] = nc.listar_empleados()
            estado["turnos"] = nc.listar_turnos()
            estado["horarios"] = nc.listar_horarios()
            estado["parametros"] = nc.obtener_parametros()

        def terminar(estado):
            try:
                if estado.get("error"):
                    self.aviso("❌ Error al actualizar: %s" % estado["error"], COLOR_ERROR)
                    return
                self.empleados = estado.get("empleados") or []
                self.turnos = estado.get("turnos") or []
                self.horarios = estado.get("horarios") or []
                self.parametros = estado.get("parametros") or self.parametros
                self.refrescar_combos_globales()
                self.cargar_empleados()
                self.cargar_turnos()
                self.cargar_horarios()
                self.cargar_asignaciones()
                self.cargar_feriados()
                self.cargar_panel()
                self.aviso("✅ Datos actualizados", COLOR_OK)
            finally:
                letrero.cerrar()

        ejecutar_en_hilo(self.parent_frame, tarea, al_terminar=terminar)


if __name__ == "__main__":
    pass
