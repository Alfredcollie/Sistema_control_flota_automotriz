# -*- coding: utf-8 -*-
"""
INSPECCION_VEHICULAR.PY - Módulo de recepción de Inspecciones Vehiculares.
- Pestañas: "Últimas inspecciones" (la última por placa) e "Históricos" (todas).
- Clic en una fila abre el detalle automáticamente.
- Reconstruye fotos (incluidos daños) y firmas, y las guarda en disco.
- Borra el registro de Supabase tras la descarga.
"""
import os
import sys
import subprocess
import json
import base64
import io
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk
from PIL import Image

from app_paths import CONFIG_FILE
from conexion import conectar_db, liberar_conexion, registrar_auditoria

COLOR_PRIMARIO = "#eb337a"   # rosa / franja de marca
COLOR_AZUL = "#1f538d"       # botones del menú
COLOR_HOVER = "#163b65"
COLOR_TEXTO = "#7f8c8d"


def _normalizar_ruta(ruta):
    if not ruta:
        return ""
    ruta = str(ruta).strip()
    if sys.platform != "win32" and len(ruta) >= 2 and ruta[1] == ":" and ruta[0].isalpha():
        return ""
    return ruta


def _carpeta_archivos():
    """Carpeta donde el programa guarda sus archivos (ruta_drive de Configuración General)."""
    config = {}
    try:
        if os.path.exists(str(CONFIG_FILE)):
            with open(str(CONFIG_FILE), "r", encoding="utf-8") as f:
                config = json.load(f)
    except Exception:
        config = {}
    ruta = _normalizar_ruta(config.get("ruta_drive", ""))
    if ruta:
        ruta = os.path.expanduser(ruta)
        if os.path.isdir(ruta):
            try:
                prueba = os.path.join(ruta, ".escritura_programa")
                with open(prueba, "w", encoding="utf-8") as f:
                    f.write("x")
                os.remove(prueba)
                return ruta
            except Exception:
                pass
    return os.path.dirname(os.path.abspath(__file__))


def decodificar_imagen(data_b64):
    if not data_b64:
        return None
    try:
        return Image.open(io.BytesIO(base64.b64decode(data_b64)))
    except Exception:
        return None


def _datos_payload(payload):
    if not payload:
        return {}
    try:
        return json.loads(payload) if isinstance(payload, str) else dict(payload)
    except Exception:
        return {}


def _fecha_archivo(reg, datos):
    """Devuelve fecha DDMMYYYY desde fecha_hora (ISO) o la fecha de hoy."""
    fecha_raw = reg.get("fecha_hora") or datos.get("fecha_hora") or ""
    if fecha_raw:
        dias = fecha_raw[:10]
        if len(dias) == 10 and dias[4] == "-" and dias[7] == "-":
            try:
                return dias[8:10] + dias[5:7] + dias[0:4]  # DDMMYYYY
            except Exception:
                pass
    return datetime.now().strftime("%d%m%Y")


class InspeccionVehicularApp:
    def __init__(self, parent_frame, usuario_activo="No autenticado"):
        self.parent_frame = parent_frame
        self.usuario_activo = usuario_activo
        self.registros = []
        self.registros_principal = []
        self._display_principal = []
        self._display_historico = []
        self._last_clicked = None
        self.crear_interfaz()
        self.cargar_inspecciones()   # carga automática al entrar

    # ---------------- UI ----------------
    def crear_interfaz(self):
        self.frame_main = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.frame_main.pack(fill="both", expand=True, padx=15, pady=15)

        cab = ctk.CTkFrame(self.frame_main, height=56, corner_radius=0, fg_color=COLOR_PRIMARIO)
        cab.pack(fill="x", pady=(0, 12))
        cab.pack_propagate(False)
        ctk.CTkLabel(cab, text="🔍  INSPECCIÓN VEHICULAR", font=("Arial", 20, "bold"),
                     text_color="white").pack(anchor="w", padx=18, pady=10)

        # Buscador (filtra por cualquier campo: placa, chofer, fecha, id)
        self.search_var = tk.StringVar()
        f_busq = ctk.CTkFrame(self.frame_main, fg_color="transparent")
        f_busq.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(f_busq, text="🔎 Buscar:", font=("Arial", 11, "bold")).pack(side="left")
        ent_busq = ctk.CTkEntry(f_busq, textvariable=self.search_var,
                                placeholder_text="Placa, chofer, fecha o ID...",
                                width=420)
        ent_busq.pack(side="left", padx=(8, 0))
        ent_busq.bind("<KeyRelease>", lambda e: self._aplicar_filtro())
        ctk.CTkButton(f_busq, text="Limpiar", font=("Arial", 10),
                      fg_color="#555555", hover_color="#333333", width=70,
                      command=self._limpiar_busqueda).pack(side="left", padx=6)

        self.tabview = ctk.CTkTabview(self.frame_main, fg_color="transparent")
        self.tabview.pack(fill="both", expand=True)
        self.tabview.add("Últimas inspecciones")
        self.tabview.add("Históricos")

        self.tabla_principal = self._crear_pestana(self.tabview.tab("Últimas inspecciones"), ver_eliminar=False)
        self.tabla_historico = self._crear_pestana(self.tabview.tab("Históricos"), ver_eliminar=True)
        self.tabla_principal.bind("<ButtonRelease-1>", lambda e: self._abrir_detalle_click())
        self.tabla_historico.bind("<ButtonRelease-1>", lambda e: self._abrir_detalle_click())

        ctk.CTkLabel(self.frame_main, text="Haz clic sobre una inspección para ver su detalle.",
                     font=("Arial", 11, "italic"), text_color=COLOR_TEXTO).pack(pady=(8, 0))

    def _crear_pestana(self, tab_frame, ver_eliminar=False):
        f_btn = ctk.CTkFrame(tab_frame, fg_color="transparent")
        f_btn.pack(side="bottom", fill="x", pady=(0, 8))
        ctk.CTkButton(f_btn, text="⬇ Descargar y Borrar", font=("Arial", 12, "bold"),
                      fg_color=COLOR_PRIMARIO, hover_color="#c92a6b", width=200,
                      command=self._btn_descargar).pack(side="left", padx=6)
        ctk.CTkButton(f_btn, text="📄 PDF", font=("Arial", 12, "bold"),
                      fg_color=COLOR_AZUL, hover_color=COLOR_HOVER, width=90,
                      command=self._btn_pdf).pack(side="left", padx=6)
        if ver_eliminar:
            ctk.CTkButton(f_btn, text="🗑 Eliminar", font=("Arial", 12, "bold"),
                          fg_color="#c0392b", hover_color="#922b21", width=120,
                          command=self._btn_eliminar).pack(side="left", padx=6)

        f_tabla = ctk.CTkFrame(tab_frame, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True)
        return self._crear_tabla(f_tabla)

    def _crear_tabla(self, parent):
        columnas = ("id", "placa", "chofer", "fecha", "danos")
        style = ttk.Style()
        style.configure("Treeview", rowheight=30, font=("Arial", 10))
        style.configure("Treeview.Heading", font=("Arial", 11, "bold"))
        tabla = ttk.Treeview(parent, columns=columnas, show="headings", style="Treeview")
        for c, titulo, ancho, ancla in [
            ("id", "ID", 55, "center"),
            ("placa", "Placa", 130, "w"),
            ("chofer", "Chofer", 190, "w"),
            ("fecha", "Fecha / Hora", 190, "w"),
            ("danos", "Daños", 60, "center"),
        ]:
            tabla.heading(c, text=titulo)
            tabla.column(c, width=ancho, anchor=ancla)
        scroll = ttk.Scrollbar(parent, orient="vertical", command=tabla.yview)
        tabla.configure(yscrollcommand=scroll.set)
        tabla.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        scroll.pack(side="right", fill="y", padx=(0, 10), pady=10)
        return tabla

    # ---------------- Datos ----------------
    def _resumen_reg(self, reg):
        datos = _datos_payload(reg.get("payload"))
        placa = reg.get("placa") or datos.get("placa") or "SIN_PLACA"
        chofer = reg.get("chofer") or datos.get("chofer") or ""
        fecha = reg.get("fecha_hora") or datos.get("fecha_hora") or ""
        n_danos = len(datos.get("danos") or [])
        return placa, chofer, fecha, n_danos

    def cargar_inspecciones(self):
        self.registros = []
        self.registros_principal = []
        conn = conectar_db()
        if not conn:
            messagebox.showerror("Sin conexión", "No se pudo conectar a la base de datos.")
            return
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM inspecciones ORDER BY fecha_hora DESC, id DESC")
            cols = [d[0] for d in cursor.description]
            vistos = set()
            for row in cursor.fetchall():
                reg = dict(zip(cols, row))
                self.registros.append(reg)
                placa, _, _, _ = self._resumen_reg(reg)
                if placa in vistos:
                    continue
                vistos.add(placa)
                self.registros_principal.append(reg)
        except Exception as e:
            messagebox.showerror("Error", "No se pudieron cargar las inspecciones:\n" + str(e))
        finally:
            liberar_conexion(conn)
        self._aplicar_filtro()

    def _tabla_activa(self):
        tab = self.tabview.get()
        if tab == "Históricos":
            return self.tabla_historico, self._display_historico
        return self.tabla_principal, self._display_principal

    def _seleccionado(self):
        tabla, data = self._tabla_activa()
        sel = tabla.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecciona una inspección.")
            return None
        idx = tabla.index(sel[0])
        if 0 <= idx < len(data):
            return data[idx]
        return None

    def _abrir_detalle_click(self, event=None):
        tabla, _ = self._tabla_activa()
        tab = self.tabview.get()
        sel = tabla.selection()
        if not sel:
            return
        item = sel[0]
        key = (tab, item)
        if self._last_clicked == key:
            reg = self._seleccionado()
            if reg:
                self.ver_detalle(reg)
        else:
            self._last_clicked = key

    def _limpiar_busqueda(self):
        self.search_var.set("")
        self._aplicar_filtro()

    def _aplicar_filtro(self):
        self._last_clicked = None
        texto = (self.search_var.get() or "").strip().lower()
        for tabla in (self.tabla_principal, self.tabla_historico):
            tabla.delete(*tabla.get_children())
        self._display_principal = []
        self._display_historico = []

        def insertar(tabla, reg, display):
            placa, chofer, fecha, nd = self._resumen_reg(reg)
            tabla.insert("", tk.END, values=(reg.get("id"), placa, chofer, fecha, nd))
            display.append(reg)

        for reg in self.registros_principal:
            if texto and not self._coincide(reg, texto):
                continue
            insertar(self.tabla_principal, reg, self._display_principal)
        for reg in self.registros:
            if texto and not self._coincide(reg, texto):
                continue
            insertar(self.tabla_historico, reg, self._display_historico)

    def _coincide(self, reg, texto):
        placa, chofer, fecha, nd = self._resumen_reg(reg)
        campos = [str(reg.get("id", "")), placa, chofer, fecha, str(nd)]
        return any(texto in (c or "").lower() for c in campos)

    def _btn_descargar(self):
        reg = self._seleccionado()
        if reg:
            self._descargar(reg)

    def _btn_eliminar(self):
        reg = self._seleccionado()
        if not reg:
            return
        if not messagebox.askyesno("Confirmar",
                                   "¿Eliminar esta inspección de Supabase SIN descargarla?"):
            return
        try:
            conn = conectar_db()
            if not conn:
                messagebox.showerror("Sin conexión",
                                     "No se pudo conectar para eliminar el registro.")
                return
            try:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM inspecciones WHERE id = %s", (reg.get("id"),))
                conn.commit()
            finally:
                liberar_conexion(conn)
            registrar_auditoria(self.usuario_activo, "Inspección Vehicular",
                                "Eliminó la inspección " + str(reg.get("id")))
            messagebox.showinfo("Listo", "Inspección eliminada.")
            self.cargar_inspecciones()
        except Exception as e:
            messagebox.showerror("Error", "No se pudo eliminar:\n" + str(e))

    def _btn_pdf(self):
        reg = self._seleccionado()
        if not reg:
            messagebox.showwarning("Aviso", "Selecciona primero una inspección.")
            return
        self._generar_pdf(reg)

    # ---------------- Detalle ----------------
    def ver_detalle(self, reg=None):
        if reg is None:
            reg = self._seleccionado()
        if not reg:
            return
        datos = _datos_payload(reg.get("payload"))

        v = ctk.CTkToplevel(self.parent_frame)
        v.title("Detalle de Inspección - " + str(reg.get("placa") or datos.get("placa") or ""))
        v.geometry("900x780")
        try:
            v.attributes("-topmost", True)
        except Exception:
            pass
        v.grab_set()

        cont = ctk.CTkScrollableFrame(v, fg_color="#f8f9fa")
        cont.pack(fill="both", expand=True, padx=15, pady=15)

        ctk.CTkLabel(cont, text="✦ DATOS DE CONTROL", font=("Arial", 14, "bold"),
                     text_color=COLOR_AZUL).pack(anchor="w", pady=(5, 8))
        self._fila(cont, "Inspector", datos.get("inspector", ""))
        self._fila(cont, "Placas", datos.get("placa", reg.get("placa", "")))
        self._fila(cont, "Chofer", datos.get("chofer", reg.get("chofer", "")))
        self._fila(cont, "Fecha / Hora", datos.get("fecha_hora", reg.get("fecha_hora", "")))

        ctk.CTkLabel(cont, text="✦ INSPECCIÓN VISUAL", font=("Arial", 14, "bold"),
                     text_color=COLOR_AZUL).pack(anchor="w", pady=(15, 8))
        self._fila(cont, "Limpieza interior", datos.get("limpieza_interior", ""))
        self._fila(cont, "Limpieza exterior", datos.get("limpieza_exterior", ""))
        danos = datos.get("danos") or []
        if danos:
            for i, d in enumerate(danos):
                txt = str(d.get("tipo", "")) + " en " + str(d.get("zona", ""))
                if d.get("descripcion"):
                    txt += " - " + str(d["descripcion"])
                self._fila(cont, "Daño " + str(i + 1), txt)
                if d.get("foto"):
                    img = decodificar_imagen(d["foto"])
                    if img is not None:
                        thumb = img.copy()
                        thumb.thumbnail((320, 240))
                        ctk_img = ctk.CTkImage(light_image=thumb, dark_image=thumb,
                                               size=(thumb.width, thumb.height))
                        marco = ctk.CTkFrame(cont, corner_radius=8, fg_color="white")
                        marco.pack(anchor="w", pady=6)
                        lbl = ctk.CTkLabel(marco, image=ctk_img, text="")
                        lbl.pack(padx=8, pady=8)
                        lbl.bind("<Button-1>", lambda e, im=img: self._mostrar_imagen(im))
        else:
            self._fila(cont, "Daños", "Ninguno")

        ctk.CTkLabel(cont, text="✦ CHECKLIST MECÁNICO", font=("Arial", 14, "bold"),
                     text_color=COLOR_AZUL).pack(anchor="w", pady=(15, 8))
        self._fila(cont, "Aceite motor", datos.get("aceite_motor", ""))
        self._fila(cont, "Anticongelante", datos.get("anticongelante", ""))
        self._fila(cont, "Líquido frenos", datos.get("liquido_frenos", ""))
        self._fila(cont, "Combustible", datos.get("nivel_combustible", ""))
        self._fila(cont, "Nivel de gas", datos.get("nivel_gas", ""))
        for ll in (datos.get("llantas") or []):
            self._fila(cont, "Llanta " + str(ll.get("posicion", "")),
                       (str(ll.get("presion", "")) or "—") + " · " + str(ll.get("desgaste", "")))

        ctk.CTkLabel(cont, text="✦ EVIDENCIA FOTOGRÁFICA", font=("Arial", 14, "bold"),
                     text_color=COLOR_AZUL).pack(anchor="w", pady=(15, 8))
        self._galeria(cont, datos.get("fotos") or {})

        ctk.CTkLabel(cont, text="✦ FIRMAS", font=("Arial", 14, "bold"),
                     text_color=COLOR_AZUL).pack(anchor="w", pady=(15, 8))
        self._firma(cont, "Firma del chofer", datos.get("firma_chofer"))
        self._firma(cont, "Firma del receptor", datos.get("firma_receptor"))

        f_btn = ctk.CTkFrame(v, fg_color="transparent")
        f_btn.pack(fill="x", padx=15, pady=10)
        ctk.CTkButton(f_btn, text="⬇ Descargar y Borrar", font=("Arial", 13, "bold"),
                      fg_color=COLOR_PRIMARIO, hover_color="#c92a6b", width=200,
                      command=lambda: self._descargar(reg, v)).pack(side="left", padx=6)
        ctk.CTkButton(f_btn, text="📄 PDF", font=("Arial", 12, "bold"),
                      fg_color=COLOR_AZUL, hover_color=COLOR_HOVER, width=90,
                      command=lambda: self._generar_pdf(reg, v)).pack(side="left", padx=6)
        ctk.CTkButton(f_btn, text="Cerrar", font=("Arial", 12, "bold"),
                      fg_color="#555555", hover_color="#333333", width=110,
                      command=v.destroy).pack(side="left", padx=6)

    def _fila(self, parent, etiqueta, valor):
        fila = ctk.CTkFrame(parent, fg_color="transparent")
        fila.pack(fill="x", pady=2)
        ctk.CTkLabel(fila, text=etiqueta + ":", font=("Arial", 11, "bold"),
                     width=230, anchor="w").pack(side="left")
        ctk.CTkLabel(fila, text=str(valor if valor is not None else ""),
                     font=("Arial", 11), anchor="w").pack(side="left")

    def _mostrar_imagen(self, img):
        if img is None:
            return
        # Liberar el grab del detalle para poder interactuar con la foto
        grab_previo = None
        try:
            grab_previo = self.parent_frame.grab_current()
            if grab_previo is not None:
                grab_previo.grab_release()
        except Exception:
            grab_previo = None

        v = ctk.CTkToplevel(self.parent_frame)
        v.title("Vista previa")
        try:
            v.attributes("-fullscreen", True)
        except Exception:
            sw = self.parent_frame.winfo_screenwidth()
            sh = self.parent_frame.winfo_screenheight()
            v.geometry(str(sw) + "x" + str(sh))
        # Traer al frente para que se vea encima de la inspección
        try:
            v.attributes("-topmost", True)
            v.lift()
            v.focus_force()
        except Exception:
            pass

        sw = self.parent_frame.winfo_screenwidth()
        sh = self.parent_frame.winfo_screenheight()
        zoom = [1.0]
        lbl = [None]

        def render():
            fit = img.copy()
            fit.thumbnail((max(1, sw - 40), max(1, sh - 60)))
            w = max(50, int(fit.width * zoom[0]))
            h = max(50, int(fit.height * zoom[0]))
            ctk_img = ctk.CTkImage(light_image=fit, dark_image=fit, size=(w, h))
            if lbl[0] is None:
                lbl[0] = ctk.CTkLabel(v, image=ctk_img, text="")
                lbl[0].pack(fill="both", expand=True)
            else:
                lbl[0].configure(image=ctk_img)

        def on_wheel(e):
            try:
                d = e.delta
            except Exception:
                d = 0
            if d > 0:
                zoom[0] = min(zoom[0] * 1.15, 8.0)
            else:
                zoom[0] = max(zoom[0] / 1.15, 0.2)
            render()

        def cerrar(e=None):
            try:
                v.destroy()
            except Exception:
                pass
            # Devolver el foco al detalle de la inspección
            if grab_previo is not None:
                try:
                    grab_previo.grab_set()
                    grab_previo.lift()
                except Exception:
                    pass

        # Botón ✕ Cerrar (arriba a la derecha)
        try:
            ctk.CTkButton(v, text="✕ Cerrar", font=("Arial", 12, "bold"),
                          fg_color="#c0392b", hover_color="#922b21", width=90,
                          command=cerrar).place(relx=1.0, x=-8, y=8, anchor="ne")
        except Exception:
            pass
        # Ayuda de zoom (arriba al centro)
        try:
            ctk.CTkLabel(v, text="Rueda = Zoom · clic o ✕ = cerrar",
                         font=("Arial", 11, "italic"), text_color="white",
                         fg_color="#1a252c", corner_radius=6).place(relx=0.5, y=8, anchor="n")
        except Exception:
            pass

        render()
        try:
            v.grab_set()
        except Exception:
            pass
        try:
            lbl[0].bind("<Button-1>", cerrar)
            v.bind("<Escape>", cerrar)
            v.bind("<MouseWheel>", on_wheel)
            v.bind("<Button-4>", on_wheel)
            v.bind("<Button-5>", on_wheel)
        except Exception:
            pass

    def _galeria(self, parent, fotos):
        if not fotos:
            ctk.CTkLabel(parent, text="Sin fotografías.", text_color=COLOR_TEXTO).pack(anchor="w")
            return
        fila = ctk.CTkFrame(parent, fg_color="transparent")
        fila.pack(fill="x", pady=5)
        for lado, data in fotos.items():
            if not data:
                continue
            img = decodificar_imagen(data)
            if img is None:
                continue
            thumb = img.copy()
            thumb.thumbnail((260, 200))
            ctk_img = ctk.CTkImage(light_image=thumb, dark_image=thumb, size=(thumb.width, thumb.height))
            marco = ctk.CTkFrame(fila, corner_radius=8)
            marco.pack(side="left", padx=6, pady=6)
            ctk.CTkLabel(marco, text=str(lado).replace("_", " ").upper(),
                         font=("Arial", 10, "bold"), text_color=COLOR_AZUL).pack(pady=(6, 2))
            lbl = ctk.CTkLabel(marco, image=ctk_img, text="")
            lbl.pack(padx=6, pady=6)
            lbl.bind("<Button-1>", lambda e, im=img: self._mostrar_imagen(im))

    def _firma(self, parent, etiqueta, data):
        ctk.CTkLabel(parent, text=etiqueta, font=("Arial", 11, "bold")).pack(anchor="w", pady=(4, 2))
        img = decodificar_imagen(data)
        if img is None:
            ctk.CTkLabel(parent, text="Sin firma.", text_color=COLOR_TEXTO).pack(anchor="w")
            return
        thumb = img.copy()
        thumb.thumbnail((420, 200))
        ctk_img = ctk.CTkImage(light_image=thumb, dark_image=thumb, size=(thumb.width, thumb.height))
        marco = ctk.CTkFrame(parent, corner_radius=8, fg_color="white")
        marco.pack(anchor="w", pady=6)
        lbl = ctk.CTkLabel(marco, image=ctk_img, text="")
        lbl.pack(padx=8, pady=8)
        lbl.bind("<Button-1>", lambda e, im=img: self._mostrar_imagen(im))

    # ---------------- Descargar y borrar ----------------
    def _descargar(self, reg, ventana=None):
        if not messagebox.askyesno("Confirmar",
                                   "¿Descargar esta inspección y borrarla de Supabase?"):
            return

        carpeta = os.path.join(_carpeta_archivos(), "Inspecciones")
        os.makedirs(carpeta, exist_ok=True)
        datos = _datos_payload(reg.get("payload"))
        placa = str(reg.get("placa") or datos.get("placa") or "SIN_PLACA").strip()
        placa = placa.replace(" ", "_") or "SIN_PLACA"
        subcarpeta = os.path.join(carpeta, placa + "_" + str(reg.get("id", "0")))
        os.makedirs(subcarpeta, exist_ok=True)
        fecha = _fecha_archivo(reg, datos)
        nombre_archivo = f"{placa}_{fecha}.json"

        try:
            with open(os.path.join(subcarpeta, nombre_archivo), "w", encoding="utf-8") as f:
                json.dump(datos, f, ensure_ascii=False, indent=2)
            self._guardar_imagenes(datos, subcarpeta)

            conn = conectar_db()
            if not conn:
                messagebox.showerror("Sin conexión",
                                     "No se pudo conectar para borrar el registro.")
                return
            try:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM inspecciones WHERE id = %s", (reg.get("id"),))
                conn.commit()
            finally:
                liberar_conexion(conn)

            registrar_auditoria(self.usuario_activo, "Inspección Vehicular",
                                "Descargó y borró la inspección " + str(reg.get("id")))
            messagebox.showinfo("Listo", "Inspección descargada y borrada de Supabase.")
            if ventana is not None:
                try:
                    ventana.destroy()
                except Exception:
                    pass
            self.cargar_inspecciones()
        except Exception as e:
            messagebox.showerror("Error", "No se pudo descargar:\n" + str(e))

    def _generar_pdf(self, reg, ventana=None):
        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import mm
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib import colors
            from reportlab.lib.utils import ImageReader
        except Exception as e:
            messagebox.showerror("Error", "reportlab no está disponible:\n" + str(e), parent=ventana)
            return

        datos = _datos_payload(reg.get("payload"))
        placa = str(reg.get("placa") or datos.get("placa") or "SIN_PLACA")
        fecha = _fecha_archivo(reg, datos)
        carpeta = os.path.join(_carpeta_archivos(), "Inspecciones_PDF")
        os.makedirs(carpeta, exist_ok=True)
        sello = datetime.now().strftime("%H%M%S")
        ruta_pdf = os.path.join(carpeta, "Inspeccion_" + placa + "_" + fecha + "_" + sello + ".pdf")

        # Datos de la empresa (Configuración General)
        cfg = {}
        try:
            if os.path.exists(str(CONFIG_FILE)):
                with open(str(CONFIG_FILE), "r", encoding="utf-8") as f:
                    cfg = json.load(f)
        except Exception:
            cfg = {}
        nom_empresa = cfg.get("razon_social_empresa", "") or "BLACK RIDERS"
        ruc_empresa = cfg.get("ruc_empresa", "") or ""

        try:
            estilos = getSampleStyleSheet()
            h0 = ParagraphStyle('h0', parent=estilos['Title'], fontSize=16,
                                textColor=colors.HexColor('#333333'), spaceAfter=2)
            h1 = ParagraphStyle('h1', parent=estilos['Heading1'], fontSize=15,
                                textColor=colors.HexColor('#eb337a'), spaceAfter=4)
            h2 = ParagraphStyle('h2', parent=estilos['Heading2'], fontSize=12,
                                textColor=colors.HexColor('#1f538d'), spaceBefore=10, spaceAfter=4)
            h3 = ParagraphStyle('h3', parent=estilos['Heading3'], fontSize=10,
                                textColor=colors.HexColor('#444444'), spaceBefore=6, spaceAfter=2)
            normal = estilos['Normal']
            pequeno = ParagraphStyle('peq', parent=estilos['Normal'], fontSize=9,
                                     textColor=colors.HexColor('#888888'))

            doc = SimpleDocTemplate(ruta_pdf, pagesize=A4, topMargin=15*mm,
                                    bottomMargin=15*mm, leftMargin=15*mm, rightMargin=15*mm)
            story = []

            # Encabezado: logo a la izquierda (encajado) + info corporativa
            ruta_logo = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")
            logo_flow = None
            if os.path.exists(ruta_logo):
                try:
                    _limg = Image.open(ruta_logo)
                    _lw, _lh = _limg.size
                    if _lw > 0 and _lh > 0:
                        max_w = 38*mm
                        max_h = 25*mm
                        _ratio = min(max_w / _lw, max_h / _lh)
                        logo_flow = RLImage(ruta_logo, width=_lw * _ratio, height=_lh * _ratio)
                except Exception:
                    logo_flow = None

            info_celda = [Paragraph(nom_empresa.upper(), h0)]
            if ruc_empresa:
                info_celda.append(Paragraph("RUC: %s" % ruc_empresa, normal))
            info_celda.append(Paragraph("Sistema de Control de Flota Automotriz", pequeno))

            logo_celda = logo_flow if logo_flow is not None else ""
            header = Table([[logo_celda, info_celda]], colWidths=[45*mm, 120*mm])
            header.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                ('TOPPADDING', (0, 0), (-1, -1), 0),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ]))
            story.append(header)
            story.append(Spacer(1, 6))
            # línea separadora rosa
            story.append(Table([[""]], colWidths=[165*mm], rowHeights=[0.7],
                               style=TableStyle([('LINEBELOW', (0, 0), (-1, -1), 1.2, colors.HexColor('#eb337a')),
                                                 ('LEFTPADDING', (0, 0), (-1, -1), 0),
                                                 ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                                                 ('TOPPADDING', (0, 0), (-1, -1), 0),
                                                 ('BOTTOMPADDING', (0, 0), (-1, -1), 0)])))
            story.append(Spacer(1, 8))
            story.append(Paragraph("INSPECCIÓN VEHICULAR", h1))
            story.append(Paragraph("Placa: <b>%s</b> &nbsp;·&nbsp; Fecha: %s" % (placa, fecha), normal))
            story.append(Spacer(1, 8))

            def seccion(titulo):
                story.append(Paragraph(titulo, h2))

            def sub(titulo):
                story.append(Paragraph(titulo, h3))

            def par(etiqueta, valor):
                story.append(Paragraph("<b>%s:</b> %s" % (etiqueta, str(valor if valor is not None else "")), normal))

            def sn(val):
                return "Sí" if val else "No"

            def imagen(data_b64, ancho=58*mm):
                img = decodificar_imagen(data_b64)
                if img is None:
                    return
                iw, ih = img.size
                if iw <= 0:
                    return
                h = ancho * ih / iw
                try:
                    buf = io.BytesIO()
                    img.save(buf, format="PNG")
                    buf.seek(0)
                    story.append(RLImage(buf, width=ancho, height=h))
                    story.append(Spacer(1, 4))
                except Exception:
                    pass

            seccion("1. DATOS DE CONTROL")
            par("Inspector", datos.get("inspector", ""))
            par("Placas", datos.get("placa", placa))
            par("Chofer", datos.get("chofer", ""))
            par("Fecha / Hora", datos.get("fecha_hora", ""))

            seccion("2. INSPECCIÓN VISUAL Y CARROCERÍA")
            par("Limpieza interior", datos.get("limpieza_interior", ""))
            par("Limpieza exterior", datos.get("limpieza_exterior", ""))
            danos = datos.get("danos") or []
            if danos:
                for i, d in enumerate(danos):
                    txt = str(d.get("tipo", "")) + " en " + str(d.get("zona", ""))
                    if d.get("descripcion"):
                        txt += " - " + str(d["descripcion"])
                    par("Daño " + str(i + 1), txt)
                    imagen(d.get("foto"))
            else:
                par("Daños", "Ninguno")
            if datos.get("observaciones_visuales"):
                par("Observaciones", datos.get("observaciones_visuales"))

            seccion("3. CHECKLIST MECÁNICO Y OPERATIVO")
            sub("Niveles de líquidos")
            par("Aceite de motor", datos.get("aceite_motor", ""))
            par("Anticongelante", datos.get("anticongelante", ""))
            par("Líquido de frenos", datos.get("liquido_frenos", ""))
            par("Líquido limpiaparabrisas", datos.get("liquido_limpiaparabrisas", ""))
            sub("Llantas (presión y desgaste)")
            for ll in (datos.get("llantas") or []):
                par("Llanta " + str(ll.get("posicion", "")), (str(ll.get("presion", "")) or "—") + " · " + str(ll.get("desgaste", "")))
            sub("Sistema eléctrico")
            par("Luces altas", sn(datos.get("luces_altas")))
            par("Luces bajas", sn(datos.get("luces_bajas")))
            par("Direccionales", sn(datos.get("direccionales")))
            par("Intermitentes", sn(datos.get("intermitentes")))
            par("Stop", sn(datos.get("stop")))
            par("Claxon", sn(datos.get("claxon")))
            sub("Tablero")
            par("Nivel de combustible", datos.get("nivel_combustible", ""))
            par("Nivel de gas", datos.get("nivel_gas", ""))
            par("Check Engine", sn(datos.get("check_engine")))
            par("ABS", sn(datos.get("abs")))
            par("Otros códigos", sn(datos.get("otros_codigos")))
            if datos.get("observaciones_mecanicas"):
                par("Observaciones", datos.get("observaciones_mecanicas"))

            seccion("4. INVENTARIO DE ACCESORIOS")
            par("Gato hidráulico", sn(datos.get("gato")))
            par("Cruceta", sn(datos.get("cruceta")))
            par("Cables de corriente", sn(datos.get("cables")))
            par("Triángulos de seguridad", sn(datos.get("triangulos")))
            if datos.get("observaciones_inventario"):
                par("Observaciones", datos.get("observaciones_inventario"))

            seccion("5. EVIDENCIA FOTOGRÁFICA")
            fotos = datos.get("fotos") or {}
            celdas_foto = []
            for lado, data in fotos.items():
                if not data:
                    continue
                img = decodificar_imagen(data)
                if img is None:
                    continue
                iw, ih = img.size
                if iw <= 0:
                    continue
                ancho = 82*mm
                alto = ancho * ih / iw
                if alto > 60*mm:
                    alto = 60*mm
                    ancho = alto * iw / ih
                try:
                    buf = io.BytesIO()
                    img.save(buf, format="PNG")
                    buf.seek(0)
                    f_img = RLImage(buf, width=ancho, height=alto)
                    cap = Paragraph(str(lado).replace("_", " ").title(), pequeno)
                    celdas_foto.append([f_img, cap])
                except Exception:
                    pass
            if celdas_foto:
                filas = [celdas_foto[i:i+2] for i in range(0, len(celdas_foto), 2)]
                t_fotos = Table(filas, colWidths=[88*mm, 88*mm])
                t_fotos.setStyle(TableStyle([
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('LEFTPADDING', (0, 0), (-1, -1), 3),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                ]))
                story.append(t_fotos)

            seccion("6. FIRMAS")
            firmas_celdas = []
            for etiqueta, data in (("Firma del chofer", datos.get("firma_chofer")),
                                   ("Firma del receptor", datos.get("firma_receptor"))):
                img = decodificar_imagen(data)
                if img is None:
                    continue
                iw, ih = img.size
                if iw <= 0:
                    continue
                ancho = 60*mm
                alto = ancho * ih / iw
                if alto > 30*mm:
                    alto = 30*mm
                    ancho = alto * iw / ih
                try:
                    buf = io.BytesIO()
                    img.save(buf, format="PNG")
                    buf.seek(0)
                    f_img = RLImage(buf, width=ancho, height=alto)
                    firmas_celdas.append([f_img, Paragraph(etiqueta, pequeno)])
                except Exception:
                    pass
            if firmas_celdas:
                t_firmas = Table([firmas_celdas], colWidths=[70*mm, 70*mm])
                t_firmas.setStyle(TableStyle([
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ]))
                story.append(t_firmas)

            story.append(Spacer(1, 10))
            story.append(Paragraph("Documento generado por Sistema de Control de Flota Automotriz · %s" % datetime.now().strftime("%d/%m/%Y %H:%M"), pequeno))

            doc.build(story)
            # Abrir el PDF automáticamente con el visor predeterminado
            try:
                if sys.platform == "win32":
                    os.startfile(ruta_pdf)
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", ruta_pdf])
                else:
                    subprocess.Popen(["xdg-open", ruta_pdf])
            except Exception:
                pass
            messagebox.showinfo("PDF", "PDF generado:\n" + ruta_pdf, parent=ventana)
            if ventana is not None:
                try:
                    ventana.destroy()
                except Exception:
                    pass
        except Exception as e:
            messagebox.showerror("Error", "No se pudo generar el PDF:\n" + str(e), parent=ventana)


    def _guardar_imagenes(self, datos, carpeta):
        fotos = datos.get("fotos") or {}
        for lado, data in fotos.items():
            img = decodificar_imagen(data)
            if img is not None:
                img.save(os.path.join(carpeta, "foto_" + str(lado) + ".png"), "PNG")
        for i, d in enumerate(datos.get("danos") or []):
            if d.get("foto"):
                img = decodificar_imagen(d["foto"])
                if img is not None:
                    img.save(os.path.join(carpeta, "dano_" + str(i + 1) + ".png"), "PNG")
        for etiqueta, data in (("firma_chofer", datos.get("firma_chofer")),
                               ("firma_receptor", datos.get("firma_receptor"))):
            img = decodificar_imagen(data)
            if img is not None:
                img.save(os.path.join(carpeta, etiqueta + ".png"), "PNG")


if __name__ == "__main__":
    pass
