# -*- coding: utf-8 -*-
"""
INSPECCION_VEHICULAR.PY - Módulo de recepción de Inspecciones Vehiculares.
- Se abre desde el menú "MÓDULOS OPERATIVOS" (posicionable en Configuración General).
- Carga automáticamente la lista al entrar (y con el botón Actualizar).
- Reconstruye fotos y firmas (base64) y las guarda en disco.
- Borra el registro de Supabase tras la descarga.
- Mantiene el mismo formato visual del resto de módulos.
"""
import os
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
        self.crear_interfaz()
        self.cargar_inspecciones()   # <-- carga automática al entrar

    # ---------------- UI ----------------
    def crear_interfaz(self):
        self.frame_main = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.frame_main.pack(fill="both", expand=True, padx=15, pady=15)

        cab = ctk.CTkFrame(self.frame_main, height=56, corner_radius=0, fg_color=COLOR_PRIMARIO)
        cab.pack(fill="x", pady=(0, 15))
        cab.pack_propagate(False)
        ctk.CTkLabel(cab, text="🔍  INSPECCIÓN VEHICULAR", font=("Arial", 20, "bold"),
                     text_color="white").pack(anchor="w", padx=18, pady=10)

        f_tabla = ctk.CTkFrame(self.frame_main, corner_radius=10)
        f_tabla.pack(fill="both", expand=True)

        columnas = ("id", "placa", "chofer", "fecha", "danos")
        style = ttk.Style()
        style.configure("Treeview", rowheight=30, font=("Arial", 10))
        style.configure("Treeview.Heading", font=("Arial", 11, "bold"))

        self.tabla = ttk.Treeview(f_tabla, columns=columnas, show="headings", style="Treeview")
        for c, titulo, ancho, ancla in [
            ("id", "ID", 55, "center"),
            ("placa", "Placa", 130, "w"),
            ("chofer", "Chofer", 190, "w"),
            ("fecha", "Fecha / Hora", 190, "w"),
            ("danos", "Daños", 60, "center"),
        ]:
            self.tabla.heading(c, text=titulo)
            self.tabla.column(c, width=ancho, anchor=ancla)

        scroll_y = ttk.Scrollbar(f_tabla, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scroll_y.set)
        self.tabla.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        scroll_y.pack(side="right", fill="y", padx=(0, 10), pady=10)

        f_btn = ctk.CTkFrame(self.frame_main, fg_color="transparent")
        f_btn.pack(fill="x", pady=(15, 0))
        ctk.CTkButton(f_btn, text="⬇ Descargar y Borrar", font=("Arial", 12, "bold"),
                      fg_color=COLOR_PRIMARIO, hover_color="#c92a6b", width=180,
                      command=self.descargar_y_borrar).pack(side="right", padx=6)
        ctk.CTkButton(f_btn, text="👁 Ver Detalle", font=("Arial", 12, "bold"),
                      fg_color=COLOR_AZUL, hover_color=COLOR_HOVER, width=140,
                      command=self.ver_detalle).pack(side="right", padx=6)
        ctk.CTkButton(f_btn, text="🔄 Actualizar", font=("Arial", 12, "bold"),
                      fg_color="#555555", hover_color="#333333",
                      command=self.cargar_inspecciones).pack(side="left")

    # ---------------- Datos ----------------
    def cargar_inspecciones(self):
        self.tabla.delete(*self.tabla.get_children())
        self.registros = []
        conn = conectar_db()
        if not conn:
            messagebox.showerror("Sin conexión", "No se pudo conectar a la base de datos.")
            return
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM inspecciones ORDER BY id DESC")
            cols = [d[0] for d in cursor.description]
            for row in cursor.fetchall():
                reg = dict(zip(cols, row))
                self.registros.append(reg)
                datos = _datos_payload(reg.get("payload"))
                placa = reg.get("placa") or datos.get("placa") or ""
                chofer = reg.get("chofer") or datos.get("chofer") or ""
                fecha = reg.get("fecha_hora") or datos.get("fecha_hora") or ""
                n_danos = len(datos.get("danos") or [])
                self.tabla.insert("", tk.END, values=(reg.get("id"), placa, chofer, fecha, n_danos))
        except Exception as e:
            messagebox.showerror("Error", "No se pudieron cargar las inspecciones:\n" + str(e))
        finally:
            liberar_conexion(conn)

    def _seleccionado(self):
        sel = self.tabla.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecciona una inspección de la lista.")
            return None
        idx = self.tabla.index(sel[0])
        if 0 <= idx < len(self.registros):
            return self.registros[idx]
        return None

    # ---------------- Detalle ----------------
    def ver_detalle(self):
        reg = self._seleccionado()
        if not reg:
            return
        datos = _datos_payload(reg.get("payload"))

        v = ctk.CTkToplevel(self.parent_frame)
        v.title("Detalle de Inspección - " + str(reg.get("placa") or datos.get("placa") or ""))
        v.geometry("900x720")
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
        for d in (datos.get("danos") or []):
            txt = str(d.get("tipo", "")) + " en " + str(d.get("zona", ""))
            if d.get("descripcion"):
                txt += " - " + str(d["descripcion"])
            self._fila(cont, "Daño", txt)

        ctk.CTkLabel(cont, text="✦ CHECKLIST MECÁNICO", font=("Arial", 14, "bold"),
                     text_color=COLOR_AZUL).pack(anchor="w", pady=(15, 8))
        self._fila(cont, "Aceite motor", datos.get("aceite_motor", ""))
        self._fila(cont, "Anticongelante", datos.get("anticongelante", ""))
        self._fila(cont, "Líquido frenos", datos.get("liquido_frenos", ""))
        self._fila(cont, "Combustible", datos.get("nivel_combustible", ""))
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

    def _fila(self, parent, etiqueta, valor):
        fila = ctk.CTkFrame(parent, fg_color="transparent")
        fila.pack(fill="x", pady=2)
        ctk.CTkLabel(fila, text=etiqueta + ":", font=("Arial", 11, "bold"),
                     width=230, anchor="w").pack(side="left")
        ctk.CTkLabel(fila, text=str(valor if valor is not None else ""),
                     font=("Arial", 11), anchor="w").pack(side="left")

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
            img.thumbnail((260, 200))
            ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=(img.width, img.height))
            marco = ctk.CTkFrame(fila, corner_radius=8)
            marco.pack(side="left", padx=6, pady=6)
            ctk.CTkLabel(marco, text=str(lado).replace("_", " ").upper(),
                         font=("Arial", 10, "bold"), text_color=COLOR_AZUL).pack(pady=(6, 2))
            ctk.CTkLabel(marco, image=ctk_img, text="").pack(padx=6, pady=6)

    def _firma(self, parent, etiqueta, data):
        ctk.CTkLabel(parent, text=etiqueta, font=("Arial", 11, "bold")).pack(anchor="w", pady=(4, 2))
        img = decodificar_imagen(data)
        if img is None:
            ctk.CTkLabel(parent, text="Sin firma.", text_color=COLOR_TEXTO).pack(anchor="w")
            return
        img.thumbnail((420, 200))
        ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=(img.width, img.height))
        marco = ctk.CTkFrame(parent, corner_radius=8, fg_color="white")
        marco.pack(anchor="w", pady=6)
        ctk.CTkLabel(marco, image=ctk_img, text="").pack(padx=8, pady=8)

    # ---------------- Descargar y borrar ----------------
    def descargar_y_borrar(self):
        reg = self._seleccionado()
        if not reg:
            return
        if not messagebox.askyesno("Confirmar",
                                   "¿Descargar esta inspección y borrarla de Supabase?"):
            return

        # Guarda en la misma carpeta de archivos configurada en Configuración General.
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
            self.cargar_inspecciones()
        except Exception as e:
            messagebox.showerror("Error", "No se pudo descargar:\n" + str(e))

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
