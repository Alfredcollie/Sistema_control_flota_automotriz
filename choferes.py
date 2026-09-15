# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk
import calendar
import os
import sys
import json
import shutil
import subprocess
import urllib.request
import time
import threading
from datetime import datetime

# 🚀 IMPORTAMOS NUESTRAS NUEVAS HERRAMIENTAS CORPORATIVAS
from conexion import conectar_db, registrar_auditoria, liberar_conexion
from buffer_memoria import cache_sistema
from dialogos_seguros import seleccionar_archivo_dialogo, guardar_archivo_dialogo

def abrir_documento_local(ruta):
    if not ruta: return False
    # Resuelve la ruta aunque el expediente se haya guardado en otro equipo/SO
    ruta = _resolver_documento(ruta) or ruta
    ruta_norm = os.path.normpath(ruta)
    if not os.path.exists(ruta_norm):
        return False
    try:
        # Se abre SIN bloquear la ventana (en macOS 'call' puede congelar la app)
        if sys.platform == "win32":
            os.startfile(ruta_norm)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", ruta_norm], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        else:
            subprocess.Popen(["xdg-open", ruta_norm], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except Exception as e:
        print(f"Error al abrir documento: {e}")
        return False

# 🚀 FIX macOS: el selector de archivos seguro es COMPARTIDO (dialogos_seguros.py).
#    Abrir el panel nativo de Tk puede cerrar la app en macOS; en Mac se usa 'osascript'.

def _carpeta_expedientes():
    """Carpeta ESCRIBIBLE donde se guardan los expedientes y las FOTOS de los choferes.

    Orden de preferencia (la primera que exista y permita escritura):
      1) Carpeta sincronizada configurada (ruta_drive), en
         'archivos_flota/expedientes_choferes': es la carpeta visible y
         compartida entre todos los equipos (Windows y macOS).
      2) Documentos del usuario: 'Documentos/ControlFlota/archivos_flota/
         expedientes_choferes', una carpeta VISIBLE y fácil de encontrar.
      3) Junto al programa: 'archivos_flota/expedientes_choferes'.
      4) Carpeta de datos de la aplicación (oculta): último recurso.
    """
    candidatas = _candidatas_carpetas_expedientes()
    for carpeta in candidatas:
        try:
            os.makedirs(carpeta, exist_ok=True)
            prueba = os.path.join(carpeta, ".cf_escritura")
            with open(prueba, "w", encoding="utf-8") as f:
                f.write("x")
            os.remove(prueba)
            return carpeta
        except Exception:
            continue
    return candidatas[-1]


def _carpetas_datos_cercanas():
    """Carpeta(s) de datos que rodean al programa (contienen 'archivos_flota', etc.).

    Sirve cuando la aplicación vive DENTRO de la carpeta sincronizada: permite
    encontrar los archivos compartidos (fotos, documentos) aunque en ese equipo
    todavía no se haya configurado la carpeta en Configuración del Sistema.
    """
    detectadas = []
    try:
        actual = os.path.dirname(os.path.abspath(__file__))
        for _ in range(4):                       # hasta 4 niveles hacia arriba
            for nombre in ("archivos_flota", "cobranzas_generadas", "facturas_recibidas"):
                if os.path.isdir(os.path.join(actual, nombre)):
                    if actual not in detectadas:
                        detectadas.append(actual)
                    break
            padre = os.path.dirname(actual)
            if padre == actual:
                break
            actual = padre
    except Exception:
        pass
    return detectadas


def _candidatas_carpetas_expedientes():
    """Carpetas candidatas (en orden de preferencia) para los expedientes, sin crearlas."""
    candidatas = []
    drive = ""
    try:
        from app_paths import ruta_drive_configurada, DATA_DIR
        drive = os.path.expanduser(ruta_drive_configurada() or "")
        if drive and os.path.isdir(drive):
            # Carpeta sincronizada (la misma que usan los demás módulos), para que
            # las fotos y documentos se vean desde CUALQUIER equipo (Windows y Mac).
            candidatas.append(os.path.join(drive, "archivos_flota", "expedientes_choferes"))
        # Carpetas de datos que rodean al programa (app dentro de la carpeta sincronizada)
        for cercana in _carpetas_datos_cercanas():
            candidatas.append(os.path.join(cercana, "archivos_flota", "expedientes_choferes"))
        # Carpeta VISIBLE del usuario (Windows: C:\Users\...\Documents, Mac: ~/Documents)
        candidatas.append(os.path.join(os.path.expanduser("~"), "Documents", "ControlFlota",
                                       "archivos_flota", "expedientes_choferes"))
        candidatas.append(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                       "archivos_flota", "expedientes_choferes"))
        candidatas.append(os.path.join(str(DATA_DIR), "archivos_flota", "expedientes_choferes"))
    except Exception:
        candidatas.append(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                       "archivos_flota", "expedientes_choferes"))
    # Ubicaciones que se siguen revisando al BUSCAR documentos ya guardados
    if drive:
        candidatas.append(os.path.join(drive, "expedientes_choferes"))
    return [os.path.normpath(c) for c in candidatas]


def _carpetas_expedientes_posibles():
    """Todas las carpetas donde puede estar un expediente (sin crearlas)."""
    return _candidatas_carpetas_expedientes()


def _resolver_documento(ruta):
    """Devuelve la ruta REAL del documento/foto en ESTE equipo (o "" si no está).

    Resuelve el problema de los archivos cargados en otro equipo o en otro
    sistema operativo (por ejemplo, una foto subida desde la máquina principal
    con Windows): se prueba la ruta portable de app_paths y, si no, se busca el
    archivo por su nombre en las carpetas de expedientes (sincronizada y local).
    """
    if not ruta:
        return ""
    texto = str(ruta).strip()
    try:
        from app_paths import resolver_ruta_archivo
        encontrada = resolver_ruta_archivo(texto)
        if encontrada:
            return encontrada
    except Exception:
        pass
    candidatas = [texto]
    nombre = os.path.basename(str(texto).replace(chr(92), "/"))
    if nombre:
        for carpeta in _carpetas_expedientes_posibles():
            candidatas.append(os.path.join(carpeta, nombre))
    for candidata in candidatas:
        try:
            if candidata and os.path.exists(os.path.normpath(candidata)):
                return os.path.normpath(candidata)
        except Exception:
            continue
    return ""


def _existe_documento(ruta):
    return bool(_resolver_documento(ruta))


# =========================================================
# CLASE: ASISTENTE DE CARGA DE DOCUMENTOS
# =========================================================
class AsistenteCargaDocs(ctk.CTkToplevel):
    def __init__(self, parent, target_dict, documentos_faltantes, callback):
        super().__init__(parent)
        self.target_dict = target_dict
        self.callback = callback
        self.documentos = documentos_faltantes
        self.indice = 0
        
        self.title("Asistente de Carga de Expediente")
        self.geometry("400x280")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (400 // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (280 // 2)
        self.geometry(f"+{x}+{y}")
        
        self.lbl_paso = ctk.CTkLabel(self, text="", font=("Arial", 12, "bold"), text_color="gray")
        self.lbl_paso.pack(pady=(20, 5))
        
        self.lbl_doc = ctk.CTkLabel(self, text="", font=("Arial", 18, "bold"), text_color="#1f538d")
        self.lbl_doc.pack(pady=(0, 20))
        
        self.btn_cargar = ctk.CTkButton(self, text="📎 Buscar Archivo (PDF/Imagen)", height=40, font=("Arial", 12, "bold"), fg_color="#27ae60", hover_color="#1e8449", command=self.cargar_actual)
        self.btn_cargar.pack(fill="x", padx=40, pady=10)
        
        self.btn_saltar = ctk.CTkButton(self, text="⏭️ Saltar (No lo tiene)", height=40, font=("Arial", 12, "bold"), fg_color="#7f8c8d", hover_color="#606b6b", command=self.siguiente)
        self.btn_saltar.pack(fill="x", padx=40, pady=10)
        
        self.protocol("WM_DELETE_WINDOW", self.cerrar_asistente)
        self.actualizar_vista()

    def actualizar_vista(self):
        if self.indice >= len(self.documentos):
            self.cerrar_asistente()
            return
            
        doc_actual = self.documentos[self.indice]
        self.lbl_paso.configure(text=f"Paso {self.indice + 1} de {len(self.documentos)}")
        self.lbl_doc.configure(text=f"Cargar: {ETIQUETAS_DOC.get(doc_actual, doc_actual)}")

    def cargar_actual(self):
        ruta = seleccionar_archivo_dialogo(titulo=f"Seleccionar {ETIQUETAS_DOC.get(self.documentos[self.indice], self.documentos[self.indice])}", tipos=[("Documentos", "*.pdf;*.png;*.jpg;*.jpeg")])
        if ruta:
            self.target_dict[self.documentos[self.indice]] = ruta
            self.siguiente()

    def siguiente(self):
        self.indice += 1
        self.actualizar_vista()
        
    def cerrar_asistente(self):
        self.callback()
        self.destroy()

# =========================================================
# CLASE: CALENDARIO NATIVO (MEJORADO CON COMBOBOX)
# =========================================================
class CalendarioNativo(ctk.CTkToplevel):
    def __init__(self, parent, target_entry):
        super().__init__(parent)
        self.target_entry = target_entry
        self.title("Seleccionar Fecha")
        self.geometry("310x320")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (310 // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (320 // 2)
        self.geometry(f"+{x}+{y}")
        
        self.current_year = datetime.now().year
        self.current_month = datetime.now().month
        self.meses_nombres = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        
        self.header_frame = ctk.CTkFrame(self, fg_color="#1f538d", corner_radius=0)
        self.header_frame.pack(fill="x")
        
        ctk.CTkButton(self.header_frame, text="<", width=25, fg_color="transparent", text_color="white", hover_color="#163b65", font=("Arial", 14, "bold"), command=self.prev_month).pack(side="left", padx=5, pady=10)
        
        self.cmb_mes = ctk.CTkComboBox(self.header_frame, values=self.meses_nombres, width=100, command=self.cambiar_mes_combo)
        self.cmb_mes.pack(side="left", padx=2, pady=10)
        
        anios = [str(y) for y in range(datetime.now().year - 80, datetime.now().year + 20)]
        self.cmb_anio = ctk.CTkComboBox(self.header_frame, values=anios, width=75, command=self.cambiar_anio_combo)
        self.cmb_anio.pack(side="left", padx=2, pady=10)
        
        ctk.CTkButton(self.header_frame, text=">", width=25, fg_color="transparent", text_color="white", hover_color="#163b65", font=("Arial", 14, "bold"), command=self.next_month).pack(side="right", padx=5, pady=10)
        
        self.days_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.days_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        dias_semana = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]
        for i, day in enumerate(dias_semana):
            ctk.CTkLabel(self.days_frame, text=day, font=("Arial", 11, "bold"), text_color="#1f538d").grid(row=0, column=i, padx=5, pady=5)
            
        self.update_calendar()
        
    def cambiar_mes_combo(self, choice):
        self.current_month = self.meses_nombres.index(choice) + 1
        self.update_calendar()

    def cambiar_anio_combo(self, choice):
        try:
            self.current_year = int(choice)
            self.update_calendar()
        except ValueError:
            pass

    def update_calendar(self):
        self.cmb_mes.set(self.meses_nombres[self.current_month - 1])
        self.cmb_anio.set(str(self.current_year))

        for widget in self.days_frame.winfo_children():
            if int(widget.grid_info()["row"]) > 0: widget.destroy()
        
        cal = calendar.monthcalendar(self.current_year, self.current_month)
        hoy = datetime.now()
        
        for row_idx, week in enumerate(cal, start=1):
            for col_idx, day in enumerate(week):
                if day != 0:
                    btn_color = "#d4edda" if day == hoy.day and self.current_month == hoy.month and self.current_year == hoy.year else "transparent"
                    txt_color = "#155724" if btn_color == "#d4edda" else "black"
                    btn = ctk.CTkButton(self.days_frame, text=str(day), width=30, height=30, fg_color=btn_color, text_color=txt_color, hover_color="#e0e0e0", font=("Arial", 11))
                    btn.configure(command=lambda d=day: self.select_date(d))
                    btn.grid(row=row_idx, column=col_idx, padx=3, pady=2)
                    
    def prev_month(self):
        self.current_month -= 1
        if self.current_month < 1: self.current_month = 12; self.current_year -= 1
        self.update_calendar()
        
    def next_month(self):
        self.current_month += 1
        if self.current_month > 12: self.current_month = 1; self.current_year += 1
        self.update_calendar()
        
    def select_date(self, day):
        self.target_entry.delete(0, tk.END)
        self.target_entry.insert(0, f"{day:02d}/{self.current_month:02d}/{self.current_year}")
        self.destroy()

# Nombres visibles de los documentos del expediente.
# La CLAVE interna ("DNI") se conserva para no perder los archivos ya guardados.
ETIQUETAS_DOC = {"DNI": "DNI / C.E."}


# =========================================================
# CLASE: DIÁLOGO DE OBSERVACIÓN DEL ESTADO
# =========================================================
class DialogoObservacionEstado(ctk.CTkToplevel):
    """Pide la observación (motivo) del cambio de estado del chofer.

    Se abre al marcar al personal como INACTIVO y también desde el botón
    '📝 Observación / motivo'. El texto escrito queda en .result
    (None si el usuario cancela).
    """

    def __init__(self, parent, texto_inicial="", obligatorio=True,
                 titulo="Observación del estado", subtitulo=""):
        super().__init__(parent)
        self.result = None
        self.obligatorio = obligatorio
        familia = "Helvetica" if sys.platform == "darwin" else "Arial"

        self.title(titulo)
        self.geometry("480x340")
        self.resizable(False, False)
        self.transient(parent)
        try:
            self.grab_set()
        except Exception:
            pass
        self.update_idletasks()
        try:
            x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (480 // 2)
            y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (340 // 2)
            self.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

        ctk.CTkLabel(self, text="📝 Observación", font=(familia, 15, "bold"),
                     text_color="#1f538d").pack(anchor="w", padx=20, pady=(18, 4))
        ctk.CTkLabel(self, text=subtitulo or "Escriba la observación.",
                     font=(familia, 11), text_color="gray", justify="left",
                     wraplength=440).pack(anchor="w", padx=20, pady=(0, 8))

        self.txt = ctk.CTkTextbox(self, height=140, font=(familia, 12))
        self.txt.pack(fill="both", expand=True, padx=20, pady=(4, 6))
        if texto_inicial:
            self.txt.insert("1.0", texto_inicial)
        try:
            self.txt.focus_set()
        except Exception:
            pass

        ctk.CTkLabel(self, text="Obligatorio: indique el motivo." if obligatorio else "Opcional.",
                     font=(familia, 10), text_color="gray").pack(anchor="w", padx=20)

        f_btns = ctk.CTkFrame(self, fg_color="transparent")
        f_btns.pack(fill="x", padx=20, pady=(8, 16))
        ctk.CTkButton(f_btns, text="💾 Guardar", width=150, height=36,
                      font=(familia, 12, "bold"), fg_color="#27ae60", hover_color="#1e8449",
                      command=self._aceptar).pack(side="left", padx=(0, 8))
        ctk.CTkButton(f_btns, text="✖ Cancelar", width=120, height=36,
                      font=(familia, 12, "bold"), fg_color="#7f8c8d", hover_color="#606b6b",
                      command=self.destroy).pack(side="left")

        self.bind("<Escape>", lambda e: self.destroy())
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.wait_window()

    def _aceptar(self):
        texto = self.txt.get("1.0", "end").strip()
        if self.obligatorio and not texto:
            messagebox.showwarning("Falta la observación",
                                   "Escriba el motivo por el que el chofer queda inactivo.",
                                   parent=self)
            return
        self.result = texto
        self.destroy()


_SCHEMA_CHOFERES_OK = False

# =========================================================
# CLASE PRINCIPAL: PADRÓN DE CHOFERES
# =========================================================
class ChoferesApp:
    def __init__(self, parent_frame, usuario_activo):
        self.parent_frame = parent_frame
        self.usuario_activo = usuario_activo
        self.id_edicion = None
        self.rutas_documentos_temp = {}
        self.rutas_documentos_db = {}
        self._observacion_estado = ""      # motivo/observación del estado (baja, suspensión, etc.)
        self._estado_anterior = "Activo"   # para restaurar el estado si se cancela el diálogo
        self._img_foto_ctk = None
        # 🔧 Imágenes anteriores de la foto carnet: se conservan VIVAS a propósito
        # (ver _conservar_foto_anterior)
        self._fotos_previas = []
        
        # 🚀 VARIABLES DE PAGINACIÓN (LAZY LOADING)
        self.pagina_actual = 1
        self.registros_por_pagina = 50
        
        self.inicializar_bd()
        self.crear_interfaz()

    # 🚀 FIX: AUTO-CURACIÓN EN SEGUNDO PLANO
    def inicializar_bd(self):
        global _SCHEMA_CHOFERES_OK
        if _SCHEMA_CHOFERES_OK: return
        
        def tarea_curacion():
            global _SCHEMA_CHOFERES_OK
            conn = conectar_db(silencioso=True)
            if not conn: return
            try:
                cursor = conn.cursor()
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS choferes (
                        id SERIAL PRIMARY KEY,
                        dni VARCHAR(20) UNIQUE NOT NULL,
                        nombres VARCHAR(255) NOT NULL,
                        ruc VARCHAR(20),
                        telefono VARCHAR(50),
                        correo VARCHAR(150),
                        licencia VARCHAR(50),
                        categoria_licencia VARCHAR(50),
                        vencimiento_licencia VARCHAR(20),
                        estado VARCHAR(50) DEFAULT 'Activo'
                    )
                """)
                conn.commit()
                
                columnas_nuevas = [
                    "ALTER TABLE choferes ADD COLUMN direccion VARCHAR(255) DEFAULT ''",
                    "ALTER TABLE choferes ADD COLUMN seguro_salud_num VARCHAR(100) DEFAULT ''",
                    "ALTER TABLE choferes ADD COLUMN seguro_salud_venc VARCHAR(20) DEFAULT ''",
                    "ALTER TABLE choferes ADD COLUMN fecha_nacimiento VARCHAR(20) DEFAULT ''",
                    "ALTER TABLE choferes ADD COLUMN numero_hijos VARCHAR(10) DEFAULT '0'",
                    "ALTER TABLE choferes ADD COLUMN sexo VARCHAR(20) DEFAULT ''",
                    "ALTER TABLE choferes ADD COLUMN seguro_vida_num VARCHAR(100) DEFAULT ''",
                    "ALTER TABLE choferes ADD COLUMN seguro_vida_venc VARCHAR(20) DEFAULT ''",
                    "ALTER TABLE choferes ADD COLUMN movil_asignado VARCHAR(100) DEFAULT 'Ninguno / Sin Asignar'",
                    "ALTER TABLE choferes ADD COLUMN ruta_documentos TEXT DEFAULT ''",
                    "ALTER TABLE choferes ADD COLUMN fecha_inicio_contrato VARCHAR(20) DEFAULT ''",
                    "ALTER TABLE choferes ADD COLUMN fecha_fin_contrato VARCHAR(20) DEFAULT ''",
                    "ALTER TABLE choferes ADD COLUMN observacion_estado VARCHAR(300) DEFAULT ''",
                    "ALTER TABLE choferes ADD COLUMN carnet_sanidad_num VARCHAR(100) DEFAULT ''",
                    "ALTER TABLE choferes ADD COLUMN carnet_sanidad_venc VARCHAR(20) DEFAULT ''",
                    "ALTER TABLE choferes ADD COLUMN licencia2 VARCHAR(50) DEFAULT ''",
                    "ALTER TABLE choferes ADD COLUMN vencimiento_licencia2 VARCHAR(20) DEFAULT ''",
                    "ALTER TABLE choferes ADD COLUMN categoria_licencia2 VARCHAR(50) DEFAULT ''",
                    "ALTER TABLE choferes ADD COLUMN telefono_emergencia VARCHAR(50) DEFAULT ''"
                ]
                
                for query in columnas_nuevas:
                    try: cursor.execute(query); conn.commit()
                    except: conn.rollback()
                _SCHEMA_CHOFERES_OK = True
            except Exception as e:
                print(f"Error BD Choferes: {e}")
            finally:
                liberar_conexion(conn)

        threading.Thread(target=tarea_curacion, daemon=True).start()

    # 🚀 FIX: CONSULTA A SUNAT ASÍNCRONA
    def buscar_documento_api(self, event=None):
        ruc = self.ent_ruc.get().strip()
        if len(ruc) != 11 or not ruc.isdigit():
            return messagebox.showwarning("RUC Inválido", "Por favor, ingrese un RUC válido de 11 dígitos para buscar en SUNAT.")
        
        self.ent_nombres.delete(0, tk.END)
        self.ent_nombres.insert(0, "Consultando...")
        
        # IMPORTANTE (macOS): el hilo SOLO consulta internet y deja el resultado
        # en 'resultado'. Tocar la interfaz (Tk) desde otro hilo congela la app,
        # por eso la ventana se actualiza con un sondeo desde el hilo principal.
        resultado = {}

        def tarea_api():
            import ssl
            import urllib.error
            try:
                # BYPASS SSL para macOS (igual que clientes.py / proveedores.py)
                try:
                    ctx = ssl._create_unverified_context()
                except AttributeError:
                    ctx = ssl.create_default_context()
                    ctx.check_hostname = False
                    ctx.verify_mode = ssl.CERT_NONE

                url = f"https://api.apis.net.pe/v1/ruc?numero={ruc}"
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, context=ctx, timeout=8) as response:
                    if response.status == 200:
                        resultado["datos"] = json.loads(response.read().decode())
                    else:
                        resultado["error"] = "No se encontró información para este RUC."
            except urllib.error.HTTPError as e:
                if e.code in (404, 422):
                    resultado["error"] = "El RUC ingresado no existe en SUNAT o no es válido."
                else:
                    resultado["error"] = f"La API de SUNAT respondió con error {e.code}."
            except Exception as e:
                resultado["error"] = f"No se pudo conectar con la API de SUNAT:\n{e}"

        threading.Thread(target=tarea_api, daemon=True).start()
        self._esperar_api(resultado)

    def _esperar_api(self, resultado, intentos=0):
        """Espera (desde el hilo principal) el resultado de la consulta a SUNAT."""
        try:
            if "datos" in resultado:
                self._aplicar_datos_api(resultado["datos"])
                return
            if "error" in resultado:
                self._error_api(resultado["error"])
                return
            if intentos < 150:            # hasta ~15 segundos
                self.parent_frame.after(100, lambda: self._esperar_api(resultado, intentos + 1))
        except Exception:
            pass

    def _aplicar_datos_api(self, data):
        self.ent_nombres.delete(0, tk.END)
        self.ent_nombres.insert(0, data.get("nombre", ""))
        self.ent_direccion.delete(0, tk.END)
        self.ent_direccion.insert(0, data.get("direccion", ""))
        messagebox.showinfo("Éxito", "Datos de SUNAT recuperados correctamente.")
        
    def _error_api(self, mensaje):
        self.ent_nombres.delete(0, tk.END)
        messagebox.showwarning("Aviso", mensaje)

    # 🚀 FIX: CARGA DE MÓVILES CON CACHÉ Y EN SEGUNDO PLANO
    def cargar_moviles_disponibles(self):
        moviles_cache = cache_sistema.obtener("lista_vehiculos_combobox")

        if moviles_cache is not None:
            self._aplicar_moviles(moviles_cache)
            return

        if hasattr(self, 'cmb_movil'):
            self.cmb_movil.set("Cargando vehículos...")

        # El hilo solo consulta la base de datos (sin tocar Tk) y el resultado se
        # aplica desde el hilo principal con un sondeo: así no se congela en macOS.
        resultado = {}

        def tarea_moviles():
            moviles = ["Ninguno / Sin Asignar"]
            conn = conectar_db(silencioso=True)
            if conn:
                try:
                    c = conn.cursor()
                    c.execute("SELECT placa, marca, modelo FROM flota_vehiculos WHERE estado = 'Operativo'")
                    for r in c.fetchall():
                        moviles.append(f"{r[0]} | {r[1]} {r[2]}")
                    cache_sistema.guardar("lista_vehiculos_combobox", moviles)
                except Exception:
                    pass
                finally:
                    liberar_conexion(conn)
            resultado["moviles"] = moviles

        threading.Thread(target=tarea_moviles, daemon=True).start()
        self._esperar_moviles(resultado)

    def _esperar_moviles(self, resultado, intentos=0):
        """Aplica la lista de vehículos cuando el hilo termina (hilo principal)."""
        try:
            if "moviles" in resultado:
                self._aplicar_moviles(resultado["moviles"])
                return
            if intentos < 100:            # hasta ~10 segundos
                self.parent_frame.after(100, lambda: self._esperar_moviles(resultado, intentos + 1))
        except Exception:
            pass

    def _aplicar_moviles(self, moviles):
        if hasattr(self, 'cmb_movil'):
            self.cmb_movil.configure(values=moviles)
            if self.cmb_movil.get() not in moviles:
                self.cmb_movil.set("Ninguno / Sin Asignar")

    def _actualizar_lbl_carpeta(self):
        """Muestra en el formulario la carpeta donde se guardan fotos y documentos."""
        lbl = getattr(self, "lbl_carpeta_expedientes", None)
        if lbl is None:
            return
        try:
            lbl.configure(text=_carpeta_expedientes())
        except Exception:
            pass

    def abrir_carpeta_expedientes(self):
        """Abre en el explorador la carpeta de expedientes y fotos de los choferes."""
        carpeta = _carpeta_expedientes()
        try:
            if sys.platform == "win32":
                os.startfile(carpeta)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", carpeta], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            else:
                subprocess.Popen(["xdg-open", carpeta], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception as e:
            messagebox.showerror("No se pudo abrir la carpeta", f"{carpeta}\n\n{e}",
                                 parent=self.parent_frame)

    def actualizar_botones_docs(self):
        todas = {**self.rutas_documentos_db, **self.rutas_documentos_temp}
        validos = sum(1 for p in todas.values() if _existe_documento(p))
        
        if validos > 0:
            self.btn_ver_docs.configure(state="normal", fg_color="#34495e")
            self.btn_cargar_docs.configure(text=f"✅ {validos} Doc(s) Listos", fg_color="#27ae60")
        else:
            self.btn_ver_docs.configure(state="disabled", fg_color="#7f8c8d")
            self.btn_cargar_docs.configure(text="📎 Cargar Documentos", fg_color="#1f538d")

    def lanzar_asistente_carga(self):
        todas_rutas_actuales = {**self.rutas_documentos_db, **self.rutas_documentos_temp}
        docs_requeridos = ["DNI", "Brevete", "Antecedentes Penales", "Antecedentes Policiales"]
        # Compatibilidad: los expedientes antiguos guardaban "Antecedentes Judiciales"
        equivalentes = {"Antecedentes Policiales": ("Antecedentes Judiciales",)}

        faltantes = []
        for doc in docs_requeridos:
            opciones = [todas_rutas_actuales.get(doc)] + [
                todas_rutas_actuales.get(alt) for alt in equivalentes.get(doc, ())]
            if not any(_existe_documento(r) for r in opciones):
                faltantes.append(doc)
                
        if not faltantes:
            messagebox.showinfo("Expediente Completo", "Ya tienes todos los documentos cargados y guardados para este personal.")
            return

        def on_asistente_cerrado():
            self.actualizar_botones_docs()
                
        AsistenteCargaDocs(self.parent_frame.winfo_toplevel(), self.rutas_documentos_temp, faltantes, on_asistente_cerrado)

    def gestionar_documentos(self):
        todas_rutas = {**self.rutas_documentos_db, **self.rutas_documentos_temp}
        # Se resuelven las rutas: el archivo puede venir de otro equipo/SO
        rutas_validas = {k: _resolver_documento(v) for k, v in todas_rutas.items() if _existe_documento(v)}
        
        if not rutas_validas:
            messagebox.showinfo("Aviso", "No hay documentos cargados para mostrar.")
            return

        v_gestor = ctk.CTkToplevel(self.parent_frame.winfo_toplevel())
        v_gestor.title("Gestor de Expediente")
        v_gestor.geometry("450x300")
        v_gestor.transient(self.parent_frame.winfo_toplevel())
        v_gestor.grab_set()

        ctk.CTkLabel(v_gestor, text="📂 Documentos del Expediente", font=("Arial", 14, "bold"), text_color="#1f538d").pack(pady=(15, 10))

        f_lista = ctk.CTkScrollableFrame(v_gestor, fg_color="transparent")
        f_lista.pack(fill="both", expand=True, padx=15, pady=5)

        def eliminar_doc(doc_name, frame_fila):
            if messagebox.askyesno("Confirmar", f"¿Eliminar {doc_name} del expediente?", parent=v_gestor):
                if doc_name in self.rutas_documentos_db: del self.rutas_documentos_db[doc_name]
                if doc_name in self.rutas_documentos_temp: del self.rutas_documentos_temp[doc_name]
                
                frame_fila.destroy()
                self.actualizar_botones_docs()
                self._mostrar_foto_carnet()
                if not {**self.rutas_documentos_db, **self.rutas_documentos_temp}:
                    v_gestor.destroy()

        for doc_name, path in rutas_validas.items():
            f_row = ctk.CTkFrame(f_lista, fg_color="#ffffff", border_width=1, border_color="#ccc", corner_radius=6)
            f_row.pack(fill="x", pady=5, padx=5)
            
            ctk.CTkLabel(f_row, text=doc_name, font=("Arial", 12, "bold"), text_color="#333333").pack(side="left", padx=10, pady=8)
            
            btn_eliminar = ctk.CTkButton(f_row, text="❌", width=30, fg_color="#e74c3c", hover_color="#c0392b", command=lambda d=doc_name, f=f_row: eliminar_doc(d, f))
            btn_eliminar.pack(side="right", padx=(5, 10))
            
            btn_ver = ctk.CTkButton(f_row, text="👁️ Ver", width=60, fg_color="#34495e", hover_color="#2c3e50", command=lambda p=path: abrir_documento_local(p))
            btn_ver.pack(side="right", padx=5)

    # =========================================================
    # 🚀 FOTO CARNET DEL CHOFER (carga en la parte superior)
    # =========================================================
    @staticmethod
    def _quitar_imagen_label(lbl):
        """Quita la imagen que muestra una CTkLabel.

        ⚠️ customtkinter NO limpia la imagen al recibir image=None (su
        _update_image() solo actúa cuando HAY imagen), así que al abrir un chofer
        sin foto se quedaba a la vista la foto del chofer anterior. Hay que
        limpiar la etiqueta Tk interna.
        """
        interior = getattr(lbl, "_label", None)
        if interior is not None:
            try:
                interior.configure(image="")
                return
            except Exception:
                pass
        try:
            lbl.configure(image="")
        except Exception:
            pass

    def _conservar_foto_anterior(self):
        """Guarda la imagen anterior para que Tk NO la destruya.

        🔧 FIX macOS/Windows: al reemplazar la foto, si el CTkImage anterior se
        liberaba, Tk eliminaba la imagen y la etiqueta quedaba apuntando a algo
        que ya no existe ("image pyimageN doesn't exist"): por eso la foto solo
        se veía la PRIMERA vez que se abría un chofer y después salía
        "FOTO NO VÁLIDA" hasta salir y volver a entrar al módulo.
        """
        anterior = getattr(self, "_img_foto_ctk", None)
        if anterior is not None:
            lista = getattr(self, "_fotos_previas", None)
            if lista is None:
                lista = self._fotos_previas = []
            lista.append(anterior)
            del lista[:-10]          # se guardan solo las últimas 10 (memoria)
        self._img_foto_ctk = None

    def _ruta_foto_carnet(self):
        """Devuelve la ruta REAL de la foto carnet en este equipo (temp -> base de datos).

        Si la foto se cargó en otro equipo (por ejemplo, la máquina principal con
        Windows), se busca el archivo por su nombre en la carpeta de expedientes
        sincronizada, de modo que también se vea en macOS.
        """
        temp = self.rutas_documentos_temp.get("Foto Carnet")
        if temp:
            return _resolver_documento(temp) or temp
        return _resolver_documento(self.rutas_documentos_db.get("Foto Carnet", ""))

    def cargar_foto_carnet(self):
        ruta = seleccionar_archivo_dialogo(
            titulo="Seleccionar Foto Carnet del Chofer",
            tipos=[("Imágenes", "*.png *.jpg *.jpeg *.webp *.bmp"), ("Todos los Archivos", "*.*")]
        )
        if not ruta:
            return
        self.rutas_documentos_temp["Foto Carnet"] = ruta
        self._mostrar_foto_carnet()
        self.actualizar_botones_docs()

    def quitar_foto_carnet(self):
        self.rutas_documentos_temp.pop("Foto Carnet", None)
        self.rutas_documentos_db.pop("Foto Carnet", None)
        self._mostrar_foto_carnet()
        self.actualizar_botones_docs()

    def _mostrar_foto_carnet(self):
        """Muestra la foto carnet en el recuadro superior (o el estado SIN FOTO)."""
        lbl = getattr(self, "lbl_foto", None)
        if lbl is None:
            return
        # La imagen que estaba puesta pasa a la lista de "no destruir"
        self._conservar_foto_anterior()

        ruta = self._ruta_foto_carnet()
        if not ruta or not os.path.exists(os.path.normpath(ruta)):          # foto no encontrada en este equipo
            self._quitar_imagen_label(lbl)      # se borra la foto del chofer anterior
            try:
                lbl.configure(image=None, text="SIN FOTO\nCarga la foto carnet del chofer")
            except Exception:
                pass
            return
        try:
            from PIL import Image as PILImage
            img = PILImage.open(ruta)
            img.load()                       # se lee a memoria (no deja el archivo abierto)
            if img.mode not in ("RGB", "RGBA"):
                img = img.convert("RGBA")
            ancho_max, alto_max = 195, 120
            iw, ih = img.size
            escala = min(ancho_max / float(iw), alto_max / float(ih))
            nw, nh = max(1, int(iw * escala)), max(1, int(ih * escala))
            self._img_foto_ctk = ctk.CTkImage(light_image=img, dark_image=img, size=(nw, nh))
            lbl.configure(image=self._img_foto_ctk, text="")
        except Exception as e:
            print("Error mostrando foto carnet:", e)
            self._img_foto_ctk = None
            self._quitar_imagen_label(lbl)
            try:
                lbl.configure(image=None, text="FOTO NO VÁLIDA")
            except Exception:
                pass

    # =========================================================
    # 🚀 FICHA PDF DE CHOFERES (LLENADO MASIVO) - sin logística ni seguros
    # =========================================================
    def generar_ficha_pdf_chofer(self):
        try:
            from crear_ficha_chofer_pdf import generar_ficha_chofer_pdf
            ruta = generar_ficha_chofer_pdf()
        except ImportError:
            messagebox.showerror("Librería Faltante", "Para generar la Ficha PDF ejecute en su consola:\npip install reportlab")
            return
        except Exception as e:
            messagebox.showerror("Error de Generación", f"No se pudo generar la Ficha PDF:\n\n{e}")
            return
        if messagebox.askyesno("Ficha Generada", f"Ficha PDF del chofer creada correctamente:\n{ruta}\n\n¿Deseas abrirla ahora?"):
            abrir_documento_local(ruta)

    def ejecutar_importacion_ficha_pdf(self):
        try:
            from pypdf import PdfReader
        except ImportError:
            messagebox.showerror("Librería Faltante", "Para activar el lector de Fichas PDF, ejecute en su consola:\npip install pypdf")
            return

        archivo_pdf = seleccionar_archivo_dialogo(
            titulo="Seleccionar Ficha PDF de Chofer",
            tipos=[("Archivos PDF de Fichas", "*.pdf *.PDF"), ("Todos los Archivos", "*.*")]
        )
        if not archivo_pdf:
            return

        try:
            reader = PdfReader(archivo_pdf)
            fields = reader.get_fields()
            if not fields:
                messagebox.showwarning("PDF Inválido", "El archivo PDF no contiene un formulario interactivo."); return

            dni = fields.get("dni", {}).get("/V", "").strip()
            nombres = fields.get("nombres", {}).get("/V", "").strip()
            ruc = fields.get("ruc", {}).get("/V", "").strip()
            direccion = fields.get("direccion", {}).get("/V", "").strip()
            fec_nac = fields.get("fecha_nacimiento", {}).get("/V", "").strip()
            sexo = fields.get("sexo", {}).get("/V", "").strip()
            hijos = fields.get("numero_hijos", {}).get("/V", "").strip()
            telefono = fields.get("telefono", {}).get("/V", "").strip()
            correo = fields.get("correo", {}).get("/V", "").strip()
            estado = fields.get("estado_laboral", {}).get("/V", "").strip()
            licencia = fields.get("licencia", {}).get("/V", "").strip()
            categoria = fields.get("categoria_licencia", {}).get("/V", "").strip()
            venc_licencia = fields.get("venc_licencia", {}).get("/V", "").strip()

            def _valor_campo(nombre_campo):
                """Valor de un campo del PDF (sin fallar si viene vacío o nulo)."""
                try:
                    return str((fields.get(nombre_campo, {}) or {}).get("/V", "") or "").strip()
                except Exception:
                    return ""

            ini_contrato = _valor_campo("inicio_contrato")
            fin_contrato = _valor_campo("fin_contrato")
            sanidad_num = _valor_campo("carnet_sanidad")
            sanidad_venc = _valor_campo("venc_sanidad")

            # NOTA: la ficha NO trae logística ni seguros (móvil / seguros salud-vida)
            self.limpiar_formulario()
            self.ent_dni.insert(0, dni)
            self.ent_nombres.insert(0, nombres)
            self.ent_ruc.insert(0, ruc)
            self.ent_direccion.insert(0, direccion)
            self.ent_fec_nac.insert(0, fec_nac)
            if sexo in ("Masculino", "Femenino", "Otro"):
                self.cmb_sexo.set(sexo)
            self.ent_hijos.insert(0, hijos)
            self.ent_telefono.insert(0, telefono)
            self.ent_correo.insert(0, correo)
            if estado in ("Activo", "Inactivo", "Suspendido"):
                self.cmb_estado.set(estado)
            self.ent_licencia.insert(0, licencia)
            self.ent_cat_licencia.insert(0, categoria)
            self.ent_venc_licencia.insert(0, venc_licencia)
            self.ent_ini_contrato.insert(0, ini_contrato)
            self.ent_fin_contrato.insert(0, fin_contrato)
            self.ent_sanidad_num.insert(0, sanidad_num)
            self.ent_sanidad_venc.insert(0, sanidad_venc)

            messagebox.showinfo("Ficha Importada", "¡Datos extraídos del PDF!\nRevisa el formulario y dale a guardar.")
        except Exception as e:
            messagebox.showerror("Error de Lectura", f"No se pudo procesar el PDF:\n\n{e}")

    def crear_interfaz(self):
        for widget in self.parent_frame.winfo_children(): widget.destroy()

        lbl_titulo = ctk.CTkLabel(self.parent_frame, text="🧑‍✈️ PADRÓN DE CHOFERES Y PERSONAL", font=("Arial", 18, "bold"), text_color="#1f538d")
        lbl_titulo.pack(anchor="w", padx=20, pady=(15, 5))

        self.main_split = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.main_split.pack(fill="both", expand=True, padx=15, pady=5)

        # PANEL IZQUIERDO: FORMULARIO (SCROLLABLE)
        self.f_form = ctk.CTkScrollableFrame(self.main_split, width=320, corner_radius=10, fg_color="#f8f9fa", border_width=1, border_color="#e0e0e0")
        self.f_form.pack(side="left", fill="y", padx=(0, 10))

        def crear_campo(texto, placeholder=""):
            ctk.CTkLabel(self.f_form, text=texto, font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
            ent = ctk.CTkEntry(self.f_form, placeholder_text=placeholder)
            ent.pack(fill="x", padx=10, pady=(0, 10))
            return ent

        def crear_campo_fecha(texto):
            ctk.CTkLabel(self.f_form, text=texto, font=("Arial", 11, "bold"), text_color="#1f538d").pack(anchor="w", padx=10)
            f_fec = ctk.CTkFrame(self.f_form, fg_color="transparent")
            f_fec.pack(fill="x", padx=10, pady=(0, 10))
            ent = ctk.CTkEntry(f_fec, placeholder_text="DD/MM/AAAA")
            ent.pack(side="left", fill="x", expand=True)
            ctk.CTkButton(f_fec, text="📅", width=35, fg_color="#1f538d", command=lambda: CalendarioNativo(self.parent_frame.winfo_toplevel(), ent)).pack(side="right", padx=(5, 0))
            return ent

        # --- FOTO CARNET DEL CHOFER (PARTE SUPERIOR DEL FORMULARIO) ---
        ctk.CTkLabel(self.f_form, text="📷 Foto Carnet del Chofer", font=("Arial", 14, "bold")).pack(pady=(15, 8))
        f_foto = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_foto.pack(fill="x", padx=10, pady=(0, 5))
        self.lbl_foto = ctk.CTkLabel(f_foto, text="SIN FOTO\nCarga la foto carnet del chofer", width=200, height=125,
                                     corner_radius=8, fg_color="#e9ecef", text_color="#6c757d",
                                     font=("Arial", 10, "bold"), justify="center")
        self.lbl_foto.pack(fill="x")
        self._conservar_foto_anterior()
        f_foto_btns = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_foto_btns.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkButton(f_foto_btns, text="📷 Cargar Foto", font=("Arial", 11, "bold"), fg_color="#1f538d", hover_color="#163b65", command=self.cargar_foto_carnet).pack(side="left", expand=True, fill="x", padx=(0, 5))
        ctk.CTkButton(f_foto_btns, text="✖ Quitar Foto", font=("Arial", 11, "bold"), fg_color="#e74c3c", hover_color="#c0392b", command=self.quitar_foto_carnet).pack(side="right", expand=True, fill="x", padx=(5, 0))

        # --- Datos Personales ---
        ctk.CTkLabel(self.f_form, text="Datos Personales", font=("Arial", 14, "bold")).pack(pady=(5, 10))
        self.ent_dni = crear_campo("DNI / C.E.: *", "Ej: 12345678  ó  001234567")
        
        self.ent_ruc = crear_campo("RUC (Enter para auto-completar):", "Ej: 10123456789")
        self.ent_ruc.bind("<Return>", self.buscar_documento_api)
        
        self.ent_nombres = crear_campo("Nombres y Apellidos: *", "Ej: Juan Pérez")
        self.ent_direccion = crear_campo("Dirección de Residencia:", "Ej: Av. Principal 123, Lima")
        
        self.ent_fec_nac = crear_campo_fecha("Fecha de Nacimiento:")
        
        ctk.CTkLabel(self.f_form, text="Sexo:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.cmb_sexo = ctk.CTkComboBox(self.f_form, values=["Masculino", "Femenino", "Otro"], state="readonly")
        self.cmb_sexo.pack(fill="x", padx=10, pady=(0, 10))
        self.cmb_sexo.set("Masculino")
        
        self.ent_hijos = crear_campo("Número de Hijos:", "Ej: 0")
        self.ent_telefono = crear_campo("Teléfono / WhatsApp:", "Ej: 999888777")
        self.ent_emergencia = crear_campo("Teléfono de Contacto de Emergencia:", "Ej: 987654321 (familiar)")
        self.ent_correo = crear_campo("Correo Electrónico:", "Ej: correo@gmail.com")

        # --- Asignación y Seguros ---
        ctk.CTkLabel(self.f_form, text="--- Logística y Seguros ---", font=("Arial", 11, "bold"), text_color="#d35400").pack(anchor="w", padx=10, pady=(10,5))
        
        ctk.CTkLabel(self.f_form, text="Vehículo Móvil Asignado:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.cmb_movil = ctk.CTkComboBox(self.f_form, state="readonly")
        self.cmb_movil.pack(fill="x", padx=10, pady=(0, 10))
        self.cargar_moviles_disponibles()

        self.ent_salud_num = crear_campo("N° Seguro de Salud (EsSalud/EPS):", "Código del seguro")
        self.ent_salud_venc = crear_campo_fecha("Vencimiento Seguro Salud:")
        
        self.ent_vida_num = crear_campo("N° Seguro Vida Ley:", "Código Póliza Vida Ley")
        self.ent_vida_venc = crear_campo_fecha("Vencimiento Vida Ley:")

        # --- Datos de Licencia ---
        ctk.CTkLabel(self.f_form, text="--- Datos de Licencia (MTC) ---", font=("Arial", 11, "bold"), text_color="#7f8c8d").pack(anchor="w", padx=10, pady=(10,5))
        self.ent_licencia = crear_campo("N° Licencia / Brevete:", "Ej: Q12345678")
        self.ent_cat_licencia = crear_campo("Categoría:", "Ej: A-IIb")
        self.ent_venc_licencia = crear_campo_fecha("Vencimiento de Licencia / Brevete:")
        self.ent_licencia2 = crear_campo("N° Licencia / Brevete 2:", "Opcional")
        self.ent_cat_licencia2 = crear_campo("Categoría Licencia / Brevete 2:", "Ej: A-IIIb")
        self.ent_venc_licencia2 = crear_campo_fecha("Vencimiento de Licencia / Brevete 2:")

        # --- Carné de Sanidad ---
        ctk.CTkLabel(self.f_form, text="--- Carné de Sanidad ---", font=("Arial", 11, "bold"), text_color="#2e86c1").pack(anchor="w", padx=10, pady=(10,5))
        self.ent_sanidad_num = crear_campo("N° Carné de Sanidad:", "Ej: CS-001234")
        self.ent_sanidad_venc = crear_campo_fecha("Vencimiento Carné de Sanidad:")

        # --- Datos de Contrato ---
        ctk.CTkLabel(self.f_form, text="--- Contrato ---", font=("Arial", 11, "bold"), text_color="#16a085").pack(anchor="w", padx=10, pady=(10,5))
        self.ent_ini_contrato = crear_campo_fecha("Inicio de Contrato:")
        self.ent_fin_contrato = crear_campo_fecha("Culminación de Contrato:")

        ctk.CTkLabel(self.f_form, text="Estado Laboral:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.cmb_estado = ctk.CTkComboBox(self.f_form, values=["Activo", "Inactivo", "Suspendido"],
                                          state="readonly", command=self._al_cambiar_estado)
        self.cmb_estado.pack(fill="x", padx=10, pady=(0, 5))
        self.cmb_estado.set("Activo")

        ctk.CTkButton(self.f_form, text="📝 Observación / motivo", font=("Arial", 11, "bold"),
                      fg_color="#8e44ad", hover_color="#6c3483",
                      command=self.editar_observacion).pack(fill="x", padx=10, pady=(0, 3))
        self.lbl_observacion = ctk.CTkLabel(self.f_form, text="", font=("Arial", 10),
                                            text_color="#c0392b", wraplength=280, justify="left")
        self.lbl_observacion.pack(anchor="w", padx=10, pady=(0, 15))
        self._actualizar_lbl_observacion()

        # --- CARGA DE EXPEDIENTE ASISTIDA ---
        ctk.CTkLabel(self.f_form, text="--- Expediente Físico del Chofer ---", font=("Arial", 11, "bold"), text_color="#8e44ad").pack(anchor="w", padx=10, pady=(5,5))
        f_docs = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_docs.pack(fill="x", padx=10, pady=(0, 15))
        
        self.btn_cargar_docs = ctk.CTkButton(f_docs, text="📎 Cargar Documentos", font=("Arial", 11, "bold"), fg_color="#1f538d", hover_color="#163b65", command=self.lanzar_asistente_carga)
        self.btn_cargar_docs.pack(side="left", expand=True, fill="x", padx=(0, 5))
        
        self.btn_ver_docs = ctk.CTkButton(f_docs, text="👁️ Gestionar Docs.", font=("Arial", 11, "bold"), fg_color="#7f8c8d", command=self.gestionar_documentos, state="disabled")
        self.btn_ver_docs.pack(side="right", expand=True, fill="x", padx=(5, 0))

        # Carpeta REAL donde se guardan los expedientes y las fotos de los choferes
        ctk.CTkLabel(self.f_form, text="📁 Carpeta de expedientes y fotos:",
                     font=("Arial", 10, "bold"), text_color="#7f8c8d").pack(anchor="w", padx=10)
        self.lbl_carpeta_expedientes = ctk.CTkLabel(self.f_form, text="", font=("Arial", 9),
                                                    text_color="gray", wraplength=290, justify="left")
        self.lbl_carpeta_expedientes.pack(anchor="w", padx=10, pady=(0, 2))
        ctk.CTkButton(self.f_form, text="📂 Abrir carpeta de expedientes", height=28,
                      font=("Arial", 11, "bold"), fg_color="#2980b9", hover_color="#1f618d",
                      command=self.abrir_carpeta_expedientes).pack(fill="x", padx=10, pady=(0, 15))
        self._actualizar_lbl_carpeta()

        # --- FICHA PDF DEL CHOFER (LLENADO MASIVO) ---
        ctk.CTkLabel(self.f_form, text="--- Ficha PDF del Chofer (llenado masivo) ---", font=("Arial", 11, "bold"), text_color="#1e8449").pack(anchor="w", padx=10, pady=(5,5))
        f_pdf = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_pdf.pack(fill="x", padx=10, pady=(0, 15))
        self.btn_crear_ficha = ctk.CTkButton(f_pdf, text="📄 Crear Ficha PDF", font=("Arial", 11, "bold"), fg_color="#34495e", hover_color="#2c3e50", command=self.generar_ficha_pdf_chofer)
        self.btn_crear_ficha.pack(side="left", expand=True, fill="x", padx=(0, 5))
        self.btn_importar_ficha = ctk.CTkButton(f_pdf, text="📥 Cargar Ficha PDF", font=("Arial", 11, "bold"), fg_color="#27ae60", hover_color="#1e8449", command=self.ejecutar_importacion_ficha_pdf)
        self.btn_importar_ficha.pack(side="right", expand=True, fill="x", padx=(5, 0))

        f_btns = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_btns.pack(fill="x", padx=10, pady=(10, 20))
        
        self.btn_guardar = ctk.CTkButton(f_btns, text="💾 Guardar Nuevo", fg_color="#27ae60", hover_color="#1e8449", font=("Arial", 12, "bold"), command=self.guardar_chofer)
        self.btn_guardar.pack(side="left", expand=True, fill="x", padx=(0, 5))
        
        btn_limpiar = ctk.CTkButton(f_btns, text="🔄 Limpiar", fg_color="#7f8c8d", hover_color="#606b6b", font=("Arial", 12, "bold"), command=self.limpiar_y_recargar_vista)
        btn_limpiar.pack(side="right", expand=True, fill="x", padx=(5, 0))

        # PANEL DERECHO: TABLA
        f_derecho = ctk.CTkFrame(self.main_split, fg_color="transparent")
        f_derecho.pack(side="right", fill="both", expand=True)

        f_busqueda = ctk.CTkFrame(f_derecho, fg_color="transparent")
        f_busqueda.pack(fill="x", pady=(0, 5))
        ctk.CTkLabel(f_busqueda, text="🔍 Buscar:", font=("Arial", 12, "bold")).pack(side="left", padx=(0, 5))
        self.ent_buscar = ctk.CTkEntry(f_busqueda, placeholder_text="Buscar por DNI / C.E., Nombres, Licencia, Móvil...")
        self.ent_buscar.pack(side="left", fill="x", expand=True)
        
        self.ent_buscar.bind("<KeyRelease>", lambda e: self.buscar_con_retraso())
        self.ent_buscar.bind("<Return>", lambda e: self.cargar_datos(reset_pagina=True))

        f_tabla = ctk.CTkFrame(f_derecho, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True)

        # 📋 Columnas del padrón: DNI/C.E., Nombre y Apellido, Teléfono,
        # Fin de Contrato, Vencimiento del Carné de Salud y Estado.
        # ("id" se mantiene oculto porque se usa internamente para editar/eliminar)
        columnas = ("id", "dni", "nombres", "telefono", "fin_contrato", "sanidad_venc", "estado")
        self.tabla = ttk.Treeview(f_tabla, columns=columnas, show="headings")
        self.tabla.heading("id", text="ID")
        self.tabla.heading("dni", text="DNI / C.E.")
        self.tabla.heading("nombres", text="Nombres y Apellidos")
        self.tabla.heading("telefono", text="Teléfono")
        self.tabla.heading("fin_contrato", text="Fin de Contrato")
        self.tabla.heading("sanidad_venc", text="Venc. Carné de Salud")
        self.tabla.heading("estado", text="Estado")

        self.tabla.column("id", width=0, stretch=tk.NO)
        self.tabla.column("dni", width=105, anchor="center")
        self.tabla.column("nombres", width=240, anchor="w")
        self.tabla.column("telefono", width=105, anchor="center")
        self.tabla.column("fin_contrato", width=110, anchor="center")
        self.tabla.column("sanidad_venc", width=140, anchor="center")
        self.tabla.column("estado", width=90, anchor="center")

        self.tabla.config(displaycolumns=("dni", "nombres", "telefono", "fin_contrato",
                                          "sanidad_venc", "estado"))

        # Colores de aviso: estado inactivo/suspendido y contrato por vencer/vencido
        self.tabla.tag_configure("inactivo", foreground="#c0392b")
        self.tabla.tag_configure("suspendido", foreground="#d35400")
        self.tabla.tag_configure("contrato_por_vencer", background="#fff4e0")
        self.tabla.tag_configure("contrato_vencido", background="#fdecea")
        self.tabla.tag_configure("sanidad_vencida", foreground="#7d3c98")   # carné de salud vencido

        scroll_y = ctk.CTkScrollbar(f_tabla, orientation="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scroll_y.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")
        
        self.tabla.bind("<Double-1>", lambda e: self.cargar_para_edicion())

        ctk.CTkLabel(f_derecho,
                     text=("Leyenda: rojo = Inactivo / Suspendido  ·  fondo naranja = contrato por vencer "
                           "(30 días o menos)  ·  fondo rojo suave = contrato vencido  ·  "
                           "morado = carné de salud vencido"),
                     font=("Arial", 10), text_color="gray").pack(anchor="w", pady=(3, 0))

        f_acciones_tabla = ctk.CTkFrame(f_derecho, fg_color="transparent")
        f_acciones_tabla.pack(fill="x", pady=10)
        
        # 🚀 BOTONES DE PAGINACIÓN
        self.btn_ant = ctk.CTkButton(f_acciones_tabla, text="◀ Ant", width=60, command=self.pagina_anterior)
        self.btn_ant.pack(side="left", padx=2)
        
        self.lbl_pagina = ctk.CTkLabel(f_acciones_tabla, text=f"Pág {self.pagina_actual}", font=("Arial", 11, "bold"))
        self.lbl_pagina.pack(side="left", padx=5)
        
        self.btn_sig = ctk.CTkButton(f_acciones_tabla, text="Sig ▶", width=60, command=self.pagina_siguiente)
        self.btn_sig.pack(side="left", padx=2)
        
        ctk.CTkButton(f_acciones_tabla, text="✏️ Editar Seleccionado", fg_color="#34495e", hover_color="#2c3e50", font=("Arial", 12, "bold"), command=self.cargar_para_edicion).pack(side="left", padx=(15, 5))
        ctk.CTkButton(f_acciones_tabla, text="❌ Eliminar", fg_color="#e74c3c", hover_color="#c0392b", font=("Arial", 12, "bold"), command=self.eliminar_chofer).pack(side="right", padx=5)

        self.parent_frame.after(100, lambda: self.cargar_datos(reset_pagina=True))

    # =========================================================
    # 🚀 ESTADO LABORAL Y OBSERVACIÓN (motivo de la baja)
    # =========================================================
    @staticmethod
    def _fecha_valida(texto):
        """Devuelve la fecha si el texto es DD/MM/AAAA; None si está vacío o mal escrito."""
        if not texto:
            return None
        try:
            return datetime.strptime(texto.strip(), "%d/%m/%Y")
        except Exception:
            return None

    def _actualizar_lbl_observacion(self):
        """Muestra en el formulario la observación guardada del estado."""
        lbl = getattr(self, "lbl_observacion", None)
        if lbl is None:
            return
        texto = (self._observacion_estado or "").strip()
        if not texto:
            lbl.configure(text="Sin observación registrada.")
        else:
            corto = texto if len(texto) <= 180 else texto[:180] + "..."
            lbl.configure(text=f"Observación: {corto}")

    def editar_observacion(self):
        """Abre el diálogo para escribir/editar la observación del estado."""
        estado = self.cmb_estado.get()
        obligatorio = (estado == "Inactivo")
        dlg = DialogoObservacionEstado(
            self.parent_frame.winfo_toplevel(), self._observacion_estado,
            obligatorio=obligatorio, titulo="Observación del estado",
            subtitulo=(f"El personal quedará en estado {estado}. Escriba el motivo."
                       if obligatorio else
                       "Escriba cualquier observación del conductor (opcional)."))
        if dlg.result is not None:
            self._observacion_estado = dlg.result
            self._actualizar_lbl_observacion()

    def _al_cambiar_estado(self, valor):
        """Al marcar al chofer como INACTIVO se pide el motivo de la baja."""
        if valor != "Inactivo":
            self._estado_anterior = valor
            return
        # El diálogo se abre un instante después para no interferir con el
        # desplegable del combo (evita que el menú quede "atrapado").
        try:
            self.parent_frame.after(60, self._pedir_motivo_baja)
        except Exception:
            self._pedir_motivo_baja()

    def _pedir_motivo_baja(self):
        """Diálogo obligatorio con el motivo cuando el estado pasa a INACTIVO."""
        if self.cmb_estado.get() != "Inactivo":
            return
        dlg = DialogoObservacionEstado(
            self.parent_frame.winfo_toplevel(), self._observacion_estado, obligatorio=True,
            titulo="Motivo de la baja",
            subtitulo=("El chofer quedará INACTIVO. Escriba el motivo "
                       "(renuncia, cese, falta, suspensión de contrato, etc.)."))
        if dlg.result is None:
            # Canceló: se restaura el estado anterior sin cambios
            self.cmb_estado.set(self._estado_anterior or "Activo")
            return
        self._observacion_estado = dlg.result
        self._estado_anterior = "Inactivo"
        self._actualizar_lbl_observacion()

    def pagina_anterior(self):
        if self.pagina_actual > 1:
            self.pagina_actual -= 1
            self.cargar_datos()
            
    def pagina_siguiente(self):
        self.pagina_actual += 1
        self.cargar_datos()

    def buscar_con_retraso(self):
        if hasattr(self, "_busqueda_job"):
            try: self.parent_frame.after_cancel(self._busqueda_job)
            except: pass
        self._busqueda_job = self.parent_frame.after(350, lambda: self.cargar_datos(reset_pagina=True))

    def limpiar_formulario(self):
        self.id_edicion = None
        self.btn_guardar.configure(text="💾 Guardar Nuevo")
        self.ent_dni.delete(0, tk.END)
        self.ent_ruc.delete(0, tk.END)
        self.ent_nombres.delete(0, tk.END)
        self.ent_direccion.delete(0, tk.END)
        self.ent_fec_nac.delete(0, tk.END)
        self.cmb_sexo.set("Masculino")
        self.ent_hijos.delete(0, tk.END)
        self.ent_telefono.delete(0, tk.END)
        self.ent_emergencia.delete(0, tk.END)
        self.ent_correo.delete(0, tk.END)
        
        self.cargar_moviles_disponibles()
        self.ent_salud_num.delete(0, tk.END)
        self.ent_salud_venc.delete(0, tk.END)
        self.ent_vida_num.delete(0, tk.END)
        self.ent_vida_venc.delete(0, tk.END)
        
        self.ent_licencia.delete(0, tk.END)
        self.ent_cat_licencia.delete(0, tk.END)
        self.ent_venc_licencia.delete(0, tk.END)
        self.ent_licencia2.delete(0, tk.END)
        self.ent_cat_licencia2.delete(0, tk.END)
        self.ent_venc_licencia2.delete(0, tk.END)
        self.ent_sanidad_num.delete(0, tk.END)
        self.ent_sanidad_venc.delete(0, tk.END)

        self.ent_ini_contrato.delete(0, tk.END)
        self.ent_fin_contrato.delete(0, tk.END)
        self.cmb_estado.set("Activo")
        self._estado_anterior = "Activo"
        self._observacion_estado = ""
        self._actualizar_lbl_observacion()
        
        self.rutas_documentos_temp = {}
        self.rutas_documentos_db = {}
        self._mostrar_foto_carnet()
        self.actualizar_botones_docs()

    # 🚀 RECARGA COMPLETA DE LA VENTANA DEL MÓDULO (como al volver a entrar)
    def recargar_interfaz(self):
        """Reconstruye toda la vista del módulo para poder seguir registrando/editando."""
        try:
            self.id_edicion = None
            self.rutas_documentos_temp = {}
            self.rutas_documentos_db = {}
            self._conservar_foto_anterior()
            self._observacion_estado = ""
            self._estado_anterior = "Activo"
            _carpeta_expedientes()          # asegura que la carpeta exista
            self.crear_interfaz()
        except Exception as e:
            print("Error recargando interfaz:", e)

    def limpiar_y_recargar_vista(self):
        """Botón Limpiar: reinicia el estado y recarga la ventana de inmediato."""
        self.recargar_interfaz()

    # 🚀 FIX: CARGA LAZY LOADING + CACHÉ
    def cargar_datos(self, reset_pagina=False):
        if reset_pagina:
            self.pagina_actual = 1
            
        self.lbl_pagina.configure(text=f"Pág {self.pagina_actual}")

        for item in self.tabla.get_children(): 
            self.tabla.delete(item)
            
        filtro = self.ent_buscar.get().strip().lower()
        offset = (self.pagina_actual - 1) * self.registros_por_pagina
        
        clave_cache = f"choferes_v2_{filtro}_pag_{self.pagina_actual}"
        datos = cache_sistema.obtener(clave_cache)
        
        if datos is not None:
            self._pintar_datos(datos)
        else:
            self.tabla.insert("", tk.END, values=("", "", "Cargando datos...", "", "", "", ""))

            # El hilo solo consulta la base de datos: el pintado se hace desde el
            # hilo principal (en macOS, tocar Tk desde otro hilo congela la app).
            self._carga_actual = getattr(self, "_carga_actual", 0) + 1
            carga_id = self._carga_actual
            resultado = {}

            def tarea_descarga():
                conn = conectar_db(silencioso=True)
                if not conn: return
                try:
                    cursor = conn.cursor()
                    if filtro:
                        cursor.execute("""
                            SELECT id, dni, nombres, telefono, fecha_fin_contrato,
                                   COALESCE(carnet_sanidad_venc, ''), estado
                            FROM choferes 
                            WHERE dni ILIKE %s OR nombres ILIKE %s OR licencia ILIKE %s OR movil_asignado ILIKE %s
                            ORDER BY nombres ASC LIMIT %s OFFSET %s
                        """, (f"%{filtro}%", f"%{filtro}%", f"%{filtro}%", f"%{filtro}%", self.registros_por_pagina, offset))
                    else:
                        cursor.execute("SELECT id, dni, nombres, telefono, fecha_fin_contrato, COALESCE(carnet_sanidad_venc, ''), estado FROM choferes ORDER BY nombres ASC LIMIT %s OFFSET %s", (self.registros_por_pagina, offset))
                    
                    datos_db = cursor.fetchall()
                    cache_sistema.guardar(clave_cache, datos_db)
                    resultado["datos"] = datos_db
                except Exception as e:
                    print(f"Error cargando tabla de choferes: {e}")
                    resultado["datos"] = []
                finally:
                    liberar_conexion(conn)

            threading.Thread(target=tarea_descarga, daemon=True).start()
            self._esperar_tabla(resultado, carga_id)

    def _esperar_tabla(self, resultado, carga_id, intentos=0):
        """Pinta la tabla cuando el hilo termina (siempre desde el hilo principal).

        Se ignoran las respuestas de búsquedas antiguas: solo se pinta la última
        carga solicitada, así no se mezclan resultados al escribir rápido.
        """
        try:
            if carga_id != getattr(self, "_carga_actual", 0):
                return                      # ya hay una búsqueda más nueva en curso
            if "datos" in resultado:
                self._pintar_datos(resultado["datos"])
                return
            if intentos < 200:              # hasta ~20 segundos
                self.parent_frame.after(100, lambda: self._esperar_tabla(resultado, carga_id, intentos + 1))
        except Exception:
            pass

    def _pintar_datos(self, datos):
        for item in self.tabla.get_children():
            self.tabla.delete(item)

        hoy = datetime.now()
        for r in datos:
            valores = tuple(r)
            if len(valores) < 7:                     # por si llega una fila antigua
                valores = valores + ("",) * (7 - len(valores))
            etiquetas = []
            # 0=id  1=dni  2=nombres  3=telefono  4=fin de contrato  5=venc. carné de salud  6=estado
            estado = str(valores[6] or "").strip().lower()
            if estado == "inactivo":
                etiquetas.append("inactivo")
            elif estado == "suspendido":
                etiquetas.append("suspendido")
            fin_contrato = self._fecha_valida(str(valores[4] or ""))
            if fin_contrato:
                if fin_contrato < hoy:
                    etiquetas.append("contrato_vencido")
                elif (fin_contrato - hoy).days <= 30:
                    etiquetas.append("contrato_por_vencer")
            sanidad = self._fecha_valida(str(valores[5] or ""))
            if sanidad and sanidad < hoy:
                etiquetas.append("sanidad_vencida")
            self.tabla.insert("", tk.END, values=valores, tags=tuple(etiquetas))
            
        if self.pagina_actual > 1:
            self.btn_ant.configure(state="normal")
        else:
            self.btn_ant.configure(state="disabled")
            
        if len(datos) == self.registros_por_pagina:
            self.btn_sig.configure(state="normal")
        else:
            self.btn_sig.configure(state="disabled")

    def guardar_chofer(self):
        dni = self.ent_dni.get().strip()
        nombres = self.ent_nombres.get().strip().upper()
        ruc = self.ent_ruc.get().strip()
        direccion = self.ent_direccion.get().strip()
        fec_nac = self.ent_fec_nac.get().strip()
        sexo = self.cmb_sexo.get()
        hijos = self.ent_hijos.get().strip() or "0"
        
        tel = self.ent_telefono.get().strip()
        tel_emergencia = self.ent_emergencia.get().strip()   # 📞 contacto de emergencia
        correo = self.ent_correo.get().strip()
        
        movil = self.cmb_movil.get()
        salud_num = self.ent_salud_num.get().strip()
        salud_venc = self.ent_salud_venc.get().strip()
        vida_num = self.ent_vida_num.get().strip()
        vida_venc = self.ent_vida_venc.get().strip()
        
        licencia = self.ent_licencia.get().strip().upper()
        cat = self.ent_cat_licencia.get().strip().upper()
        venc = self.ent_venc_licencia.get().strip()

        licencia2 = self.ent_licencia2.get().strip().upper()
        cat2 = self.ent_cat_licencia2.get().strip().upper()
        venc2 = self.ent_venc_licencia2.get().strip()

        sanidad_num = self.ent_sanidad_num.get().strip()
        sanidad_venc = self.ent_sanidad_venc.get().strip()

        ini_contrato = self.ent_ini_contrato.get().strip()
        fin_contrato = self.ent_fin_contrato.get().strip()
        estado = self.cmb_estado.get()
        observacion = (self._observacion_estado or "").strip()

        if not dni or not nombres:
            return messagebox.showwarning("Atención", "El DNI / C.E. y los Nombres son obligatorios.")

        # Validación de las fechas de contrato (DD/MM/AAAA)
        f_ini = self._fecha_valida(ini_contrato)
        f_fin = self._fecha_valida(fin_contrato)
        if ini_contrato and f_ini is None:
            return messagebox.showwarning("Fecha inválida",
                                          "La fecha de INICIO de contrato debe tener el formato DD/MM/AAAA.")
        if fin_contrato and f_fin is None:
            return messagebox.showwarning("Fecha inválida",
                                          "La fecha de CULMINACIÓN de contrato debe tener el formato DD/MM/AAAA.")
        if f_ini and f_fin and f_fin < f_ini:
            return messagebox.showwarning("Fechas incoherentes",
                                          "La fecha de culminación no puede ser anterior a la fecha de inicio de contrato.")
        if sanidad_venc and self._fecha_valida(sanidad_venc) is None:
            return messagebox.showwarning("Fecha inválida",
                                          "El vencimiento del Carné de Sanidad debe tener el formato DD/MM/AAAA.")
        if venc2 and self._fecha_valida(venc2) is None:
            return messagebox.showwarning("Fecha inválida",
                                          "El vencimiento de la Licencia 2 debe tener el formato DD/MM/AAAA.")

        # Si queda INACTIVO es obligatorio registrar el motivo
        if estado == "Inactivo" and not observacion:
            dlg = DialogoObservacionEstado(
                self.parent_frame.winfo_toplevel(), "", obligatorio=True,
                titulo="Motivo de la baja",
                subtitulo=("El chofer quedará INACTIVO. Escriba el motivo "
                           "(renuncia, cese, falta, suspensión de contrato, etc.)."))
            if dlg.result is None:
                return
            self._observacion_estado = dlg.result
            observacion = dlg.result
            self._actualizar_lbl_observacion()

        diccionario_final = self.rutas_documentos_db.copy()
        
        if self.rutas_documentos_temp:
            # 🔒 Política de almacenamiento: sin autorización no se guarda ningún archivo
            try:
                from politica_almacenamiento import exigir_permiso
                if not exigir_permiso(self):
                    return
            except ImportError:
                pass
            try:
                # En Mac (app empaquetada) la carpeta del programa puede ser de solo lectura:
                # _carpeta_expedientes() cae automáticamente a la carpeta de datos del usuario.
                carpeta_destino = _carpeta_expedientes()
                
                for doc_name, temp_path in self.rutas_documentos_temp.items():
                    if temp_path and os.path.exists(os.path.normpath(temp_path)):
                        ext = os.path.splitext(temp_path)[1]
                        nombre_seguro = doc_name.replace(" ", "_")
                        nombre_archivo = f"Expediente_{nombre_seguro}_{dni}_{datetime.now().strftime('%Y%m%d%H%M%S')}{ext}"
                        ruta_final = os.path.join(carpeta_destino, nombre_archivo)
                        
                        shutil.copy2(os.path.normpath(temp_path), ruta_final)
                        # Ruta PORTABLE (relativa a la carpeta sincronizada): sirve en Windows y en Mac
                        try:
                            from app_paths import ruta_para_guardar
                            diccionario_final[doc_name] = ruta_para_guardar(ruta_final)
                        except Exception:
                            diccionario_final[doc_name] = ruta_final
                        
            except Exception as e:
                return messagebox.showerror("Error de Archivo", f"No se pudieron copiar los archivos del expediente:\n{e}")

        json_rutas_finales = json.dumps(diccionario_final)

        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            if self.id_edicion:
                cursor.execute("""
                    UPDATE choferes SET 
                    dni=%s, nombres=%s, ruc=%s, telefono=%s, correo=%s, licencia=%s, 
                    categoria_licencia=%s, vencimiento_licencia=%s, estado=%s,
                    direccion=%s, fecha_nacimiento=%s, sexo=%s, numero_hijos=%s,
                    movil_asignado=%s, seguro_salud_num=%s, seguro_salud_venc=%s, seguro_vida_num=%s, seguro_vida_venc=%s,
                    ruta_documentos=%s, fecha_inicio_contrato=%s, fecha_fin_contrato=%s,
                    observacion_estado=%s, carnet_sanidad_num=%s, carnet_sanidad_venc=%s,
                    licencia2=%s, categoria_licencia2=%s, vencimiento_licencia2=%s,
                    telefono_emergencia=%s
                    WHERE id=%s
                """, (dni, nombres, ruc, tel, correo, licencia, cat, venc, estado,
                      direccion, fec_nac, sexo, hijos, movil, salud_num, salud_venc, vida_num, vida_venc,
                      json_rutas_finales, ini_contrato, fin_contrato, observacion,
                      sanidad_num, sanidad_venc, licencia2, cat2, venc2, tel_emergencia, self.id_edicion))
                detalle_estado = f" — marcado INACTIVO. Motivo: {observacion}" if estado == "Inactivo" else ""
                registrar_auditoria(self.usuario_activo, "Choferes",
                                    f"Actualizó datos de {nombres}{detalle_estado}"[:240])
                messagebox.showinfo("Éxito", "Datos actualizados correctamente.")
            else:
                cursor.execute("SELECT id FROM choferes WHERE dni = %s", (dni,))
                if cursor.fetchone():
                    liberar_conexion(conn)
                    return messagebox.showwarning("Duplicado", f"El DNI / C.E. {dni} ya existe en el sistema.")
                    
                cursor.execute("""
                    INSERT INTO choferes (dni, nombres, ruc, telefono, correo, licencia, categoria_licencia, vencimiento_licencia, estado, 
                    direccion, fecha_nacimiento, sexo, numero_hijos, movil_asignado, seguro_salud_num, seguro_salud_venc, seguro_vida_num, seguro_vida_venc, ruta_documentos,
                    fecha_inicio_contrato, fecha_fin_contrato, observacion_estado,
                    carnet_sanidad_num, carnet_sanidad_venc, licencia2, categoria_licencia2,
                    vencimiento_licencia2, telefono_emergencia) 
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (dni, nombres, ruc, tel, correo, licencia, cat, venc, estado,
                      direccion, fec_nac, sexo, hijos, movil, salud_num, salud_venc, vida_num, vida_venc,
                      json_rutas_finales, ini_contrato, fin_contrato, observacion,
                      sanidad_num, sanidad_venc, licencia2, cat2, venc2, tel_emergencia))
                detalle_estado = f" — INACTIVO. Motivo: {observacion}" if estado == "Inactivo" else ""
                registrar_auditoria(self.usuario_activo, "Choferes",
                                    f"Registró nuevo conductor/personal: {nombres}{detalle_estado}"[:240])
                messagebox.showinfo("Éxito", "Personal registrado correctamente.")
            
            try:
                c_crono = conn.cursor()
                identificador_crono = f"{nombres} (DNI: {dni})"
                c_crono.execute("DELETE FROM tareas_evento WHERE evento_asociado = 'FLOTA | Vencimientos' AND responsable = %s", (identificador_crono,))
                
                vencimientos = [
                    (f"Venc. Licencia ({cat})", venc),
                    (f"Venc. Licencia 2 ({cat2})" if cat2 else "Venc. Licencia 2", venc2),
                    ("Venc. Seguro Salud (EsSalud/EPS)", salud_venc),
                    ("Venc. Seguro Vida Ley", vida_venc),
                    ("Venc. Carné de Sanidad", sanidad_venc),
                    ("Fin de Contrato", fin_contrato)
                ]
                
                if estado == 'Activo':
                    for nom_doc, fec_doc in vencimientos:
                        if fec_doc and fec_doc.strip():
                            c_crono.execute("SELECT COALESCE(MAX(orden), 0) FROM tareas_evento WHERE evento_asociado = 'FLOTA | Vencimientos'")
                            nuevo_orden = c_crono.fetchone()[0] + 1
                            
                            c_crono.execute("""
                                INSERT INTO tareas_evento (evento_asociado, nombre_tarea, responsable, fecha_limite, estado, notas, orden, tipo_pago)
                                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                            """, ("FLOTA | Vencimientos", nom_doc, identificador_crono, fec_doc, "Pendiente", f"Alerta automática de RRHH/Logística para {nombres}.", nuevo_orden, "No aplica"))
            except Exception as e_crono: print("Aviso - Sincronización Cronograma:", e_crono)

            conn.commit()
            cache_sistema.invalidar() # 🚀 FIX: Borrar caché
            # 🚀 RECARGA COMPLETA: la vista se reconstruye para poder seguir registrando/editando
            self.parent_frame.after(150, self.recargar_interfaz)
        except Exception as e:
            conn.rollback()
            messagebox.showerror("Error", str(e))
        finally:
            liberar_conexion(conn)

    def cargar_para_edicion(self):
        sel = self.tabla.selection()
        if not sel: return
        vid = self.tabla.item(sel[0], "values")[0]
        
        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id, dni, nombres, ruc, telefono, correo, licencia, categoria_licencia, vencimiento_licencia, estado,
                direccion, fecha_nacimiento, sexo, numero_hijos, movil_asignado, seguro_salud_num, seguro_salud_venc, seguro_vida_num, seguro_vida_venc,
                ruta_documentos, fecha_inicio_contrato, fecha_fin_contrato, observacion_estado,
                carnet_sanidad_num, carnet_sanidad_venc, licencia2, categoria_licencia2,
                vencimiento_licencia2, telefono_emergencia
                FROM choferes WHERE id = %s
            """, (vid,))
            r = cursor.fetchone()
            if r:
                self.limpiar_formulario()
                self.id_edicion = r[0]
                self.btn_guardar.configure(text="💾 Actualizar Conductor")
                
                self.ent_dni.insert(0, r[1] if r[1] else "")
                self.ent_nombres.insert(0, r[2] if r[2] else "")
                self.ent_ruc.insert(0, r[3] if r[3] else "")
                self.ent_telefono.insert(0, r[4] if r[4] else "")
                self.ent_correo.insert(0, r[5] if r[5] else "")
                self.ent_licencia.insert(0, r[6] if r[6] else "")
                self.ent_cat_licencia.insert(0, r[7] if r[7] else "")
                self.ent_venc_licencia.insert(0, r[8] if r[8] else "")
                if r[9]: self.cmb_estado.set(r[9])
                
                self.ent_direccion.insert(0, r[10] if r[10] else "")
                self.ent_fec_nac.insert(0, r[11] if r[11] else "")
                if r[12]: self.cmb_sexo.set(r[12])
                self.ent_hijos.insert(0, r[13] if r[13] else "")
                
                if r[14]: self.cmb_movil.set(r[14])
                self.ent_salud_num.insert(0, r[15] if r[15] else "")
                self.ent_salud_venc.insert(0, r[16] if r[16] else "")
                self.ent_vida_num.insert(0, r[17] if r[17] else "")
                self.ent_vida_venc.insert(0, r[18] if r[18] else "")

                self.ent_sanidad_num.insert(0, r[23] if len(r) > 23 and r[23] else "")
                self.ent_sanidad_venc.insert(0, r[24] if len(r) > 24 and r[24] else "")
                self.ent_licencia2.insert(0, r[25] if len(r) > 25 and r[25] else "")
                self.ent_cat_licencia2.insert(0, r[26] if len(r) > 26 and r[26] else "")
                self.ent_venc_licencia2.insert(0, r[27] if len(r) > 27 and r[27] else "")
                self.ent_emergencia.insert(0, r[28] if len(r) > 28 and r[28] else "")
                self.ent_ini_contrato.insert(0, r[20] if len(r) > 20 and r[20] else "")
                self.ent_fin_contrato.insert(0, r[21] if len(r) > 21 and r[21] else "")
                self._observacion_estado = (r[22] or "") if len(r) > 22 else ""
                self._estado_anterior = self.cmb_estado.get() or "Activo"
                self._actualizar_lbl_observacion()
                
                json_str = r[19] if len(r) > 19 and r[19] else "{}"
                try:
                    self.rutas_documentos_db = json.loads(json_str)
                except Exception:
                    if json_str and not json_str.startswith("{"):
                        self.rutas_documentos_db = {"Expediente Clásico": json_str}
                    else:
                        self.rutas_documentos_db = {}

                self.actualizar_botones_docs()
                self._mostrar_foto_carnet()
                    
        except Exception as e:
            print("Error cargando edición:", e)
        finally:
            liberar_conexion(conn)

    def eliminar_chofer(self):
        sel = self.tabla.selection()
        if not sel: return messagebox.showwarning("Atención", "Seleccione un conductor para eliminar.")
        
        dni = self.tabla.item(sel[0], "values")[1]
        nombres = self.tabla.item(sel[0], "values")[2]
        vid = self.tabla.item(sel[0], "values")[0]
        
        if messagebox.askyesno("Confirmar", f"¿Eliminar permanentemente al personal {nombres}?"):
            conn = conectar_db()
            if not conn: return
            try:
                cursor = conn.cursor()
                cursor.execute("SELECT ruta_documentos FROM choferes WHERE id = %s", (vid,))
                res_archivo = cursor.fetchone()
                if res_archivo and res_archivo[0]:
                    try:
                        rutas = json.loads(res_archivo[0])
                        for ruta in rutas.values():
                            if os.path.exists(os.path.normpath(ruta)): os.remove(os.path.normpath(ruta))
                    except Exception:
                        if os.path.exists(os.path.normpath(res_archivo[0])): os.remove(os.path.normpath(res_archivo[0]))
                    
                identificador_crono = f"{nombres} (DNI: {dni})"
                cursor.execute("DELETE FROM choferes WHERE id = %s", (vid,))
                cursor.execute("DELETE FROM tareas_evento WHERE evento_asociado = 'FLOTA | Vencimientos' AND responsable = %s", (identificador_crono,))
                conn.commit()
                
                cache_sistema.invalidar()
                registrar_auditoria(self.usuario_activo, "Choferes", f"Eliminó al conductor {nombres}")
                
                self.limpiar_formulario()
                self.cargar_datos(reset_pagina=True)
            except Exception as e:
                messagebox.showerror("Error", str(e))
            finally:
                liberar_conexion(conn)