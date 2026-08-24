# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk
import calendar
import re
import urllib.request
import urllib.error
import ssl
import json
import os
import sys
import shutil
import subprocess
import threading
from datetime import datetime

# 🚀 IMPORTAMOS NUESTRAS NUEVAS HERRAMIENTAS CORPORATIVAS
from conexion import conectar_db, registrar_auditoria, liberar_conexion
from buffer_memoria import cache_sistema

try:
    import fitz  
    import base64
    import requests
except ImportError:
    pass


def _contexto_ssl_seguro():
    """Contexto SSL que funciona en Windows y macOS.

    En macOS, Python (instalado desde python.org o dentro de un .app de
    PyInstaller) no tiene el bundle de certificados raíz y urllib lanza:
        CERTIFICATE_VERIFY_FAILED: unable to get local issuer certificate
    Por eso usamos certifi (incluido con 'requests'), que trae su propio
    bundle de CAs. Si certifi no está disponible, se desactiva la
    verificación como último recurso.
    """
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        try:
            return ssl._create_unverified_context()
        except AttributeError:
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            return ctx


# =========================================================
# CLASE: MINI CALENDARIO (ACTUALIZADO CON COMBOBOX)
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
        fecha_seleccionada = f"{day:02d}/{self.current_month:02d}/{self.current_year}"
        self.target_entry.delete(0, tk.END)
        self.target_entry.insert(0, fecha_seleccionada)
        self.destroy()

_SCHEMA_VEHICULOS_OK = False

# =========================================================
# MÓDULO PRINCIPAL DE FLOTA
# =========================================================
class FlotaAutomotrizApp:
    def __init__(self, parent_frame, usuario_activo):
        self.parent_frame = parent_frame
        self.usuario_activo = usuario_activo
        self.id_edicion = None
        self.ruta_tarjeta_temp = ""
        self.ruta_tarjeta_db = ""
        
        # 🚀 VARIABLES DE PAGINACIÓN (LAZY LOADING)
        self.pagina_actual = 1
        self.registros_por_pagina = 50
        
        self.inicializar_bd()
        self.crear_interfaz()

    # 🚀 FIX: AUTO-CURACIÓN EN SEGUNDO PLANO
    def inicializar_bd(self):
        global _SCHEMA_VEHICULOS_OK
        if _SCHEMA_VEHICULOS_OK: return

        def tarea_curacion():
            global _SCHEMA_VEHICULOS_OK
            conn = conectar_db(silencioso=True)
            if not conn: return
            try:
                cursor = conn.cursor()
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS flota_vehiculos (
                        id SERIAL PRIMARY KEY,
                        placa VARCHAR(20) UNIQUE NOT NULL,
                        marca VARCHAR(100),
                        modelo VARCHAR(100),
                        anio VARCHAR(4),
                        color VARCHAR(50),
                        serial_motor VARCHAR(150),
                        serial_carroceria VARCHAR(150),
                        tipo_combustible VARCHAR(100),
                        vencimiento_soat VARCHAR(20),
                        vencimiento_rt VARCHAR(20),
                        estado VARCHAR(50) DEFAULT 'Operativo'
                    )
                """)
                conn.commit()
                
                columnas_nuevas = [
                    "ALTER TABLE flota_vehiculos ADD COLUMN kilometraje VARCHAR(50) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN emision_soat VARCHAR(20) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN emision_seguro VARCHAR(20) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN vencimiento_seguro VARCHAR(20) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN categoria VARCHAR(100) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN nro_titulo VARCHAR(100) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN fecha_titulo VARCHAR(50) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN ruta_tarjeta TEXT DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN fec_aceite VARCHAR(20) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN fec_correa VARCHAR(20) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN fec_rev_gas VARCHAR(20) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN fec_compra_bat VARCHAR(20) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN fec_venc_bat VARCHAR(20) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN extintor_num VARCHAR(100) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN fec_venc_extintor VARCHAR(20) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN km_prox_correa VARCHAR(50) DEFAULT ''" 
                ]
                
                for query in columnas_nuevas:
                    try:
                        cursor.execute(query)
                        conn.commit()
                    except Exception:
                        conn.rollback() 

                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS tareas_evento (
                        id SERIAL PRIMARY KEY, evento_asociado VARCHAR(255), nombre_tarea VARCHAR(255),
                        responsable VARCHAR(100), fecha_limite VARCHAR(20), estado VARCHAR(50),
                        notas TEXT, orden INTEGER DEFAULT 0, tipo_pago VARCHAR(50) DEFAULT 'Crédito', archivo_pago TEXT
                    )
                """)
                conn.commit()
                _SCHEMA_VEHICULOS_OK = True
                
            except Exception as e:
                print(f"Error BD Flota: {e}")
            finally:
                liberar_conexion(conn)
                
        threading.Thread(target=tarea_curacion, daemon=True).start()

    def extraer_datos_tarjeta_pdf(self):
        try:
            import fitz  
            import base64
            import requests
            import json
            import re
        except ImportError:
            return messagebox.showerror("Librería Faltante", "Para leer este documento con IA, abre tu consola y ejecuta:\n\npip install PyMuPDF requests")

        ruta_archivo = filedialog.askopenfilename(title="Seleccionar Tarjeta SUNARP", filetypes=[("Archivos", "*.pdf;*.png;*.jpg;*.jpeg")])
        if not ruta_archivo: return

        self.ruta_tarjeta_temp = ruta_archivo
        self.btn_ver_tarjeta.configure(state="normal", fg_color="#27ae60")

        messagebox.showinfo("Procesando", "La Inteligencia Artificial está analizando el documento.\nEsto tomará unos segundos; la ventana seguirá respondiendo...")

        # 🚀 RENDIMIENTO: el análisis de la IA (PyMuPDF + API) corre en un hilo para
        # no congelar la interfaz. Los campos se rellenan en el hilo principal con after(0).
        def procesar_en_hilo():
            try:
                if ruta_archivo.lower().endswith(".pdf"):
                    doc = fitz.open(ruta_archivo)
                    page = doc.load_page(0)
                    pix = page.get_pixmap(dpi=150)
                    img_bytes = pix.tobytes("jpeg")
                    mime_type = "image/jpeg"
                    doc.close()
                else:
                    with open(ruta_archivo, "rb") as f:
                        img_bytes = f.read()
                    mime_type = "image/jpeg" if ruta_archivo.lower().endswith(("jpg", "jpeg")) else "image/png"

                img_b64 = base64.b64encode(img_bytes).decode('utf-8')

                QWEN_API_KEY = "sk-ws-H.XXIIPI.p7Tl.MEUCIQDguE3Ocd7FjxHPFFi1_wroePYr_MVppA0wmOuUC9K8YgIgPisI2c7VCjgcuZ0Rv5U0yCwj3JIz_7omprW1jEoTqcg"
                QWEN_API_URL = "https://dashscope-intl.aliyuncs.com/compatible-mode/v1/chat/completions"

                prompt = """
                Eres un auditor experto en vehículos de Perú. Lee esta Tarjeta de Identificación Vehicular (TIVe) y extrae los datos en formato JSON estricto:
                - "placa": (Ej: ABC-123)
                - "titulo": (Ej: 2025-2927390)
                - "fecha_titulo": (Ej: 01/10/2025)
                - "categoria": (Ej: N1)
                - "marca": (Ej: DFSK)
                - "modelo": (Ej: C35)
                - "anio": (Año modelo o fabricación, Ej: 2026)
                - "color": (Ej: BLANCO)
                - "vin": (Número de VIN o Chasis)
                - "serie": (Número de serie)
                - "motor": (Número de motor)
                - "combustible": (Ej: BI-COMBUSTIBLE GNV)
                Si algún dato no existe, devuélvelo como "".
                """

                headers = {
                    "Authorization": f"Bearer {QWEN_API_KEY}",
                    "Content-Type": "application/json"
                }

                payload = {
                    "model": "qwen-vl-plus",
                    "messages": [
                        {
                            "role": "user",
                            "content": [
                                {
                                    "type": "image_url",
                                    "image_url": {"url": f"data:{mime_type};base64,{img_b64}"}
                                },
                                {"type": "text", "text": prompt}
                            ]
                        }
                    ]
                }

                respuesta = requests.post(QWEN_API_URL, headers=headers, json=payload)
                respuesta.raise_for_status()

                datos_respuesta = respuesta.json()
                texto_ia = datos_respuesta['choices'][0]['message']['content']

                match = re.search(r'\{.*\}', texto_ia, re.DOTALL)
                if not match: raise ValueError("No se encontró JSON válido en la respuesta de la IA.")

                return json.loads(match.group(0))
            except Exception as e:
                return e

        def aplicar_resultado(datos):
            if isinstance(datos, Exception):
                messagebox.showerror("Error IA", f"No se pudo procesar el documento con Inteligencia Artificial.\nDetalle: {datos}")
                return

            def set_val(entry, clave):
                val = datos.get(clave, "")
                if val and val != "-":
                    entry.delete(0, tk.END)
                    entry.insert(0, str(val).replace("-", "") if clave == "placa" else str(val))

            set_val(self.ent_placa, "placa")
            set_val(self.ent_titulo, "titulo")
            set_val(self.ent_fecha_titulo, "fecha_titulo")
            set_val(self.ent_categoria, "categoria")
            set_val(self.ent_marca, "marca")
            set_val(self.ent_modelo, "modelo")
            set_val(self.ent_anio, "anio")
            set_val(self.ent_color, "color")
            set_val(self.ent_motor, "motor")

            vin = datos.get("vin", "")
            serie = datos.get("serie", "")
            chasis_final = vin if vin else serie
            if chasis_final:
                self.ent_carroceria.delete(0, tk.END)
                self.ent_carroceria.insert(0, chasis_final)

            combustible = datos.get("combustible", "").upper()
            if combustible:
                es_dual = any(palabra in combustible for palabra in ["BI-COMBUSTIBLE", "BICOMBUSTIBLE", "DUAL", "GASOLINA"])
                if es_dual and "GNV" in combustible:
                    self.cmb_combustible.set("Dual (Gasolina + GNV)")
                elif es_dual and "GLP" in combustible:
                    self.cmb_combustible.set("Dual (Gasolina + GLP)")
                elif "DIESEL" in combustible:
                    self.cmb_combustible.set("Diésel")
                elif "GASOLINA" in combustible:
                    self.cmb_combustible.set("Gasolina")
                elif "GNV" in combustible:
                    self.cmb_combustible.set("GNV (Solo Gas)")
                elif "GLP" in combustible:
                    self.cmb_combustible.set("GLP (Solo Gas)")

            messagebox.showinfo("Lectura Exitosa", "La IA ha extraído los datos. El archivo está listo para ser guardado con el vehículo.")

        def correr():
            resultado = procesar_en_hilo()
            self.parent_frame.after(0, lambda: aplicar_resultado(resultado))

        threading.Thread(target=correr, daemon=True).start()

    def abrir_tarjeta(self):
        ruta = self.ruta_tarjeta_db if self.ruta_tarjeta_db else self.ruta_tarjeta_temp
        if ruta and os.path.exists(ruta):
            try:
                if sys.platform == "win32": os.startfile(ruta)
                elif sys.platform == "darwin": subprocess.call(["open", ruta])
                else: subprocess.call(["xdg-open", ruta])
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")
        else:
            messagebox.showwarning("No encontrado", "No se encontró el archivo físico de la tarjeta de propiedad.")

    def consultar_placa_api(self, event=None):
        placa = self.ent_placa.get().strip().replace("-", "").upper()
        if len(placa) < 6 or not placa.isalnum():
            return messagebox.showwarning("Placa Inválida", "Ingrese un número de placa válido (Ej: ABC123).")

        def tarea_api():
            try:
                # Contexto SSL seguro y compatible con macOS (ver _contexto_ssl_seguro)
                ctx = _contexto_ssl_seguro()

                url = f"https://api.apis.net.pe/v1/vehiculos?placa={placa}"
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                
                with urllib.request.urlopen(req, context=ctx, timeout=5) as response:
                    if response.status == 200:
                        data = json.loads(response.read().decode())
                        self.parent_frame.after(0, lambda: self._aplicar_datos_api(data))
                    else:
                        self.parent_frame.after(0, lambda: messagebox.showwarning("Sin Resultados", "No se encontró la placa en la base de datos pública."))
            except urllib.error.HTTPError as e:
                if e.code in [401, 403, 404]:
                    self.parent_frame.after(0, lambda: messagebox.showinfo("API Restringida", "La consulta automática requiere API Key. Usa el botón 'Leer Tarjeta (IA)'."))
                else:
                    self.parent_frame.after(0, lambda: messagebox.showwarning("Error de API", f"No se pudo contactar al servidor ({e.code})."))
            except Exception:
                self.parent_frame.after(0, lambda: messagebox.showinfo("Servicio no disponible", "La conexión a la base de datos vehicular no está disponible. Usa 'Leer Tarjeta (IA)'."))

        threading.Thread(target=tarea_api, daemon=True).start()

    def _aplicar_datos_api(self, data):
        if "marca" in data: self.ent_marca.delete(0, tk.END); self.ent_marca.insert(0, data.get("marca", ""))
        if "modelo" in data: self.ent_modelo.delete(0, tk.END); self.ent_modelo.insert(0, data.get("modelo", ""))
        if "color" in data: self.ent_color.delete(0, tk.END); self.ent_color.insert(0, data.get("color", ""))
        messagebox.showinfo("Consulta", "Datos obtenidos exitosamente de la base de datos.")

    def crear_interfaz(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="#ffffff", foreground="#000000", rowheight=28, font=("Arial", 10))
        style.map("Treeview", background=[("selected", "#1f538d")], foreground=[("selected", "#ffffff")])
        style.configure("Treeview.Heading", background="#f0f0f0", font=("Arial", 10, "bold"))

        lbl_titulo = ctk.CTkLabel(self.parent_frame, text="🚙 GESTIÓN DE FLOTA AUTOMOTRIZ", font=("Arial", 18, "bold"), text_color="#1f538d")
        lbl_titulo.pack(anchor="w", padx=20, pady=(15, 5))

        self.main_split = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.main_split.pack(fill="both", expand=True, padx=15, pady=5)

        # PANEL IZQUIERDO: FORMULARIO (Scrollable)
        self.f_form = ctk.CTkScrollableFrame(self.main_split, width=330, corner_radius=10, fg_color="#f8f9fa", border_width=1, border_color="#e0e0e0")
        self.f_form.pack(side="left", fill="y", padx=(0, 10))

        # NUEVO BLOQUE DE BOTONES PARA DOCUMENTOS
        f_lector = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_lector.pack(fill="x", padx=10, pady=(5, 10))
        
        btn_pdf = ctk.CTkButton(f_lector, text="📄 Leer Tarjeta (IA)", font=("Arial", 12, "bold"), fg_color="#8e44ad", hover_color="#732d91", command=self.extraer_datos_tarjeta_pdf)
        btn_pdf.pack(side="left", expand=True, fill="x", padx=(0, 5))
        
        self.btn_ver_tarjeta = ctk.CTkButton(f_lector, text="👁️ Ver Tarjeta", font=("Arial", 12, "bold"), fg_color="#34495e", command=self.abrir_tarjeta, state="disabled")
        self.btn_ver_tarjeta.pack(side="right", expand=True, fill="x", padx=(5, 0))

        ctk.CTkLabel(self.f_form, text="Datos del Vehículo", font=("Arial", 14, "bold")).pack(pady=(5, 15))

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

        # 1. Datos Principales
        ctk.CTkLabel(self.f_form, text="--- Datos Principales ---", font=("Arial", 11, "bold"), text_color="#1f538d").pack(anchor="w", padx=10, pady=(5,5))
        self.ent_placa = crear_campo("Matrícula / Placa: *", "Ej: ABC-123 (Presiona Enter para buscar)")
        self.ent_placa.bind("<Return>", self.consultar_placa_api)
        self.ent_titulo = crear_campo("N° de Título:", "Ej: 2024-12345678")
        self.ent_fecha_titulo = crear_campo_fecha("Fecha del Título:")
        self.ent_categoria = crear_campo("Categoría:", "Ej: M1 / N1")
        self.ent_marca = crear_campo("Marca:", "Ej: Toyota")
        self.ent_modelo = crear_campo("Modelo:", "Ej: Hilux")
        self.ent_anio = crear_campo("Año de Fabricación:", "Ej: 2024")
        self.ent_color = crear_campo("Color:", "Ej: Blanco")
        
        # 2. Identificadores Físicos
        ctk.CTkLabel(self.f_form, text="--- Identificadores Físicos ---", font=("Arial", 11, "bold"), text_color="#1f538d").pack(anchor="w", padx=10, pady=(10,5))
        self.ent_kilometraje = crear_campo("Kilometraje Actual:", "Ej: 45000")
        self.ent_motor = crear_campo("Serial del Motor:", "N° de Motor")
        self.ent_carroceria = crear_campo("Serial Carrocería (VIN):", "N° de Chasis/VIN")

        ctk.CTkLabel(self.f_form, text="Tipo de Combustible:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        opciones_combustible = ["Gasolina", "Diésel", "GNV (Solo Gas)", "GLP (Solo Gas)", "Dual (Gasolina + GNV)", "Dual (Gasolina + GLP)", "Híbrido", "Eléctrico"]
        self.cmb_combustible = ctk.CTkComboBox(self.f_form, values=opciones_combustible, state="readonly")
        self.cmb_combustible.pack(fill="x", padx=10, pady=(0, 10))
        self.cmb_combustible.set("Gasolina")

        # 3. Mantenimientos (Nuevos campos)
        ctk.CTkLabel(self.f_form, text="--- Fechas de Mantenimiento ---", font=("Arial", 11, "bold"), text_color="#e67e22").pack(anchor="w", padx=10, pady=(10,5))
        self.ent_fec_aceite = crear_campo_fecha("Último Cambio de Aceite:")
        self.ent_fec_correa = crear_campo_fecha("Cambio Correa/Cadena Distribución:")
        self.ent_km_prox_correa = crear_campo("KM Próx. Cambio Correa:", "Ej: 100000")
        self.ent_fec_compra_bat = crear_campo_fecha("Compra de Batería:")
        self.ent_fec_venc_bat = crear_campo_fecha("Vencimiento Garantía Batería:")

        # 4. Seguridad (Extintor)
        ctk.CTkLabel(self.f_form, text="--- Seguridad y Extintores ---", font=("Arial", 11, "bold"), text_color="#c0392b").pack(anchor="w", padx=10, pady=(10,5))
        self.ent_extintor_num = crear_campo("N° Certificado Extintor:")
        self.ent_fec_venc_ext = crear_campo_fecha("Vencimiento del Extintor:")

        # 5. SOAT y Seguro
        ctk.CTkLabel(self.f_form, text="--- SOAT y Póliza de Seguro ---", font=("Arial", 11, "bold"), text_color="#27ae60").pack(anchor="w", padx=10, pady=(10,5))
        self.ent_emi_soat = crear_campo_fecha("Emisión SOAT:")
        self.ent_soat = crear_campo_fecha("Vencimiento SOAT:")
        self.ent_emi_seguro = crear_campo_fecha("Emisión del Seguro:")
        self.ent_venc_seguro = crear_campo_fecha("Vencimiento del Seguro:")
        
        # 6. Revisiones y Estado
        ctk.CTkLabel(self.f_form, text="--- Revisiones y Estado ---", font=("Arial", 11, "bold"), text_color="#8e44ad").pack(anchor="w", padx=10, pady=(10,5))
        self.ent_rt = crear_campo_fecha("Venc. Revisión Técnica:")
        self.ent_fec_rev_gas = crear_campo_fecha("Venc. Rev. Sist. Gas:")

        ctk.CTkLabel(self.f_form, text="Estado Operativo:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.cmb_estado = ctk.CTkComboBox(self.f_form, values=["Operativo", "En Mantenimiento", "Fuera de Servicio", "Vendido"], state="readonly")
        self.cmb_estado.pack(fill="x", padx=10, pady=(0, 15))
        self.cmb_estado.set("Operativo")

        f_btns = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_btns.pack(fill="x", padx=10, pady=(10, 20))
        
        self.btn_guardar = ctk.CTkButton(f_btns, text="💾 Guardar", fg_color="#27ae60", hover_color="#1e8449", font=("Arial", 12, "bold"), command=self.guardar_vehiculo)
        self.btn_guardar.pack(side="left", expand=True, fill="x", padx=(0, 5))
        
        btn_limpiar = ctk.CTkButton(f_btns, text="🔄 Limpiar", fg_color="#7f8c8d", hover_color="#606b6b", font=("Arial", 12, "bold"), command=self.limpiar_formulario)
        btn_limpiar.pack(side="right", expand=True, fill="x", padx=(5, 0))

        # PANEL DERECHO: TABLA Y BUSCADOR
        f_derecho = ctk.CTkFrame(self.main_split, fg_color="transparent")
        f_derecho.pack(side="right", fill="both", expand=True)

        f_busqueda = ctk.CTkFrame(f_derecho, fg_color="transparent")
        f_busqueda.pack(fill="x", pady=(0, 5))
        ctk.CTkLabel(f_busqueda, text="🔍 Buscar:", font=("Arial", 12, "bold")).pack(side="left", padx=(0, 5))
        self.ent_buscar = ctk.CTkEntry(f_busqueda, placeholder_text="Buscar por placa, marca, modelo...")
        self.ent_buscar.pack(side="left", fill="x", expand=True)
        
        self.ent_buscar.bind("<KeyRelease>", lambda e: self.buscar_con_retraso())
        self.ent_buscar.bind("<Return>", lambda e: self.cargar_datos(reset_pagina=True))

        f_tabla = ctk.CTkFrame(f_derecho, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True)

        columnas = ("id", "placa", "vehiculo", "categoria", "color", "combustible", "kilometraje", "estado", "titulo", "fecha_titulo")
        self.tabla = ttk.Treeview(f_tabla, columns=columnas, show="headings")
        self.tabla.heading("id", text="ID")
        self.tabla.heading("placa", text="Placa")
        self.tabla.heading("vehiculo", text="Marca / Modelo (Año)")
        self.tabla.heading("categoria", text="Categoría")
        self.tabla.heading("color", text="Color")
        self.tabla.heading("combustible", text="Combustible")
        self.tabla.heading("kilometraje", text="Kilometraje")
        self.tabla.heading("estado", text="Estado")
        self.tabla.heading("titulo", text="N° Título")
        self.tabla.heading("fecha_titulo", text="Fecha Título")

        self.tabla.column("id", width=0, stretch=tk.NO)
        self.tabla.column("placa", width=80, anchor="center")
        self.tabla.column("vehiculo", width=180, anchor="w")
        self.tabla.column("categoria", width=80, anchor="center")
        self.tabla.column("color", width=80, anchor="center")
        self.tabla.column("combustible", width=130, anchor="center")
        self.tabla.column("kilometraje", width=90, anchor="center")
        self.tabla.column("estado", width=100, anchor="center")
        self.tabla.column("titulo", width=0, stretch=tk.NO)
        self.tabla.column("fecha_titulo", width=0, stretch=tk.NO)
        
        self.tabla.config(displaycolumns=("placa", "vehiculo", "categoria", "color", "combustible", "kilometraje", "estado"))

        scroll_y = ctk.CTkScrollbar(f_tabla, orientation="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scroll_y.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")
        
        self.tabla.bind("<Double-1>", lambda e: self.cargar_para_edicion())

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
        ctk.CTkButton(f_acciones_tabla, text="❌ Eliminar", fg_color="#e74c3c", hover_color="#c0392b", font=("Arial", 12, "bold"), command=self.eliminar_vehiculo).pack(side="right", padx=5)

        self.parent_frame.after(100, lambda: self.cargar_datos(reset_pagina=True))

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
        self.ruta_tarjeta_temp = ""
        self.ruta_tarjeta_db = ""
        self.btn_guardar.configure(text="💾 Guardar Nuevo")
        self.btn_ver_tarjeta.configure(state="disabled", fg_color="#34495e")
        self.ent_placa.delete(0, tk.END)
        self.ent_titulo.delete(0, tk.END)
        self.ent_fecha_titulo.delete(0, tk.END)
        self.ent_categoria.delete(0, tk.END)
        self.ent_marca.delete(0, tk.END)
        self.ent_modelo.delete(0, tk.END)
        self.ent_anio.delete(0, tk.END)
        self.ent_color.delete(0, tk.END)
        self.ent_kilometraje.delete(0, tk.END)
        self.ent_motor.delete(0, tk.END)
        self.ent_carroceria.delete(0, tk.END)
        self.cmb_combustible.set("Gasolina")
        self.ent_fec_aceite.delete(0, tk.END)
        self.ent_fec_correa.delete(0, tk.END)
        self.ent_km_prox_correa.delete(0, tk.END)
        self.ent_fec_rev_gas.delete(0, tk.END)
        self.ent_fec_compra_bat.delete(0, tk.END)
        self.ent_fec_venc_bat.delete(0, tk.END)
        self.ent_extintor_num.delete(0, tk.END)
        self.ent_fec_venc_ext.delete(0, tk.END)
        self.ent_emi_soat.delete(0, tk.END)
        self.ent_soat.delete(0, tk.END)
        self.ent_emi_seguro.delete(0, tk.END)
        self.ent_venc_seguro.delete(0, tk.END)
        self.ent_rt.delete(0, tk.END)
        self.cmb_estado.set("Operativo")

    # 🚀 FIX: LAZY LOADING Y CACHÉ
    def cargar_datos(self, reset_pagina=False):
        if reset_pagina:
            self.pagina_actual = 1
            
        self.lbl_pagina.configure(text=f"Pág {self.pagina_actual}")

        for item in self.tabla.get_children(): 
            self.tabla.delete(item)
            
        filtro = self.ent_buscar.get().strip().lower()
        offset = (self.pagina_actual - 1) * self.registros_por_pagina
        
        clave_cache = f"vehiculos_flota_{filtro}_pag_{self.pagina_actual}"
        datos = cache_sistema.obtener(clave_cache)
        
        if datos is not None:
            self._pintar_vehiculos(datos)
        else:
            self.tabla.insert("", tk.END, values=("", "", "Cargando datos...", "", "", "", "", "", "", ""))
            
            def tarea_descarga():
                conn = conectar_db(silencioso=True)
                if not conn: return
                try:
                    cursor = conn.cursor()
                    if filtro:
                        cursor.execute("""
                            SELECT id, placa, marca, modelo, anio, color, tipo_combustible, kilometraje, estado, categoria, nro_titulo, fecha_titulo
                            FROM flota_vehiculos 
                            WHERE placa ILIKE %s OR marca ILIKE %s OR modelo ILIKE %s 
                            ORDER BY id DESC LIMIT %s OFFSET %s
                        """, (f"%{filtro}%", f"%{filtro}%", f"%{filtro}%", self.registros_por_pagina, offset))
                    else:
                        cursor.execute("""
                            SELECT id, placa, marca, modelo, anio, color, tipo_combustible, kilometraje, estado, categoria, nro_titulo, fecha_titulo 
                            FROM flota_vehiculos 
                            ORDER BY id DESC LIMIT %s OFFSET %s
                        """, (self.registros_por_pagina, offset))
                    
                    datos_db = cursor.fetchall()
                    cache_sistema.guardar(clave_cache, datos_db)
                    self.parent_frame.after(0, lambda: self._pintar_vehiculos(datos_db))
                except Exception as e:
                    print(f"Error cargando tabla de flota: {e}")
                finally:
                    liberar_conexion(conn)

            threading.Thread(target=tarea_descarga, daemon=True).start()

    def _pintar_vehiculos(self, datos):
        for item in self.tabla.get_children():
            self.tabla.delete(item)

        for r in datos:
            v_id, placa, marca, modelo, anio, color, comb, km, est, categoria, nro_titulo, fecha_titulo = r
            vehiculo_nom = f"{marca} {modelo} ({anio})"
            km_mostrar = km if km else "0"
            cat_mostrar = categoria if categoria else "-"
            
            self.tabla.insert("", tk.END, values=(v_id, placa, vehiculo_nom, cat_mostrar, color, comb, km_mostrar, est, nro_titulo, fecha_titulo))
            
        if self.pagina_actual > 1:
            self.btn_ant.configure(state="normal")
        else:
            self.btn_ant.configure(state="disabled")
            
        if len(datos) == self.registros_por_pagina:
            self.btn_sig.configure(state="normal")
        else:
            self.btn_sig.configure(state="disabled")

    def guardar_vehiculo(self):
        placa = self.ent_placa.get().strip().upper()
        if not placa:
            return messagebox.showwarning("Atención", "La placa es obligatoria.")
            
        titulo = self.ent_titulo.get().strip()
        fecha_titulo = self.ent_fecha_titulo.get().strip()
        categoria = self.ent_categoria.get().strip()
        marca = self.ent_marca.get().strip()
        modelo = self.ent_modelo.get().strip()
        anio = self.ent_anio.get().strip()
        color = self.ent_color.get().strip()
        kilometraje = self.ent_kilometraje.get().strip()
        motor = self.ent_motor.get().strip()
        carroceria = self.ent_carroceria.get().strip()
        combustible = self.cmb_combustible.get()
        
        fec_aceite = self.ent_fec_aceite.get().strip()
        fec_correa = self.ent_fec_correa.get().strip()
        km_prox_correa = self.ent_km_prox_correa.get().strip()
        fec_gas = self.ent_fec_rev_gas.get().strip()
        fec_compra_bat = self.ent_fec_compra_bat.get().strip()
        fec_venc_bat = self.ent_fec_venc_bat.get().strip()
        num_extintor = self.ent_extintor_num.get().strip()
        fec_venc_ext = self.ent_fec_venc_ext.get().strip()
        
        emi_soat = self.ent_emi_soat.get().strip()
        venc_soat = self.ent_soat.get().strip()
        emi_seguro = self.ent_emi_seguro.get().strip()
        venc_seguro = self.ent_venc_seguro.get().strip()
        rt = self.ent_rt.get().strip()
        estado = self.cmb_estado.get()

        ruta_final_archivo = self.ruta_tarjeta_db
        if self.ruta_tarjeta_temp and os.path.exists(self.ruta_tarjeta_temp):
            try:
                base_dir = os.path.dirname(os.path.abspath(__file__))
                carpeta_destino = os.path.join(base_dir, "archivos_flota", "tarjetas_propiedad")
                os.makedirs(carpeta_destino, exist_ok=True)
                
                ext = os.path.splitext(self.ruta_tarjeta_temp)[1]
                nombre_archivo = f"Tarjeta_{placa}_{datetime.now().strftime('%Y%m%d%H%M%S')}{ext}"
                ruta_final_archivo = os.path.join(carpeta_destino, nombre_archivo)
                
                shutil.copy2(self.ruta_tarjeta_temp, ruta_final_archivo)
            except Exception as e:
                return messagebox.showerror("Error de Archivo", f"No se pudo guardar la tarjeta física:\n{e}")
                
        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            if self.id_edicion:
                cursor.execute("""
                    UPDATE flota_vehiculos SET 
                    placa=%s, marca=%s, modelo=%s, anio=%s, color=%s, kilometraje=%s, serial_motor=%s, 
                    serial_carroceria=%s, tipo_combustible=%s, emision_soat=%s, vencimiento_soat=%s, 
                    emision_seguro=%s, vencimiento_seguro=%s, vencimiento_rt=%s, estado=%s,
                    categoria=%s, nro_titulo=%s, fecha_titulo=%s, ruta_tarjeta=%s,
                    fec_aceite=%s, fec_correa=%s, km_prox_correa=%s, fec_rev_gas=%s, fec_compra_bat=%s, fec_venc_bat=%s, 
                    extintor_num=%s, fec_venc_extintor=%s
                    WHERE id=%s
                """, (placa, marca, modelo, anio, color, kilometraje, motor, carroceria, combustible, 
                      emi_soat, venc_soat, emi_seguro, venc_seguro, rt, estado, categoria, titulo, 
                      fecha_titulo, ruta_final_archivo, fec_aceite, fec_correa, km_prox_correa, fec_gas, fec_compra_bat, 
                      fec_venc_bat, num_extintor, fec_venc_ext, self.id_edicion))
                registrar_auditoria(self.usuario_activo, "Flota", f"Actualizó datos del vehículo placa {placa}")
                messagebox.showinfo("Éxito", "Vehículo actualizado correctamente.")
            else:
                cursor.execute("SELECT id FROM flota_vehiculos WHERE placa = %s", (placa,))
                if cursor.fetchone():
                    liberar_conexion(conn)
                    return messagebox.showwarning("Duplicado", f"La placa {placa} ya existe en el sistema.")
                    
                cursor.execute("""
                    INSERT INTO flota_vehiculos (placa, marca, modelo, anio, color, kilometraje, serial_motor, serial_carroceria, 
                    tipo_combustible, emision_soat, vencimiento_soat, emision_seguro, vencimiento_seguro, vencimiento_rt, estado, 
                    categoria, nro_titulo, fecha_titulo, ruta_tarjeta, fec_aceite, fec_correa, km_prox_correa, fec_rev_gas, fec_compra_bat, 
                    fec_venc_bat, extintor_num, fec_venc_extintor) 
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (placa, marca, modelo, anio, color, kilometraje, motor, carroceria, combustible, 
                      emi_soat, venc_soat, emi_seguro, venc_seguro, rt, estado, categoria, titulo, 
                      fecha_titulo, ruta_final_archivo, fec_aceite, fec_correa, km_prox_correa, fec_gas, fec_compra_bat, 
                      fec_venc_bat, num_extintor, fec_venc_ext))
                registrar_auditoria(self.usuario_activo, "Flota", f"Registró nuevo vehículo placa {placa}")
                messagebox.showinfo("Éxito", "Vehículo registrado correctamente.")
            
            try:
                c_crono = conn.cursor()
                c_crono.execute("DELETE FROM tareas_evento WHERE evento_asociado = 'FLOTA | Vencimientos' AND responsable = %s", (placa,))
                
                vencimientos = [
                    ("Vencimiento SOAT", venc_soat),
                    ("Vencimiento Seguro", venc_seguro),
                    ("Vencimiento Revisión Técnica", rt),
                    ("Venc. Rev. Sist. Gas", fec_gas),
                    ("Vencimiento Garantía Batería", fec_venc_bat),
                    ("Vencimiento Extintor", fec_venc_ext)
                ]
                for nom_doc, fec_doc in vencimientos:
                    if fec_doc and fec_doc.strip():
                        c_crono.execute("SELECT COALESCE(MAX(orden), 0) FROM tareas_evento WHERE evento_asociado = 'FLOTA | Vencimientos'")
                        nuevo_orden = c_crono.fetchone()[0] + 1
                        
                        c_crono.execute("""
                            INSERT INTO tareas_evento (evento_asociado, nombre_tarea, responsable, fecha_limite, estado, notas, orden, tipo_pago)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                        """, ("FLOTA | Vencimientos", nom_doc, placa, fec_doc, "Pendiente", f"Alerta automática de {nom_doc} para la unidad {placa}", nuevo_orden, "No aplica"))
            except Exception as e_crono: print("Aviso - Sincronización Cronograma:", e_crono)

            conn.commit()
            cache_sistema.invalidar() # 🚀 FIX: Borrar caché
            cache_sistema.invalidar("lista_vehiculos_combobox") # Refrescar lista desplegable de otros módulos
            
            self.limpiar_formulario()
            self.cargar_datos(reset_pagina=True)
        except Exception as e:
            conn.rollback()
            messagebox.showerror("Error", str(e))
        finally:
            liberar_conexion(conn)

    def cargar_para_edicion(self):
        sel = self.tabla.selection()
        if not sel: 
            return messagebox.showwarning("Atención", "Por favor, seleccione un vehículo de la tabla primero.")
            
        vid = self.tabla.item(sel[0], "values")[0]
        
        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id, placa, marca, modelo, anio, color, serial_motor, serial_carroceria, 
                tipo_combustible, kilometraje, emision_soat, vencimiento_soat, 
                emision_seguro, vencimiento_seguro, vencimiento_rt, estado,
                categoria, nro_titulo, fecha_titulo, ruta_tarjeta,
                fec_aceite, fec_correa, fec_rev_gas, fec_compra_bat, fec_venc_bat, extintor_num, fec_venc_extintor,
                km_prox_correa
                FROM flota_vehiculos WHERE id = %s
            """, (vid,))
            r = cursor.fetchone()
            if r:
                self.limpiar_formulario()
                self.id_edicion = r[0]
                self.btn_guardar.configure(text="💾 Actualizar Vehículo")
                
                self.ent_placa.insert(0, r[1])
                self.ent_marca.insert(0, r[2] if r[2] else "")
                self.ent_modelo.insert(0, r[3] if r[3] else "")
                self.ent_anio.insert(0, r[4] if r[4] else "")
                self.ent_color.insert(0, r[5] if r[5] else "")
                self.ent_motor.insert(0, r[6] if r[6] else "")
                self.ent_carroceria.insert(0, r[7] if r[7] else "")
                if r[8]: self.cmb_combustible.set(r[8])
                self.ent_kilometraje.insert(0, r[9] if r[9] else "")
                self.ent_emi_soat.insert(0, r[10] if r[10] else "")
                self.ent_soat.insert(0, r[11] if r[11] else "")
                self.ent_emi_seguro.insert(0, r[12] if r[12] else "")
                self.ent_venc_seguro.insert(0, r[13] if r[13] else "")
                self.ent_rt.insert(0, r[14] if r[14] else "")
                if r[15]: self.cmb_estado.set(r[15])
                
                self.ent_categoria.insert(0, r[16] if r[16] else "")
                self.ent_titulo.insert(0, r[17] if r[17] else "")
                self.ent_fecha_titulo.insert(0, r[18] if r[18] else "")
                
                self.ruta_tarjeta_db = r[19] if r[19] else ""
                
                self.ent_fec_aceite.insert(0, r[20] if r[20] else "")
                self.ent_fec_correa.insert(0, r[21] if r[21] else "")
                self.ent_fec_rev_gas.insert(0, r[22] if r[22] else "")
                self.ent_fec_compra_bat.insert(0, r[23] if r[23] else "")
                self.ent_fec_venc_bat.insert(0, r[24] if r[24] else "")
                self.ent_extintor_num.insert(0, r[25] if r[25] else "")
                self.ent_fec_venc_ext.insert(0, r[26] if r[26] else "")
                self.ent_km_prox_correa.insert(0, r[27] if r[27] else "")

                if self.ruta_tarjeta_db and os.path.exists(self.ruta_tarjeta_db):
                    self.btn_ver_tarjeta.configure(state="normal", fg_color="#27ae60")
                else:
                    self.btn_ver_tarjeta.configure(state="disabled", fg_color="#34495e")
                    
        except Exception as e:
            messagebox.showerror("Error Base de Datos", f"No se pudo cargar el registro.\nDetalle: {e}")
        finally:
            liberar_conexion(conn)

    def eliminar_vehiculo(self):
        sel = self.tabla.selection()
        if not sel: return messagebox.showwarning("Atención", "Seleccione un vehículo para eliminar.")
        
        placa = self.tabla.item(sel[0], "values")[1]
        vid = self.tabla.item(sel[0], "values")[0]
        
        if messagebox.askyesno("Confirmar", f"¿Eliminar permanentemente el vehículo con placa {placa}?"):
            conn = conectar_db()
            if not conn: return
            try:
                cursor = conn.cursor()
                
                cursor.execute("SELECT ruta_tarjeta FROM flota_vehiculos WHERE id = %s", (vid,))
                res_archivo = cursor.fetchone()
                if res_archivo and res_archivo[0] and os.path.exists(res_archivo[0]):
                    os.remove(res_archivo[0])
                    
                cursor.execute("DELETE FROM flota_vehiculos WHERE id = %s", (vid,))
                cursor.execute("DELETE FROM tareas_evento WHERE evento_asociado = 'FLOTA | Vencimientos' AND responsable = %s", (placa,))
                conn.commit()
                
                cache_sistema.invalidar()
                cache_sistema.invalidar("lista_vehiculos_combobox")
                
                registrar_auditoria(self.usuario_activo, "Flota", f"Eliminó el vehículo placa {placa}")
                self.limpiar_formulario()
                self.cargar_datos(reset_pagina=True)
                
                messagebox.showinfo("Éxito", "El vehículo ha sido eliminado correctamente del sistema.")
                
            except Exception as e:
                messagebox.showerror("Error", f"Fallo al intentar eliminar el registro:\n{e}")
            finally:
                liberar_conexion(conn)