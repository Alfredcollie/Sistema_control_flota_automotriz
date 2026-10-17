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
from dialogos_seguros import seleccionar_archivo_dialogo, guardar_archivo_dialogo
from tareas_seguras import ejecutar_en_hilo
from app_paths import CONFIG_FILE


# =========================================================
# 🚀 CONFIGURACIÓN GENERAL (límite de km para el cambio de aceite)
# =========================================================
def cargar_configuracion_regional():
    config = {"alerta_aceite_km": "5000", "alerta_cambio_aceite_km": "5000"}
    try:
        if os.path.exists(str(CONFIG_FILE)):
            with open(str(CONFIG_FILE), "r", encoding="utf-8") as f:
                config.update(json.load(f))
    except Exception:
        pass
    return config


CONFIG_REGIONAL = cargar_configuracion_regional()


def limite_km_aceite():
    """Cada cuántos kilómetros toca el cambio de aceite (Configuración General)."""
    for clave in ("alerta_aceite_km", "alerta_cambio_aceite_km"):
        try:
            v = str(CONFIG_REGIONAL.get(clave, "") or "").replace(",", ".").strip()
            if v:
                n = float(v)
                if n > 0:
                    return n
        except Exception:
            continue
    return 5000.0


def limite_km_gas():
    """Cada cuántos kilómetros toca el mantenimiento del sistema de gas.

    Se configura en Configuración General → Ajustes de Flota Automotriz
    ('alerta_gas_km'); por defecto 10.000 km.
    """
    for clave in ("alerta_gas_km", "alerta_mantenimiento_gas_km"):
        try:
            v = str(CONFIG_REGIONAL.get(clave, "") or "").replace(",", ".").strip()
            if v:
                n = float(v)
                if n > 0:
                    return n
        except Exception:
            continue
    return 10000.0


def a_numero(valor):
    """Convierte '45000', '45,000', '45.000 km' o '' a número (0.0 si no se puede)."""
    txt = re.sub(r"[^0-9.,]", "", str(valor or ""))
    if not txt:
        return 0.0
    try:
        if "," in txt and "." in txt:
            txt = (txt.replace(".", "").replace(",", ".") if txt.rfind(",") > txt.rfind(".")
                   else txt.replace(",", ""))
        elif "," in txt:
            txt = txt.replace(",", ".") if len(txt.split(",")[-1]) <= 2 else txt.replace(",", "")
        elif "." in txt:
            txt = txt if len(txt.split(".")[-1]) <= 2 else txt.replace(".", "")
        return float(txt)
    except Exception:
        return 0.0

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
def fechas_de_certificado(texto):
    """(emision, vencimiento) leídas del texto de un certificado.

    Sirve para el SOAT y para el certificado de extintor: se buscan las fechas
    dd/mm/aaaa y se toma la menor como emisión y la mayor como vencimiento.
    Devuelve ("", "") si no encuentra dos fechas coherentes.
    """
    import re as _re
    from datetime import datetime as _dt
    encontradas = []
    for m in _re.finditer(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})", str(texto or "")):
        d, mes, anio = m.groups()
        try:
            anio = int(anio) + (2000 if len(anio) == 2 else 0)
            fecha = _dt(anio, int(mes), int(d))
        except Exception:
            continue
        if 2000 <= fecha.year <= 2100:
            encontradas.append(fecha)
    if not encontradas:
        return "", ""
    unicas = sorted(set(encontradas))
    if len(unicas) == 1:
        return unicas[0].strftime("%d/%m/%Y"), ""
    # La primera fecha del certificado suele ser la emisión y la última el vencimiento
    return unicas[0].strftime("%d/%m/%Y"), unicas[-1].strftime("%d/%m/%Y")


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
                    "ALTER TABLE flota_vehiculos ADD COLUMN km_prox_correa VARCHAR(50) DEFAULT ''",
                    # 🛢️ Kilometraje del último cambio de aceite (para calcular el próximo).
                    # Son las mismas columnas que usa el Cronograma para sus proyecciones.
                    "ALTER TABLE flota_vehiculos ADD COLUMN km_ultimo_aceite NUMERIC DEFAULT 0",
                    "ALTER TABLE flota_vehiculos ADD COLUMN fecha_ultimo_aceite VARCHAR(20) DEFAULT ''",
                    # 📎 Certificados escaneados (SOAT y extintor) guardados como PDF/JPG
                    "ALTER TABLE flota_vehiculos ADD COLUMN soat_archivo TEXT DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN extintor_archivo TEXT DEFAULT ''",
                    # 🛡️ N° de póliza del seguro y soportes de RT y de gas
                    "ALTER TABLE flota_vehiculos ADD COLUMN seguro_num VARCHAR(100) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN rt_archivo TEXT DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN gas_archivo TEXT DEFAULT ''",
                    # 💨 MANTENIMIENTO del sistema de gas (distinto de la REVISIÓN):
                    #    fecha de emisión + kilómetros del mantenimiento
                    "ALTER TABLE flota_vehiculos ADD COLUMN emision_mant_gas VARCHAR(20) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN km_mant_gas VARCHAR(50) DEFAULT ''",
                    # 🔋 Marca y modelo de la batería
                    "ALTER TABLE flota_vehiculos ADD COLUMN bat_marca VARCHAR(60) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN bat_modelo VARCHAR(60) DEFAULT ''",
                    # 🛞 Neumáticos: marca, modelo, medidas, kilómetros y fecha
                    "ALTER TABLE flota_vehiculos ADD COLUMN neu_marca VARCHAR(60) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN neu_modelo VARCHAR(60) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN neu_medidas VARCHAR(60) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN neu_km VARCHAR(50) DEFAULT ''",
                    "ALTER TABLE flota_vehiculos ADD COLUMN neu_fecha VARCHAR(20) DEFAULT ''",
                    # (renombre por si ya se había creado con el nombre anterior)
                    "ALTER TABLE flota_vehiculos RENAME COLUMN emision_rev_gas TO emision_mant_gas",
                    "ALTER TABLE flota_vehiculos RENAME COLUMN km_rev_gas TO km_mant_gas"
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

        ruta_archivo = seleccionar_archivo_dialogo(titulo="Seleccionar Tarjeta SUNARP", tipos=[("Archivos", "*.pdf;*.png;*.jpg;*.jpeg")])
        if not ruta_archivo: return

        self.ruta_tarjeta_temp = ruta_archivo
        self.btn_ver_tarjeta.configure(state="normal", fg_color="#27ae60")

        messagebox.showinfo("Procesando", "La Inteligencia Artificial está analizando el documento.\nEsto tomará unos segundos; la ventana seguirá respondiendo...")

        # 🚀 RENDIMIENTO: el análisis de la IA (PyMuPDF + API) corre en un hilo para
        # no congelar la interfaz. Los campos se rellenan en el hilo principal
        # mediante 'al_terminar' de 'tareas_seguras.ejecutar_en_hilo'.
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

                # 🔑 La clave se puede cambiar en Configuración General → Ajustes (ia_qwen_api_key)
                QWEN_API_KEY = str(CONFIG_REGIONAL.get("ia_qwen_api_key", "") or "").strip() or "sk-ws-H.XXIIPI.p7Tl.MEUCIQDguE3Ocd7FjxHPFFi1_wroePYr_MVppA0wmOuUC9K8YgIgPisI2c7VCjgcuZ0Rv5U0yCwj3JIz_7omprW1jEoTqcg"
                QWEN_API_URL = (str(CONFIG_REGIONAL.get("ia_qwen_url", "") or "").strip()
                                or "https://dashscope-intl.aliyuncs.com/compatible-mode/v1/chat/completions")
                # 🤖 Modelos a intentar en orden: qwen-vl-plus fue RETIRADO (da 404 "Model not exist")
                _preferido = str(CONFIG_REGIONAL.get("ia_qwen_modelo", "") or "").strip()
                MODELOS_IA = ([_preferido] if _preferido else []) + [
                    "qwen3-vl-plus", "qwen-vl-max", "qwen-vl-plus",
                    "qwen2.5-vl-72b-instruct", "qwen-vl-max-latest"]
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

                mensajes = [
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

                # 🤖 Se prueban los modelos en orden: si uno no existe (404) o no tiene
                # permiso/cuota (403), se intenta con el siguiente.
                datos_respuesta = None
                problemas = []
                for modelo in MODELOS_IA:
                    payload = {"model": modelo, "messages": mensajes}
                    try:
                        respuesta = requests.post(QWEN_API_URL, headers=headers, json=payload, timeout=120)
                    except Exception as e:
                        problemas.append("%s → no se pudo conectar (%s)" % (modelo, str(e)[:90]))
                        continue
                    if respuesta.status_code == 200:
                        try:
                            datos_respuesta = respuesta.json()
                            break
                        except Exception as e:
                            problemas.append("%s → respuesta ilegible (%s)" % (modelo, str(e)[:60]))
                            continue
                    detalle = ""
                    try:
                        detalle = str((respuesta.json() or {}).get("error", {}).get("message", ""))[:150]
                    except Exception:
                        detalle = str(respuesta.text or "")[:150]
                    problemas.append("%s → %s %s" % (modelo, respuesta.status_code, detalle))
                    if respuesta.status_code in (401, 402):
                        break   # clave inválida o sin saldo: no tiene sentido seguir probando

                if datos_respuesta is None:
                    raise RuntimeError("Ningún modelo de IA respondió.\n" + "\n".join(problemas))

                texto_ia = datos_respuesta['choices'][0]['message']['content']

                match = re.search(r'\{.*\}', texto_ia, re.DOTALL)
                if not match: raise ValueError("No se encontró JSON válido en la respuesta de la IA.")

                return json.loads(match.group(0))
            except Exception as e:
                return e

        def aplicar_resultado(datos):
            if isinstance(datos, Exception):
                pista = ""
                txt_err = str(datos)
                if "quota" in txt_err.lower() or "exhaust" in txt_err.lower():
                    pista = ("\n\n💡 La cuota gratuita del modelo se agotó. Puede poner su propia clave "
                             "de Alibaba/DashScope en: Configuración General → Ajustes (campo «Clave IA Qwen»).")
                elif "404" in txt_err or "not exist" in txt_err.lower():
                    pista = "\n\n💡 El modelo fue retirado; el sistema ya prueba otros automáticamente."
                messagebox.showerror("Error IA", f"No se pudo procesar el documento con Inteligencia Artificial.\nDetalle: {datos}{pista}")
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

        def correr(estado):
            # Hilo secundario: solo el trabajo pesado (PyMuPDF + API)
            estado["resultado"] = procesar_en_hilo()

        def pintar(estado):
            # Hilo principal: se rellenan los campos del formulario
            aplicar_resultado(estado.get("resultado"))

        ejecutar_en_hilo(self.parent_frame, correr, al_terminar=pintar)

    def abrir_tarjeta(self):
        from app_paths import resolver_ruta_archivo
        ruta = self.ruta_tarjeta_db if self.ruta_tarjeta_db else self.ruta_tarjeta_temp
        ruta = resolver_ruta_archivo(ruta) if ruta else ""
        if ruta:
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

        def tarea_api(estado):
            try:
                # Contexto SSL seguro y compatible con macOS (ver _contexto_ssl_seguro)
                ctx = _contexto_ssl_seguro()

                url = f"https://api.apis.net.pe/v1/vehiculos?placa={placa}"
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                
                with urllib.request.urlopen(req, context=ctx, timeout=5) as response:
                    if response.status == 200:
                        estado["datos"] = json.loads(response.read().decode())
                    else:
                        estado["aviso"] = ("warning", "Sin Resultados", "No se encontró la placa en la base de datos pública.")
            except urllib.error.HTTPError as e:
                if e.code in [401, 403, 404]:
                    estado["aviso"] = ("info", "API Restringida", "La consulta automática requiere API Key. Usa el botón 'Leer Tarjeta (IA)'.")
                else:
                    estado["aviso"] = ("warning", "Error de API", f"No se pudo contactar al servidor ({e.code}).")
            except Exception:
                estado["aviso"] = ("info", "Servicio no disponible", "La conexión a la base de datos vehicular no está disponible. Usa 'Leer Tarjeta (IA)'.")

        def pintar_api(estado):
            # Hilo principal: se actualizan los campos y se avisa al usuario
            if "datos" in estado:
                return self._aplicar_datos_api(estado["datos"])
            if estado.get("aviso"):
                tipo, titulo, mensaje = estado["aviso"]
                if tipo == "info":
                    messagebox.showinfo(titulo, mensaje)
                else:
                    messagebox.showwarning(titulo, mensaje)

        ejecutar_en_hilo(self.parent_frame, tarea_api, al_terminar=pintar_api)

    def _aplicar_datos_api(self, data):
        if "marca" in data: self.ent_marca.delete(0, tk.END); self.ent_marca.insert(0, data.get("marca", ""))
        if "modelo" in data: self.ent_modelo.delete(0, tk.END); self.ent_modelo.insert(0, data.get("modelo", ""))
        if "color" in data: self.ent_color.delete(0, tk.END); self.ent_color.insert(0, data.get("color", ""))
        messagebox.showinfo("Consulta", "Datos obtenidos exitosamente de la base de datos.")

    # =====================================================================
    # 📎 CERTIFICADOS (SOAT y EXTINTOR): se cargan en PDF o JPG/PNG
    # =====================================================================
    def _abrir_archivo_local(self, ruta):
        """Abre un archivo del equipo (Windows y macOS)."""
        try:
            if not ruta or not os.path.exists(ruta):
                return messagebox.showwarning("Certificado", "El archivo ya no está en esa carpeta.")
            if sys.platform == "darwin":
                subprocess.Popen(["open", ruta])
            elif sys.platform == "win32":
                os.startfile(ruta)          # noqa: S606 (solo Windows)
            else:
                subprocess.Popen(["xdg-open", ruta])
        except Exception as e:
            messagebox.showerror("Certificado", "No se pudo abrir el archivo:\n%s" % e)

    def _elegir_certificado(self, tipo):
        """Pide el PDF/JPG del certificado y devuelve la ruta (o "" si se cancela).

        Si ya había uno cargado, ofrece verlo o reemplazarlo.
        """
        ATRIBUTOS = {"soat": ("_soat_archivo", "SOAT"),
                     "extintor": ("_extintor_archivo", "Certificado de Extintor"),
                     "rt": ("_rt_archivo", "Revisión Técnica"),
                     "gas": ("_gas_archivo", "Revisión de Gas")}
        atributo, etiqueta = ATRIBUTOS.get(tipo, ("_soat_archivo", "Certificado"))
        actual = str(getattr(self, atributo, "") or "").strip()
        if actual and os.path.exists(actual):
            reemplazar = messagebox.askyesno(
                "Certificado ya cargado",
                "Ya hay un %s cargado:\n%s\n\n¿Quiere REEMPLAZARLO?\n"
                "(elija «No» para verlo)" % (etiqueta, os.path.basename(actual)))
            if not reemplazar:
                self._abrir_archivo_local(actual)
                return ""
        ruta = filedialog.askopenfilename(
            title="Seleccionar %s (PDF o imagen)" % etiqueta,
            filetypes=[("PDF o imagen", "*.pdf *.jpg *.jpeg *.png"), ("PDF", "*.pdf"),
                       ("Imágenes", "*.jpg *.jpeg *.png"), ("Todos", "*.*")])
        return ruta or ""

    def _texto_de_certificado(self, ruta):
        """Devuelve el texto del certificado: lo lee del PDF o le hace OCR a la imagen."""
        texto = ""
        try:
            if str(ruta).lower().endswith(".pdf"):
                try:
                    import pdfplumber
                    with pdfplumber.open(ruta) as pdf:
                        for pagina in pdf.pages:
                            texto += (pagina.extract_text() or "") + "\n"
                except Exception:
                    texto = ""
                if not texto.strip():
                    try:
                        import fitz
                        doc = fitz.open(ruta)
                        for pagina in doc:
                            texto += (pagina.get_text() or "") + "\n"
                        doc.close()
                    except Exception:
                        texto = ""
            if not texto.strip():
                import escaner
                resultado = escaner.ocr_paginas([ruta])
                if resultado.get("ok"):
                    texto = "\n".join(str(p.get("texto", "")) for p in (resultado.get("resultados") or []))
        except Exception as e:
            print("Aviso - lectura del certificado:", e)
        return texto or ""

    def _valor(self, nombre):
        """Valor de un campo del formulario ('' si no existe o falla)."""
        try:
            return str(getattr(self, nombre).get()).strip()
        except Exception:
            return ""

    def _completar_fechas_del_certificado(self, tipo, ruta):
        """Lee las fechas del certificado y completa los campos que estén vacíos."""
        try:
            texto = self._texto_de_certificado(ruta)
            if not texto.strip():
                return "No se pudo leer el texto del certificado (complete las fechas a mano)."
            emision, vencimiento = fechas_de_certificado(texto)
            if not emision and not vencimiento:
                return "No se encontraron fechas en el certificado (complete las a mano)."
            # Con una sola fecha, para RT/Gas/extintor se toma como VENCIMIENTO
            if emision and not vencimiento and tipo in ("rt", "gas", "extintor"):
                vencimiento, emision = emision, ""
            aplicados = []
            if tipo == "soat":
                if emision and not self.ent_emi_soat.get().strip():
                    self.ent_emi_soat.insert(0, emision); aplicados.append("emisión SOAT " + emision)
                if vencimiento and not self.ent_soat.get().strip():
                    self.ent_soat.insert(0, vencimiento); aplicados.append("vencimiento SOAT " + vencimiento)
            elif tipo == "extintor":
                if vencimiento and not self.ent_fec_venc_ext.get().strip():
                    self.ent_fec_venc_ext.insert(0, vencimiento); aplicados.append("vencimiento del extintor " + vencimiento)
            elif tipo == "rt":
                if vencimiento and not self.ent_rt.get().strip():
                    self.ent_rt.insert(0, vencimiento); aplicados.append("vencimiento de la revisión técnica " + vencimiento)
            elif tipo == "gas":
                if vencimiento and not self.ent_fec_rev_gas.get().strip():
                    self.ent_fec_rev_gas.insert(0, vencimiento); aplicados.append("vencimiento de la revisión de gas " + vencimiento)
            if not aplicados:
                return "El certificado trae fechas, pero los campos ya tenían valor (no se tocaron)."
            return "Se completaron las fechas: " + ", ".join(aplicados) + "."
        except Exception as e:
            return "No se pudieron leer las fechas del certificado: %s" % e

    def _ver_certificado(self, tipo):
        """Abre el certificado guardado (SOAT o extintor) tal como se cargó."""
        ATRIBUTOS = {"soat": ("_soat_archivo", "SOAT", "soat_archivo"),
                     "extintor": ("_extintor_archivo", "Certificado de Extintor", "extintor_archivo"),
                     "rt": ("_rt_archivo", "Revisión Técnica", "rt_archivo"),
                     "gas": ("_gas_archivo", "Revisión de Gas", "gas_archivo")}
        atributo, etiqueta, columna = ATRIBUTOS.get(tipo, ("_soat_archivo", "Certificado", "soat_archivo"))
        ruta = str(getattr(self, atributo, "") or "").strip()
        if not ruta:
            # Al editar un vehículo ya guardado puede estar en la base
            try:
                if getattr(self, "id_edicion", None):
                    conn = conectar_db()
                    if conn:
                        try:
                            cursor = conn.cursor()
                            cursor.execute("SELECT COALESCE(%s, '') FROM flota_vehiculos WHERE id = %%s"
                                           % columna, (self.id_edicion,))
                            fila = cursor.fetchone()
                            ruta = str(fila[0] or "") if fila else ""
                            if ruta:
                                setattr(self, atributo, ruta)
                        finally:
                            liberar_conexion(conn)
            except Exception as e:
                print("Aviso - certificado guardado:", e)
        if not ruta:
            return messagebox.showinfo("Certificado",
                                       "Todavía no hay %s cargado.\n\nUse el botón «📎 Cargar %s (PDF/JPG)»."
                                       % (etiqueta, etiqueta),
                                       parent=self.winfo_toplevel())
        self._abrir_archivo_local(ruta)

    def ver_certificado_soat(self):
        """Abre el SOAT cargado (PDF o imagen)."""
        return self._ver_certificado("soat")

    def ver_certificado_extintor(self):
        """Abre el certificado del extintor cargado (PDF o imagen)."""
        return self._ver_certificado("extintor")

    # 📎 Soportes de Revisión Técnica y de Gas (mismo comportamiento que SOAT/extintor)
    def cargar_soporte_rt(self):
        ruta = self._elegir_certificado("rt")
        if not ruta:
            return
        self._rt_archivo = ruta
        try:
            self.lbl_rt_archivo.configure(text="📎 %s" % os.path.basename(ruta), text_color="#1e8449")
        except Exception:
            pass
        aviso = self._completar_fechas_del_certificado("rt", ruta)
        messagebox.showinfo("Revisión Técnica cargada", "%s\n\n%s" % (os.path.basename(ruta), aviso),
                            parent=self.winfo_toplevel())

    def ver_soporte_rt(self):
        return self._ver_certificado("rt")

    def cargar_soporte_gas(self):
        ruta = self._elegir_certificado("gas")
        if not ruta:
            return
        self._gas_archivo = ruta
        try:
            self.lbl_gas_archivo.configure(text="📎 %s" % os.path.basename(ruta), text_color="#1e8449")
        except Exception:
            pass
        aviso = self._completar_fechas_del_certificado("gas", ruta)
        messagebox.showinfo("Revisión de Gas cargada", "%s\n\n%s" % (os.path.basename(ruta), aviso),
                            parent=self.winfo_toplevel())

    def ver_soporte_gas(self):
        return self._ver_certificado("gas")

    def cargar_certificado_soat(self):
        """Carga (o reemplaza) el certificado del SOAT del vehículo seleccionado."""
        ruta = self._elegir_certificado("soat")
        if not ruta:
            return
        self._soat_archivo = ruta
        try:
            self.lbl_soat_archivo.configure(text="📎 %s" % os.path.basename(ruta), text_color="#1e8449")
        except Exception:
            pass
        # 📅 Se leen las fechas del certificado y se completan los campos vacíos
        aviso = self._completar_fechas_del_certificado("soat", ruta)
        try:
            self.actualizar_proximo_aceite() if hasattr(self, "actualizar_proximo_aceite") else None
        except Exception:
            pass
        try:
            messagebox.showinfo("SOAT cargado", "%s\n\n%s" % (os.path.basename(ruta), aviso),
                                parent=ventana_para_mensaje(self) if "ventana_para_mensaje" in globals() else None)
        except Exception:
            pass

    def cargar_certificado_extintor(self):
        """Carga (o reemplaza) el certificado del extintor del vehículo seleccionado."""
        ruta = self._elegir_certificado("extintor")
        if not ruta:
            return
        self._extintor_archivo = ruta
        try:
            self.lbl_extintor_archivo.configure(text="📎 %s" % os.path.basename(ruta), text_color="#1e8449")
        except Exception:
            pass
        aviso_ext = self._completar_fechas_del_certificado("extintor", ruta)
        try:
            messagebox.showinfo("Certificado de extintor cargado",
                                "%s\n\n%s" % (os.path.basename(ruta), aviso_ext),
                                parent=ventana_para_mensaje(self) if "ventana_para_mensaje" in globals() else None)
        except Exception:
            pass

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
        # 🛢️ Kilometraje al que se hizo ese cambio: sirve para calcular el próximo
        self.ent_km_aceite = crear_campo("Km del Último Cambio de Aceite:", "Ej: 45000")
        self.lbl_prox_aceite = ctk.CTkLabel(self.f_form, text="", font=("Arial", 11),
                                            text_color="#e67e22", justify="left", anchor="w",
                                            wraplength=330)
        self.lbl_prox_aceite.pack(fill="x", padx=10, pady=(0, 8))
        # El cálculo del próximo cambio se actualiza mientras se escriben los km
        self.ent_km_aceite.bind("<KeyRelease>", lambda e: self.actualizar_proximo_aceite())
        self.ent_kilometraje.bind("<KeyRelease>", lambda e: self.actualizar_proximo_aceite())
        self.ent_fec_correa = crear_campo_fecha("Cambio Correa/Cadena Distribución:")
        self.ent_km_prox_correa = crear_campo("KM Próx. Cambio Correa:", "Ej: 100000")
        self.ent_bat_marca = crear_campo("Marca de Batería:", "Ej: Bosch")
        self.ent_bat_modelo = crear_campo("Modelo de Batería:", "Ej: S4 12V 70Ah")
        self.ent_fec_compra_bat = crear_campo_fecha("Compra de Batería:")
        self.ent_fec_venc_bat = crear_campo_fecha("Vencimiento Garantía Batería:")

        # --- 🛞 NEUMÁTICOS ---
        ctk.CTkLabel(self.f_form, text="--- Neumáticos ---", font=("Arial", 11, "bold"),
                     text_color="#2c3e50").pack(anchor="w", padx=10, pady=(6, 2))
        self.ent_neu_marca = crear_campo("Marca de Neumáticos:", "Ej: Michelin")
        self.ent_neu_modelo = crear_campo("Modelo de Neumáticos:", "Ej: Primacy 4")
        self.ent_neu_medidas = crear_campo("Medidas:", "Ej: 195/65 R15")
        self.ent_neu_km = crear_campo("Kilómetros del Cambio:", "Ej: 52000")
        self.ent_neu_fecha = crear_campo_fecha("Fecha del Cambio de Neumáticos:")

        # 4. Seguridad (Extintor)
        ctk.CTkLabel(self.f_form, text="--- Seguridad y Extintores ---", font=("Arial", 11, "bold"), text_color="#c0392b").pack(anchor="w", padx=10, pady=(10,5))
        self.ent_extintor_num = crear_campo("N° Certificado Extintor:")
        self.ent_fec_venc_ext = crear_campo_fecha("Vencimiento del Extintor:")
        # 📎 Certificado del extintor (PDF o JPG)
        f_ext_botones = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_ext_botones.pack(fill="x", padx=10, pady=(0, 2))
        ctk.CTkButton(f_ext_botones, text="📎 Cargar Certificado de Extintor (PDF/JPG)",
                      fg_color="#c0392b", hover_color="#96281b", font=("Arial", 11, "bold"),
                      command=self.cargar_certificado_extintor).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(f_ext_botones, text="👁️ Ver", width=64, fg_color="#5d6d7e", hover_color="#4a5967",
                      font=("Arial", 11, "bold"),
                      command=self.ver_certificado_extintor).pack(side="left", padx=(4, 0))
        self.lbl_extintor_archivo = ctk.CTkLabel(self.f_form, text="Sin certificado cargado",
                                                 font=("Arial", 10), text_color="#999999")
        self.lbl_extintor_archivo.pack(anchor="w", padx=10, pady=(0, 8))

        # 5. SOAT y Seguro
        ctk.CTkLabel(self.f_form, text="--- SOAT y Póliza de Seguro ---", font=("Arial", 11, "bold"), text_color="#27ae60").pack(anchor="w", padx=10, pady=(10,5))
        self.ent_emi_soat = crear_campo_fecha("Emisión SOAT:")
        self.ent_soat = crear_campo_fecha("Vencimiento SOAT:")
        # 📎 Certificado del SOAT (PDF o JPG)
        f_soat_botones = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_soat_botones.pack(fill="x", padx=10, pady=(0, 2))
        ctk.CTkButton(f_soat_botones, text="📎 Cargar SOAT (PDF/JPG)",
                      fg_color="#27ae60", hover_color="#1e8449", font=("Arial", 11, "bold"),
                      command=self.cargar_certificado_soat).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(f_soat_botones, text="👁️ Ver", width=64, fg_color="#5d6d7e", hover_color="#4a5967",
                      font=("Arial", 11, "bold"),
                      command=self.ver_certificado_soat).pack(side="left", padx=(4, 0))
        self.lbl_soat_archivo = ctk.CTkLabel(self.f_form, text="Sin SOAT cargado",
                                             font=("Arial", 10), text_color="#999999")
        self.lbl_soat_archivo.pack(anchor="w", padx=10, pady=(0, 8))
        self.ent_emi_seguro = crear_campo_fecha("Emisión del Seguro:")
        self.ent_venc_seguro = crear_campo_fecha("Vencimiento del Seguro:")
        self.ent_seguro_num = crear_campo("N° de Póliza / Seguro:")
        
        # 6. Revisiones y Estado
        ctk.CTkLabel(self.f_form, text="--- Revisiones y Estado ---", font=("Arial", 11, "bold"), text_color="#8e44ad").pack(anchor="w", padx=10, pady=(10,5))
        self.ent_rt = crear_campo_fecha("Venc. Revisión Técnica:")
        f_rt_botones = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_rt_botones.pack(fill="x", padx=10, pady=(0, 2))
        ctk.CTkButton(f_rt_botones, text="📎 Cargar Revisión Técnica (PDF/JPG)",
                      fg_color="#8e44ad", hover_color="#703688", font=("Arial", 11, "bold"),
                      command=self.cargar_soporte_rt).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(f_rt_botones, text="👁️ Ver", width=64, fg_color="#5d6d7e", hover_color="#4a5967",
                      font=("Arial", 11, "bold"),
                      command=self.ver_soporte_rt).pack(side="left", padx=(4, 0))
        self.lbl_rt_archivo = ctk.CTkLabel(self.f_form, text="Sin soporte cargado",
                                                  font=("Arial", 10), text_color="#999999")
        self.lbl_rt_archivo.pack(anchor="w", padx=10, pady=(0, 8))
        # --- REVISIÓN del sistema de gas (con su certificado) ---
        ctk.CTkLabel(self.f_form, text="--- Revisión del Sistema de Gas ---", font=("Arial", 11, "bold"),
                     text_color="#d35400").pack(anchor="w", padx=10, pady=(6, 2))
        self.ent_fec_rev_gas = crear_campo_fecha("Venc. Rev. Sist. Gas:")
        f_gas_botones = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_gas_botones.pack(fill="x", padx=10, pady=(0, 2))
        ctk.CTkButton(f_gas_botones, text="📎 Cargar Revisión de Gas (PDF/JPG)",
                      fg_color="#d35400", hover_color="#a04000", font=("Arial", 11, "bold"),
                      command=self.cargar_soporte_gas).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(f_gas_botones, text="👁️ Ver", width=64, fg_color="#5d6d7e", hover_color="#4a5967",
                      font=("Arial", 11, "bold"),
                      command=self.ver_soporte_gas).pack(side="left", padx=(4, 0))
        self.lbl_gas_archivo = ctk.CTkLabel(self.f_form, text="Sin soporte cargado",
                                                  font=("Arial", 10), text_color="#999999")
        self.lbl_gas_archivo.pack(anchor="w", padx=10, pady=(0, 8))

        # --- MANTENIMIENTO del sistema de gas (emisión + kilómetros) ---
        ctk.CTkLabel(self.f_form, text="--- Mantenimiento del Sistema de Gas ---", font=("Arial", 11, "bold"),
                     text_color="#a04000").pack(anchor="w", padx=10, pady=(6, 2))
        self.ent_emi_rev_gas = crear_campo_fecha("Fecha de Emisión del Mantenimiento:")
        self.ent_km_rev_gas = crear_campo("Km del Mantenimiento:", "Ej: 52000")
        self.ent_km_rev_gas.bind("<KeyRelease>", lambda e: self.actualizar_proximo_gas())
        self.lbl_prox_gas = ctk.CTkLabel(self.f_form, text="", font=("Arial", 11),
                                         text_color="#a04000", justify="left")
        self.lbl_prox_gas.pack(fill="x", padx=10, pady=(0, 8))

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

        columnas = ("id", "placa", "vehiculo", "categoria", "color", "combustible", "kilometraje", "km_aceite", "estado", "titulo", "fecha_titulo")
        self.tabla = ttk.Treeview(f_tabla, columns=columnas, show="headings")
        self.tabla.heading("id", text="ID")
        self.tabla.heading("placa", text="Placa")
        self.tabla.heading("vehiculo", text="Marca / Modelo (Año)")
        self.tabla.heading("categoria", text="Categoría")
        self.tabla.heading("color", text="Color")
        self.tabla.heading("combustible", text="Combustible")
        self.tabla.heading("kilometraje", text="Kilometraje")
        self.tabla.heading("km_aceite", text="Km Últ. Aceite")
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
        self.tabla.column("km_aceite", width=100, anchor="center")
        self.tabla.column("estado", width=100, anchor="center")
        self.tabla.column("titulo", width=0, stretch=tk.NO)
        self.tabla.column("fecha_titulo", width=0, stretch=tk.NO)
        
        self.tabla.config(displaycolumns=("placa", "vehiculo", "categoria", "color", "combustible", "kilometraje", "km_aceite", "estado"))

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
        try:
            self.ent_km_aceite.delete(0, tk.END)
            self.lbl_prox_aceite.configure(text="")
        except Exception:
            pass
        self.ent_motor.delete(0, tk.END)
        self.ent_carroceria.delete(0, tk.END)
        self.cmb_combustible.set("Gasolina")
        self.ent_fec_aceite.delete(0, tk.END)
        self.ent_fec_correa.delete(0, tk.END)
        self.ent_km_prox_correa.delete(0, tk.END)
        self.ent_fec_rev_gas.delete(0, tk.END)
        try:
            self.ent_emi_rev_gas.delete(0, tk.END)
            self.ent_km_rev_gas.delete(0, tk.END)
            self.lbl_prox_gas.configure(text="")
        except Exception:
            pass
        self.ent_fec_compra_bat.delete(0, tk.END)
        self.ent_fec_venc_bat.delete(0, tk.END)
        for campo in ("ent_bat_marca", "ent_bat_modelo", "ent_neu_marca", "ent_neu_modelo",
                      "ent_neu_medidas", "ent_neu_km", "ent_neu_fecha"):
            try:
                getattr(self, campo).delete(0, tk.END)
            except Exception:
                pass
        self.ent_extintor_num.delete(0, tk.END)
        self.ent_fec_venc_ext.delete(0, tk.END)
        self.ent_emi_soat.delete(0, tk.END)
        self.ent_soat.delete(0, tk.END)
        self.ent_emi_seguro.delete(0, tk.END)
        self.ent_venc_seguro.delete(0, tk.END)
        self.ent_rt.delete(0, tk.END)
        self.cmb_estado.set("Operativo")
        # 📎 Certificados
        self._soat_archivo = ""
        self._extintor_archivo = ""
        self._rt_archivo = ""
        self._gas_archivo = ""
        try:
            self.ent_seguro_num.delete(0, tk.END)
        except Exception:
            pass
        try:
            self.lbl_soat_archivo.configure(text="Sin SOAT cargado", text_color="#999999")
            self.lbl_extintor_archivo.configure(text="Sin certificado cargado", text_color="#999999")
            self.lbl_rt_archivo.configure(text="Sin soporte cargado", text_color="#999999")
            self.lbl_gas_archivo.configure(text="Sin soporte cargado", text_color="#999999")
        except Exception:
            pass

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
            self.tabla.insert("", tk.END, values=("", "", "Cargando datos...", "", "", "", "", "", "", "", ""))
            
            def tarea_descarga(estado):
                conn = conectar_db(silencioso=True)
                if not conn: return
                try:
                    cursor = conn.cursor()
                    if filtro:
                        cursor.execute("""
                            SELECT id, placa, marca, modelo, anio, color, tipo_combustible, kilometraje, estado, categoria, nro_titulo, fecha_titulo, COALESCE(km_ultimo_aceite, 0)
                            FROM flota_vehiculos 
                            WHERE placa ILIKE %s OR marca ILIKE %s OR modelo ILIKE %s 
                            ORDER BY id DESC LIMIT %s OFFSET %s
                        """, (f"%{filtro}%", f"%{filtro}%", f"%{filtro}%", self.registros_por_pagina, offset))
                    else:
                        cursor.execute("""
                            SELECT id, placa, marca, modelo, anio, color, tipo_combustible, kilometraje, estado, categoria, nro_titulo, fecha_titulo, COALESCE(km_ultimo_aceite, 0)
                            FROM flota_vehiculos 
                            ORDER BY id DESC LIMIT %s OFFSET %s
                        """, (self.registros_por_pagina, offset))
                    
                    datos_db = cursor.fetchall()
                    cache_sistema.guardar(clave_cache, datos_db)
                    estado["vehiculos"] = datos_db
                except Exception as e:
                    print(f"Error cargando tabla de flota: {e}")
                finally:
                    liberar_conexion(conn)

            def pintar_descarga(estado):
                # Hilo principal: si no hubo conexión o falló la consulta, la tabla queda igual
                if "vehiculos" not in estado:
                    return
                self._pintar_vehiculos(estado["vehiculos"])

            ejecutar_en_hilo(self.parent_frame, tarea_descarga, al_terminar=pintar_descarga)

    def _pintar_vehiculos(self, datos):
        for item in self.tabla.get_children():
            self.tabla.delete(item)

        for r in datos:
            if len(r) < 13:                      # caché de una versión anterior
                r = tuple(r) + (0,) * (13 - len(r))
            v_id, placa, marca, modelo, anio, color, comb, km, est, categoria, nro_titulo, fecha_titulo, km_aceite = r
            vehiculo_nom = f"{marca} {modelo} ({anio})"
            km_mostrar = km if km else "0"
            cat_mostrar = categoria if categoria else "-"
            # 🛢️ Km del último cambio de aceite (y aviso si ya toca el próximo)
            km_aceite_num = a_numero(km_aceite)
            if km_aceite_num > 0:
                prox = km_aceite_num + limite_km_aceite()
                km_aceite_txt = f"{km_aceite_num:,.0f}" + (" ⚠️" if a_numero(km) >= prox else "")
            else:
                km_aceite_txt = "-"
            
            self.tabla.insert("", tk.END, values=(v_id, placa, vehiculo_nom, cat_mostrar, color, comb, km_mostrar, km_aceite_txt, est, nro_titulo, fecha_titulo))
            
        if self.pagina_actual > 1:
            self.btn_ant.configure(state="normal")
        else:
            self.btn_ant.configure(state="disabled")
            
        if len(datos) == self.registros_por_pagina:
            self.btn_sig.configure(state="normal")
        else:
            self.btn_sig.configure(state="disabled")

    def actualizar_proximo_gas(self):
        """Muestra a cuántos km toca el próximo mantenimiento del sistema de gas."""
        try:
            limite = limite_km_gas()
            km_actual = a_numero(self.ent_kilometraje.get())
            km_gas = a_numero(self.ent_km_rev_gas.get())
            if km_gas <= 0:
                txt = ("💨 Indique los km del mantenimiento de gas para calcular el próximo "
                       f"(cada {limite:,.0f} km).")
            else:
                prox = km_gas + limite
                txt = f"💨 Próximo mantenimiento de gas: a los {prox:,.0f} km"
                if km_actual <= 0:
                    txt += " (registre el kilometraje actual para ver cuánto falta)."
                else:
                    faltan = prox - km_actual
                    if faltan <= 0:
                        txt += f" · ⚠️ ¡VENCIDO! Ya pasó por {abs(faltan):,.0f} km."
                    else:
                        txt += f" · Faltan {faltan:,.0f} km."
            self.lbl_prox_gas.configure(text=txt)
        except Exception:
            pass

    def actualizar_proximo_aceite(self):
        """Muestra cuándo toca el próximo cambio de aceite, según los km.

        Se calcula con el kilometraje del ÚLTIMO cambio + el límite configurado
        en Configuración General ('alerta_aceite_km', por defecto 5.000 km).
        """
        try:
            limite = limite_km_aceite()
            km_actual = a_numero(self.ent_kilometraje.get())
            km_aceite = a_numero(self.ent_km_aceite.get())

            if km_aceite <= 0:
                txt = (f"🛢️ Indique los km del último cambio para calcular el próximo "
                       f"(cada {limite:,.0f} km).")
            else:
                prox = km_aceite + limite
                txt = f"🛢️ Próximo cambio de aceite: a los {prox:,.0f} km"
                if km_actual <= 0:
                    txt += " (registre el kilometraje actual para ver cuánto falta)."
                else:
                    faltan = prox - km_actual
                    if faltan <= 0:
                        txt += f" · ⚠️ ¡VENCIDO! Ya pasó por {abs(faltan):,.0f} km."
                    else:
                        txt += f" · Faltan {faltan:,.0f} km."
            self.lbl_prox_aceite.configure(text=txt)
        except Exception:
            pass

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
        km_aceite = a_numero(self.ent_km_aceite.get())   # 🛢️ km del último cambio de aceite
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
            # 🔒 Política de almacenamiento: sin autorización no se guarda ningún archivo
            try:
                from politica_almacenamiento import exigir_permiso
                if not exigir_permiso(self):
                    return
            except ImportError:
                pass
            try:
                # 🔗 El archivo va SIEMPRE a la carpeta configurada en
                # "Configuración General" (la misma que sincroniza Rclone con la
                # cuenta del equipo principal), NO a la carpeta del programa: así la
                # tarjeta se ve desde cualquier equipo (Windows y Mac) y se respalda.
                try:
                    from politica_almacenamiento import ruta_base_autorizada
                    ruta_base = ruta_base_autorizada(mostrar_alerta=False)
                except Exception:
                    ruta_base = ""
                if not ruta_base:
                    return messagebox.showwarning(
                        "Configuración Requerida",
                        "No se pudo guardar la Tarjeta de Propiedad:\n"
                        "no hay una carpeta válida configurada en Configuración General.")
                carpeta_destino = os.path.join(ruta_base, "archivos_flota", "tarjetas_propiedad")
                os.makedirs(carpeta_destino, exist_ok=True)
                
                ext = os.path.splitext(self.ruta_tarjeta_temp)[1]
                nombre_archivo = f"Tarjeta_{placa}_{datetime.now().strftime('%Y%m%d%H%M%S')}{ext}"
                ruta_final_archivo = os.path.join(carpeta_destino, nombre_archivo)
                
                shutil.copy2(self.ruta_tarjeta_temp, ruta_final_archivo)
                # Se guarda la ruta PORTABLE (relativa a la carpeta sincronizada):
                # sirve igual en Windows y en Mac.
                try:
                    from app_paths import ruta_para_guardar
                    ruta_final_archivo = ruta_para_guardar(ruta_final_archivo)
                except Exception:
                    pass
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
                    extintor_num=%s, fec_venc_extintor=%s, km_ultimo_aceite=%s, fecha_ultimo_aceite=%s,
                    soat_archivo=%s, extintor_archivo=%s,
                    seguro_num=%s, rt_archivo=%s, gas_archivo=%s,
                    emision_mant_gas=%s, km_mant_gas=%s,
                    bat_marca=%s, bat_modelo=%s,
                    neu_marca=%s, neu_modelo=%s, neu_medidas=%s, neu_km=%s, neu_fecha=%s
                    WHERE id=%s
                """, (placa, marca, modelo, anio, color, kilometraje, motor, carroceria, combustible, 
                      emi_soat, venc_soat, emi_seguro, venc_seguro, rt, estado, categoria, titulo, 
                      fecha_titulo, ruta_final_archivo, fec_aceite, fec_correa, km_prox_correa, fec_gas, fec_compra_bat, 
                      fec_venc_bat, num_extintor, fec_venc_ext, km_aceite, fec_aceite,
                      str(getattr(self, "_soat_archivo", "") or ""),
                      str(getattr(self, "_extintor_archivo", "") or ""),
                      self.ent_seguro_num.get().strip() if hasattr(self, "ent_seguro_num") else "",
                      str(getattr(self, "_rt_archivo", "") or ""),
                      str(getattr(self, "_gas_archivo", "") or ""),
                      self.ent_emi_rev_gas.get().strip() if hasattr(self, "ent_emi_rev_gas") else "",
                      self.ent_km_rev_gas.get().strip() if hasattr(self, "ent_km_rev_gas") else "",
                      self._valor("ent_bat_marca"), self._valor("ent_bat_modelo"),
                      self._valor("ent_neu_marca"), self._valor("ent_neu_modelo"),
                      self._valor("ent_neu_medidas"), self._valor("ent_neu_km"), self._valor("ent_neu_fecha"),
                      self.id_edicion))
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
                    fec_venc_bat, extintor_num, fec_venc_extintor, km_ultimo_aceite, fecha_ultimo_aceite,
                    soat_archivo, extintor_archivo, seguro_num, rt_archivo, gas_archivo,
                    emision_mant_gas, km_mant_gas,
                    bat_marca, bat_modelo,
                    neu_marca, neu_modelo, neu_medidas, neu_km, neu_fecha) 
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (placa, marca, modelo, anio, color, kilometraje, motor, carroceria, combustible, 
                      emi_soat, venc_soat, emi_seguro, venc_seguro, rt, estado, categoria, titulo, 
                      fecha_titulo, ruta_final_archivo, fec_aceite, fec_correa, km_prox_correa, fec_gas, fec_compra_bat, 
                      fec_venc_bat, num_extintor, fec_venc_ext, km_aceite, fec_aceite,
                      str(getattr(self, "_soat_archivo", "") or ""),
                      str(getattr(self, "_extintor_archivo", "") or ""),
                      self.ent_seguro_num.get().strip() if hasattr(self, "ent_seguro_num") else "",
                      str(getattr(self, "_rt_archivo", "") or ""),
                      str(getattr(self, "_gas_archivo", "") or ""),
                      self.ent_emi_rev_gas.get().strip() if hasattr(self, "ent_emi_rev_gas") else "",
                      self.ent_km_rev_gas.get().strip() if hasattr(self, "ent_km_rev_gas") else "",
                      self._valor("ent_bat_marca"), self._valor("ent_bat_modelo"),
                      self._valor("ent_neu_marca"), self._valor("ent_neu_modelo"),
                      self._valor("ent_neu_medidas"), self._valor("ent_neu_km"), self._valor("ent_neu_fecha")))
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
                km_prox_correa, COALESCE(km_ultimo_aceite, 0), COALESCE(fecha_ultimo_aceite, ''),
                COALESCE(soat_archivo, ''), COALESCE(extintor_archivo, ''),
                COALESCE(seguro_num, ''), COALESCE(rt_archivo, ''), COALESCE(gas_archivo, ''),
                COALESCE(emision_mant_gas, ''), COALESCE(km_mant_gas, ''),
                COALESCE(bat_marca, ''), COALESCE(bat_modelo, ''),
                COALESCE(neu_marca, ''), COALESCE(neu_modelo, ''), COALESCE(neu_medidas, ''),
                COALESCE(neu_km, ''), COALESCE(neu_fecha, '')
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
                # 📎 Certificados guardados (índices al final del SELECT)
                try:
                    # El SELECT los trae al final: fecha_ultimo_aceite es r[29],
                    # así que los certificados son r[30] y r[31].
                    self._soat_archivo = str(r[30] or "")
                    self._extintor_archivo = str(r[31] or "")
                    if self._soat_archivo:
                        self.lbl_soat_archivo.configure(text="📎 %s" % os.path.basename(self._soat_archivo),
                                                        text_color="#1e8449")
                    if self._extintor_archivo:
                        self.lbl_extintor_archivo.configure(
                            text="📎 %s" % os.path.basename(self._extintor_archivo), text_color="#1e8449")
                    # 🛡️ N° de seguro y soportes de RT / gas (r[32], r[33], r[34])
                    try:
                        if r[32]:
                            self.ent_seguro_num.insert(0, str(r[32]))
                    except Exception:
                        pass
                    self._rt_archivo = str(r[33] or "")
                    self._gas_archivo = str(r[34] or "")
                    if self._rt_archivo:
                        self.lbl_rt_archivo.configure(text="📎 %s" % os.path.basename(self._rt_archivo),
                                                      text_color="#1e8449")
                    if self._gas_archivo:
                        self.lbl_gas_archivo.configure(text="📎 %s" % os.path.basename(self._gas_archivo),
                                                       text_color="#1e8449")
                    # 💨 Mantenimiento del sistema de gas (r[35] y r[36])
                    try:
                        if r[35]:
                            self.ent_emi_rev_gas.insert(0, str(r[35]))
                        if r[36]:
                            self.ent_km_rev_gas.insert(0, str(r[36]))
                    except Exception:
                        pass
                    try:
                        self.actualizar_proximo_gas()
                    except Exception:
                        pass
                    # 🔋 Batería (r[37], r[38]) y 🛞 neumáticos (r[39] a r[43])
                    for campo, indice in (("ent_bat_marca", 37), ("ent_bat_modelo", 38),
                                          ("ent_neu_marca", 39), ("ent_neu_modelo", 40),
                                          ("ent_neu_medidas", 41), ("ent_neu_km", 42),
                                          ("ent_neu_fecha", 43)):
                        try:
                            if r[indice]:
                                getattr(self, campo).insert(0, str(r[indice]))
                        except Exception:
                            continue
                except Exception as e_cert:
                    print("Aviso - certificados del vehículo:", e_cert)
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
                # 🛢️ Km del último cambio de aceite (y su fecha, si no estaba en el campo antiguo)
                if r[28]:
                    self.ent_km_aceite.insert(0, f"{float(r[28]):,.0f}")
                if not self.ent_fec_aceite.get().strip() and r[29]:
                    self.ent_fec_aceite.insert(0, r[29])
                self.actualizar_proximo_aceite()

                # El botón se habilita si la tarjeta se puede encontrar en ESTE equipo
                # (la ruta guardada puede venir de otro equipo o ser relativa).
                ruta_tarjeta_real = ""
                if self.ruta_tarjeta_db:
                    try:
                        from app_paths import resolver_ruta_archivo
                        ruta_tarjeta_real = resolver_ruta_archivo(self.ruta_tarjeta_db) or ""
                    except Exception:
                        ruta_tarjeta_real = ""
                if ruta_tarjeta_real:
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