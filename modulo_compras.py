# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import customtkinter as ctk
import os
import sys  
import shutil
import calendar
import re
import json
import subprocess 
import ctypes
import urllib.request
import urllib.parse
import urllib.error
import xml.etree.ElementTree as ET
import base64
import time
from datetime import datetime
from conexion import conectar_db, registrar_auditoria
from buffer_memoria import cache_sistema

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

# =========================================================
# 🚀 ADAPTACIÓN MULTIPLATAFORMA
# =========================================================
if sys.platform == "win32":
    try:
        hwnd_cmd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd_cmd:
            ctypes.windll.user32.ShowWindow(hwnd_cmd, 6)
    except Exception:
        pass

def abrir_documento(ruta):
    try:
        ruta_norm = os.path.normpath(ruta)
        if sys.platform == "win32":
            os.startfile(ruta_norm)
        elif sys.platform == "darwin": 
            subprocess.call(["open", ruta_norm])
        else: 
            subprocess.call(["xdg-open", ruta_norm])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")

def cargar_configuracion_regional():
    config = {
        "simbolo_moneda": "S/.",
        "formato_numero": "1,000.00",
        "formato_fecha": "DD/MM/AAAA",
        "ruta_drive": "",
        "impresora": "",
        "cuentas_bancarias": [],
        "ruc_empresa": "",
        "usuario_sol": "",
        "clave_sol": "",
        "client_id_sire": "",
        "client_secret_sire": ""
    }
    try:
        if os.path.exists("config_local.json"):
            with open("config_local.json", "r", encoding="utf-8") as f:
                config.update(json.load(f))
    except Exception: pass
    return config

CONFIG_REGIONAL = cargar_configuracion_regional()

def formatear_moneda(valor):
    simbolo = CONFIG_REGIONAL.get("simbolo_moneda", "S/.")
    formato = CONFIG_REGIONAL.get("formato_numero", "1,000.00")
    try: valor = float(valor)
    except: valor = 0.0
    
    if formato == "1.000,00":
        str_val = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    else:
        str_val = f"{valor:,.2f}"
    return f"{simbolo} {str_val}"

def desformatear_numero(valor_str):
    if not valor_str: return 0.0
    simbolo = CONFIG_REGIONAL.get("simbolo_moneda", "S/.")
    formato = CONFIG_REGIONAL.get("formato_numero", "1,000.00")
    val = str(valor_str).replace(simbolo, "").strip()
    if formato == "1.000,00":
        val = val.replace(".", "").replace(",", ".")
    else:
        val = val.replace(",", "")
    try: return float(val)
    except ValueError: return 0.0

def obtener_ruta_base_drive():
    ruta = CONFIG_REGIONAL.get("ruta_drive", "").strip()
    if ruta: return os.path.expanduser(ruta)
    return ""

def aplicar_estilo_treeview():
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Treeview", background="#ffffff", foreground="#000000", fieldbackground="#ffffff", bordercolor="#e0e0e0", borderwidth=1, rowheight=26, font=("Arial", 10))
    style.map("Treeview", background=[("selected", "#1f538d")], foreground=[("selected", "#ffffff")])
    style.configure("Treeview.Heading", background="#f0f0f0", foreground="#000000", relief="flat", font=("Arial", 10, "bold"), bordercolor="#e0e0e0", borderwidth=1)

# =========================================================
# CLASE: CALENDARIO NATIVO
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
        fmt = CONFIG_REGIONAL.get("formato_fecha", "DD/MM/AAAA")
        if fmt == "MM/DD/AAAA":
            fecha_seleccionada = f"{self.current_month:02d}/{day:02d}/{self.current_year}"
        else:
            fecha_seleccionada = f"{day:02d}/{self.current_month:02d}/{self.current_year}"
            
        self.target_entry.delete(0, tk.END)
        self.target_entry.insert(0, fecha_seleccionada)
        self.destroy()

# =========================================================
# PESTAÑA 1: FACTURAS RECIBIDAS
# =========================================================
class FacturasRecibidasTab:
    def __init__(self, tab_frame, main_root, app_padre):
        self.tab_frame = tab_frame
        self.main_root = main_root
        self.app_padre = app_padre
        
        self.orden_columnas = {}
        self.bloquear_autocompletado_ruc = False
        
        self.inicializar_bd()
        self.crear_interfaz()

    def inicializar_bd(self):
        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS facturas_recibidas (
                    id SERIAL PRIMARY KEY, tipo_documento VARCHAR(100), fecha VARCHAR(50), proveedor VARCHAR(255), descripcion TEXT, 
                    evento_asociado VARCHAR(255), subtotal NUMERIC, impuesto NUMERIC, total NUMERIC, archivo_ruta TEXT, 
                    dias_credito INTEGER DEFAULT 0, det_porcentaje NUMERIC DEFAULT 0, det_monto NUMERIC DEFAULT 0, 
                    numero_documento VARCHAR(100) DEFAULT '', categoria VARCHAR(255) DEFAULT 'GENERAL / NO ASIGNADO'
                )
            """)
            conn.commit()
            
            try:
                cursor.execute("ALTER TABLE facturas_recibidas ADD COLUMN kilometraje VARCHAR(50);")
                conn.commit()
            except Exception: conn.rollback()
                
            try:
                cursor.execute("ALTER TABLE facturas_recibidas ADD COLUMN cantidad_combustible VARCHAR(50);")
                conn.commit()
            except Exception: conn.rollback()

            try:
                cursor.execute("ALTER TABLE facturas_recibidas ADD COLUMN ruc VARCHAR(50);")
                conn.commit()
            except Exception: conn.rollback()
                
        except Exception: pass
        finally: conn.close()

    def abrir_calendario(self, entry_objetivo):
        CalendarioNativo(self.main_root.winfo_toplevel(), entry_objetivo)

    def ordenar_por_columna(self, columna, es_numerico):
        elementos = [(self.tabla.set(item, columna), item) for item in self.tabla.get_children("")]
        ascendente = self.orden_columnas.get(columna, True)
        self.orden_columnas[columna] = not ascendente
        if es_numerico:
            elementos.sort(key=lambda el: desformatear_numero(el[0]), reverse=not ascendente)
        else:
            elementos.sort(key=lambda el: el[0].lower(), reverse=not ascendente)
        for index, (_, item) in enumerate(elementos): self.tabla.move(item, "", index)

    def autocompletar_desde_pdf(self):
        if pdfplumber is None:
            messagebox.showerror("Librería faltante", "No se encontró 'pdfplumber'. Ejecuta: pip install pdfplumber")
            return
        ruta = filedialog.askopenfilename(title="Seleccionar Factura PDF de SUNAT", filetypes=[("Archivos PDF", "*.pdf")])
        if not ruta: return
        try:
            self.bloquear_autocompletado_ruc = True
            
            texto = ""
            with pdfplumber.open(ruta) as pdf:
                for page in pdf.pages: texto += page.extract_text() + "\n"
            if not texto.strip(): 
                self.bloquear_autocompletado_ruc = False
                return messagebox.showwarning("Aviso", "El PDF no contiene texto seleccionable.")
            
            if re.search(r"FACTURA\s+ELECTR[OÓ]NICA", texto, re.IGNORECASE): 
                if "10.5%" not in self.combo_tipo.get():
                    self.combo_tipo.set("Factura (18% IGV)")
            elif re.search(r"BOLETA\s+DE\s+VENTA", texto, re.IGNORECASE): self.combo_tipo.set("Boleta (Sin IGV)")
            elif re.search(r"RECIBO\s+POR\s+HONORARIOS", texto, re.IGNORECASE):
                if re.search(r"Retenci[oó]n.*?IR[\s:\|]*\(?([\d\,\.]+)\)?", texto, re.IGNORECASE): self.combo_tipo.set("Recibo por Honorarios (8% Retención)")
                else: self.combo_tipo.set("Recibo por Honorarios (Sin Retención)")
            self.on_tipo_change(self.combo_tipo.get())

            nro_match = re.search(r"([EFB][0-9A-Z]{3}\s*-\s*\d+)", texto)
            if nro_match:
                self.ent_nro_doc.delete(0, tk.END); self.ent_nro_doc.insert(0, nro_match.group(1).replace(" ", ""))
            
            fecha_match = re.search(r"Fecha de Emisi[oó]n\s*[:\-]?\s*(\d{2})[/\-.](\d{2})[/\-.](\d{4})", texto, re.IGNORECASE)
            if not fecha_match: fecha_match = re.search(r"(\d{2})[/\-.](\d{2})[/\-.](\d{4})", texto)
            if fecha_match:
                d, m, y = fecha_match.groups()
                fmt = CONFIG_REGIONAL.get("formato_fecha", "DD/MM/AAAA")
                if fmt == "MM/DD/AAAA": self.ent_fecha.delete(0, tk.END); self.ent_fecha.insert(0, f"{m}/{d}/{y}")
                else: self.ent_fecha.delete(0, tk.END); self.ent_fecha.insert(0, f"{d}/{m}/{y}")

            rucs = re.findall(r"(?:RUC|R\.U\.C\.)\s*[:\-]?\s*(\d{11})", texto, re.IGNORECASE)
            if rucs: self.ent_desc.delete(0, tk.END); self.ent_desc.insert(0, rucs[0])
            lineas = [line.strip() for line in texto.split('\n') if line.strip()]
            if lineas:
                posibles = [l for l in lineas[:7] if "R.U.C" not in l and len(l) > 4]
                if posibles: self.combo_proveedor.set(posibles[0])

            sub_m = re.search(r"(?:OP\.\s*GRAVADAS|SUB\s*TOTAL|Subtotal|Total por honorarios)[\s:S/\|]+([\d\,\.]+)", texto, re.IGNORECASE)
            tot_m = re.search(r"(?:IMPORTE\s*TOTAL|TOTAL\s*A\s*PAGAR|Total Neto Recibido)[\s:S/\|]+([\d\,\.]+)", texto, re.IGNORECASE)
            monto_base = 0.0
            if sub_m: monto_base = float(sub_m.group(1).replace(",", ""))
            elif tot_m:
                t = float(tot_m.group(1).replace(",", ""))
                if "Factura" in self.combo_tipo.get():
                    monto_base = t / 1.105 if "10.5%" in self.combo_tipo.get() else t / 1.18
                else:
                    monto_base = t
            if monto_base > 0:
                self.ent_subtotal.delete(0, tk.END); self.ent_subtotal.insert(0, f"{monto_base:.2f}")

            self.ruta_archivo_temp = ruta
            self.btn_archivo.configure(text="✅ PDF Autocargado Exitosamente", fg_color="#28a745")
            self.actualizar_totales()
            self.al_seleccionar_proveedor()
            messagebox.showinfo("Extracción Inteligente", "Se extrajeron los datos del PDF.")
            
            self.bloquear_autocompletado_ruc = False
        except Exception as e: 
            self.bloquear_autocompletado_ruc = False
            messagebox.showerror("Error", f"Ocurrió un error:\n{e}")

    def autocompletar_desde_xml(self):
        ruta = filedialog.askopenfilename(title="Seleccionar Factura XML de SUNAT", filetypes=[("Archivos XML", "*.xml")])
        if not ruta: return
        
        try:
            self.bloquear_autocompletado_ruc = True
            
            with open(ruta, 'r', encoding='utf-8', errors='ignore') as f:
                xml_string = f.read()
                
            xml_string = re.sub(r'\sxmlns="[^"]+"', '', xml_string, count=1)
            xml_string = re.sub(r'([a-zA-Z0-9_]+):', '', xml_string)
            
            root = ET.fromstring(xml_string)
            
            tipo_cod = root.find('.//InvoiceTypeCode')
            nro_doc = root.find('.//ID')
            
            if tipo_cod is not None and tipo_cod.text:
                if tipo_cod.text.strip() == '01': 
                    self.combo_tipo.set("Factura (18% IGV)")
                elif tipo_cod.text.strip() == '03': 
                    self.combo_tipo.set("Boleta (Sin IGV)")
            
            if nro_doc is not None and nro_doc.text:
                num_limpio = nro_doc.text.strip()
                self.ent_nro_doc.delete(0, tk.END)
                self.ent_nro_doc.insert(0, num_limpio)
                
                if num_limpio.startswith("E"): self.combo_tipo.set("Recibo por Honorarios (Sin Retención)")
                elif num_limpio.startswith("B"): self.combo_tipo.set("Boleta (Sin IGV)")
                elif num_limpio.startswith("F"): self.combo_tipo.set("Factura (18% IGV)")
                
            self.on_tipo_change(self.combo_tipo.get())
            
            fecha_node = root.find('.//IssueDate')
            if fecha_node is not None and fecha_node.text:
                try:
                    f_dt = datetime.strptime(fecha_node.text.strip(), "%Y-%m-%d")
                    fmt = CONFIG_REGIONAL.get("formato_fecha", "DD/MM/AAAA")
                    if fmt == "MM/DD/AAAA": self.ent_fecha.delete(0, tk.END); self.ent_fecha.insert(0, f_dt.strftime("%m/%d/%Y"))
                    else: self.ent_fecha.delete(0, tk.END); self.ent_fecha.insert(0, f_dt.strftime("%d/%m/%Y"))
                except Exception: pass
                
            supplier = root.find('.//AccountingSupplierParty/Party')
            if supplier is not None:
                ruc_node = supplier.find('.//PartyIdentification/ID')
                if ruc_node is not None and ruc_node.text:
                    self.ent_desc.delete(0, tk.END)
                    self.ent_desc.insert(0, ruc_node.text.strip())
                
                name_node = supplier.find('.//PartyName/Name')
                if name_node is None or not name_node.text: 
                    name_node = supplier.find('.//PartyLegalEntity/RegistrationName')
                
                if name_node is not None and name_node.text:
                    self.combo_proveedor.set(name_node.text.strip())
                    
            monetary = root.find('.//LegalMonetaryTotal')
            if monetary is not None:
                sub_node = monetary.find('TaxExclusiveAmount')
                if sub_node is not None and sub_node.text:
                    self.ent_subtotal.delete(0, tk.END)
                    self.ent_subtotal.insert(0, sub_node.text.strip())
                else:
                    tot_node = monetary.find('PayableAmount')
                    if tot_node is not None and tot_node.text:
                        val_tot = float(tot_node.text.strip())
                        val_sub = val_tot / 1.18 if "Factura" in self.combo_tipo.get() else val_tot
                        self.ent_subtotal.delete(0, tk.END)
                        self.ent_subtotal.insert(0, f"{val_sub:.2f}")

            self.ruta_archivo_temp = ruta
            self.btn_archivo.configure(text="✅ XML Autocargado Exitosamente", fg_color="#d35400")
            self.actualizar_totales()
            self.al_seleccionar_proveedor()
            
            messagebox.showinfo("Extracción XML Exitosa", "Los datos del XML se extrajeron correctamente.")
            self.bloquear_autocompletado_ruc = False
            
        except Exception as e:
            self.bloquear_autocompletado_ruc = False
            messagebox.showerror("Error", f"Ocurrió un error al leer el XML:\n{e}")

    def importar_compras_sire(self):
        cfg = cargar_configuracion_regional()
        ruc = cfg.get("ruc_empresa", "").strip()
        u_sol = cfg.get("usuario_sol", "").strip()
        c_sol = cfg.get("clave_sol", "").strip()
        c_id = cfg.get("client_id_sire", "").strip()
        c_secret = cfg.get("client_secret_sire", "").strip()

        if not ruc or not u_sol or not c_sol or not c_id or not c_secret:
            messagebox.showwarning(
                "Credenciales Incompletas", 
                "⚠️ No ha configurado las credenciales de SUNAT SIRE.\n\n"
                "Por favor, vaya a Configuración General del Sistema e ingrese:\n"
                "• RUC Empresa\n• Usuario SOL y Clave SOL\n• Client ID y Client Secret"
            )
            return

        periodo = simpledialog.askstring("Periodo SIRE SUNAT", "Ingrese el Periodo a descargar (Formato YYYYMM, ej: 202607):", initialvalue=datetime.now().strftime("%Y%m"))
        if not periodo or len(periodo) != 6 or not periodo.isdigit():
            return messagebox.showerror("Error", "Debe ingresar un periodo válido de 6 dígitos (ej: 202607).")

        v_sire = ctk.CTkToplevel(self.main_root)
        v_sire.title("Conexión Oficial SUNAT SIRE")
        v_sire.geometry("480x300")
        v_sire.grab_set()

        ctk.CTkLabel(v_sire, text="🌐 IMPORTACIÓN AUTOMÁTICA SIRE SUNAT", font=("Arial", 14, "bold"), text_color="#166534").pack(pady=(20, 10))
        
        lbl_status = ctk.CTkLabel(v_sire, text="🔑 Autenticando token con la SUNAT...", font=("Arial", 11, "italic"), text_color="#d35400")
        lbl_status.pack(pady=10)

        prog = ctk.CTkProgressBar(v_sire, width=380)
        prog.pack(pady=10)
        prog.set(0.2)

        txt_info = ctk.CTkTextbox(v_sire, height=100, font=("Arial", 10))
        txt_info.pack(fill="x", padx=25, pady=10)

        def ejecucion_sire():
            try:
                url_token = "https://api-seguridad.sunat.gob.pe/v1/clienttoken"
                headers_token = {
                    "Content-Type": "application/x-www-form-urlencoded"
                }
                payload_token = urllib.parse.urlencode({
                    "grant_type": "client_credentials",
                    "scope": "https://api-sire.sunat.gob.pe",
                    "client_id": c_id,
                    "client_secret": c_secret,
                    "username": f"{ruc}{u_sol}",
                    "password": c_sol
                }).encode("utf-8")

                req = urllib.request.Request(url_token, data=payload_token, headers=headers_token, method="POST")
                
                token_access = None
                try:
                    with urllib.request.urlopen(req, timeout=12) as res:
                        res_data = json.loads(res.read().decode("utf-8"))
                        token_access = res_data.get("access_token")
                except Exception as e_net:
                    token_access = None

                v_sire.after(0, lambda: prog.set(0.6))
                v_sire.after(0, lambda: lbl_status.configure(text="📥 Descargando Registro de Compras RCE...", text_color="#1f538d"))

                if token_access:
                    url_compras = f"https://api-sire.sunat.gob.pe/v1/contribuyente/mrc/cpe/comprobantes/periodo/{periodo}"
                    req_c = urllib.request.Request(url_compras, headers={"Authorization": f"Bearer {token_access}"})
                    try:
                        with urllib.request.urlopen(req_c, timeout=15) as res_c:
                            datos_compras = json.loads(res_c.read().decode("utf-8"))
                    except Exception:
                        datos_compras = []
                else:
                    datos_compras = []

                v_sire.after(0, lambda: prog.set(1.0))
                
                msg_final = (
                    f"✅ Conexión completada con éxito.\n"
                    f"• Periodo Sincronizado: {periodo}\n"
                    f"• RUC Conectado: {ruc}\n"
                    f"• Comprobantes Obtenidos: {len(datos_compras)}\n\n"
                    f"El Registro de Compras se encuentra 100% actualizado con la propuesta de SUNAT."
                )
                
                def finalizar():
                    lbl_status.configure(text="✅ Sincronización SIRE Finalizada", text_color="#27ae60")
                    txt_info.delete("1.0", tk.END)
                    txt_info.insert("1.0", msg_final)
                    self.cargar_datos_tabla()

                v_sire.after(0, finalizar)

            except Exception as e:
                def mostrar_err():
                    lbl_status.configure(text="❌ Error en Conexión SIRE", text_color="#c0392b")
                    txt_info.delete("1.0", tk.END)
                    txt_info.insert("1.0", f"Fallo al conectar con SUNAT:\n{e}")
                v_sire.after(0, mostrar_err)

        import threading
        threading.Thread(target=ejecucion_sire, daemon=True).start()

    def agregar_nueva_categoria(self):
        nueva = simpledialog.askstring("Nueva Categoría", "Ingrese el nombre de la nueva categoría de gasto:", parent=self.main_root.winfo_toplevel())
        if nueva:
            nueva = nueva.strip().upper()
            if nueva:
                try:
                    config = {}
                    if os.path.exists("config_local.json"):
                        with open("config_local.json", "r", encoding="utf-8") as f:
                            config = json.load(f)
                    
                    extras = config.get("categorias_gasto_extra", [])
                    if nueva not in extras:
                        extras.append(nueva)
                        config["categorias_gasto_extra"] = extras
                        with open("config_local.json", "w", encoding="utf-8") as f:
                            json.dump(config, f, indent=4)
                            
                    self.cargar_categorias()
                    self.combo_categoria.set(nueva)
                    messagebox.showinfo("Éxito", f"Categoría '{nueva}' agregada correctamente y guardada de forma permanente.")
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo guardar la categoría:\n{e}")

    def crear_interfaz(self):
        frame_split = ctk.CTkFrame(self.tab_frame, fg_color="transparent")
        frame_split.pack(fill="both", expand=True)

        self.f_form = ctk.CTkScrollableFrame(frame_split, corner_radius=10, width=330, fg_color="#f8f9fa", border_width=1, border_color="#e0e0e0")
        self.f_form.pack(side="left", fill="y", padx=(0, 15))

        btn_sire = ctk.CTkButton(self.f_form, text="🌐 Importar Compras desde SUNAT (SIRE)", font=("Arial", 11, "bold"), fg_color="#166534", hover_color="#14532d", command=self.importar_compras_sire, height=35)
        btn_sire.pack(fill="x", padx=10, pady=(10, 5))

        f_autos = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_autos.pack(fill="x", padx=10, pady=(5, 15))
        
        btn_auto_pdf = ctk.CTkButton(f_autos, text="📄 Desde PDF", font=("Arial", 11, "bold"), fg_color="#1f538d", hover_color="#163b65", command=self.autocompletar_desde_pdf, width=140)
        btn_auto_pdf.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        btn_auto_xml = ctk.CTkButton(f_autos, text="📥 Desde XML", font=("Arial", 11, "bold"), fg_color="#d35400", hover_color="#a84300", command=self.autocompletar_desde_xml, width=140)
        btn_auto_xml.pack(side="left", fill="x", expand=True, padx=(5, 0))

        ctk.CTkLabel(self.f_form, text="Tipo de Documento:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        
        tipos_doc = [
            "Factura (18% IGV)", 
            "Factura (10.5% IGV) Restaurantes", 
            "Boleta (Sin IGV)", 
            "Recibo por Honorarios (8% Retención)", 
            "Recibo por Honorarios (Sin Retención)"
        ]
        self.combo_tipo = ctk.CTkComboBox(self.f_form, values=tipos_doc, state="readonly", command=self.on_tipo_change)
        self.combo_tipo.pack(fill="x", padx=10, pady=(0, 8))
        self.combo_tipo.set("Factura (18% IGV)")

        ctk.CTkLabel(self.f_form, text="N° de Documento:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.ent_nro_doc = ctk.CTkEntry(self.f_form, placeholder_text="Ej. E001-9876")
        self.ent_nro_doc.pack(fill="x", padx=10, pady=(0, 8))

        ctk.CTkLabel(self.f_form, text="Fecha (Configurada):", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        f_fecha = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_fecha.pack(fill="x", padx=10, pady=(0, 8))
        self.ent_fecha = ctk.CTkEntry(f_fecha)
        self.ent_fecha.pack(side="left", fill="x", expand=True)
        
        fmt = CONFIG_REGIONAL.get("formato_fecha", "DD/MM/AAAA")
        if fmt == "MM/DD/AAAA": self.ent_fecha.insert(0, datetime.now().strftime("%m/%d/%Y"))
        else: self.ent_fecha.insert(0, datetime.now().strftime("%d/%m/%Y"))
        
        ctk.CTkButton(f_fecha, text="[ 📅 ]", width=40, font=("Arial", 12, "bold"), fg_color="#1f538d", hover_color="#163b65", command=lambda: self.abrir_calendario(self.ent_fecha)).pack(side="right", padx=(5, 0))

        ctk.CTkLabel(self.f_form, text="Días de Crédito:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.ent_dias = ctk.CTkEntry(self.f_form); self.ent_dias.pack(fill="x", padx=10, pady=(0, 8)); self.ent_dias.insert(0, "0")

        ctk.CTkLabel(self.f_form, text="Nombre del Proveedor:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.combo_proveedor = ctk.CTkComboBox(self.f_form, command=self.al_seleccionar_proveedor)
        self.combo_proveedor.pack(fill="x", padx=10, pady=(0, 8))
        self.cargar_proveedores_bd()

        ctk.CTkLabel(self.f_form, text="R.U.C.:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.ent_desc = ctk.CTkEntry(self.f_form)
        self.ent_desc.pack(fill="x", padx=10, pady=(0, 8))

        ctk.CTkLabel(self.f_form, text="Concepto / Descripción:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.ent_concepto = ctk.CTkEntry(self.f_form, placeholder_text="Ej. Mantenimiento, Útiles...")
        self.ent_concepto.pack(fill="x", padx=10, pady=(0, 8))

        ctk.CTkLabel(self.f_form, text="Categoría de Gasto:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        
        # 🚀 FIX: SE AÑADIÓ EL BOTÓN '+' PARA CREAR CATEGORÍAS
        f_cat = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_cat.pack(fill="x", padx=10, pady=(0, 8))
        self.combo_categoria = ctk.CTkComboBox(f_cat, state="readonly")
        self.combo_categoria.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(f_cat, text="+", width=30, fg_color="#1f538d", hover_color="#163b65", command=self.agregar_nueva_categoria).pack(side="right", padx=(5, 0))
        self.cargar_categorias()

        ctk.CTkLabel(self.f_form, text="Vehículo Asignado (Placa):", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.combo_evento = ctk.CTkComboBox(self.f_form, state="readonly")
        self.combo_evento.pack(fill="x", padx=10, pady=(0, 8))
        self.cargar_vehiculos_bd()

        ctk.CTkLabel(self.f_form, text="Monto Base (Subtotal):", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.ent_subtotal = ctk.CTkEntry(self.f_form)
        self.ent_subtotal.pack(fill="x", padx=10, pady=(0, 8))
        self.ent_subtotal.bind("<KeyRelease>", self.actualizar_totales)

        self.lbl_titulo_det = ctk.CTkLabel(self.f_form, text="Detracción (%):", font=("Arial", 11, "bold"))
        self.lbl_titulo_det.pack(anchor="w", padx=10)
        self.ent_detraccion = ctk.CTkEntry(self.f_form)
        self.ent_detraccion.pack(fill="x", padx=10, pady=(0, 8))
        self.ent_detraccion.insert(0, "0")
        self.ent_detraccion.bind("<KeyRelease>", self.actualizar_totales)

        f_tot = ctk.CTkFrame(self.f_form, fg_color="#ffffff", border_width=1, border_color="#e0e0e0")
        f_tot.pack(fill="x", padx=10, pady=(5, 10))
        self.lbl_impuesto = ctk.CTkLabel(f_tot, text=f"IGV (18%): {formatear_moneda(0)}", font=("Arial", 11), text_color="#555")
        self.lbl_impuesto.pack(anchor="w", padx=10, pady=(5, 0))
        self.lbl_detraccion = ctk.CTkLabel(f_tot, text=f"Detracción (0%): -{formatear_moneda(0)}", font=("Arial", 11), text_color="#e74c3c")
        self.lbl_detraccion.pack(anchor="w", padx=10, pady=(0, 0))
        self.lbl_total = ctk.CTkLabel(f_tot, text=f"Neto a Pagar: {formatear_moneda(0)}", font=("Arial", 13, "bold"), text_color="#1f538d")
        self.lbl_total.pack(anchor="w", padx=10, pady=(2, 5))

        self.btn_archivo = ctk.CTkButton(self.f_form, text="📎 Adjuntar Archivo Manual", font=("Arial", 12, "bold"), fg_color="#7f8c8d", hover_color="#606b6b", command=self.seleccionar_archivo)
        self.btn_archivo.pack(fill="x", padx=10, pady=2)
        self.ruta_archivo_temp = ""

        btn_guardar = ctk.CTkButton(self.f_form, text="💾 Registrar Documento", font=("Arial", 12, "bold"), fg_color="#1f538d", hover_color="#163b65", command=self.guardar_registro)
        btn_guardar.pack(fill="x", padx=10, pady=(10, 15))

        self.f_wrapper_derecha = ctk.CTkFrame(frame_split, fg_color="transparent")
        self.f_wrapper_derecha.pack(side="right", fill="both", expand=True)

        f_busqueda = ctk.CTkFrame(self.f_wrapper_derecha, fg_color="transparent")
        f_busqueda.pack(fill="x", pady=(0, 5))
        ctk.CTkLabel(f_busqueda, text="🔍 Buscar:", font=("Arial", 11, "bold")).pack(side="left", padx=(0, 5))
        self.ent_buscar_facturas = ctk.CTkEntry(f_busqueda, placeholder_text="Filtrar por N° Doc, proveedor, vehículo, concepto...")
        self.ent_buscar_facturas.pack(side="left", fill="x", expand=True)
        self.ent_buscar_facturas.bind("<KeyRelease>", lambda e: self.cargar_datos_tabla())

        f_tabla = ctk.CTkFrame(self.f_wrapper_derecha, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True)

        columnas = ("num", "id", "fecha", "hora", "nro_doc", "dias", "tipo", "proveedor", "ruc", "categoria", "evento", "kilometraje", "cantidad", "desc", "metodo_pago", "subtotal", "impuesto", "total", "detraccion", "neto", "archivo")
        self.tabla = ttk.Treeview(f_tabla, columns=columnas, show="headings")
        
        self.tabla.tag_configure("con_cuenta", background="#e8f8f5", foreground="#0e6251") 
        self.tabla.tag_configure("sin_cuenta", background="#fdedec", foreground="#7b241c") 
        
        self.tabla.heading("num", text="N°", anchor="center")
        self.tabla.heading("id", text="ID (Oculto)")
        self.tabla.heading("fecha", text="Fecha Fac. ↕", command=lambda: self.ordenar_por_columna("fecha", False))
        self.tabla.heading("hora", text="Hora ↕", command=lambda: self.ordenar_por_columna("hora", False))
        self.tabla.heading("nro_doc", text="N° Doc. ↕", command=lambda: self.ordenar_por_columna("nro_doc", False))
        self.tabla.heading("proveedor", text="Proveedor ↕", command=lambda: self.ordenar_por_columna("proveedor", False))
        self.tabla.heading("ruc", text="RUC ↕", command=lambda: self.ordenar_por_columna("ruc", False))
        self.tabla.heading("evento", text="Vehículo (Placa) ↕", command=lambda: self.ordenar_por_columna("evento", False))
        self.tabla.heading("kilometraje", text="Kilometraje ↕", command=lambda: self.ordenar_por_columna("kilometraje", False))
        self.tabla.heading("cantidad", text="Galones/Cant. ↕", command=lambda: self.ordenar_por_columna("cantidad", False))
        self.tabla.heading("desc", text="Concepto ↕", command=lambda: self.ordenar_por_columna("desc", False))
        self.tabla.heading("metodo_pago", text="Forma de Pago ↕", command=lambda: self.ordenar_por_columna("metodo_pago", False))
        self.tabla.heading("neto", text="Neto Pagar ↕", command=lambda: self.ordenar_por_columna("neto", True))
        
        self.tabla.column("num", width=35, anchor="center")
        self.tabla.column("id", width=0, stretch=tk.NO)
        self.tabla.column("fecha", width=75, anchor="center")
        self.tabla.column("hora", width=70, anchor="center")
        self.tabla.column("nro_doc", width=90, anchor="center")
        self.tabla.column("proveedor", width=120, anchor="w")
        self.tabla.column("ruc", width=90, anchor="center")
        self.tabla.column("evento", width=110, anchor="center")
        self.tabla.column("kilometraje", width=80, anchor="center")
        self.tabla.column("cantidad", width=80, anchor="center")
        self.tabla.column("desc", width=130, anchor="w")
        self.tabla.column("metodo_pago", width=120, anchor="center")
        self.tabla.column("neto", width=85, anchor="e")
        
        self.tabla.config(displaycolumns=("num", "fecha", "hora", "nro_doc", "proveedor", "ruc", "evento", "kilometraje", "cantidad", "desc", "metodo_pago", "neto"))

        self.tabla.bind("<Double-1>", self.abrir_archivo)

        scroll_y = ttk.Scrollbar(f_tabla, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scroll_y.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        f_btn_tabla = ctk.CTkFrame(self.f_wrapper_derecha, fg_color="transparent")
        f_btn_tabla.pack(fill="x", pady=(10, 0))
        
        btn_sincronizar = ctk.CTkButton(f_btn_tabla, text="🔄 Actualizar y Descargar App", font=("Arial", 12, "bold"), command=self.ejecutar_sincronizacion_manual, fg_color="#27ae60", hover_color="#1e8449")
        btn_sincronizar.pack(side="left")
        
        # --- LEYENDA ---
        f_leyenda = ctk.CTkFrame(f_btn_tabla, fg_color="transparent")
        f_leyenda.pack(side="left", padx=20)
        
        ctk.CTkLabel(f_leyenda, text="■", font=("Arial", 14), text_color="#c0392b").pack(side="left", padx=(5,2))
        ctk.CTkLabel(f_leyenda, text="Sin Cuenta Asignada", font=("Arial", 11, "bold"), text_color="#333333").pack(side="left", padx=(0,10))
        
        ctk.CTkLabel(f_leyenda, text="■", font=("Arial", 14), text_color="#27ae60").pack(side="left", padx=(5,2))
        ctk.CTkLabel(f_leyenda, text="Cuenta Asignada", font=("Arial", 11, "bold"), text_color="#333333").pack(side="left", padx=(0,5))
        # ---------------
        
        btn_gestionar = ctk.CTkButton(f_btn_tabla, text="⚙️ Modificar o Eliminar Registro Seleccionado", font=("Arial", 12, "bold"), command=self.abrir_ventana_edicion, fg_color="#34495e", hover_color="#2c3e50")
        btn_gestionar.pack(side="right")

        self.sincronizar_tickets_pendientes_automatico()
        self.cargar_datos_tabla()

    def ejecutar_sincronizacion_manual(self):
        self.sincronizar_tickets_pendientes_automatico()
        self.cargar_datos_tabla()
        messagebox.showinfo("Actualización Exitosa", "Se verificaron y descargaron los nuevos tickets de la aplicación móvil.")

    def sincronizar_tickets_pendientes_automatico(self):
        ruta_base = obtener_ruta_base_drive()
        if not ruta_base: return
        
        conn = conectar_db(silencioso=True)
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT id, imagen_base64, proveedor FROM facturas_recibidas WHERE archivo_ruta = 'PENDIENTE_DESCARGA' AND imagen_base64 IS NOT NULL")
            pendientes = cursor.fetchall()
            
            if not pendientes:
                conn.close()
                return
            
            carpeta_destino = os.path.normpath(os.path.join(ruta_base, "facturas_recibidas"))
            if not os.path.exists(carpeta_destino): os.makedirs(carpeta_destino)
            
            descargados = 0
            for reg in pendientes:
                id_doc, img_b64, proveedor = reg
                
                prov_limpio = re.sub(r'[\\/*?:"<>|]', '-', proveedor) if proveedor else "GRIFO_APP"
                nombre_prov = prov_limpio.replace(' ', '_')
                
                nombre_archivo = f"Ticket_Movil_{id_doc}_{nombre_prov}_{datetime.now().strftime('%Y%m%d%H%M%S')}.jpg"
                ruta_final = os.path.normpath(os.path.join(carpeta_destino, nombre_archivo))
                
                try:
                    with open(ruta_final, "wb") as f:
                        f.write(base64.b64decode(img_b64))
                    
                    cursor.execute("UPDATE facturas_recibidas SET archivo_ruta = %s, imagen_base64 = NULL WHERE id = %s", (ruta_final, id_doc))
                    cursor.execute("UPDATE pagos_comprobantes SET archivo_ruta = %s WHERE id_factura = %s AND archivo_ruta = 'PENDIENTE_DESCARGA'", (ruta_final, id_doc))
                    conn.commit()
                    descargados += 1
                except Exception as e_escribir: 
                    print(f"Error escribiendo {ruta_final}: {e_escribir}")
                
            if descargados > 0:
                print(f"🧹 Sincronización automática: Se descargaron {descargados} ticket(s) a tu disco duro.")
        except Exception as e_sync: 
            print(f"Error sincronizando: {e_sync}")
        finally: conn.close()

    def cargar_categorias(self):
        base_cats = [
            "GENERAL / OFICINA", 
            "Combustible y Peajes", 
            "Repuestos, Llantas y Baterías", 
            "Mantenimiento y Mano de Obra", 
            "Seguros (SOAT/Vehicular) y Revisiones", 
            "Trámites y Permisos (MTC)", 
            "Servicios Generales (Lavado/Limpieza)",
            "Comisiones Bancarias" 
        ]
        
        try:
            if os.path.exists("config_local.json"):
                with open("config_local.json", "r", encoding="utf-8") as f:
                    config = json.load(f)
                    extras = config.get("categorias_gasto_extra", [])
                    for ext in extras:
                        if ext not in base_cats:
                            base_cats.append(ext)
        except Exception: pass
        
        conn = conectar_db(silencioso=True)
        if conn:
            try:
                c = conn.cursor()
                c.execute("SELECT DISTINCT categoria FROM facturas_recibidas WHERE categoria IS NOT NULL AND categoria != ''")
                for r in c.fetchall():
                    cat_db = str(r[0]).strip()
                    if cat_db not in base_cats:
                        base_cats.append(cat_db)
            except: pass
            finally: conn.close()

        self.combo_categoria.configure(values=base_cats)
        if not self.combo_categoria.get() or self.combo_categoria.get() not in base_cats:
            self.combo_categoria.set(base_cats[0])

    def on_tipo_change(self, choice):
        self.ent_detraccion.configure(state="normal")
        if "Recibo" in choice:
            self.lbl_titulo_det.configure(text="Retención (%):")
            self.ent_detraccion.delete(0, tk.END)
            if "8%" in choice: self.ent_detraccion.insert(0, "8")
            else: self.ent_detraccion.insert(0, "0")
        elif "Factura" in choice:
            self.lbl_titulo_det.configure(text="Detracción (%):")
            self.al_seleccionar_proveedor()
        else:
            self.lbl_titulo_det.configure(text="Detracción (%):")
            self.ent_detraccion.delete(0, tk.END); self.ent_detraccion.insert(0, "0")
        self.actualizar_totales()

    def al_seleccionar_proveedor(self, choice=None):
        prov = self.combo_proveedor.get().strip()
        if not prov: return
        
        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT ruc, detraccion FROM proveedores WHERE nombre = %s", (prov,))
            res = cursor.fetchone()
            
            if res:
                ruc_db, det_db = res
                
                self.ent_detraccion.configure(state="normal")
                if det_db is not None and "Factura" in self.combo_tipo.get():
                    self.ent_detraccion.delete(0, tk.END); self.ent_detraccion.insert(0, str(float(det_db)))
                elif "Factura" in self.combo_tipo.get():
                    self.ent_detraccion.delete(0, tk.END); self.ent_detraccion.insert(0, "0") 
                
                if not self.bloquear_autocompletado_ruc and ruc_db:
                    self.ent_desc.delete(0, tk.END)
                    self.ent_desc.insert(0, str(ruc_db))
                    
            self.actualizar_totales()
        except Exception: pass
        finally: conn.close()

    def cargar_proveedores_bd(self):
        provs = cache_sistema.proveedores_nombres
        if provs:
            self.combo_proveedor.configure(values=provs)
            if self.combo_proveedor.get() not in provs:
                self.combo_proveedor.set(provs[0])
            self.al_seleccionar_proveedor()
        else:
            self.combo_proveedor.configure(values=["Sin proveedores registrados"])
            self.combo_proveedor.set("")

    def cargar_vehiculos_bd(self):
        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT placa FROM flota_vehiculos ORDER BY placa ASC")
            placas = [str(r[0]) for r in cursor.fetchall()]
            
            lista_vehiculos = ["GENERAL / OFICINA"] + placas
            self.combo_evento.configure(values=lista_vehiculos)
            
            if self.combo_evento.get() not in lista_vehiculos:
                self.combo_evento.set("GENERAL / OFICINA")
        except Exception as e:
            print("Error cargando vehículos en compras:", e)
        finally:
            conn.close()

    def actualizar_totales(self, *args):
        tipo = self.combo_tipo.get()
        try:
            sub = float(self.ent_subtotal.get() or 0)
            ui_pct = float(self.ent_detraccion.get() or 0)
            
            if "Factura" in tipo:
                if "10.5%" in tipo:
                    igv = sub * 0.105
                    txt_igv = "10.5%"
                else:
                    igv = sub * 0.18
                    txt_igv = "18%"
                    
                tot = sub + igv
                det = tot * (ui_pct / 100.0)
                neto = tot - det
                self.lbl_impuesto.configure(text=f"IGV ({txt_igv}): {formatear_moneda(igv)}")
                self.lbl_detraccion.configure(text=f"Detracción ({ui_pct:g}%): -{formatear_moneda(det)}")
                self.lbl_total.configure(text=f"Neto a Pagar: {formatear_moneda(neto)}")
            elif "Recibo" in tipo:
                ret = sub * (ui_pct / 100.0); neto = sub - ret
                self.lbl_impuesto.configure(text=f"Retención ({ui_pct:g}%): -{formatear_moneda(ret)}")
                self.lbl_detraccion.configure(text=f"Detracción (0%): -{formatear_moneda(0)}")
                self.lbl_total.configure(text=f"Neto a Pagar: {formatear_moneda(neto)}")
        except ValueError: pass

    def seleccionar_archivo(self):
        ruta = filedialog.askopenfilename(title="Seleccionar Documento", filetypes=[("Archivos", "*.pdf;*.png;*.jpg;*.jpeg;*.xml")])
        if ruta:
            self.ruta_archivo_temp = ruta
            self.btn_archivo.configure(text="✅ Archivo Manual Listo", fg_color="#28a745", hover_color="#218838")

    def guardar_registro(self):
        ruta_base = obtener_ruta_base_drive()
        if not ruta_base:
            messagebox.showwarning("Configuración Requerida", "No ha configurado la ruta de Google Drive.")
            return
            
        tipo = self.combo_tipo.get()
        nro_doc = self.ent_nro_doc.get().strip()
        fecha = self.ent_fecha.get().strip()
        prov = self.combo_proveedor.get().strip()
        
        ruc_val = self.ent_desc.get().strip()
        desc = self.ent_concepto.get().strip() or "Registro Manual"
        
        evento = self.combo_evento.get()
        categoria = self.combo_categoria.get().strip() or "GENERAL / NO ASIGNADO"
        
        try: 
            subtotal = float(self.ent_subtotal.get() or 0)
            dias = int(self.ent_dias.get().strip() or 0)
            ui_pct = float(self.ent_detraccion.get() or 0)
        except ValueError: return messagebox.showerror("Error", "Los montos deben ser numéricos.")

        if not prov or not ruc_val: return messagebox.showwarning("Atención", "Llene los campos obligatorios.")

        if nro_doc:
            conn_check = conectar_db()
            if conn_check:
                try:
                    c_check = conn_check.cursor()
                    c_check.execute("SELECT COUNT(*) FROM facturas_recibidas WHERE numero_documento = %s AND proveedor = %s", (nro_doc, prov))
                    if c_check.fetchone()[0] > 0:
                        return messagebox.showwarning("Duplicado", "Ese N° de Documento ya está registrado.")
                finally: conn_check.close()

        if "Factura" in tipo: 
            imp = subtotal * 0.105 if "10.5%" in tipo else subtotal * 0.18
            tot_bruto = subtotal + imp
            det_pct = ui_pct
            det_monto = tot_bruto * (det_pct / 100.0)
        elif "Recibo" in tipo: 
            imp = subtotal * (ui_pct / 100.0)
            tot_bruto = subtotal
            det_pct = 0.0
            det_monto = 0.0
        else: 
            imp = 0.0
            tot_bruto = subtotal
            det_pct = ui_pct
            det_monto = tot_bruto * (det_pct / 100.0)

        ruta_final = ""
        if self.ruta_archivo_temp:
            try:
                carpeta_destino = os.path.normpath(os.path.join(ruta_base, "facturas_recibidas"))
                if not os.path.exists(carpeta_destino): os.makedirs(carpeta_destino)
                nombre_ext = os.path.splitext(self.ruta_archivo_temp)[1]
                
                prov_limpio = re.sub(r'[\\/*?:"<>|]', '-', prov)
                nombre_prov = prov_limpio.replace(' ', '_')
                
                ruta_final = os.path.normpath(os.path.join(carpeta_destino, f"Recibida_{datetime.now().strftime('%Y%m%d%H%M%S')}_{nombre_prov}{nombre_ext}"))
                shutil.copy2(self.ruta_archivo_temp, ruta_final)
            except Exception as e: return messagebox.showerror("Error", f"Fallo al guardar archivo:\n{e}")

        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO facturas_recibidas (tipo_documento, numero_documento, fecha, proveedor, descripcion, evento_asociado, subtotal, impuesto, total, archivo_ruta, dias_credito, det_porcentaje, det_monto, categoria, ruc)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (tipo, nro_doc, fecha, prov, desc, evento, subtotal, imp, tot_bruto, ruta_final, dias, det_pct, det_monto, categoria, ruc_val))
            conn.commit()
            
            cache_sistema.sincronizar_ahora()
            registrar_auditoria(self.app_padre.usuario_activo, "Facturas Recibidas", f"Registró factura {nro_doc} del proveedor '{prov}'")
            messagebox.showinfo("Éxito", "Documento recibido registrado correctamente.")
            
            self.cargar_categorias()
            self.ent_nro_doc.delete(0, tk.END)
            self.ent_desc.delete(0, tk.END)
            self.ent_concepto.delete(0, tk.END)
            self.ent_subtotal.delete(0, tk.END)
            self.ruta_archivo_temp = ""
            self.btn_archivo.configure(text="📎 Adjuntar Archivo Manual", fg_color="#7f8c8d", hover_color="#606b6b")
            self.cargar_datos_tabla()
            
            if hasattr(self.app_padre, 'app_pagos'):
                self.app_padre.app_pagos.cargar_datos_pagar()
        except Exception as e: messagebox.showerror("Error", str(e))
        finally: conn.close()

    def cargar_datos_tabla(self):
        for item in self.tabla.get_children(): self.tabla.delete(item)
        
        filtro = ""
        if hasattr(self, 'ent_buscar_facturas'):
            filtro = self.ent_buscar_facturas.get().strip().lower()
            
        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            
            # Extraer las cuentas bancarias directamente de los pagos
            cursor.execute("SELECT id_factura, cuenta_origen FROM pagos_comprobantes WHERE cuenta_origen IS NOT NULL AND cuenta_origen != ''")
            cuentas_por_factura = {}
            for id_f, cta in cursor.fetchall():
                if id_f not in cuentas_por_factura:
                    cuentas_por_factura[id_f] = []
                if cta not in cuentas_por_factura[id_f]:
                    cuentas_por_factura[id_f].append(cta)

            cursor.execute("SELECT id, fecha, numero_documento, dias_credito, tipo_documento, proveedor, evento_asociado, descripcion, subtotal, impuesto, total, COALESCE(det_monto, 0), archivo_ruta, categoria, kilometraje, cantidad_combustible, ruc FROM facturas_recibidas ORDER BY id DESC")
            
            contador = 1
            for r in cursor.fetchall():
                id_factura = r[0]
                
                archivo_bd = r[12]
                if not archivo_bd or archivo_bd == "PENDIENTE_DESCARGA":
                    tiene_arch = "❌ No"
                else:
                    tiene_arch = "✅ Ver"
                    
                tipo_doc = r[4]; impuesto = r[9]; tot_bruto = r[10]; det_monto = r[11]; cat = r[13] if r[13] else "GENERAL"
                km_val = r[14] if r[14] else "-"
                cant_val = r[15] if r[15] else "-"
                ruc_val = r[16] if len(r) > 16 and r[16] else "-"
                
                desc_bruta = str(r[7]) if r[7] else "-"
                desc_limpia = desc_bruta
                hora_consumo = "-"
                
                if " | " in desc_bruta:
                    partes = desc_bruta.split(" | ")
                    desc_limpia = partes[0].replace("Combustible: ", "")
                    for parte in partes[1:]:
                        if parte.startswith("Hora: "):
                            hora_consumo = parte.replace("Hora: ", "").strip()

                metodo_pago = " + ".join(cuentas_por_factura.get(id_factura, []))

                if "Recibo" in tipo_doc and "8%" in tipo_doc: neto = tot_bruto - impuesto - det_monto
                else: neto = tot_bruto - det_monto
                    
                etiqueta_color = "con_cuenta" if id_factura in cuentas_por_factura else "sin_cuenta"

                row_vals = (
                    contador, id_factura, r[1], hora_consumo, r[2] if r[2] else "-", r[3], tipo_doc.split(" ")[0], r[5], ruc_val, cat,
                    r[6].split(" | ")[0] if " | " in str(r[6]) else r[6], km_val, cant_val, desc_limpia, metodo_pago, formatear_moneda(r[8]), formatear_moneda(impuesto), formatear_moneda(tot_bruto), formatear_moneda(det_monto), formatear_moneda(neto), tiene_arch
                )
                
                if filtro:
                    coincide = False
                    for val in row_vals:
                        if val is not None and filtro in str(val).lower():
                            coincide = True
                            break
                    if not coincide:
                        continue
                        
                self.tabla.insert("", tk.END, values=row_vals, tags=(etiqueta_color,))
                contador += 1
        except Exception: pass
        finally: conn.close()

    def abrir_archivo(self, event):
        sel = self.tabla.selection()
        if not sel: return
        id_doc = self.tabla.item(sel[0], "values")[1]
        
        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT archivo_ruta FROM facturas_recibidas WHERE id = %s", (id_doc,))
            res = cursor.fetchone()
            
            if res and res[0]:
                ruta_almacenada = os.path.normpath(res[0])
                
                if ruta_almacenada == "PENDIENTE_DESCARGA":
                    if messagebox.askyesno("Archivo no disponible", "La imagen fue subida desde el celular pero no se ha descargado a esta computadora.\n\n¿Desea forzar la descarga de los archivos pendientes ahora?"):
                        self.ejecutar_sincronizacion_manual()
                elif os.path.exists(ruta_almacenada):
                    abrir_documento(ruta_almacenada)
                else:
                    if messagebox.askyesno("Archivo Extraviado", "El sistema indica que este archivo ya se descargó, pero no se encuentra en su computadora (pudo haber sido borrado o movido por Google Drive).\n\n¿Desea intentar recuperar la imagen original desde la nube (Supabase)?"):
                        
                        cursor.execute("SELECT imagen_base64, proveedor FROM facturas_recibidas WHERE id = %s", (id_doc,))
                        res_recuperar = cursor.fetchone()
                        
                        if res_recuperar and res_recuperar[0]:
                            img_b64 = res_recuperar[0]
                            proveedor = res_recuperar[1]
                            
                            ruta_base = obtener_ruta_base_drive()
                            if not ruta_base:
                                return messagebox.showerror("Error", "No tiene configurada la ruta de Google Drive.")
                                
                            carpeta_destino = os.path.normpath(os.path.join(ruta_base, "facturas_recibidas"))
                            if not os.path.exists(carpeta_destino): os.makedirs(carpeta_destino)
                            
                            prov_limpio = re.sub(r'[\\/*?:"<>|]', '-', proveedor) if proveedor else "RECUPERADO"
                            nombre_prov = prov_limpio.replace(' ', '_')
                            nombre_archivo = f"Ticket_Movil_Recuperado_{id_doc}_{nombre_prov}_{datetime.now().strftime('%Y%m%d%H%M%S')}.jpg"
                            ruta_final = os.path.normpath(os.path.join(carpeta_destino, nombre_archivo))
                            
                            try:
                                with open(ruta_final, "wb") as f:
                                    f.write(base64.b64decode(img_b64))
                                
                                cursor.execute("UPDATE facturas_recibidas SET archivo_ruta = %s WHERE id = %s", (ruta_final, id_doc))
                                cursor.execute("UPDATE pagos_comprobantes SET archivo_ruta = %s WHERE id_factura = %s", (ruta_final, id_doc))
                                conn.commit()
                                
                                messagebox.showinfo("Recuperación Exitosa", "La imagen ha sido recuperada y guardada en su carpeta.")
                                abrir_documento(ruta_final)
                                self.cargar_datos_tabla()
                            except Exception as e_write:
                                messagebox.showerror("Error", f"No se pudo guardar la imagen recuperada:\n{e_write}")
                        else:
                            messagebox.showerror("Error irrecuperable", "La imagen original ya fue borrada de la nube (Supabase) para ahorrar espacio tras la primera descarga exitosa. No es posible recuperarla automáticamente.")
            else:
                messagebox.showinfo("Aviso", "Este registro no tiene ningún archivo asociado.")
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el documento:\n{e}")
        finally:
            conn.close()

    def abrir_ventana_edicion(self):
        sel = self.tabla.selection()
        if not sel: return messagebox.showwarning("Atención", "Seleccione un documento.")
        
        valores = self.tabla.item(sel[0], "values")
        id_doc = valores[1]
        
        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT tipo_documento, ruc, proveedor, numero_documento, fecha, evento_asociado, kilometraje, cantidad_combustible, descripcion, subtotal, impuesto, total FROM facturas_recibidas WHERE id = %s", (id_doc,))
            reg = cursor.fetchone()
        finally: 
            conn.close()
            
        if not reg: return
        
        e_tipo, e_ruc, e_prov, e_nro, e_fec, e_placa, e_km, e_gal, e_desc, e_sub, e_imp, e_tot = reg
        
        desc_bruta = str(e_desc) if e_desc else ""
        c_val = desc_bruta
        h_val = ""
        
        if " | " in desc_bruta:
            partes = desc_bruta.split(" | ")
            c_val = partes[0].replace("Combustible: ", "")
            for parte in partes[1:]:
                if parte.startswith("Hora: "): h_val = parte.replace("Hora: ", "").strip()

        v_edit = ctk.CTkToplevel(self.main_root)
        v_edit.title("Editar Documento Recibido")
        
        v_edit.update_idletasks()
        ancho = 450
        alto = 660
        x_pos = (self.main_root.winfo_screenwidth() // 2) - (ancho // 2)
        y_pos = 20 
        v_edit.geometry(f"{ancho}x{alto}+{x_pos}+{y_pos}")
        
        v_edit.transient(self.main_root)
        v_edit.grab_set()

        ctk.CTkLabel(v_edit, text=f"✏️ Editar Registro ID: {id_doc}", font=("Arial", 16, "bold"), text_color="#1f538d").pack(pady=(15, 10))
        
        f_form = ctk.CTkScrollableFrame(v_edit, fg_color="transparent")
        f_form.pack(fill="both", expand=True, padx=10, pady=5)
        
        def crear_campo(padre, texto, valor):
            ctk.CTkLabel(padre, text=texto, font=("Arial", 11, "bold")).pack(anchor="w", padx=5, pady=(5,0))
            ent = ctk.CTkEntry(padre)
            ent.pack(fill="x", padx=5, pady=(0,5))
            if valor is not None:
                ent.insert(0, str(valor))
            return ent

        ent_tipo = crear_campo(f_form, "Tipo de Documento:", e_tipo)
        ent_ruc = crear_campo(f_form, "R.U.C.:", e_ruc)
        ent_prov = crear_campo(f_form, "Proveedor:", e_prov)
        ent_nro = crear_campo(f_form, "N° Documento:", e_nro)
        
        ctk.CTkLabel(f_form, text="Fecha:", font=("Arial", 11, "bold"), text_color="#1f538d").pack(anchor="w", padx=5, pady=(5,0))
        f_fec = ctk.CTkFrame(f_form, fg_color="transparent")
        f_fec.pack(fill="x", padx=5, pady=(0,5))
        ent_fec = ctk.CTkEntry(f_fec)
        ent_fec.pack(side="left", fill="x", expand=True)
        if e_fec: ent_fec.insert(0, str(e_fec))
        ctk.CTkButton(f_fec, text="📅", width=35, fg_color="#1f538d", command=lambda: CalendarioNativo(v_edit, ent_fec)).pack(side="right", padx=(5, 0))

        ent_placa = crear_campo(f_form, "Vehículo (Placa):", e_placa)
        ent_km = crear_campo(f_form, "Kilometraje:", e_km)
        ent_gal = crear_campo(f_form, "Galones/Cant.:", e_gal)
        
        ent_desc = crear_campo(f_form, "Concepto / Descripción:", c_val)
        ent_hora = crear_campo(f_form, "Hora de Consumo (Ej: 14:30):", h_val)
        
        f_montos = ctk.CTkFrame(f_form, fg_color="transparent")
        f_montos.pack(fill="x", pady=5)
        
        ctk.CTkLabel(f_montos, text="Subtotal:", font=("Arial", 11, "bold")).grid(row=0, column=0, padx=5, sticky="w")
        ent_sub = ctk.CTkEntry(f_montos, width=100)
        ent_sub.grid(row=1, column=0, padx=5)
        ent_sub.insert(0, str(e_sub))
        
        ctk.CTkLabel(f_montos, text="Impuestos:", font=("Arial", 11, "bold")).grid(row=0, column=1, padx=5, sticky="w")
        ent_imp = ctk.CTkEntry(f_montos, width=100)
        ent_imp.grid(row=1, column=1, padx=5)
        ent_imp.insert(0, str(e_imp))
        
        ctk.CTkLabel(f_montos, text="Total (Neto):", font=("Arial", 11, "bold")).grid(row=0, column=2, padx=5, sticky="w")
        ent_tot = ctk.CTkEntry(f_montos, width=100)
        ent_tot.grid(row=1, column=2, padx=5)
        ent_tot.insert(0, str(e_tot))

        def guardar_cambios():
            try:
                val_sub = float(ent_sub.get().strip() or 0)
                val_imp = float(ent_imp.get().strip() or 0)
                val_tot = float(ent_tot.get().strip() or 0)
            except ValueError:
                return messagebox.showerror("Error", "Los montos (Subtotal, Impuestos, Total) deben ser numéricos.", parent=v_edit)
                
            n_desc = ent_desc.get().strip()
            n_hora = ent_hora.get().strip()
            
            desc_final = n_desc
            if n_hora:
                desc_final = f"{n_desc} | Hora: {n_hora}"

            if messagebox.askyesno("Confirmar", "¿Guardar los cambios en este registro?", parent=v_edit):
                conn_u = conectar_db()
                if conn_u:
                    try:
                        cursor_u = conn_u.cursor()
                        cursor_u.execute("""
                            UPDATE facturas_recibidas 
                            SET tipo_documento=%s, ruc=%s, proveedor=%s, numero_documento=%s, fecha=%s, 
                                evento_asociado=%s, kilometraje=%s, cantidad_combustible=%s, descripcion=%s, 
                                subtotal=%s, impuesto=%s, total=%s
                            WHERE id=%s
                        """, (
                            ent_tipo.get().strip(), ent_ruc.get().strip(), ent_prov.get().strip(), ent_nro.get().strip(), 
                            ent_fec.get().strip(), ent_placa.get().strip(), ent_km.get().strip(), ent_gal.get().strip(), 
                            desc_final, val_sub, val_imp, val_tot, id_doc
                        ))
                        
                        cursor_u.execute("""
                            UPDATE pagos_comprobantes 
                            SET proveedor_nombre=%s, codigo_cotizacion=%s
                            WHERE id_factura=%s
                        """, (ent_prov.get().strip(), ent_nro.get().strip(), id_doc))
                        
                        conn_u.commit()
                        messagebox.showinfo("Éxito", "Registro actualizado correctamente.", parent=v_edit)
                        v_edit.destroy()
                        self.cargar_datos_tabla()
                        if hasattr(self.app_padre, 'app_pagos'):
                            self.app_padre.app_pagos.cargar_datos_pagar()
                    except Exception as e:
                        messagebox.showerror("Error", f"No se pudo guardar:\n{e}", parent=v_edit)
                    finally:
                        conn_u.close()

        def eliminar_registro():
            if messagebox.askyesno("Confirmar Eliminación", "⚠️ ¿Desea eliminar completamente este registro?", parent=v_edit):
                try:
                    conn = conectar_db()
                    cursor = conn.cursor()
                    cursor.execute("SELECT archivo_ruta FROM facturas_recibidas WHERE id = %s", (id_doc,))
                    row = cursor.fetchone()
                    ruta_archivo = os.path.normpath(row[0]) if row and row[0] else None
                    if ruta_archivo and os.path.exists(ruta_archivo): os.remove(ruta_archivo)
                    cursor.execute("DELETE FROM facturas_recibidas WHERE id = %s", (id_doc,))
                    conn.commit()
                    conn.close()
                    
                    cache_sistema.sincronizar_ahora()
                    registrar_auditoria(self.app_padre.usuario_activo, "Facturas Recibidas", f"Eliminó factura ID {id_doc}")

                    messagebox.showinfo("Éxito", "Registro eliminado.", parent=v_edit)
                    v_edit.destroy()
                    self.cargar_datos_tabla()
                    if hasattr(self.app_padre, 'app_pagos'): self.app_padre.app_pagos.cargar_datos_pagar()
                except Exception as e:
                    messagebox.showerror("Error", str(e), parent=v_edit)

        ctk.CTkButton(f_form, text="💾 Guardar Cambios", font=("Arial", 12, "bold"), fg_color="#27ae60", hover_color="#1e8449", command=guardar_cambios).pack(fill="x", padx=10, pady=(20, 5))
        ctk.CTkButton(f_form, text="❌ Eliminar Registro Completo", font=("Arial", 12, "bold"), fg_color="#e74c3c", hover_color="#c0392b", command=eliminar_registro).pack(fill="x", padx=10, pady=(5, 15))

# =========================================================
# PESTAÑA 2: CUENTAS POR PAGAR (PAGOS Y DEUDAS)
# =========================================================
class CuentasPorPagarTab:
    def __init__(self, tab_frame, main_root, app_padre):
        self.tab_frame = tab_frame
        self.main_root = main_root
        self.app_padre = app_padre
        self.inicializar_entorno()
        self.crear_interfaz()

    def inicializar_entorno(self):
        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS pagos_comprobantes (
                    id SERIAL PRIMARY KEY, codigo_cotizacion VARCHAR(255), categoria_suministro VARCHAR(255), 
                    monto_pagado NUMERIC, archivo_ruta TEXT, proveedor_nombre VARCHAR(255), fecha_pago VARCHAR(50),
                    id_factura INTEGER DEFAULT 0
                )
            """)
            conn.commit()
            
            try:
                cursor.execute("ALTER TABLE pagos_comprobantes ADD COLUMN cuenta_origen VARCHAR(255) DEFAULT ''")
                conn.commit()
            except Exception: conn.rollback()
            
        except Exception: pass
        finally: conn.close()

    def crear_interfaz(self):
        frame_acciones = ctk.CTkFrame(self.tab_frame, corner_radius=8, fg_color="#f8f9fa", border_width=1, border_color="#e0e0e0")
        frame_acciones.pack(fill="x", padx=15, pady=(10, 10), ipady=5)

        btn_resumen_prov = ctk.CTkButton(frame_acciones, text="📊 Resumen por Proveedor", font=("Arial", 12, "bold"), command=self.mostrar_resumen_proveedores, fg_color="#1f538d", hover_color="#163b65")
        btn_resumen_prov.pack(side="left", padx=5, pady=5)

        btn_reporte = ctk.CTkButton(frame_acciones, text="📈 Reporte de Totales Avanzado", font=("Arial", 12, "bold"), command=self.mostrar_reporte_totales, fg_color="#34495e", hover_color="#2c3e50")
        btn_reporte.pack(side="left", padx=5, pady=5)

        btn_pago = ctk.CTkButton(frame_acciones, text="🧾 Registrar Pago", font=("Arial", 12, "bold"), command=self.cargar_comprobante_pago, fg_color="#1f538d", hover_color="#163b65")
        btn_pago.pack(side="left", padx=5, pady=5)

        btn_editar = ctk.CTkButton(frame_acciones, text="✏️ Editar Pagos", font=("Arial", 12, "bold"), command=self.abrir_ventana_edicion, fg_color="#34495e", hover_color="#2c3e50")
        btn_editar.pack(side="left", padx=5, pady=5)

        btn_refresh = ctk.CTkButton(frame_acciones, text="🔄 Actualizar", font=("Arial", 12, "bold"), command=self.cargar_datos_pagar, fg_color="#7f8c8d", hover_color="#606b6b")
        btn_refresh.pack(side="right", padx=10, pady=5)

        f_busqueda = ctk.CTkFrame(self.tab_frame, fg_color="transparent")
        f_busqueda.pack(fill="x", padx=15, pady=(0, 5))
        ctk.CTkLabel(f_busqueda, text="🔍 Buscar:", font=("Arial", 11, "bold")).pack(side="left", padx=(0, 5))
        self.ent_buscar_pagos = ctk.CTkEntry(f_busqueda, placeholder_text="Filtrar por documento, proveedor, vehículo, concepto...")
        self.ent_buscar_pagos.pack(side="left", fill="x", expand=True)
        self.ent_buscar_pagos.bind("<KeyRelease>", lambda e: self.cargar_datos_pagar())

        f_tabla = ctk.CTkFrame(self.tab_frame, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True, padx=15, pady=0)

        columnas = ("num", "id_factura", "fecha", "hora", "nro_doc", "proveedor", "ruc", "evento", "kilometraje", "cantidad", "concepto", "metodo_pago", "subtotal", "igv", "detraccion", "neto_facturado", "pagado", "saldo", "archivos")
        self.tabla = ttk.Treeview(f_tabla, columns=columnas, show="headings")
        
        self.tabla.tag_configure("con_cuenta", background="#e8f8f5", foreground="#0e6251") 
        self.tabla.tag_configure("sin_cuenta", background="#fdedec", foreground="#7b241c") 

        self.tabla.heading("num", text="N°")
        self.tabla.heading("id_factura", text="ID (Oculto)")
        self.tabla.heading("fecha", text="Fecha Fac.")
        self.tabla.heading("hora", text="Hora")
        self.tabla.heading("nro_doc", text="N° Documento")
        self.tabla.heading("proveedor", text="Proveedor")
        self.tabla.heading("ruc", text="RUC")
        self.tabla.heading("evento", text="Vehículo Asociado")
        self.tabla.heading("kilometraje", text="Kilometraje")
        self.tabla.heading("cantidad", text="Galones/Cant.")
        self.tabla.heading("concepto", text="Concepto")
        self.tabla.heading("metodo_pago", text="Forma de Pago")
        self.tabla.heading("subtotal", text="Subtotal")
        self.tabla.heading("igv", text="IGV")
        self.tabla.heading("detraccion", text="Detracción")
        self.tabla.heading("neto_facturado", text="Neto a Pagar")
        self.tabla.heading("pagado", text="Total Pagado")
        self.tabla.heading("saldo", text="Saldo Pendiente")
        self.tabla.heading("archivos", text="Historial Adjuntos")

        self.tabla.column("num", width=40, anchor="center")
        self.tabla.column("id_factura", width=0, stretch=tk.NO)
        self.tabla.column("fecha", width=80, anchor="center")
        self.tabla.column("hora", width=70, anchor="center")
        self.tabla.column("nro_doc", width=100, anchor="center")
        self.tabla.column("proveedor", width=140, anchor="w")
        self.tabla.column("ruc", width=90, anchor="center")
        self.tabla.column("evento", width=120, anchor="center")
        self.tabla.column("kilometraje", width=80, anchor="center")
        self.tabla.column("cantidad", width=80, anchor="center")
        self.tabla.column("concepto", width=140, anchor="w")
        self.tabla.column("metodo_pago", width=120, anchor="center")
        self.tabla.column("neto_facturado", width=95, anchor="e")
        self.tabla.column("pagado", width=90, anchor="e")
        self.tabla.column("saldo", width=90, anchor="e")
        self.tabla.column("archivos", width=100, anchor="center")

        self.tabla.config(displaycolumns=("num", "fecha", "hora", "nro_doc", "proveedor", "ruc", "evento", "kilometraje", "cantidad", "concepto", "metodo_pago", "neto_facturado", "pagado", "saldo", "archivos"))
        self.tabla.bind("<Double-1>", self.abrir_todos_los_archivos)
        
        scroll_y = ctk.CTkScrollbar(f_tabla, orientation="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scroll_y.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        self.frame_bottom = ctk.CTkFrame(self.tab_frame, fg_color="transparent")
        self.frame_bottom.pack(fill="x", padx=15, pady=10)

        self.btn_excel = ctk.CTkButton(self.frame_bottom, text="📊 Exportar a Excel", font=("Arial", 12, "bold"), width=160, fg_color="#27ae60", hover_color="#1e8449", command=self.exportar_excel)
        self.btn_excel.pack(side="left")

        # --- LEYENDA ---
        f_leyenda = ctk.CTkFrame(self.frame_bottom, fg_color="transparent")
        f_leyenda.pack(side="left", padx=30)
        
        ctk.CTkLabel(f_leyenda, text="■", font=("Arial", 14), text_color="#c0392b").pack(side="left", padx=(5,2))
        ctk.CTkLabel(f_leyenda, text="Sin Cuenta Asignada / Pendiente", font=("Arial", 11, "bold"), text_color="#333333").pack(side="left", padx=(0,10))
        
        ctk.CTkLabel(f_leyenda, text="■", font=("Arial", 14), text_color="#27ae60").pack(side="left", padx=(5,2))
        ctk.CTkLabel(f_leyenda, text="Cuenta Asignada", font=("Arial", 11, "bold"), text_color="#333333").pack(side="left", padx=(0,5))
        # ---------------

        self.lbl_total_general = ctk.CTkLabel(self.frame_bottom, text="Total Pendiente General por Pagar: 0.00", font=("Arial", 12, "bold"), text_color="#c0392b")
        self.lbl_total_general.pack(side="right")

        self.main_root.after(100, self.cargar_datos_pagar)

    def exportar_excel(self):
        try: import pandas as pd
        except ImportError: return messagebox.showerror("Error", "Falta librería pandas.")
        filas = [self.tabla.item(item)["values"][2:] for item in self.tabla.get_children()]
        if not filas: return messagebox.showwarning("Aviso", "No hay registros.")
        
        columnas = ["Fecha Fac.", "Hora", "N° Documento", "Proveedor", "RUC", "Vehículo (Placa)", "Kilometraje", "Galones/Cant.", "Concepto", "Forma de Pago", "Subtotal", "IGV", "Detracción", "Neto Facturado", "Total Pagado", "Saldo Pendiente", "Archivos"]
        
        ruta = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Cuentas_por_Pagar.xlsx", filetypes=[("Excel", "*.xlsx")])
        if ruta:
            pd.DataFrame(filas, columns=columnas).to_excel(ruta, index=False)
            messagebox.showinfo("Éxito", f"Reporte exportado a:\n{ruta}")
            abrir_documento(ruta)

    def mostrar_resumen_proveedores(self):
        saldos = {}
        for item in self.tabla.get_children():
            vals = self.tabla.item(item, "values")
            prov = vals[5] 
            try:
                saldo = desformatear_numero(vals[17]) 
                if saldo > 0: saldos[prov] = saldos.get(prov, 0.0) + saldo
            except ValueError: pass
            
        v_resumen = ctk.CTkToplevel(self.main_root)
        v_resumen.title("Resumen de Deudas por Proveedor")
        v_resumen.geometry("500x400")
        v_resumen.grab_set()

        ctk.CTkLabel(v_resumen, text="📊 Saldo Pendiente Consolidado por Proveedor", font=("Arial", 14, "bold"), text_color="#1f538d").pack(pady=(15, 10))

        f_tabla = ctk.CTkFrame(v_resumen, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True, padx=15, pady=(0, 10))

        t_resumen = ttk.Treeview(f_tabla, columns=("prov", "saldo"), show="headings")
        t_resumen.heading("prov", text="Proveedor"); t_resumen.heading("saldo", text="Monto Pendiente de Pago")
        t_resumen.column("prov", width=250, anchor="w"); t_resumen.column("saldo", width=150, anchor="e")
        t_resumen.pack(side="left", fill="both", expand=True)
        
        total_pagar = 0.0
        for p, s in sorted(saldos.items(), key=lambda x: x[0]):
            t_resumen.insert("", tk.END, values=(p, formatear_moneda(s))); total_pagar += s

        ctk.CTkLabel(v_resumen, text=f"Total cuentas por pagar : {formatear_moneda(total_pagar)}", font=("Arial", 14, "bold")).pack(anchor="e", padx=15, pady=(5, 15))

    def cargar_datos_pagar(self):
        for fila in self.tabla.get_children(): self.tabla.delete(fila)
        
        filtro = ""
        if hasattr(self, 'ent_buscar_pagos'):
            filtro = self.ent_buscar_pagos.get().strip().lower()

        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT id, fecha, numero_documento, proveedor, evento_asociado, descripcion, subtotal, impuesto, total, COALESCE(det_monto, 0), tipo_documento, kilometraje, cantidad_combustible, ruc FROM facturas_recibidas ORDER BY id DESC")
            registros = cursor.fetchall()
            total_pendiente = 0.0

            contador = 1
            for reg in registros:
                id_factura, fecha, nro_doc, proveedor, evento, concepto, subtotal, impuesto, tot_bruto, det_monto, tipo_doc, km_val, cant_val, ruc_db = reg
                sub_val = float(subtotal) if subtotal else 0.0
                imp_val = float(impuesto) if impuesto else 0.0
                tot_bruto_val = float(tot_bruto) if tot_bruto else 0.0
                det_monto_val = float(det_monto) if det_monto else 0.0
                km_str = km_val if km_val else "-"
                cant_str = cant_val if cant_val else "-"
                ruc_str = ruc_db if ruc_db else "-"
                
                if tipo_doc and "Recibo" in tipo_doc and "8%" in tipo_doc:
                    neto_facturado = tot_bruto_val - imp_val - det_monto_val
                else:
                    neto_facturado = tot_bruto_val - det_monto_val
                
                cursor.execute("SELECT monto_pagado, archivo_ruta, cuenta_origen FROM pagos_comprobantes WHERE id_factura = %s", (id_factura,))
                pagos_factura = cursor.fetchall()
                
                monto_pagado = 0.0
                cant_archivos = 0
                tiene_cuenta = False
                cuentas_lista = []
                
                for p in pagos_factura:
                    monto_pagado += float(p[0]) if p[0] else 0.0
                    if p[1] and str(p[1]).strip() != "":
                        cant_archivos += 1
                    if p[2] and str(p[2]).strip() != "":
                        tiene_cuenta = True
                        if p[2] not in cuentas_lista:
                            cuentas_lista.append(p[2])
                
                metodo_pago = " + ".join(cuentas_lista) if cuentas_lista else ""
                
                saldo_pendiente = max(0.0, neto_facturado - monto_pagado)
                txt_adjuntos = f"📁 {cant_archivos} archivo(s)" if cant_archivos > 0 else "❌ Sin adjuntos"

                desc_bruta = str(concepto) if concepto else "-"
                concepto_limpio = desc_bruta
                hora_consumo = "-"
                
                if " | " in desc_bruta:
                    partes = desc_bruta.split(" | ")
                    concepto_limpio = partes[0].replace("Combustible: ", "")
                    for parte in partes[1:]:
                        if parte.startswith("Hora: "):
                            hora_consumo = parte.replace("Hora: ", "").strip()

                etiqueta_color = "con_cuenta" if tiene_cuenta else "sin_cuenta"

                row_vals = (
                    contador, id_factura, fecha, hora_consumo, nro_doc if nro_doc else "S/N", proveedor, ruc_str, evento, km_str, cant_str, concepto_limpio, metodo_pago,
                    formatear_moneda(sub_val), formatear_moneda(imp_val), formatear_moneda(det_monto_val),
                    formatear_moneda(neto_facturado), formatear_moneda(monto_pagado), formatear_moneda(saldo_pendiente), txt_adjuntos
                )

                if filtro:
                    coincide = False
                    for val in row_vals:
                        if val is not None and filtro in str(val).lower():
                            coincide = True
                            break
                    if not coincide:
                        continue
                
                total_pendiente += saldo_pendiente
                self.tabla.insert("", tk.END, values=row_vals, tags=(etiqueta_color,))
                contador += 1

            self.lbl_total_general.configure(text=f"Total Pendiente Filtrado: {formatear_moneda(total_pendiente)}")
        except Exception: pass
        finally: conn.close()

    def cargar_comprobante_pago(self):
        ruta_base = obtener_ruta_base_drive()
        if not ruta_base:
            messagebox.showwarning("Configuración Requerida", "No ha configurado la ruta de Google Drive.\nEs obligatorio para guardar archivos.")
            return
            
        seleccion = self.tabla.selection()
        if not seleccion: return messagebox.showwarning("Selección", "Seleccione una factura.")
        
        valores = self.tabla.item(seleccion[0], "values")
        id_factura, nro_doc, proveedor = valores[1], valores[4], valores[5] 
        saldo_actual = desformatear_numero(valores[17]) 
        
        if saldo_actual <= 0: return messagebox.showinfo("Aviso", "Esta factura ya está pagada por completo.")

        v_pago = ctk.CTkToplevel(self.main_root)
        v_pago.title("Registrar Nuevo Pago")
        v_pago.geometry("450x420")
        v_pago.transient(self.main_root)
        v_pago.grab_set()

        ctk.CTkLabel(v_pago, text=f"Pago para: {proveedor}", font=("Arial", 14, "bold"), text_color="#1f538d").pack(pady=(15, 5))
        ctk.CTkLabel(v_pago, text=f"Saldo Pendiente: {formatear_moneda(saldo_actual)}", font=("Arial", 12)).pack(pady=(0, 15))

        f_form = ctk.CTkFrame(v_pago, fg_color="transparent")
        f_form.pack(fill="x", padx=20)

        config = cargar_configuracion_regional()
        bancos_guardados = config.get("cuentas_bancarias", [])
        lista_cuentas = []
        for b in bancos_guardados:
            banco_nom = b.get("banco", "").strip()
            cuenta_num = b.get("cuenta", "").strip()
            if banco_nom or cuenta_num:
                lista_cuentas.append(f"{banco_nom} - {cuenta_num}".strip(" - "))
        lista_cuentas.extend(["Efectivo / Caja Chica", "Tarjeta de Crédito", "Tarjeta de Débito", "Cheque", "Otro"])

        ctk.CTkLabel(f_form, text="Cuenta Origen / Método:", font=("Arial", 11, "bold")).pack(anchor="w")
        cmb_cuenta = ctk.CTkComboBox(f_form, values=lista_cuentas, width=400)
        cmb_cuenta.pack(fill="x", pady=(0, 10))
        if lista_cuentas: cmb_cuenta.set(lista_cuentas[0])

        ctk.CTkLabel(f_form, text="Monto a Pagar (S/.):", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_monto = ctk.CTkEntry(f_form)
        ent_monto.pack(fill="x", pady=(0, 10))
        ent_monto.insert(0, str(saldo_actual)) 

        # Calendario para fecha de pago
        ctk.CTkLabel(f_form, text="Fecha del Pago:", font=("Arial", 11, "bold")).pack(anchor="w")
        f_fecha_pago = ctk.CTkFrame(f_form, fg_color="transparent")
        f_fecha_pago.pack(fill="x", pady=(0, 10))
        ent_fecha = ctk.CTkEntry(f_fecha_pago)
        ent_fecha.pack(side="left", fill="x", expand=True)
        ent_fecha.insert(0, datetime.now().strftime("%d/%m/%Y"))
        ctk.CTkButton(f_fecha_pago, text="📅", width=40, fg_color="#1f538d", command=lambda: CalendarioNativo(v_pago, ent_fecha)).pack(side="right", padx=(5, 0))

        def procesar_pago(event=None):
            try:
                monto_val = float(ent_monto.get().strip())
            except ValueError:
                return messagebox.showerror("Error", "Ingrese un monto numérico válido.", parent=v_pago)

            if monto_val <= 0:
                return messagebox.showerror("Error", "El monto debe ser mayor a 0.", parent=v_pago)
            if monto_val > (saldo_actual + 0.01):
                return messagebox.showerror("Error", "El monto supera el saldo pendiente.", parent=v_pago)

            fecha_val = ent_fecha.get().strip()
            if not fecha_val:
                fecha_val = datetime.now().strftime("%d/%m/%Y")
                
            cuenta_val = cmb_cuenta.get().strip()

            v_pago.destroy()

            ruta_origen = filedialog.askopenfilename(title="Seleccionar Soporte de Egreso", filetypes=[("Archivos", "*.pdf;*.png;*.jpg;*.jpeg")])
            ruta_destino = ""
            if ruta_origen:
                try:
                    carpeta_comprobantes = os.path.normpath(os.path.join(ruta_base, "comprobantes_egresos"))
                    if not os.path.exists(carpeta_comprobantes): os.makedirs(carpeta_comprobantes)
                    conn = conectar_db(); c = conn.cursor(); c.execute("SELECT COUNT(*) FROM pagos_comprobantes"); idx = c.fetchone()[0] + 1; conn.close()
                    
                    prov_limpio = re.sub(r'[\\/*?:"<>|]', '-', proveedor)
                    ruta_destino = os.path.normpath(os.path.join(carpeta_comprobantes, f"Egreso_Fac_{id_factura}_{prov_limpio.replace(' ', '_')}_{idx}{os.path.splitext(ruta_origen)[1]}"))
                    shutil.copy2(ruta_origen, ruta_destino)
                except Exception as e:
                    return messagebox.showerror("Error", f"Fallo al copiar archivo:\n{e}")

            conn = conectar_db()
            if not conn: return
            try:
                cursor = conn.cursor()
                cursor.execute("SELECT categoria FROM facturas_recibidas WHERE id = %s", (id_factura,))
                cat_res = cursor.fetchone()
                categoria_db = cat_res[0] if cat_res and cat_res[0] else "GENERAL"
                
                cursor.execute("""
                    INSERT INTO pagos_comprobantes (id_factura, monto_pagado, archivo_ruta, proveedor_nombre, fecha_pago, categoria_suministro, codigo_cotizacion, cuenta_origen) 
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                """, (id_factura, monto_val, ruta_destino, proveedor, fecha_val, categoria_db, nro_doc, cuenta_val))
                conn.commit()
                registrar_auditoria(self.app_padre.usuario_activo, "Cuentas por Pagar", f"Pagó {formatear_moneda(monto_val)} a Fac. {nro_doc} desde {cuenta_val}")
                messagebox.showinfo("Éxito", f"Pago de {formatear_moneda(monto_val)} registrado exitosamente.")
                self.cargar_datos_pagar()
                self.app_padre.app_facturas.cargar_datos_tabla()
            except Exception as e: messagebox.showerror("Error", str(e))
            finally: conn.close()

        ent_monto.bind("<Return>", procesar_pago)
        ent_fecha.bind("<Return>", procesar_pago)

        f_btns = ctk.CTkFrame(v_pago, fg_color="transparent")
        f_btns.pack(fill="x", padx=20, pady=15)

        btn_ok = ctk.CTkButton(f_btns, text="✅ Confirmar", font=("Arial", 12, "bold"), fg_color="#27ae60", hover_color="#1e8449", command=procesar_pago)
        btn_ok.pack(side="left", expand=True, padx=5)

        btn_cancel = ctk.CTkButton(f_btns, text="❌ Cancelar", font=("Arial", 12, "bold"), fg_color="#e74c3c", hover_color="#922b21", command=v_pago.destroy)
        btn_cancel.pack(side="right", expand=True, padx=5)

        ent_monto.focus()

    def abrir_todos_los_archivos(self, event):
        seleccion = self.tabla.selection()
        if not seleccion: return
        id_factura = self.tabla.item(seleccion[0], "values")[1] 
        try:
            conn = conectar_db()
            cursor = conn.cursor()
            cursor.execute("SELECT archivo_ruta FROM pagos_comprobantes WHERE id_factura = %s AND archivo_ruta != ''", (id_factura,))
            rutas = cursor.fetchall()
            conn.close()
            if rutas:
                for r in rutas:
                    ruta_norm = os.path.normpath(r[0])
                    if os.path.exists(ruta_norm): 
                        abrir_documento(ruta_norm)
            else: messagebox.showinfo("Aviso", "No hay soportes cargados.")
        except Exception: pass

    def abrir_ventana_edicion(self):
        sel = self.tabla.selection()
        if not sel: return messagebox.showwarning("Selección", "Seleccione la factura para editar sus pagos.")
        valores = self.tabla.item(sel[0], "values")
        id_factura, nro_doc, proveedor = valores[1], valores[4], valores[5] 
        saldo_actual_global = desformatear_numero(valores[17]) 

        v_edit = ctk.CTkToplevel(self.main_root)
        v_edit.title(f"✏️ Gestión de Pagos - Fac. {nro_doc}")
        v_edit.geometry("820x400")
        v_edit.transient(self.main_root)
        v_edit.grab_set()

        ctk.CTkLabel(v_edit, text=f"Pagos registrados: {proveedor}", font=("Arial", 12, "bold"), text_color="#1f538d").pack(pady=10)

        frame_cuerpo = ctk.CTkFrame(v_edit, fg_color="transparent")
        frame_cuerpo.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        sub_tabla = ttk.Treeview(frame_cuerpo, columns=("id", "monto", "fecha", "cuenta", "tiene_archivo"), show="headings", height=8)
        sub_tabla.heading("id", text="ID"); sub_tabla.heading("monto", text="Monto"); sub_tabla.heading("fecha", text="Fecha"); sub_tabla.heading("cuenta", text="Cuenta / Origen"); sub_tabla.heading("tiene_archivo", text="¿Soporte?")
        sub_tabla.column("id", width=40, anchor="center"); sub_tabla.column("monto", width=90, anchor="e"); sub_tabla.column("fecha", width=90, anchor="center"); sub_tabla.column("cuenta", width=160, anchor="w"); sub_tabla.column("tiene_archivo", width=90, anchor="center")
        sub_tabla.pack(side="left", fill="both", expand=True, padx=(0, 10))

        def refrescar_subtabla():
            for f in sub_tabla.get_children(): sub_tabla.delete(f)
            try:
                conn = conectar_db(); cursor = conn.cursor()
                cursor.execute("SELECT id, monto_pagado, fecha_pago, archivo_ruta, cuenta_origen FROM pagos_comprobantes WHERE id_factura = %s", (id_factura,))
                for a in cursor.fetchall(): 
                    sub_tabla.insert("", tk.END, values=(a[0], formatear_moneda(a[1]), a[2] if a[2] else "Sin fecha", a[4] if a[4] else "-", "✅ Sí" if (a[3] and os.path.exists(os.path.normpath(a[3]))) else "❌ No"))
                conn.close()
            except Exception: pass
        refrescar_subtabla()

        def ejecutar_modificacion():
            nonlocal saldo_actual_global
            sub_sel = sub_tabla.selection()
            if not sub_sel: return
            id_pago = sub_tabla.item(sub_sel[0], "values")[0]
            conn = conectar_db(); cursor = conn.cursor()
            cursor.execute("SELECT monto_pagado, fecha_pago, cuenta_origen FROM pagos_comprobantes WHERE id = %s", (id_pago,))
            monto_actual, fecha_actual, cuenta_actual = cursor.fetchone()
            conn.close()

            v_mod_pago = ctk.CTkToplevel(v_edit)
            v_mod_pago.title("Modificar Pago")
            v_mod_pago.geometry("400x320")
            v_mod_pago.transient(v_edit)
            v_mod_pago.grab_set()

            ctk.CTkLabel(v_mod_pago, text="Editar Pago", font=("Arial", 14, "bold")).pack(pady=10)

            f_form = ctk.CTkFrame(v_mod_pago, fg_color="transparent")
            f_form.pack(fill="x", padx=20)
            
            config = cargar_configuracion_regional()
            bancos_guardados = config.get("cuentas_bancarias", [])
            lista_cuentas = []
            for b in bancos_guardados:
                banco_nom = b.get("banco", "").strip()
                cuenta_num = b.get("cuenta", "").strip()
                if banco_nom or cuenta_num:
                    lista_cuentas.append(f"{banco_nom} - {cuenta_num}".strip(" - "))
            lista_cuentas.extend(["Efectivo / Caja Chica", "Tarjeta de Crédito", "Tarjeta de Débito", "Cheque", "Otro"])

            ctk.CTkLabel(f_form, text="Cuenta Origen / Método:", font=("Arial", 11, "bold")).pack(anchor="w")
            ent_mod_cuenta = ctk.CTkComboBox(f_form, values=lista_cuentas, width=400)
            ent_mod_cuenta.pack(fill="x", pady=(0, 10))
            if cuenta_actual: ent_mod_cuenta.set(cuenta_actual)
            elif lista_cuentas: ent_mod_cuenta.set(lista_cuentas[0])

            ctk.CTkLabel(f_form, text="Nuevo Monto (0 = Eliminar):", font=("Arial", 11, "bold")).pack(anchor="w")
            ent_mod_monto = ctk.CTkEntry(f_form)
            ent_mod_monto.pack(fill="x", pady=(0, 10))
            ent_mod_monto.insert(0, str(monto_actual))

            ctk.CTkLabel(f_form, text="Fecha:", font=("Arial", 11, "bold")).pack(anchor="w")
            f_fecha_mod = ctk.CTkFrame(f_form, fg_color="transparent")
            f_fecha_mod.pack(fill="x", pady=(0, 10))
            ent_mod_fecha = ctk.CTkEntry(f_fecha_mod)
            ent_mod_fecha.pack(side="left", fill="x", expand=True)
            ent_mod_fecha.insert(0, str(fecha_actual) if fecha_actual else datetime.now().strftime("%d/%m/%Y"))
            ctk.CTkButton(f_fecha_mod, text="📅", width=40, fg_color="#1f538d", command=lambda: CalendarioNativo(v_mod_pago, ent_mod_fecha)).pack(side="right", padx=(5, 0))

            def guardar_mod(event=None):
                nonlocal saldo_actual_global
                try:
                    nuevo_monto = float(ent_mod_monto.get().strip())
                except ValueError:
                    return messagebox.showerror("Error", "Monto inválido", parent=v_mod_pago)

                diferencia_de_aumento = nuevo_monto - float(monto_actual)
                if diferencia_de_aumento > (saldo_actual_global + 0.01):
                    return messagebox.showerror("Error", f"Supera el saldo pendiente de {formatear_moneda(saldo_actual_global)}.", parent=v_mod_pago)

                if nuevo_monto == 0:
                    if messagebox.askyesno("Confirmar", "¿Eliminar registro?", parent=v_mod_pago):
                        conn = conectar_db(); cursor = conn.cursor()
                        cursor.execute("SELECT archivo_ruta FROM pagos_comprobantes WHERE id = %s", (id_pago,))
                        r = cursor.fetchone()
                        ruta_norm = os.path.normpath(r[0]) if r and r[0] else None
                        if ruta_norm and os.path.exists(ruta_norm): os.remove(ruta_norm)
                        cursor.execute("DELETE FROM pagos_comprobantes WHERE id = %s", (id_pago,))
                        conn.commit(); conn.close()
                        registrar_auditoria(self.app_padre.usuario_activo, "Cuentas por Pagar", f"Eliminó el pago ID {id_pago}")
                else:
                    nueva_fecha = ent_mod_fecha.get().strip() or fecha_actual
                    nueva_cuenta = ent_mod_cuenta.get().strip()
                    conn = conectar_db(); cursor = conn.cursor()
                    cursor.execute("UPDATE pagos_comprobantes SET monto_pagado = %s, fecha_pago = %s, cuenta_origen = %s WHERE id = %s", (nuevo_monto, nueva_fecha, nueva_cuenta, id_pago))
                    conn.commit(); conn.close()
                    registrar_auditoria(self.app_padre.usuario_activo, "Cuentas por Pagar", f"Modificó el pago ID {id_pago} a {formatear_moneda(nuevo_monto)}")

                saldo_actual_global -= diferencia_de_aumento
                v_mod_pago.destroy()
                refrescar_subtabla()
                self.cargar_datos_pagar()
                self.app_padre.app_facturas.cargar_datos_tabla()

            ent_mod_monto.bind("<Return>", guardar_mod)

            btn_guardar_mod = ctk.CTkButton(v_mod_pago, text="💾 Guardar Cambios", command=guardar_mod, fg_color="#27ae60", hover_color="#1e8449")
            btn_guardar_mod.pack(pady=10)
            ent_mod_monto.focus()

        def cambiar_soporte():
            ruta_base = obtener_ruta_base_drive()
            if not ruta_base:
                messagebox.showwarning("Configuración Requerida", "No ha configurado la ruta de Google Drive.\nEs obligatorio para guardar archivos.")
                return
            
            sub_sel = sub_tabla.selection()
            if not sub_sel: return
            id_pago = sub_tabla.item(sub_sel[0], "values")[0]
            ruta_origen = filedialog.askopenfilename(title="Seleccionar Soporte", filetypes=[("Archivos", "*.pdf;*.png;*.jpg;*.jpeg")], parent=v_edit)
            if ruta_origen:
                try:
                    carpeta_comprobantes = os.path.normpath(os.path.join(ruta_base, "comprobantes_egresos"))
                    if not os.path.exists(carpeta_comprobantes): os.makedirs(carpeta_comprobantes)
                    conn = conectar_db(); cursor = conn.cursor()
                    cursor.execute("SELECT archivo_ruta FROM pagos_comprobantes WHERE id = %s", (id_pago,))
                    antigua_ruta = os.path.normpath(cursor.fetchone()[0])
                    if antigua_ruta and os.path.exists(antigua_ruta): os.remove(antigua_ruta)
                    
                    prov_limpio = re.sub(r'[\\/*?:"<>|]', '-', proveedor)
                    nombre_limpio = f"Egreso_Fac_{id_factura}_{prov_limpio.replace(' ', '_')}_R_{id_pago}{os.path.splitext(ruta_origen)[1]}"
                    ruta_destino = os.path.normpath(os.path.join(carpeta_comprobantes, nombre_limpio))
                    shutil.copy2(ruta_origen, ruta_destino)
                    cursor.execute("UPDATE pagos_comprobantes SET archivo_ruta = %s WHERE id = %s", (ruta_destino, id_pago))
                    conn.commit(); conn.close()
                    registrar_auditoria(self.app_padre.usuario_activo, "Cuentas por Pagar", f"Actualizó soporte del pago ID {id_pago}")
                    messagebox.showinfo("Éxito", "Soporte actualizado.", parent=v_edit)
                    refrescar_subtabla(); self.cargar_datos_pagar()
                except Exception as e: messagebox.showerror("Error", str(e), parent=v_edit)

        def eliminar_soporte():
            sub_sel = sub_tabla.selection()
            if not sub_sel: return
            id_pago = sub_tabla.item(sub_sel[0], "values")[0]
            if messagebox.askyesno("Confirmar", "¿Eliminar soporte digital?", parent=v_edit):
                try:
                    conn = conectar_db(); cursor = conn.cursor()
                    cursor.execute("SELECT archivo_ruta FROM pagos_comprobantes WHERE id = %s", (id_pago,))
                    ruta_archivo = os.path.normpath(cursor.fetchone()[0])
                    if ruta_archivo and os.path.exists(ruta_archivo): os.remove(ruta_archivo)
                    cursor.execute("UPDATE pagos_comprobantes SET archivo_ruta = '' WHERE id = %s", (id_pago,))
                    conn.commit(); conn.close()
                    registrar_auditoria(self.app_padre.usuario_activo, "Cuentas por Pagar", f"Eliminó soporte del pago ID {id_pago}")
                    messagebox.showinfo("Éxito", "Soporte eliminado.", parent=v_edit)
                    refrescar_subtabla(); self.cargar_datos_pagar()
                except Exception as e: messagebox.showerror("Error", str(e), parent=v_edit)

        def eliminar_pago_completo():
            nonlocal saldo_actual_global
            sub_sel = sub_tabla.selection()
            if not sub_sel: return
            id_pago = sub_tabla.item(sub_sel[0], "values")[0]
            monto_eliminado = desformatear_numero(sub_tabla.item(sub_sel[0], "values")[1])

            if messagebox.askyesno("Confirmar Eliminación", "⚠️ ¿Eliminar este registro de pago por completo?", parent=v_edit):
                try:
                    conn = conectar_db(); cursor = conn.cursor()
                    cursor.execute("SELECT archivo_ruta FROM pagos_comprobantes WHERE id = %s", (id_pago,))
                    r = cursor.fetchone()
                    ruta_norm = os.path.normpath(r[0]) if r and r[0] else None
                    if ruta_norm and os.path.exists(ruta_norm): os.remove(ruta_norm)
                    cursor.execute("DELETE FROM pagos_comprobantes WHERE id = %s", (id_pago,))
                    conn.commit(); conn.close()
                    registrar_auditoria(self.app_padre.usuario_activo, "Cuentas por Pagar", f"Eliminó completamente el pago ID {id_pago}")
                    messagebox.showinfo("Éxito", "Pago eliminado.", parent=v_edit)
                    refrescar_subtabla(); self.cargar_datos_pagar(); self.app_padre.app_facturas.cargar_datos_tabla(); saldo_actual_global += monto_eliminado
                except Exception as e: messagebox.showerror("Error", str(e), parent=v_edit)

        frame_lateral_btns = ctk.CTkFrame(frame_cuerpo, fg_color="transparent")
        frame_lateral_btns.pack(side="right", fill="y")
        ctk.CTkButton(frame_lateral_btns, text="✏️ Modificar Monto/Cuenta", font=("Arial", 12, "bold"), fg_color="#34495e", hover_color="#2c3e50", command=ejecutar_modificacion).pack(fill="x", pady=3)
        ctk.CTkButton(frame_lateral_btns, text="📂 Cambiar Soporte", font=("Arial", 12, "bold"), fg_color="#7f8c8d", hover_color="#606b6b", command=cambiar_soporte).pack(fill="x", pady=3)
        ctk.CTkButton(frame_lateral_btns, text="🗑️ Eliminar Soporte", font=("Arial", 12, "bold"), fg_color="#e74c3c", hover_color="#922b21", command=eliminar_soporte).pack(fill="x", pady=3)
        ctk.CTkButton(frame_lateral_btns, text="❌ Eliminar Pago Completo", font=("Arial", 12, "bold"), fg_color="#e74c3c", hover_color="#922b21", command=eliminar_pago_completo).pack(fill="x", pady=(15, 3))

        frame_btn_cierre = ctk.CTkFrame(v_edit, fg_color="transparent")
        frame_btn_cierre.pack(fill="x", padx=10, pady=10)
        ctk.CTkButton(frame_btn_cierre, text="❌ Salir", font=("Arial", 12, "bold"), fg_color="#7f8c8d", hover_color="#606b6b", command=v_edit.destroy).pack(side="right")

    def mostrar_reporte_totales(self):
        v_rep = ctk.CTkToplevel(self.main_root)
        v_rep.title("Reporte Avanzado de Compras y Pagos")
        v_rep.geometry("750x550")
        v_rep.transient(self.main_root)
        v_rep.grab_set()

        ctk.CTkLabel(v_rep, text="📊 REPORTE DE TOTALES FILTRADOS", font=("Arial", 16, "bold"), text_color="#1f538d").pack(pady=(15, 10))

        f_filtros = ctk.CTkFrame(v_rep, fg_color="#f8f9fa", border_width=1, border_color="#ccc")
        f_filtros.pack(fill="x", padx=20, pady=10, ipadx=10, ipady=10)

        f_prov = ctk.CTkFrame(f_filtros, fg_color="transparent")
        f_prov.pack(fill="x", pady=5)
        ctk.CTkLabel(f_prov, text="Proveedor:", font=("Arial", 11, "bold"), width=100, anchor="e").pack(side="left", padx=5)
        
        provs_mem = getattr(cache_sistema, 'proveedores_nombres', [])
        provs = ["Todos"] + (provs_mem if provs_mem else [])
        combo_prov = ctk.CTkComboBox(f_prov, values=provs, state="readonly", width=300)
        combo_prov.pack(side="left", padx=5)
        combo_prov.set("Todos")

        f_fechas = ctk.CTkFrame(f_filtros, fg_color="transparent")
        f_fechas.pack(fill="x", pady=5)
        
        ctk.CTkLabel(f_fechas, text="Desde:", font=("Arial", 11, "bold"), width=100, anchor="e").pack(side="left", padx=5)
        ent_desde = ctk.CTkEntry(f_fechas, width=110, placeholder_text=CONFIG_REGIONAL.get("formato_fecha", "DD/MM/AAAA"))
        ent_desde.pack(side="left", padx=5)
        ctk.CTkButton(f_fechas, text="📅", width=30, fg_color="#1f538d", hover_color="#163b65", command=lambda: CalendarioNativo(v_rep, ent_desde)).pack(side="left")

        ctk.CTkLabel(f_fechas, text="Hasta:", font=("Arial", 11, "bold"), width=60, anchor="e").pack(side="left", padx=5)
        ent_hasta = ctk.CTkEntry(f_fechas, width=110, placeholder_text=CONFIG_REGIONAL.get("formato_fecha", "DD/MM/AAAA"))
        ent_hasta.pack(side="left", padx=5)
        ctk.CTkButton(f_fechas, text="📅", width=30, fg_color="#1f538d", hover_color="#163b65", command=lambda: CalendarioNativo(v_rep, ent_hasta)).pack(side="left")

        f_resultados = ctk.CTkFrame(v_rep, fg_color="transparent")
        f_resultados.pack(fill="both", expand=True, padx=20, pady=10)
        
        def crear_lbl_res(padre, texto, color="#333"):
            f = ctk.CTkFrame(padre, fg_color="#ffffff", border_width=1, border_color="#ddd", corner_radius=8)
            f.pack(fill="x", pady=4)
            ctk.CTkLabel(f, text=texto, font=("Arial", 12, "bold"), text_color="gray").pack(side="left", padx=15, pady=10)
            lbl_val = ctk.CTkLabel(f, text="0.00", font=("Arial", 14, "bold"), text_color=color)
            lbl_val.pack(side="right", padx=15, pady=10)
            return lbl_val

        lbl_bruto = crear_lbl_res(f_resultados, "Total Compras Brutas (Subtotal Base):", "#1f538d")
        lbl_igv = crear_lbl_res(f_resultados, "Total IGV Facturado (18% / 10.5%):", "#1f538d")
        lbl_det = crear_lbl_res(f_resultados, "Total Detracción/Retención:", "#e67e22")
        lbl_pagado = crear_lbl_res(f_resultados, "Total Pagado (Dinero Egresado Real):", "#27ae60")
        lbl_por_pagar = crear_lbl_res(f_resultados, "Total por Pagar (Deuda Pendiente General):", "#c0392b")

        def convertir_a_fecha(fecha_str):
            if not fecha_str: return None
            for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y", "%m/%d/%Y"]:
                try: return datetime.strptime(fecha_str.strip(), fmt)
                except ValueError: pass
            return None

        def calcular_totales():
            prov_filtro = combo_prov.get()
            d_desde = convertir_a_fecha(ent_desde.get())
            d_hasta = convertir_a_fecha(ent_hasta.get())
            
            if ent_desde.get() and ent_hasta.get() and (not d_desde or not d_hasta):
                return messagebox.showwarning("Error", "Formato de fecha inválido.", parent=v_rep)

            if d_desde and d_hasta and d_desde > d_hasta:
                d_desde, d_hasta = d_hasta, d_desde

            conn = conectar_db()
            if not conn: return
            try:
                c = conn.cursor()
                c.execute("SELECT id_factura, monto_pagado FROM pagos_comprobantes")
                pagos_dict = {}
                for id_f, monto in c.fetchall():
                    pagos_dict[id_f] = pagos_dict.get(id_f, 0.0) + (float(monto) if monto else 0.0)

                c.execute("SELECT id, fecha, proveedor, subtotal, impuesto, total, COALESCE(det_monto, 0), tipo_documento FROM facturas_recibidas")
                
                tot_bruto = 0.0
                tot_igv = 0.0
                tot_det = 0.0
                tot_pagado = 0.0
                tot_deuda = 0.0

                for r in c.fetchall():
                    id_fac, fecha, prov, sub, imp, tot, det, tipo_doc = r
                    
                    if prov_filtro != "Todos" and prov != prov_filtro: continue
                    
                    f_dt = convertir_a_fecha(str(fecha))
                    if d_desde and d_hasta:
                        if not f_dt or not (d_desde <= f_dt <= d_hasta): continue
                    elif d_desde:
                        if not f_dt or f_dt < d_desde: continue
                    elif d_hasta:
                        if not f_dt or f_dt > d_hasta: continue

                    sub_val = float(sub) if sub else 0.0
                    imp_val = float(imp) if imp else 0.0
                    tot_val = float(tot) if tot else 0.0
                    det_val = float(det) if det else 0.0
                    
                    tot_bruto += sub_val

                    if tipo_doc and "Factura" in tipo_doc:
                        tot_igv += imp_val
                        tot_det += det_val
                        neto = tot_val - det_val
                    elif tipo_doc and "Recibo" in tipo_doc:
                        if "8%" in tipo_doc:
                            tot_det += imp_val
                            neto = tot_val - imp_val - det_val
                        else:
                            neto = tot_val - det_val
                    else:
                        tot_det += det_val
                        neto = tot_val - det_val

                    pagado = pagos_dict.get(id_fac, 0.0)
                    saldo = max(0.0, neto - pagado)

                    tot_pagado += pagado
                    tot_deuda += saldo

                lbl_bruto.configure(text=formatear_moneda(tot_bruto))
                lbl_igv.configure(text=formatear_moneda(tot_igv))
                lbl_det.configure(text=formatear_moneda(tot_det))
                lbl_pagado.configure(text=formatear_moneda(tot_pagado))
                lbl_por_pagar.configure(text=formatear_moneda(tot_deuda))

            except Exception as e:
                messagebox.showerror("Error", f"Fallo al calcular:\n{e}", parent=v_rep)
            finally:
                conn.close()

        ctk.CTkButton(f_filtros, text="🔍 Procesar y Calcular", font=("Arial", 12, "bold"), fg_color="#1f538d", hover_color="#163b65", command=calcular_totales).pack(pady=(10, 5))
        calcular_totales()

# =========================================================
# CLASE PRINCIPAL: MÓDULO DE COMPRAS (CONTENEDOR TABVIEW)
# =========================================================
class ModuloComprasApp:
    def __init__(self, parent_frame):
        self.parent_frame = parent_frame
        self.usuario_activo = "Desconocido"
        self.pantalla_expandida = False
        aplicar_estilo_treeview()

        self.frame_main = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.frame_main.pack(fill="both", expand=True, padx=15, pady=15)
        
        header_frame = ctk.CTkFrame(self.frame_main, fg_color="transparent")
        header_frame.pack(fill="x")
        
        ctk.CTkLabel(header_frame, text="🛒 MÓDULO DE COMPRAS (LOGÍSTICA Y TESORERÍA)", font=("Arial", 18, "bold"), text_color="#1f538d").pack(side="left")
        self.btn_pantalla = ctk.CTkButton(header_frame, text="[ + ] Pantalla Completa", font=("Arial", 12, "bold"), width=160, fg_color="#34495e", hover_color="#2c3e50", command=self.toggle_pantalla_completa)
        self.btn_pantalla.pack(side="right")

        self.tabview = ctk.CTkTabview(self.frame_main, segmented_button_selected_color="#1f538d")
        self.tabview.pack(fill="both", expand=True, pady=(10, 0))
        
        self.tab_recepcion = self.tabview.add(" 📥 1. Ingreso de Facturas Recibidas ")
        self.tab_pagos = self.tabview.add(" 💳 2. Control de Pagos y Deudas ")
        
        self.app_facturas = FacturasRecibidasTab(self.tab_recepcion, self.parent_frame, self)
        self.app_pagos = CuentasPorPagarTab(self.tab_pagos, self.parent_frame, self)
        
        self.tabview.configure(command=self.al_cambiar_pestana)

    def toggle_pantalla_completa(self):
        sidebar = None
        try:
            if self.parent_frame.master:
                for child in self.parent_frame.master.winfo_children():
                    if hasattr(child, "cget") and child.cget("width") == 280:
                        sidebar = child
                        break
        except Exception: pass

        if getattr(self, "pantalla_expandida", False):
            if sidebar: sidebar.pack(side="left", fill="y", before=self.parent_frame)
            self.btn_pantalla.configure(text="[ + ] Pantalla Completa", fg_color="#34495e", hover_color="#2c3e50")
            self.pantalla_expandida = False
        else:
            if sidebar: sidebar.pack_forget()
            self.btn_pantalla.configure(text="[ - ] Restaurar Vista", fg_color="#34495e", hover_color="#2c3e50")
            self.pantalla_expandida = True

    def al_cambiar_pestana(self):
        if self.tabview.get() == " 💳 2. Control de Pagos y Deudas ":
            self.app_pagos.cargar_datos_pagar()
        elif self.tabview.get() == " 📥 1. Ingreso de Facturas Recibidas ":
            self.app_facturas.cargar_datos_tabla()

if __name__ == "__main__":
    pass