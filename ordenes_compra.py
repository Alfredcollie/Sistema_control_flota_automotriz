# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import customtkinter as ctk
from datetime import datetime
import sys
import os
import subprocess
import webbrowser
import urllib.parse
import json
import re
import shutil

from conexion import conectar_db, registrar_auditoria

try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.utils import ImageReader
    REPORTLAB_DISPONIBLE = True
except ImportError:
    REPORTLAB_DISPONIBLE = False

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

# =========================================================
# 🚀 FUNCIONES MULTIPLATAFORMA Y EXTRACCIÓN DE DATOS
# =========================================================
def abrir_documento(ruta):
    try:
        ruta_abs = os.path.abspath(ruta)
        if sys.platform == "win32": os.startfile(ruta_abs)
        elif sys.platform == "darwin": subprocess.call(["open", ruta_abs])
        else: subprocess.call(["xdg-open", ruta_abs])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el archivo o carpeta:\n{e}")

def copiar_archivo_portapapeles(ruta):
    try:
        ruta_absoluta = os.path.abspath(ruta)
        if sys.platform == "darwin": 
            os.system(f'osascript -e \'set the clipboard to POSIX file "{ruta_absoluta}"\'')
        elif sys.platform == "win32": 
            os.system(f'powershell -command "Set-Clipboard -Path \'{ruta_absoluta}\'"')
    except Exception as e: print("Error copiando al portapapeles:", e)

def maximizar_ventana(ventana):
    if sys.platform == "win32":
        try: ventana.state("zoomed")
        except: pass
    else:
        try:
            w = ventana.winfo_screenwidth(); h = ventana.winfo_screenheight()
            ventana.geometry(f"{w}x{h}+0+0")
        except: pass

def obtener_ruta_logo():
    try:
        if os.path.exists("config_local.json"):
            with open("config_local.json", "r", encoding="utf-8") as f:
                return json.load(f).get("ruta_logo_cotizacion", "")
    except Exception: pass
    return ""

def obtener_telefono_proveedor(prov):
    conn = conectar_db()
    if not conn: return ""
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT whatsapp FROM proveedores WHERE nombre = %s", (prov,))
        res = cursor.fetchone()
        if res and res[0]:
            digits = re.sub(r'\D', '', str(res[0]))
            if len(digits) == 9 and digits.startswith("9"): return f"51{digits}"
            elif len(digits) >= 11: return digits
    except Exception: pass
    finally: conn.close()
    return ""

def obtener_email_proveedor(prov):
    conn = conectar_db()
    if not conn: return ""
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT correo FROM proveedores WHERE nombre = %s", (prov,))
        res = cursor.fetchone()
        if res and res[0]: return str(res[0]).strip()
    except Exception: pass
    finally: conn.close()
    return ""

# =========================================================
# CLASE PRINCIPAL: ÓRDENES DE SERVICIO Y REPARACIÓN
# =========================================================
class OrdenesCompraApp:
    def __init__(self, parent_frame, usuario_activo="Desconocido"):
        self.parent_frame = parent_frame
        self.usuario_activo = usuario_activo
        self.ultima_ruta_pdf = ""
        self.orden_columnas_hist = {}
        
        self.inicializar_bd()
        self.crear_interfaz()

    def inicializar_bd(self):
        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS ordenes_servicio_flota (
                    id SERIAL PRIMARY KEY, 
                    numero_orden VARCHAR(100), 
                    placa VARCHAR(50), 
                    vehiculo_info VARCHAR(255), 
                    proveedor VARCHAR(255), 
                    servicio VARCHAR(150), 
                    descripcion TEXT, 
                    costo_total NUMERIC, 
                    fecha_emision VARCHAR(50), 
                    pdf_ruta TEXT, 
                    version INTEGER DEFAULT 0,
                    estado VARCHAR(50) DEFAULT 'Activa'
                )
            """)
            conn.commit()
        except Exception as e: print("Error DB Ordenes Flota:", e)
        finally: conn.close()

    def abrir_carpeta_anuladas(self):
        carpeta = "ordenes_anuladas"
        if not os.path.exists(carpeta):
            os.makedirs(carpeta)
        abrir_documento(carpeta)

    def crear_interfaz(self):
        self.frame_main = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.frame_main.pack(fill="both", expand=True, padx=15, pady=15)
        
        ctk.CTkLabel(self.frame_main, text="🛠️ GESTIÓN DE ÓRDENES DE SERVICIO Y REPARACIÓN", font=("Arial", 18, "bold"), text_color="#1f538d").pack(anchor="w", pady=(0, 10))

        self.tabview = ctk.CTkTabview(self.frame_main, segmented_button_selected_color="#1f538d")
        self.tabview.pack(fill="both", expand=True)
        
        self.tab_generar = self.tabview.add(" ➕ 1. Generar Nueva Orden ")
        self.tab_historial = self.tabview.add(" 🗂️ 2. Historial de Órdenes ")
        
        self.crear_pestaña_generar()
        self.crear_pestaña_historial()
        
        self.tabview.configure(command=self.al_cambiar_pestana)

    def crear_pestaña_generar(self):
        f_top = ctk.CTkFrame(self.tab_generar, fg_color="#f8f9fa", border_width=1, border_color="#e0e0e0", corner_radius=8)
        f_top.pack(fill="x", pady=(0, 10), ipadx=10, ipady=10)

        ctk.CTkLabel(f_top, text="1. Vehículo (Placa):", font=("Arial", 12, "bold")).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.cmb_placa = ctk.CTkComboBox(f_top, width=350, state="readonly")
        self.cmb_placa.grid(row=0, column=1, padx=10, pady=10)

        ctk.CTkLabel(f_top, text="2. Proveedor Asignado:", font=("Arial", 12, "bold")).grid(row=0, column=2, padx=10, pady=10, sticky="w")
        self.cmb_proveedor = ctk.CTkComboBox(f_top, width=300, state="readonly")
        self.cmb_proveedor.grid(row=0, column=3, padx=10, pady=10)

        f_centro = ctk.CTkFrame(self.tab_generar, fg_color="transparent")
        f_centro.pack(fill="both", expand=True)
        
        f_izq = ctk.CTkFrame(f_centro, corner_radius=10)
        f_izq.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        ctk.CTkLabel(f_izq, text="Detalles del Trabajo Solicitado", font=("Arial", 14, "bold"), text_color="#1f538d").pack(anchor="w", padx=15, pady=15)
        
        f_serv = ctk.CTkFrame(f_izq, fg_color="transparent")
        f_serv.pack(fill="x", padx=15, pady=5)
        
        ctk.CTkLabel(f_serv, text="Tipo de Servicio:", font=("Arial", 12, "bold")).pack(side="left")
        opciones_servicio = ["Mantenimiento Preventivo", "Mantenimiento Correctivo", "Cambio de Aceite / Filtros", "Llantas / Alineación / Balanceo", "Frenos / Suspensión", "Sistema Eléctrico / Batería", "Planchado y Pintura", "Lavado / Limpieza", "Otro"]
        self.cmb_servicio = ctk.CTkComboBox(f_serv, values=opciones_servicio, width=250)
        self.cmb_servicio.pack(side="left", padx=15)
        
        ctk.CTkLabel(f_serv, text="Costo Acordado (S/):", font=("Arial", 12, "bold")).pack(side="left", padx=(20, 5))
        self.ent_costo = ctk.CTkEntry(f_serv, width=120, placeholder_text="0.00")
        self.ent_costo.pack(side="left")

        ctk.CTkLabel(f_izq, text="Descripción Detallada del Trabajo a Realizar:", font=("Arial", 12, "bold")).pack(anchor="w", padx=15, pady=(15, 5))
        self.txt_descripcion = ctk.CTkTextbox(f_izq, height=150, border_width=1)
        self.txt_descripcion.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        f_botones = ctk.CTkFrame(self.tab_generar, fg_color="transparent")
        f_botones.pack(fill="x", pady=5)
        
        ctk.CTkButton(f_botones, text="📄 Solo Generar Orden", font=("Arial", 13, "bold"), fg_color="#1f538d", hover_color="#163b65", height=40, command=lambda: self.generar_orden("solo")).pack(side="left", expand=True, fill="x", padx=3)
        ctk.CTkButton(f_botones, text="💬 Generar y Enviar WA", font=("Arial", 13, "bold"), fg_color="#27ae60", hover_color="#1e8449", height=40, command=lambda: self.generar_orden("whatsapp")).pack(side="left", expand=True, fill="x", padx=3)
        ctk.CTkButton(f_botones, text="📧 Generar y Enviar Mail", font=("Arial", 13, "bold"), fg_color="#e67e22", hover_color="#d35400", height=40, command=lambda: self.generar_orden("email")).pack(side="left", expand=True, fill="x", padx=3)

        self.cargar_datos_combos()

    def crear_pestaña_historial(self):
        f_busqueda = ctk.CTkFrame(self.tab_historial, fg_color="transparent")
        f_busqueda.pack(fill="x", padx=10, pady=(10, 0))
        ctk.CTkLabel(f_busqueda, text="🔍 Buscar:", font=("Arial", 11, "bold")).pack(side="left", padx=(0, 5))
        self.ent_buscar_historial = ctk.CTkEntry(f_busqueda, placeholder_text="Filtrar por N° Orden, Placa, Proveedor o Servicio...")
        self.ent_buscar_historial.pack(side="left", fill="x", expand=True)
        self.ent_buscar_historial.bind("<KeyRelease>", lambda e: self.cargar_historial_ordenes())

        f_tabla = ctk.CTkFrame(self.tab_historial, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True, pady=10)

        self.tree_historial = ttk.Treeview(f_tabla, columns=("id", "num_orden", "placa", "servicio", "proveedor", "fecha", "total"), show="headings")
        
        self.tree_historial.heading("id", text="ID ↕", command=lambda: self.ordenar_por_columna("id", True))
        self.tree_historial.heading("num_orden", text="N° Orden ↕", command=lambda: self.ordenar_por_columna("num_orden", False))
        self.tree_historial.heading("placa", text="Placa Unidad ↕", command=lambda: self.ordenar_por_columna("placa", False))
        self.tree_historial.heading("servicio", text="Servicio Solicitado ↕", command=lambda: self.ordenar_por_columna("servicio", False))
        self.tree_historial.heading("proveedor", text="Proveedor ↕", command=lambda: self.ordenar_por_columna("proveedor", False))
        self.tree_historial.heading("fecha", text="Emisión ↕", command=lambda: self.ordenar_por_columna("fecha", False))
        self.tree_historial.heading("total", text="Total S/ ↕", command=lambda: self.ordenar_por_columna("total", True))
        
        self.tree_historial.column("id", width=40, anchor="center")
        self.tree_historial.column("num_orden", width=120, anchor="center")
        self.tree_historial.column("placa", width=100, anchor="center")
        self.tree_historial.column("servicio", width=180, anchor="w")
        self.tree_historial.column("proveedor", width=180, anchor="w")
        self.tree_historial.column("fecha", width=110, anchor="center")
        self.tree_historial.column("total", width=90, anchor="e")
        
        scroll_y = ctk.CTkScrollbar(f_tabla, command=self.tree_historial.yview)
        self.tree_historial.configure(yscrollcommand=scroll_y.set)
        self.tree_historial.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y", padx=5)

        self.tree_historial.bind("<Double-1>", lambda e: self.abrir_ventana_modificacion())

        f_botones_hist = ctk.CTkFrame(self.tab_historial, fg_color="transparent")
        f_botones_hist.pack(fill="x", pady=10)
        
        ctk.CTkButton(f_botones_hist, text="📂 Ver Anuladas", font=("Arial", 12, "bold"), fg_color="#8e44ad", hover_color="#732d91", command=self.abrir_carpeta_anuladas).pack(side="left", padx=5)
        ctk.CTkButton(f_botones_hist, text="🔄 Actualizar", font=("Arial", 12, "bold"), fg_color="#7f8c8d", hover_color="#606b6b", command=self.cargar_historial_ordenes).pack(side="left", padx=5)
        
        ctk.CTkButton(f_botones_hist, text="🗑️ Anular Orden", font=("Arial", 12, "bold"), fg_color="#e74c3c", hover_color="#c0392b", command=self.eliminar_orden_historial).pack(side="right", padx=5)
        ctk.CTkButton(f_botones_hist, text="✏️ Modificar y Reenviar", font=("Arial", 12, "bold"), fg_color="#d35400", hover_color="#a84300", command=self.abrir_ventana_modificacion).pack(side="right", padx=5)
        ctk.CTkButton(f_botones_hist, text="👁️ Ver PDF Generado", font=("Arial", 12, "bold"), fg_color="#34495e", hover_color="#2c3e50", command=self.ver_pdf_historial).pack(side="right", padx=5)

    def ordenar_por_columna(self, columna, es_numerico):
        elementos = [(self.tree_historial.set(item, columna), item) for item in self.tree_historial.get_children("")]
        ascendente = self.orden_columnas_hist.get(columna, True)
        self.orden_columnas_hist[columna] = not ascendente
        
        if es_numerico:
            def parsear_numero(val):
                try: return float(val.replace(",", "").replace("S/", "").strip())
                except ValueError: return 0.0
            elementos.sort(key=lambda el: parsear_numero(el[0]), reverse=not ascendente)
        else:
            elementos.sort(key=lambda el: el[0].lower() if el[0] else "", reverse=not ascendente)
            
        for index, (_, item) in enumerate(elementos):
            self.tree_historial.move(item, "", index)

    def al_cambiar_pestana(self):
        if self.tabview.get() == " 🗂️ 2. Historial de Órdenes ":
            self.cargar_historial_ordenes()
        else:
            self.cargar_datos_combos()

    def cargar_datos_combos(self):
        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            
            # Cargar Vehículos
            try:
                cursor.execute("SELECT placa, marca, modelo FROM flota_vehiculos ORDER BY placa ASC")
                vehiculos = [f"{r[0]} | {r[1]} {r[2]}" for r in cursor.fetchall()]
                if vehiculos:
                    self.cmb_placa.configure(values=vehiculos); self.cmb_placa.set("Seleccione una unidad...")
                else:
                    self.cmb_placa.configure(values=["No hay vehículos registrados"]); self.cmb_placa.set("No hay vehículos registrados")
            except Exception: pass

            # Cargar Proveedores
            try:
                cursor.execute("SELECT nombre FROM proveedores ORDER BY nombre ASC")
                proveedores = [str(r[0]) for r in cursor.fetchall()]
                if proveedores:
                    self.cmb_proveedor.configure(values=proveedores); self.cmb_proveedor.set("Seleccione proveedor...")
                else:
                    self.cmb_proveedor.configure(values=["No hay proveedores registrados"]); self.cmb_proveedor.set("No hay proveedores registrados")
            except Exception: pass
            
        except Exception as e: print("Error combos:", e)
        finally: conn.close()

    def cargar_historial_ordenes(self):
        for item in self.tree_historial.get_children(): self.tree_historial.delete(item)
        
        filtro = ""
        if hasattr(self, 'ent_buscar_historial'):
            filtro = self.ent_buscar_historial.get().strip().lower()

        conn = conectar_db()
        if conn:
            try:
                cursor = conn.cursor()
                query = """
                    SELECT id, numero_orden, version, placa, vehiculo_info, proveedor, servicio, fecha_emision, costo_total 
                    FROM ordenes_servicio_flota 
                    WHERE estado != 'Anulada' OR estado IS NULL
                    ORDER BY id DESC
                """
                cursor.execute(query)
                for r in cursor.fetchall():
                    num_base = r[1] if r[1] else f"OS-{r[3]}-L{r[0]}"
                    ver = r[2] or 0
                    n_imprimir = num_base if ver == 0 else f"{num_base}-{ver}"
                    
                    try: costo_val = float(r[8]) if r[8] else 0.0
                    except: costo_val = 0.0
                    
                    row_vals = (r[0], n_imprimir, r[3], r[6], r[5], r[7], f"{costo_val:,.2f}")
                    
                    if filtro:
                        coincide = False
                        for val in row_vals:
                            if val is not None and filtro in str(val).lower():
                                coincide = True
                                break
                        if not coincide:
                            continue
                            
                    self.tree_historial.insert("", tk.END, values=row_vals)
            except Exception as e: print("Error al cargar historial:", e)
            finally: conn.close()

    def ver_pdf_historial(self):
        sel = self.tree_historial.selection()
        if not sel: return messagebox.showwarning("Aviso", "Seleccione una orden del historial.")
        id_orden = self.tree_historial.item(sel[0], "values")[0]
        conn = conectar_db()
        if conn:
            try:
                cursor = conn.cursor()
                cursor.execute("SELECT pdf_ruta FROM ordenes_servicio_flota WHERE id = %s", (id_orden,))
                ruta = cursor.fetchone()
                if ruta and ruta[0] and os.path.exists(ruta[0]): abrir_documento(ruta[0])
                else: messagebox.showerror("Error", "El archivo PDF no se encuentra en la ruta especificada.")
            except Exception: pass
            finally: conn.close()

    def eliminar_orden_historial(self):
        sel = self.tree_historial.selection()
        if not sel: return messagebox.showwarning("Aviso", "Seleccione una orden del historial para anular.")
        
        id_orden = self.tree_historial.item(sel[0], "values")[0]
        n_orden = self.tree_historial.item(sel[0], "values")[1]
        prov = self.tree_historial.item(sel[0], "values")[4]
        
        msg = f"¿Estás seguro de que deseas anular y archivar la orden {n_orden} de {prov}?\n\nLa orden pasará a la carpeta de 'Anuladas'."
        if messagebox.askyesno("Confirmar Anulación", msg):
            conn = conectar_db()
            if not conn: return
            try:
                cursor = conn.cursor()
                cursor.execute("UPDATE ordenes_servicio_flota SET estado = 'Anulada' WHERE id=%s", (id_orden,))
                conn.commit()
                
                cursor.execute("SELECT pdf_ruta FROM ordenes_servicio_flota WHERE id = %s", (id_orden,))
                ruta_pdf = cursor.fetchone()[0]
                
                carpeta_anuladas = "ordenes_anuladas"
                if not os.path.exists(carpeta_anuladas): os.makedirs(carpeta_anuladas)
                
                if ruta_pdf and os.path.exists(ruta_pdf):
                    try: shutil.move(ruta_pdf, os.path.join(carpeta_anuladas, os.path.basename(ruta_pdf)))
                    except Exception as e: print(f"No se pudo archivar {ruta_pdf}: {e}")
                
                registrar_auditoria(self.usuario_activo, "Órdenes Servicio", f"Anuló la O/S {n_orden} de {prov}")
                messagebox.showinfo("Éxito", "Orden anulada y archivada correctamente.")
                
                self.cargar_historial_ordenes()
            except Exception as e: messagebox.showerror("Error", str(e))
            finally: conn.close()

    def fabricar_pdf(self, placa, vehiculo_info, prov, servicio, detalles, fecha, total_orden, num_orden_imprimir):
        try: total_orden = float(total_orden)
        except ValueError: total_orden = 0.0
        
        carpeta_destino = "ordenes_generadas"
        if not os.path.exists(carpeta_destino): os.makedirs(carpeta_destino)
        marca_tiempo = datetime.now().strftime("%H%M%S")
        nombre_archivo = os.path.join(carpeta_destino, f"Orden_Servicio_{placa}_{marca_tiempo}.pdf")
        
        c = canvas.Canvas(nombre_archivo, pagesize=letter)
        ancho_util = 532.0
        y_cursor = 740.0
        ruta_logo = obtener_ruta_logo()
        
        if ruta_logo and os.path.exists(ruta_logo):
            try:
                img = ImageReader(ruta_logo)
                w_orig, h_orig = img.getSize()
                alto_calculado = ancho_util * (float(h_orig) / float(w_orig))
                y_logo = 780.0 - alto_calculado
                c.drawImage(img, 40, y_logo, width=ancho_util, height=alto_calculado)
                
                c.setFont("Helvetica", 10)
                c.setFillColorRGB(0, 0, 0)
                c.drawRightString(558, y_logo + 22, "RUC. 20613989146")
                
                y_cursor = y_logo - 20.0
            except Exception:
                y_cursor = 740.0
        else:
            y_cursor = 740.0
            c.setFont("Helvetica-Bold", 10)
            c.drawRightString(572, 750.0, "RUC. 20613989146")

        c.setFont("Helvetica-Bold", 16)
        c.setFillColorRGB(0.12, 0.32, 0.55)
        c.drawString(40, y_cursor, "ORDEN DE SERVICIO Y REPARACIÓN")

        c.setFont("Helvetica", 10)
        c.setFillColorRGB(0, 0, 0)
        c.drawString(410, y_cursor, f"Fecha: {fecha}")
        c.drawString(410, y_cursor - 15, f"N° Orden: {num_orden_imprimir}")
        
        y_cursor -= 45.0
        c.setLineWidth(1)
        c.line(40, y_cursor, 572, y_cursor)
        
        y_cursor -= 20.0
        c.setFont("Helvetica-Bold", 10)
        c.drawString(40, y_cursor, "DATOS DEL PROVEEDOR:")
        c.drawString(300, y_cursor, "DATOS DEL VEHÍCULO:")
        
        c.setFont("Helvetica", 10)
        c.drawString(40, y_cursor - 15, f"Empresa / Taller: {prov}")
        
        c.drawString(300, y_cursor - 15, f"Placa / Matrícula: {placa}")
        c.drawString(300, y_cursor - 30, f"Unidad: {vehiculo_info}")
        
        y_cursor -= 55.0
        c.line(40, y_cursor, 572, y_cursor)
        
        y_cursor -= 25.0
        c.setFont("Helvetica-Bold", 11)
        c.drawString(40, y_cursor, "SERVICIO SOLICITADO:")
        c.setFont("Helvetica", 11)
        c.drawString(190, y_cursor, servicio)

        y_pos = y_cursor - 30.0
        
        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(40, y_pos, "DESCRIPCIÓN DEL TRABAJO A REALIZAR:")
        y_pos -= 20.0
        c.setFont("Helvetica", 9)
        for linea in detalles.split("\n"):
            while len(linea) > 0:
                if y_pos < 150:
                    c.showPage()
                    y_pos = 730.0
                    c.setFont("Helvetica", 9)
                c.drawString(40, y_pos, linea[:100])
                linea = linea[100:]
                y_pos -= 12.0
                
        y_pos -= 20.0
        
        subtotal = total_orden / 1.18
        igv = total_orden - subtotal

        c.line(40, y_pos + 10, 572, y_pos + 10)
        
        c.setFont("Helvetica", 10)
        c.drawString(380, y_pos - 10, "SUBTOTAL APROX:")
        c.drawString(480, y_pos - 10, f"S/ {subtotal:,.2f}")
        
        c.drawString(380, y_pos - 25, "IGV (18%):")
        c.drawString(480, y_pos - 25, f"S/ {igv:,.2f}")
        
        c.setFont("Helvetica-Bold", 12)
        c.drawString(380, y_pos - 45, "COSTO ACORDADO:")
        c.setFillColorRGB(0.75, 0.22, 0.16)
        c.drawString(500, y_pos - 45, f"S/ {total_orden:,.2f}")

        c.save()
        return nombre_archivo

    def generar_orden(self, accion="solo"):
        vehiculo_str = self.cmb_placa.get()
        prov = self.cmb_proveedor.get()
        servicio = self.cmb_servicio.get().strip()
        costo_str = self.ent_costo.get().strip()
        detalles = self.txt_descripcion.get("1.0", "end-1c").strip()
        
        if "Seleccione" in vehiculo_str or "No hay" in vehiculo_str or "Seleccione" in prov or "No hay" in prov:
            return messagebox.showwarning("Incompleto", "Seleccione una unidad y un proveedor válido.")
            
        if not REPORTLAB_DISPONIBLE: return messagebox.showerror("Librería", "Falta ReportLab para generar PDFs.")
            
        placa = vehiculo_str.split(" | ")[0].strip()
        vehiculo_info = vehiculo_str.split(" | ")[1].strip() if " | " in vehiculo_str else ""
        fecha_emision = datetime.now().strftime("%d/%m/%Y")
        
        try: costo_total = float(costo_str.replace(",", "")) if costo_str else 0.0
        except ValueError: return messagebox.showwarning("Error", "El costo ingresado no es un número válido.")

        try:
            conn = conectar_db()
            if not conn: return
            cursor = conn.cursor()
            
            cursor.execute("SELECT COUNT(id) FROM ordenes_servicio_flota")
            count = cursor.fetchone()[0]
            numero_orden_base = f"OS-{placa}-{(count + 1):03d}"
            version_inicial = 0
            
            ruta_pdf = self.fabricar_pdf(placa, vehiculo_info, prov, servicio, detalles, fecha_emision, costo_total, numero_orden_base)
            self.ultima_ruta_pdf = ruta_pdf
            
            cursor.execute("""
                INSERT INTO ordenes_servicio_flota (numero_orden, placa, vehiculo_info, proveedor, servicio, descripcion, costo_total, fecha_emision, pdf_ruta, version, estado)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (numero_orden_base, placa, vehiculo_info, prov, servicio, detalles, costo_total, fecha_emision, ruta_pdf, version_inicial, 'Activa'))
            conn.commit(); conn.close()
            
            registrar_auditoria(self.usuario_activo, "Órdenes Servicio", f"Generó O/S {numero_orden_base} para {prov}")

            if accion == "solo":
                messagebox.showinfo("Éxito", f"Orden de Servicio generada exitosamente.\nN° Orden: {numero_orden_base}")
                abrir_documento(ruta_pdf)
            else:
                datos_mod = {"prov": prov, "placa": placa, "vehiculo": vehiculo_info, "servicio": servicio, "ruta_pdf": ruta_pdf}
                if accion == "whatsapp":
                    self.ejecutar_envio_whatsapp(datos_mod, es_modificacion=False)
                elif accion == "email":
                    self.ejecutar_envio_email(datos_mod, es_modificacion=False)
            
            self.ent_costo.delete(0, tk.END)
            self.txt_descripcion.delete("1.0", tk.END)
            
        except Exception as e: messagebox.showerror("Error", f"No se pudo generar la orden:\n{e}")

    def ejecutar_envio_whatsapp(self, datos_mod, es_modificacion=False):
        prov = datos_mod["prov"]
        placa = datos_mod["placa"]
        vehiculo = datos_mod["vehiculo"]
        servicio = datos_mod["servicio"]
        ruta_pdf = datos_mod["ruta_pdf"]

        telefono_proveedor = obtener_telefono_proveedor(prov)

        if not telefono_proveedor:
            tel_manual = simpledialog.askstring("Celular no detectado", f"No se detectó automáticamente el celular de:\n{prov}\n\nIngresa su número para enviar el mensaje (ej: 987654321):", parent=self.parent_frame.winfo_toplevel())
            if tel_manual:
                digits = re.sub(r'\D', '', tel_manual)
                telefono_proveedor = f"51{digits}" if len(digits) == 9 and digits.startswith("9") else digits

        if es_modificacion:
            mensaje = (
                f"Hola {prov},\n\n"
                f"Te compartimos la *Orden de Servicio ACTUALIZADA* para nuestra unidad con placa *{placa}* ({vehiculo}).\n\n"
                f"⚠️ *NOTA LOGÍSTICA: Hubo cambios en la orden original.*\n\n"
                f"🛠️ *Servicio Requerido:* {servicio}\n\n"
                f"Por favor, revisa el archivo PDF adjunto para validar las actualizaciones del trabajo solicitado.\n"
                f"¡Quedamos atentos a tu confirmación!"
            )
        else:
            mensaje = (
                f"Hola {prov},\n\n"
                f"Te compartimos la Orden de Servicio Oficial para nuestra unidad con placa *{placa}* ({vehiculo}).\n\n"
                f"🛠️ *Servicio Requerido:* {servicio}\n\n"
                f"Por favor, revisa el archivo PDF adjunto con las especificaciones y descripción del trabajo a realizar.\n"
                f"¡Quedamos atentos a tu confirmación!"
            )
        
        mensaje_codificado = urllib.parse.quote(mensaje)
        respuesta = messagebox.askyesnocancel("WhatsApp", "¿Abrir WhatsApp de Escritorio?\n\n[Sí] = App de Escritorio\n[No] = WhatsApp Web\n[Cancelar] = Cancelar")
        if respuesta is None: return
            
        if telefono_proveedor:
            url_whatsapp = f"{'whatsapp://send' if respuesta else 'https://api.whatsapp.com/send'}?phone={telefono_proveedor}&text={mensaje_codificado}"
        else:
            url_whatsapp = f"{'whatsapp://send' if respuesta else 'https://web.whatsapp.com/send'}?text={mensaje_codificado}"
        
        try:
            copiar_archivo_portapapeles(ruta_pdf)
            messagebox.showinfo("¡Listo!", f"PDF copiado al portapapeles.\n\nAbriendo el chat de {prov}...\nHaz clic en la caja de mensaje y presiona Pegar (Ctrl+V / Cmd+V).")
            webbrowser.open(url_whatsapp)
        except Exception as e: messagebox.showerror("Error", str(e))

    def ejecutar_envio_email(self, datos_mod, es_modificacion=False):
        prov = datos_mod["prov"]
        placa = datos_mod["placa"]
        vehiculo = datos_mod["vehiculo"]
        servicio = datos_mod["servicio"]
        ruta_pdf = datos_mod["ruta_pdf"]

        email_prov = obtener_email_proveedor(prov)

        if not email_prov:
            email_prov = simpledialog.askstring("Correo no detectado", f"No se detectó automáticamente el correo de:\n{prov}\n\nIngresa su correo electrónico:", parent=self.parent_frame.winfo_toplevel())
            if not email_prov: return 

        if es_modificacion:
            asunto = f"ACTUALIZACIÓN: Orden de Servicio Unidad {placa}"
            cuerpo = (
                f"Hola {prov},\n\n"
                f"Te compartimos la Orden de Servicio ACTUALIZADA para nuestra unidad con placa {placa} ({vehiculo}).\n\n"
                f"NOTA LOGÍSTICA: Hubo cambios en la orden original.\n"
                f"- Servicio Requerido: {servicio}\n\n"
                f"Por favor, revisa el archivo PDF que adjuntamos para validar las actualizaciones del trabajo solicitado.\n"
                f"Quedamos atentos a tu confirmación."
            )
        else:
            asunto = f"Orden de Servicio Automotriz - Unidad {placa}"
            cuerpo = (
                f"Hola {prov},\n\n"
                f"Te compartimos la Orden de Servicio Oficial para nuestra unidad con placa {placa} ({vehiculo}).\n\n"
                f"- Servicio Requerido: {servicio}\n\n"
                f"Por favor, revisa el archivo PDF que adjuntamos con las especificaciones técnicas completas del trabajo a realizar.\n"
                f"Quedamos atentos a tu confirmación."
            )
        
        asunto_enc = urllib.parse.quote(asunto)
        cuerpo_enc = urllib.parse.quote(cuerpo)
        mailto_url = f"mailto:{email_prov}?subject={asunto_enc}&body={cuerpo_enc}"

        try:
            copiar_archivo_portapapeles(ruta_pdf)
            messagebox.showinfo("¡Listo para enviar!", f"El PDF de la Orden fue copiado al portapapeles.\n\nSe abrirá tu gestor de correos ({email_prov}).\n\n1. Haz clic derecho y selecciona 'Pegar' (o presiona Ctrl+V) para adjuntar el PDF.\n2. Haz clic en Enviar.")
            webbrowser.open(mailto_url)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def abrir_ventana_modificacion(self):
        sel = self.tree_historial.selection()
        if not sel: return messagebox.showwarning("Aviso", "Seleccione una orden del historial para modificar.")
        
        id_orden = self.tree_historial.item(sel[0], "values")[0]
        
        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT numero_orden, placa, vehiculo_info, proveedor, servicio, descripcion, costo_total, pdf_ruta, version FROM ordenes_servicio_flota WHERE id=%s", (id_orden,))
            ord_data = cursor.fetchone()
            if not ord_data: return
            
            num_orden_db, placa_db, vehiculo_db, prov_db, serv_db, desc_db, costo_db, ruta_pdf_antigua, version_db = ord_data
                
        except Exception as e:
            return messagebox.showerror("Error", str(e))
        finally: conn.close()

        v_mod = ctk.CTkToplevel(self.parent_frame.winfo_toplevel())
        v_mod.title(f"Modificar Orden de Unidad {placa_db}")
        v_mod.geometry("500x550")
        v_mod.grab_set()

        ctk.CTkLabel(v_mod, text="✏️ Modificar Trabajo y/o Costos", font=("Arial", 16, "bold"), text_color="#1f538d").pack(pady=15)
        
        f_form = ctk.CTkFrame(v_mod, fg_color="transparent")
        f_form.pack(fill="both", expand=True, padx=20)
        
        ctk.CTkLabel(f_form, text="Tipo de Servicio:", font=("Arial", 11, "bold")).pack(anchor="w", pady=(5,2))
        ent_serv_mod = ctk.CTkEntry(f_form, width=400)
        ent_serv_mod.pack(fill="x", pady=2); ent_serv_mod.insert(0, serv_db)
        
        ctk.CTkLabel(f_form, text="Costo Total (S/):", font=("Arial", 11, "bold")).pack(anchor="w", pady=(10,2))
        ent_costo_mod = ctk.CTkEntry(f_form, width=400)
        ent_costo_mod.pack(fill="x", pady=2); ent_costo_mod.insert(0, str(costo_db))
        
        ctk.CTkLabel(f_form, text="Descripción Detallada:", font=("Arial", 11, "bold")).pack(anchor="w", pady=(10,2))
        txt_det_mod = ctk.CTkTextbox(f_form, height=150); txt_det_mod.pack(fill="x", pady=2)
        txt_det_mod.insert("1.0", desc_db)

        def ejecutar_modificacion(accion="solo"):
            n_serv = ent_serv_mod.get().strip()
            n_costo_str = ent_costo_mod.get().strip()
            n_det = txt_det_mod.get("1.0", "end-1c").strip()
            n_fecha_emision = datetime.now().strftime("%d/%m/%Y (Modif.)")
            
            try: n_costo = float(n_costo_str.replace(",", "")) if n_costo_str else 0.0
            except ValueError: return messagebox.showwarning("Error", "Costo inválido.")
            
            n_version = (version_db or 0) + 1
            num_orden_imprimir = f"{num_orden_db}-{n_version}" if version_db == 0 else num_orden_db.rsplit('-', 1)[0] + f"-{n_version}"
            
            if ruta_pdf_antigua and os.path.exists(ruta_pdf_antigua):
                carpeta_anuladas = "ordenes_anuladas"
                if not os.path.exists(carpeta_anuladas): os.makedirs(carpeta_anuladas)
                try: shutil.move(ruta_pdf_antigua, os.path.join(carpeta_anuladas, os.path.basename(ruta_pdf_antigua)))
                except: pass
            
            try:
                n_ruta_pdf = self.fabricar_pdf(placa_db, vehiculo_db, prov_db, n_serv, n_det, n_fecha_emision, n_costo, num_orden_imprimir)
                
                c2 = conectar_db()
                cursor = c2.cursor()
                cursor.execute("""
                    UPDATE ordenes_servicio_flota SET servicio=%s, descripcion=%s, costo_total=%s, fecha_emision=%s, pdf_ruta=%s, version=%s WHERE id=%s
                """, (n_serv, n_det, n_costo, n_fecha_emision, n_ruta_pdf, n_version, id_orden))
                c2.commit(); c2.close()
                
                registrar_auditoria(self.usuario_activo, "Órdenes Servicio", f"Modificó O/S a versión {num_orden_imprimir} para unidad {placa_db}")
                self.cargar_historial_ordenes()
                v_mod.destroy()
                
                if accion == "solo":
                    messagebox.showinfo("Éxito", f"Orden modificada a la versión {n_version}.\nEl PDF anterior se movió a 'Anuladas'.")
                    abrir_documento(n_ruta_pdf)
                else:
                    datos_mod = {"prov": prov_db, "placa": placa_db, "vehiculo": vehiculo_db, "servicio": n_serv, "ruta_pdf": n_ruta_pdf}
                    if accion == "whatsapp":
                        self.ejecutar_envio_whatsapp(datos_mod, es_modificacion=True)
                    elif accion == "email":
                        self.ejecutar_envio_email(datos_mod, es_modificacion=True)
                        
            except Exception as e: messagebox.showerror("Error", str(e))

        f_btn_mod = ctk.CTkFrame(v_mod, fg_color="transparent")
        f_btn_mod.pack(fill="x", padx=20, pady=20)
        ctk.CTkButton(f_btn_mod, text="💾 Guardar", fg_color="#1f538d", hover_color="#163b65", command=lambda: ejecutar_modificacion("solo")).pack(side="left", fill="x", expand=True, padx=2)
        ctk.CTkButton(f_btn_mod, text="💬 Guardar + WA", fg_color="#27ae60", hover_color="#1e8449", command=lambda: ejecutar_modificacion("whatsapp")).pack(side="left", fill="x", expand=True, padx=2)
        ctk.CTkButton(f_btn_mod, text="📧 Guardar + Mail", fg_color="#e67e22", hover_color="#d35400", command=lambda: ejecutar_modificacion("email")).pack(side="left", fill="x", expand=True, padx=2)

if __name__ == "__main__":
    pass