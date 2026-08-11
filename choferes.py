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
from datetime import datetime
from conexion import conectar_db, registrar_auditoria

def abrir_documento_local(ruta):
    if not ruta: return False
    ruta_norm = os.path.normpath(ruta)
    if not os.path.exists(ruta_norm):
        return False
    try:
        if sys.platform == "win32": os.startfile(ruta_norm)
        elif sys.platform == "darwin": subprocess.call(["open", ruta_norm])
        else: subprocess.call(["xdg-open", ruta_norm])
        return True
    except Exception as e:
        print(f"Error al abrir documento: {e}")
        return False

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
        self.lbl_doc.configure(text=f"Cargar: {doc_actual}")

    def cargar_actual(self):
        ruta = filedialog.askopenfilename(title=f"Seleccionar {self.documentos[self.indice]}", filetypes=[("Documentos", "*.pdf;*.png;*.jpg;*.jpeg")])
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
        
        # Controles superiores (Desplegables en lugar de solo texto)
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
        # Actualizar los combos a la selección actual
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
        self.inicializar_bd()
        self.crear_interfaz()

    def inicializar_bd(self):
        conn = conectar_db()
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
                "ALTER TABLE choferes ADD COLUMN ruta_documentos TEXT DEFAULT ''"
            ]
            
            for query in columnas_nuevas:
                try: cursor.execute(query); conn.commit()
                except: conn.rollback()
                
        except Exception as e:
            print(f"Error BD Choferes: {e}")
        finally:
            conn.close()

    def buscar_documento_api(self, event=None):
        ruc = self.ent_ruc.get().strip()
        if len(ruc) != 11 or not ruc.isdigit():
            return messagebox.showwarning("RUC Inválido", "Por favor, ingrese un RUC válido de 11 dígitos para buscar en SUNAT.")
        
        try:
            import json
            url = f"https://api.apis.net.pe/v1/ruc?numero={ruc}"
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=5) as response:
                if response.status == 200:
                    data = json.loads(response.read().decode())
                    self.ent_nombres.delete(0, tk.END)
                    self.ent_nombres.insert(0, data.get("nombre", ""))
                    self.ent_direccion.delete(0, tk.END)
                    self.ent_direccion.insert(0, data.get("direccion", ""))
                    messagebox.showinfo("Éxito", "Datos de SUNAT recuperados correctamente.")
                else:
                    messagebox.showwarning("Sin Resultados", "No se encontró información para este RUC.")
        except Exception as e:
            messagebox.showwarning("Error de Conexión", f"No se pudo conectar con la API de SUNAT:\n{e}")

    def cargar_moviles_disponibles(self):
        conn = conectar_db()
        moviles = ["Ninguno / Sin Asignar"]
        if conn:
            try:
                c = conn.cursor()
                c.execute("SELECT placa, marca, modelo FROM flota_vehiculos WHERE estado = 'Operativo'")
                for r in c.fetchall():
                    moviles.append(f"{r[0]} | {r[1]} {r[2]}")
            except: pass
            finally: conn.close()
            
        if hasattr(self, 'cmb_movil'):
            self.cmb_movil.configure(values=moviles)
            if self.cmb_movil.get() not in moviles:
                self.cmb_movil.set("Ninguno / Sin Asignar")

    def actualizar_botones_docs(self):
        todas = {**self.rutas_documentos_db, **self.rutas_documentos_temp}
        validos = sum(1 for p in todas.values() if p and os.path.exists(os.path.normpath(p)))
        
        if validos > 0:
            self.btn_ver_docs.configure(state="normal", fg_color="#34495e")
            self.btn_cargar_docs.configure(text=f"✅ {validos} Doc(s) Listos", fg_color="#27ae60")
        else:
            self.btn_ver_docs.configure(state="disabled", fg_color="#7f8c8d")
            self.btn_cargar_docs.configure(text="📎 Cargar Documentos", fg_color="#1f538d")

    def lanzar_asistente_carga(self):
        todas_rutas_actuales = {**self.rutas_documentos_db, **self.rutas_documentos_temp}
        docs_requeridos = ["DNI", "Brevete", "Antecedentes Penales", "Antecedentes Judiciales"]
        
        faltantes = []
        for doc in docs_requeridos:
            ruta = todas_rutas_actuales.get(doc)
            if not ruta or not os.path.exists(os.path.normpath(ruta)):
                faltantes.append(doc)
                
        if not faltantes:
            messagebox.showinfo("Expediente Completo", "Ya tienes todos los documentos cargados y guardados para este personal.")
            return

        def on_asistente_cerrado():
            self.actualizar_botones_docs()
                
        AsistenteCargaDocs(self.parent_frame.winfo_toplevel(), self.rutas_documentos_temp, faltantes, on_asistente_cerrado)

    def gestionar_documentos(self):
        todas_rutas = {**self.rutas_documentos_db, **self.rutas_documentos_temp}
        rutas_validas = {k: v for k, v in todas_rutas.items() if v and os.path.exists(os.path.normpath(v))}
        
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

        # --- Datos Personales ---
        ctk.CTkLabel(self.f_form, text="Datos Personales", font=("Arial", 14, "bold")).pack(pady=(15, 10))
        self.ent_dni = crear_campo("DNI: *", "Ej: 12345678")
        
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
        self.ent_venc_licencia = crear_campo_fecha("Vencimiento de Licencia:")

        ctk.CTkLabel(self.f_form, text="Estado Laboral:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.cmb_estado = ctk.CTkComboBox(self.f_form, values=["Activo", "Inactivo", "Suspendido"], state="readonly")
        self.cmb_estado.pack(fill="x", padx=10, pady=(0, 15))
        self.cmb_estado.set("Activo")

        # --- CARGA DE EXPEDIENTE ASISTIDA ---
        ctk.CTkLabel(self.f_form, text="--- Expediente Físico del Chofer ---", font=("Arial", 11, "bold"), text_color="#8e44ad").pack(anchor="w", padx=10, pady=(5,5))
        f_docs = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_docs.pack(fill="x", padx=10, pady=(0, 15))
        
        self.btn_cargar_docs = ctk.CTkButton(f_docs, text="📎 Cargar Documentos", font=("Arial", 11, "bold"), fg_color="#1f538d", hover_color="#163b65", command=self.lanzar_asistente_carga)
        self.btn_cargar_docs.pack(side="left", expand=True, fill="x", padx=(0, 5))
        
        self.btn_ver_docs = ctk.CTkButton(f_docs, text="👁️ Gestionar Docs.", font=("Arial", 11, "bold"), fg_color="#7f8c8d", command=self.gestionar_documentos, state="disabled")
        self.btn_ver_docs.pack(side="right", expand=True, fill="x", padx=(5, 0))

        f_btns = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_btns.pack(fill="x", padx=10, pady=(10, 20))
        
        self.btn_guardar = ctk.CTkButton(f_btns, text="💾 Guardar Nuevo", fg_color="#27ae60", hover_color="#1e8449", font=("Arial", 12, "bold"), command=self.guardar_chofer)
        self.btn_guardar.pack(side="left", expand=True, fill="x", padx=(0, 5))
        
        btn_limpiar = ctk.CTkButton(f_btns, text="🔄 Limpiar", fg_color="#7f8c8d", hover_color="#606b6b", font=("Arial", 12, "bold"), command=self.limpiar_formulario)
        btn_limpiar.pack(side="right", expand=True, fill="x", padx=(5, 0))

        # PANEL DERECHO: TABLA
        f_derecho = ctk.CTkFrame(self.main_split, fg_color="transparent")
        f_derecho.pack(side="right", fill="both", expand=True)

        f_busqueda = ctk.CTkFrame(f_derecho, fg_color="transparent")
        f_busqueda.pack(fill="x", pady=(0, 5))
        ctk.CTkLabel(f_busqueda, text="🔍 Buscar:", font=("Arial", 12, "bold")).pack(side="left", padx=(0, 5))
        self.ent_buscar = ctk.CTkEntry(f_busqueda, placeholder_text="Buscar por DNI, Nombres, Licencia, Móvil...")
        self.ent_buscar.pack(side="left", fill="x", expand=True)
        self.ent_buscar.bind("<KeyRelease>", lambda e: self.cargar_datos())

        f_tabla = ctk.CTkFrame(f_derecho, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True)

        columnas = ("id", "dni", "nombres", "telefono", "licencia", "vencimiento", "movil", "estado")
        self.tabla = ttk.Treeview(f_tabla, columns=columnas, show="headings")
        self.tabla.heading("id", text="ID")
        self.tabla.heading("dni", text="DNI")
        self.tabla.heading("nombres", text="Nombres y Apellidos")
        self.tabla.heading("telefono", text="Teléfono")
        self.tabla.heading("licencia", text="N° Licencia")
        self.tabla.heading("vencimiento", text="Venc. Licencia")
        self.tabla.heading("movil", text="Móvil Asignado")
        self.tabla.heading("estado", text="Estado")

        self.tabla.column("id", width=0, stretch=tk.NO)
        self.tabla.column("dni", width=80, anchor="center")
        self.tabla.column("nombres", width=200, anchor="w")
        self.tabla.column("telefono", width=90, anchor="center")
        self.tabla.column("licencia", width=90, anchor="center")
        self.tabla.column("vencimiento", width=95, anchor="center")
        self.tabla.column("movil", width=120, anchor="center")
        self.tabla.column("estado", width=80, anchor="center")
        
        self.tabla.config(displaycolumns=("dni", "nombres", "telefono", "licencia", "movil", "estado"))

        scroll_y = ctk.CTkScrollbar(f_tabla, orientation="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scroll_y.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")
        
        self.tabla.bind("<Double-1>", lambda e: self.cargar_para_edicion())

        f_acciones_tabla = ctk.CTkFrame(f_derecho, fg_color="transparent")
        f_acciones_tabla.pack(fill="x", pady=10)
        
        ctk.CTkButton(f_acciones_tabla, text="✏️ Editar Seleccionado", fg_color="#34495e", hover_color="#2c3e50", font=("Arial", 12, "bold"), command=self.cargar_para_edicion).pack(side="left", padx=5)
        ctk.CTkButton(f_acciones_tabla, text="❌ Eliminar", fg_color="#e74c3c", hover_color="#c0392b", font=("Arial", 12, "bold"), command=self.eliminar_chofer).pack(side="right", padx=5)

        self.cargar_datos()

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
        self.ent_correo.delete(0, tk.END)
        
        self.cargar_moviles_disponibles()
        self.ent_salud_num.delete(0, tk.END)
        self.ent_salud_venc.delete(0, tk.END)
        self.ent_vida_num.delete(0, tk.END)
        self.ent_vida_venc.delete(0, tk.END)
        
        self.ent_licencia.delete(0, tk.END)
        self.ent_cat_licencia.delete(0, tk.END)
        self.ent_venc_licencia.delete(0, tk.END)
        self.cmb_estado.set("Activo")
        
        self.rutas_documentos_temp = {}
        self.rutas_documentos_db = {}
        self.actualizar_botones_docs()

    def cargar_datos(self):
        for item in self.tabla.get_children(): self.tabla.delete(item)
        filtro = self.ent_buscar.get().strip().lower()
        
        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            if filtro:
                cursor.execute("""
                    SELECT id, dni, nombres, telefono, licencia, vencimiento_licencia, movil_asignado, estado 
                    FROM choferes 
                    WHERE dni ILIKE %s OR nombres ILIKE %s OR licencia ILIKE %s OR movil_asignado ILIKE %s
                    ORDER BY nombres ASC
                """, (f"%{filtro}%", f"%{filtro}%", f"%{filtro}%", f"%{filtro}%"))
            else:
                cursor.execute("SELECT id, dni, nombres, telefono, licencia, vencimiento_licencia, movil_asignado, estado FROM choferes ORDER BY nombres ASC")
            
            for r in cursor.fetchall():
                self.tabla.insert("", tk.END, values=r)
        except Exception as e:
            print(f"Error cargando tabla: {e}")
        finally:
            conn.close()

    def guardar_chofer(self):
        dni = self.ent_dni.get().strip()
        nombres = self.ent_nombres.get().strip().upper()
        ruc = self.ent_ruc.get().strip()
        direccion = self.ent_direccion.get().strip()
        fec_nac = self.ent_fec_nac.get().strip()
        sexo = self.cmb_sexo.get()
        hijos = self.ent_hijos.get().strip() or "0"
        
        tel = self.ent_telefono.get().strip()
        correo = self.ent_correo.get().strip()
        
        movil = self.cmb_movil.get()
        salud_num = self.ent_salud_num.get().strip()
        salud_venc = self.ent_salud_venc.get().strip()
        vida_num = self.ent_vida_num.get().strip()
        vida_venc = self.ent_vida_venc.get().strip()
        
        licencia = self.ent_licencia.get().strip().upper()
        cat = self.ent_cat_licencia.get().strip().upper()
        venc = self.ent_venc_licencia.get().strip()
        estado = self.cmb_estado.get()

        if not dni or not nombres:
            return messagebox.showwarning("Atención", "El DNI y los Nombres son obligatorios.")

        # Lógica de guardado masivo de documentos del expediente
        diccionario_final = self.rutas_documentos_db.copy()
        
        if self.rutas_documentos_temp:
            try:
                base_dir = os.path.dirname(os.path.abspath(__file__))
                carpeta_destino = os.path.normpath(os.path.join(base_dir, "archivos_flota", "expedientes_choferes"))
                os.makedirs(carpeta_destino, exist_ok=True)
                
                for doc_name, temp_path in self.rutas_documentos_temp.items():
                    if temp_path and os.path.exists(os.path.normpath(temp_path)):
                        ext = os.path.splitext(temp_path)[1]
                        nombre_seguro = doc_name.replace(" ", "_")
                        nombre_archivo = f"Expediente_{nombre_seguro}_{dni}_{datetime.now().strftime('%Y%m%d%H%M%S')}{ext}"
                        ruta_final = os.path.join(carpeta_destino, nombre_archivo)
                        
                        shutil.copy2(os.path.normpath(temp_path), ruta_final)
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
                    ruta_documentos=%s
                    WHERE id=%s
                """, (dni, nombres, ruc, tel, correo, licencia, cat, venc, estado,
                      direccion, fec_nac, sexo, hijos, movil, salud_num, salud_venc, vida_num, vida_venc, json_rutas_finales, self.id_edicion))
                registrar_auditoria(self.usuario_activo, "Choferes", f"Actualizó datos de {nombres}")
                messagebox.showinfo("Éxito", "Datos actualizados correctamente.")
            else:
                cursor.execute("SELECT id FROM choferes WHERE dni = %s", (dni,))
                if cursor.fetchone():
                    conn.close()
                    return messagebox.showwarning("Duplicado", f"El DNI {dni} ya existe en el sistema.")
                    
                cursor.execute("""
                    INSERT INTO choferes (dni, nombres, ruc, telefono, correo, licencia, categoria_licencia, vencimiento_licencia, estado, 
                    direccion, fecha_nacimiento, sexo, numero_hijos, movil_asignado, seguro_salud_num, seguro_salud_venc, seguro_vida_num, seguro_vida_venc, ruta_documentos) 
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (dni, nombres, ruc, tel, correo, licencia, cat, venc, estado,
                      direccion, fec_nac, sexo, hijos, movil, salud_num, salud_venc, vida_num, vida_venc, json_rutas_finales))
                registrar_auditoria(self.usuario_activo, "Choferes", f"Registró nuevo conductor/personal: {nombres}")
                messagebox.showinfo("Éxito", "Personal registrado correctamente.")
            
            # 🚀 INYECCIÓN AL CRONOGRAMA DE VENCIMIENTOS
            try:
                c_crono = conn.cursor()
                identificador_crono = f"{nombres} (DNI: {dni})"
                c_crono.execute("DELETE FROM tareas_evento WHERE evento_asociado = 'FLOTA | Vencimientos' AND responsable = %s", (identificador_crono,))
                
                vencimientos = [
                    (f"Venc. Licencia ({cat})", venc),
                    ("Venc. Seguro Salud (EsSalud/EPS)", salud_venc),
                    ("Venc. Seguro Vida Ley", vida_venc)
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
            self.limpiar_formulario()
            self.cargar_datos()
        except Exception as e:
            conn.rollback()
            messagebox.showerror("Error", str(e))
        finally:
            conn.close()

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
                ruta_documentos
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
                
                json_str = r[19] if len(r) > 19 and r[19] else "{}"
                try:
                    self.rutas_documentos_db = json.loads(json_str)
                except Exception:
                    if json_str and not json_str.startswith("{"):
                        self.rutas_documentos_db = {"Expediente Clásico": json_str}
                    else:
                        self.rutas_documentos_db = {}

                self.actualizar_botones_docs()
                    
        except Exception as e:
            print("Error cargando edición:", e)
        finally:
            conn.close()

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
                registrar_auditoria(self.usuario_activo, "Choferes", f"Eliminó al conductor {nombres}")
                self.cargar_datos()
                self.limpiar_formulario()
            except Exception as e:
                messagebox.showerror("Error", str(e))
            finally:
                conn.close()