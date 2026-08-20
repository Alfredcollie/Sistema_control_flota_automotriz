# -*- coding: utf-8 -*-
import os
import sys 
import tempfile
import subprocess 
import webbrowser
import urllib.parse
import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk
import calendar
import ctypes
from datetime import datetime, timedelta
import threading

# 🚀 IMPORTAMOS NUESTRAS NUEVAS HERRAMIENTAS CORPORATIVAS
from conexion import conectar_db, registrar_auditoria, liberar_conexion
from buffer_memoria import cache_sistema

# =========================================================
# 🚀 ADAPTACIÓN MULTIPLATAFORMA: Ocultar consola solo en Windows
# =========================================================
if sys.platform == "win32":
    try:
        hwnd_cmd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd_cmd:
            ctypes.windll.user32.ShowWindow(hwnd_cmd, 6)
    except Exception:
        pass

# =========================================================
# 🚀 ADAPTACIÓN MULTIPLATAFORMA: Función universal para abrir archivos
# =========================================================
def abrir_documento(ruta):
    try:
        if sys.platform == "win32":
            os.startfile(ruta)
        elif sys.platform == "darwin": 
            subprocess.call(["open", ruta])
        else: 
            subprocess.call(["xdg-open", ruta])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")

def maximizar_ventana(ventana):
    if sys.platform == "win32":
        try:
            ventana.state("zoomed")
        except Exception:
            pass
    else:
        try:
            w = ventana.winfo_screenwidth()
            h = ventana.winfo_screenheight()
            ventana.geometry(f"{w}x{h}+0+0")
        except Exception:
            pass


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
        fecha_seleccionada = f"{day:02d}/{self.current_month:02d}/{self.current_year}"
        self.target_entry.delete(0, tk.END)
        self.target_entry.insert(0, fecha_seleccionada)
        self.destroy()

# =========================================================
# CLASE: DASHBOARD DE CALENDARIO GENERAL A PANTALLA COMPLETA
# =========================================================
class CalendarioDashboard(ctk.CTkToplevel):
    def __init__(self, parent, usuario_activo):
        super().__init__(parent)
        self.usuario_activo = usuario_activo
        self.title("Calendario General de Flota y Tareas")
        self.geometry("1200x700")
        
        self.after(100, lambda: maximizar_ventana(self))
        
        self.transient(parent)
        self.grab_set()

        self.mes_actual = datetime.now().month
        self.anio_actual = datetime.now().year
        self.tareas_db = {} 
        self.pop_detalle = None 

        self.crear_interfaz()
        self.cargar_datos_db()

    def crear_interfaz(self):
        self.f_header = ctk.CTkFrame(self, corner_radius=0, fg_color="#1a252c")
        self.f_header.pack(fill="x")

        f_controles = ctk.CTkFrame(self.f_header, fg_color="transparent")
        f_controles.pack(side="left", padx=20, pady=8)
        
        ctk.CTkButton(f_controles, text="<", width=40, fg_color="#34495e", hover_color="#2c3e50", command=self.mes_anterior).pack(side="left", padx=5)
        self.lbl_mes_anio = ctk.CTkLabel(f_controles, text="", font=("Arial", 16, "bold"), text_color="white", width=150)
        self.lbl_mes_anio.pack(side="left", padx=5)
        ctk.CTkButton(f_controles, text=">", width=40, fg_color="#34495e", hover_color="#2c3e50", command=self.mes_siguiente).pack(side="left", padx=5)

        f_filtros = ctk.CTkFrame(self.f_header, fg_color="transparent")
        f_filtros.pack(side="left", padx=20, pady=8)
        
        ctk.CTkLabel(f_filtros, text="Filtrar por:", text_color="white", font=("Arial", 12, "bold")).pack(side="left", padx=5)
        
        self.combo_filtro_principal = ctk.CTkComboBox(f_filtros, values=["Todo", "Por Vehículo", "Trabajo Interno", "Por Proveedor"], command=self.actualizar_opciones_filtro, width=150)
        self.combo_filtro_principal.pack(side="left", padx=5)
        self.combo_filtro_principal.set("Todo")

        self.combo_filtro_secundario = ctk.CTkComboBox(f_filtros, values=["-"], state="disabled", command=self.aplicar_filtro, width=250)
        self.combo_filtro_secundario.pack(side="left", padx=5)

        f_btns = ctk.CTkFrame(self.f_header, fg_color="transparent")
        f_btns.pack(side="right", padx=20, pady=8)

        ctk.CTkButton(f_btns, text="📥 Exportar Todo", font=("Arial", 12, "bold"), fg_color="#28a745", hover_color="#218838", height=32, command=self.exportar_calendario_completo_ics).pack(side="left", padx=5)

        ctk.CTkButton(f_btns, text="[ + ] Mantenimiento / Prov.", font=("Arial", 12, "bold"), fg_color="#1f538d", hover_color="#163b65", height=32, command=self.abrir_agendar_proveedor).pack(side="left", padx=5)
        ctk.CTkButton(f_btns, text="[ + ] Trabajo (Oficina)", font=("Arial", 12, "bold"), fg_color="#34495e", hover_color="#2c3e50", height=32, command=self.abrir_agendar_interno).pack(side="left", padx=5)

        self.f_grid = ctk.CTkFrame(self, fg_color="#ecf0f1", corner_radius=0)
        self.f_grid.pack(fill="both", expand=True, padx=5, pady=5)

        for i in range(7):
            self.f_grid.grid_columnconfigure(i, weight=1, uniform="col")
        
        dias = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO", "DOMINGO"]
        for i, d in enumerate(dias):
            ctk.CTkLabel(self.f_grid, text=d, font=("Arial", 12, "bold"), fg_color="#2980b9", text_color="white", corner_radius=5).grid(row=0, column=i, sticky="ew", padx=1, pady=1, ipadx=5, ipady=3)

    def actualizar_opciones_filtro(self, choice):
        if choice in ["Todo", "Trabajo Interno"]:
            self.combo_filtro_secundario.configure(values=["-"], state="disabled")
            self.combo_filtro_secundario.set("-")
            self.cargar_datos_db()
        else:
            self.combo_filtro_secundario.configure(state="normal")
            self.combo_filtro_secundario.set("Cargando...")
            
            def tarea_filtro():
                opciones = ["Sin registros"]
                conn = conectar_db(silencioso=True)
                if conn:
                    try:
                        cursor = conn.cursor()
                        if choice == "Por Vehículo":
                            cursor.execute("SELECT DISTINCT evento_asociado FROM tareas_evento WHERE evento_asociado != 'OFICINA | Trabajos Internos' AND evento_asociado IS NOT NULL")
                            opc_db = [str(r[0]) for r in cursor.fetchall()]
                            if opc_db: opciones = opc_db
                        elif choice == "Por Proveedor":
                            cursor.execute("SELECT DISTINCT responsable FROM tareas_evento WHERE responsable IS NOT NULL AND responsable != ''")
                            opc_db = [str(r[0]) for r in cursor.fetchall()]
                            if opc_db: opciones = opc_db
                    except Exception as e:
                        print("Error cargando filtros:", e)
                    finally:
                        liberar_conexion(conn)
                self.after(0, lambda: self._aplicar_opciones_filtro(opciones))

            threading.Thread(target=tarea_filtro, daemon=True).start()

    def _aplicar_opciones_filtro(self, opciones):
        self.combo_filtro_secundario.configure(values=opciones, state="readonly")
        self.combo_filtro_secundario.set(opciones[0])
        self.cargar_datos_db()

    def aplicar_filtro(self, choice=None):
        self.cargar_datos_db()

    def cargar_datos_db(self):
        self.tareas_db.clear()
        
        filtro_p = self.combo_filtro_principal.get() if hasattr(self, 'combo_filtro_principal') else "Todo"
        filtro_s = self.combo_filtro_secundario.get() if hasattr(self, 'combo_filtro_secundario') else "-"
        
        clave_cache = f"calendario_{filtro_p}_{filtro_s}"
        datos_calendario = cache_sistema.obtener(clave_cache)

        if datos_calendario is not None:
            self._procesar_y_dibujar(datos_calendario)
        else:
            self.lbl_mes_anio.configure(text="Cargando...")
            def tarea_dash():
                datos_db = []
                conn = conectar_db(silencioso=True)
                if conn:
                    try:
                        query = "SELECT id, fecha_limite, nombre_tarea, evento_asociado, responsable, notas, tipo_entrada, ubicacion, repeticion, tiempo_aviso FROM tareas_evento"
                        params = []
                        condiciones = []

                        if filtro_p == "Trabajo Interno":
                            condiciones.append("evento_asociado = 'OFICINA | Trabajos Internos'")
                        elif filtro_p == "Por Vehículo" and filtro_s != "-" and filtro_s != "Sin registros":
                            condiciones.append("evento_asociado = %s")
                            params.append(filtro_s)
                        elif filtro_p == "Por Proveedor" and filtro_s != "-" and filtro_s != "Sin registros":
                            condiciones.append("responsable = %s")
                            params.append(filtro_s)

                        if condiciones:
                            query += " WHERE " + " AND ".join(condiciones)

                        cursor = conn.cursor()
                        cursor.execute(query, tuple(params))
                        datos_db = cursor.fetchall()
                        cache_sistema.guardar(clave_cache, datos_db)
                    except Exception as e:
                        print("Error SQL Calendario:", e)
                    finally:
                        liberar_conexion(conn)
                
                self.after(0, lambda: self._procesar_y_dibujar(datos_db))

            threading.Thread(target=tarea_dash, daemon=True).start()
            
    def _procesar_y_dibujar(self, datos_calendario):
        self.tareas_db.clear()
        if datos_calendario:
            for t_id, f_limite, t_nombre, evento, resp, notas, t_entrada, ubic, repeticion, aviso in datos_calendario:
                if f_limite:
                    if f_limite not in self.tareas_db:
                        self.tareas_db[f_limite] = []
                        
                    evento_nombre_limpio = evento.split(" | ")[1] if " | " in evento else evento
                        
                    self.tareas_db[f_limite].append({
                        "id": t_id,
                        "tarea": t_nombre,
                        "evento_completo": evento,
                        "evento_nombre": evento_nombre_limpio,
                        "responsable": resp,
                        "notas": notas if notas else "",
                        "fecha_limite": f_limite,
                        "tipo_entrada": t_entrada,
                        "ubicacion": ubic,
                        "repeticion": repeticion,
                        "aviso": aviso
                    })
                    
        self.dibujar_calendario()

    def exportar_calendario_completo_ics(self):
        if not self.tareas_db:
            messagebox.showinfo("Información", "No hay tareas cargadas en el calendario actual para exportar.")
            return

        ics_content = [
            "BEGIN:VCALENDAR",
            "VERSION:2.0",
            "PRODID:-//BlackCube//CronogramaApp//ES",
            "CALSCALE:GREGORIAN"
        ]
        
        contador_tareas = 0
        for f_limite, tareas_dia in self.tareas_db.items():
            for td in tareas_dia:
                try:
                    dt_start = datetime.strptime(f_limite, "%d/%m/%Y")
                    dt_end = dt_start + timedelta(days=1)
                    
                    str_start = dt_start.strftime("%Y%m%d")
                    str_end = dt_end.strftime("%Y%m%d")
                    dtstamp = datetime.now().strftime('%Y%m%dT%H%M%SZ')
                    
                    desc = f"📌 Vehículo/Proyecto: {td['evento_completo']}\\n👤 Responsable: {td['responsable']}\\n📋 Detalles: {td['notas']}".replace("\n", "\\n")
                    
                    ubic_ics = f"\nLOCATION:{td.get('ubicacion', '')}" if td.get('ubicacion', '') else ""

                    ics_content.extend([
                        "BEGIN:VEVENT",
                        f"UID:tarea_global_{contador_tareas}_{dtstamp}@blackcube",
                        f"DTSTAMP:{dtstamp}",
                        f"DTSTART;VALUE=DATE:{str_start}",
                        f"DTEND;VALUE=DATE:{str_end}",
                        f"SUMMARY:{td['tarea']} - {td['evento_nombre']}",
                        f"DESCRIPTION:{desc}{ubic_ics}",
                        "END:VEVENT"
                    ])
                    contador_tareas += 1
                except Exception as e:
                    print(f"Error procesando fecha {f_limite}: {e}")

        ics_content.append("END:VCALENDAR")
        
        try:
            fd, path = tempfile.mkstemp(suffix=".ics", prefix="Calendario_Completo_BlackCube_")
            with os.fdopen(fd, 'w', encoding='utf-8') as f:
                f.write("\n".join(ics_content))
            
            abrir_documento(path)

            messagebox.showinfo("Exportación Exitosa", 
                f"✅ El calendario se ha generado correctamente.\n\n"
                f"Ruta: {path}\n\n"
                "• Windows/Mac: Se abrirá tu aplicación predeterminada (Ej: Outlook / Apple Calendar).\n"
                "• Google Calendar: Ve a 'Configuración' -> 'Importar y exportar' y sube este archivo."
            )
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar el calendario:\n{e}")

    def exportar_tarea_google(self, td):
        try:
            f_limite = td["fecha_limite"]
            dt_start = datetime.strptime(f_limite, "%d/%m/%Y")
            dt_end = dt_start + timedelta(days=1)
            
            str_start = dt_start.strftime("%Y%m%d")
            str_end = dt_end.strftime("%Y%m%d")
            
            titulo = f"{td['tarea']} - {td['evento_nombre']}"
            desc = f"📌 Vehículo/Proyecto: {td['evento_completo']}\n👤 Responsable: {td['responsable']}\n📋 Detalles: {td['notas']}"
            
            base_url = "https://calendar.google.com/calendar/render?action=TEMPLATE"
            text_param = urllib.parse.quote(titulo)
            dates_param = f"{str_start}/{str_end}"
            details_param = urllib.parse.quote(desc)
            loc_param = urllib.parse.quote(td.get('ubicacion', ''))
            
            final_url = f"{base_url}&text={text_param}&dates={dates_param}&details={details_param}&location={loc_param}"
            webbrowser.open(final_url)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir Google Calendar:\n{e}")

    def exportar_tarea_ics(self, td):
        try:
            f_limite = td["fecha_limite"]
            dt_start = datetime.strptime(f_limite, "%d/%m/%Y")
            dt_end = dt_start + timedelta(days=1)
            
            str_start = dt_start.strftime("%Y%m%d")
            str_end = dt_end.strftime("%Y%m%d")
            dtstamp = datetime.now().strftime('%Y%m%dT%H%M%SZ')
            
            desc = f"📌 Vehículo/Proyecto: {td['evento_completo']}\\n👤 Responsable: {td['responsable']}\\n📋 Detalles: {td['notas']}".replace("\n", "\\n")
            ubic_ics = f"\nLOCATION:{td.get('ubicacion', '')}" if td.get('ubicacion', '') else ""

            ics_content = [
                "BEGIN:VCALENDAR",
                "VERSION:2.0",
                "PRODID:-//BlackCube//CronogramaApp//ES",
                "CALSCALE:GREGORIAN",
                "BEGIN:VEVENT",
                f"UID:tarea_{datetime.now().strftime('%Y%m%d%H%M%S')}@blackcube",
                f"DTSTAMP:{dtstamp}",
                f"DTSTART;VALUE=DATE:{str_start}",
                f"DTEND;VALUE=DATE:{str_end}",
                f"SUMMARY:{td['tarea']} - {td['evento_nombre']}",
                f"DESCRIPTION:{desc}{ubic_ics}",
                "END:VEVENT",
                "END:VCALENDAR"
            ]
            
            fd, path = tempfile.mkstemp(suffix=".ics", prefix="Tarea_BlackCube_")
            with os.fdopen(fd, 'w', encoding='utf-8') as f:
                f.write("\n".join(ics_content))
            
            abrir_documento(path)
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar la tarea al calendario:\n{e}")

    def mostrar_detalle_tarea(self, td):
        det_id = td.get('id')
        det_evento = td['evento_completo']
        det_tarea = td['tarea']
        det_resp = td['responsable']
        det_notas = td['notas']

        if self.pop_detalle is not None and self.pop_detalle.winfo_exists():
            for widget in self.pop_detalle.winfo_children():
                widget.destroy()
            pop = self.pop_detalle
        else:
            pop = ctk.CTkToplevel(self)
            self.pop_detalle = pop
            pop.title("Detalles de la Entrada")
            pop.geometry("480x500")
            pop.resizable(False, False)
            pop.transient(self) 
            
            pop.update_idletasks()
            x = self.winfo_rootx() + (self.winfo_width() // 2) - (480 // 2)
            y = self.winfo_rooty() + (self.winfo_height() // 2) - (500 // 2)
            pop.geometry(f"+{x}+{y}")

        f_top = ctk.CTkFrame(pop, fg_color="transparent")
        f_top.pack(fill="x", padx=15, pady=(15, 5))
        
        ctk.CTkLabel(f_top, text="🔍 DETALLES DE LA ENTRADA", font=("Arial", 16, "bold"), text_color="#1f538d").pack(side="left", expand=True, anchor="w", padx=(5, 0))
        
        f_botones_export = ctk.CTkFrame(f_top, fg_color="transparent")
        f_botones_export.pack(side="right")

        btn_export_ics = ctk.CTkButton(f_botones_export, text="📥 .ics (Mac/Outlook)", font=("Arial", 10, "bold"), fg_color="#7f8c8d", hover_color="#606b6b", width=130, height=26, command=lambda: self.exportar_tarea_ics(td))
        btn_export_ics.pack(side="top", pady=(0, 4))

        btn_google = ctk.CTkButton(f_botones_export, text="🌐 Google Calendar", font=("Arial", 10, "bold"), fg_color="#2980b9", hover_color="#1c5982", width=130, height=26, command=lambda: self.exportar_tarea_google(td))
        btn_google.pack(side="top")

        f_info = ctk.CTkFrame(pop, fg_color="transparent")
        f_info.pack(fill="both", expand=True, padx=20, pady=5)

        ctk.CTkLabel(f_info, text="📌 VEHÍCULO / PROYECTO:", font=("Arial", 11, "bold")).pack(anchor="w", pady=(5, 0))
        ctk.CTkLabel(f_info, text=det_evento, font=("Arial", 13), justify="left").pack(anchor="w", padx=5)

        ctk.CTkLabel(f_info, text="📝 TÍTULO / DESCRIPCIÓN:", font=("Arial", 11, "bold")).pack(anchor="w", pady=(10, 0))
        ent_tarea = ctk.CTkEntry(f_info, font=("Arial", 12))
        ent_tarea.pack(fill="x", padx=5, pady=(2, 0))
        ent_tarea.insert(0, det_tarea)
        
        if td.get('ubicacion'):
            ctk.CTkLabel(f_info, text="📍 UBICACIÓN:", font=("Arial", 11, "bold")).pack(anchor="w", pady=(10, 0))
            ctk.CTkLabel(f_info, text=td.get('ubicacion'), font=("Arial", 13), justify="left").pack(anchor="w", padx=5)

        if td.get('tipo_entrada') != 'Cumpleaños':
            ctk.CTkLabel(f_info, text="👤 RESPONSABLE / TALLER:", font=("Arial", 11, "bold")).pack(anchor="w", pady=(10, 0))
            ent_resp = ctk.CTkEntry(f_info, font=("Arial", 12))
            ent_resp.pack(fill="x", padx=5, pady=(2, 0))
            ent_resp.insert(0, det_resp)

        ctk.CTkLabel(f_info, text="📋 DETALLES ADICIONALES:", font=("Arial", 11, "bold")).pack(anchor="w", pady=(10, 0))
        txt_notas = ctk.CTkTextbox(f_info, height=90, wrap="word", border_width=1, fg_color="#ffffff", text_color="black")
        txt_notas.pack(fill="x", padx=5, pady=(2, 0))
        txt_notas.insert("1.0", det_notas)

        f_botones = ctk.CTkFrame(pop, fg_color="transparent")
        f_botones.pack(fill="x", pady=(5, 15), padx=20)

        def click_guardar_rapido():
            r_val = ent_resp.get() if td.get('tipo_entrada') != 'Cumpleaños' else ""
            self.guardar_edicion_rapida(det_id, ent_tarea.get(), r_val, txt_notas.get("1.0", "end-1c"), pop)

        btn_guardar = ctk.CTkButton(f_botones, text="💾 Guardar Cambios", font=("Arial", 12, "bold"), fg_color="#1f538d", hover_color="#163b65", height=32, command=click_guardar_rapido)
        btn_guardar.pack(side="left", expand=True, fill="x", padx=(0, 5))

        btn_cerrar = ctk.CTkButton(f_botones, text="Cerrar", font=("Arial", 12, "bold"), fg_color="#7f8c8d", hover_color="#606b6b", height=32, command=pop.destroy)
        btn_cerrar.pack(side="right", expand=True, fill="x", padx=(5, 0))

    def guardar_edicion_rapida(self, t_id, nueva_tarea, nuevo_resp, nuevas_notas, pop_window):
        if not nueva_tarea.strip():
            messagebox.showwarning("Atención", "El título de la entrada es obligatorio.", parent=pop_window)
            return

        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE tareas_evento 
                SET nombre_tarea=%s, responsable=%s, notas=%s
                WHERE id=%s
            """, (nueva_tarea.strip(), nuevo_resp.strip(), nuevas_notas.strip(), t_id))
            conn.commit()
            cache_sistema.invalidar()
            registrar_auditoria(self.usuario_activo, "Cronograma", f"Edición rápida de entrada ID {t_id}")
            
            pop_window.destroy()
            self.cargar_datos_db()
            messagebox.showinfo("Éxito", "Cambios guardados correctamente.")
        except Exception as ex:
            messagebox.showerror("Error", f"No se pudieron guardar los cambios:\n{ex}", parent=pop_window)
        finally:
            liberar_conexion(conn)

    def dibujar_calendario(self):
        for widget in self.f_grid.winfo_children():
            if int(widget.grid_info()["row"]) > 0:
                widget.destroy()

        meses = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        self.lbl_mes_anio.configure(text=f"{meses[self.mes_actual].upper()} {self.anio_actual}")

        cal = calendar.monthcalendar(self.anio_actual, self.mes_actual)
        num_weeks = len(cal)

        for i in range(1, 7):
            self.f_grid.grid_rowconfigure(i, weight=0, uniform="") 
            
        for i in range(1, num_weeks + 1):
            self.f_grid.grid_rowconfigure(i, weight=1, uniform="row") 

        hoy = datetime.now()

        for row_idx, week in enumerate(cal, start=1):
            for col_idx, day in enumerate(week):
                if day != 0:
                    fecha_str = f"{day:02d}/{self.mes_actual:02d}/{self.anio_actual}"
                    
                    bg_color = "#ffffff"
                    if day == hoy.day and self.mes_actual == hoy.month and self.anio_actual == hoy.year:
                        bg_color = "#eafaf1" 

                    cell_frame = ctk.CTkFrame(self.f_grid, fg_color=bg_color, corner_radius=5, border_width=1, border_color="#bdc3c7")
                    cell_frame.grid(row=row_idx, column=col_idx, sticky="nsew", padx=1, pady=1)
                    
                    lbl_day = ctk.CTkLabel(cell_frame, text=str(day), font=("Arial", 11, "bold"), text_color="#2c3e50", height=12)
                    lbl_day.pack(side="top", anchor="ne", padx=5, pady=(0, 0))

                    if fecha_str in self.tareas_db:
                        tareas_del_dia = self.tareas_db[fecha_str]
                        
                        if len(tareas_del_dia) <= 5:
                            f_contenedor = ctk.CTkFrame(cell_frame, fg_color="transparent")
                            f_contenedor.pack(fill="both", expand=True, padx=1, pady=0)
                        else:
                            f_contenedor = ctk.CTkScrollableFrame(cell_frame, fg_color="transparent", height=90)
                            f_contenedor.pack(fill="both", expand=True, padx=0, pady=0)
                        
                        for tarea_data in tareas_del_dia:
                            t_tipo = tarea_data.get("tipo_entrada", "Tarea")
                            
                            color_estado = "#fff3cd"
                            text_color = "#856404"
                            
                            if t_tipo == "Renovación Documento":
                                color_estado = "#e8daef"
                                text_color = "#8e44ad"
                            elif t_tipo == "Mantenimiento":
                                color_estado = "#d1ecf1"
                                text_color = "#0c5460"
                                
                            titulo_corto = tarea_data["tarea"][:34] + ".." if len(tarea_data["tarea"]) > 34 else tarea_data["tarea"]
                            
                            lbl_tarea = ctk.CTkLabel(
                                f_contenedor, 
                                text=f"• {titulo_corto}", 
                                font=("Arial", 9, "bold"), 
                                fg_color=color_estado, 
                                text_color=text_color, 
                                corner_radius=3, 
                                anchor="w",
                                padx=4,
                                cursor="hand2",
                                height=16 
                            )
                            lbl_tarea.pack(side="top", fill="x", pady=(0, 1))
                            lbl_tarea.bind("<Button-1>", lambda e, td=tarea_data: self.mostrar_detalle_tarea(td))

    def mes_anterior(self):
        self.mes_actual -= 1
        if self.mes_actual < 1:
            self.mes_actual = 12
            self.anio_actual -= 1
        self.dibujar_calendario()

    def mes_siguiente(self):
        self.mes_actual += 1
        if self.mes_actual > 12:
            self.mes_actual = 1
            self.anio_actual += 1
        self.dibujar_calendario()

    def abrir_agendar_proveedor(self):
        v_prov = ctk.CTkToplevel(self.parent_frame.winfo_toplevel())
        v_prov.title("Agendar Mantenimiento a Flota")
        v_prov.geometry("420x450")
        v_prov.transient(self.parent_frame.winfo_toplevel())
        v_prov.grab_set()

        ctk.CTkLabel(v_prov, text="📋 AGENDAR TAREA A FLOTA", font=("Arial", 14, "bold"), text_color="#1f538d").pack(pady=(15, 10))

        f_cont = ctk.CTkFrame(v_prov, fg_color="transparent")
        f_cont.pack(fill="both", expand=True, padx=20, pady=5)

        ctk.CTkLabel(f_cont, text="Vehículo / Flota:", font=("Arial", 11, "bold")).pack(anchor="w")
        cmb_flota = ctk.CTkComboBox(f_cont, state="readonly")
        cmb_flota.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(f_cont, text="Proveedor Asignado (Taller/Mecánico):", font=("Arial", 11, "bold")).pack(anchor="w")
        cmb_prov = ctk.CTkComboBox(f_cont, state="readonly", values=["Cargando..."])
        cmb_prov.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(f_cont, text="Título / Descripción de la Tarea:", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_t = ctk.CTkEntry(f_cont, placeholder_text="Ej: Cambio de Aceite / Revisión Técnica")
        ent_t.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(f_cont, text="Fecha Programada (DD/MM/AAAA):", font=("Arial", 11, "bold")).pack(anchor="w")
        f_fecha = ctk.CTkFrame(f_cont, fg_color="transparent")
        f_fecha.pack(fill="x", pady=(0, 10))
        ent_f = ctk.CTkEntry(f_fecha, placeholder_text="Seleccione fecha...")
        ent_f.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(f_fecha, text="[ 📅 ]", width=40, fg_color="#1f538d", hover_color="#163b65", command=lambda: CalendarioNativo(v_prov, ent_f)).pack(side="right", padx=(5, 0))

        conn = conectar_db(silencioso=True)
        if conn:
            try:
                cursor = conn.cursor()
                cursor.execute("SELECT placa, marca FROM flota_vehiculos ORDER BY placa ASC")
                eventos = [f"{r[0]} | {r[1]}" for r in cursor.fetchall()]
                if eventos:
                    cmb_flota.configure(values=eventos)
                    cmb_flota.set(eventos[0])
                else:
                    cmb_flota.configure(values=["Sin vehículos registrados"])
                    cmb_flota.set("Sin vehículos registrados")
            except Exception as e: 
                print("Error cargando flota_vehiculos:", e)
            finally: 
                liberar_conexion(conn)

        def cargar_proveedores():
            conn2 = conectar_db(silencioso=True)
            if conn2:
                try:
                    c2 = conn2.cursor()
                    c2.execute("SELECT nombre FROM proveedores ORDER BY nombre ASC")
                    provs = [str(r[0]).strip() for r in c2.fetchall() if r[0]]
                    if provs:
                        cmb_prov.configure(values=provs)
                        cmb_prov.set(provs[0])
                    else:
                        cmb_prov.configure(values=["Sin proveedores registrados"])
                        cmb_prov.set("Sin proveedores registrados")
                except Exception as e: pass
                finally: liberar_conexion(conn2)

        cargar_proveedores()

        def guardar_prov():
            ev = cmb_flota.get()
            pr = cmb_prov.get()
            t_base = ent_t.get().strip()
            f = ent_f.get().strip()
            
            if ev == "Sin vehículos registrados" or not ev:
                messagebox.showwarning("Atención", "Debe registrar un vehículo en el sistema primero.", parent=v_prov)
                return
            if not t_base:
                messagebox.showwarning("Atención", "La descripción de la tarea es obligatoria.", parent=v_prov)
                return

            conn3 = conectar_db()
            if not conn3: return
            try:
                c3 = conn3.cursor()
                
                c3.execute("SELECT id FROM tareas_evento WHERE evento_asociado = %s AND nombre_tarea = %s", (ev, t_base))
                if c3.fetchone():
                    if not messagebox.askyesno("Tarea Existente", f"Ya registraste:\n'{t_base}'\n\n¿Seguro que deseas agregarla otra vez para este mismo vehículo?", parent=v_prov):
                        liberar_conexion(conn3)
                        return

                c3.execute("SELECT COALESCE(MAX(orden), 0) FROM tareas_evento WHERE evento_asociado = %s", (ev,))
                nuevo_orden = c3.fetchone()[0] + 1
                
                c3.execute("""
                    INSERT INTO tareas_evento (evento_asociado, nombre_tarea, responsable, fecha_limite, estado, notas, orden, tipo_pago, tipo_entrada, repeticion, ubicacion, tiempo_aviso)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (ev, t_base, pr, f, "Pendiente", "Entrada registrada directamente desde el calendario.", nuevo_orden, "No aplica", "Mantenimiento", "No se repite", "", "Sin aviso"))
                conn3.commit()
                cache_sistema.invalidar()
                registrar_auditoria(self.usuario_activo, "Cronograma", f"Agendó mantenimiento: '{t_base}' para proveedor {pr}")
                
                self.cargar_tareas_tabla()
                
                ent_t.delete(0, tk.END)
                messagebox.showinfo("Éxito", "Entrada de mantenimiento guardada correctamente.", parent=v_prov)
            except Exception as ex:
                messagebox.showerror("Error", str(ex), parent=v_prov)
            finally:
                liberar_conexion(conn3)

        ctk.CTkButton(f_cont, text="💾 Guardar y Agregar Otra Entrada", font=("Arial", 12, "bold"), fg_color="#1f538d", hover_color="#163b65", command=guardar_prov).pack(fill="x", pady=10)

    def abrir_agendar_interno(self):
        v_int = ctk.CTkToplevel(self.parent_frame.winfo_toplevel())
        v_int.title("Agendar Trabajo de Oficina")
        v_int.geometry("380x380")
        v_int.transient(self.parent_frame.winfo_toplevel())
        v_int.grab_set()

        ctk.CTkLabel(v_int, text="🏢 NUEVA TAREA INTERNA", font=("Arial", 14, "bold"), text_color="#1f538d").pack(pady=(15, 10))

        f_cont = ctk.CTkFrame(v_int, fg_color="transparent")
        f_cont.pack(fill="both", expand=True, padx=20, pady=5)

        ctk.CTkLabel(f_cont, text="Título / Descripción de la Tarea:", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_t = ctk.CTkEntry(f_cont)
        ent_t.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(f_cont, text="Responsable / Área:", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_r = ctk.CTkEntry(f_cont)
        ent_r.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(f_cont, text="Fecha Programada (DD/MM/AAAA):", font=("Arial", 11, "bold")).pack(anchor="w")
        f_fecha = ctk.CTkFrame(f_cont, fg_color="transparent")
        f_fecha.pack(fill="x", pady=(0, 10))
        ent_f = ctk.CTkEntry(f_fecha, placeholder_text="Seleccione fecha...")
        ent_f.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(f_fecha, text="[ 📅 ]", width=40, fg_color="#1f538d", hover_color="#163b65", command=lambda: CalendarioNativo(v_int, ent_f)).pack(side="right", padx=(5, 0))

        def guardar_interno():
            t_base = ent_t.get().strip()
            r = ent_r.get().strip()
            f = ent_f.get().strip()
            
            if not t_base:
                messagebox.showwarning("Atención", "La descripción de la tarea es obligatoria.", parent=v_int)
                return

            conn = conectar_db()
            if not conn: return
            try:
                cursor = conn.cursor()
                
                cursor.execute("SELECT id FROM tareas_evento WHERE evento_asociado = %s AND nombre_tarea = %s", ("OFICINA | Trabajos Internos", t_base))
                if cursor.fetchone():
                    if not messagebox.askyesno("Entrada Existente", f"Ya registraste:\n'{t_base}'\n\n¿Seguro que deseas agregarla otra vez?", parent=v_int):
                        liberar_conexion(conn)
                        return

                cursor.execute("SELECT COALESCE(MAX(orden), 0) FROM tareas_evento WHERE evento_asociado = 'OFICINA | Trabajos Internos'")
                nuevo_orden = cursor.fetchone()[0] + 1
                
                cursor.execute("""
                    INSERT INTO tareas_evento (evento_asociado, nombre_tarea, responsable, fecha_limite, estado, notas, orden, tipo_pago, tipo_entrada, repeticion, ubicacion, tiempo_aviso)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, ("OFICINA | Trabajos Internos", t_base, r, f, "Pendiente", "Asignación interna de oficina.", nuevo_orden, "No aplica", "Tarea", "No se repite", "", "Sin aviso"))
                conn.commit()
                cache_sistema.invalidar()
                registrar_auditoria(self.usuario_activo, "Cronograma", f"Agendó trabajo interno: '{t_base}'")
                v_int.destroy()
                self.cargar_tareas_tabla()
                messagebox.showinfo("Éxito", "Trabajo interno agendado y visible en el calendario.")
            except Exception as ex:
                messagebox.showerror("Error", str(ex), parent=v_int)
            finally:
                liberar_conexion(conn)

        ctk.CTkButton(f_cont, text="💾 Agendar Tarea", font=("Arial", 12, "bold"), fg_color="#1f538d", hover_color="#163b65", command=guardar_interno).pack(fill="x", pady=10)


# =========================================================
# CLASE PRINCIPAL: GANTT Y CRONOGRAMA
# =========================================================
class CronogramaApp:
    def __init__(self, parent_frame):
        self.parent_frame = parent_frame
        self.usuario_activo = "Desconocido"
        self.pantalla_expandida = False
        
        self.inicializar_bd()
        self.crear_interfaz()

    # 🚀 FIX: AUTO-CURACIÓN EN SEGUNDO PLANO
    def inicializar_bd(self):
        def tarea_init():
            conn = conectar_db(silencioso=True)
            if not conn: return
            try:
                cursor = conn.cursor()
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS tareas_evento (
                        id SERIAL PRIMARY KEY,
                        evento_asociado VARCHAR(255),
                        nombre_tarea VARCHAR(255),
                        responsable VARCHAR(100),
                        fecha_limite VARCHAR(20),
                        estado VARCHAR(50),
                        notas TEXT
                    )
                """)
                conn.commit()
                try: cursor.execute("ALTER TABLE tareas_evento ADD COLUMN orden INTEGER DEFAULT 0"); conn.commit()
                except: conn.rollback()
                try: cursor.execute("ALTER TABLE tareas_evento ADD COLUMN tipo_pago VARCHAR(50) DEFAULT 'No aplica'"); conn.commit()
                except: conn.rollback()
                try: cursor.execute("ALTER TABLE tareas_evento ADD COLUMN archivo_pago TEXT DEFAULT ''"); conn.commit()
                except: conn.rollback()
                try: cursor.execute("UPDATE tareas_evento SET orden = id WHERE orden = 0"); conn.commit()
                except: conn.rollback()

                try: cursor.execute("ALTER TABLE tareas_evento ADD COLUMN tipo_entrada VARCHAR(50) DEFAULT 'Tarea'"); conn.commit()
                except: conn.rollback()
                try: cursor.execute("ALTER TABLE tareas_evento ADD COLUMN repeticion VARCHAR(50) DEFAULT 'No se repite'"); conn.commit()
                except: conn.rollback()
                try: cursor.execute("ALTER TABLE tareas_evento ADD COLUMN ubicacion TEXT DEFAULT ''"); conn.commit()
                except: conn.rollback()
                try: cursor.execute("ALTER TABLE tareas_evento ADD COLUMN tiempo_aviso VARCHAR(50) DEFAULT 'Sin aviso'"); conn.commit()
                except: conn.rollback()

                # 🚀 TABLA PARA TIPOS DE ENTRADA DINÁMICOS
                try:
                    cursor.execute("""
                        CREATE TABLE IF NOT EXISTS tipos_entrada_cronograma (
                            id SERIAL PRIMARY KEY,
                            nombre VARCHAR(255) UNIQUE NOT NULL
                        )
                    """)
                    conn.commit()
                    cursor.execute("SELECT COUNT(*) FROM tipos_entrada_cronograma")
                    if cursor.fetchone()[0] == 0:
                        cursor.execute("INSERT INTO tipos_entrada_cronograma (nombre) VALUES ('Tarea'), ('Mantenimiento'), ('Renovación Documento')")
                        conn.commit()
                except Exception:
                    conn.rollback()

            except Exception as e:
                print("Error BD Tareas:", e)
            finally:
                liberar_conexion(conn)

        threading.Thread(target=tarea_init, daemon=True).start()

    def abrir_calendario(self, entry_objetivo):
        CalendarioNativo(self.parent_frame.winfo_toplevel(), entry_objetivo)

    def abrir_calendario_dashboard(self):
        CalendarioDashboard(self.parent_frame.winfo_toplevel(), self.usuario_activo)

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
            self.f_form.pack(side="left", fill="y", padx=(0, 15), before=self.f_wrapper_derecha)
            self.btn_pantalla.configure(text="[ + ] Pantalla Completa", fg_color="#34495e", hover_color="#2c3e50")
            self.pantalla_expandida = False
        else:
            if sidebar: sidebar.pack_forget()
            self.f_form.pack_forget()
            self.btn_pantalla.configure(text="[ - ] Restaurar Vista", fg_color="#34495e", hover_color="#2c3e50")
            self.pantalla_expandida = True

    def cargar_tipos_combos(self, combobox, valor_default=None):
        def tarea():
            conn = conectar_db(silencioso=True)
            tipos = ["Tarea", "Mantenimiento", "Renovación Documento"] 
            if conn:
                try:
                    cursor = conn.cursor()
                    cursor.execute("SELECT nombre FROM tipos_entrada_cronograma ORDER BY id ASC")
                    tipos_db = [str(r[0]) for r in cursor.fetchall()]
                    if tipos_db:
                        tipos = tipos_db
                except Exception: pass
                finally: liberar_conexion(conn)
            
            if hasattr(self, 'parent_frame') and self.parent_frame.winfo_exists():
                self.parent_frame.after(0, lambda: self._actualizar_combo_tipos(combobox, tipos, valor_default))
        threading.Thread(target=tarea, daemon=True).start()

    def _actualizar_combo_tipos(self, combobox, tipos, valor_default):
        if combobox.winfo_exists():
            combobox.configure(values=tipos)
            if valor_default and valor_default in tipos:
                combobox.set(valor_default)
            elif not valor_default and "Tarea" in tipos:
                combobox.set("Tarea")
            elif tipos:
                combobox.set(tipos[0])

    def ventana_emergente_tipos(self, combo_widget, string_var, callback_toggle=None):
        v_cat = ctk.CTkToplevel(self.parent_frame.winfo_toplevel())
        v_cat.title("Tipos de Entrada")
        v_cat.geometry("400x500")
        v_cat.grab_set() 
        v_cat.resizable(False, False)
        
        ctk.CTkLabel(v_cat, text="Gestione los tipos de entrada:", font=("Arial", 12, "bold")).pack(pady=10)
        
        frame_lista = ctk.CTkFrame(v_cat)
        frame_lista.pack(fill="both", expand=True, padx=20, pady=5)
        
        lista_box = tk.Listbox(frame_lista, font=("Arial", 10), bg="#ffffff", fg="black", selectbackground="#1f538d", borderwidth=1)
        scroll_list = ttk.Scrollbar(frame_lista, orient="vertical", command=lista_box.yview)
        lista_box.configure(yscrollcommand=scroll_list.set)
        lista_box.pack(side="left", fill="both", expand=True, padx=(10,0), pady=10)
        scroll_list.pack(side="right", fill="y", pady=10, padx=(0,10))
        
        def cargar_lista():
            lista_box.delete(0, tk.END)
            conn = conectar_db(silencioso=True)
            if not conn: return
            try:
                cursor = conn.cursor()
                cursor.execute("SELECT nombre FROM tipos_entrada_cronograma ORDER BY id ASC")
                for row in cursor.fetchall():
                    lista_box.insert(tk.END, str(row[0]).strip())
            except Exception: pass
            finally: liberar_conexion(conn)
                
        cargar_lista()
        
        frame_ctrl = ctk.CTkFrame(v_cat, fg_color="transparent")
        frame_ctrl.pack(fill="x", padx=20, pady=5)
        
        ent_nueva = ctk.CTkEntry(frame_ctrl, width=180, placeholder_text="Nuevo tipo...")
        ent_nueva.pack(side="left", padx=5)
        
        def agregar():
            nueva = ent_nueva.get().strip()
            if nueva:
                conn = conectar_db()
                if not conn: return
                try:
                    cursor = conn.cursor()
                    cursor.execute("SELECT id FROM tipos_entrada_cronograma WHERE nombre ILIKE %s", (nueva,))
                    if cursor.fetchone():
                        messagebox.showerror("Error", "El tipo ya existe.", parent=v_cat)
                        return
                    cursor.execute("INSERT INTO tipos_entrada_cronograma (nombre) VALUES (%s)", (nueva,))
                    conn.commit()
                    cache_sistema.invalidar()
                    registrar_auditoria(self.usuario_activo, "Cronograma", f"Creó tipo de entrada '{nueva}'")
                    ent_nueva.delete(0, tk.END)
                    cargar_lista()
                    self.cargar_tipos_combos(combo_widget, string_var.get())
                except psycopg2.IntegrityError:
                    conn.rollback()
                    messagebox.showerror("Error", "Ya existe.", parent=v_cat)
                except Exception:
                    conn.rollback()
                finally:
                    liberar_conexion(conn)

        def eliminar():
            if not lista_box.curselection(): return
            tipo_a_borrar = lista_box.get(lista_box.curselection())
            
            if messagebox.askyesno("Confirmar", f"¿Eliminar el tipo '{tipo_a_borrar}'?", parent=v_cat):
                conn = conectar_db()
                if not conn: return
                try:
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM tipos_entrada_cronograma WHERE nombre = %s", (tipo_a_borrar,))
                    cursor.execute("UPDATE tareas_evento SET tipo_entrada = 'Tarea' WHERE tipo_entrada = %s", (tipo_a_borrar,))
                    conn.commit()
                    cache_sistema.invalidar()
                    registrar_auditoria(self.usuario_activo, "Cronograma", f"Eliminó tipo '{tipo_a_borrar}'")
                    
                    if string_var.get() == tipo_a_borrar:
                        string_var.set("Tarea")
                    
                    cargar_lista()
                    self.cargar_tipos_combos(combo_widget, string_var.get())
                except Exception:
                    conn.rollback()
                finally:
                    liberar_conexion(conn)

        def seleccionar():
            if lista_box.curselection():
                item = lista_box.get(lista_box.curselection())
                string_var.set(item)
                if callback_toggle:
                    callback_toggle()
                v_cat.destroy()

        btn_add = ctk.CTkButton(frame_ctrl, text="[ + ]", width=60, command=agregar)
        btn_add.pack(side="left", padx=2)
        
        btn_eli = ctk.CTkButton(frame_ctrl, text="[ - ]", width=60, fg_color="#D32F2F", hover_color="#B71C1C", command=eliminar)
        btn_eli.pack(side="left", padx=2)
        
        btn_sel = ctk.CTkButton(v_cat, text="[ OK ] Seleccionar", width=200, command=seleccionar)
        btn_sel.pack(pady=(15, 15))


    def crear_interfaz(self):
        self.frame_main = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.frame_main.pack(fill="both", expand=True, padx=15, pady=15)

        ctk.CTkLabel(self.frame_main, text="📅 CRONOGRAMA DE MANTENIMIENTO Y TAREAS", font=("Arial", 18, "bold"), text_color="#1f538d").pack(anchor="w", pady=(0, 10))

        f_top = ctk.CTkFrame(self.frame_main, fg_color="#f8f9fa", border_width=1, border_color="#e0e0e0", corner_radius=8)
        f_top.pack(fill="x", pady=(0, 15), ipadx=10, ipady=5)
        
        ctk.CTkLabel(f_top, text="Seleccione el Grupo / Proyecto de Trabajo:", font=("Arial", 12, "bold"), text_color="#333333").pack(side="left", padx=(10, 10), pady=10)
        self.combo_evento_global = ctk.CTkComboBox(f_top, width=400, state="readonly", command=self.cargar_tareas_tabla)
        self.combo_evento_global.pack(side="left", padx=10, pady=10)
        
        self.cargar_flota_activa()

        frame_split = ctk.CTkFrame(self.frame_main, fg_color="transparent")
        frame_split.pack(fill="both", expand=True)

        self.f_form = ctk.CTkScrollableFrame(frame_split, corner_radius=10, width=320)
        self.f_form.pack(side="left", fill="y", padx=(0, 15))

        ctk.CTkLabel(self.f_form, text="Gestión de Entradas", font=("Arial", 14, "bold"), text_color="#1f538d").pack(pady=(10, 15))

        ctk.CTkLabel(self.f_form, text="Tipo de Entrada:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        
        # 🚀 FIX: COMBOBOX DE TIPO DE ENTRADA CON BOTÓN GESTIONAR
        f_tipo_inline = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_tipo_inline.pack(fill="x", padx=10, pady=(0, 10))
        
        self.tipo_entrada_var = tk.StringVar(value="Tarea")
        self.combo_tipo_entrada = ctk.CTkComboBox(f_tipo_inline, variable=self.tipo_entrada_var, state="readonly", command=self.toggle_vista_google)
        self.combo_tipo_entrada.pack(side="left", fill="x", expand=True)
        
        btn_gestionar_tipo = ctk.CTkButton(f_tipo_inline, text="Gestionar", width=70, command=lambda: self.ventana_emergente_tipos(self.combo_tipo_entrada, self.tipo_entrada_var, self.toggle_vista_google))
        btn_gestionar_tipo.pack(side="right", padx=(5, 0))
        
        self.cargar_tipos_combos(self.combo_tipo_entrada)
        
        ctk.CTkLabel(self.f_form, text="Título / Descripción:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.ent_tarea = ctk.CTkEntry(self.f_form, placeholder_text="Ej: Cambio de aceite")
        self.ent_tarea.pack(fill="x", padx=10, pady=(0, 10))
        
        ctk.CTkLabel(self.f_form, text="Fecha Programada:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.f_fecha = ctk.CTkFrame(self.f_form, fg_color="transparent")
        self.f_fecha.pack(fill="x", padx=10, pady=(0, 10))
        self.ent_fecha = ctk.CTkEntry(self.f_fecha, placeholder_text="Seleccione fecha...")
        self.ent_fecha.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(self.f_fecha, text="[ 📅 ]", width=40, fg_color="#1f538d", hover_color="#163b65", command=lambda: self.abrir_calendario(self.ent_fecha)).pack(side="right", padx=(5, 0))

        self.f_repeticion = ctk.CTkFrame(self.f_form, fg_color="transparent")
        self.f_repeticion.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkLabel(self.f_repeticion, text="Repetición:", font=("Arial", 11, "bold")).pack(anchor="w")
        self.combo_repeticion = ctk.CTkComboBox(self.f_repeticion, values=["No se repite", "Diariamente", "Semanalmente", "Mensualmente", "Anualmente"], state="readonly")
        self.combo_repeticion.pack(fill="x")
        self.combo_repeticion.set("No se repite")

        self.f_evento_options = ctk.CTkFrame(self.f_form, fg_color="transparent")
        ctk.CTkLabel(self.f_evento_options, text="Ubicación / Taller:", font=("Arial", 11, "bold")).pack(anchor="w")
        self.ent_ubicacion = ctk.CTkEntry(self.f_evento_options, placeholder_text="Ej. Sede Central / Dirección")
        self.ent_ubicacion.pack(fill="x", pady=(0, 10))
        
        ctk.CTkLabel(self.f_evento_options, text="Aviso / Recordatorio:", font=("Arial", 11, "bold")).pack(anchor="w")
        self.combo_aviso = ctk.CTkComboBox(self.f_evento_options, values=["Sin aviso", "10 minutos antes", "1 hora antes", "1 día antes", "1 semana antes"], state="readonly")
        self.combo_aviso.pack(fill="x", pady=(0, 10))
        self.combo_aviso.set("Sin aviso")
        self.f_evento_options.pack(fill="x", padx=10, pady=0) 

        self.f_operativo = ctk.CTkFrame(self.f_form, fg_color="transparent")
        ctk.CTkLabel(self.f_operativo, text="Responsable Asignado:", font=("Arial", 11, "bold")).pack(anchor="w")
        self.ent_responsable = ctk.CTkEntry(self.f_operativo, placeholder_text="Ej. Juan Pérez / Taller")
        self.ent_responsable.pack(fill="x", pady=(0, 10))
        self.f_operativo.pack(fill="x", padx=10, pady=0)

        ctk.CTkLabel(self.f_form, text="Detalles / Notas Adicionales:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.txt_notas = ctk.CTkTextbox(self.f_form, height=80, border_width=1)
        self.txt_notas.pack(fill="x", padx=10, pady=(0, 15))

        self.btn_guardar = ctk.CTkButton(self.f_form, text="💾 Agregar Entrada", font=("Arial", 12, "bold"), fg_color="#1f538d", hover_color="#163b65", command=self.guardar_tarea)
        self.btn_guardar.pack(fill="x", padx=10, pady=(10, 5))

        self.btn_limpiar = ctk.CTkButton(self.f_form, text="🧹 Limpiar Formulario", font=("Arial", 12, "bold"), fg_color="#7f8c8d", hover_color="#606b6b", command=self.limpiar_formulario)
        self.btn_limpiar.pack(fill="x", padx=10, pady=5)

        self.toggle_vista_google()
        self.id_tarea_seleccionada = None

        self.f_wrapper_derecha = ctk.CTkFrame(frame_split, fg_color="transparent")
        self.f_wrapper_derecha.pack(side="right", fill="both", expand=True)

        ctk.CTkLabel(self.f_wrapper_derecha, text="Lista de Entradas (Seleccione para editar/eliminar/mover)", font=("Arial", 13, "bold")).pack(anchor="w", pady=(0, 5))

        f_tabla = ctk.CTkFrame(self.f_wrapper_derecha, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True)

        columnas = ("num", "id", "tipo_entrada", "tarea", "responsable", "fecha_limite", "notas", "repeticion", "ubicacion", "tiempo_aviso")
        style = ttk.Style()
        style.configure("Treeview", rowheight=30, font=("Arial", 10))
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        self.tabla = ttk.Treeview(f_tabla, columns=columnas, show="headings", style="Treeview")
        
        self.tabla.heading("num", text="N°")
        self.tabla.heading("id", text="ID (Oculto)")
        self.tabla.heading("tipo_entrada", text="Tipo")
        self.tabla.heading("tarea", text="Título / Descripción")
        self.tabla.heading("responsable", text="Responsable")
        self.tabla.heading("fecha_limite", text="Fecha Programada")
        
        self.tabla.column("num", width=40, anchor="center")
        self.tabla.column("id", width=0, stretch=tk.NO) 
        self.tabla.column("tipo_entrada", width=100, anchor="center")
        self.tabla.column("tarea", width=220, anchor="w")
        self.tabla.column("responsable", width=120, anchor="w")
        self.tabla.column("fecha_limite", width=120, anchor="center")

        for col in ["notas", "repeticion", "ubicacion", "tiempo_aviso"]:
            self.tabla.heading(col, text="")
            self.tabla.column(col, width=0, stretch=tk.NO)

        self.tabla.config(displaycolumns=("num", "tipo_entrada", "tarea", "responsable", "fecha_limite"))
        
        # 🚀 COLORES DE TABLA ADAPTADOS A FLOTA
        self.tabla.tag_configure("Tarea", background="#fff3cd", foreground="#856404")     
        self.tabla.tag_configure("Mantenimiento", background="#d1ecf1", foreground="#0c5460")   
        self.tabla.tag_configure("Renovación Documento", background="#e8daef", foreground="#8e44ad")    
        self.tabla.bind("<<TreeviewSelect>>", self.seleccionar_tarea_tabla)

        scroll_y = ttk.Scrollbar(f_tabla, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scroll_y.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        f_btn_tabla = ctk.CTkFrame(self.f_wrapper_derecha, fg_color="transparent")
        f_btn_tabla.pack(fill="x", pady=(15, 0))
        
        self.btn_pantalla = ctk.CTkButton(f_btn_tabla, text="[ + ] Pantalla Completa", font=("Arial", 12, "bold"), width=160, fg_color="#34495e", hover_color="#2c3e50", command=self.toggle_pantalla_completa)
        self.btn_pantalla.pack(side="left")

        btn_ver_calendario = ctk.CTkButton(f_btn_tabla, text="🗓️ Ver Calendario General", font=("Arial", 12, "bold"), width=180, fg_color="#1f538d", hover_color="#163b65", command=self.abrir_calendario_dashboard)
        btn_ver_calendario.pack(side="left", padx=(15, 5))

        btn_subir = ctk.CTkButton(f_btn_tabla, text="▲", width=40, fg_color="#e0e0e0", text_color="black", hover_color="#c8c8c8", command=lambda: self.mover_tarea("ARRIBA"))
        btn_subir.pack(side="left", padx=5)
        btn_bajar = ctk.CTkButton(f_btn_tabla, text="▼", width=40, fg_color="#e0e0e0", text_color="black", hover_color="#c8c8c8", command=lambda: self.mover_tarea("ABAJO"))
        btn_bajar.pack(side="left", padx=5)

        btn_editar = ctk.CTkButton(f_btn_tabla, text="✏️ Editar", font=("Arial", 12, "bold"), width=90, fg_color="#34495e", hover_color="#2c3e50", command=self.abrir_ventana_edicion)
        btn_editar.pack(side="left", padx=5)

        btn_eliminar = ctk.CTkButton(f_btn_tabla, text="❌ Eliminar", font=("Arial", 12, "bold"), width=90, fg_color="#e74c3c", hover_color="#c0392b", command=self.eliminar_tarea)
        btn_eliminar.pack(side="right")

        # 🚀 AL FINAL: CARGAR DATOS CUANDO LA UI ESTÉ LISTA
        self.cargar_flota_activa()

    def toggle_vista_google(self, choice=None):
        tipo = self.tipo_entrada_var.get()
        
        self.f_evento_options.pack_forget()
        self.f_operativo.pack_forget()
        
        if tipo == "Mantenimiento":
            self.f_evento_options.pack(fill="x", padx=10, pady=(0, 0), after=self.f_repeticion)
            self.f_operativo.pack(fill="x", padx=10, pady=(0, 0), after=self.f_evento_options)
        elif tipo == "Renovación Documento":
            self.combo_repeticion.set("Anualmente")
            self.f_operativo.pack(fill="x", padx=10, pady=(0, 0), after=self.f_repeticion)
        else:
            self.f_operativo.pack(fill="x", padx=10, pady=(0, 0), after=self.f_repeticion)

    def abrir_ventana_edicion(self):
        sel = self.tabla.selection()
        if not sel:
            messagebox.showwarning("Atención", "Seleccione una entrada de la lista para editar.")
            return

        valores = self.tabla.item(sel[0], "values")
        id_tarea = valores[1]
        
        v_edit = ctk.CTkToplevel(self.parent_frame.winfo_toplevel())
        v_edit.title(f"Editar Entrada ID: {id_tarea}")
        v_edit.geometry("400x550")
        v_edit.transient(self.parent_frame.winfo_toplevel())
        v_edit.grab_set()

        ctk.CTkLabel(v_edit, text="Modificar Detalles", font=("Arial", 14, "bold"), text_color="#1f538d").pack(pady=(15, 10))

        f_cont = ctk.CTkScrollableFrame(v_edit, fg_color="transparent")
        f_cont.pack(fill="both", expand=True, padx=20, pady=5)

        ctk.CTkLabel(f_cont, text="Tipo de Entrada:", font=("Arial", 11, "bold")).pack(anchor="w")
        
        # 🚀 FIX: COMBOBOX DE EDICIÓN CON BOTÓN GESTIONAR
        f_tipo_e_inline = ctk.CTkFrame(f_cont, fg_color="transparent")
        f_tipo_e_inline.pack(fill="x", pady=(0, 10))
        
        var_tipo_e = tk.StringVar(value=valores[2])
        cmb_tipo_e = ctk.CTkComboBox(f_tipo_e_inline, variable=var_tipo_e, state="readonly")
        cmb_tipo_e.pack(side="left", fill="x", expand=True)
        
        btn_gestionar_tipo_e = ctk.CTkButton(f_tipo_e_inline, text="Gestionar", width=80, command=lambda: self.ventana_emergente_tipos(cmb_tipo_e, var_tipo_e, toggle_edit))
        btn_gestionar_tipo_e.pack(side="right", padx=(5, 0))
        
        self.cargar_tipos_combos(cmb_tipo_e, var_tipo_e.get())

        ctk.CTkLabel(f_cont, text="Título / Descripción:", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_t = ctk.CTkEntry(f_cont)
        ent_t.pack(fill="x", pady=(0, 10))
        ent_t.insert(0, valores[3]) 

        ctk.CTkLabel(f_cont, text="Fecha Programada:", font=("Arial", 11, "bold")).pack(anchor="w")
        f_fecha_edit = ctk.CTkFrame(f_cont, fg_color="transparent")
        f_fecha_edit.pack(fill="x", pady=(0, 10))
        ent_f = ctk.CTkEntry(f_fecha_edit)
        ent_f.pack(side="left", fill="x", expand=True)
        ent_f.insert(0, valores[5]) 
        ctk.CTkButton(f_fecha_edit, text="[ 📅 ]", width=40, fg_color="#1f538d", hover_color="#163b65", command=lambda: self.abrir_calendario(ent_f)).pack(side="right", padx=(5, 0))

        f_rep = ctk.CTkFrame(f_cont, fg_color="transparent")
        f_rep.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(f_rep, text="Repetición:", font=("Arial", 11, "bold")).pack(anchor="w")
        cmb_rep = ctk.CTkComboBox(f_rep, values=["No se repite", "Diariamente", "Semanalmente", "Mensualmente", "Anualmente"], state="readonly")
        cmb_rep.pack(fill="x")
        cmb_rep.set(valores[7] if len(valores) > 7 and valores[7] else "No se repite")

        # FRAME MANTENIMIENTO
        f_ev = ctk.CTkFrame(f_cont, fg_color="transparent")
        ctk.CTkLabel(f_ev, text="Ubicación / Taller:", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_ub = ctk.CTkEntry(f_ev)
        ent_ub.pack(fill="x", pady=(0, 10))
        ent_ub.insert(0, valores[8] if len(valores) > 8 else "")
        
        ctk.CTkLabel(f_ev, text="Aviso / Recordatorio:", font=("Arial", 11, "bold")).pack(anchor="w")
        cmb_avi = ctk.CTkComboBox(f_ev, values=["Sin aviso", "10 minutos antes", "1 hora antes", "1 día antes", "1 semana antes"], state="readonly")
        cmb_avi.pack(fill="x", pady=(0, 10))
        cmb_avi.set(valores[9] if len(valores) > 9 and valores[9] else "Sin aviso")

        # FRAME OPERATIVO
        f_op = ctk.CTkFrame(f_cont, fg_color="transparent")
        ctk.CTkLabel(f_op, text="Responsable Asignado:", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_r = ctk.CTkEntry(f_op)
        ent_r.pack(fill="x", pady=(0, 10))
        ent_r.insert(0, valores[4]) 

        ctk.CTkLabel(f_cont, text="Notas Adicionales:", font=("Arial", 11, "bold")).pack(anchor="w")
        txt_n = ctk.CTkTextbox(f_cont, height=80, border_width=1)
        txt_n.pack(fill="x", pady=(0, 10))
        txt_n.insert("1.0", valores[6] if len(valores) > 6 else "") 

        def toggle_edit(choice=None):
            t = cmb_tipo_e.get()
            f_ev.pack_forget()
            f_op.pack_forget()
            
            if t == "Mantenimiento":
                f_ev.pack(fill="x", pady=(0, 0), after=f_rep)
                f_op.pack(fill="x", pady=(0, 0), after=f_ev)
            elif t == "Renovación Documento":
                cmb_rep.set("Anualmente")
                f_op.pack(fill="x", pady=(0, 0), after=f_rep)
            else:
                f_op.pack(fill="x", pady=(0, 0), after=f_rep)

        cmb_tipo_e.configure(command=toggle_edit)
        toggle_edit()

        def guardar_edicion():
            tipo_e = cmb_tipo_e.get()
            t_base = ent_t.get().strip()
            f = ent_f.get().strip()
            rep = cmb_rep.get()
            
            ub = ent_ub.get().strip() if tipo_e == "Mantenimiento" else ""
            avi = cmb_avi.get() if tipo_e == "Mantenimiento" else "Sin aviso"
            r = ent_r.get().strip() if tipo_e != "Renovación Documento" else ""
            n = txt_n.get("1.0", "end-1c").strip()
            
            if not t_base:
                messagebox.showwarning("Atención", "El título de la entrada es obligatorio.", parent=v_edit)
                return

            conn = conectar_db()
            if not conn: return
            try:
                cursor = conn.cursor()
                cursor.execute("""
                    UPDATE tareas_evento 
                    SET nombre_tarea=%s, responsable=%s, fecha_limite=%s, notas=%s, tipo_entrada=%s, repeticion=%s, ubicacion=%s, tiempo_aviso=%s
                    WHERE id=%s
                """, (t_base, r, f, n, tipo_e, rep, ub, avi, id_tarea))
                conn.commit()
                cache_sistema.invalidar()
                registrar_auditoria(self.usuario_activo, "Cronograma", f"Actualizó entrada ID {id_tarea}")
                v_edit.destroy()
                self.limpiar_formulario()
                self.cargar_tareas_tabla()
            except Exception as ex:
                messagebox.showerror("Error", str(ex), parent=v_edit)
            finally:
                liberar_conexion(conn)

        ctk.CTkButton(f_cont, text="💾 Guardar Cambios", font=("Arial", 12, "bold"), fg_color="#1f538d", hover_color="#163b65", command=guardar_edicion).pack(fill="x", pady=15)

    # 🚀 FIX: CARGA DIRECTA DESDE LA TABLA 'flota_vehiculos'
    def cargar_flota_activa(self):
        clave_cache = "lista_flotas_activas"
        eventos = cache_sistema.obtener(clave_cache)
        
        if eventos is not None:
            self._pintar_eventos(eventos)
        else:
            self.combo_evento_global.set("Cargando vehículos...")
            def tarea_eventos():
                evts = ["OFICINA | Trabajos Internos"]
                conn = conectar_db(silencioso=True)
                if conn:
                    try:
                        cursor = conn.cursor()
                        cursor.execute("SELECT placa, marca FROM flota_vehiculos ORDER BY placa ASC")
                        for r in cursor.fetchall():
                            evts.append(f"{r[0]} | {r[1]}")
                        
                        cache_sistema.guardar(clave_cache, evts)
                    except Exception as e:
                        print("Error descargando flota activa:", e)
                    finally:
                        liberar_conexion(conn)
                        
                if hasattr(self, 'parent_frame') and self.parent_frame.winfo_exists():
                    self.parent_frame.after(0, lambda: self._pintar_eventos(evts))

            threading.Thread(target=tarea_eventos, daemon=True).start()

    def _pintar_eventos(self, eventos):
        if eventos and hasattr(self, 'combo_evento_global') and self.combo_evento_global.winfo_exists():
            self.combo_evento_global.configure(values=eventos)
            self.combo_evento_global.set(eventos[0])
            self.cargar_tareas_tabla()

    def limpiar_formulario(self):
        self.id_tarea_seleccionada = None
        self.tipo_entrada_var.set("Tarea")
        self.ent_tarea.delete(0, tk.END)
        self.ent_responsable.delete(0, tk.END)
        self.ent_fecha.delete(0, tk.END)
        self.txt_notas.delete("1.0", tk.END)
        self.combo_repeticion.set("No se repite")
        self.ent_ubicacion.delete(0, tk.END)
        self.combo_aviso.set("Sin aviso")
        self.btn_guardar.configure(text="💾 Agregar Entrada", fg_color="#1f538d", hover_color="#163b65")
        self.toggle_vista_google()

    def seleccionar_tarea_tabla(self, event):
        sel = self.tabla.selection()
        if not sel: return
        valores = self.tabla.item(sel[0], "values")
        
        self.id_tarea_seleccionada = valores[1] 
        self.tipo_entrada_var.set(valores[2]) 
        self.toggle_vista_google()
        
        self.ent_tarea.delete(0, tk.END)
        self.ent_tarea.insert(0, valores[3])
        
        self.ent_responsable.delete(0, tk.END)
        self.ent_responsable.insert(0, valores[4]) 
        
        self.ent_fecha.delete(0, tk.END)
        self.ent_fecha.insert(0, valores[5]) 
        
        self.txt_notas.delete("1.0", tk.END)
        self.txt_notas.insert("1.0", valores[6]) 

        self.combo_repeticion.set(valores[7] if len(valores) > 7 and valores[7] else "No se repite")
        
        self.ent_ubicacion.delete(0, tk.END)
        self.ent_ubicacion.insert(0, valores[8] if len(valores) > 8 else "")
        
        self.combo_aviso.set(valores[9] if len(valores) > 9 and valores[9] else "Sin aviso")

        self.btn_guardar.configure(text="✏️ Guardar Cambios", fg_color="#34495e", hover_color="#2c3e50")

    def mover_tarea(self, direccion):
        sel = self.tabla.selection()
        if not sel: return
        
        items = self.tabla.get_children()
        idx_act = items.index(sel[0])
        
        if direccion == "ARRIBA" and idx_act > 0:
            idx_dest = idx_act - 1
        elif direccion == "ABAJO" and idx_act < len(items) - 1:
            idx_dest = idx_act + 1
        else:
            return
            
        id_act = self.tabla.item(items[idx_act], "values")[1] 
        id_dest = self.tabla.item(items[idx_dest], "values")[1] 
        
        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT orden FROM tareas_evento WHERE id = %s", (id_act,))
            ord_act = cursor.fetchone()[0]
            cursor.execute("SELECT orden FROM tareas_evento WHERE id = %s", (id_dest,))
            ord_dest = cursor.fetchone()[0]
            
            cursor.execute("UPDATE tareas_evento SET orden = %s WHERE id = %s", (ord_dest, id_act))
            cursor.execute("UPDATE tareas_evento SET orden = %s WHERE id = %s", (ord_act, id_dest))
            conn.commit()
            cache_sistema.invalidar()
            self.cargar_tareas_tabla()
            
            for child in self.tabla.get_children():
                if str(self.tabla.item(child, "values")[1]) == str(id_act):
                    self.tabla.selection_set(child)
                    break
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo mover la tarea:\n{e}")
        finally:
            liberar_conexion(conn)

    def guardar_tarea(self):
        evento = self.combo_evento_global.get()
        if "Sin vehículos" in evento or "Cargando" in evento or not evento.strip():
            messagebox.showwarning("Atención", "Debe seleccionar un vehículo válido.")
            return

        tipo_entrada = self.tipo_entrada_var.get()
        tarea = self.ent_tarea.get().strip()
        fecha_limite = self.ent_fecha.get().strip()
        repeticion = self.combo_repeticion.get()
        ubicacion = self.ent_ubicacion.get().strip() if tipo_entrada == "Mantenimiento" else ""
        tiempo_aviso = self.combo_aviso.get() if tipo_entrada == "Mantenimiento" else "Sin aviso"
        responsable = self.ent_responsable.get().strip() if tipo_entrada != "Renovación Documento" else ""
        notas = self.txt_notas.get("1.0", "end-1c").strip()

        if not tarea:
            messagebox.showwarning("Atención", "El Título / Descripción es obligatorio.")
            return

        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            
            if not self.id_tarea_seleccionada:
                cursor.execute("SELECT id FROM tareas_evento WHERE evento_asociado = %s AND nombre_tarea = %s", (evento, tarea))
                if cursor.fetchone():
                    if not messagebox.askyesno("Entrada Existente", f"Ya registraste la entrada:\n'{tarea}'\n\n¿Seguro que deseas agregarla otra vez?", parent=self.parent_frame.winfo_toplevel()):
                        liberar_conexion(conn)
                        return
                        
            if self.id_tarea_seleccionada:
                cursor.execute("""
                    UPDATE tareas_evento 
                    SET nombre_tarea=%s, responsable=%s, fecha_limite=%s, notas=%s, tipo_entrada=%s, repeticion=%s, ubicacion=%s, tiempo_aviso=%s
                    WHERE id=%s
                """, (tarea, responsable, fecha_limite, notas, tipo_entrada, repeticion, ubicacion, tiempo_aviso, self.id_tarea_seleccionada))
                registrar_auditoria(self.usuario_activo, "Cronograma", f"Actualizó la entrada '{tarea}'")
            else:
                cursor.execute("SELECT COALESCE(MAX(orden), 0) FROM tareas_evento WHERE evento_asociado = %s", (evento,))
                nuevo_orden = cursor.fetchone()[0] + 1
                
                cursor.execute("""
                    INSERT INTO tareas_evento (evento_asociado, nombre_tarea, responsable, fecha_limite, estado, notas, orden, tipo_pago, archivo_pago, tipo_entrada, repeticion, ubicacion, tiempo_aviso)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (evento, tarea, responsable, fecha_limite, "Pendiente", notas, nuevo_orden, "No aplica", "", tipo_entrada, repeticion, ubicacion, tiempo_aviso))
                registrar_auditoria(self.usuario_activo, "Cronograma", f"Creó nueva entrada '{tarea}'")
            
            conn.commit()
            cache_sistema.invalidar()
            self.limpiar_formulario()
            self.cargar_tareas_tabla()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar la entrada:\n{e}")
        finally:
            liberar_conexion(conn)

    # 🚀 FIX: CARGA DE TAREAS ASÍNCRONA
    def cargar_tareas_tabla(self, choice=None):
        if not hasattr(self, 'tabla') or not self.tabla.winfo_exists(): return
        for item in self.tabla.get_children(): self.tabla.delete(item)
        evento_seleccionado = self.combo_evento_global.get()
        if "Sin vehículos" in evento_seleccionado or "Cargando" in evento_seleccionado or not evento_seleccionado.strip(): return

        clave_cache = f"tareas_evento_{evento_seleccionado}"
        datos = cache_sistema.obtener(clave_cache)
        
        if datos is not None:
            self._pintar_tareas_tabla(datos)
        else:
            self.tabla.insert("", tk.END, values=("", "", "", "Cargando datos...", "", "", "", "", "", ""))
            def tarea_descarga():
                datos_db = []
                conn = conectar_db(silencioso=True)
                if conn:
                    try:
                        cursor = conn.cursor()
                        cursor.execute("""
                            SELECT id, tipo_entrada, nombre_tarea, responsable, fecha_limite, notas, repeticion, ubicacion, tiempo_aviso 
                            FROM tareas_evento 
                            WHERE evento_asociado = %s 
                            ORDER BY orden ASC, id ASC
                        """, (evento_seleccionado,))
                        
                        datos_db = cursor.fetchall()
                        cache_sistema.guardar(clave_cache, datos_db)
                    except Exception as e:
                        print("Error leyendo tareas:", e)
                    finally:
                        liberar_conexion(conn)
                if hasattr(self, 'parent_frame') and self.parent_frame.winfo_exists():
                    self.parent_frame.after(0, lambda: self._pintar_tareas_tabla(datos_db))
                
            threading.Thread(target=tarea_descarga, daemon=True).start()

    def _pintar_tareas_tabla(self, datos):
        if not hasattr(self, 'tabla') or not self.tabla.winfo_exists(): return
        for item in self.tabla.get_children(): self.tabla.delete(item)
        contador_visual = 1
        for r in datos:
            tipo = r[1]
            valores = (contador_visual, r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8])
            self.tabla.insert("", tk.END, values=valores, tags=(tipo,))
            contador_visual += 1

    def eliminar_tarea(self):
        if not self.id_tarea_seleccionada:
            messagebox.showwarning("Atención", "Debe seleccionar una entrada de la lista primero.")
            return

        if messagebox.askyesno("Confirmar", "¿Desea eliminar permanentemente esta entrada del cronograma?"):
            conn = conectar_db()
            if not conn: return
            try:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM tareas_evento WHERE id = %s", (self.id_tarea_seleccionada,))
                conn.commit()
                cache_sistema.invalidar()
                registrar_auditoria(self.usuario_activo, "Cronograma", f"Eliminó la entrada ID {self.id_tarea_seleccionada}")
                self.limpiar_formulario()
                self.cargar_tareas_tabla()
            except Exception as e:
                messagebox.showerror("Error", str(e))
            finally:
                liberar_conexion(conn)

if __name__ == "__main__":
    pass