# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import os
import sys  
import json
import subprocess 
import webbrowser
import customtkinter as ctk
from conexion import conectar_db, liberar_conexion

# =========================================================
# 🚀 ADAPTACIÓN MULTIPLATAFORMA: Función universal para abrir archivos
# =========================================================
def abrir_documento(ruta):
    try:
        if sys.platform == "win32":
            os.startfile(ruta)
        elif sys.platform == "darwin": # macOS
            subprocess.call(["open", ruta])
        else: # Linux
            subprocess.call(["xdg-open", ruta])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")

# =========================================================
# 🚀 MOTOR DE CONFIGURACIÓN REGIONAL
# =========================================================
def cargar_configuracion_regional():
    config = {
        "simbolo_moneda": "S/.",
        "formato_numero": "1,000.00",
        "cuentas_bancarias": []
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

class LibroDiarioApp:
    def __init__(self, root):
        self.root = root
        self.pantalla_expandida = False
        
        try:
            self.root.configure(fg_color="#f8f9fa")
        except Exception:
            try: self.root.configure(bg="#f8f9fa")
            except: pass
        
        if hasattr(self.root, 'title'):
            self.root.title("📖 Libro Diario - Black Cube")
        if hasattr(self.root, 'geometry') and isinstance(self.root, (tk.Tk, tk.Toplevel)):
            self.root.geometry("1200x650")
            self.root.update_idletasks()
            ancho = self.root.winfo_width()
            alto = self.root.winfo_height()
            x = (self.root.winfo_screenwidth() // 2) - (ancho // 2)
            y = (self.root.winfo_screenheight() // 2) - (alto // 2)
            self.root.geometry(f"{ancho}x{alto}+{x}+{y}")

        style = ttk.Style()
        style.theme_use("clam")
        
        style.configure("Treeview", background="#ffffff", foreground="#000000", fieldbackground="#ffffff", bordercolor="#e0e0e0", borderwidth=1, rowheight=28, font=("Arial", 10))
        style.map("Treeview", background=[("selected", "#1f538d")], foreground=[("selected", "#ffffff")])
        style.configure("Treeview.Heading", background="#f0f0f0", foreground="#000000", font=("Arial", 11, "bold"), bordercolor="#e0e0e0", borderwidth=1, relief="flat")
        
        # 🚀 HEADER CON BOTÓN DE PANTALLA COMPLETA Y FILTROS
        header_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        header_frame.pack(fill="x", padx=15, pady=(15, 5))
        
        ctk.CTkLabel(header_frame, text="LIBRO DIARIO GENERAL DE OPERACIONES", font=("Arial", 16, "bold"), text_color="#1f538d").pack(side="left")
        
        self.btn_pantalla = ctk.CTkButton(header_frame, text="[ + ] Pantalla Completa", font=("Arial", 12, "bold"), width=160, fg_color="#34495e", hover_color="#2c3e50", command=self.toggle_pantalla_completa)
        self.btn_pantalla.pack(side="right", padx=(10, 0))

        # --- FILTRO 1: CUENTA O FORMA DE PAGO ---
        bancos_guardados = CONFIG_REGIONAL.get("cuentas_bancarias", [])
        lista_cuentas = ["Cualquier Cuenta"]
        for b in bancos_guardados:
            banco_nom = b.get("banco", "").strip()
            cuenta_num = b.get("cuenta", "").strip()
            if banco_nom or cuenta_num:
                lista_cuentas.append(f"{banco_nom} - {cuenta_num}".strip(" - "))
        lista_cuentas.extend(["Efectivo / Caja Chica", "Tarjeta de Crédito", "Tarjeta de Débito", "Cheque", "Otro", "No Especificado", "Reversión (N.C.)"])

        self.combo_cuenta = ctk.CTkComboBox(header_frame, values=lista_cuentas, width=170, command=self.aplicar_filtro)
        self.combo_cuenta.pack(side="right", padx=(5, 10))
        self.combo_cuenta.set("Cualquier Cuenta")
        ctk.CTkLabel(header_frame, text="Cuenta / Método:", font=("Arial", 11, "bold")).pack(side="right", padx=(10, 2))

        # --- FILTRO 2: TIPO DE ASIENTO ---
        self.combo_filtro = ctk.CTkComboBox(header_frame, values=["Ver Todos (General)", "Solo Ingresos (Ventas/Cobros)", "Solo Egresos (Compras/Pagos)", "Solo Anulaciones (N.C.)"], width=190, command=self.aplicar_filtro)
        self.combo_filtro.pack(side="right", padx=(5, 5))
        self.combo_filtro.set("Ver Todos (General)")
        ctk.CTkLabel(header_frame, text="Filtro Asientos:", font=("Arial", 11, "bold")).pack(side="right", padx=(10, 2))
        # ---------------------------

        frame_principal = tk.Frame(root, bg="#f8f9fa")
        frame_principal.pack(fill="both", expand=True, padx=15, pady=5)
        
        frame_acciones = tk.LabelFrame(frame_principal, text=" Panel de Acciones ", font=("Arial", 11, "bold"), fg="#34495e", bg="#f8f9fa", padx=10, pady=10)
        frame_acciones.pack(side="right", fill="y", padx=(15, 0))
        
        btn_refresh = ctk.CTkButton(frame_acciones, text="🔄 Actualizar Tabla", font=("Arial", 12, "bold"), fg_color="#7f8c8d", hover_color="#606b6b", command=self.cargar_datos_diario)
        btn_refresh.pack(fill="x", pady=(10, 5))

        btn_export = ctk.CTkButton(frame_acciones, text="📊 Exportar Excel", font=("Arial", 12, "bold"), fg_color="#27ae60", hover_color="#1e8449", command=self.exportar_excel)
        btn_export.pack(fill="x", pady=5)
        
        frame_tabla = tk.Frame(frame_principal, bg="#f8f9fa")
        frame_tabla.pack(side="left", fill="both", expand=True)
        
        lbl_hint = tk.Label(frame_tabla, text="💡 Haz doble clic sobre cualquier asiento contable para visualizar su comprobante o PDF de SUNAT asociado.", font=("Arial", 10, "italic"), fg="gray", bg="#f8f9fa")
        lbl_hint.pack(side="bottom", anchor="w", pady=(5, 0))

        columnas = ("id", "origen", "fecha", "codigo", "concepto", "tipo", "cuenta", "debe", "haber", "ruta_archivo", "categoria_real")
        self.tabla = ttk.Treeview(frame_tabla, columns=columnas, show="headings", style="Treeview")
        
        # --- CONFIGURACIÓN DE ETIQUETAS DE COLOR ---
        self.tabla.tag_configure("ingreso", background="#e8f8f5", foreground="#0e6251") 
        self.tabla.tag_configure("egreso", background="#fdedec", foreground="#7b241c")  
        self.tabla.tag_configure("anulacion", background="#fef9e7", foreground="#d35400") 
        # -------------------------------------------

        self.tabla.heading("fecha", text="Fecha Reg.")
        self.tabla.heading("codigo", text="Doc. / Factura") 
        self.tabla.heading("concepto", text="Beneficiario / Concepto")
        self.tabla.heading("tipo", text="Categoría Suministro")
        self.tabla.heading("cuenta", text="Cuenta / Pago")
        self.tabla.heading("debe", text="Debe (Ingresos)")
        self.tabla.heading("haber", text="Haber (Egresos)")
        
        self.tabla["displaycolumns"] = ("fecha", "codigo", "concepto", "tipo", "cuenta", "debe", "haber")
        
        self.tabla.column("fecha", width=100, anchor="center")
        self.tabla.column("codigo", width=110, anchor="center")
        self.tabla.column("concepto", width=220, anchor="w")
        self.tabla.column("tipo", width=150, anchor="center")
        self.tabla.column("cuenta", width=140, anchor="center")
        self.tabla.column("debe", width=115, anchor="e")
        self.tabla.column("haber", width=115, anchor="e")
        
        scr_y = ttk.Scrollbar(frame_tabla, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scr_y.set)
        
        self.tabla.pack(side="left", fill="both", expand=True)
        scr_y.pack(side="right", fill="y")
        
        self.tabla.bind("<Double-1>", self.abrir_documento_asiento)
        
        self.frame_totales = tk.Frame(root, bg="#f8f9fa")
        self.frame_totales.pack(fill="x", side="bottom", pady=(10, 15))
        
        self.lbl_total_debe = tk.Label(self.frame_totales, text="Total Debe: 0.00", font=("Arial", 14, "bold"), fg="#27ae60", bg="#f8f9fa")
        self.lbl_total_debe.pack(side="left", padx=30)
        
        self.lbl_saldo_neto = tk.Label(self.frame_totales, text="Saldo Neto: 0.00", font=("Arial", 16, "bold"), fg="#1f538d", bg="#f8f9fa")
        self.lbl_saldo_neto.pack(side="left", expand=True)
        
        self.lbl_total_haber = tk.Label(self.frame_totales, text="Total Haber: 0.00", font=("Arial", 14, "bold"), fg="#c0392b", bg="#f8f9fa")
        self.lbl_total_haber.pack(side="right", padx=30)
        
        # Variable para almacenar todos los movimientos en memoria (para los filtros)
        self.todos_los_movimientos = []

        if hasattr(self.root, 'after'):
            self.root.after(100, self.cargar_datos_diario)
        else:
            self.cargar_datos_diario()

    def toggle_pantalla_completa(self):
        sidebar = None
        try:
            if self.root.master:
                for child in self.root.master.winfo_children():
                    if hasattr(child, "cget") and child.cget("width") == 280:
                        sidebar = child
                        break
        except Exception: pass

        if getattr(self, "pantalla_expandida", False):
            if sidebar: sidebar.pack(side="left", fill="y", before=self.root)
            self.btn_pantalla.configure(text="[ + ] Pantalla Completa", fg_color="#34495e", hover_color="#2c3e50")
            self.pantalla_expandida = False
        else:
            if sidebar: sidebar.pack_forget()
            self.btn_pantalla.configure(text="[ - ] Restaurar Vista", fg_color="#34495e", hover_color="#2c3e50")
            self.pantalla_expandida = True

    def abrir_documento_asiento(self, event):
        seleccion = self.tabla.selection()
        if not seleccion: return
        valores = self.tabla.item(seleccion[0], "values")
        
        if len(valores) > 9:
            ruta = str(valores[9]).strip()
            if ruta and ruta != "None":
                if ruta.startswith("http"):
                    webbrowser.open(ruta)
                elif os.path.exists(ruta):
                    abrir_documento(ruta)
                else:
                    messagebox.showinfo("Aviso", "El archivo físico ha sido movido o ya no se encuentra en la ruta especificada.")
            else:
                messagebox.showinfo("Aviso", "Este asiento no tiene un documento digital o PDF asociado en la base de datos.")

    def exportar_excel(self):
        try:
            import pandas as pd
        except ImportError:
            return messagebox.showerror("Librería Faltante", "Instala pandas: pip install pandas openpyxl")
        
        filas = []
        for item in self.tabla.get_children():
            vals = self.tabla.item(item)["values"]
            filas.append([vals[2], vals[3], vals[4], vals[5], vals[6], vals[7], vals[8]])

        if not filas: return messagebox.showwarning("Aviso", "No hay datos para exportar con el filtro actual.")
        
        cols = ["Fecha", "Doc. / Factura", "Beneficiario / Concepto", "Categoría Suministro", "Cuenta / Pago", "Debe (Ingresos)", "Haber (Egresos)"]
        
        ruta = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Libro_Diario_BlackCube.xlsx", filetypes=[("Excel", "*.xlsx")])
        if ruta:
            try:
                pd.DataFrame(filas, columns=cols).to_excel(ruta, index=False)
                messagebox.showinfo("Éxito", f"Reporte exportado exitosamente a:\n{ruta}")
                abrir_documento(ruta)
            except Exception as e:
                messagebox.showerror("Error", f"Fallo al guardar:\n{e}")

    def cargar_datos_diario(self):
        self.todos_los_movimientos = []
        conn = conectar_db()
        if not conn: 
            self.aplicar_filtro()
            return
            
        try:
            cursor = conn.cursor()
            
            # Garantizar que la columna cuenta_destino exista (Previene errores si es bd vieja)
            try:
                cursor.execute("ALTER TABLE pagos_clientes ADD COLUMN cuenta_destino VARCHAR(255) DEFAULT ''")
                conn.commit()
            except Exception:
                conn.rollback()

            # 🚀 1. INGRESOS (CLIENTES) -> DEBE
            # OJO: Se usa INNER JOIN para que si la factura se eliminó en Ventas, no aparezcan pagos huérfanos.
            try:
                cursor.execute("""
                    SELECT p.id, 'cliente', p.fecha_pago, 
                           COALESCE(f.numero_documento, p.codigo_cotizacion, 'Sin Ref'), 
                           p.cliente_nombre, 'Cobro Cliente', p.monto_pagado,
                           COALESCE(NULLIF(p.archivo_ruta, ''), f.enlace_pdf_sunat, ''),
                           COALESCE(NULLIF(p.cuenta_destino, ''), 'No Especificado')
                    FROM pagos_clientes p
                    INNER JOIN facturas_emitidas f ON p.id_factura = f.id
                """)
                for r in cursor.fetchall():
                    fecha = r[2] if r[2] and str(r[2]).strip() != "" else "Sin Fecha"
                    doc_ref = str(r[3]).strip()
                    if not doc_ref or doc_ref.lower() == "none": doc_ref = "Sin Ref"
                    monto = float(r[6]) if r[6] else 0.0
                    ruta_arch = r[7] if r[7] else ""
                    cuenta_pago = r[8] if r[8] else "No Especificado"
                    self.todos_los_movimientos.append((r[0], r[1], fecha, doc_ref, r[4], r[5], monto, 0.0, ruta_arch, 'ingreso', cuenta_pago))
            except Exception as e:
                conn.rollback()
                print("Error Ingresos:", e)
                
            # 🚀 2. EGRESOS (PROVEEDORES) -> HABER
            # OJO: INNER JOIN para omitir registros huérfanos de facturas borradas en Compras.
            try:
                cursor.execute("""
                    SELECT p.id, 'proveedor', p.fecha_pago, 
                           COALESCE(f.numero_documento, p.codigo_cotizacion, 'Sin Ref'), 
                           p.proveedor_nombre, p.categoria_suministro, p.monto_pagado,
                           COALESCE(NULLIF(p.archivo_ruta, ''), f.archivo_ruta, ''),
                           COALESCE(NULLIF(p.cuenta_origen, ''), 'No Especificado')
                    FROM pagos_comprobantes p
                    INNER JOIN facturas_recibidas f ON p.id_factura = f.id
                """)
                for r in cursor.fetchall():
                    fecha = r[2] if r[2] and str(r[2]).strip() != "" else "Sin Fecha"
                    doc_ref = str(r[3]).strip()
                    if not doc_ref or doc_ref.lower() == "none": doc_ref = "Sin Ref"
                    monto = float(r[6]) if r[6] else 0.0
                    ruta_arch = r[7] if r[7] else ""
                    cuenta_pago = r[8] if r[8] else "No Especificado"
                    self.todos_los_movimientos.append((r[0], r[1], fecha, doc_ref, r[4], r[5], 0.0, monto, ruta_arch, 'egreso', cuenta_pago))
            except Exception as e:
                conn.rollback()
                print("Error Egresos:", e)
            
            # 🚀 3. NOTAS DE CRÉDITO Y FACTURAS ANULADAS -> ANULACION
            try:
                cursor.execute("""
                    SELECT id, fecha, numero_documento, cliente, total, COALESCE(det_monto, 0), 
                           enlace_pdf_sunat, enlace_pdf_nc, archivo_ruta, tipo_documento
                    FROM facturas_emitidas
                    WHERE estado_sunat LIKE '%Anulada%'
                """)
                for r in cursor.fetchall():
                    id_fac = r[0]
                    fecha = r[1] if r[1] and str(r[1]).strip() != "" else "Sin Fecha"
                    nro_doc = str(r[2]).strip() if r[2] else "S/N"
                    cliente = str(r[3]).strip()
                    tot_bruto = float(r[4]) if r[4] else 0.0
                    det_monto = float(r[5]) if r[5] else 0.0
                    neto = tot_bruto - det_monto
                    
                    pdf_fac = r[6] if r[6] else (r[8] if r[8] else "")
                    pdf_nc = r[7] if r[7] else ""
                    tipo_doc = r[9] if r[9] else ""
                    
                    serie_nc = "FC01" if "Factura" in tipo_doc else "BC01"
                    num_orig = nro_doc.split("-")[1] if "-" in nro_doc else "1"
                    doc_nc = f"{serie_nc}-{num_orig}"
                    
                    # Fila 1: Factura Original (Suma al Debe/Ingresos pero pintada como anulación)
                    self.todos_los_movimientos.append((id_fac, 'factura_anulada', fecha, nro_doc, f"{cliente} (FACTURA ANULADA)", "Ingreso Revertido", neto, 0.0, pdf_fac, 'anulacion', "Reversión (N.C.)"))
                    
                    # Fila 2: Nota de Crédito (Suma al Haber/Egresos para cruzar contablemente)
                    self.todos_los_movimientos.append((id_fac, 'nota_credito', fecha, doc_nc, f"{cliente} (NOTA DE CRÉDITO)", "Anulación de Ingreso", 0.0, neto, pdf_nc, 'anulacion', "Reversión (N.C.)"))
            except Exception:
                conn.rollback()
                
        except Exception as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudo compilar el diario:\n{str(e)}")
        finally:
            liberar_conexion(conn)
            
        # Ordenar cronológicamente por la fecha (índice 2)
        self.todos_los_movimientos.sort(key=lambda x: x[2])
        self.aplicar_filtro()

    def aplicar_filtro(self, event=None):
        filtro_asiento = self.combo_filtro.get()
        filtro_cuenta = self.combo_cuenta.get()
        
        self.tabla.config(displaycolumns="")
        for fila in self.tabla.get_children():
            self.tabla.delete(fila)
            
        total_debe = 0.0
        total_haber = 0.0
        
        for mov in self.todos_los_movimientos:
            cat_real = mov[9]
            cuenta_asociada = str(mov[10]).strip()
            
            # 1. Filtro Lógico por Tipo (Ingreso / Egreso)
            mostrar_tipo = False
            if filtro_asiento == "Ver Todos (General)":
                mostrar_tipo = True
            elif filtro_asiento == "Solo Ingresos (Ventas/Cobros)" and cat_real == 'ingreso':
                mostrar_tipo = True
            elif filtro_asiento == "Solo Egresos (Compras/Pagos)" and cat_real == 'egreso':
                mostrar_tipo = True
            elif filtro_asiento == "Solo Anulaciones (N.C.)" and cat_real == 'anulacion':
                mostrar_tipo = True
                
            # 2. Filtro Lógico por Cuenta / Método de Pago
            mostrar_cuenta = False
            if filtro_cuenta == "Cualquier Cuenta":
                mostrar_cuenta = True
            elif filtro_cuenta in cuenta_asociada:
                mostrar_cuenta = True
                
            if mostrar_tipo and mostrar_cuenta:
                str_debe = formatear_moneda(mov[6]) if mov[6] > 0 else "-"
                str_haber = formatear_moneda(mov[7]) if mov[7] > 0 else "-"
                
                # Inserción con etiqueta de color correspondiente
                self.tabla.insert("", tk.END, values=(
                    mov[0], mov[1], mov[2], mov[3], mov[4], mov[5], cuenta_asociada, str_debe, str_haber, mov[8], cat_real
                ), tags=(cat_real,))
                
                total_debe += mov[6]
                total_haber += mov[7]
            
        self.lbl_total_debe.config(text=f"Total Debe (Ingresos): {formatear_moneda(total_debe)}")
        self.lbl_total_haber.config(text=f"Total Haber (Egresos): {formatear_moneda(total_haber)}")
        
        # Calcular saldo neto dinámico según lo que se esté mostrando
        saldo_neto = total_debe - total_haber
        color_saldo = "#27ae60" if saldo_neto >= 0 else "#c0392b"
        self.lbl_saldo_neto.config(text=f"Saldo Neto: {formatear_moneda(saldo_neto)}", fg=color_saldo)
        
        self.tabla.config(displaycolumns=("fecha", "codigo", "concepto", "tipo", "cuenta", "debe", "haber"))

if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style()
    style.theme_use("clam")
    app = LibroDiarioApp(root)
    root.mainloop()