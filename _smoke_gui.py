# -*- coding: utf-8 -*-
import io, os, sys, time, traceback, warnings
warnings.filterwarnings("ignore")
sys.path.insert(0, r"C:\\Users\\Alberto\\Desktop\\Proyecto Control Flota 140926")
registro_dialogos = []
out = io.StringIO()
def p(*a): out.write(" ".join(str(x) for x in a) + "\n")

import tkinter as tk
from tkinter import messagebox, simpledialog
import customtkinter as ctk

def _falso(titulo="", mensaje="", **kw):
    registro_dialogos.append(("info/warn", str(titulo), str(mensaje)[:400]))
    return True
messagebox.showinfo = _falso
messagebox.showerror = lambda t="", m="", **k: (registro_dialogos.append(("ERROR", str(t), str(m)[:800])), True)[1]
messagebox.showwarning = lambda t="", m="", **k: (registro_dialogos.append(("WARN", str(t), str(m)[:400])), True)[1]
messagebox.askyesno = lambda t="", m="", **k: (registro_dialogos.append(("askyesno", str(t), str(m)[:200])), True)[1]
simpledialog.askstring = lambda *a, **k: None

import modulo_nomina

root = ctk.CTk()
root.geometry("1600x950")
frame = ctk.CTkFrame(root)
frame.pack(fill="both", expand=True)

p("### CONSTRUCCION DEL MODULO")
try:
    app = modulo_nomina.ModuloNominaApp(frame, "TEST_SMOKE")
    p("Constructor OK")
except Exception:
    p("FALLO EN EL CONSTRUCTOR:\n" + traceback.format_exc())
    open(r"C:\\Users\\Alberto\\Desktop\\Proyecto Control Flota 140926\_smoke_gui.txt", "w", encoding="utf-8").write(out.getvalue())
    sys.exit(1)

def bombear(segundos=1.0):
    fin = time.time() + segundos
    while time.time() < fin:
        root.update()
        time.sleep(0.02)

p("Esperando carga inicial en segundo plano...")
bombear(12)
p("empleados=%d turnos=%d horarios=%d parametros=%d" % (len(app.empleados), len(app.turnos), len(app.horarios), len(app.parametros)))
p("estado barra: %s" % app.lbl_estado.cget("text"))

p("\n### RECORRIDO DE PESTAÑAS PRINCIPALES")
for nombre in [" 📊 Panel ", " 📥 Marcaciones ", " 🗓️ Asistencia ", " 👥 Personal ", " 💵 Planilla ",
               " 📈 Reportes ", " ⚙️ Configuración "]:
    try:
        app.tabview.set(nombre)
        bombear(0.8)
        p("  OK %s" % nombre)
    except Exception:
        p("  FALLO %s -> %s" % (nombre, traceback.format_exc()))

p("\n### SUB-PESTAÑAS")
for valor in ["Importar Excel del reloj", "Personas del reloj", "Historial"]:
    try:
        app.cambiar_sub_marcaciones(valor); bombear(1.5); p("  OK marcaciones/%s" % valor)
    except Exception: p("  FALLO %s -> %s" % (valor, traceback.format_exc()))
for valor in ["Matriz mensual", "Detalle por día", "Incidencias", "Horas extra"]:
    try:
        app.cambiar_sub_asistencia(valor); bombear(2.5); p("  OK asistencia/%s" % valor)
    except Exception: p("  FALLO %s -> %s" % (valor, traceback.format_exc()))
for valor in ["Empleados", "Turnos", "Horarios", "Asignaciones"]:
    try:
        app.cambiar_sub_personal(valor); bombear(2.0); p("  OK personal/%s" % valor)
    except Exception: p("  FALLO %s -> %s" % (valor, traceback.format_exc()))
for valor in ["Resumen del período", "Boleta de pago", "Historial"]:
    try:
        app.cambiar_sub_planilla(valor); bombear(2.0); p("  OK planilla/%s" % valor)
    except Exception: p("  FALLO %s -> %s" % (valor, traceback.format_exc()))
for valor in ["Parámetros de cálculo", "Feriados y calendario", "Herramientas"]:
    try:
        app.cambiar_sub_config(valor); bombear(1.5); p("  OK config/%s" % valor)
    except Exception: p("  FALLO %s -> %s" % (valor, traceback.format_exc()))

p("\n### ANALISIS DEL EXCEL REAL DESDE LA INTERFAZ")
F = r"C:\Users\Alberto\Downloads\Transacciones_2026-04-01_2026-09-30.xlsx"
app.tabview.set(" 📥 Marcaciones ")
app.ent_archivo.delete(0, tk.END); app.ent_archivo.insert(0, F)
mock_calls = []
app.analizar_archivo()
bombear(10)
p("filas en vista previa: %d" % len(app.tabla_preview.get_children()))
p("etiqueta analisis: %s" % app.lbl_analisis.cget("text")[:300])
p("mapeo detectado en combos: %s" % app.mapeo_desde_combos())
p("filas personas reloj: %d" % len(app.tabla_personas.get_children()))
p("filas mapeos: %d" % len(app.tabla_mapeos.get_children()))

p("\n### MATRIZ DE ASISTENCIA Y BOLETA")
app.tabview.set(" 🗓️ Asistencia ")
app.cambiar_sub_asistencia("Matriz mensual"); bombear(4)
p("periodo=%s filas matriz=%d" % (app.periodo_actual, len(app.tabla_matriz.get_children())))
app.periodo_actual = "2026-09"
app.ent_periodo.delete(0, tk.END); app.ent_periodo.insert(0, "09/2026")
app.cargar_matriz_asistencia(); bombear(6)
p("2026-09 -> filas matriz=%d columnas=%d" % (len(app.tabla_matriz.get_children()), len(app.tabla_matriz.cget("columns"))))
app.cambiar_sub_asistencia("Detalle por día"); bombear(3)
p("detalle dia -> filas=%d" % len(app.tabla_detalle.get_children()))
app.cambiar_sub_asistencia("Incidencias"); bombear(2)
p("incidencias -> filas=%d" % len(app.tabla_incidencias.get_children()))
app.cambiar_sub_asistencia("Horas extra"); bombear(2)
p("horas extra -> filas=%d" % len(app.tabla_horas_extra.get_children()))

p("\n### PLANILLA")
app.tabview.set(" 💵 Planilla ")
app.ent_periodo_planilla.delete(0, tk.END); app.ent_periodo_planilla.insert(0, "09/2026")
app.cargar_planilla(); bombear(4)
p("planilla -> filas=%d | estado=%s" % (len(app.tabla_planilla.get_children()), app.lbl_estado_planilla.cget("text")))
p("totales: %s" % app.lbl_totales_planilla.cget("text")[:250])
app.cambiar_sub_planilla("Historial"); bombear(3)
p("historial planilla -> filas=%d" % len(app.tabla_historial_planilla.get_children()))
app.sub_planilla.set("Boleta de pago"); app.cambiar_sub_planilla("Boleta de pago")
if app.empleados:
    app.cmb_boleta_empleado.set(app.opciones_empleados()[0])
    app.mostrar_boleta(); bombear(3)
    p("boleta mostrada -> widgets=%d" % len(app.frame_boleta.winfo_children()))

p("\n### REPORTES")
app.tabview.set(" 📈 Reportes ")
for reporte in ["Puntualidad", "Tardanzas", "Faltas", "Horas extra", "Marcaciones del reloj"]:
    app.sub_reportes.set(reporte)
    app.ent_periodo_reporte.delete(0, tk.END); app.ent_periodo_reporte.insert(0, "09/2026")
    app.generar_reporte(); bombear(4)
    filas = len(app.tabla_reporte.get_children()) if app.tabla_reporte else -1
    p("  %-24s -> %d filas" % (reporte, filas))

p("\n### DIALOGOS DISPARADOS")
for tipo, titulo, mensaje in registro_dialogos:
    p("  [%s] %s :: %s" % (tipo, titulo, mensaje.replace("\n", " | ")[:300]))

p("\n### FIN")
try:
    root.destroy()
except Exception:
    pass
open(r"C:\\Users\\Alberto\\Desktop\\Proyecto Control Flota 140926\_smoke_gui.txt", "w", encoding="utf-8").write(out.getvalue())
print("SMOKE OK")
