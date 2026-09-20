# -*- coding: utf-8 -*-
"""Detecta controles cortados (abajo o a la derecha) en cada vista del modulo."""
import io, sys, time, warnings
warnings.filterwarnings("ignore")
sys.path.insert(0, r"C:\\Users\\Alberto\\Desktop\\Proyecto Control Flota 140926")
out = io.StringIO()
def p(*a): out.write(" ".join(str(x) for x in a) + "\n")
import tkinter as tk
from tkinter import messagebox
messagebox.showinfo = lambda *a, **k: True
messagebox.showerror = lambda *a, **k: True
messagebox.showwarning = lambda *a, **k: True
messagebox.askyesno = lambda *a, **k: True
import customtkinter as ctk
import modulo_nomina

ANCHO, ALTO = 1065, 568
root = ctk.CTk(); root.geometry("%dx%d+10+10" % (ANCHO, ALTO))
bl = ctk.CTkFrame(root, width=280, corner_radius=0); bl.pack(side="left", fill="y"); bl.pack_propagate(False)
central = ctk.CTkFrame(root, corner_radius=0); central.pack(side="right", fill="both", expand=True)

def bombear(s=0.8):
    fin = time.time() + s
    while time.time() < fin:
        root.update(); time.sleep(0.02)

app = modulo_nomina.ModuloNominaApp(central, "TEST_LAYOUT")
bombear(8)
p("### VENTANA %dx%d | area del modulo %dx%d" % (ANCHO, ALTO, central.winfo_width(),
                                                  central.winfo_height()))

def marco_pestana():
    return app.tabview.tab(app.tabview.get())

def recolectar(widget, lista):
    for hijo in widget.winfo_children():
        if hijo.__class__.__name__ in ("CTkButton", "CTkEntry", "CTkComboBox", "CTkCheckBox"):
            lista.append(hijo)
        recolectar(hijo, lista)

def vista_activa(widget, tab_frame):
    """Sube hasta el hijo directo del marco de la pestana y mira si esa vista esta puesta."""
    hijo, padre = widget, getattr(widget, "master", None)
    while padre is not None and padre is not tab_frame:
        hijo, padre = padre, getattr(padre, "master", None)
    if padre is None:
        return False
    try:
        return bool(hijo.winfo_ismapped())
    except Exception:
        return False

def en_scrollable(widget, tab_frame):
    padre = getattr(widget, "master", None)
    while padre is not None and padre is not tab_frame:
        if "Scrollable" in padre.__class__.__name__:
            return True
        padre = getattr(padre, "master", None)
    return False

def etiqueta(widget):
    try:
        texto = widget.cget("text")
        if texto:
            return str(texto)[:28]
    except Exception:
        pass
    return "(campo)"

def revisar(nombre):
    tab = marco_pestana()
    controles = []
    recolectar(tab, controles)
    abajo_max = tab.winfo_rooty() + tab.winfo_height()
    derecha_max = tab.winfo_rootx() + tab.winfo_width()
    cortados = []
    for control in controles:
        try:
            if en_scrollable(control, tab) or not vista_activa(control, tab):
                continue
            abajo = control.winfo_rooty() + control.winfo_height()
            derecha = control.winfo_rootx() + control.winfo_width()
            problemas = []
            if not control.winfo_ismapped():
                problemas.append("oculto")
            if abajo > abajo_max + 2:
                problemas.append("abajo +%dpx" % (abajo - abajo_max))
            if derecha > derecha_max + 2:
                problemas.append("derecha +%dpx" % (derecha - derecha_max))
            if problemas:
                cortados.append((etiqueta(control), " ".join(problemas)))
        except Exception:
            pass
    if cortados:
        p("  %-30s *** %d: %s" % (nombre, len(cortados),
                                  " | ".join("%s (%s)" % (t, e) for t, e in cortados[:6])))
    else:
        p("  %-30s OK" % nombre)
    return len(cortados)

p("")
p("### PESTANAS")
fallos = 0
for nombre in [" 📊 Panel ", " 📥 Marcaciones ", " 🗓️ Asistencia ", " 👥 Personal ", " 💵 Planilla ",
               " 📈 Reportes ", " ⚙️ Configuración "]:
    app.tabview.set(nombre); bombear(0.7); fallos += revisar(nombre.strip())

p("")
p("### VISTAS INTERNAS")
vistas = [("Importar Excel del reloj", "marcaciones"), ("Personas del reloj", "marcaciones"),
          ("Historial", "marcaciones"), ("Matriz mensual", "asistencia"),
          ("Detalle por día", "asistencia"), ("Incidencias", "asistencia"),
          ("Horas extra", "asistencia"), ("Empleados", "personal"), ("Turnos", "personal"),
          ("Horarios", "personal"), ("Asignaciones", "personal"),
          ("Resumen del período", "planilla"), ("Boleta de pago", "planilla"),
          ("Historial", "planilla"), ("Parámetros de cálculo", "config"),
          ("Feriados y calendario", "config"), ("Herramientas", "config")]
cambiadores = {"marcaciones": app.cambiar_sub_marcaciones, "asistencia": app.cambiar_sub_asistencia,
               "personal": app.cambiar_sub_personal, "planilla": app.cambiar_sub_planilla,
               "config": app.cambiar_sub_config}
for nombre, grupo in vistas:
    try:
        cambiadores[grupo](nombre); bombear(0.8); fallos += revisar(nombre)
    except Exception as e:
        p("  %-30s ERROR %s" % (nombre, e)); fallos += 1
p("")
p("### VISTAS CON CONTROLES CORTADOS: %d" % fallos)

try:
    root.destroy()
except Exception:
    pass
open(r"C:\\Users\\Alberto\\Desktop\\Proyecto Control Flota 140926\_layout.txt", "w", encoding="utf-8").write(out.getvalue())
print("OK")
