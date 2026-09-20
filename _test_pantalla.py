# -*- coding: utf-8 -*-
import io, sys, time, traceback, warnings
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

root = ctk.CTk()
root.geometry("1400x900+40+40")
frame = ctk.CTkFrame(root)
frame.pack(fill="both", expand=True)

def bombear(segundos=0.6):
    fin = time.time() + segundos
    while time.time() < fin:
        root.update()
        time.sleep(0.02)

app = modulo_nomina.ModuloNominaApp(frame, "TEST_PANTALLA")
bombear(1.0)

p("### ESTADO INICIAL")
p("  boton existe:", hasattr(app, "btn_pantalla"), "| texto:", app.btn_pantalla.cget("text"))
p("  pantalla completa activa:", app.pantalla_completa_activa())
p("  atributo -fullscreen en la ventana:", root.attributes("-fullscreen"))
p("  geometria:", root.geometry())

p("\n### ACTIVAR CON EL BOTON")
app.btn_pantalla.invoke()
bombear(0.8)
p("  activa:", app.pantalla_completa_activa())
p("  texto del boton:", app.btn_pantalla.cget("text"))
p("  color del boton:", app.btn_pantalla.cget("fg_color"))
p("  atributo -fullscreen:", root.attributes("-fullscreen"))

p("\n### DESACTIVAR CON EL BOTON (debe volver a ventana normal)")
app.btn_pantalla.invoke()
bombear(0.8)
p("  activa:", app.pantalla_completa_activa())
p("  texto del boton:", app.btn_pantalla.cget("text"))
p("  atributo -fullscreen:", root.attributes("-fullscreen"))
p("  geometria restaurada:", root.geometry())

p("\n### ATAJOS DE TECLADO")
root.event_generate("<F11>"); bombear(0.7)
p("  tras F11 -> activa:", app.pantalla_completa_activa(), "| texto:", app.btn_pantalla.cget("text"))
root.event_generate("<Escape>"); bombear(0.7)
p("  tras Esc -> activa:", app.pantalla_completa_activa(), "| texto:", app.btn_pantalla.cget("text"))
root.event_generate("<F11>"); bombear(0.5)
p("  F11 otra vez -> activa:", app.pantalla_completa_activa())
app.salir_pantalla_completa(); bombear(0.5)
p("  metodo salir -> activa:", app.pantalla_completa_activa())

p("\n### HELPER DE LA BARRA LATERAL (control_general)")
try:
    import control_general as cg
    p("  import OK | helper:", callable(getattr(cg, "alternar_pantalla_completa", None)))
    ventana = ctk.CTkToplevel(root)
    ventana.geometry("800x600+60+60")
    boton = ctk.CTkButton(ventana, text="⛶ Pantalla completa")
    boton.pack()
    bombear(0.4)
    p("  estado 1:", cg.alternar_pantalla_completa(ventana, boton), "|", boton.cget("text"))
    p("  -fullscreen:", ventana.attributes("-fullscreen"))
    p("  estado 2:", cg.alternar_pantalla_completa(ventana, boton), "|", boton.cget("text"))
    p("  -fullscreen:", ventana.attributes("-fullscreen"))
    ventana.destroy()
    bombear(0.3)
except Exception:
    p("  FALLO:\n" + traceback.format_exc())

p("\n### ESTADO COMPARTIDO ENTRE MODULO Y BARRA LATERAL")
cg.alternar_pantalla_completa(root, None)
bombear(0.5)
p("  activado desde la barra lateral ->", root._pantalla_completa_activa)
app.actualizar_boton_pantalla()
p("  el boton del modulo se sincroniza:", app.btn_pantalla.cget("text"))
app.btn_pantalla.invoke(); bombear(0.6)
p("  desactivado desde el modulo ->", getattr(root, "_pantalla_completa_activa", None))

try:
    root.destroy()
except Exception:
    pass
open(r"C:\\Users\\Alberto\\Desktop\\Proyecto Control Flota 140926\_test_pantalla.txt", "w", encoding="utf-8").write(out.getvalue())
print("TEST PANTALLA OK")
