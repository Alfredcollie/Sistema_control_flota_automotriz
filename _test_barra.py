# -*- coding: utf-8 -*-
import io, sys, time, types, traceback, warnings
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
import control_general as cg

root = ctk.CTk()
root.geometry("1400x900+40+40")

class Falso:
    pass

f = Falso()
f.root = root
# Estructura real de construir_dashboard_spa
f.sidebar = ctk.CTkFrame(root, width=280, corner_radius=0, fg_color="#1a252c")
f.sidebar.pack(side="left", fill="y")
f.sidebar.pack_propagate(False)
f.contenedor_central = ctk.CTkFrame(root, corner_radius=0, fg_color="transparent")
f.contenedor_central.pack(side="right", fill="both", expand=True)
f.btn_flotante_salir = ctk.CTkButton(root, text="Salir de pantalla completa", width=215, height=28)
f.alternar_barra_lateral = types.MethodType(cg.ControlGeneralEventos.alternar_barra_lateral, f)
root._alternar_barra_lateral = f.alternar_barra_lateral
root._boton_flotante_salir = f.btn_flotante_salir
root._pantalla_completa_activa = False

def bombear(s=0.8):
    fin = time.time() + s
    while time.time() < fin:
        root.update(); time.sleep(0.02)

bombear(1.0)
p("### ESTADO INICIAL (ventana normal)")
p("  barra lateral visible: %s | ancho=%d" % (f.sidebar.winfo_ismapped(), f.sidebar.winfo_width()))
p("  contenido ancho=%d | boton flotante visible=%s" % (f.contenedor_central.winfo_width(),
                                                        f.btn_flotante_salir.winfo_ismapped()))

p("\n### ENTRAR A PANTALLA COMPLETA (con el helper del sistema)")
estado = cg.alternar_pantalla_completa(root, None)
bombear(1.0)
p("  -fullscreen: %s" % root.attributes("-fullscreen"))
p("  barra lateral visible: %s  <- debe ser False" % f.sidebar.winfo_ismapped())
p("  contenido ancho=%d (ventana=%d)  <- debe ocupar todo" % (f.contenedor_central.winfo_width(),
                                                              root.winfo_width()))
p("  boton flotante visible: %s  <- debe ser True" % f.btn_flotante_salir.winfo_ismapped())
p("  texto del boton flotante: %s" % f.btn_flotante_salir.cget("text"))

p("\n### SALIR DE PANTALLA COMPLETA")
estado = cg.alternar_pantalla_completa(root, None)
bombear(1.2)
p("  -fullscreen: %s" % root.attributes("-fullscreen"))
p("  barra lateral visible: %s | ancho=%d  <- debe recuperar 280" % (f.sidebar.winfo_ismapped(),
                                                                     f.sidebar.winfo_width()))
p("  contenido ancho=%d  <- debe reducirse" % f.contenedor_central.winfo_width())
p("  boton flotante visible: %s  <- debe ser False" % f.btn_flotante_salir.winfo_ismapped())

p("\n### ENTRAR Y SALIR 3 VECES (comprobar que no se descuadra)")
for vuelta in range(3):
    cg.alternar_pantalla_completa(root, None); bombear(0.5)
    dentro = (f.sidebar.winfo_ismapped(), f.btn_flotante_salir.winfo_ismapped())
    cg.alternar_pantalla_completa(root, None); bombear(0.6)
    fuera = (f.sidebar.winfo_ismapped(), f.btn_flotante_salir.winfo_ismapped())
    p("  vuelta %d -> dentro barra=%s flotante=%s | fuera barra=%s flotante=%s | ancho barra=%d" % (
        vuelta + 1, dentro[0], dentro[1], fuera[0], fuera[1], f.sidebar.winfo_width()))

p("\n### EL MODULO DE NOMINA SIN GANCHO (no debe fallar)")
try:
    import modulo_nomina
    ventana = ctk.CTkToplevel(root)
    ventana.geometry("900x600+80+80")
    marco = ctk.CTkFrame(ventana)
    marco.pack(fill="both", expand=True)
    root._alternar_barra_lateral = None   # simulamos que el sistema no expone el gancho
    app = modulo_nomina.ModuloNominaApp(marco, "TEST")
    bombear(2)
    app.alternar_pantalla_completa(); bombear(0.6)
    p("  alternar sin gancho -> activa=%s | texto=%s" % (app.pantalla_completa_activa(),
                                                         app.btn_pantalla.cget("text")))
    app.alternar_pantalla_completa(); bombear(0.6)
    p("  vuelta a la normal -> activa=%s" % app.pantalla_completa_activa())
    ventana.destroy(); bombear(0.5)
except Exception:
    p("  FALLO:\n" + traceback.format_exc())

try:
    root.destroy()
except Exception:
    pass
open(r"C:\\Users\\Alberto\\Desktop\\Proyecto Control Flota 140926\_test_barra.txt", "w", encoding="utf-8").write(out.getvalue())
print("OK")
