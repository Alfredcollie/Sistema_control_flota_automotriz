# -*- coding: utf-8 -*-
import io, sys, time, os, warnings
warnings.filterwarnings("ignore")
sys.path.insert(0, r"C:\\Users\\Alberto\\Desktop\\Proyecto Control Flota 140926")
out = io.StringIO()
def p(*a): out.write(" ".join(str(x) for x in a) + "\n")
import tkinter as tk
from tkinter import messagebox
for f in ("showinfo", "showerror", "showwarning"):
    setattr(messagebox, f, lambda *a, **k: True)
messagebox.askyesno = lambda *a, **k: True
import customtkinter as ctk
import modulo_nomina

p("### UBICACION DEL INSTRUCTIVO")
ruta = modulo_nomina.ruta_instructivo()
p("  encontrado: %s" % (ruta or "*** NO ENCONTRADO ***"))
if ruta:
    p("  tamano: %.1f KB" % (os.path.getsize(ruta) / 1024.0))

root = ctk.CTk(); root.geometry("1200x740+10+10")
frame = ctk.CTkFrame(root); frame.pack(fill="both", expand=True)
def bombear(s):
    fin = time.time() + s
    while time.time() < fin:
        root.update(); time.sleep(0.05)
app = modulo_nomina.ModuloNominaApp(frame, "TEST_AYUDA")
bombear(14)
p("")
p("### BOTON DE AYUDA")
p("  texto del boton: %r | visible: %s" % (app.btn_ayuda.cget("text"), bool(app.btn_ayuda.winfo_ismapped())))

p("")
p("### ABRIR LA AYUDA (visor interno del PDF)")
t0 = time.time()
app.btn_ayuda.invoke()
bombear(6)
visor = getattr(app, "_visor", None)
p("  visor creado: %s (%.1f s)" % (visor is not None and visor.winfo_exists(), time.time() - t0))
if visor is None or not visor.winfo_exists():
    p("  *** No se abrio el visor ***"); root.destroy(); sys.exit(1)
p("  titulo: %s" % visor.title())
p("  paginas del documento: %d" % visor.total)
p("  estado: %s" % visor.lbl_estado.cget("text"))
p("  zoom: %s" % visor.lbl_zoom.cget("text"))
p("  imagen dibujada: %s (%sx%s px)" % (visor._imagen is not None,
                                        visor._imagen.cget("size")[0], visor._imagen.cget("size")[1]))

p("")
p("### NAVEGACION")
visor.mostrar_pagina(4); bombear(1.2)
p("  ir a la 5      -> %s | campo=%s" % (visor.lbl_estado.cget("text"), visor.ent_pagina.get()))
visor.mostrar_pagina(visor.numero + 1); bombear(1.0)
p("  siguiente      -> %s" % visor.lbl_estado.cget("text"))
visor.mostrar_pagina(visor.numero - 1); bombear(1.0)
p("  anterior       -> %s" % visor.lbl_estado.cget("text"))
visor.ent_pagina.delete(0, tk.END); visor.ent_pagina.insert(0, "12"); visor.ir_a_pagina(); bombear(1.2)
p("  escribir 12    -> %s" % visor.lbl_estado.cget("text"))
visor.mostrar_pagina(999); bombear(1.2)
p("  pedir la 999   -> %s (debe quedarse en la ultima)" % visor.lbl_estado.cget("text"))
visor.mostrar_pagina(-5); bombear(1.0)
p("  pedir la -5    -> %s (debe quedarse en la primera)" % visor.lbl_estado.cget("text"))

p("")
p("### ZOOM")
visor.cambiar_zoom(0.3); bombear(1.2)
p("  acercar        -> %s" % visor.lbl_zoom.cget("text"))
visor.cambiar_zoom(-0.6); bombear(1.2)
p("  alejar         -> %s" % visor.lbl_zoom.cget("text"))
visor.ajustar_ancho(); bombear(1.2)
p("  ajustar ancho  -> %s" % visor.lbl_zoom.cget("text"))

p("")
p("### CERRAR Y REABRIR")
visor.destroy(); bombear(1.5)
p("  cerrado: %s" % (not getattr(app, "_visor").winfo_exists()))
app.abrir_ayuda(); bombear(5)
p("  reabierto: %s | paginas: %d" % (app._visor.winfo_exists(), app._visor.total))
app._visor.destroy(); bombear(1)
p("")
p("### SIN EL ARCHIVO (comportamiento de aviso)")
original = modulo_nomina.ruta_instructivo
modulo_nomina.ruta_instructivo = lambda: ""
app.abrir_ayuda(); bombear(1)
p("  aviso mostrado sin cortar la aplicacion: OK")
modulo_nomina.ruta_instructivo = original
try:
    root.destroy()
except Exception:
    pass
open(r"C:\\Users\\Alberto\\Desktop\\Proyecto Control Flota 140926\_test_ayuda.txt", "w", encoding="utf-8").write(out.getvalue())
print("OK")
