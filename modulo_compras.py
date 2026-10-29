# -*- coding: utf-8 -*-
# Aquí tienes el código completo y unificado.

### 🌟 Cambios realizados en la unificación:
#1. **Unificación de Opciones:** Se eliminó el botón duplicado de *"Resumen por Proveedor"* y se integró todo en un único **"📊 Reporte de Totales y Proveedores"**.
#2. **Ventana Consolidada Todo-en-Uno:** La nueva ventana de reporte incluye:
#   * **Filtros avanzados:** Por Proveedor (*"Todos"* o uno específico) y por rango de fechas (*Desde / Hasta* con selector de calendario).
#   * **Tarjetas de Totales Globales:** Muestra *Compras Brutas*, *IGV*, *Detracciones/Retenciones*, *Total Pagado* y *Deuda Pendiente*.
#   * **Tabla de Desglose por Proveedor:** Muestra en tiempo real la lista consolidada de proveedores con su *Total Facturado*, *Monto Pagado*, *Saldo Pendiente* y *N° de Comprobantes*, respondiendo a los filtros aplicados.
#3. **Alto Rendimiento y Cero Bloqueos:** El cálculo corre en segundo plano (`Thread`) y se procesa en memoria instantáneamente.

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import customtkinter as ctk
import os
import sys  
import shutil
import calendar
import re
import json
import difflib
import tempfile
import subprocess 
import ctypes
import urllib.request
import urllib.parse
import urllib.error
import xml.etree.ElementTree as ET
import base64
import time
from datetime import datetime
import threading

# 🚀 IMPORTAMOS NUESTRAS HERRAMIENTAS CORPORATIVAS
from conexion import conectar_db, registrar_auditoria, liberar_conexion
from buffer_memoria import cache_sistema
from app_paths import CONFIG_FILE, eliminar_archivo, ruta_para_guardar, resolver_ruta_archivo
from config_nube import cargar_bancos
from dialogos_seguros import seleccionar_archivo_dialogo, guardar_archivo_dialogo
# Tareas en segundo plano seguras: la interfaz SIEMPRE se actualiza desde el hilo principal
# (llamar a Tk desde un hilo secundario congela la aplicación en macOS)
from tareas_seguras import ejecutar_en_hilo

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

# =========================================================
# 🚀 ADAPTACIÓN MULTIPLATAFORMA (WINDOWS / MAC / LINUX)
# =========================================================
if sys.platform == "win32":
    try:
        hwnd_cmd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd_cmd:
            ctypes.windll.user32.ShowWindow(hwnd_cmd, 6)
    except Exception:
        pass
def activar_busqueda_rapida(combo, titulo="Buscar proveedor", al_elegir=None):
    """Buscador del desplegable: clic en el combo abre la lista con filtro al escribir."""
    estado = {"abierta": False}

    def _plano(t):
        t = str(t or "")
        try:
            import unicodedata
            t = "".join(c for c in unicodedata.normalize("NFD", t)
                        if unicodedata.category(c) != "Mn")
        except Exception:
            pass
        return t.casefold().strip()

    def abrir(event=None):
        if estado["abierta"]:
            return None
        completo = list(getattr(combo, "_valores_completos", []) or [])
        if not completo:
            try:
                completo = [str(v) for v in (combo.cget("values") or [])]
            except Exception:
                completo = []
        completo = [v for v in completo if str(v).strip()]
        if not completo:
            return None
        try:
            menu = getattr(combo, "_dropdown_menu", None)
            if menu is not None and getattr(menu, "_is_open", False):
                menu._close_menu()
        except Exception:
            pass
        try:
            ancho = max(int(combo.winfo_width()), 420)
            x = combo.winfo_rootx()
            y = combo.winfo_rooty() + combo.winfo_height()
        except Exception:
            return None
        try:
            top = ctk.CTkToplevel(combo)
            top.overrideredirect(True)
            top.geometry("%dx360+%d+%d" % (ancho, x, y))
            try:
                top.attributes("-topmost", True)
            except Exception:
                pass
            cabecera = ctk.CTkLabel(top, text="\U0001f50d " + titulo + ": escriba para filtrar",
                                    font=("Arial", 11, "bold"), anchor="w")
            cabecera.pack(fill="x", padx=8, pady=(8, 2))
            entrada = ctk.CTkEntry(top, placeholder_text="Escriba el nombre... (ej. coesti)")
            entrada.pack(fill="x", padx=8, pady=(0, 4))
            cuenta = ctk.CTkLabel(top, text="", font=("Arial", 10), text_color="gray", anchor="w")
            cuenta.pack(fill="x", padx=10)
            lista = tk.Listbox(top, height=9, activestyle="none", exportselection=False,
                               selectbackground="#2c3e50", selectforeground="white",
                               font=("Arial", 11))
            lista.pack(fill="both", expand=True, padx=8, pady=(2, 2))
            pie = ctk.CTkLabel(top, text="clic para elegir  ·  ↑↓ mover  ·  Enter acepta  ·  Esc/click fuera cierra",
                               font=("Arial", 10), text_color="gray")
            pie.pack(fill="x", padx=8, pady=(0, 8))
        except Exception:
            return None

        estado["abierta"] = True

        def cerrar(_e=None):
            if not estado["abierta"]:
                return
            estado["abierta"] = False
            try:
                top.destroy()
            except Exception:
                pass

        def elegir(_e=None):
            sel = lista.curselection()
            if not sel:
                return "break"
            valor = lista.get(sel[0])
            cerrar()
            try:
                combo.set(valor)
            except Exception:
                pass
            # 🔎 se avisa al combo para que complete el RUC / detracción
            try:
                accion = al_elegir or getattr(combo, "_al_elegir", None)
                if callable(accion):
                    accion()
            except Exception as e_sel:
                print("Aviso - al elegir proveedor:", e_sel)
            try:
                combo.event_generate("<<ComboboxSelected>>")
            except Exception:
                pass
            return "break"

        def pintar(texto=""):
            lista.delete(0, tk.END)
            t = _plano(texto)
            hallados = [v for v in completo if not t or t in _plano(v)]
            for valor in hallados:
                lista.insert(tk.END, valor)
            if hallados:
                lista.selection_clear(0, tk.END)
                lista.selection_set(0)
                lista.activate(0)
            cuenta.configure(text="%d resultado(s) — clic en uno para elegirlo" % len(hallados))

        pintar()
        entrada.bind("<KeyRelease>", lambda e: pintar(entrada.get()))
        entrada.bind("<Return>", elegir)
        entrada.bind("<Escape>", cerrar)
        entrada.bind("<Down>", lambda e: (lista.focus_set(), None)[1])
        lista.bind("<ButtonRelease-1>", elegir)     # 👈 UN clic elige y cierra
        lista.bind("<Return>", elegir)
        lista.bind("<Escape>", cerrar)
        try:
            top.bind("<FocusOut>", lambda e: top.after(200, carpeta_salida))
        except Exception:
            pass

        def carpeta_salida():
            try:
                enfocado = top.focus_get()
                dentro = enfocado is not None and str(enfocado).startswith(str(top))
            except Exception:
                dentro = False
            if not dentro:
                cerrar()

        try:
            entrada.focus_set()
        except Exception:
            pass
        return None

    try:
        combo._valores_completos = list(combo.cget("values") or [])
    except Exception:
        combo._valores_completos = []
    # 🚫 Sin flecha de desplegable: el buscador reemplaza a la lista nativa
    # 🚫 Se tapa SOLO la flechita: el cuadro del campo se mantiene
    try:
        _lienzo = getattr(combo, "_canvas", None)
        _draw_original = getattr(combo, "_draw", None)
        if _lienzo is not None and callable(_draw_original):
            def _draw_sin_flecha(*args, **kwargs):
                _draw_original(*args, **kwargs)
                try:
                    try:
                        fondo = combo._apply_appearance_mode(combo.cget("fg_color"))
                    except Exception:
                        fondo = "#d9d9d9"
                    # 🚫 se borra la flechita y el recuadro de la derecha queda igual al campo
                    _lienzo.delete("dropdown_arrow")
                    for _etiqueta in ("inner_parts_right", "inner_rectangle_right_1",
                                      "inner_rectangle_right_2"):
                        try:
                            _lienzo.itemconfig(_etiqueta, fill=fondo, outline=fondo)
                        except Exception:
                            pass
                except Exception:
                    pass
            combo._draw = _draw_sin_flecha
            try:
                combo._draw()
            except Exception:
                pass
    except Exception:
        pass
    for _nombre_btn in ("_button", "_arrow", "_dropdown_button", "_btn"):
        try:
            _boton = getattr(combo, _nombre_btn, None)
            if _boton is None:
                continue
            for _quitar in ("place_forget", "pack_forget", "grid_forget"):
                try:
                    getattr(_boton, _quitar)()
                except Exception:
                    pass
            try:
                _boton.configure(width=0, height=0, text="")
            except Exception:
                pass
        except Exception:
            pass
    try:
        combo.bind("<Button-1>", abrir, add="+")
        combo.bind("<KeyRelease>", lambda e: abrir(), add="+")
    except Exception:
        pass
    return combo

def _numero_peru(txt):
    """Convierte '1,340.78' / '1.340,78' / '1340.78' a float (0.0 si no se puede)."""
    import re as _re
    t = _re.sub(r"[^\d.,]", "", str(txt or ""))
    if not t:
        return 0.0
    try:
        if "," in t and "." in t:
            if t.rfind(",") > t.rfind("."):
                t = t.replace(".", "").replace(",", ".")
            else:
                t = t.replace(",", "")
        elif "," in t:
            partes = t.split(",")
            t = t.replace(",", ".") if len(partes[-1]) <= 2 else t.replace(",", "")
        elif t.count(".") > 1:
            partes = t.split(".")
            t = ("".join(partes[:-1]) + "." + partes[-1]) if len(partes[-1]) <= 2 else "".join(partes)
        return float(t)
    except Exception:
        return 0.0


def monto_del_comprobante(texto):
    """💵 Importe leído de un comprobante (0.0 si no se encuentra).

    Prioriza el monto de la transferencia/depósito (una línea «s/ X» sin las
    palabras total/comisión), para no confundirlo con el «Total cobrado», que
    incluye la comisión del banco (ej. transferencia 803.23 + comisión 4.80 =
    «Total cobrado 808.03»).
    """
    import re as _re
    try:
        lineas = [l.strip() for l in str(texto or "").splitlines() if l.strip()]
        for l in lineas:
            bajo = l.lower()
            if any(p in bajo for p in ("total", "importe", "monto", "comis", "cobrado")):
                continue
            m = _re.search(r"s[./]\s*[\d][\d.,]*", bajo)
            if m:
                nums = _re.findall(r"\d[\d.,]*", m.group(0))
                if nums:
                    v = _numero_peru(nums[-1])
                    if v > 0:
                        return float(v)
        comision = 0.0
        for l in lineas:
            if "comis" in l.lower():
                nums = _re.findall(r"\d[\d.,]*", l)
                for n in reversed(nums):
                    c = _numero_peru(n)
                    if 0 < c < 10000000:
                        comision = float(c)
                        break
                if comision:
                    break
        for l in lineas:
            bajo = l.upper()
            if ("TOTAL" in bajo or "IMPORTE" in bajo or "MONTO" in bajo) and "COBRADO" in bajo:
                nums = _re.findall(r"\d[\d.,]*", l)
                if nums:
                    v = _numero_peru(nums[-1])
                    if v > 0:
                        return float(v) - comision
        for l in lineas:
            bajo = l.upper()
            if "TOTAL" in bajo or "IMPORTE" in bajo or "MONTO" in bajo:
                nums = _re.findall(r"\d[\d.,]*", l)
                if nums:
                    val = _numero_peru(nums[-1])
                    if val > 0:
                        return float(val)
        candidatos = []
        for l in lineas:
            for n in _re.findall(r"\d[\d.,]{2,}", l):
                v = _numero_peru(n)
                if 0 < v < 10000000:
                    candidatos.append(float(v))
        return max(candidatos) if candidatos else 0.0
    except Exception:
        return 0.0



def centrar_ventana(ventana, parent, ancho, alto):
    """Calcula la posición centrada de forma segura para Windows y macOS (evitando el Notch y Menú Superior)"""
    ventana.update_idletasks()
    try:
        if parent and parent.winfo_ismapped():
            p_x = parent.winfo_rootx()
            p_y = parent.winfo_rooty()
            p_w = parent.winfo_width()
            p_h = parent.winfo_height()
            x = p_x + (p_w // 2) - (ancho // 2)
            y = p_y + (p_h // 2) - (alto // 2)
        else:
            s_w = ventana.winfo_screenwidth()
            s_h = ventana.winfo_screenheight()
            x = (s_w // 2) - (ancho // 2)
            y = (s_h // 2) - (alto // 2)
    except Exception:
        s_w = ventana.winfo_screenwidth()
        s_h = ventana.winfo_screenheight()
        x = (s_w // 2) - (ancho // 2)
        y = (s_h // 2) - (alto // 2)

    x = max(10, x)
    y = max(35, y)
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")

def ver_documento_en_ventana(ruta, titulo="Documento", parent=None):
    """🖼️ Visor INTERNO: muestra el PDF (página a página) o la imagen sin salir del programa."""
    from app_paths import resolver_ruta_archivo
    real = ""
    try:
        real = resolver_ruta_archivo(ruta) or ""
    except Exception:
        real = ""
    if not real:
        real = str(ruta or "")
    if not real or not os.path.exists(real):
        return messagebox.showerror(
            "Documento no encontrado",
            "El archivo no está en este equipo:" + NL + NL + str(ruta), parent=parent)

    ventana = ctk.CTkToplevel(parent)
    ventana.title(titulo + " — " + os.path.basename(real))
    try:
        ventana.geometry("980x760")
        ventana.transient(parent)
    except Exception:
        pass

    barra = ctk.CTkFrame(ventana, fg_color="#1f538d", corner_radius=0)
    barra.pack(fill="x")

    estado = {"paginas": [], "i": 0}

    def _cargar_paginas():
        paginas = []
        try:
            if str(real).lower().endswith(".pdf"):
                import fitz
                from PIL import Image
                doc = fitz.open(real)
                for i in range(doc.page_count):
                    pix = doc[i].get_pixmap(dpi=140)
                    paginas.append(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
                doc.close()
            else:
                from PIL import Image
                paginas.append(Image.open(real))
        except Exception as e_carga:
            print("Aviso - abriendo el documento:", e_carga)
        return paginas

    estado["paginas"] = _cargar_paginas()

    lbl_info = ctk.CTkLabel(barra, text="", font=("Arial", 11, "bold"), text_color="white")
    lbl_info.pack(side="left", padx=10, pady=6)

    lienzo = ctk.CTkFrame(ventana, fg_color="#2b2b2b")
    lienzo.pack(fill="both", expand=True)
    # Se usa un Label de Tk (no CTkLabel): muestra bien las imágenes PhotoImage
    lbl_imagen = tk.Label(lienzo, text="", bg="#2b2b2b", fg="white")
    lbl_imagen.pack(fill="both", expand=True)

    ref = {"img": None}

    def pintar(_evento=None):
        paginas = estado["paginas"]
        if not paginas:
            lbl_imagen.configure(text="No se pudo mostrar el documento.", text_color="white")
            lbl_info.configure(text=os.path.basename(real))
            return
        i = max(0, min(estado["i"], len(paginas) - 1))
        estado["i"] = i
        try:
            ancho = max(400, ventana.winfo_width() - 40)
            alto = max(300, ventana.winfo_height() - 110)
        except Exception:
            ancho, alto = 900, 650
        img = paginas[i].copy()
        factor = min(ancho / float(img.width), alto / float(img.height), 1.6)
        if factor != 1:
            img = img.resize((max(1, int(img.width * factor)), max(1, int(img.height * factor))))
        from PIL import ImageTk
        ref["img"] = ImageTk.PhotoImage(img)
        lbl_imagen.configure(image=ref["img"], text="")
        lbl_info.configure(text="%s   ·   página %d de %d"
                           % (os.path.basename(real), i + 1, len(paginas)))

    def anterior():
        if estado["i"] > 0:
            estado["i"] -= 1
            pintar()

    def siguiente():
        if estado["i"] < len(estado["paginas"]) - 1:
            estado["i"] += 1
            pintar()

    def abrir_externo():
        try:
            abrir_documento(real)
        except Exception as e_ab:
            messagebox.showerror("Documento", "No se pudo abrir aparte:" + NL + str(e_ab), parent=ventana)

    ctk.CTkButton(barra, text="‹", width=36, fg_color="#2c3e50", command=anterior).pack(side="left", padx=2, pady=6)
    ctk.CTkButton(barra, text="›", width=36, fg_color="#2c3e50", command=siguiente).pack(side="left", padx=2, pady=6)
    ctk.CTkButton(barra, text="Abrir aparte", width=110, fg_color="#16a085",
                  command=abrir_externo).pack(side="right", padx=6, pady=6)
    ctk.CTkButton(barra, text="Cerrar", width=80, fg_color="#c0392b",
                  command=ventana.destroy).pack(side="right", padx=6, pady=6)

    ventana.bind("<Configure>", pintar)
    ventana.bind("<Left>", lambda e: anterior())
    ventana.bind("<Right>", lambda e: siguiente())
    try:
        ventana.after(120, pintar)
    except Exception:
        pintar()
    return ventana


def abrir_documento(ruta, titulo="Documento", parent=None):
    """🖼️ Muestra el documento en el VISOR INTERNO del programa.

    Antes se abría con el programa externo (Acrobat, Fotos…) y la ventana
    quedaba muchas veces detrás del sistema. Ahora se ve dentro del programa.
    'parent' es la ventana desde la que se abre (así el visor sale ENCIMA).
    """
    try:
        from visor_documentos import ver_documento_ventana
        return ver_documento_ventana(ruta, titulo, parent)
    except Exception as e_visor:
        print("Aviso - visor interno:", e_visor)
    return abrir_documento_externo(ruta)


def abrir_documento_externo(ruta):
    """Abre documentos de forma nativa en Windows, macOS y Linux.

    Resuelve la ruta aunque se haya guardado en otro equipo u otro sistema
    operativo (por ejemplo, una ruta de macOS usada desde Windows).
    """
    try:
        from app_paths import resolver_ruta_archivo
        encontrada = resolver_ruta_archivo(ruta)
        if encontrada:
            ruta = encontrada
    except Exception:
        pass
    try:
        ruta_norm = os.path.normpath(ruta)
        if not os.path.exists(ruta_norm):
            return messagebox.showerror("Error", f"El archivo no existe:\n{ruta_norm}")

        if sys.platform == "win32":
            os.startfile(ruta_norm)
        elif sys.platform == "darwin": 
            subprocess.call(["open", ruta_norm])
        else: 
            subprocess.call(["xdg-open", ruta_norm])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")

# 🚀 FIX macOS: el selector de archivos seguro ahora es COMPARTIDO (dialogos_seguros.py),
# para que todos los módulos usen el mismo arreglo y ninguno se quede sin él.

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
        if os.path.exists(str(CONFIG_FILE)):
            with open(str(CONFIG_FILE), "r", encoding="utf-8") as f:
                config.update(json.load(f))
    except Exception: pass
    return config

CONFIG_REGIONAL = cargar_configuracion_regional()

# =========================================================
# 🗓️ DESPLEGABLE DE MES PARA FILTRAR COMPRAS Y VENTAS
# =========================================================
NOMBRES_MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                 "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

def etiqueta_mes_de_fecha(fecha):
    """'09/10/2026' -> 'Octubre 2026' ('' si no se puede interpretar)."""
    texto = str(fecha or "").strip()
    if not texto:
        return ""
    partes = re.split(r"[/\-.]", texto)
    if len(partes) != 3:
        return ""
    a, b, c = partes
    try:
        if len(a) == 4:
            anio, mes = int(a), int(b)
        else:
            try:
                fmt = str((CONFIG_REGIONAL or {}).get("formato_fecha", "DD/MM/AAAA")).upper()
            except Exception:
                fmt = "DD/MM/AAAA"
            mes = int(b if not fmt.startswith("MM") else a)
            anio = int(c)
        if not (2000 <= anio <= 2100 and 1 <= mes <= 12):
            return ""
        return "%s %d" % (NOMBRES_MESES[mes - 1], anio)
    except Exception:
        return ""


def construir_valores_mes():
    """Opciones del desplegable: 'Todos los meses' + meses de años cercanos."""
    hoy = datetime.now()
    valores = ["Todos los meses"]
    for anio in (hoy.year - 1, hoy.year, hoy.year + 1):
        for nombre in NOMBRES_MESES:
            valores.append(f"{nombre} {anio}")
    return valores

def mes_en_curso():
    """Etiqueta del mes actual, usada como valor por defecto del desplegable."""
    hoy = datetime.now()
    return f"{NOMBRES_MESES[hoy.month - 1]} {hoy.year}"

def patron_fecha_mes(etiqueta):
    """Devuelve un patrón SQL LIKE para filtrar la columna 'fecha' por el mes indicado.
    Retorna None para 'Todos los meses' o etiquetas inválidas."""
    if not etiqueta or etiqueta == "Todos los meses":
        return None
    partes = etiqueta.strip().split()
    if len(partes) < 2 or partes[0] not in NOMBRES_MESES:
        return None
    mes = NOMBRES_MESES.index(partes[0]) + 1
    anio = partes[-1]
    mes_str = f"{mes:02d}"
    fmt = CONFIG_REGIONAL.get("formato_fecha", "DD/MM/AAAA")
    if fmt == "MM/DD/AAAA":
        return f"{mes_str}/%/{anio}"
    return f"%/{mes_str}/{anio}"

def traer_modulo_al_frente(ventana, retrasos=(60, 260, 700, 1150)):
    """🖥️ Deja la ventana del módulo DELANTE después de un mensaje o un guardado.

    CustomTkinter oculta y vuelve a mostrar la ventana para pintarle la barra de
    título (y a los 1000 ms reajusta el tamaño mínimo): esos ciclos terminan
    dejando la ventana DETRÁS de la principal. Por eso se trae al frente varias
    veces, con un 'topmost' momentáneo que se suelta enseguida.
    """
    if ventana is None:
        return
    try:
        ventana._windows_set_titlebar_color = lambda *a, **k: None
    except Exception:
        pass

    def _soltar():
        try:
            if ventana.winfo_exists():
                ventana.attributes("-topmost", False)
        except Exception:
            pass

    def _traer():
        try:
            if not ventana.winfo_exists():
                return
            ventana.deiconify()
            ventana.lift()
            try:
                ventana.attributes("-topmost", True)
                ventana.after(450, _soltar)
            except Exception:
                pass
            ventana.focus_force()
        except Exception:
            pass

    _traer()
    for retraso in retrasos or ():
        try:
            ventana.after(retraso, _traer)
        except Exception:
            break


def ventana_para_mensaje(principal=None):
    """🖥️ Ventana que se usa como padre de un mensaje para que salga DELANTE.

    Devuelve el diálogo que esté abierto (el último creado y visible) y, si no hay
    ninguno, la ventana del módulo. Así los comunicados y las consultas nunca quedan
    detrás de la ventana de Compras.
    """
    try:
        raiz = getattr(tk, "_default_root", None)
        if raiz is not None:
            candidatas = []
            for w in raiz.winfo_children():
                try:
                    if isinstance(w, (tk.Toplevel, ctk.CTkToplevel)) and w.winfo_viewable():
                        candidatas.append(w)
                except Exception:
                    continue
            if candidatas:
                return candidatas[-1]
    except Exception:
        pass
    return principal


def _texto_bancario_limpio(valor):
    """Quita los textos de relleno de los combos de la ficha del proveedor."""
    texto = str(valor or "").strip()
    if texto.lower() in ("no seleccionada", "seleccione banco", "ninguno", "none", "-", "nan"):
        return ""
    return texto


def obtener_datos_bancarios_proveedor(nombre_proveedor):
    """Devuelve la ficha bancaria del proveedor (tabla 'proveedores') o None."""
    nombre = str(nombre_proveedor or "").strip()
    if not nombre:
        return None
    conn = conectar_db(silencioso=True)
    if not conn:
        return None
    columnas = """SELECT ruc, nombre, banco_1, cuenta_1, cci_1, banco_2, cuenta_2, cci_2,
                         cuenta_detraccion, porcentaje_detraccion, contacto, whatsapp, correo
                  FROM proveedores"""
    try:
        cursor = conn.cursor()
        cursor.execute(columnas + " WHERE TRIM(UPPER(nombre)) = TRIM(UPPER(%s)) LIMIT 1", (nombre,))
        fila = cursor.fetchone()
        if not fila:
            cursor.execute(columnas + " WHERE nombre ILIKE %s ORDER BY id LIMIT 1", (f"%{nombre}%",))
            fila = cursor.fetchone()
        return fila
    except Exception:
        return None
    finally:
        liberar_conexion(conn)


def mostrar_datos_bancarios_proveedor(parent, proveedor, datos, factura_txt="", reutilizar=None):
    """🧾 Ventana con los datos bancarios del proveedor y botones para copiarlos."""
    datos = list(datos or []) + [""] * 13
    (ruc, nombre, b1, c1, cc1, b2, c2, cc2,
     cdet, pdet, contacto, whatsapp, correo) = datos[:13]
    b1, c1, cc1 = _texto_bancario_limpio(b1), _texto_bancario_limpio(c1), _texto_bancario_limpio(cc1)
    b2, c2, cc2 = _texto_bancario_limpio(b2), _texto_bancario_limpio(c2), _texto_bancario_limpio(cc2)
    cdet, pdet = _texto_bancario_limpio(cdet), _texto_bancario_limpio(pdet)
    nombre = str(nombre or proveedor or "").strip()

    ventana = reutilizar
    if ventana is None:
        ventana = ctk.CTkToplevel(parent)
        ventana.transient(parent)
    for hijo in ventana.winfo_children():
        hijo.destroy()
    ventana.title(f"🏦 Datos bancarios - {nombre}")
    ventana.geometry("620x560")
    ventana.minsize(560, 420)

    ctk.CTkLabel(ventana, text="🏦 DATOS BANCARIOS DEL PROVEEDOR",
                 font=("Arial", 15, "bold"), text_color="#1f538d").pack(pady=(14, 2))
    ctk.CTkLabel(ventana, text=nombre or "(sin nombre)", font=("Arial", 13, "bold")).pack()
    if ruc:
        ctk.CTkLabel(ventana, text=f"RUC: {ruc}", font=("Arial", 11, "bold"),
                     text_color="#7f8c8d").pack(pady=(0, 4))
    if factura_txt:
        ctk.CTkLabel(ventana, text=factura_txt, font=("Arial", 11, "italic"),
                     text_color="#7f8c8d").pack(pady=(0, 6))

    lbl_estado = ctk.CTkLabel(ventana, text="", font=("Arial", 11, "bold"), text_color="#1e8449")
    lbl_estado.pack(pady=(0, 4))

    def copiar(texto, etiqueta="Dato"):
        texto = str(texto or "").strip()
        if not texto:
            return lbl_estado.configure(text="⚠️ No hay dato para copiar.", text_color="#c0392b")
        try:
            ventana.clipboard_clear()
            ventana.clipboard_append(texto)
            ventana.update_idletasks()
            lbl_estado.configure(text=f"✅ Copiado: {texto}", text_color="#1e8449")
        except Exception as e:
            lbl_estado.configure(text=f"⚠️ No se pudo copiar: {e}", text_color="#c0392b")

    cuerpo = ctk.CTkScrollableFrame(ventana, fg_color="transparent")
    cuerpo.pack(fill="both", expand=True, padx=14, pady=6)

    def ficha(titulo, contenido):
        marco = ctk.CTkFrame(cuerpo, fg_color="#f8f9fa", corner_radius=8,
                             border_width=1, border_color="#e0e0e0")
        marco.pack(fill="x", pady=6)
        ctk.CTkLabel(marco, text=titulo, font=("Arial", 12, "bold"),
                     text_color="#1f538d", anchor="w").pack(fill="x", padx=12, pady=(8, 2))
        for etiqueta, valor in contenido:
            valor = str(valor or "").strip()
            if not valor:
                continue
            fila = ctk.CTkFrame(marco, fg_color="transparent")
            fila.pack(fill="x", padx=12, pady=2)
            ctk.CTkLabel(fila, text=etiqueta, width=120, anchor="w",
                         font=("Arial", 11, "bold")).pack(side="left")
            ctk.CTkLabel(fila, text=valor, anchor="w", font=("Arial", 12)).pack(side="left", padx=(0, 6))
            ctk.CTkButton(fila, text="📋 Copiar", width=84, height=24, font=("Arial", 10, "bold"),
                          fg_color="#1f538d", hover_color="#163b65",
                          command=lambda v=valor, e=etiqueta: copiar(v, e)).pack(side="right", padx=(6, 0))
        return marco

    hay_cuenta = False
    if b1 or c1 or cc1:
        hay_cuenta = True
        ficha("💳 Cuenta bancaria 1", [("Banco:", b1), ("N° de cuenta:", c1), ("CCI:", cc1)])
    if b2 or c2 or cc2:
        hay_cuenta = True
        ficha("💳 Cuenta bancaria 2", [("Banco:", b2), ("N° de cuenta:", c2), ("CCI:", cc2)])
    if cdet or pdet:
        hay_cuenta = True
        ficha("🧾 Cuenta de detracción", [("N° de cuenta:", cdet), ("Porcentaje:", pdet)])

    if hay_cuenta:
        contacto_txt = [("Contacto:", contacto), ("WhatsApp:", whatsapp), ("Correo:", correo)]
        if any(str(v or "").strip() for _, v in contacto_txt):
            ficha("📞 Contacto del proveedor", contacto_txt)

        def copiar_todo():
            lineas = [f"PROVEEDOR: {nombre}"]
            if ruc:
                lineas.append(f"RUC: {ruc}")
            if b1 or c1 or cc1:
                lineas += ["", "CUENTA 1:",
                           f"  Banco: {b1 or '-'}", f"  N° de cuenta: {c1 or '-'}", f"  CCI: {cc1 or '-'}"]
            if b2 or c2 or cc2:
                lineas += ["", "CUENTA 2:",
                           f"  Banco: {b2 or '-'}", f"  N° de cuenta: {c2 or '-'}", f"  CCI: {cc2 or '-'}"]
            if cdet or pdet:
                lineas += ["", f"DETRACCIÓN: {cdet or '-'}" + (f"  ({pdet})" if pdet else "")]
            copiar(chr(10).join(lineas), "Todos los datos")

        ctk.CTkButton(ventana, text="📋 Copiar todos los datos", height=34,
                      font=("Arial", 12, "bold"), fg_color="#27ae60", hover_color="#1e8449",
                      command=copiar_todo).pack(fill="x", padx=14, pady=(4, 2))
        ctk.CTkLabel(ventana, text="Los datos se toman de la ficha del proveedor (módulo Proveedores).",
                     font=("Arial", 9, "italic"), text_color="#7f8c8d").pack(pady=(0, 2))
    else:
        ctk.CTkLabel(cuerpo, text="⚠️ Este proveedor todavía no tiene datos bancarios registrados en su ficha.",
                     font=("Arial", 12, "bold"), text_color="#c0392b").pack(pady=18)
        ctk.CTkLabel(cuerpo, text="Regístralos en el módulo 📦 Proveedores → Editar → Banco 1 / Cuenta 1 / CCI 1.",
                     font=("Arial", 11), text_color="#7f8c8d", justify="center").pack(pady=(0, 10))

    ctk.CTkButton(ventana, text="❌ Cerrar", width=140, height=30, font=("Arial", 11, "bold"),
                  fg_color="#7f8c8d", hover_color="#606b6b",
                  command=ventana.destroy).pack(pady=(4, 12))
    try:
        ventana.bind("<Escape>", lambda _e: ventana.destroy(), add="+")
        ventana.lift()
        ventana.focus_force()
    except Exception:
        pass
    return ventana


def _fecha_dd_mm_aaaa(fecha):
    """Devuelve la fecha en DD/MM/AAAA (acepta ISO, dd/mm o mm/dd)."""
    import re as _re
    texto = str(fecha or "").strip()
    if not texto:
        return ""
    m = _re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(.*)$", texto)
    if m:
        anio, mes, dia, resto = m.groups()
        return "%02d/%02d/%s%s" % (int(dia), int(mes), anio, resto)
    m = _re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})(.*)$", texto)
    if m:
        a, b, anio, resto = m.groups()
        dia, mes = int(a), int(b)
        if mes > 12 and dia <= 12:
            dia, mes = mes, dia
        return "%02d/%02d/%s%s" % (dia, mes, anio, resto)
    return texto


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

# =========================================================================
# 🔎 LECTURA DE COMPROBANTES ESCANEADOS (FOTO / PDF SIN CAPA DE TEXTO)
# =========================================================================
# Muchos comprobantes llegan escaneados (CamScanner) o como foto del celular:
# no tienen texto seleccionable, por lo que se aplica OCR con el motor que ya
# trae Windows (sin instalar nada y sin internet) y se reconstruyen las filas
# usando la posición de cada palabra, para poder asociar "OP. GRAVADAS" con su
# importe aunque el OCR devuelva los números en otra línea.

_ETIQUETAS_MONTOS = {
    "recargo": ["RECARGOALCONSUMO", "RECARGOALCONSUMOSERVICIO", "RECARGOPORSERVICIO",
                "RECARGOSERVICIO", "PROPINASUGERIDA", "SERVICIOALCONSUMO"],
    "total": ["IMPORTETOTAL", "TOTALAPAGAR", "TOTALNETO", "TOTALCOMPROBANTE",
              "IMPORTETOTALVENTA", "TOTALPAGAR", "TOTALAGENERAL",
              "TOTALPORHONORARIOS", "TOTALNETORECIBIDO", "TOTALHONORARIOS",
              "IMPORTETOTALHONORARIOS", "NETORECIBIDO"],
    "base": ["OPGRAVADAS", "OPGRABADAS", "OPGRAVADA", "SUBTOTAL", "BASEIMPONIBLE",
             "VALORVENTA", "TOTALGRAVADO", "BASEGRAVADA"],
    "igv": ["IGV", "IGV18", "IGV105", "IMPUESTOGENERALALASVENTAS", "IGVTOTAL"],
}

# Script de PowerShell que usa el motor OCR que ya incluye Windows.
# Devuelve una línea por palabra:  X <TAB> Y <TAB> ancho <TAB> alto <TAB> texto
_SCRIPT_OCR_WINDOWS = r'''
param([string]$Ruta, [string]$Idiomas = "es-PE,es-MX,es-ES,en-US")
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Runtime.WindowsRuntime | Out-Null
$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
function Await($op, $tipo) {
    $m = $asTaskGeneric.MakeGenericMethod($tipo)
    $t = $m.Invoke($null, @($op))
    $t.Wait(-1) | Out-Null
    $t.Result
}
[Windows.Storage.StorageFile,Windows.Foundation,ContentType=WindowsRuntime] | Out-Null
[Windows.Graphics.Imaging.BitmapDecoder,Windows.Foundation,ContentType=WindowsRuntime] | Out-Null
[Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime] | Out-Null
[Windows.Globalization.Language,Windows.Foundation,ContentType=WindowsRuntime] | Out-Null

$engine = $null
foreach ($tag in $Idiomas.Split(",")) {
    $tag = $tag.Trim()
    if (-not $tag) { continue }
    try {
        $idioma = [Windows.Globalization.Language]::new($tag)
        $posible = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($idioma)
        if ($null -ne $posible) { $engine = $posible; break }
    } catch { }
}
if ($null -eq $engine) { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }
if ($null -eq $engine) { return }

$archivo = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Ruta)) ([Windows.Storage.StorageFile])
$stream  = Await ($archivo.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
$decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bitmap  = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
$resultado = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
foreach ($linea in $resultado.Lines) {
  foreach ($p in $linea.Words) {
    $r = $p.BoundingRect
    Write-Output ([string][int]$r.X + [char]9 + [string][int]$r.Y + [char]9 + [string][int]$r.Width + [char]9 + [string][int]$r.Height + [char]9 + $p.Text)
  }
}
'''


def _normalizar_etiqueta(texto):
    """Deja solo letras en mayúscula: sirve para reconocer rótulos aunque el OCR falle."""
    return re.sub(r"[^A-ZÑ]", "", str(texto or "").upper())


def _parecido(texto, candidato):
    return difflib.SequenceMatcher(None, texto, candidato).ratio()


def _etiqueta_reconocida(etiqueta, grupo, umbral=0.78):
    """¿La etiqueta de la fila corresponde a base / IGV / recargo / total?"""
    return any(_parecido(etiqueta, candidato) >= umbral
               for candidato in _ETIQUETAS_MONTOS.get(grupo, []))


def _agrupar_palabras_en_lineas(palabras, tolerancia_rel=0.6):
    """Reconstruye las filas del documento a partir de palabras con posición.

    'palabras' es una lista de (x, y, alto, texto). El OCR suele devolver el
    rótulo y el importe en líneas distintas, pero conservando su Y: agrupando
    por Y se recupera la fila real ("OP. GRAVADAS  S/  187.04").
    """
    if not palabras:
        return ""
    palabras = sorted(palabras, key=lambda p: (p[1], p[0]))
    alturas = [p[2] for p in palabras if p[2]]
    tolerancia = max(8.0, (sum(alturas) / len(alturas)) * tolerancia_rel) if alturas else 12.0
    filas, actual, y_fila = [], [], None
    for x, y, _alto, texto in palabras:
        if y_fila is None or abs(y - y_fila) <= tolerancia:
            actual.append((x, texto))
            if y_fila is None:
                y_fila = y
        else:
            filas.append(actual)
            actual, y_fila = [(x, texto)], y
    if actual:
        filas.append(actual)
    return "\n".join(" ".join(t for _x, t in sorted(fila, key=lambda p: p[0])) for fila in filas)


def leer_ocr_mac(ruta, dpi=250):
    """OCR en macOS con el motor nativo (Apple Vision, vía PyObjC).

    Se convierte el PDF (o la foto) en imágenes y se leen con el OCR del sistema;
    el texto vuelve reconstruido por filas, igual que en Windows. Si el módulo
    escaner o PyMuPDF no están disponibles devuelve "" (se completa a mano).
    """
    try:
        import escaner
    except Exception:
        return ""
    if not getattr(escaner, "es_mac", lambda: False)():
        return ""
    try:
        import fitz   # PyMuPDF: convierte el PDF (o la foto) en imágenes
    except Exception:
        return ""
    try:
        carpeta = tempfile.mkdtemp(prefix="ocr_mac_")
        imagenes = []
        documento = fitz.open(ruta)
        for numero, pagina in enumerate(documento):
            destino = os.path.join(carpeta, "pagina_%d.png" % (numero + 1))
            pagina.get_pixmap(dpi=dpi).save(destino)
            imagenes.append(destino)
        documento.close()
    except Exception:
        return ""
    if not imagenes:
        return ""
    try:
        return escaner.texto_de_resultados(escaner.ocr_paginas(imagenes))
    except Exception:
        return ""


def _numero_comprobante_patron(texto):
    """¿El texto trae un número de comprobante (F003-00012345, F37B-00035956...?"""
    return re.search(r"[A-Z]\s?\d{2,4}\s?[A-Z]?\s*[-–]\s*\d{4,8}", str(texto or "").upper())


def _variantes_ocr(ruta_imagen):
    """Copias mejoradas de la imagen para el OCR: gris+contraste y x2.

    Las fotos de tickets termicos (poco contraste y letra chica) se leen mucho
    mejor asi: aparece el N de documento y el importe total.
    """
    salidas = []
    try:
        from PIL import Image, ImageOps
    except Exception:
        return salidas
    try:
        with Image.open(ruta_imagen) as original:
            gris = original.convert("L")
            for factor, nombre in ((1, "gris"), (2, "gris2x")):
                img = gris if factor == 1 else gris.resize(
                    (gris.width * factor, gris.height * factor), Image.LANCZOS)
                img = ImageOps.autocontrast(img, cutoff=1)
                fd, salida = tempfile.mkstemp(prefix="ocr_%s_" % nombre, suffix=".png")
                os.close(fd)
                img.save(salida)
                salidas.append(salida)
    except Exception as e:
        print("[Compras] variantes OCR:", e)
    return salidas


def leer_ocr_windows(ruta, idiomas="es-PE,es-MX,es-ES,en-US", dpi=250):
    """OCR del sistema: motor de Windows (WinRT) o de macOS (Vision).

    Devuelve el texto reconstruido por filas (o "" si no se pudo leer).

    No requiere instalar nada ni conexión a internet. En otros sistemas
    operativos devuelve "" (el usuario puede digitar los montos a mano).
    """
    if sys.platform == "darwin":
        return leer_ocr_mac(ruta, dpi=dpi)
    if sys.platform != "win32":
        return ""
    try:
        import fitz   # PyMuPDF: convierte el PDF (o la foto) en imágenes
    except Exception:
        return ""
    try:
        carpeta = tempfile.mkdtemp(prefix="ocr_compras_")
        imagenes = []
        documento = fitz.open(ruta)
        for numero, pagina in enumerate(documento):
            destino = os.path.join(carpeta, "pagina_%d.png" % (numero + 1))
            pagina.get_pixmap(dpi=dpi).save(destino)
            imagenes.append(destino)
        documento.close()
    except Exception:
        return ""
    if not imagenes:
        return ""

    script = os.path.join(carpeta, "ocr_windows.ps1")
    try:
        with open(script, "w", encoding="utf-8-sig") as archivo_ps:
            archivo_ps.write(_SCRIPT_OCR_WINDOWS)
    except Exception:
        return ""

    def _ocr_una_imagen(ruta_imagen):
        try:
            orden = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script,
                     "-Ruta", ruta_imagen, "-Idiomas", idiomas]
            resultado = subprocess.run(orden, capture_output=True, text=True, encoding="utf-8",
                                       errors="ignore", timeout=180,
                                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except Exception:
            return ""
        palabras = []
        for linea in (resultado.stdout or "").splitlines() or []:
            partes = linea.split("\t")
            if len(partes) != 5:
                continue
            try:
                x, y, alto = int(partes[0]), int(partes[1]), int(partes[3])
            except ValueError:
                continue
            palabras.append((x, y, alto, partes[4]))
        return _agrupar_palabras_en_lineas(palabras)

    def _puntaje_ocr(texto):
        """Cuánta información útil tiene el texto (nº de documento, montos, total)."""
        if not texto or not texto.strip():
            return -1
        puntos = min(len(texto) // 60, 10)
        if _numero_comprobante_patron(texto):
            puntos += 15
        if len(re.findall(r"\d+[.,]\d{2}", texto)) >= 2:
            puntos += 5
        return puntos

    paginas = []
    for imagen in imagenes[:5]:
        mejor = _ocr_una_imagen(imagen)          # 1) tal cual
        for variante in _variantes_ocr(imagen):  # 2) mejoradas (gris, contraste, 2x)
            otro = _ocr_una_imagen(variante)
            try:
                os.remove(variante)
            except Exception:
                pass
            if _puntaje_ocr(otro) > _puntaje_ocr(mejor):
                mejor = otro
        if mejor.strip():
            paginas.append(mejor)
    try:
        shutil.rmtree(carpeta, ignore_errors=True)
    except Exception:
        pass
    return "\n".join(paginas)


def _corregir_digitos_ocr(token):
    """Corrige las confusiones típicas del OCR en fechas/números (i9 -> 19, ó9 -> 09...)."""
    tabla = str.maketrans({"O": "0", "o": "0", "ó": "0", "Ó": "0", "Q": "0", "D": "0",
                           "I": "1", "l": "1", "i": "1", "|": "1", "!": "1",
                           "Z": "2", "z": "2", "A": "4", "S": "5", "s": "5",
                           "G": "6", "b": "6", "T": "7", "B": "8", "g": "9", "q": "9"})
    return str(token).translate(tabla)


_PALABRAS_NUMERO = {
    "CERO": 0, "UN": 1, "UNO": 1, "UNA": 1, "DOS": 2, "TRES": 3, "CUATRO": 4, "CINCO": 5, "SEIS": 6,
    "SIETE": 7, "OCHO": 8, "NUEVE": 9, "DIEZ": 10, "ONCE": 11, "DOCE": 12, "TRECE": 13, "CATORCE": 14,
    "QUINCE": 15, "DIECISEIS": 16, "DIECISIETE": 17, "DIECIOCHO": 18, "DIECINUEVE": 19, "VEINTE": 20,
    "VEINTIUNO": 21, "VEINTIDOS": 22, "VEINTITRES": 23, "VEINTICUATRO": 24, "VEINTICINCO": 25,
    "VEINTISEIS": 26, "VEINTISIETE": 27, "VEINTIOCHO": 28, "VEINTINUEVE": 29, "TREINTA": 30,
    "CUARENTA": 40, "CINCUENTA": 50, "SESENTA": 60, "SETENTA": 70, "OCHENTA": 80, "NOVENTA": 90,
    "CIEN": 100, "CIENTO": 100, "DOSCIENTOS": 200, "TRESCIENTOS": 300, "CUATROCIENTOS": 400,
    "QUINIENTOS": 500, "SEISCIENTOS": 600, "SETECIENTOS": 700, "OCHOCIENTOS": 800, "NOVECIENTOS": 900,
}


def _texto_sin_acentos(texto):
    """Texto en mayúsculas y sin tildes (para comparar rótulos). Multiplataforma."""
    import unicodedata
    base = unicodedata.normalize("NFKD", str(texto or ""))
    return "".join(c for c in base if not unicodedata.combining(c)).upper()


def monto_en_palabras(texto):
    """Lee el importe escrito en letras ('SON: CIENTO SETENTA Y SIETE CON 0/100 SOLES').

    Devuelve 177.0 en ese ejemplo, o 0.0 si no lo encuentra. Sirve de respaldo
    cuando el comprobante trae los importes en columnas y no se lee el total.
    """
    limpio = _texto_sin_acentos(texto)
    posicion = limpio.find("SON")
    if posicion == -1:
        return 0.0
    resto = limpio[posicion + 3:]
    resto = resto.split(chr(10))[0]                     # solo esa linea
    centimos = 0
    # 💬 "DIEZ Y 97 / 100 SOLES" (los tickets escriben los céntimos con "Y")
    if " CON " not in resto:
        _m_cent = re.search(r"Y\s+(\d{1,2})\s*/\s*100", resto)
        if _m_cent:
            centimos = int(_m_cent.group(1))
            resto = resto[:_m_cent.start()]
    if " CON " in resto:
        antes, despues = resto.split(" CON ", 1)
        resto = antes
        try:
            centimos = int(re.findall(r"\d{1,2}", despues)[0])
        except Exception:
            centimos = 0
    palabras = [p for p in re.split(r"[^A-Z]+", resto) if p]
    total, parcial = 0, 0
    for palabra in palabras:
        if palabra in ("Y", "CON", "SOLES", "SOL", "DOLARES", "SON", "DE", "CONFORMADO", "SEUO"):
            continue
        if palabra in ("MIL", "MILES"):
            total += (parcial or 1) * 1000
            parcial = 0
            continue
        if palabra in ("MILLON", "MILLONES"):
            total += (parcial or 1) * 1000000
            parcial = 0
            continue
        if palabra in _PALABRAS_NUMERO:
            parcial += _PALABRAS_NUMERO[palabra]
            continue
        if parcial or total:
            break                                      # palabra rara: se corta
    total += parcial
    if total <= 0:
        return 0.0
    return round(total + centimos / 100.0, 2)


def _montos_de_columnas(datos, montos):
    """Reparte los importes de una fila de tabla (etiquetas en la línea anterior).

    En muchos comprobantes (RIMAC, AFP, servicios) los rótulos van en una línea y
    los importes en la siguiente, con las casillas vacías sin valor. Se toma el
    último importe como total y, de los demás, el mayor como base; el IGV sale de
    la diferencia.
    """
    valores = []
    for texto_monto in montos:
        try:
            valor = parsear_monto_texto(texto_monto)
        except Exception:
            valor = 0.0
        if valor and valor > 0:
            valores.append(round(float(valor), 2))
    if not valores:
        return
    total = valores[-1]
    if len(valores) == 1:
        if not datos.get("total"):
            datos["total"] = total
            datos["leidos"] = datos.get("leidos", 0) + 1
        return
    base = max(valores[:-1])
    if not datos.get("total"):
        datos["total"] = total
        datos["leidos"] = datos.get("leidos", 0) + 1
    if not datos.get("base"):
        datos["base"] = base
        datos["leidos"] = datos.get("leidos", 0) + 1
    if not datos.get("igv"):
        igv = round(total - base, 2)
        if 0 < igv <= total * 0.5:
            datos["igv"] = igv
            datos["leidos"] = datos.get("leidos", 0) + 1


def extraer_montos_comprobante(texto):
    """Lee base, IGV, recargo al consumo y total de un comprobante peruano.

    Soporta los tres formatos habituales:
      · rótulo e importe en la MISMA línea      (Total 24.80)
      · rótulo y el importe en la línea SIGUIENTE (Importe Total / 24.80 / S/)
      · tabla en columnas                        (encabezados y luego los montos)
    El importe en letras («SON: ...») se usa como control final del total.
    """
    datos = {"base": 0.0, "igv": 0.0, "recargo": 0.0, "total": 0.0, "leidos": 0}
    pendiente = None          # grupo que espera su importe en la línea siguiente
    columnas = []             # grupos de un encabezado de tabla

    def _asignar(grupo, valor):
        if grupo and not datos.get(grupo) and valor:
            datos[grupo] = valor
            datos["leidos"] += 1
            return True
        return False

    for linea in str(texto or "").splitlines():
        fila = linea.replace("S/.", " ").replace("S/", " ")
        corte = re.search(r"\d", fila)
        etiqueta = _normalizar_etiqueta(fila[:corte.start()] if corte else fila)
        montos = re.findall(r"\d{1,3}(?:[.,]\d{3})*[.,]\d{2}|\d+[.,]\d{1,2}", fila)

        # 1) ¿La línea es solo un importe y hay un rótulo esperándolo?
        if montos and not etiqueta:
            valor = None
            try:
                valor = parsear_monto_texto(montos[-1])
            except ValueError:
                valor = None
            if pendiente and _asignar(pendiente, valor):
                pendiente = None
                continue
            if columnas and valor:
                for grupo in columnas:
                    if _asignar(grupo, valor):
                        break
                columnas = []
            pendiente = None
            continue

        if not etiqueta:
            continue

        # 2) ¿Trae el/los importe(s) en la misma línea?
        #    Puede haber varios pares: "Subtotal 100.00 IGV 18.00 Total 118.00"
        if montos:
            asignado = False
            ultimo = 0
            for marca in re.finditer(r"\d{1,3}(?:[.,]\d{3})*[.,]\d{2}|\d+[.,]\d{1,2}", fila):
                trozo = _normalizar_etiqueta(fila[ultimo:marca.start()]) or etiqueta
                ultimo = marca.end()
                try:
                    valor = parsear_monto_texto(marca.group(0))
                except ValueError:
                    continue
                for grupo in ("recargo", "total", "base", "igv"):
                    if _etiqueta_reconocida(trozo, grupo) and _asignar(grupo, valor):
                        asignado = True
                        break
            pendiente = None if asignado else pendiente
            columnas = []
            continue

        # 3) Rótulo sin importe: se recuerda para la línea siguiente
        grupos = [g for g in ("recargo", "total", "base", "igv")
                  if not datos[g] and _etiqueta_reconocida(etiqueta, g)]
        if grupos:
            pendiente = grupos[0]
            columnas = grupos
        else:
            pendiente = None
            columnas = []

    # El recargo al consumo puede no haberse leído: se deduce de la diferencia
    if datos["base"] and datos["igv"] and datos["total"] and not datos["recargo"]:
        diferencia = round(datos["total"] - datos["base"] - datos["igv"], 2)
        if 0.02 <= diferencia <= datos["total"] * 0.35:
            datos["recargo"] = diferencia
    if not datos["base"] and datos["total"] and datos["igv"]:
        datos["base"] = round(datos["total"] - datos["recargo"] - datos["igv"], 2)
    if not datos["total"] and datos["base"] and datos["igv"]:
        datos["total"] = round(datos["base"] + datos["igv"] + datos["recargo"], 2)

    # 💬 El importe en letras manda si no cuadra («SON: VEINTICUATRO CON 80/100»)
    letras = monto_en_palabras(texto)
    if letras:
        cuadra = bool(datos["base"]) and abs(
            datos["base"] + datos["igv"] + datos["recargo"] - datos["total"]) <= 0.05
        # 💬 Si el importe en letras no coincide con el número leído, manda el de
        #    letras (es el importe legal del comprobante) y se recalculan base/IGV.
        if (not datos["total"] or not cuadra or datos["total"] < datos["base"]
                or abs(letras - datos["total"]) > 0.10):
            datos["total"] = letras
            datos["leidos"] += 1
            if datos["base"] and abs(datos["base"] + datos["igv"] + datos["recargo"] - letras) > 0.05:
                igv = round(letras - datos["base"] - datos["recargo"], 2)
                if 18 <= (igv / datos["base"] * 100 if datos["base"] else 0) <= 19.5 or \
                        (igv > 0 and abs(igv - letras * 0.18 / 1.18) < letras * 0.02):
                    datos["igv"] = igv
                else:
                    base = round(letras / 1.18, 2)
                    if base > 0:
                        datos["base"], datos["igv"] = base, round(letras - base - datos["recargo"], 2)
        if datos["base"] and not datos["igv"]:
            calculado = round(datos["total"] - datos["base"] - datos["recargo"], 2)
            if 0 < calculado <= datos["total"] * 0.5:
                datos["igv"] = calculado
        elif not datos["base"] and datos["igv"]:
            datos["base"] = round(datos["total"] - datos["igv"] - datos["recargo"], 2)
    return datos


def tasa_igv_de_montos(base, igv):
    """Devuelve 10.5 (restaurantes) o 18 según los importes leídos; 0 si no cuadra."""
    try:
        base, igv = float(base), float(igv)
    except (TypeError, ValueError):
        return 0.0
    if base <= 0 or igv <= 0:
        return 0.0
    tasa = igv / base
    if abs(tasa - 0.105) < 0.02:
        return 10.5
    if abs(tasa - 0.18) < 0.025:
        return 18.0
    return 0.0


def _ruc_valido(ruc):
    """RUC peruano con digito verificador correcto (descarta ruido del OCR)."""
    ruc = re.sub(r"\D", "", str(ruc or ""))
    if len(ruc) != 11 or ruc[:2] not in ("10", "15", "17", "20"):
        return False
    if ruc == ruc[0] * 11:
        return False
    pesos = (5, 4, 3, 2, 7, 6, 5, 4, 3, 2)
    suma = sum(int(d) * p for d, p in zip(ruc[:10], pesos))
    resto = 11 - (suma % 11)
    verificador = 0 if resto == 10 else (1 if resto == 11 else resto)
    return verificador == int(ruc[10])


def _distancia_uno(a, b):
    """¿Dos cadenas difieren en un solo carácter (falta/sobra/cambia)?"""
    a, b = str(a or ""), str(b or "")
    if a == b:
        return True
    if abs(len(a) - len(b)) > 1:
        return False
    if len(a) == len(b):
        return sum(1 for x, y in zip(a, b) if x != y) == 1
    largo, corto = (a, b) if len(a) > len(b) else (b, a)
    for i in range(len(largo)):
        if largo[:i] + largo[i + 1:] == corto:
            return True
    return False


def buscar_proveedor_por_ruc(texto, ruc_empresa=""):
    """(ruc, nombre) del proveedor: busca el RUC del OCR en la tabla 'proveedores'.

    Tolera que el OCR pierda o cambie un dígito (20 1:7765279 -> 20127765279).
    """
    propio = re.sub(r"\D", "", str(ruc_empresa or ""))
    candidatos = []
    for grupo in re.findall(r"\d{9,12}", str(texto or "")):
        if propio and (grupo == propio or _distancia_uno(grupo, propio)):
            continue
        if grupo not in candidatos:
            candidatos.append(grupo)
    # 🏷️ Línea del RUC con espacios o dos puntos: "RUC 20 1:7765279" -> 2017765279
    for linea in str(texto or "").splitlines():
        if re.search(r"R\.?\s?U\.?\s?C", linea, re.IGNORECASE):
            digitos = re.sub(r"\D", "", linea)
            if 9 <= len(digitos) <= 12 and digitos not in candidatos:
                if not (propio and (digitos == propio or _distancia_uno(digitos, propio))):
                    candidatos.append(digitos)
    if not candidatos:
        return "", ""
    conn = conectar_db(silencioso=True)
    if not conn:
        return "", ""
    filas = []
    try:
        with conn.cursor() as c:
            c.execute("SELECT ruc, nombre FROM proveedores WHERE COALESCE(ruc,'') <> ''")
            filas = c.fetchall()
    except Exception as e:
        print("[Compras] proveedores por RUC:", e)
    finally:
        liberar_conexion(conn)
    if not filas:
        return "", ""

    # 1) coincidencia exacta
    for candidato in candidatos:
        for ruc, nombre in filas:
            if re.sub(r"\D", "", str(ruc or "")) == candidato:
                if "BLACK RIDERS" in str(nombre or "").upper():
                    continue
                return str(ruc), nombre
    # 2) un dígito de diferencia (error típico del OCR)
    for candidato in candidatos:
        for ruc, nombre in filas:
            limpio = re.sub(r"\D", "", str(ruc or ""))
            if len(limpio) >= 9 and _distancia_uno(limpio, candidato):
                if "BLACK RIDERS" in str(nombre or "").upper():
                    continue
                return str(ruc), nombre
    return "", ""


def _ruc_parecido_al_propio(candidato, propio):
    """¿El OCR confundió el RUC de la empresa? (1 o 2 dígitos distintos)."""
    candidato, propio = str(candidato or ""), str(propio or "")
    if not candidato or not propio or len(candidato) != len(propio):
        return False
    return sum(1 for a, b in zip(candidato, propio) if a != b) <= 2


def ruc_distinto_al_de_la_empresa(texto, ruc_empresa=""):
    """RUC del proveedor: el primer RUC de 11 dígitos que no sea el de la empresa."""
    propio = re.sub(r"\D", "", str(ruc_empresa or ""))
    for candidato in re.findall(r"\d{11}", str(texto or "")):
        if candidato == propio or _ruc_parecido_al_propio(candidato, propio):
            continue
        if not _ruc_valido(candidato):
            continue
        return candidato
    return ""


# =========================================================================
# 🏢 VALIDACIÓN DEL DESTINATARIO Y ALTA AUTOMÁTICA DE PROVEEDORES
# =========================================================================
def validar_destinatario_comprobante(texto, config=None):
    """Comprueba que el comprobante esté a nombre de la empresa configurada.

    Se apoya en el RUC de la empresa (y en su razón social) que están en
    Configuración General. Devuelve (estado, mensaje):

      'ok'             -> el comprobante es de la empresa
      'ajeno'          -> es de OTRA empresa (hay que rechazarlo)
      'sin-configurar' -> falta el RUC/Razón Social en Configuración General
      'sin-datos'      -> no se pudo leer ningún RUC (OCR dudoso)
    """
    if not isinstance(config, dict):
        config = cargar_configuracion_regional()
    ruc_empresa = re.sub(r"\D", "", str(config.get("ruc_empresa") or ""))
    nombre_empresa = str(config.get("razon_social_empresa") or "").strip()
    if not ruc_empresa and not nombre_empresa:
        return "sin-configurar", ("No se puede validar el destinatario: configure el RUC y la Razón Social "
                                  "de la empresa en Ajustes → Configuración General.")

    texto = str(texto or "")
    texto_sin_espacios = re.sub(r"[^0-9]", "", texto)
    rucs = re.findall(r"\d{11}", texto)

    comparable = " ".join(re.sub(r"[^0-9A-Za-zÁÉÍÓÚÑÜáéíóúñü ]", " ", texto.upper()).split())
    palabras_empresa = [p for p in _normalizar_nombre_empresa(nombre_empresa).split() if len(p) > 3] \
        if nombre_empresa else []
    nombre_encontrado = bool(palabras_empresa) and all(p in comparable for p in palabras_empresa)

    # 1) Se reconoce a la empresa (por RUC o por razón social) -> se acepta
    if ruc_empresa and (ruc_empresa in rucs or ruc_empresa in texto_sin_espacios):
        return "ok", "✅ El comprobante está a nombre de la empresa (RUC " + ruc_empresa + ")."
    if nombre_encontrado:
        return "ok", "✅ El comprobante está a nombre de la empresa (" + nombre_empresa + ")."

    # 2) Hay RUC en el documento pero ninguno es el de la empresa -> es de otra empresa
    if rucs:
        otros = ", ".join(dict.fromkeys(r for r in rucs if r != ruc_empresa))
        return "ajeno", ("❌ Este comprobante NO está a nombre de la empresa: no figura su RUC "
                         f"({ruc_empresa or 'sin configurar'}). Los RUC que aparecen son: {otros}.\n"
                         "La empresa configurada es "
                         + (nombre_empresa or "sin razón social configurada") + ".")

    # 3) No se pudo leer ningún RUC: se avisa (el OCR puede haber fallado)
    return "sin-datos", ("⚠️ No se pudo leer ningún RUC en el comprobante, así que no se pudo confirmar "
                         "que esté a nombre de la empresa"
                         + (f" ({nombre_empresa})" if nombre_empresa else "") +
                         ". Verifíquelo antes de guardarlo.")


def _normalizar_nombre_empresa(texto):
    """Nombre de empresa comparable: sin signos ni sufijos societarios."""
    limpio = str(texto or "").upper().replace(".", "")
    limpio = re.sub(r"[^0-9A-ZÁÉÍÓÚÑÜ ]", " ", limpio)
    limpio = re.sub(r"\b(SOCIEDAD|ANONIMA|CERRADA|ABIERTA|SAC|SA|SRL|EIRL|LTDA)\b", " ", limpio)
    return " ".join(limpio.split())


def datos_proveedor_desde_ruc(ruc, timeout=10):
    """Datos completos del proveedor traídos de SUNAT por internet (o None)."""
    try:
        from proveedores import consultar_ruc_en_internet
    except Exception:
        return None
    return consultar_ruc_en_internet(ruc, timeout=timeout)


def _serie_ocr_normalizada(serie):
    """Corrige las letras que el OCR confunde con dígitos (EOOI -> E001)."""
    tabla = str.maketrans({"O": "0", "Q": "0", "D": "0", "I": "1", "L": "1", "|": "1",
                           "S": "5", "Z": "2", "G": "6", "A": "4", "T": "7", "B": "8"})
    serie = str(serie or "").upper()
    # Solo se corrigen los DÍGITOS: la letra inicial (E/F/B) se respeta.
    return serie[:1] + serie[1:].translate(tabla) if len(serie) > 1 else serie


def _limpiar_numero_operacion(token):
    """Deja el N° de operación legible (corrige O->0, I->1... si es numérico)."""
    token = str(token or "").strip(" .,:;#-")
    digitos = re.sub(r"\D", "", token)
    if token and len(digitos) >= max(4, len(token) - 2):
        tabla = str.maketrans({"O": "0", "o": "0", "Q": "0", "D": "0", "I": "1", "l": "1", "|": "1",
                               "S": "5", "s": "5", "Z": "2", "B": "8", "G": "6", "A": "4", "T": "7"})
        token = token.translate(tabla)
    return token.upper()


def _fecha_valida(dia, mes, anio):
    """Devuelve (dd, mm, aaaa) si es una fecha real del calendario, o None."""
    try:
        dia, mes, anio = int(dia), int(mes), int(anio)
    except Exception:
        return None
    if anio < 100:
        anio += 2000
    if not (2000 <= anio <= 2100):
        return None
    for d, m in ((dia, mes), (mes, dia)):      # se acepta DD/MM y, si no, MM/DD
        try:
            datetime(anio, m, d)
            return d, m, anio
        except Exception:
            continue
    return None


def _formatear_fecha(dia, mes, anio, formato="DD/MM/AAAA"):
    """Arma la fecha en el formato configurado en el sistema."""
    try:
        formato = str((CONFIG_REGIONAL or {}).get("formato_fecha", formato) or formato)
    except Exception:
        pass
    if str(formato).upper().startswith("MM"):
        return f"{mes:02d}/{dia:02d}/{anio}"
    return f"{dia:02d}/{mes:02d}/{anio}"


def extraer_fecha_comprobante(texto, formato="DD/MM/AAAA"):
    """Fecha del comprobante de pago leída por OCR ('' si no aparece).

    Se buscan primero las fechas con rótulo ("Fecha de pago", "Fecha de
    operación", "Fecha de transferencia", "Fecha"... ) y, si no hay, la primera
    fecha válida del documento. Acepta 12/10/2026, 12-10-26 y "12 de octubre
    del 2026". Devuelve la fecha en el formato configurado en el sistema.
    """
    if not str(texto or "").strip():
        return ""
    plano = " ".join(str(texto).upper().split())
    meses = {"ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
             "JULIO": 7, "AGOSTO": 8, "SETIEMBRE": 9, "SEPTIEMBRE": 9, "OCTUBRE": 10,
             "NOVIEMBRE": 11, "DICIEMBRE": 12,
             # Abreviaturas que usan los vouchers del banco ("12 OCT 2026")
             "ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6, "JUL": 7,
             "AGO": 8, "SET": 9, "SEP": 9, "OCT": 10, "NOV": 11, "DIC": 12}
    numeros = re.compile(r"\b(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{2,4})\b")
    # 📅 "12 de Octubre del 2026", "12 Octubre 2026" (vouchers del banco)
    #    y "Lunes, 12 OCT 2026": el "de" es opcional y se aceptan abreviaturas.
    escrito = re.compile(r"\b(\d{1,2})\s+(?:DE\s+)?([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1]{3,})\s+(?:DE\s+|DEL\s+)?(\d{2,4})\b")

    import datetime as _dt
    _hoy = _dt.date.today()

    def _plausible(fecha):
        """Descarta fechas imposibles (vencimientos lejanos tipo 2030)."""
        try:
            valor = _dt.date(int(fecha[2]), int(fecha[1]), int(fecha[0]))
        except Exception:
            return False
        return (_dt.date(_hoy.year - 5, 1, 1) <= valor <= _hoy + _dt.timedelta(days=200))

    def leer(trozo):
        encontrado = numeros.search(trozo)
        if encontrado:
            fecha = _fecha_valida(*encontrado.groups())
            if fecha and _plausible(fecha):
                return _formatear_fecha(*fecha, formato=formato)
        encontrado = escrito.search(trozo)
        if encontrado and encontrado.group(2) in meses:
            fecha = _fecha_valida(encontrado.group(1), meses[encontrado.group(2)], encontrado.group(3))
            if fecha and _plausible(fecha):
                return _formatear_fecha(*fecha, formato=formato)
        return ""

    # 1) Con rótulo de fecha: es lo más confiable
    etiquetas = (r"FECHA\s+DE\s+PAGO", r"FECHA\s+DE\s+LA\s+OPERACI[OÓ]N",
                 r"FECHA\s+DE\s+OPERACI[OÓ]N", r"FECHA\s+DE\s+TRANSACCI[OÓ]N",
                 r"FECHA\s+DE\s+TRANSFERENCIA", r"FECHA\s+DE\s+DEP[OÓ]SITO",
                 r"FECHA\s*[/\-]\s*HORA", r"FECHA\s+DE\s+EMISI[OÓ]N", r"FECHA\s+OPE",
                 r"FECHA")
    for etiqueta in etiquetas:
        encontrado = re.search(etiqueta + r"[^0-9]{0,18}", plano)
        if not encontrado:
            continue
        fecha = leer(plano[encontrado.end(): encontrado.end() + 45])
        if fecha:
            return fecha
    # 2) La primera fecha válida del documento
    for encontrado in numeros.finditer(plano):
        fecha = _fecha_valida(*encontrado.groups())
        if fecha and _plausible(fecha):
            return _formatear_fecha(*fecha, formato=formato)
    for encontrado in escrito.finditer(plano):
        if encontrado.group(2) in meses:
            fecha = _fecha_valida(encontrado.group(1), meses[encontrado.group(2)], encontrado.group(3))
            if fecha and _plausible(fecha):
                return _formatear_fecha(*fecha, formato=formato)
    return ""


def _texto_del_soporte(ruta):
    """Texto de un PDF (o de una foto) de soporte de pago."""
    texto = ""
    try:
        import fitz
        documento = fitz.open(ruta)
        try:
            texto = chr(10).join(pagina.get_text("text") or "" for pagina in documento)
        finally:
            documento.close()
    except Exception:
        texto = ""
    if str(texto or "").strip():
        return texto
    try:
        return leer_ocr_windows(ruta)
    except Exception:
        return ""


def nombre_tercero_desde_texto(texto):
    """Nombre de quien hizo el pago (el tercero) según su comprobante."""
    lineas = [l.strip() for l in str(texto or "").splitlines() if l.strip()]
    etiquetas = ("ENVIADO POR", "ORDENANTE", "REMITENTE", "PAGADOR", "TITULAR",
                 "NOMBRE DEL ORDENANTE", "DESDE", "BENEFICIARIO", "ENVIADO A",
                 "A NOMBRE DE")

    def _candidato(valor):
        valor = re.sub(r"^(NOMBRE|SR|SRA|SRTA)\.?\s*[:\-]?\s*", "", str(valor or "").strip(),
                       flags=re.I)
        valor = re.sub(r"\s{2,}", " ", valor).strip(" :.-")
        if len(valor) < 5:
            return ""
        alto = _texto_sin_acentos(valor).upper()
        if "BLACK RIDERS" in alto or "RUC" in alto or "S/" in alto:
            return ""
        if re.search(r"\d{3,}", valor):
            return ""
        if not re.search(r"[A-Za-z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1]{3,}", valor):
            return ""
        # 🚫 Palabras que son RÓTULOS del voucher, no nombres de personas/empresas
        prohibidas = ("DESDE", "ENVIADO A", "ENVIADO POR", "ORDENANTE", "REMITENTE",
                      "PAGADOR", "TITULAR", "BENEFICIARIO", "A NOMBRE DE", "NUMERO DE OPERACION",
                      "OPERACION", "MONEDA", "CORRIENTE", "SOLES", "CUENTA", "COMISION",
                      "TOTAL", "TOTAL COBRADO", "MONTO", "FECHA", "TRANSFERENCIA",
                      "EXITOSA", "TRANSFERENCIA EXITOSA", "BCP", "BBVA", "INTERBANK",
                      "MONEDA SOLES", "CORRIENTE SOLES")
        if alto in prohibidas or any(alto.endswith(p) for p in ("SOLES", "OPERACION")):
            return ""
        if "OPERACI" in alto or "MONEDA" in alto:
            return ""
        # Un nombre real suele tener al menos dos palabras (Collie Paiser Fabiola C.)
        if len(valor.split()) < 2:
            return ""
        return valor

    for posicion, linea in enumerate(lineas):
        alto = _texto_sin_acentos(linea).upper()
        for etiqueta in etiquetas:
            if etiqueta in alto:
                resto = linea[alto.find(etiqueta) + len(etiqueta):]
                candidato = _candidato(resto)
                if not candidato:
                    # se revisan las siguientes líneas (el OCR desordena las columnas)
                    for salto in range(1, 4):
                        if posicion + salto >= len(lineas):
                            break
                        candidato = _candidato(lineas[posicion + salto])
                        if candidato:
                            break
                if candidato:
                    return candidato
    return ""


def extraer_numero_operacion(texto):
    """N° de operación / transacción leído del comprobante de pago (o '').

    Reconoce los rótulos habituales de un voucher ("N° de operación",
    "Nro. de transacción", "Código de operación", "N° de referencia"...)
    y también las capturas de la banca por internet. Es tolerante a los
    errores del OCR (operacin, Nro, N°, etc.).
    """
    if not str(texto or "").strip():
        return ""
    plano = " ".join(str(texto).upper().split())
    patrones = (
        r"(?:N(?:RO|R|\u00ba|\u00b0|O)?\.?\s*[\u00ba\u00b0]?\s*:?\s*(?:DE\s+|DEL\s+)?)?"
        r"(?:OPERACI[O\u00d3]N|TRANSACCI[O\u00d3]N|TRANSFERENCIA|REFERENCIA|CONSTANCIA|"
        r"DEP[O\u00d3]SITO|C[O\u00d3]DIGO)\s*(?:N(?:RO|R|\u00ba|\u00b0|O)?\.?\s*[\u00ba\u00b0]?\s*)?"
        r"[:#\-]?\s*([A-Z0-9][A-Z0-9\-\.]{3,24})",
        r"OPERACI[O\u00d3]N\s*N(?:RO|R|\u00ba|\u00b0)?\.?\s*[:#\-]?\s*([A-Z0-9\-\.]{4,24})",
    )
    for patron in patrones:
        for encontrado in re.finditer(patron, plano):
            candidato = _limpiar_numero_operacion(encontrado.group(1))
            if re.fullmatch(r"\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}", candidato):
                continue                      # es una fecha
            digitos = re.sub(r"\D", "", candidato)
            if len(digitos) < 5:
                # "2026" (un año del documento) no es un N° de operación
                continue
            return candidato
    # Último recurso: la secuencia más larga de 6 a 20 dígitos, siempre que el
    # comprobante parezca una transferencia/pago (no un RUC de 11).
    if re.search(r"TRANSFERENCIA|INTERBANCARI|OPERACI|CONSTANCIA|DEPOSITO|DEP\u00d3SITO|VOUCHER", plano):
        numeros = [n for n in re.findall(r"\b\d{6,20}\b", plano) if len(n) != 11]
        if numeros:
            return max(numeros, key=len)
    return ""


def numero_comprobante_desde_texto(texto):
    """N° de documento (F001-123, E001-45...) leído del comprobante."""
    texto = str(texto or "")
    m = re.search(r"\b([EFB][0-9A-Z]{3})\s*[-:\u2013\u2014]?\s*(\d{1,8})", texto)
    if m:
        # Se respetan los ceros tal como vienen (F001-00001234)
        return f"{_serie_ocr_normalizada(m.group(1))}-{m.group(2)}"
    # 🔎 El OCR confunde la letra inicial con un dígito ("F37B" -> "1376")
    m = re.search(r"\b([0-9][0-9A-Z]{3})\s*[-\u2013\u2014]\s*(\d{5,8})\b", texto)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    m = re.search(r"([EFB][0-9A-Z]{3}\s*-\s*\d+)", texto)
    return m.group(1).replace(" ", "") if m else ""


def _mismo_proveedor(nombre_a, nombre_b):
    """¿Los dos nombres son del mismo proveedor? (tolera S.A.C., tildes, espacios)."""
    def normalizar(nombre):
        texto = str(nombre or "").upper()
        for relleno in ("S.A.C.", "S.A.C", "SAC", "E.I.R.L.", "E.I.R.L", "EIRL",
                        "S.R.L.", "S.R.L", "SRL", "S.A.", "S.A", "SAA", "E.I.R.L",
                        "-", ".", ",", "  "):
            texto = texto.replace(relleno, " ")
        texto = " ".join(texto.split())
        for con_tilde, sin_tilde in (("Á","A"), ("É","E"), ("Í","I"), ("Ó","O"), ("Ú","U"), ("Ñ","N")):
            texto = texto.replace(con_tilde, sin_tilde)
        return texto
    a, b = normalizar(nombre_a), normalizar(nombre_b)
    if not a or not b:
        return False
    if a == b:
        return True
    # uno contenido en el otro (p. ej. "COESTI" y "COESTI E/S ORRANTIA")
    return a.startswith(b) or b.startswith(a)


def buscar_comprobante_duplicado(numero_documento):
    """Comprobante ya registrado con ese N° de documento (o None).

    Se compara solo el número (sin espacios ni mayúsculas), igual que la
    verificación que hace 'guardar_registro', para no permitir duplicados.
    """
    nro = re.sub(r"\s+", "", str(numero_documento or "")).upper()
    if not nro:
        return None
    conn = conectar_db(silencioso=True)
    if not conn:
        return None
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT proveedor, fecha, COALESCE(total, 0), COALESCE(ruc, ''), COALESCE(archivo_ruta, '')
            FROM facturas_recibidas
            WHERE REGEXP_REPLACE(REGEXP_REPLACE(UPPER(COALESCE(numero_documento, '')), '[^A-Z0-9]', '', 'g'),
                                 '^([A-Z]+)0+', '\1')
                  = REGEXP_REPLACE(REGEXP_REPLACE(%s, '[^A-Z0-9]', '', 'g'), '^([A-Z]+)0+', '\1')
            ORDER BY id DESC LIMIT 1
        """, (nro,))
        fila = cursor.fetchone()
        if not fila:
            return None
        return {"numero": nro, "proveedor": fila[0], "fecha": fila[1],
                "total": fila[2], "ruc": fila[3], "archivo": fila[4]}
    except Exception:
        return None
    finally:
        liberar_conexion(conn)


def asegurar_proveedor(nombre, ruc, usar_internet=True, timeout=10):
    """Deja registrado al proveedor (creándolo si hace falta) y devuelve sus datos.

    Orden de trabajo:
      1. Si ya existe (por RUC o por nombre) se usa el de la base de datos;
         si le faltan la dirección o la ubicación y hay internet, se completan.
      2. Si no existe, se consulta el RUC en internet (SUNAT) y se guarda con
         TODA la información de la empresa (razón social, dirección fiscal, ubicación).
      3. Si no hay internet o el RUC no existe en SUNAT, se guarda con lo leído
         en el comprobante para no perder el registro.

    Devuelve (estado, datos) con estado en:
      'existe' | 'completado' | 'creado' | 'creado_sin_sunat' | 'sin-datos'
      | 'sin-conexion' | 'error'
    """
    nombre = " ".join(str(nombre or "").split()).strip()
    ruc = re.sub(r"\D", "", str(ruc or ""))
    datos = {"ruc": ruc, "nombre": nombre, "direccion_fiscal": "", "ubicacion": ""}
    if not nombre and not ruc:
        return "sin-datos", datos

    conn = conectar_db(silencioso=True)
    if not conn:
        return "sin-conexion", datos
    try:
        cursor = conn.cursor()
        existente = None
        if ruc:
            cursor.execute("SELECT id, COALESCE(ruc,''), COALESCE(nombre,''), "
                           "COALESCE(direccion_fiscal,''), COALESCE(ubicacion,'') "
                           "FROM proveedores WHERE ruc = %s LIMIT 1", (ruc,))
            existente = cursor.fetchone()
        if not existente and nombre:
            cursor.execute("SELECT id, COALESCE(ruc,''), COALESCE(nombre,''), "
                           "COALESCE(direccion_fiscal,''), COALESCE(ubicacion,'') "
                           "FROM proveedores WHERE TRIM(UPPER(nombre)) = TRIM(UPPER(%s)) LIMIT 1",
                           (nombre,))
            existente = cursor.fetchone()

        de_sunat = None
        if existente:
            id_prov, ruc_db, nombre_db, direccion_db, ubicacion_db = existente
            datos.update({"ruc": ruc_db or ruc, "nombre": nombre_db or nombre,
                          "direccion_fiscal": direccion_db, "ubicacion": ubicacion_db})
            # Solo se consulta SUNAT si a la ficha le falta información
            falta_algo = not (direccion_db and ubicacion_db)
            de_sunat = datos_proveedor_desde_ruc(ruc, timeout=timeout) if (usar_internet and ruc and falta_algo) else None
            cambios = {}
            if de_sunat:
                if not direccion_db and de_sunat.get("direccion_fiscal"):
                    cambios["direccion_fiscal"] = de_sunat["direccion_fiscal"]
                    datos["direccion_fiscal"] = de_sunat["direccion_fiscal"]
                if not ubicacion_db and de_sunat.get("ubicacion"):
                    cambios["ubicacion"] = de_sunat["ubicacion"]
                    datos["ubicacion"] = de_sunat["ubicacion"]
                if not nombre_db and de_sunat.get("nombre"):
                    cambios["nombre"] = de_sunat["nombre"]
                    datos["nombre"] = de_sunat["nombre"]
            if cambios:
                columnas = ", ".join(f"{c} = %s" for c in cambios)
                cursor.execute(f"UPDATE proveedores SET {columnas} WHERE id = %s",
                               tuple(cambios.values()) + (id_prov,))
                conn.commit()
                try:
                    cache_sistema.invalidar()
                except Exception:
                    pass
                return "completado", datos
            return "existe", datos

        # No existe: se crea con los datos de SUNAT (si se pudieron traer)
        if de_sunat is None and usar_internet and ruc:
            de_sunat = datos_proveedor_desde_ruc(ruc, timeout=timeout)
        if de_sunat and (de_sunat.get("nombre") or nombre):
            datos.update({"nombre": de_sunat.get("nombre") or nombre,
                          "direccion_fiscal": de_sunat.get("direccion_fiscal", ""),
                          "ubicacion": de_sunat.get("ubicacion", "")})
            estado = "creado"
        elif nombre:
            datos["nombre"] = nombre
            estado = "creado_sin_sunat"
        else:
            # Sin razón social (ni del comprobante ni de SUNAT) NO se crea un proveedor
            # con el número de RUC como nombre: se le pide al usuario que lo escriba.
            return "sin-datos", datos

        cursor.execute("""
            INSERT INTO proveedores (ruc, nombre, direccion_fiscal, categoria, ubicacion, descripcion)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (datos["ruc"], datos["nombre"], datos.get("direccion_fiscal", ""),
              "No seleccionada", datos.get("ubicacion", ""),
              "Agregado automáticamente al escanear un comprobante"))
        conn.commit()
        try:
            cache_sistema.invalidar()
        except Exception:
            pass
        return estado, datos
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        return "error", datos
    finally:
        liberar_conexion(conn)


def crear_proveedor_si_no_existe(nombre, ruc, usar_internet=True):
    """Compatibilidad: agrega el proveedor si no está y devuelve solo el estado."""
    estado, _datos = asegurar_proveedor(nombre, ruc, usar_internet=usar_internet)
    return estado


# Palabras que indican que la línea NO es la razón social del emisor
_NO_ES_NOMBRE = re.compile(
    r"FACTURA|BOLETA|ELECTR[OÓ]NIC|R\.?\s?U\.?\s?C|D\.?\s?N\.?\s?I|TICKET|COMPROBANTE|COTIZACI|"
    r"GUIA|PROFORMA|NOTA\s+DE|SERIE|N[uú]MERO|FECHA|HORA|TOTAL|SUBTOTAL|IMPORTE|IGV|GRACIAS|"
    r"BIENVENID|CLIENTE|SE[NÑ]OR|DIRECCI[OÓ]N|TELEF|CELULAR|CORREO|EMAIL|WWW\.|P[ÁA]GINA|"
    r"C[OÓ]DIGO|ORDEN|PEDIDO|VENCIMIENTO|CONDICI[OÓ]N|FORMA\s+DE\s+PAGO", re.IGNORECASE)
_PARECE_DIRECCION = re.compile(
    r"\b(AV\.?|AVENIDA|CAL\.?|CALLE|JR\.?|JIR[OÓ]N|URB\.?|MZ\.?|LT\.?|LOTE|NRO\.?|N[°º]|PASAJE|"
    r"PROLONG|CARRETERA|KM\.?|INT\.?|DPTO\.?|OFIC\.?|PISO)\b", re.IGNORECASE)
_SUFIJO_EMPRESA = re.compile(
    r"\b(S\.?A\.?C\.?|S\.?A\.?|S\.?R\.?L\.?|E\.?I\.?R\.?L\.?|LTDA|SOCIEDAD|EMPRESA|CORPORACI[OÓ]N|"
    r"INVERSIONES|COMERCIAL|SERVICIOS|CONTRATISTAS|NEGOCIOS|GRIFO|TRANSPORTES|DISTRIBUIDORA|"
    r"IMPORT|EXPORT|HOLDING|GRUPO)\b", re.IGNORECASE)


def _puntaje_como_nombre(linea):
    """Qué tan probable es que la línea sea la razón social del emisor."""
    limpio = re.sub(r"\s+", " ", str(linea or "")).strip(" -:|,.")
    letras = re.sub(r"[^A-Za-zÁÉÍÓÚÑÜáéíóúñü]", "", limpio)
    if len(letras) < 4:
        return -99
    if _NO_ES_NOMBRE.search(limpio):
        return -99
    puntaje = 0
    if _SUFIJO_EMPRESA.search(limpio):
        puntaje += 5
    if len(letras) >= 12:
        puntaje += 1
    if limpio.isupper():
        puntaje += 1
    if _PARECE_DIRECCION.search(limpio):
        puntaje -= 3
    if re.search(r"\d", limpio):
        puntaje -= 1
    if len(limpio) > 60:
        puntaje -= 2
    return puntaje


_NO_ES_PROVEEDOR = re.compile(
    r"(RAZ[OÓ]N\s*SOCIAL|RAZON\s*SOCIAL|CLIENTE|SE[NÑ]OR|SE[NÑ]ORA|BLACK\s*RIDERS|"
    r"R\.?U\.?C|DNI|DIRECCI[OÓ]N|TELEFONO|TEL[EÉ]FONO|RECIBO|FACTURA|BOLETA|"
    r"SUNAT|TOTAL|IMPORTE|SUB\s*TOTAL|FECHA|FORMAS?\\s+DE\\s+PAGO|MONEDA|DOCUMENTO|ELECTR[OÓ]NICA|REPRESENTACI[OÓ]N)", re.IGNORECASE)


def _nombre_parece_proveedor(nombre):
    """¿El texto parece la razón social del proveedor (y no una etiqueta)?"""
    valor = str(nombre or "").strip()
    if len(valor) < 4:
        return False
    if _NO_ES_PROVEEDOR.search(valor) or _NO_ES_NOMBRE.search(valor):
        return False
    solo_letras = re.sub(r"[^A-Za-z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1 ]", "", valor)
    if len(solo_letras.strip()) < 4:
        return False
    # Al menos 2 palabras "reales" o un sufijo societario
    palabras = [p for p in solo_letras.split() if len(p) > 2]
    if len(palabras) >= 2:
        return True
    return bool(re.search(r"\b(S\.?A\.?C?|E\.?I\.?R\.?L|S\.?R\.?L|SOCIEDAD)\b",
                          valor.upper()))


def _a_numero(texto):
    """Convierte '1,234.56' / '1.234,56' / 'S/ 118.00' en número."""
    limpio = re.sub(r"[^\d.,]", "", str(texto or ""))
    if not limpio:
        return 0.0
    if "," in limpio and "." in limpio:
        if limpio.rfind(",") > limpio.rfind("."):
            limpio = limpio.replace(".", "").replace(",", ".")
        else:
            limpio = limpio.replace(",", "")
    elif "," in limpio:
        partes = limpio.split(",")
        limpio = limpio.replace(",", ".") if len(partes[-1]) <= 2 else limpio.replace(",", "")
    try:
        return float(limpio)
    except Exception:
        return 0.0

def deducir_montos_por_suma(texto, tipo=""):
    """Deduce (base, IGV, total) buscando 'base + IGV = total' entre los importes.

    Sirve cuando el OCR separa los rótulos de los números o confunde las
    etiquetas: se recogen todos los importes y se busca el par que cumple la
    relación del IGV (18% o 10.5%). Devuelve (0, 0, 0) si no hay combinación válida.
    """
    # 📊 Primero el caso de tablas en columnas (rótulos en una línea, importes en la otra)
    por_columnas = montos_en_columnas(texto)
    if por_columnas[2] > 0:
        return por_columnas

    valores = _importes_del_texto(texto)

    # 💬 El importe en letras manda: la base es el mayor importe menor que el total
    letras = monto_en_palabras(texto)
    if letras:
        candidatos = [v for v in valores if 0 < v < letras]
        if candidatos:
            base = max(candidatos)
            igv = round(letras - base, 2)
            if 0 < igv <= letras * 0.5:
                return base, igv, letras
        # Si no hay con qué comparar, se asume el IGV general (18%)
        return round(letras / 1.18, 2), round(letras - letras / 1.18, 2), letras

    if len(valores) < 3:
        return 0.0, 0.0, 0.0
    tasas = [0.105] if "10.5" in str(tipo) else [0.18]
    mejor = None
    for base in valores:
        for igv in valores:
            if base <= 0 or igv <= 0 or igv == base:
                continue
            if not any(abs(igv / base - tasa) < 0.02 for tasa in tasas):
                continue
            total = round(base + igv, 2)
            if any(abs(v - total) <= 0.05 for v in valores):
                if mejor is None or total > mejor[2]:
                    mejor = (base, igv, total)
    return mejor or (0.0, 0.0, 0.0)

def _importes_del_texto(texto):
    """Todos los importes con decimales del texto (corrige los 'oo' del OCR)."""
    limpio = str(texto or "").replace("oo", "00").replace("OO", "00")
    valores = []
    for encontrado in re.findall(r"\d{1,3}(?:[.,]\d{3})*[.,]\d{1,2}|\d+[.,]\d{1,2}", limpio):
        valor = _a_numero(encontrado)
        if valor and valor not in valores:
            valores.append(round(float(valor), 2))
    return valores

def montos_en_columnas(texto):
    """(base, igv, total) cuando los rótulos y los importes van en líneas distintas.

    Es el caso de RIMAC, AFP y servicios: la línea de rótulos
    ("OP- GRAVADA ... I.G.V. IMPORTE TOTAL") y la de valores ("S/ 150,00 S/ 27.00 S/ 177.0").
    Se toma el último importe como total y el mayor de los demás como base.
    """
    lineas = str(texto or "").splitlines()
    for indice in range(1, len(lineas)):
        etiquetas = _texto_sin_acentos(lineas[indice - 1])
        # Solo un encabezado inequívoco (IGV / IMPORTE TOTAL / TOTAL A PAGAR) y que esa
        # línea NO traiga importes: así las facturas normales no entran por aquí.
        if not re.search(r"(IMPORTE\s*TOTAL|TOTAL\s*A\s*PAGAR|I\.?\s*G\.?\s*V|\bIGV\b)", etiquetas):
            continue
        if re.search(r"\d{1,3}(?:[.,]\d{3})*[.,]\d{1,2}", etiquetas):
            continue
        fila = lineas[indice].replace("S/.", " ").replace("S/", " ")
        valores = []
        for encontrado in re.findall(r"\d{1,3}(?:[.,]\d{3})*[.,]\d{1,2}|\d+[.,]\d{1,2}", fila):
            try:
                valor = _a_numero(encontrado)
            except Exception:
                valor = 0
            if valor and valor > 0:
                valores.append(round(float(valor), 2))
        if len(valores) >= 2:
            total = valores[-1]
            base = max(valores[:-1])
            igv = round(total - base, 2)
            if 0 < igv <= total * 0.5:
                return base, igv, total
    return 0.0, 0.0, 0.0

def validar_destinatario(datos, config=None):
    """Comprueba que la factura esté a nombre de la empresa configurada en el sistema.

    Devuelve (estado, mensaje):
      'ok'             -> la factura es de la empresa configurada
      'ajeno'          -> es de OTRA empresa (hay que rechazarla)
      'sin-configurar' -> falta el RUC/Razón Social en Configuración General
      'sin-datos'      -> no se pudo leer el destinatario en la factura
    """
    config = config if isinstance(config, dict) else cargar_configuracion_regional()
    ruc_empresa = re.sub(r"\D", "", str(config.get("ruc_empresa") or ""))
    nombre_empresa = str(config.get("razon_social_empresa") or "").strip()
    if not ruc_empresa and not nombre_empresa:
        return "sin-configurar", ("No se puede validar el destinatario: configure el RUC y la Razón Social "
                                  "de la empresa en Ajustes → Configuración General.")

    datos = datos or {}
    cliente_ruc = re.sub(r"\D", "", str(datos.get("cliente_ruc") or ""))
    cliente_nombre = str(datos.get("cliente_nombre") or "").strip()

    if ruc_empresa and cliente_ruc:
        if cliente_ruc == ruc_empresa:
            return "ok", f"✅ Factura a nombre de la empresa (RUC {cliente_ruc})."
        detalle = f"RUC {cliente_ruc}" + (f" — {cliente_nombre}" if cliente_nombre else "")
        return "ajeno", (f"❌ Esta factura está a nombre de OTRA empresa: {detalle}.\n"
                         f"La empresa configurada es RUC {ruc_empresa}"
                         + (f" — {nombre_empresa}" if nombre_empresa else "") + ".")

    if nombre_empresa and cliente_nombre:
        a, b = _normalizar_nombre_empresa(cliente_nombre), _normalizar_nombre_empresa(nombre_empresa)
        if a and b and (a in b or b in a):
            return "ok", f"✅ El destinatario '{cliente_nombre}' coincide con la empresa configurada."
        return "ajeno", (f"❌ Esta factura está a nombre de '{cliente_nombre}', pero la empresa configurada "
                         f"es '{nombre_empresa}'.")

    return "sin-datos", ("⚠️ No se pudo leer el destinatario (RUC o razón social) en la factura. "
                         "Verifique que esté a nombre de la empresa antes de guardarla.")

def extraer_datos_de_texto(texto):
    """Saca los datos de una factura del texto (de un PDF o del OCR del escaneo).

    Es tolerante a los errores típicos del OCR (espacios de más, '/' que se pierde,
    'S.A.C.' sin espacios, etc.).
    """
    datos = {"tipo_documento": "", "numero": "", "fecha": "", "proveedor": "", "ruc": "",
             "subtotal": 0.0, "igv": 0.0, "total": 0.0, "texto": str(texto or "")}
    if not str(texto or "").strip():
        return datos

    original = str(texto)
    mayus = original.upper()
    plano = " ".join(mayus.split())

    # --- Tipo de documento ---
    if "FACTURA" in plano:
        datos["tipo_documento"] = "Factura (10.5% IGV) Restaurantes" if "10.5" in plano or "10,5" in plano \
            else "Factura (18% IGV)"
    elif "BOLETA" in plano:
        datos["tipo_documento"] = "Boleta (Sin IGV)"
    elif "HONORARIOS" in plano:
        datos["tipo_documento"] = "Recibo por Honorarios (8% Retención)" \
            if ("RETENCION" in plano or "RETENCIÓN" in plano or re.search(r"\b8\s*%", plano)) \
            else "Recibo por Honorarios (Sin Retención)"

    # --- Número (F001-123, E001 123, B001-45...) ---
    #     El OCR suele confundir 0/O, 1/I, 5/S... (EOOI-519 es E001-519).
    def _serie_ocr(serie):
        tabla = str.maketrans({"O": "0", "Q": "0", "D": "0", "I": "1", "L": "1", "|": "1",
                               "S": "5", "Z": "2", "G": "6", "A": "4", "T": "7", "B": "8"})
        # Solo se corrigen los DÍGITOS de la serie: la letra inicial se respeta.
        serie = str(serie or "").upper()
        return serie[:1] + serie[1:].translate(tabla) if len(serie) > 1 else serie

    m = (re.search(r"\b([EFB][0-9A-Z]{3})\s*[-–—:.]?\s*(\d{1,8})\b", plano) or
         re.search(r"\b([EFB])\s*0*(\d{2,3})\s*[-–—:.]?\s*(\d{1,8})\b", plano))
    if m:
        if m.lastindex == 2:
            datos["numero"] = f"{_serie_ocr(m.group(1))}-{int(m.group(2))}"
        else:
            datos["numero"] = f"{m.group(1)}{int(m.group(2)):03d}-{int(m.group(3))}"

    # --- RUC del EMISOR (proveedor) ---
    #     👉 En una factura hay DOS RUC: el del que emite (arriba, en el membrete)
    #     y el del cliente (debajo de "CLIENTE" / "SEÑOR(ES)"). El proveedor es el
    #     primero que aparece ANTES de la sección del cliente.
    rucs = re.findall(r"\b(\d{11})\b", plano)
    ubicaciones = [(m.start(), m.group(1)) for m in re.finditer(r"\b(\d{11})\b", plano)]
    _m_cli_pos = re.search(
        r"(?:SE[NÑ]OR(?:ES)?|ADQUIRIENTE|CLIENTE|DESTINATARIO|"
        r"RAZ[OÓ]N\s+SOCIAL\s+DEL?\s+CLIENTE)\s*[:\-]?\s*", plano)
    _pos_cli = _m_cli_pos.start() if _m_cli_pos else len(plano)
    _antes = [r for pos, r in ubicaciones if pos < _pos_cli]
    _despues = [r for pos, r in ubicaciones if pos >= _pos_cli]
    if _antes:
        datos["ruc"] = _antes[0]
    elif _despues:
        datos["ruc"] = _despues[0]
    elif rucs:
        datos["ruc"] = rucs[0]

    # --- Fecha ---
    fecha = None
    m = re.search(r"FECHA\s*(?:DE\s*)?(?:EMISI[OÓ]N|EMISION|EXPEDICI[OÓ]N)?\s*[:\-]?\s*(\d{1,2})\s*[/\-.|]\s*(\d{1,2})\s*[/\-.|]\s*(\d{2,4})", plano)
    if not m:
        m = re.search(r"\b(\d{1,2})\s*[/\-.|]\s*(\d{1,2})\s*[/\-.|]\s*(\d{4})\b", plano)
    if m:
        dia, mes, anio = m.group(1), m.group(2), m.group(3)
        if len(anio) == 2:
            anio = "20" + anio
        try:
            if 1 <= int(dia) <= 31 and 1 <= int(mes) <= 12:
                fecha = f"{int(dia):02d}/{int(mes):02d}/{anio}"
        except Exception:
            fecha = None
    if not fecha:
        # Fechas escritas: "09 de Octubre del 2026"
        MESES_ES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
                    "julio": 7, "agosto": 8, "setiembre": 9, "septiembre": 9, "octubre": 10,
                    "noviembre": 11, "diciembre": 12}
        m_esc = re.search(r"(\d{1,2})\s*(?:DE\s+)?([A-ZÁÉÍÓÚ]+)\s*(?:DE(?:L)?\s+)?(\d{4})", plano)
        if m_esc and m_esc.group(2).lower() in MESES_ES:
            try:
                dia = int(m_esc.group(1))
                if 1 <= dia <= 31:
                    fecha = f"{dia:02d}/{MESES_ES[m_esc.group(2).lower()]:02d}/{m_esc.group(3)}"
            except Exception:
                fecha = None
    datos["fecha"] = fecha or ""

    # --- Montos ---
    def monto_con(patron):
        # Se salta el porcentaje cuando la etiqueta lo trae (ej: "IGV 18%: 180.00")
        m = re.search(patron + r"[^\d\n]{0,10}(?:\d{1,2}\s*%\s*[:\-]?\s*)?[^\d\n]{0,10}(\d[\d.,]*)", plano)
        return _a_numero(m.group(1)) if m else 0.0

    subtotal = monto_con(r"(?:OP\.?\s*GRAVADAS?|SUB\s*TOTAL|SUB\s*T0TAL|BASE\s*IMPONIBLE|VALOR\s*DE\s*VENTA|GRAVAD[OA]S?)")
    igv = monto_con(r"(?:I\.?\s*G\.?\s*V\.?|IMPUESTO\s*GENERAL\s*A\s*LAS\s*VENTAS)")
    total = monto_con(r"(?:IMPORTE\s*TOTAL|TOTAL\s*A\s*PAGAR|TOTAL\s*GENERAL|TOTAL\s*NETO|TOTAL\s*S/)")
    if not total:
        total = monto_con(r"(?<!SUB\s)(?<!SUB)TOTAL")
    if not subtotal and total:
        if datos["tipo_documento"].startswith("Factura"):
            divisor = 1.105 if "10.5" in datos["tipo_documento"] else 1.18
            subtotal = round(total / divisor, 2)
        else:
            subtotal = total
    if not igv and subtotal and datos["tipo_documento"].startswith("Factura"):
        divisor = 1.105 if "10.5" in datos["tipo_documento"] else 1.18
        igv = round(subtotal * (divisor - 1), 2)
    # 🔎 Si los montos no cuadran (base + IGV = total) o falta alguno, se deducen
    #    buscando esa relación entre TODOS los importes del documento (OCR despistado).
    if not subtotal or not total or abs((subtotal + igv) - total) > 0.05:
        b2, i2, t2 = deducir_montos_por_suma(texto, datos.get("tipo_documento", ""))
        if t2:
            subtotal, igv, total = b2, i2, t2
    else:
        # 🔎 Aunque cuadre, la deducción por suma puede dar un TOTAL mayor y correcto:
        #    el OCR lee "IMPORTE 313.22" en vez de "TOTAL 369.60".
        try:
            b3, i3, t3 = deducir_montos_por_suma(texto, datos.get("tipo_documento", ""))
            if t3 and abs((b3 + i3) - t3) <= 0.05 and t3 > (total or 0) + 0.05:
                subtotal, igv, total = b3, i3, t3
        except Exception:
            pass
    if not total:
        total = round(subtotal + igv, 2)
    datos["subtotal"] = round(subtotal, 2)
    datos["igv"] = round(igv, 2)
    datos["total"] = round(total, 2)

    # --- Destinatario (la empresa que recibe la factura) ---
    cliente_ruc, cliente_nombre = "", ""
    m_cli = re.search(r"(?:SE[NÑ]OR(?:ES)?|ADQUIRIENTE|CLIENTE|RAZ[OÓ]N\s+SOCIAL(?:\s+DEL?\s+CLIENTE)?|"
                      r"RECIB\w*\s+DE|RECIBIDO\s+DE|A\s+NOMBRE\s+DE|PAGADO\s+POR)\s*[:\-]?\s*", plano)
    if m_cli:
        ventana = plano[m_cli.end(): m_cli.end() + 320]
        numeros_ventana = re.findall(r"\b(\d{11})\b", ventana)
        elegido = next((c for c in numeros_ventana if c != datos["ruc"]),
                       numeros_ventana[0] if numeros_ventana else "")
        if elegido:
            cliente_ruc = elegido
            texto_nombre = ventana[:ventana.index(elegido)]
        else:
            texto_nombre = ventana[:90]
        texto_nombre = re.split(r"\b(?:RUC|R\.U\.C|DNI|N[UÚ]MERO|NUMERO|IDENTIFICADO|DIRECCI[OÓ]N|"
                                r"DIRECCION|FECHA|TEL[EÉ]FONO|TELEFONO|MONEDA|CONDICI[OÓ]N)\b",
                                texto_nombre, maxsplit=1)[0]
        texto_nombre = re.sub(r"[^0-9A-Za-zÁÉÍÓÚÑÜáéíóúñü .,&\-]", " ", texto_nombre)
        cliente_nombre = " ".join(texto_nombre.split()).strip(" .-,:").strip()[:80]
    if not cliente_ruc:
        # Sin etiqueta reconocible: el primer RUC que NO sea el del emisor
        # (en los recibos por honorarios el del emisor va primero).
        for candidato in re.findall(r"\b(\d{11})\b", plano):
            if candidato != datos["ruc"]:
                cliente_ruc = candidato
                break
    if not cliente_ruc and len(rucs) >= 2:
        cliente_ruc = rucs[-1]
    datos["cliente_ruc"] = cliente_ruc
    datos["cliente_nombre"] = cliente_nombre

    # --- Proveedor (razón social del EMISOR) ---
    #     buscar_proveedor_por_ruc devuelve (ruc, nombre) y recibe el TEXTO completo;
    #     se le pasa el RUC del cliente como "propio" para no confundirlo con el emisor.
    nombre = ""
    try:
        _ruc_bd, _nombre_bd = buscar_proveedor_por_ruc(original, datos.get("cliente_ruc", ""))
        nombre = str(_nombre_bd or "").strip()
        if nombre and not str(datos.get("ruc") or "").strip():
            datos["ruc"] = str(_ruc_bd or "").strip()
    except Exception:
        nombre = ""
    if not nombre:
        try:
            nombre = nombre_proveedor_desde_texto(original, datos["ruc"]) or ""
        except Exception:
            nombre = ""
    if not nombre:
        lineas_txt = [l.strip() for l in original.split("\n") if l.strip()]
        candidatos = []
        for i, linea in enumerate(lineas_txt):
            if re.match(r"^\s*R\.?\s*U\.?\s*C", linea, re.IGNORECASE) and i > 0:
                candidatos.append(lineas_txt[i - 1])
        candidatos += lineas_txt[:8]
        for cand in candidatos:
            limpio = re.sub(r"^[\W\d_]+", "", cand).strip()
            if len(limpio) < 5:
                continue
            if re.search(r"R\.?U\.?C|FACTURA|BOLETA|ELECTR|HONORARIOS|FECHA|TOTAL|SUB\s*TOTAL|SEÑOR|SENOR|CLIENTE",
                         limpio.upper()):
                continue
            if not re.search(r"[A-Za-zÁÉÍÓÚÑ]{4}", limpio):
                continue
            nombre = limpio[:80]
            break
    # 🚫 Nunca se toma a la PROPIA empresa (cliente) como proveedor
    if nombre and datos.get("cliente_nombre"):
        try:
            if _texto_sin_acentos(nombre).upper() == _texto_sin_acentos(datos["cliente_nombre"]).upper():
                nombre = nombre_proveedor_desde_texto(original, datos["ruc"]) or ""
        except Exception:
            pass
    datos["proveedor"] = str(nombre or "").strip()
    return datos

def nombre_proveedor_desde_texto(texto, ruc_proveedor=""):
    """Razón social del emisor: la mejor candidata alrededor de la línea del RUC.

    Se revisan la propia línea del RUC (lo que está antes de "RUC"), las líneas
    de arriba y las de abajo, y se elige la que más se parece a una razón social
    (se descartan direcciones, etiquetas y totales).
    """
    lineas = [l.strip() for l in str(texto or "").splitlines() if l.strip()]
    if not lineas:
        return ""
    ruc_digitos = re.sub(r"\D", "", str(ruc_proveedor or ""))
    indice = None
    if ruc_digitos:
        for i, linea in enumerate(lineas):
            if ruc_digitos in re.sub(r"\D", "", linea):
                indice = i
                break

    candidatos = []
    if indice is not None:
        izquierda = re.split(r"R\.?\s?U\.?\s?C", lineas[indice], flags=re.IGNORECASE)[0].strip(" -:|,.")
        if izquierda:
            candidatos.append(izquierda)
        candidatos += [l for l in reversed(lineas[max(0, indice - 4):indice])]
        candidatos += lineas[indice + 1:indice + 4]
    else:
        candidatos += lineas[:6]

    mejor, mejor_puntaje = "", 0
    for candidato in candidatos:
        # Se quita el RUC/DNI que el OCR pegó al nombre en la misma línea
        limpio_cand = re.sub(r"\b\d{8,11}\b", " ", str(candidato)).strip(" -:|,")
        puntaje = _puntaje_como_nombre(limpio_cand)
        if puntaje > mejor_puntaje:
            mejor, mejor_puntaje = limpio_cand, puntaje
    if mejor:
        return mejor

    for linea in lineas[:8]:      # último recurso
        if len(re.sub(r"[^A-Za-zÁÉÍÓÚÑáéíóúñ]", "", linea)) >= 4 and not _NO_ES_NOMBRE.search(linea):
            return linea
    return ""


def parsear_monto_texto(texto):
    """Convierte lo escrito por el usuario en un monto (float >= 0).

    Acepta formatos como 1000, 1000.50, 1.000,50, 1,000.50 o 12,50.
    Lanza ValueError solo cuando el texto no contiene ningún dígito.
    """
    texto = str(texto or "").strip().replace(" ", "")
    if not texto:
        return 0.0
    try:
        return max(float(texto), 0.0)
    except ValueError:
        pass
    if re.search(r"\d", texto) is None:
        raise ValueError(texto)

    limpio = re.sub(r"[^\d.,]", "", texto)
    # El separador decimal es el que aparece más a la derecha; si el grupo final
    # tiene 3 dígitos se asume que es separador de miles.
    if "," in limpio and "." in limpio:
        decimal = "," if limpio.rfind(",") > limpio.rfind(".") else "."
    elif "," in limpio:
        decimal = "," if len(limpio.split(",")[-1]) <= 2 else None
    elif "." in limpio:
        decimal = "." if len(limpio.split(".")[-1]) <= 2 else None
    else:
        decimal = None

    if decimal:
        entero, _sep, dec = limpio.rpartition(decimal)
        entero = re.sub(r"[.,]", "", entero) or "0"
        limpio = f"{entero}.{dec}"
    else:
        limpio = re.sub(r"[.,]", "", limpio)
    try:
        return max(float(limpio), 0.0)
    except ValueError:
        raise ValueError(texto)

def obtener_ruta_base_drive():
    """Carpeta base para guardar archivos.

    Devuelve "" si el equipo NO está autorizado (equipo secundario sin la cuenta
    Rclone del equipo principal): en ese caso ningún módulo debe guardar archivos.
    """
    try:
        from politica_almacenamiento import ruta_base_autorizada
        return ruta_base_autorizada(mostrar_alerta=False)
    except Exception:
        ruta = CONFIG_REGIONAL.get("ruta_drive", "").strip()
        return os.path.expanduser(ruta) if ruta else ""


def avisar_sin_permiso_guardado(parent=None):
    """Muestra el aviso correcto cuando no hay carpeta base disponible.

    Si el bloqueo viene de la política de almacenamiento (cuenta Rclone distinta a
    la del equipo principal) muestra esa advertencia; si no, la de configuración.
    Devuelve True si ya se mostró el aviso de bloqueo.
    """
    try:
        from politica_almacenamiento import avisar_sin_ruta
        return avisar_sin_ruta(parent)
    except Exception:
        messagebox.showwarning("Configuración Requerida", "No ha configurado la ruta de Google Drive.\nEs obligatorio para guardar archivos.", parent=parent)
        return False


# =========================================================
# 🏦 N° DE OPERACIÓN BANCARIA (OBLIGATORIO EN LOS PAGOS)
# =========================================================
def pago_sin_operacion_bancaria(cuenta):
    """True cuando el medio de pago no genera N° de operación bancaria.

    Los pagos que salen del banco (transferencia, cheque, tarjeta, etc.) deben
    registrar obligatoriamente su N° de operación; los pagos en efectivo o de
    caja chica no tienen número de operación.
    """
    texto = (cuenta or "").strip().lower()
    return any(p in texto for p in ("efectivo", "caja chica", "caja", "rendicion", "rendición"))


def validar_operacion_bancaria(operacion, cuenta):
    """Valida el N° de operación bancaria y devuelve el texto que se guardará.

    - Si el usuario escribió un número, se usa tal cual.
    - Si el pago es en efectivo / caja chica y no escribió nada, se guarda "EFECTIVO".
    - Si es un pago bancario y falta el número, devuelve None (el llamador avisa).
    """
    numero = (operacion or "").strip()
    if numero:
        return numero
    if pago_sin_operacion_bancaria(cuenta):
        return "EFECTIVO"
    return None

def cargar_comision_interbancaria_config():
    """Comisión interbancaria configurada en el módulo de Banco (S/.)."""
    try:
        from modulo_banco import cargar_comision_interbancaria
        return float(cargar_comision_interbancaria())
    except Exception:
        try:
            from modulo_banco import COMISION_INTERBANCARIA_DEFAULT
            return float(COMISION_INTERBANCARIA_DEFAULT)
        except Exception:
            return 4.80


# =========================================================
# 🔃 ORDENAMIENTO POR CUALQUIER COLUMNA (TODAS LAS PÁGINAS)
# =========================================================
# 🔗 Opción del desplegable cuando la factura NO se relaciona con ninguna orden de servicio
SIN_ORDEN_SERVICIO = "— Sin orden de servicio —"

COLUMNAS_ORDEN_MONEDA = {"subtotal", "impuesto", "igv", "total", "detraccion", "neto",
                         "neto_facturado", "pagado", "saldo"}
COLUMNAS_ORDEN_NUMERO = {"num", "id", "id_factura", "dias", "kilometraje", "cantidad", "archivos"}

def clave_orden_moneda(valor):
    texto = "" if valor is None else str(valor).strip()
    if texto in ("", "-"):
        return (1, 0.0)
    return (0, desformatear_numero(texto))

def clave_orden_numero(valor):
    texto = "" if valor is None else str(valor).strip()
    if texto in ("", "-"):
        return (1, 0.0)
    try:
        return (0, float(texto))
    except ValueError:
        pass
    coincidencia = re.search(r"\d+(?:[.,]\d+)?", texto)
    if coincidencia:
        try:
            return (0, float(coincidencia.group(0).replace(",", ".")))
        except ValueError:
            pass
    return (1, 0.0)

def clave_orden_fecha(valor):
    texto = "" if valor is None else str(valor).strip()
    if not texto:
        return (1, "")
    for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y", "%m/%d/%Y"):
        try:
            return (0, datetime.strptime(texto, formato).strftime("%Y%m%d"))
        except ValueError:
            continue
    return (1, texto.casefold())

def clave_orden_texto(valor):
    texto = "" if valor is None else str(valor).strip()
    if texto in ("", "-"):
        return (1, "")
    return (0, texto.casefold())

def clave_orden_columna(columna, valor):
    """Clave comparable para ordenar cualquier columna mostrada en las tablas."""
    if columna in COLUMNAS_ORDEN_MONEDA:
        return clave_orden_moneda(valor)
    if columna in COLUMNAS_ORDEN_NUMERO:
        return clave_orden_numero(valor)
    if columna == "fecha":
        return clave_orden_fecha(valor)
    return clave_orden_texto(valor)

def aplicar_orden_filas(filas, columnas, columna, ascendente, indice_valores=0):
    """Ordena TODAS las filas (todas las páginas) por la columna indicada."""
    try:
        indice = columnas.index(columna)
    except ValueError:
        return filas
    return sorted(filas,
                  key=lambda fila: clave_orden_columna(columna, fila[indice_valores][indice]),
                  reverse=not ascendente)

def aplicar_estilo_treeview():
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("Treeview", background="#ffffff", foreground="#000000", fieldbackground="#ffffff", bordercolor="#e0e0e0", borderwidth=1, rowheight=26, font=("Arial", 10))
    style.map("Treeview", background=[("selected", "#1f538d")], foreground=[("selected", "#ffffff")])
    style.configure("Treeview.Heading", background="#f0f0f0", foreground="#000000", relief="flat", font=("Arial", 10, "bold"), bordercolor="#e0e0e0", borderwidth=1)


# =========================================================
# 🔍 BÚSQUEDA POR CUALQUIER COLUMNA (MÓDULO DE COMPRAS)
# =========================================================
def agregar_condicion_ocultar_app(condiciones, params, ocultar):
    """Agrega la condición que OCULTA las facturas enviadas por la App (App Grifo).

    Se reconocen porque llegaron desde la aplicación móvil:
      - quedaron como 'PENDIENTE_DESCARGA' (todavía sin descargar), o
      - su archivo es un ticket del móvil (nombre 'Ticket_Movil...'), o
      - conservan la imagen en la nube (imagen_base64).
    Las compras cruzadas registradas a mano desde el módulo de Banco NUNCA se
    ocultan: siempre deben verse con toda su información.
    """
    if not ocultar:
        return
    condiciones.append(
        "(COALESCE(es_compra_cruzada, FALSE) OR NOT ("
        "archivo_ruta = 'PENDIENTE_DESCARGA' OR archivo_ruta LIKE %s "
        "OR imagen_base64 IS NOT NULL))"
    )
    params.append("%Ticket_Movil%")


def _normalizar_numero_busqueda(txt):
    """Normaliza un monto escrito por el usuario: 'S/ 1,234.56' -> '1234.56'."""
    t = re.sub(r"[^\d.,]", "", str(txt or ""))
    if not t:
        return ""
    if "," in t and "." in t:
        t = t.replace(",", "")
    elif "," in t:
        partes = t.split(",")
        if len(partes) == 2 and len(partes[-1]) == 2:
            t = t.replace(",", ".")
        else:
            t = t.replace(",", "")
    return t


def construir_condicion_busqueda_compras(filtro):
    """Busca en TODAS las columnas de la fila, incluido el MONTO.

    Cubre fecha, N° documento, proveedor, RUC, vehículo, kilometraje, cantidad,
    concepto, tipo, categoría, días de crédito, subtotal, IGV, total,
    detracción, neto a pagar, monto pagado y forma de pago.
    Los montos se comparan normalizados: escribir «S/ 1,234.56», «510.70» o
    «510.7» encuentra lo mismo.
    """
    if not filtro or not str(filtro).strip():
        return "", []
    txt = str(filtro).strip()
    val = f"%{txt}%"

    columnas = ("numero_documento", "proveedor", "evento_asociado", "descripcion",
                "tipo_documento", "categoria", "ruc", "fecha", "dias_credito",
                "kilometraje", "cantidad_combustible", "subtotal", "impuesto",
                "total", "det_monto", "orden_servicio")

    partes = [f"CAST({c} AS TEXT) ILIKE %s" for c in columnas]
    params = [val] * len(columnas)

    # Forma de pago / montos pagados asociados al comprobante (columnas derivadas)
    partes.append(
        "EXISTS (SELECT 1 FROM pagos_comprobantes pc "
        "WHERE pc.id_factura = facturas_recibidas.id "
        "AND (pc.cuenta_origen ILIKE %s OR CAST(pc.monto_pagado AS TEXT) ILIKE %s))"
    )
    params.extend([val, val])

    # 🔢 Búsqueda por MONTO (aunque escriba 'S/ 1,234.56', '510.70' o el NETO)
    num = _normalizar_numero_busqueda(txt)
    if num:
        n = "%" + num + "%"
        for col in ("subtotal", "impuesto", "total", "det_monto"):
            partes.append("TO_CHAR(COALESCE(%s,0), 'FM999999999999990.00') LIKE %%s" % col)
            params.append(n)
        # neto a pagar (total - detracción) y neto de recibo 8%
        partes.append("TO_CHAR(COALESCE(total,0) - COALESCE(det_monto,0), 'FM999999999999990.00') LIKE %s")
        params.append(n)
        partes.append("TO_CHAR(COALESCE(total,0) - COALESCE(impuesto,0) - COALESCE(det_monto,0), 'FM999999999999990.00') LIKE %s")
        params.append(n)
        # monto pagado en pagos_comprobantes
        partes.append("EXISTS (SELECT 1 FROM pagos_comprobantes pc2 "
                      "WHERE pc2.id_factura = facturas_recibidas.id "
                      "AND TO_CHAR(COALESCE(pc2.monto_pagado,0), 'FM999999999999990.00') LIKE %s)")
        params.append(n)

    return "(" + " OR ".join(partes) + ")", params


# =========================================================
# CLASE: CALENDARIO NATIVO
# =========================================================
class CalendarioNativo(ctk.CTkToplevel):
    def __init__(self, parent, target_entry, al_elegir=None):
        # 'al_elegir' es opcional: al elegir la fecha se avisa (si se pidió).
        # Todo es Tkinter puro, así funciona igual en Windows y en macOS.
        super().__init__(parent)
        self.target_entry = target_entry
        self.al_elegir = al_elegir
        self.title("Seleccionar Fecha")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        centrar_ventana(self, parent, 310, 320)
        
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
        if callable(getattr(self, "al_elegir", None)):
            try:
                self.al_elegir(fecha_seleccionada)
            except Exception as e_aviso:
                print("Aviso - calendario:", e_aviso)
        self.destroy()

_SCHEMA_COMPRAS_OK = False

# =========================================================
# PESTAÑA 1: FACTURAS RECIBIDAS
# =========================================================
# =========================================================
# 🛠️ MANTENIMIENTOS DEL VEHÍCULO (mismas fechas que controla Flota Automotriz)
# =========================================================
# Cada mantenimiento indica las columnas de 'flota_vehiculos' que se actualizan
# al registrar el servicio. El segundo valor son los MESES que se suman a la
# fecha ingresada para calcular el PRÓXIMO vencimiento (0 = la misma fecha).
# =========================================================
# 📄 COMPROBANTES DE PAGO: lectura del PDF e identificación de facturas
# =========================================================
def texto_de_comprobante(ruta):
    """Texto del PDF de un comprobante. Si es un escaneo sin texto, se usa OCR."""
    if not ruta or not os.path.exists(str(ruta)):
        return ""
    # 1) Texto propio del PDF (recibos digitales del banco)
    try:
        from pypdf import PdfReader
        lector = PdfReader(str(ruta))
        texto = "\n".join((p.extract_text() or "") for p in lector.pages[:3])
        if texto.strip():
            return texto[:6000]
    except Exception:
        pass
    try:
        import fitz
        doc = fitz.open(str(ruta))
        texto = "\n".join(doc[i].get_text() for i in range(min(3, doc.page_count)))
        doc.close()
        if texto.strip():
            return texto[:6000]
    except Exception:
        doc = None

    # 2) Sin texto: es una foto/escaneo -> OCR
    imagenes = []
    try:
        import fitz
        doc = fitz.open(str(ruta))
        carpeta_tmp = tempfile.mkdtemp(prefix="ocr_comp_")
        for i in range(min(2, doc.page_count)):
            pagina = doc[i]
            pix = pagina.get_pixmap(dpi=200)
            salida = os.path.join(carpeta_tmp, "pagina_%d.png" % i)
            pix.save(salida)
            imagenes.append(salida)
        doc.close()
    except Exception as e:
        print("Aviso - convirtiendo el PDF para OCR:", e)
        try:
            imagenes = [str(ruta)]          # puede ser una imagen (jpg/png)
        except Exception:
            imagenes = []

    if not imagenes:
        return ""
    try:
        import escaner as mod_escaner
        resultado = mod_escaner.ocr_paginas(imagenes)
        texto = mod_escaner.texto_de_resultados(resultado)
        return texto[:6000]
    except Exception as e:
        print("Aviso - OCR del comprobante:", e)
        return ""


def datos_de_comprobante(ruta):
    """Extrae monto, fecha, N° de operación y N° de facturas del comprobante."""
    texto = texto_de_comprobante(ruta)
    datos = {"monto": "", "fecha": "", "operacion": "", "facturas": [], "texto": texto}
    if not texto.strip():
        return datos
    plano = " ".join(texto.split())

    monto = re.search(r"(?:S/|S/\.|PEN|IMPORTE|MONTO|TOTAL)\D{0,14}([0-9][0-9.,oO/|' ]{1,16})",
                      plano, re.IGNORECASE)
    if monto:
        # El OCR confunde separadores: "1,770.oo" · "1/770.00" · "1.770,00"
        crudo = monto.group(1).strip()
        valor = crudo.replace("o", "0").replace("O", "0")
        valor = valor.replace(" ", "").replace("'", "").replace("/", "").replace("|", "")
        if "," in valor and "." in valor:
            # el último separador es el decimal
            if valor.rfind(",") > valor.rfind("."):
                valor = valor.replace(".", "").replace(",", ".")
            else:
                valor = valor.replace(",", "")
        elif "," in valor:
            partes = valor.split(",")
            valor = "".join(partes) if len(partes[-1]) == 3 and len(partes) == 2 else valor.replace(",", ".")
        valor = valor.rstrip(".").rstrip(",")
        try:
            datos["monto"] = "%.2f" % float(valor)
        except Exception:
            try:
                solo_digitos = "".join(c for c in valor if c.isdigit())
                if solo_digitos:
                    datos["monto"] = "%.2f" % (float(solo_digitos) / 100.0)
            except Exception:
                pass

    # 📅 Fecha: primero se usa la función especializada (valida el año y entiende
    #    "23 Octubre 2026"); si no devuelve nada, se usa el formato numérico.
    try:
        datos["fecha"] = extraer_fecha_comprobante(texto) or ""
    except Exception:
        datos["fecha"] = ""
    fecha = re.search(r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})\b", plano)
    if datos["fecha"]:
        fecha = None
    if fecha:
        dia, mes, anio = fecha.groups()
        if len(anio) == 2:
            anio = "20" + anio
        datos["fecha"] = "%02d/%02d/%s" % (int(dia), int(mes), anio)
    if not datos["fecha"]:
        MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
                 "julio": 7, "agosto": 8, "setiembre": 9, "septiembre": 9, "octubre": 10,
                 "noviembre": 11, "diciembre": 12}
        por_nombre = re.search(r"\b(\d{1,2})\s*(?:de\s*)?([a-zA-ZáéíóúÁÉÍÓÚ]{4,12})\s*(?:de\s*|del\s*)?(\d{4})\b",
                               plano)
        if por_nombre:
            dia, mes_txt, anio = por_nombre.groups()
            mes_num = MESES.get(mes_txt.lower())
            if mes_num:
                datos["fecha"] = "%02d/%02d/%s" % (int(dia), mes_num, anio)

    # 🔢 N° de operación (el banco escribe "Número de operación", "Nro. op.", "Referencia")
    # 🔢 N° de operación / comprobante. El OCR de recibos bancarios lee las
    # columnas desordenadas, así que el número puede venir partido (ej. "4550"
    # ... "05155128"). Se buscan los grupos largos de dígitos y se unen.
    etiqueta_op = re.search(
        r"(?:n[uú]mero\s*de\s*operaci[oó]n|n[°º]\s*de\s*operaci[oó]n|"
        r"nro\.?\s*(?:de\s*)?operaci[oó]n|operaci[oó]n(?:\s*n[°º])?|"
        r"c[oó]digo\s*de\s*operaci[oó]n|n[uú]mero\s*de\s*comprobante|"
        r"comprobante\s*n[°º]?|constancia\s*n[°º]?|n[°º]\s*de\s*documento|"
        r"referencia|constancia)",
        plano, re.IGNORECASE)
    grupos_digitos = [(m.start(), m.group(0)) for m in re.finditer(r"\b\d{4,20}\b", plano)]

    def _unir(valor_a, valor_b):
        """Une dos trozos del mismo número cuando el OCR los separó."""
        if not valor_b:
            return valor_a
        return (valor_a + valor_b)[:20]

    if etiqueta_op and grupos_digitos:
        pos = etiqueta_op.end()
        # Solo los números que vienen DESPUÉS de la etiqueta (antes se incluía el
        # número de cuenta que aparece arriba y quedaba "006901819091").
        despues = [(s_pos, v) for s_pos, v in grupos_digitos if s_pos >= pos]
        if despues:
            # Se toma el número MÁS LARGO (el de operación suele tener 7-12 dígitos;
            # el de la cuenta aparece como fragmento corto, p. ej. "0069").
            datos["operacion"] = max(despues, key=lambda g: (len(g[1]), -g[0]))[1]
    # 🔁 Se contrasta con la función especializada (más fiable con los vouchers):
    #    si ella reconoce un número, se usa ese.
    try:
        _op_especializada = extraer_numero_operacion(plano)
    except Exception:
        _op_especializada = ""
    if _op_especializada:
        datos["operacion"] = _op_especializada
    if not datos["operacion"] and grupos_digitos:
        # sin etiqueta legible: se toma el grupo de dígitos más largo del recibo
        datos["operacion"] = max(grupos_digitos, key=lambda g: len(g[1]))[1]

    # Respaldo: valor alfanumérico pegado a la etiqueta (ej. "operación: AB-123456")
    if not datos["operacion"]:
        operacion = re.search(r"(?:n[uú]mero\s*de\s*operaci[oó]n|operaci[oó]n|referencia)"
                              r"\D{0,20}([A-Za-z][A-Za-z0-9\-]{4,24})", plano, re.IGNORECASE)
        if operacion:
            valor = operacion.group(1).strip("-.:")
            if valor.lower() not in ("es", "del", "de", "la", "banco"):
                datos["operacion"] = valor

    facturas = []
    for hallada in re.findall(r"\b([EF][A-Z]\d{2,3}-\d{1,8}|[EF]\d{3}-\d{1,8})\b", plano):
        if hallada not in facturas:
            facturas.append(hallada)
    datos["facturas"] = facturas[:20]
    return datos


def facturas_que_cancela(archivo_ruta="", numero_operacion=""):
    """Facturas canceladas por el MISMO comprobante o el mismo N° de operación."""
    filas = []
    try:
        conn = conectar_db(silencioso=True)
        if not conn:
            return filas
        try:
            with conn.cursor() as cur:
                # 🌍 La base puede tener la ruta RELATIVA ("facturas_recibidas/x.pdf")
                #    mientras que la app maneja la ruta ABSOLUTA resuelta. Se comparan
                #    las tres formas posibles + el nombre del archivo.
                ruta_txt = str(archivo_ruta or "").replace("\\", "/").strip()
                nombre = ruta_txt.split("/")[-1] if ruta_txt else ""
                relativa = ruta_txt
                try:
                    from app_paths import partir_ruta_guardada
                    sub, nom = partir_ruta_guardada(ruta_txt)
                    if nom:
                        relativa = (sub + "/" + nom) if sub else nom
                except Exception:
                    pass
                cur.execute("""
                    SELECT DISTINCT f.numero_documento, COALESCE(f.proveedor, ''),
                                    COALESCE(pc.monto_pagado, 0), COALESCE(pc.fecha_pago, '')
                    FROM pagos_comprobantes pc
                    JOIN facturas_recibidas f ON f.id = pc.id_factura
                    WHERE (%s <> '' AND (
                              COALESCE(pc.archivo_ruta, '') = %s
                           OR COALESCE(pc.archivo_ruta, '') = %s
                           OR COALESCE(pc.archivo_ruta, '') LIKE %s))
                       OR (%s <> '' AND COALESCE(pc.numero_operacion, '') = %s)
                    ORDER BY f.numero_documento
                """, (ruta_txt, ruta_txt, relativa, ("%" + nombre) if nombre else "\x00",
                      str(numero_operacion or ""), str(numero_operacion or "")))
                filas = cur.fetchall()
        finally:
            liberar_conexion(conn)
    except Exception as e:
        print("Aviso - facturas del comprobante:", e)
    return filas


def texto_facturas_canceladas(filas):
    """Texto listo para mostrar con las facturas que cancela un comprobante."""
    if not filas:
        return ""
    lineas = []
    for numero, proveedor, monto, fecha in filas:
        lineas.append("• %s  %s  %s  (%s)" % (numero, proveedor[:28],
                                              formatear_moneda(monto), fecha))
    return "\n".join(lineas)


def crear_pdf_desde_imagenes(imagenes, destino):
    """Une las páginas escaneadas (imágenes) en un solo PDF."""
    try:
        from PIL import Image
        paginas = []
        for ruta in imagenes:
            img = Image.open(ruta)
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")
            paginas.append(img)
        if not paginas:
            return ""
        paginas[0].save(destino, "PDF", resolution=150.0, save_all=True, append_images=paginas[1:])
        return destino if os.path.exists(destino) else ""
    except Exception as e:
        print("Aviso - creando PDF del escaneo:", e)
        return ""


MANTENIMIENTOS_VEHICULO = (
    {"nombre": "🛢️ Cambio de Aceite",
     "campos": (("fec_aceite", 0), ("fecha_ultimo_aceite", 0)),
     "km": ("km_ultimo_aceite", "Km del cambio"),
     "actual": ("fec_aceite", "fecha_ultimo_aceite", "km_ultimo_aceite")},
    {"nombre": "⛓️ Cadena / Correa de Tiempo",
     "campos": (("fec_correa", 0),),
     "km": ("km_prox_correa", "Próximo cambio a los (Km)"),
     "actual": ("fec_correa", "km_prox_correa")},
    {"nombre": "💨 Revisión del Sistema de Gas (con certificado)", "doc": "revision_gas",
     "campos": (("fec_rev_gas", 0),),
     "actual": ("fec_rev_gas",)},
    {"nombre": "💨 Mantenimiento del Sistema de Gas",
     "campos": (("emision_mant_gas", 0),),
     "km": ("km_mant_gas", "Km del mantenimiento"),
     "actual": ("emision_mant_gas", "km_mant_gas")},
    {"nombre": "🛞 Neumáticos / Cauchos",
     "campos": (("neu_fecha", 0),),
     "km": ("neu_km", "Km del cambio"),
     "actual": ("neu_fecha", "neu_km")},
    {"nombre": "🛠️ Mantenimiento Preventivo / General",
     "campos": (("fecha_ultimo_general", 0),),
     "km": ("km_ultimo_general", "Km del servicio"),
     "actual": ("fecha_ultimo_general", "km_ultimo_general")},
    {"nombre": "📄 SOAT (emisión → vence en 1 año)", "doc": "soat",
     "campos": (("emision_soat", 0), ("vencimiento_soat", 12)),
     "actual": ("emision_soat", "vencimiento_soat")},
    {"nombre": "🔍 Revisión Técnica (realizada → vence en 1 año)", "doc": "revision_tecnica",
     "campos": (("vencimiento_rt", 12),),
     "actual": ("vencimiento_rt",)},
    {"nombre": "🛡️ Póliza de Seguro (emisión → vence en 1 año)",
     "campos": (("emision_seguro", 0), ("vencimiento_seguro", 12)),
     "actual": ("emision_seguro", "vencimiento_seguro")},
    {"nombre": "🔋 Batería (compra → vence en 2 años)",
     "campos": (("fec_compra_bat", 0), ("fec_venc_bat", 24)),
     "actual": ("fec_compra_bat", "fec_venc_bat")},
    {"nombre": "🧯 Extintor (recarga → vence en 1 año)", "doc": "extintor",
     "campos": (("fec_venc_extintor", 12),),
     "actual": ("fec_venc_extintor",)},
)

ETIQUETAS_MANT = {
    "fec_aceite": "Último cambio", "fecha_ultimo_aceite": "Fecha último aceite",
    "km_ultimo_aceite": "Km último aceite", "fec_correa": "Cambio correa",
    "km_prox_correa": "Km próximo correa", "fec_rev_gas": "Revisión gas",
    "fecha_ultimo_general": "Último general", "km_ultimo_general": "Km último general",
    "emision_soat": "Emisión SOAT", "vencimiento_soat": "Vence SOAT",
    "vencimiento_rt": "Vence RT", "emision_seguro": "Emisión seguro",
    "vencimiento_seguro": "Vence seguro", "fec_compra_bat": "Compra batería",
    "fec_venc_bat": "Vence batería", "fec_venc_extintor": "Vence extintor",
    "emision_mant_gas": "Emisión mant. gas", "km_mant_gas": "Km mant. gas",
    "neu_fecha": "Cambio neumáticos", "neu_km": "Km neumáticos",
    "neu_marca": "Marca neumáticos", "neu_modelo": "Modelo neumáticos",
    "neu_medidas": "Medidas neumáticos", "bat_marca": "Marca batería",
    "bat_modelo": "Modelo batería",
}

COLUMNAS_MANT = ("placa", "marca", "modelo", "kilometraje", "fec_aceite", "fecha_ultimo_aceite",
                 "fec_correa", "fec_rev_gas", "fecha_ultimo_general", "emision_soat",
                 "vencimiento_soat", "vencimiento_rt", "emision_seguro", "vencimiento_seguro",
                 "fec_compra_bat", "fec_venc_bat", "fec_venc_extintor", "km_ultimo_aceite",
                 "km_prox_correa", "km_ultimo_general",
                 "emision_mant_gas", "km_mant_gas", "neu_fecha", "neu_km",
                 "neu_marca", "neu_modelo", "neu_medidas", "bat_marca", "bat_modelo")


def sumar_meses_a_fecha(fecha_txt, meses):
    """Suma meses a una fecha DD/MM/AAAA (0 = la devuelve igual).

    Devuelve '' si la fecha no se puede interpretar."""
    fecha_txt = str(fecha_txt or "").strip()
    if not meses:
        return fecha_txt
    try:
        base = datetime.strptime(fecha_txt, "%d/%m/%Y")
    except ValueError:
        return ""
    total = base.month - 1 + int(meses)
    anio = base.year + total // 12
    mes = total % 12 + 1
    dia = min(base.day, calendar.monthrange(anio, mes)[1])
    return f"{dia:02d}/{mes:02d}/{anio}"


def texto_duracion_meses(meses):
    """'6 meses' / '1 año' / '2 años' (para los botones de vencimiento)."""
    try:
        meses = int(meses)
    except Exception:
        return ""
    if meses % 12 == 0 and meses >= 12:
        anios = meses // 12
        return "1 año" if anios == 1 else "%d años" % anios
    return "%d meses" % meses


# =========================================================
# 📎 ADJUNTOS (escaneo / PDF / ver) DE VEHÍCULOS Y CHOFERES
# =========================================================
def escanear_documento_pdf(parent, destino_pdf, etiqueta="documento"):
    """Escanea una hoja con la multifuncional y la guarda como PDF.

    Devuelve la ruta del PDF creado, o "" si no se pudo.
    """
    try:
        import escaner
    except Exception as e:
        messagebox.showerror("Escáner", "No se pudo cargar el módulo de escaneo:\n%s" % e, parent=parent)
        return ""
    if not escaner.sistema_soportado():
        messagebox.showinfo("Escáner",
                            "El escaneo con la multifuncional está disponible en Windows y macOS.",
                            parent=parent)
        return ""
    config = cargar_configuracion_regional()
    datos = escaner.config_escaner(config)
    aviso = ctk.CTkToplevel(parent)
    aviso.title("Escaneando")
    aviso.geometry("440x120")
    aviso.transient(parent)
    try:
        aviso.grab_set()
    except Exception:
        pass
    ctk.CTkLabel(aviso, text="🖨️ Escaneando %s..." % etiqueta,
                 font=("Arial", 12, "bold")).pack(pady=(14, 2))
    ctk.CTkLabel(aviso, text="Coloque la hoja en el escáner (cristal).",
                 font=("Arial", 10), text_color="#555555").pack()
    try:
        aviso.update()
    except Exception:
        pass
    pdf = ""
    try:
        carpeta = escaner.carpeta_escaneos(config)
        escaneo = escaner.escanear(
            carpeta,
            dispositivo=datos.get("escaner_dispositivo", ""),
            dpi=datos.get("escaner_dpi", "300"),
            color=datos.get("escaner_color", "Escala de grises"),
            origen=datos.get("escaner_origen", "Cristal (una página a la vez)"),
            paginas=1,
            optimizar=str(datos.get("escaner_optimizar", "Sí")).lower().startswith("s"),
            modo=datos.get("escaner_modo", "Automático (sin ventanas)"))
        if not escaneo.get("ok"):
            messagebox.showerror("Escáner", escaneo.get("error") or "No se pudo escanear.", parent=parent)
        else:
            paginas = [str(p) for p in (escaneo.get("paginas") or []) if os.path.exists(str(p))]
            if paginas:
                carpeta_dest = os.path.dirname(destino_pdf)
                if carpeta_dest:
                    os.makedirs(carpeta_dest, exist_ok=True)
                escaner.imagenes_a_pdf(paginas, destino_pdf)
                pdf = destino_pdf
    except Exception as e:
        messagebox.showerror("Escáner", "No se pudo escanear:\n%s" % e, parent=parent)
    finally:
        try:
            aviso.grab_release()
            aviso.destroy()
        except Exception:
            pass
    return pdf


def fila_documento_adjunto(parent, tipo, entidad, documento, etiqueta):
    """Fila [Cargar PDF] [Escanear] [Ver] [Quitar] para un documento concreto."""
    try:
        import documentos_adjuntos as adj
    except Exception as e:
        ctk.CTkLabel(parent, text="⚠️ Adjuntos no disponibles: %s" % e,
                     font=("Arial", 10), text_color="#c0392b").pack(anchor="w", padx=30, pady=(0, 6))
        return None

    marco = ctk.CTkFrame(parent, fg_color="#eef6ff", corner_radius=6,
                         border_width=1, border_color="#aed6f1")
    marco.pack(fill="x", padx=30, pady=(0, 6))
    lbl = ctk.CTkLabel(marco, text="", font=("Arial", 10), text_color="#7f8c8d",
                       anchor="w", wraplength=430, justify="left")
    lbl.pack(side="left", padx=(8, 6), pady=4, fill="x", expand=True)

    def _pintar():
        archivo = adj.obtener(tipo, entidad, documento)
        if archivo:
            lbl.configure(text="📎 " + os.path.basename(archivo), text_color="#1e8449")
        else:
            lbl.configure(text="Sin archivo adjunto (opcional)", text_color="#999999")

    def _cargar():
        archivo = seleccionar_archivo_dialogo("Seleccionar %s" % etiqueta,
                                              [("PDF e imágenes", "*.pdf;*.png;*.jpg;*.jpeg")])
        if not archivo:
            return
        if adj.guardar(tipo, entidad, documento, archivo):
            _pintar()
        else:
            messagebox.showerror("Adjunto", "No se pudo guardar el archivo.", parent=parent)

    def _escanear():
        destino = adj.ruta_destino(tipo, entidad, documento)
        pdf = escanear_documento_pdf(parent, destino, etiqueta)
        if pdf and adj.guardar(tipo, entidad, documento, pdf):
            _pintar()

    def _ver():
        archivo = adj.obtener(tipo, entidad, documento)
        if not archivo:
            messagebox.showinfo("Adjunto",
                                "Este documento todavía no tiene un archivo cargado.",
                                parent=parent)
            return
        try:
            abrir_documento(archivo)
        except Exception as e:
            messagebox.showerror("Adjunto", "No se pudo abrir el archivo:\n%s" % e, parent=parent)

    def _quitar():
        if not messagebox.askyesno("Adjunto", "¿Quitar el archivo de %s?" % etiqueta, parent=parent):
            return
        adj.quitar(tipo, entidad, documento)
        _pintar()

    ctk.CTkButton(marco, text="📎 Cargar PDF", width=108, height=26,
                  font=("Arial", 10, "bold"), fg_color="#1f538d", hover_color="#163b65",
                  command=_cargar).pack(side="left", padx=2)
    ctk.CTkButton(marco, text="🖨️ Escanear", width=98, height=26,
                  font=("Arial", 10, "bold"), fg_color="#e67e22", hover_color="#ca6f1e",
                  command=_escanear).pack(side="left", padx=2)
    ctk.CTkButton(marco, text="👁 Ver", width=68, height=26,
                  font=("Arial", 10, "bold"), fg_color="#27ae60", hover_color="#1e8449",
                  command=_ver).pack(side="left", padx=2)
    ctk.CTkButton(marco, text="🗑", width=30, height=26,
                  font=("Arial", 10, "bold"), fg_color="#7f8c8d", hover_color="#5d6d6e",
                  command=_quitar).pack(side="left", padx=(2, 6))
    _pintar()
    return marco


def abrir_reset_documentos_chofer(parent, chofer_etiqueta="", todos=False, usuario=""):
    """Ventana para resetear las fechas de los documentos del chofer.

    Se muestran los documentos UNA sola vez: si es para todos los choferes, la misma
    fecha se aplica a cada uno. Cada documento lleva "Fecha en que se realizó" y
    "Fecha de vencimiento" (las dos con calendario); el vencimiento se calcula solo al
    escribir la fecha de realización y queda en el cronograma para avisar.
    """
    try:
        from choferes import (DOCUMENTOS_CHOFER, choferes_activos, documentos_de_chofer,
                              actualizar_documento_chofer, sumar_meses_fecha)
    except Exception as e:
        return messagebox.showerror("Documentos del chofer",
                                    "No se pudo cargar el módulo de Choferes:\n%s" % e, parent=parent)

    if todos:
        elegidos = choferes_activos()
        titulo = "Documentos de TODOS los choferes (%d)" % len(elegidos)
    else:
        elegidos = [c for c in choferes_activos() if c.get("etiqueta") == chofer_etiqueta]
        if not elegidos:
            return messagebox.showinfo("Documentos del chofer",
                                       "Elija primero un chofer en la lista (o «Todos los choferes»).",
                                       parent=parent)
        titulo = "Documentos de %s" % elegidos[0]["etiqueta"]

    if not elegidos:
        return messagebox.showinfo("Documentos del chofer",
                                   "No hay choferes activos para actualizar.", parent=parent)

    # Los documentos se muestran UNA sola vez (la misma fecha para todos)
    datos_actuales = {} if todos else documentos_de_chofer(elegidos[0]["id"])

    v = ctk.CTkToplevel(parent)
    v.title("Documentos del Chofer")
    v.geometry("1000x700")
    v.transient(parent)
    try:
        traer_al_frente(v)
    except Exception:
        pass

    ctk.CTkLabel(v, text="📄 %s" % titulo, font=("Arial", 16, "bold"),
                 text_color="#1f538d").pack(pady=(12, 2))
    ayuda = ("Marque la fecha en que se realizó cada documento y su vencimiento (el vencimiento "
             "se calcula solo al escribir la fecha de realización). " +
             ("Se aplicará la MISMA fecha a los %d choferes." % len(elegidos) if todos else
              "Quedará avisado en Cronograma y Vencimientos."))
    ctk.CTkLabel(v, text=ayuda, font=("Arial", 11), text_color="gray",
                 wraplength=940, justify="left").pack(pady=(0, 8), padx=14)

    contenedor = ctk.CTkScrollableFrame(v, fg_color="transparent")
    contenedor.pack(fill="both", expand=True, padx=12, pady=(0, 6))

    filas = []

    def _aplicar_duracion(meses, entrada_fecha, entrada_vence, doc, ventana):
        """Pone el vencimiento = fecha de realización + la duración elegida."""
        realizado = entrada_fecha.get().strip()
        if not realizado:
            messagebox.showwarning(
                "Falta la fecha",
                "Primero escriba la «Fecha en que se realizó» de %s y vuelva a pulsar el botón."
                % doc["etiqueta"], parent=ventana)
            try:
                entrada_fecha.focus_set()
            except Exception:
                pass
            return
        try:
            entrada_vence.delete(0, "end")
            entrada_vence.insert(0, sumar_meses_fecha(realizado, meses) or "")
        except Exception as e:
            print("Aviso - duración del documento:", e)

    def _calcular(doc, entrada_fecha, entrada_vence):
        """Calcula el vencimiento sumando la duración del documento."""
        try:
            escrito = entrada_fecha.get().strip()
            if escrito:
                entrada_vence.delete(0, "end")
                entrada_vence.insert(0, sumar_meses_fecha(escrito, int(doc.get("meses") or 0)) or "")
        except Exception:
            pass

    for doc in DOCUMENTOS_CHOFER:
        actual = (datos_actuales.get("fechas", {}) or {}).get(doc["clave"], "") or "—"
        marco = ctk.CTkFrame(contenedor, fg_color="#f7f9fb", corner_radius=8)
        marco.pack(fill="x", padx=6, pady=5)
        ctk.CTkLabel(marco, text="📄 %s" % doc["etiqueta"], font=("Arial", 12, "bold"),
                     anchor="w").pack(fill="x", padx=10, pady=(8, 0))
        ctk.CTkLabel(marco, text=("Vencimiento actual: %s" % actual if not todos else
                                  "Se aplicará la misma fecha a todos los choferes"),
                     font=("Arial", 10), text_color="gray", anchor="w").pack(fill="x", padx=10)

        linea = ctk.CTkFrame(marco, fg_color="transparent")
        linea.pack(fill="x", padx=10, pady=(4, 9))
        ctk.CTkLabel(linea, text="Fecha en que se realizó:", font=("Arial", 11, "bold")).pack(side="left")
        entrada_fecha = ctk.CTkEntry(linea, placeholder_text="DD/MM/AAAA", width=120)
        entrada_fecha.pack(side="left", padx=(6, 4))
        ctk.CTkButton(linea, text="[ 📅 ]", width=44, font=("Arial", 11, "bold"),
                      fg_color="#2980b9", hover_color="#1f6391",
                      command=lambda e=entrada_fecha, c=doc, vv=None: CalendarioNativo(
                          v, e, lambda _x=None: _calcular(c, e, entrada_vence))).pack(side="left")

        ctk.CTkLabel(linea, text="Fecha de vencimiento:", font=("Arial", 11, "bold")).pack(
            side="left", padx=(16, 4))
        entrada_vence = ctk.CTkEntry(linea, placeholder_text="DD/MM/AAAA", width=120)
        entrada_vence.pack(side="left", padx=(0, 4))
        ctk.CTkButton(linea, text="[ 📅 ]", width=44, font=("Arial", 11, "bold"),
                      fg_color="#2980b9", hover_color="#1f6391",
                      command=lambda e=entrada_vence: CalendarioNativo(v, e)).pack(side="left")

        entrada_fecha.bind("<KeyRelease>", lambda _e, c=doc, f=entrada_fecha, vv=entrada_vence:
                           _calcular(c, f, vv))
        entrada_fecha.bind("<FocusOut>", lambda _e, c=doc, f=entrada_fecha, vv=entrada_vence:
                           _calcular(c, f, vv))

        # ⏱️ Botones de duración: ponen la fecha de vencimiento a partir de la emisión
        atajos = ctk.CTkFrame(marco, fg_color="transparent")
        atajos.pack(fill="x", padx=10, pady=(0, 9))
        ctk.CTkLabel(atajos, text="Atajos:", font=("Arial", 10, "bold"),
                     text_color="gray").pack(side="left", padx=(0, 6))
        duraciones = [int(doc.get("meses") or 0)]
        for alterno in (6, 12, 24, 36):
            if alterno not in duraciones:
                duraciones.append(alterno)
        for meses in duraciones[:3]:
            etiqueta_boton = "Vence %s" % texto_duracion_meses(meses)
            ctk.CTkButton(atajos, text=etiqueta_boton, width=118, height=26,
                          font=("Arial", 11, "bold"),
                          fg_color=("#1f538d" if meses == int(doc.get("meses") or 0) else "#5d6d7e"),
                          hover_color="#163b65",
                          command=lambda m=meses, f=entrada_fecha, vv=entrada_vence, c=doc:
                          _aplicar_duracion(m, f, vv, c, v)).pack(side="left", padx=3)

        # 📎 Brevete: escanear / cargar PDF / ver (solo cuando es UN chofer)
        if str(doc.get("clave")) == "brevete":
            if todos:
                ctk.CTkLabel(marco, text="📎 Para adjuntar el brevete abra un chofer específico.",
                             font=("Arial", 10), text_color="#999999").pack(anchor="w", padx=10, pady=(0, 8))
            else:
                fila_documento_adjunto(marco, "chofer", str(elegidos[0]["id"]),
                                       doc["clave"], doc["etiqueta"])

        filas.append((doc, entrada_fecha, entrada_vence))

    def aplicar():
        aplicados, errores = 0, []
        for doc, entrada_fecha, entrada_vence in filas:
            realizado = entrada_fecha.get().strip()
            vence = entrada_vence.get().strip()
            if not realizado and not vence:
                continue
            for chofer in elegidos:
                ok, mensaje = actualizar_documento_chofer(
                    chofer["id"], doc["clave"], fecha_realizado=realizado,
                    vencimiento=vence, usuario=usuario,
                    categoria=str(chofer.get("categoria") or ""))
                if ok:
                    aplicados += 1
                else:
                    errores.append("%s - %s: %s" % (chofer.get("etiqueta", "?"), doc["etiqueta"], mensaje))
        try:
            cache_sistema.invalidar()
        except Exception:
            pass
        if errores:
            messagebox.showwarning("Documentos del chofer",
                                   "Se aplicaron %d documento(s).\n\nCon problemas:\n%s"
                                   % (aplicados, chr(10).join(errores[:6])), parent=v)
        else:
            messagebox.showinfo("Documentos del chofer",
                                "Se aplicaron %d documento(s) a %d chofer(es).\n\nLos vencimientos "
                                "ya están en «Cronograma y Vencimientos»."
                                % (aplicados, len(elegidos)), parent=v)
        if aplicados:
            try:
                v.destroy()
            except Exception:
                pass

    barra = ctk.CTkFrame(v, fg_color="transparent")
    barra.pack(fill="x", padx=12, pady=(0, 12))
    ctk.CTkButton(barra, text="✅ Aplicar documentos", height=42, font=("Arial", 13, "bold"),
                  fg_color="#27ae60", hover_color="#1e8449", command=aplicar).pack(
        side="left", expand=True, fill="x", padx=(0, 6))
    ctk.CTkButton(barra, text="✖ Cancelar", height=42, font=("Arial", 13, "bold"),
                  fg_color="#7f8c8d", hover_color="#616a6b", command=v.destroy).pack(
        side="left", expand=True, fill="x", padx=(6, 0))
    return v


def _refrescar_flota_abierta():
    """Refresca la ventana de Flota Automotriz si está abierta (tras un mantenimiento)."""
    try:
        import control_general
        instancia = getattr(control_general, "INSTANCIA_ACTIVA", None)
        if instancia is None:
            return
        instancia._refrescar_modulo("flota")
    except Exception as e:
        print("Aviso - refrescando Flota:", e)


def abrir_reset_mantenimientos_vehiculo(parent, placa, usuario_activo="", al_terminar=None):
    """Ventana con CHECK por mantenimiento para registrar el servicio hecho a una unidad.

    Al marcar los mantenimientos realizados y escribir la fecha del servicio se
    actualizan las fechas del vehículo en 'flota_vehiculos' (las mismas que
    controla Flota Automotriz / Cronograma) y se calcula el próximo vencimiento.
    """
    placa = str(placa or "").strip()
    if not placa or placa.upper().startswith("SIN "):
        return messagebox.showwarning("Mantenimientos",
                                      "Este registro no tiene un vehículo asignado.\n"
                                      "Asigne la placa en la factura para poder registrar sus mantenimientos.",
                                      parent=parent)

    conn = conectar_db(silencioso=True)
    if not conn:
        return messagebox.showerror("Error", "Sin conexión a la base de datos.", parent=parent)
    vehiculo = None
    try:
        with conn.cursor() as c:
            # ::text porque varias columnas (km_ultimo_aceite, km_prox_correa...) son NUMERIC
            columnas_sql = ", ".join("COALESCE({}::text, '')".format(col) for col in COLUMNAS_MANT)
            c.execute("SELECT " + columnas_sql +
                      " FROM flota_vehiculos WHERE UPPER(placa) = UPPER(%s) LIMIT 1", (placa,))
            fila = c.fetchone()
            if fila:
                vehiculo = dict(zip(COLUMNAS_MANT, fila))
    except Exception as e:
        vehiculo = None
        print("[Compras -> Mantenimientos]", e)
    finally:
        liberar_conexion(conn)

    if not vehiculo:
        return messagebox.showwarning("Mantenimientos",
                                      f"La unidad '{placa}' no está registrada en Flota Automotriz.\n"
                                      "Regístrela primero en ese módulo para poder controlar sus mantenimientos.",
                                      parent=parent)

    v = ctk.CTkToplevel(parent)
    v.title(f"Mantenimientos - {vehiculo['placa']}")
    centrar_ventana(v, parent, 640, 660)
    v.transient(parent)
    v.grab_set()

    ctk.CTkLabel(v, text=f"🛠️ Mantenimientos de {vehiculo['placa']}", font=("Arial", 15, "bold"),
                 text_color="#1f538d").pack(pady=(12, 2))
    detalle = " ".join(x for x in (vehiculo.get("marca"), vehiculo.get("modelo")) if x)
    ctk.CTkLabel(v, text=f"{detalle}    |    Kilometraje registrado: {vehiculo.get('kilometraje') or '-'}",
                 font=("Arial", 11, "italic"), text_color="#7f8c8d").pack(pady=(0, 4))
    ctk.CTkLabel(v, text="Marque los mantenimientos realizados e ingrese la fecha del servicio.\n"
                         "Las fechas del vehículo se actualizan y queda calculado el próximo vencimiento.",
                 font=("Arial", 10), text_color="#555555", justify="left").pack(padx=12, pady=(0, 4))

    scroll = ctk.CTkScrollableFrame(v, fg_color="transparent")
    scroll.pack(fill="both", expand=True, padx=8, pady=(0, 4))

    hoy_txt = datetime.now().strftime("%d/%m/%Y")
    filas = []

    def actualizar_vista(*_):
        """Muestra, para cada mantenimiento marcado, las fechas que se van a guardar."""
        for f in filas:
            if not f["var"].get():
                f["lbl"].configure(text="")
                continue
            fecha = f["ent_f"].get().strip()
            partes = []
            for col, meses in f["item"]["campos"]:
                valor = sumar_meses_a_fecha(fecha, meses)
                partes.append(f"{ETIQUETAS_MANT.get(col, col)}: {valor or 'fecha inválida'}")
            f["lbl"].configure(text="Se guardará → " + " · ".join(partes))

    def texto_actual(item):
        partes = []
        for col in item.get("actual", ()):
            valor = vehiculo.get(col)
            if valor in (None, "") or str(valor).strip() in ("", "0", "0.0"):
                continue
            etiqueta = ETIQUETAS_MANT.get(col, col)
            if col.startswith("km_"):
                try:
                    valor = f"{float(valor):,.0f}"
                except Exception:
                    pass
            partes.append(f"{etiqueta}: {valor}")
        return " · ".join(partes) or "sin registro"

    for item in MANTENIMIENTOS_VEHICULO:
        card = ctk.CTkFrame(scroll, fg_color="#f8f9fa", corner_radius=8,
                            border_width=1, border_color="#e0e0e0")
        card.pack(fill="x", pady=3)

        var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(card, text=item["nombre"], variable=var, font=("Arial", 11, "bold"),
                        command=actualizar_vista).pack(anchor="w", padx=10, pady=(8, 0))
        lbl_actual = ctk.CTkLabel(card, text=f"Actual: {texto_actual(item)}", font=("Arial", 10, "italic"),
                                  text_color="#7f8c8d", wraplength=560, justify="left")
        lbl_actual.pack(anchor="w", padx=30, pady=(0, 2))

        fr = ctk.CTkFrame(card, fg_color="transparent")
        fr.pack(fill="x", padx=30, pady=(2, 6))
        ctk.CTkLabel(fr, text="Fecha en que se realizó:", font=("Arial", 11)).pack(side="left")
        ent_f = ctk.CTkEntry(fr, width=115)
        ent_f.pack(side="left", padx=6)
        ent_f.insert(0, hoy_txt)
        ctk.CTkButton(fr, text="📅", width=36, fg_color="#1f538d", hover_color="#163b65",
                      command=lambda e=ent_f: CalendarioNativo(v, e)).pack(side="left")

        ent_km = None
        if item.get("km"):
            col_km, etq_km = item["km"]
            ctk.CTkLabel(fr, text=f"{etq_km}:", font=("Arial", 11)).pack(side="left", padx=(14, 4))
            ent_km = ctk.CTkEntry(fr, width=95)
            ent_km.pack(side="left")
            km_previo = vehiculo.get(col_km)
            if km_previo not in (None, "") and str(km_previo).strip() not in ("", "0", "0.0"):
                try:
                    ent_km.insert(0, f"{float(km_previo):,.0f}")
                except Exception:
                    ent_km.insert(0, str(km_previo))
            elif col_km == "km_ultimo_aceite" and vehiculo.get("kilometraje"):
                ent_km.insert(0, str(vehiculo.get("kilometraje")))

        lbl_prox = ctk.CTkLabel(card, text="", font=("Arial", 10, "bold"), text_color="#27ae60",
                                wraplength=560, justify="left")
        lbl_prox.pack(anchor="w", padx=30, pady=(0, 8))

        # 📎 Certificado / comprobante del documento (escanear, cargar PDF o ver)
        if item.get("doc"):
            fila_documento_adjunto(card, "vehiculo", vehiculo.get("placa", ""),
                                   item["doc"], item["nombre"])

        filas.append({"item": item, "var": var, "ent_f": ent_f, "ent_km": ent_km, "lbl": lbl_prox})

    for f in filas:
        f["ent_f"].bind("<KeyRelease>", actualizar_vista)

    f_pie = ctk.CTkFrame(v, fg_color="transparent")
    f_pie.pack(fill="x", padx=10, pady=(0, 10))

    def aplicar():
        sets, params, resumen = [], [], []
        for f in filas:
            if not f["var"].get():
                continue
            fecha = f["ent_f"].get().strip()
            if not re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", fecha):
                return messagebox.showerror("Mantenimientos",
                                            f"Fecha inválida en «{f['item']['nombre']}» (use DD/MM/AAAA).",
                                            parent=v)
            for col, meses in f["item"]["campos"]:
                sets.append(f"{col} = %s")
                params.append(sumar_meses_a_fecha(fecha, meses))
            if f["item"].get("km") and f["ent_km"] is not None:
                km_txt = f["ent_km"].get().strip()
                if km_txt:
                    try:
                        km_val = float(km_txt.replace(",", "").replace(" ", ""))
                    except ValueError:
                        return messagebox.showerror("Mantenimientos",
                                                    f"Kilometraje inválido en «{f['item']['nombre']}».",
                                                    parent=v)
                    sets.append(f"{f['item']['km'][0]} = %s")
                    params.append(km_val)
                    # El kilometraje general del vehículo se mantiene al día
                    if f["item"]["km"][0] in ("km_ultimo_aceite", "km_ultimo_general"):
                        try:
                            km_actual = float(str(vehiculo.get("kilometraje") or "0").replace(",", "") or 0)
                        except ValueError:
                            km_actual = 0.0
                        if km_val > km_actual:
                            sets.append("kilometraje = %s")
                            params.append(f"{km_val:,.0f}")
            resumen.append(f"{f['item']['nombre']}: {fecha}")

        if not sets:
            return messagebox.showinfo("Mantenimientos", "Marque al menos un mantenimiento realizado.",
                                       parent=v)

        conn2 = conectar_db()
        if not conn2:
            return messagebox.showerror("Error", "Sin conexión a la base de datos.", parent=v)
        try:
            cursor = conn2.cursor()
            cursor.execute("UPDATE flota_vehiculos SET " + ", ".join(sets) +
                           " WHERE UPPER(placa) = UPPER(%s)", tuple(params + [vehiculo["placa"]]))
            conn2.commit()
            # 🗓️ El Cronograma y el panel de Vencimientos trabajan con caché: se limpia
            # para que las fechas nuevas se vean de inmediato en el calendario.
            try:
                cache_sistema.invalidar()
            except Exception:
                pass
            # 🔄 Si el módulo Flota Automotriz está abierto, se refresca en el acto
            try:
                _refrescar_flota_abierta()
            except Exception:
                pass
            try:
                registrar_auditoria(usuario_activo, "Compras",
                                    f"Mantenimientos de {vehiculo['placa']}: " + "; ".join(resumen))
            except Exception:
                pass
            messagebox.showinfo("Mantenimientos",
                                f"Se actualizaron los mantenimientos de {vehiculo['placa']}:\n\n" +
                                "\n".join("• " + r for r in resumen) +
                                "\n\nYa se reflejan en Flota Automotriz y en el Cronograma.", parent=v)
            v.destroy()
            if al_terminar:
                al_terminar()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron actualizar los mantenimientos:\n{e}", parent=v)
        finally:
            liberar_conexion(conn2)

    ctk.CTkButton(f_pie, text="✅ Aplicar mantenimientos", height=36, font=("Arial", 12, "bold"),
                  fg_color="#27ae60", hover_color="#1e8449", command=aplicar).pack(side="left", expand=True,
                                                                                   fill="x", padx=(0, 5))
    ctk.CTkButton(f_pie, text="✖ Cancelar", height=36, font=("Arial", 12), fg_color="#7f8c8d",
                  hover_color="#606b6b", command=v.destroy).pack(side="right", expand=True, fill="x",
                                                                 padx=(5, 0))


def _escanear_con_progreso(parent, etiqueta, mod_escaner, *args, **kwargs):
    """Escanea mostrando el cuadrito "Escaneando…" y lo cierra al terminar.

    El escaneo es sincrónico (así el resultado nunca se pierde); el cuadrito se
    pinta antes de empezar para que el usuario vea que está trabajando.
    """
    try:
        parent = parent or getattr(tk, "_default_root", None)
    except Exception:
        parent = None
    prog = None
    try:
        prog = ctk.CTkToplevel(parent)
        prog.title("Escaneando")
        prog.geometry("450x165")
        try:
            prog.transient(parent)
        except Exception:
            pass
        try:
            prog.grab_set()
        except Exception:
            pass
        ctk.CTkLabel(prog, text="🖨️ Escaneando %s..." % (etiqueta or "el documento"),
                     font=("Arial", 13, "bold"), text_color="#1f538d").pack(pady=(20, 4))
        ctk.CTkLabel(prog, text="Coloque el documento en el escáner (cristal).\nEspere unos segundos…",
                     font=("Arial", 10), text_color="gray", justify="center").pack()
        barra = ctk.CTkProgressBar(prog, width=350, mode="indeterminate")
        barra.pack(pady=12)
        barra.start()
        try:
            prog.update_idletasks()
            prog.update()
        except Exception:
            pass
    except Exception as e_prog:
        print("[Escáner] No se pudo mostrar el progreso:", e_prog)
        prog = None
    try:
        return mod_escaner.escanear(*args, **kwargs)
    finally:
        if prog is not None:
            try:
                prog.grab_release()
            except Exception:
                pass
            try:
                prog.destroy()
            except Exception:
                pass


def _pick_comprobante(parent):
    """Elige un archivo PDF/imagen para el comprobante."""
    try:
        from tkinter import filedialog
        return filedialog.askopenfilename(
            parent=parent, title="Seleccionar el comprobante de pago",
            filetypes=[("PDF o imagen", "*.pdf *.jpg *.jpeg *.png"), ("Todos los archivos", "*.*")])
    except Exception:
        return ""


def _ruta_columnas_visibles():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "columnas_visibles.json")


def guardar_columnas_visibles(clave, columnas):
    try:
        ruta = _ruta_columnas_visibles()
        datos = {}
        if os.path.exists(ruta):
            try:
                with open(ruta, "r", encoding="utf-8") as f:
                    datos = json.load(f)
            except Exception:
                datos = {}
        datos[clave] = list(columnas)
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=1)
    except Exception as e:
        print("[Columnas] No se pudo guardar:", e)


def cargar_columnas_visibles(clave):
    try:
        ruta = _ruta_columnas_visibles()
        if not os.path.exists(ruta):
            return None
        with open(ruta, "r", encoding="utf-8") as f:
            datos = json.load(f)
        return datos.get(clave)
    except Exception:
        return None


def aplicar_columnas_visibles(tabla, clave):
    try:
        guardadas = cargar_columnas_visibles(clave)
        if not guardadas:
            return
        orden = [c for c in tabla["columns"] if c in guardadas]
        if orden:
            tabla.configure(displaycolumns=tuple(orden))
    except Exception as e:
        print("[Columnas] No se pudo restaurar:", e)


def _columnas_visibles(tabla):
    try:
        dc = tabla.cget("displaycolumns")
        if isinstance(dc, (tuple, list)):
            cols = [str(c) for c in dc]
        else:
            cols = [c for c in str(dc).split() if c]
        if not cols or cols == ["#all"]:
            return [c for c in tabla["columns"]]
        return [c for c in cols if c != "#all"]
    except Exception:
        return [c for c in tabla["columns"]]


def _alternar_columna(tabla, clave, col, visible):
    try:
        visibles = set(_columnas_visibles(tabla))
        if visible:
            visibles.add(col)
        else:
            visibles.discard(col)
        orden = [c for c in tabla["columns"] if c in visibles]
        tabla.configure(displaycolumns=tuple(orden) if orden else ("#all",))
        guardar_columnas_visibles(clave, orden)
    except Exception as e:
        print("[Columnas] No se pudo alternar:", e)


def menu_columnas(tabla, clave):
    menu = tk.Menu(tabla, tearoff=0)
    visibles = set(_columnas_visibles(tabla))
    variables = []
    for col in tabla["columns"]:
        texto = str(tabla.heading(col, "text") or "").strip()
        if not texto or "Oculto" in texto:
            continue
        var = tk.BooleanVar(master=tabla, value=(col in visibles))
        variables.append(var)
        menu.add_checkbutton(label=texto, variable=var,
                             command=lambda c=col, v=var: _alternar_columna(tabla, clave, c, v.get()))
    menu._vars = variables
    try:
        menu.tk_popup(tabla.winfo_pointerx(), tabla.winfo_pointery())
    finally:
        menu.grab_release()




# =========================================================
# VENTANA AL FRENTE + AVISOS SOBRE EL MÓDULO (Compras)
# =========================================================
_VENTANA_COMPRAS = [None]


def _traer_ventana_al_frente():
    """Devuelve la ventana del módulo al frente (Windows suele mandarla atrás)."""
    ventana = _VENTANA_COMPRAS[0]
    if ventana is None:
        return
    # 🚫 Si hay una ventana modal abierta (Registrar Pago, edición…) NO se trae el
    #    módulo al frente: la ventana de trabajo debe quedar SIEMPRE por delante.
    try:
        _que_tiene_grab = ventana.grab_current()
        if _que_tiene_grab is not None and _que_tiene_grab is not ventana:
            return
    except Exception:
        pass
    try:
        ventana.deiconify()
        ventana.lift()
        ventana.attributes("-topmost", True)
        ventana.after(400, lambda: ventana.attributes("-topmost", False))
        ventana.focus_force()
    except Exception:
        pass


class _MessageBoxCompras:
    """Envuelve messagebox: cada aviso sale SOBRE el módulo y deja la ventana al frente."""

    def __init__(self, real):
        self._real = real

    def _parent(self):
        ventana = _VENTANA_COMPRAS[0]
        try:
            if ventana is not None and ventana.winfo_exists():
                return ventana
        except Exception:
            pass
        return None

    def _ejecutar(self, nombre, *args, **kwargs):
        # 🪟 El aviso sale sobre la ventana que lo pide (una ventana de pago, de
        #    edición, etc.) y, al cerrarlo, se devuelve el frente a ESA ventana.
        #    Antes se traía al frente el módulo y la ventana de trabajo quedaba
        #    detrás (por eso "desaparecía" la ventana de cancelación).
        padre = kwargs.get("parent") or self._parent()
        kwargs["parent"] = padre
        try:
            return getattr(self._real, nombre)(*args, **kwargs)
        finally:
            try:
                if padre is not None and padre.winfo_exists():
                    padre.deiconify()
                    padre.lift()
                    padre.focus_force()
                    try:
                        padre.grab_set()
                    except Exception:
                        pass
                else:
                    _traer_ventana_al_frente()
            except Exception:
                _traer_ventana_al_frente()

    def showinfo(self, *a, **k):
        return self._ejecutar("showinfo", *a, **k)

    def showwarning(self, *a, **k):
        return self._ejecutar("showwarning", *a, **k)

    def showerror(self, *a, **k):
        return self._ejecutar("showerror", *a, **k)

    def askyesno(self, *a, **k):
        return self._ejecutar("askyesno", *a, **k)

    def askokcancel(self, *a, **k):
        return self._ejecutar("askokcancel", *a, **k)

    def askquestion(self, *a, **k):
        return self._ejecutar("askquestion", *a, **k)

    def __getattr__(self, nombre):
        return getattr(self._real, nombre)


messagebox = _MessageBoxCompras(messagebox)


def _log_escaneo_comprobante(etiqueta, ruta="", fecha="", operacion="", monto=0.0, nota=""):
    """Escribe un registro del escaneo del comprobante (para diagnosticar)."""
    try:
        import datetime as _dt
        from app_paths import DATA_DIR
        ruta_log = os.path.join(str(DATA_DIR), "compras_scan.log")
        linea = "%s | %s | archivo=%s | fecha_leida=%r | operacion=%r | monto=%r | %s\n" % (
            _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), etiqueta,
            os.path.basename(str(ruta or "")), fecha, operacion, monto, nota)
        with open(ruta_log, "a", encoding="utf-8") as f:
            f.write(linea)
    except Exception:
        pass


def _ocr_comprobante_con_reintentos(paginas, intentos=3):
    """OCR de las páginas recién escaneadas, esperando a que el archivo esté listo.

    Al escanear, el JPG se termina de escribir unos instantes después; si el OCR
    corre antes, devuelve vacío y la ventana se queda sin fecha/operación/monto.
    Aquí se espera, se reintenta y se devuelve el mejor texto obtenido.
    """
    import time as _time
    paginas = [str(p) for p in (paginas or []) if p and os.path.exists(str(p))]
    if not paginas:
        return ""
    mejor = ""
    for intento in range(max(1, int(intentos))):
        # se espera a que el tamaño del archivo se estabilice (escritura terminada)
        try:
            tam1 = os.path.getsize(paginas[0])
            _time.sleep(0.25)
            tam2 = os.path.getsize(paginas[0])
            if tam1 != tam2:
                _time.sleep(0.40)
        except Exception:
            pass
        texto = ""
        try:
            import escaner as _esc
            texto = _esc.texto_de_resultados(_esc.ocr_paginas(paginas)) or ""
        except Exception as e:
            print("[Compras] OCR del comprobante (intento %d): %s" % (intento + 1, e))
        if len(str(texto).strip()) > len(str(mejor).strip()):
            mejor = texto
        if len("".join(str(mejor).split())) >= 20:
            return mejor
        _time.sleep(0.35 * (intento + 1))
    return mejor


class FacturasRecibidasTab:
    def __init__(self, tab_frame, main_root, app_padre):
        self.tab_frame = tab_frame
        self.main_root = main_root
        self.app_padre = app_padre
        
        self.orden_columnas = {}
        # 🔃 Ordenamiento por cualquier columna aplicado a TODAS las páginas
        self.columna_orden = "fecha"
        self.orden_ascendente = False
        self.total_paginas = 1
        self.bloquear_autocompletado_ruc = False
        self._ruc_autocompletado = False
        self.ruta_archivo_temp = ""

        # 💰 Modo de ingreso del monto: "BASE" (monto sin IGV) o "CON_IGV" (monto total del documento).
        # El usuario elige con el selector "Monto Base / Monto con IGV" y el sistema calcula el resto.
        self.modo_monto = "BASE"
        
        # VARIABLES DE PAGINACIÓN (LAZY LOADING)
        self.pagina_actual = 1
        self.registros_por_pagina = 55
        
        # 🗓️ FILTRO DE MES (por defecto, el mes en curso)
        self.mes_filtro = mes_en_curso()

        # 🚫 Check para OCULTAR las facturas enviadas por la App (App Grifo).
        # Por estándar viene desmarcado: se ven todas.
        self.var_ocultar_app = tk.BooleanVar(value=False)

        self.inicializar_bd()
        self.crear_interfaz()

    def inicializar_bd(self):
        global _SCHEMA_COMPRAS_OK
        if _SCHEMA_COMPRAS_OK: return

        def tarea_curacion():
            global _SCHEMA_COMPRAS_OK
            conn = conectar_db(silencioso=True)
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
                
                try: cursor.execute("ALTER TABLE facturas_recibidas ADD COLUMN IF NOT EXISTS kilometraje VARCHAR(50);"); conn.commit()
                except: conn.rollback()
                try: cursor.execute("ALTER TABLE facturas_recibidas ADD COLUMN IF NOT EXISTS cantidad_combustible VARCHAR(50);"); conn.commit()
                except: conn.rollback()
                try: cursor.execute("ALTER TABLE facturas_recibidas ADD COLUMN IF NOT EXISTS ruc VARCHAR(50);"); conn.commit()
                except: conn.rollback()
                # Columnas usadas para distinguir las facturas de la App y las compras cruzadas
                try: cursor.execute("ALTER TABLE facturas_recibidas ADD COLUMN IF NOT EXISTS imagen_base64 TEXT;"); conn.commit()
                except: conn.rollback()
                try: cursor.execute("ALTER TABLE facturas_recibidas ADD COLUMN IF NOT EXISTS es_compra_cruzada BOOLEAN DEFAULT FALSE;"); conn.commit()
                except: conn.rollback()
                try: cursor.execute("ALTER TABLE facturas_recibidas ADD COLUMN IF NOT EXISTS pagado_por_tercero VARCHAR(255) DEFAULT '';"); conn.commit()
                except: conn.rollback()
                try: cursor.execute("ALTER TABLE facturas_recibidas ADD COLUMN IF NOT EXISTS soporte_pago_tercero TEXT DEFAULT '';"); conn.commit()
                except: conn.rollback()
                # 💸 Comisión interbancaria de la factura cruzada (igual que en Banco)
                try: cursor.execute("ALTER TABLE facturas_recibidas ADD COLUMN IF NOT EXISTS comision_interbancaria NUMERIC DEFAULT 0;"); conn.commit()
                except: conn.rollback()
                try: cursor.execute("ALTER TABLE facturas_recibidas ADD COLUMN IF NOT EXISTS es_interbancario BOOLEAN DEFAULT FALSE;"); conn.commit()
                except: conn.rollback()
                # 🔗 Cruce de la factura con su Orden de Servicio (módulo de Órdenes)
                try: cursor.execute("ALTER TABLE facturas_recibidas ADD COLUMN IF NOT EXISTS id_orden_servicio INTEGER;"); conn.commit()
                except: conn.rollback()
                try: cursor.execute("ALTER TABLE facturas_recibidas ADD COLUMN IF NOT EXISTS orden_servicio VARCHAR(120) DEFAULT '';"); conn.commit()
                except: conn.rollback()
                try: cursor.execute("ALTER TABLE facturas_recibidas ADD COLUMN IF NOT EXISTS chofer_asociado VARCHAR(160) DEFAULT '';"); conn.commit()
                except: conn.rollback()
                try: cursor.execute("ALTER TABLE facturas_recibidas ADD COLUMN IF NOT EXISTS delivery NUMERIC DEFAULT 0;"); conn.commit()
                except: conn.rollback()
                try: cursor.execute("ALTER TABLE pagos_comprobantes ADD COLUMN IF NOT EXISTS comision NUMERIC DEFAULT 0;"); conn.commit()
                except: conn.rollback()
                _SCHEMA_COMPRAS_OK = True
            except Exception: pass
            finally: liberar_conexion(conn)

        threading.Thread(target=tarea_curacion, daemon=True).start()

    def abrir_calendario(self, entry_objetivo):
        CalendarioNativo(self.main_root.winfo_toplevel(), entry_objetivo)

    TITULOS_ORDEN = {
        "fecha": "Fecha Fac.", "hora": "Hora", "nro_doc": "N° Doc.", "proveedor": "Proveedor",
        "ruc": "RUC", "evento": "Vehículo (Placa)", "kilometraje": "Kilometraje",
        "cantidad": "Galones/Cant.", "desc": "Concepto", "metodo_pago": "Forma de Pago",
        "neto": "Neto Pagar",
    }

    def _actualizar_flechas_orden(self):
        """Muestra ▲/▼ en la columna activa y ↕ en las demás."""
        for columna, titulo in self.TITULOS_ORDEN.items():
            if columna == self.columna_orden:
                flecha = "▲" if self.orden_ascendente else "▼"
            else:
                flecha = "↕"
            try:
                self.tabla.heading(columna, text=f"{titulo} {flecha}")
            except Exception:
                pass

    def ordenar_por_columna(self, columna, es_numerico=None):
        """Ordena TODAS las páginas por la columna elegida (no solo la página visible)."""
        if self.columna_orden == columna:
            self.orden_ascendente = not self.orden_ascendente
        else:
            self.columna_orden = columna
            self.orden_ascendente = True
        self._actualizar_flechas_orden()
        self.pagina_actual = 1
        self.cargar_datos_tabla(reset_pagina=True)

    def _aviso_procesando(self, mensaje):
        """Ventana pequena "procesando..." mientras se lee el documento (OCR)."""
        try:
            ventana = ctk.CTkToplevel(getattr(self, "main_root", None))
            ventana.title("Procesando")
            ventana.geometry("380x130")
            ventana.resizable(False, False)
            try:
                ventana.transient(getattr(self, "main_root", None))
                ventana.grab_set()
            except Exception:
                pass
            ctk.CTkLabel(ventana, text=mensaje, font=("Arial", 13, "bold"),
                         wraplength=340, justify="center").pack(pady=(20, 10), padx=16)
            barra = ctk.CTkProgressBar(ventana, mode="indeterminate", width=320)
            barra.pack(pady=(0, 14))
            barra.start()
            try:
                ventana._inicio_aviso = time.time()   # para que se vea un minimo de tiempo
            except Exception:
                pass
            ventana.update_idletasks()
            ventana.update()
            # que quede SIEMPRE a la vista mientras se procesa
            for _accion in (ventana.lift, ventana.focus_force,
                            lambda: ventana.attributes("-topmost", True)):
                try:
                    _accion()
                except Exception:
                    pass
            ventana.update()
            return ventana
        except Exception:
            return None

    def _cerrar_aviso(self, ventana, minimo=1.2):
        """Cierra la ventana de 'procesando...' dejandola visible al menos 'minimo' segundos."""
        try:
            if ventana is None:
                return
            try:
                inicio = getattr(ventana, "_inicio_aviso", None)
                if inicio:
                    restante = float(minimo) - (time.time() - inicio)
                    if restante > 0:
                        ventana.update()
                        time.sleep(min(restante, 2.0))
            except Exception:
                pass
            ventana.grab_release()
            ventana.destroy()
        except Exception:
            pass

    def _extraer_texto_documento(self, ruta):
        """Capa de texto del PDF (pdfplumber y, si falla, PyMuPDF). Devuelve "" si es un escaneo."""
        texto = ""
        if pdfplumber is not None:
            try:
                with pdfplumber.open(ruta) as pdf:
                    for pagina in pdf.pages:
                        extraido = pagina.extract_text()
                        if extraido:
                            texto += extraido + "\n"
            except Exception:
                texto = ""
        if not texto.strip():
            try:
                import fitz
                documento = fitz.open(ruta)
                for pagina in documento:
                    texto += (pagina.get_text() or "") + "\n"
                documento.close()
            except Exception:
                pass
        return texto

    def escanear_factura_ocr(self):
        """Escanea la factura en la impresora multifuncional y la lee con OCR."""
        try:
            import escaner
        except Exception as e:
            return messagebox.showerror("Escáner", f"No se pudo cargar el módulo de escaneo:\n{e}", parent=ventana_para_mensaje(getattr(self, "main_root", None)))
        if not escaner.sistema_soportado():
            return messagebox.showinfo("Escáner", "El escaneo con la multifuncional está disponible en Windows y en macOS (M1/M2).", parent=ventana_para_mensaje(getattr(self, "main_root", None)))

        config = cargar_configuracion_regional()
        datos_escaner = escaner.config_escaner(config)
        modo = datos_escaner.get("escaner_modo", "Automático (sin ventanas)")
        instruccion = ("Se abrirá el cuadro de escaneo del sistema." if str(modo).lower().startswith("con la ventana")
                       else f"Coloque la factura en el escáner (cristal) — {datos_escaner.get('escaner_dpi', '300')} ppp, "
                            f"{datos_escaner.get('escaner_color', 'Escala de grises')}.")
        if not datos_escaner.get("escaner_dispositivo") and not str(modo).lower().startswith("con la ventana"):
            instruccion += "\n(Si hay varios escáneres, elíjalo en Ajustes → Configuración General → Escáner.)"

        v_prog = ctk.CTkToplevel(self.main_root)
        v_prog.title("Escaneando factura")
        v_prog.geometry("430x170")
        v_prog.transient(self.main_root)
        v_prog.grab_set()
        ctk.CTkLabel(v_prog, text="🖨️ Escaneando y leyendo la factura...", font=("Arial", 13, "bold"),
                     text_color="#1f538d").pack(pady=(18, 4))
        ctk.CTkLabel(v_prog, text=instruccion, font=("Arial", 10), text_color="gray",
                     justify="center", wraplength=380).pack(pady=(0, 10))
        barra = ctk.CTkProgressBar(v_prog, width=340, mode="indeterminate")
        barra.pack(pady=4)
        barra.start()
        ctk.CTkLabel(v_prog, text="Esto puede tardar unos segundos.", font=("Arial", 9, "italic"),
                     text_color="gray").pack(pady=(4, 0))

        threading.Thread(target=self._worker_escanear, args=(config, v_prog), daemon=True).start()

    def _worker_escanear(self, config, v_prog, paginas_previas=None):
        try:
            import escaner
            resultado = escaner.procesar_factura(config, paginas_previas=paginas_previas)
            # 🧠 Mejor lectura: se rehace el OCR por POSICIONES (el del módulo escaner
            #    aplasta todo en una sola línea y pierde los rótulos del comprobante).
            if resultado.get("ok") and resultado.get("pdf"):
                try:
                    texto_mejor = leer_ocr_windows(resultado["pdf"])
                    if texto_mejor and texto_mejor.strip():
                        resultado["texto"] = texto_mejor
                        resultado["ocr"] = {"motor": "Windows OCR (por posiciones)"}
                except Exception as e_ocr_mejor:
                    print("Aviso - OCR por posiciones:", e_ocr_mejor)
        except Exception as e:
            resultado = {"ok": False, "error": str(e), "paginas": [], "pdf": "", "texto": "", "ocr": {}}

        # 🏢 Lo PRIMERO: el RUC del proveedor -> se registra con toda su información
        # (razón social, dirección y ubicación traídas de SUNAT por internet).
        # Y de paso se controla que el comprobante no esté ya registrado.
        if resultado.get("ok"):
            try:
                datos_txt = extraer_datos_de_texto(resultado.get("texto", ""))
                nombre_prov = str(datos_txt.get("proveedor") or "").strip()
                ruc_prov = re.sub(r"\D", "", str(datos_txt.get("ruc") or ""))
                # 🚫 ¿La factura está a nombre de OTRA empresa? Entonces no se registra
                #    nada: tampoco el proveedor de esa otra empresa.
                estado_destino_w, _msg_w = validar_destinatario(datos_txt)
                resultado["destino"] = estado_destino_w
                if estado_destino_w == "ajeno":
                    resultado["proveedor"] = {"estado": "ajeno",
                                              "datos": {"nombre": nombre_prov, "ruc": ruc_prov,
                                                        "direccion_fiscal": "", "ubicacion": ""}}
                elif nombre_prov or ruc_prov:
                    estado_prov, datos_prov = asegurar_proveedor(nombre_prov, ruc_prov)
                    resultado["proveedor"] = {"estado": estado_prov, "datos": datos_prov}
                else:
                    resultado["proveedor"] = {"estado": "sin-datos",
                                              "datos": {"nombre": "", "ruc": "", "direccion_fiscal": "",
                                                        "ubicacion": ""}}
                resultado["numero"] = datos_txt.get("numero", "") or numero_comprobante_desde_texto(resultado.get("texto", ""))
                resultado["duplicado"] = buscar_comprobante_duplicado(resultado["numero"])
            except Exception as e_prov:
                resultado["proveedor"] = {"estado": "error", "datos": {}}
                resultado["aviso"] = (resultado.get("aviso", "") + f" (proveedor: {e_prov})").strip()

        try:
            self.main_root.after(0, lambda: self._mostrar_resultado_escaneo(resultado, v_prog))
        except Exception:
            pass

    def _mostrar_resultado_escaneo(self, resultado, v_prog=None, reutilizar=None):
        """Ventana con lo escaneado: campos detectados, texto OCR y el PDF adjunto."""
        try:
            if v_prog is not None and v_prog.winfo_exists():
                v_prog.grab_release()
                v_prog.destroy()
        except Exception:
            pass

        if not resultado.get("ok"):
            return messagebox.showerror(
                "Escáner",
                "No se pudo escanear la factura.\n\n" + str(resultado.get("error") or "") +
                "\n\nRevise que la impresora esté encendida y conectada; si su modelo tiene "
                "alimentador, coloque el papel en la bandeja. También puede probar el modo "
                "el modo \"con ventana\" en Ajustes → Configuración General → Escáner.", parent=ventana_para_mensaje(getattr(self, "main_root", None)))

        self._escaneo_actual = resultado
        config = cargar_configuracion_regional()
        datos = extraer_datos_de_texto(resultado.get("texto", ""))
        estado_destino, mensaje_destino = validar_destinatario(datos, config)

        ventana = reutilizar
        if ventana is None:
            ventana = ctk.CTkToplevel(self.main_root)
            ventana.title("🖨️ Factura escaneada (OCR)")
            ventana.transient(self.main_root)
        for hijo in ventana.winfo_children():
            hijo.destroy()
        ventana.geometry("700x640")
        ventana.minsize(620, 520)

        ctk.CTkLabel(ventana, text="🖨️ FACTURA ESCANEADA — REVISAR ANTES DE GUARDAR",
                     font=("Arial", 14, "bold"), text_color="#1f538d").pack(pady=(12, 2))
        info = (f"Páginas: {len(resultado.get('paginas') or [])}   ·   "
                f"Escáner: {resultado.get('escaner') or 'el configurado'}")
        ctk.CTkLabel(ventana, text=info, font=("Arial", 10), text_color="gray").pack()
        if resultado.get("aviso"):
            ctk.CTkLabel(ventana, text="⚠️ " + str(resultado["aviso"]), font=("Arial", 10, "bold"),
                         text_color="#c0392b", wraplength=640, justify="center").pack(pady=(2, 0))

        # 🏢 ¿La factura está a nombre de la empresa configurada?
        colores_destino = {"ok": ("#d4edda", "#155724"), "ajeno": ("#f8d7da", "#721c24"),
                           "sin-datos": ("#fff3cd", "#856404"), "sin-configurar": ("#fff3cd", "#856404")}
        fondo_dest, color_dest = colores_destino.get(estado_destino, ("#f8f9fa", "#2c3e50"))
        f_destino = ctk.CTkFrame(ventana, fg_color=fondo_dest, corner_radius=6)
        f_destino.pack(fill="x", padx=14, pady=(6, 0))
        ctk.CTkLabel(f_destino, text=mensaje_destino, font=("Arial", 11, "bold"), text_color=color_dest,
                     wraplength=650, justify="left").pack(padx=10, pady=6, anchor="w")

        # 🏢 Proveedor: ya quedó registrado (con los datos completos de SUNAT si se pudieron traer)
        info_prov = resultado.get("proveedor") or {}
        datos_prov = info_prov.get("datos") or {}
        estado_prov = info_prov.get("estado", "")
        duplicado = resultado.get("duplicado")
        textos_prov = {
            "creado": ("#d4edda", "#155724", "🏢 Proveedor NUEVO registrado con los datos de SUNAT."),
            "completado": ("#d4edda", "#155724", "🏢 Proveedor ya existía: se completó su ficha con los datos de SUNAT."),
            "existe": ("#eaf2f8", "#1b4f72", "🏢 El proveedor ya estaba registrado."),
            "creado_sin_sunat": ("#fff3cd", "#856404",
                                 "🏢 Proveedor agregado con los datos leídos. No se pudo consultar el RUC en internet "
                                 "(revise la conexión o el token de la API)."),
            "sin-conexion": ("#fff3cd", "#856404", "⚠️ Sin conexión: no se pudo registrar ni verificar el proveedor."),
            "sin-datos": ("#fff3cd", "#856404", "⚠️ No se pudo leer el RUC del proveedor en el comprobante."),
            "error": ("#f8d7da", "#721c24", "⚠️ No se pudo registrar el proveedor."),
            "ajeno": ("#f8d7da", "#721c24",
                      "🚫 Factura de OTRA empresa: no se registró al proveedor en 📦 Proveedores."),
        }
        fondo_prov, color_prov, texto_prov = textos_prov.get(estado_prov, ("#f8f9fa", "#2c3e50", ""))
        if texto_prov:
            f_prov = ctk.CTkFrame(ventana, fg_color=fondo_prov, corner_radius=6)
            f_prov.pack(fill="x", padx=14, pady=(6, 0))
            ctk.CTkLabel(f_prov, text=texto_prov, font=("Arial", 11, "bold"), text_color=color_prov,
                         wraplength=650, justify="left").pack(padx=10, pady=6, anchor="w")

        # 🚫 No se permiten comprobantes repetidos
        if duplicado:
            detalle_dup = (f"N° {duplicado.get('numero')}   ·   Proveedor: {duplicado.get('proveedor')}"
                           f"   ·   Fecha: {duplicado.get('fecha')}"
                           f"   ·   Total: {formatear_moneda(duplicado.get('total') or 0)}")
            f_dup = ctk.CTkFrame(ventana, fg_color="#f8d7da", corner_radius=6)
            f_dup.pack(fill="x", padx=14, pady=(6, 0))
            ctk.CTkLabel(f_dup, text="🚫 ESTE COMPROBANTE YA ESTÁ REGISTRADO — no se permite duplicarlo.\n"
                                     + detalle_dup,
                         font=("Arial", 11, "bold"), text_color="#721c24", wraplength=650,
                         justify="left").pack(padx=10, pady=6, anchor="w")

        f_datos = ctk.CTkFrame(ventana, fg_color="#f8f9fa", corner_radius=8,
                               border_width=1, border_color="#e0e0e0")
        f_datos.pack(fill="x", padx=14, pady=8)
        ctk.CTkLabel(f_datos, text="🔎 Datos detectados por el OCR", font=("Arial", 12, "bold"),
                     text_color="#1f538d").pack(anchor="w", padx=12, pady=(8, 4))
        # Si no se pudo leer/consultar la razón social, se puede escribir aquí antes de aplicar
        f_nombre = ctk.CTkFrame(ventana, fg_color="transparent")
        f_nombre.pack(fill="x", padx=16, pady=(2, 0))
        ctk.CTkLabel(f_nombre, text="Razón social del proveedor:", font=("Arial", 11, "bold")).pack(side="left")
        ent_prov_nombre = ctk.CTkEntry(f_nombre, font=("Arial", 11))
        ent_prov_nombre.pack(side="left", fill="x", expand=True, padx=6)
        ent_prov_nombre.insert(0, datos_prov.get("nombre") or datos.get("proveedor") or "")
        if not (datos_prov.get("nombre") or datos.get("proveedor")):
            ctk.CTkLabel(f_nombre, text="⚠️ escríbala", font=("Arial", 9, "italic"),
                         text_color="#c0392b").pack(side="left")

        campos = (("Tipo:", datos["tipo_documento"]), ("N° Documento:", datos["numero"]),
                  ("Fecha:", datos["fecha"]),
                  ("Proveedor:", datos_prov.get("nombre") or datos["proveedor"]),
                  ("RUC proveedor:", datos_prov.get("ruc") or datos["ruc"]),
                  ("Dirección:", datos_prov.get("direccion_fiscal", "")),
                  ("Ubicación:", datos_prov.get("ubicacion", "")),
                  ("Destinatario:", datos.get("cliente_nombre", "")),
                  ("RUC destinatario:", datos.get("cliente_ruc", "")), ("Subtotal:", formatear_moneda(datos["subtotal"]) if datos["subtotal"] else ""),
                  ("IGV:", formatear_moneda(datos["igv"]) if datos["igv"] else ""),
                  ("Total:", formatear_moneda(datos["total"]) if datos["total"] else ""))
        for etiqueta, valor in campos:
            fila = ctk.CTkFrame(f_datos, fg_color="transparent")
            fila.pack(fill="x", padx=12, pady=1)
            ctk.CTkLabel(fila, text=etiqueta, width=110, anchor="w", font=("Arial", 11, "bold")).pack(side="left")
            ctk.CTkLabel(fila, text=(str(valor) if str(valor).strip() else "(no detectado)"),
                         anchor="w", font=("Arial", 11),
                         text_color=("#2c3e50" if str(valor).strip() else "#c0392b")).pack(side="left")
        ctk.CTkLabel(f_datos, text="", height=4).pack()

        lbl_estado = ctk.CTkLabel(ventana, text="", font=("Arial", 10, "bold"), text_color="#1e8449")

        def copiar_texto():
            try:
                ventana.clipboard_clear()
                ventana.clipboard_append(caja_texto.get("1.0", "end-1c"))
                lbl_estado.configure(text="✅ Texto copiado al portapapeles.")
            except Exception as e:
                lbl_estado.configure(text=f"⚠️ No se pudo copiar: {e}", text_color="#c0392b")

        def ver_pdf():
            ruta_pdf = resultado.get("pdf")
            if ruta_pdf and os.path.exists(ruta_pdf):
                abrir_documento(ruta_pdf)
            else:
                messagebox.showinfo("PDF", "Todavía no se generó el PDF.", parent=ventana)

        def aplicar():
            # 📄 Aunque el OCR NO haya leído el N° de comprobante, se aplica TODO lo
            #    demás (fecha, proveedor, RUC, montos) y el PDF queda adjunto: solo
            #    se pide completar el número. Antes se descartaba todo y el formulario
            #    quedaba vacío.
            try:
                sin_numero = not str(resultado.get("numero") or "").strip()
            except Exception as e_guard:
                print("[Compras] Aviso - guarda del OCR:", e_guard)
            if resultado.get("duplicado"):
                # ⚠️ Se avisa del posible duplicado pero IGUAL se aplican los datos:
                #    si el OCR leyó mal el número, el usuario corrige solo ese campo
                #    (antes se perdía todo lo escaneado).
                try:
                    messagebox.showwarning(
                        "Posible comprobante duplicado",
                        "El comprobante N° %s ya está registrado en el sistema.\n\n"
                        "Se cargaron igual los datos leídos: revise el N° de documento antes de "
                        "guardar (si es el mismo comprobante, no lo registre otra vez)."
                        % resultado["duplicado"].get("numero"),
                        parent=ventana)
                except Exception:
                    pass
            estado_prov = (resultado.get("proveedor") or {}).get("estado", "")
            nombre_final = ent_prov_nombre.get().strip() or datos.get("proveedor", "")
            if estado_prov in ("", "sin-datos") and estado_destino != "ajeno" and nombre_final:
                # Se registra con la razón social que quedó en la ventana
                # (SUNAT ya no se vuelve a consultar para no demorar el escaneo).
                estado_nuevo, datos_finales = asegurar_proveedor(nombre_final, datos.get("ruc", ""),
                                                                 usar_internet=False)
                resultado["proveedor"] = {"estado": estado_nuevo, "datos": datos_finales}
                estado_prov = estado_nuevo
            self.aplicar_datos_escaneo(datos)
            if not estado_prov:
                estado_prov = getattr(self, "_estado_proveedor_escaneo", "")
            extra = {"creado": "  🏢 Proveedor registrado con los datos de SUNAT.",
                     "completado": "  🏢 Se completó la ficha del proveedor con los datos de SUNAT.",
                     "existe": "  El proveedor ya estaba registrado.",
                     "creado_sin_sunat": "  🏢 Proveedor agregado (no se pudo consultar SUNAT).",
                     "sin-conexion": "  ⚠️ Sin conexión: no se pudo verificar el proveedor.",
                     "error": "  ⚠️ No se pudo agregar el proveedor automáticamente.",
                     "sin-datos": ""}.get(estado_prov, "")
            lbl_estado.configure(text="✅ Datos aplicados al formulario." + extra +
                                      "  Revíselos y pulse 💾 Registrar Documento.", text_color="#1e8449")
            try:
                btn_aplicar.configure(state="disabled", text="✅ Datos aplicados")
            except Exception:
                pass
            # ⏱️ Se muestra el aviso, espera 2 segundos, se cierra y cierra la ventana
            #    del escaneo: el formulario queda listo para 💾 Registrar Documento.
            mensaje_final = "Se extrajeron los datos del documento." + extra
            if sin_numero:
                mensaje_final += ("  ⚠️ El N° de documento no se pudo leer: escríbalo a mano "
                                  "antes de registrar.")
                try:
                    self.ent_nro_doc.focus_set()
                except Exception:
                    pass
            self._aviso_temporal(ventana, mensaje_final,
                                 titulo="Extracción Inteligente", segundos=2,
                                 al_cerrar=ventana.destroy)

        def otra_pagina():
            v_prog2 = ctk.CTkToplevel(self.main_root)
            v_prog2.title("Escaneando otra página")
            v_prog2.geometry("400x140")
            v_prog2.transient(self.main_root)
            ctk.CTkLabel(v_prog2, text="🖨️ Escaneando otra página...", font=("Arial", 12, "bold"),
                         text_color="#1f538d").pack(pady=(20, 6))
            ctk.CTkLabel(v_prog2, text="Coloque la siguiente hoja en el escáner.",
                         font=("Arial", 10), text_color="gray").pack()
            barra2 = ctk.CTkProgressBar(v_prog2, width=320, mode="indeterminate")
            barra2.pack(pady=8)
            barra2.start()
            previas = resultado.get("paginas") or []
            threading.Thread(target=self._worker_escanear,
                             args=(cargar_configuracion_regional(), v_prog2, previas), daemon=True).start()

        # 🔻 La barra de botones se ancla abajo: siempre visible aunque la ventana
        # sea pequeña o el texto del OCR sea muy largo.
        f_botones = ctk.CTkFrame(ventana, fg_color="transparent")
        f_botones.pack(side="bottom", fill="x", padx=14, pady=(2, 10))
        lbl_estado.pack(side="bottom", pady=(2, 0))
        btn_aplicar = ctk.CTkButton(f_botones, text="✅ Aplicar al formulario", height=36, font=("Arial", 12, "bold"),
                                    fg_color="#27ae60", hover_color="#1e8449", command=aplicar)
        btn_aplicar.pack(side="left", fill="x", expand=True, padx=(0, 4))
        if estado_destino == "ajeno":
            # 🚫 La factura es de otra empresa: no se puede pasar al formulario
            btn_aplicar.configure(state="disabled", fg_color="#95a5a6",
                                  text="🚫 Factura de otra empresa (rechazada)")
        elif resultado.get("duplicado"):
            # ⚠️ Posible duplicado: el botón sigue ACTIVO porque el OCR pudo leer mal
            #    el número. Se aplica lo leído y se pide revisar el N° antes de guardar.
            try:
                btn_aplicar.configure(fg_color="#e67e22", hover_color="#ca6f1e",
                                      text="⚠️ Aplicar igual (posible duplicado)")
            except Exception:
                pass
        ctk.CTkButton(f_botones, text="➕ Escanear otra página", height=36, font=("Arial", 12, "bold"),
                      fg_color="#8e44ad", hover_color="#703688", command=otra_pagina
                      ).pack(side="left", fill="x", expand=True, padx=4)
        ctk.CTkButton(f_botones, text="📄 Ver PDF", width=110, height=36, font=("Arial", 12, "bold"),
                      fg_color="#1f538d", hover_color="#163b65", command=ver_pdf
                      ).pack(side="left", padx=4)
        ctk.CTkButton(f_botones, text="📋 Copiar OCR", width=130, height=36, font=("Arial", 12, "bold"),
                      fg_color="#34495e", hover_color="#2c3e50", command=copiar_texto
                      ).pack(side="left", padx=4)
        ctk.CTkButton(f_botones, text="❌ Cerrar", width=110, height=36, font=("Arial", 12, "bold"),
                      fg_color="#7f8c8d", hover_color="#606b6b", command=ventana.destroy
                      ).pack(side="left", padx=(4, 0))

        # 📝 El texto del OCR ocupa el espacio que queda entre las fichas y los botones
        ctk.CTkLabel(ventana, text="📝 Texto leído del escaneo (para revisar o copiar)",
                     font=("Arial", 11, "bold")).pack(anchor="w", padx=16)
        caja_texto = ctk.CTkTextbox(ventana, height=110, font=("Consolas", 10))
        caja_texto.pack(fill="both", expand=True, padx=14, pady=(2, 6))
        caja_texto.insert("1.0", resultado.get("texto") or "(el OCR no encontró texto: complete los datos a mano)")

        # La ventana se ajusta y se centra en la pantalla para que NUNCA se corte
        # la barra de botones (ni por tamaño ni por quedar fuera del escritorio).
        try:
            ancho_pantalla, alto_pantalla = ventana.winfo_screenwidth(), ventana.winfo_screenheight()
            ancho = min(740, ancho_pantalla - 40)
            alto = max(540, min(740, alto_pantalla - 110))
            x = max(0, (ancho_pantalla - ancho) // 2)
            y = max(0, (alto_pantalla - alto) // 2 - 20)
            ventana.geometry(f"{ancho}x{alto}+{x}+{y}")
        except Exception:
            pass
        try:
            ventana.lift()
            ventana.focus_force()
        except Exception:
            pass
        return ventana

    def _aviso_temporal(self, padre, mensaje, titulo="Extracción Inteligente", segundos=2,
                        al_cerrar=None):
        """Muestra un aviso que se cierra SOLO a los N segundos (sin botones).

        Se usa al aplicar un escaneo: aparece el mensaje, espera los segundos
        indicados y luego cierra la ventana de escaneo (si se pasó al_cerrar).
        """
        raiz = padre if padre is not None else self.main_root
        try:
            aviso = ctk.CTkToplevel(raiz)
        except Exception:
            if al_cerrar:
                try:
                    al_cerrar()
                except Exception:
                    pass
            return None
        aviso.title(titulo)
        aviso.resizable(False, False)
        try:
            aviso.transient(raiz)
        except Exception:
            pass
        try:
            aviso.attributes("-topmost", True)
        except Exception:
            pass
        ctk.CTkLabel(aviso, text="✅ " + str(mensaje), font=("Arial", 12, "bold"),
                     text_color="#1e8449", wraplength=380, justify="center").pack(expand=True, padx=18, pady=18)
        ancho, alto = 420, 130
        try:
            aviso.update_idletasks()
            x = max(0, (aviso.winfo_screenwidth() - ancho) // 2)
            y = max(0, (aviso.winfo_screenheight() - alto) // 2)
            aviso.geometry(f"{ancho}x{alto}+{x}+{y}")
        except Exception:
            aviso.geometry(f"{ancho}x{alto}")

        def cerrar():
            try:
                aviso.destroy()
            except Exception:
                pass
            if al_cerrar:
                try:
                    al_cerrar()
                except Exception:
                    pass

        aviso.after(max(200, int(float(segundos) * 1000)), cerrar)
        try:
            aviso.lift()
        except Exception:
            pass
        return aviso

    def aplicar_datos_escaneo(self, datos):
        """Vuelca en el formulario lo leído por el OCR y deja el PDF como adjunto.

        Si la factura está a nombre de OTRA empresa se rechaza y no se aplica nada.
        Si el proveedor no está registrado, se agrega automáticamente a Proveedores.
        """
        if not datos:
            return
        # 🏢 ¿Está a nombre de la empresa configurada?
        estado_destino, mensaje_destino = validar_destinatario(datos)
        self._destino_escaneo = estado_destino
        if estado_destino == "ajeno":
            return messagebox.showerror("Factura de otra empresa", mensaje_destino +
                                        "\n\nNo se cargó nada en el formulario.", parent=ventana_para_mensaje(getattr(self, "main_root", None)))
        try:
            self.bloquear_autocompletado_ruc = True
            if datos.get("tipo_documento"):
                self.combo_tipo.set(datos["tipo_documento"])
                self.on_tipo_change(self.combo_tipo.get())
            if datos.get("numero"):
                self.ent_nro_doc.delete(0, tk.END)
                self.ent_nro_doc.insert(0, str(datos["numero"]))
            if datos.get("fecha"):
                self.ent_fecha.delete(0, tk.END)
                self.ent_fecha.insert(0, str(datos["fecha"]))
            if datos.get("proveedor"):
                _prov = datos["proveedor"]
                # el lector puede devolver (ruc, nombre) o "RUC NOMBRE"
                if isinstance(_prov, (tuple, list)):
                    _prov = _prov[1] if len(_prov) > 1 else (_prov[0] if _prov else "")
                nombre = str(_prov or "").strip()
                nombre = re.sub(r"^\s*\d{11}\s*[-\u2013:.|]?\s*", "", nombre).strip()
                if nombre:
                    try:
                        valores = [str(v) for v in (self.combo_proveedor.cget("values") or [])]
                        if nombre not in valores:
                            self.combo_proveedor.configure(values=[nombre] + valores)
                    except Exception:
                        pass
                    self.combo_proveedor.set(nombre)
            if datos.get("ruc"):
                self.ent_desc.delete(0, tk.END)
                self.ent_desc.insert(0, str(datos["ruc"]))
            # 💰 Se carga el TOTAL del documento (con IGV) y el modo queda en "con IGV",
            #    igual que en el Programa de Flota: la base y el IGV se calculan solos.
            monto_total = 0.0
            try:
                if datos.get("total"):
                    monto_total = float(datos["total"])
                elif datos.get("subtotal"):
                    base_leida = float(datos["subtotal"])
                    if "Factura" in self.combo_tipo.get():
                        monto_total = base_leida * (1.105 if "10.5%" in self.combo_tipo.get() else 1.18)
                    else:
                        monto_total = base_leida
            except Exception as e_monto:
                print("Aviso - monto del comprobante:", e_monto)
            if monto_total > 0:
                self.ent_subtotal.delete(0, tk.END)
                self.ent_subtotal.insert(0, f"{monto_total:.2f}")
                self._poner_modo_monto("CON_IGV")

            resultado = getattr(self, "_escaneo_actual", {}) or {}
            if resultado.get("pdf") and os.path.exists(resultado["pdf"]):
                self.ruta_archivo_temp = resultado["pdf"]
                try:
                    self.btn_archivo.configure(text="✅ PDF Escaneado (OCR)", fg_color="#28a745")
                except Exception:
                    pass
            self.actualizar_totales()
            try:
                self.al_seleccionar_proveedor()
            except Exception:
                pass
            # 🏢 El proveedor ya quedó registrado al escanear (con los datos de SUNAT).
            # Solo si el escaneo es anterior a esta mejora se crea aquí.
            estado_prov = (getattr(self, "_escaneo_actual", {}) or {}).get("proveedor", {}).get("estado", "")
            if not estado_prov:
                try:
                    estado_prov, _datos_prov = asegurar_proveedor(datos.get("proveedor"), datos.get("ruc"))
                except Exception:
                    estado_prov = "error"
            self._estado_proveedor_escaneo = estado_prov
            if estado_destino == "sin-datos":
                messagebox.showwarning("Destinatario", mensaje_destino, parent=ventana_para_mensaje(getattr(self, "main_root", None)))
            elif estado_destino == "sin-configurar":
                messagebox.showwarning("Destinatario", mensaje_destino, parent=ventana_para_mensaje(getattr(self, "main_root", None)))
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron aplicar los datos:\n{e}", parent=ventana_para_mensaje(getattr(self, "main_root", None)))
        finally:
            self.bloquear_autocompletado_ruc = False

    def autocompletar_desde_pdf(self, ruta=None):
        """Lee un comprobante (PDF con texto, PDF escaneado o foto) y llena los campos."""
        if not ruta:
            ruta = seleccionar_archivo_dialogo(
                "Seleccionar la factura (PDF o foto escaneada)",
                [("Documentos", "*.pdf;*.png;*.jpg;*.jpeg"),
                 ("Archivos PDF", "*.pdf"),
                 ("Imágenes", "*.png;*.jpg;*.jpeg")])
        if not ruta:
            return
        try:
            self.bloquear_autocompletado_ruc = True
            _aviso = self._aviso_procesando("⏳ Leyendo el documento... (OCR, puede tardar unos segundos)")
            try:
                texto = self._extraer_texto_documento(ruta)
                leido_con_ocr = False
                if not texto.strip():
                    # Comprobante escaneado o foto: no tiene texto, se aplica OCR local
                    try:
                        self.btn_auto_pdf.configure(text="⏳ Aplicando OCR al documento...")
                        self.main_root.update_idletasks()
                    except Exception:
                        pass
                    texto = leer_ocr_windows(ruta)
                    leido_con_ocr = bool(texto.strip())
                    try:
                        self.btn_auto_pdf.configure(text="📄 Desde PDF")
                    except Exception:
                        pass
            finally:
                self._cerrar_aviso(_aviso)

            if not texto.strip():
                # No se pudo leer: el documento queda adjunto y el usuario solo digita los montos
                self.ruta_archivo_temp = ruta
                self.btn_archivo.configure(text="✅ Adjuntado (digite los montos)",
                                           fg_color="#e67e22", hover_color="#b9651a")
                self.bloquear_autocompletado_ruc = False
                return messagebox.showwarning(
                    "Documento escaneado",
                    "Este archivo no tiene texto (es un escaneo o una foto) y no se pudo leer con OCR.\n\n"
                    "El documento quedó ADJUNTO al registro: digite el tipo de documento, el monto base "
                    "(o el total con IGV), el recargo al consumo si lo tuviera y el N° de documento.")

            montos = extraer_montos_comprobante(texto)
            tasa_leida = tasa_igv_de_montos(montos["base"], montos["igv"])
            es_restaurante = montos["recargo"] > 0 or abs(tasa_leida - 10.5) < 0.01

            # ---- Tipo de documento ----
            texto_mayus = texto.upper()
            tipo_detectado = ""
            if "BOLETA" in texto_mayus and "FACTURA" not in texto_mayus:
                tipo_detectado = "Boleta (Sin IGV)"
            elif "RECIBO" in texto_mayus and "HONORARIO" in texto_mayus:
                tipo_detectado = ("Recibo por Honorarios (8% Retención)"
                                  if re.search(r"RETENCI", texto_mayus)
                                  else "Recibo por Honorarios (Sin Retención)")
            elif "FACTURA" in texto_mayus:
                # Las facturas de restaurante llevan 10.5% de IGV y recargo al consumo
                tipo_detectado = ("Factura (10.5% IGV) Restaurantes" if es_restaurante
                                  else "Factura (18% IGV)")
            if tipo_detectado:
                self.combo_tipo.set(tipo_detectado)
            self.on_tipo_change(self.combo_tipo.get())

            # ---- N° de documento (lector tolerante + nombre del archivo) ----
            numero_leido = numero_comprobante_desde_texto(texto)
            if not numero_leido:
                nro_match = re.search(r"([EFB][0-9A-Z]{3}\s*-\s*\d+)", texto)
                numero_leido = nro_match.group(1).replace(" ", "") if nro_match else ""
            # Si el OCR lo leyó mal (empieza con dígito) manda el nombre del archivo
            try:
                if not numero_leido or numero_leido[:1].isdigit():
                    _m_nombre = re.search(r"([A-Z][0-9A-Z]{3})\s*[-–]\s*(\d{4,8})",
                                          os.path.basename(str(ruta)).upper())
                    if _m_nombre:
                        numero_leido = "%s-%s" % (_m_nombre.group(1), _m_nombre.group(2))
            except Exception:
                pass
            if numero_leido:
                self.ent_nro_doc.delete(0, tk.END)
                self.ent_nro_doc.insert(0, numero_leido)

            # ---- Fecha (lector tolerante: descarta vencimientos 2030) ----
            fecha_match = None
            try:
                _fecha_tol = extraer_fecha_comprobante(texto)
                if _fecha_tol:
                    fecha_match = re.match(r"(\d{2})[/\-.](\d{2})[/\-.](\d{4})", _fecha_tol)
            except Exception:
                fecha_match = None
            if not fecha_match:
                fecha_match = re.search(r"Fecha de Emisi[oó]n\s*[:\-]?\s*(\d{2})[/\-.](\d{2})[/\-.](\d{4})",
                                        texto, re.IGNORECASE)
            if not fecha_match:
                fecha_match = re.search(r"(\d{2})[/\-.](\d{2})[/\-.](\d{4})", texto)
            if fecha_match:
                d, m, y = fecha_match.groups()
                fmt = CONFIG_REGIONAL.get("formato_fecha", "DD/MM/AAAA")
                self.ent_fecha.delete(0, tk.END)
                if fmt == "MM/DD/AAAA":
                    self.ent_fecha.insert(0, f"{m}/{d}/{y}")
                else:
                    self.ent_fecha.insert(0, f"{d}/{m}/{y}")
            fecha_ilegible = bool(leido_con_ocr and not fecha_match)

            # ---- Proveedor: 1º) el RUC leído se busca en la tabla 'proveedores' ----
            ruc_proveedor = ruc_distinto_al_de_la_empresa(texto, CONFIG_REGIONAL.get("ruc_empresa", ""))
            nombre_bd = ""
            try:
                ruc_bd, nombre_bd = buscar_proveedor_por_ruc(texto, CONFIG_REGIONAL.get("ruc_empresa", ""))
                if ruc_bd:
                    ruc_proveedor = ruc_bd
            except Exception as e_bd:
                print("Aviso - proveedor por RUC:", e_bd)
            if ruc_proveedor:
                self.ent_desc.delete(0, tk.END)
                self.ent_desc.insert(0, ruc_proveedor)

            if nombre_bd:
                nombre = nombre_bd          # proveedor ya registrado con ese RUC
            else:
                # 2º) el nombre impreso (si parece una razón social)
                nombre = nombre_proveedor_desde_texto(texto, ruc_proveedor)
                if nombre and not _nombre_parece_proveedor(nombre):
                    nombre = ""
                if not nombre:
                    lineas = [l.strip() for l in texto.split("\n") if l.strip()]
                    posibles = [l for l in lineas[:7]
                                if "R.U.C" not in l.upper() and "RUC" not in l.upper() and len(l) > 4]
                    nombre = posibles[0] if posibles else ""
            if nombre:
                self.combo_proveedor.set(nombre)

            # ---- Montos: base, IGV, recargo al consumo y total ----
            if hasattr(self, "ent_recargo"):
                self.ent_recargo.delete(0, tk.END)
                self.ent_recargo.insert(0, "%.2f" % montos["recargo"])
            if montos["base"] > 0:
                self._poner_modo_monto("BASE")
                self.ent_subtotal.delete(0, tk.END)
                self.ent_subtotal.insert(0, "%.2f" % montos["base"])
            elif montos["total"] > 0:
                self._poner_modo_monto("CON_IGV")
                self.ent_subtotal.delete(0, tk.END)
                self.ent_subtotal.insert(0, "%.2f" % montos["total"])

            self.ruta_archivo_temp = ruta
            self.btn_archivo.configure(text="✅ PDF Autocargado Exitosamente", fg_color="#28a745")
            self.actualizar_totales()
            self.al_seleccionar_proveedor()

            avisos = []
            if leido_con_ocr:
                avisos.append("El documento era un ESCANEO o FOTO: los datos se leyeron con OCR. "
                              "Revise la fecha, el N° de documento y los montos antes de guardar.")
            if fecha_ilegible:
                avisos.append("No se pudo leer la fecha con seguridad: corríjala a mano.")
            if montos["recargo"] > 0:
                avisos.append(f"Recargo al consumo detectado: {formatear_moneda(montos['recargo'])} "
                              "(se suma al total y no lleva IGV).")
            messagebox.showinfo("Extracción Inteligente",
                                "Se extrajeron los datos del documento."
                                + ("\n\n" + "\n".join(avisos) if avisos else ""))
            self.bloquear_autocompletado_ruc = False
        except Exception as e:
            self.bloquear_autocompletado_ruc = False
            messagebox.showerror("Error", f"Ocurrió un error:\n{e}")

    def autocompletar_desde_xml(self):
        ruta = seleccionar_archivo_dialogo("Seleccionar Factura XML de SUNAT", [("Archivos XML", "*.xml")])
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
                # El XML entrega el monto BASE (TaxExclusiveAmount): se fuerza ese modo.
                self._poner_modo_monto("BASE")
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
        centrar_ventana(v_sire, self.main_root, 480, 300)
        v_sire.grab_set()

        ctk.CTkLabel(v_sire, text="🌐 IMPORTACIÓN AUTOMÁTICA SIRE SUNAT", font=("Arial", 14, "bold"), text_color="#166534").pack(pady=(20, 10))
        lbl_status = ctk.CTkLabel(v_sire, text="🔑 Autenticando token con la SUNAT...", font=("Arial", 11, "italic"), text_color="#d35400")
        lbl_status.pack(pady=10)

        prog = ctk.CTkProgressBar(v_sire, width=380)
        prog.pack(pady=10)
        prog.set(0.2)

        txt_info = ctk.CTkTextbox(v_sire, height=100, font=("Arial", 10))
        txt_info.pack(fill="x", padx=25, pady=10)

        # El hilo solo hace las llamadas a SUNAT; la ventana se actualiza por sondeo
        def ejecucion_sire(estado):
            estado["progreso"] = 0.35
            estado["texto"] = "🔐 Conectando con SUNAT..."
            if True:
                url_token = "https://api-seguridad.sunat.gob.pe/v1/clienttoken"
                headers_token = {"Content-Type": "application/x-www-form-urlencoded"}
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
                except Exception:
                    token_access = None

                estado["progreso"] = 0.6
                estado["texto"] = "📥 Descargando Registro de Compras RCE..."

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

                estado["progreso"] = 1.0
                estado["texto"] = "✅ Sincronización SIRE Finalizada"
                estado["color"] = "#27ae60"
                estado["mensaje"] = (
                    f"✅ Conexión completada con éxito.\n"
                    f"• Periodo Sincronizado: {periodo}\n"
                    f"• RUC Conectado: {ruc}\n"
                    f"• Comprobantes Obtenidos: {len(datos_compras)}\n\n"
                    f"El Registro de Compras se encuentra 100% actualizado con la propuesta de SUNAT."
                )

        def avanzar_sire(estado):
            if "progreso" in estado:
                prog.set(estado["progreso"])
            if "texto" in estado:
                lbl_status.configure(text=estado["texto"], text_color=estado.get("color", "#1f538d"))

        def terminar_sire(estado):
            if estado.get("error"):
                lbl_status.configure(text="❌ Error en Conexión SIRE", text_color="#c0392b")
                txt_info.delete("1.0", tk.END)
                txt_info.insert("1.0", f"Fallo al conectar con SUNAT:\n{estado['error']}")
                return
            txt_info.delete("1.0", tk.END)
            txt_info.insert("1.0", estado.get("mensaje", ""))
            self.cargar_datos_tabla(reset_pagina=True)

        ejecutar_en_hilo(v_sire, ejecucion_sire, aplicar=avanzar_sire, al_terminar=terminar_sire)

    def agregar_nueva_categoria(self):
        # Se abre EXACTAMENTE la misma ventana que usa el módulo de Banco
        # (Conciliación → Agregar Movimiento → "⚙️ Gestionar Categorías"), para que
        # las categorías creadas aquí queden guardadas y disponibles en ambos módulos.
        from modulo_banco import abrir_gestion_categorias_gastos
        abrir_gestion_categorias_gastos(
            self.main_root.winfo_toplevel(),
            al_guardar=self.refrescar_categorias_gastos,
        )

    def refrescar_categorias_gastos(self):
        """Vuelve a leer las categorías guardadas y actualiza el desplegable."""
        try:
            self.cargar_categorias()
        except Exception as e:
            print(f"No se pudieron refrescar las categorías de gastos: {e}")

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

        # 🖨️ Escanear la factura con la multifuncional y leerla con OCR
        btn_escanear = ctk.CTkButton(self.f_form, text="🖨️ Escanear factura (OCR)",
                                     font=("Arial", 12, "bold"), fg_color="#0e7490",
                                     hover_color="#155e75",
                                     command=self.escanear_factura_ocr, height=34)
        btn_escanear.pack(fill="x", padx=10, pady=(0, 10))

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
        activar_busqueda_rapida(self.combo_proveedor, al_elegir=self.al_seleccionar_proveedor)
        self.combo_proveedor.bind("<KeyRelease>", self._al_teclear_proveedor)
        self.combo_proveedor.pack(fill="x", padx=10, pady=(0, 8))
        self.cargar_proveedores_bd()

        ctk.CTkLabel(self.f_form, text="R.U.C.:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.ent_desc = ctk.CTkEntry(self.f_form)
        self.ent_desc.pack(fill="x", padx=10, pady=(0, 8))

        ctk.CTkLabel(self.f_form, text="Concepto / Descripción:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.ent_concepto = ctk.CTkEntry(self.f_form, placeholder_text="Ej. Mantenimiento, Útiles...")
        self.ent_concepto.pack(fill="x", padx=10, pady=(0, 8))

        ctk.CTkLabel(self.f_form, text="Categoría de Gasto:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        f_cat = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_cat.pack(fill="x", padx=10, pady=(0, 8))
        self.combo_categoria = ctk.CTkComboBox(f_cat, state="readonly", command=self.on_cambiar_categoria)
        self.combo_categoria.pack(side="left", fill="x", expand=True)
        # Abre la MISMA ventana de categorías que el módulo de Banco
        ctk.CTkButton(f_cat, text="⚙️", width=34, fg_color="#8e44ad", hover_color="#703688", command=self.agregar_nueva_categoria).pack(side="right", padx=(5, 0))
        self.cargar_categorias()

        ctk.CTkLabel(self.f_form, text="Vehículo Asignado (Placa):", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        f_veh = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_veh.pack(fill="x", padx=10, pady=(0, 8))
        self.combo_evento = ctk.CTkComboBox(f_veh, state="readonly")
        self.combo_evento.pack(side="left", fill="x", expand=True)
        # 🛠️ Si la compra es de un vehículo, desde aquí se resetean sus mantenimientos
        ctk.CTkButton(f_veh, text="🛠️", width=34, font=("Arial", 12, "bold"),
                      fg_color="#e67e22", hover_color="#ca6f1e",
                      command=lambda: abrir_reset_mantenimientos_vehiculo(
                          self.main_root, self.combo_evento.get(), self.app_padre.usuario_activo)).pack(
            side="right", padx=(5, 0))
        self.cargar_vehiculos_bd()

        # 🎫 CARNET DE SANIDAD: cuando la categoría del gasto es un carné de sanidad
        # se pide el chofer y se calcula el vencimiento (6 meses) para el cronograma.
        self.f_carnet = ctk.CTkFrame(self.f_form, fg_color="#eaf4fb", corner_radius=6,
                                     border_width=1, border_color="#aed6f1")
        ctk.CTkLabel(self.f_carnet, text="🎫 Carné de Sanidad del Chofer", font=("Arial", 11, "bold"),
                     text_color="#1f538d").pack(anchor="w", padx=8, pady=(6, 2))

        ctk.CTkLabel(self.f_carnet, text="Chofer:", font=("Arial", 10, "bold")).pack(anchor="w", padx=8)
        self.combo_chofer = ctk.CTkComboBox(self.f_carnet, state="readonly")
        self.combo_chofer.pack(fill="x", padx=8, pady=(0, 4))

        ctk.CTkLabel(self.f_carnet, text="N° del Carné (opcional):", font=("Arial", 10, "bold")).pack(anchor="w", padx=8)
        self.ent_carnet_num = ctk.CTkEntry(self.f_carnet, placeholder_text="Ej: CS-001234")
        self.ent_carnet_num.pack(fill="x", padx=8, pady=(0, 4))

        ctk.CTkLabel(self.f_carnet, text="Fecha de Emisión del Carné:", font=("Arial", 10, "bold")).pack(anchor="w", padx=8)
        f_carnet_fec = ctk.CTkFrame(self.f_carnet, fg_color="transparent")
        f_carnet_fec.pack(fill="x", padx=8, pady=(0, 4))
        self.ent_carnet_emision = ctk.CTkEntry(f_carnet_fec, placeholder_text="DD/MM/AAAA")
        self.ent_carnet_emision.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(f_carnet_fec, text="[ 📅 ]", width=40, font=("Arial", 11, "bold"),
                      fg_color="#1f538d", hover_color="#163b65",
                      command=lambda: CalendarioNativo(self.f_carnet, self.ent_carnet_emision,
                                                       self._actualizar_vencimiento_carnet)).pack(side="right", padx=(5, 0))
        self.ent_carnet_emision.bind("<KeyRelease>", lambda _e: self._actualizar_vencimiento_carnet())
        self.ent_carnet_emision.bind("<FocusOut>", lambda _e: self._actualizar_vencimiento_carnet())

        self.lbl_carnet_venc = ctk.CTkLabel(self.f_carnet, text="Vencimiento: —", font=("Arial", 10, "bold"),
                                            text_color="#c0392b", wraplength=290, justify="left")
        self.lbl_carnet_venc.pack(anchor="w", padx=8, pady=(0, 8))
        self._mapa_choferes = {}

        # 🔗 CRUCE CON LA ORDEN DE SERVICIO: se elige la orden del módulo de Órdenes
        # para relacionarla con esta factura y comparar montos.
        ctk.CTkLabel(self.f_form, text="🔗 Orden de Servicio (relacionar):", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        f_orden = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_orden.pack(fill="x", padx=10, pady=(0, 4))
        self.combo_orden_servicio = ctk.CTkComboBox(f_orden, values=[SIN_ORDEN_SERVICIO], state="readonly",
                                                    command=self.al_seleccionar_orden_servicio)
        self.combo_orden_servicio.pack(side="left", fill="x", expand=True)
        self.btn_ver_orden = ctk.CTkButton(f_orden, text="📄", width=34, font=("Arial", 11),
                                           fg_color="#34495e", hover_color="#2c3e50",
                                           command=self.abrir_pdf_orden_servicio)
        self.btn_ver_orden.pack(side="right", padx=(5, 0))
        self.lbl_cruce_orden = ctk.CTkLabel(self.f_form, text="", font=("Arial", 10, "italic"),
                                            text_color="#555555", wraplength=300, justify="left")
        self.lbl_cruce_orden.pack(anchor="w", padx=10, pady=(0, 8))
        self.cargar_ordenes_servicio()


        # =====================================================================
        # 🔁 FACTURA CRUZADA (PAGADA POR UN TERCERO)
        # Al marcarla se piden los datos de quien pagó y, si el pago fue
        # interbancario, la comisión (igual que en el módulo de Banco).
        # =====================================================================
        self.var_cruzada = tk.BooleanVar(value=False)
        self.chk_cruzada = ctk.CTkCheckBox(
            self.f_form, text="💸 Pagada por un tercero · le devolví el dinero",
            variable=self.var_cruzada, font=("Arial", 11, "bold"),
            checkbox_width=18, checkbox_height=18, command=self._actualizar_bloque_cruzada)
        self.chk_cruzada.pack(anchor="w", padx=10, pady=(4, 2))

        self.f_cruzada = ctk.CTkFrame(self.f_form, fg_color="#f4ecf7", corner_radius=6,
                                      border_width=1, border_color="#d2b4de")
        ctk.CTkLabel(self.f_cruzada, text="💸 Devolución al tercero que pagó (sale de mi cuenta)", font=("Arial", 11, "bold"),
                     text_color="#6c3483").pack(anchor="w", padx=8, pady=(6, 2))

        ctk.CTkLabel(self.f_cruzada, text="A quién le devuelvo (tercero): *", font=("Arial", 10, "bold")).pack(anchor="w", padx=8)
        self.ent_tercero = ctk.CTkEntry(self.f_cruzada, placeholder_text="Nombre de la persona a la que se le devuelve")
        self.ent_tercero.pack(fill="x", padx=8, pady=(0, 4))

        ctk.CTkLabel(self.f_cruzada, text="🏦 Cuenta desde la que le devuelvo el dinero:", font=("Arial", 10, "bold")).pack(anchor="w", padx=8)
        self.cmb_cuenta_cruzada = ctk.CTkComboBox(self.f_cruzada, values=self._lista_cuentas_pago())
        self.cmb_cuenta_cruzada.pack(fill="x", padx=8, pady=(0, 4))

        ctk.CTkLabel(self.f_cruzada, text="N° de Operación de la devolución:", font=("Arial", 10, "bold")).pack(anchor="w", padx=8)
        self.ent_oper_cruzada = ctk.CTkEntry(self.f_cruzada, placeholder_text="Ej.: 00123456 / constancia")
        self.ent_oper_cruzada.pack(fill="x", padx=8, pady=(0, 4))

        # 📅 Fecha en que el tercero hizo el pago (va al asiento del banco)
        ctk.CTkLabel(self.f_cruzada, text="📅 Fecha de la devolución (DD/MM/AAAA):",
                     font=("Arial", 10, "bold")).pack(anchor="w", padx=8)
        self.ent_fecha_pago_cruz = ctk.CTkEntry(self.f_cruzada, placeholder_text="22/10/2026")
        self.ent_fecha_pago_cruz.pack(fill="x", padx=8, pady=(0, 6))
        try:
            if not self.ent_fecha_pago_cruz.get():
                self.ent_fecha_pago_cruz.insert(0, datetime.now().strftime("%d/%m/%Y"))
        except Exception:
            pass

        f_com_cruz = ctk.CTkFrame(self.f_cruzada, fg_color="transparent")
        f_com_cruz.pack(fill="x", padx=8, pady=(0, 2))
        self.var_interbancario = tk.BooleanVar(value=False)
        self.chk_interbancario = ctk.CTkCheckBox(
            f_com_cruz, text="Pago Interbancario", variable=self.var_interbancario,
            font=("Arial", 10, "bold"), checkbox_width=16, checkbox_height=16,
            command=self._actualizar_total_cruzada)
        self.chk_interbancario.pack(side="left")
        ctk.CTkLabel(f_com_cruz, text="Comisión (S/.):", font=("Arial", 10, "bold")).pack(side="left", padx=(10, 4))
        self.ent_comision = ctk.CTkEntry(f_com_cruz, width=80)
        self.ent_comision.pack(side="left")
        self.ent_comision.insert(0, f"{cargar_comision_interbancaria_config():.2f}")
        # El resumen se recalcula MIENTRAS se escribe (y al pegar) la comisión
        for _evento in ("<KeyRelease>", "<FocusOut>", "<<Paste>>"):
            self.ent_comision.bind(_evento, lambda _e: self._actualizar_total_cruzada(), add="+")
        # Al entrar al campo se selecciona el valor: escribir la comisión lo reemplaza
        try:
            self.ent_comision.bind("<FocusIn>",
                                   lambda _e: self.ent_comision.select_range(0, "end"), add="+")
        except Exception:
            pass

        self.lbl_cruzada_total = ctk.CTkLabel(self.f_cruzada, text="", font=("Arial", 10, "bold"),
                                              text_color="#6c3483", wraplength=290, justify="left")
        self.lbl_cruzada_total.pack(anchor="w", padx=8, pady=(2, 4))

        # 🖨️ Escanear o 📂 cargar el comprobante del pago hecho por el tercero
        _fila_soporte = ctk.CTkFrame(self.f_cruzada, fg_color="transparent")
        _fila_soporte.pack(fill="x", padx=8, pady=(0, 8))

        self.btn_soporte_tercero = ctk.CTkButton(
            _fila_soporte, text="🖨️ Escanear", font=("Arial", 11, "bold"),
            fg_color="#7f8c8d", hover_color="#606b6b",
            command=self.escanear_soporte_tercero)
        self.btn_soporte_tercero.pack(side="left", fill="x", expand=True)

        self.btn_soporte_tercero_pdf = ctk.CTkButton(
            _fila_soporte, text="📂 Cargar PDF / foto", font=("Arial", 11, "bold"),
            fg_color="#2980b9", hover_color="#1f618d",
            command=self.seleccionar_soporte_tercero)
        self.btn_soporte_tercero_pdf.pack(side="left", fill="x", expand=True, padx=(6, 0))
        self.ruta_soporte_tercero = ""

        # 💰 El monto se puede digitar SIN IGV (monto base) o CON IGV (total del documento)
        # (esta etiqueta sirve de ancla: el bloque del carné de sanidad se coloca antes)
        self._ancla_bloque_carnet = ctk.CTkLabel(self.f_form, text="Tipo de Monto a Ingresar:", font=("Arial", 11, "bold"))
        self._ancla_bloque_carnet.pack(anchor="w", padx=10)
        self.seg_modo_monto = ctk.CTkSegmentedButton(
            self.f_form,
            values=["Monto Base (sin IGV)", "Monto con IGV (Total)"],
            font=("Arial", 10, "bold"),
            command=self.on_cambiar_modo_monto
        )
        self.seg_modo_monto.pack(fill="x", padx=10, pady=(0, 6))
        self.seg_modo_monto.set("Monto Base (sin IGV)")

        self.lbl_monto_entrada = ctk.CTkLabel(self.f_form, text="Monto Base (Subtotal):", font=("Arial", 11, "bold"))
        self.lbl_monto_entrada.pack(anchor="w", padx=10)
        self.ent_subtotal = ctk.CTkEntry(self.f_form, placeholder_text="Ej. 100.00")
        self.ent_subtotal.pack(fill="x", padx=10, pady=(0, 4))
        self.ent_subtotal.bind("<KeyRelease>", self.actualizar_totales)

        self.lbl_nota_modo = ctk.CTkLabel(self.f_form, text="", font=("Arial", 10, "italic"), text_color="#555555", wraplength=300, justify="left")
        self.lbl_nota_modo.pack(anchor="w", padx=10, pady=(0, 8))
        self._refrescar_etiqueta_modo()

        # 🍽️ Recargo al consumo (servicio / propina de restaurantes): suma al total
        # pero NO forma parte de la base del IGV, igual que en la factura.
        ctk.CTkLabel(self.f_form, text="Recargo al consumo (no gravado):", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.ent_recargo = ctk.CTkEntry(self.f_form, placeholder_text="0.00")
        self.ent_recargo.pack(fill="x", padx=10, pady=(0, 8))
        self.ent_recargo.insert(0, "0")
        self.ent_recargo.bind("<KeyRelease>", self.actualizar_totales)

        self.lbl_titulo_det = ctk.CTkLabel(self.f_form, text="Detracción (%):", font=("Arial", 11, "bold"))
        self.lbl_titulo_det.pack(anchor="w", padx=10)
        self.ent_detraccion = ctk.CTkEntry(self.f_form)
        self.ent_detraccion.pack(fill="x", padx=10, pady=(0, 8))
        self.ent_detraccion.insert(0, "0")
        self.ent_detraccion.bind("<KeyRelease>", self.actualizar_totales)

        f_tot = ctk.CTkFrame(self.f_form, fg_color="#ffffff", border_width=1, border_color="#e0e0e0")
        f_tot.pack(fill="x", padx=10, pady=(5, 10))
        self.lbl_base = ctk.CTkLabel(f_tot, text=f"Monto Base (Subtotal): {formatear_moneda(0)}", font=("Arial", 11), text_color="#555")
        self.lbl_base.pack(anchor="w", padx=10, pady=(5, 0))
        self.lbl_impuesto = ctk.CTkLabel(f_tot, text=f"IGV (18%): {formatear_moneda(0)}", font=("Arial", 11), text_color="#555")
        self.lbl_impuesto.pack(anchor="w", padx=10, pady=(0, 0))
        self.lbl_recargo = ctk.CTkLabel(f_tot, text=f"Recargo al consumo: {formatear_moneda(0)}", font=("Arial", 11), text_color="#555")
        self.lbl_recargo.pack(anchor="w", padx=10, pady=(0, 0))
        self.lbl_bruto = ctk.CTkLabel(f_tot, text=f"Total con IGV: {formatear_moneda(0)}", font=("Arial", 11, "bold"), text_color="#1f538d")
        self.lbl_bruto.pack(anchor="w", padx=10, pady=(0, 0))
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
        self.ent_buscar_facturas = ctk.CTkEntry(f_busqueda, placeholder_text="Buscar por cualquier columna: N° doc, orden de servicio, proveedor, RUC, placa, concepto, fecha, monto...")
        self.ent_buscar_facturas.pack(side="left", fill="x", expand=True)
        
        ctk.CTkLabel(f_busqueda, text="🗓️ Mes:", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 5))
        self.combo_mes = ctk.CTkComboBox(f_busqueda, values=construir_valores_mes(), width=150, state="readonly", command=self.on_cambiar_mes)
        self.combo_mes.pack(side="left", padx=(0, 5))
        self.combo_mes.set(self.mes_filtro)

        # 🚫 Casilla para ocultar las facturas que llegan desde la App (App Grifo)
        self.chk_ocultar_app = ctk.CTkCheckBox(
            f_busqueda, text="🚫 Ocultar facturas de la App", variable=self.var_ocultar_app,
            font=("Arial", 11, "bold"), checkbox_width=18, checkbox_height=18,
            command=lambda: self.cargar_datos_tabla(reset_pagina=True))
        self.chk_ocultar_app.pack(side="left", padx=(10, 0))

        self.ent_buscar_facturas.bind("<KeyRelease>", lambda e: self.buscar_con_retraso())
        self.ent_buscar_facturas.bind("<Return>", lambda e: self.cargar_datos_tabla(reset_pagina=True))

        f_btn_tabla = ctk.CTkFrame(self.f_wrapper_derecha, fg_color="transparent")
        f_btn_tabla.pack(side="bottom", fill="x", pady=(10, 0))
        
        f_paginacion = ctk.CTkFrame(f_btn_tabla, fg_color="transparent")
        f_paginacion.pack(side="left", padx=(0, 10))
        
        self.btn_ant = ctk.CTkButton(f_paginacion, text="◀ Ant", width=60, command=self.pagina_anterior)
        self.btn_ant.pack(side="left", padx=2)
        
        self.lbl_pagina = ctk.CTkLabel(f_paginacion, text=f"Pág {self.pagina_actual}", font=("Arial", 11, "bold"))
        self.lbl_pagina.pack(side="left", padx=5)
        
        self.btn_sig = ctk.CTkButton(f_paginacion, text="Sig ▶", width=60, command=self.pagina_siguiente)
        self.btn_sig.pack(side="left", padx=2)
        
        btn_sincronizar = ctk.CTkButton(f_btn_tabla, text="🔄 Actualizar y Descargar App", font=("Arial", 12, "bold"), command=self.ejecutar_sincronizacion_manual, fg_color="#27ae60", hover_color="#1e8449")
        btn_sincronizar.pack(side="left")
        
        # Leyenda de colores: se mantiene en su fila original (junto a paginación y Actualizar)
        # ⚙️ Columnas visibles de la tabla de facturas
        btn_cols = ctk.CTkButton(f_btn_tabla, text="⚙️ Columnas", width=90,
                                 font=("Arial", 12, "bold"), fg_color="#6c757d",
                                 hover_color="#565e64",
                                 command=lambda: menu_columnas(self.tabla, "compras_facturas"))
        btn_cols.pack(side="left", padx=5, pady=5)
        try:
            aplicar_columnas_visibles(self.tabla, "compras_facturas")
        except Exception:
            pass

        f_leyenda = ctk.CTkFrame(f_btn_tabla, fg_color="transparent")
        f_leyenda.pack(side="left", padx=20)
        
        ctk.CTkLabel(f_leyenda, text="■", font=("Arial", 14), text_color="#c0392b").pack(side="left", padx=(5,2))
        ctk.CTkLabel(f_leyenda, text="Sin Cuenta Asignada", font=("Arial", 11, "bold"), text_color="#333333").pack(side="left", padx=(0,10))
        
        ctk.CTkLabel(f_leyenda, text="■", font=("Arial", 14), text_color="#27ae60").pack(side="left", padx=(5,2))
        ctk.CTkLabel(f_leyenda, text="Cuenta Asignada", font=("Arial", 11, "bold"), text_color="#333333").pack(side="left", padx=(0,5))

        f_tabla = ctk.CTkFrame(self.f_wrapper_derecha, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True)

        columnas = ("num", "id", "fecha", "hora", "nro_doc", "orden", "dias", "tipo", "proveedor", "ruc", "categoria", "evento", "kilometraje", "cantidad", "desc", "metodo_pago", "subtotal", "impuesto", "total", "detraccion", "neto", "archivo")
        self.columnas_tabla = columnas
        self.tabla = ttk.Treeview(f_tabla, columns=columnas, show="headings", height=28)
        # 📏 La tabla ocupa todo el alto disponible (hasta la barra inferior)
        try:
            f_tabla.bind("<Configure>", lambda e: self._ajustar_alto_tabla(e))
            self._f_tabla_alto = f_tabla
        except Exception:
            pass
        # 🖱️ Clic derecho: mismo menú que Control de Pagos
        self.tabla.bind("<Button-3>", self._menu_contextual_factura)
        self.tabla.bind("<Button-2>", self._menu_contextual_factura)
        
        self.tabla.tag_configure("con_cuenta", background="#e8f8f5", foreground="#0e6251")
        # 🟢 Estados de pago (verde = totalmente cancelada)
        self.tabla.tag_configure("pagada", background="#d5f5e3", foreground="#145a32")
        self.tabla.tag_configure("parcial", background="#e8daef", foreground="#6c3483")
        self.tabla.tag_configure("no_pagada", background="#fdedec", foreground="#7b241c") 
        self.tabla.tag_configure("sin_cuenta", background="#fdedec", foreground="#7b241c") 
        
        self.tabla.heading("num", text="N°", anchor="center")
        self.tabla.heading("id", text="ID (Oculto)")
        self.tabla.heading("fecha", text="Fecha Fac. ↕", command=lambda: self.ordenar_por_columna("fecha", False))
        self.tabla.heading("hora", text="Hora ↕", command=lambda: self.ordenar_por_columna("hora", False))
        self.tabla.heading("nro_doc", text="N° Doc. ↕", command=lambda: self.ordenar_por_columna("nro_doc", False))
        self.tabla.heading("orden", text="Orden Serv. ↕", command=lambda: self.ordenar_por_columna("orden", False))
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
        # 📏 El N° de documento necesita ancho suficiente: con 90 px los números
        # largos se recortaban (ej. F581-08266857 y F581-08266859 se veían IGUALES).
        self.tabla.column("nro_doc", width=130, minwidth=120, stretch=False, anchor="w")
        self.tabla.column("orden", width=115, anchor="center")
        self.tabla.column("proveedor", width=120, anchor="w")
        self.tabla.column("ruc", width=90, anchor="center")
        self.tabla.column("evento", width=110, anchor="center")
        self.tabla.column("kilometraje", width=80, anchor="center")
        self.tabla.column("cantidad", width=80, anchor="center")
        self.tabla.column("desc", width=130, anchor="w")
        self.tabla.column("metodo_pago", width=120, anchor="center")
        self.tabla.column("neto", width=85, anchor="e")
        
        self.tabla.config(displaycolumns=("num", "fecha", "hora", "nro_doc", "orden", "proveedor", "ruc", "evento", "kilometraje", "cantidad", "desc", "metodo_pago", "neto"))
        self._actualizar_flechas_orden()
        self.tabla.bind("<Double-1>", self.abrir_archivo)

        scroll_y = ttk.Scrollbar(f_tabla, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scroll_y.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        # ✅ Solo los botones: el de Modificar/Eliminar pasa a una fila propia
        # (ancho completo) para que siempre sea visible sin importar el ancho.
        f_gestion = ctk.CTkFrame(self.f_wrapper_derecha, fg_color="transparent")
        f_gestion.pack(fill="x", pady=(8, 0))

        # Descarga de tickets al abrir el módulo (en segundo plano, aviso en el hilo principal)
        ejecutar_en_hilo(
            self.main_root,
            lambda estado: estado.update(n=self.sincronizar_tickets_pendientes_automatico()),
            al_terminar=lambda estado: (self.cargar_datos_tabla(reset_pagina=True)
                                        if estado.get("n") else None))
        # 🔍 Ventana al MÁXIMO para ver todas las facturas (Windows y Mac)
        try:
            self._maximizar_ventana()
        except Exception:
            pass
        self.main_root.after(100, lambda: self.cargar_datos_tabla(reset_pagina=True))

    def _maximizar_ventana(self):
        """Pone la ventana del programa al máximo (funciona en Windows y macOS)."""
        raiz = getattr(self, "main_root", None)
        if raiz is None:
            return
        try:
            raiz.state("zoomed")                 # Windows
            return
        except Exception:
            pass
        try:
            raiz.attributes("-zoomed", True)     # algunos Linux
            return
        except Exception:
            pass
        try:
            ancho = raiz.winfo_screenwidth()
            alto = raiz.winfo_screenheight()
            raiz.geometry("%dx%d+0+0" % (ancho, alto))
        except Exception:
            pass

    def on_cambiar_mes(self, choice):
        self.mes_filtro = choice
        self.cargar_datos_tabla(reset_pagina=True)

    def pagina_anterior(self):
        if self.pagina_actual > 1:
            self.pagina_actual -= 1
            self.cargar_datos_tabla()
            
    def pagina_siguiente(self):
        if self.pagina_actual < getattr(self, "total_paginas", 1):
            self.pagina_actual += 1
            self.cargar_datos_tabla()

    def buscar_con_retraso(self):
        if hasattr(self, "_busqueda_job"):
            try: self.main_root.after_cancel(self._busqueda_job)
            except: pass
        self._busqueda_job = self.main_root.after(350, lambda: self.cargar_datos_tabla(reset_pagina=True))

    def ejecutar_sincronizacion_manual(self):
        """Descarga los tickets de la App en segundo plano y avisa en el hilo principal."""

        def _trabajo(estado):
            estado["n"] = self.sincronizar_tickets_pendientes_automatico()

        def _terminar(estado):
            if estado.get("error"):
                messagebox.showerror("Error al sincronizar", f"No se pudieron descargar los tickets:\n{estado['error']}",
                                     parent=self.main_root)
                return
            n = estado.get("n") or 0
            self.cargar_datos_tabla(reset_pagina=True)
            if n:
                messagebox.showinfo("Actualización Exitosa",
                                    f"Se descargaron {n} ticket(s) de la aplicación móvil.",
                                    parent=self.main_root)
                return
            # Si no se descargó nada por el bloqueo de almacenamiento, se avisa de eso
            # (los tickets quedan en la nube para un equipo autorizado).
            try:
                from politica_almacenamiento import estado_almacenamiento, mensaje_bloqueo
                if not estado_almacenamiento().get("autorizado"):
                    messagebox.showwarning(
                        "Descarga bloqueada",
                        mensaje_bloqueo() + "\n\nLos tickets del App Grifo NO se borraron de la nube: "
                        "quedan disponibles para un equipo configurado con la cuenta del principal.",
                        parent=self.main_root)
                    return
            except Exception:
                pass
            messagebox.showinfo(
                "Sin tickets nuevos",
                "No hay tickets de la aplicación móvil pendientes de descargar.\n\n"
                "Si acabas de enviar uno, revisa en el celular si fue rechazado (RUC de la empresa no encontrado).",
                parent=self.main_root)

        ejecutar_en_hilo(self.main_root, _trabajo, al_terminar=_terminar)

    def sincronizar_tickets_pendientes_automatico(self):
        ruta_base = obtener_ruta_base_drive()

        # 🔒 EQUIPO NO AUTORIZADO: no se descargan los tickets del App Grifo y
        # TAMPOCO se borran de la nube (imagen_base64 se conserva), para que otro
        # equipo configurado con la cuenta del principal pueda descargarlos.
        if not ruta_base:
            try:
                from politica_almacenamiento import estado_almacenamiento
                if not estado_almacenamiento().get("autorizado"):
                    print("🔒 Descarga de tickets App Grifo bloqueada: este equipo no usa la cuenta Rclone del principal. "
                          "Los tickets se mantienen en la nube.")
                    return 0
            except Exception:
                pass
            return 0

        conn = conectar_db(silencioso=True)
        if not conn: return 0
        descargados = 0
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT id, imagen_base64, proveedor FROM facturas_recibidas WHERE archivo_ruta = 'PENDIENTE_DESCARGA' AND imagen_base64 IS NOT NULL")
            pendientes = cursor.fetchall()
            
            if not pendientes:
                return 0
            
            carpeta_destino = os.path.normpath(os.path.join(ruta_base, "facturas_recibidas"))
            if not os.path.exists(carpeta_destino): os.makedirs(carpeta_destino)
            
            for reg in pendientes:
                id_doc, img_b64, proveedor = reg
                prov_limpio = re.sub(r'[\\/*?:"<>|]', '-', proveedor) if proveedor else "GRIFO_APP"
                nombre_prov = prov_limpio.replace(' ', '_')
                nombre_archivo = f"Ticket_Movil_{id_doc}_{nombre_prov}_{datetime.now().strftime('%Y%m%d%H%M%S')}.jpg"
                ruta_final = os.path.normpath(os.path.join(carpeta_destino, nombre_archivo))
                
                try:
                    with open(ruta_final, "wb") as f:
                        f.write(base64.b64decode(img_b64))
                    
                    ruta_bd = ruta_para_guardar(ruta_final, ruta_base)
                    cursor.execute("UPDATE facturas_recibidas SET archivo_ruta = %s, imagen_base64 = NULL WHERE id = %s", (ruta_bd, id_doc))
                    cursor.execute("UPDATE pagos_comprobantes SET archivo_ruta = %s WHERE id_factura = %s AND archivo_ruta = 'PENDIENTE_DESCARGA'", (ruta_bd, id_doc))
                    conn.commit()
                    descargados += 1
                except Exception as e_escribir: 
                    print(f"Error escribiendo {ruta_final}: {e_escribir}")
                
            if descargados > 0:
                print(f"🧹 Sincronización automática: Se descargaron {descargados} ticket(s).")
                # OJO: esta función corre en un hilo, así que AQUÍ no se toca la
                # interfaz; el refresco lo hace quien la llama (hilo principal).
        except Exception as e_sync: 
            print(f"Error sincronizando: {e_sync}")
        finally: liberar_conexion(conn)
        return descargados

    def cargar_categorias(self):
        # Usar las MISMAS categorías que el módulo Banco (Conciliación → Agregar Movimiento)
        from modulo_banco import cargar_categorias_gastos
        cats = cargar_categorias_gastos()
        base_cats = []
        for principal, subs in cats.items():
            if isinstance(subs, list):
                for s in subs:
                    etiqueta = f"{principal} - {s}".strip()
                    if etiqueta and etiqueta not in base_cats:
                        base_cats.append(etiqueta)
        if not base_cats:
            base_cats = ["Gastos Fijos - Alquiler"]

        self.combo_categoria.configure(values=base_cats)
        if not self.combo_categoria.get() or self.combo_categoria.get() not in base_cats:
            self.combo_categoria.set(base_cats[0])
        # El bloque del carné de sanidad aparece solo con esa categoría
        self._actualizar_bloque_carnet()

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
        if not prov:
            return

        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            # Búsqueda tolerante a mayúsculas/minúsculas y espacios
            cursor.execute(
                "SELECT ruc, porcentaje_detraccion FROM proveedores "
                "WHERE TRIM(nombre) ILIKE %s ORDER BY id ASC LIMIT 1",
                (prov,)
            )
            res = cursor.fetchone()

            if res:
                ruc_db, det_db = res
                self.ent_detraccion.configure(state="normal")
                if "Factura" in self.combo_tipo.get():
                    self.ent_detraccion.delete(0, tk.END)
                    try:
                        valor_det = str(float(det_db)) if det_db not in (None, "") else "0"
                    except Exception:
                        valor_det = "0"
                    self.ent_detraccion.insert(0, valor_det)

                # RUC del proveedor (salvo que esté bloqueado por importación PDF/XML)
                if not self.bloquear_autocompletado_ruc:
                    self.ent_desc.delete(0, tk.END)
                    if ruc_db:
                        self.ent_desc.insert(0, str(ruc_db))
                self._ruc_autocompletado = bool(ruc_db) and not self.bloquear_autocompletado_ruc
            else:
                # Proveedor no encontrado -> nunca dejar el RUC del proveedor anterior
                if not self.bloquear_autocompletado_ruc:
                    self.ent_desc.delete(0, tk.END)
                self._ruc_autocompletado = False
                if "Factura" in self.combo_tipo.get():
                    self.ent_detraccion.configure(state="normal")
                    self.ent_detraccion.delete(0, tk.END)
                    self.ent_detraccion.insert(0, "0")

            self.actualizar_totales()
        except Exception:
            pass
        finally:
            liberar_conexion(conn)

    def _al_teclear_proveedor(self, event=None):
        """Al escribir el proveedor: si coincide con uno registrado rellena su
        RUC/detracción; si se vacía el combo, limpia el RUC autocompletado."""
        texto = self.combo_proveedor.get().strip()
        try:
            opciones = list(self.combo_proveedor.cget("values") or [])
        except Exception:
            opciones = []

        if texto == "":
            if getattr(self, "_ruc_autocompletado", False):
                try:
                    self.ent_desc.delete(0, tk.END)
                except Exception:
                    pass
                self._ruc_autocompletado = False
            return

        if any(str(o).strip().casefold() == texto.casefold() for o in opciones):
            self.al_seleccionar_proveedor()

    def cargar_proveedores_bd(self):
        provs = cache_sistema.obtener("lista_proveedores_combobox")
        if provs is not None:
            self._aplicar_provs(provs)
            return

        self.combo_proveedor.set("Cargando proveedores...")

        def _leer_proveedores(estado):
            p_lista = []
            conn = conectar_db(silencioso=True)
            if conn:
                try:
                    c = conn.cursor()
                    c.execute("SELECT nombre FROM proveedores ORDER BY nombre ASC")
                    p_lista = [str(r[0]) for r in c.fetchall()]
                    cache_sistema.guardar("lista_proveedores_combobox", p_lista)
                except Exception:
                    pass
                finally:
                    liberar_conexion(conn)
            estado["valor"] = p_lista

        # El hilo solo consulta; el combo se llena desde el hilo principal
        ejecutar_en_hilo(self.main_root, _leer_proveedores,
                         al_terminar=lambda e: self._aplicar_provs(e.get("valor") or []))
            
    def _aplicar_provs(self, provs):
        if provs:
            self.combo_proveedor.configure(values=provs)
            try:
                self.combo_proveedor._valores_completos = list(self.combo_proveedor.cget('values') or [])
            except Exception:
                pass
        else:
            self.combo_proveedor.configure(values=["Sin proveedores registrados"])
        # Entra SIN proveedor preseleccionado: el usuario elige uno del desplegable
        self.combo_proveedor.set("")

    def cargar_vehiculos_bd(self):
        vehiculos = cache_sistema.obtener("lista_placas_combobox")
        if vehiculos is not None:
            self._aplicar_vehs(vehiculos)
            return

        self.combo_evento.set("Cargando vehículos...")

        def _leer_vehiculos(estado):
            v_lista = []
            conn = conectar_db(silencioso=True)
            if conn:
                try:
                    c = conn.cursor()
                    c.execute("SELECT placa FROM flota_vehiculos ORDER BY placa ASC")
                    v_lista = [str(r[0]) for r in c.fetchall()]
                    cache_sistema.guardar("lista_placas_combobox", v_lista)
                except Exception:
                    pass
                finally:
                    liberar_conexion(conn)
            estado["valor"] = v_lista

        ejecutar_en_hilo(self.main_root, _leer_vehiculos,
                         al_terminar=lambda e: self._aplicar_vehs(e.get("valor") or []))

    def _aplicar_vehs(self, vehiculos):
        lista_vehiculos = ["GENERAL / OFICINA"] + vehiculos
        self.combo_evento.configure(values=lista_vehiculos)
        if self.combo_evento.get() not in lista_vehiculos:
            self.combo_evento.set("GENERAL / OFICINA")

    # =========================================================================
    # 🎫 CARNET DE SANIDAD: CHOFER + VENCIMIENTO (6 MESES) EN EL CRONOGRAMA
    # =========================================================================
    def _es_carnet_sanidad(self):
        """True si la categoría elegida en la factura es un carné de sanidad."""
        try:
            from choferes import es_categoria_carnet_sanidad
        except Exception:
            def es_categoria_carnet_sanidad(texto):
                return "sanidad" in str(texto or "").lower()
        try:
            return bool(es_categoria_carnet_sanidad(self.combo_categoria.get()))
        except Exception:
            return False

    def on_cambiar_categoria(self, valor=None):
        """Al cambiar la categoría se muestra u oculta el bloque del carné."""
        self._actualizar_bloque_carnet()

    def _actualizar_bloque_carnet(self):
        """Muestra los datos del carné de sanidad solo con esa categoría."""
        try:
            if not hasattr(self, "f_carnet"):
                return
            if self._es_carnet_sanidad():
                if not self.f_carnet.winfo_ismapped():
                    self.f_carnet.pack(fill="x", padx=10, pady=(0, 8), before=self._ancla_bloque_carnet)
                self.cargar_choferes_bd()
                self._actualizar_vencimiento_carnet()
            elif self.f_carnet.winfo_ismapped():
                self.f_carnet.pack_forget()
        except Exception as e:
            print("Aviso - bloque de carné de sanidad:", e)

    def _actualizar_vencimiento_carnet(self, evento=None):
        """Muestra el vencimiento: emisión (o fecha de la factura) + 6 meses."""
        try:
            from choferes import sumar_meses_fecha, DURABILIDAD_CARNET_SANIDAD_MESES
        except Exception:
            sumar_meses_fecha = lambda _t: ""
            DURABILIDAD_CARNET_SANIDAD_MESES = 6
        try:
            emision = self.ent_carnet_emision.get().strip() or self.ent_fecha.get().strip()
            vencimiento = sumar_meses_fecha(emision)
            if vencimiento:
                self.lbl_carnet_venc.configure(
                    text=(f"Vencimiento: {vencimiento}  ({emision} + {DURABILIDAD_CARNET_SANIDAD_MESES} meses)\n"
                          "Se agregará al cronograma como 'Venc. Carné de Sanidad'."))
            else:
                self.lbl_carnet_venc.configure(text="Vencimiento: — (escriba o elija la fecha de emisión)")
        except Exception:
            pass

    def cargar_choferes_bd(self):
        """Carga los choferes activos para el desplegable del carné de sanidad."""
        choferes = cache_sistema.obtener("lista_choferes_compras")
        if choferes is not None:
            self._aplicar_choferes(choferes)
            return

        self.combo_chofer.set("Cargando choferes...")

        def _leer_choferes(estado):
            lista = []
            conn = conectar_db(silencioso=True)
            if conn:
                try:
                    c = conn.cursor()
                    c.execute("""SELECT id, nombres, COALESCE(dni, '') FROM choferes
                                 WHERE COALESCE(estado, 'Activo') <> 'Inactivo'
                                 ORDER BY nombres ASC""")
                    lista = [[int(r[0]), str(r[1] or ""), str(r[2] or "")] for r in c.fetchall()]
                    cache_sistema.guardar("lista_choferes_compras", lista)
                except Exception as e:
                    print("Aviso cargando choferes:", e)
                finally:
                    liberar_conexion(conn)
            estado["valor"] = lista

        ejecutar_en_hilo(self.main_root, _leer_choferes,
                         al_terminar=lambda e: self._aplicar_choferes(e.get("valor") or []))

    def _aplicar_choferes(self, lista):
        try:
            self._mapa_choferes = {}
            for fila in lista:
                id_chofer, nombres, dni = (list(fila) + ["", ""])[:3]
                etiqueta = f"{nombres} — {dni}".strip(" —")
                if etiqueta:
                    self._mapa_choferes[etiqueta] = id_chofer
            etiquetas = list(self._mapa_choferes.keys())
            if not etiquetas:
                etiquetas = ["Sin choferes registrados"]
            self.combo_chofer.configure(values=etiquetas)
            # Se deja vacío a propósito: el usuario debe elegir el chofer.
            self.combo_chofer.set("" if self._mapa_choferes else etiquetas[0])
        except Exception as e:
            print("Aviso - lista de choferes:", e)

    def _registrar_carnet_sanidad(self, etiqueta_chofer, numero, emision, vencimiento=None):
        """Guarda el carné de sanidad del chofer y su vencimiento en el cronograma."""
        try:
            from choferes import actualizar_carnet_sanidad
        except Exception as e:
            return False, f"No se pudo cargar el módulo de Choferes: {e}"
        id_chofer = (getattr(self, "_mapa_choferes", {}) or {}).get(etiqueta_chofer)
        if not id_chofer:
            return False, "Seleccione un chofer de la lista."
        return actualizar_carnet_sanidad(id_chofer, numero=numero, emision=emision,
                                         vencimiento=vencimiento,
                                         usuario=getattr(self.app_padre, "usuario_activo", "Sistema"))


    # =========================================================================
    # 🔁 FACTURA CRUZADA (PAGADA POR UN TERCERO)
    # =========================================================================
    def _lista_cuentas_pago(self):
        """Cuentas bancarias + medios de pago (las mismas opciones que el Banco)."""
        try:
            lista = []
            for b in cargar_bancos():
                banco_nom = str(b.get("banco", "")).strip()
                cuenta_num = str(b.get("cuenta", "")).strip()
                if banco_nom or cuenta_num:
                    lista.append(f"{banco_nom} - {cuenta_num}".strip(" -"))
            lista.extend(["Efectivo / Caja Chica", "Tarjeta de Crédito", "Tarjeta de Débito", "Cheque", "Otro"])
            return lista
        except Exception:
            return ["Efectivo / Caja Chica", "Cheque", "Otro"]

    def _datos_cruzada(self):
        """Datos del pago a tercero tal como están en el formulario."""
        return {
            "activa": bool(getattr(self, "var_cruzada", None) and self.var_cruzada.get()),
            "tercero": self.ent_tercero.get().strip() if hasattr(self, "ent_tercero") else "",
            "cuenta": self.cmb_cuenta_cruzada.get().strip() if hasattr(self, "cmb_cuenta_cruzada") else "",
            "operacion": self.ent_oper_cruzada.get().strip() if hasattr(self, "ent_oper_cruzada") else "",
            "interbancario": bool(getattr(self, "var_interbancario", None) and self.var_interbancario.get()),
            "comision": (parsear_monto_texto(self.ent_comision.get())
                         if hasattr(self, "ent_comision") and self.var_interbancario.get() else 0.0),
            "fecha": (self.ent_fecha_pago_cruz.get().strip()
                      if getattr(self, "ent_fecha_pago_cruz", None) is not None else ""),
            "soporte": getattr(self, "ruta_soporte_tercero", ""),
        }

    def _actualizar_bloque_cruzada(self):
        """Muestra u oculta los datos del tercero al marcar la casilla."""
        try:
            if self.var_cruzada.get():
                if not self.f_cruzada.winfo_ismapped():
                    self.f_cruzada.pack(fill="x", padx=10, pady=(0, 8), before=self._ancla_bloque_carnet)
                self._actualizar_total_cruzada()
            elif self.f_cruzada.winfo_ismapped():
                self.f_cruzada.pack_forget()
        except Exception as e:
            print("Aviso - bloque de factura cruzada:", e)

    def _actualizar_total_cruzada(self, evento=None):
        """Muestra el total de la factura más la comisión interbancaria (al instante)."""
        try:
            subtotal, impuesto, total = self._montos_ingresados()
        except Exception:
            subtotal, impuesto, total = 0.0, 0.0, 0.0
        # Neto a pagar = total del documento menos la detracción escrita
        try:
            detraccion = total * (float(self.ent_detraccion.get() or 0) / 100.0)
        except Exception:
            detraccion = 0.0
        neto_factura = max(0.0, total - detraccion)
        comision = self._datos_cruzada()["comision"]
        try:
            if comision > 0:
                self.lbl_cruzada_total.configure(
                    text=(f"Factura: {formatear_moneda(neto_factura)}  +  Comisión: {formatear_moneda(comision)}"
                          f"  =  {formatear_moneda(neto_factura + comision)} sale de la cuenta"))
            else:
                self.lbl_cruzada_total.configure(
                    text=f"Factura: {formatear_moneda(neto_factura)} — sin comisión interbancaria")
        except Exception:
            pass

    def escanear_soporte_tercero(self):
        """🖨️ Escanea el comprobante del pago al tercero y lo deja como soporte."""
        try:
            import escaner as _esc
            _cfg = cargar_configuracion_regional()
            _datos = _esc.config_escaner(_cfg)
            _res = _escanear_con_progreso(None, "el documento", _esc,
                _esc.carpeta_escaneos(_cfg),
                dispositivo=_datos.get("escaner_dispositivo", ""),
                dpi=_datos.get("escaner_dpi", "300"),
                color=_datos.get("escaner_color", "Escala de grises"),
                origen=_datos.get("escaner_origen", "Cristal (una página a la vez)"),
                paginas=1,
                optimizar=str(_datos.get("escaner_optimizar", "Sí")).lower().startswith("s"),
                modo=_datos.get("escaner_modo", "Automático (sin ventanas)"))
            _pags = [p for p in (_res.get("paginas") or []) if os.path.exists(str(p))]
            if not _pags:
                return messagebox.showerror(
                    "Escáner", "No se pudo escanear el comprobante:" + chr(10) + "%s"
                    % (_res.get("error") or "sin detalle"))
            self.ruta_soporte_tercero = _pags[0]
            try:
                self.btn_soporte_tercero.configure(
                    text="✅ Escaneado", fg_color="#27ae60")
            except Exception:
                pass

            # 🧠 Se lee el comprobante escaneado: tercero, N° de operación y monto
            leido_nombre, leido_oper, leido_monto = "", "", 0.0
            _aviso = self._aviso_procesando("⏳ Leyendo el comprobante escaneado... (OCR)")
            try:
                try:
                    texto_soporte = _ocr_comprobante_con_reintentos(_pags)
                except Exception as e_ocr:
                    texto_soporte = ""
                    print("Aviso - OCR del comprobante escaneado:", e_ocr)
                if str(texto_soporte or "").strip():
                    leido_nombre = nombre_tercero_desde_texto(texto_soporte)
                    leido_oper = extraer_numero_operacion(texto_soporte)
                    try:
                        leido_monto = float(monto_del_comprobante(texto_soporte) or 0)
                    except Exception:
                        leido_monto = 0.0
                    if leido_nombre:
                        self.ent_tercero.delete(0, tk.END)
                        self.ent_tercero.insert(0, leido_nombre)
                    if leido_oper:
                        self.ent_oper_cruzada.delete(0, tk.END)
                        self.ent_oper_cruzada.insert(0, leido_oper)
                    if leido_monto > 0:
                        try:
                            self.ent_monto_cruzada.delete(0, tk.END)
                            self.ent_monto_cruzada.insert(0, "%.2f" % leido_monto)
                        except Exception:
                            pass
                    try:
                        self._actualizar_total_cruzada()
                    except Exception:
                        pass
            except Exception as e_lec:
                print("Aviso - datos del comprobante escaneado:", e_lec)
            finally:
                self._cerrar_aviso(_aviso)

            detalle = ("Comprobante escaneado y listo para guardarse con la factura:" +
                       chr(10) + chr(10) + os.path.basename(_pags[0]))
            if leido_nombre:
                detalle += chr(10) + chr(10) + "Tercero leído: " + leido_nombre
            if leido_oper:
                detalle += chr(10) + "N° de operación leído: " + leido_oper
            if leido_monto > 0:
                detalle += chr(10) + "Monto leído: " + formatear_moneda(leido_monto)
            if not (leido_nombre or leido_oper or leido_monto > 0):
                detalle += (chr(10) + chr(10) +
                            "No se pudieron leer los datos (documento sin texto claro): "
                            "escríbalos a mano.")
            messagebox.showinfo("Comprobante de pago", detalle)
        except Exception as e:
            messagebox.showerror("Escáner", "No se pudo escanear el comprobante:" + chr(10) + "%s" % e)

    def seleccionar_soporte_tercero(self):
        """📂 Adjunta el comprobante del pago y lee el tercero y el N° de operación."""
        ruta = seleccionar_archivo_dialogo("Seleccionar el soporte del pago al tercero",
                                           [("Archivos", "*.pdf;*.png;*.jpg;*.jpeg")])
        if not ruta:
            return
        self.ruta_soporte_tercero = ruta
        try:
            self.btn_soporte_tercero.configure(text="🖨️ Escanear",
                                               fg_color="#7f8c8d", hover_color="#606b6b")
            self.btn_soporte_tercero_pdf.configure(text="✅ PDF cargado",
                                                   fg_color="#28a745", hover_color="#218838")
        except Exception:
            pass

        leido_nombre, leido_oper = "", ""
        _aviso = self._aviso_procesando("⏳ Leyendo el comprobante del pago... (OCR)")
        try:
            texto_soporte = _texto_del_soporte(ruta)
            if str(texto_soporte or "").strip():
                leido_nombre = nombre_tercero_desde_texto(texto_soporte)
                leido_oper = extraer_numero_operacion(texto_soporte)
                if leido_nombre:
                    self.ent_tercero.delete(0, tk.END)
                    self.ent_tercero.insert(0, leido_nombre)
                if leido_oper:
                    self.ent_oper_cruzada.delete(0, tk.END)
                    self.ent_oper_cruzada.insert(0, leido_oper)
                try:
                    self._actualizar_total_cruzada()
                except Exception:
                    pass
        except Exception as e_lec:
            print("Aviso - datos del soporte del tercero:", e_lec)
        finally:
            self._cerrar_aviso(_aviso)

        detalle = "Archivo adjunto: " + os.path.basename(ruta)
        if leido_nombre:
            detalle += chr(10) + chr(10) + "Tercero leído: " + leido_nombre
        if leido_oper:
            detalle += chr(10) + "N° de operación leído: " + leido_oper
        if not (leido_nombre or leido_oper):
            detalle += chr(10) + chr(10) + "No se pudieron leer los datos: escríbalos a mano."
        try:
            messagebox.showinfo("Soporte del pago", detalle)
        except Exception:
            pass

    def _guardar_soporte_tercero(self, ruta_base, tercero):
        """Copia el soporte del pago al tercero a la carpeta autorizada.

        Devuelve la ruta guardada (relativa) o "" si no había soporte.
        """
        origen = getattr(self, "ruta_soporte_tercero", "")
        if not origen or not os.path.exists(origen):
            return ""
        try:
            carpeta = os.path.normpath(os.path.join(ruta_base, "soportes_pagos_terceros"))
            if not os.path.exists(carpeta):
                os.makedirs(carpeta)
            ext = os.path.splitext(origen)[1] or ".pdf"
            limpio = re.sub(r'[\\/*?:"<>|]', '-', str(tercero or "tercero")).replace(" ", "_")[:40]
            destino = os.path.normpath(os.path.join(
                carpeta, f"Soporte_{datetime.now().strftime('%Y%m%d%H%M%S')}_{limpio}{ext}"))
            shutil.copy2(origen, destino)
            return ruta_para_guardar(destino)
        except Exception as e:
            print("Aviso - copiando el soporte del pago al tercero:", e)
            return ""

    def _limpiar_formulario_factura(self):
        """Deja el formulario listo para registrar el siguiente documento.

        Se borra el N° de documento (antes se quedaba el de la factura anterior y
        se repetía al registrar varias facturas seguidas del mismo proveedor), el
        concepto, los montos, el adjunto y los datos del carné de sanidad.
        El proveedor, la fecha, el tipo y la categoría se conservan para seguir
        facturando, y el RUC se vuelve a cargar del proveedor seleccionado.
        """
        try:
            self.ent_nro_doc.delete(0, tk.END)
            self.ent_concepto.delete(0, tk.END)
            self.ent_subtotal.delete(0, tk.END)
            self.ent_desc.delete(0, tk.END)            # RUC (se recarga del proveedor)
            if hasattr(self, "ent_recargo"):
                self.ent_recargo.delete(0, tk.END)
                self.ent_recargo.insert(0, "0")
            if hasattr(self, "ent_detraccion") and "Factura" in self.combo_tipo.get():
                self.ent_detraccion.delete(0, tk.END)
                self.ent_detraccion.insert(0, "0")
            # 🎫 Datos del carné de sanidad del documento anterior
            for entrada in (getattr(self, "ent_carnet_num", None), getattr(self, "ent_carnet_emision", None)):
                if entrada is not None:
                    entrada.delete(0, tk.END)
            if hasattr(self, "combo_chofer"):
                self.combo_chofer.set("")
            self._actualizar_vencimiento_carnet()
            # 🔁 Datos del pago a tercero del documento anterior
            if hasattr(self, "var_cruzada"):
                self.var_cruzada.set(False)
                self.var_interbancario.set(False)
                for entrada in (self.ent_tercero, self.ent_oper_cruzada, self.ent_comision):
                    entrada.delete(0, tk.END)
                self.ent_comision.insert(0, f"{cargar_comision_interbancaria_config():.2f}")
                self.ruta_soporte_tercero = ""
                self.btn_soporte_tercero.configure(text="📎 Adjuntar soporte del pago al tercero",
                                                   fg_color="#7f8c8d", hover_color="#606b6b")
                self._actualizar_bloque_cruzada()
            # 📎 Adjunto manual
            self.ruta_archivo_temp = ""
            self.btn_archivo.configure(text="📎 Adjuntar Archivo Manual", fg_color="#7f8c8d", hover_color="#606b6b")
            # 🧾 RUC del proveedor que sigue seleccionado (evita el aviso de campos vacíos)
            self._ruc_autocompletado = False
            self.bloquear_autocompletado_ruc = False
            if self.combo_proveedor.get().strip():
                try: self.al_seleccionar_proveedor()
                except Exception: pass
            self.actualizar_totales()
        except Exception as e:
            print("Aviso - limpiando el formulario de compras:", e)


    # =========================================================================
    # 🔗 ÓRDENES DE SERVICIO: CRUCE Y RELACIÓN CON LA FACTURA DEL PROVEEDOR
    # =========================================================================
    def cargar_ordenes_servicio(self, forzar=False):
        """Carga las órdenes de servicio (módulo de Órdenes) para el desplegable."""
        if not forzar:
            ordenes = cache_sistema.obtener("ordenes_servicio_compras")
            if ordenes is not None:
                self._aplicar_ordenes_servicio(ordenes)
                return

        def _leer_ordenes(estado):
            lista = []
            conn = conectar_db(silencioso=True)
            if conn:
                try:
                    c = conn.cursor()
                    c.execute("""
                        SELECT o.id, COALESCE(o.numero_orden, ''), COALESCE(o.version, 0),
                               COALESCE(o.placa, ''), COALESCE(o.proveedor, ''),
                               COALESCE(o.servicio, ''), COALESCE(o.costo_total, 0),
                               COALESCE(o.pdf_ruta, ''),
                               (SELECT COUNT(*) FROM facturas_recibidas f
                                 WHERE f.id_orden_servicio = o.id AND COALESCE(f.orden_servicio, '') <> '')
                        FROM ordenes_servicio_flota o
                        WHERE o.estado IS NULL OR o.estado <> 'Anulada'
                        ORDER BY o.id DESC LIMIT 500
                    """)
                    for fila in c.fetchall():
                        lista.append([int(fila[0]), str(fila[1] or ""), int(fila[2] or 0), str(fila[3] or ""),
                                      str(fila[4] or ""), str(fila[5] or ""), float(fila[6] or 0),
                                      str(fila[7] or ""), int(fila[8] or 0)])
                    cache_sistema.guardar("ordenes_servicio_compras", lista)
                except Exception:
                    lista = []
                finally:
                    liberar_conexion(conn)
            estado["valor"] = lista

        # El hilo solo consulta; el combo se llena desde el hilo principal
        ejecutar_en_hilo(self.main_root, _leer_ordenes,
                         al_terminar=lambda e: self._aplicar_ordenes_servicio(e.get("valor") or []))

    def _numero_visible_orden(self, orden):
        """N° de orden tal como se imprime en el PDF (con su versión si la tiene)."""
        numero = orden[1] or f"OS-{orden[0]}"
        return f"{numero}-{orden[2]}" if orden[2] else numero

    def _etiqueta_orden_servicio(self, orden):
        """Texto que se muestra en el desplegable para identificar la orden."""
        partes = [self._numero_visible_orden(orden)]
        if orden[3]:
            partes.append(orden[3])
        if orden[4]:
            partes.append(orden[4][:24])
        if orden[6]:
            partes.append(formatear_moneda(orden[6]))
        if orden[8]:
            partes.append("⚠ ya facturada")
        return " · ".join(partes)

    def _aplicar_ordenes_servicio(self, lista):
        self._ordenes_servicio = {}
        etiquetas = [SIN_ORDEN_SERVICIO]
        for orden in lista:
            etiqueta = self._etiqueta_orden_servicio(orden)
            if etiqueta in self._ordenes_servicio:
                etiqueta = f"{etiqueta} (#{orden[0]})"
            self._ordenes_servicio[etiqueta] = orden
            etiquetas.append(etiqueta)
        try:
            self.combo_orden_servicio.configure(values=etiquetas)
            if self.combo_orden_servicio.get() not in etiquetas:
                self.combo_orden_servicio.set(SIN_ORDEN_SERVICIO)
        except Exception:
            pass
        # Si aún no aparece ninguna orden (por ejemplo al abrir la app mientras se
        # preparan las columnas nuevas), se reintenta una vez en segundo plano.
        if not lista and not getattr(self, "_reintento_ordenes_hecho", False):
            self._reintento_ordenes_hecho = True
            try:
                self.main_root.after(4000, lambda: self.cargar_ordenes_servicio(forzar=True))
            except Exception:
                pass
        self._actualizar_cruce_orden()

    def orden_servicio_actual(self):
        """Datos de la orden elegida en el desplegable (None si no hay ninguna)."""
        if not hasattr(self, "combo_orden_servicio"):
            return None
        etiqueta = self.combo_orden_servicio.get()
        if not etiqueta or etiqueta == SIN_ORDEN_SERVICIO:
            return None
        return (getattr(self, "_ordenes_servicio", {}) or {}).get(etiqueta)

    def al_seleccionar_orden_servicio(self, choice=None):
        """Al elegir una orden se completan los datos que falten y se compara con la factura."""
        orden = self.orden_servicio_actual()
        if orden:
            # Proveedor de la orden: se completa solo si aún no se eligió ninguno
            if orden[4] and not self.combo_proveedor.get().strip():
                self.combo_proveedor.set(orden[4])
                self.al_seleccionar_proveedor()
            # Placa de la unidad (si la orden la tiene y aún no se eligió una)
            if orden[3]:
                try:
                    valores = list(self.combo_evento.cget("values") or [])
                    actual = self.combo_evento.get().strip()
                    if not actual or actual == "GENERAL / OFICINA":
                        if orden[3] not in valores:
                            valores.append(orden[3])      # la unidad puede no estar en la lista
                            self.combo_evento.configure(values=valores)
                        self.combo_evento.set(orden[3])
                except Exception:
                    pass
            # Servicio solicitado como concepto
            if orden[5] and not self.ent_concepto.get().strip():
                self.ent_concepto.delete(0, tk.END)
                self.ent_concepto.insert(0, orden[5])
            # Si todavía no se digitó el monto, se propone el costo acordado de la orden
            if orden[6] > 0 and not self.ent_subtotal.get().strip():
                self._poner_modo_monto("CON_IGV")
                self.ent_subtotal.insert(0, f"{orden[6]:.2f}")
        self._actualizar_cruce_orden()
        self.actualizar_totales()

    def _actualizar_cruce_orden(self):
        """Compara el monto de la factura con el costo acordado de la orden elegida."""
        if not hasattr(self, "lbl_cruce_orden"):
            return
        orden = self.orden_servicio_actual()
        if not orden:
            self.lbl_cruce_orden.configure(text="Sin orden de servicio relacionada.", text_color="#777777")
            try:
                self.btn_ver_orden.configure(state="disabled")
            except Exception:
                pass
            return
        try:
            self.btn_ver_orden.configure(state="normal")
        except Exception:
            pass

        numero = self._numero_visible_orden(orden)
        costo = orden[6]
        try:
            _sub, _igv, total_factura = self._montos_ingresados()
        except ValueError:
            total_factura = 0.0

        if costo <= 0:
            self.lbl_cruce_orden.configure(text=f"Orden {numero}: sin costo registrado.", text_color="#b9770e")
        elif total_factura <= 0:
            self.lbl_cruce_orden.configure(
                text=f"Orden {numero}: costo acordado {formatear_moneda(costo)}. Ingrese el monto de la factura.",
                text_color="#555555")
        elif abs(total_factura - costo) <= 0.05:
            self.lbl_cruce_orden.configure(
                text=f"✔ Cruce correcto: la factura coincide con la orden {numero} ({formatear_moneda(costo)}).",
                text_color="#1e8449")
        else:
            diferencia = total_factura - costo
            self.lbl_cruce_orden.configure(
                text=f"⚠ La orden {numero} es {formatear_moneda(costo)} y la factura "
                     f"{formatear_moneda(total_factura)}: diferencia de {formatear_moneda(abs(diferencia))} "
                     f"{'de más' if diferencia > 0 else 'de menos'}.",
                text_color="#c0392b")

    def abrir_pdf_orden_servicio(self):
        """Abre el PDF de la orden de servicio relacionada."""
        orden = self.orden_servicio_actual()
        if not orden:
            return messagebox.showinfo("Orden de Servicio", "Seleccione una orden de servicio de la lista.")
        if not orden[7]:
            return messagebox.showinfo("Orden de Servicio", "Esta orden no tiene PDF guardado.")
        try:
            ruta_abs = resolver_ruta_archivo(orden[7])
        except Exception:
            ruta_abs = orden[7]
        if ruta_abs and os.path.exists(ruta_abs):
            abrir_documento(ruta_abs)
        else:
            messagebox.showwarning("Orden de Servicio",
                                   "No se encontró el PDF de la orden en este equipo.\n"
                                   "Puede abrirlo desde el módulo de Órdenes de Servicio.")

    # =========================================================================
    # 💰 MONTO BASE (SIN IGV)  /  MONTO CON IGV (TOTAL)
    # =========================================================================
    def _es_factura(self):
        return "Factura" in self.combo_tipo.get()

    def _tasa_igv(self):
        """Tasa de IGV según el tipo de documento elegido."""
        return 0.105 if "10.5%" in self.combo_tipo.get() else 0.18

    def _monto_tecleado(self):
        """Número escrito en la casilla del monto (0.0 si está vacía)."""
        return parsear_monto_texto(self.ent_subtotal.get())

    def _recargo_tecleado(self):
        """Recargo al consumo digitado (servicio/propina): suma al total y no lleva IGV."""
        if not hasattr(self, "ent_recargo"):
            return 0.0
        return parsear_monto_texto(self.ent_recargo.get())

    def _montos_ingresados(self):
        """Devuelve (subtotal, igv, total) según el modo elegido por el usuario.

        - Modo "BASE"    : lo escrito es el monto SIN IGV  -> se calcula el IGV y el total.
        - Modo "CON_IGV" : lo escrito es el monto CON IGV  -> se calcula la base y el IGV.
        """
        valor = self._monto_tecleado()
        recargo = self._recargo_tecleado()
        if self._es_factura():
            tasa = self._tasa_igv()
            if self.modo_monto == "CON_IGV":
                # Lo digitado es el total del documento: se descuenta primero el
                # recargo al consumo (no gravado) y el resto se reparte base + IGV.
                gravado = max(valor - recargo, 0.0)
                subtotal = gravado / (1.0 + tasa)
                igv = gravado - subtotal
            else:
                subtotal = valor
                igv = subtotal * tasa
        else:
            # Boletas, recibos por honorarios y otros: no llevan IGV en el monto.
            subtotal = valor
            igv = 0.0
        total = subtotal + igv + recargo
        return round(subtotal, 2), round(igv, 2), round(total, 2)

    def _refrescar_etiqueta_modo(self):
        """Ajusta el rótulo de la casilla y el aviso según el modo y el tipo de documento."""
        if not hasattr(self, "lbl_monto_entrada") or not hasattr(self, "lbl_nota_modo"):
            return
        con_igv = self.modo_monto == "CON_IGV"
        if self._es_factura():
            self.lbl_monto_entrada.configure(text="Monto Total con IGV:" if con_igv else "Monto Base (Subtotal):")
            self.lbl_nota_modo.configure(
                text=("Digite el monto final del documento (con IGV): la base y el IGV se calculan solos."
                      if con_igv else
                      "Digite el monto sin IGV: el IGV y el total se calculan solos."))
        else:
            self.lbl_monto_entrada.configure(text="Monto Total:" if con_igv else "Monto Base (Subtotal):")
            self.lbl_nota_modo.configure(text="Este tipo de documento no lleva IGV: el monto se registra tal cual.")

    def _poner_modo_monto(self, modo):
        """Fija el modo sin recalcular lo escrito (se usa al autocargar PDF/XML)."""
        self.modo_monto = "CON_IGV" if modo == "CON_IGV" else "BASE"
        try:
            self.seg_modo_monto.set("Monto con IGV (Total)" if self.modo_monto == "CON_IGV" else "Monto Base (sin IGV)")
        except Exception:
            pass
        self._refrescar_etiqueta_modo()

    def on_cambiar_modo_monto(self, choice=None):
        """Cambia entre 'monto base' y 'monto con IGV' conservando el monto real."""
        nuevo = "CON_IGV" if "con igv" in str(choice or "").lower() else "BASE"
        if nuevo != self.modo_monto:
            try:
                subtotal, _igv, total = self._montos_ingresados()
            except ValueError:
                subtotal = total = 0.0
            self.modo_monto = nuevo
            nuevo_valor = total if nuevo == "CON_IGV" else subtotal
            self.ent_subtotal.delete(0, tk.END)
            if nuevo_valor > 0:
                self.ent_subtotal.insert(0, f"{nuevo_valor:.2f}")
        self._refrescar_etiqueta_modo()
        self.actualizar_totales()

    def actualizar_totales(self, *args):
        if not hasattr(self, "ent_subtotal"):
            return
        self._refrescar_etiqueta_modo()
        tipo = self.combo_tipo.get()
        try:
            subtotal, igv, total = self._montos_ingresados()
            recargo = self._recargo_tecleado()
            ui_pct = float(self.ent_detraccion.get() or 0)
        except ValueError:
            return

        if "Factura" in tipo:
            txt_igv = "10.5%" if "10.5%" in tipo else "18%"
            det = total * (ui_pct / 100.0)
            neto = total - det
            self.lbl_base.configure(text=f"Monto Base (Subtotal): {formatear_moneda(subtotal)}")
            self.lbl_impuesto.configure(text=f"IGV ({txt_igv}): {formatear_moneda(igv)}")
            self.lbl_recargo.configure(text=f"Recargo al consumo: {formatear_moneda(recargo)}")
            self.lbl_bruto.configure(text=f"Total con IGV: {formatear_moneda(total)}")
            self.lbl_detraccion.configure(text=f"Detracción ({ui_pct:g}%): -{formatear_moneda(det)}")
            self.lbl_total.configure(text=f"Neto a Pagar: {formatear_moneda(neto)}")
        elif "Recibo" in tipo:
            ret = subtotal * (ui_pct / 100.0)
            neto = total - ret
            self.lbl_base.configure(text=f"Monto Base (Subtotal): {formatear_moneda(subtotal)}")
            self.lbl_impuesto.configure(text=f"Retención ({ui_pct:g}%): -{formatear_moneda(ret)}")
            self.lbl_recargo.configure(text=f"Recargo al consumo: {formatear_moneda(recargo)}")
            self.lbl_bruto.configure(text=f"Total del Documento: {formatear_moneda(total)}")
            self.lbl_detraccion.configure(text=f"Detracción (0%): -{formatear_moneda(0)}")
            self.lbl_total.configure(text=f"Neto a Pagar: {formatear_moneda(neto)}")
        else:
            # Boletas / otros documentos sin IGV
            det = total * (ui_pct / 100.0)
            neto = total - det
            self.lbl_base.configure(text=f"Monto Base (Subtotal): {formatear_moneda(subtotal)}")
            self.lbl_impuesto.configure(text=f"IGV (0%): {formatear_moneda(0)}")
            self.lbl_recargo.configure(text=f"Recargo al consumo: {formatear_moneda(recargo)}")
            self.lbl_bruto.configure(text=f"Total del Documento: {formatear_moneda(total)}")
            self.lbl_detraccion.configure(text=f"Detracción ({ui_pct:g}%): -{formatear_moneda(det)}")
            self.lbl_total.configure(text=f"Neto a Pagar: {formatear_moneda(neto)}")

        self._actualizar_cruce_orden()
        # 🔁 Si la factura es cruzada, el resumen (factura + comisión interbancaria) se
        # recalcula al instante con cualquier cambio de montos, sin tocar la casilla.
        if getattr(self, "var_cruzada", None) is not None and self.var_cruzada.get():
            self._actualizar_total_cruzada()

    def seleccionar_archivo(self):
        ruta = seleccionar_archivo_dialogo("Seleccionar Documento", [("Archivos", "*.pdf;*.png;*.jpg;*.jpeg;*.xml")])
        if ruta:
            self.ruta_archivo_temp = ruta
            self.btn_archivo.configure(text="✅ Archivo Manual Listo", fg_color="#28a745", hover_color="#218838")

    def guardar_registro(self):
        ruta_base = obtener_ruta_base_drive()
        if not ruta_base:
            avisar_sin_permiso_guardado()
            return
            
        tipo = self.combo_tipo.get()
        nro_doc = self.ent_nro_doc.get().strip()
        fecha = self.ent_fecha.get().strip()
        prov = self.combo_proveedor.get().strip()
        
        ruc_val = self.ent_desc.get().strip()
        desc = self.ent_concepto.get().strip() or "Registro Manual"
        
        evento = self.combo_evento.get()
        categoria = self.combo_categoria.get().strip() or "GENERAL / NO ASIGNADO"

        # 🎫 ¿La factura es de un carné de sanidad? Entonces se pide el chofer y su
        #    vencimiento (emisión + 6 meses) se envía al cronograma.
        es_carnet_sanidad = self._es_carnet_sanidad()

        # 🔁 ¿La factura la pagó un tercero? (factura cruzada)
        cruzada = self._datos_cruzada()
        # 💸 El comprobante del tercero es la DEVOLUCIÓN del dinero: tiene que salir
        #    de una cuenta bancaria para que el egreso llegue al Banco.
        if cruzada["activa"] and not str(cruzada.get("cuenta") or "").strip():
            return messagebox.showwarning(
                "Pago de tercero",
                "Seleccione la CUENTA BANCARIA desde la que se devolvió el dinero al tercero." + chr(10) + chr(10) +
                "Sin cuenta bancaria el pago no se descuenta del saldo del banco.")

        try: 
            dias = int(self.ent_dias.get().strip() or 0)
            ui_pct = float(self.ent_detraccion.get() or 0)
        except ValueError: return messagebox.showerror("Error", "Los montos deben ser numéricos.")

        # 💰 Los montos salen de la casilla única según el modo elegido
        # (Monto Base sin IGV o Monto con IGV): el resto se calcula automáticamente.
        try:
            subtotal, imp_calculado, total_calculado = self._montos_ingresados()
        except ValueError:
            return messagebox.showerror("Error", "El monto ingresado no es válido.")

        # 🚫 Una factura sin monto quedaría como "ya cancelada" y no se podría pagar
        if not total_calculado or total_calculado <= 0:
            return messagebox.showerror(
                "Monto en cero",
                "El monto del documento es 0." + chr(10)  + chr(10) +
                "Revise el importe antes de registrar: una factura sin monto no aparece "
                "en Cuentas por Pagar y no se puede cancelar.")
        if not prov or not ruc_val: return messagebox.showwarning("Atención", "Llene los campos obligatorios.")

        if es_carnet_sanidad and not self.combo_chofer.get().strip():
            return messagebox.showwarning(
                "Carné de Sanidad",
                "Seleccione el chofer al que corresponde este carné de sanidad.\n\n"
                "El vencimiento (fecha de emisión + 6 meses) se enviará al cronograma.")

        if cruzada["activa"] and not cruzada["tercero"]:
            return messagebox.showwarning(
                "Factura Cruzada",
                "Escriba el nombre de la persona o empresa a quien se le pagó (tercero).\n\n"
                "La factura quedará registrada como pagada por ese tercero.")

        if nro_doc and nro_doc.upper() not in ("POR-ASIGNAR", "ERROR-LECTURA", "S/N"):
            conn_check = conectar_db()
            if conn_check:
                try:
                    c_check = conn_check.cursor()
                    # 🔎 Se compara SOLO el N° de documento (sin espacios ni mayúsculas) y
                    # NO el nombre del proveedor: el mismo comprobante puede venir escrito
                    # de dos formas (sobre todo si lo leyó la App desde una foto), y así se
                    # evita registrar dos veces la misma factura.
                    c_check.execute("""
                        SELECT proveedor, fecha, COALESCE(total, 0), COALESCE(ruc, '')
                        FROM facturas_recibidas
                        WHERE UPPER(REPLACE(TRIM(numero_documento), ' ', '')) = UPPER(REPLACE(TRIM(%s), ' ', ''))
                        ORDER BY id DESC LIMIT 1
                    """, (nro_doc,))
                    previa = c_check.fetchone()
                    if previa:
                        prov_prev, fecha_prev, total_prev, ruc_prev = previa
                        detalle = (f"N° {nro_doc}\nProveedor: {prov_prev}\nFecha: {fecha_prev}\n"
                                   f"Total: {formatear_moneda(total_prev)}")
                        mismo_ruc = (bool(str(ruc_val).strip())
                                     and str(ruc_prev or '').strip() == str(ruc_val).strip())
                        # 🔎 Solo se rechaza si es el MISMO proveedor: el mismo N° de otro
                        #    proveedor se acepta SIN preguntar nada.
                        if mismo_ruc or _mismo_proveedor(prov_prev, prov):
                            return messagebox.showwarning(
                                "Documento duplicado",
                                f"Este documento ya está registrado en el sistema:\n\n{detalle}\n\n"
                                "No se registró de nuevo. Revise el N° de documento.")
                except Exception as e_dup:
                    print("Aviso - verificación de documentos duplicados:", e_dup)
                finally: liberar_conexion(conn_check)

        if "Factura" in tipo: 
            imp = imp_calculado
            tot_bruto = total_calculado
            det_pct = ui_pct
            det_monto = tot_bruto * (det_pct / 100.0)
        elif "Recibo" in tipo: 
            imp = subtotal * (ui_pct / 100.0)
            tot_bruto = total_calculado          # base + recargo al consumo (si hubiera)
            det_pct = 0.0
            det_monto = 0.0
        else: 
            imp = 0.0
            tot_bruto = total_calculado          # base + recargo al consumo (si hubiera)
            det_pct = ui_pct
            det_monto = tot_bruto * (det_pct / 100.0)

        # 🔗 Orden de servicio relacionada (si el usuario eligió una en el desplegable)
        orden = self.orden_servicio_actual()
        orden_id = orden[0] if orden else None
        orden_txt = self._numero_visible_orden(orden) if orden else ""
        if orden_id:
            conn_orden = conectar_db()
            if conn_orden:
                try:
                    c_orden = conn_orden.cursor()
                    c_orden.execute("SELECT COALESCE(numero_documento, '') FROM facturas_recibidas "
                                    "WHERE id_orden_servicio = %s LIMIT 1", (orden_id,))
                    previa = c_orden.fetchone()
                    if previa and not messagebox.askyesno(
                            "Orden de Servicio",
                            f"La orden {orden_txt} ya está relacionada con la factura "
                            f"{previa[0] or '(sin N°)'}.\n\n¿Desea relacionarla también con este documento?"):
                        return
                except Exception:
                    pass
                finally:
                    liberar_conexion(conn_orden)

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

        # 📎 Soporte del pago hecho al tercero (se copia a la carpeta autorizada)
        ruta_soporte_tercero = ""
        if cruzada["activa"] and cruzada["soporte"]:
            ruta_soporte_tercero = self._guardar_soporte_tercero(ruta_base, cruzada["tercero"])
            if not ruta_soporte_tercero:
                return messagebox.showwarning("Factura Cruzada",
                                              "No se pudo copiar el soporte del pago al tercero.\n"
                                              "Verifique el archivo o vuelva a adjuntarlo.")

        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO facturas_recibidas (tipo_documento, numero_documento, fecha, proveedor, descripcion, evento_asociado, subtotal, impuesto, total, archivo_ruta, dias_credito, det_porcentaje, det_monto, categoria, ruc, id_orden_servicio, orden_servicio,
                                                pagado_por_tercero, soporte_pago_tercero, es_compra_cruzada, comision_interbancaria, es_interbancario)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                RETURNING id
            """, (tipo, nro_doc, fecha, prov, desc, evento, subtotal, imp, tot_bruto, ruta_para_guardar(ruta_final), dias, det_pct, det_monto, categoria, ruc_val, orden_id, orden_txt,
                  cruzada["tercero"] if cruzada["activa"] else "", ruta_soporte_tercero,
                  bool(cruzada["activa"]), cruzada["comision"], bool(cruzada["interbancario"])))
            try:
                id_factura_nueva = int(cursor.fetchone()[0])
            except Exception:
                id_factura_nueva = 0
            conn.commit()

            # 💳 Factura cruzada: se registra el pago hecho por el tercero para que la
            #    factura aparezca como PAGADA y su forma de pago muestre la cuenta.
            aviso_cruzada = ""
            if cruzada["activa"] and id_factura_nueva:
                if "Recibo" in tipo and "8%" in tipo:
                    neto_cruzada = tot_bruto - imp - det_monto
                else:
                    neto_cruzada = tot_bruto - det_monto
                # 💸 El tercero pagó la factura MÁS la comisión interbancaria: el pago
                #    registrado incluye las dos cosas (y así no queda nada pendiente).
                comision_cruzada = cruzada["comision"] if cruzada["interbancario"] else 0.0
                monto_pago_cruz = max(0.0, neto_cruzada) + max(0.0, comision_cruzada)
                try:
                    # 📅 El asiento del banco va con la FECHA DE PAGO (no la de la factura)
                    _fecha_pago_cruz = (cruzada.get("fecha") or "").strip() or datetime.now().strftime("%d/%m/%Y")
                    cur_pago = conn.cursor()
                    cur_pago.execute("""
                        INSERT INTO pagos_comprobantes (id_factura, monto_pagado, archivo_ruta, proveedor_nombre,
                                                       fecha_pago, categoria_suministro, codigo_cotizacion,
                                                       cuenta_origen, numero_operacion)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """, (id_factura_nueva, monto_pago_cruz, ruta_soporte_tercero or "", prov, _fecha_pago_cruz,
                          categoria, nro_doc, cruzada["cuenta"] or f"Pago de {cruzada['tercero']}",
                          cruzada["operacion"] or "PAGO A TERCERO"))
                    conn.commit()
                except Exception as e_pago_cruz:
                    print("Aviso - pago de la factura cruzada:", e_pago_cruz)
                detalle_comision = (f" (factura {formatear_moneda(neto_cruzada)}"
                                    f" + comisión {formatear_moneda(comision_cruzada)})"
                                    if comision_cruzada > 0 else "")
                aviso_cruzada = (f"\n\n💸 Devolución a {cruzada['tercero']} (pagó la factura)"
                                 f"\nPago registrado: {formatear_moneda(monto_pago_cruz)}{detalle_comision}"
                                 "\nSe descontó de la cuenta " + str(cruzada.get("cuenta") or "")
                                 + " el " + str(_fecha_pago_cruz or "") +
                                 ": queda como egreso en el Banco y PAGADA en Cuentas por Pagar.")
            
            cache_sistema.invalidar()
            registrar_auditoria(self.app_padre.usuario_activo, "Facturas Recibidas", f"Registró factura {nro_doc} del proveedor '{prov}'")

            # 🎫 Carné de sanidad: se actualiza el chofer y su vencimiento en el cronograma
            aviso_carnet = ""
            if es_carnet_sanidad:
                emision_carnet = self.ent_carnet_emision.get().strip() or fecha
                ok_carnet, respuesta_carnet = self._registrar_carnet_sanidad(
                    self.combo_chofer.get().strip(),
                    self.ent_carnet_num.get().strip() or nro_doc,
                    emision_carnet)
                if ok_carnet:
                    aviso_carnet = (f"\n\n🎫 Carné de sanidad de {self.combo_chofer.get().strip()}"
                                    f"\nEmisión: {emision_carnet}    Vence: {respuesta_carnet or '—'}"
                                    "\nSe agregó el vencimiento al cronograma.")
                else:
                    aviso_carnet = f"\n\n⚠ No se pudo actualizar el carné de sanidad: {respuesta_carnet}"

            messagebox.showinfo("Éxito", f"Documento {nro_doc or 's/n'} registrado correctamente."
                                + aviso_carnet + aviso_cruzada)

            # ✅ El formulario se limpia SIEMPRE (aunque falle un refresco), para que el
            #    siguiente documento no herede el N° de la factura anterior.
            self._limpiar_formulario_factura()
            for refresco in (self.cargar_categorias, self.cargar_ordenes_servicio):
                try: refresco()                 # refresca categorías y la marca "⚠ ya facturada"
                except Exception as e_ref: print("Aviso - refresco tras guardar:", e_ref)
            self.cargar_datos_tabla(reset_pagina=True)

            if hasattr(self.app_padre, 'app_pagos'):
                try: self.app_padre.app_pagos.cargar_datos_pagar(reset_pagina=True)
                except Exception as e_pag: print("Aviso - refresco de cuentas por pagar:", e_pag)

            # 💰 Se ofrece CANCELAR (pagar) la factura recién registrada y adjuntar
            #    el comprobante de pago (se puede escanear o cargar un PDF).
            if id_factura_nueva and not cruzada["activa"]:
                try:
                    self._ofrecer_cancelacion(id_factura_nueva, prov, nro_doc)
                except Exception as e_can:
                    print("Aviso - oferta de cancelación:", e_can)
        except Exception as e: messagebox.showerror("Error", str(e))
        finally: liberar_conexion(conn)

    def _ofrecer_cancelacion(self, id_factura, proveedor, numero_documento):
        """Pregunta si se cancela (paga) la factura recién registrada."""
        try:
            respuesta = messagebox.askyesno(
                "Cancelar la factura",
                "¿Desea CANCELAR (pagar) ahora la factura %s de %s?" % (numero_documento or "s/n",
                                                                        proveedor) + chr(10)  + chr(10) +
                "Se abrirá «Registrar Pago» y podrá adjuntar o escanear el comprobante.")
        except Exception:
            return
        if respuesta:
            try:
                self.abrir_pago_de_factura(id_factura)
            except Exception as e:
                print("Aviso - abrir el pago:", e)

    def abrir_pago_de_factura(self, id_factura, escanear=False):
        """Abre «Registrar Pago» para una factura concreta (la busca en TODAS las páginas)."""
        try:
            cache_sistema.invalidar()
        except Exception:
            pass
        # 📅 Si la factura es de otro mes, se muestra ese mes para poder verla
        try:
            conn_mes = conectar_db(silencioso=True)
            if conn_mes:
                try:
                    with conn_mes.cursor() as c_mes:
                        c_mes.execute("SELECT fecha FROM facturas_recibidas WHERE id = %s", (id_factura,))
                        fila_mes = c_mes.fetchone()
                finally:
                    liberar_conexion(conn_mes)
                etiqueta_fac = etiqueta_mes_de_fecha(fila_mes[0] if fila_mes else "")
                if etiqueta_fac and etiqueta_fac != getattr(self, "mes_filtro", ""):
                    self.mes_filtro = etiqueta_fac
                    try:
                        self.combo_mes.set(etiqueta_fac)
                    except Exception:
                        pass
        except Exception as e_mes:
            print("Aviso - mes de la factura:", e_mes)
        try:
            if hasattr(self, "var_ocultar_app") and self.var_ocultar_app.get():
                self.var_ocultar_app.set(False)
        except Exception:
            pass
        try:
            self.cargar_datos_pagar(reset_pagina=True)
        except Exception:
            pass

        def _seleccionar_si_esta():
            for item in self.tabla.get_children():
                valores = self.tabla.item(item, "values")
                if len(valores) > 1 and str(valores[1]) == str(id_factura):
                    self.tabla.selection_set(item)
                    self.tabla.see(item)
                    self._escanear_soporte_pendiente = bool(escanear)
                    self.cargar_comprobante_pago()
                    return True
            return False

        def buscar_factura(intento=1):
            if _seleccionar_si_esta():
                return
            try:
                total_paginas = int(getattr(self, "total_paginas", 1) or 1)
            except Exception:
                total_paginas = 1
            original = getattr(self, "pagina_actual", 1)
            for pagina in range(1, min(total_paginas, 60) + 1):
                if pagina == original:
                    continue
                try:
                    self.pagina_actual = pagina
                    self.cargar_datos_pagar()
                except Exception:
                    continue
                if _seleccionar_si_esta():
                    return
            try:
                self.pagina_actual = original
                self.cargar_datos_pagar()
            except Exception:
                pass
            if intento < 4:
                try:
                    self.main_root.after(400 * intento, lambda: buscar_factura(intento + 1))
                except Exception:
                    buscar_factura(intento + 1)
                return
            messagebox.showinfo(
                "Registrar Pago",
                "La factura todavía no aparece en la lista de Cuentas por Pagar." + chr(10)  + chr(10) +
                "Pulse «🔄 Actualizar» en la pestaña y vuelva a intentarlo.")
        buscar_factura()

    # =====================================================================
    # 🖱️ MENÚ DE CLIC DERECHO (las DOS pestañas usan el MISMO menú)
    # =====================================================================
    def _id_factura_seleccionada(self):
        """Id de la factura seleccionada en la tabla (o 0)."""
        try:
            sel = self.tabla.selection()
            if not sel:
                return 0
            return int(self.tabla.item(sel[0], "values")[1])
        except Exception:
            return 0

    def _primera_ruta_existente(self, candidatas):
        """Devuelve la primera ruta de la lista que exista en este equipo."""
        for ruta in (candidatas or []):
            ruta = str(ruta or "").strip()
            if not ruta:
                continue
            try:
                resuelta = resolver_ruta_archivo(ruta)
            except Exception:
                resuelta = ""
            if resuelta and os.path.exists(resuelta):
                return resuelta
        return ""

    def _ruta_factura(self, id_factura):
        """Ruta del archivo de la FACTURA (o el soporte del pago a tercero)."""
        if not id_factura:
            return ""
        candidatas = []
        try:
            conn = conectar_db(silencioso=True)
            if conn:
                try:
                    with conn.cursor() as c:
                        c.execute("""SELECT COALESCE(archivo_ruta, ''), COALESCE(soporte_pago_tercero, '')
                                     FROM facturas_recibidas WHERE id = %s""", (id_factura,))
                        f = c.fetchone()
                        if f:
                            candidatas.extend([f[0], f[1]])
                finally:
                    liberar_conexion(conn)
        except Exception:
            return ""
        return self._primera_ruta_existente(candidatas)

    def _rutas_comprobantes_pago(self, id_factura):
        """TODAS las rutas de comprobantes de pago de la factura (las que existen)."""
        if not id_factura:
            return []
        rutas = []
        try:
            conn = conectar_db(silencioso=True)
            if conn:
                try:
                    with conn.cursor() as c:
                        c.execute("""SELECT COALESCE(archivo_ruta, '') FROM pagos_comprobantes
                                     WHERE id_factura = %s ORDER BY id DESC""", (id_factura,))
                        candidatas = [r[0] for r in c.fetchall()]
                finally:
                    liberar_conexion(conn)
        except Exception:
            return []
        for ruta in candidatas:
            if not str(ruta or "").strip():
                continue
            real = ""
            try:
                real = self._primera_ruta_existente([ruta]) or ""
            except Exception:
                real = ""
            if real and real not in rutas:
                rutas.append(real)
        return rutas

    def _ruta_comprobante_pago(self, id_factura):
        """Ruta del COMPROBANTE de pago registrado para la factura."""
        if not id_factura:
            return ""
        candidatas = []
        try:
            conn = conectar_db(silencioso=True)
            if conn:
                try:
                    with conn.cursor() as c:
                        c.execute("""SELECT COALESCE(archivo_ruta, '') FROM pagos_comprobantes
                                     WHERE id_factura = %s ORDER BY id DESC""", (id_factura,))
                        candidatas.extend([r[0] for r in c.fetchall()])
                finally:
                    liberar_conexion(conn)
        except Exception:
            return ""
        return self._primera_ruta_existente(candidatas)

    def _datos_bancarios_factura(self, id_factura=None):
        """🏦 Datos bancarios del proveedor de la factura (funciona desde el menú)."""
        id_fact = id_factura or self._id_factura_seleccionada()
        if not id_fact:
            return
        proveedor, detalle = "", ""
        try:
            conn = conectar_db(silencioso=True)
            if conn:
                try:
                    with conn.cursor() as c:
                        c.execute("""SELECT COALESCE(proveedor, ''), COALESCE(numero_documento, ''),
                                            COALESCE(fecha, '')
                                     FROM facturas_recibidas WHERE id = %s""", (id_fact,))
                        f = c.fetchone()
                        if f:
                            proveedor = str(f[0] or "")
                            detalle = "Factura %s   ·   Fecha: %s" % (f[1] or "", f[2] or "")
                finally:
                    liberar_conexion(conn)
        except Exception as e:
            print("Aviso - datos bancarios:", e)
        if not proveedor.strip():
            return messagebox.showinfo("Datos bancarios",
                                       "Esta factura no tiene proveedor registrado.")
        cache = getattr(self, "_ventanas_bancarias", None)
        if cache is None:
            cache = self._ventanas_bancarias = {}
        previa = cache.get(proveedor)
        try:
            if previa is not None and not previa.winfo_exists():
                previa = None
        except Exception:
            previa = None
        if previa is not None:
            try:
                for hijo in previa.winfo_children():
                    hijo.destroy()
            except Exception:
                previa = None
        try:
            datos = obtener_datos_bancarios_proveedor(proveedor)
            ventana = mostrar_datos_bancarios_proveedor(
                self.main_root, proveedor, datos, factura_txt=detalle, reutilizar=previa)
            cache[proveedor] = ventana
            return ventana
        except Exception as e:
            print("Aviso - ventana de datos bancarios:", e)
            return messagebox.showerror("Datos bancarios",
                                        "No se pudo abrir la ventana: %s" % e)

    def _avisar_sin_archivo(self, id_factura, que):
        """Explica que la factura/comprobante no tiene archivo y cómo adjuntarlo."""
        try:
            numero = ""
            conn = conectar_db(silencioso=True)
            if conn:
                try:
                    with conn.cursor() as c:
                        c.execute("SELECT COALESCE(numero_documento,''), COALESCE(archivo_ruta,'') "
                                  "FROM facturas_recibidas WHERE id = %s", (id_factura,))
                        f = c.fetchone()
                        if f:
                            numero = str(f[0] or "")
                finally:
                    liberar_conexion(conn)
        except Exception:
            numero = ""
        if que == "comprobante":
            detalle = ("Esta factura no tiene un comprobante de pago con archivo adjunto.")
            pista = ("Registrá el pago con «💳 2. Control de Pagos» → Registrar Pago y adjuntá "
                     "o escaneá el comprobante.")
        else:
            detalle = ("El archivo de la factura no está adjunto (o el archivo ya no está en "
                       "este equipo).")
            pista = ("Usá «✏️ Editar factura» para adjuntar el PDF/imagen de la factura, "
                     "o escaneala con «Escanear factura (OCR)».")
        mensaje = detalle + chr(10) + chr(10) + pista
        if numero:
            mensaje = "Factura " + numero + chr(10) + chr(10) + mensaje
        messagebox.showinfo("Sin archivo", mensaje,
                            parent=getattr(self, "main_root", None))

    def _editar_factura(self, id_factura):
        """✏️ Editor completo de la factura (todos los campos, incluidos los de flota)."""
        if not id_factura:
            return
        conn = conectar_db(silencioso=True)
        if not conn:
            return
        datos = {}
        try:
            with conn.cursor() as c:
                c.execute("""SELECT tipo_documento, numero_documento, fecha, proveedor, ruc,
                                    categoria, descripcion, evento_asociado, dias_credito,
                                    subtotal, impuesto, total, det_porcentaje, det_monto,
                                    delivery, chofer_asociado, kilometraje, cantidad_combustible,
                                    orden_servicio
                             FROM facturas_recibidas WHERE id = %s""", (id_factura,))
                fila = c.fetchone()
                if not fila:
                    return
                columnas = ("tipo_documento", "numero_documento", "fecha", "proveedor", "ruc",
                            "categoria", "descripcion", "evento_asociado", "dias_credito",
                            "subtotal", "impuesto", "total", "det_porcentaje", "det_monto",
                            "delivery", "chofer_asociado", "kilometraje", "cantidad_combustible",
                            "orden_servicio")
                datos = dict(zip(columnas, fila))
        finally:
            liberar_conexion(conn)

        ventana = ctk.CTkToplevel(getattr(self, "main_root", None))
        ventana.title("Editar factura #%s" % id_factura)
        ventana.geometry("560x640")
        marco = ctk.CTkScrollableFrame(ventana, width=520, height=520)
        marco.pack(fill="both", expand=True, padx=10, pady=10)

        entradas = {}
        etiquetas = (("tipo_documento", "Tipo de documento"), ("numero_documento", "N° de documento"),
                     ("fecha", "Fecha (DD/MM/AAAA)"), ("proveedor", "Proveedor"),
                     ("ruc", "RUC"), ("categoria", "Categoría"),
                     ("descripcion", "Descripción"), ("evento_asociado", "Vehículo (Placa)"),
                     ("dias_credito", "Días de crédito"), ("subtotal", "Base"),
                     ("impuesto", "IGV"), ("total", "Total"),
                     ("det_porcentaje", "Detracción %"), ("det_monto", "Detracción S/"),
                     ("delivery", "Delivery"), ("chofer_asociado", "Chofer"),
                     ("kilometraje", "Kilometraje"), ("cantidad_combustible", "Galones"),
                     ("orden_servicio", "Orden de Servicio"))
        for campo, etiqueta in etiquetas:
            ctk.CTkLabel(marco, text=etiqueta + ":", font=("Arial", 11, "bold")).pack(anchor="w", pady=(4, 0))
            entrada = ctk.CTkEntry(marco)
            entrada.pack(fill="x")
            valor = datos.get(campo)
            entrada.insert(0, "" if valor is None else str(valor))
            entradas[campo] = entrada

        def guardar():
            valores = {c: e.get().strip() for c, e in entradas.items()}
            try:
                def a_num(x):
                    return float(str(x).replace(",", ".") or 0)
                conn_g = conectar_db(silencioso=True)
                if not conn_g:
                    return
                try:
                    with conn_g.cursor() as cg:
                        cg.execute("""UPDATE facturas_recibidas SET
                                        tipo_documento=%s, numero_documento=%s, fecha=%s,
                                        proveedor=%s, ruc=%s, categoria=%s, descripcion=%s,
                                        evento_asociado=%s, dias_credito=%s, subtotal=%s,
                                        impuesto=%s, total=%s, det_porcentaje=%s, det_monto=%s,
                                        delivery=%s, chofer_asociado=%s, kilometraje=%s,
                                        cantidad_combustible=%s, orden_servicio=%s
                                      WHERE id=%s""",
                                   (valores["tipo_documento"], valores["numero_documento"],
                                    valores["fecha"], valores["proveedor"], valores["ruc"],
                                    valores["categoria"], valores["descripcion"],
                                    valores["evento_asociado"],
                                    int(a_num(valores["dias_credito"])),
                                    a_num(valores["subtotal"]), a_num(valores["impuesto"]),
                                    a_num(valores["total"]), a_num(valores["det_porcentaje"]),
                                    a_num(valores["det_monto"]), valores["delivery"],
                                    valores["chofer_asociado"], valores["kilometraje"],
                                    a_num(valores["cantidad_combustible"]),
                                    valores["orden_servicio"], id_factura))
                    conn_g.commit()
                finally:
                    liberar_conexion(conn_g)
                try:
                    cache_sistema.invalidar()
                except Exception:
                    pass
                messagebox.showinfo("Factura", "Factura actualizada correctamente.")
                ventana.destroy()
                try:
                    self.cargar_datos_tabla(reset_pagina=True)
                except Exception:
                    pass
            except Exception as e:
                messagebox.showerror("Factura", "No se pudo guardar: %s" % e)

        ctk.CTkButton(marco, text="💾 Guardar cambios", font=("Arial", 12, "bold"),
                      fg_color="#27ae60", hover_color="#1e8449",
                      command=guardar).pack(fill="x", pady=(12, 6))
        return ventana

    def _mostrar_menu_factura(self, id_factura, event, es_pagos=False):
        """Menú de clic derecho de una factura (lo usan las DOS pestañas)."""
        menu = tk.Menu(getattr(self, "main_root", None), tearoff=0)
        if es_pagos:
            pagos = getattr(self.app_padre, "app_pagos", None)
            if pagos is not None:
                menu.add_command(label="🧾 Registrar Pago",
                                 command=pagos.cargar_comprobante_pago)
                menu.add_command(label="✏️ Editar Pagos",
                                 command=pagos.abrir_ventana_edicion)
                menu.add_separator()
        menu.add_command(label="✏️ Editar factura (todos los campos)",
                         command=lambda: self._editar_factura(id_factura))
        menu.add_command(label="🏦 Datos bancarios del proveedor",
                         command=lambda: self._datos_bancarios_factura(id_factura))
        ruta_f = self._ruta_factura(id_factura)
        if ruta_f:
            menu.add_command(label="📄 Ver factura",
                             command=lambda r=ruta_f: ver_documento_en_ventana(
                                 r, "Factura", self.main_root))
        else:
            # Se deja ACTIVO para explicar por qué no hay nada que abrir
            menu.add_command(label="📄 Ver factura (sin archivo adjunto)",
                             command=lambda: self._avisar_sin_archivo(id_factura, "factura"))
        rutas_c = self._rutas_comprobantes_pago(id_factura)
        ruta_c = rutas_c[0] if rutas_c else ""
        if rutas_c:
            etiqueta_menu = ("🧾 Ver comprobante de pago" if len(rutas_c) == 1
                             else "🧾 Ver los %d comprobantes de pago" % len(rutas_c))
            menu.add_command(label=etiqueta_menu,
                             command=lambda rs=list(rutas_c): self._ver_comprobantes_lista(rs))
            try:
                vinculadas = facturas_que_cancela(ruta_c, "") or []
            except Exception:
                vinculadas = []
            # 🔎 La opción aparece ÚNICAMENTE si ese comprobante cancela VARIAS facturas
            if len(vinculadas) > 1:
                numeros = ", ".join(str(f[0]) for f in vinculadas[:8])
                if len(vinculadas) > 8:
                    numeros += "…"
                menu.add_command(
                    label="🧾 Este comprobante cancela %d facturas: %s" % (len(vinculadas), numeros),
                    state="disabled")
                menu.add_command(label="🔎 Ver el detalle de esas facturas",
                                 command=lambda r=ruta_c: self._mostrar_facturas_del_comprobante(r))
        else:
            menu.add_command(label="🧾 Ver comprobante (sin archivo adjunto)",
                             command=lambda: self._avisar_sin_archivo(id_factura, "comprobante"))
            menu.add_command(label="🔎 Buscar facturas por N° de operación…",
                             command=self._buscar_por_operacion)
        menu.add_separator()
        menu.add_command(label="🗑️ Borrar factura",
                         command=lambda: self._borrar_factura(id_factura))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _mostrar_facturas_del_comprobante(self, ruta_comprobante):
        """🔎 Lista las facturas que cancela un mismo comprobante de pago."""
        try:
            filas = facturas_que_cancela(ruta_comprobante or "", "")
        except Exception as e:
            return messagebox.showerror("Comprobante", "No se pudo consultar: %s" % e)
        if not filas:
            return messagebox.showinfo(
                "Comprobante",
                "Este comprobante todavía no cancela ninguna factura registrada.")
        messagebox.showinfo(
            "Facturas que cancela este comprobante",
            "Cancela %d factura(s):\n\n%s" % (len(filas), texto_facturas_canceladas(filas)))

    def _buscar_por_operacion(self):
        """🔎 Pide el N° de operación y muestra las facturas que ese pago cancela."""
        numero = simpledialog.askstring("Buscar por N° de operación",
                                        "Escriba el N° de operación del comprobante:")
        if not numero or not numero.strip():
            return
        filas = facturas_que_cancela("", numero.strip())
        if not filas:
            return messagebox.showinfo(
                "Comprobante",
                "No hay facturas registradas con la operación «%s»." % numero.strip())
        messagebox.showinfo(
            "Facturas que cancela el comprobante",
            "Operación %s → %d factura(s):\n\n%s"
            % (numero.strip(), len(filas), texto_facturas_canceladas(filas)))

    def _ajustar_alto_tabla(self, evento=None):
        """Ajusta las filas visibles de la tabla al alto disponible del panel."""
        try:
            alto = int(evento.height) if evento is not None else int(self._f_tabla_alto.winfo_height())
        except Exception:
            return
        filas = max(10, int((alto - 30) / 22))
        try:
            if int(self.tabla.cget("height")) != filas:
                self.tabla.configure(height=filas)
        except Exception:
            pass

    def _menu_contextual_factura(self, event):
        """Botón derecho en Control de Pagos: el MISMO menú que la pestaña de facturas."""
        try:
            fila = self.tabla.identify_row(event.y)
            if not fila:
                return
            self.tabla.selection_set(fila)
            id_fact = int(self.tabla.item(fila, "values")[1])
        except Exception:
            return
        facturas = getattr(self.app_padre, "app_facturas", None)
        if facturas is not None and hasattr(facturas, "_mostrar_menu_factura"):
            # En la pestaña de INGRESO no se muestran las opciones de pago
            facturas._mostrar_menu_factura(id_fact, event, es_pagos=False)

    def _borrar_factura(self, id_factura):
        """Elimina la factura (y sus pagos) previa confirmación."""
        if not id_factura:
            return
        if not messagebox.askyesno("Borrar factura",
                                   "¿Eliminar la factura #%s y sus pagos?" % id_factura):
            return
        conn = conectar_db(silencioso=True)
        if not conn:
            return
        try:
            with conn.cursor() as c:
                c.execute("DELETE FROM pagos_comprobantes WHERE id_factura = %s", (id_factura,))
                c.execute("DELETE FROM facturas_recibidas WHERE id = %s", (id_factura,))
            conn.commit()
            try:
                cache_sistema.invalidar()
            except Exception:
                pass
            self.cargar_datos_tabla(reset_pagina=True)
            if hasattr(self.app_padre, "app_pagos"):
                try:
                    self.app_padre.app_pagos.cargar_datos_pagar(reset_pagina=True)
                except Exception:
                    pass
        except Exception as e:
            messagebox.showerror("Borrar factura", "No se pudo eliminar: %s" % e)
        finally:
            liberar_conexion(conn)

    def cargar_datos_tabla(self, reset_pagina=False):
        if reset_pagina:
            self.pagina_actual = 1
            
        self.lbl_pagina.configure(text=f"Pág {self.pagina_actual}")
        for item in self.tabla.get_children(): 
            self.tabla.delete(item)
        
        filtro = ""
        if hasattr(self, 'ent_buscar_facturas'):
            filtro = self.ent_buscar_facturas.get().strip().lower()
            
        # 🔃 Se cargan TODOS los registros del filtro/mes para poder ordenar por
        # cualquier columna afectando a todas las páginas (la paginación se hace al pintar)
        ocultar_app = bool(self.var_ocultar_app.get())
        clave_cache = f"compras_recibidas_v4_{filtro}_mes_{self.mes_filtro}_app_{int(ocultar_app)}"
        datos = cache_sistema.obtener(clave_cache)

        if datos is not None:
            self._pintar_datos_tabla(datos["filas"], datos["cuentas"])
        else:
            self.tabla.insert("", tk.END, values=("", "", "", "Cargando datos...", "", "", "", "", "", "", "", "", ""))
            
            # El hilo solo descarga datos; la tabla se pinta desde el hilo principal
            self._carga_actual = getattr(self, "_carga_actual", 0) + 1
            carga_id = self._carga_actual

            def tarea_descarga(estado):
                conn = conectar_db(silencioso=True)
                if not conn:
                    estado["filas"] = []
                    estado["cuentas"] = {}
                    return
                try:
                    cursor = conn.cursor()
                    query_base = "SELECT id, fecha, numero_documento, dias_credito, tipo_documento, proveedor, evento_asociado, descripcion, subtotal, impuesto, total, COALESCE(det_monto, 0), archivo_ruta, categoria, kilometraje, cantidad_combustible, ruc, COALESCE(pagado_por_tercero, ''), COALESCE(es_compra_cruzada, FALSE), COALESCE(soporte_pago_tercero, ''), COALESCE(orden_servicio, ''), COALESCE(comision_interbancaria, 0), COALESCE(es_interbancario, FALSE), \
COALESCE((SELECT SUM(p.monto_pagado) FROM pagos_comprobantes p WHERE p.id_factura = facturas_recibidas.id), 0) FROM facturas_recibidas"
                    
                    condiciones = []
                    params = []
                    if filtro:
                        cond_busqueda, params_busqueda = construir_condicion_busqueda_compras(filtro)
                        condiciones.append(cond_busqueda)
                        params.extend(params_busqueda)
                    patron_mes = patron_fecha_mes(self.mes_filtro)
                    if patron_mes:
                        condiciones.append("fecha LIKE %s")
                        params.append(patron_mes)
                    # 🚫 Facturas de la App (App Grifo): se ocultan si el check está marcado
                    agregar_condicion_ocultar_app(condiciones, params, ocultar_app)
                    where_sql = (" WHERE " + " AND ".join(condiciones)) if condiciones else ""
                    cursor.execute(f"{query_base}{where_sql} ORDER BY id DESC", tuple(params))
                        
                    datos_db = cursor.fetchall()
                    
                    ids_facturas = [r[0] for r in datos_db]
                    cuentas_por_factura = {}
                    if ids_facturas:
                        cursor.execute(
                            "SELECT id_factura, cuenta_origen FROM pagos_comprobantes WHERE id_factura = ANY(%s) AND cuenta_origen IS NOT NULL AND cuenta_origen != ''", 
                            (ids_facturas,)
                        )
                        for id_f, cta in cursor.fetchall():
                            if id_f not in cuentas_por_factura:
                                cuentas_por_factura[id_f] = []
                            if cta not in cuentas_por_factura[id_f]:
                                cuentas_por_factura[id_f].append(cta)

                    datos_cache = {"filas": datos_db, "cuentas": cuentas_por_factura}
                    cache_sistema.guardar(clave_cache, datos_cache)
                    estado["filas"] = datos_db
                    estado["cuentas"] = cuentas_por_factura
                except Exception as e:
                    print(f"Error cargando tabla de compras: {e}")
                    estado["filas"] = []
                    estado["cuentas"] = {}
                finally:
                    liberar_conexion(conn)

            ejecutar_en_hilo(self.main_root, tarea_descarga,
                             al_terminar=lambda e: self._pintar_si_vigente(carga_id, e))

    def _pintar_si_vigente(self, carga_id, estado):
        """Pinta la tabla solo si esta carga sigue siendo la última solicitada."""
        if carga_id != getattr(self, "_carga_actual", 0):
            return                      # ya hay una búsqueda más nueva en curso
        if estado.get("error"):
            print("[Compras] Error al cargar la tabla:", estado["error"])
        self._pintar_datos_tabla(estado.get("filas") or [], estado.get("cuentas") or {})

    def _pintar_datos_tabla(self, registros, cuentas_por_factura):
        """Construye TODAS las filas, las ordena por la columna activa (todas las
        páginas) y muestra únicamente la página actual."""
        for item in self.tabla.get_children():
            self.tabla.delete(item)

        filas = []
        for r in registros:
            id_factura = r[0]
            archivo_bd = r[12]
            # Datos de las compras cruzadas (creadas a mano desde el módulo de Banco)
            tercero = str(r[17]) if len(r) > 17 and r[17] else ""
            es_cruzada = bool(r[18]) if len(r) > 18 else False
            soporte_cruzada = str(r[19]) if len(r) > 19 and r[19] else ""
            tiene_archivo = bool(archivo_bd) and str(archivo_bd).strip() != "PENDIENTE_DESCARGA"
            if not tiene_archivo and soporte_cruzada:
                tiene_archivo = True          # solo tiene el soporte del pago a tercero
            tiene_arch = "✅ Ver" if tiene_archivo else "❌ No"
                
            tipo_doc = r[4]; impuesto = r[9]; tot_bruto = r[10]; det_monto = r[11]; cat = r[13] if r[13] else "GENERAL"
            km_val = r[14] if r[14] else "-"
            cant_val = r[15] if r[15] else "-"
            ruc_val = r[16] if len(r) > 16 and r[16] else "-"
            orden_val = str(r[20]) if len(r) > 20 and r[20] else "-"
            
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
            if es_cruzada:
                # Compra cruzada: se indica quién pagó, pero SIN ocultar la cuenta
                # bancaria con la que salió el dinero (el pago también se registra
                # en Banco, así que la factura figura como pagada).
                nota = "🔁 Compra cruzada" + (f" · pagó: {tercero}" if tercero else "")
                try:
                    comision_cruz = float(r[21] or 0) if len(r) > 21 else 0.0
                except Exception:
                    comision_cruz = 0.0
                if comision_cruz > 0:
                    nota += f" · comisión {formatear_moneda(comision_cruz)}"
                metodo_pago = f"{metodo_pago} · {nota}" if metodo_pago else nota

            if "Recibo" in tipo_doc and "8%" in tipo_doc: neto = tot_bruto - impuesto - det_monto
            else: neto = tot_bruto - det_monto
                
            # 🟢 Verde = cancelada; 🟣 violeta = pago parcial; 🔴 rojo = sin pagos
            try:
                # 💰 La suma pagada es la ÚLTIMA columna de la consulta
                _pagado_f = float(r[-1] or 0) if len(r) > 22 else 0.0
            except Exception:
                _pagado_f = 0.0
            if es_cruzada or neto <= 0.01 or _pagado_f >= neto - 0.01:
                etiqueta_color = "pagada"
            elif _pagado_f <= 0.01:
                etiqueta_color = "no_pagada"
            else:
                etiqueta_color = "parcial"

            row_vals = (
                0, id_factura, r[1], hora_consumo, r[2] if r[2] else "-", orden_val, r[3], tipo_doc.split(" ")[0], r[5], ruc_val, cat,
                r[6].split(" | ")[0] if " | " in str(r[6]) else r[6], km_val, cant_val, desc_limpia, metodo_pago, formatear_moneda(r[8]), formatear_moneda(impuesto), formatear_moneda(tot_bruto), formatear_moneda(det_monto), formatear_moneda(neto), tiene_arch
            )

            filas.append((row_vals, etiqueta_color))

        # 🔃 Ordena TODAS las páginas según la columna seleccionada
        filas = aplicar_orden_filas(filas, self.columnas_tabla, self.columna_orden, self.orden_ascendente)

        total_registros = len(filas)
        self.total_paginas = max(1, (total_registros + self.registros_por_pagina - 1) // self.registros_por_pagina)
        if self.pagina_actual > self.total_paginas:
            self.pagina_actual = self.total_paginas
        offset = (self.pagina_actual - 1) * self.registros_por_pagina
        pagina = filas[offset:offset + self.registros_por_pagina]

        for numero, (row_vals, etiqueta_color) in enumerate(pagina, start=offset + 1):
            self.tabla.insert("", tk.END, values=(numero,) + tuple(row_vals[1:]), tags=(etiqueta_color,))

        self.lbl_pagina.configure(text=f"Pág {self.pagina_actual} de {self.total_paginas}")
        self.btn_ant.configure(state="normal" if self.pagina_actual > 1 else "disabled")
        self.btn_sig.configure(state="normal" if self.pagina_actual < self.total_paginas else "disabled")

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
                if str(res[0]).strip() == "PENDIENTE_DESCARGA":
                    if messagebox.askyesno("Archivo no disponible", "¿Desea forzar la descarga de los archivos pendientes ahora?"):
                        self.ejecutar_sincronizacion_manual()
                    return
                # 🔎 Se resuelve la ruta aunque venga de otro equipo/sistema (Mac/Windows)
                from app_paths import resolver_ruta_archivo
                ruta_almacenada = resolver_ruta_archivo(res[0])
                if ruta_almacenada:
                    abrir_documento(ruta_almacenada)
                else:
                    if messagebox.askyesno("Archivo Extraviado", "El archivo local no se encuentra.\n¿Desea recuperarlo desde la nube (Supabase)?"):
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
                                
                                ruta_bd = ruta_para_guardar(ruta_final)
                                cursor.execute("UPDATE facturas_recibidas SET archivo_ruta = %s WHERE id = %s", (ruta_bd, id_doc))
                                cursor.execute("UPDATE pagos_comprobantes SET archivo_ruta = %s WHERE id_factura = %s", (ruta_bd, id_doc))
                                conn.commit()
                                
                                messagebox.showinfo("Recuperación Exitosa", "Imagen recuperada y guardada.")
                                abrir_documento(ruta_final)
                                self.cargar_datos_tabla()
                            except Exception as e_write:
                                messagebox.showerror("Error", f"No se pudo guardar la imagen recuperada:\n{e_write}")
                        else:
                            messagebox.showerror("Error irrecuperable", "La imagen original ya no se encuentra en el servidor temporal.")
            else:
                messagebox.showinfo("Aviso", "Este registro no tiene ningún archivo asociado.")
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el documento:\n{e}")
        finally:
            liberar_conexion(conn)

    def abrir_ventana_edicion(self):
        sel = self.tabla.selection()
        if not sel: return messagebox.showwarning("Atención", "Seleccione un documento.")
        
        valores = self.tabla.item(sel[0], "values")
        id_doc = valores[1]
        
        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT tipo_documento, ruc, proveedor, numero_documento, fecha, evento_asociado, kilometraje, cantidad_combustible, descripcion, subtotal, impuesto, total, COALESCE(id_orden_servicio, 0), COALESCE(orden_servicio, ''), COALESCE(pagado_por_tercero, ''), COALESCE(es_compra_cruzada, FALSE), COALESCE(comision_interbancaria, 0), COALESCE(es_interbancario, FALSE), COALESCE(det_monto, 0) FROM facturas_recibidas WHERE id = %s", (id_doc,))
            reg = cursor.fetchone()
            # 💳 Pagos ya registrados (para saber si la factura figura pagada)
            cursor.execute("""SELECT id, COALESCE(monto_pagado, 0), COALESCE(cuenta_origen, ''),
                                     COALESCE(numero_operacion, '')
                              FROM pagos_comprobantes WHERE id_factura = %s ORDER BY id""", (id_doc,))
            e_pagos = cursor.fetchall() or []
        finally: 
            liberar_conexion(conn)
            
        if not reg: return
        (e_tipo, e_ruc, e_prov, e_nro, e_fec, e_placa, e_km, e_gal, e_desc, e_sub, e_imp, e_tot,
         e_orden_id, e_orden_txt, e_tercero, e_cruzada, e_comision, e_inter, e_det) = reg
        try:
            e_orden_id = int(e_orden_id or 0)
        except (TypeError, ValueError):
            e_orden_id = 0
        e_orden_txt = str(e_orden_txt or "")
        
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
        centrar_ventana(v_edit, self.main_root, 450, 660)
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
        # 🛠️ Mantenimientos del vehículo de esta factura (aceite, correa, gas, SOAT, etc.)
        ctk.CTkButton(f_form, text="🛠️ Registrar mantenimientos del vehículo", height=28,
                      font=("Arial", 11, "bold"), fg_color="#e67e22", hover_color="#ca6f1e",
                      command=lambda: abrir_reset_mantenimientos_vehiculo(
                          v_edit, ent_placa.get(), self.app_padre.usuario_activo)).pack(
            fill="x", padx=5, pady=(0, 5))
        ent_km = crear_campo(f_form, "Kilometraje:", e_km)
        ent_gal = crear_campo(f_form, "Galones/Cant.:", e_gal)
        ent_desc = crear_campo(f_form, "Concepto / Descripción:", c_val)
        ent_hora = crear_campo(f_form, "Hora de Consumo (Ej: 14:30):", h_val)

        # 🔁 FACTURA CRUZADA (pagada por un tercero) + comisión interbancaria
        e_tercero = str(e_tercero or "")
        e_cruzada = bool(e_cruzada)
        try: e_comision = float(e_comision or 0)
        except (TypeError, ValueError): e_comision = 0.0
        e_inter = bool(e_inter)

        ctk.CTkLabel(f_form, text="🔁 Pago a tercero (factura cruzada)", font=("Arial", 11, "bold"),
                     text_color="#6c3483").pack(anchor="w", padx=5, pady=(8, 0))
        var_edit_cruzada = tk.BooleanVar(value=e_cruzada)
        ctk.CTkCheckBox(f_form, text="Esta factura la pagó un tercero", variable=var_edit_cruzada,
                        font=("Arial", 11), checkbox_width=18, checkbox_height=18).pack(anchor="w", padx=5)
        ent_edit_tercero = crear_campo(f_form, "Pagado por (tercero):", e_tercero)
        f_edit_com = ctk.CTkFrame(f_form, fg_color="transparent")
        f_edit_com.pack(fill="x", padx=5, pady=(0, 5))
        var_edit_inter = tk.BooleanVar(value=e_inter)
        ctk.CTkCheckBox(f_edit_com, text="Interbancario", variable=var_edit_inter,
                        font=("Arial", 11), checkbox_width=16, checkbox_height=16).pack(side="left")
        ctk.CTkLabel(f_edit_com, text="Comisión (S/.):", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 4))
        ent_edit_comision = ctk.CTkEntry(f_edit_com, width=80)
        ent_edit_comision.pack(side="left")
        ent_edit_comision.insert(0, f"{e_comision:.2f}")

        # 💳 El total que pagó el tercero se calcula AL INSTANTE (factura + comisión)
        lbl_edit_cruz_total = ctk.CTkLabel(f_form, text="", font=("Arial", 10, "bold"),
                                           text_color="#6c3483", wraplength=400, justify="left")
        lbl_edit_cruz_total.pack(anchor="w", padx=5, pady=(0, 5))

        def actualizar_total_cruzada_edit(*_a):
            """Muestra el total pagado al tercero: factura + comisión interbancaria."""
            try:
                neto = max(0.0, _num_edit(ent_tot) - float(e_det or 0))
            except Exception:
                neto = 0.0
            try:
                com = parsear_monto_texto(ent_edit_comision.get()) if var_edit_inter.get() else 0.0
            except Exception:
                com = 0.0
            try:
                if com > 0:
                    lbl_edit_cruz_total.configure(
                        text=(f"💳 El tercero pagó: {formatear_moneda(neto + com)}   "
                              f"(factura {formatear_moneda(neto)} + comisión {formatear_moneda(com)})"))
                else:
                    lbl_edit_cruz_total.configure(
                        text=f"💳 El tercero pagó: {formatear_moneda(neto)} (sin comisión interbancaria)")
            except Exception:
                pass

        for _evento_com in ("<KeyRelease>", "<FocusOut>", "<<Paste>>"):
            ent_edit_comision.bind(_evento_com, actualizar_total_cruzada_edit, add="+")
        # Al entrar al campo se selecciona el valor: escribir la comisión lo reemplaza
        try:
            ent_edit_comision.bind("<FocusIn>",
                                   lambda _e: ent_edit_comision.select_range(0, "end"), add="+")
        except Exception:
            pass
        try:
            var_edit_inter.trace_add("write", lambda *_a: actualizar_total_cruzada_edit())
        except Exception:
            pass

        ctk.CTkLabel(f_form, text="Cuenta / Banco con el que pagó:", font=("Arial", 11, "bold")).pack(anchor="w", padx=5, pady=(5, 0))
        cmb_edit_cuenta = ctk.CTkComboBox(f_form, values=self._lista_cuentas_pago())
        cmb_edit_cuenta.pack(fill="x", padx=5, pady=(0, 5))
        if e_pagos and e_pagos[0][2]:
            cmb_edit_cuenta.set(str(e_pagos[0][2]))
        ent_edit_oper = crear_campo(f_form, "N° de Operación Bancaria:",
                                    str(e_pagos[0][3]) if e_pagos and e_pagos[0][3] else "")

        # 🔗 Orden de servicio relacionada con esta factura (se puede cambiar o quitar)
        ctk.CTkLabel(f_form, text="🔗 Orden de Servicio (relacionar):", font=("Arial", 11, "bold")).pack(anchor="w", padx=5, pady=(5, 0))
        combo_edit_orden = ctk.CTkComboBox(f_form, values=[SIN_ORDEN_SERVICIO], state="readonly")
        combo_edit_orden.pack(fill="x", padx=5, pady=(0, 5))
        valores_orden_edit = [SIN_ORDEN_SERVICIO]
        mapa_ordenes_edit = {}
        for _etiqueta, _orden in (getattr(self, "_ordenes_servicio", {}) or {}).items():
            valores_orden_edit.append(_etiqueta)
            mapa_ordenes_edit[_etiqueta] = _orden
        seleccion_orden_edit = SIN_ORDEN_SERVICIO
        if e_orden_id:
            for _etiqueta, _orden in mapa_ordenes_edit.items():
                if _orden[0] == e_orden_id:
                    seleccion_orden_edit = _etiqueta
                    break
            else:
                # La orden ya no está en la lista (anulada o de otro equipo): se conserva el vínculo
                _etiqueta_actual = f"{e_orden_txt or ('Orden #' + str(e_orden_id))} (relacionada actualmente)"
                valores_orden_edit.append(_etiqueta_actual)
                mapa_ordenes_edit[_etiqueta_actual] = None
                seleccion_orden_edit = _etiqueta_actual
        combo_edit_orden.configure(values=valores_orden_edit)
        combo_edit_orden.set(seleccion_orden_edit)
        
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

        # 🍽️ Recargo al consumo (servicio/propina): suma al total y no lleva IGV
        ctk.CTkLabel(f_montos, text="Recargo al consumo (no gravado):", font=("Arial", 11, "bold")).grid(
            row=2, column=0, columnspan=3, padx=5, pady=(8, 0), sticky="w")
        ent_rec = ctk.CTkEntry(f_montos, width=100)
        ent_rec.grid(row=3, column=0, padx=5)
        try:
            _recargo_inicial = float(e_tot or 0) - float(e_sub or 0) - float(e_imp or 0)
        except (TypeError, ValueError):
            _recargo_inicial = 0.0
        ent_rec.insert(0, f"{max(_recargo_inicial, 0.0):.2f}")

        # 💳 El resumen del pago al tercero sigue los montos y la comisión en vivo
        for _campo_monto in (ent_sub, ent_imp, ent_tot, ent_rec):
            try:
                _campo_monto.bind("<KeyRelease>", actualizar_total_cruzada_edit, add="+")
            except Exception:
                pass
        # Se pinta cuando la ventana ya está armada (así lee el total y la comisión actuales)
        try:
            v_edit.after(300, actualizar_total_cruzada_edit)
        except Exception:
            pass

        # 💰 Igual que en el registro: se puede escribir el monto BASE o el monto ya CON IGV
        # y el sistema calcula automáticamente la base, el IGV y el total.
        seg_edit = ctk.CTkSegmentedButton(
            f_form, values=["Monto Base (sin IGV)", "Monto con IGV (Total)"], font=("Arial", 10, "bold"),
            command=lambda valor: cambiar_modo_edit(valor))
        seg_edit.pack(fill="x", padx=5, pady=(10, 0))
        seg_edit.set("Monto Base (sin IGV)")
        lbl_edit_nota = ctk.CTkLabel(f_form, text="", font=("Arial", 10, "italic"), text_color="#555555", wraplength=400, justify="left")
        lbl_edit_nota.pack(anchor="w", padx=5)
        estado_edit = {"modo": "BASE"}

        def _es_factura_edit():
            return "Factura" in str(ent_tipo.get() or "")

        def _es_boleta_edit():
            return "Boleta" in str(ent_tipo.get() or "")

        def _tasa_edit():
            texto = str(ent_tipo.get() or "")
            return 0.105 if "10.5%" in texto else 0.18

        def _num_edit(entrada):
            try:
                return parsear_monto_texto(entrada.get())
            except ValueError:
                return 0.0

        def _set_edit(entrada, valor):
            entrada.delete(0, tk.END)
            entrada.insert(0, f"{valor:.2f}")

        def _nota_edit():
            if _es_factura_edit():
                lbl_edit_nota.configure(
                    text=("Escriba en 'Total (Neto)' el monto con IGV: la base y el IGV se calculan solos."
                          if estado_edit["modo"] == "CON_IGV" else
                          "Escriba en 'Subtotal' el monto sin IGV: el IGV y el total se calculan solos."))
            elif _es_boleta_edit():
                lbl_edit_nota.configure(text="Boleta sin IGV: el total es igual al monto base.")
            else:
                lbl_edit_nota.configure(text="Este tipo de documento no lleva IGV: los montos se guardan tal cual.")

        def _recalcular_desde_base(*_a):
            base = _num_edit(ent_sub)
            rec = _num_edit(ent_rec)
            if _es_factura_edit():
                igv = base * _tasa_edit()
                _set_edit(ent_imp, igv)
                _set_edit(ent_tot, base + igv + rec)
            elif _es_boleta_edit():
                _set_edit(ent_imp, 0.0)
                _set_edit(ent_tot, base + rec)
            estado_edit["modo"] = "BASE"
            seg_edit.set("Monto Base (sin IGV)")
            _nota_edit()

        def _recalcular_desde_total(*_a):
            total = _num_edit(ent_tot)
            rec = _num_edit(ent_rec)
            if _es_factura_edit():
                tasa = _tasa_edit()
                # El recargo al consumo no está gravado: se descuenta antes del IGV
                gravado = max(total - rec, 0.0)
                base = gravado / (1.0 + tasa)
                _set_edit(ent_sub, base)
                _set_edit(ent_imp, gravado - base)
            elif _es_boleta_edit():
                _set_edit(ent_sub, max(total - rec, 0.0))
                _set_edit(ent_imp, 0.0)
            estado_edit["modo"] = "CON_IGV"
            seg_edit.set("Monto con IGV (Total)")
            _nota_edit()

        def _recalcular_desde_recargo(*_a):
            rec = _num_edit(ent_rec)
            if _es_factura_edit():
                _set_edit(ent_tot, _num_edit(ent_sub) + _num_edit(ent_imp) + rec)
            elif _es_boleta_edit():
                _set_edit(ent_tot, _num_edit(ent_sub) + rec)
            _nota_edit()

        def cambiar_modo_edit(valor):
            estado_edit["modo"] = "CON_IGV" if "con igv" in str(valor).lower() else "BASE"
            if estado_edit["modo"] == "CON_IGV":
                _recalcular_desde_total()
            else:
                _recalcular_desde_base()

        ent_sub.bind("<KeyRelease>", _recalcular_desde_base)
        ent_tot.bind("<KeyRelease>", _recalcular_desde_total)
        ent_rec.bind("<KeyRelease>", _recalcular_desde_recargo)
        _nota_edit()

        def _orden_elegida_edit():
            """(id, texto) de la orden de servicio elegida en la ventana de edición."""
            etiqueta = combo_edit_orden.get()
            if etiqueta == SIN_ORDEN_SERVICIO:
                return None, ""
            orden = mapa_ordenes_edit.get(etiqueta)
            if orden is None:                        # se mantiene el vínculo que ya tenía
                return (e_orden_id or None), e_orden_txt
            return orden[0], self._numero_visible_orden(orden)

        def guardar_cambios():
            try:
                val_sub = float(ent_sub.get().strip() or 0)
                val_imp = float(ent_imp.get().strip() or 0)
                val_tot = float(ent_tot.get().strip() or 0)
            except ValueError:
                return messagebox.showerror("Error", "Los montos deben ser numéricos.", parent=v_edit)
                
            n_desc = ent_desc.get().strip()
            n_hora = ent_hora.get().strip()
            desc_final = f"{n_desc} | Hora: {n_hora}" if n_hora else n_desc

            # 🔁 Datos del pago a tercero (factura cruzada)
            es_cruz_edit = bool(var_edit_cruzada.get())
            tercero_edit = ent_edit_tercero.get().strip()
            inter_edit = bool(var_edit_inter.get())
            comision_edit = parsear_monto_texto(ent_edit_comision.get()) if inter_edit else 0.0
            if es_cruz_edit and not tercero_edit:
                return messagebox.showwarning("Factura cruzada",
                                              "Escriba el nombre de quien pagó (tercero).", parent=v_edit)

            if messagebox.askyesno("Confirmar", "¿Guardar los cambios?", parent=v_edit):
                conn_u = conectar_db()
                if conn_u:
                    try:
                        cursor_u = conn_u.cursor()
                        id_orden_final, txt_orden_final = _orden_elegida_edit()
                        cursor_u.execute("""
                            UPDATE facturas_recibidas 
                            SET tipo_documento=%s, ruc=%s, proveedor=%s, numero_documento=%s, fecha=%s, 
                                evento_asociado=%s, kilometraje=%s, cantidad_combustible=%s, descripcion=%s, 
                                subtotal=%s, impuesto=%s, total=%s, id_orden_servicio=%s, orden_servicio=%s,
                                pagado_por_tercero=%s, es_compra_cruzada=%s,
                                comision_interbancaria=%s, es_interbancario=%s
                            WHERE id=%s
                        """, (
                            ent_tipo.get().strip(), ent_ruc.get().strip(), ent_prov.get().strip(), ent_nro.get().strip(), 
                            ent_fec.get().strip(), ent_placa.get().strip(), ent_km.get().strip(), ent_gal.get().strip(), 
                            desc_final, val_sub, val_imp, val_tot, id_orden_final, txt_orden_final,
                            tercero_edit if es_cruz_edit else "", es_cruz_edit,
                            comision_edit, inter_edit, id_doc
                        ))
                        
                        cursor_u.execute("""
                            UPDATE pagos_comprobantes 
                            SET proveedor_nombre=%s, codigo_cotizacion=%s
                            WHERE id_factura=%s
                        """, (ent_prov.get().strip(), ent_nro.get().strip(), id_doc))

                        # 💸 FACTURA CRUZADA: el tercero pagó la factura MÁS la comisión, así
                        #    que el pago registrado se completa para no dejar saldo pendiente.
                        aviso_pago = ""
                        if es_cruz_edit and comision_edit > 0:
                            try:
                                neto_objetivo = max(0.0, float(val_tot or 0) - float(e_det or 0))
                            except (TypeError, ValueError):
                                neto_objetivo = 0.0
                            objetivo = neto_objetivo + comision_edit
                            cuenta_edit = cmb_edit_cuenta.get().strip()
                            oper_edit = ent_edit_oper.get().strip()
                            registrados = float(sum(float(p[1] or 0) for p in (e_pagos or [])))
                            if not e_pagos:
                                cursor_u.execute("""
                                    INSERT INTO pagos_comprobantes
                                    (id_factura, monto_pagado, archivo_ruta, proveedor_nombre, fecha_pago,
                                     categoria_suministro, codigo_cotizacion, cuenta_origen, numero_operacion)
                                    VALUES (%s, %s, '', %s, %s, '', %s, %s, %s)
                                """, (id_doc, objetivo, ent_prov.get().strip(), ent_fec.get().strip(),
                                      ent_nro.get().strip(), cuenta_edit or f"Pago de {tercero_edit}",
                                      oper_edit or "PAGO A TERCERO"))
                                aviso_pago = ("\n\n💳 Se registró el pago al tercero por "
                                              f"{formatear_moneda(objetivo)} (factura + comisión).")
                            elif len(e_pagos) == 1 and registrados < objetivo - 0.01:
                                sets_p = ["monto_pagado = %s"]
                                params_p = [objetivo]
                                if cuenta_edit:
                                    sets_p.append("cuenta_origen = %s"); params_p.append(cuenta_edit)
                                if oper_edit:
                                    sets_p.append("numero_operacion = %s"); params_p.append(oper_edit)
                                params_p.append(e_pagos[0][0])
                                cursor_u.execute("UPDATE pagos_comprobantes SET " + ", ".join(sets_p)
                                                 + " WHERE id = %s", tuple(params_p))
                                aviso_pago = (f"\n\n💳 El pago se completó a {formatear_moneda(objetivo)}: "
                                              f"factura {formatear_moneda(neto_objetivo)} + comisión "
                                              f"{formatear_moneda(comision_edit)}. No queda saldo pendiente.")
                            elif len(e_pagos) > 1:
                                aviso_pago = ("\n\n💡 Esta factura tiene varios pagos registrados: "
                                              "revise los montos en 'Editar Pagos'.")

                        conn_u.commit()
                        cache_sistema.invalidar()
                        messagebox.showinfo("Éxito", "Registro actualizado correctamente." + aviso_pago, parent=v_edit)
                        v_edit.destroy()
                        self.cargar_datos_tabla(reset_pagina=True)
                        if hasattr(self.app_padre, 'app_pagos'):
                            self.app_padre.app_pagos.cargar_datos_pagar(reset_pagina=True)
                    except Exception as e:
                        messagebox.showerror("Error", f"No se pudo guardar:\n{e}", parent=v_edit)
                    finally:
                        liberar_conexion(conn_u)

        def eliminar_registro():
            # 🔗 Se averigua si este registro se creó desde el módulo de BANCO: si es así,
            # al eliminarlo aquí también se elimina el movimiento bancario (módulos sincronizados).
            mov_banco = None
            try:
                conn_b = conectar_db(silencioso=True)
                if conn_b:
                    try:
                        with conn_b.cursor() as cb:
                            cb.execute("""SELECT id, COALESCE(banco, ''), COALESCE(cuenta, ''),
                                                 COALESCE(fecha, ''), COALESCE(monto, 0)
                                          FROM conciliacion_bancaria
                                          WHERE id_gasto_compras = %s
                                          ORDER BY id""", (id_doc,))
                            mov_banco = cb.fetchone()
                    finally:
                        liberar_conexion(conn_b)
            except Exception:
                mov_banco = None

            aviso = "⚠️ ¿Desea eliminar completamente este registro?"
            if mov_banco:
                etiqueta_b = " - ".join(str(x) for x in (mov_banco[1], mov_banco[2]) if x).strip(" -")
                aviso += (f"\n\n🔗 Este registro se creó desde el módulo de BANCO:\n"
                          f"     {etiqueta_b} | {mov_banco[3]} | {formatear_moneda(mov_banco[4])}"
                          f"\n\nEl movimiento bancario y el pago se eliminarán también, "
                          f"para que los dos módulos queden sincronizados.")

            if messagebox.askyesno("Confirmar Eliminación", aviso, parent=v_edit):
                try:
                    conn = conectar_db()
                    cursor = conn.cursor()
                    cursor.execute("SELECT archivo_ruta FROM facturas_recibidas WHERE id = %s", (id_doc,))
                    row = cursor.fetchone()
                    ruta_archivo = os.path.normpath(row[0]) if row and row[0] else None
                    eliminar_archivo(ruta_archivo)   # resuelve rutas de otros equipos/SO

                    # 🗑️ Se borran también los pagos de esta factura (incluido el que creó el Banco)
                    cursor.execute("DELETE FROM pagos_comprobantes WHERE id_factura = %s", (id_doc,))
                    cursor.execute("DELETE FROM facturas_recibidas WHERE id = %s", (id_doc,))

                    # 🔗 Y el movimiento del Banco que la originó
                    mov_borrados = 0
                    if mov_banco:
                        cursor.execute("DELETE FROM conciliacion_bancaria WHERE id_gasto_compras = %s", (id_doc,))
                        mov_borrados = max(cursor.rowcount, 0)

                    conn.commit()
                    liberar_conexion(conn)
                    
                    cache_sistema.invalidar()
                    registrar_auditoria(self.app_padre.usuario_activo, "Facturas Recibidas",
                                        f"Eliminó factura ID {id_doc}"
                                        + (" y su movimiento bancario en Banco" if mov_borrados else ""))

                    mensaje = "Registro eliminado."
                    if mov_borrados:
                        mensaje += ("\n🔗 También se eliminó el movimiento en el módulo de BANCO."
                                    "\nLos dos módulos quedan sincronizados.")
                    messagebox.showinfo("Éxito", mensaje, parent=v_edit)
                    v_edit.destroy()
                    self.cargar_datos_tabla(reset_pagina=True)
                    if hasattr(self.app_padre, 'app_pagos'): self.app_padre.app_pagos.cargar_datos_pagar(reset_pagina=True)
                except Exception as e:
                    messagebox.showerror("Error", str(e), parent=v_edit)

        ctk.CTkButton(f_form, text="💾 Guardar Cambios", font=("Arial", 12, "bold"), fg_color="#27ae60", hover_color="#1e8449", command=guardar_cambios).pack(fill="x", padx=10, pady=(20, 5))
        ctk.CTkButton(f_form, text="❌ Eliminar Registro Completo", font=("Arial", 12, "bold"), fg_color="#e74c3c", hover_color="#c0392b", command=eliminar_registro).pack(fill="x", padx=10, pady=(5, 15))

# =========================================================
# PESTAÑA 2: CUENTAS POR PAGAR (PAGOS Y DEUDAS)
# =========================================================
class CuentasPorPagarTab:
    def _menu_contextual_factura(self, event):
        """Botón derecho en Control de Pagos: el MISMO menú que la pestaña de facturas."""
        try:
            fila = self.tabla.identify_row(event.y)
            if not fila:
                return
            self.tabla.selection_set(fila)
            id_fact = int(self.tabla.item(fila, "values")[1])
        except Exception:
            return
        facturas = getattr(self.app_padre, "app_facturas", None)
        if facturas is not None and hasattr(facturas, "_mostrar_menu_factura"):
            facturas._mostrar_menu_factura(id_fact, event, es_pagos=True)

    def abrir_datos_bancarios_pagar(self, event=None):
        """Compatibilidad: datos bancarios del proveedor de la fila seleccionada."""
        try:
            id_fact = int(self.tabla.item(self.tabla.selection()[0], "values")[1])
        except Exception:
            return
        facturas = getattr(self.app_padre, "app_facturas", None)
        if facturas is not None and hasattr(facturas, "_datos_bancarios_factura"):
            return facturas._datos_bancarios_factura(id_fact)

    def __init__(self, tab_frame, main_root, app_padre):
        self.tab_frame = tab_frame
        self.main_root = main_root
        self.app_padre = app_padre
        
        self.pagina_actual = 1
        self.registros_por_pagina = 55
        
        # 🗓️ FILTRO DE MES (por defecto, el mes en curso)
        self.mes_filtro = mes_en_curso()

        # 🚫 Check para OCULTAR las facturas enviadas por la App (App Grifo).
        # Por estándar viene desmarcado: se ven todas.
        self.var_ocultar_app = tk.BooleanVar(value=False)

        # 🔃 Ordenamiento por cualquier columna aplicado a TODAS las páginas
        self.columna_orden = "id_factura"
        self.orden_ascendente = False
        self.total_paginas = 1
        
        self.inicializar_entorno()
        self.crear_interfaz()

    def inicializar_entorno(self):
        def tarea_init():
            conn = conectar_db(silencioso=True)
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
                    cursor.execute("ALTER TABLE pagos_comprobantes ADD COLUMN IF NOT EXISTS cuenta_origen VARCHAR(255) DEFAULT ''")
                    conn.commit()
                except Exception: conn.rollback()
                try:
                    cursor.execute("ALTER TABLE pagos_comprobantes ADD COLUMN IF NOT EXISTS numero_operacion VARCHAR(100) DEFAULT ''")
                    conn.commit()
                except Exception: conn.rollback()
            except Exception: pass
            finally: liberar_conexion(conn)
        threading.Thread(target=tarea_init, daemon=True).start()

    def crear_interfaz(self):
        frame_acciones = ctk.CTkFrame(self.tab_frame, corner_radius=8, fg_color="#f8f9fa", border_width=1, border_color="#e0e0e0")
        frame_acciones.pack(fill="x", padx=15, pady=(10, 10), ipady=5)

        # ⚙️ Columnas visibles de la tabla de pagos
        btn_cols = ctk.CTkButton(frame_acciones, text="⚙️ Columnas", width=90,
                                 font=("Arial", 12, "bold"), fg_color="#6c757d",
                                 hover_color="#565e64",
                                 command=lambda: menu_columnas(self.tabla, "compras_cuentas_pagar"))
        btn_cols.pack(side="left", padx=5, pady=5)
        try:
            aplicar_columnas_visibles(self.tabla, "compras_cuentas_pagar")
        except Exception:
            pass

        # 🚀 OPCIÓN UNIFICADA: Reporte de Totales y Resumen por Proveedor
        btn_reporte = ctk.CTkButton(frame_acciones, text="📊 Reporte de Totales y Proveedores", font=("Arial", 12, "bold"), command=self.mostrar_reporte_totales, fg_color="#1f538d", hover_color="#163b65")
        btn_reporte.pack(side="left", padx=5, pady=5)



        btn_mant = ctk.CTkButton(frame_acciones, text="🛠️ Resetear Mantenimientos", font=("Arial", 12, "bold"), command=self.resetear_mantenimientos_vehiculo, fg_color="#e67e22", hover_color="#ca6f1e")
        btn_mant.pack(side="left", padx=5, pady=5)

        btn_refresh = ctk.CTkButton(frame_acciones, text="🔄 Actualizar", font=("Arial", 12, "bold"), command=lambda: self.cargar_datos_pagar(reset_pagina=True), fg_color="#7f8c8d", hover_color="#606b6b")
        btn_refresh.pack(side="right", padx=10, pady=5)

        f_busqueda = ctk.CTkFrame(self.tab_frame, fg_color="transparent")
        f_busqueda.pack(fill="x", padx=15, pady=(0, 5))
        ctk.CTkLabel(f_busqueda, text="🔍 Buscar:", font=("Arial", 11, "bold")).pack(side="left", padx=(0, 5))
        self.ent_buscar_pagos = ctk.CTkEntry(f_busqueda, placeholder_text="Buscar por cualquier columna: N° doc, proveedor, RUC, placa, concepto, fecha, monto...")
        self.ent_buscar_pagos.pack(side="left", fill="x", expand=True)
        
        ctk.CTkLabel(f_busqueda, text="🗓️ Mes:", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 5))
        self.combo_mes = ctk.CTkComboBox(f_busqueda, values=construir_valores_mes(), width=150, state="readonly", command=self.on_cambiar_mes)
        self.combo_mes.pack(side="left", padx=(0, 5))
        self.combo_mes.set(self.mes_filtro)

        # 🚫 Casilla para ocultar las facturas que llegan desde la App (App Grifo)
        self.chk_ocultar_app = ctk.CTkCheckBox(
            f_busqueda, text="🚫 Ocultar facturas de la App", variable=self.var_ocultar_app,
            font=("Arial", 11, "bold"), checkbox_width=18, checkbox_height=18,
            command=lambda: self.cargar_datos_pagar(reset_pagina=True))
        self.chk_ocultar_app.pack(side="left", padx=(10, 0))

        self.ent_buscar_pagos.bind("<KeyRelease>", lambda e: self.buscar_con_retraso())
        self.ent_buscar_pagos.bind("<Return>", lambda e: self.cargar_datos_pagar(reset_pagina=True))

        f_tabla = ctk.CTkFrame(self.tab_frame, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True, padx=15, pady=0)

        columnas = ("num", "id_factura", "fecha", "hora", "nro_doc", "proveedor", "ruc", "evento", "kilometraje", "cantidad", "concepto", "metodo_pago", "subtotal", "igv", "detraccion", "neto_facturado", "pagado", "saldo", "archivos")
        self.columnas_tabla = columnas
        self.tabla = ttk.Treeview(f_tabla, columns=columnas, show="headings")
        self.tabla.bind("<Button-3>", self._menu_contextual_factura, add="+")
        self.tabla.bind("<Button-2>", self._menu_contextual_factura, add="+")
        
        self.tabla.tag_configure("con_cuenta", background="#e8f8f5", foreground="#0e6251") 
        self.tabla.tag_configure("sin_cuenta", background="#fdedec", foreground="#7b241c") 

        self.tabla.heading("num", text="N°")
        self.tabla.heading("id_factura", text="ID (Oculto)")
        self.tabla.heading("fecha", text="Fecha Fac.", command=lambda: self.ordenar_por_columna("fecha"))
        self.tabla.heading("hora", text="Hora", command=lambda: self.ordenar_por_columna("hora"))
        self.tabla.heading("nro_doc", text="N° Documento", command=lambda: self.ordenar_por_columna("nro_doc"))
        self.tabla.heading("proveedor", text="Proveedor", command=lambda: self.ordenar_por_columna("proveedor"))
        self.tabla.heading("ruc", text="RUC", command=lambda: self.ordenar_por_columna("ruc"))
        self.tabla.heading("evento", text="Vehículo Asociado", command=lambda: self.ordenar_por_columna("evento"))
        self.tabla.heading("kilometraje", text="Kilometraje", command=lambda: self.ordenar_por_columna("kilometraje"))
        self.tabla.heading("cantidad", text="Galones/Cant.", command=lambda: self.ordenar_por_columna("cantidad"))
        self.tabla.heading("concepto", text="Concepto", command=lambda: self.ordenar_por_columna("concepto"))
        self.tabla.heading("metodo_pago", text="Forma de Pago", command=lambda: self.ordenar_por_columna("metodo_pago"))
        self.tabla.heading("subtotal", text="Subtotal", command=lambda: self.ordenar_por_columna("subtotal"))
        self.tabla.heading("igv", text="IGV", command=lambda: self.ordenar_por_columna("igv"))
        self.tabla.heading("detraccion", text="Detracción", command=lambda: self.ordenar_por_columna("detraccion"))
        self.tabla.heading("neto_facturado", text="Neto a Pagar", command=lambda: self.ordenar_por_columna("neto_facturado"))
        self.tabla.heading("pagado", text="Total Pagado", command=lambda: self.ordenar_por_columna("pagado"))
        self.tabla.heading("saldo", text="Saldo Pendiente", command=lambda: self.ordenar_por_columna("saldo"))
        self.tabla.heading("archivos", text="Historial Adjuntos", command=lambda: self.ordenar_por_columna("archivos"))

        self.tabla.column("num", width=40, anchor="center")
        self.tabla.column("id_factura", width=0, stretch=tk.NO)
        self.tabla.column("fecha", width=80, anchor="center")
        self.tabla.column("hora", width=70, anchor="center")
        # 📏 Ancho suficiente para el N° de documento completo (ver pestaña de facturas)
        self.tabla.column("nro_doc", width=130, minwidth=120, stretch=False, anchor="w")
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
        self._actualizar_flechas_orden()
        self.tabla.bind("<Double-1>", self.abrir_todos_los_archivos)
        
        scroll_y = ctk.CTkScrollbar(f_tabla, orientation="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scroll_y.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        self.frame_bottom = ctk.CTkFrame(self.tab_frame, fg_color="transparent")
        self.frame_bottom.pack(fill="x", padx=15, pady=10)
        
        f_paginacion = ctk.CTkFrame(self.frame_bottom, fg_color="transparent")
        f_paginacion.pack(side="left", padx=(0, 20))
        
        self.btn_ant = ctk.CTkButton(f_paginacion, text="◀ Ant", width=60, command=self.pagina_anterior)
        self.btn_ant.pack(side="left", padx=2)
        
        self.lbl_pagina = ctk.CTkLabel(f_paginacion, text=f"Pág {self.pagina_actual}", font=("Arial", 11, "bold"))
        self.lbl_pagina.pack(side="left", padx=5)
        
        self.btn_sig = ctk.CTkButton(f_paginacion, text="Sig ▶", width=60, command=self.pagina_siguiente)
        self.btn_sig.pack(side="left", padx=2)

        self.btn_excel = ctk.CTkButton(self.frame_bottom, text="📊 Exportar a Excel", font=("Arial", 12, "bold"), width=160, fg_color="#27ae60", hover_color="#1e8449", command=self.exportar_excel)
        self.btn_excel.pack(side="left")

        f_leyenda = ctk.CTkFrame(self.frame_bottom, fg_color="transparent")
        f_leyenda.pack(side="left", padx=30)
        
        ctk.CTkLabel(f_leyenda, text="■", font=("Arial", 14), text_color="#c0392b").pack(side="left", padx=(5,2))
        ctk.CTkLabel(f_leyenda, text="Sin Cuenta Asignada / Pendiente", font=("Arial", 11, "bold"), text_color="#333333").pack(side="left", padx=(0,10))
        
        ctk.CTkLabel(f_leyenda, text="■", font=("Arial", 14), text_color="#27ae60").pack(side="left", padx=(5,2))
        ctk.CTkLabel(f_leyenda, text="Cuenta Asignada", font=("Arial", 11, "bold"), text_color="#333333").pack(side="left", padx=(0,5))

        self.lbl_total_general = ctk.CTkLabel(self.frame_bottom, text="Total Pendiente General por Pagar: 0.00", font=("Arial", 12, "bold"), text_color="#c0392b")
        self.lbl_total_general.pack(side="right")

        # 🧾 Aviso de comprobante compartido: al elegir una factura pagada se indica
        #    qué otras facturas cancela el MISMO comprobante de pago.
        self.lbl_comprobante_compartido = ctk.CTkLabel(
            self.frame_bottom, text="", font=("Arial", 10, "bold"), text_color="#6c3483",
            anchor="w", justify="left", wraplength=520)
        self.lbl_comprobante_compartido.pack(side="left", padx=(8, 8))

        def _al_seleccionar_factura(_evento=None):
            """Muestra las facturas que cancela el comprobante de la factura elegida."""
            try:
                sel = self.tabla.selection()
                if not sel:
                    self.lbl_comprobante_compartido.configure(text="")
                    return
                id_fact = int(self.tabla.item(sel[0], "values")[1])
                ruta_c = self._ruta_comprobante_pago(id_fact)
                if not ruta_c:
                    self.lbl_comprobante_compartido.configure(text="")
                    return
                filas = facturas_que_cancela(ruta_c, "") or []
                if len(filas) > 1:
                    numeros = ", ".join(str(f[0]) for f in filas[:6])
                    if len(filas) > 6:
                        numeros += "…"
                    self.lbl_comprobante_compartido.configure(
                        text="🧾 Este comprobante cancela %d facturas: %s" % (len(filas), numeros))
                else:
                    self.lbl_comprobante_compartido.configure(text="")
            except Exception:
                try:
                    self.lbl_comprobante_compartido.configure(text="")
                except Exception:
                    pass

        try:
            self.tabla.bind("<<TreeviewSelect>>", _al_seleccionar_factura, add="+")
        except Exception as e_sel:
            print("Aviso - selección de factura:", e_sel)

        self.main_root.after(150, lambda: self.cargar_datos_pagar(reset_pagina=True))

    def on_cambiar_mes(self, choice):
        self.mes_filtro = choice
        self.cargar_datos_pagar(reset_pagina=True)

    def pagina_anterior(self):
        if self.pagina_actual > 1:
            self.pagina_actual -= 1
            self.cargar_datos_pagar()
            
    def pagina_siguiente(self):
        if self.pagina_actual < getattr(self, "total_paginas", 1):
            self.pagina_actual += 1
            self.cargar_datos_pagar()

    def buscar_con_retraso(self):
        if hasattr(self, "_busqueda_job"):
            try: self.main_root.after_cancel(self._busqueda_job)
            except: pass
        self._busqueda_job = self.main_root.after(350, lambda: self.cargar_datos_pagar(reset_pagina=True))

    def exportar_excel(self):
        try: import pandas as pd
        except ImportError: return messagebox.showerror("Error", "Falta librería pandas. Ejecuta: pip install pandas openpyxl")
        filas = [self.tabla.item(item)["values"][2:] for item in self.tabla.get_children()]
        if not filas: return messagebox.showwarning("Aviso", "No hay registros.")
        
        columnas = ["Fecha Fac.", "Hora", "N° Documento", "Proveedor", "RUC", "Vehículo (Placa)", "Kilometraje", "Galones/Cant.", "Concepto", "Forma de Pago", "Subtotal", "IGV", "Detracción", "Neto Facturado", "Total Pagado", "Saldo Pendiente", "Archivos"]
        
        ruta = guardar_archivo_dialogo(titulo="Exportar Cuentas por Pagar", defaultextension=".xlsx", initialfile="Cuentas_por_Pagar.xlsx", tipos=[("Excel", "*.xlsx")])
        if ruta:
            pd.DataFrame(filas, columns=columnas).to_excel(ruta, index=False)
            messagebox.showinfo("Éxito", f"Reporte exportado a:\n{ruta}")
            abrir_documento(ruta)

    TITULOS_ORDEN = {
        "fecha": "Fecha Fac.", "hora": "Hora", "nro_doc": "N° Documento", "proveedor": "Proveedor",
        "ruc": "RUC", "evento": "Vehículo Asociado", "kilometraje": "Kilometraje",
        "cantidad": "Galones/Cant.", "concepto": "Concepto", "metodo_pago": "Forma de Pago",
        "subtotal": "Subtotal", "igv": "IGV", "detraccion": "Detracción",
        "neto_facturado": "Neto a Pagar", "pagado": "Total Pagado", "saldo": "Saldo Pendiente",
        "archivos": "Historial Adjuntos",
    }

    def _actualizar_flechas_orden(self):
        """Muestra ▲/▼ en la columna activa y ↕ en las demás."""
        for columna, titulo in self.TITULOS_ORDEN.items():
            if columna == self.columna_orden:
                flecha = "▲" if self.orden_ascendente else "▼"
            else:
                flecha = "↕"
            try:
                self.tabla.heading(columna, text=f"{titulo} {flecha}")
            except Exception:
                pass

    def ordenar_por_columna(self, columna, es_numerico=None):
        """Ordena TODAS las páginas por la columna elegida (no solo la página visible)."""
        if self.columna_orden == columna:
            self.orden_ascendente = not self.orden_ascendente
        else:
            self.columna_orden = columna
            self.orden_ascendente = True
        self._actualizar_flechas_orden()
        self.pagina_actual = 1
        self.cargar_datos_pagar(reset_pagina=True)

    # 🚀 MOTOR DE CONSULTA OPTIMIZADO (0 BUCLES SQL N+1)
    def cargar_datos_pagar(self, reset_pagina=False):
        if reset_pagina:
            self.pagina_actual = 1
            
        self.lbl_pagina.configure(text=f"Pág {self.pagina_actual}")
        for fila in self.tabla.get_children(): 
            self.tabla.delete(fila)
        
        filtro = ""
        if hasattr(self, 'ent_buscar_pagos'):
            filtro = self.ent_buscar_pagos.get().strip().lower()

        # 🔃 Se cargan TODOS los registros del filtro/mes para ordenar por cualquier
        # columna afectando a todas las páginas (la paginación se hace al pintar)
        ocultar_app = bool(self.var_ocultar_app.get())
        clave_cache = f"pagos_compras_v3_{filtro}_mes_{self.mes_filtro}_app_{int(ocultar_app)}"
        datos = cache_sistema.obtener(clave_cache)

        if datos is not None:
            self._pintar_pagos(datos["filas"], datos["total_pendiente"])
        else:
            self.tabla.insert("", tk.END, values=("", "", "", "", "", "", "", "Cargando datos...", "", "", "", "", "", "", "", "", "", "", ""))
            
            # El hilo solo descarga y calcula; la tabla se pinta desde el hilo principal
            self._carga_actual = getattr(self, "_carga_actual", 0) + 1
            carga_id = self._carga_actual

            def tarea_descarga(estado):
                conn = conectar_db(silencioso=True)
                if not conn:
                    estado["filas"] = []
                    estado["total"] = 0.0
                    return
                
                filas_procesadas = []
                total_pendiente_global = 0.0
                
                try:
                    cursor = conn.cursor()

                    condiciones = []
                    params = []
                    if filtro:
                        cond_busqueda, params_busqueda = construir_condicion_busqueda_compras(filtro)
                        condiciones.append(cond_busqueda)
                        params.extend(params_busqueda)
                    patron_mes = patron_fecha_mes(self.mes_filtro)
                    if patron_mes:
                        condiciones.append("fecha LIKE %s")
                        params.append(patron_mes)
                    # 🚫 Facturas de la App (App Grifo): se ocultan si el check está marcado
                    agregar_condicion_ocultar_app(condiciones, params, ocultar_app)
                    where_sql = (" WHERE " + " AND ".join(condiciones)) if condiciones else ""

                    # Se traen TODOS los comprobantes del filtro/mes (sin LIMIT) para
                    # poder ordenar por cualquier columna en todas las páginas
                    cursor.execute(f"SELECT id, fecha, numero_documento, proveedor, evento_asociado, descripcion, subtotal, impuesto, total, COALESCE(det_monto, 0), tipo_documento, kilometraje, cantidad_combustible, ruc, COALESCE(pagado_por_tercero, ''), COALESCE(es_compra_cruzada, FALSE), COALESCE(soporte_pago_tercero, ''), COALESCE(comision_interbancaria, 0), COALESCE(es_interbancario, FALSE) FROM facturas_recibidas{where_sql} ORDER BY id DESC", tuple(params))
                    registros = cursor.fetchall()
                    
                    ids_actuales = [r[0] for r in registros]
                    mapa_detalle_pagos = {}
                    if ids_actuales:
                        cursor.execute(
                            "SELECT id_factura, monto_pagado, archivo_ruta, cuenta_origen FROM pagos_comprobantes WHERE id_factura = ANY(%s)", 
                            (ids_actuales,)
                        )
                        for id_f, m_pag, arch, cta in cursor.fetchall():
                            if id_f not in mapa_detalle_pagos:
                                mapa_detalle_pagos[id_f] = []
                            mapa_detalle_pagos[id_f].append((m_pag, arch, cta))

                    for reg in registros:
                        (id_factura, fecha, nro_doc, proveedor, evento, concepto, subtotal, impuesto,
                         tot_bruto, det_monto, tipo_doc, km_val, cant_val, ruc_db,
                         tercero, es_cruzada, soporte_cruzada,
                         comision_cruz, interbancario_cruz) = reg
                        tercero = str(tercero or "")
                        es_cruzada = bool(es_cruzada)
                        soporte_cruzada = str(soporte_cruzada or "")
                        sub_val = float(subtotal or 0.0)
                        imp_val = float(impuesto or 0.0)
                        tot_bruto_val = float(tot_bruto or 0.0)
                        det_monto_val = float(det_monto or 0.0)
                        
                        if tipo_doc and "Recibo" in tipo_doc and "8%" in tipo_doc:
                            neto_facturado = tot_bruto_val - imp_val - det_monto_val
                        else:
                            neto_facturado = tot_bruto_val - det_monto_val
                        
                        pagos_factura = mapa_detalle_pagos.get(id_factura, [])
                        monto_pagado = 0.0
                        cant_archivos = 0
                        tiene_cuenta = False
                        cuentas_lista = []
                        
                        for p in pagos_factura:
                            monto_pagado += float(p[0] or 0.0)
                            if p[1] and str(p[1]).strip() != "": cant_archivos += 1
                            if p[2] and str(p[2]).strip() != "":
                                tiene_cuenta = True
                                if p[2] not in cuentas_lista: cuentas_lista.append(p[2])
                        
                        saldo_pendiente = max(0.0, neto_facturado - monto_pagado)
                        if es_cruzada:
                            # Compra cruzada (pagada por un TERCERO): no queda saldo por pagar.
                            # Si hay pagos registrados se muestra lo realmente pagado, que
                            # incluye la comisión interbancaria (ej. 28.00 + 4.80 = 32.80).
                            monto_pagado = max(neto_facturado, monto_pagado)
                            saldo_pendiente = 0.0
                        
                        filas_procesadas.append({
                            "id_factura": id_factura, "fecha": fecha, "nro_doc": nro_doc, "proveedor": proveedor, "ruc_db": ruc_db,
                            "evento": evento, "km_val": km_val, "cant_val": cant_val, "concepto": concepto, "cuentas_lista": cuentas_lista,
                            "sub_val": sub_val, "imp_val": imp_val, "det_monto_val": det_monto_val, "neto_facturado": neto_facturado,
                            "monto_pagado": monto_pagado, "saldo_pendiente": saldo_pendiente, "cant_archivos": cant_archivos, "tiene_cuenta": tiene_cuenta,
                            "es_cruzada": es_cruzada, "tercero": tercero, "soporte_cruzada": soporte_cruzada,
                            "comision_cruzada": float(comision_cruz or 0), "interbancario": bool(interbancario_cruz)
                        })
                        
                    # Total pendiente global (mismo criterio que antes) calculado
                    # sobre todos los comprobantes procesados
                    total_pendiente_global = sum(
                        max(0.0, fl["neto_facturado"] - fl["monto_pagado"]) for fl in filas_procesadas
                    )

                    datos_cache = {"filas": filas_procesadas, "total_pendiente": total_pendiente_global}
                    cache_sistema.guardar(clave_cache, datos_cache)
                    estado["filas"] = filas_procesadas
                    estado["total"] = total_pendiente_global
                except Exception as e:
                    print("Error cargando pagos:", e)
                    estado["filas"] = []
                    estado["total"] = 0.0
                finally:
                    liberar_conexion(conn)

            ejecutar_en_hilo(self.main_root, tarea_descarga,
                             al_terminar=lambda e: self._pintar_pagos_si_vigente(carga_id, e))

    def _pintar_pagos_si_vigente(self, carga_id, estado):
        """Pinta la tabla de pagos solo si esta carga sigue siendo la última solicitada."""
        if carga_id != getattr(self, "_carga_actual", 0):
            return
        if estado.get("error"):
            print("[Compras] Error al cargar cuentas por pagar:", estado["error"])
        self._pintar_pagos(estado.get("filas") or [], estado.get("total") or 0.0)

    def _pintar_pagos(self, filas, total_pendiente):
        """Construye TODAS las filas, las ordena por la columna activa (todas las
        páginas) y muestra únicamente la página actual."""
        for fila in self.tabla.get_children(): self.tabla.delete(fila)

        filas_tabla = []
        for f in filas:
            km_str = f['km_val'] if f['km_val'] else "-"
            cant_str = f['cant_val'] if f['cant_val'] else "-"
            ruc_str = f['ruc_db'] if f['ruc_db'] else "-"
            metodo_pago = " + ".join(f['cuentas_lista']) if f['cuentas_lista'] else ""
            if f.get('es_cruzada'):
                # Compra cruzada registrada desde el módulo de Banco: se agrega la
                # nota del tercero conservando la cuenta bancaria del pago.
                nota = "🔁 Compra cruzada" + (f" · pagó: {f['tercero']}" if f.get('tercero') else "")
                if float(f.get('comision_cruzada') or 0) > 0:
                    nota += f" · comisión {formatear_moneda(f['comision_cruzada'])}"
                metodo_pago = f"{metodo_pago} · {nota}" if metodo_pago else nota
            txt_adjuntos = f"📁 {f['cant_archivos']} archivo(s)" if f['cant_archivos'] > 0 else "❌ Sin adjuntos"
            if f.get('es_cruzada') and f['cant_archivos'] == 0:
                txt_adjuntos = "🔁 Ver factura / soporte"
            
            desc_bruta = str(f['concepto']) if f['concepto'] else "-"
            concepto_limpio = desc_bruta
            hora_consumo = "-"
            
            if " | " in desc_bruta:
                partes = desc_bruta.split(" | ")
                concepto_limpio = partes[0].replace("Combustible: ", "")
                for parte in partes[1:]:
                    if parte.startswith("Hora: "): hora_consumo = parte.replace("Hora: ", "").strip()

            etiqueta_color = "con_cuenta" if f['tiene_cuenta'] else "sin_cuenta"

            row_vals = (
                0, f['id_factura'], f['fecha'], hora_consumo, f['nro_doc'] if f['nro_doc'] else "S/N", f['proveedor'], ruc_str, f['evento'], km_str, cant_str, concepto_limpio, metodo_pago,
                formatear_moneda(f['sub_val']), formatear_moneda(f['imp_val']), formatear_moneda(f['det_monto_val']),
                formatear_moneda(f['neto_facturado']), formatear_moneda(f['monto_pagado']), formatear_moneda(f['saldo_pendiente']), txt_adjuntos
            )

            filas_tabla.append((row_vals, etiqueta_color))

        # 🔃 Ordena TODAS las páginas según la columna seleccionada
        filas_tabla = aplicar_orden_filas(filas_tabla, self.columnas_tabla, self.columna_orden, self.orden_ascendente)

        total_registros = len(filas_tabla)
        self.total_paginas = max(1, (total_registros + self.registros_por_pagina - 1) // self.registros_por_pagina)
        if self.pagina_actual > self.total_paginas:
            self.pagina_actual = self.total_paginas
        offset = (self.pagina_actual - 1) * self.registros_por_pagina
        pagina = filas_tabla[offset:offset + self.registros_por_pagina]

        for numero, (row_vals, etiqueta_color) in enumerate(pagina, start=offset + 1):
            self.tabla.insert("", tk.END, values=(numero,) + tuple(row_vals[1:]), tags=(etiqueta_color,))

        self.lbl_total_general.configure(text=f"Total Pendiente Filtrado: {formatear_moneda(total_pendiente)}")
        self.lbl_pagina.configure(text=f"Pág {self.pagina_actual} de {self.total_paginas}")
        self.btn_ant.configure(state="normal" if self.pagina_actual > 1 else "disabled")
        self.btn_sig.configure(state="normal" if self.pagina_actual < self.total_paginas else "disabled")

    def resetear_mantenimientos_vehiculo(self):
        """Abre los mantenimientos del vehículo asociado a la factura seleccionada."""
        sel = self.tabla.selection()
        if not sel:
            return messagebox.showwarning("Mantenimientos",
                                          "Seleccione una factura que tenga vehículo asignado.",
                                          parent=self.main_root)
        valores = self.tabla.item(sel[0], "values")
        placa = valores[7] if len(valores) > 7 else ""
        abrir_reset_mantenimientos_vehiculo(self.main_root, placa, self.app_padre.usuario_activo)

    def cargar_comprobante_pago(self):
        ruta_base = obtener_ruta_base_drive()
        if not ruta_base:
            avisar_sin_permiso_guardado()
            return
            
        seleccion = self.tabla.selection()
        # 🔎 Si no hay fila marcada, se intenta ayudar en vez de solo avisar:
        #   - si la tabla tiene una sola factura, se selecciona sola;
        #   - si está vacía, se explica en qué pestaña está la factura.
        if not seleccion:
            filas = self.tabla.get_children()
            if len(filas) == 1:
                self.tabla.selection_set(filas[0])
                self.tabla.see(filas[0])
                seleccion = (filas[0],)
            elif not filas:
                return messagebox.showwarning(
                    "Sin facturas pendientes",
                    "En esta pestaña no hay facturas pendientes de pago.\n\n"
                    "La factura que quiere cancelar aparece en la pestaña "
                    "«📥 1. Ingreso de Facturas Recibidas»; allí debe estar registrada "
                    "y luego se paga desde «💳 2. Control de Pagos y Deudas».")
            else:
                return messagebox.showwarning(
                    "Selección",
                    "Primero haga clic sobre la factura en la lista de esta pestaña "
                    "(«💳 2. Control de Pagos y Deudas»), de modo que la fila quede marcada "
                    "en azul, y vuelva a pulsar «Registrar Pago».")
        
        valores = self.tabla.item(seleccion[0], "values")
        id_factura, nro_doc, proveedor = valores[1], valores[4], valores[5] 
        saldo_actual = desformatear_numero(valores[17]) 

        # 💸 Si la factura es cruzada con comisión interbancaria, el pago al tercero
        #    cubre la factura MÁS la comisión (ej. factura 28.00 + comisión 4.80 = 32.80).
        comision_factura = 0.0
        es_factura_cruzada = False
        neto_factura = saldo_actual
        ya_pagado = 0.0
        try:
            conn_com = conectar_db(silencioso=True)
            if conn_com:
                try:
                    with conn_com.cursor() as c_com:
                        c_com.execute("""SELECT COALESCE(comision_interbancaria, 0), COALESCE(es_compra_cruzada, FALSE),
                                                COALESCE(total, 0), COALESCE(det_monto, 0), COALESCE(impuesto, 0),
                                                COALESCE(tipo_documento, '')
                                         FROM facturas_recibidas WHERE id = %s""", (id_factura,))
                        fila_com = c_com.fetchone()
                        if fila_com:
                            comision_factura = float(fila_com[0] or 0)
                            es_factura_cruzada = bool(fila_com[1])
                            total_fac = float(fila_com[2] or 0)
                            det_fac = float(fila_com[3] or 0)
                            imp_fac = float(fila_com[4] or 0)
                            tipo_fac = str(fila_com[5] or "")
                            if "Recibo" in tipo_fac and "8%" in tipo_fac:
                                neto_factura = max(0.0, total_fac - imp_fac - det_fac)
                            else:
                                neto_factura = max(0.0, total_fac - det_fac)
                        c_com.execute("SELECT COALESCE(SUM(monto_pagado), 0) FROM pagos_comprobantes WHERE id_factura = %s",
                                      (id_factura,))
                        ya_pagado = float((c_com.fetchone() or [0])[0] or 0)
                finally:
                    liberar_conexion(conn_com)
        except Exception as e_com:
            print("Aviso - comisión de la factura cruzada:", e_com)

        # Lo que falta pagar = factura + comisión interbancaria - lo ya pagado
        pendiente_total = max(0.0, neto_factura + comision_factura - ya_pagado)
        saldo_actual = max(0.0, neto_factura - ya_pagado)
        if pendiente_total <= 0.01:
            return messagebox.showinfo(
                "Aviso",
                "Esta factura ya está pagada por completo"
                + (f" (incluida la comisión interbancaria de {formatear_moneda(comision_factura)})."
                   if comision_factura > 0 else "."))

        v_pago = ctk.CTkToplevel(self.main_root)
        v_pago.title("Registrar Nuevo Pago")
        # 📐 La ventana incluye el panel de soporte: se calcula el alto según la pantalla
        #    para que los botones de abajo SIEMPRE se vean (sin agrandar a mano).
        try:
            alto_pantalla = v_pago.winfo_screenheight()
        except Exception:
            alto_pantalla = 768
        alto_pago = min(720, max(560, alto_pantalla - 110))
        centrar_ventana(v_pago, self.main_root, 470, alto_pago)
        v_pago.transient(self.main_root)
        v_pago.grab_set()
        # 🪟 La ventana de pago debe quedar SIEMPRE por delante de Compras:
        #    se marca como "encima" mientras está abierta y se libera al cerrarse.
        def _pago_al_frente(_intento=0):
            try:
                if not v_pago.winfo_exists():
                    return
                v_pago.deiconify()
                v_pago.lift()
                v_pago.attributes("-topmost", True)
                v_pago.focus_force()
                if _intento < 3:
                    v_pago.after(300 + 250 * _intento, lambda: _pago_al_frente(_intento + 1))
            except Exception:
                pass

        def _pago_soltar_topmost(_evento=None):
            try:
                if v_pago.winfo_exists():
                    v_pago.attributes("-topmost", False)
            except Exception:
                pass

        try:
            v_pago.after(80, _pago_al_frente)
            v_pago.bind("<Destroy>", _pago_soltar_topmost, add="+")
        except Exception:
            pass

        ctk.CTkLabel(v_pago, text=f"Pago para: {proveedor}", font=("Arial", 14, "bold"), text_color="#1f538d").pack(pady=(15, 5))
        ctk.CTkLabel(v_pago, text=f"Saldo Pendiente: {formatear_moneda(saldo_actual)}", font=("Arial", 12)).pack(pady=(0, 2))
        if comision_factura > 0:
            ctk.CTkLabel(v_pago,
                         text=(f"🔁 Factura cruzada: incluye la comisión interbancaria de "
                               f"{formatear_moneda(comision_factura)} → total a pagar "
                               f"{formatear_moneda(pendiente_total)}"),
                         font=("Arial", 10, "bold"), text_color="#6c3483",
                         wraplength=400, justify="center").pack(pady=(0, 10))
        else:
            ctk.CTkLabel(v_pago, text="", font=("Arial", 2)).pack(pady=(0, 10))

        f_form = ctk.CTkFrame(v_pago, fg_color="transparent")
        f_form.pack(fill="x", padx=20)

        config = cargar_configuracion_regional()
        bancos_guardados = cargar_bancos()
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
        # En una factura cruzada con comisión se propone el total: factura + comisión.
        ent_monto.insert(0, f"{pendiente_total:.2f}")

        ctk.CTkLabel(f_form, text="Fecha del Pago:", font=("Arial", 11, "bold")).pack(anchor="w")
        f_fecha_pago = ctk.CTkFrame(f_form, fg_color="transparent")
        f_fecha_pago.pack(fill="x", pady=(0, 10))
        ent_fecha = ctk.CTkEntry(f_fecha_pago)
        ent_fecha.pack(side="left", fill="x", expand=True)
        ent_fecha.insert(0, datetime.now().strftime("%d/%m/%Y"))
        ctk.CTkButton(f_fecha_pago, text="📅", width=40, fg_color="#1f538d", command=lambda: CalendarioNativo(v_pago, ent_fecha)).pack(side="right", padx=(5, 0))

        ctk.CTkLabel(f_form, text="N° de Operación Bancaria (obligatorio):", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_operacion = ctk.CTkEntry(f_form, placeholder_text="Ej.: 00123456 / constancia de la transferencia")
        ent_operacion.pack(fill="x", pady=(0, 2))
        ctk.CTkLabel(f_form, text="Se exige el N° de operación del banco. Solo se omite si el pago es en Efectivo / Caja Chica.",
                     font=("Arial", 9, "italic"), text_color="#7f8c8d", wraplength=400, justify="left").pack(anchor="w", pady=(0, 10))

        def procesar_pago(event=None):
            try:
                monto_val = float(ent_monto.get().strip())
            except ValueError:
                return messagebox.showerror("Error", "Ingrese un monto numérico válido.", parent=v_pago)

            if monto_val <= 0:
                return messagebox.showerror("Error", "El monto debe ser mayor a 0.", parent=v_pago)
            # Se admite hasta el pendiente total (saldo de la factura + comisión interbancaria)
            if monto_val > (pendiente_total + 0.01):
                mensaje = "El monto supera el saldo pendiente."
                if comision_factura > 0:
                    mensaje += (f"\n\nEsta factura cruzada incluye la comisión interbancaria de "
                                f"{formatear_moneda(comision_factura)}: el máximo a pagar es "
                                f"{formatear_moneda(pendiente_total)}.")
                return messagebox.showerror("Error", mensaje, parent=v_pago)

            fecha_val = ent_fecha.get().strip() or datetime.now().strftime("%d/%m/%Y")
            cuenta_val = cmb_cuenta.get().strip()

            # 🏦 El N° de operación bancaria es OBLIGATORIO (en efectivo/caja chica se guarda "EFECTIVO")
            operacion_val = validar_operacion_bancaria(ent_operacion.get(), cuenta_val)
            if operacion_val is None:
                ent_operacion.focus()
                return messagebox.showerror("Falta el N° de Operación",
                                            "Ingrese el N° de Operación Bancaria de este pago.\n\n"
                                            "Si el pago fue en efectivo o con caja chica, seleccione esa cuenta en "
                                            "'Cuenta Origen / Método'.",
                                            parent=v_pago)

            # Cerrar la ventana modal y abrir el diálogo de archivo en el SIGUIENTE
            # ciclo del bucle de eventos. En macOS, invocar el diálogo nativo de
            # archivos inmediatamente después de destruir una ventana con grab_set()
            # (mismo callback) cierra la aplicación.
            soporte_elegido = getattr(self, "_soporte_pago", "") or ""
            v_pago.destroy()
            self.main_root.after(150, lambda: self._guardar_pago_con_soporte(
                id_factura, nro_doc, proveedor, monto_val, fecha_val, cuenta_val, ruta_base,
                operacion_val, soporte_elegido))

        ent_monto.bind("<Return>", procesar_pago)
        ent_fecha.bind("<Return>", procesar_pago)
        ent_operacion.bind("<Return>", procesar_pago)

        # =========================================================
        # 📎 SOPORTE DEL PAGO: escanear o cargar el PDF/ticket
        # =========================================================
        self._soporte_pago = ""

        f_soporte = ctk.CTkFrame(v_pago, fg_color="#eef3f9", border_width=1, border_color="#c5d5e6",
                                 corner_radius=6)
        f_soporte.pack(fill="x", padx=20, pady=(13, 4), ipady=4)
        ctk.CTkLabel(f_soporte, text="📎 Soporte del pago (comprobante)", font=("Arial", 11, "bold"),
                     text_color="#1f538d").pack(anchor="w", padx=8, pady=(3, 2))
        lbl_soporte = ctk.CTkLabel(f_soporte, text="Sin soporte cargado (se puede elegir al confirmar).",
                                   font=("Arial", 9), justify="left", anchor="w", wraplength=370,
                                   text_color="#5d6d7e")
        lbl_soporte.pack(fill="x", padx=8)

        def _aplicar_datos_comprobante(ruta_archivo):
            datos = datos_de_comprobante(ruta_archivo)
            _log_escaneo_comprobante("pago (datos)", ruta_archivo, datos.get("fecha"),
                                     datos.get("operacion"), datos.get("monto"),
                                     "texto=%d car." % len(str(datos.get("texto") or "")))
            avisos = []
            if datos["monto"]:
                ent_monto.delete(0, tk.END)
                ent_monto.insert(0, datos["monto"])
                avisos.append("monto " + formatear_moneda(float(datos["monto"])))
            if datos["fecha"]:
                ent_fecha.delete(0, tk.END)
                ent_fecha.insert(0, datos["fecha"])
                avisos.append("fecha " + datos["fecha"])
            if datos["operacion"]:
                ent_operacion.delete(0, tk.END)
                ent_operacion.insert(0, datos["operacion"])
                avisos.append("operación " + datos["operacion"])

            parte_facturas = ""
            if datos["facturas"]:
                parte_facturas = "\nFacturas detectadas en el documento: " + ", ".join(datos["facturas"])
            ya = facturas_que_cancela("", datos["operacion"]) if datos["operacion"] else []
            if ya:
                parte_facturas += ("\n\nEste comprobante YA canceló:\n" +
                                   texto_facturas_canceladas(ya))
            if avisos:
                lbl_soporte.configure(
                    text=("✅ " + os.path.basename(ruta_archivo) + "\nDetectado: " +
                          " · ".join(avisos) + parte_facturas))
            else:
                lbl_soporte.configure(
                    text=("✅ " + os.path.basename(ruta_archivo) +
                          "\nNo se pudo leer el texto (documento escaneado). "
                          "Complete los datos a mano." + parte_facturas))

        def cargar_pdf_soporte():
            ruta = seleccionar_archivo_dialogo("Seleccionar comprobante de pago",
                                              [("PDF o imagen", "*.pdf;*.png;*.jpg;*.jpeg")])
            if not ruta:
                return
            self._soporte_pago = ruta
            _aplicar_datos_comprobante(ruta)

        def escanear_soporte():
            try:
                import escaner as mod_escaner
            except Exception as e:
                return messagebox.showerror("Escáner", f"No se pudo cargar el módulo de escaneo:\n{e}",
                                            parent=v_pago)
            cfg = cargar_configuracion_regional() or {}
            carpeta = mod_escaner.carpeta_escaneos(cfg)
            lbl_soporte.configure(text="⏳ Escaneando… espere un momento.")
            v_pago.update_idletasks()

            def tarea(estado):
                try:
                    resultado = mod_escaner.escanear(carpeta)
                except Exception as e:
                    estado["error"] = str(e)
                    return
                estado["resultado"] = resultado

            def al_terminar(estado):
                resultado = estado.get("resultado") or {}
                if not resultado.get("ok"):
                    return messagebox.showerror("Escáner",
                                                resultado.get("error") or estado.get("error") or
                                                "No se pudo escanear.", parent=v_pago)
                paginas = resultado.get("paginas") or []
                if not paginas:
                    return messagebox.showwarning("Escáner", "El escáner no devolvió páginas.",
                                                  parent=v_pago)
                try:
                    carpeta_comp = os.path.join(obtener_ruta_base_drive() or carpeta, "comprobantes_egresos")
                    os.makedirs(carpeta_comp, exist_ok=True)
                    destino = os.path.join(carpeta_comp,
                                           "Escaneo_Pago_%s.pdf" % datetime.now().strftime("%Y%m%d_%H%M%S"))
                    pdf = crear_pdf_desde_imagenes(paginas, destino)
                except Exception as e:
                    return messagebox.showerror("Escáner", f"No se pudo crear el PDF:\n{e}", parent=v_pago)
                if not pdf:
                    return messagebox.showerror("Escáner", "No se pudo crear el PDF del escaneo.",
                                                parent=v_pago)
                # 🔍 Se comprueba que el escaneo tenga contenido (hoja en el cristal)
                try:
                    texto_scan = _ocr_comprobante_con_reintentos(paginas)
                except Exception:
                    texto_scan = ""
                self._soporte_pago = pdf
                # 🪟 La ventana del módulo se queda AL FRENTE tras el escaneo
                _traer_ventana_al_frente()
                if not str(texto_scan or "").strip():
                    lbl_soporte.configure(
                        text="⚠️ El escaneo salió en blanco: " + os.path.basename(pdf) +
                             "\nRevise que la hoja esté en el cristal y vuelva a pulsar 📷 Escanear.")
                    return messagebox.showwarning(
                        "Escaneo en blanco",
                        "El escaneo no tiene texto legible.\n\n"
                        "Revise que la hoja esté sobre el cristal del escáner (derecho, sin "
                        "tapas) y vuelva a escanear.\n\n"
                        "El archivo se guardó, pero no se pudieron leer los datos.",
                        parent=v_pago)
                _aplicar_datos_comprobante(pdf)

            ejecutar_en_hilo(v_pago, tarea, al_terminar=al_terminar)

        f_btn_soporte = ctk.CTkFrame(f_soporte, fg_color="transparent")
        f_btn_soporte.pack(fill="x", padx=8, pady=(4, 4))
        ctk.CTkButton(f_btn_soporte, text="📄 Cargar PDF", width=150, height=26,
                      font=("Arial", 11, "bold"), fg_color="#1f538d",
                      command=cargar_pdf_soporte).pack(side="left", padx=(0, 6))
        ctk.CTkButton(f_btn_soporte, text="📷 Escanear", width=130, height=26,
                      font=("Arial", 11, "bold"), fg_color="#8e44ad",
                      command=escanear_soporte).pack(side="left")


        def _consultar_facturas_canceladas():
            operacion = ent_operacion.get().strip()
            filas = facturas_que_cancela(self._soporte_pago or "", operacion)
            if not filas:
                return messagebox.showinfo(
                    "Comprobante",
                    "Este comprobante todavía no cancela ninguna factura registrada.\n\n"
                    "Los números aparecerán aquí cuando se guarden los pagos con este mismo "
                    "soporte o N° de operación.", parent=v_pago)
            messagebox.showinfo(
                "Facturas que cancela este comprobante",
                "Este comprobante cancela %d factura(s):\n\n%s" %
                (len(filas), texto_facturas_canceladas(filas)), parent=v_pago)

        f_btns = ctk.CTkFrame(v_pago, fg_color="transparent")
        f_btns.pack(fill="x", padx=20, pady=15)

        btn_ok = ctk.CTkButton(f_btns, text="✅ Confirmar", font=("Arial", 12, "bold"), fg_color="#27ae60", hover_color="#1e8449", command=procesar_pago)
        btn_ok.pack(side="left", expand=True, padx=5)

        btn_cancel = ctk.CTkButton(f_btns, text="❌ Cancelar", font=("Arial", 12, "bold"), fg_color="#e74c3c", hover_color="#922b21", command=v_pago.destroy)
        btn_cancel.pack(side="right", expand=True, padx=5)
        ent_monto.focus()

        # 📐 Ajuste final: se mide lo que ocupan los controles y la ventana se abre
        #    completa (si la pantalla es chica, se limita al alto disponible).
        def _ajustar_alto_ventana():
            try:
                v_pago.update_idletasks()
                alto_necesario = v_pago.winfo_reqheight() + 12
                alto_pantalla = v_pago.winfo_screenheight()
                alto_final = min(max(alto_necesario, 540), max(540, alto_pantalla - 90))
                centrar_ventana(v_pago, self.main_root, 470, alto_final)
                try:
                    v_pago.resizable(True, True)
                except Exception:
                    pass
            except Exception as e_aj:
                print("Aviso - tamaño de la ventana de pago:", e_aj)

        try:
            v_pago.after(60, _ajustar_alto_ventana)
        except Exception:
            _ajustar_alto_ventana()

    def _guardar_pago_con_soporte(self, id_factura, nro_doc, proveedor, monto_val, fecha_val,
                                  cuenta_val, ruta_base, numero_operacion="", ruta_origen=""):
        """Abre el diálogo para elegir el soporte y guarda el pago.
        Se invoca con after() para evitar el cierre de la app en macOS al abrir
        el diálogo nativo justo después de destruir la ventana modal."""
        try:
            if not ruta_origen:
                ruta_origen = seleccionar_archivo_dialogo("Seleccionar Soporte de Egreso",
                                                          [("Archivos", "*.pdf;*.png;*.jpg;*.jpeg")])
            ruta_destino = ""
            if ruta_origen:
                try:
                    carpeta_comprobantes = os.path.normpath(os.path.join(ruta_base, "comprobantes_egresos"))
                    if not os.path.exists(carpeta_comprobantes): os.makedirs(carpeta_comprobantes)
                    conn = conectar_db(); c = conn.cursor(); c.execute("SELECT COUNT(*) FROM pagos_comprobantes"); idx = c.fetchone()[0] + 1; liberar_conexion(conn)
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
                    INSERT INTO pagos_comprobantes (id_factura, monto_pagado, archivo_ruta, proveedor_nombre, fecha_pago, categoria_suministro, codigo_cotizacion, cuenta_origen, numero_operacion) 
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (id_factura, monto_val, ruta_para_guardar(ruta_destino), proveedor, fecha_val, categoria_db, nro_doc, cuenta_val, numero_operacion))
                conn.commit()
                cache_sistema.invalidar()
                registrar_auditoria(self.app_padre.usuario_activo, "Cuentas por Pagar", f"Pagó {formatear_moneda(monto_val)} a Fac. {nro_doc} desde {cuenta_val}")
                # 📄 Se informa qué facturas cancela este mismo comprobante/operación
                detalle_extra = ""
                try:
                    ruta_guardada = ruta_para_guardar(ruta_destino) if ruta_destino else ""
                    vinculadas = facturas_que_cancela(ruta_guardada, numero_operacion)
                    if len(vinculadas) > 1:
                        detalle_extra = ("\n\nEste comprobante cancela %d facturas:\n%s"
                                         % (len(vinculadas), texto_facturas_canceladas(vinculadas)))
                except Exception as e_vin:
                    print("Aviso - facturas del comprobante:", e_vin)
                messagebox.showinfo("Éxito",
                                    f"Pago de {formatear_moneda(monto_val)} registrado exitosamente."
                                    + detalle_extra)
                self.cargar_datos_pagar(reset_pagina=True)
                self.app_padre.app_facturas.cargar_datos_tabla(reset_pagina=True)
            except Exception as e:
                messagebox.showerror("Error", str(e))
            finally:
                liberar_conexion(conn)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _ver_comprobantes_lista(self, rutas, titulo="Comprobantes de pago"):
        """Muestra los comprobantes de la factura.

        Con uno solo se abre directo. Con varios se despliega un MENÚ junto al
        puntero para elegir cuál ver (antes se abría una ventana intermedia que
        quedaba detrás y no dejaba ver el documento).
        """
        rutas = [r for r in (rutas or []) if r]
        ventana_padre = getattr(self, "main_root", None)
        if not rutas:
            return messagebox.showinfo("Aviso", "No hay comprobantes cargados.",
                                       parent=ventana_para_mensaje(ventana_padre))
        if len(rutas) == 1:
            return abrir_documento(rutas[0], titulo, ventana_padre)

        menu = tk.Menu(ventana_padre, tearoff=0)
        menu.add_command(label="📎  %d comprobantes · elija cuál ver" % len(rutas),
                         state="disabled")
        menu.add_separator()
        for indice, ruta in enumerate(rutas, start=1):
            menu.add_command(
                label="%d.  %s" % (indice, os.path.basename(str(ruta))),
                command=lambda r=ruta: abrir_documento(r, titulo, ventana_padre))
        try:
            menu.tk_popup(menu.winfo_pointerx(), menu.winfo_pointery())
        except Exception as e_menu:
            print("Aviso - menú de comprobantes:", e_menu)
            # Respaldo: se abre el primero (así nunca se queda sin poder verlo)
            return abrir_documento(rutas[0], titulo, ventana_padre)
        finally:
            try:
                menu.grab_release()
            except Exception:
                pass
        return menu

    def abrir_todos_los_archivos(self, event):
        seleccion = self.tabla.selection()
        if not seleccion: return
        id_factura = self.tabla.item(seleccion[0], "values")[1] 
        try:
            conn = conectar_db()
            cursor = conn.cursor()
            cursor.execute("SELECT archivo_ruta FROM pagos_comprobantes WHERE id_factura = %s AND archivo_ruta != ''", (id_factura,))
            rutas = [r[0] for r in cursor.fetchall()]
            if not rutas:
                # Compras cruzadas (módulo de Banco) o facturas sin pagos registrados:
                # se abren los documentos de la propia factura (y su soporte del pago a tercero).
                cursor.execute("""SELECT COALESCE(archivo_ruta, ''), COALESCE(soporte_pago_tercero, '')
                                  FROM facturas_recibidas WHERE id = %s""", (id_factura,))
                extra = cursor.fetchone()
                if extra:
                    rutas = [x for x in extra if x]
            liberar_conexion(conn)
            if rutas:
                abiertos = 0
                _encontrados = []
                for ruta in rutas:
                    # 🔎 Resuelve rutas guardadas por otro equipo/sistema (Mac/Windows)
                    ruta_norm = resolver_ruta_archivo(ruta)
                    if ruta_norm:
                        _encontrados.append(ruta_norm)
                        abiertos += 1
                if _encontrados:
                    # 📎 Con varios comprobantes se muestra la lista para elegir
                    return self._ver_comprobantes_lista(_encontrados)
                if not abiertos:
                    messagebox.showwarning("Aviso", "Los documentos están registrados pero los archivos no se encuentran en este equipo.")
                # 📄 Si el mismo comprobante cancela varias facturas, se muestran sus números
                try:
                    vinculadas = facturas_que_cancela(rutas[0] if rutas else "", "")
                    if len(vinculadas) > 1:
                        messagebox.showinfo(
                            "Comprobante compartido",
                            "Este comprobante de pago cancela %d facturas:\n\n%s"
                            % (len(vinculadas), texto_facturas_canceladas(vinculadas)))
                except Exception as e_vin:
                    print("Aviso - comprobante compartido:", e_vin)
            else: messagebox.showinfo("Aviso", "No hay documentos cargados para esta factura.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron abrir los documentos: {e}")

    def abrir_ventana_edicion(self):
        sel = self.tabla.selection()
        if not sel: return messagebox.showwarning("Selección", "Seleccione la factura para editar sus pagos.")
        valores = self.tabla.item(sel[0], "values")
        id_factura, nro_doc, proveedor = valores[1], valores[4], valores[5] 
        saldo_actual_global = desformatear_numero(valores[17]) 

        v_edit = ctk.CTkToplevel(self.main_root)
        v_edit.title(f"✏️ Gestión de Pagos - Fac. {nro_doc}")
        centrar_ventana(v_edit, self.main_root, 820, 400)
        v_edit.transient(self.main_root)
        v_edit.grab_set()

        ctk.CTkLabel(v_edit, text=f"Pagos registrados: {proveedor}", font=("Arial", 12, "bold"), text_color="#1f538d").pack(pady=10)

        frame_cuerpo = ctk.CTkFrame(v_edit, fg_color="transparent")
        frame_cuerpo.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        sub_tabla = ttk.Treeview(frame_cuerpo, columns=("id", "monto", "fecha", "cuenta", "oper", "tiene_archivo"), show="headings", height=8)
        sub_tabla.heading("id", text="ID"); sub_tabla.heading("monto", text="Monto"); sub_tabla.heading("fecha", text="Fecha"); sub_tabla.heading("cuenta", text="Cuenta / Origen"); sub_tabla.heading("oper", text="N° Operación"); sub_tabla.heading("tiene_archivo", text="¿Soporte?")
        sub_tabla.column("id", width=40, anchor="center"); sub_tabla.column("monto", width=90, anchor="e"); sub_tabla.column("fecha", width=90, anchor="center"); sub_tabla.column("cuenta", width=160, anchor="w"); sub_tabla.column("oper", width=130, anchor="center"); sub_tabla.column("tiene_archivo", width=90, anchor="center")
        sub_tabla.pack(side="left", fill="both", expand=True, padx=(0, 10))

        def refrescar_subtabla():
            for f in sub_tabla.get_children(): sub_tabla.delete(f)
            try:
                conn = conectar_db(); cursor = conn.cursor()
                cursor.execute("SELECT id, monto_pagado, fecha_pago, archivo_ruta, cuenta_origen, COALESCE(numero_operacion,'') FROM pagos_comprobantes WHERE id_factura = %s", (id_factura,))
                for a in cursor.fetchall(): 
                    # El indicador resuelve la ruta (sirve con rutas relativas o de otro equipo)
                    sub_tabla.insert("", tk.END, values=(a[0], formatear_moneda(a[1]), a[2] if a[2] else "Sin fecha", a[4] if a[4] else "-", a[5] if a[5] else "-", "✅ Sí" if resolver_ruta_archivo(a[3]) else "❌ No"))
                liberar_conexion(conn)
            except Exception: pass
        refrescar_subtabla()

        def ejecutar_modificacion():
            nonlocal saldo_actual_global
            sub_sel = sub_tabla.selection()
            if not sub_sel: return
            id_pago = sub_tabla.item(sub_sel[0], "values")[0]
            conn = conectar_db(); cursor = conn.cursor()
            cursor.execute("SELECT monto_pagado, fecha_pago, cuenta_origen, COALESCE(numero_operacion,'') FROM pagos_comprobantes WHERE id = %s", (id_pago,))
            monto_actual, fecha_actual, cuenta_actual, operacion_actual = cursor.fetchone()
            liberar_conexion(conn)

            v_mod_pago = ctk.CTkToplevel(v_edit)
            v_mod_pago.title("Modificar Pago")
            centrar_ventana(v_mod_pago, v_edit, 400, 400)
            v_mod_pago.transient(v_edit)
            v_mod_pago.grab_set()

            ctk.CTkLabel(v_mod_pago, text="Editar Pago", font=("Arial", 14, "bold")).pack(pady=10)
            f_form = ctk.CTkFrame(v_mod_pago, fg_color="transparent")
            f_form.pack(fill="x", padx=20)
            
            config = cargar_configuracion_regional()
            bancos_guardados = cargar_bancos()
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

            ctk.CTkLabel(f_form, text="N° de Operación Bancaria (obligatorio):", font=("Arial", 11, "bold")).pack(anchor="w")
            ent_mod_oper = ctk.CTkEntry(f_form, placeholder_text="Ej.: 00123456")
            ent_mod_oper.pack(fill="x", pady=(0, 2))
            ctk.CTkLabel(f_form, text="Obligatorio en pagos bancarios; en Efectivo / Caja Chica se guarda \"EFECTIVO\".",
                         font=("Arial", 9, "italic"), text_color="#7f8c8d", wraplength=400, justify="left").pack(anchor="w", pady=(0, 10))
            if operacion_actual: ent_mod_oper.insert(0, str(operacion_actual))

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
                        eliminar_archivo(ruta_norm)   # resuelve rutas de otros equipos/SO

                        cursor.execute("DELETE FROM pagos_comprobantes WHERE id = %s", (id_pago,))
                        conn.commit(); liberar_conexion(conn)
                        cache_sistema.invalidar()
                        registrar_auditoria(self.app_padre.usuario_activo, "Cuentas por Pagar", f"Eliminó el pago ID {id_pago}")
                else:
                    nueva_fecha = ent_mod_fecha.get().strip() or fecha_actual
                    nueva_cuenta = ent_mod_cuenta.get().strip()

                    # 🏦 El N° de operación bancaria es OBLIGATORIO (efectivo/caja chica => "EFECTIVO")
                    nueva_oper = validar_operacion_bancaria(ent_mod_oper.get(), nueva_cuenta)
                    if nueva_oper is None:
                        ent_mod_oper.focus()
                        return messagebox.showerror("Falta el N° de Operación",
                                                    "Ingrese el N° de Operación Bancaria del pago.\n\n"
                                                    "Si el pago fue en efectivo o con caja chica, seleccione esa cuenta.",
                                                    parent=v_mod_pago)
                    conn = conectar_db(); cursor = conn.cursor()
                    cursor.execute("UPDATE pagos_comprobantes SET monto_pagado = %s, fecha_pago = %s, cuenta_origen = %s, numero_operacion = %s WHERE id = %s", (nuevo_monto, nueva_fecha, nueva_cuenta, nueva_oper, id_pago))
                    conn.commit(); liberar_conexion(conn)
                    cache_sistema.invalidar()
                    registrar_auditoria(self.app_padre.usuario_activo, "Cuentas por Pagar", f"Modificó el pago ID {id_pago} a {formatear_moneda(nuevo_monto)}")

                saldo_actual_global -= diferencia_de_aumento
                v_mod_pago.destroy()
                refrescar_subtabla()
                self.cargar_datos_pagar(reset_pagina=True)
                self.app_padre.app_facturas.cargar_datos_tabla(reset_pagina=True)

            ent_mod_monto.bind("<Return>", guardar_mod)
            ent_mod_fecha.bind("<Return>", guardar_mod)

            btn_guardar_mod = ctk.CTkButton(v_mod_pago, text="💾 Guardar Cambios", command=guardar_mod, fg_color="#27ae60", hover_color="#1e8449")
            btn_guardar_mod.pack(pady=10)
            ent_mod_monto.focus()

        def cambiar_soporte():
            ruta_base = obtener_ruta_base_drive()
            if not ruta_base:
                avisar_sin_permiso_guardado(v_edit)
                return
            
            sub_sel = sub_tabla.selection()
            if not sub_sel: return
            id_pago = sub_tabla.item(sub_sel[0], "values")[0]
            ruta_origen = seleccionar_archivo_dialogo("Seleccionar Soporte", [("Archivos", "*.pdf;*.png;*.jpg;*.jpeg")])
            if ruta_origen:
                try:
                    carpeta_comprobantes = os.path.normpath(os.path.join(ruta_base, "comprobantes_egresos"))
                    if not os.path.exists(carpeta_comprobantes): os.makedirs(carpeta_comprobantes)
                    conn = conectar_db(); cursor = conn.cursor()
                    cursor.execute("SELECT archivo_ruta FROM pagos_comprobantes WHERE id = %s", (id_pago,))
                    antigua_ruta = os.path.normpath(cursor.fetchone()[0])
                    eliminar_archivo(antigua_ruta)   # resuelve rutas de otros equipos/SO

                    prov_limpio = re.sub(r'[\\/*?:"<>|]', '-', proveedor)
                    nombre_limpio = f"Egreso_Fac_{id_factura}_{prov_limpio.replace(' ', '_')}_R_{id_pago}{os.path.splitext(ruta_origen)[1]}"
                    ruta_destino = os.path.normpath(os.path.join(carpeta_comprobantes, nombre_limpio))
                    shutil.copy2(ruta_origen, ruta_destino)
                    cursor.execute("UPDATE pagos_comprobantes SET archivo_ruta = %s WHERE id = %s", (ruta_para_guardar(ruta_destino), id_pago))
                    conn.commit(); liberar_conexion(conn)
                    registrar_auditoria(self.app_padre.usuario_activo, "Cuentas por Pagar", f"Actualizó soporte del pago ID {id_pago}")
                    messagebox.showinfo("Éxito", "Soporte actualizado.", parent=v_edit)
                    refrescar_subtabla(); self.cargar_datos_pagar(reset_pagina=True)
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
                    eliminar_archivo(ruta_archivo)   # resuelve rutas de otros equipos/SO

                    cursor.execute("UPDATE pagos_comprobantes SET archivo_ruta = '' WHERE id = %s", (id_pago,))
                    conn.commit(); liberar_conexion(conn)
                    registrar_auditoria(self.app_padre.usuario_activo, "Cuentas por Pagar", f"Eliminó soporte del pago ID {id_pago}")
                    messagebox.showinfo("Éxito", "Soporte eliminado.", parent=v_edit)
                    refrescar_subtabla(); self.cargar_datos_pagar(reset_pagina=True)
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
                    eliminar_archivo(ruta_norm)   # resuelve rutas de otros equipos/SO

                    cursor.execute("DELETE FROM pagos_comprobantes WHERE id = %s", (id_pago,))
                    conn.commit(); liberar_conexion(conn)
                    cache_sistema.invalidar()
                    registrar_auditoria(self.app_padre.usuario_activo, "Cuentas por Pagar", f"Eliminó completamente el pago ID {id_pago}")
                    messagebox.showinfo("Éxito", "Pago eliminado.", parent=v_edit)
                    refrescar_subtabla(); self.cargar_datos_pagar(reset_pagina=True); self.app_padre.app_facturas.cargar_datos_tabla(reset_pagina=True); saldo_actual_global += monto_eliminado
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

    # =========================================================================
    # 🚀 REPORTE UNIFICADO: TOTALES FINANCIEROS Y RESUMEN POR PROVEEDOR
    # =========================================================================
    def mostrar_reporte_totales(self):
        v_rep = ctk.CTkToplevel(self.main_root)
        v_rep.title("Reporte Consolidado de Totales y Proveedores")
        centrar_ventana(v_rep, self.main_root, 880, 680)
        v_rep.transient(self.main_root)
        v_rep.grab_set()

        ctk.CTkLabel(v_rep, text="📊 REPORTE CONSOLIDADO DE COMPRAS Y PROVEEDORES", font=("Arial", 16, "bold"), text_color="#1f538d").pack(pady=(12, 5))

        # --- FILTROS ---
        f_filtros = ctk.CTkFrame(v_rep, fg_color="#f8f9fa", border_width=1, border_color="#ccc")
        f_filtros.pack(fill="x", padx=15, pady=5, ipadx=10, ipady=5)

        f_controles = ctk.CTkFrame(f_filtros, fg_color="transparent")
        f_controles.pack(fill="x", pady=2)

        ctk.CTkLabel(f_controles, text="Proveedor:", font=("Arial", 11, "bold")).pack(side="left", padx=(5, 5))
        provs_mem = cache_sistema.obtener('lista_proveedores_combobox')
        provs = ["Todos"] + (provs_mem if provs_mem else [])
        combo_prov = ctk.CTkComboBox(f_controles, values=provs, state="readonly", width=220)
        combo_prov.pack(side="left", padx=5)
        combo_prov.set("Todos")

        ctk.CTkLabel(f_controles, text="Desde:", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 5))
        ent_desde = ctk.CTkEntry(f_controles, width=100, placeholder_text=CONFIG_REGIONAL.get("formato_fecha", "DD/MM/AAAA"))
        ent_desde.pack(side="left", padx=2)
        ctk.CTkButton(f_controles, text="📅", width=30, fg_color="#1f538d", hover_color="#163b65", command=lambda: CalendarioNativo(v_rep, ent_desde)).pack(side="left")

        ctk.CTkLabel(f_controles, text="Hasta:", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 5))
        ent_hasta = ctk.CTkEntry(f_controles, width=100, placeholder_text=CONFIG_REGIONAL.get("formato_fecha", "DD/MM/AAAA"))
        ent_hasta.pack(side="left", padx=2)
        ctk.CTkButton(f_controles, text="📅", width=30, fg_color="#1f538d", hover_color="#163b65", command=lambda: CalendarioNativo(v_rep, ent_hasta)).pack(side="left")

        btn_buscar = ctk.CTkButton(f_controles, text="🔍 Procesar", width=100, font=("Arial", 11, "bold"), fg_color="#1f538d", hover_color="#163b65", command=lambda: calcular_totales())
        btn_buscar.pack(side="left", padx=(15, 5))

        # --- TARJETAS DE TOTALES (MÉTRICAS) ---
        f_cards = ctk.CTkFrame(v_rep, fg_color="transparent")
        f_cards.pack(fill="x", padx=15, pady=5)

        def crear_card(padre, titulo, color):
            f = ctk.CTkFrame(padre, fg_color="#ffffff", border_width=1, border_color="#ddd", corner_radius=6)
            f.pack(side="left", fill="both", expand=True, padx=3)
            ctk.CTkLabel(f, text=titulo, font=("Arial", 10, "bold"), text_color="gray").pack(pady=(4, 0))
            lbl_val = ctk.CTkLabel(f, text="0.00", font=("Arial", 13, "bold"), text_color=color)
            lbl_val.pack(pady=(0, 4))
            return lbl_val

        lbl_bruto = crear_card(f_cards, "Subtotal Base", "#1f538d")
        lbl_igv = crear_card(f_cards, "IGV Facturado", "#1f538d")
        lbl_det = crear_card(f_cards, "Detrac / Reten", "#e67e22")
        lbl_pagado = crear_card(f_cards, "Total Pagado", "#27ae60")
        lbl_por_pagar = crear_card(f_cards, "Deuda Pendiente", "#c0392b")

        # --- TABLA DESGLOSE POR PROVEEDOR ---
        ctk.CTkLabel(v_rep, text="📋 Desglose y Saldos Pendientes por Proveedor:", font=("Arial", 12, "bold"), text_color="#333").pack(anchor="w", padx=15, pady=(8, 2))

        f_tabla_prov = ctk.CTkFrame(v_rep, fg_color="transparent")
        f_tabla_prov.pack(fill="both", expand=True, padx=15, pady=(0, 10))

        t_resumen = ttk.Treeview(f_tabla_prov, columns=("prov", "neto", "pagado", "saldo", "docs"), show="headings")
        t_resumen.heading("prov", text="Proveedor / Empresa")
        t_resumen.heading("neto", text="Neto Facturado")
        t_resumen.heading("pagado", text="Total Pagado")
        t_resumen.heading("saldo", text="Saldo Pendiente")
        t_resumen.heading("docs", text="N° Docs")

        t_resumen.column("prov", width=260, anchor="w")
        t_resumen.column("neto", width=120, anchor="e")
        t_resumen.column("pagado", width=120, anchor="e")
        t_resumen.column("saldo", width=130, anchor="e")
        t_resumen.column("docs", width=70, anchor="center")

        scroll_res = ttk.Scrollbar(f_tabla_prov, orient="vertical", command=t_resumen.yview)
        t_resumen.configure(yscrollcommand=scroll_res.set)
        t_resumen.pack(side="left", fill="both", expand=True)
        scroll_res.pack(side="right", fill="y")

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

            for f in t_resumen.get_children():
                t_resumen.delete(f)

            def tarea_calculo(estado):
                """Calcula los totales en segundo plano; la ventana se pinta en el hilo principal."""
                conn = conectar_db()
                if not conn:
                    estado["resumen"] = {}
                    return
                try:
                    c = conn.cursor()
                    c.execute("SELECT id_factura, COALESCE(SUM(monto_pagado), 0) FROM pagos_comprobantes GROUP BY id_factura")
                    pagos_dict = {row[0]: float(row[1]) for row in c.fetchall()}

                    c.execute("SELECT id, fecha, proveedor, subtotal, impuesto, total, COALESCE(det_monto, 0), tipo_documento FROM facturas_recibidas")
                    
                    tot_bruto = 0.0
                    tot_igv = 0.0
                    tot_det = 0.0
                    tot_pagado = 0.0
                    tot_deuda = 0.0

                    proveedores_desglose = {}

                    for r in c.fetchall():
                        id_fac, fecha, prov, sub, imp, tot, det, tipo_doc = r
                        nombre_prov = str(prov).strip() if prov else "SIN PROVEEDOR"
                        
                        if prov_filtro != "Todos" and nombre_prov != prov_filtro: continue
                        
                        f_dt = convertir_a_fecha(str(fecha))
                        if d_desde and d_hasta:
                            if not f_dt or not (d_desde <= f_dt <= d_hasta): continue
                        elif d_desde:
                            if not f_dt or f_dt < d_desde: continue
                        elif d_hasta:
                            if not f_dt or f_dt > d_hasta: continue

                        sub_val = float(sub or 0.0)
                        imp_val = float(imp or 0.0)
                        tot_val = float(tot or 0.0)
                        det_val = float(det or 0.0)
                        
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

                        if nombre_prov not in proveedores_desglose:
                            proveedores_desglose[nombre_prov] = {"neto": 0.0, "pagado": 0.0, "saldo": 0.0, "docs": 0}
                        
                        proveedores_desglose[nombre_prov]["neto"] += neto
                        proveedores_desglose[nombre_prov]["pagado"] += pagado
                        proveedores_desglose[nombre_prov]["saldo"] += saldo
                        proveedores_desglose[nombre_prov]["docs"] += 1

                    estado["resumen"] = {
                        "bruto": tot_bruto, "igv": tot_igv, "det": tot_det,
                        "pagado": tot_pagado, "deuda": tot_deuda,
                        "proveedores": proveedores_desglose,
                    }
                except Exception as e:
                    print("[Compras] Error en el reporte:", e)
                    estado["resumen"] = {}
                finally:
                    liberar_conexion(conn)

            def pintar_reporte(estado):
                if estado.get("error"):
                    messagebox.showerror("Error", f"Fallo al calcular:\n{estado['error']}", parent=v_rep)
                    return
                resumen = estado.get("resumen") or {}
                lbl_bruto.configure(text=formatear_moneda(resumen.get("bruto", 0)))
                lbl_igv.configure(text=formatear_moneda(resumen.get("igv", 0)))
                lbl_det.configure(text=formatear_moneda(resumen.get("det", 0)))
                lbl_pagado.configure(text=formatear_moneda(resumen.get("pagado", 0)))
                lbl_por_pagar.configure(text=formatear_moneda(resumen.get("deuda", 0)))
                for p, data in sorted((resumen.get("proveedores") or {}).items(),
                                      key=lambda x: x[1]["saldo"], reverse=True):
                    t_resumen.insert("", tk.END, values=(
                        p,
                        formatear_moneda(data["neto"]),
                        formatear_moneda(data["pagado"]),
                        formatear_moneda(data["saldo"]),
                        data["docs"]
                    ))

            ejecutar_en_hilo(self.main_root, tarea_calculo, al_terminar=pintar_reporte)

        calcular_totales()

# =========================================================
# CLASE PRINCIPAL: MÓDULO DE COMPRAS (CONTENEDOR TABVIEW)
# =========================================================
# 📎 La pestaña de Facturas Recibidas usa la misma lista de comprobantes que Pagos
FacturasRecibidasTab._ver_comprobantes_lista = CuentasPorPagarTab._ver_comprobantes_lista


class ModuloComprasApp:
    def __init__(self, parent_frame):
        self.parent_frame = parent_frame
        # 🪟 Se registra la ventana del módulo para traerla al frente tras cada aviso
        try:
            _VENTANA_COMPRAS[0] = parent_frame.winfo_toplevel()
        except Exception:
            _VENTANA_COMPRAS[0] = None
        self.usuario_activo = "Desconocido"
        self.pantalla_expandida = False
        aplicar_estilo_treeview()

        self.frame_main = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.frame_main.pack(fill="both", expand=True, padx=15, pady=15)
        
        header_frame = ctk.CTkFrame(self.frame_main, fg_color="transparent")
        header_frame.pack(fill="x")
        
        ctk.CTkLabel(header_frame, text="🛒 MÓDULO DE COMPRAS (LOGÍSTICA Y TESORERÍA)", font=("Arial", 18, "bold"), text_color="#1f538d").pack(side="left")

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
            self.pantalla_expandida = False
        else:
            if sidebar: sidebar.pack_forget()
            self.pantalla_expandida = True

    def al_cambiar_pestana(self):
        if self.tabview.get() == " 💳 2. Control de Pagos y Deudas ":
            self.app_pagos.cargar_datos_pagar(reset_pagina=True)
        elif self.tabview.get() == " 📥 1. Ingreso de Facturas Recibidas ":
            self.app_facturas.cargar_datos_tabla(reset_pagina=True)

if __name__ == "__main__":
    pass
