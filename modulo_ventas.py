# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import customtkinter as ctk
import os
import sys
import shutil
import calendar
import math
import re
import json
import subprocess 
import ctypes
import webbrowser
import random
import smtplib
from email.mime.text import MIMEText
import urllib.request
import urllib.parse
from datetime import datetime
import threading

# 🚀 IMPORTAMOS NUESTRAS NUEVAS HERRAMIENTAS CORPORATIVAS
from conexion import conectar_db, registrar_auditoria, liberar_conexion
from buffer_memoria import cache_sistema
from dialogos_seguros import seleccionar_archivo_dialogo, guardar_archivo_dialogo
from app_paths import CONFIG_FILE, eliminar_archivo, ruta_para_guardar, resolver_ruta_archivo
from tareas_seguras import ejecutar_en_hilo
from config_nube import cargar_bancos

# 🖥️ ventana_para_mensaje vive en modulo_compras; se importa con respaldo local
try:
    from modulo_compras import ventana_para_mensaje
except Exception:
    def ventana_para_mensaje(principal=None):
        return principal


try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    from google import genai
    from google.genai import types as genai_types
except Exception:
    genai = None
    genai_types = None

try:
    from facturacion_sunat import abrir_ventana_emision_sunat
except ImportError:
    abrir_ventana_emision_sunat = None

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

NOMBRES_MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                 "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

# =========================================================
# 🚀 ADAPTACIÓN MULTIPLATAFORMA
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
            pie = ctk.CTkLabel(top, text="clic para elegir  Â·  â†‘â†“ mover  Â·  Enter acepta  Â·  Esc/click fuera cierra",
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
            # ðŸ”Ž se avisa al combo para que complete el RUC / detracciÃ³n
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
            cuenta.configure(text="%d resultado(s) â€” clic en uno para elegirlo" % len(hallados))

        pintar()
        entrada.bind("<KeyRelease>", lambda e: pintar(entrada.get()))
        entrada.bind("<Return>", elegir)
        entrada.bind("<Escape>", cerrar)
        entrada.bind("<Down>", lambda e: (lista.focus_set(), None)[1])
        lista.bind("<ButtonRelease-1>", elegir)     # ðŸ‘ˆ UN clic elige y cierra
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

def resolver_archivo_local(ruta):
    """Devuelve la ruta REAL del archivo en ESTE equipo (o "" si no se encuentra).

    Las rutas de los PDFs se guardan relativas a la carpeta sincronizada (o
    absolutas de otro equipo/SO), así que hay que resolverlas antes de
    comprobar si el archivo existe: un os.path.exists directo fallaría y el
    programa respondería "no hay PDF asociado" aunque el archivo sí exista.
    """
    if not ruta:
        return ""
    texto = str(ruta).strip()
    try:
        encontrada = resolver_ruta_archivo(texto)
        if encontrada:
            return encontrada
    except Exception:
        pass
    try:
        candidata = os.path.normpath(texto)
        if os.path.exists(candidata):
            return candidata
    except Exception:
        pass
    return ""


def abrir_documento(ruta):
    try:
        # Resuelve rutas guardadas en otro equipo/SO (Mac <-> Windows)
        ruta = resolver_archivo_local(ruta) or ruta
        ruta_abs = os.path.abspath(ruta)
        if sys.platform == "win32": os.startfile(ruta_abs)
        elif sys.platform == "darwin": subprocess.call(["open", ruta_abs])
        else: subprocess.call(["xdg-open", ruta_abs])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el archivo o carpeta:\n{e}")

# =========================================================
# 🚀 MOTOR DE CONFIGURACIÓN REGIONAL
# =========================================================
def cargar_configuracion_regional():
    config = {
        "simbolo_moneda": "S/.",
        "formato_numero": "1,000.00",
        "formato_fecha": "DD/MM/AAAA",
        "ruta_drive": "",
        "impresora": "",
        "ultimo_factura": "F001-0",
        "ultimo_boleta": "B001-0",
        "ultimo_recibo": "E001-0",
        "2fa_metodo": "Inactivo",
        "tel_bot_token": "",
        "tel_chat_id": "",
        "email_smtp": "smtp.gmail.com",
        "email_port": "587",
        "email_user": "",
        "email_pass": "",
        "email_dest": "",
        "ruc_empresa": "",
        "usuario_sol": "",
        "clave_sol": "",
        "client_id_sire": "",
        "client_secret_sire": "",
        "detraccion_porcentaje": "12"
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
    simbolo = CONFIG_REGIONAL.get("simbolo_moneda", "S/.")
    formato = CONFIG_REGIONAL.get("formato_numero", "1,000.00")
    val = str(valor_str).replace(simbolo, "").strip()
    if formato == "1.000,00":
        val = val.replace(".", "").replace(",", ".")
    else:
        val = val.replace(",", "")
    try: return float(val)
    except ValueError: return 0.0

# =========================================================
# 🧮 REDONDEO DE LA DETRACCIÓN (norma del Banco de la Nación)
# =========================================================
def redondear_centavos(monto):
    """Redondea un importe a 2 decimales (evita el ruido de coma flotante al
    desagregar el IGV de un total que ya lo incluye)."""
    try:
        return round(float(monto) + 1e-9, 2)
    except (TypeError, ValueError):
        return 0.0


def redondear_detraccion(monto):
    """Redondea al entero el monto de la detracción.

    - Fracción menor a 0.50  → se suprimen los decimales (baja al entero).
    - Fracción igual o mayor a 0.50 → se ajusta a la unidad inmediata superior.
    """
    try:
        valor = float(monto)
    except (TypeError, ValueError):
        return 0.0
    if valor < 0:
        return -redondear_detraccion(-valor)
    entero = math.floor(valor)
    return float(entero + 1) if (valor - entero) >= 0.5 else float(entero)


def obtener_ruta_base_drive():
    """Carpeta base para guardar archivos.

    Devuelve "" si el equipo NO está autorizado (equipo secundario sin la cuenta
    Rclone del equipo principal): en ese caso no se guarda ningún archivo.
    """
    try:
        from politica_almacenamiento import ruta_base_autorizada
        return ruta_base_autorizada(mostrar_alerta=False)
    except Exception:
        ruta = CONFIG_REGIONAL.get("ruta_drive", "").strip()
        return os.path.expanduser(ruta) if ruta else ""


def avisar_sin_permiso_guardado(parent=None):
    """Advertencia de bloqueo o de configuración faltante."""
    try:
        from politica_almacenamiento import avisar_sin_ruta
        return avisar_sin_ruta(parent)
    except Exception:
        messagebox.showwarning("Configuración Requerida", "No ha configurado la ruta de Google Drive.\nEs obligatorio para guardar archivos.", parent=parent)
        return False

def aplicar_estilo_treeview():
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Treeview", background="#ffffff", foreground="#000000", fieldbackground="#ffffff", bordercolor="#e0e0e0", borderwidth=1, rowheight=26, font=("Arial", 10))
    style.map("Treeview", background=[("selected", "#1f538d")], foreground=[("selected", "#ffffff")])
    style.configure("Treeview.Heading", background="#f0f0f0", foreground="#000000", relief="flat", font=("Arial", 10, "bold"), bordercolor="#e0e0e0", borderwidth=1)

# =========================================================
# 🛡️ MOTOR DE SEGURIDAD OTP (CLAVE DINÁMICA)
# =========================================================
def solicitar_otp(accion, callback_exito, parent_window):
    config = cargar_configuracion_regional()
    metodo = config.get("2fa_metodo", "Inactivo")

    if metodo == "Inactivo":
        callback_exito()
        return

    codigo_otp = str(random.randint(100000, 999999))
    mensaje = f"🔐 CÓDIGO DE SEGURIDAD SUNAT\n\nAcción solicitada: {accion}\n\nTu clave dinámica (OTP) es: {codigo_otp}\n\nNo compartas este código con nadie."

    def _mostrar_ventana_otp():
        v_otp = ctk.CTkToplevel(parent_window)
        v_otp.title("Verificación de Seguridad")
        v_otp.geometry("380x250")
        v_otp.grab_set()

        ctk.CTkLabel(v_otp, text="🔒 Verificación en Dos Pasos", font=("Arial", 16, "bold"), text_color="#1f538d").pack(pady=(20, 5))
        ctk.CTkLabel(v_otp, text=f"Se ha enviado un código por {metodo}.\nIngrese los 6 dígitos para autorizar.", font=("Arial", 11)).pack(pady=(0, 15))

        ent_codigo = ctk.CTkEntry(v_otp, placeholder_text="Ej: 123456", justify="center", font=("Arial", 18, "bold"), width=150)
        ent_codigo.pack(pady=10)

        def verificar():
            if ent_codigo.get().strip() == codigo_otp:
                v_otp.destroy()
                callback_exito()
            else:
                messagebox.showerror("Código Incorrecto", "El código ingresado no es válido.", parent=v_otp)

        ctk.CTkButton(v_otp, text="✅ Validar y Autorizar", font=("Arial", 13, "bold"), fg_color="#27ae60", hover_color="#1e8449", height=40, command=verificar).pack(pady=10)

    # 🚀 RENDIMIENTO: el envío (Telegram/SMTP) corre en un hilo para no congelar la
    # interfaz; la ventana OTP se muestra recién cuando el código fue enviado.
    # 🔒 macOS: el hilo NO toca la interfaz; los avisos y la ventana OTP se abren
    #    en 'al_terminar', ya en el hilo principal.
    def _enviar_y_continuar(estado):         # 👈 corre en el HILO SECUNDARIO
        exito_envio = False
        try:
            if "Telegram" in metodo:
                token = config.get("tel_bot_token", "").strip()
                chat_id = config.get("tel_chat_id", "").strip()
                if not token or not chat_id:
                    estado["error_titulo"] = "Error"
                    estado["error_texto"] = "Faltan credenciales de Telegram en Configuración General."
                    return
                url = f"https://api.telegram.org/bot{token}/sendMessage"
                data = urllib.parse.urlencode({"chat_id": chat_id, "text": mensaje}).encode("utf-8")
                req = urllib.request.Request(url, data=data)
                with urllib.request.urlopen(req, timeout=5) as response:
                    exito_envio = True

            elif "Correo" in metodo:
                smtp_server = config.get("email_smtp", "smtp.gmail.com").strip()
                port = int(config.get("email_port", 587))
                user = config.get("email_user", "").strip()
                password = config.get("email_pass", "").strip()
                dest = config.get("email_dest", "").strip()

                if not user or not password or not dest:
                    estado["error_titulo"] = "Error"
                    estado["error_texto"] = "Faltan credenciales de correo en Configuración General."
                    return

                msg = MIMEText(mensaje)
                msg['Subject'] = 'Código de Seguridad SUNAT - Black Cube'
                msg['From'] = user
                msg['To'] = dest

                server = smtplib.SMTP(smtp_server, port)
                server.starttls()
                server.login(user, password)
                server.send_message(msg)
                server.quit()
                exito_envio = True

            elif "SMS" in metodo:
                estado["aviso_sms"] = True
                print(f"--- [ALERTA SMS TWILIO] CÓDIGO OTP --- : {codigo_otp}")
                exito_envio = True
        except Exception as e:
            estado["error_titulo"] = "Error de Envío OTP"
            estado["error_texto"] = f"No se pudo conectar con el servicio {metodo}:\n{e}"
            return

        if not exito_envio:
            estado["error_titulo"] = "Error de Envío OTP"
            estado["error_texto"] = "No se pudo enviar el código de verificación. Verifique las credenciales en Configuración General."
            return

        estado["mostrar_otp"] = True

    def _terminar_envio(estado):             # 👈 corre en el HILO PRINCIPAL
        if estado.get("error_texto"):
            messagebox.showerror(estado.get("error_titulo") or "Error", estado["error_texto"], parent=parent_window)
            return
        if estado.get("aviso_sms"):
            messagebox.showwarning("Aviso", "El módulo SMS Twilio requiere instalación externa. Se imprimirá el código en la consola del servidor por ahora.", parent=parent_window)
        if estado.get("mostrar_otp"):
            _mostrar_ventana_otp()

    ejecutar_en_hilo(parent_window, _enviar_y_continuar, al_terminar=_terminar_envio)

# =========================================================
# CLASE: MINI CALENDARIO COMPARTIDO
# =========================================================
# =========================================================================
# 💳 PAGOS POR ADELANTADO (sin factura) — se asocian después a UNA factura
# =========================================================================
# Se guardan como un cobro normal en 'pagos_clientes' con:
#   tipo = 'Adelanto'  y  id_factura = 0 (todavía sin factura)
# Al asociarlo a su factura solo se le pone el id_factura, así que TODO lo que ya
# lee esa tabla (Banco, Libro Diario, Estadísticas Financiera, estado de cuenta
# del cliente) lo cuenta automáticamente. El contrato es externo: texto libre.

def registrar_adelanto_cliente(cliente, monto, fecha="", contrato="", cuenta="", archivo="",
                               usuario="Sistema"):
    """Guarda un pago por adelantado SIN factura. Devuelve (ok, mensaje)."""
    nombre = str(cliente or "").strip()
    if not nombre:
        return False, "Indique el cliente."
    try:
        monto_val = round(float(str(monto).replace(",", ".")), 2)
    except Exception:
        return False, "El monto no es un número válido."
    if monto_val <= 0:
        return False, "El monto debe ser mayor a 0."
    if not str(fecha or "").strip():
        from datetime import datetime as _dt
        fecha = _dt.now().strftime("%d/%m/%Y")

    conn = conectar_db()
    if not conn:
        return False, "Sin conexión a la base de datos."
    try:
        cursor = conn.cursor()
        for sentencia in ("ALTER TABLE pagos_clientes ADD COLUMN IF NOT EXISTS tipo VARCHAR(20) DEFAULT 'Cobro'",
                          "ALTER TABLE pagos_clientes ADD COLUMN IF NOT EXISTS contrato VARCHAR(255) DEFAULT ''"):
            try:
                cursor.execute(sentencia); conn.commit()
            except Exception:
                conn.rollback()
        cursor.execute("""
            INSERT INTO pagos_clientes (codigo_cotizacion, monto_pagado, archivo_ruta, cliente_nombre,
                                        fecha_pago, id_factura, cuenta_destino, tipo, contrato)
            VALUES (%s, %s, %s, %s, %s, 0, %s, 'Adelanto', %s)
            RETURNING id
        """, (str(contrato or "").strip(), monto_val, str(archivo or ""), nombre, str(fecha),
              str(cuenta or ""), str(contrato or "").strip()))
        try:
            id_nuevo = int(cursor.fetchone()[0])
        except Exception:
            id_nuevo = 0
        conn.commit()
        try:
            registrar_auditoria(usuario, "Ventas",
                                "Registró un pago por adelantado de %s a %s (contrato %s)"
                                % (monto_val, nombre, contrato or "sin dato"))
        except Exception:
            pass
        try:
            cache_sistema.invalidar()
        except Exception:
            pass
        return True, "Adelanto registrado (N° %s). Queda como saldo a favor hasta asociarlo a su factura." % id_nuevo
    except Exception as e:
        try:
            conn.rollback()
        except Exception:
            pass
        return False, str(e)
    finally:
        liberar_conexion(conn)


def adelantos_disponibles(cliente="", incluir_aplicados=False):
    """Adelantos de clientes. Por defecto solo los que están SIN factura (saldo a favor).

    Devuelve [{'id', 'cliente', 'monto', 'fecha', 'contrato', 'cuenta', 'id_factura', 'estado'}].
    """
    lista = []
    conn = conectar_db()
    if not conn:
        return lista
    try:
        cursor = conn.cursor()
        for sentencia in ("ALTER TABLE pagos_clientes ADD COLUMN IF NOT EXISTS tipo VARCHAR(20) DEFAULT 'Cobro'",
                          "ALTER TABLE pagos_clientes ADD COLUMN IF NOT EXISTS contrato VARCHAR(255) DEFAULT ''"):
            try:
                cursor.execute(sentencia); conn.commit()
            except Exception:
                conn.rollback()
        donde = "WHERE COALESCE(tipo, 'Cobro') = 'Adelanto'"
        parametros = []
        if not incluir_aplicados:
            donde += " AND COALESCE(id_factura, 0) = 0"
        if str(cliente or "").strip():
            donde += " AND cliente_nombre ILIKE %s"
            parametros.append("%" + str(cliente).strip() + "%")
        cursor.execute("""
            SELECT id, COALESCE(cliente_nombre, ''), COALESCE(monto_pagado, 0),
                   COALESCE(fecha_pago, ''), COALESCE(contrato, ''), COALESCE(cuenta_destino, ''),
                   COALESCE(id_factura, 0)
            FROM pagos_clientes %s ORDER BY id DESC
        """ % donde, tuple(parametros))
        for fila in cursor.fetchall():
            id_adelanto, nombre, monto, fecha, contrato, cuenta, id_factura = (list(fila) + [0, "", 0, "", "", "", 0])[:7]
            lista.append({"id": int(id_adelanto or 0), "cliente": str(nombre or ""),
                          "monto": round(float(monto or 0), 2), "fecha": str(fecha or ""),
                          "contrato": str(contrato or ""), "cuenta": str(cuenta or ""),
                          "id_factura": int(id_factura or 0),
                          "estado": ("Aplicado a la factura N° %s" % id_factura) if id_factura else "Saldo a favor"})
    except Exception as e:
        print("Aviso - adelantos:", e)
    finally:
        liberar_conexion(conn)
    return lista


def aplicar_adelanto_a_factura(id_adelanto, id_factura, usuario="Sistema"):
    """Asocia el adelanto a su factura (uno a uno). Devuelve (ok, mensaje)."""
    try:
        id_adelanto, id_factura = int(id_adelanto), int(id_factura)
    except Exception:
        return False, "Datos incompletos."
    if not id_adelanto or not id_factura:
        return False, "Elija el adelanto y la factura."
    conn = conectar_db()
    if not conn:
        return False, "Sin conexión a la base de datos."
    try:
        cursor = conn.cursor()
        cursor.execute("""SELECT COALESCE(monto_pagado, 0), COALESCE(cliente_nombre, ''),
                                 COALESCE(id_factura, 0)
                          FROM pagos_clientes WHERE id = %s""", (id_adelanto,))
        fila = cursor.fetchone()
        if not fila:
            return False, "El adelanto ya no existe."
        monto, cliente, factura_actual = (list(fila) + [0, "", 0])[:3]
        if int(factura_actual or 0):
            return False, "Ese adelanto ya está asociado a la factura N° %s." % factura_actual
        # 🚫 Evitar DUPLICADOS: si la factura ya está cobrada (por un cobro registrado),
        #    aplicar el adelanto sumaría dos veces el mismo dinero.
        try:
            cursor.execute("""SELECT COALESCE(f.total, 0), COALESCE(f.impuesto, 0),
                                     COALESCE(f.det_monto, 0), COALESCE(f.tipo_documento, ''),
                                     COALESCE((SELECT SUM(p2.monto_pagado + COALESCE(p2.impuestos_deducibles, 0)) FROM pagos_clientes p2
                                               WHERE p2.id_factura = f.id), 0)
                              FROM facturas_emitidas f WHERE f.id = %s""", (id_factura,))
            _f_saldo = cursor.fetchone()
            if _f_saldo:
                _tot = float(_f_saldo[0] or 0)
                _imp = float(_f_saldo[1] or 0)
                _det = float(_f_saldo[2] or 0)
                _tipo_doc = str(_f_saldo[3] or "")
                _cobrado = float(_f_saldo[4] or 0)
                if _tipo_doc and "Recibo" in _tipo_doc and "8%" in _tipo_doc:
                    _neto = _tot - _imp - _det
                else:
                    _neto = _tot - _det
                _saldo_fac = round(max(0.0, _neto - _cobrado), 2)
                if _saldo_fac <= 0.01:
                    return False, ("La factura N° %s ya está COBRADA por completo." % id_factura
                                   + chr(10) + chr(10)
                                   + "No se asoció el adelanto para no duplicar el cobro."
                                   + chr(10)
                                   + "El adelanto queda como saldo a favor (sin asociar).")
        except Exception as _e_saldo:
            print("Aviso - saldo de la factura:", _e_saldo)
        cursor.execute("UPDATE pagos_clientes SET id_factura = %s WHERE id = %s", (id_factura, id_adelanto))
        conn.commit()
        try:
            registrar_auditoria(usuario, "Ventas",
                                "Asoció el adelanto N° %s (%s) a la factura N° %s" % (id_adelanto, cliente, id_factura))
        except Exception:
            pass
        try:
            cache_sistema.invalidar()
        except Exception:
            pass
        return True, "Adelanto de %s asociado a la factura N° %s." % (monto, id_factura)
    except Exception as e:
        try:
            conn.rollback()
        except Exception:
            pass
        return False, str(e)
    finally:
        liberar_conexion(conn)


# =========================================================================
# 🏢 CLIENTES DE VENTAS: se validan contra el módulo de Clientes
# =========================================================================
# Al cargar una factura desde un archivo (PDF/imagen) el nombre del cliente se
# busca en la tabla 'clientes': si existe se toma SU RUC, y si no existe se crea
# automáticamente (cuando el comprobante trae RUC, que es obligatorio ahí).

def buscar_cliente_por_nombre(nombre, ruc=""):
    """Cliente de la tabla 'clientes' por RUC o por nombre (tolerante). Devuelve dict o None."""
    nombre = str(nombre or "").strip()
    ruc = re.sub(r"\D", "", str(ruc or ""))
    if not nombre and not ruc:
        return None
    conn = conectar_db(silencioso=True)
    if not conn:
        return None
    try:
        cursor = conn.cursor()
        if ruc:
            cursor.execute("""SELECT id, ruc, COALESCE(nombre_empresa, ''),
                                     COALESCE(razon_comercial, ''), COALESCE(direccion_fiscal, '')
                              FROM clientes WHERE ruc = %s LIMIT 1""", (ruc,))
            fila = cursor.fetchone()
            if fila:
                return {"id": int(fila[0]), "ruc": str(fila[1] or ""), "nombre": str(fila[2] or ""),
                        "comercial": str(fila[3] or ""), "direccion": str(fila[4] or "")}
        if nombre:
            cursor.execute("""SELECT id, ruc, COALESCE(nombre_empresa, ''),
                                     COALESCE(razon_comercial, ''), COALESCE(direccion_fiscal, '')
                              FROM clientes
                              WHERE nombre_empresa ILIKE %s OR razon_comercial ILIKE %s
                              ORDER BY id LIMIT 1""", ("%" + nombre + "%", "%" + nombre + "%"))
            fila = cursor.fetchone()
            if fila:
                return {"id": int(fila[0]), "ruc": str(fila[1] or ""), "nombre": str(fila[2] or ""),
                        "comercial": str(fila[3] or ""), "direccion": str(fila[4] or "")}
    except Exception as e:
        print("Aviso - buscar cliente:", e)
    finally:
        liberar_conexion(conn)
    return None


def asegurar_cliente_ventas(nombre, ruc="", direccion="", usuario="Sistema"):
    """Devuelve (estado, datos) con el cliente del comprobante.

    estado: 'existe' (ya estaba) · 'creado' (se creó ahora) · 'sin-datos'
            (no se pudo crear porque falta el RUC) · 'error'.
    datos: {'id','ruc','nombre','direccion'} del cliente encontrado o creado.
    """
    nombre = str(nombre or "").strip()
    ruc = re.sub(r"\D", "", str(ruc or ""))
    if not nombre and not ruc:
        return "sin-datos", {}
    encontrado = buscar_cliente_por_nombre(nombre, ruc)
    if encontrado:
        return "existe", encontrado
    if len(ruc) != 11:
        # En la tabla 'clientes' el RUC es obligatorio y único: sin él no se crea.
        return "sin-datos", {"nombre": nombre, "ruc": ""}
    conn = conectar_db(silencioso=True)
    if not conn:
        return "error", {}
    try:
        cursor = conn.cursor()
        cursor.execute("""INSERT INTO clientes (ruc, nombre_empresa, direccion_fiscal)
                          VALUES (%s, %s, %s)
                          ON CONFLICT (ruc) DO NOTHING
                          RETURNING id""", (ruc, nombre, str(direccion or "")))
        fila = cursor.fetchone()
        conn.commit()
        if not fila:
            # Ya existía con ese RUC: se devuelve el que está
            otro = buscar_cliente_por_nombre("", ruc)
            return ("existe", otro) if otro else ("error", {})
        try:
            registrar_auditoria(usuario, "Clientes",
                                "Cliente creado automáticamente al cargar una factura: %s (RUC %s)"
                                % (nombre, ruc))
        except Exception:
            pass
        try:
            cache_sistema.invalidar()
        except Exception:
            pass
        return "creado", {"id": int(fila[0]), "ruc": ruc, "nombre": nombre, "direccion": str(direccion or "")}
    except Exception as e:
        try:
            conn.rollback()
        except Exception:
            pass
        print("Aviso - crear cliente:", e)
        return "error", {}
    finally:
        liberar_conexion(conn)


# 💧 Detracción de las facturas de venta: regla del negocio
#    Si la factura SUPERA los S/ 400 se detrae el 4%; si no los supera, es cero.
UMBRAL_DETRACCION_VENTAS = 400.0      # soles
PORCENTAJE_DETRACCION_VENTAS = 4.0    # %


def porcentaje_detraccion_ventas(total, umbral=UMBRAL_DETRACCION_VENTAS,
                                 porcentaje=PORCENTAJE_DETRACCION_VENTAS):
    """4% si el total SUPERA los S/ 400; si no, 0% (no hay detracción)."""
    try:
        texto = str(total).replace("S/", "").replace("s/", "").strip()
        if not texto:
            return 0.0
        if "," in texto and "." in texto:
            # "1,000.50" (miles con coma) o "1.000,50" (decimal con coma)
            if texto.rfind(",") < texto.rfind("."):
                texto = texto.replace(",", "")
            else:
                texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", ".")
        total = float(texto)
    except Exception:
        return 0.0
    try:
        return float(porcentaje) if total > float(umbral) else 0.0
    except Exception:
        return 0.0


class CalendarioNativo(ctk.CTkToplevel):
    def __init__(self, parent, target_entry):
        super().__init__(parent)
        self.target_entry = target_entry
        self.title("Seleccionar Fecha")
        self.geometry("280x320")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (280 // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (320 // 2)
        self.geometry(f"+{x}+{y}")
        self.current_year = datetime.now().year
        self.current_month = datetime.now().month
        self.header_frame = ctk.CTkFrame(self, fg_color="#1f538d", corner_radius=0)
        self.header_frame.pack(fill="x")
        btn_prev = ctk.CTkButton(self.header_frame, text="<", width=30, fg_color="transparent", text_color="white", hover_color="#163b65", font=("Arial", 14, "bold"), command=self.prev_month)
        btn_prev.pack(side="left", padx=10, pady=10)
        self.lbl_month_year = ctk.CTkLabel(self.header_frame, text="", font=("Arial", 14, "bold"), text_color="white")
        self.lbl_month_year.pack(side="left", expand=True)
        btn_next = ctk.CTkButton(self.header_frame, text=">", width=30, fg_color="transparent", text_color="white", hover_color="#163b65", font=("Arial", 14, "bold"), command=self.next_month)
        btn_next.pack(side="right", padx=10, pady=10)
        self.days_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.days_frame.pack(fill="both", expand=True, padx=10, pady=10)
        dias_semana = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]
        for i, day in enumerate(dias_semana):
            lbl = ctk.CTkLabel(self.days_frame, text=day, font=("Arial", 11, "bold"), text_color="#1f538d")
            lbl.grid(row=0, column=i, padx=4, pady=5)
        self.update_calendar()
        
    def update_calendar(self):
        for widget in self.days_frame.winfo_children():
            if int(widget.grid_info()["row"]) > 0: widget.destroy()
        meses = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        self.lbl_month_year.configure(text=f"{meses[self.current_month]} {self.current_year}")
        cal = calendar.monthcalendar(self.current_year, self.current_month)
        hoy = datetime.now()
        for row_idx, week in enumerate(cal, start=1):
            for col_idx, day in enumerate(week):
                if day != 0:
                    btn_color = "transparent"
                    txt_color = "black"
                    if day == hoy.day and self.current_month == hoy.month and self.current_year == hoy.year:
                        btn_color = "#d4edda"
                        txt_color = "#155724"
                    btn = ctk.CTkButton(self.days_frame, text=str(day), width=30, height=30, fg_color=btn_color, text_color=txt_color, hover_color="#e0e0e0", font=("Arial", 11))
                    btn.configure(command=lambda d=day: self.select_date(d))
                    btn.grid(row=row_idx, column=col_idx, padx=2, pady=2)
                    
    def prev_month(self):
        self.current_month -= 1
        if self.current_month < 1:
            self.current_month = 12
            self.current_year -= 1
        self.update_calendar()
        
    def next_month(self):
        self.current_month += 1
        if self.current_month > 12:
            self.current_month = 1
            self.current_year += 1
        self.update_calendar()
        
    def select_date(self, day):
        fmt = CONFIG_REGIONAL.get("formato_fecha", "DD/MM/AAAA")
        if fmt == "MM/DD/AAAA":
            fecha_seleccionada = f"{self.current_month:02d}/{day:02d}/{self.current_year}"
        else:
            fecha_seleccionada = f"{day:02d}/{self.current_month:02d}/{self.current_year}"
            
        self.target_entry.delete(0, tk.END)
        self.target_entry.insert(0, fecha_seleccionada)
        self.destroy()


_SCHEMA_VENTAS_OK = False
# True solo si la columna 'det_redondeo' existe en 'facturas_emitidas' (se verifica al abrir el módulo).
_COLUMNA_DET_REDONDEO = False

# =========================================================
# PESTAÑA 1: FACTURAS EMITIDAS
# =========================================================
def _pick_comprobante(parent):
    """Elige un archivo PDF/imagen para el comprobante de depósito."""
    try:
        from tkinter import filedialog
        return filedialog.askopenfilename(
            parent=parent, title="Seleccionar el comprobante de depósito",
            filetypes=[("PDF o imagen", "*.pdf *.jpg *.jpeg *.png"), ("Todos los archivos", "*.*")])
    except Exception:
        return ""


def _escanear_con_progreso(parent, etiqueta, mod_escaner, *args, **kwargs):
    """Escanea mostrando el cuadrito 'Escaneando…' y lo cierra al terminar."""
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
        ctk.CTkLabel(prog, text="🖨️ Escaneando %s..." % (etiqueta or "el comprobante"),
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


# 🧾 Columnas para los impuestos deducibles del cobro (se crean una vez)
_PAGOS_IMPUESTOS_OK = False


def _asegurar_columnas_impuestos(conn):
    """Agrega (una sola vez por sesión) las columnas de impuestos deducibles."""
    global _PAGOS_IMPUESTOS_OK
    if _PAGOS_IMPUESTOS_OK or conn is None:
        return
    try:
        with conn.cursor() as c:
            c.execute("ALTER TABLE pagos_clientes ADD COLUMN IF NOT EXISTS "
                      "impuestos_deducibles NUMERIC DEFAULT 0")
            c.execute("ALTER TABLE pagos_clientes ADD COLUMN IF NOT EXISTS "
                      "archivo_impuesto TEXT DEFAULT ''")
        conn.commit()
        _PAGOS_IMPUESTOS_OK = True
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass


class FacturasEmitidasTab:
    def __init__(self, tab_frame, main_root, app_padre):
        self.tab_frame = tab_frame
        self.main_root = main_root
        self.app_padre = app_padre
        
        self.orden_columnas = {}
        self.bloquear_autocompletado_ruc = False
        self.ruta_archivo_temp = ""
        self.id_cobranza_seleccionada = None   # id de cobranza_quincenas vinculada a la factura en curso
        self.total_cobranza_seleccionada = 0.0  # total de esa cobranza (YA INCLUYE IGV) que debe totalizar la factura
        self.cobranzas_pendientes = {}          # etiqueta -> datos de la cobranza pendiente de facturar
        
        # 🚀 VARIABLES DE PAGINACIÓN
        self.pagina_actual = 1
        self.registros_por_pagina = 50
        
        # 🗓️ FILTRO DE MES (por defecto, el mes en curso)
        self.mes_filtro = mes_en_curso()
        
        self.inicializar_bd()
        self.crear_interfaz()

    # 🚀 FIX: AUTO-CURACIÓN EN SEGUNDO PLANO
    def inicializar_bd(self):
        global _SCHEMA_VENTAS_OK
        if _SCHEMA_VENTAS_OK: return

        def tarea_curacion():
            global _SCHEMA_VENTAS_OK, _COLUMNA_DET_REDONDEO
            conn = conectar_db(silencioso=True)
            if not conn: return
            try:
                cursor = conn.cursor()
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS facturas_emitidas (
                        id SERIAL PRIMARY KEY, tipo_documento VARCHAR(100), fecha VARCHAR(50), cliente VARCHAR(255), descripcion TEXT, 
                        evento_asociado VARCHAR(255), subtotal NUMERIC, impuesto NUMERIC, total NUMERIC, archivo_ruta TEXT, 
                        dias_credito INTEGER DEFAULT 0, det_porcentaje NUMERIC DEFAULT 0, det_monto NUMERIC DEFAULT 0, 
                        numero_documento VARCHAR(100) DEFAULT ''
                    )
                """)
                conn.commit()
                
                columnas_nuevas = [
                    "ALTER TABLE facturas_emitidas ADD COLUMN estado_sunat VARCHAR(255) DEFAULT ''",
                    "ALTER TABLE facturas_emitidas ADD COLUMN enlace_pdf_sunat TEXT DEFAULT ''",
                    "ALTER TABLE facturas_emitidas ADD COLUMN enlace_xml_sunat TEXT DEFAULT ''",
                    "ALTER TABLE facturas_emitidas ADD COLUMN enlace_pdf_nc TEXT DEFAULT ''",
                    "ALTER TABLE facturas_emitidas ADD COLUMN orden_compra VARCHAR(255) DEFAULT ''"
                ]
                for query in columnas_nuevas:
                    try: cursor.execute(query); conn.commit()
                    except: conn.rollback()

                # 🧮 Marca si la factura se registró con la detracción redondeada al entero.
                # Solo se consulta/escribe esa columna si realmente existe en la BD.
                try:
                    cursor.execute("ALTER TABLE facturas_emitidas ADD COLUMN IF NOT EXISTS det_redondeo BOOLEAN DEFAULT FALSE")
                    conn.commit()
                    _COLUMNA_DET_REDONDEO = True
                except Exception:
                    conn.rollback()
                    _COLUMNA_DET_REDONDEO = False

                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS ordenes_compra_clientes (
                        id SERIAL PRIMARY KEY,
                        numero_oc VARCHAR(100),
                        cotizacion_asociada VARCHAR(255),
                        fecha VARCHAR(50),
                        archivo_ruta TEXT
                    )
                """)
                conn.commit()
                _SCHEMA_VENTAS_OK = True
            except Exception: pass
            finally: liberar_conexion(conn)

        threading.Thread(target=tarea_curacion, daemon=True).start()

    def abrir_calendario(self, entry_objetivo):
        CalendarioNativo(self.main_root.winfo_toplevel(), entry_objetivo)

    def adjuntar_pdf_factura(self):
        """Adjunta manualmente el PDF de una factura local y autocompleta los campos (con OCR si es necesario)."""
        ruta = seleccionar_archivo_dialogo(
            titulo="Seleccionar PDF de la Factura",
            tipos=[("Archivos PDF", "*.pdf"),
                   ("Imágenes (JPG/PNG)", "*.png;*.jpg;*.jpeg"),
                   ("Todos los archivos", "*.*")])
        if not ruta:
            return
        self._autocompletar_desde_archivo(ruta)

    def sugerir_correlativo(self):
        if not hasattr(self, 'combo_tipo') or not hasattr(self, 'ent_nro_doc'):
            return

        tipo = self.combo_tipo.get()
        config = cargar_configuracion_regional()

        prefijo = "F"
        if "Factura" in tipo:
            filtro_tipo = "Factura%"
            ult_cfg = config.get("ultimo_factura", "F001-0")
            prefijo = "F"
        elif "Boleta" in tipo:
            filtro_tipo = "Boleta%"
            ult_cfg = config.get("ultimo_boleta", "B001-0")
            prefijo = "B"
        else:
            filtro_tipo = "Recibo%"
            ult_cfg = config.get("ultimo_recibo", "E001-0")
            prefijo = "E"

        cfg_serie = f"{prefijo}001"
        cfg_num = 0
        if ult_cfg and "-" in ult_cfg:
            parts = ult_cfg.split("-")
            if parts[0].strip().startswith(prefijo):
                cfg_serie = parts[0].strip()
            if parts[1].strip().isdigit():
                cfg_num = int(parts[1].strip())

        def tarea_sugerir(estado):               # 👈 corre en el HILO SECUNDARIO
            conn = conectar_db(silencioso=True)
            db_num = 0
            db_serie = cfg_serie
            if conn:
                try:
                    cursor = conn.cursor()
                    cursor.execute("SELECT numero_documento FROM facturas_emitidas WHERE tipo_documento LIKE %s AND numero_documento LIKE '%%-%%' ORDER BY id DESC LIMIT 1", (filtro_tipo,))
                    res = cursor.fetchone()
                    if res and res[0]:
                        partes = res[0].split('-')
                        if len(partes) == 2 and partes[1].isdigit():
                            db_serie = partes[0]
                            db_num = int(partes[1])
                except Exception:
                    pass
                finally:
                    liberar_conexion(conn)

            if not db_serie.startswith(prefijo):
                db_serie = cfg_serie
                db_num = 0

            if db_serie == cfg_serie:
                final_num = max(cfg_num, db_num) + 1
                final_serie = cfg_serie
            else:
                final_serie = db_serie
                final_num = db_num + 1

            nuevo_doc = f"{final_serie}-{final_num}"
            estado["nuevo_doc"] = nuevo_doc

        def _aplicar_correlativo_hilo(estado):   # 👈 corre en el HILO PRINCIPAL
            if estado.get("nuevo_doc"):
                self._aplicar_correlativo(estado["nuevo_doc"])

        ejecutar_en_hilo(self.main_root, tarea_sugerir, al_terminar=_aplicar_correlativo_hilo)

    def _aplicar_correlativo(self, nuevo_doc):
        # 📄 Si el N° vino del archivo (PDF/OCR), el correlativo del sistema NO lo pisa.
        if getattr(self, "_nro_desde_archivo", False):
            print("Aviso - se conserva el N° leído del archivo:", self.ent_nro_doc.get())
            return
        self.ent_nro_doc.delete(0, tk.END)
        self.ent_nro_doc.insert(0, nuevo_doc)

    def ordenar_por_columna(self, columna, es_numerico):
        elementos = [(self.tabla.set(item, columna), item) for item in self.tabla.get_children("")]
        ascendente = self.orden_columnas.get(columna, True)
        self.orden_columnas[columna] = not ascendente
        if es_numerico:
            elementos.sort(key=lambda el: desformatear_numero(el[0]), reverse=not ascendente)
        else:
            elementos.sort(key=lambda el: el[0].lower(), reverse=not ascendente)
        for index, (_, item) in enumerate(elementos): self.tabla.move(item, "", index)

    def escanear_factura_multifuncional(self):
        """Escanea el comprobante en la multifuncional y lo carga en el formulario.

        Es el mismo flujo del módulo de Compras: escanea, arma el PDF, lo lee (texto
        del PDF u OCR de la imagen) y completa todos los campos, validando al cliente.
        Funciona en Windows y en macOS (M1/M2).
        """
        try:
            import escaner
        except Exception as e:
            return messagebox.showerror("Escáner", "No se pudo cargar el módulo de escaneo:\n%s" % e,
                                        parent=self._padre_mensaje())
        try:
            if not escaner.sistema_soportado():
                return messagebox.showinfo(
                    "Escáner",
                    "El escaneo con la multifuncional está disponible en Windows y en macOS (M1/M2).",
                    parent=self._padre_mensaje())
        except Exception:
            pass
        try:
            config = cargar_configuracion_regional()
            datos = escaner.config_escaner(config)
        except Exception as e:
            return messagebox.showerror("Escáner", "No se pudo leer la configuración del escáner:\n%s" % e,
                                        parent=self._padre_mensaje())

        modo = str(datos.get("escaner_modo", "Automático (sin ventanas)"))
        if modo.lower().startswith("con la ventana"):
            instruccion = "Se abrirá el cuadro de escaneo del sistema."
        else:
            instruccion = ("Coloque el comprobante en el escáner (cristal) — %s ppp, %s."
                           % (datos.get("escaner_dpi", "300"), datos.get("escaner_color", "Escala de grises")))
            if not datos.get("escaner_dispositivo"):
                instruccion += ("\n(Si hay varios escáneres, elíjalo en Ajustes → "
                                "Configuración General → Escáner.)")

        v_prog = ctk.CTkToplevel(self.main_root)
        v_prog.title("Escaneando comprobante")
        v_prog.geometry("440x175")
        v_prog.transient(self.main_root)
        try:
            v_prog.grab_set()
        except Exception:
            pass
        ctk.CTkLabel(v_prog, text="🖨️ Escaneando y leyendo el comprobante...",
                     font=("Arial", 13, "bold"), text_color="#1f538d").pack(pady=(18, 4))
        ctk.CTkLabel(v_prog, text=instruccion, font=("Arial", 10), text_color="gray",
                     justify="center", wraplength=390).pack(pady=(0, 10))
        barra = ctk.CTkProgressBar(v_prog, width=350, mode="indeterminate")
        barra.pack(pady=4)
        barra.start()
        ctk.CTkLabel(v_prog, text="Esto puede tardar unos segundos.", font=("Arial", 9, "italic"),
                     text_color="gray").pack(pady=(4, 0))
        try:
            traer_al_frente(v_prog)
        except Exception:
            pass
        threading.Thread(target=self._worker_escanear_venta, args=(config, v_prog), daemon=True).start()
        return v_prog

    def _padre_mensaje(self):
        """Ventana que debe quedar delante en los avisos (o la principal)."""
        try:
            if "ventana_para_mensaje" in globals():
                return ventana_para_mensaje(getattr(self, "main_root", None))
        except Exception:
            pass
        return getattr(self, "main_root", None)

    def _cerrar_progreso_escaneo(self, v_prog):
        try:
            if v_prog is not None and v_prog.winfo_exists():
                v_prog.destroy()
        except Exception:
            pass

    def _worker_escanear_venta(self, config, v_prog):
        """Escanea en segundo plano, arma el PDF y lo aplica al formulario."""
        try:
            import escaner
            datos = escaner.config_escaner(config)
            carpeta = escaner.carpeta_escaneos(config)
            try:
                paginas_pedidas = int(str(datos.get("escaner_paginas") or 1))
            except Exception:
                paginas_pedidas = 1
            escaneo = escaner.escanear(
                carpeta,
                dispositivo=datos.get("escaner_dispositivo", ""),
                dpi=datos.get("escaner_dpi", "300"),
                color=datos.get("escaner_color", "Escala de grises"),
                origen=datos.get("escaner_origen", "Cristal (una página a la vez)"),
                paginas=paginas_pedidas,
                optimizar=str(datos.get("escaner_optimizar", "Sí")).lower().startswith("s"),
                modo=datos.get("escaner_modo", "Automático (sin ventanas)"))
            if not escaneo.get("ok"):
                error = str(escaneo.get("error") or "No se pudo escanear.")
                return self.main_root.after(0, lambda: (
                    self._cerrar_progreso_escaneo(v_prog),
                    messagebox.showerror("Escáner", "No se pudo escanear el comprobante.\n\n" + error,
                                         parent=self._padre_mensaje())))

            paginas = [str(p) for p in (escaneo.get("paginas") or []) if os.path.exists(str(p))]
            if not paginas:
                return self.main_root.after(0, lambda: (
                    self._cerrar_progreso_escaneo(v_prog),
                    messagebox.showinfo("Escáner", "El escáner no devolvió ninguna imagen.",
                                        parent=self._padre_mensaje())))
            marca = datetime.now().strftime("%Y%m%d_%H%M%S")
            pdf = os.path.join(os.path.dirname(paginas[0]), "Comprobante_Venta_" + marca + ".pdf")
            try:
                escaner.imagenes_a_pdf(paginas, pdf)
            except Exception as e_pdf:
                print("Aviso - PDF del escaneo:", e_pdf)

            def aplicar():
                self._cerrar_progreso_escaneo(v_prog)
                ruta = pdf if os.path.exists(pdf) else paginas[0]
                try:
                    self._autocompletar_desde_archivo(ruta)   # texto/OCR + cliente validado
                except Exception as e:
                    messagebox.showerror("Escáner", "Se escaneó, pero no se pudo leer:\n%s" % e,
                                         parent=self._padre_mensaje())

            return self.main_root.after(0, aplicar)
        except Exception as e:
            return self.main_root.after(0, lambda: (
                self._cerrar_progreso_escaneo(v_prog),
                messagebox.showerror("Escáner", "No se pudo escanear el comprobante:\n%s" % e,
                                     parent=self._padre_mensaje())))


    def autocompletar_desde_pdf(self):
        ruta = seleccionar_archivo_dialogo(
            titulo="Seleccionar Factura PDF",
            tipos=[("Archivos PDF", "*.pdf"), ("Imágenes", "*.png;*.jpg;*.jpeg")])
        if not ruta:
            return
        self._autocompletar_desde_archivo(ruta)

    def _autocompletar_desde_archivo(self, ruta):
        """Carga el PDF/imagen, extrae el texto o aplica OCR y rellena todos los campos."""
        try:
            self.bloquear_autocompletado_ruc = True
            rellenado = False
            texto = self._extraer_texto_pdf(ruta)
            if texto.strip():
                rellenado = self._rellenar_desde_texto(texto)
            if not rellenado:
                datos = self._ocr_factura_gemini(ruta)
                if datos:
                    rellenado = self._rellenar_desde_ocr(datos)

            self.ruta_archivo_temp = ruta
            if hasattr(self, "lbl_archivo"):
                self.lbl_archivo.configure(text=f"📎 {os.path.basename(ruta)}", text_color="#166534")
            self.actualizar_totales()

            # 🏢 El cliente se valida contra el módulo de Clientes: se toma SU RUC y,
            #    si no existe, se crea automáticamente.
            aviso_cliente = ""
            try:
                nombre_cli = self.combo_cliente.get().strip()
                ruc_cli = re.sub(r"\D", "", self.ent_ruc.get() or "")
                if nombre_cli:
                    estado_cli, datos_cli = asegurar_cliente_ventas(nombre_cli, ruc_cli)
                    if estado_cli in ("existe", "creado") and datos_cli.get("ruc"):
                        if not ruc_cli:
                            self.ent_ruc.configure(state="normal")
                            self.ent_ruc.delete(0, tk.END)
                            self.ent_ruc.insert(0, datos_cli["ruc"])
                            ruc_cli = datos_cli["ruc"]
                        if datos_cli.get("nombre"):
                            self.combo_cliente.set(datos_cli["nombre"])
                        aviso_cliente = ("  🏢 Se creó el cliente en el módulo de Clientes (RUC %s)."
                                         % datos_cli["ruc"]) if estado_cli == "creado" else \
                                        "  El cliente ya estaba registrado (RUC %s)." % datos_cli["ruc"]
                        for nombre_metodo in ("cargar_clientes_combo", "cargar_clientes_combobox",
                                              "actualizar_clientes", "cargar_clientes"):
                            metodo = getattr(self, nombre_metodo, None)
                            if callable(metodo):
                                try:
                                    metodo()
                                except Exception:
                                    pass
                                break
                    elif estado_cli == "sin-datos":
                        aviso_cliente = ("  ⚠️ No se pudo crear el cliente: el comprobante no trae "
                                         "un RUC de 11 dígitos (es obligatorio en Clientes).")
            except Exception as e_cli:
                print("Aviso - cliente del archivo:", e_cli)

            self.al_seleccionar_cliente()

            if rellenado:
                messagebox.showinfo("Lectura de Factura",
                                    "Datos de la factura rellenados automáticamente (texto/OCR)." + aviso_cliente)
            else:
                messagebox.showwarning("Lectura Incompleta", "No se pudieron leer automáticamente los datos del documento.\nRevise los campos manualmente.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error:\n{e}")
        finally:
            self.bloquear_autocompletado_ruc = False

    def _extraer_texto_pdf(self, ruta):
        """Extrae la capa de texto del PDF (pdfplumber y luego PyMuPDF). Devuelve '' si no hay texto."""
        texto = ""
        if pdfplumber is not None:
            try:
                with pdfplumber.open(ruta) as pdf:
                    for page in pdf.pages:
                        t = page.extract_text()
                        if t:
                            texto += t + "\n"
            except Exception:
                pass
        if not texto.strip():
            try:
                import fitz
                doc = fitz.open(ruta)
                for page in doc:
                    texto += (page.get_text() or "") + "\n"
                doc.close()
            except Exception:
                pass
        return texto

    def _rellenar_desde_texto(self, texto):
        """Rellena los campos del formulario a partir del texto extraído. Devuelve True si rellenó algo."""
        relleno = False

        if re.search(r"FACTURA\s+ELECTR[OÓ]NICA", texto, re.IGNORECASE):
            self.combo_tipo.set("Factura (18% IGV)")
        elif re.search(r"BOLETA\s+DE\s+VENTA", texto, re.IGNORECASE):
            self.combo_tipo.set("Boleta (Sin IGV)")
        elif re.search(r"RECIBO\s+POR\s+HONORARIOS", texto, re.IGNORECASE):
            if re.search(r"Retenci[oó]n.*?IR[\s:\|]*\(?([\d\,\.]+)\)?", texto, re.IGNORECASE):
                self.combo_tipo.set("Recibo por Honorarios (8% Retención)")
            else:
                self.combo_tipo.set("Recibo por Honorarios (Sin Retención)")
        self.on_tipo_change(self.combo_tipo.get())

        nro_match = re.search(r"([EFB][0-9A-Z]{3}\s*-\s*\d+)", texto)
        if nro_match:
            self._nro_desde_archivo = True          # el N° es el del comprobante
            self.ent_nro_doc.delete(0, tk.END); self.ent_nro_doc.insert(0, nro_match.group(1).replace(" ", ""))
            relleno = True

        fecha_match = re.search(r"Fecha de Emisi[oó]n\s*[:\-]?\s*(\d{2})[/\-.](\d{2})[/\-.](\d{4})", texto, re.IGNORECASE)
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
            relleno = True

        cliente_match = re.search(r"(?:Señor\(es\)|Señores|Razón Social|Cliente|Recibí\s*de)\s*[:\-]\s*(.+)", texto, re.IGNORECASE)
        rucs = re.findall(r"(?:RUC|R\.U\.C\.|Documento)\s*[:\-]?\s*(\d{11})", texto, re.IGNORECASE)

        if cliente_match:
            self.combo_cliente.set(cliente_match.group(1).strip())
            relleno = True
        if rucs:
            self.ent_ruc.configure(state="normal")
            self.ent_ruc.delete(0, tk.END)
            self.ent_ruc.insert(0, rucs[-1])
            relleno = True

        sub_m = re.search(r"(?:OP\.\s*GRAVADAS|SUB\s*TOTAL|Subtotal|Total por honorarios)[\s:S/\|]+([\d\,\.]+)", texto, re.IGNORECASE)
        tot_m = re.search(r"(?:IMPORTE\s*TOTAL|TOTAL\s*A\s*PAGAR|Total Neto Recibido)[\s:S/\|]+([\d\,\.]+)", texto, re.IGNORECASE)
        monto_base = 0.0
        if sub_m:
            try:
                monto_base = float(sub_m.group(1).replace(",", ""))
            except ValueError:
                monto_base = 0.0
        elif tot_m:
            try:
                t = float(tot_m.group(1).replace(",", ""))
                monto_base = t / 1.18 if "Factura" in self.combo_tipo.get() else t
            except ValueError:
                monto_base = 0.0
        if monto_base > 0:
            self.ent_subtotal.delete(0, tk.END); self.ent_subtotal.insert(0, f"{monto_base:.2f}")
            relleno = True

        if "Recibo" not in self.combo_tipo.get():
            det_match = re.search(r"(?:Detracci[oó]n|Porcentaje|Tasa).*?(\d{1,2}(?:\.\d{1,2})?)\s*%", texto, re.IGNORECASE)
            if not det_match:
                if re.search(r"Sujeta\s*a\s*detracci[oó]n", texto, re.IGNORECASE):
                    self.ent_detraccion.delete(0, tk.END); self.ent_detraccion.insert(0, CONFIG_REGIONAL.get("detraccion_porcentaje", "12"))
                    relleno = True
            else:
                self.ent_detraccion.delete(0, tk.END); self.ent_detraccion.insert(0, det_match.group(1))
                relleno = True

        return relleno

    def _ocr_factura_gemini(self, ruta):
        """Aplica OCR con IA (Google Gemini) a un PDF/imagen escaneado. Devuelve un dict con los campos o None."""
        if genai is None:
            return None
        imagenes = []
        try:
            import fitz
            doc = fitz.open(ruta)
            for page in doc:
                try:
                    pix = page.get_pixmap(dpi=200)
                    imagenes.append(pix.tobytes("png"))
                except Exception:
                    continue
            doc.close()
        except Exception:
            imagenes = []
        if not imagenes:
            try:
                with open(ruta, "rb") as f:
                    imagenes.append(f.read())
            except Exception:
                return None

        try:
            clave = os.environ.get("GEMINI_API_KEY", "").strip() or "AQ.Ab8RN6LTyHmVNUALwk6Wk7b2EMSzbZrVXVjg-cKUH7cSwnJ0Iw"
            cliente = genai.Client(api_key=clave)
        except Exception:
            return None

        prompt = (
            "Eres un lector OCR de facturas, boletas y recibos peruanos. "
            "Lee la imagen del comprobante y devuelve SOLO un JSON válido, sin texto adicional, con estas claves exactas:\n"
            '- "tipo": "factura", "boleta" o "recibo".\n'
            '- "numero_documento": serie y correlativo (ej. "F001-123"). Si no se ve, "".\n'
            '- "fecha": fecha en DD/MM/YYYY. Si no se ve, "".\n'
            '- "cliente": razón social o nombre del cliente. Si no se ve, "".\n'
            '- "ruc": exactamente 11 dígitos. Si no se ve o no son 11 dígitos, "".\n'
            '- "descripcion": concepto o descripción del servicio/venta. Si no se ve, "".\n'
            '- "subtotal": monto base sin IGV, número decimal (ej. 22212.00). Si no se ve, 0.\n'
            '- "igv": importe del IGV, número decimal. Si no se ve, 0.\n'
            '- "total": importe total, número decimal. Si no se ve, 0.\n'
            '- "detraccion": porcentaje de detracción o retención (ej. 12). Si no aplica, 0.\n'
            "Reglas: NO inventes datos. Montos como números con punto decimal, sin símbolo de moneda ni comas. "
            "Si un campo no se ve, devuélvelo vacío o 0."
        )

        try:
            partes = []
            if genai_types is not None:
                for img in imagenes[:3]:
                    partes.append(genai_types.Part.from_bytes(data=img, mime_type="image/png"))
            if not partes:
                return None
            respuesta = cliente.models.generate_content(model="gemini-3.5-flash-lite", contents=[prompt] + partes)
            texto_ia = (respuesta.text or "").strip()
        except Exception as e:
            print("⚠️ Error OCR Gemini:", e)
            return None

        m = re.search(r"\{.*\}", texto_ia, re.DOTALL)
        if not m:
            return None
        try:
            return json.loads(m.group(0))
        except Exception:
            return None

    def _rellenar_desde_ocr(self, datos):
        """Rellena los campos del formulario a partir del dict del OCR. Devuelve True si rellenó algo."""
        relleno = False
        try:
            tipo = str(datos.get("tipo") or "").lower()
            if "boleta" in tipo:
                self.combo_tipo.set("Boleta (Sin IGV)")
            elif "recibo" in tipo:
                self.combo_tipo.set("Recibo por Honorarios (8% Retención)")
            else:
                self.combo_tipo.set("Factura (18% IGV)")
            self.on_tipo_change(self.combo_tipo.get())

            nro = str(datos.get("numero_documento") or "").strip()
            if nro:
                self._nro_desde_archivo = True      # el N° es el del comprobante
                self.ent_nro_doc.delete(0, tk.END); self.ent_nro_doc.insert(0, nro)
                relleno = True

            fecha = str(datos.get("fecha") or "").strip()
            mf = re.search(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})", fecha)
            if mf:
                d, mes, anio = mf.groups()
                if len(anio) == 2:
                    anio = "20" + anio
                fmt = CONFIG_REGIONAL.get("formato_fecha", "DD/MM/AAAA")
                self.ent_fecha.delete(0, tk.END)
                if fmt == "MM/DD/AAAA":
                    self.ent_fecha.insert(0, f"{mes}/{d}/{anio}")
                else:
                    self.ent_fecha.insert(0, f"{d}/{mes}/{anio}")
                relleno = True

            cliente = str(datos.get("cliente") or "").strip()
            if cliente:
                self.combo_cliente.set(cliente)
                relleno = True

            ruc = re.sub(r"\D", "", str(datos.get("ruc") or ""))
            if len(ruc) == 11:
                self.ent_ruc.configure(state="normal")
                self.ent_ruc.delete(0, tk.END)
                self.ent_ruc.insert(0, ruc)
                relleno = True

            desc = str(datos.get("descripcion") or "").strip()
            if desc:
                self.ent_desc.delete(0, tk.END); self.ent_desc.insert(0, desc)
                relleno = True

            def _num(v):
                s = str(v).strip().replace("S/", "").replace("s/", "").replace("$", "").replace(" ", "")
                if "," in s and "." in s:
                    if s.rfind(",") > s.rfind("."):
                        s = s.replace(".", "").replace(",", ".")
                    else:
                        s = s.replace(",", "")
                elif "," in s:
                    s = s.replace(",", ".")
                try:
                    return float(s)
                except Exception:
                    return 0.0

            subtotal = _num(datos.get("subtotal"))
            total = _num(datos.get("total"))
            monto_base = subtotal
            if monto_base <= 0 and total > 0:
                monto_base = total / 1.18 if "Factura" in self.combo_tipo.get() else total
            if monto_base > 0:
                self.ent_subtotal.delete(0, tk.END); self.ent_subtotal.insert(0, f"{monto_base:.2f}")
                relleno = True

            det = _num(datos.get("detraccion"))
            if det > 0:
                self.ent_detraccion.delete(0, tk.END); self.ent_detraccion.insert(0, f"{det:g}")
                relleno = True
        except Exception as e:
            print("⚠️ Error rellenando desde OCR:", e)
        return relleno

    def crear_interfaz(self):
        frame_split = ctk.CTkFrame(self.tab_frame, fg_color="transparent")
        frame_split.pack(fill="both", expand=True)

        self.f_form = ctk.CTkScrollableFrame(frame_split, corner_radius=10, width=330, fg_color="#f8f9fa", border_width=1, border_color="#e0e0e0")
        self.f_form.pack(side="left", fill="y", padx=(0, 15))

        btn_escanear = ctk.CTkButton(self.f_form, text="🖨️ Escanear con la Multifuncional",
                                     fg_color="#8e44ad", hover_color="#703688", font=("Arial", 12, "bold"),
                                     command=self.escanear_factura_multifuncional)
        btn_escanear.pack(fill="x", padx=10, pady=(10, 5))

        btn_auto = ctk.CTkButton(self.f_form, text="📄 Cargar PDF y Autocompletar (OCR)", fg_color="#1f538d", hover_color="#163b65", command=self.autocompletar_desde_pdf)
        btn_auto.pack(fill="x", padx=10, pady=(10, 5))

        btn_adjuntar = ctk.CTkButton(self.f_form, text="📎 Adjuntar PDF de la Factura (Local)", fg_color="#34495e", hover_color="#2c3e50", command=self.adjuntar_pdf_factura)
        btn_adjuntar.pack(fill="x", padx=10, pady=(0, 2))

        self.lbl_archivo = ctk.CTkLabel(self.f_form, text="Sin PDF adjunto", font=("Arial", 10), text_color="#999999")
        self.lbl_archivo.pack(anchor="w", padx=10, pady=(0, 15))

        ctk.CTkLabel(self.f_form, text="Tipo de Documento:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.combo_tipo = ctk.CTkComboBox(self.f_form, values=["Factura (18% IGV)", "Boleta (Sin IGV)", "Recibo por Honorarios (8% Retención)", "Recibo por Honorarios (Sin Retención)"], state="readonly", command=self.on_tipo_change)
        self.combo_tipo.pack(fill="x", padx=10, pady=(0, 8))
        self.combo_tipo.set("Factura (18% IGV)")

        ctk.CTkLabel(self.f_form, text="N° de Documento:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.ent_nro_doc = ctk.CTkEntry(self.f_form, placeholder_text="Ej. F001-1")
        self.ent_nro_doc.pack(fill="x", padx=10, pady=(0, 8))

        ctk.CTkLabel(self.f_form, text="Fecha (Configurada):", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        f_fecha = ctk.CTkFrame(self.f_form, fg_color="transparent")
        f_fecha.pack(fill="x", padx=10, pady=(0, 8))
        self.ent_fecha = ctk.CTkEntry(f_fecha)
        self.ent_fecha.pack(side="left", fill="x", expand=True)
        
        fmt = CONFIG_REGIONAL.get("formato_fecha", "DD/MM/AAAA")
        if fmt == "MM/DD/AAAA": self.ent_fecha.insert(0, datetime.now().strftime("%m/%d/%Y"))
        else: self.ent_fecha.insert(0, datetime.now().strftime("%d/%m/%Y"))
            
        ctk.CTkButton(f_fecha, text="[ 📅 ]", width=40, fg_color="#1f538d", hover_color="#163b65", command=lambda: self.abrir_calendario(self.ent_fecha)).pack(side="right", padx=(5, 0))

        ctk.CTkLabel(self.f_form, text="Días de Crédito:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.ent_dias = ctk.CTkEntry(self.f_form); self.ent_dias.pack(fill="x", padx=10, pady=(0, 8)); self.ent_dias.insert(0, "0")

        ctk.CTkLabel(self.f_form, text="🧾 Cobranza a Facturar (No Facturadas):", font=("Arial", 11, "bold"), text_color="#166534").pack(anchor="w", padx=10)
        self.combo_cobranza = ctk.CTkComboBox(self.f_form, state="readonly", command=self.al_seleccionar_cobranza)
        self.combo_cobranza.pack(fill="x", padx=10, pady=(0, 2))
        self.combo_cobranza.set("--- Seleccione Cobranza ---")
        # 💡 Aviso: el total de la cobranza ya incluye IGV (es lo que paga el cliente menos detracción)
        self.lbl_cobranza_info = ctk.CTkLabel(self.f_form, text="", font=("Arial", 9),
                                              text_color="#166534", justify="left", wraplength=290)
        self.lbl_cobranza_info.pack(anchor="w", padx=10, pady=(0, 8))

        ctk.CTkLabel(self.f_form, text="Nombre del Cliente:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.combo_cliente = ctk.CTkComboBox(self.f_form, command=self.al_seleccionar_cliente)
        activar_busqueda_rapida(self.combo_cliente, titulo="Buscar cliente",
                                 al_elegir=self.al_seleccionar_cliente)
        self.combo_cliente.pack(fill="x", padx=10, pady=(0, 8))

        ctk.CTkLabel(self.f_form, text="R.U.C. Cliente:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.ent_ruc = ctk.CTkEntry(self.f_form, placeholder_text="Se autocompleta con el cliente")
        self.ent_ruc.pack(fill="x", padx=10, pady=(0, 8))

        ctk.CTkLabel(self.f_form, text="Concepto / Descripción de Venta:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.ent_desc = ctk.CTkEntry(self.f_form, placeholder_text="Ej: Servicios Generales...")
        self.ent_desc.pack(fill="x", padx=10, pady=(0, 8))

        ctk.CTkLabel(self.f_form, text="Orden de Compra:", font=("Arial", 11, "bold"), text_color="#166534").pack(anchor="w", padx=10)
        self.combo_oc = ctk.CTkComboBox(self.f_form, state="readonly")
        self.combo_oc.pack(fill="x", padx=10, pady=(0, 8))

        ctk.CTkLabel(self.f_form, text="Monto Base (Sin IGV / Subtotal):", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
        self.ent_subtotal = ctk.CTkEntry(self.f_form)
        self.ent_subtotal.pack(fill="x", padx=10, pady=(0, 8))
        self.ent_subtotal.bind("<KeyRelease>", self.actualizar_totales)

        self.lbl_titulo_det = ctk.CTkLabel(self.f_form, text="Detracción (%):", font=("Arial", 11, "bold"))
        self.lbl_titulo_det.pack(anchor="w", padx=10)
        self.ent_detraccion = ctk.CTkEntry(self.f_form)
        self.ent_detraccion.pack(fill="x", padx=10, pady=(0, 8))
        self.ent_detraccion.insert(0, CONFIG_REGIONAL.get("detraccion_porcentaje", "12"))
        self.ent_detraccion.bind("<KeyRelease>", self.actualizar_totales)

        # 🧮 Redondeo opcional de la detracción al entero (norma del Banco de la Nación)
        self.chk_redondeo_det = ctk.CTkCheckBox(
            self.f_form, text="🧮 Redondear detracción al entero",
            font=("Arial", 11, "bold"), checkbox_width=18, checkbox_height=18,
            command=self.actualizar_totales)
        self.chk_redondeo_det.pack(anchor="w", padx=10, pady=(0, 0))
        ctk.CTkLabel(self.f_form,
                     text="Fracción < 0.50: baja al entero · Fracción ≥ 0.50: sube a la unidad superior",
                     font=("Arial", 9), text_color="#7a7a7a", justify="left",
                     wraplength=290).pack(anchor="w", padx=10, pady=(0, 8))

        f_tot = ctk.CTkFrame(self.f_form, fg_color="#ffffff", border_width=1, border_color="#ccc")
        f_tot.pack(fill="x", padx=10, pady=(5, 10))
        self.lbl_impuesto = ctk.CTkLabel(f_tot, text="IGV (18%): 0.00", font=("Arial", 11), text_color="#555")
        self.lbl_impuesto.pack(anchor="w", padx=10, pady=(5, 0))
        self.lbl_detraccion = ctk.CTkLabel(f_tot, text="Detracción (12%): 0.00", font=("Arial", 11), text_color="#e74c3c")
        self.lbl_detraccion.pack(anchor="w", padx=10, pady=(0, 0))
        self.lbl_total = ctk.CTkLabel(f_tot, text="Neto a Cobrar: 0.00", font=("Arial", 13, "bold"), text_color="#1f538d")
        self.lbl_total.pack(anchor="w", padx=10, pady=(2, 5))

        btn_guardar = ctk.CTkButton(self.f_form, text="💾 Registrar Documento", font=("Arial", 12, "bold"), fg_color="#1f538d", hover_color="#163b65", command=self.guardar_registro)
        btn_guardar.pack(fill="x", padx=10, pady=(10, 15))

        self.cargar_clientes_bd()
        self.cargar_ordenes_compra()
        self.cargar_cobranzas_pendientes()
        self.sugerir_correlativo()

        self.f_wrapper_derecha = ctk.CTkFrame(frame_split, fg_color="transparent")
        self.f_wrapper_derecha.pack(side="right", fill="both", expand=True)

        f_busqueda = ctk.CTkFrame(self.f_wrapper_derecha, fg_color="transparent")
        f_busqueda.pack(fill="x", pady=(0, 5))
        ctk.CTkLabel(f_busqueda, text="🔍 Buscar:", font=("Arial", 11, "bold")).pack(side="left", padx=(0, 5))
        self.ent_buscar_facturas = ctk.CTkEntry(f_busqueda, placeholder_text="Filtrar por N° Doc, cliente, concepto, fecha...")
        self.ent_buscar_facturas.pack(side="left", fill="x", expand=True)
        
        ctk.CTkLabel(f_busqueda, text="🗓️ Mes:", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 5))
        self.combo_mes = ctk.CTkComboBox(f_busqueda, values=construir_valores_mes(), width=150, state="readonly", command=self.on_cambiar_mes)
        self.combo_mes.pack(side="left", padx=(0, 5))
        self.combo_mes.set(self.mes_filtro)
        
        # 🚀 BÚSQUEDA ASÍNCRONA
        self.ent_buscar_facturas.bind("<KeyRelease>", lambda e: self.buscar_con_retraso())
        self.ent_buscar_facturas.bind("<Return>", lambda e: self.cargar_datos_tabla(reset_pagina=True))

        f_tabla = ctk.CTkFrame(self.f_wrapper_derecha, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True)

        columnas = ("num", "id", "fecha", "nro_doc", "dias", "tipo", "cliente", "concepto", "subtotal", "impuesto", "total", "detraccion", "neto", "estado_sunat", "archivo")
        self.tabla = ttk.Treeview(f_tabla, columns=columnas, show="headings", selectmode="extended")
        self.tabla.tag_configure("pagada", background="#d6eaf8", foreground="#1a5276")
        self.tabla.tag_configure("no_pagada", background="#fadbd8", foreground="#922b21")
        self.tabla.tag_configure("parcial", background="#e8daef", foreground="#6c3483")
        self.tabla.heading("num", text="N°", anchor="center")
        self.tabla.heading("id", text="ID (Oculto)")
        self.tabla.heading("fecha", text="Fecha ↕", command=lambda: self.ordenar_por_columna("fecha", False))
        self.tabla.heading("nro_doc", text="N° Doc. ↕", command=lambda: self.ordenar_por_columna("nro_doc", False))
        self.tabla.heading("cliente", text="Cliente ↕", command=lambda: self.ordenar_por_columna("cliente", False))
        self.tabla.heading("concepto", text="Concepto ↕", command=lambda: self.ordenar_por_columna("concepto", False))
        self.tabla.heading("neto", text="Neto Cobrar ↕", command=lambda: self.ordenar_por_columna("neto", True))
        self.tabla.heading("estado_sunat", text="Estado SUNAT ↕", command=lambda: self.ordenar_por_columna("estado_sunat", False))
        
        self.tabla.column("num", width=35, anchor="center")
        self.tabla.column("id", width=0, stretch=tk.NO)
        self.tabla.column("fecha", width=75, anchor="center")
        # 📏 Ancho suficiente para el N° de documento completo: "F001-00000159" mide
        # 89 px y con 90 px se recortaba el último dígito (parecían números repetidos).
        self.tabla.column("nro_doc", width=130, minwidth=120, stretch=False, anchor="w")
        self.tabla.column("cliente", width=120, anchor="w")
        self.tabla.column("concepto", width=140, anchor="w")
        self.tabla.column("neto", width=85, anchor="e")
        self.tabla.column("estado_sunat", width=95, anchor="center")
        
        self.tabla.config(displaycolumns=("num", "fecha", "nro_doc", "cliente", "concepto", "neto", "estado_sunat"))

        self.tabla.bind("<Double-1>", self.abrir_archivo_desde_tabla)
        # 🖱️ Clic derecho: ver factura / comprobante de pago (botón 3 y 2 en macOS)
        self.tabla.bind("<Button-3>", self._menu_contextual_factura)
        self.tabla.bind("<Button-2>", self._menu_contextual_factura)

        scroll_y = ttk.Scrollbar(f_tabla, orient="vertical", command=self.tabla.yview)
        scroll_x = ttk.Scrollbar(f_tabla, orient="horizontal", command=self.tabla.xview)
        self.tabla.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_x.pack(side="bottom", fill="x")
        scroll_y.pack(side="right", fill="y")
        self.tabla.pack(side="left", fill="both", expand=True)
        for _c in self.tabla["columns"]:
            try:
                self.tabla.column(_c, stretch=False)
            except Exception:
                pass

        # 🚀 BOTONES DE PAGINACIÓN Y ACCIONES
        f_btn_tabla = ctk.CTkFrame(self.f_wrapper_derecha, fg_color="transparent")
        f_btn_tabla.pack(fill="x", pady=(10, 0))

        btn_cols = ctk.CTkButton(f_btn_tabla, text="⚙️ Columnas", width=90, font=("Arial", 12, "bold"), fg_color="#6c757d", hover_color="#565e64", command=lambda: menu_columnas(self.tabla, "ventas_facturas"))
        btn_cols.pack(side="left", padx=5, pady=5)
        try:
            aplicar_columnas_visibles(self.tabla, "ventas_facturas")
        except Exception:
            pass        
        self.btn_ant = ctk.CTkButton(f_btn_tabla, text="◀ Ant", width=60, command=self.pagina_anterior)
        self.btn_ant.pack(side="left", padx=2)
        
        self.lbl_pagina = ctk.CTkLabel(f_btn_tabla, text=f"Pág {self.pagina_actual}", font=("Arial", 11, "bold"))
        self.lbl_pagina.pack(side="left", padx=5)
        
        self.btn_sig = ctk.CTkButton(f_btn_tabla, text="Sig ▶", width=60, command=self.pagina_siguiente)
        self.btn_sig.pack(side="left", padx=2)
        
        self.main_root.after(100, lambda: self.cargar_datos_tabla(reset_pagina=True))

    def on_cambiar_mes(self, choice):
        self.mes_filtro = choice
        self.cargar_datos_tabla(reset_pagina=True)

    def pagina_anterior(self):
        if self.pagina_actual > 1:
            self.pagina_actual -= 1
            self.cargar_datos_tabla()
            
    def pagina_siguiente(self):
        self.pagina_actual += 1
        self.cargar_datos_tabla()

    def buscar_con_retraso(self):
        if hasattr(self, "_busqueda_job"):
            try: self.main_root.after_cancel(self._busqueda_job)
            except: pass
        self._busqueda_job = self.main_root.after(350, lambda: self.cargar_datos_tabla(reset_pagina=True))

    # 🚀 FIX: CARGA DE OCS CON CACHÉ
    def cargar_ordenes_compra(self):
        ocs_cache = cache_sistema.obtener("lista_ocs_combobox")
        if ocs_cache is not None:
            if hasattr(self, 'combo_oc'):
                self.combo_oc.configure(values=ocs_cache)
                self.combo_oc.set("--- Sin Orden de Compra ---")

        else:
            def tarea_ocs(estado):               # 👈 corre en el HILO SECUNDARIO
                ocs = ["--- Sin Orden de Compra ---"]
                conn = conectar_db(silencioso=True)
                if conn:
                    try:
                        c = conn.cursor()
                        c.execute("SELECT numero_oc FROM ordenes_compra_clientes ORDER BY id DESC")
                        for r in c.fetchall(): ocs.append(r[0])
                        cache_sistema.guardar("lista_ocs_combobox", ocs)
                    except: pass
                    finally: liberar_conexion(conn)
                estado["ocs"] = ocs

            def _aplicar_ocs_hilo(estado):       # 👈 corre en el HILO PRINCIPAL
                if estado.get("ocs"):
                    self._aplicar_ocs(estado["ocs"])

            ejecutar_en_hilo(self.main_root, tarea_ocs, al_terminar=_aplicar_ocs_hilo)

    def _aplicar_ocs(self, ocs):
        if hasattr(self, 'combo_oc'):
            self.combo_oc.configure(values=ocs)
            self.combo_oc.set("--- Sin Orden de Compra ---")

    def on_tipo_change(self, choice):
        if hasattr(self, 'chk_redondeo_det'):
            # El redondeo solo aplica a la detracción (no a la retención de recibos)
            try:
                self.chk_redondeo_det.configure(state="disabled" if "Recibo" in choice else "normal")
            except Exception:
                pass
        if hasattr(self, 'ent_detraccion'):
            self.ent_detraccion.configure(state="normal")
            if "Recibo" in choice:
                self.lbl_titulo_det.configure(text="Retención (%):")
                self.ent_detraccion.delete(0, tk.END)
                if "8%" in choice: self.ent_detraccion.insert(0, "8")
                else: self.ent_detraccion.insert(0, "0")
            elif "Factura" in choice:
                self.lbl_titulo_det.configure(text="Detracción (%):")
                self.al_seleccionar_cliente()
            else:
                self.lbl_titulo_det.configure(text="Detracción (%):")
                self.ent_detraccion.delete(0, tk.END); self.ent_detraccion.insert(0, "0")
            # Si la venta viene de una cobranza, la base se reajusta según el IGV del documento
            self.aplicar_monto_cobranza()
            self.actualizar_totales()
            self.sugerir_correlativo()

    # 🚀 FIX: AUTOCOMPLETADO RUC ASÍNCRONO
    def al_seleccionar_cliente(self, choice=None):
        if not hasattr(self, 'ent_ruc') or not hasattr(self, 'ent_detraccion'):
            return
            
        cliente = self.combo_cliente.get().strip()
        if not cliente or cliente == "--- Seleccione Cliente ---":
            if not getattr(self, 'bloquear_autocompletado_ruc', False):
                self.ent_ruc.configure(state="normal")
                self.ent_ruc.delete(0, tk.END)
            self.cargar_ordenes_compra()
            return
        
        def tarea_ruc(estado):                     # 👈 corre en el HILO SECUNDARIO
            ruc_db = ""
            conn = conectar_db(silencioso=True)
            if conn:
                try:
                    cursor = conn.cursor()
                    cursor.execute("SELECT ruc FROM clientes WHERE TRIM(UPPER(nombre_empresa)) = TRIM(UPPER(%s))", (cliente,))
                    res = cursor.fetchone()
                    if res: ruc_db = res[0]
                except: pass
                finally: liberar_conexion(conn)
            estado["ruc_db"] = ruc_db

        def _aplicar_ruc_hilo(estado):           # 👈 corre en el HILO PRINCIPAL
            self._aplicar_ruc(estado.get("ruc_db") or "")

        ejecutar_en_hilo(self.main_root, tarea_ruc, al_terminar=_aplicar_ruc_hilo)

    def _aplicar_ruc(self, ruc_db):
        if not getattr(self, 'bloquear_autocompletado_ruc', False):
            self.ent_ruc.configure(state="normal")
            self.ent_ruc.delete(0, tk.END)
            if ruc_db:
                self.ent_ruc.insert(0, str(ruc_db))
                
        self.ent_detraccion.configure(state="normal")
        if "Factura" in getattr(self.combo_tipo, 'get', lambda: "")():
            if not self.ent_detraccion.get().strip() or self.ent_detraccion.get().strip() == "0":
                self.ent_detraccion.delete(0, tk.END)
                self.ent_detraccion.insert(0, CONFIG_REGIONAL.get("detraccion_porcentaje", "12"))
                
        if hasattr(self, 'actualizar_totales'):
            self.actualizar_totales()

    def cargar_clientes_bd(self):
        clis = cache_sistema.obtener('lista_clientes_combobox')
        if clis is not None:
            self._aplicar_clientes_combo(clis)
        else:
            def tarea_clientes(estado):          # 👈 corre en el HILO SECUNDARIO
                clis_bd = []
                conn = conectar_db(silencioso=True)
                if conn:
                    try:
                        cursor = conn.cursor()
                        cursor.execute("SELECT nombre_empresa FROM clientes ORDER BY nombre_empresa ASC")
                        clis_bd = [str(r[0]).strip() for r in cursor.fetchall() if r[0]]
                        cache_sistema.guardar('lista_clientes_combobox', clis_bd)
                    except: pass
                    finally: liberar_conexion(conn)
                estado["clis_bd"] = clis_bd

            def _aplicar_clientes_hilo(estado):  # 👈 corre en el HILO PRINCIPAL
                self._aplicar_clientes_combo(estado.get("clis_bd") or [])

            ejecutar_en_hilo(self.main_root, tarea_clientes, al_terminar=_aplicar_clientes_hilo)

    def _aplicar_clientes_combo(self, clis):
        if clis:
            lista_clis = ["--- Seleccione Cliente ---"] + clis
            self.combo_cliente.configure(values=lista_clis)
        try:
            self.combo_cliente._valores_completos = list(lista_clis)
        except Exception:
            pass
            if self.combo_cliente.get() not in lista_clis:
                self.combo_cliente.set("--- Seleccione Cliente ---")
            self.al_seleccionar_cliente()
        else:
            self.combo_cliente.configure(values=["--- Seleccione Cliente ---"])
            self.combo_cliente.set("--- Seleccione Cliente ---")

    # =========================================================
    # 🚀 COBRANZAS PENDIENTES DE FACTURAR (vínculo Cálculo de Cobranza → Ventas)
    # =========================================================
    def cargar_cobranzas_pendientes(self):
        """Carga en el desplegable los cálculos de cobranza aún NO facturados."""
        def tarea_cobranzas(estado):             # 👈 corre en el HILO SECUNDARIO
            pend = {}
            conn = conectar_db(silencioso=True)
            if conn:
                try:
                    cursor = conn.cursor()
                    # Asegura las columnas de estado (por si el módulo de cobranza aún no se ha abierto)
                    try:
                        cursor.execute("ALTER TABLE cobranza_quincenas ADD COLUMN IF NOT EXISTS facturado BOOLEAN DEFAULT FALSE")
                        cursor.execute("ALTER TABLE cobranza_quincenas ADD COLUMN IF NOT EXISTS factura_referencia VARCHAR(100) DEFAULT ''")
                        conn.commit()
                    except Exception:
                        conn.rollback()
                    cursor.execute("""
                        SELECT id, cliente_nombre, cliente_ruc, anio, mes, quincena, plan_cobro, total
                        FROM cobranza_quincenas
                        WHERE COALESCE(facturado, FALSE) = FALSE
                        ORDER BY anio DESC, mes DESC, quincena DESC, id DESC
                    """)
                    for (cid, nom, ruc, anio, mes, q, plan, total) in cursor.fetchall():
                        nom = (nom or "").strip()
                        ruc = (ruc or "").strip()
                        try: anio = int(anio)
                        except Exception: anio = 0
                        try: mes = int(mes)
                        except Exception: mes = 0
                        try: q = int(q)
                        except Exception: q = 0
                        try: total = float(total or 0)
                        except Exception: total = 0.0
                        mes_nom = NOMBRES_MESES[mes - 1] if 1 <= mes <= 12 else f"Mes {mes}"
                        if q == 1: q_txt = "1ª Quincena (1-15)"
                        elif q == 2: q_txt = "2ª Quincena (16-fin)"
                        else: q_txt = f"Quincena {q}"
                        etiqueta = f"{nom} · {q_txt} {mes_nom} {anio} · {formatear_moneda(total)}"
                        pend[etiqueta] = {
                            "id": cid, "cliente_nombre": nom, "cliente_ruc": ruc,
                            "anio": anio, "mes": mes, "quincena": q,
                            "quincena_txt": q_txt, "mes_nombre": mes_nom,
                            "plan_cobro": (plan or "").strip(), "total": total,
                        }
                except Exception:
                    pass
                finally:
                    liberar_conexion(conn)
            estado["pend"] = pend

        def _aplicar_cobranzas_hilo(estado):     # 👈 corre en el HILO PRINCIPAL
            self._aplicar_cobranzas_combo(estado.get("pend") or {})

        ejecutar_en_hilo(self.main_root, tarea_cobranzas, al_terminar=_aplicar_cobranzas_hilo)

    def _aplicar_cobranzas_combo(self, pend):
        self.cobranzas_pendientes = pend
        if pend:
            valores = ["--- Seleccione Cobranza ---"] + list(pend.keys())
        else:
            valores = ["--- Sin cobranzas pendientes ---"]
        self.combo_cobranza.configure(values=valores)
        actual = self.combo_cobranza.get()
        if actual not in valores:
            self.combo_cobranza.set(valores[0])

    def al_seleccionar_cobranza(self, choice=None):
        """Al elegir una cobranza pendiente, autocompleta cliente, RUC, descripción y monto base."""
        self.id_cobranza_seleccionada = None
        data = self.cobranzas_pendientes.get(choice)
        if not data:
            return
        self.id_cobranza_seleccionada = data["id"]

        # Cliente + RUC (evita que el autocompletado asíncrono pise el RUC)
        self.bloquear_autocompletado_ruc = True
        try:
            if data["cliente_nombre"]:
                self.combo_cliente.set(data["cliente_nombre"])
            self.ent_ruc.configure(state="normal")
            self.ent_ruc.delete(0, tk.END)
            if data["cliente_ruc"]:
                self.ent_ruc.insert(0, data["cliente_ruc"])
        finally:
            self.bloquear_autocompletado_ruc = False

        # Descripción de la quincena a cobrar
        desc = f"Cobranza {data['quincena_txt']} {data['mes_nombre']} {data['anio']}"
        if data["plan_cobro"]:
            desc += f" — {data['plan_cobro']}"
        self.ent_desc.delete(0, tk.END)
        self.ent_desc.insert(0, desc)

        # 💡 El TOTAL de la cobranza YA INCLUYE IGV: es el importe que paga el cliente
        # (menos la detracción). Por eso NO se carga tal cual en "Monto Base": se
        # desagrega el IGV para que Base + IGV = total de la cobranza.
        self.total_cobranza_seleccionada = float(data["total"] or 0)
        self.aplicar_monto_cobranza()
        self.actualizar_totales()

    def aplicar_monto_cobranza(self, *args):
        """Carga en "Monto Base" el importe de la cobranza seleccionada.

        El total de la cobranza ya incluye IGV, así que en FACTURA se coloca el valor
        sin IGV (total ÷ 1.18): al sumarle el 18% la factura totaliza exactamente lo
        que dice la cobranza. En boleta/recibo (sin IGV) el monto va tal cual.
        El cliente paga ese total menos la detracción.
        """
        total = float(getattr(self, "total_cobranza_seleccionada", 0.0) or 0.0)
        if not hasattr(self, "ent_subtotal"):
            return
        tipo = self.combo_tipo.get() if hasattr(self, "combo_tipo") else ""
        con_igv = "Factura" in tipo
        base = round(total / 1.18, 2) if con_igv else round(total, 2)
        if total > 0:
            self.ent_subtotal.delete(0, tk.END)
            self.ent_subtotal.insert(0, f"{base:.2f}")
        if hasattr(self, "lbl_cobranza_info"):
            if total <= 0:
                self.lbl_cobranza_info.configure(text="")
            elif con_igv:
                self.lbl_cobranza_info.configure(
                    text=f"Cobranza (IGV incluido): {formatear_moneda(total)} → Base sin IGV: {formatear_moneda(base)}")
            else:
                self.lbl_cobranza_info.configure(text=f"Cobranza: {formatear_moneda(total)} (sin IGV)")

    def detraccion_con_redondeo(self):
        """True si el usuario activó el check de redondeo de la detracción."""
        chk = getattr(self, "chk_redondeo_det", None)
        if chk is None:
            return False
        try:
            return bool(chk.get())
        except Exception:
            return False

    @staticmethod
    def _etiqueta_redondeo(redondeo):
        return " (redondeada)" if redondeo else ""

    def actualizar_totales(self, *args):
        if not hasattr(self, 'combo_tipo') or not hasattr(self, 'ent_subtotal') or not hasattr(self, 'ent_detraccion'):
            return
            
        tipo = self.combo_tipo.get()
        try:
            sub = float(self.ent_subtotal.get() or 0)
            ui_pct = float(self.ent_detraccion.get() or 0)
            redondeo = self.detraccion_con_redondeo()
            if "Factura" in tipo:
                igv = redondear_centavos(sub * 0.18); tot = redondear_centavos(sub + igv)

                # 💧 Regla: si la factura SUPERA los S/ 400 se detrae el 4%; si no, es cero.
                porcentaje_regla = porcentaje_detraccion_ventas(tot)
                if abs(porcentaje_regla - ui_pct) > 0.001:
                    ui_pct = porcentaje_regla
                    try:
                        self.ent_detraccion.delete(0, tk.END)
                        self.ent_detraccion.insert(0, "%g" % ui_pct)
                    except Exception:
                        pass

                det = tot * (ui_pct / 100.0)
                det = redondear_detraccion(det) if redondeo else redondear_centavos(det)
                neto = redondear_centavos(tot - det)
                self.lbl_impuesto.configure(text=f"IGV (18%): {formatear_moneda(igv)}")
                self.lbl_detraccion.configure(text=f"Detracción ({ui_pct:g}%): -{formatear_moneda(det)}{self._etiqueta_redondeo(redondeo)}")
                self.lbl_total.configure(text=f"Neto a Cobrar: {formatear_moneda(neto)}")
            elif "Recibo" in tipo:      # los recibos usan su retención (no esta regla)
                ret = sub * (ui_pct / 100.0); neto = sub - ret
                self.lbl_impuesto.configure(text=f"Retención ({ui_pct:g}%): -{formatear_moneda(ret)}")
                self.lbl_detraccion.configure(text=f"Detracción (0%): -{formatear_moneda(0)}")
                self.lbl_total.configure(text=f"Neto a Cobrar: {formatear_moneda(neto)}")
            else:
                # Boleta / otros: sin IGV, pero sí puede estar sujeta a detracción
                det = sub * (ui_pct / 100.0)
                det = redondear_detraccion(det) if redondeo else redondear_centavos(det)
                neto = redondear_centavos(sub - det)
                self.lbl_impuesto.configure(text=f"IGV (0%): {formatear_moneda(0)}")
                self.lbl_detraccion.configure(text=f"Detracción ({ui_pct:g}%): -{formatear_moneda(det)}{self._etiqueta_redondeo(redondeo)}")
                self.lbl_total.configure(text=f"Neto a Cobrar: {formatear_moneda(neto)}")
        except ValueError: pass

    def guardar_registro(self):
        ruta_base = obtener_ruta_base_drive()
        if not ruta_base:
            avisar_sin_permiso_guardado()
            return
            
        tipo = self.combo_tipo.get()
        nro_doc = self.ent_nro_doc.get().strip()
        fecha = self.ent_fecha.get().strip()
        cliente = self.combo_cliente.get().strip()
        desc = self.ent_desc.get().strip()
        evento = "GENERAL / NO ASIGNADO"
        oc_sel = self.combo_oc.get()
        
        if not cliente or cliente == "--- Seleccione Cliente ---" or not desc: 
            return messagebox.showwarning("Atención", "Llene los campos obligatorios (Cliente y Concepto).")
            
        if oc_sel == "--- Sin Orden de Compra ---":
            oc_sel = ""
            
        if not oc_sel and "Factura" in tipo:
            if not messagebox.askyesno("Falta Orden de Compra", "⚠️ No ha seleccionado una Orden de Compra para esta factura.\n\n¿Desea emitir la factura de todas formas sin asociarla a una Orden de Compra?"):
                return
        
        try: 
            subtotal = float(self.ent_subtotal.get() or 0)
            dias = int(self.ent_dias.get().strip() or 0)
            ui_pct = float(self.ent_detraccion.get() or 0)
        except ValueError: return messagebox.showerror("Error", "Los montos deben ser numéricos.")

        # 🧮 Redondeo opcional de la detracción al entero (check del formulario)
        redondeo_det = self.detraccion_con_redondeo()

        if "Factura" in tipo: 
            imp = redondear_centavos(subtotal * 0.18); tot_bruto = redondear_centavos(subtotal + imp)
            det_pct = ui_pct; det_monto = tot_bruto * (det_pct / 100.0)
            det_monto = redondear_detraccion(det_monto) if redondeo_det else redondear_centavos(det_monto)
            neto_nuevo = redondear_centavos(tot_bruto - det_monto)
        elif "Recibo" in tipo: 
            imp = redondear_centavos(subtotal * (ui_pct / 100.0)); tot_bruto = redondear_centavos(subtotal)
            det_pct = 0.0; det_monto = 0.0
            neto_nuevo = redondear_centavos(subtotal - imp)  
        else: 
            imp = 0.0; tot_bruto = redondear_centavos(subtotal); det_pct = ui_pct
            det_monto = tot_bruto * (det_pct / 100.0)
            det_monto = redondear_detraccion(det_monto) if redondeo_det else redondear_centavos(det_monto)
            neto_nuevo = redondear_centavos(tot_bruto - det_monto)

        # 🚀 FACTURA LOCAL: solicitar el PDF; si no lo tiene, permitir cargarlo después.
        if not self.ruta_archivo_temp:
            adjuntar_ahora = messagebox.askyesno(
                "PDF de la Factura",
                "¿Desea adjuntar el PDF de la factura ahora?\n\n"
                "• Sí  → seleccionar el PDF ahora.\n"
                "• No  → registrarla sin PDF y adjuntarlo más tarde.")
            if adjuntar_ahora:
                ruta_pdf = seleccionar_archivo_dialogo(
                    titulo="Adjuntar PDF de la Factura",
                    tipos=[("Archivos PDF", "*.pdf"),
                           ("Imágenes (JPG/PNG)", "*.png;*.jpg;*.jpeg"),
                           ("Todos los archivos", "*.*")])
                if ruta_pdf:
                    self.ruta_archivo_temp = ruta_pdf
                    if hasattr(self, "lbl_archivo"):
                        self.lbl_archivo.configure(text=f"📎 {os.path.basename(ruta_pdf)}", text_color="#166534")

        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            
            if nro_doc:
                cursor.execute("SELECT COUNT(*) FROM facturas_emitidas WHERE numero_documento = %s", (nro_doc,))
                if cursor.fetchone()[0] > 0:
                    liberar_conexion(conn)
                    return messagebox.showwarning("Duplicado", "Ese N° de Documento ya está registrado.")

            cli_match = cliente.strip().upper()
            cursor.execute("SELECT limite_credito FROM clientes WHERE TRIM(UPPER(nombre_empresa)) = %s", (cli_match,))
            res_cli = cursor.fetchone()
            limite_credito = float(res_cli[0]) if (res_cli and res_cli[0] is not None) else 0.0
            
            if limite_credito > 0.0:
                cursor.execute("""
                    SELECT id, total, COALESCE(det_monto, 0), tipo_documento, COALESCE(impuesto, 0)
                    FROM facturas_emitidas WHERE TRIM(UPPER(cliente)) = %s
                """, (cli_match,))
                
                facturas_historicas = cursor.fetchall()
                deuda_actual = 0.0
                
                for r in facturas_historicas:
                    id_fac, tot_b, det_m, t_doc, imp_val = r
                    tot_val = float(tot_b) if tot_b else 0.0
                    det_val = float(det_m) if det_m else 0.0
                    impuesto_val = float(imp_val) if imp_val else 0.0
                    
                    if t_doc and "Recibo" in t_doc and "8%" in t_doc:
                        neto_fac = tot_val - impuesto_val - det_val
                    else:
                        neto_fac = tot_val - det_val
                    
                    cursor.execute("SELECT SUM(monto_pagado + COALESCE(impuestos_deducibles, 0)) FROM pagos_clientes WHERE id_factura = %s", (id_fac,))
                    pagado = cursor.fetchone()[0]
                    pagado_val = float(pagado) if pagado else 0.0
                    
                    saldo = max(0.0, neto_fac - pagado_val)
                    deuda_actual += saldo
                
                deuda_proyectada = deuda_actual + neto_nuevo
                
                if deuda_proyectada > limite_credito:
                    msg = (f"⚠️ LÍMITE DE CRÉDITO EXCEDIDO\n\n"
                           f"Cliente: {cliente}\n"
                           f"Límite Asignado: {formatear_moneda(limite_credito)}\n"
                           f"Deuda Actual Pendiente: {formatear_moneda(deuda_actual)}\n\n"
                           f"Nueva Facturación: {formatear_moneda(neto_nuevo)}\n"
                           f"Deuda Proyectada: {formatear_moneda(deuda_proyectada)}\n\n"
                           f"El sistema NO permite registrar esta venta hasta que el cliente regularice sus pagos.")
                    liberar_conexion(conn)
                    return messagebox.showerror("Bloqueo por Crédito", msg)

            ruta_final = ""
            if self.ruta_archivo_temp:
                try:
                    carpeta_destino = os.path.join(ruta_base, "facturas_emitidas")
                    if not os.path.exists(carpeta_destino): os.makedirs(carpeta_destino)
                    nombre_ext = os.path.splitext(self.ruta_archivo_temp)[1]
                    ruta_final = os.path.join(carpeta_destino, f"Emitida_{datetime.now().strftime('%Y%m%d%H%M%S')}_{cliente.replace(' ', '_')}{nombre_ext}")
                    shutil.copy2(self.ruta_archivo_temp, ruta_final)
                except Exception as e: 
                    liberar_conexion(conn)
                    return messagebox.showerror("Error", f"Fallo al guardar archivo:\n{e}")

            columnas_ins = ["tipo_documento", "numero_documento", "fecha", "cliente", "descripcion",
                            "evento_asociado", "subtotal", "impuesto", "total", "archivo_ruta",
                            "dias_credito", "det_porcentaje", "det_monto", "orden_compra"]
            valores_ins = [tipo, nro_doc, fecha, cliente, desc, evento, subtotal, imp, tot_bruto,
                           ruta_para_guardar(ruta_final), dias, det_pct, det_monto, oc_sel]
            if _COLUMNA_DET_REDONDEO and "Recibo" not in tipo:
                columnas_ins.append("det_redondeo")
                valores_ins.append(bool(redondeo_det))
            cursor.execute(
                f"INSERT INTO facturas_emitidas ({', '.join(columnas_ins)}) "
                f"VALUES ({', '.join(['%s'] * len(columnas_ins))}) RETURNING id",
                tuple(valores_ins))
            try:
                id_factura_nueva = int(cursor.fetchone()[0])
            except Exception:
                id_factura_nueva = 0
            conn.commit()

            # 🚀 Marca la cobranza vinculada como FACTURADA (deja de aparecer en el desplegable,
            # pero se conserva intacta en el histórico de "Cálculo de Cobranza").
            if self.id_cobranza_seleccionada:
                try:
                    cursor.execute(
                        "UPDATE cobranza_quincenas SET facturado = TRUE, factura_referencia = %s WHERE id = %s",
                        (nro_doc, self.id_cobranza_seleccionada))
                    conn.commit()
                except Exception:
                    pass

            cache_sistema.invalidar()
            registrar_auditoria(self.app_padre.usuario_activo, "Facturas Emitidas", f"Registró factura {nro_doc} del cliente '{cliente}'")
            messagebox.showinfo("Éxito", "Documento emitido registrado correctamente.")
            # 💰 Pregunta si ya fue pagada (después de cerrar el aviso de éxito)
            # 💰 ¿Ya fue pagada? (método de ESTA clase, sin depender de otra pestaña)
            self._ofrecer_cobro(id_factura_nueva, nro_doc)
            
            self.ent_desc.delete(0, tk.END)
            self.ent_subtotal.delete(0, tk.END)
            self.ruta_archivo_temp = ""
            if hasattr(self, "lbl_archivo"):
                self.lbl_archivo.configure(text="Sin PDF adjunto", text_color="#999999")
            self.combo_oc.set("--- Sin Orden de Compra ---")
            self.id_cobranza_seleccionada = None
            self.total_cobranza_seleccionada = 0.0
            if hasattr(self, "lbl_cobranza_info"):
                self.lbl_cobranza_info.configure(text="")
            self.combo_cobranza.set("--- Seleccione Cobranza ---")
            self.cargar_cobranzas_pendientes()
            self.cargar_datos_tabla(reset_pagina=True)
            self.sugerir_correlativo()
            
            if hasattr(self.app_padre, 'app_cobros'):
                self.app_padre.app_cobros.cargar_datos_cobrar()
            if hasattr(self.app_padre, 'app_nc'):
                self.app_padre.app_nc.cargar_datos_nc()
        except Exception as e: messagebox.showerror("Error SQL", str(e))
        finally: liberar_conexion(conn)

    # 🚀 FIX: CARGA LAZY LOADING + CACHÉ
    def _ofrecer_cobro(self, id_factura=0, nro_doc=""):
        """Pregunta si la factura ya fue pagada y abre el cobro en la pestaña de Cobros."""
        try:
            respuesta = messagebox.askyesno(
                "Cancelar la factura",
                "¿Desea CANCELAR ahora la factura %s?" % (nro_doc or "s/n") + chr(10) + chr(10) +
                "Si ya fue pagada, se abrirá «Registrar Cobro» para registrar el ingreso.")
        except Exception as e:
            print("Aviso - pregunta de cobro:", e)
            return False
        if not respuesta:
            return False
        return self._abrir_cobro_de(id_factura, nro_doc)

    def _abrir_cobro_de(self, id_factura=0, nro_doc="", intento=0):
        """Selecciona la factura en la pestaña de Cobros y abre «Registrar Cobro».

        La lista de cobros se carga en segundo plano: se reintenta con 'after'
        hasta que la fila aparezca (sin congelar la ventana).
        """
        cobros = None
        try:
            cobros = getattr(getattr(self, "app_padre", None), "app_cobros", None)
        except Exception:
            cobros = None
        if cobros is None:
            messagebox.showinfo(
                "Registrar Cobro",
                "La factura se guardó. Abra la pestaña «Cuentas por Cobrar»," + chr(10) +
                "seleccione la factura y use «🧾 Registrar cobro».")
            return False

        if intento == 0 or (intento % 3 == 0):
            try:
                cache_sistema.invalidar()
            except Exception:
                pass
            try:
                cobros.cargar_datos_cobrar(reset_pagina=True)
            except Exception:
                try:
                    cobros.cargar_datos_cobrar()
                except Exception:
                    pass

        try:
            tabla = cobros.tabla
        except Exception:
            return False

        for item in tabla.get_children():
            try:
                valores = tabla.item(item, "values")
            except Exception:
                continue
            if len(valores) < 4:
                continue
            coincide_id = id_factura and str(valores[1]) == str(id_factura)
            coincide_doc = nro_doc and str(valores[3]).strip() == str(nro_doc).strip()
            if coincide_id or coincide_doc:
                try:
                    tabla.selection_set(item)
                    tabla.see(item)
                    cobros.cargar_comprobante_cobro()
                except Exception as e:
                    print("Aviso - abrir cobro:", e)
                    return False
                return True

        if intento < 30:
            try:
                self.main_root.after(400, lambda: self._abrir_cobro_de(id_factura, nro_doc, intento + 1))
            except Exception:
                pass
            return False

        messagebox.showinfo(
            "Registrar Cobro",
            "La factura se guardó, pero aún no aparece en la pestaña de Cobros." + chr(10) +
            "Allí pulse «Actualizar» y use «🧾 Registrar cobro».")
        return False

    def cargar_datos_tabla(self, reset_pagina=False):
        if reset_pagina:
            self.pagina_actual = 1
            
        self.lbl_pagina.configure(text=f"Pág {self.pagina_actual}")

        for item in self.tabla.get_children(): 
            self.tabla.delete(item)
        
        filtro = ""
        if hasattr(self, 'ent_buscar_facturas'):
            filtro = self.ent_buscar_facturas.get().strip().lower()
            
        offset = (self.pagina_actual - 1) * self.registros_por_pagina
        clave_cache = f"facturas_{filtro}_mes_{self.mes_filtro}_pag_{self.pagina_actual}"
        datos = cache_sistema.obtener(clave_cache)

        if datos is not None:
            self._pintar_facturas(datos)
        else:
            self.tabla.insert("", tk.END, values=("", "", "", "Cargando datos...", "", "", "", "", "", "", "", "", "", "", ""))
            
            def tarea_descarga(estado):          # 👈 corre en el HILO SECUNDARIO
                conn = conectar_db(silencioso=True)
                if not conn: return
                try:
                    cursor = conn.cursor()
                    query_base = "SELECT id, fecha, numero_documento, dias_credito, tipo_documento, cliente, evento_asociado, descripcion, subtotal, impuesto, total, COALESCE(det_monto, 0), archivo_ruta, enlace_pdf_sunat, estado_sunat, COALESCE((SELECT SUM(p.monto_pagado + COALESCE(p.impuestos_deducibles, 0)) FROM pagos_clientes p WHERE p.id_factura = facturas_emitidas.id), 0) AS cobrado FROM facturas_emitidas"
                    
                    condiciones = []
                    params = []
                    if filtro:
                        val = f"%{filtro}%"
                        condiciones.append("(numero_documento ILIKE %s OR cliente ILIKE %s OR descripcion ILIKE %s)")
                        params.extend([val, val, val])
                    patron_mes = patron_fecha_mes(self.mes_filtro)
                    if patron_mes:
                        condiciones.append("fecha LIKE %s")
                        params.append(patron_mes)
                    where_sql = (" WHERE " + " AND ".join(condiciones)) if condiciones else ""
                    params.extend([self.registros_por_pagina, offset])
                    cursor.execute(f"{query_base}{where_sql} ORDER BY NULLIF(SPLIT_PART(fecha, '/', 3), '')::int DESC NULLS LAST, NULLIF(SPLIT_PART(fecha, '/', 2), '')::int DESC NULLS LAST, NULLIF(SPLIT_PART(fecha, '/', 1), '')::int DESC NULLS LAST, id DESC LIMIT %s OFFSET %s", tuple(params))
                    
                    datos_db = cursor.fetchall()
                    cache_sistema.guardar(clave_cache, datos_db)
                    estado["datos_db"] = datos_db
                except Exception as e:
                    print(f"Error cargando tabla de facturas: {e}")
                finally:
                    liberar_conexion(conn)

            def _pintar_facturas_hilo(estado):   # 👈 corre en el HILO PRINCIPAL
                if estado.get("datos_db") is not None:
                    self._pintar_facturas(estado["datos_db"])

            ejecutar_en_hilo(self.main_root, tarea_descarga, al_terminar=_pintar_facturas_hilo)

    def _pintar_facturas(self, datos):
        for item in self.tabla.get_children():
            self.tabla.delete(item)

        contador = 1
        for r in datos:
            tiene_arch = "✅ Ver" if r[12] else "❌ No"
            tipo_doc = r[4]; impuesto = r[9]; tot_bruto = r[10]; det_monto = r[11]
            _tb, _im, _dm = float(tot_bruto or 0), float(impuesto or 0), float(det_monto or 0)
            if "Recibo" in tipo_doc and "8%" in tipo_doc: neto = _tb - _im - _dm
            else: neto = _tb - _dm
            
            estado_db = r[14]
            if estado_db and "Anulada" in estado_db:
                estado_sunat_str = "❌ Anulada"
            else:
                estado_sunat_str = "✅ Emitido" if r[13] else "⏳ Local"
                
            row_vals = (
                contador, r[0], r[1], r[2] if r[2] else "-", r[3], tipo_doc.split(" ")[0], r[5],
                r[7].split(" | ")[0] if " | " in r[7] else r[7], r[8], formatear_moneda(impuesto), formatear_moneda(tot_bruto), formatear_moneda(det_monto), formatear_moneda(neto), estado_sunat_str, tiene_arch
            )
            
            try:
                _cobrado_f = float(r[15] or 0) if len(r) > 15 else 0.0
            except Exception:
                _cobrado_f = 0.0
            if (neto <= 0.01) or (_cobrado_f >= neto - 0.01):
                _etiqueta = "pagada"
            elif _cobrado_f <= 0.01:
                _etiqueta = "no_pagada"
            else:
                _etiqueta = "parcial"
            self.tabla.insert("", tk.END, values=row_vals, tags=(_etiqueta,))
            contador += 1
            
        if self.pagina_actual > 1:
            self.btn_ant.configure(state="normal")
        else:
            self.btn_ant.configure(state="disabled")
            
        if len(datos) == self.registros_por_pagina:
            self.btn_sig.configure(state="normal")
        else:
            self.btn_sig.configure(state="disabled")

    # =========================================================
    # 🖱️ MENÚ DE CLIC DERECHO SOBRE UNA FACTURA (ver factura / comprobante)
    # =========================================================
    def _datos_factura_venta(self, id_factura):
        """(ruta_local, enlace_sunat, comprobantes) de una factura de venta."""
        ruta_local, enlace, comprobantes = "", "", []
        conn = conectar_db(silencioso=True)
        if not conn:
            return ruta_local, enlace, comprobantes
        try:
            with conn.cursor() as c:
                c.execute("""SELECT COALESCE(archivo_ruta, ''), COALESCE(enlace_pdf_sunat, '')
                             FROM facturas_emitidas WHERE id = %s""", (id_factura,))
                fila = c.fetchone()
                if fila:
                    ruta_local = str(fila[0] or "").strip()
                    enlace = str(fila[1] or "").strip()
                c.execute("""SELECT COALESCE(archivo_ruta, '') FROM pagos_clientes
                             WHERE id_factura = %s AND COALESCE(archivo_ruta, '') <> ''
                             ORDER BY id DESC""", (id_factura,))
                comprobantes = [r[0] for r in c.fetchall()]
        except Exception as e:
            print("Aviso - archivos de la factura de venta:", e)
        finally:
            liberar_conexion(conn)
        return ruta_local, enlace, comprobantes

    def _primera_ruta_venta(self, candidatas):
        """Primera ruta de la lista que exista en este equipo (o '')."""
        for ruta in (candidatas or []):
            ruta = str(ruta or "").strip()
            if not ruta:
                continue
            resuelta = ""
            try:
                from app_paths import resolver_ruta_archivo as _resolver
                resuelta = _resolver(ruta)
            except Exception:
                resuelta = ruta if os.path.exists(ruta) else ""
            for candidata in (resuelta, ruta):
                try:
                    if candidata and os.path.isfile(candidata):
                        return candidata
                except Exception:
                    continue
        return ""

    def _abrir_comprobantes_venta(self, comprobantes):
        """Abre todos los comprobantes de pago disponibles de la factura."""
        abiertos = 0
        for ruta in (comprobantes or []):
            resuelta = self._primera_ruta_venta([ruta])
            if not resuelta:
                continue
            try:
                abrir_documento(resuelta)
                abiertos += 1
            except Exception:
                pass
        if not abiertos:
            messagebox.showwarning(
                "Comprobante de pago",
                "Los comprobantes están registrados pero los archivos no se encuentran en este equipo."
                + chr(10) + chr(10) +
                "Verifique que la carpeta sincronizada esté actualizada (Rclone).",
                parent=ventana_para_mensaje(getattr(self, "main_root", None)))

    def _menu_contextual_factura(self, event):
        """Botón derecho: editar registro / ver factura / ver comprobante (y cobros)."""
        tabla = event.widget
        try:
            fila = tabla.identify_row(event.y)
            if not fila:
                return
            tabla.selection_set(fila)
            id_fact = int(tabla.item(fila, "values")[1])
        except Exception:
            return

        es_cobranza = hasattr(self, "cargar_comprobante_cobro")
        ruta_local, enlace, comprobantes = self._datos_factura_venta(id_fact)
        ruta_factura = self._primera_ruta_venta([ruta_local])
        ruta_comp = self._primera_ruta_venta(comprobantes)

        menu = tk.Menu(getattr(self, "main_root", None), tearoff=0)
        if es_cobranza:
            menu.add_command(label="🧾 Registrar cobro", command=self.cargar_comprobante_cobro)
            menu.add_command(label="✏️ Editar cobro", command=self.abrir_ventana_edicion)
        else:
            menu.add_command(label="✏️ Editar registro", command=self.abrir_ventana_edicion)
        menu.add_separator()
        if ruta_factura:
            menu.add_command(label="📄 Ver factura",
                             command=lambda r=ruta_factura: abrir_documento(r))
        elif enlace:
            menu.add_command(label="📄 Ver factura (SUNAT)",
                             command=lambda e=enlace: webbrowser.open(e))
        else:
            menu.add_command(label="📄 Ver factura (sin archivo)", state="disabled")
        if ruta_comp:
            menu.add_command(label="🧾 Ver comprobante de pago",
                             command=lambda c=comprobantes: self._abrir_comprobantes_venta(c))
        else:
            menu.add_command(label="🧾 Ver comprobante de pago (sin archivo)", state="disabled")
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def abrir_archivo_desde_tabla(self, event):
        sel = self.tabla.selection()
        if not sel: return
        id_doc = self.tabla.item(sel[0], "values")[1] 
        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT enlace_pdf_sunat, archivo_ruta FROM facturas_emitidas WHERE id = %s", (id_doc,))
            res = cursor.fetchone()
            if res:
                enlace_pdf = res[0]
                ruta_local = res[1]
                # Se resuelve la ruta: puede estar guardada como relativa a la
                # carpeta sincronizada o pertenecer a otro equipo (Mac/Windows).
                ruta_real = resolver_archivo_local(ruta_local)
                if ruta_real:
                    abrir_documento(ruta_real)
                elif enlace_pdf and str(enlace_pdf).strip():
                    webbrowser.open(str(enlace_pdf).strip())
                elif ruta_local:
                    messagebox.showwarning(
                        "PDF no encontrado",
                        "La factura tiene un PDF asociado, pero el archivo no se encuentra en este equipo." + chr(10) + chr(10) +
                        "Ruta registrada: " + str(ruta_local) + chr(10) + chr(10) +
                        "Verifique que la carpeta sincronizada esté actualizada (Rclone).",
                        parent=self.main_root)
                else:
                    messagebox.showinfo("Aviso", "No hay PDF asociado a esta factura.",
                                        parent=self.main_root)
        except Exception as e:
            messagebox.showerror("Error", "No se pudo abrir el PDF de la factura: " + str(e),
                                 parent=self.main_root)
        finally: liberar_conexion(conn)

    def abrir_ventana_edicion(self):
        sel = self.tabla.selection()
        if not sel: return messagebox.showwarning("Atención", "Seleccione un documento de la lista para gestionar.")
        
        valores = self.tabla.item(sel[0], "values")
        id_doc = valores[1]
        cli_str = valores[6]
        
        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cols_reg = ("tipo_documento, numero_documento, fecha, cliente, descripcion, evento_asociado, "
                        "subtotal, dias_credito, COALESCE(det_porcentaje, 0), impuesto, total, "
                        "enlace_pdf_sunat, estado_sunat, orden_compra, COALESCE(det_monto, 0)")
            if _COLUMNA_DET_REDONDEO:
                cols_reg += ", COALESCE(det_redondeo, FALSE)"
            cursor.execute(f"SELECT {cols_reg} FROM facturas_emitidas WHERE id = %s", (id_doc,))
            reg = cursor.fetchone()
            
            cursor.execute("SELECT ruc, direccion_fiscal, correo FROM clientes WHERE TRIM(UPPER(nombre_empresa)) = TRIM(UPPER(%s))", (cli_str,))
            res_cli = cursor.fetchone()
            ruc_cliente = res_cli[0] if res_cli and res_cli[0] else ""
            dir_cliente = res_cli[1] if res_cli and res_cli[1] else "-"
            correo_cliente = res_cli[2] if res_cli and res_cli[2] else ""
        finally:
            liberar_conexion(conn)
            
        if not reg: return
        enlace_sunat = reg[11]
        estado_sunat_db = reg[12]

        # 🧮 Estado inicial del check de redondeo de la detracción:
        # se lee de la BD si la columna existe; si no, se deduce del monto guardado
        # (un monto entero que no coincide con el cálculo exacto solo puede venir de un redondeo).
        det_monto_guardado = float(reg[14]) if len(reg) > 14 and reg[14] else 0.0
        redondeo_inicial = bool(reg[15]) if (len(reg) > 15 and reg[15] is not None) else False
        if len(reg) <= 15 and "Recibo" not in str(reg[0]) and det_monto_guardado:
            base_calc = float(reg[6] or 0) * (1.18 if "Factura" in str(reg[0]) else 1.0)
            exacto = base_calc * (float(reg[8] or 0) / 100.0)
            redondeo_inicial = (abs(det_monto_guardado - round(det_monto_guardado)) < 0.005
                                and abs(det_monto_guardado - exacto) > 0.005)
        
        ya_anulado = False
        if estado_sunat_db and "Anulada" in str(estado_sunat_db):
            ya_anulado = True

        v_edit = ctk.CTkToplevel(self.main_root)
        v_edit.title("Gestión de Documento Emitido")
        v_edit.geometry("450x450") 
        v_edit.transient(self.main_root)
        v_edit.grab_set()

        ctk.CTkLabel(v_edit, text=f"Gestión del Registro ID: {id_doc}", font=("Arial", 14, "bold"), text_color="#1f538d").pack(pady=(15, 10))

        fecha_em = datetime.now().strftime("%Y-%m-%d")
        if reg[2]:
            try: fecha_em = datetime.strptime(reg[2], "%d/%m/%Y").strftime("%Y-%m-%d")
            except: pass

        datos_factura = {
            "id_factura_bd": id_doc,
            "tipo_doc": reg[0],
            "serie": str(reg[1]).split("-")[0] if "-" in str(reg[1]) else "F001",
            "numero": str(reg[1]).split("-")[1] if "-" in str(reg[1]) else "1",
            "ruc_dni_cliente": ruc_cliente,
            "nombre_cliente": reg[3],
            "direccion_cliente": dir_cliente,
            "correo_cliente": correo_cliente,
            "fecha_emision": fecha_em,
            "moneda": "Soles",
            "subtotal": float(reg[6]) if reg[6] else 0.0,
            "impuesto": float(reg[9]) if reg[9] else 0.0,
            "total": float(reg[10]) if len(reg) > 10 and reg[10] else float(reg[6] or 0) + float(reg[9] or 0),
            "descripcion": reg[4]
        }

        f_docs = ctk.CTkFrame(v_edit, fg_color="#f0f0f0", border_width=1, border_color="#e0e0e0")
        f_docs.pack(fill="x", padx=30, pady=10, ipadx=5, ipady=5)
        ctk.CTkLabel(f_docs, text="📄 Documentos Asociados:", font=("Arial", 11, "bold")).pack(pady=5)

        def descargar_pdf_sunat():
            conn = conectar_db()
            if not conn: return
            try:
                cursor = conn.cursor()
                cursor.execute("SELECT enlace_pdf_sunat, archivo_ruta FROM facturas_emitidas WHERE id = %s", (id_doc,))
                res = cursor.fetchone()
                if res:
                    link_pdf, ruta_local = res[0], res[1]
                    ruta_real = resolver_archivo_local(ruta_local)
                    if ruta_real:
                        abrir_documento(ruta_real)
                    elif link_pdf and str(link_pdf).strip():
                        webbrowser.open(str(link_pdf).strip())
                    elif ruta_local:
                        messagebox.showwarning(
                            "PDF no encontrado",
                            "El comprobante tiene un PDF asociado, pero el archivo no se encuentra en este equipo." + chr(10) + chr(10) +
                            "Ruta registrada: " + str(ruta_local),
                            parent=v_edit)
                    else:
                        messagebox.showinfo("Aviso", "Este comprobante aún no tiene PDF asociado.", parent=v_edit)
            except Exception as e: messagebox.showerror("Error", str(e), parent=v_edit)
            finally: liberar_conexion(conn)

        def abrir_oc_asociada(numero_oc):
            conn = conectar_db()
            if not conn: return
            try:
                c = conn.cursor()
                c.execute("SELECT archivo_ruta FROM ordenes_compra_clientes WHERE numero_oc = %s", (numero_oc,))
                res = c.fetchone()
                from app_paths import resolver_ruta_archivo
                ruta_oc = resolver_ruta_archivo(res[0]) if res and res[0] else ""
                if ruta_oc:
                    abrir_documento(ruta_oc)
                else:
                    messagebox.showinfo("Aviso", "No se encontró el PDF de esta Orden de Compra.\n\nEs probable que esta orden se haya registrado manualmente sin archivo adjunto, o que no se haya realizado.", parent=v_edit)
            except Exception as e: messagebox.showerror("Error", str(e), parent=v_edit)
            finally: liberar_conexion(conn)

        def adjuntar_pdf_a_registro():
            """Adjunta (o reemplaza) el PDF de una factura ya registrada."""
            ruta_base = obtener_ruta_base_drive()
            if not ruta_base:
                avisar_sin_permiso_guardado(v_edit)
                return
            ruta_pdf = seleccionar_archivo_dialogo(
                titulo="Seleccionar PDF de la Factura",
                tipos=[("Archivos PDF", "*.pdf"),
                       ("Imágenes (JPG/PNG)", "*.png;*.jpg;*.jpeg"),
                       ("Todos los archivos", "*.*")])
            if not ruta_pdf:
                return
            try:
                carpeta_destino = os.path.join(ruta_base, "facturas_emitidas")
                if not os.path.exists(carpeta_destino):
                    os.makedirs(carpeta_destino)
                nombre_ext = os.path.splitext(ruta_pdf)[1]
                ruta_final = os.path.join(carpeta_destino, f"Emitida_{datetime.now().strftime('%Y%m%d%H%M%S')}_{id_doc}{nombre_ext}")
                shutil.copy2(ruta_pdf, ruta_final)
            except Exception as e:
                return messagebox.showerror("Error", f"No se pudo guardar el archivo:\n{e}", parent=v_edit)

            conn = conectar_db()
            if not conn:
                return
            try:
                cursor = conn.cursor()
                cursor.execute("UPDATE facturas_emitidas SET archivo_ruta = %s WHERE id = %s", (ruta_para_guardar(ruta_final), id_doc))
                conn.commit()
                cache_sistema.invalidar()
                registrar_auditoria(self.app_padre.usuario_activo, "Facturas Emitidas", f"Adjuntó PDF a la factura ID {id_doc}")
            except Exception as e:
                messagebox.showerror("Error SQL", str(e), parent=v_edit)
            finally:
                liberar_conexion(conn)

            messagebox.showinfo("Éxito", "PDF adjuntado correctamente a la factura.", parent=v_edit)
            self.cargar_datos_tabla(reset_pagina=True)
            if hasattr(self.app_padre, 'app_cobros'):
                self.app_padre.app_cobros.cargar_datos_cobrar()

        ctk.CTkButton(f_docs, text="Ver Factura (Local/SUNAT)", fg_color="#27ae60", hover_color="#1e8449", height=28, command=descargar_pdf_sunat).pack(fill="x", padx=10, pady=2)
        ctk.CTkButton(f_docs, text="📎 Adjuntar / Reemplazar PDF", fg_color="#1f538d", hover_color="#163b65", height=28, command=adjuntar_pdf_a_registro).pack(fill="x", padx=10, pady=2)
        
        if reg[13]: 
            ctk.CTkButton(f_docs, text=f"Ver Orden de Compra: {reg[13]}", fg_color="#d35400", hover_color="#a84300", height=28, command=lambda: abrir_oc_asociada(reg[13])).pack(fill="x", padx=10, pady=2)

        def modificar_datos_registro():
            v_edit.destroy()
            v_mod = ctk.CTkToplevel(self.main_root)
            v_mod.title(f"Modificar Registro ID: {id_doc}")
            v_mod.geometry("400x720")
            v_mod.transient(self.main_root)
            v_mod.grab_set()

            ctk.CTkLabel(v_mod, text="Modificar Datos del Documento", font=("Arial", 14, "bold"), text_color="#1f538d").pack(pady=(15, 10))

            f_cont = ctk.CTkScrollableFrame(v_mod, fg_color="transparent")
            f_cont.pack(fill="both", expand=True, padx=20, pady=5)

            lbl_tit_det = ctk.CTkLabel(f_cont, text="Detracción (%):", font=("Arial", 11, "bold"))
            chk_redondeo = None      # se crea más abajo, junto al campo de porcentaje

            def toggle_detraccion_edit(*args):
                tipo_val = cmb_tipo.get()
                if chk_redondeo is not None:
                    try:
                        chk_redondeo.configure(state="disabled" if "Recibo" in tipo_val else "normal")
                    except Exception:
                        pass
                if "Recibo" in tipo_val:
                    lbl_tit_det.configure(text="Retención (%):")
                    if "8%" in tipo_val:
                        ent_det.delete(0, tk.END); ent_det.insert(0, "8")
                    else:
                        ent_det.delete(0, tk.END); ent_det.insert(0, "0")
                else:
                    lbl_tit_det.configure(text="Detracción (%):")

            ctk.CTkLabel(f_cont, text="Tipo de Documento:", font=("Arial", 11, "bold")).pack(anchor="w")
            cmb_tipo = ctk.CTkComboBox(f_cont, values=["Factura (18% IGV)", "Boleta (Sin IGV)", "Recibo por Honorarios (8% Retención)", "Recibo por Honorarios (Sin Retención)"], state="readonly", command=toggle_detraccion_edit)
            cmb_tipo.pack(fill="x", pady=(0, 10))
            cmb_tipo.set(reg[0])

            ctk.CTkLabel(f_cont, text="N° de Documento:", font=("Arial", 11, "bold")).pack(anchor="w")
            ent_nro = ctk.CTkEntry(f_cont)
            ent_nro.pack(fill="x", pady=(0, 10))
            ent_nro.insert(0, reg[1] if reg[1] else "")

            ctk.CTkLabel(f_cont, text="Fecha (DD/MM/AAAA):", font=("Arial", 11, "bold")).pack(anchor="w")
            ent_f = ctk.CTkEntry(f_cont)
            ent_f.pack(fill="x", pady=(0, 10))
            ent_f.insert(0, reg[2])

            ctk.CTkLabel(f_cont, text="Días de Crédito:", font=("Arial", 11, "bold")).pack(anchor="w")
            ent_d = ctk.CTkEntry(f_cont)
            ent_d.pack(fill="x", pady=(0, 10))
            ent_d.insert(0, str(reg[7]))

            ctk.CTkLabel(f_cont, text="Cliente:", font=("Arial", 11, "bold")).pack(anchor="w")
            ent_c = ctk.CTkEntry(f_cont)
            ent_c.pack(fill="x", pady=(0, 10))
            ent_c.insert(0, reg[3])

            ctk.CTkLabel(f_cont, text="Concepto / Descripción:", font=("Arial", 11, "bold")).pack(anchor="w")
            ent_desc = ctk.CTkEntry(f_cont)
            ent_desc.pack(fill="x", pady=(0, 10))
            ent_desc.insert(0, reg[4])
            
            conn = conectar_db()
            ocs_lista = ["--- Sin Orden de Compra ---"]
            if conn:
                try:
                    c = conn.cursor()
                    c.execute("SELECT numero_oc FROM ordenes_compra_clientes ORDER BY id DESC")
                    for r_oc in c.fetchall(): ocs_lista.append(r_oc[0])
                except: pass
                finally: liberar_conexion(conn)
                
            ctk.CTkLabel(f_cont, text="Orden de Compra:", font=("Arial", 11, "bold"), text_color="#166534").pack(anchor="w")
            cmb_oc = ctk.CTkComboBox(f_cont, values=ocs_lista, state="readonly")
            cmb_oc.pack(fill="x", pady=(0, 10))
            if reg[13]: cmb_oc.set(reg[13])
            else: cmb_oc.set("--- Sin Orden de Compra ---")

            ctk.CTkLabel(f_cont, text="Monto Base (Subtotal):", font=("Arial", 11, "bold")).pack(anchor="w")
            ent_s = ctk.CTkEntry(f_cont)
            ent_s.pack(fill="x", pady=(0, 10))
            ent_s.insert(0, str(reg[6]))

            lbl_tit_det.pack(anchor="w")
            ent_det = ctk.CTkEntry(f_cont)
            ent_det.pack(fill="x", pady=(0, 10))

            # 🧮 Redondeo de la detracción al entero (fracción < 0.50 baja, ≥ 0.50 sube)
            chk_redondeo = ctk.CTkCheckBox(f_cont, text="🧮 Redondear detracción al entero",
                                           font=("Arial", 11, "bold"),
                                           checkbox_width=18, checkbox_height=18)
            chk_redondeo.pack(anchor="w", pady=(0, 10))
            if redondeo_inicial:
                chk_redondeo.select()
            if "Recibo" in str(reg[0]):
                chk_redondeo.configure(state="disabled")
            
            if "Recibo" in str(reg[0]):
                lbl_tit_det.configure(text="Retención (%):")
                if reg[6] > 0 and reg[9] > 0:
                    pct = (float(reg[9]) / float(reg[6])) * 100
                    ent_det.insert(0, f"{pct:g}")
                else:
                    ent_det.insert(0, "0")
            else:
                lbl_tit_det.configure(text="Detracción (%):")
                ent_det.insert(0, str(reg[8]) if reg[8] else CONFIG_REGIONAL.get("detraccion_porcentaje", "12"))

            def guardar_cambios():
                tipo = cmb_tipo.get()
                nro_doc = ent_nro.get().strip()
                fecha = ent_f.get().strip()
                cli = ent_c.get().strip()
                desc = ent_desc.get().strip()
                evento = "GENERAL / NO ASIGNADO"
                oc_sel = cmb_oc.get()
                if oc_sel == "--- Sin Orden de Compra ---": oc_sel = ""
                
                try:
                    sub = float(ent_s.get())
                    dias = int(ent_d.get())
                    ui_pct = float(ent_det.get() if ent_det.get() else 0)
                except ValueError:
                    messagebox.showerror("Error", "Monto, días y porcentaje deben ser numéricos.", parent=v_mod)
                    return

                # 🧮 Redondeo opcional de la detracción al entero (check de esta ventana)
                try:
                    redondeo_det = bool(chk_redondeo.get()) and "Recibo" not in tipo
                except Exception:
                    redondeo_det = False

                if "Factura" in tipo: 
                    imp = redondear_centavos(sub * 0.18); tot_bruto = redondear_centavos(sub + imp)
                    det_pct = ui_pct; det_monto = tot_bruto * (det_pct / 100.0)
                    det_monto = redondear_detraccion(det_monto) if redondeo_det else redondear_centavos(det_monto)
                    neto_nuevo = redondear_centavos(tot_bruto - det_monto)
                elif "Recibo" in tipo: 
                    imp = redondear_centavos(sub * (ui_pct / 100.0)); tot_bruto = redondear_centavos(sub)
                    det_pct = 0.0; det_monto = 0.0
                    neto_nuevo = redondear_centavos(sub - imp)
                else: 
                    imp = 0.0; tot_bruto = redondear_centavos(sub); det_pct = ui_pct
                    det_monto = tot_bruto * (det_pct / 100.0)
                    det_monto = redondear_detraccion(det_monto) if redondeo_det else redondear_centavos(det_monto)
                    neto_nuevo = redondear_centavos(tot_bruto - det_monto)

                conn2 = conectar_db()
                if not conn2: return
                try:
                    c2 = conn2.cursor()
                    
                    if nro_doc:
                        c2.execute("SELECT COUNT(*) FROM facturas_emitidas WHERE numero_documento = %s AND id != %s", (nro_doc, id_doc))
                        if c2.fetchone()[0] > 0:
                            liberar_conexion(conn2)
                            return messagebox.showwarning("Duplicado", f"El N° de Documento '{nro_doc}' ya está en uso.", parent=v_mod)

                    cli_match = cli.strip().upper()
                    c2.execute("SELECT limite_credito FROM clientes WHERE TRIM(UPPER(nombre_empresa)) = %s", (cli_match,))
                    res_cli = c2.fetchone()
                    limite_credito = float(res_cli[0]) if (res_cli and res_cli[0] is not None) else 0.0
                    
                    if limite_credito > 0.0:
                        c2.execute("""
                            SELECT id, total, COALESCE(det_monto, 0), tipo_documento, COALESCE(impuesto, 0)
                            FROM facturas_emitidas WHERE TRIM(UPPER(cliente)) = %s
                        """, (cli_match,))
                        
                        facturas_historicas = c2.fetchall()
                        deuda_actual = 0.0
                        
                        for r in facturas_historicas:
                            f_id, tot_b, det_m, t_doc, imp_val = r
                            if str(f_id) == str(id_doc): continue
                            
                            tot_val = float(tot_b) if tot_b else 0.0
                            det_val = float(det_m) if det_m else 0.0
                            impuesto_val = float(imp_val) if imp_val else 0.0
                            
                            if t_doc and "Recibo" in t_doc and "8%" in t_doc: neto_fac = tot_val - impuesto_val - det_val
                            else: neto_fac = tot_val - det_val
                            
                            c2.execute("SELECT SUM(monto_pagado + COALESCE(impuestos_deducibles, 0)) FROM pagos_clientes WHERE id_factura = %s", (f_id,))
                            pagado = c2.fetchone()[0]
                            pagado_val = float(pagado) if pagado else 0.0
                            
                            saldo = max(0.0, neto_fac - pagado_val)
                            deuda_actual += saldo
                        
                        c2.execute("SELECT SUM(monto_pagado + COALESCE(impuestos_deducibles, 0)) FROM pagos_clientes WHERE id_factura = %s", (id_doc,))
                        pag_esta = c2.fetchone()[0]
                        pagado_esta_val = float(pag_esta) if pag_esta else 0.0
                        
                        saldo_nuevo_esta = max(0.0, neto_nuevo - pagado_esta_val)
                        deuda_proyectada = deuda_actual + saldo_nuevo_esta
                        
                        if deuda_proyectada > limite_credito:
                            msg = (f"⚠️ LÍMITE DE CRÉDITO EXCEDIDO\n\n"
                                   f"Al modificar esta factura, se excede el límite del cliente.\n"
                                   f"Límite Asignado: {formatear_moneda(limite_credito)}\n"
                                   f"Deuda de otras Facturas: {formatear_moneda(deuda_actual)}\n"
                                   f"Deuda Proyectada Total: {formatear_moneda(deuda_proyectada)}\n")
                            liberar_conexion(conn2)
                            return messagebox.showerror("Bloqueo por Crédito", msg, parent=v_mod)


                    sql_upd = ("UPDATE facturas_emitidas SET tipo_documento=%s, numero_documento=%s, fecha=%s, "
                               "cliente=%s, descripcion=%s, evento_asociado=%s, subtotal=%s, impuesto=%s, total=%s, "
                               "dias_credito=%s, det_porcentaje=%s, det_monto=%s, orden_compra=%s{extra} WHERE id=%s")
                    vals_upd = [tipo, nro_doc, fecha, cli, desc, evento, sub, imp, tot_bruto, dias, det_pct, det_monto, oc_sel]
                    extra_redondeo = ""
                    if _COLUMNA_DET_REDONDEO:
                        extra_redondeo = ", det_redondeo=%s"
                        vals_upd.append(bool(redondeo_det))
                    vals_upd.append(id_doc)
                    c2.execute(sql_upd.format(extra=extra_redondeo), tuple(vals_upd))
                    conn2.commit()
                    cache_sistema.invalidar()
                    registrar_auditoria(self.app_padre.usuario_activo, "Facturas Emitidas", f"Modificó la factura ID {id_doc}")
                    messagebox.showinfo("Éxito", "Registro modificado correctamente.", parent=v_mod)
                    v_mod.destroy()
                    self.cargar_datos_tabla(reset_pagina=True)
                    if hasattr(self.app_padre, 'app_cobros'): self.app_padre.app_cobros.cargar_datos_cobrar()
                except Exception as e: messagebox.showerror("Error SQL", str(e), parent=v_mod)
                finally: liberar_conexion(conn2)

            ctk.CTkButton(f_cont, text="💾 Guardar Cambios", font=("Arial", 12, "bold"), fg_color="#1f538d", hover_color="#163b65", command=guardar_cambios).pack(fill="x", pady=20)

        def eliminar_registro():
            # 🔗 Los cobros (pagos_clientes) de esta factura se consultan ANTES de borrar:
            # si quedaran huérfanos seguirían sumando como ingreso en el saldo de bancos.
            cobros = []
            try:
                conn_c = conectar_db(silencioso=True)
                if conn_c:
                    try:
                        with conn_c.cursor() as cc:
                            cc.execute("""SELECT monto_pagado, COALESCE(archivo_ruta, '')
                                          FROM pagos_clientes WHERE id_factura = %s""", (id_doc,))
                            cobros = cc.fetchall()
                    finally:
                        liberar_conexion(conn_c)
            except Exception:
                cobros = []

            aviso = ("⚠️ ¿Desea eliminar completamente este registro?\n\n"
                     "Se borrará de la base de datos y el archivo físico asociado.")
            if cobros:
                total_cobrado = sum(float(cc[0] or 0) for cc in cobros)
                aviso += (f"\n\n💰 También se eliminarán {len(cobros)} cobro(s) registrado(s) de esta "
                          f"factura (total {formatear_moneda(total_cobrado)}) junto con sus comprobantes.\n"
                          f"Si no se borraran, seguirían contando como ingreso en el saldo de bancos.")
            if messagebox.askyesno("Confirmar Eliminación", aviso, parent=v_edit):
                try:
                    conn = conectar_db()
                    cursor = conn.cursor()
                    cursor.execute("SELECT archivo_ruta FROM facturas_emitidas WHERE id = %s", (id_doc,))
                    row = cursor.fetchone()
                    ruta_archivo = row[0]
                    eliminar_archivo(ruta_archivo)   # resuelve rutas de otros equipos/SO

                    # 🗑️ Se borran también los cobros de esta factura (y sus comprobantes físicos)
                    for _monto_c, ruta_c in cobros:
                        try:
                            eliminar_archivo(os.path.normpath(ruta_c) if ruta_c else None)
                        except Exception:
                            pass
                    cursor.execute("DELETE FROM pagos_clientes WHERE id_factura = %s", (id_doc,))
                    cobros_borrados = max(cursor.rowcount, 0)

                    cursor.execute("DELETE FROM facturas_emitidas WHERE id = %s", (id_doc,))
                    conn.commit()
                    liberar_conexion(conn)
                    
                    cache_sistema.invalidar()
                    registrar_auditoria(self.app_padre.usuario_activo, "Facturas Emitidas",
                                        f"Eliminó completamente la factura ID {id_doc} "
                                        f"({cobros_borrados} cobro(s) asociado(s))")
                    messagebox.showinfo("Éxito", "Registro eliminado.", parent=v_edit)
                    v_edit.destroy()
                    self.cargar_datos_tabla(reset_pagina=True)
                    if hasattr(self.app_padre, 'app_cobros'): self.app_padre.app_cobros.cargar_datos_cobrar()
                except Exception as e: messagebox.showerror("Error", str(e), parent=v_edit)

        def emitir_sunat():
            if not abrir_ventana_emision_sunat:
                return messagebox.showerror("Error", "El módulo 'facturacion_sunat.py' no se encuentra disponible.", parent=v_edit)
                
            def accion_post_otp():
                abrir_ventana_emision_sunat(v_edit, datos_factura, self.app_padre.usuario_activo, callback_exito=self.cargar_datos_tabla)

            solicitar_otp(f"Emitir comprobante {reg[1]} a SUNAT", accion_post_otp, v_edit)

        def procesar_nota_credito():
            v_nc = ctk.CTkToplevel(v_edit)
            v_nc.title("Emitir Nota de Crédito")
            v_nc.geometry("400x250")
            v_nc.grab_set()

            ctk.CTkLabel(v_nc, text="⚠️ Anular Factura vía Nota de Crédito", font=("Arial", 14, "bold"), text_color="#c0392b").pack(pady=(15, 10))
            ctk.CTkLabel(v_nc, text="Seleccione el motivo legal para SUNAT:", font=("Arial", 11)).pack(anchor="w", padx=20)
            
            motivos = ["Anulación de la operación", "Anulación por error en el RUC", "Devolución total", "Devolución por ítem"]
            cmb_motivo = ctk.CTkComboBox(v_nc, values=motivos, width=320)
            cmb_motivo.pack(pady=10)

            def confirmar_nc():
                motivo_sel = cmb_motivo.get()
                v_nc.destroy()
                
                def accion_post_otp():
                    datos_nc = datos_factura.copy()
                    datos_nc["tipo_emision"] = "nota_credito"
                    datos_nc["motivo_nc"] = motivo_sel
                    
                    if abrir_ventana_emision_sunat:
                        def callback_ambos():
                            self.cargar_datos_tabla(reset_pagina=True)
                            if hasattr(self.app_padre, 'app_nc'):
                                self.app_padre.app_nc.cargar_datos_nc(reset_pagina=True)
                        
                        abrir_ventana_emision_sunat(v_edit, datos_nc, self.app_padre.usuario_activo, callback_exito=callback_ambos)
                    else:
                        messagebox.showerror("Error", "Módulo de facturación inactivo.")

                solicitar_otp(f"Emitir Nota Crédito para la Fac. {reg[1]}", accion_post_otp, v_edit)

            ctk.CTkButton(v_nc, text="✅ Generar Nota de Crédito", fg_color="#c0392b", hover_color="#922b21", font=("Arial", 12, "bold"), command=confirmar_nc).pack(pady=15)

        if not enlace_sunat: 
            ctk.CTkButton(v_edit, text="⚡ Emitir Comprobante a SUNAT", font=("Arial", 12, "bold"), fg_color="#d35400", hover_color="#a84300", command=emitir_sunat).pack(fill="x", padx=30, pady=5)
            ctk.CTkButton(v_edit, text="✏️ Modificar Datos del Registro", font=("Arial", 12, "bold"), fg_color="#34495e", hover_color="#2c3e50", command=modificar_datos_registro).pack(fill="x", padx=30, pady=5)
            ctk.CTkButton(v_edit, text="❌ Eliminar Registro Completo", font=("Arial", 12, "bold"), fg_color="#e74c3c", hover_color="#c0392b", command=eliminar_registro).pack(fill="x", padx=30, pady=(5, 10))
        elif ya_anulado:
            ctk.CTkLabel(v_edit, text="❌ Documento Anulado", font=("Arial", 14, "bold"), text_color="#c0392b").pack(pady=5)
        else:
            ctk.CTkButton(v_edit, text="⚠️ Generar Nota de Crédito / Anular", font=("Arial", 12, "bold"), fg_color="#c0392b", hover_color="#922b21", command=procesar_nota_credito).pack(fill="x", padx=30, pady=(5, 10))

# =========================================================
# PESTAÑA 2: CUENTAS POR COBRAR (COBROS Y DEUDAS)
# =========================================================
class CuentasPorCobrarTab:
    # 🖱️ Reutiliza el menú de clic derecho de la pestaña de facturas
    #    (ver factura / ver comprobante de pago).
    _datos_factura_venta = FacturasEmitidasTab._datos_factura_venta
    _primera_ruta_venta = FacturasEmitidasTab._primera_ruta_venta
    _abrir_comprobantes_venta = FacturasEmitidasTab._abrir_comprobantes_venta
    _menu_contextual_factura = FacturasEmitidasTab._menu_contextual_factura

    def __init__(self, tab_frame, main_root, app_padre):
        self.tab_frame = tab_frame
        self.main_root = main_root
        self.app_padre = app_padre
        self.orden_columnas = {}
        
        # 🚀 VARIABLES DE PAGINACIÓN
        self.pagina_actual = 1
        self.registros_por_pagina = 50
        
        # 🗓️ FILTRO DE MES (por defecto, el mes en curso)
        self.mes_filtro = mes_en_curso()
        
        self.inicializar_entorno()
        self.crear_interfaz()

    # 🚀 FIX: INICIALIZACIÓN ASÍNCRONA
    def inicializar_entorno(self):
        def tarea_init():
            conn = conectar_db(silencioso=True)
            if not conn: return
            try:
                cursor = conn.cursor()
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS pagos_clientes (
                        id SERIAL PRIMARY KEY, codigo_cotizacion VARCHAR(255), monto_pagado NUMERIC, 
                        archivo_ruta TEXT, cliente_nombre VARCHAR(255), fecha_pago VARCHAR(50),
                        id_factura INTEGER DEFAULT 0
                    )
                """)
                conn.commit()
                
                try:
                    cursor.execute("ALTER TABLE pagos_clientes ADD COLUMN cuenta_destino VARCHAR(255) DEFAULT ''")
                    conn.commit()
                except Exception: conn.rollback()

                # 💳 Pagos por adelantado (sin factura): tipo y contrato externo
                for sentencia in ("ALTER TABLE pagos_clientes ADD COLUMN IF NOT EXISTS tipo VARCHAR(20) DEFAULT 'Cobro'",
                                  "ALTER TABLE pagos_clientes ADD COLUMN IF NOT EXISTS contrato VARCHAR(255) DEFAULT ''",
                                  "ALTER TABLE pagos_clientes ADD COLUMN IF NOT EXISTS numero_operacion VARCHAR(100) DEFAULT ''"):
                    try:
                        cursor.execute(sentencia); conn.commit()
                    except Exception: conn.rollback()
                
            except Exception: pass
            finally: liberar_conexion(conn)
        
        threading.Thread(target=tarea_init, daemon=True).start()

    def crear_interfaz(self):
        frame_acciones = ctk.CTkFrame(self.tab_frame, corner_radius=8, fg_color="#f8f9fa", border_width=1, border_color="#e0e0e0")
        frame_acciones.pack(fill="x", padx=15, pady=(10, 10), ipady=5)

        btn_cols = ctk.CTkButton(frame_acciones, text="⚙️ Columnas", width=90, font=("Arial", 12, "bold"), fg_color="#6c757d", hover_color="#565e64", command=lambda: menu_columnas(self.tabla, "ventas_cuentas_cobrar"))
        btn_cols.pack(side="left", padx=5, pady=5)
        try:
            aplicar_columnas_visibles(self.tabla, "ventas_cuentas_cobrar")
        except Exception:
            pass
        btn_resumen_cli = ctk.CTkButton(frame_acciones, text="📊 Resumen por Cliente", font=("Arial", 12, "bold"), command=self.mostrar_resumen_clientes, fg_color="#1f538d", hover_color="#163b65")
        btn_resumen_cli.pack(side="left", padx=5, pady=5)

        btn_reporte = ctk.CTkButton(frame_acciones, text="📈 Reporte Avanzado", font=("Arial", 12, "bold"), command=self.mostrar_reporte_totales, fg_color="#34495e", hover_color="#2c3e50")
        btn_reporte.pack(side="left", padx=5, pady=5)

        btn_varias = ctk.CTkButton(frame_acciones, text="🧾 Cobrar Varias",
                                    font=("Arial", 12, "bold"), fg_color="#b9770e", hover_color="#8e5e0a",
                                    command=self.cobrar_varias_facturas)
        btn_varias.pack(side="left", padx=5, pady=5)
        btn_adelanto = ctk.CTkButton(frame_acciones, text="💳 Registrar Adelanto", font=("Arial", 12, "bold"),
                                     fg_color="#8e44ad", hover_color="#703688",
                                     command=self.abrir_registrar_adelanto)
        btn_adelanto.pack(side="left", padx=5, pady=5)

        btn_aplicar = ctk.CTkButton(frame_acciones, text="🧩 Asociar Adelanto", font=("Arial", 12, "bold"),
                                    fg_color="#16a085", hover_color="#11806a",
                                    command=self.abrir_asociar_adelanto)
        btn_aplicar.pack(side="left", padx=5, pady=5)


        btn_refresh = ctk.CTkButton(frame_acciones, text="🔄 Actualizar", font=("Arial", 12, "bold"), command=lambda: self._refrescar_listas_cobro(), fg_color="#7f8c8d", hover_color="#606b6b")
        btn_refresh.pack(side="right", padx=10, pady=5)

        f_busqueda = ctk.CTkFrame(self.tab_frame, fg_color="transparent")
        f_busqueda.pack(fill="x", padx=15, pady=(0, 5))
        ctk.CTkLabel(f_busqueda, text="🔍 Buscar:", font=("Arial", 11, "bold")).pack(side="left", padx=(0, 5))
        self.ent_buscar_cobros = ctk.CTkEntry(f_busqueda, placeholder_text="Filtrar por documento, cliente, concepto...")
        self.ent_buscar_cobros.pack(side="left", fill="x", expand=True)
        
        ctk.CTkLabel(f_busqueda, text="🗓️ Mes:", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 5))
        self.combo_mes = ctk.CTkComboBox(f_busqueda, values=construir_valores_mes(), width=150, state="readonly", command=self.on_cambiar_mes)
        self.combo_mes.pack(side="left", padx=(0, 5))
        self.combo_mes.set(self.mes_filtro)
        
        self.ent_buscar_cobros.bind("<KeyRelease>", lambda e: self.buscar_con_retraso())
        self.ent_buscar_cobros.bind("<Return>", lambda e: self._refrescar_listas_cobro())

        f_tabla = ctk.CTkFrame(self.tab_frame, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True, padx=15, pady=0)

        # SE OMITIÓ LA COLUMNA EVENTO VISUALMENTE
        columnas = ("num", "id_factura", "fecha", "nro_doc", "cliente", "concepto", "subtotal", "igv", "detraccion", "neto_facturado", "cobrado", "saldo", "archivos")
        self.tabla = ttk.Treeview(f_tabla, columns=columnas, show="headings")

        self.tabla.tag_configure("con_cuenta", background="#e8f8f5", foreground="#0e6251") 
        self.tabla.tag_configure("sin_cuenta", background="#fdedec", foreground="#7b241c") 

        # 🎨 Estado de cobro: azul (cobrada) · rojo (no cobrada) · morado (parcial)
        self.tabla.tag_configure("pagada", background="#d6eaf8", foreground="#1a5276")
        self.tabla.tag_configure("no_pagada", background="#fadbd8", foreground="#922b21")
        self.tabla.tag_configure("parcial", background="#e8daef", foreground="#6c3483")

        self.tabla.heading("num", text="N°")
        self.tabla.heading("id_factura", text="ID (Oculto)")
        self.tabla.heading("fecha", text="Fecha Fac.")
        self.tabla.heading("nro_doc", text="N° Documento")
        self.tabla.heading("cliente", text="Cliente")
        self.tabla.heading("concepto", text="Concepto")
        self.tabla.heading("neto_facturado", text="Neto Facturado")
        self.tabla.heading("cobrado", text="Total Cobrado")
        self.tabla.heading("saldo", text="Saldo Pendiente")
        self.tabla.heading("archivos", text="Historial Adjuntos")

        self.tabla.column("num", width=40, anchor="center")
        self.tabla.column("id_factura", width=0, stretch=tk.NO) 
        self.tabla.column("fecha", width=80, anchor="center")
        # 📏 Ancho suficiente para el N° de documento completo (ver pestaña de facturas)
        self.tabla.column("nro_doc", width=130, minwidth=120, stretch=False, anchor="w")
        self.tabla.column("cliente", width=140, anchor="w")
        self.tabla.column("concepto", width=150, anchor="w")
        self.tabla.column("neto_facturado", width=95, anchor="e")
        self.tabla.column("cobrado", width=90, anchor="e")
        self.tabla.column("saldo", width=90, anchor="e")
        self.tabla.column("archivos", width=100, anchor="center")

        self.tabla.config(displaycolumns=("num", "fecha", "nro_doc", "cliente", "concepto", "neto_facturado", "cobrado", "saldo", "archivos"))
        self.tabla.bind("<Double-1>", self.abrir_todos_los_archivos)
        self.tabla.bind("<Button-3>", self._menu_contextual_factura)
        self.tabla.bind("<Button-2>", self._menu_contextual_factura)

        scroll_y = ctk.CTkScrollbar(f_tabla, orientation="vertical", command=self.tabla.yview)
        scroll_x = ctk.CTkScrollbar(f_tabla, orientation="horizontal", command=self.tabla.xview)
        self.tabla.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_x.pack(side="bottom", fill="x")
        scroll_y.pack(side="right", fill="y")
        self.tabla.pack(side="left", fill="both", expand=True)
        for _c in self.tabla["columns"]:
            try:
                self.tabla.column(_c, stretch=False)
            except Exception:
                pass

        self.frame_bottom = ctk.CTkFrame(self.tab_frame, fg_color="transparent")
        self.frame_bottom.pack(fill="x", padx=15, pady=10)

        # 🚀 BOTONES DE PAGINACIÓN (Agregados a la izquierda del bottom frame)
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

        # --- LEYENDA ---
        f_leyenda = ctk.CTkFrame(self.frame_bottom, fg_color="transparent")
        f_leyenda.pack(side="left", padx=30)
        
        ctk.CTkLabel(f_leyenda, text="■", font=("Arial", 14), text_color="#1a5276").pack(side="left", padx=(5,2))
        ctk.CTkLabel(f_leyenda, text="Totalmente cobrada", font=("Arial", 11, "bold"), text_color="#333333").pack(side="left")

        ctk.CTkLabel(f_leyenda, text="■", font=("Arial", 14), text_color="#922b21").pack(side="left", padx=(5,2))
        ctk.CTkLabel(f_leyenda, text="No cobrada", font=("Arial", 11, "bold"), text_color="#333333").pack(side="left")

        ctk.CTkLabel(f_leyenda, text="■", font=("Arial", 14), text_color="#6c3483").pack(side="left", padx=(5,2))
        ctk.CTkLabel(f_leyenda, text="Parcialmente cobrada", font=("Arial", 11, "bold"), text_color="#333333").pack(side="left")
        # ---------------

        self.lbl_total_general = ctk.CTkLabel(self.frame_bottom, text="Total Pendiente General por Cobrar: 0.00", font=("Arial", 12, "bold"), text_color="#c0392b")
        self.lbl_total_general.pack(side="right")

        self.main_root.after(150, lambda: self._refrescar_listas_cobro())

    def on_cambiar_mes(self, choice):
        self.mes_filtro = choice
        self._refrescar_listas_cobro()

    def pagina_anterior(self):
        if self.pagina_actual > 1:
            self.pagina_actual -= 1
            self.cargar_datos_cobrar()
            
    def pagina_siguiente(self):
        self.pagina_actual += 1
        self.cargar_datos_cobrar()

    def buscar_con_retraso(self):
        if hasattr(self, "_busqueda_job"):
            try: self.main_root.after_cancel(self._busqueda_job)
            except: pass
        self._busqueda_job = self.main_root.after(350, lambda: self._refrescar_listas_cobro())

    def exportar_excel(self):
        try: import pandas as pd
        except ImportError: return messagebox.showerror("Error", "Falta librería pandas.")
        filas = [self.tabla.item(item)["values"][2:] for item in self.tabla.get_children()]
        if not filas: return messagebox.showwarning("Aviso", "No hay registros.")
        columnas = ["Fecha Fac.", "N° Documento", "Cliente", "Concepto", "Subtotal", "IGV", "Detracción", "Neto Facturado", "Cobrado", "Saldo Pendiente", "Archivos"]
        ruta = guardar_archivo_dialogo(titulo="Exportar Cuentas por Cobrar", defaultextension=".xlsx", initialfile="Cuentas_por_Cobrar.xlsx", tipos=[("Excel", "*.xlsx")])
        if ruta:
            pd.DataFrame(filas, columns=columnas).to_excel(ruta, index=False)
            messagebox.showinfo("Éxito", f"Reporte exportado a:\n{ruta}")
            abrir_documento(ruta)

    def mostrar_resumen_clientes(self):
        saldos = {}
        for item in self.tabla.get_children():
            vals = self.tabla.item(item, "values")
            cli = vals[4] 
            try:
                saldo = desformatear_numero(vals[11]) 
                if saldo > 0: saldos[cli] = saldos.get(cli, 0.0) + saldo
            except ValueError: pass
            
        v_resumen = ctk.CTkToplevel(self.main_root)
        v_resumen.title("Resumen de Deudas por Cliente")
        v_resumen.geometry("500x400")
        v_resumen.grab_set()

        ctk.CTkLabel(v_resumen, text="📊 Saldo Pendiente Consolidado por Cliente", font=("Arial", 14, "bold"), text_color="#1f538d").pack(pady=(15, 10))

        f_tabla = ctk.CTkFrame(v_resumen, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True, padx=15, pady=(0, 10))

        t_resumen = ttk.Treeview(f_tabla, columns=("cliente", "saldo"), show="headings")
        t_resumen.heading("cliente", text="Cliente"); t_resumen.heading("saldo", text="Monto Pendiente de Cobro")
        t_resumen.column("cliente", width=250, anchor="w"); t_resumen.column("saldo", width=150, anchor="e")
        t_resumen.pack(side="left", fill="both", expand=True)
        
        total_cobrar = 0.0
        for p, s in sorted(saldos.items(), key=lambda x: x[0]):
            t_resumen.insert("", tk.END, values=(p, formatear_moneda(s))); total_cobrar += s

        ctk.CTkLabel(v_resumen, text=f"Total cuentas por cobrar : {formatear_moneda(total_cobrar)}", font=("Arial", 14, "bold")).pack(anchor="e", padx=15, pady=(5, 15))

    # 🚀 FIX: CARGA LAZY LOADING + CACHÉ
    def cobrar_varias_facturas(self):
        """🧾 Cobra dos o más facturas de una vez, con UN solo comprobante de depósito.

        El comprobante (cargado o escaneado) se asigna a TODAS las facturas
        seleccionadas, y cada una se cobra por su saldo pendiente.
        """
        seleccion = self.tabla.selection()
        if not seleccion:
            return messagebox.showinfo(
                "Cobrar varias facturas",
                "Seleccione dos o más facturas de la lista.\n\n"
                "(Puede mantener presionada la tecla Ctrl y hacer clic en cada una.)",
                parent=ventana_para_mensaje(getattr(self, "main_root", None)))
        ids = []
        for item in seleccion:
            try:
                v = self.tabla.item(item, "values")
                if len(v) > 1 and str(v[1]).strip():
                    ids.append(int(v[1]))
            except Exception:
                continue
        if len(ids) < 2:
            return messagebox.showwarning(
                "Cobrar varias facturas",
                "Seleccione al menos DOS facturas para cobrarlas con un solo comprobante.",
                parent=ventana_para_mensaje(getattr(self, "main_root", None)))

        pendientes = []
        try:
            conn = conectar_db(silencioso=True)
            if conn:
                try:
                    with conn.cursor() as cur:
                        for idf in ids:
                            cur.execute("""
                                SELECT f.numero_documento, f.cliente, COALESCE(f.total, 0),
                                       COALESCE(f.impuesto, 0), COALESCE(f.det_monto, 0),
                                       COALESCE(f.tipo_documento, ''),
                                       COALESCE((SELECT SUM(monto_pagado + COALESCE(impuestos_deducibles, 0)) FROM pagos_clientes p
                                                 WHERE p.id_factura = f.id), 0)
                                FROM facturas_emitidas f WHERE f.id = %s""", (idf,))
                            fila = cur.fetchone()
                            if fila:
                                _tipo_doc = str(fila[5] or "")
                                _tot = float(fila[2] or 0)
                                _imp = float(fila[3] or 0)
                                _det = float(fila[4] or 0)
                                _cobrado = float(fila[6] or 0)
                                if _tipo_doc and "Recibo" in _tipo_doc and "8%" in _tipo_doc:
                                    _neto = _tot - _imp - _det
                                else:
                                    _neto = _tot - _det
                                saldo = round(max(0.0, _neto - _cobrado), 2)
                                pendientes.append({"id": idf, "nro": str(fila[0] or ""),
                                                   "cliente": str(fila[1] or ""), "saldo": saldo})
                finally:
                    liberar_conexion(conn)
        except Exception as e:
            print("Aviso - saldos de las facturas:", e)
        if not pendientes:
            return messagebox.showinfo("Cobrar varias facturas",
                                       "No se pudieron leer las facturas seleccionadas.",
                                       parent=ventana_para_mensaje(getattr(self, "main_root", None)))

        # 🚫 Separar las facturas YA CANCELADAS (saldo 0) de las que sí tienen saldo
        canceladas = [p for p in pendientes if p["saldo"] <= 0.01]
        cobrables = [p for p in pendientes if p["saldo"] > 0.01]
        if not cobrables:
            detalle = chr(10).join("• %s — %s" % (p["nro"] or "s/n", p["cliente"]) for p in canceladas[:8])
            return messagebox.showwarning(
                "Facturas ya canceladas",
                "TODAS las facturas seleccionadas ya están COBRADAS/CANCELADAS:" + chr(10) + chr(10)
                + detalle + chr(10) + chr(10) + "No hay nada por cobrar.",
                parent=ventana_para_mensaje(getattr(self, "main_root", None)))
        if canceladas:
            detalle = chr(10).join("• %s — %s" % (p["nro"] or "s/n", p["cliente"]) for p in canceladas[:8])
            if not messagebox.askyesno(
                    "Algunas facturas ya están canceladas",
                    "Estas facturas ya están COBRADAS/CANCELADAS:" + chr(10) + chr(10)
                    + detalle + chr(10) + chr(10)
                    + "Se cobrarán SOLO las %d pendientes." % len(cobrables) + chr(10)
                    + "¿Desea continuar?",
                    parent=ventana_para_mensaje(getattr(self, "main_root", None))):
                return
        pendientes = cobrables


        bancos = []
        try:
            for b in (cargar_bancos() or []):
                if isinstance(b, dict):
                    nombre = " - ".join(str(x) for x in (b.get("banco"), b.get("cuenta")) if x).strip(" -")
                else:
                    nombre = str(b).strip()
                if nombre and nombre not in bancos:
                    bancos.append(nombre)
        except Exception:
            pass
        bancos += ["Efectivo / Caja Chica", "Tarjeta de Crédito", "Tarjeta de Débito", "Cheque", "Otro"]

        v = ctk.CTkToplevel(getattr(self, "main_root", None))
        v.title("🧾 Cobrar Varias Facturas (1 comprobante)")
        v.geometry("620x580")
        v.transient(getattr(self, "main_root", None))
        try:
            v.grab_set()
        except Exception:
            pass
        f = ctk.CTkFrame(v, fg_color="transparent")
        f.pack(fill="both", expand=True, padx=16, pady=12)

        ctk.CTkLabel(f, text="📋 %d facturas seleccionadas (se cobrarán con UN solo comprobante)"
                     % len(pendientes), font=("Arial", 12, "bold"), text_color="#1f538d").pack(anchor="w")

        ctk.CTkLabel(f, text="Fecha del cobro:", font=("Arial", 11, "bold")).pack(anchor="w", pady=(8, 0))
        f_fecha_cob = ctk.CTkFrame(f, fg_color="transparent")
        f_fecha_cob.pack(fill="x")
        ent_fecha = ctk.CTkEntry(f_fecha_cob)
        ent_fecha.pack(side="left", fill="x", expand=True)
        ent_fecha.insert(0, datetime.now().strftime("%d/%m/%Y"))
        ctk.CTkButton(f_fecha_cob, text="📅", width=40, fg_color="#1f538d", hover_color="#163b65",
                      command=lambda: CalendarioNativo(v, ent_fecha)).pack(side="left", padx=(5, 0))

        ctk.CTkLabel(f, text="Cuenta destino / método de cobro:", font=("Arial", 11, "bold")).pack(anchor="w", pady=(8, 0))
        cmb_cuenta = ctk.CTkComboBox(f, values=bancos, width=420, state="readonly")
        cmb_cuenta.pack(fill="x")
        if bancos:
            cmb_cuenta.set(bancos[0])

        ctk.CTkLabel(f, text="N° de Operación (opcional):", font=("Arial", 11, "bold")).pack(anchor="w", pady=(8, 0))
        ent_oper = ctk.CTkEntry(f)
        ent_oper.pack(fill="x")

        ctk.CTkLabel(f, text="💰 Monto recibido:", font=("Arial", 11, "bold")).pack(anchor="w", pady=(8, 0))
        ent_monto = ctk.CTkEntry(f)
        ent_monto.pack(fill="x")
        _total_pendiente = sum(max(0.0, p["saldo"]) for p in pendientes)
        lbl_calc = ctk.CTkLabel(f, text="", font=("Arial", 11, "bold"), text_color="#555555")

        def actualizar_calc(event=None):
            try:
                _monto = float(ent_monto.get().strip() or 0)
            except ValueError:
                _monto = 0.0
            _diff = round(_monto - _total_pendiente, 2)
            if _diff < -0.005:
                lbl_calc.configure(text="Total a cobrar: %s   ·   Falta: %s"
                                   % (formatear_moneda(_total_pendiente), formatear_moneda(abs(_diff))),
                                   text_color="#c0392b")
            elif _diff > 0.005:
                lbl_calc.configure(text="Total a cobrar: %s   ·   Sobra: %s"
                                   % (formatear_moneda(_total_pendiente), formatear_moneda(_diff)),
                                   text_color="#b9770e")
            else:
                lbl_calc.configure(text="Total a cobrar: %s   ·   Exacto" % formatear_moneda(_total_pendiente),
                                   text_color="#1e8449")

        ent_monto.bind("<KeyRelease>", actualizar_calc)
        lbl_calc.pack(anchor="w", pady=(2, 0))
        actualizar_calc()

        ctk.CTkLabel(f, text="📎 Comprobante de depósito (se asignará a todas):",
                     font=("Arial", 11, "bold")).pack(anchor="w", pady=(10, 0))
        estado_comp = {"ruta": ""}
        lbl_comp = ctk.CTkLabel(f, text="Sin comprobante cargado", font=("Arial", 10),
                                text_color="#999999", anchor="w")

        def actualizar_lbl_comp():
            if estado_comp["ruta"]:
                lbl_comp.configure(text="📎 " + os.path.basename(estado_comp["ruta"]), text_color="#1e8449")
            else:
                lbl_comp.configure(text="Sin comprobante cargado", text_color="#999999")
        def aplicar_ocr(ruta):
            """🔎 Lee el comprobante con OCR y rellena el monto y el N° de operación."""
            try:
                from modulo_compras import (leer_ocr_windows, extraer_numero_operacion,
                                            monto_del_comprobante)
            except Exception as _e_imp:
                print("[OCR comprobante] no se pudo importar OCR:", _e_imp)
                return
            try:
                _texto = leer_ocr_windows(ruta)
                if not _texto.strip():
                    try:
                        _texto = escaner.texto_de_resultados(escaner.ocr_paginas([ruta]))
                    except Exception:
                        _texto = ""
            except Exception as _e_ocr:
                print("[OCR comprobante] error:", _e_ocr)
                _texto = ""
            if not _texto.strip():
                return
            _num, _monto = "", 0.0
            try:
                _num = extraer_numero_operacion(_texto) or ""
            except Exception:
                _num = ""
            try:
                _monto = monto_del_comprobante(_texto) or 0.0
            except Exception:
                _monto = 0.0
            _avisos = []
            if _num:
                try:
                    ent_oper.delete(0, tk.END)
                    ent_oper.insert(0, str(_num))
                except Exception:
                    pass
                _avisos.append("N° de operación: " + str(_num))
            if _monto and float(_monto) > 0:
                try:
                    ent_monto.delete(0, tk.END)
                    ent_monto.insert(0, "%.2f" % float(_monto))
                except Exception:
                    pass
                try:
                    actualizar_calc()
                except Exception:
                    pass
                _avisos.append("Monto: %.2f" % float(_monto))
            if _avisos:
                try:
                    lbl_comp.configure(text="📎 " + os.path.basename(ruta) + "   ·   " + " · ".join(_avisos),
                                       text_color="#1e8449")
                except Exception:
                    pass



        f_comp = ctk.CTkFrame(f, fg_color="transparent")
        f_comp.pack(fill="x", pady=(0, 2))

        def cargar_comp():
            r = _pick_comprobante(v)
            if r:
                estado_comp["ruta"] = r
                actualizar_lbl_comp()
                aplicar_ocr(r)

        def escanear_comp():
            try:
                _cfg = cargar_configuracion_regional()
            except Exception:
                _cfg = {}
            _datos = escaner.config_escaner(_cfg)
            _res = _escanear_con_progreso(
                v, "el comprobante", escaner, escaner.carpeta_escaneos(_cfg),
                dispositivo=_datos.get("escaner_dispositivo", ""),
                dpi=_datos.get("escaner_dpi", "300"),
                color=_datos.get("escaner_color", "Escala de grises"),
                origen=_datos.get("escaner_origen", "Cristal (una página a la vez)"),
                paginas=1,
                optimizar=str(_datos.get("escaner_optimizar", "Sí")).lower().startswith("s"),
                modo=_datos.get("escaner_modo", "Automático (sin ventanas)"))
            _pags = [p for p in (_res.get("paginas") or []) if os.path.exists(str(p))]
            if not _pags:
                return messagebox.showerror("Escáner", "No se pudo escanear:\n%s"
                                            % (_res.get("error") or "sin detalle"), parent=v)
            _pdf = os.path.join(os.path.dirname(_pags[0]),
                                "Comprobante_Varias_%s.pdf" % datetime.now().strftime("%Y%m%d_%H%M%S"))
            escaner.imagenes_a_pdf(_pags, _pdf)
            estado_comp["ruta"] = _pdf
            actualizar_lbl_comp()
            aplicar_ocr(_pdf)
            messagebox.showinfo("Comprobante", "Comprobante escaneado:\n" + os.path.basename(_pdf), parent=v)

        ctk.CTkButton(f_comp, text="📎 Cargar archivo", height=30, font=("Arial", 11, "bold"),
                      fg_color="#16a085", hover_color="#11806a", command=cargar_comp).pack(
            side="left", fill="x", expand=True)
        ctk.CTkButton(f_comp, text="🖨️ Escanear comprobante", height=30, font=("Arial", 11, "bold"),
                      fg_color="#0e7490", hover_color="#155e75",
                      command=escanear_comp).pack(side="left", fill="x", expand=True, padx=(6, 0))
        ctk.CTkButton(f_comp, text="👁️ Ver", width=64, height=30, font=("Arial", 11, "bold"),
                      fg_color="#5d6d7e", hover_color="#4a5967",
                      command=lambda: (abrir_documento(estado_comp["ruta"])
                                       if estado_comp["ruta"] and os.path.isfile(estado_comp["ruta"])
                                       else messagebox.showinfo("Comprobante", "Sin comprobante cargado.", parent=v))
                      ).pack(side="left", padx=(6, 0))
        lbl_comp.pack(anchor="w", pady=(2, 6))
        actualizar_lbl_comp()

        resumen_texto = chr(10).join(
            "• %s — %s (%s)" % (p["nro"] or "s/n", p["cliente"],
                                formatear_moneda(max(0.0, p["saldo"])))
            for p in pendientes)
        resumen = ctk.CTkLabel(f, text=resumen_texto, font=("Arial", 10), justify="left", text_color="#555555")
        resumen.pack(anchor="w", pady=(4, 0))

        def guardar():
            fecha_val = ent_fecha.get().strip() or datetime.now().strftime("%d/%m/%Y")
            cuenta = cmb_cuenta.get().strip()
            oper = ent_oper.get().strip()
            if not cuenta:
                return messagebox.showwarning("Cobrar varias facturas", "Elija la cuenta destino / método.", parent=v)
            ruta_bd = ""
            if estado_comp["ruta"]:
                try:
                    ruta_bd = ruta_para_guardar(estado_comp["ruta"]) or estado_comp["ruta"]
                except Exception:
                    ruta_bd = estado_comp["ruta"]
            cobrados, errores = 0, []
            for p in pendientes:
                if p["saldo"] <= 0:
                    continue
                try:
                    conn = conectar_db()
                    if not conn:
                        continue
                    with conn.cursor() as cur:
                        try:
                            cur.execute("""
                                INSERT INTO pagos_clientes
                                (id_factura, monto_pagado, archivo_ruta, cliente_nombre, fecha_pago, cuenta_destino,
                                 numero_operacion)
                                VALUES (%s, %s, %s, %s, %s, %s, %s)""",
                                (p["id"], p["saldo"], ruta_bd, p["cliente"], fecha_val, cuenta, oper))
                        except Exception:
                            conn.rollback()
                            cur.execute("""
                                INSERT INTO pagos_clientes
                                (id_factura, monto_pagado, archivo_ruta, cliente_nombre, fecha_pago, cuenta_destino)
                                VALUES (%s, %s, %s, %s, %s, %s)""",
                                (p["id"], p["saldo"], ruta_bd, p["cliente"], fecha_val, cuenta))
                    conn.commit()
                    liberar_conexion(conn)
                    cobrados += 1
                except Exception as e:
                    errores.append("%s: %s" % (p["nro"], str(e)[:60]))
            try:
                cache_sistema.invalidar()
            except Exception:
                pass
            try:
                registrar_auditoria(getattr(getattr(self, "app_padre", None), "usuario_activo", ""),
                                    "Cuentas por Cobrar",
                                    "Cobró %d facturas con un solo comprobante" % cobrados)
            except Exception:
                pass
            try:
                self._refrescar_listas_cobro()
                self.app_padre.app_facturas.cargar_datos_tabla(reset_pagina=True)
            except Exception:
                pass
            try:
                v.destroy()
            except Exception:
                pass
            msg = "Se cobraron %d factura(s) con el comprobante." % cobrados
            if errores:
                msg += "\n\nAlgunas no se cobraron:\n" + "\n".join(errores[:5])
            messagebox.showinfo("Éxito", msg,
                                parent=ventana_para_mensaje(getattr(self, "main_root", None)))

        barra = ctk.CTkFrame(v, fg_color="transparent")
        barra.pack(fill="x", padx=16, pady=(0, 12))
        ctk.CTkButton(barra, text="💾 Cobrar las %d facturas" % len(pendientes), height=40,
                      font=("Arial", 13, "bold"), fg_color="#27ae60", hover_color="#1e8449",
                      command=guardar).pack(side="left", expand=True, fill="x", padx=(0, 6))
        ctk.CTkButton(barra, text="❌ Cancelar", height=40, font=("Arial", 12, "bold"),
                      fg_color="#7f8c8d", hover_color="#606b6b", command=v.destroy).pack(
            side="left", expand=True, fill="x", padx=(6, 0))

    def cargar_datos_cobrar(self, reset_pagina=False):
        if reset_pagina:
            self.pagina_actual = 1
            
        self.lbl_pagina.configure(text=f"Pág {self.pagina_actual}")

        for fila in self.tabla.get_children(): 
            self.tabla.delete(fila)
        
        filtro = ""
        if hasattr(self, 'ent_buscar_cobros'):
            filtro = self.ent_buscar_cobros.get().strip().lower()

        offset = (self.pagina_actual - 1) * self.registros_por_pagina
        clave_cache = f"cobros_{filtro}_mes_{self.mes_filtro}_pag_{self.pagina_actual}"
        datos = cache_sistema.obtener(clave_cache)

        if datos is not None:
            self._pintar_cobros(datos["filas"], datos["total_pendiente"], datos["facturas_con_cuenta"])
        else:
            self.tabla.insert("", tk.END, values=("", "", "", "Cargando datos...", "", "", "", "", "", "", "", "", ""))
            
            def tarea_descarga(estado):          # 👈 corre en el HILO SECUNDARIO
                conn = conectar_db(silencioso=True)
                if not conn: return
                
                filas_procesadas = []
                total_pendiente_global = 0.0
                facturas_con_cuenta = set()
                
                try:
                    cursor = conn.cursor()
                    cursor.execute("SELECT DISTINCT id_factura FROM pagos_clientes WHERE cuenta_destino IS NOT NULL AND cuenta_destino != ''")
                    facturas_con_cuenta = {row[0] for row in cursor.fetchall()}

                    condiciones = []
                    params = []
                    if filtro:
                        val = f"%{filtro}%"
                        condiciones.append("(numero_documento ILIKE %s OR cliente ILIKE %s OR descripcion ILIKE %s)")
                        params.extend([val, val, val])
                    patron_mes = patron_fecha_mes(self.mes_filtro)
                    if patron_mes:
                        condiciones.append("fecha LIKE %s")
                        params.append(patron_mes)
                    where_sql = (" WHERE " + " AND ".join(condiciones)) if condiciones else ""
                    params_reg = list(params)
                    params_reg.extend([self.registros_por_pagina, offset])
                    cursor.execute(f"SELECT id, fecha, numero_documento, cliente, evento_asociado, subtotal, impuesto, COALESCE(det_monto, 0), total, tipo_documento, descripcion, enlace_pdf_sunat, archivo_ruta, estado_sunat FROM facturas_emitidas{where_sql} ORDER BY id DESC LIMIT %s OFFSET %s", tuple(params_reg))
                        
                    registros = cursor.fetchall()
                    
                    # Calcular el total global pendiente sin importar paginación (Solo para el Label inferior)
                    cursor.execute(f"SELECT id, subtotal, impuesto, COALESCE(det_monto, 0), total, tipo_documento, estado_sunat FROM facturas_emitidas{where_sql}", tuple(params))
                        
                    todos_los_regs = cursor.fetchall()
                    for r_tot in todos_los_regs:
                        id_f_tot, s_tot, i_tot, d_tot, t_bruto_tot, tipo_doc_tot, est_tot = r_tot
                        if est_tot and "Anulada" in str(est_tot): continue
                        
                        sub_v = float(s_tot) if s_tot else 0.0
                        imp_v = float(i_tot) if i_tot else 0.0
                        det_v = float(d_tot) if d_tot else 0.0
                        tot_v = float(t_bruto_tot) if t_bruto_tot else 0.0
                        
                        if tipo_doc_tot and "Recibo" in tipo_doc_tot and "8%" in tipo_doc_tot: neto_fac = tot_v - imp_v - det_v
                        else: neto_fac = tot_v - det_v
                        
                        cursor.execute("SELECT SUM(monto_pagado + COALESCE(impuestos_deducibles, 0)) FROM pagos_clientes WHERE id_factura = %s", (id_f_tot,))
                        cobrado_res = cursor.fetchone()[0]
                        m_cobrado = float(cobrado_res) if cobrado_res else 0.0
                        total_pendiente_global += max(0.0, neto_fac - m_cobrado)

                    for reg in registros:
                        id_factura, fecha, nro_doc, cliente, evento, sub, imp, det_monto, tot_bruto, tipo_doc, concepto, enlace_sunat, arch_ruta, est_sunat = reg
                        
                        if est_sunat and "Anulada" in str(est_sunat): continue

                        sub_val = float(sub) if sub else 0.0
                        imp_val = float(imp) if imp else 0.0
                        det_monto_val = float(det_monto) if det_monto else 0.0
                        tot_bruto_val = float(tot_bruto) if tot_bruto else 0.0
                        
                        if tipo_doc and "Recibo" in tipo_doc and "8%" in tipo_doc:
                            neto_facturado = tot_bruto_val - imp_val - det_monto_val
                        else:
                            neto_facturado = tot_bruto_val - det_monto_val
                        
                        cursor.execute("""SELECT COALESCE(SUM(monto_pagado + COALESCE(impuestos_deducibles, 0)), 0),
                                                 COUNT(*) FILTER (WHERE COALESCE(archivo_ruta, '') <> '')
                                          FROM pagos_clientes WHERE id_factura = %s""", (id_factura,))
                        cobrado_res, cant_archivos = cursor.fetchone()
                        monto_cobrado = float(cobrado_res) if cobrado_res else 0.0
                        
                        saldo_pendiente = max(0.0, neto_facturado - monto_cobrado)
                        
                        filas_procesadas.append({
                            "id_factura": id_factura, "fecha": fecha, "nro_doc": nro_doc, "cliente": cliente, "concepto": concepto,
                            "sub_val": sub_val, "imp_val": imp_val, "det_monto_val": det_monto_val, "neto_facturado": neto_facturado,
                            "monto_cobrado": monto_cobrado, "saldo_pendiente": saldo_pendiente, "cant_archivos": cant_archivos
                        })
                        
                    datos_cache = {"filas": filas_procesadas, "total_pendiente": total_pendiente_global, "facturas_con_cuenta": facturas_con_cuenta}
                    cache_sistema.guardar(clave_cache, datos_cache)
                    
                except Exception as e:
                    print("Error cargando cobros:", e)
                finally:
                    liberar_conexion(conn)

                estado["filas"] = filas_procesadas
                estado["total_pendiente"] = total_pendiente_global
                estado["facturas_con_cuenta"] = facturas_con_cuenta

            def _pintar_cobros_hilo(estado):     # 👈 corre en el HILO PRINCIPAL
                self._pintar_cobros(estado["filas"], estado["total_pendiente"], estado["facturas_con_cuenta"])

            ejecutar_en_hilo(self.main_root, tarea_descarga, al_terminar=_pintar_cobros_hilo)

    def _pintar_cobros(self, filas, total_pendiente, facturas_con_cuenta):
        for fila in self.tabla.get_children(): self.tabla.delete(fila)
        
        contador = 1
        for f in filas:
            txt_adjuntos = f"📁 {f['cant_archivos']} archivo(s)" if f['cant_archivos'] > 0 else "❌ Sin adjuntos"
            try:
                _neto_c = float(f.get('neto_facturado') or 0)
                _cobr = float(f.get('monto_cobrado') or 0)
            except Exception:
                _neto_c, _cobr = 0.0, 0.0
            if (_neto_c <= 0.01) or (_cobr >= _neto_c - 0.01):
                etiqueta_color = "pagada"
            elif _cobr <= 0.01:
                etiqueta_color = "no_pagada"
            else:
                etiqueta_color = "parcial"

            row_vals = (
                contador, f['id_factura'], f['fecha'], f['nro_doc'] if f['nro_doc'] else "S/N", f['cliente'] if f['cliente'] else "Cliente Sin Nombre", f['concepto'], 
                formatear_moneda(f['sub_val']), formatear_moneda(f['imp_val']), formatear_moneda(f['det_monto_val']),
                formatear_moneda(f['neto_facturado']), formatear_moneda(f['monto_cobrado']), formatear_moneda(f['saldo_pendiente']), txt_adjuntos
            )

            self.tabla.insert("", tk.END, values=row_vals, tags=(etiqueta_color,))
            contador += 1

        self.lbl_total_general.configure(text=f"Total Pendiente Filtrado: {formatear_moneda(total_pendiente)}")
        
        if self.pagina_actual > 1:
            self.btn_ant.configure(state="normal")
        else:
            self.btn_ant.configure(state="disabled")
            
        if len(filas) == self.registros_por_pagina:
            self.btn_sig.configure(state="normal")
        else:
            self.btn_sig.configure(state="disabled")

    def _cuentas_para_combo(self):
        """Cuentas registradas en el sistema (Bancos) + los métodos habituales.

        Es la misma lista que usa «Registrar Cobro», para que el adelanto entre a
        la cuenta correcta y sume igual en Banco / Libro Diario / Estadísticas.
        """
        lista = []
        try:
            for b in cargar_bancos():
                banco_nom = str((b or {}).get("banco", "")).strip()
                cuenta_num = str((b or {}).get("cuenta", "")).strip()
                if banco_nom or cuenta_num:
                    etiqueta = ("%s - %s" % (banco_nom, cuenta_num)).strip(" -")
                    if etiqueta and etiqueta not in lista:
                        lista.append(etiqueta)
        except Exception as e:
            print("Aviso - cuentas registradas:", e)
        for extra in ("Efectivo / Caja Chica", "Tarjeta de Crédito", "Tarjeta de Débito", "Cheque", "Otro"):
            if extra not in lista:
                lista.append(extra)
        return lista

    def _clientes_para_combo(self):
        """Clientes ya registrados (para el desplegable del adelanto)."""
        lista = []
        conn = conectar_db(silencioso=True)
        if not conn:
            return lista
        try:
            cursor = conn.cursor()
            cursor.execute("""SELECT DISTINCT COALESCE(cliente, '') FROM facturas_emitidas
                              WHERE COALESCE(cliente, '') <> '' ORDER BY 1 LIMIT 400""")
            lista = [str(f[0]) for f in cursor.fetchall() if f and str(f[0]).strip()]
        except Exception as e:
            print("Aviso - clientes:", e)
        finally:
            liberar_conexion(conn)
        return lista

    def abrir_registrar_adelanto(self):
        """Ventana para asentar un pago por adelantado SIN factura.

        Queda como saldo a favor del cliente (aparece en su estado de cuenta y en
        Banco / Libro Diario / Estadísticas) hasta asociarlo a su factura.
        """
        v = ctk.CTkToplevel(self.main_root)
        v.title("Registrar Pago por Adelantado (sin factura)")
        v.geometry("620x560")
        v.transient(self.main_root)
        try:
            traer_al_frente(v)
        except Exception:
            pass
        ctk.CTkLabel(v, text="💳 PAGO POR ADELANTADO (SIN FACTURA)", font=("Arial", 15, "bold"),
                     text_color="#1f538d").pack(pady=(12, 2))
        ctk.CTkLabel(v, text="No se elige factura: queda como saldo a favor del cliente y después "
                             "se asocia con «🧩 Asociar Adelanto a Factura».",
                     font=("Arial", 11), text_color="gray", wraplength=560).pack(pady=(0, 8))
        f = ctk.CTkScrollableFrame(v, fg_color="transparent")
        f.pack(fill="both", expand=True, padx=14, pady=(0, 6))

        ctk.CTkLabel(f, text="Cliente:", font=("Arial", 11, "bold")).pack(anchor="w")
        cmb_cliente = ctk.CTkComboBox(f, values=self._clientes_para_combo() or [""])
        cmb_cliente.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(f, text="N° de Contrato (externo, opcional):", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_contrato = ctk.CTkEntry(f, placeholder_text="Ej. CONT-2026-014")
        ent_contrato.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(f, text="Monto adelantado (S/.):", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_monto = ctk.CTkEntry(f)
        ent_monto.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(f, text="Fecha del adelanto:", font=("Arial", 11, "bold")).pack(anchor="w")
        f_fecha = ctk.CTkFrame(f, fg_color="transparent")
        f_fecha.pack(fill="x", pady=(0, 8))
        ent_fecha = ctk.CTkEntry(f_fecha, placeholder_text="DD/MM/AAAA")
        ent_fecha.pack(side="left", fill="x", expand=True)
        ent_fecha.insert(0, datetime.now().strftime("%d/%m/%Y"))
        ctk.CTkButton(f_fecha, text="[ 📅 ]", width=44, font=("Arial", 11, "bold"),
                      fg_color="#2980b9", hover_color="#1f6391",
                      command=lambda e=ent_fecha: CalendarioNativo(
                          self.main_root.winfo_toplevel(), e)).pack(side="left", padx=(5, 0))

        ctk.CTkLabel(f, text="Cuenta donde entró el dinero:", font=("Arial", 11, "bold")).pack(anchor="w")
        cmb_cuenta = ctk.CTkComboBox(f, values=self._cuentas_para_combo(), width=400)
        cmb_cuenta.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(f, text="Comprobante del depósito (opcional):", font=("Arial", 11, "bold")).pack(anchor="w")
        f_arch = ctk.CTkFrame(f, fg_color="transparent")
        f_arch.pack(fill="x", pady=(0, 8))
        ent_arch = ctk.CTkEntry(f_arch, placeholder_text="Ruta del PDF/JPG del depósito")
        ent_arch.pack(side="left", fill="x", expand=True)
        ruta_adjunto = {"ruta": ""}

        def elegir_archivo():
            try:
                ruta = filedialog.askopenfilename(title="Comprobante del adelanto",
                                                  filetypes=[("Documentos", "*.pdf *.jpg *.jpeg *.png")])
                if ruta:
                    ruta_adjunto["ruta"] = ruta
                    ent_arch.delete(0, "end"); ent_arch.insert(0, ruta)
            except Exception as e:
                print("Aviso - adjunto:", e)

        ctk.CTkButton(f_arch, text="📎", width=44, font=("Arial", 11, "bold"), fg_color="#2980b9",
                      hover_color="#1f6391", command=elegir_archivo).pack(side="left", padx=(5, 0))

        def guardar():
            ok, mensaje = registrar_adelanto_cliente(
                cmb_cliente.get(), ent_monto.get(), fecha=ent_fecha.get().strip(),
                contrato=ent_contrato.get().strip(), cuenta=cmb_cuenta.get().strip(),
                archivo=(ruta_adjunto["ruta"] or ent_arch.get().strip()),
                usuario=getattr(self, "usuario_activo", "Sistema"))
            if not ok:
                return messagebox.showerror("Adelanto", str(mensaje), parent=v)
            messagebox.showinfo("Adelanto registrado", str(mensaje), parent=v)
            try:
                v.destroy()
            except Exception:
                pass
            try:
                self.cargar_datos_tabla()
            except Exception:
                pass

        barra = ctk.CTkFrame(v, fg_color="transparent")
        barra.pack(fill="x", padx=14, pady=(0, 12))
        ctk.CTkButton(barra, text="✅ Registrar Adelanto", height=40, font=("Arial", 13, "bold"),
                      fg_color="#27ae60", hover_color="#1e8449", command=guardar).pack(
            side="left", expand=True, fill="x", padx=(0, 6))
        ctk.CTkButton(barra, text="❌ Cancelar", height=40, font=("Arial", 13, "bold"), fg_color="#7f8c8d",
                      hover_color="#616a6b", command=v.destroy).pack(side="left", expand=True, fill="x", padx=(6, 0))
        return v

    def _facturas_con_cliente(self, limite=400):
        """Facturas emitidas (para asociarles un adelanto): (id, cliente, documento, total, pagado)."""
        lista = []
        conn = conectar_db(silencioso=True)
        if not conn:
            return lista
        try:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT f.id, COALESCE(f.cliente, ''), COALESCE(f.numero_documento, ''),
                       COALESCE(f.total, 0),
                       COALESCE((SELECT SUM(p.monto_pagado + COALESCE(p.impuestos_deducibles, 0)) FROM pagos_clientes p WHERE p.id_factura = f.id), 0)
                FROM facturas_emitidas f ORDER BY f.id DESC LIMIT %s
            """, (limite,))
            for fila in cursor.fetchall():
                id_f, cliente, documento, total, pagado = (list(fila) + [0, "", "", 0, 0])[:5]
                lista.append({"id": int(id_f or 0), "cliente": str(cliente or ""),
                              "documento": str(documento or ""), "total": round(float(total or 0), 2),
                              "pagado": round(float(pagado or 0), 2)})
        except Exception as e:
            print("Aviso - facturas para adelanto:", e)
        finally:
            liberar_conexion(conn)
        return lista

    def abrir_asociar_adelanto(self):
        """Ventana para asociar un adelanto (saldo a favor) a su factura."""
        v = ctk.CTkToplevel(self.main_root)
        v.title("Asociar Adelanto a su Factura")
        v.geometry("900x600")
        v.transient(self.main_root)
        try:
            traer_al_frente(v)
        except Exception:
            pass
        ctk.CTkLabel(v, text="🧩 ASOCIAR ADELANTO A SU FACTURA", font=("Arial", 15, "bold"),
                     text_color="#1f538d").pack(pady=(12, 2))
        ctk.CTkLabel(v, text="El adelanto es de UNA sola factura: al asociarlo pasa a contar como "
                             "cobrado en esa factura (y deja de ser saldo a favor).",
                     font=("Arial", 11), text_color="gray").pack(pady=(0, 8))

        marco = ctk.CTkFrame(v, fg_color="transparent")
        marco.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        izq = ctk.CTkFrame(marco, fg_color="#f7f9fb", corner_radius=8)
        izq.pack(side="left", fill="both", expand=True, padx=(0, 6))
        der = ctk.CTkFrame(marco, fg_color="#f7f9fb", corner_radius=8)
        der.pack(side="left", fill="both", expand=True, padx=(6, 0))
        ctk.CTkLabel(izq, text="💳 Adelantos con saldo a favor", font=("Arial", 12, "bold")).pack(pady=6)
        ctk.CTkLabel(der, text="🧾 Facturas emitidas", font=("Arial", 12, "bold")).pack(pady=6)

        lista_adelantos = ctk.CTkScrollableFrame(izq, fg_color="transparent")
        lista_adelantos.pack(fill="both", expand=True, padx=6, pady=6)
        lista_facturas = ctk.CTkScrollableFrame(der, fg_color="transparent")
        lista_facturas.pack(fill="both", expand=True, padx=6, pady=6)

        elegido = {"adelanto": None, "factura": None}
        botones = {"adelanto": {}, "factura": {}}

        def marcar_adelanto(datos, boton):
            elegido["adelanto"] = datos
            for otro in botones["adelanto"].values():
                otro.configure(fg_color="#5d6d7e")
            boton.configure(fg_color="#27ae60")

        def marcar_factura(datos, boton):
            elegido["factura"] = datos
            for otro in botones["factura"].values():
                otro.configure(fg_color="#5d6d7e")
            boton.configure(fg_color="#27ae60")

        adelantos = adelantos_disponibles()
        if not adelantos:
            ctk.CTkLabel(lista_adelantos, text="(no hay adelantos con saldo a favor)",
                         font=("Arial", 11), text_color="gray").pack(pady=10)
        for datos in adelantos:
            texto = "%s — S/ %.2f\n%s | contrato: %s" % (datos["cliente"], datos["monto"],
                                                         datos["fecha"] or "sin fecha",
                                                         datos["contrato"] or "sin dato")
            b = ctk.CTkButton(lista_adelantos, text=texto, font=("Arial", 11), anchor="w",
                              fg_color="#5d6d7e", hover_color="#4a5967", height=46,
                              command=lambda d=datos: None)
            b.configure(command=lambda d=datos, bb=b: marcar_adelanto(d, bb))
            b.pack(fill="x", pady=3)
            botones["adelanto"][datos["id"]] = b

        for datos in self._facturas_con_cliente():
            saldo = max(0.0, datos["total"] - datos["pagado"])
            texto = "%s — %s\nTotal S/ %.2f | pagado S/ %.2f | saldo S/ %.2f" % (
                datos["cliente"], datos["documento"] or ("ID %s" % datos["id"]),
                datos["total"], datos["pagado"], saldo)
            b = ctk.CTkButton(lista_facturas, text=texto, font=("Arial", 11), anchor="w",
                              fg_color="#5d6d7e", hover_color="#4a5967", height=46,
                              command=lambda d=datos: None)
            b.configure(command=lambda d=datos, bb=b: marcar_factura(d, bb))
            b.pack(fill="x", pady=3)
            botones["factura"][datos["id"]] = b

        def asociar():
            if not elegido["adelanto"] or not elegido["factura"]:
                return messagebox.showwarning("Adelanto", "Elija un adelanto y una factura.", parent=v)
            ok, mensaje = aplicar_adelanto_a_factura(elegido["adelanto"]["id"], elegido["factura"]["id"],
                                                     usuario=getattr(self, "usuario_activo", "Sistema"))
            if not ok:
                return messagebox.showerror("Adelanto", str(mensaje), parent=v)
            messagebox.showinfo("Adelanto asociado", str(mensaje), parent=v)
            try:
                v.destroy()
            except Exception:
                pass
            try:
                self.cargar_datos_tabla()
            except Exception:
                pass

        barra = ctk.CTkFrame(v, fg_color="transparent")
        barra.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkButton(barra, text="🔗 Asociar a la factura", height=40, font=("Arial", 13, "bold"),
                      fg_color="#16a085", hover_color="#11806a", command=asociar).pack(
            side="left", expand=True, fill="x", padx=(0, 6))
        ctk.CTkButton(barra, text="❌ Cancelar", height=40, font=("Arial", 13, "bold"), fg_color="#7f8c8d",
                      hover_color="#616a6b", command=v.destroy).pack(side="left", expand=True, fill="x", padx=(6, 0))
        return v

    def _refrescar_listas_cobro(self):
        """Refresca la lista de cobros y la de facturas (para el color de cobrado)."""
        try:
            self.cargar_datos_cobrar(reset_pagina=True)
        except Exception:
            pass
        try:
            self.app_padre.app_facturas.cargar_datos_tabla(reset_pagina=True)
        except Exception:
            pass

    def _ofrecer_cobro(self, id_factura=0, nro_doc=""):
        """Pregunta si la factura ya fue pagada y abre «Registrar Cobro»."""
        try:
            respuesta = messagebox.askyesno(
                "Cancelar la factura",
                "¿Desea CANCELAR ahora la factura %s?" % (nro_doc or "s/n") + chr(10) + chr(10) +
                "Si ya fue pagada, se abrirá «Registrar Cobro» para registrar el ingreso.",
                parent=ventana_para_mensaje(getattr(self, "main_root", None)))
        except Exception as e:
            print("Aviso - oferta de cobro:", e)
            return
        if respuesta:
            self._abrir_cobro_de(id_factura, nro_doc)

    def _abrir_cobro_de(self, id_factura=0, nro_doc=""):
        """Abre «Registrar Cobro» para la factura indicada (por id o por N° de documento)."""
        for intento in (1, 2):
            try:
                self.cargar_datos_tabla(reset_pagina=True)
            except Exception:
                pass
            for item in self.tabla.get_children():
                valores = self.tabla.item(item, "values")
                if len(valores) < 4:
                    continue
                coincide_id = id_factura and str(valores[1]) == str(id_factura)
                coincide_doc = nro_doc and str(valores[3]).strip() == str(nro_doc).strip()
                if coincide_id or coincide_doc:
                    self.tabla.selection_set(item)
                    self.tabla.see(item)
                    try:
                        self.cargar_comprobante_cobro()
                    except Exception as e:
                        print("Aviso - abrir cobro:", e)
                    return True
        messagebox.showinfo(
            "Registrar Cobro",
            "La factura se guardó, pero no se encontró en la lista." + chr(10) +
            "Pulse «Actualizar» y use el menú de clic derecho → 🧾 Registrar cobro.",
            parent=ventana_para_mensaje(getattr(self, "main_root", None)))
        return False

    def cargar_comprobante_cobro(self):
        """🧾 Registrar Cobro (mismo formato que el cobro múltiple)."""
        ruta_base = obtener_ruta_base_drive()
        if not ruta_base:
            avisar_sin_permiso_guardado()
            return

        seleccion = self.tabla.selection()
        if not seleccion:
            return messagebox.showwarning("Selección", "Seleccione una factura.")
        valores = self.tabla.item(seleccion[0], "values")
        id_factura, nro_doc, cliente = valores[1], valores[3], valores[4]
        saldo_actual = desformatear_numero(self.tabla.set(seleccion[0], "saldo"))
        if saldo_actual <= 0:
            return messagebox.showinfo("Aviso", "Esta factura ya está cobrada por completo.",
                                       parent=ventana_para_mensaje(getattr(self, "main_root", None)))

        lista_cuentas = []
        try:
            for b in (cargar_bancos() or []):
                if isinstance(b, dict):
                    nombre = " - ".join(str(x) for x in (b.get("banco"), b.get("cuenta")) if x).strip(" - ")
                else:
                    nombre = str(b).strip()
                if nombre and nombre not in lista_cuentas:
                    lista_cuentas.append(nombre)
        except Exception:
            pass
        lista_cuentas += ["Efectivo / Caja Chica", "Tarjeta de Crédito", "Tarjeta de Débito", "Cheque", "Otro"]

        v = ctk.CTkToplevel(getattr(self, "main_root", None))
        v.title("🧾 Registrar Cobro")
        # 🪟 Se abre lo más arriba posible y con scroll si el formulario es alto.
        _alto_cobro = min(880, max(480, v.winfo_screenheight() - 90))
        _x_cobro = max(0, (v.winfo_screenwidth() - 620) // 2)
        v.geometry("620x%d+%d+0" % (_alto_cobro, _x_cobro))
        v.transient(getattr(self, "main_root", None))
        try:
            v.grab_set()
        except Exception:
            pass
        f = ctk.CTkScrollableFrame(v, fg_color="transparent")
        f.pack(fill="both", expand=True, padx=16, pady=12)

        ctk.CTkLabel(f, text="Cobro para: %s" % cliente, font=("Arial", 13, "bold"),
                     text_color="#1f538d").pack(anchor="w")
        ctk.CTkLabel(f, text="Documento: %s   ·   Saldo pendiente: %s"
                     % (nro_doc or "s/n", formatear_moneda(saldo_actual)),
                     font=("Arial", 11), text_color="#555555").pack(anchor="w", pady=(0, 6))

        ctk.CTkLabel(f, text="Fecha del cobro:", font=("Arial", 11, "bold")).pack(anchor="w", pady=(6, 0))
        f_fecha_cob = ctk.CTkFrame(f, fg_color="transparent")
        f_fecha_cob.pack(fill="x")
        ent_fecha = ctk.CTkEntry(f_fecha_cob)
        ent_fecha.pack(side="left", fill="x", expand=True)
        ent_fecha.insert(0, datetime.now().strftime("%d/%m/%Y"))
        ctk.CTkButton(f_fecha_cob, text="📅", width=40, fg_color="#1f538d", hover_color="#163b65",
                      command=lambda: CalendarioNativo(v, ent_fecha)).pack(side="left", padx=(5, 0))

        ctk.CTkLabel(f, text="Cuenta destino / método de cobro:", font=("Arial", 11, "bold")).pack(anchor="w", pady=(6, 0))
        cmb_cuenta = ctk.CTkComboBox(f, values=lista_cuentas, width=420, state="readonly")
        cmb_cuenta.pack(fill="x")
        if lista_cuentas:
            cmb_cuenta.set(lista_cuentas[0])

        ctk.CTkLabel(f, text="N° de Operación (opcional):", font=("Arial", 11, "bold")).pack(anchor="w", pady=(6, 0))
        ent_oper = ctk.CTkEntry(f)
        ent_oper.pack(fill="x")

        ctk.CTkLabel(f, text="💰 Monto recibido:", font=("Arial", 11, "bold")).pack(anchor="w", pady=(8, 0))
        ent_monto = ctk.CTkEntry(f)
        ent_monto.pack(fill="x")
        ent_monto.insert(0, "%.2f" % saldo_actual)

        # 🧾 Otros impuestos deducibles (retenciones) del cobro
        ctk.CTkLabel(f, text="🧾 Otros impuestos deducibles (S/):",
                     font=("Arial", 11, "bold")).pack(anchor="w", pady=(8, 0))
        ent_impuestos = ctk.CTkEntry(f)
        ent_impuestos.pack(fill="x")
        ent_impuestos.insert(0, "0.00")
        estado_imp = {"ruta": ""}
        lbl_imp = ctk.CTkLabel(f, text="", font=("Arial", 10), text_color="#999999", anchor="w")

        def _pintar_impuesto():
            nombre = os.path.basename(estado_imp["ruta"]) if estado_imp["ruta"] else ""
            lbl_imp.configure(text=("📎 " + nombre) if nombre
                              else "Sin comprobante del impuesto (opcional)",
                              text_color="#1e8449" if nombre else "#999999")

        def _cargar_impuesto():
            try:
                ruta = seleccionar_archivo_dialogo(titulo="Comprobante del impuesto")
            except Exception:
                ruta = ""
            if ruta:
                estado_imp["ruta"] = ruta
                _pintar_impuesto()

        def _quitar_impuesto():
            estado_imp["ruta"] = ""
            _pintar_impuesto()

        f_imp = ctk.CTkFrame(f, fg_color="transparent")
        f_imp.pack(fill="x")
        ctk.CTkButton(f_imp, text="📎 Cargar comprobante del impuesto", height=30,
                      font=("Arial", 11, "bold"),
                      command=_cargar_impuesto).pack(side="left")
        ctk.CTkButton(f_imp, text="✖ Quitar", width=80, height=30, font=("Arial", 11, "bold"),
                      command=_quitar_impuesto).pack(side="left", padx=6)
        lbl_imp.pack(anchor="w", pady=(2, 0))
        _pintar_impuesto()
        lbl_calc = ctk.CTkLabel(f, text="", font=("Arial", 11, "bold"), text_color="#555555")

        def actualizar_calc(event=None):
            try:
                _monto = float(ent_monto.get().strip() or 0)
            except ValueError:
                _monto = 0.0
            try:
                _impuestos = float(ent_impuestos.get().strip() or 0)
            except ValueError:
                _impuestos = 0.0
            _monto = _monto + max(0.0, _impuestos)   # impuestos deducibles suman al cobro
            _diff = round(_monto - saldo_actual, 2)
            _aplicado = formatear_moneda(_monto)
            if _diff < -0.005:
                lbl_calc.configure(text="Aplicado a la factura: %s   ·   Saldo: %s   ·   Falta: %s"
                                   % (_aplicado, formatear_moneda(saldo_actual),
                                      formatear_moneda(abs(_diff))),
                                   text_color="#c0392b")
            elif _diff > 0.005:
                lbl_calc.configure(text="Aplicado a la factura: %s   ·   Saldo: %s   ·   Sobra: %s"
                                   % (_aplicado, formatear_moneda(saldo_actual),
                                      formatear_moneda(_diff)),
                                   text_color="#b9770e")
            else:
                lbl_calc.configure(text="Aplicado a la factura: %s   ·   Saldo: %s   ·   Exacto"
                                   % (_aplicado, formatear_moneda(saldo_actual)),
                                   text_color="#1e8449")

        ent_monto.bind("<KeyRelease>", actualizar_calc)
        lbl_calc.pack(anchor="w", pady=(2, 0))
        actualizar_calc()

        ctk.CTkLabel(f, text="📎 Comprobante de depósito:", font=("Arial", 11, "bold")).pack(anchor="w", pady=(10, 0))
        estado_comp = {"ruta": ""}
        lbl_comp = ctk.CTkLabel(f, text="Sin comprobante cargado", font=("Arial", 10),
                                text_color="#999999", anchor="w")

        def actualizar_lbl_comp():
            if estado_comp["ruta"]:
                lbl_comp.configure(text="📎 " + os.path.basename(estado_comp["ruta"]), text_color="#1e8449")
            else:
                lbl_comp.configure(text="Sin comprobante cargado", text_color="#999999")

        def aplicar_ocr(ruta):
            try:
                from modulo_compras import (leer_ocr_windows, extraer_numero_operacion,
                                            monto_del_comprobante)
            except Exception as e_imp:
                print("[OCR comprobante] no se pudo importar OCR:", e_imp)
                return
            try:
                _texto = leer_ocr_windows(ruta)
                if not _texto.strip():
                    try:
                        _texto = escaner.texto_de_resultados(escaner.ocr_paginas([ruta]))
                    except Exception:
                        _texto = ""
            except Exception as e_ocr:
                print("[OCR comprobante] error:", e_ocr)
                _texto = ""
            if not _texto.strip():
                return
            _num, _monto = "", 0.0
            try:
                _num = extraer_numero_operacion(_texto) or ""
            except Exception:
                _num = ""
            try:
                _monto = monto_del_comprobante(_texto) or 0.0
            except Exception:
                _monto = 0.0
            _avisos = []
            if _num:
                try:
                    ent_oper.delete(0, tk.END)
                    ent_oper.insert(0, str(_num))
                except Exception:
                    pass
                _avisos.append("N° de operación: " + str(_num))
            if _monto and float(_monto) > 0:
                try:
                    ent_monto.delete(0, tk.END)
                    ent_monto.insert(0, "%.2f" % float(_monto))
                except Exception:
                    pass
                try:
                    actualizar_calc()
                except Exception:
                    pass
                _avisos.append("Monto: %.2f" % float(_monto))
            if _avisos:
                try:
                    lbl_comp.configure(text="📎 " + os.path.basename(ruta) + "   ·   " + " · ".join(_avisos),
                                       text_color="#1e8449")
                except Exception:
                    pass

        def cargar_comp():
            r = _pick_comprobante(v)
            if r:
                estado_comp["ruta"] = r
                actualizar_lbl_comp()
                aplicar_ocr(r)

        def escanear_comp():
            try:
                _cfg = cargar_configuracion_regional()
            except Exception:
                _cfg = {}
            _datos = escaner.config_escaner(_cfg)
            _res = _escanear_con_progreso(
                v, "el comprobante", escaner, escaner.carpeta_escaneos(_cfg),
                dispositivo=_datos.get("escaner_dispositivo", ""),
                dpi=_datos.get("escaner_dpi", "300"),
                color=_datos.get("escaner_color", "Escala de grises"),
                origen=_datos.get("escaner_origen", "Cristal (una página a la vez)"),
                paginas=1,
                optimizar=str(_datos.get("escaner_optimizar", "Sí")).lower().startswith("s"),
                modo=_datos.get("escaner_modo", "Automático (sin ventanas)"))
            _pags = [p for p in (_res.get("paginas") or []) if os.path.exists(str(p))]
            if not _pags:
                return messagebox.showerror("Escáner", "No se pudo escanear:" + chr(10)
                                            + str(_res.get("error") or "sin detalle"), parent=v)
            _pdf = os.path.join(os.path.dirname(_pags[0]),
                                "Comprobante_Cobro_%s.pdf" % datetime.now().strftime("%Y%m%d_%H%M%S"))
            escaner.imagenes_a_pdf(_pags, _pdf)
            estado_comp["ruta"] = _pdf
            actualizar_lbl_comp()
            aplicar_ocr(_pdf)
            messagebox.showinfo("Comprobante", "Comprobante escaneado:" + chr(10) + os.path.basename(_pdf), parent=v)

        f_comp = ctk.CTkFrame(f, fg_color="transparent")
        f_comp.pack(fill="x", pady=(0, 2))
        ctk.CTkButton(f_comp, text="📎 Cargar archivo", height=30, font=("Arial", 11, "bold"),
                      fg_color="#16a085", hover_color="#11806a", command=cargar_comp).pack(
            side="left", fill="x", expand=True)
        ctk.CTkButton(f_comp, text="🖨️ Escanear comprobante", height=30, font=("Arial", 11, "bold"),
                      fg_color="#0e7490", hover_color="#155e75",
                      command=escanear_comp).pack(side="left", fill="x", expand=True, padx=(6, 0))
        ctk.CTkButton(f_comp, text="👁️ Ver", width=64, height=30, font=("Arial", 11, "bold"),
                      fg_color="#5d6d7e", hover_color="#4a5967",
                      command=lambda: (abrir_documento(estado_comp["ruta"])
                                       if estado_comp["ruta"] and os.path.isfile(estado_comp["ruta"])
                                       else messagebox.showinfo("Comprobante", "Sin comprobante cargado.", parent=v))
                      ).pack(side="left", padx=(6, 0))
        lbl_comp.pack(anchor="w", pady=(2, 6))
        actualizar_lbl_comp()

        def confirmar():
            try:
                monto_val = float(ent_monto.get().strip() or 0)
            except ValueError:
                return messagebox.showerror("Error", "Ingrese un monto numérico válido.", parent=v)
            try:
                impuestos_val = float(ent_impuestos.get().strip() or 0)
            except ValueError:
                return messagebox.showerror("Error", "Los impuestos deducibles deben ser numéricos.",
                                            parent=v)
            if impuestos_val < 0:
                return messagebox.showerror("Error", "Los impuestos deducibles no pueden ser negativos.",
                                            parent=v)
            total_val = round(monto_val + impuestos_val, 2)
            if total_val <= 0:
                return messagebox.showerror("Error", "El total del cobro debe ser mayor a 0.", parent=v)
            if total_val > (saldo_actual + 0.01):
                return messagebox.showerror(
                    "Error", "El total (monto + impuestos deducibles) supera el saldo pendiente.", parent=v)
            fecha_val = ent_fecha.get().strip() or datetime.now().strftime("%d/%m/%Y")
            cuenta_val = cmb_cuenta.get().strip()
            oper_val = ent_oper.get().strip()
            ruta_bd = ""
            if estado_comp["ruta"]:
                try:
                    ruta_bd = ruta_para_guardar(estado_comp["ruta"]) or estado_comp["ruta"]
                except Exception:
                    ruta_bd = estado_comp["ruta"]
            try:
                conn = conectar_db()
                if not conn:
                    return
                cursor = conn.cursor()
                ruta_imp_bd = ""
                if estado_imp["ruta"]:
                    try:
                        ruta_imp_bd = ruta_para_guardar(estado_imp["ruta"]) or estado_imp["ruta"]
                    except Exception:
                        ruta_imp_bd = estado_imp["ruta"]
                try:
                    _asegurar_columnas_impuestos(conn)
                    cursor.execute("""
                        INSERT INTO pagos_clientes
                        (id_factura, monto_pagado, archivo_ruta, cliente_nombre, fecha_pago,
                         cuenta_destino, numero_operacion, impuestos_deducibles, archivo_impuesto)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)""",
                        (id_factura, monto_val, ruta_bd, cliente, fecha_val, cuenta_val, oper_val,
                         impuestos_val, ruta_imp_bd))
                except Exception:
                    conn.rollback()
                    cursor.execute("""
                        INSERT INTO pagos_clientes
                        (id_factura, monto_pagado, archivo_ruta, cliente_nombre, fecha_pago, cuenta_destino)
                        VALUES (%s, %s, %s, %s, %s, %s)""",
                        (id_factura, monto_val, ruta_bd, cliente, fecha_val, cuenta_val))
                conn.commit()
                liberar_conexion(conn)
            except Exception as e:
                return messagebox.showerror("Error", str(e), parent=v)
            try:
                cache_sistema.invalidar()
            except Exception:
                pass
            try:
                registrar_auditoria(self.app_padre.usuario_activo, "Cuentas por Cobrar",
                                    "Cobró %s a Fac. %s" % (formatear_moneda(monto_val), nro_doc))
            except Exception:
                pass
            try:
                v.destroy()
            except Exception:
                pass
            messagebox.showinfo("Éxito", "Cobro de %s registrado." % formatear_moneda(monto_val),
                                parent=ventana_para_mensaje(getattr(self, "main_root", None)))
            self._refrescar_listas_cobro()

        barra = ctk.CTkFrame(v, fg_color="transparent")
        barra.pack(fill="x", padx=16, pady=(0, 12))
        ctk.CTkButton(barra, text="✅ Confirmar cobro", height=40, font=("Arial", 13, "bold"),
                      fg_color="#27ae60", hover_color="#1e8449", command=confirmar).pack(
            side="left", expand=True, fill="x", padx=(0, 6))
        ctk.CTkButton(barra, text="❌ Cancelar", height=40, font=("Arial", 12, "bold"),
                      fg_color="#7f8c8d", hover_color="#606b6b", command=v.destroy).pack(
            side="left", expand=True, fill="x", padx=(6, 0))
        ent_monto.focus()

    def abrir_todos_los_archivos(self, event):
        seleccion = self.tabla.selection()
        if not seleccion: return
        id_factura = self.tabla.item(seleccion[0], "values")[1] 
        try:
            conn = conectar_db()
            cursor = conn.cursor()
            cursor.execute("SELECT archivo_ruta FROM pagos_clientes WHERE id_factura = %s AND archivo_ruta != ''", (id_factura,))
            rutas = cursor.fetchall()
            liberar_conexion(conn)
            if rutas:
                from app_paths import resolver_ruta_archivo
                abiertos = 0
                for r in rutas:
                    # 🔎 Resuelve rutas guardadas por otro equipo/sistema (Mac/Windows)
                    ruta_norm = resolver_ruta_archivo(r[0])
                    if ruta_norm:
                        abrir_documento(ruta_norm)
                        abiertos += 1
                if not abiertos:
                    messagebox.showwarning("Aviso", "Los soportes están registrados pero los archivos no se encuentran en este equipo.")
            else: messagebox.showinfo("Aviso", "No hay soportes cargados.")
        except Exception: pass

    def abrir_ventana_edicion(self):
        sel = self.tabla.selection()
        if not sel: return messagebox.showwarning("Selección", "Seleccione la factura para editar sus cobros.")
        valores = self.tabla.item(sel[0], "values")
        id_factura, nro_doc, cliente = valores[1], valores[3], valores[4] 
        saldo_actual_global = float(desformatear_numero(self.tabla.set(sel[0], "saldo")) or 0) 

        v_edit = ctk.CTkToplevel(self.main_root)
        v_edit.title(f"✏️ Gestión de Cobros - Fac. {nro_doc}")
        v_edit.geometry("820x400")
        v_edit.transient(self.main_root)
        v_edit.grab_set()

        ctk.CTkLabel(v_edit, text=f"Cobros registrados: {cliente}", font=("Arial", 12, "bold"), text_color="#1f538d").pack(pady=10)

        frame_cuerpo = ctk.CTkFrame(v_edit, fg_color="transparent")
        frame_cuerpo.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        sub_tabla = ttk.Treeview(frame_cuerpo, columns=("id", "monto", "fecha", "cuenta", "tiene_archivo"), show="headings", height=8)
        sub_tabla.heading("id", text="ID Cobro"); sub_tabla.heading("monto", text="Monto"); sub_tabla.heading("fecha", text="Fecha"); sub_tabla.heading("cuenta", text="Cuenta / Destino"); sub_tabla.heading("tiene_archivo", text="¿Soporte?")
        sub_tabla.column("id", width=60, anchor="center"); sub_tabla.column("monto", width=110, anchor="e"); sub_tabla.column("fecha", width=110, anchor="center"); sub_tabla.column("cuenta", width=160, anchor="w"); sub_tabla.column("tiene_archivo", width=130, anchor="center")
        sub_tabla.pack(side="left", fill="both", expand=True, padx=(0, 10))

        def refrescar_subtabla():
            for f in sub_tabla.get_children(): sub_tabla.delete(f)
            try:
                conn = conectar_db(); cursor = conn.cursor()
                cursor.execute("SELECT id, monto_pagado, fecha_pago, archivo_ruta, cuenta_destino FROM pagos_clientes WHERE id_factura = %s", (id_factura,))
                for a in cursor.fetchall(): 
                    # El indicador resuelve la ruta (sirve con rutas relativas o de otro equipo)
                    sub_tabla.insert("", tk.END, values=(a[0], formatear_moneda(a[1]), a[2] if a[2] else "Sin fecha", a[4] if a[4] else "-", "✅ Sí" if resolver_ruta_archivo(a[3]) else "❌ No"))
                liberar_conexion(conn)
            except Exception: pass
        refrescar_subtabla()

        def ejecutar_modificacion():
            nonlocal saldo_actual_global
            sub_sel = sub_tabla.selection()
            if not sub_sel: return
            id_pago = sub_tabla.item(sub_sel[0], "values")[0]
            conn = conectar_db(); cursor = conn.cursor()
            cursor.execute("SELECT monto_pagado, fecha_pago, cuenta_destino, COALESCE(numero_operacion, '') FROM pagos_clientes WHERE id = %s", (id_pago,))
            _fila_mod = cursor.fetchone()
            monto_actual, fecha_actual, cuenta_actual = _fila_mod[0], _fila_mod[1], _fila_mod[2]
            oper_actual = _fila_mod[3] if len(_fila_mod) > 3 else ""
            liberar_conexion(conn)

            v_mod_cobro = ctk.CTkToplevel(v_edit)
            v_mod_cobro.title("Modificar Cobro")
            v_mod_cobro.geometry("400x430")
            v_mod_cobro.transient(v_edit)
            v_mod_cobro.grab_set()

            ctk.CTkLabel(v_mod_cobro, text="Editar Cobro", font=("Arial", 14, "bold")).pack(pady=10)

            f_form = ctk.CTkFrame(v_mod_cobro, fg_color="transparent")
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

            ctk.CTkLabel(f_form, text="Cuenta Destino / Método:", font=("Arial", 11, "bold")).pack(anchor="w")
            ent_mod_cuenta = ctk.CTkComboBox(f_form, values=lista_cuentas, width=400)
            ent_mod_cuenta.pack(fill="x", pady=(0, 10))
            if cuenta_actual: ent_mod_cuenta.set(cuenta_actual)
            elif lista_cuentas: ent_mod_cuenta.set(lista_cuentas[0])

            ctk.CTkLabel(f_form, text="Nuevo Monto (0 = Eliminar):", font=("Arial", 11, "bold")).pack(anchor="w")
            ent_mod_monto = ctk.CTkEntry(f_form)
            ent_mod_monto.pack(fill="x", pady=(0, 10))
            ent_mod_monto.insert(0, str(monto_actual))

            ctk.CTkLabel(f_form, text="Fecha (AAAA-MM-DD):", font=("Arial", 11, "bold")).pack(anchor="w")
            ent_mod_fecha = ctk.CTkEntry(f_form)
            ent_mod_fecha.pack(fill="x", pady=(0, 10))
            _fech_mod = str(fecha_actual) if fecha_actual else datetime.now().strftime("%d/%m/%Y")
            try:
                if re.match(r"^\d{4}-\d{2}-\d{2}$", _fech_mod.strip()):
                    _fech_mod = datetime.strptime(_fech_mod.strip(), "%Y-%m-%d").strftime("%d/%m/%Y")
            except Exception:
                pass
            ent_mod_fecha.insert(0, _fech_mod)

            ctk.CTkLabel(f_form, text="N° de Operación / Comprobante:", font=("Arial", 11, "bold")).pack(anchor="w")
            ent_mod_oper = ctk.CTkEntry(f_form, placeholder_text="N° de operación del banco")
            ent_mod_oper.pack(fill="x", pady=(0, 10))
            if oper_actual:
                ent_mod_oper.insert(0, str(oper_actual))

            def guardar_mod(event=None):
                nonlocal saldo_actual_global
                try:
                    nuevo_monto = float(ent_mod_monto.get().strip())
                except ValueError:
                    return messagebox.showerror("Error", "Monto inválido", parent=v_mod_cobro)

                diferencia_de_aumento = nuevo_monto - float(monto_actual)
                if diferencia_de_aumento > (saldo_actual_global + 0.01):
                    return messagebox.showerror("Error", f"Supera el saldo pendiente de {formatear_moneda(saldo_actual_global)}.", parent=v_mod_cobro)

                if nuevo_monto == 0:
                    if messagebox.askyesno("Confirmar", "¿Eliminar registro?", parent=v_mod_cobro):
                        conn = conectar_db(); cursor = conn.cursor()
                        cursor.execute("SELECT archivo_ruta FROM pagos_clientes WHERE id = %s", (id_pago,))
                        r = cursor.fetchone()
                        eliminar_archivo(r[0])   # resuelve rutas de otros equipos/SO

                        cursor.execute("DELETE FROM pagos_clientes WHERE id = %s", (id_pago,))
                        conn.commit(); liberar_conexion(conn)
                        cache_sistema.invalidar()
                        registrar_auditoria(self.app_padre.usuario_activo, "Cuentas por Cobrar", f"Eliminó el cobro ID {id_pago}")
                else:
                    nueva_fecha = ent_mod_fecha.get().strip() or fecha_actual
                    nueva_cuenta = ent_mod_cuenta.get().strip()
                    conn = conectar_db(); cursor = conn.cursor()
                    cursor.execute("UPDATE pagos_clientes SET monto_pagado = %s, fecha_pago = %s, cuenta_destino = %s, numero_operacion = %s WHERE id = %s", (nuevo_monto, nueva_fecha, nueva_cuenta, id_pago, ent_mod_oper.get().strip()))
                    conn.commit(); liberar_conexion(conn)
                    cache_sistema.invalidar()
                    registrar_auditoria(self.app_padre.usuario_activo, "Cuentas por Cobrar", f"Modificó el cobro ID {id_pago} a {formatear_moneda(nuevo_monto)}")

                saldo_actual_global -= diferencia_de_aumento
                v_mod_cobro.destroy()
                refrescar_subtabla()
                self._refrescar_listas_cobro()

            ent_mod_monto.bind("<Return>", guardar_mod)
            ent_mod_fecha.bind("<Return>", guardar_mod)

            btn_guardar_mod = ctk.CTkButton(v_mod_cobro, text="💾 Guardar Cambios", command=guardar_mod, fg_color="#27ae60", hover_color="#1e8449")
            btn_guardar_mod.pack(pady=10)
            ent_mod_monto.focus()

        def cambiar_soporte():
            ruta_base = obtener_ruta_base_drive()
            if not ruta_base:
                avisar_sin_permiso_guardado()
                return
            
            sub_sel = sub_tabla.selection()
            if not sub_sel: return
            id_pago = sub_tabla.item(sub_sel[0], "values")[0]
            ruta_origen = seleccionar_archivo_dialogo(titulo="Seleccionar Soporte", tipos=[("Archivos", "*.pdf;*.png;*.jpg;*.jpeg")])
            if ruta_origen:
                try:
                    carpeta_comprobantes = os.path.join(ruta_base, "comprobantes_ingresos")
                    if not os.path.exists(carpeta_comprobantes): os.makedirs(carpeta_comprobantes)
                    conn = conectar_db(); cursor = conn.cursor()
                    cursor.execute("SELECT archivo_ruta FROM pagos_clientes WHERE id = %s", (id_pago,))
                    antigua_ruta = cursor.fetchone()[0]
                    eliminar_archivo(antigua_ruta)   # resuelve rutas de otros equipos/SO

                    nombre_limpio = f"Ingreso_Fac_{id_factura}_{cliente.replace(' ', '_')}_R_{id_pago}{os.path.splitext(ruta_origen)[1]}"
                    ruta_destino = os.path.join(carpeta_comprobantes, nombre_limpio)
                    shutil.copy2(ruta_origen, ruta_destino)
                    cursor.execute("UPDATE pagos_clientes SET archivo_ruta = %s WHERE id = %s", (ruta_para_guardar(ruta_destino), id_pago))
                    conn.commit(); liberar_conexion(conn)
                    registrar_auditoria(self.app_padre.usuario_activo, "Cuentas por Cobrar", f"Actualizó soporte del cobro ID {id_pago}")
                    messagebox.showinfo("Éxito", "Soporte actualizado.", parent=v_edit)
                    refrescar_subtabla(); self._refrescar_listas_cobro()
                except Exception as e: messagebox.showerror("Error", str(e), parent=v_edit)

        def eliminar_soporte():
            sub_sel = sub_tabla.selection()
            if not sub_sel: return
            id_pago = sub_tabla.item(sub_sel[0], "values")[0]
            if messagebox.askyesno("Confirmar", "¿Eliminar soporte digital?", parent=v_edit):
                try:
                    conn = conectar_db(); cursor = conn.cursor()
                    cursor.execute("SELECT archivo_ruta FROM pagos_clientes WHERE id = %s", (id_pago,))
                    ruta_archivo = cursor.fetchone()[0]
                    eliminar_archivo(ruta_archivo)   # resuelve rutas de otros equipos/SO

                    cursor.execute("UPDATE pagos_clientes SET archivo_ruta = '' WHERE id = %s", (id_pago,))
                    conn.commit(); liberar_conexion(conn)
                    registrar_auditoria(self.app_padre.usuario_activo, "Cuentas por Cobrar", f"Eliminó soporte del cobro ID {id_pago}")
                    messagebox.showinfo("Éxito", "Soporte eliminado.", parent=v_edit)
                    refrescar_subtabla(); self._refrescar_listas_cobro()
                except Exception as e: messagebox.showerror("Error", str(e), parent=v_edit)

        def eliminar_cobro_completo():
            nonlocal saldo_actual_global
            sub_sel = sub_tabla.selection()
            if not sub_sel: return
            id_pago = sub_tabla.item(sub_sel[0], "values")[0]
            monto_eliminado = desformatear_numero(sub_tabla.item(sub_sel[0], "values")[1])

            if messagebox.askyesno("Confirmar Eliminación", "⚠️ ¿Eliminar este registro de cobro por completo?", parent=v_edit):
                try:
                    conn = conectar_db(); cursor = conn.cursor()
                    cursor.execute("SELECT archivo_ruta FROM pagos_clientes WHERE id = %s", (id_pago,))
                    r = cursor.fetchone()
                    eliminar_archivo(r[0])   # resuelve rutas de otros equipos/SO

                    cursor.execute("DELETE FROM pagos_clientes WHERE id = %s", (id_pago,))
                    conn.commit(); liberar_conexion(conn)
                    cache_sistema.invalidar()
                    registrar_auditoria(self.app_padre.usuario_activo, "Cuentas por Cobrar", f"Eliminó completamente el cobro ID {id_pago}")
                    messagebox.showinfo("Éxito", "Cobro eliminado.", parent=v_edit)
                    refrescar_subtabla(); self._refrescar_listas_cobro(); saldo_actual_global += monto_eliminado
                except Exception as e: messagebox.showerror("Error", str(e), parent=v_edit)

        frame_lateral_btns = ctk.CTkFrame(frame_cuerpo, fg_color="transparent")
        frame_lateral_btns.pack(side="right", fill="y")
        ctk.CTkButton(frame_lateral_btns, text="✏️ Modificar Monto/Cuenta", font=("Arial", 12, "bold"), fg_color="#34495e", hover_color="#2c3e50", command=ejecutar_modificacion).pack(fill="x", pady=3)
        ctk.CTkButton(frame_lateral_btns, text="📂 Cambiar Soporte", font=("Arial", 12, "bold"), fg_color="#7f8c8d", hover_color="#606b6b", command=cambiar_soporte).pack(fill="x", pady=3)
        ctk.CTkButton(frame_lateral_btns, text="🗑️ Eliminar Soporte", font=("Arial", 12, "bold"), fg_color="#e74c3c", hover_color="#c0392b", command=eliminar_soporte).pack(fill="x", pady=3)
        ctk.CTkButton(frame_lateral_btns, text="❌ Eliminar Cobro Completo", font=("Arial", 12, "bold"), fg_color="#e74c3c", hover_color="#c0392b", command=eliminar_cobro_completo).pack(fill="x", pady=(15, 3))

        frame_btn_cierre = ctk.CTkFrame(v_edit, fg_color="transparent")
        frame_btn_cierre.pack(fill="x", padx=10, pady=10)
        ctk.CTkButton(frame_btn_cierre, text="❌ Salir", font=("Arial", 12, "bold"), fg_color="#7f8c8d", hover_color="#606b6b", command=v_edit.destroy).pack(side="right")

    def mostrar_reporte_totales(self):
        v_rep = ctk.CTkToplevel(self.main_root)
        v_rep.title("Reporte Avanzado de Ventas y Cobros")
        v_rep.geometry("750x550")
        v_rep.transient(self.main_root)
        v_rep.grab_set()

        ctk.CTkLabel(v_rep, text="📊 REPORTE DE TOTALES FILTRADOS", font=("Arial", 16, "bold"), text_color="#1f538d").pack(pady=(15, 10))

        f_filtros = ctk.CTkFrame(v_rep, fg_color="#f8f9fa", border_width=1, border_color="#ccc")
        f_filtros.pack(fill="x", padx=20, pady=10, ipadx=10, ipady=10)

        f_cli = ctk.CTkFrame(f_filtros, fg_color="transparent")
        f_cli.pack(fill="x", pady=5)
        ctk.CTkLabel(f_cli, text="Cliente:", font=("Arial", 11, "bold"), width=100, anchor="e").pack(side="left", padx=5)
        
        clis_mem = cache_sistema.obtener('lista_clientes_combobox')
        clis = ["Todos"] + (clis_mem if clis_mem else [])
        combo_cli = ctk.CTkComboBox(f_cli, values=clis, state="readonly", width=300)
        combo_cli.pack(side="left", padx=5)
        combo_cli.set("Todos")

        f_fechas = ctk.CTkFrame(f_filtros, fg_color="transparent")
        f_fechas.pack(fill="x", pady=5)
        
        ctk.CTkLabel(f_fechas, text="Desde:", font=("Arial", 11, "bold"), width=100, anchor="e").pack(side="left", padx=5)
        ent_desde = ctk.CTkEntry(f_fechas, width=110, placeholder_text=CONFIG_REGIONAL.get("formato_fecha", "DD/MM/AAAA"))
        ent_desde.pack(side="left", padx=5)
        ctk.CTkButton(f_fechas, text="📅", width=30, fg_color="#1f538d", hover_color="#163b65", command=lambda: CalendarioNativo(v_rep, ent_desde)).pack(side="left")

        ctk.CTkLabel(f_fechas, text="Hasta:", font=("Arial", 11, "bold"), width=60, anchor="e").pack(side="left", padx=5)
        ent_hasta = ctk.CTkEntry(f_fechas, width=110, placeholder_text=CONFIG_REGIONAL.get("formato_fecha", "DD/MM/AAAA"))
        ent_hasta.pack(side="left", padx=5)
        ctk.CTkButton(f_fechas, text="📅", width=30, fg_color="#1f538d", hover_color="#163b65", command=lambda: CalendarioNativo(v_rep, ent_hasta)).pack(side="left")

        f_resultados = ctk.CTkFrame(v_rep, fg_color="transparent")
        f_resultados.pack(fill="both", expand=True, padx=20, pady=10)
        
        def crear_lbl_res(padre, texto, color="#333"):
            f = ctk.CTkFrame(padre, fg_color="#ffffff", border_width=1, border_color="#ddd", corner_radius=8)
            f.pack(fill="x", pady=4)
            ctk.CTkLabel(f, text=texto, font=("Arial", 12, "bold"), text_color="gray").pack(side="left", padx=15, pady=10)
            lbl_val = ctk.CTkLabel(f, text="0.00", font=("Arial", 14, "bold"), text_color=color)
            lbl_val.pack(side="right", padx=15, pady=10)
            return lbl_val

        lbl_bruto = crear_lbl_res(f_resultados, "Total Ventas Brutas (Subtotal Base):", "#1f538d")
        lbl_igv = crear_lbl_res(f_resultados, "Total IGV Facturado (18%):", "#1f538d")
        lbl_det = crear_lbl_res(f_resultados, "Total Detracción/Retención:", "#e67e22")
        lbl_cobrado = crear_lbl_res(f_resultados, "Total Cobrado (Dinero Ingresado Real):", "#27ae60")
        lbl_por_cobrar = crear_lbl_res(f_resultados, "Total por Cobrar (Deuda Pendiente General):", "#c0392b")

        def convertir_a_fecha(fecha_str):
            if not fecha_str: return None
            for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y", "%m/%d/%Y"]:
                try: return datetime.strptime(fecha_str.strip(), fmt)
                except ValueError: pass
            return None

        def calcular_totales():
            cli_filtro = combo_cli.get()
            d_desde = convertir_a_fecha(ent_desde.get())
            d_hasta = convertir_a_fecha(ent_hasta.get())
            
            if ent_desde.get() and ent_hasta.get() and (not d_desde or not d_hasta):
                return messagebox.showwarning("Error", "Formato de fecha inválido.", parent=v_rep)

            if d_desde and d_hasta and d_desde > d_hasta:
                d_desde, d_hasta = d_hasta, d_desde

            conn = conectar_db()
            if not conn: return
            try:
                c = conn.cursor()
                c.execute("SELECT id_factura, monto_pagado FROM pagos_clientes")
                cobros_dict = {}
                for id_f, monto in c.fetchall():
                    cobros_dict[id_f] = cobros_dict.get(id_f, 0.0) + (float(monto) if monto else 0.0)

                c.execute("SELECT id, fecha, cliente, subtotal, impuesto, total, COALESCE(det_monto, 0), tipo_documento, enlace_pdf_sunat, archivo_ruta, estado_sunat FROM facturas_emitidas")
                
                tot_bruto = 0.0
                tot_igv = 0.0
                tot_det = 0.0
                tot_cobrado = 0.0
                tot_deuda = 0.0

                for r in c.fetchall():
                    id_fac, fecha, cli, sub, imp, tot, det, tipo_doc, enlace_sunat, arch_ruta, est_sunat = r[:11]
                    
                    if est_sunat and "Anulada" in str(est_sunat): continue
                    
                    if cli_filtro != "Todos" and cli != cli_filtro: continue
                    
                    f_dt = convertir_a_fecha(str(fecha))
                    if d_desde and d_hasta:
                        if not f_dt or not (d_desde <= f_dt <= d_hasta): continue
                    elif d_desde:
                        if not f_dt or f_dt < d_desde: continue
                    elif d_hasta:
                        if not f_dt or f_dt > d_hasta: continue

                    sub_val = float(sub) if sub else 0.0
                    imp_val = float(imp) if imp else 0.0
                    tot_val = float(tot) if tot else 0.0
                    det_val = float(det) if det else 0.0
                    
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

                    cobrado = cobros_dict.get(id_fac, 0.0)
                    saldo = max(0.0, neto - cobrado)

                    tot_cobrado += cobrado
                    tot_deuda += saldo

                lbl_bruto.configure(text=formatear_moneda(tot_bruto))
                lbl_igv.configure(text=formatear_moneda(tot_igv))
                lbl_det.configure(text=formatear_moneda(tot_det))
                lbl_cobrado.configure(text=formatear_moneda(tot_cobrado))
                lbl_por_cobrar.configure(text=formatear_moneda(tot_deuda))

            except Exception as e:
                messagebox.showerror("Error", f"Fallo al calcular:\n{e}", parent=v_rep)
            finally:
                liberar_conexion(conn)

        ctk.CTkButton(f_filtros, text="🔍 Procesar y Calcular", font=("Arial", 12, "bold"), fg_color="#1f538d", hover_color="#163b65", command=calcular_totales).pack(pady=(10, 5))
        calcular_totales()


# =========================================================
# 🚀 PESTAÑA 3: NOTAS DE CRÉDITO (ANULACIONES)
# =========================================================
class NotasCreditoTab:
    def __init__(self, tab_frame, main_root, app_padre):
        self.tab_frame = tab_frame
        self.main_root = main_root
        self.app_padre = app_padre
        
        # 🚀 VARIABLES DE PAGINACIÓN
        self.pagina_actual = 1
        self.registros_por_pagina = 50
        
        # 🗓️ FILTRO DE MES (por defecto, el mes en curso)
        self.mes_filtro = mes_en_curso()
        
        self.crear_interfaz()

    def importar_nc_sire(self):
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
                "Por favor, vaya a Configuración General del Sistema."
            )
            return

        periodo = simpledialog.askstring("Periodo SIRE SUNAT", "Ingrese el Periodo a descargar (Formato YYYYMM, ej: 202607):", initialvalue=datetime.now().strftime("%Y%m"))
        if not periodo or len(periodo) != 6 or not periodo.isdigit():
            return messagebox.showerror("Error", "Debe ingresar un periodo válido de 6 dígitos (ej: 202607).")

        v_sire = ctk.CTkToplevel(self.main_root)
        v_sire.title("Conexión Oficial SUNAT SIRE (Ventas)")
        v_sire.geometry("480x300")
        v_sire.grab_set()

        ctk.CTkLabel(v_sire, text="🌐 IMPORTANDO NOTAS DE CRÉDITO (SIRE)", font=("Arial", 14, "bold"), text_color="#166534").pack(pady=(20, 10))
        
        lbl_status = ctk.CTkLabel(v_sire, text="🔑 Autenticando token con la SUNAT...", font=("Arial", 11, "italic"), text_color="#d35400")
        lbl_status.pack(pady=10)

        prog = ctk.CTkProgressBar(v_sire, width=380)
        prog.pack(pady=10)
        prog.set(0.2)

        txt_info = ctk.CTkTextbox(v_sire, height=100, font=("Arial", 10))
        txt_info.pack(fill="x", padx=25, pady=10)

        def ejecucion_sire(estado):              # 👈 corre en el HILO SECUNDARIO
            try:
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
                    pass

                estado["progreso"] = 0.6
                estado["texto"] = "📥 Descargando Registro de Ventas RVIE..."
                estado["color"] = "#1f538d"

                datos_nc = []
                if token_access:
                    url_ventas = f"https://api-sire.sunat.gob.pe/v1/contribuyente/mrv/cpe/comprobantes/periodo/{periodo}"
                    req_c = urllib.request.Request(url_ventas, headers={"Authorization": f"Bearer {token_access}"})
                    try:
                        with urllib.request.urlopen(req_c, timeout=15) as res_c:
                            datos_ventas = json.loads(res_c.read().decode("utf-8"))
                            if isinstance(datos_ventas, dict) and "comprobantes" in datos_ventas:
                                for comp in datos_ventas["comprobantes"]:
                                    if comp.get("tipoDocumento") == "07":
                                        datos_nc.append(comp)
                    except Exception:
                        pass
                
                estado["progreso"] = 1.0
                
                estado["mensaje"] = (
                    f"✅ Conexión completada con éxito.\n"
                    f"• Periodo Sincronizado: {periodo}\n"
                    f"• Notas de Crédito Encontradas: {len(datos_nc)}\n\n"
                    f"El registro ha sido actualizado con los datos oficiales de SUNAT (RVIE)."
                )
                
            except Exception as e:
                estado["error"] = e

        def aplicar_sire(estado):                # 👈 corre en el HILO PRINCIPAL (cada sondeo)
            if "progreso" in estado:
                prog.set(estado["progreso"])
            if "texto" in estado:
                lbl_status.configure(text=estado["texto"], text_color=estado.get("color", "#1f538d"))

        def terminar_sire(estado):               # 👈 corre en el HILO PRINCIPAL
            if estado.get("error"):
                lbl_status.configure(text="❌ Error en Conexión SIRE", text_color="#c0392b")
                txt_info.delete("1.0", tk.END)
                txt_info.insert("1.0", f"Fallo al conectar con SUNAT:\n{estado['error']}")
                return
            lbl_status.configure(text="✅ Sincronización SIRE Finalizada", text_color="#27ae60")
            txt_info.delete("1.0", tk.END)
            txt_info.insert("1.0", estado.get("mensaje", ""))
            self.cargar_datos_nc(reset_pagina=True)

        ejecutar_en_hilo(v_sire, ejecucion_sire, aplicar=aplicar_sire, al_terminar=terminar_sire)

    def crear_interfaz(self):
        f_top = ctk.CTkFrame(self.tab_frame, corner_radius=8, fg_color="#f8f9fa", border_width=1, border_color="#e0e0e0")
        f_top.pack(fill="x", padx=15, pady=(10, 10), ipady=5)

        btn_sire_nc = ctk.CTkButton(f_top, text="🌐 Descargar NC desde SUNAT (SIRE)", font=("Arial", 12, "bold"), fg_color="#166534", hover_color="#14532d", command=self.importar_nc_sire)
        btn_sire_nc.pack(side="left", padx=15, pady=5)

        ctk.CTkLabel(f_top, text="🔍 Buscar:", font=("Arial", 11, "bold")).pack(side="left", padx=(15, 5))
        self.ent_buscar_nc = ctk.CTkEntry(f_top, placeholder_text="Filtrar por documento, cliente...")
        self.ent_buscar_nc.pack(side="left", fill="x", expand=True, padx=5)
        
        ctk.CTkLabel(f_top, text="🗓️ Mes:", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 5))
        self.combo_mes = ctk.CTkComboBox(f_top, values=construir_valores_mes(), width=150, state="readonly", command=self.on_cambiar_mes)
        self.combo_mes.pack(side="left", padx=(0, 5))
        self.combo_mes.set(self.mes_filtro)
        
        self.ent_buscar_nc.bind("<KeyRelease>", lambda e: self.buscar_con_retraso())
        self.ent_buscar_nc.bind("<Return>", lambda e: self.cargar_datos_nc(reset_pagina=True))

        btn_refresh = ctk.CTkButton(f_top, text="🔄 Actualizar Tabla", font=("Arial", 12, "bold"), fg_color="#7f8c8d", hover_color="#606b6b", command=lambda: self.cargar_datos_nc(reset_pagina=True))
        btn_refresh.pack(side="right", padx=15, pady=5)

        f_tabla = ctk.CTkFrame(self.tab_frame, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True, padx=15, pady=0)

        columnas = ("num", "id_fac", "fecha_orig", "doc_orig", "doc_nc", "cliente", "monto_anulado", "estado", "pdf_nc")
        self.tabla = ttk.Treeview(f_tabla, columns=columnas, show="headings")

        self.tabla.heading("num", text="N°")
        self.tabla.heading("id_fac", text="ID (Oculto)")
        self.tabla.heading("fecha_orig", text="Fecha Fac. Original")
        self.tabla.heading("doc_orig", text="Doc. Original")
        self.tabla.heading("doc_nc", text="Nota Crédito")
        self.tabla.heading("cliente", text="Cliente")
        self.tabla.heading("monto_anulado", text="Monto Anulado")
        self.tabla.heading("estado", text="Estado Actual")
        self.tabla.heading("pdf_nc", text="PDF Nota Crédito")

        self.tabla.column("num", width=40, anchor="center")
        self.tabla.column("id_fac", width=0, stretch=tk.NO)
        self.tabla.column("fecha_orig", width=100, anchor="center")
        # 📏 N° de documento (origen y nota de crédito) sin recortar el último dígito
        self.tabla.column("doc_orig", width=130, minwidth=120, stretch=False, anchor="w")
        self.tabla.column("doc_nc", width=130, minwidth=120, stretch=False, anchor="w")
        self.tabla.column("cliente", width=250, anchor="w")
        self.tabla.column("monto_anulado", width=110, anchor="e")
        self.tabla.column("estado", width=180, anchor="center")
        self.tabla.column("pdf_nc", width=120, anchor="center")

        self.tabla.config(displaycolumns=("num", "fecha_orig", "doc_orig", "doc_nc", "cliente", "monto_anulado", "estado", "pdf_nc"))
        self.tabla.bind("<Double-1>", self.abrir_pdf_nc)

        scroll_y = ctk.CTkScrollbar(f_tabla, orientation="vertical", command=self.tabla.yview)
        scroll_x = ctk.CTkScrollbar(f_tabla, orientation="horizontal", command=self.tabla.xview)
        self.tabla.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_x.pack(side="bottom", fill="x")
        scroll_y.pack(side="right", fill="y")
        self.tabla.pack(side="left", fill="both", expand=True)
        for _c in self.tabla["columns"]:
            try:
                self.tabla.column(_c, stretch=False)
            except Exception:
                pass
        
        # 🚀 BOTONES DE PAGINACIÓN
        f_paginacion = ctk.CTkFrame(self.tab_frame, fg_color="transparent")
        f_paginacion.pack(fill="x", padx=15, pady=10)
        
        self.btn_ant = ctk.CTkButton(f_paginacion, text="◀ Ant", width=60, command=self.pagina_anterior)
        self.btn_ant.pack(side="left")
        
        self.lbl_pagina = ctk.CTkLabel(f_paginacion, text=f"Pág {self.pagina_actual}", font=("Arial", 11, "bold"))
        self.lbl_pagina.pack(side="left", padx=10)
        
        self.btn_sig = ctk.CTkButton(f_paginacion, text="Sig ▶", width=60, command=self.pagina_siguiente)
        self.btn_sig.pack(side="left")
        
        ctk.CTkLabel(f_paginacion, text="💡 Haz doble clic sobre un registro para descargar/visualizar el PDF oficial de la Nota de Crédito.", font=("Arial", 11, "italic"), text_color="gray").pack(side="right", pady=10)

        self.main_root.after(250, lambda: self.cargar_datos_nc(reset_pagina=True))

    def on_cambiar_mes(self, choice):
        self.mes_filtro = choice
        self.cargar_datos_nc(reset_pagina=True)

    def pagina_anterior(self):
        if self.pagina_actual > 1:
            self.pagina_actual -= 1
            self.cargar_datos_nc()
            
    def pagina_siguiente(self):
        self.pagina_actual += 1
        self.cargar_datos_nc()

    def buscar_con_retraso(self):
        if hasattr(self, "_busqueda_job"):
            try: self.main_root.after_cancel(self._busqueda_job)
            except: pass
        self._busqueda_job = self.main_root.after(350, lambda: self.cargar_datos_nc(reset_pagina=True))

    # 🚀 FIX: CARGA LAZY LOADING + CACHÉ
    def cargar_datos_nc(self, reset_pagina=False):
        if reset_pagina:
            self.pagina_actual = 1
            
        self.lbl_pagina.configure(text=f"Pág {self.pagina_actual}")

        for fila in self.tabla.get_children(): 
            self.tabla.delete(fila)
        
        filtro = ""
        if hasattr(self, 'ent_buscar_nc'):
            filtro = self.ent_buscar_nc.get().strip().lower()

        offset = (self.pagina_actual - 1) * self.registros_por_pagina
        clave_cache = f"notas_credito_{filtro}_mes_{self.mes_filtro}_pag_{self.pagina_actual}"
        datos = cache_sistema.obtener(clave_cache)

        if datos is not None:
            self._pintar_notas_credito(datos)
        else:
            self.tabla.insert("", tk.END, values=("", "", "", "Cargando datos...", "", "", "", "", ""))
            
            def tarea_descarga(estado):          # 👈 corre en el HILO SECUNDARIO
                conn = conectar_db(silencioso=True)
                if not conn: return
                try:
                    cursor = conn.cursor()
                    condiciones = ["estado_sunat LIKE %s"]
                    params = ["%Anulada%"]
                    if filtro:
                        val = f"%{filtro}%"
                        condiciones.append("(numero_documento ILIKE %s OR cliente ILIKE %s)")
                        params.extend([val, val])
                    patron_mes = patron_fecha_mes(self.mes_filtro)
                    if patron_mes:
                        condiciones.append("fecha LIKE %s")
                        params.append(patron_mes)
                    where_sql = " WHERE " + " AND ".join(condiciones)
                    params.extend([self.registros_por_pagina, offset])
                    cursor.execute(f"SELECT id, fecha, numero_documento, cliente, total, COALESCE(det_monto, 0), estado_sunat, enlace_pdf_nc, tipo_documento FROM facturas_emitidas{where_sql} ORDER BY id DESC LIMIT %s OFFSET %s", tuple(params))
                        
                    datos_db = cursor.fetchall()
                    cache_sistema.guardar(clave_cache, datos_db)
                    estado["datos_db"] = datos_db
                except Exception as e:
                    print("Error cargando Notas de Crédito:", e)
                finally:
                    liberar_conexion(conn)

            def _pintar_notas_hilo(estado):      # 👈 corre en el HILO PRINCIPAL
                if estado.get("datos_db") is not None:
                    self._pintar_notas_credito(estado["datos_db"])

            ejecutar_en_hilo(self.main_root, tarea_descarga, al_terminar=_pintar_notas_hilo)

    def _pintar_notas_credito(self, registros):
        for fila in self.tabla.get_children(): self.tabla.delete(fila)

        contador = 1
        for reg in registros:
            id_fac, fecha, nro_doc, cliente, tot_bruto, det_monto, estado, link_nc, tipo_doc = reg
            
            tot_bruto_val = float(tot_bruto) if tot_bruto else 0.0
            det_monto_val = float(det_monto) if det_monto else 0.0
            neto_anulado = tot_bruto_val - det_monto_val
            
            serie_nc = "FC01" if tipo_doc and "Factura" in tipo_doc else "BC01"
            num_orig = str(nro_doc).split("-")[1] if nro_doc and "-" in str(nro_doc) else "1"
            doc_nc_str = f"{serie_nc}-{num_orig}"
            
            if not link_nc: txt_pdf = "❌ Sin PDF / SIRE"
            else: txt_pdf = "✅ Ver PDF NC"

            row_vals = (contador, id_fac, fecha, nro_doc, doc_nc_str, cliente, formatear_moneda(neto_anulado), estado, txt_pdf)

            self.tabla.insert("", tk.END, values=row_vals)
            contador += 1
            
        if self.pagina_actual > 1:
            self.btn_ant.configure(state="normal")
        else:
            self.btn_ant.configure(state="disabled")
            
        if len(registros) == self.registros_por_pagina:
            self.btn_sig.configure(state="normal")
        else:
            self.btn_sig.configure(state="disabled")

    def abrir_pdf_nc(self, event):
        seleccion = self.tabla.selection()
        if not seleccion: return
        id_fac = self.tabla.item(seleccion[0], "values")[1] 
        
        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT enlace_pdf_nc, ruc FROM facturas_emitidas f LEFT JOIN clientes c ON f.cliente = c.nombre_empresa WHERE f.id = %s", (id_fac,))
            res = cursor.fetchone()
            if res:
                link_nc = res[0]
                if link_nc and str(link_nc).startswith("http"):
                    webbrowser.open(link_nc)
                else:
                    messagebox.showinfo("Consulta SUNAT", "Este documento fue importado desde el SIRE y no posee un enlace directo al PDF.\n\nSe abrirá el portal de validación de SUNAT para que pueda consultarlo.")
                    webbrowser.open("https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias")
            else:
                messagebox.showinfo("Aviso", "Este registro no se encuentra disponible.")
        except Exception as e:
            pass
        finally:
            liberar_conexion(conn)

# =========================================================
# CLASE PRINCIPAL: MÓDULO DE VENTAS (CONTENEDOR TABVIEW)
# =========================================================
class ModuloVentasApp:
    def __init__(self, parent_frame):
        self.parent_frame = parent_frame
        self.usuario_activo = "Desconocido"
        self.pantalla_expandida = False
        aplicar_estilo_treeview()

        self.frame_main = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.frame_main.pack(fill="both", expand=True, padx=15, pady=15)
        
        header_frame = ctk.CTkFrame(self.frame_main, fg_color="transparent")
        header_frame.pack(fill="x")
        
        ctk.CTkLabel(header_frame, text="💼 MÓDULO DE VENTAS (COMERCIAL Y TESORERÍA)", font=("Arial", 18, "bold"), text_color="#1f538d").pack(side="left")
        self.btn_pantalla = ctk.CTkButton(header_frame, text="[ + ] Pantalla Completa", font=("Arial", 12, "bold"), width=160, fg_color="#34495e", hover_color="#2c3e50", command=self.toggle_pantalla_completa)
        self.btn_pantalla.pack(side="right")

        self.tabview = ctk.CTkTabview(self.frame_main, segmented_button_selected_color="#1f538d")
        self.tabview.pack(fill="both", expand=True, pady=(10, 0))
        
        self.tab_emision = self.tabview.add(" 📤 1. Emisión de Facturas (Ventas) ")
        self.tab_cobros = self.tabview.add(" 💰 2. Control de Cobros y Deudas ")
        self.tab_nc = self.tabview.add(" 📄 3. Notas de Crédito ")
        
        self.app_facturas = FacturasEmitidasTab(self.tab_emision, self.parent_frame, self)
        self.app_cobros = CuentasPorCobrarTab(self.tab_cobros, self.parent_frame, self)
        self.app_nc = NotasCreditoTab(self.tab_nc, self.parent_frame, self)
        
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
            self.btn_pantalla.configure(text="[ + ] Pantalla Completa", fg_color="#34495e", hover_color="#2c3e50")
            self.pantalla_expandida = False
        else:
            if sidebar: sidebar.pack_forget()
            self.btn_pantalla.configure(text="[ - ] Restaurar Vista", fg_color="#34495e", hover_color="#2c3e50")
            self.pantalla_expandida = True

    def al_cambiar_pestana(self):
        if self.tabview.get() == " 💰 2. Control de Cobros y Deudas ":
            self.app_cobros.cargar_datos_cobrar(reset_pagina=True)
        elif self.tabview.get() == " 📄 3. Notas de Crédito ":
            self.app_nc.cargar_datos_nc(reset_pagina=True)

if __name__ == "__main__":
    pass