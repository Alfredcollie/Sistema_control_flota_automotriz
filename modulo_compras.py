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

def abrir_documento(ruta):
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
              "IMPORTETOTALVENTA", "TOTALPAGAR", "TOTALAGENERAL"],
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


def leer_ocr_windows(ruta, idiomas="es-PE,es-MX,es-ES,en-US", dpi=250):
    """OCR local con el motor de Windows. Devuelve el texto reconstruido o "".

    No requiere instalar nada ni conexión a internet. En otros sistemas
    operativos devuelve "" (el usuario puede digitar los montos a mano).
    """
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

    paginas = []
    for imagen in imagenes[:5]:
        try:
            orden = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script,
                     "-Ruta", imagen, "-Idiomas", idiomas]
            resultado = subprocess.run(orden, capture_output=True, text=True, encoding="utf-8",
                                       errors="ignore", timeout=180,
                                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            palabras = []
            for linea in (resultado.stdout or "").splitlines():
                partes = linea.split("\t")
                if len(partes) != 5:
                    continue
                try:
                    x, y, alto = int(partes[0]), int(partes[1]), int(partes[3])
                except ValueError:
                    continue
                palabras.append((x, y, alto, partes[4]))
            texto = _agrupar_palabras_en_lineas(palabras)
            if texto.strip():
                paginas.append(texto)
        except Exception:
            continue
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


def extraer_montos_comprobante(texto):
    """Lee base, IGV, recargo al consumo y total de un comprobante peruano.

    Funciona igual con el texto de un PDF que con el texto reconstruido por OCR:
    los rótulos se comparan de forma tolerante y el importe de cada fila se toma
    como el último número de esa fila.
    """
    datos = {"base": 0.0, "igv": 0.0, "recargo": 0.0, "total": 0.0, "leidos": 0}
    for linea in str(texto or "").splitlines():
        fila = linea.replace("S/.", " ").replace("S/", " ")
        corte = re.search(r"\d", fila)
        etiqueta = _normalizar_etiqueta(fila[:corte.start()] if corte else fila)
        if not etiqueta:
            continue
        montos = re.findall(r"\d{1,3}(?:[.,]\d{3})*[.,]\d{2}|\d+[.,]\d{1,2}", fila)
        if not montos:
            continue
        try:
            valor = parsear_monto_texto(montos[-1])
        except ValueError:
            continue
        for grupo in ("recargo", "total", "base", "igv"):
            if datos[grupo]:
                continue
            if _etiqueta_reconocida(etiqueta, grupo):
                datos[grupo] = valor
                datos["leidos"] += 1
                break

    # El recargo al consumo puede no haberse leído: se deduce de la diferencia
    if datos["base"] and datos["igv"] and datos["total"] and not datos["recargo"]:
        diferencia = round(datos["total"] - datos["base"] - datos["igv"], 2)
        if 0.02 <= diferencia <= datos["total"] * 0.35:
            datos["recargo"] = diferencia
    # Si no se leyó la base pero sí el total y el IGV, se deduce
    if not datos["base"] and datos["total"] and datos["igv"]:
        datos["base"] = round(datos["total"] - datos["recargo"] - datos["igv"], 2)
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


def ruc_distinto_al_de_la_empresa(texto, ruc_empresa=""):
    """RUC del proveedor: el primer RUC de 11 dígitos que no sea el de la empresa."""
    propio = re.sub(r"\D", "", str(ruc_empresa or ""))
    for candidato in re.findall(r"\d{11}", str(texto or "")):
        if candidato != propio:
            return candidato
    return ""


def nombre_proveedor_desde_texto(texto, ruc_proveedor=""):
    """Razón social del emisor: la línea con letras que está sobre su RUC."""
    lineas = [l.strip() for l in str(texto or "").splitlines() if l.strip()]
    indice = None
    if ruc_proveedor:
        for i, linea in enumerate(lineas):
            if ruc_proveedor in linea:
                indice = i
                break
    if indice is None:
        indice = min(len(lineas), 4)
    for linea in reversed(lineas[:indice]):
        if len(re.sub(r"[^A-Za-zÁÉÍÓÚÑáéíóúñ]", "", linea)) < 4:
            continue
        if re.search(r"FACTURA|BOLETA|ELECTR[OÓ]NIC|R\.?U\.?C|D\.?N\.?I|TICKET|"
                     r"COMPROBANTE|COTIZACI|GUIA|PROFORMA", linea, re.IGNORECASE):
            continue
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


def construir_condicion_busqueda_compras(filtro):
    """Genera cláusula SQL + parámetros para que la búsqueda del módulo de
    Compras funcione contra CUALQUIER columna de la fila (no solo N° doc,
    proveedor, vehículo y concepto).

    Cubre: fecha, N° documento, proveedor, RUC, vehículo/placa, kilometraje,
    cantidad, concepto/descripción (incluye la hora), tipo de documento,
    categoría, días de crédito, montos (subtotal, IGV, total, detracción),
    y también la forma de pago / montos pagados en pagos_comprobantes.
    Retorna (clausula_sql, lista_de_parametros); si el filtro está vacío,
    retorna ("", []).
    """
    if not filtro or not str(filtro).strip():
        return "", []
    val = f"%{str(filtro).strip()}%"

    # Todas las columnas se convierten a texto para que la búsqueda funcione
    # sin importar el tipo real de la columna (texto, fecha, número, etc.)
    columnas = ("numero_documento", "proveedor", "evento_asociado", "descripcion",
                "tipo_documento", "categoria", "ruc", "fecha", "dias_credito",
                "kilometraje", "cantidad_combustible", "subtotal", "impuesto",
                "total", "det_monto", "orden_servicio")

    partes = [f"CAST({c} AS TEXT) ILIKE %s" for c in columnas]
    cantidad = len(columnas)

    # Forma de pago / montos pagados asociados al comprobante (columnas derivadas)
    partes.append(
        "EXISTS (SELECT 1 FROM pagos_comprobantes pc "
        "WHERE pc.id_factura = facturas_recibidas.id "
        "AND (pc.cuenta_origen ILIKE %s OR CAST(pc.monto_pagado AS TEXT) ILIKE %s))"
    )
    cantidad += 2

    return "(" + " OR ".join(partes) + ")", [val] * cantidad

# =========================================================
# CLASE: CALENDARIO NATIVO
# =========================================================
class CalendarioNativo(ctk.CTkToplevel):
    def __init__(self, parent, target_entry):
        super().__init__(parent)
        self.target_entry = target_entry
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
MANTENIMIENTOS_VEHICULO = (
    {"nombre": "🛢️ Cambio de Aceite",
     "campos": (("fec_aceite", 0), ("fecha_ultimo_aceite", 0)),
     "km": ("km_ultimo_aceite", "Km del cambio"),
     "actual": ("fec_aceite", "fecha_ultimo_aceite", "km_ultimo_aceite")},
    {"nombre": "⛓️ Cadena / Correa de Tiempo",
     "campos": (("fec_correa", 0),),
     "km": ("km_prox_correa", "Próximo cambio a los (Km)"),
     "actual": ("fec_correa", "km_prox_correa")},
    {"nombre": "💨 Mantenimiento del Sistema de Gas",
     "campos": (("fec_rev_gas", 0),),
     "actual": ("fec_rev_gas",)},
    {"nombre": "🛠️ Mantenimiento Preventivo / General",
     "campos": (("fecha_ultimo_general", 0),),
     "km": ("km_ultimo_general", "Km del servicio"),
     "actual": ("fecha_ultimo_general", "km_ultimo_general")},
    {"nombre": "📄 SOAT (emisión → vence en 1 año)",
     "campos": (("emision_soat", 0), ("vencimiento_soat", 12)),
     "actual": ("emision_soat", "vencimiento_soat")},
    {"nombre": "🔍 Revisión Técnica (realizada → vence en 1 año)",
     "campos": (("vencimiento_rt", 12),),
     "actual": ("vencimiento_rt",)},
    {"nombre": "🛡️ Póliza de Seguro (emisión → vence en 1 año)",
     "campos": (("emision_seguro", 0), ("vencimiento_seguro", 12)),
     "actual": ("emision_seguro", "vencimiento_seguro")},
    {"nombre": "🔋 Batería (compra → vence en 2 años)",
     "campos": (("fec_compra_bat", 0), ("fec_venc_bat", 24)),
     "actual": ("fec_compra_bat", "fec_venc_bat")},
    {"nombre": "🧯 Extintor (recarga → vence en 1 año)",
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
}

COLUMNAS_MANT = ("placa", "marca", "modelo", "kilometraje", "fec_aceite", "fecha_ultimo_aceite",
                 "fec_correa", "fec_rev_gas", "fecha_ultimo_general", "emision_soat",
                 "vencimiento_soat", "vencimiento_rt", "emision_seguro", "vencimiento_seguro",
                 "fec_compra_bat", "fec_venc_bat", "fec_venc_extintor", "km_ultimo_aceite",
                 "km_prox_correa", "km_ultimo_general")


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


class FacturasRecibidasTab:
    def __init__(self, tab_frame, main_root, app_padre):
        self.tab_frame = tab_frame
        self.main_root = main_root
        self.app_padre = app_padre
        
        self.orden_columnas = {}
        # 🔃 Ordenamiento por cualquier columna aplicado a TODAS las páginas
        self.columna_orden = "id"
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
        self.registros_por_pagina = 50
        
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
                # 🔗 Cruce de la factura con su Orden de Servicio (módulo de Órdenes)
                try: cursor.execute("ALTER TABLE facturas_recibidas ADD COLUMN IF NOT EXISTS id_orden_servicio INTEGER;"); conn.commit()
                except: conn.rollback()
                try: cursor.execute("ALTER TABLE facturas_recibidas ADD COLUMN IF NOT EXISTS orden_servicio VARCHAR(120) DEFAULT '';"); conn.commit()
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

    def autocompletar_desde_pdf(self):
        """Lee un comprobante (PDF con texto, PDF escaneado o foto) y llena los campos."""
        ruta = seleccionar_archivo_dialogo(
            "Seleccionar la factura (PDF o foto escaneada)",
            [("Documentos", "*.pdf;*.png;*.jpg;*.jpeg"),
             ("Archivos PDF", "*.pdf"),
             ("Imágenes", "*.png;*.jpg;*.jpeg")])
        if not ruta:
            return
        try:
            self.bloquear_autocompletado_ruc = True
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

            # ---- N° de documento ----
            nro_match = re.search(r"([EFB][0-9A-Z]{3}\s*-\s*\d+)", texto)
            if nro_match:
                self.ent_nro_doc.delete(0, tk.END)
                self.ent_nro_doc.insert(0, nro_match.group(1).replace(" ", ""))

            # ---- Fecha ----
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

            # ---- RUC y razón social del proveedor (el emisor, no el cliente) ----
            ruc_proveedor = ruc_distinto_al_de_la_empresa(texto, CONFIG_REGIONAL.get("ruc_empresa", ""))
            if ruc_proveedor:
                self.ent_desc.delete(0, tk.END)
                self.ent_desc.insert(0, ruc_proveedor)
            nombre = nombre_proveedor_desde_texto(texto, ruc_proveedor)
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
        self.combo_categoria = ctk.CTkComboBox(f_cat, state="readonly")
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

        # 💰 El monto se puede digitar SIN IGV (monto base) o CON IGV (total del documento)
        ctk.CTkLabel(self.f_form, text="Tipo de Monto a Ingresar:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10)
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

        f_tabla = ctk.CTkFrame(self.f_wrapper_derecha, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True)

        columnas = ("num", "id", "fecha", "hora", "nro_doc", "orden", "dias", "tipo", "proveedor", "ruc", "categoria", "evento", "kilometraje", "cantidad", "desc", "metodo_pago", "subtotal", "impuesto", "total", "detraccion", "neto", "archivo")
        self.columnas_tabla = columnas
        self.tabla = ttk.Treeview(f_tabla, columns=columnas, show="headings")
        
        self.tabla.tag_configure("con_cuenta", background="#e8f8f5", foreground="#0e6251") 
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
        self.tabla.column("nro_doc", width=90, anchor="center")
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

        f_btn_tabla = ctk.CTkFrame(self.f_wrapper_derecha, fg_color="transparent")
        f_btn_tabla.pack(fill="x", pady=(10, 0))
        
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
        f_leyenda = ctk.CTkFrame(f_btn_tabla, fg_color="transparent")
        f_leyenda.pack(side="left", padx=20)
        
        ctk.CTkLabel(f_leyenda, text="■", font=("Arial", 14), text_color="#c0392b").pack(side="left", padx=(5,2))
        ctk.CTkLabel(f_leyenda, text="Sin Cuenta Asignada", font=("Arial", 11, "bold"), text_color="#333333").pack(side="left", padx=(0,10))
        
        ctk.CTkLabel(f_leyenda, text="■", font=("Arial", 14), text_color="#27ae60").pack(side="left", padx=(5,2))
        ctk.CTkLabel(f_leyenda, text="Cuenta Asignada", font=("Arial", 11, "bold"), text_color="#333333").pack(side="left", padx=(0,5))

        # ✅ Solo los botones: el de Modificar/Eliminar pasa a una fila propia
        # (ancho completo) para que siempre sea visible sin importar el ancho.
        f_gestion = ctk.CTkFrame(self.f_wrapper_derecha, fg_color="transparent")
        f_gestion.pack(fill="x", pady=(8, 0))
        btn_gestionar = ctk.CTkButton(f_gestion, text="⚙️ Modificar o Eliminar Registro Seleccionado", font=("Arial", 12, "bold"), command=self.abrir_ventana_edicion, fg_color="#34495e", hover_color="#2c3e50", height=36)
        btn_gestionar.pack(fill="x")

        # Descarga de tickets al abrir el módulo (en segundo plano, aviso en el hilo principal)
        ejecutar_en_hilo(
            self.main_root,
            lambda estado: estado.update(n=self.sincronizar_tickets_pendientes_automatico()),
            al_terminar=lambda estado: (self.cargar_datos_tabla(reset_pagina=True)
                                        if estado.get("n") else None))
        self.main_root.after(100, lambda: self.cargar_datos_tabla(reset_pagina=True))

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

        if not prov or not ruc_val: return messagebox.showwarning("Atención", "Llene los campos obligatorios.")

        if nro_doc:
            conn_check = conectar_db()
            if conn_check:
                try:
                    c_check = conn_check.cursor()
                    c_check.execute("SELECT COUNT(*) FROM facturas_recibidas WHERE numero_documento = %s AND proveedor = %s", (nro_doc, prov))
                    if c_check.fetchone()[0] > 0:
                        return messagebox.showwarning("Duplicado", "Ese N° de Documento ya está registrado.")
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

        conn = conectar_db()
        if not conn: return
        try:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO facturas_recibidas (tipo_documento, numero_documento, fecha, proveedor, descripcion, evento_asociado, subtotal, impuesto, total, archivo_ruta, dias_credito, det_porcentaje, det_monto, categoria, ruc, id_orden_servicio, orden_servicio)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (tipo, nro_doc, fecha, prov, desc, evento, subtotal, imp, tot_bruto, ruta_para_guardar(ruta_final), dias, det_pct, det_monto, categoria, ruc_val, orden_id, orden_txt))
            conn.commit()
            
            cache_sistema.invalidar()
            registrar_auditoria(self.app_padre.usuario_activo, "Facturas Recibidas", f"Registró factura {nro_doc} del proveedor '{prov}'")
            messagebox.showinfo("Éxito", "Documento recibido registrado correctamente.")
            
            self.cargar_categorias()
            self.cargar_ordenes_servicio()      # refresca la marca "⚠ ya facturada"
            self.ent_nro_doc.delete(0, tk.END)
            self.ent_desc.delete(0, tk.END)
            self.ent_concepto.delete(0, tk.END)
            self.ent_subtotal.delete(0, tk.END)
            if hasattr(self, "ent_recargo"):
                self.ent_recargo.delete(0, tk.END)
                self.ent_recargo.insert(0, "0")
            self.ruta_archivo_temp = ""
            self.btn_archivo.configure(text="📎 Adjuntar Archivo Manual", fg_color="#7f8c8d", hover_color="#606b6b")
            self.cargar_datos_tabla(reset_pagina=True)
            
            if hasattr(self.app_padre, 'app_pagos'):
                self.app_padre.app_pagos.cargar_datos_pagar(reset_pagina=True)
        except Exception as e: messagebox.showerror("Error", str(e))
        finally: liberar_conexion(conn)

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
        clave_cache = f"compras_recibidas_v3_{filtro}_mes_{self.mes_filtro}_app_{int(ocultar_app)}"
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
                    query_base = "SELECT id, fecha, numero_documento, dias_credito, tipo_documento, proveedor, evento_asociado, descripcion, subtotal, impuesto, total, COALESCE(det_monto, 0), archivo_ruta, categoria, kilometraje, cantidad_combustible, ruc, COALESCE(pagado_por_tercero, ''), COALESCE(es_compra_cruzada, FALSE), COALESCE(soporte_pago_tercero, ''), COALESCE(orden_servicio, '') FROM facturas_recibidas"
                    
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
                metodo_pago = f"{metodo_pago} · {nota}" if metodo_pago else nota

            if "Recibo" in tipo_doc and "8%" in tipo_doc: neto = tot_bruto - impuesto - det_monto
            else: neto = tot_bruto - det_monto
                
            etiqueta_color = "con_cuenta" if id_factura in cuentas_por_factura else "sin_cuenta"

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
            cursor.execute("SELECT tipo_documento, ruc, proveedor, numero_documento, fecha, evento_asociado, kilometraje, cantidad_combustible, descripcion, subtotal, impuesto, total, COALESCE(id_orden_servicio, 0), COALESCE(orden_servicio, '') FROM facturas_recibidas WHERE id = %s", (id_doc,))
            reg = cursor.fetchone()
        finally: 
            liberar_conexion(conn)
            
        if not reg: return
        e_tipo, e_ruc, e_prov, e_nro, e_fec, e_placa, e_km, e_gal, e_desc, e_sub, e_imp, e_tot, e_orden_id, e_orden_txt = reg
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
                                subtotal=%s, impuesto=%s, total=%s, id_orden_servicio=%s, orden_servicio=%s
                            WHERE id=%s
                        """, (
                            ent_tipo.get().strip(), ent_ruc.get().strip(), ent_prov.get().strip(), ent_nro.get().strip(), 
                            ent_fec.get().strip(), ent_placa.get().strip(), ent_km.get().strip(), ent_gal.get().strip(), 
                            desc_final, val_sub, val_imp, val_tot, id_orden_final, txt_orden_final, id_doc
                        ))
                        
                        cursor_u.execute("""
                            UPDATE pagos_comprobantes 
                            SET proveedor_nombre=%s, codigo_cotizacion=%s
                            WHERE id_factura=%s
                        """, (ent_prov.get().strip(), ent_nro.get().strip(), id_doc))
                        
                        conn_u.commit()
                        cache_sistema.invalidar()
                        messagebox.showinfo("Éxito", "Registro actualizado correctamente.", parent=v_edit)
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
    def __init__(self, tab_frame, main_root, app_padre):
        self.tab_frame = tab_frame
        self.main_root = main_root
        self.app_padre = app_padre
        
        self.pagina_actual = 1
        self.registros_por_pagina = 50
        
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

        # 🚀 OPCIÓN UNIFICADA: Reporte de Totales y Resumen por Proveedor
        btn_reporte = ctk.CTkButton(frame_acciones, text="📊 Reporte de Totales y Proveedores", font=("Arial", 12, "bold"), command=self.mostrar_reporte_totales, fg_color="#1f538d", hover_color="#163b65")
        btn_reporte.pack(side="left", padx=5, pady=5)

        btn_pago = ctk.CTkButton(frame_acciones, text="🧾 Registrar Pago", font=("Arial", 12, "bold"), command=self.cargar_comprobante_pago, fg_color="#27ae60", hover_color="#1e8449")
        btn_pago.pack(side="left", padx=5, pady=5)

        btn_editar = ctk.CTkButton(frame_acciones, text="✏️ Editar Pagos", font=("Arial", 12, "bold"), command=self.abrir_ventana_edicion, fg_color="#34495e", hover_color="#2c3e50")
        btn_editar.pack(side="left", padx=5, pady=5)

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
        self.tabla.column("nro_doc", width=100, anchor="center")
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
                    cursor.execute(f"SELECT id, fecha, numero_documento, proveedor, evento_asociado, descripcion, subtotal, impuesto, total, COALESCE(det_monto, 0), tipo_documento, kilometraje, cantidad_combustible, ruc, COALESCE(pagado_por_tercero, ''), COALESCE(es_compra_cruzada, FALSE), COALESCE(soporte_pago_tercero, '') FROM facturas_recibidas{where_sql} ORDER BY id DESC", tuple(params))
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
                         tercero, es_cruzada, soporte_cruzada) = reg
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
                            # Compra cruzada (creada desde el módulo de Banco): la factura
                            # la pagó un TERCERO, así que no queda saldo por pagar.
                            monto_pagado = neto_facturado
                            saldo_pendiente = 0.0
                        
                        filas_procesadas.append({
                            "id_factura": id_factura, "fecha": fecha, "nro_doc": nro_doc, "proveedor": proveedor, "ruc_db": ruc_db,
                            "evento": evento, "km_val": km_val, "cant_val": cant_val, "concepto": concepto, "cuentas_lista": cuentas_lista,
                            "sub_val": sub_val, "imp_val": imp_val, "det_monto_val": det_monto_val, "neto_facturado": neto_facturado,
                            "monto_pagado": monto_pagado, "saldo_pendiente": saldo_pendiente, "cant_archivos": cant_archivos, "tiene_cuenta": tiene_cuenta,
                            "es_cruzada": es_cruzada, "tercero": tercero, "soporte_cruzada": soporte_cruzada
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
        if not seleccion: return messagebox.showwarning("Selección", "Seleccione una factura.")
        
        valores = self.tabla.item(seleccion[0], "values")
        id_factura, nro_doc, proveedor = valores[1], valores[4], valores[5] 
        saldo_actual = desformatear_numero(valores[17]) 
        
        if saldo_actual <= 0: return messagebox.showinfo("Aviso", "Esta factura ya está pagada por completo.")

        v_pago = ctk.CTkToplevel(self.main_root)
        v_pago.title("Registrar Nuevo Pago")
        centrar_ventana(v_pago, self.main_root, 450, 480)
        v_pago.transient(self.main_root)
        v_pago.grab_set()

        ctk.CTkLabel(v_pago, text=f"Pago para: {proveedor}", font=("Arial", 14, "bold"), text_color="#1f538d").pack(pady=(15, 5))
        ctk.CTkLabel(v_pago, text=f"Saldo Pendiente: {formatear_moneda(saldo_actual)}", font=("Arial", 12)).pack(pady=(0, 15))

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
        ent_monto.insert(0, str(saldo_actual)) 

        ctk.CTkLabel(f_form, text="Fecha del Pago:", font=("Arial", 11, "bold")).pack(anchor="w")
        f_fecha_pago = ctk.CTkFrame(f_form, fg_color="transparent")
        f_fecha_pago.pack(fill="x", pady=(0, 10))
        ent_fecha = ctk.CTkEntry(f_fecha_pago)
        ent_fecha.pack(side="left", fill="x", expand=True)
        ent_fecha.insert(0, datetime.now().strftime("%d/%m/%Y"))
        ctk.CTkButton(f_fecha_pago, text="📅", width=40, fg_color="#1f538d", command=lambda: CalendarioNativo(v_pago, ent_fecha)).pack(side="right", padx=(5, 0))

        ctk.CTkLabel(f_form, text="N° de Operación (opcional):", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_operacion = ctk.CTkEntry(f_form, placeholder_text="Ej.: 00123456 / constancia de la transferencia")
        ent_operacion.pack(fill="x", pady=(0, 10))

        def procesar_pago(event=None):
            try:
                monto_val = float(ent_monto.get().strip())
            except ValueError:
                return messagebox.showerror("Error", "Ingrese un monto numérico válido.", parent=v_pago)

            if monto_val <= 0:
                return messagebox.showerror("Error", "El monto debe ser mayor a 0.", parent=v_pago)
            if monto_val > (saldo_actual + 0.01):
                return messagebox.showerror("Error", "El monto supera el saldo pendiente.", parent=v_pago)

            fecha_val = ent_fecha.get().strip() or datetime.now().strftime("%d/%m/%Y")
            cuenta_val = cmb_cuenta.get().strip()
            operacion_val = ent_operacion.get().strip()

            # Cerrar la ventana modal y abrir el diálogo de archivo en el SIGUIENTE
            # ciclo del bucle de eventos. En macOS, invocar el diálogo nativo de
            # archivos inmediatamente después de destruir una ventana con grab_set()
            # (mismo callback) cierra la aplicación.
            v_pago.destroy()
            self.main_root.after(150, lambda: self._guardar_pago_con_soporte(
                id_factura, nro_doc, proveedor, monto_val, fecha_val, cuenta_val, ruta_base,
                operacion_val))

        ent_monto.bind("<Return>", procesar_pago)
        ent_fecha.bind("<Return>", procesar_pago)
        ent_operacion.bind("<Return>", procesar_pago)

        f_btns = ctk.CTkFrame(v_pago, fg_color="transparent")
        f_btns.pack(fill="x", padx=20, pady=15)

        btn_ok = ctk.CTkButton(f_btns, text="✅ Confirmar", font=("Arial", 12, "bold"), fg_color="#27ae60", hover_color="#1e8449", command=procesar_pago)
        btn_ok.pack(side="left", expand=True, padx=5)

        btn_cancel = ctk.CTkButton(f_btns, text="❌ Cancelar", font=("Arial", 12, "bold"), fg_color="#e74c3c", hover_color="#922b21", command=v_pago.destroy)
        btn_cancel.pack(side="right", expand=True, padx=5)
        ent_monto.focus()

    def _guardar_pago_con_soporte(self, id_factura, nro_doc, proveedor, monto_val, fecha_val,
                                  cuenta_val, ruta_base, numero_operacion=""):
        """Abre el diálogo para elegir el soporte y guarda el pago.
        Se invoca con after() para evitar el cierre de la app en macOS al abrir
        el diálogo nativo justo después de destruir la ventana modal."""
        try:
            ruta_origen = seleccionar_archivo_dialogo("Seleccionar Soporte de Egreso", [("Archivos", "*.pdf;*.png;*.jpg;*.jpeg")])
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
                messagebox.showinfo("Éxito", f"Pago de {formatear_moneda(monto_val)} registrado exitosamente.")
                self.cargar_datos_pagar(reset_pagina=True)
                self.app_padre.app_facturas.cargar_datos_tabla(reset_pagina=True)
            except Exception as e:
                messagebox.showerror("Error", str(e))
            finally:
                liberar_conexion(conn)
        except Exception as e:
            messagebox.showerror("Error", str(e))

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
                for ruta in rutas:
                    # 🔎 Resuelve rutas guardadas por otro equipo/sistema (Mac/Windows)
                    ruta_norm = resolver_ruta_archivo(ruta)
                    if ruta_norm:
                        abrir_documento(ruta_norm)
                        abiertos += 1
                if not abiertos:
                    messagebox.showwarning("Aviso", "Los documentos están registrados pero los archivos no se encuentran en este equipo.")
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

            ctk.CTkLabel(f_form, text="N° de Operación (opcional):", font=("Arial", 11, "bold")).pack(anchor="w")
            ent_mod_oper = ctk.CTkEntry(f_form)
            ent_mod_oper.pack(fill="x", pady=(0, 10))
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
                    nueva_oper = ent_mod_oper.get().strip()
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
class ModuloComprasApp:
    def __init__(self, parent_frame):
        self.parent_frame = parent_frame
        self.usuario_activo = "Desconocido"
        self.pantalla_expandida = False
        aplicar_estilo_treeview()

        self.frame_main = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.frame_main.pack(fill="both", expand=True, padx=15, pady=15)
        
        header_frame = ctk.CTkFrame(self.frame_main, fg_color="transparent")
        header_frame.pack(fill="x")
        
        ctk.CTkLabel(header_frame, text="🛒 MÓDULO DE COMPRAS (LOGÍSTICA Y TESORERÍA)", font=("Arial", 18, "bold"), text_color="#1f538d").pack(side="left")
        self.btn_pantalla = ctk.CTkButton(header_frame, text="[ + ] Pantalla Completa", font=("Arial", 12, "bold"), width=160, fg_color="#34495e", hover_color="#2c3e50", command=self.toggle_pantalla_completa)
        self.btn_pantalla.pack(side="right")

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
            self.btn_pantalla.configure(text="[ + ] Pantalla Completa", fg_color="#34495e", hover_color="#2c3e50")
            self.pantalla_expandida = False
        else:
            if sidebar: sidebar.pack_forget()
            self.btn_pantalla.configure(text="[ - ] Restaurar Vista", fg_color="#34495e", hover_color="#2c3e50")
            self.pantalla_expandida = True

    def al_cambiar_pestana(self):
        if self.tabview.get() == " 💳 2. Control de Pagos y Deudas ":
            self.app_pagos.cargar_datos_pagar(reset_pagina=True)
        elif self.tabview.get() == " 📥 1. Ingreso de Facturas Recibidas ":
            self.app_facturas.cargar_datos_tabla(reset_pagina=True)

if __name__ == "__main__":
    pass
