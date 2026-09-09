# -*- coding: utf-8 -*-
"""
MODULO_BANCO.PY - MODULO DE BANCO (SALDOS Y CONCILIACION BANCARIA)
- Saldo de cada cuenta bancaria configurada en "Configuracion General":
    Saldo = Saldo Inicial + Ingresos Cobrados - Egresos Pagados.
- Conciliacion bancaria: subir el PDF del estado de cuenta del banco,
  extraer sus movimientos, compararlos contra los movimientos registrados
  en el sistema (cobros y pagos) y generar un reporte en pantalla.
- Permite conciliar, corregir montos/fechas y agregar movimientos faltantes
  (tanto ajustes manuales como cobros/pagos reales en el sistema).
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk
import os
import re
import json
import threading
from datetime import datetime

from conexion import conectar_db, registrar_auditoria, liberar_conexion
from app_paths import CONFIG_FILE

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    from pypdf import PdfReader as _PypdfReader
except ImportError:
    try:
        from PyPDF2 import PdfReader as _PypdfReader
    except ImportError:
        _PypdfReader = None

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")


# =========================================================
# CONFIGURACION Y UTILIDADES
# =========================================================
def cargar_config():
    config = {
        "simbolo_moneda": "S/.",
        "formato_numero": "1,000.00",
        "formato_fecha": "DD/MM/AAAA",
        "cuentas_bancarias": [],
    }
    try:
        if os.path.exists(str(CONFIG_FILE)):
            with open(str(CONFIG_FILE), "r", encoding="utf-8") as f:
                config.update(json.load(f))
    except Exception:
        pass
    return config


CONFIG = cargar_config()


def normalizar_monto(v):
    """Convierte cualquier texto de monto a float, aceptando ambos formatos
    (1,234.56 y 1.234,56) y simbolos de moneda."""
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v or "").strip()
    if not s:
        return 0.0
    s = re.sub(r"[^\d.,+-]", "", s)
    if not s:
        return 0.0
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        if re.fullmatch(r"[+-]?\d{1,3}(,\d{3})+", s):
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        dig = re.sub(r"[^\d.]", "", s)
        try:
            return float(dig)
        except ValueError:
            return 0.0


def formatear_monto(valor):
    simbolo = CONFIG.get("simbolo_moneda", "S/.")
    formato = CONFIG.get("formato_numero", "1,000.00")
    try:
        valor = float(valor)
    except Exception:
        valor = 0.0
    signo = "-" if valor < 0 else ""
    valor = abs(valor)
    if formato == "1.000,00":
        txt = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    else:
        txt = f"{valor:,.2f}"
    return f"{signo}{simbolo} {txt}"


def normalizar_fecha(s):
    s = (s or "").strip()
    if not s:
        return ""
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y", "%Y%m%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except Exception:
            continue
    return s


def construir_etiqueta_banco(b):
    return f"{b.get('banco', '')} - {b.get('cuenta', '')}".strip(" -")


def _norm_texto(s):
    return re.sub(r"\s+", " ", (s or "")).strip().lower()


def banco_matchea(banco, texto_cuenta):
    """Determina si un texto de cuenta (cuenta_destino / cuenta_origen)
    corresponde al banco indicado."""
    t = _norm_texto(texto_cuenta)
    if not t:
        return False
    nombre = _norm_texto(banco.get("banco", ""))
    cuenta = _norm_texto(banco.get("cuenta", ""))
    etiqueta = _norm_texto(construir_etiqueta_banco(banco))
    if t == etiqueta and etiqueta:
        return True
    if nombre and nombre in t:
        return True
    if cuenta:
        digitos_cuenta = re.sub(r"\D", "", cuenta)
        digitos_t = re.sub(r"\D", "", t)
        if len(digitos_cuenta) >= 4 and digitos_cuenta in digitos_t:
            return True
    return False


def aplicar_estilo_treeview():
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("Treeview", background="#ffffff", foreground="#000000",
                    fieldbackground="#ffffff", bordercolor="#e0e0e0",
                    borderwidth=1, rowheight=26, font=("Arial", 10))
    style.map("Treeview", background=[("selected", "#1f538d")],
              foreground=[("selected", "#ffffff")])
    style.configure("Treeview.Heading", background="#f0f0f0", foreground="#000000",
                    relief="flat", font=("Arial", 10, "bold"),
                    bordercolor="#e0e0e0", borderwidth=1)


# =========================================================
# EXTRACCION Y ANALISIS DE PDF (ESTADO DE CUENTA)
# =========================================================
def extraer_texto_pdf(ruta):
    texto = ""
    if pdfplumber is not None:
        try:
            with pdfplumber.open(ruta) as pdf:
                partes = [p.extract_text() or "" for p in pdf.pages]
            texto = "\n".join(partes).strip()
        except Exception:
            texto = ""
    if not texto and fitz is not None:
        try:
            doc = fitz.open(ruta)
            partes = [page.get_text() or "" for page in doc]
            doc.close()
            texto = "\n".join(partes).strip()
        except Exception:
            texto = ""
    if not texto and _PypdfReader is not None:
        try:
            reader = _PypdfReader(ruta)
            partes = [(p.extract_text() or "") for p in reader.pages]
            texto = "\n".join(partes).strip()
        except Exception:
            texto = ""
    return texto


FECHA_RE = re.compile(r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b")
MONTO_RE = re.compile(r"[-+]?\s?\d{1,3}(?:[.,]\d{3})+(?:[.,]\d{1,2})?|[-+]?\s?\d+(?:[.,]\d{1,2})?")

PALABRAS_ABONO = ["abono", "abon", "deposito", "depósito", "credito", "crédito", "ingreso",
                  "entrada", "transferencia recibida", "recibido", "cobro", "cobranza", "dep"]
PALABRAS_CARGO = ["cargo", "retiro", "debito", "débito", "pago", "egreso", "salida",
                  "transferencia enviada", "compra", "ret", "cargos"]


def _signo_por_palabras(linea):
    low = linea.lower()
    for p in PALABRAS_ABONO:
        if p in low:
            return 1
    for p in PALABRAS_CARGO:
        if p in low:
            return -1
    return None


def _seleccionar_monto(linea, montos):
    """montos: lista de (texto, posicion). Devuelve (valor, token) del movimiento."""
    if not montos:
        return None
    if len(montos) >= 2:
        candidatos = montos[:-1]
        vals = []
        for tok, pos in candidatos:
            v = normalizar_monto(tok)
            vals.append((v, tok))
        no_cero = [x for x in vals if abs(x[0]) > 0.0009]
        if no_cero:
            return max(no_cero, key=lambda x: abs(x[0]))
        return vals[-1]
    tok, pos = montos[0]
    return normalizar_monto(tok), tok


def parsear_movimientos(texto):
    """Heuristica para extraer movimientos de un estado de cuenta."""
    movimientos = []
    if not texto:
        return movimientos
    for linea in texto.splitlines():
        linea = linea.strip()
        if not linea:
            continue
        low = linea.lower()
        # Encabezados / rangos de periodo no son transacciones
        if any(k in low for k in ("periodo", "period", "estado de cuenta", "saldo inicial")):
            continue
        fechas = FECHA_RE.findall(linea)
        if len(fechas) >= 2:
            continue
        mf = FECHA_RE.search(linea)
        if not mf:
            continue
        fecha = mf.group(1)
        resto = linea[:mf.start()] + " " + linea[mf.end():]
        montos = [(m.group(0), m.start()) for m in MONTO_RE.finditer(resto)]
        if not montos:
            continue
        seleccion = _seleccionar_monto(resto, montos)
        if seleccion is None:
            continue
        valor, token = seleccion
        signo_palabras = _signo_por_palabras(linea)
        signo_explicito = token.strip().startswith("+") or token.strip().startswith("-")
        if not signo_explicito and signo_palabras is not None:
            valor = signo_palabras * abs(valor)
        descripcion = re.sub(r"\s+", " ", resto).strip()
        descripcion = MONTO_RE.sub("", descripcion)
        descripcion = re.sub(r"\s+", " ", descripcion).strip(" -|")
        movimientos.append({
            "fecha": normalizar_fecha(fecha) or fecha,
            "descripcion": descripcion or "Movimiento sin descripción",
            "monto": valor,
        })
    return movimientos


def detectar_saldo_final(texto):
    """Intenta detectar el saldo final del estado de cuenta."""
    if not texto:
        return None
    lineas_saldo = [ln for ln in texto.splitlines() if "saldo" in ln.lower()]
    if not lineas_saldo:
        return None
    ultima = lineas_saldo[-1]
    montos = [(m.group(0), m.start()) for m in MONTO_RE.finditer(ultima)]
    if not montos:
        return None
    return normalizar_monto(montos[-1][0])


# =========================================================
# CLASE PRINCIPAL
# =========================================================
class ModuloBancoApp:
    def __init__(self, parent_frame, usuario_activo=""):
        self.parent_frame = parent_frame
        self.usuario_activo = usuario_activo or "Desconocido"
        self.config = cargar_config()
        self.bancos = self.config.get("cuentas_bancarias", []) or []
        self.movimientos_pdf = []
        self.texto_pdf = ""
        self.filas_conciliacion = []
        self.saldo_final_estado = None
        self.banco_conciliado = None
        aplicar_estilo_treeview()
        self.inicializar_db()

        self.frame_main = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.frame_main.pack(fill="both", expand=True, padx=15, pady=15)

        header = ctk.CTkFrame(self.frame_main, fg_color="transparent")
        header.pack(fill="x")
        ctk.CTkLabel(header, text="🏦 MÓDULO DE BANCO (SALDOS Y CONCILIACIÓN BANCARIA)",
                     font=("Arial", 18, "bold"), text_color="#1f538d").pack(side="left")

        self.tabview = ctk.CTkTabview(self.frame_main, segmented_button_selected_color="#1f538d")
        self.tabview.pack(fill="both", expand=True, pady=(10, 0))
        self.tab_saldos = self.tabview.add(" 💵 Saldo de Bancos ")
        self.tab_conciliacion = self.tabview.add(" 🧾 Conciliación Bancaria ")
        self.tab_transferencias = self.tabview.add(" 🔄 Transferencias ")

        self.construir_tab_saldos()
        self.construir_tab_conciliacion()
        self.construir_tab_transferencias()

    # -----------------------------------------------------
    # BASE DE DATOS DE CONCILIACION
    # -----------------------------------------------------
    def inicializar_db(self):
        conn = conectar_db(silencioso=True)
        if not conn:
            return
        try:
            with conn.cursor() as c:
                c.execute("""
                    CREATE TABLE IF NOT EXISTS conciliacion_bancaria (
                        id SERIAL PRIMARY KEY,
                        banco VARCHAR(255),
                        cuenta VARCHAR(255),
                        fecha VARCHAR(20),
                        descripcion TEXT,
                        monto NUMERIC,
                        tipo VARCHAR(20),
                        origen VARCHAR(30),
                        id_movimiento INTEGER DEFAULT 0,
                        estado VARCHAR(20) DEFAULT 'pendiente',
                        fecha_conciliacion VARCHAR(20) DEFAULT ''
                    )
                """)
                c.execute("""
                    CREATE TABLE IF NOT EXISTS transferencias_bancarias (
                        id SERIAL PRIMARY KEY,
                        banco_origen VARCHAR(255),
                        cuenta_origen VARCHAR(255),
                        banco_destino VARCHAR(255),
                        cuenta_destino VARCHAR(255),
                        fecha VARCHAR(20),
                        descripcion TEXT,
                        monto NUMERIC
                    )
                """)
                conn.commit()
        except Exception:
            pass
        finally:
            liberar_conexion(conn)

    # -----------------------------------------------------
    # TAB: SALDO DE BANCOS
    # -----------------------------------------------------
    def construir_tab_saldos(self):
        f_top = ctk.CTkFrame(self.tab_saldos, fg_color="transparent")
        f_top.pack(fill="x", pady=(5, 10))
        ctk.CTkLabel(f_top, text="Saldo calculado como: Saldo Inicial + Ingresos Cobrados − Egresos Pagados.",
                     font=("Arial", 12, "italic"), text_color="gray").pack(side="left")
        ctk.CTkButton(f_top, text="🔄 Actualizar", width=130, font=("Arial", 12, "bold"),
                      fg_color="#1f538d", hover_color="#163b65",
                      command=self.refrescar_saldos).pack(side="right")

        self.scroll_saldos = ctk.CTkScrollableFrame(self.tab_saldos, fg_color="transparent")
        self.scroll_saldos.pack(fill="both", expand=True)
        self.refrescar_saldos()

    def refrescar_saldos(self):
        for w in self.scroll_saldos.winfo_children():
            w.destroy()

        if not self.bancos:
            f_aviso = ctk.CTkFrame(self.scroll_saldos, corner_radius=10, fg_color="#fff3cd",
                                   border_width=1, border_color="#ffeeba")
            f_aviso.pack(fill="x", pady=10, padx=5)
            ctk.CTkLabel(f_aviso, text="⚠️ No hay cuentas bancarias configuradas.\n"
                                       "Agréguelas en: Ajustes de Sistema → ⚙️ Configuración General → Cuentas Bancarias.",
                         font=("Arial", 13), text_color="#856404", justify="left").pack(padx=15, pady=15)
            return

        total_general = 0.0
        for idx, banco in enumerate(self.bancos):
            resumen = self.calcular_resumen(banco)
            total_general += resumen["saldo_actual"]
            self._crear_tarjeta_banco(banco, resumen, idx)

        f_total = ctk.CTkFrame(self.scroll_saldos, corner_radius=10, fg_color="#1f538d")
        f_total.pack(fill="x", pady=(15, 5), padx=5)
        ctk.CTkLabel(f_total, text="TOTAL DISPONIBLE EN BANCOS", font=("Arial", 14, "bold"),
                     text_color="white").pack(side="left", padx=15, pady=12)
        ctk.CTkLabel(f_total, text=formatear_monto(total_general), font=("Arial", 16, "bold"),
                     text_color="white").pack(side="right", padx=15, pady=12)

    def _crear_tarjeta_banco(self, banco, resumen, idx):
        etiqueta = construir_etiqueta_banco(banco) or f"Banco {idx + 1}"
        card = ctk.CTkFrame(self.scroll_saldos, corner_radius=10, fg_color="#f8f9fa",
                            border_width=1, border_color="#e0e0e0")
        card.pack(fill="x", pady=6, padx=5)

        fila_titulo = ctk.CTkFrame(card, fg_color="transparent")
        fila_titulo.pack(fill="x", padx=12, pady=(10, 4))
        ctk.CTkLabel(fila_titulo, text=f"🏦 {etiqueta}", font=("Arial", 15, "bold"),
                     text_color="#1f538d").pack(side="left")
        ctk.CTkLabel(fila_titulo, text=resumen["saldo_actual_txt"], font=("Arial", 16, "bold"),
                     text_color="#27ae60" if resumen["saldo_actual"] >= 0 else "#c0392b").pack(side="right")

        fila_det = ctk.CTkFrame(card, fg_color="transparent")
        fila_det.pack(fill="x", padx=12, pady=(0, 6))
        ctk.CTkLabel(fila_det, text=f"Saldo inicial: {formatear_monto(resumen['saldo_inicial'])}",
                     font=("Arial", 11), text_color="gray").pack(side="left", padx=(0, 15))
        ctk.CTkLabel(fila_det, text=f"+ Ingresos: {formatear_monto(resumen['ingresos'])}",
                     font=("Arial", 11), text_color="#27ae60").pack(side="left", padx=(0, 15))
        ctk.CTkLabel(fila_det, text=f"− Egresos: {formatear_monto(resumen['egresos'])}",
                     font=("Arial", 11), text_color="#c0392b").pack(side="left", padx=(0, 15))
        ctk.CTkLabel(fila_det, text=f"({resumen['n_movimientos']} movimientos)",
                     font=("Arial", 10, "italic"), text_color="#7f8c8d").pack(side="left")

        fila_btn = ctk.CTkFrame(card, fg_color="transparent")
        fila_btn.pack(fill="x", padx=12, pady=(0, 10))
        ctk.CTkButton(fila_btn, text="👁 Ver Movimientos", width=160, font=("Arial", 11, "bold"),
                      fg_color="#34495e", hover_color="#2c3e50",
                      command=lambda b=banco: self.ver_movimientos(b)).pack(side="left", padx=(0, 8))
        ctk.CTkButton(fila_btn, text="🧾 Conciliar este Banco", width=180, font=("Arial", 11, "bold"),
                      fg_color="#1f538d", hover_color="#163b65",
                      command=lambda b=banco: self.ir_a_conciliacion(b)).pack(side="left")

    def calcular_resumen(self, banco):
        movs = self.cargar_movimientos_sistema(banco)
        ingresos = sum(m["monto"] for m in movs if m["monto"] > 0)
        egresos = sum(-m["monto"] for m in movs if m["monto"] < 0)
        saldo_inicial = normalizar_monto(banco.get("saldo_inicial", ""))
        saldo = saldo_inicial + ingresos - egresos
        return {
            "saldo_inicial": saldo_inicial,
            "ingresos": ingresos,
            "egresos": egresos,
            "saldo_actual": saldo,
            "saldo_actual_txt": formatear_monto(saldo),
            "n_movimientos": len(movs),
            "movimientos": movs,
        }

    def cargar_movimientos_sistema(self, banco):
        """Cobros (ingresos) y pagos (egresos) registrados para este banco."""
        conn = conectar_db(silencioso=True)
        if not conn:
            return []
        movs = []
        try:
            with conn.cursor() as c:
                try:
                    c.execute("""
                        SELECT p.id, p.fecha_pago, p.cliente_nombre, p.monto_pagado,
                               COALESCE(p.cuenta_destino, ''), p.id_factura,
                               COALESCE(f.numero_documento, ''), COALESCE(p.codigo_cotizacion, '')
                        FROM pagos_clientes p
                        LEFT JOIN facturas_emitidas f ON p.id_factura = f.id
                    """)
                    for idp, fecha, nombre, monto, cuenta, idf, num_doc, ref in c.fetchall():
                        if banco_matchea(banco, cuenta):
                            doc = (num_doc or ref or "").strip()
                            desc = f"Cobro {nombre or ''}".replace("  ", " ").strip()
                            movs.append({
                                "id": idp, "fecha": fecha or "", "descripcion": desc,
                                "monto": float(monto or 0), "tipo": "ingreso",
                                "origen": "sistema", "tabla": "pagos_clientes", "documento": doc,
                            })
                except Exception:
                    conn.rollback()
                try:
                    c.execute("""
                        SELECT p.id, p.fecha_pago, p.proveedor_nombre, p.monto_pagado,
                               COALESCE(p.cuenta_origen, ''), p.id_factura,
                               COALESCE(f.numero_documento, ''), COALESCE(p.codigo_cotizacion, '')
                        FROM pagos_comprobantes p
                        LEFT JOIN facturas_recibidas f ON p.id_factura = f.id
                    """)
                    for idp, fecha, nombre, monto, cuenta, idf, num_doc, ref in c.fetchall():
                        if banco_matchea(banco, cuenta):
                            doc = (num_doc or ref or "").strip()
                            desc = f"Pago {nombre or ''}".replace("  ", " ").strip()
                            movs.append({
                                "id": idp, "fecha": fecha or "", "descripcion": desc,
                                "monto": -float(monto or 0), "tipo": "egreso",
                                "origen": "sistema", "tabla": "pagos_comprobantes", "documento": doc,
                            })
                except Exception:
                    conn.rollback()
                try:
                    c.execute("""
                        SELECT id, banco_origen, cuenta_origen, banco_destino, cuenta_destino,
                               fecha, descripcion, monto
                        FROM transferencias_bancarias
                    """)
                    for idt, bo, co, bd, cd, fecha, desc, monto in c.fetchall():
                        monto_v = float(monto or 0)
                        et_origen = construir_etiqueta_banco({"banco": bo, "cuenta": co})
                        et_destino = construir_etiqueta_banco({"banco": bd, "cuenta": cd})
                        if banco_matchea(banco, et_origen):
                            desc_eg = f"Transferencia a {bd or 'otra cuenta'}"
                            if desc:
                                desc_eg += f" ({desc})"
                            movs.append({
                                "id": idt, "fecha": fecha or "", "descripcion": desc_eg,
                                "monto": -monto_v, "tipo": "transferencia",
                                "origen": "sistema", "tabla": "transferencias_bancarias", "documento": "",
                            })
                        if banco_matchea(banco, et_destino):
                            desc_in = f"Transferencia de {bo or 'otra cuenta'}"
                            if desc:
                                desc_in += f" ({desc})"
                            movs.append({
                                "id": idt, "fecha": fecha or "", "descripcion": desc_in,
                                "monto": monto_v, "tipo": "transferencia",
                                "origen": "sistema", "tabla": "transferencias_bancarias", "documento": "",
                            })
                except Exception:
                    conn.rollback()
        except Exception:
            pass
        finally:
            liberar_conexion(conn)
        movs.sort(key=lambda m: normalizar_fecha(m["fecha"]))
        return movs

    def ver_movimientos(self, banco):
        resumen = self.calcular_resumen(banco)
        movs = resumen["movimientos"]
        v = ctk.CTkToplevel(self.parent_frame)
        v.title(f"Movimientos - {construir_etiqueta_banco(banco)}")
        v.geometry("760x480")
        v.transient(self.parent_frame)

        ctk.CTkLabel(v, text=f"Movimientos de {construir_etiqueta_banco(banco)}",
                     font=("Arial", 15, "bold"), text_color="#1f538d").pack(pady=(15, 5))

        f_tabla = ctk.CTkFrame(v, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True, padx=15, pady=10)
        columnas = ("fecha", "documento", "descripcion", "tipo", "monto")
        tabla = ttk.Treeview(f_tabla, columns=columnas, show="headings", height=14)
        tabla.heading("fecha", text="Fecha")
        tabla.heading("documento", text="N° Documento")
        tabla.heading("descripcion", text="Descripción")
        tabla.heading("tipo", text="Tipo")
        tabla.heading("monto", text="Monto")
        tabla.column("fecha", width=95, anchor="center")
        tabla.column("documento", width=150, anchor="center")
        tabla.column("descripcion", width=260, anchor="w")
        tabla.column("tipo", width=110, anchor="center")
        tabla.column("monto", width=130, anchor="e")
        vsb = ttk.Scrollbar(f_tabla, orient="vertical", command=tabla.yview)
        tabla.configure(yscrollcommand=vsb.set)
        tabla.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        movs = sorted(movs, key=lambda m: (normalizar_fecha(m["fecha"]), m["monto"]))
        for m in movs:
            if m.get("tipo") == "transferencia":
                tipo_txt = "Transferencia (Ingreso)" if m["monto"] > 0 else "Transferencia (Egreso)"
            else:
                tipo_txt = "Ingreso (Cobro)" if m["monto"] > 0 else "Egreso (Pago)"
            tabla.insert("", tk.END, values=(m["fecha"], m.get("documento", ""), m["descripcion"],
                                             tipo_txt, formatear_monto(m["monto"])))
        ctk.CTkButton(v, text="Cerrar", width=120, fg_color="#34495e",
                      command=v.destroy).pack(pady=(0, 15))

    # -----------------------------------------------------
    # TAB: TRANSFERENCIAS ENTRE CUENTAS
    # -----------------------------------------------------
    def construir_tab_transferencias(self):
        f_form = ctk.CTkFrame(self.tab_transferencias, fg_color="#f8f9fa",
                              border_width=1, border_color="#e0e0e0", corner_radius=8)
        f_form.pack(fill="x", padx=5, pady=(5, 10))

        etiquetas = [construir_etiqueta_banco(b) for b in self.bancos]
        if not etiquetas:
            etiquetas = ["(Sin bancos configurados)"]

        r1 = ctk.CTkFrame(f_form, fg_color="transparent"); r1.pack(fill="x", padx=12, pady=(12, 4))
        ctk.CTkLabel(r1, text="Banco Origen:", width=130, anchor="w", font=("Arial", 12, "bold")).pack(side="left")
        self.cmb_tx_origen = ctk.CTkComboBox(r1, values=etiquetas, width=340, state="readonly")
        self.cmb_tx_origen.pack(side="left", padx=6)
        if etiquetas:
            self.cmb_tx_origen.set(etiquetas[0])

        r2 = ctk.CTkFrame(f_form, fg_color="transparent"); r2.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(r2, text="Banco Destino:", width=130, anchor="w", font=("Arial", 12, "bold")).pack(side="left")
        self.cmb_tx_destino = ctk.CTkComboBox(r2, values=etiquetas, width=340, state="readonly")
        self.cmb_tx_destino.pack(side="left", padx=6)
        if len(etiquetas) > 1:
            self.cmb_tx_destino.set(etiquetas[1])
        elif etiquetas:
            self.cmb_tx_destino.set(etiquetas[0])

        r3 = ctk.CTkFrame(f_form, fg_color="transparent"); r3.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(r3, text="Fecha:", width=130, anchor="w", font=("Arial", 12, "bold")).pack(side="left")
        self.ent_tx_fecha = ctk.CTkEntry(r3, width=170)
        self.ent_tx_fecha.pack(side="left", padx=6)
        self.ent_tx_fecha.insert(0, datetime.now().strftime("%d/%m/%Y"))
        ctk.CTkLabel(r3, text="Monto:", width=90, anchor="w", font=("Arial", 12, "bold")).pack(side="left", padx=(16, 0))
        self.ent_tx_monto = ctk.CTkEntry(r3, width=170)
        self.ent_tx_monto.pack(side="left", padx=6)

        r4 = ctk.CTkFrame(f_form, fg_color="transparent"); r4.pack(fill="x", padx=12, pady=(4, 4))
        ctk.CTkLabel(r4, text="Descripción:", width=130, anchor="w", font=("Arial", 12, "bold")).pack(side="left")
        self.ent_tx_desc = ctk.CTkEntry(r4)
        self.ent_tx_desc.pack(side="left", fill="x", expand=True, padx=6)

        r5 = ctk.CTkFrame(f_form, fg_color="transparent"); r5.pack(fill="x", padx=12, pady=(4, 12))
        ctk.CTkButton(r5, text="💸 Realizar Transferencia", width=240, height=38,
                      font=("Arial", 13, "bold"), fg_color="#27ae60", hover_color="#1e8449",
                      command=self.realizar_transferencia).pack(side="left")

        ctk.CTkLabel(self.tab_transferencias, text="Historial de transferencias:",
                     font=("Arial", 13, "bold"), text_color="#1f538d").pack(anchor="w", padx=5, pady=(0, 4))
        f_tabla = ctk.CTkFrame(self.tab_transferencias, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True, padx=5, pady=(0, 5))
        columnas = ("fecha", "origen", "destino", "descripcion", "monto")
        self.tabla_tx = ttk.Treeview(f_tabla, columns=columnas, show="headings")
        for col, txt, w, anc in (("fecha", "Fecha", 100, "center"),
                                 ("origen", "Banco Origen", 200, "center"),
                                 ("destino", "Banco Destino", 200, "center"),
                                 ("descripcion", "Descripción", 260, "w"),
                                 ("monto", "Monto", 130, "e")):
            self.tabla_tx.heading(col, text=txt)
            self.tabla_tx.column(col, width=w, anchor=anc)
        vsb = ttk.Scrollbar(f_tabla, orient="vertical", command=self.tabla_tx.yview)
        self.tabla_tx.configure(yscrollcommand=vsb.set)
        self.tabla_tx.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.refrescar_transferencias()

    def realizar_transferencia(self):
        etiquetas = [construir_etiqueta_banco(b) for b in self.bancos]
        origen = self.cmb_tx_origen.get()
        destino = self.cmb_tx_destino.get()
        if origen not in etiquetas or destino not in etiquetas:
            messagebox.showwarning("Transferencia", "Seleccione bancos válidos.", parent=self.parent_frame)
            return
        if origen == destino:
            messagebox.showwarning("Transferencia", "El banco origen y destino deben ser diferentes.", parent=self.parent_frame)
            return
        monto = normalizar_monto(self.ent_tx_monto.get())
        if monto <= 0:
            messagebox.showerror("Error", "Ingrese un monto mayor a 0.", parent=self.parent_frame)
            return
        fecha = self.ent_tx_fecha.get().strip() or datetime.now().strftime("%d/%m/%Y")
        desc = self.ent_tx_desc.get().strip() or "Transferencia entre cuentas"

        bo = self.bancos[etiquetas.index(origen)]
        bd = self.bancos[etiquetas.index(destino)]

        conn = conectar_db(silencioso=True)
        if not conn:
            messagebox.showerror("Error", "Sin conexión a la base de datos.", parent=self.parent_frame)
            return
        try:
            with conn.cursor() as c:
                c.execute("""
                    INSERT INTO transferencias_bancarias
                    (banco_origen, cuenta_origen, banco_destino, cuenta_destino, fecha, descripcion, monto)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                """, (bo.get("banco", ""), bo.get("cuenta", ""), bd.get("banco", ""), bd.get("cuenta", ""),
                      fecha, desc, monto))
                conn.commit()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo registrar la transferencia:\n{e}", parent=self.parent_frame)
            return
        finally:
            liberar_conexion(conn)

        registrar_auditoria(self.usuario_activo, "Banco",
                            f"Transferencia {formatear_monto(monto)} de {origen} a {destino}")
        self.ent_tx_monto.delete(0, tk.END)
        self.ent_tx_desc.delete(0, tk.END)
        self.refrescar_transferencias()
        self.refrescar_saldos()
        messagebox.showinfo("Éxito", f"Transferencia de {formatear_monto(monto)} registrada:\n"
                                     f"• Egreso en {origen}\n• Ingreso en {destino}", parent=self.parent_frame)

    def cargar_transferencias(self):
        conn = conectar_db(silencioso=True)
        if not conn:
            return []
        rows = []
        try:
            with conn.cursor() as c:
                c.execute("""
                    SELECT id, banco_origen, cuenta_origen, banco_destino, cuenta_destino,
                           fecha, descripcion, monto
                    FROM transferencias_bancarias
                    ORDER BY id DESC
                """)
                for idt, bo, co, bd, cd, fecha, desc, monto in c.fetchall():
                    rows.append({
                        "id": idt,
                        "origen": construir_etiqueta_banco({"banco": bo, "cuenta": co}),
                        "destino": construir_etiqueta_banco({"banco": bd, "cuenta": cd}),
                        "fecha": fecha or "", "descripcion": desc or "", "monto": float(monto or 0),
                    })
        except Exception:
            pass
        finally:
            liberar_conexion(conn)
        return rows

    def refrescar_transferencias(self):
        self.tabla_tx.delete(*self.tabla_tx.get_children())
        for t in self.cargar_transferencias():
            self.tabla_tx.insert("", tk.END, values=(t["fecha"], t["origen"], t["destino"],
                                                     t["descripcion"], formatear_monto(t["monto"])))

    # -----------------------------------------------------
    # TAB: CONCILIACION BANCARIA
    # -----------------------------------------------------
    def construir_tab_conciliacion(self):
        f_sel = ctk.CTkFrame(self.tab_conciliacion, fg_color="#f8f9fa",
                             border_width=1, border_color="#e0e0e0", corner_radius=8)
        f_sel.pack(fill="x", padx=5, pady=(5, 10))

        ctk.CTkLabel(f_sel, text="Banco a conciliar:", font=("Arial", 12, "bold")).pack(side="left", padx=(12, 6), pady=10)
        etiquetas = [construir_etiqueta_banco(b) for b in self.bancos]
        if not etiquetas:
            etiquetas = ["(Sin bancos configurados)"]
        self.cmb_banco = ctk.CTkComboBox(f_sel, values=etiquetas, width=320, state="readonly")
        self.cmb_banco.pack(side="left", padx=6, pady=10)
        if etiquetas:
            self.cmb_banco.set(etiquetas[0])

        f_btns = ctk.CTkFrame(self.tab_conciliacion, fg_color="transparent")
        f_btns.pack(fill="x", padx=5, pady=(0, 8))
        ctk.CTkButton(f_btns, text="📄 Subir PDF Estado de Cuenta", font=("Arial", 12, "bold"),
                      fg_color="#e67e22", hover_color="#ca6f1e",
                      command=self.subir_pdf).pack(side="left", padx=4)
        ctk.CTkButton(f_btns, text="🔄 Cargar Movimientos del Sistema", font=("Arial", 12, "bold"),
                      fg_color="#1f538d", hover_color="#163b65",
                      command=self.generar_reporte).pack(side="left", padx=4)
        ctk.CTkButton(f_btns, text="✅ Conciliar Seleccionados", font=("Arial", 12, "bold"),
                      fg_color="#27ae60", hover_color="#1e8449",
                      command=self.conciliar_seleccionados).pack(side="left", padx=4)
        ctk.CTkButton(f_btns, text="➕ Agregar Movimiento", font=("Arial", 12, "bold"),
                      fg_color="#8e44ad", hover_color="#703688",
                      command=self.agregar_movimiento_manual).pack(side="left", padx=4)
        ctk.CTkButton(f_btns, text="🧾 Registrar Cobro/Pago", font=("Arial", 12, "bold"),
                      fg_color="#2980b9", hover_color="#1f618d",
                      command=self.registrar_pago_sistema).pack(side="left", padx=4)
        ctk.CTkButton(f_btns, text="✏️ Corregir", font=("Arial", 12, "bold"),
                      fg_color="#34495e", hover_color="#2c3e50",
                      command=self.corregir_movimiento).pack(side="left", padx=4)
        ctk.CTkButton(f_btns, text="💾 Exportar", font=("Arial", 12, "bold"),
                      fg_color="#7f8c8d", hover_color="#606b6b",
                      command=self.exportar_reporte).pack(side="left", padx=4)

        self.f_resumen = ctk.CTkFrame(self.tab_conciliacion, fg_color="#eaf2f8",
                                      border_width=1, border_color="#d4e6f1", corner_radius=8)
        self.f_resumen.pack(fill="x", padx=5, pady=(0, 8))
        self.lbl_resumen = ctk.CTkLabel(self.f_resumen, text="Seleccione un banco, suba el PDF del estado de cuenta "
                                                             "y cargue los movimientos del sistema.",
                                        font=("Arial", 12), text_color="#1f538d", justify="left")
        self.lbl_resumen.pack(anchor="w", padx=12, pady=10)

        f_tabla = ctk.CTkFrame(self.tab_conciliacion, fg_color="transparent")
        f_tabla.pack(fill="both", expand=True, padx=5, pady=(0, 5))
        columnas = ("estado", "fecha", "documento", "descripcion", "tipo", "monto", "origen")
        self.tabla_conc = ttk.Treeview(f_tabla, columns=columnas, show="headings",
                                       selectmode="extended")
        for col, txt, w, anc in (("estado", "Estado", 140, "center"),
                                 ("fecha", "Fecha", 90, "center"),
                                 ("documento", "N° Documento", 140, "center"),
                                 ("descripcion", "Descripción", 250, "w"),
                                 ("tipo", "Tipo", 100, "center"),
                                 ("monto", "Monto", 110, "e"),
                                 ("origen", "Origen", 120, "center")):
            self.tabla_conc.heading(col, text=txt)
            self.tabla_conc.column(col, width=w, anchor=anc)
        vsb = ttk.Scrollbar(f_tabla, orient="vertical", command=self.tabla_conc.yview)
        self.tabla_conc.configure(yscrollcommand=vsb.set)
        self.tabla_conc.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tabla_conc.tag_configure("conciliado", background="#d5f5e3")
        self.tabla_conc.tag_configure("diferencia", background="#fadbd8")
        self.tabla_conc.tag_configure("estado_cuenta", background="#d6eaf8")

        f_pie = ctk.CTkFrame(self.tab_conciliacion, fg_color="transparent")
        f_pie.pack(fill="x", padx=5)
        ctk.CTkButton(f_pie, text="👁 Ver Texto Extraído del PDF", width=220, font=("Arial", 11),
                      fg_color="#7f8c8d", hover_color="#606b6b",
                      command=self.mostrar_texto_pdf).pack(side="left")

    def banco_seleccionado(self):
        etiquetas = [construir_etiqueta_banco(b) for b in self.bancos]
        sel = self.cmb_banco.get()
        if sel in etiquetas:
            return self.bancos[etiquetas.index(sel)]
        if self.bancos:
            return self.bancos[0]
        return None

    def ir_a_conciliacion(self, banco):
        etiquetas = [construir_etiqueta_banco(b) for b in self.bancos]
        et = construir_etiqueta_banco(banco)
        if et in etiquetas:
            self.cmb_banco.set(et)
        self.tabview.set(" 🧾 Conciliación Bancaria ")

    def subir_pdf(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar Estado de Cuenta (PDF)",
            filetypes=[("Archivos PDF", "*.pdf")])
        if not ruta:
            return
        self.lbl_resumen.configure(text="⏳ Extrayendo texto del PDF...")
        self.parent_frame.update_idletasks()

        def tarea():
            texto = extraer_texto_pdf(ruta)
            if not texto.strip():
                self.lbl_resumen.configure(text="❌ No se pudo extraer texto del PDF. "
                                                "Verifique que no sea un PDF escaneado (imagen).")
                return
            movs = parsear_movimientos(texto)
            self.texto_pdf = texto
            self.movimientos_pdf = movs
            self.saldo_final_estado = detectar_saldo_final(texto)
            self.lbl_resumen.configure(text=f"✅ PDF procesado: {len(movs)} movimientos detectados.")
            self.generar_reporte()

        threading.Thread(target=tarea, daemon=True).start()

    def generar_reporte(self):
        banco = self.banco_seleccionado()
        if not banco:
            self.lbl_resumen.configure(text="⚠️ No hay banco seleccionado.")
            return

        sys_movs = self.cargar_movimientos_sistema(banco)
        manuales = self.cargar_manuales(banco)
        bank_movs = list(self.movimientos_pdf)

        filas = []
        usados = set()

        for s in sys_movs:
            s_fecha = normalizar_fecha(s["fecha"])
            idx_match = None
            for i, b in enumerate(bank_movs):
                if i in usados:
                    continue
                if s_fecha and normalizar_fecha(b["fecha"]) == s_fecha and abs(b["monto"] - s["monto"]) <= 0.02:
                    idx_match = i
                    break
            if idx_match is None:
                cands = [i for i, b in enumerate(bank_movs)
                         if i not in usados and abs(b["monto"] - s["monto"]) <= 0.02]
                if len(cands) == 1:
                    idx_match = cands[0]
            if idx_match is not None:
                usados.add(idx_match)
                b = bank_movs[idx_match]
                filas.append({**s, "estado": "conciliado"})
                filas.append({**b, "id": 0, "tipo": "ingreso" if b["monto"] > 0 else "egreso",
                              "origen": "estado_cuenta", "tabla": "", "estado": "conciliado"})
            else:
                filas.append({**s, "estado": "diferencia"})

        for i, b in enumerate(bank_movs):
            if i not in usados:
                filas.append({**b, "id": 0, "tipo": "ingreso" if b["monto"] > 0 else "egreso",
                              "origen": "estado_cuenta", "tabla": "", "estado": "diferencia"})

        for m in manuales:
            filas.append({**m, "estado": m.get("estado", "pendiente"), "origen": "manual"})

        filas.sort(key=lambda f: (normalizar_fecha(f["fecha"]), f["monto"]))
        self.filas_conciliacion = filas
        self.refrescar_tree()

        saldo_inicial = normalizar_monto(banco.get("saldo_inicial", ""))
        ingresos = sum(m["monto"] for m in sys_movs if m["monto"] > 0)
        egresos = sum(-m["monto"] for m in sys_movs if m["monto"] < 0)
        saldo_sistema = saldo_inicial + ingresos - egresos

        if self.saldo_final_estado is not None:
            diferencia = saldo_sistema - self.saldo_final_estado
            txt = (f"Banco: {construir_etiqueta_banco(banco)}   |   "
                   f"Saldo Sistema: {formatear_monto(saldo_sistema)}   |   "
                   f"Saldo Estado de Cuenta: {formatear_monto(self.saldo_final_estado)}   |   "
                   f"Diferencia: {formatear_monto(diferencia)}")
        else:
            neto_pdf = sum(b["monto"] for b in bank_movs)
            txt = (f"Banco: {construir_etiqueta_banco(banco)}   |   "
                   f"Saldo Sistema: {formatear_monto(saldo_sistema)}   |   "
                   f"Movimientos PDF (neto): {formatear_monto(neto_pdf)}   |   "
                   f"(No se detectó saldo final en el PDF)")
        self.lbl_resumen.configure(text=txt)

    def cargar_manuales(self, banco):
        conn = conectar_db(silencioso=True)
        if not conn:
            return []
        movs = []
        try:
            with conn.cursor() as c:
                c.execute("""
                    SELECT id, fecha, descripcion, monto, tipo, estado
                    FROM conciliacion_bancaria
                    WHERE origen = 'manual' AND banco = %s
                    ORDER BY id
                """, (banco.get("banco", ""),))
                for idp, fecha, desc, monto, tipo, estado in c.fetchall():
                    movs.append({
                        "id": idp, "fecha": fecha or "", "descripcion": desc or "",
                        "monto": float(monto or 0), "tipo": tipo or "ingreso",
                        "tabla": "conciliacion", "estado": estado or "pendiente",
                    })
        except Exception:
            pass
        finally:
            liberar_conexion(conn)
        return movs

    def refrescar_tree(self):
        self.tabla_conc.delete(*self.tabla_conc.get_children())
        for i, f in enumerate(self.filas_conciliacion):
            estado = f.get("estado", "pendiente")
            if estado == "conciliado":
                estado_txt = "✅ Conciliado"
                tag = "conciliado"
            elif f.get("origen") == "estado_cuenta":
                estado_txt = "❌ Sin conciliar (Banco)"
                tag = "diferencia"
            elif estado == "diferencia":
                estado_txt = "⚠️ Sin conciliar (Sistema)"
                tag = "diferencia"
            else:
                estado_txt = "📝 Manual / Pendiente"
                tag = "estado_cuenta"
            if f.get("tipo") == "transferencia":
                tipo_txt = "Transferencia (Ingreso)" if f["monto"] >= 0 else "Transferencia (Egreso)"
            else:
                tipo_txt = "Ingreso" if f["monto"] >= 0 else "Egreso"
            origen_txt = {"sistema": "Sistema", "estado_cuenta": "Estado de Cuenta",
                          "manual": "Manual"}.get(f.get("origen"), f.get("origen", ""))
            self.tabla_conc.insert("", tk.END, iid=str(i),
                                   values=(estado_txt, f["fecha"], f.get("documento", ""),
                                           f["descripcion"], tipo_txt, formatear_monto(f["monto"]),
                                           origen_txt),
                                   tags=(tag,))

    # -----------------------------------------------------
    # ACCIONES DE CONCILIACION
    # -----------------------------------------------------
    def _filas_seleccionadas(self):
        sel = self.tabla_conc.selection()
        return [self.filas_conciliacion[int(i)] for i in sel if i.isdigit()]

    def conciliar_seleccionados(self):
        filas = self._filas_seleccionadas()
        if not filas:
            messagebox.showinfo("Conciliación", "Seleccione uno o más movimientos para conciliar.",
                                parent=self.parent_frame)
            return
        banco = self.banco_seleccionado()
        for f in filas:
            f["estado"] = "conciliado"
            self._persistir_conciliado(f, banco)
        self.refrescar_tree()
        registrar_auditoria(self.usuario_activo, "Banco",
                            f"Concilió {len(filas)} movimiento(s) en {construir_etiqueta_banco(banco) if banco else 'banco'}")
        messagebox.showinfo("Conciliación", f"{len(filas)} movimiento(s) marcado(s) como conciliado(s).",
                            parent=self.parent_frame)

    def _persistir_conciliado(self, f, banco):
        if not banco:
            return
        conn = conectar_db(silencioso=True)
        if not conn:
            return
        try:
            with conn.cursor() as c:
                c.execute("""
                    SELECT COUNT(*) FROM conciliacion_bancaria
                    WHERE origen = %s AND id_movimiento = %s AND banco = %s
                """, (f.get("origen", ""), f.get("id", 0), banco.get("banco", "")))
                if c.fetchone()[0] == 0:
                    c.execute("""
                        INSERT INTO conciliacion_bancaria
                        (banco, cuenta, fecha, descripcion, monto, tipo, origen, id_movimiento, estado, fecha_conciliacion)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, 'conciliado', %s)
                    """, (banco.get("banco", ""), banco.get("cuenta", ""), f.get("fecha", ""),
                          f.get("descripcion", ""), f.get("monto", 0), f.get("tipo", "ingreso"),
                          f.get("origen", ""), f.get("id", 0), datetime.now().strftime("%Y-%m-%d")))
                    conn.commit()
        except Exception:
            pass
        finally:
            liberar_conexion(conn)

    def agregar_movimiento_manual(self):
        banco = self.banco_seleccionado()
        if not banco:
            messagebox.showwarning("Banco", "Seleccione un banco.", parent=self.parent_frame)
            return
        v = ctk.CTkToplevel(self.parent_frame)
        v.title("Agregar Movimiento (Conciliación)")
        v.geometry("460x420")
        v.transient(self.parent_frame)
        v.grab_set()

        ctk.CTkLabel(v, text="➕ Agregar movimiento manual", font=("Arial", 15, "bold"),
                     text_color="#1f538d").pack(pady=(15, 5))
        f = ctk.CTkFrame(v, fg_color="transparent")
        f.pack(fill="x", padx=20)

        ctk.CTkLabel(f, text="Tipo:", font=("Arial", 11, "bold")).pack(anchor="w")
        cmb_tipo = ctk.CTkComboBox(f, values=["Ingreso (Abono)", "Egreso (Cargo)"], width=300, state="readonly")
        cmb_tipo.pack(fill="x", pady=(0, 8)); cmb_tipo.set("Ingreso (Abono)")

        ctk.CTkLabel(f, text="Fecha (DD/MM/AAAA):", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_fecha = ctk.CTkEntry(f)
        ent_fecha.pack(fill="x", pady=(0, 8))
        ent_fecha.insert(0, datetime.now().strftime("%d/%m/%Y"))

        ctk.CTkLabel(f, text="Descripción:", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_desc = ctk.CTkEntry(f)
        ent_desc.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(f, text="Monto:", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_monto = ctk.CTkEntry(f)
        ent_monto.pack(fill="x", pady=(0, 8))

        def guardar():
            try:
                monto = float(ent_monto.get().strip())
            except ValueError:
                messagebox.showerror("Error", "Monto inválido.", parent=v)
                return
            if monto <= 0:
                messagebox.showerror("Error", "El monto debe ser mayor a 0.", parent=v)
                return
            tipo = "ingreso" if cmb_tipo.get().startswith("Ingreso") else "egreso"
            monto = abs(monto) if tipo == "ingreso" else -abs(monto)
            fecha = ent_fecha.get().strip() or datetime.now().strftime("%d/%m/%Y")
            desc = ent_desc.get().strip() or "Ajuste manual de conciliación"

            conn = conectar_db(silencioso=True)
            if conn:
                try:
                    with conn.cursor() as c:
                        c.execute("""
                            INSERT INTO conciliacion_bancaria
                            (banco, cuenta, fecha, descripcion, monto, tipo, origen, id_movimiento, estado)
                            VALUES (%s, %s, %s, %s, %s, %s, 'manual', 0, 'pendiente')
                        """, (banco.get("banco", ""), banco.get("cuenta", ""), fecha, desc, monto, tipo))
                        conn.commit()
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo guardar:\n{e}", parent=v)
                    liberar_conexion(conn)
                    return
                finally:
                    liberar_conexion(conn)
            registrar_auditoria(self.usuario_activo, "Banco",
                                f"Agregó movimiento manual {formatear_monto(monto)} en {construir_etiqueta_banco(banco)}")
            v.destroy()
            self.generar_reporte()

        ctk.CTkButton(v, text="✅ Guardar", width=140, fg_color="#27ae60", command=guardar).pack(pady=10)

    def registrar_pago_sistema(self):
        """Registra un cobro/pago real (factura faltante) vinculado al banco."""
        banco = self.banco_seleccionado()
        if not banco:
            messagebox.showwarning("Banco", "Seleccione un banco.", parent=self.parent_frame)
            return
        v = ctk.CTkToplevel(self.parent_frame)
        v.title("Registrar Cobro/Pago en el Sistema")
        v.geometry("480x480")
        v.transient(self.parent_frame)
        v.grab_set()

        ctk.CTkLabel(v, text="🧾 Registrar movimiento real (factura faltante)", font=("Arial", 14, "bold"),
                     text_color="#1f538d").pack(pady=(15, 5))
        ctk.CTkLabel(v, text=f"Banco: {construir_etiqueta_banco(banco)}", font=("Arial", 11),
                     text_color="gray").pack(pady=(0, 8))

        f = ctk.CTkFrame(v, fg_color="transparent")
        f.pack(fill="x", padx=20)

        ctk.CTkLabel(f, text="Tipo de movimiento:", font=("Arial", 11, "bold")).pack(anchor="w")
        cmb_tipo = ctk.CTkComboBox(f, values=["Ingreso (Cobro a Cliente)", "Egreso (Pago a Proveedor)"],
                                   width=300, state="readonly")
        cmb_tipo.pack(fill="x", pady=(0, 8)); cmb_tipo.set("Ingreso (Cobro a Cliente)")

        ctk.CTkLabel(f, text="Cliente / Proveedor:", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_nombre = ctk.CTkEntry(f)
        ent_nombre.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(f, text="N° Factura / Referencia:", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_ref = ctk.CTkEntry(f)
        ent_ref.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(f, text="ID Factura (opcional, 0 si no aplica):", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_id = ctk.CTkEntry(f)
        ent_id.pack(fill="x", pady=(0, 8)); ent_id.insert(0, "0")

        ctk.CTkLabel(f, text="Monto:", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_monto = ctk.CTkEntry(f)
        ent_monto.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(f, text="Fecha:", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_fecha = ctk.CTkEntry(f)
        ent_fecha.pack(fill="x", pady=(0, 8))

        def actualizar_fecha(_=None):
            es_ingreso = cmb_tipo.get().startswith("Ingreso")
            ent_fecha.delete(0, tk.END)
            ent_fecha.insert(0, datetime.now().strftime("%Y-%m-%d" if es_ingreso else "%d/%m/%Y"))
        cmb_tipo.configure(command=actualizar_fecha)
        actualizar_fecha()

        def guardar():
            try:
                monto = float(ent_monto.get().strip())
            except ValueError:
                messagebox.showerror("Error", "Monto inválido.", parent=v)
                return
            if monto <= 0:
                messagebox.showerror("Error", "El monto debe ser mayor a 0.", parent=v)
                return
            try:
                id_factura = int(ent_id.get().strip() or "0")
            except ValueError:
                id_factura = 0
            es_ingreso = cmb_tipo.get().startswith("Ingreso")
            nombre = ent_nombre.get().strip() or ("Cliente" if es_ingreso else "Proveedor")
            ref = ent_ref.get().strip()
            fecha = ent_fecha.get().strip() or datetime.now().strftime("%d/%m/%Y")
            etiqueta = construir_etiqueta_banco(banco)

            conn = conectar_db(silencioso=True)
            if not conn:
                messagebox.showerror("Error", "Sin conexión a la base de datos.", parent=v)
                return
            try:
                with conn.cursor() as c:
                    if es_ingreso:
                        c.execute("""
                            INSERT INTO pagos_clientes
                            (id_factura, monto_pagado, archivo_ruta, cliente_nombre, fecha_pago, cuenta_destino)
                            VALUES (%s, %s, '', %s, %s, %s)
                        """, (id_factura, monto, nombre, fecha, etiqueta))
                    else:
                        c.execute("""
                            INSERT INTO pagos_comprobantes
                            (id_factura, monto_pagado, archivo_ruta, proveedor_nombre, fecha_pago,
                             categoria_suministro, codigo_cotizacion, cuenta_origen)
                            VALUES (%s, %s, '', %s, %s, 'GENERAL', %s, %s)
                        """, (id_factura, monto, nombre, fecha, ref, etiqueta))
                    conn.commit()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo registrar:\n{e}", parent=v)
                return
            finally:
                liberar_conexion(conn)

            registrar_auditoria(self.usuario_activo, "Banco",
                                f"Registró {'cobro' if es_ingreso else 'pago'} {formatear_monto(monto)} en {etiqueta}")
            v.destroy()
            self.generar_reporte()
            messagebox.showinfo("Éxito", "Movimiento registrado correctamente.", parent=self.parent_frame)

        ctk.CTkButton(v, text="✅ Guardar", width=140, fg_color="#27ae60", command=guardar).pack(pady=10)

    def corregir_movimiento(self):
        filas = self._filas_seleccionadas()
        if len(filas) != 1:
            messagebox.showinfo("Corregir", "Seleccione exactamente un movimiento para corregir.",
                                parent=self.parent_frame)
            return
        f = filas[0]
        v = ctk.CTkToplevel(self.parent_frame)
        v.title("Corregir Movimiento")
        v.geometry("460x400")
        v.transient(self.parent_frame)
        v.grab_set()

        ctk.CTkLabel(v, text=f"✏️ Corregir: {f['descripcion'][:40]}", font=("Arial", 13, "bold"),
                     text_color="#1f538d").pack(pady=(15, 5))
        frm = ctk.CTkFrame(v, fg_color="transparent")
        frm.pack(fill="x", padx=20)

        ctk.CTkLabel(frm, text="Fecha:", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_fecha = ctk.CTkEntry(frm)
        ent_fecha.pack(fill="x", pady=(0, 8)); ent_fecha.insert(0, f["fecha"])

        ctk.CTkLabel(frm, text="Descripción:", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_desc = ctk.CTkEntry(frm)
        ent_desc.pack(fill="x", pady=(0, 8)); ent_desc.insert(0, f["descripcion"])

        ctk.CTkLabel(frm, text="Monto (negativo = egreso):", font=("Arial", 11, "bold")).pack(anchor="w")
        ent_monto = ctk.CTkEntry(frm)
        ent_monto.pack(fill="x", pady=(0, 8)); ent_monto.insert(0, f"{f['monto']:.2f}")

        def guardar():
            try:
                monto = float(ent_monto.get().strip())
            except ValueError:
                messagebox.showerror("Error", "Monto inválido.", parent=v)
                return
            fecha = ent_fecha.get().strip()
            desc = ent_desc.get().strip()
            origen = f.get("origen", "")

            if origen == "sistema":
                tabla = f.get("tabla", "")
                idp = f.get("id", 0)
                conn = conectar_db(silencioso=True)
                if conn and idp:
                    try:
                        with conn.cursor() as c:
                            if tabla == "pagos_clientes":
                                c.execute("UPDATE pagos_clientes SET monto_pagado=%s, fecha_pago=%s WHERE id=%s",
                                          (abs(monto), fecha, idp))
                            elif tabla == "pagos_comprobantes":
                                c.execute("UPDATE pagos_comprobantes SET monto_pagado=%s, fecha_pago=%s WHERE id=%s",
                                          (abs(monto), fecha, idp))
                            conn.commit()
                    except Exception as e:
                        messagebox.showerror("Error", f"No se pudo corregir:\n{e}", parent=v)
                        liberar_conexion(conn)
                        return
                    finally:
                        liberar_conexion(conn)
            elif origen == "manual" and f.get("id"):
                conn = conectar_db(silencioso=True)
                if conn:
                    try:
                        with conn.cursor() as c:
                            c.execute("UPDATE conciliacion_bancaria SET monto=%s, fecha=%s, descripcion=%s WHERE id=%s",
                                      (monto, fecha, desc, f["id"]))
                            conn.commit()
                    except Exception as e:
                        messagebox.showerror("Error", f"No se pudo corregir:\n{e}", parent=v)
                        liberar_conexion(conn)
                        return
                    finally:
                        liberar_conexion(conn)

            f["monto"] = monto
            f["fecha"] = fecha
            f["descripcion"] = desc
            registrar_auditoria(self.usuario_activo, "Banco", f"Corrigió movimiento: {desc[:60]}")
            v.destroy()
            self.refrescar_tree()

        ctk.CTkButton(v, text="✅ Guardar", width=140, fg_color="#27ae60", command=guardar).pack(pady=10)

    def exportar_reporte(self):
        if not self.filas_conciliacion:
            messagebox.showinfo("Exportar", "No hay movimientos para exportar.", parent=self.parent_frame)
            return
        ruta = filedialog.asksaveasfilename(
            title="Guardar Reporte de Conciliación",
            defaultextension=".txt",
            filetypes=[("Archivo de texto", "*.txt"), ("CSV", "*.csv")])
        if not ruta:
            return
        try:
            es_csv = ruta.lower().endswith(".csv")
            sep = ";" if es_csv else "\t"
            with open(ruta, "w", encoding="utf-8-sig") as out:
                cab = ["Estado", "Fecha", "Descripción", "Tipo", "Monto", "Origen"]
                out.write(sep.join(cab) + "\n")
                for f in self.filas_conciliacion:
                    tipo = "Ingreso" if f["monto"] >= 0 else "Egreso"
                    estado = f.get("estado", "pendiente")
                    out.write(sep.join([estado, f["fecha"], f["descripcion"],
                                        tipo, f"{f['monto']:.2f}", f.get("origen", "")]) + "\n")
            messagebox.showinfo("Exportar", f"Reporte guardado en:\n{ruta}", parent=self.parent_frame)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar:\n{e}", parent=self.parent_frame)

    def mostrar_texto_pdf(self):
        if not self.texto_pdf:
            messagebox.showinfo("Texto PDF", "Aún no se ha cargado ningún PDF.", parent=self.parent_frame)
            return
        v = ctk.CTkToplevel(self.parent_frame)
        v.title("Texto Extraído del Estado de Cuenta")
        v.geometry("820x560")
        v.transient(self.parent_frame)
        txt = tk.Text(v, wrap="word", font=("Consolas", 10))
        txt.pack(fill="both", expand=True, padx=10, pady=10)
        txt.insert("1.0", self.texto_pdf)
        txt.configure(state="disabled")
        ctk.CTkButton(v, text="Cerrar", width=120, fg_color="#34495e",
                      command=v.destroy).pack(pady=(0, 10))
