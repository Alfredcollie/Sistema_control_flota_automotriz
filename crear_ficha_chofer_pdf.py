# -*- coding: utf-8 -*-
"""Genera la FICHA DE REGISTRO DE CHOFERES / PERSONAL en PDF (con formulario).

🔑 El listado CAMPOS_FICHA es la ÚNICA fuente de verdad: lo usan
   - este generador, para dibujar la ficha, y
   - 'ficha_chofer_lector.py', para leer los datos cuando el PDF pierde el
     formulario (ficha aplanada/impresa).
Si se agrega un campo aquí, la ficha y el lector quedan sincronizados solos.
"""

import os

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.pdfmetrics import stringWidth

from bancos_peru import BANCOS_PDF

NOMBRE_ARCHIVO = "Ficha_Registro_Chofer.pdf"

# Alto de la hoja carta en puntos (lo usa también el lector de fichas)
ALTO_CARTA = 792.0
ANCHO_CARTA = 612.0

# Valor "vacío" para los desplegables (reportlab exige un valor no vacío)
VALOR_VACIO = " "
TALLAS = ["XS", "S", "M", "L", "XL", "XXL", "XXXL"]

# --- SECCIONES (título, y) ---
SECCIONES = [
    ("1. DATOS PERSONALES DEL CHOFER", 662),
    ("2. DATOS DE LICENCIA (MTC)", 348),
    ("3. DATOS DE CONTRATO", 254),
    ("4. CARNÉ DE SANIDAD", 182),
    ("5. DATOS BANCARIOS (para pagos)", 110),
]

# --- CAMPOS: (nombre, etiqueta, y, x_campo, ancho, opciones, maxlen, tooltip) ---
#    opciones=None -> caja de texto | lista -> desplegable
CAMPOS_FICHA = [
    ("dni", "DNI / C.E. *:", 640, 205, 375, None, 12, "DNI (8 dígitos) o Carné de Extranjería"),
    ("nombres", "Nombres y Apellidos *:", 618, 205, 375, None, 80, "Nombres completos del chofer"),
    ("ruc", "RUC (11 dígitos, si aplica):", 596, 205, 375, None, 11, "RUC del chofer"),
    ("direccion", "Dirección de Residencia:", 574, 205, 375, None, 100, "Domicilio del chofer"),
    ("fecha_nacimiento", "Fecha de Nacimiento (DD/MM/AAAA):", 552, 205, 375, None, 10,
     "Fecha de nacimiento"),
    ("sexo", "Sexo:", 530, 205, 375, ["Masculino", "Femenino", "Otro"], 0, "Sexo"),
    ("numero_hijos", "Número de Hijos:", 508, 205, 375, None, 2, "Cantidad de hijos"),
    ("telefono", "Teléfono / WhatsApp:", 486, 205, 375, None, 20,
     "Teléfono de contacto del chofer"),
    ("telefono_emergencia", "Teléfono de Emergencia:", 464, 205, 375, None, 20,
     "Teléfono del contacto de emergencia (familiar)"),
    ("contacto_emergencia_nombre", "Nombre de Contacto de Emergencia:", 442, 205, 375, None, 80,
     "Nombre y parentesco (Ej: Maria Perez - esposa)"),
    ("talla_polo", "Talla de Polo:", 420, 205, 180, [VALOR_VACIO] + TALLAS, 0,
     "Talla del polo del uniforme"),
    ("talla_casaca", "Talla de Casaca:", 420, 480, 100, [VALOR_VACIO] + TALLAS, 0,
     "Talla de la casaca del uniforme"),
    ("correo", "Correo Electrónico:", 398, 205, 375, None, 80, "Correo del chofer"),
    ("estado_laboral", "Estado Laboral:", 376, 205, 375,
     ["Activo", "Inactivo", "Suspendido"], 0, "Estado laboral"),

    ("licencia", "N° Licencia / Brevete:", 326, 205, 375, None, 30, "Número de brevete"),
    ("categoria_licencia", "Categoría:", 304, 205, 375, None, 20,
     "Categoría de licencia (Ej: A-IIb)"),
    ("venc_licencia", "Venc. Licencia (DD/MM/AAAA):", 282, 205, 375, None, 10,
     "Vencimiento del brevete"),

    ("inicio_contrato", "Inicio de Contrato (DD/MM/AAAA):", 232, 205, 375, None, 10,
     "Fecha de inicio de contrato"),
    ("fin_contrato", "Fin de Contrato (DD/MM/AAAA):", 210, 205, 375, None, 10,
     "Fecha de culminación del contrato"),

    ("carnet_sanidad", "N° Carné de Sanidad:", 160, 205, 145, None, 30,
     "Número del carné de sanidad"),
    ("sanidad_emision", "Emisión (DD/MM/AAAA):", 160, 470, 115, None, 10,
     "Fecha de emisión del carné de sanidad"),
    ("venc_sanidad", "Venc. Carné Sanidad (DD/MM/AAAA):", 138, 205, 375, None, 10,
     "Vencimiento del carné de sanidad (emisión + 6 meses)"),

    # 🏦 NUEVO: datos bancarios para el pago de sus servicios
    ("banco", "Banco:", 88, 205, 375, list(BANCOS_PDF), 0,
     "Banco donde recibe sus pagos"),
    ("cuenta_bancaria", "N° de Cuenta:", 66, 205, 375, None, 40,
     "Número de cuenta bancaria"),
    ("cci", "CCI (Cuenta Interbancaria):", 44, 205, 375, None, 30,
     "Código de Cuenta Interbancario (20 dígitos)"),
]


# Ancho REAL (en puntos) de cada etiqueta, para que el lector no la confunda
# con un dato cuando la ficha viene aplanada.
ETIQUETAS_FICHA = []
for _campo in CAMPOS_FICHA:
    _x_etiq = 40 if _campo[3] <= 205 else _campo[3] - 100
    ETIQUETAS_FICHA.append((_x_etiq, _campo[2],
                            stringWidth(str(_campo[1]), "Helvetica", 9.5)))


def _dibujar_campo(c, form, campo):
    """Dibuja la etiqueta y su caja/desplegable."""
    nombre, etiqueta, y, x_campo, ancho, opciones, maxlen, tooltip = campo
    c.setFont("Helvetica", 9.5)
    x_etiqueta = 40 if x_campo <= 205 else x_campo - 100
    c.drawString(x_etiqueta, y, etiqueta)
    if opciones:
        form.choice(name=nombre, tooltip=tooltip or etiqueta, value=(opciones[0] if opciones else ""),
                    options=list(opciones), x=x_campo, y=y - 5, width=ancho, height=16, fontSize=9)
    else:
        form.textfield(name=nombre, tooltip=tooltip or etiqueta, maxlen=maxlen or 40,
                       x=x_campo, y=y - 5, width=ancho, height=16, fontSize=9)


def generar_ficha_chofer_pdf(ruta_salida=None):
    if not ruta_salida:
        ruta_salida = os.path.join(os.path.dirname(os.path.abspath(__file__)), NOMBRE_ARCHIVO)

    c = canvas.Canvas(ruta_salida, pagesize=letter)
    form = c.acroForm

    # ---------------- encabezado ----------------
    ruta_logo = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")
    if os.path.exists(ruta_logo):
        c.drawImage(ruta_logo, 40, 726, width=50, height=50, mask="auto")

    c.setFont("Helvetica-Bold", 14)
    c.setFillColorRGB(0.12, 0.33, 0.55)
    c.drawString(102, 766, "FICHA DE REGISTRO DE CHOFERES / PERSONAL")
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Oblique", 9)
    c.drawString(102, 752, "Datos personales, emergencia, tallas, licencia, sanidad, "
                           "contrato y datos bancarios.")

    # recuadro de foto
    x0, y0, x1, y1 = 468, 700, 586, 758
    c.setLineWidth(1.1)
    c.rect(x0, y0, x1 - x0, y1 - y0)
    c.setFont("Helvetica-Bold", 8.5)
    c.drawCentredString((x0 + x1) / 2.0, y1 - 16, "FOTO CARNET")
    c.setFont("Helvetica", 6.5)
    c.drawCentredString((x0 + x1) / 2.0, y1 - 28, "Pegar aquí la foto")
    c.drawCentredString((x0 + x1) / 2.0, y1 - 36, "(DNI / C.E. o brevete)")

    c.setLineWidth(1)
    c.line(40, 690, 585, 690)

    # ---------------- secciones y campos ----------------
    for titulo, y_titulo in SECCIONES:
        c.setFont("Helvetica-Bold", 11)
        c.setFillColorRGB(0.12, 0.33, 0.55)
        c.drawString(40, y_titulo, titulo)
        c.setFillColorRGB(0, 0, 0)

    for campo in CAMPOS_FICHA:
        _dibujar_campo(c, form, campo)

    # líneas separadoras
    c.setLineWidth(1)
    for y_linea in (364, 270, 198, 126):
        c.line(40, y_linea, 585, y_linea)

    # ---------------- nota final ----------------
    c.setFont("Helvetica-BoldOblique", 8)
    c.drawString(40, 26, "Nota:")
    c.setFont("Helvetica-Oblique", 7.5)
    c.drawString(68, 26, "Ficha sin datos de asignación logística ni seguros. "
                         "Llenar con Adobe Acrobat Reader / Foxit (no imprimir ni aplanar el PDF).")
    c.drawString(40, 16, "Guarde el PDF con los campos rellenados: el sistema los lee "
                         "automáticamente al importarlo.")

    c.save()
    return ruta_salida


if __name__ == "__main__":
    ruta = generar_ficha_chofer_pdf()
    print("Ficha interactiva generada en: %s" % ruta)
