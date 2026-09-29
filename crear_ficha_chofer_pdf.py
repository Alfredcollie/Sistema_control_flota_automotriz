# -*- coding: utf-8 -*-
"""
Generador de la Ficha PDF interactiva de Registro de Choferes / Personal.

Mismo mecanismo que la 'Ficha_Registro_Proveedor.pdf': se genera un PDF con
campos interactivos (acroform) para que sea llenado a mano (o en lote) y luego
el modulo de Choferes lo importe de forma masiva leyendo los campos con pypdf.

Incluye DATOS PERSONALES (con contacto de emergencia y tallas de uniforme),
DATOS DE LICENCIA (MTC), CONTRATO y CARNE DE SANIDAD.
NO incluye la seccion 'Logistica y Seguros' (movil asignado ni seguros salud/vida).
"""
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

NOMBRE_ARCHIVO = "Ficha_Registro_Chofer.pdf"

# Tallas disponibles para el uniforme (mismas opciones que el modulo de Choferes).
# Se usan como DESPLEGABLE en el PDF para evitar errores al llenarlo.
TALLAS_UNIFORME = ["XS", "S", "M", "L", "XL", "XXL", "XXXL"]

# Primera opcion del desplegable: una fila EN BLANCO ("sin elegir"). Se usa el
# valor " " (un espacio) porque el generador de PDF exige un valor no vacio para
# poder dibujar el desplegable; el modulo de Choferes lo interpreta como vacio.
OPCIONES_TALLA = [["", " "]] + list(TALLAS_UNIFORME)
VALOR_TALLA_VACIO = " "


def generar_ficha_chofer_pdf(ruta_salida=None):
    if not ruta_salida:
        ruta_salida = os.path.join(os.path.dirname(os.path.abspath(__file__)), NOMBRE_ARCHIVO)

    c = canvas.Canvas(ruta_salida, pagesize=letter)
    form = c.acroForm

    # --- CONTROL DE LOGOTIPO (opcional) ---
    ruta_logo = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")
    if os.path.exists(ruta_logo):
        c.drawImage(ruta_logo, 40, 726, width=50, height=50, mask='auto')

    # =============================================================
    # ENCABEZADO PRINCIPAL + RECUADRO DE FOTO CARNET (parte superior)
    # =============================================================
    c.setFont("Helvetica-Bold", 14)
    c.setFillColorRGB(0.12, 0.33, 0.55)          # #1f538d corporativo
    c.drawString(102, 766, "FICHA DE REGISTRO DE CHOFERES / PERSONAL")
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Oblique", 9)
    c.drawString(102, 752, "Datos personales, contacto de emergencia, tallas, licencia, sanidad y contrato.")

    # --- RECUADRO FOTO CARNET (dibujo; quien imprime la ficha pega la foto) ---
    x0, y0, x1, y1 = 468, 700, 586, 758
    c.setLineWidth(1.1)
    c.rect(x0, y0, x1 - x0, y1 - y0)
    c.setDash(1, 0)
    c.setFont("Helvetica-Bold", 8.5)
    c.drawCentredString((x0 + x1) / 2.0, y1 - 16, "FOTO CARNET")
    c.setFont("Helvetica", 6.5)
    c.drawCentredString((x0 + x1) / 2.0, y1 - 28, "Pegar aqui la foto")
    c.drawCentredString((x0 + x1) / 2.0, y1 - 36, "(DNI / C.E. o brevete)")

    c.setLineWidth(1)
    c.line(40, 690, 585, 690)

    # =============================================================
    # Utilidad de dibujo de campos
    # =============================================================
    def _campo(y, etiqueta, nombre, tooltip, opciones=None, valor=None, maxlen=40,
               x_etiqueta=40, x_campo=205, ancho=375, tam_etiqueta=9.5):
        c.setFont("Helvetica", tam_etiqueta)
        c.drawString(x_etiqueta, y, etiqueta)
        if opciones is not None:
            # Desplegable. Si no se indica 'valor', queda VACIO (no se inventa una talla).
            form.choice(name=nombre, tooltip=tooltip, value=(valor or ""), options=opciones,
                        x=x_campo, y=y - 5, width=ancho, height=16, fontSize=9)
        else:
            form.textfield(name=nombre, tooltip=tooltip, maxlen=maxlen,
                           x=x_campo, y=y - 5, width=ancho, height=16, fontSize=9)

    # =============================================================
    # SECCION 1: DATOS PERSONALES  (sin Logistica y Seguros)
    # =============================================================
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, 662, "1. DATOS PERSONALES DEL CHOFER")

    _campo(636, "DNI / C.E. *:", "dni", "DNI (8 digitos) o Carnet de Extranjeria", maxlen=12)
    _campo(612, "Nombres y Apellidos *:", "nombres", "Nombres completos del chofer", maxlen=80)
    _campo(588, "RUC (11 digitos, si aplica):", "ruc", "RUC del chofer", maxlen=11)
    _campo(564, "Direccion de Residencia:", "direccion", "Domicilio del chofer", maxlen=100)
    _campo(540, "Fecha de Nacimiento (DD/MM/AAAA):", "fecha_nacimiento", "Fecha de nacimiento", maxlen=10)
    _campo(516, "Sexo:", "sexo", "Sexo", opciones=["Masculino", "Femenino", "Otro"], valor="Masculino")
    _campo(492, "Numero de Hijos:", "numero_hijos", "Cantidad de hijos", maxlen=2)
    _campo(468, "Telefono / WhatsApp:", "telefono", "Telefono de contacto del chofer", maxlen=20)
    _campo(444, "Telefono de Emergencia:", "telefono_emergencia",
           "Telefono del contacto de emergencia (familiar)", maxlen=20)
    _campo(420, "Nombre de Contacto de Emergencia:", "contacto_emergencia_nombre",
           "Nombre y parentesco del contacto de emergencia (Ej: Maria Perez - esposa)", maxlen=80)
    _campo(396, "Talla de Polo:", "talla_polo",
           "Talla del polo del uniforme (lista desplegable)",
           opciones=list(OPCIONES_TALLA), valor=VALOR_TALLA_VACIO,
           x_campo=205, ancho=180)
    _campo(396, "Talla de Casaca:", "talla_casaca",
           "Talla de la casaca del uniforme (lista desplegable)",
           opciones=list(OPCIONES_TALLA), valor=VALOR_TALLA_VACIO,
           x_etiqueta=395, x_campo=480, ancho=100)
    _campo(372, "Correo Electronico:", "correo", "Correo del chofer", maxlen=80)
    _campo(348, "Estado Laboral:", "estado_laboral", "Estado laboral",
           opciones=["Activo", "Inactivo", "Suspendido"], valor="Activo")

    c.line(40, 336, 585, 336)

    # =============================================================
    # SECCION 2: DATOS DE LICENCIA (MTC)
    # =============================================================
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, 316, "2. DATOS DE LICENCIA (MTC)")

    _campo(290, "N° Licencia / Brevete:", "licencia", "Numero de brevete", maxlen=30)
    _campo(266, "Categoria:", "categoria_licencia", "Categoria de licencia (Ej: A-IIb)", maxlen=20)
    _campo(242, "Venc. Licencia (DD/MM/AAAA):", "venc_licencia", "Vencimiento del brevete", maxlen=10)

    c.line(40, 230, 585, 230)

    # =============================================================
    # SECCION 3: DATOS DE CONTRATO
    # =============================================================
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, 210, "3. DATOS DE CONTRATO")

    _campo(184, "Inicio de Contrato (DD/MM/AAAA):", "inicio_contrato", "Fecha de inicio de contrato", maxlen=10)
    _campo(160, "Fin de Contrato (DD/MM/AAAA):", "fin_contrato", "Fecha de culminacion del contrato", maxlen=10)

    c.line(40, 148, 585, 148)

    # =============================================================
    # SECCION 4: CARNE DE SANIDAD
    # =============================================================
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, 128, "4. CARNE DE SANIDAD")

    _campo(102, "N° Carne de Sanidad:", "carnet_sanidad", "Numero del carne de sanidad", maxlen=30)
    _campo(78, "Venc. Carne Sanidad (DD/MM/AAAA):", "venc_sanidad", "Vencimiento del carne de sanidad", maxlen=10)

    # =============================================================
    # INSTRUCCIONES FINALES
    # =============================================================
    c.setFont("Helvetica-BoldOblique", 9)
    c.drawString(40, 55, "Nota:")
    c.setFont("Helvetica-Oblique", 9)
    c.drawString(70, 55, "Esta ficha NO solicita datos de asignacion logistica (movil) ni de seguros.")
    c.drawString(40, 42, "Guarde el archivo PDF con los campos rellenados; el sistema lo leera electronicamente.")

    c.save()
    return ruta_salida


if __name__ == "__main__":
    ruta = generar_ficha_chofer_pdf()
    print(f"Ficha interactiva oficial '{NOMBRE_ARCHIVO}' generada con exito en: {ruta}")
