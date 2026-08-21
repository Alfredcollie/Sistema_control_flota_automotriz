# -*- coding: utf-8 -*-
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

def generar_pdf_interactivo_proveedor():
    # Ruta absoluta del directorio de este script: funciona igual en Windows,
    # macOS (incl. M1/M2/M3) y Linux aunque el programa se lance desde otro lugar.
    directorio_salida = os.path.dirname(os.path.abspath(__file__))
    nombre_archivo = os.path.join(directorio_salida, "Ficha_Registro_Proveedor.pdf")
    c = canvas.Canvas(nombre_archivo, pagesize=letter)
    form = c.acroForm
    
    # 🚀 CONTROL DE LOGOTIPO
    ruta_logo = os.path.join(directorio_salida, "logo.png")
    if os.path.exists(ruta_logo):
        c.drawImage(ruta_logo, 40, 735, width=45, height=45, mask='auto')
        
    # --- ENCABEZADO PRINCIPAL ESTILIZADO ---
    c.setFont("Helvetica-Bold", 13)
    c.drawString(95, 755, "FICHA OFICIAL DE REGISTRO E INCORPORACION DE PROVEEDORES")
    c.setFont("Helvetica-Oblique", 8.5)
    c.drawString(95, 742, "Por favor, complete todos los campos interactivos exactamente como se solicita para la integracion automatica.")
    
    c.setLineWidth(1)
    c.line(40, 725, 570, 725)
    
    # === SECCIÓN 1: DATOS PRINCIPALES Y CATEGORÍA ===
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, 705, "1. Informacion de Identificacion y Categoria")
    
    c.setFont("Helvetica", 9.5)
    c.drawString(40, 680, "Numero de RUC (11 digitos):")
    form.textfield(name="ruc", tooltip="RUC de la empresa", x=180, y=675, width=120, height=16, fontSize=9.5)
    
    c.drawString(310, 680, "Nombre / Razon Social:")
    form.textfield(name="razon_social", tooltip="Nombre o Razon Social", x=420, y=675, width=150, height=16, fontSize=9.5)
    
    # Tu menú de categorías intacto y aprobado
    lista_categorias = [
        "Seleccione una opcion",
        "Luces", "Estructuras", "Sonido", "Video", "Generadores",
        "Catering", "Bebidas", "Menaje", "Mobiliario", "Manteleria",
        "Decoracion", "Impresiones", "Merchandising",
        "Personal", "Seguridad", "Movilidad", "Fotografia", "Artistas",
        "Otros"
    ]
    c.drawString(40, 650, "Categoria Suministro:")
    form.choice(name="categoria", tooltip="Seleccione su rubro", value="Seleccione una opcion", options=lista_categorias, x=180, y=645, width=160, height=16, fontSize=9.5)
    
    c.drawString(350, 650, "Especifique si marco 'Otros':")
    form.textfield(name="especifique_otros", tooltip="Escriba rubro si marco Otros", x=480, y=645, width=90, height=16, fontSize=9.5)
    
    c.drawString(40, 620, "Descripcion Proveedor:\n(Max 400 carac.)")
    form.textfield(name="descripcion_proveedor", tooltip="Resumen o descripcion comercial", x=180, y=590, width=390, height=45, fontSize=9)
    
    c.line(40, 575, 570, 575)
    
    # === SECCIÓN 2: CONTACTOS Y ENLACES ===
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, 555, "2. Informacion de Contacto, Redes y Ubicacion")
    
    contact_fields_left = [
        ("Nombre Contacto 1:", "contacto_principal", 530),
        ("Nombre Contacto 2:", "contacto_alternativo", 505),
        ("Correo Electronico:", "correo", 480),
        ("Link Web:", "link_web", 455),
        ("Enlace Catalogo:", "enlace_catalogo", 430)
    ]
    for label, name, y in contact_fields_left:
        c.setFont("Helvetica", 9.5)
        c.drawString(40, y, label)
        form.textfield(name=name, tooltip=label, x=150, y=y-4, width=150, height=16, fontSize=9)

    contact_fields_right = [
        ("WhatsApp Principal:", "whatsapp_principal", 530),
        ("WhatsApp Alternativo:", "whatsapp_alternativo", 505),
        ("Zona / Distrito Especifico:", "zona_distrito", 480)
    ]
    for label, name, y in contact_fields_right:
        c.setFont("Helvetica", 9.5)
        c.drawString(320, y, label)
        form.textfield(name=name, tooltip=label, x=440, y=y-4, width=130, height=16, fontSize=9)
        
    c.line(40, 410, 570, 410)
    
    # === SECCIÓN 3: INFORMACIÓN FINANCIERA Y DETRACCIONES ===
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, 390, "3. Informacion Financiera y Estructura de Detracciones")
    
    lista_bancos_peru = [
        "Seleccione Banco",
        "BCP", "BBVA", "Interbank", "Scotiabank", 
        "Banco de la Nacion", "BanBif", "Banco Pichincha", 
        "MiBanco", "Banco GNB", "Banco Falabella", 
        "Banco Ripley", "Santander", "Otros"
    ]
    
    # --- CUENTA PRINCIPAL ---
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(40, 365, "CUENTA PRINCIPAL")
    c.setFont("Helvetica", 9.5)
    c.drawString(40, 345, "Banco Principal:")
    form.choice(name="banco_1", tooltip="Seleccione Banco Principal", value="Seleccione Banco", options=lista_bancos_peru, x=140, y=340, width=110, height=16, fontSize=9)
    c.drawString(255, 345, "N° Cuenta:")
    form.textfield(name="cuenta_1", tooltip="Numero de cuenta 1", x=315, y=340, width=110, height=16, fontSize=9)
    c.drawString(435, 345, "CCI:")
    form.textfield(name="cci_1", tooltip="CCI cuenta 1", x=465, y=340, width=105, height=16, fontSize=9)
    
    # --- CUENTA SECUNDARIA ---
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(40, 310, "CUENTA SECUNDARIA (OPCIONAL)")
    c.setFont("Helvetica", 9.5)
    c.drawString(40, 290, "Banco Secundario:")
    form.choice(name="banco_2", tooltip="Seleccione Banco Secundario", value="Seleccione Banco", options=lista_bancos_peru, x=140, y=285, width=110, height=16, fontSize=9)
    c.drawString(255, 290, "N° Cuenta:")
    form.textfield(name="cuenta_2", tooltip="Numero de cuenta 2", x=315, y=285, width=110, height=16, fontSize=9)
    c.drawString(435, 290, "CCI:")
    form.textfield(name="cci_2", tooltip="CCI cuenta 2", x=465, y=285, width=105, height=16, fontSize=9)
    
    # --- SISTEMA DE DETRACCIONES ---
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(40, 265, "SISTEMA DE DETRACCIONES (BANCO DE LA NACION)")
    c.setFont("Helvetica", 9.5)
    c.drawString(40, 245, "Cuenta Detraccion N°:")
    form.textfield(name="cuenta_detraccion", tooltip="N° Cuenta de Detracciones", x=150, y=240, width=160, height=16, fontSize=9)
    c.drawString(330, 245, "Porcentaje Detraccion (%):")
    form.textfield(name="porcentaje_detraccion", tooltip="Porcentaje de detraccion aplicable", x=460, y=240, width=110, height=16, fontSize=9)
    
    c.line(40, 215, 570, 215)
    
    # Instrucciones de Envío Finales
    c.setFont("Helvetica-BoldOblique", 9)
    c.drawString(40, 185, "Nota importante de validacion:")
    c.setFont("Helvetica-Oblique", 9)
    c.drawString(40, 180, "Una vez completado, guarde el archivo PDF conservando los campos interactivos rellenados.")
    c.drawString(40, 167, "No escanee ni imprima este documento fisico; el sistema lo leera electronicamente en segundos.")
    
    c.save()
    print("Ficha interactiva oficial 'Ficha_Registro_Proveedor.pdf' generada con éxito con las nuevas etiquetas de contacto.")

if __name__ == "__main__":
    generar_pdf_interactivo_proveedor()
