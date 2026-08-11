# -*- coding: utf-8 -*-
import subprocess
import sys

def instalar_dependencias():
    # Lista de librerías externas que utiliza tu sistema
    librerias = [
        "customtkinter",    # Para la interfaz gráfica moderna
        "psycopg2-binary",  # Para la conexión a la base de datos PostgreSQL (usamos binary para evitar errores en Mac/Windows)
        "reportlab",        # Para la generación de PDFs
        "pdfplumber",       # Para la extracción inteligente de datos de los PDFs (OCR)
        "pandas",           # Para la exportación de reportes y cruce de datos
        "openpyxl",         # Motor necesario para que pandas pueda crear archivos Excel (.xlsx)
        "Pillow"            # Para el manejo de imágenes y logos en la interfaz
    ]

    print("=====================================================")
    print("🚀 INICIANDO INSTALACIÓN DE DEPENDENCIAS - BLACK CUBE")
    print("=====================================================")
    
    # Actualizar pip primero por seguridad
    try:
        print("Actualizando el instalador 'pip'...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"])
    except Exception as e:
        print(f"Nota: No se pudo actualizar pip, continuando... ({e})")

    # Instalar cada librería
    for lib in librerias:
        print(f"\n📦 Instalando: {lib}...")
        try:
            # sys.executable asegura que usemos el Python correcto, sea en Windows o Mac
            subprocess.check_call([sys.executable, "-m", "pip", "install", lib])
            print(f"✅ {lib} instalado correctamente.")
        except subprocess.CalledProcessError:
            print(f"❌ Error al instalar {lib}. Por favor revise su conexión o permisos.")
    
    print("\n=====================================================")
    print("✨ INSTALACIÓN FINALIZADA. EL SISTEMA ESTÁ LISTO.")
    print("=====================================================")

if __name__ == "__main__":
    instalar_dependencias()