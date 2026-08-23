# -*- coding: utf-8 -*-
import subprocess
import sys

def instalar_dependencias():
    # Lista de librerías externas que utiliza el sistema.
    # psycopg2-binary incluye binarios nativos para Windows y macOS (Intel y
    # Apple Silicon M1/M2/M3/M4), por lo que no requiere compilador.
    librerias = [
        "customtkinter",        # Interfaz gráfica moderna
        "psycopg2-binary",      # Conexión a PostgreSQL (binarios para Mac/Windows)
        "keyring",              # Llavero del sistema (Windows Credential Manager / macOS Keychain)
        "bcrypt",               # Hash de contraseñas de usuarios
        "reportlab",            # Generación de PDFs
        "pdfplumber",           # Extracción de datos de PDFs
        "pypdf",                # Lectura de fichas PDF interactivas
        "PyMuPDF",              # Lectura de tarjetas de propiedad con IA
        "requests",             # Llamadas a APIs (RUC SUNAT, IA, etc.)
        "pandas",               # Exportación de reportes a Excel
        "openpyxl",             # Motor para que pandas cree .xlsx
        "Pillow",               # Imágenes y logos en la interfaz
        "fastapi",              # API móvil (endpoints)
        "uvicorn",              # Servidor de la API móvil
        "python-multipart",     # Necesario para subir archivos/formularios en FastAPI
        "google-genai",         # IA de Google para la lectura de tickets (api_movil)
        "PyMuPDF",              # Creador de PDF para Mac
    ]

    print("=====================================================")
    print("\U0001F680 INICIANDO INSTALACIÓN DE DEPENDENCIAS - BLACK CUBE")
    print("=====================================================")

    # Actualizar pip primero por seguridad
    try:
        print("Actualizando el instalador 'pip'...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"])
    except Exception as e:
        print(f"Nota: No se pudo actualizar pip, continuando... ({e})")

    # Instalar cada librería
    for lib in librerias:
        print(f"\n\U0001F4E6 Instalando: {lib}...")
        try:
            # sys.executable asegura que usemos el Python correcto, sea en Windows o Mac
            subprocess.check_call([sys.executable, "-m", "pip", "install", lib])
            print(f"\u2705 {lib} instalado correctamente.")
        except subprocess.CalledProcessError:
            # En macOS/Linux con Python del sistema, a veces se requiere --user
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", lib])
                print(f"\u2705 {lib} instalado correctamente (modo --user).")
            except subprocess.CalledProcessError:
                print(f"\u274C Error al instalar {lib}. Por favor revise su conexión o permisos.")

    print("\n=====================================================")
    print("\u2728 INSTALACIÓN FINALIZADA. EL SISTEMA ESTÁ LISTO.")
    print("=====================================================")

if __name__ == "__main__":
    instalar_dependencias()
