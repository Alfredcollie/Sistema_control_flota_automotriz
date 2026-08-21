# -*- coding: utf-8 -*-

import getpass
import keyring

SERVICE_NAME = "ControlEventos"

print("Configuración segura de credenciales")
print("Las credenciales se guardarán en el llavero del sistema.")
print()

host = input("Host de Supabase: ").strip()
port = input("Puerto [5432]: ").strip() or "5432"
dbname = input("Nombre de base de datos [postgres]: ").strip() or "postgres"
user = input("Usuario de base de datos: ").strip()

while True:
    password = getpass.getpass("Contraseña de base de datos: ")
    password2 = getpass.getpass("Repetir contraseña: ")

    if password == password2:
        break

    print("Las contraseñas no coinciden. Intenta nuevamente.")
    print()

keyring.set_password(SERVICE_NAME, "SUPABASE_DB_HOST", host)
keyring.set_password(SERVICE_NAME, "SUPABASE_DB_PORT", port)
keyring.set_password(SERVICE_NAME, "SUPABASE_DB_NAME", dbname)
keyring.set_password(SERVICE_NAME, "SUPABASE_DB_USER", user)
keyring.set_password(SERVICE_NAME, "SUPABASE_DB_PASSWORD", password)

print()
print("Credenciales guardadas correctamente en el llavero del sistema.")