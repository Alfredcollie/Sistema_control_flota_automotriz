# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
from conexion import conectar_db, liberar_conexion
from buffer_memoria import cache_sistema

class LoginApp:
    # ... tu código anterior de __init__ y creación de interfaz ...

    def ingresar_al_sistema(self):
        usuario = self.ent_usuario.get().strip()
        password = self.ent_password.get().strip()

        if not usuario:
            messagebox.showwarning("Atención", "Ingrese su usuario.")
            return

        #Intentamos conectar a la base de datos (silencioso=True para manejar nosotros la respuesta)
        conn = conectar_db(silencioso=True)

        # ------------------------------------------------------------------
        # 📌 CASO 1: SIN INTERNET (MODO LECTURA OFFLINE)
        # ------------------------------------------------------------------
        if not conn:
            messagebox.showinfo(
                "📡 MODO LECTURA ACTIVADO", 
                f"¡Bienvenido {usuario}!\n\n"
                "No hay conexión a Internet. El sistema iniciará en MODO LECTURA OFFLINE:\n"
                "• Podrás consultar todos tus datos registrados.\n"
                "• No se podrán guardar nuevos cambios hasta reconectarse."
            )
            # Cargamos la copia de respaldo local
            cache_sistema.cargar_copia_local()
            cache_sistema.modo_lectura = True

            # Abrimos la ventana principal pasando el usuario ingresado
            self.abrir_ventana_principal(usuario_activo=usuario)
            return

        # ------------------------------------------------------------------
        # 📌 CASO 2: CON INTERNET (VALIDACIÓN NORMAL EN NUBE)
        # ------------------------------------------------------------------
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT password, rol FROM usuarios WHERE usuario ILIKE %s", (usuario,))
            user_data = cursor.fetchone()

            if user_data and user_data[0] == password: # O tu método de hash/password
                # Sincronizamos RAM e iniciamos el ciclo
                cache_sistema.sincronizar_ahora()
                cache_sistema.iniciar_ciclo()

                self.abrir_ventana_principal(usuario_activo=usuario)
            else:
                messagebox.showerror("Error de Credenciales", "Usuario o contraseña incorrectos.")
        except Exception as e:
            messagebox.showerror("Error", f"Fallo al validar usuario:\n{e}")
        finally:
            liberar_conexion(conn)

    def abrir_ventana_principal(self, usuario_activo):
        self.root.destroy() # Cerramos el login
        
        # Importamos tu ventana/dashboard principal (ej. main.py)
        import main 
        root_main = ctk.CTk()
        app = main.SistemaPrincipalApp(root_main, usuario_activo=usuario_activo)
        root_main.mainloop()