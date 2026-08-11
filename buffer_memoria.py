# -*- coding: utf-8 -*-
import threading
import time
from conexion import conectar_db

class BufferDatos:
    def __init__(self):
        self.categorias_generales = ["Equipos Audiovisuales", "Mobiliario", "Decoración", "Otros"]
        self.proveedores_nombres = []
        self.clientes_nombres = []
        self.equipos_operativos = []
        self.eventos_aprobados = []
        
        self.en_ejecucion = False
        self.intervalo_segundos = 60

    def iniciar_ciclo(self):
        if self.en_ejecucion: return
        self.en_ejecucion = True
        hilo = threading.Thread(target=self._bucle_actualizacion)
        hilo.daemon = True
        hilo.start()

    def detener_ciclo(self):
        self.en_ejecucion = False

    def _bucle_actualizacion(self):
        while self.en_ejecucion:
            self.sincronizar_ahora()
            time.sleep(self.intervalo_segundos)

    def sincronizar_ahora(self):
        conn = conectar_db(silencioso=True)
        if not conn: return

        try:
            cursor = conn.cursor()
            
            # 1. Cargar Categorías
            cursor.execute("SELECT DISTINCT categoria FROM proveedores WHERE categoria IS NOT NULL AND categoria != ''")
            cats_p = [str(r[0]).strip() for r in cursor.fetchall()]
            try:
                cursor.execute("SELECT DISTINCT categoria FROM inventario_equipos WHERE categoria IS NOT NULL AND categoria != ''")
                cats_i = [str(r[0]).strip() for r in cursor.fetchall()]
            except: cats_i = []
            cats_unidas = list(set(cats_p + cats_i))
            cats_unidas.sort()
            if cats_unidas: self.categorias_generales = cats_unidas

            # 2. Cargar Proveedores
            try:
                cursor.execute("SELECT nombre FROM proveedores ORDER BY nombre ASC")
                self.proveedores_nombres = [str(r[0]).strip() for r in cursor.fetchall()]
            except: pass

            # 3. Cargar Clientes
            try:
                cursor.execute("SELECT nombre_empresa FROM clientes ORDER BY nombre_empresa ASC")
                self.clientes_nombres = [str(r[0]).strip() for r in cursor.fetchall()]
            except: pass

            # 4. Cargar Equipos
            try:
                cursor.execute("SELECT id, codigo, nombre FROM inventario_equipos WHERE estado = 'Operativo' ORDER BY nombre")
                self.equipos_operativos = [f"[{r[0]}] {r[1]} - {r[2]}" for r in cursor.fetchall()]
            except: pass

            # 5. Cargar Eventos
            try:
                cursor.execute("SELECT codigo_cotizacion, nombre_evento FROM cotizaciones WHERE status = 'Aprobada'")
                self.eventos_aprobados = [f"{r[0]} | {r[1]}" for r in cursor.fetchall()]
            except: pass

        except Exception as e:
            pass
        finally:
            conn.close()

cache_sistema = BufferDatos()