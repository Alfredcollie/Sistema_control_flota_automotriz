# -*- coding: utf-8 -*-
"""
PROBAR_RELOJ_HIKVISION.PY - Comprueba si el captahuella se puede leer por red
==============================================================================
Los equipos Hikvision de control de acceso (como el DS-K1T321EFWX-B) traen un
servicio web llamado ISAPI con el que se pueden descargar las marcaciones SIN
exportar el Excel a mano.

Esta herramienta SOLO LEE: no borra ni modifica nada en el equipo. Sirve para
averiguar si el equipo es alcanzable y si sus datos se pueden descargar.

Uso:
    python probar_reloj_hikvision.py                 (pregunta los datos por consola)
    python probar_reloj_hikvision.py --buscar        (busca el equipo en la red local)
    python probar_reloj_hikvision.py --ip 192.168.1.64 --dias 3

Datos que pide:
    * IP del equipo (o nombre de red). Se ve en el menú del propio equipo, en
      Configuración de red, o en el portal de Hikvision donde lo registró.
    * Usuario y contraseña del equipo (el mismo del portal web del reloj).
      Por defecto el usuario es "admin".
"""
import argparse
import json
import socket
import ssl
import sys
import urllib.error
import urllib.request
from datetime import datetime, timedelta, timezone

TIEMPO_ESPERA = 6
RUTA_INFO = "/ISAPI/System/deviceInfo"
PUERTO_SDK = 8000          # puerto típico del SDK de Hikvision: delata a sus equipos

# En consolas de Windows con codificación antigua los emojis rompen el programa
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

OK = "[SI]"
MAL = "[NO]"
AVISO = "[!]"


def contexto_ssl():
    """Los equipos usan certificados propios: se acepta sin verificar."""
    try:
        return ssl._create_unverified_context()
    except AttributeError:
        contexto = ssl.create_default_context()
        contexto.check_hostname = False
        contexto.verify_mode = ssl.CERT_NONE
        return contexto


def abrir(base, usuario, clave):
    """Prepara un cliente HTTP con autenticación digest (la que usa Hikvision)."""
    manejador = urllib.request.HTTPDigestAuthHandler()
    manejador.add_password(None, base, usuario, clave)
    return urllib.request.build_opener(manejador, urllib.request.HTTPSHandler(context=contexto_ssl()))


def peticion(cliente, url, datos=None):
    """Hace una petición y devuelve (codigo, texto). No lanza excepciones."""
    peticion_http = urllib.request.Request(url, data=datos)
    if datos is not None:
        peticion_http.add_header("Content-Type", "application/json")
    try:
        with cliente.open(peticion_http, timeout=TIEMPO_ESPERA) as respuesta:
            return respuesta.status, respuesta.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as e:
        detalle = ""
        try:
            detalle = e.read().decode("utf-8", "replace")[:300]
        except Exception:
            pass
        return e.code, detalle
    except urllib.error.URLError as e:
        return 0, str(getattr(e, "reason", e))
    except Exception as e:
        return 0, str(e)


def puerto_abierto(host, puerto=80, espera=0.35):
    try:
        conexion = socket.create_connection((host, puerto), timeout=espera)
        conexion.close()
        return True
    except Exception:
        return False


def direcciones_locales():
    """Todas las direcciones IPv4 de esta computadora (puede haber varias redes)."""
    direcciones = set()
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            direcciones.add(info[4][0])
    except Exception:
        pass
    try:
        import subprocess
        if sys.platform == "win32":
            salida = subprocess.run(["ipconfig"], capture_output=True, text=True, timeout=10).stdout
            for linea in (salida or "").splitlines():
                if "IPv4" in linea and ":" in linea:
                    posible = linea.split(":")[-1].strip()
                    if posible.count(".") == 3:
                        direcciones.add(posible)
        else:
            salida = subprocess.run(["hostname", "-I"], capture_output=True, text=True, timeout=10).stdout
            for posible in (salida or "").split():
                if posible.count(".") == 3:
                    direcciones.add(posible)
    except Exception:
        pass
    return sorted(d for d in direcciones if not d.startswith("127."))


def buscar_en_red(puerto=80):
    """Busca equipos Hikvision con ISAPI en todas las redes locales."""
    import concurrent.futures
    direcciones = direcciones_locales()
    if not direcciones:
        print("No se pudo determinar la red local de esta computadora.")
        return []
    prefijos = []
    for direccion in direcciones:
        prefijo = direccion.rsplit(".", 1)[0]
        if prefijo not in prefijos:
            prefijos.append(prefijo)
    print("Redes locales de esta computadora: %s" % ", ".join("%s.0/24" % p for p in prefijos))
    encontrados = []
    for prefijo in prefijos[:4]:
        print("   Buscando en %s.0/24 (puerto %d)..." % (prefijo, puerto))
        equipos = ["%s.%d" % (prefijo, n) for n in range(1, 255)]
        with concurrent.futures.ThreadPoolExecutor(max_workers=64) as ejecutor:
            resultados = list(ejecutor.map(lambda h: (h, puerto_abierto(h, puerto)), equipos))
        candidatos = [h for h, abierto in resultados if abierto]
        # El puerto 8000 (SDK de Hikvision) es la señal más fiable de que hay un equipo suyo
        with concurrent.futures.ThreadPoolExecutor(max_workers=64) as ejecutor:
            resultados_sdk = list(ejecutor.map(lambda h: (h, puerto_abierto(h, PUERTO_SDK)), equipos))
        con_sdk = [h for h, abierto in resultados_sdk if abierto]
        if con_sdk:
            print("      Equipos Hikvision probables (puerto %d abierto): %s"
                  % (PUERTO_SDK, ", ".join(con_sdk)))
        if candidatos:
            print("      Equipos con el puerto web abierto: %s" % ", ".join(candidatos))
        for host in sorted(set(candidatos) | set(con_sdk)):
            detectado = False
            for base in ("http://%s%s" % (host, "" if puerto == 80 else ":%d" % puerto),
                         "https://%s" % host):
                codigo, texto = peticion(abrir(base, "admin", ""), base + RUTA_INFO)
                if codigo == 401:
                    encontrados.append(base)
                    print("      -> %s pide usuario y contraseña: es un equipo Hikvision (%s)"
                          % (host, base.split(":")[0]))
                    detectado = True
                    break
                if codigo == 200 and "DeviceInfo" in (texto or ""):
                    encontrados.append(base)
                    datos = resumir_equipo(texto)
                    print("      -> %s es un equipo Hikvision: %s %s" % (host, datos.get("Modelo", ""),
                                                                         datos.get("Serie", "")))
                    detectado = True
                    break
            if not detectado and host in con_sdk:
                print("      -> %s tiene el puerto %d abierto (probable Hikvision), pero su web no "
                      "respondió: pruebe a mano con --ip %s" % (host, PUERTO_SDK, host))
    return encontrados


def resumir_equipo(texto):
    """Extrae los datos útiles del XML de deviceInfo."""
    import re
    datos = {}
    for etiqueta, nombre in (("deviceName", "Nombre"), ("model", "Modelo"), ("serialNumber", "Serie"),
                             ("firmwareVersion", "Firmware"), ("deviceType", "Tipo"),
                             ("macAddress", "MAC")):
        coincidencia = re.search(r"<%s>(.*?)</%s>" % (etiqueta, etiqueta), texto or "")
        if coincidencia:
            datos[nombre] = coincidencia.group(1).strip()
    return datos


def limites_fecha(dias):
    """Rango de búsqueda con la zona horaria local (el equipo la exige)."""
    ahora = datetime.now().astimezone()
    desde = (ahora - timedelta(days=dias)).replace(hour=0, minute=0, second=0, microsecond=0)
    desplazamiento = ahora.utcoffset() or timedelta(0)
    signo = "+" if desplazamiento >= timedelta(0) else "-"
    horas = int(abs(desplazamiento).total_seconds() // 3600)
    minutos = int((abs(desplazamiento).total_seconds() % 3600) // 60)
    zona = "%s%02d:%02d" % (signo, horas, minutos)
    return desde.strftime("%Y-%m-%dT%H:%M:%S") + zona, ahora.strftime("%Y-%m-%dT%H:%M:%S") + zona


def buscar_marcaciones(base, usuario, clave, dias=1, maximo=20):
    """Pide las marcaciones al equipo con el servicio de eventos de control de acceso."""
    desde, hasta = limites_fecha(dias)
    cuerpo = {
        "AcsEventCond": {
            "searchID": "collie-nomina-001",
            "searchResultPosition": 0,
            "maxResults": min(max(maximo, 1), 1000),
            "major": 0,
            "minor": 0,
            "startTime": desde,
            "endTime": hasta,
        }
    }
    cliente = abrir(base, usuario, clave)
    codigo, texto = peticion(cliente, base + "/ISAPI/AccessControl/AcsEvent?format=json",
                             json.dumps(cuerpo).encode("utf-8"))
    if codigo != 200:
        return codigo, texto, None
    try:
        return codigo, texto, json.loads(texto)
    except Exception:
        return codigo, texto, None


def mostrar_marcaciones(datos, maximo=20):
    lista = ((datos or {}).get("AcsEvent") or {}).get("InfoList") or []
    total = ((datos or {}).get("AcsEvent") or {}).get("totalMatches")
    print("   Marcaciones encontradas: %s (se muestran %d)" % (total if total is not None else len(lista),
                                                              min(len(lista), maximo)))
    for evento in lista[:maximo]:
        hora = str(evento.get("time") or "").replace("T", " ")
        print("     %s | codigo: %-12s | %-28s | tipo: %s/%s | verificacion: %s"
              % (hora[:19], evento.get("employeeNoString") or evento.get("cardNo") or "-",
                 str(evento.get("name") or "-")[:28], evento.get("major"), evento.get("minor"),
                 evento.get("currentVerifyMode") or "-"))


def main():
    analizador = argparse.ArgumentParser(description="Comprueba la conexión con el captahuella Hikvision")
    analizador.add_argument("--ip", help="IP o nombre del equipo")
    analizador.add_argument("--usuario", default="admin", help="Usuario del equipo (por defecto admin)")
    analizador.add_argument("--clave", help="Contraseña del equipo")
    analizador.add_argument("--dias", type=int, default=1, help="Cuántos días atrás descargar (por defecto 1)")
    analizador.add_argument("--buscar", action="store_true", help="Buscar el equipo en la red local")
    analizador.add_argument("--puerto", type=int, default=80, help="Puerto web del equipo (por defecto 80)")
    argumentos = analizador.parse_args()

    print("=" * 74)
    print(" PRUEBA DE CONEXIÓN CON EL RELOJ BIOMÉTRICO (HIKVISION / ISAPI)")
    print("=" * 74)
    print("Esta prueba NO modifica nada en el equipo: solo lee información.\n")

    ip = argumentos.ip
    if argumentos.buscar and not ip:
        encontrados = buscar_en_red(argumentos.puerto)
        if len(encontrados) == 1:
            ip = encontrados[0].replace("http://", "").replace("https://", "")
            print("\nSe usará el equipo encontrado: %s" % ip)
        elif encontrados:
            ip = input("Escriba la IP del equipo que desea usar: ").strip()
        else:
            print("\nNo se encontró ningún equipo Hikvision en la red local.")
            print("Puede indicar la IP a mano con:  --ip 192.168.1.xx")
            return 1
    while not ip:
        ip = input("IP del reloj (por ejemplo 192.168.1.64): ").strip()
    usuario = argumentos.usuario or "admin"
    clave = argumentos.clave
    while not clave:
        import getpass
        clave = getpass.getpass("Contraseña del usuario %s del equipo: " % usuario)

    for esquema in ("http", "https"):
        base = "%s://%s" % (esquema, ip)
        print("\n1) Probando %s ..." % base)
        cliente = abrir(base, usuario, clave)
        codigo, texto = peticion(cliente, base + RUTA_INFO)
        if codigo == 200:
            datos = resumir_equipo(texto)
            print("   %s El equipo respondió correctamente." % OK)
            for etiqueta, valor in datos.items():
                print("      %-10s %s" % (etiqueta + ":", valor))
            print("\n2) Descargando las marcaciones de los últimos %d día(s) ..." % argumentos.dias)
            codigo_marcas, detalle, datos_json = buscar_marcaciones(base, usuario, clave, argumentos.dias, 20)
            if codigo_marcas == 200 and datos_json:
                mostrar_marcaciones(datos_json, 20)
                print("\nRESULTADO: %s SI se puede leer el reloj por red." % OK)
                print("Con esto se puede crear un botón en el módulo para descargar las")
                print("marcaciones directamente, sin exportar el Excel a mano.")
                return 0
            print("   %s El equipo responde, pero la descarga de marcaciones falló (código %s)."
                  % (AVISO, codigo_marcas))
            print("   Detalle: %s" % str(detalle)[:300])
            print("\nRESULTADO: se puede conectar al equipo, pero habría que ajustar el método")
            print("de descarga (algunos modelos usan otro servicio para los eventos).")
            return 2
        if codigo == 401:
            print("   %s El equipo está ahí, pero el usuario o la contraseña no son correctos." % MAL)
            print("      Use el mismo usuario y clave con los que entra al portal web del reloj.")
            return 3
        if codigo == 0:
            print("   %s No se pudo conectar: %s" % (MAL, str(texto)[:160]))
        else:
            print("   %s El equipo respondió con código %s." % (MAL, codigo))

    print("\nRESULTADO: %s El equipo NO es alcanzable desde esta computadora." % MAL)
    print("")
    print("Esto suele significar una de estas cosas:")
    print("  • El reloj está en otra red (solo conectado por la nube de Hikvision).")
    print("  • La computadora no está en la misma red WiFi/cable que el reloj.")
    print("  • El equipo tiene el acceso web desactivado o el puerto cambiado.")
    print("")
    print("Qué hacer:")
    print("  1) Conecte esta computadora a la misma red del reloj (WiFi o cable) y repita la")
    print("     prueba, o use --buscar para localizarlo automáticamente.")
    print("  2) Si el reloj solo está en la nube de Hikvision y no se puede alcanzar por red,")
    print("     quedan dos caminos: seguir cargando el Excel que exporta del portal, o")
    print("     configurar el reloj para que ENVÍE las marcaciones a un servidor (notificación")
    print("     HTTP), lo que permitiría tenerlas en tiempo real.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
