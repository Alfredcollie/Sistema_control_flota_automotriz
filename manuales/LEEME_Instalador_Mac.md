# 🍏 Instalador de ControlFlota para macOS

Este instalador se **corre en la Mac** y **descarga automáticamente el archivo desde GitHub**
(repositorio `Alfredcollie/Sistema_control_flota_automotriz`), para luego instalar
la aplicación **ControlFlota** en la carpeta de Aplicaciones.

## Archivos

| Archivo                          | Qué es                                                        |
|----------------------------------|---------------------------------------------------------------|
| `Instalar_ControlFlota_Mac.command`  | El instalador. Es un script de macOS listo para doble clic.   |
| `LEEME_Instalador_Mac.md`            | Este instructivo.                                             |

---

## 🚀 Cómo instalarlo en la Mac

### Opción A — Doble clic (lo más fácil)
1. Descarga el archivo `Instalar_ControlFlota_Mac.command` (desde Releases o desde el repositorio).
2. Si macOS lo bloquea: haz **clic derecho** sobre el archivo → **Abrir** → **Abrir**.
   *(Developer: Collie Software)*
3. Se abre la Terminal y el instalador:
   - Descarga el **.dmg** desde GitHub (la última versión).
   - Monta la imagen.
   - Copia **ControlFlota.app** a **Aplicaciones**.
   - Quita la cuarentena de Gatekeeper.
   - Abre la aplicación.

### Opción B — Desde Terminal
Si el doble clic no funciona (por permisos), abre **Terminal** y ejecuta:

```bash
bash Instalar_ControlFlota_Mac.command
```

---

## ⚙️ Cómo actualizarlo

Vuelve a correr el mismo instalador. Como siempre descarga la **última versión**
desde GitHub (`releases/latest/download`), instala la actualización y reemplaza la app.

---

## 🧩 Detalles técnicos

- **Repositorio:** `https://github.com/Alfredcollie/Sistema_control_flota_automotriz`
- **Archivo que descarga:** `Instalador.ControlFlota.dmg`
- **URL de descarga dinámica:** `https://github.com/<repo>/releases/latest/download/<asset>`
  (baja siempre el último release, sin necesidad de conocer el tag).
- **Destino de instalación:** `/Applications/ControlFlota.app`
  (si `/Applications` no permite escritura, cae a `~/Applications`).
- El script detecta solo el nombre del `.app` dentro de la imagen, así que también
  funciona si el bundle se llama diferente (`BlackRiders`, etc.).

### Para cambiar el archivo a descargar
Abre `Instalar_ControlFlota_Mac.command` y edita estas variables al inicio:

```bash
REPO="Alfredcollie/Sistema_control_flota_automotriz"
ASSET_NAME="Instalador.ControlFlota.dmg"   # o "ControlFlota-Mac.zip"
```

---

## 🛠️ Solución de problemas

| Problema                                          | Solución                                                                                  |
|---------------------------------------------------|-------------------------------------------------------------------------------------------|
| «No se puede abrir porque el desarrollador no es de confianza» | Clic derecho → **Abrir** → **Abrir**.                               |
| «No se pudo descargar»                            | Revisa la conexión a internet y reintenta. La app debe estar publicada en GitHub Releases. |
| Error de arquitectura (Apple Silicon / Intel)     | El binario se compiló para la otra arquitectura. Compilar el DMG para la arquitectura correcta. |
| Se instaló pero no abre                           | Intenta: `xattr -d com.apple.quarantine "/Applications/ControlFlota.app"` |

---

## 📦 Notas para el administrador

- Si el instalador descarga un **.dmg** y dentro está la app, el script la instala directamente.
- Cuando el binario se compile para una sola arquitectura, el script muestra un **AVISO** de
  arquitectura antes de abrir la app, para que sepas si puede no ejecutarse en tu Mac.
- Para publicar este archivo en Releases, el archivo `Instalar_ControlFlota_Mac.command`
  ya está referenciado en el workflow `.github/workflows/compilar_windows_mac.yml`.
