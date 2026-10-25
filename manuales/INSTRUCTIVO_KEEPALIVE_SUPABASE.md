# Mantener activo el proyecto gratuito de Supabase (anti-pausa)

## 1. Por que se pausa

Supabase pausa los proyectos del **plan Free** cuando detecta **poca actividad de
base de datos durante 7 dias seguidos**. La documentacion oficial lo dice asi:

> *"A Free plan project is considered inactive if it does not receive sufficient
> user database activity over the past week. Typically a few user requests to the
> database each day over the previous week is enough to keep the project from
> being paused."*

- Antes de pausarlo, Supabase envia **dos correos** al dueno: un aviso ~1 semana
  antes y una confirmacion cuando ya se pauso.
- Un proyecto pausado se puede **restaurar hasta 1 ano** desde el dashboard
  (boton **Resume project**). No se pierden datos.
- Los proyectos con plan **pago (Pro) no se pausan nunca**.

Este proyecto (ControlFlota) usa el proyecto Supabase
`lnmuzwlxcxewobdvlgrh` (pooler `aws-0-us-east-2.pooler.supabase.com`).

## 2. Que se instalo aqui (ya funcionando)

| Archivo | Para que sirve |
|---|---|
| `keepalive_supabase.py` | Se conecta a la base de datos y ejecuta 3 consultas reales. Eso cuenta como actividad de usuario. |
| `instalar_tarea_keepalive.ps1` | Crea / quita la tarea programada de Windows. |
| `Instalar_Tarea_KeepAlive.bat` | Igual que el anterior, pero con doble clic. |
| `.github/workflows/keepalive_supabase.yml` | Version en la nube (GitHub Actions), funciona aunque la PC este apagada. |
| `keepalive_log.txt` | Registro de las ultimas 300 ejecuciones (se crea solo). |

### Tarea programada de Windows (INSTALADA)

- Nombre: **ControlFlota KeepAlive Supabase**
- Horario: **todos los dias a las 08:00 y a las 20:00** (hora local del equipo)
- Si el equipo estaba apagado a esa hora, se ejecuta al encenderlo
  (`StartWhenAvailable`).
- Al terminar: 0 = todo bien.

Comandos utiles (desde esta carpeta):

```powershell
# Ver el estado
Get-ScheduledTaskInfo -TaskName "ControlFlota KeepAlive Supabase"

# Ejecutarla ahora mismo (prueba)
Start-ScheduledTask -TaskName "ControlFlota KeepAlive Supabase"

# Ver el registro
Get-Content .\keepalive_log.txt -Tail 20

# Probar el script a mano
python keepalive_supabase.py

# Eliminarla
powershell -ExecutionPolicy Bypass -File .\instalar_tarea_keepalive.ps1 -Quitar
```

> **Limite:** la tarea solo corre cuando la PC esta encendida. Si se van de
> vacaciones mas de 7 dias con todo apagado, el proyecto igual se pausaria.
> Para eso esta la opcion 3.

## 3. GitHub Actions (funciona con la PC apagada) - RECOMENDADO

El repositorio `Alfredcollie/Sistema_control_flota_automotriz` es **publico**,
asi que los minutos de Actions son **gratis e ilimitados**.

El workflow ya esta escrito en `.github/workflows/keepalive_supabase.yml` y:

- Corre **2 veces al dia** (12:00 UTC = 07:00 Peru, y 00:00 UTC = 19:00 Peru).
- **No necesita secretos nuevos**: reutiliza los mismos `SUPABASE_DB_HOST`,
  `SUPABASE_DB_PORT`, `SUPABASE_DB_NAME`, `SUPABASE_DB_USER` y
  `SUPABASE_DB_PASSWORD` que ya usa `compilar_windows_mac.yml`
  (Settings > Secrets and variables > Actions).
- Se puede lanzar a mano desde **Actions > Mantener activo Supabase > Run workflow**.

**Para activarlo solo falta subir el archivo al repositorio** (rama `main`):

- Opcion A: en GitHub, *Add file > Create new file*, escribir la ruta
  `.github/workflows/keepalive_supabase.yml` y pegar el contenido.
- Opcion B: con Git, desde una copia clonada del repositorio:

```bash
git add .github/workflows/keepalive_supabase.yml
git commit -m "keepalive: ping diario a Supabase para evitar la pausa"
git push origin main
```

> **Ojo:** GitHub desactiva los workflows programados (*schedule*) en
> repositorios que pasan **60 dias sin ninguna actividad** (sin commits). Como
> este repositorio recibe cambios seguido, no deberia pasar; si alguna vez
> ocurre, basta con volver a habilitarlo desde la pestana **Actions**.

## 4. Alternativa sin PC ni repositorio: cron-job.org

Si quieres una tercera red de seguridad que no dependa de tu PC ni de GitHub:

1. Crea una cuenta gratis en https://cron-job.org
2. Crea un cron job con:
   - **URL:** `https://lnmuzwlxcxewobdvlgrh.supabase.co/rest/v1/config_general?select=*&limit=1`
   - **Horario:** cada 12 horas (o diario).
   - **Headers:**
     - `apikey: <TU_ANON_KEY>`
     - `Authorization: Bearer <TU_ANON_KEY>`
3. La **anon key** se copia del dashboard de Supabase: *Project Settings >
   API keys > anon public*.

Con eso la peticion llega a la API REST del proyecto y cuenta como actividad.
(Guarda esa key solo en cron-job.org; es publica, pero igual no conviene
difundirla.)

## 5. Reglas de oro

1. **No borrar** la tarea programada ni el workflow: son el seguro de vida del
   proyecto.
2. Si llega el correo de aviso *"Your project will be paused"*, entra al
   dashboard de Supabase (*Studio*) y pulsa **Resume project**; con eso ya se
   reinicia el contador.
3. Si el proyecto aparece **Paused**, nada funcionara hasta restaurarlo: la app
   daria error de conexion (el script lo detecta y lo avisa en el log).
4. Plan Pro (~25 USD/mes) = proyectos que **nunca** se pausan. Es la unica
   solucion definitiva.
