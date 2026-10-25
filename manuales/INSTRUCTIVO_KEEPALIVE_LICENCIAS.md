# Keepalive de la base de datos "Licencias Colliesoftware"

## Qué es
Una tarea automática de GitHub que mantiene **activa** la base de datos de
licencias en Supabase, para que **no se pause** por inactividad (el plan
gratuito pausa el proyecto tras ~7 días sin actividad).

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `keepalive_licencias.py` | Conecta a la base y ejecuta 3 consultas reales (cuentan como actividad) |
| `.github/workflows/keepalive_licencias.yml` | Lo ejecuta automáticamente 2 veces al día |

## Horario
- 07:30 (Perú) = 12:30 UTC
- 19:30 (Perú) = 00:30 UTC

También se puede ejecutar a mano: **GitHub → Actions → "Mantener activo Supabase (Licencias Colliesoftware)" → Run workflow**.

## Credenciales
El script las busca en este orden:

1. **Variables de entorno / Secrets** (recomendado):
   `LICENCIAS_DB_HOST`, `LICENCIAS_DB_PORT`, `LICENCIAS_DB_NAME`,
   `LICENCIAS_DB_USER`, `LICENCIAS_DB_PASSWORD`
   *(también acepta los nombres `SUPABASE_LIC_DB_*`)*
2. **`config_licencia.json`** junto al script (el mismo que usa la validación de licencias)
3. **Llavero del sistema**, servicio `ControlFlotaLicencias`

> Para configurar los secrets: GitHub → repositorio → *Settings → Secrets and variables → Actions → New repository secret*, uno por cada nombre de arriba.

## Cómo probarlo en tu computadora
```bash
python keepalive_licencias.py
```

Salidas:
- `0` = OK (el proyecto respondió y quedó registrada la actividad)
- `2` = no se pudo conectar / el proyecto está pausado
- `3` = faltan credenciales

Se registra todo en `keepalive_licencias_log.txt`.
