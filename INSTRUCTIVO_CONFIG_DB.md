# 📦 Empaquetar config_db.json automáticamente (sin subir la contraseña a git)

La app lee las credenciales de Supabase desde un archivo `config_db.json`
ubicado **junto al ejecutable** (dentro del `.app` en `Contents/MacOS`, o
junto al `.exe` en Windows). No conviene subir ese archivo al repositorio
porque contiene la contraseña. La solución: guardar las credenciales como
**secretos de GitHub** y generar el archivo durante la compilación.

---

## Paso 1 — Guardar las credenciales como secretos de GitHub

En GitHub → tu repositorio → **Settings → Secrets and variables → Actions →
New repository secret**. Crea estos 5 secretos:

| Nombre del secreto | Valor |
|---|---|
| `SUPABASE_DB_HOST` | `aws-0-us-east-2.pooler.supabase.com` |
| `SUPABASE_DB_PORT` | `6543` |
| `SUPABASE_DB_NAME` | `postgres` |
| `SUPABASE_DB_USER` | `postgres.lnmuzwlxcxewobdvlgrh` |
| `SUPABASE_DB_PASSWORD` | *(tu contraseña — la misma del `config_db.json` local)* |

> El host / puerto / base / usuario **no** son secretos; la contraseña sí.
> Por eso la contraseña vive solo como secreto de GitHub y no en el código.

---

## Paso 2 — Agregar un paso al flujo (después de compilar, antes de empaquetar)

### En macOS (`.github/workflows/compilar_windows_mac.yml`)

Agrega este paso **después** de `Compilar con PyInstaller (macOS)` y **antes**
de `Comprimir en ZIP (macOS)`:

```yaml
    - name: Inyectar credenciales de Supabase (config_db.json)
      run: |
        cat > "dist/ControlFlota.app/Contents/MacOS/config_db.json" << 'EOF'
        {
          "host": "${{ secrets.SUPABASE_DB_HOST }}",
          "port": "${{ secrets.SUPABASE_DB_PORT }}",
          "dbname": "${{ secrets.SUPABASE_DB_NAME }}",
          "user": "${{ secrets.SUPABASE_DB_USER }}",
          "password": "${{ secrets.SUPABASE_DB_PASSWORD }}"
        }
        EOF
```

> Si usas el otro flujo (`compilar_mac.yml`, app llamada `BlackRiders`),
> cambia la ruta por `dist/BlackRiders.app/Contents/MacOS/config_db.json`.

### En Windows (mismo archivo `compilar_windows_mac.yml`)

Agrega este paso **después** de `Compilar con PyInstaller (Windows)` y **antes**
de `Comprimir en ZIP (Windows)`:

```yaml
    - name: Inyectar credenciales de Supabase (config_db.json)
      shell: pwsh
      run: |
        @"
        {
          "host": "${{ secrets.SUPABASE_DB_HOST }}",
          "port": "${{ secrets.SUPABASE_DB_PORT }}",
          "dbname": "${{ secrets.SUPABASE_DB_NAME }}",
          "user": "${{ secrets.SUPABASE_DB_USER }}",
          "password": "${{ secrets.SUPABASE_DB_PASSWORD }}"
        }
        "@ | Out-File -FilePath "dist/ControlFlota/config_db.json" -Encoding utf8
```

---

## Notas importantes

- **No** hagas commit de `config_db.json`: ya está en `.gitignore` como red de seguridad.
- El archivo se genera en cada compilación a partir de los secretos, así que cada
  `.app`/`.exe` sale con su `config_db.json` adentro, sin intervención manual.
- Si en el futuro cambias la contraseña de Supabase, solo actualiza el secreto
  `SUPABASE_DB_PASSWORD` en GitHub y recompila.
