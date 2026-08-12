# Presupuestador V4 — Supabase/PostgreSQL + autenticación

## Cambios principales

- Autenticación con `streamlit-authenticator`.
- Cookie de reautenticación guardada en el navegador; duración configurable.
- Roles: `admin` y `usuario`.
- Solo `admin` puede abrir **Base interna**.
- `GEMINI_API_KEY`, `DATABASE_URL`, cookie y usuarios viven en Streamlit Secrets.
- PostgreSQL/Supabase es la base persistente.
- SQLite queda únicamente como modo local cuando NO existe `DATABASE_URL`.
- Si `DATABASE_URL` está configurada pero falla, la app se detiene en vez de guardar accidentalmente en SQLite.
- Compatible con poolers de Supabase: se desactivan prepared statements de Psycopg.
- Conserva modo simulación, Excel/TXT/ZIP y administración de conceptos, precios, proyectos y presupuestos.

## Supabase

1. Cree una cuenta y un proyecto en Supabase.
2. Pulse **Connect**.
3. Copie preferentemente la cadena **Session pooler** (puerto 5432), especialmente para Streamlit Community Cloud.
4. Sustituya `[YOUR-PASSWORD]` por la contraseña de la base.
5. Pegue la cadena completa como `DATABASE_URL` en Streamlit Secrets.
6. La app creará automáticamente sus tablas la primera vez que conecte.

## Streamlit Secrets

Copie el contenido de `secrets_ejemplo.toml` en:

Streamlit Community Cloud > su app > Settings > Secrets

Cambie todos los valores de ejemplo antes de guardar.

## Seguridad

La API key de Gemini ya no se guarda en `localStorage`. Se usa la clave empresarial de Streamlit Secrets.
El navegador solo conserva la cookie de autenticación gestionada por `streamlit-authenticator`.
