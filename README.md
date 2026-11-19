# Pizarra Digital Táctil · Envasado RP06 (Planta Minalba)

Tablero Kanban en tiempo real con **Flask**, **SSE (`/api/stream`)**, **SQLite / PostgreSQL (`DATABASE_URL`)** y seguridad homologada con Dashboard 4.0.

## Archivos del Proyecto
- `app.py`: servidor Flask, lógica de casillas fijas por línea, importador Excel/CSV y SSE en tiempo real.
- `seguridad.py`: autenticación obligatoria, sesiones seguras, rate limiting, anti-CSRF, CORS restringido y auditoría.
- `db_compat.py`: adaptador dual SQLite (local) / PostgreSQL (Render + Neon/Supabase) y hashes de contraseñas.
- `admin_usuarios.py`: herramienta CLI para crear/eliminar cuentas encriptadas en la base de datos.
- `index.html`: interfaz principal con panel lateral **Personal Disponible** y arrastre por casillas.
- `templates/login.html`: acceso protegido al sistema.
- `static/`: íconos SVG de máquinas, botellas, logo e historial de JS (`gsap.min.js`, `Flip.min.js`).
- `render.yaml` & `requirements.txt`: configuración lista para despliegue en Render.

## Despliegue en Render + PostgreSQL (Neon o Supabase)
1. **Crea tu base de datos PostgreSQL permanente** gratis en [Neon.tech](https://neon.tech) o [Supabase](https://supabase.com) y copia la cadena `DATABASE_URL` (`postgresql://...`).
2. **Sube el proyecto a GitHub** (`.gitignore` excluye automáticamente `.env`, `*.db`, `*.xlsx`, `data/.flask_secret` y `.agents/`).
3. **Crea el servicio en Render** (`New + → Web Service`, conecta tu repo):
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --threads 16 --timeout 0`
   - **Environment Variables**:
     - `FLASK_SECRET_KEY` *(clave aleatoria de 64 caracteres)*
     - `DATABASE_URL` *(tu URL de Neon o Supabase)*
     - `SESION_MAX_SEGUNDOS` *(opcional, ej. `7200` = 2 h o `43200` = 12 h para turno completo)*
4. **Crea tu usuario único encriptado en la BD de Render**: ejecuta en tu PC con la misma `DATABASE_URL` en `.env` (o desde Render Shell):
   ```powershell
   python admin_usuarios.py crear envasado admin "Envasado RP06"
   ```
5. Entra a tu URL `https://tu-app.onrender.com`, inicia sesión y sube tu Excel en **📤 Excel**.
