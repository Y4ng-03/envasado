"""
seguridad.py — Endurecimiento de seguridad homologado con Dashboard 4.0 (Minalba / Polar)

Incluye:
  1. Login obligatorio para TODO (tablero, APIs y stream SSE en tiempo real).
  2. Clave secreta y configuración cargadas desde variables de entorno (.env / Render Env),
     sin contraseñas ni claves quemadas en el código fuente.
  3. Sesión con duración máxima absoluta (SESION_MAX_SEGUNDOS, 2 h por defecto), HttpOnly,
     SameSite=Lax y Secure automático bajo HTTPS / proxy de Render.
  4. Rate limiting por IP y por usuario contra fuerza bruta (login) y abuso de API.
  5. Anti-CSRF estricto por verificación de Origin/Referer + CORS cerrado por defecto.
  6. Roles (admin, gerencia, supervisor, tv solo lectura) validados en servidor.
  7. Saneamiento de entradas JSON (< > neutralizados), validación de fotos data:image/(png|jpeg|webp)
     y validación de fechas lógicas.
  8. Auditoría persistente en BD (tabla auditoria_seguridad) + cabeceras HTTP seguras.
"""

import ipaddress
import os
import re
import secrets
import threading
import time
from collections import deque
from datetime import date, datetime, timedelta
from urllib.parse import urlparse

from flask import (
    current_app,
    jsonify,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from flask.sessions import SecureCookieSessionInterface
from flask_cors import CORS
from werkzeug.middleware.proxy_fix import ProxyFix

from db_compat import (
    BASE_DIR,
    get_db,
    init_auth_tables,
    registrar_auditoria,
    verificar_usuario,
)

# ── Parámetros configurables mediante .env / Render Environment Variables ───────
SESION_MAX_S = max(300, int(os.environ.get("SESION_MAX_SEGUNDOS", 2 * 60 * 60)))
LOGIN_MAX_FALLOS_IP = int(os.environ.get("LOGIN_MAX_FALLOS_IP", 8))
LOGIN_MAX_FALLOS_USUARIO = int(os.environ.get("LOGIN_MAX_FALLOS_USUARIO", 15))
LOGIN_VENTANA_S = int(os.environ.get("LOGIN_VENTANA_SEGUNDOS", 300))
API_MAX_POR_MIN_IP = int(os.environ.get("API_MAX_POR_MIN_IP", 3000))
API_MAX_POR_MIN_USUARIO = int(os.environ.get("API_MAX_POR_MIN_USUARIO", 3000))
ESCRITURAS_MAX_POR_MIN_USUARIO = int(os.environ.get("ESCRITURAS_MAX_POR_MIN_USUARIO", 300))

JSON_MAX_BYTES = 512 * 1024
JSON_MAX_BYTES_GRANDE = 10 * 1024 * 1024
MAX_CAMPO_TEXTO = 4000
MAX_BUCKETS_MEMORIA = 2000
FECHA_MINIMA = date(2024, 1, 1)

RUTAS_JSON_GRANDES = ("/api/importar/preview", "/api/importar/aplicar")
CAMPOS_BASE64_FOTO = {"foto_url"}
CAMPOS_NO_AUDITAR = {"foto_url", "password"}
_PHOTO_DATA_URL_RE = re.compile(r"^data:image/(png|jpeg|jpg|webp);base64,[A-Za-z0-9+/=]+$")

ROLES_ADMIN = {"admin", "gerencia"}
ROLES_SUPERVISOR = {"admin", "gerencia", "supervisor"}

PERMISOS = [
    ("POST", r"^/api/tablero/vaciar$", ROLES_ADMIN),
    ("GET", r"^/api/auditoria$", ROLES_ADMIN),
    ("POST", r"^/api/.+$", ROLES_SUPERVISOR),
    ("PUT", r"^/api/.+$", ROLES_SUPERVISOR),
    ("PATCH", r"^/api/.+$", ROLES_SUPERVISOR),
    ("DELETE", r"^/api/.+$", ROLES_SUPERVISOR),
]
PERMISOS = [(m, re.compile(rx), roles) for m, rx, roles in PERMISOS]

_RUTAS_AUDITABLES = re.compile(
    r"^/api/(asignaciones/|operadores|importar/aplicar|tablero/(limpiar_lineas|reset|vaciar))"
)

_lock = threading.Lock()
_login_fallos_ip = {}
_login_fallos_user = {}
_ventanas = {}


def cargar_secret_key() -> str:
    """Obtiene FLASK_SECRET_KEY desde .env/entorno; nunca usa una clave fija en código."""
    env_key = os.environ.get("FLASK_SECRET_KEY", "").strip()
    if env_key:
        return env_key

    if os.environ.get("RENDER"):
        print("[SEGURIDAD ADVERTENCIA] FLASK_SECRET_KEY no configurada en Render; usando clave efímera de proceso.")
        return secrets.token_hex(32)

    data_dir = os.path.join(BASE_DIR, "data")
    ruta = os.path.join(data_dir, ".flask_secret")
    try:
        if os.path.isfile(ruta):
            val = open(ruta, "r", encoding="utf-8").read().strip()
            if len(val) >= 32:
                return val
        os.makedirs(data_dir, exist_ok=True)
        val = secrets.token_hex(32)
        with open(ruta, "w", encoding="utf-8") as fh:
            fh.write(val)
        return val
    except OSError:
        return secrets.token_hex(32)


def ip_cliente() -> str:
    """Devuelve IP válida del cliente (soporta ProxyFix/Render y Cloudflare sin spoofing)."""
    remota = request.remote_addr or "?"
    if remota in ("127.0.0.1", "::1") and request.headers.get("Cf-Ray"):
        cf = (request.headers.get("Cf-Connecting-Ip") or "").strip()
        if cf:
            try:
                return str(ipaddress.ip_address(cf))
            except ValueError:
                pass
    try:
        return str(ipaddress.ip_address(remota))
    except ValueError:
        return remota


def next_url_seguro(destino: str) -> str:
    """Solo permite redirecciones relativas internas (anti open-redirect)."""
    d = (destino or "").strip()
    if d.startswith("/") and not d.startswith("//") and "\\" not in d:
        return d
    return "/"


def validar_fecha_logica(fecha_val):
    if fecha_val is None:
        return True, None
    s = str(fecha_val).strip()
    if not s:
        return True, None
    if len(s) > 10 or not re.match(r"^\d{4}-\d{2}-\d{2}$", s):
        return False, "Formato de fecha inválido. Use YYYY-MM-DD."
    try:
        dt = datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        return False, "Fecha de calendario inexistente."
    fecha_max = date(datetime.now().year + 2, 12, 31)
    if dt < FECHA_MINIMA or dt > fecha_max:
        return False, f"Fecha fuera de rango permitido ({FECHA_MINIMA.isoformat()} a {fecha_max.isoformat()})."
    return True, None


def _exceso(clave, limite, ventana_s=60) -> bool:
    ahora = time.time()
    with _lock:
        dq = _ventanas.get(clave)
        if dq is None:
            if len(_ventanas) > 5000:
                for k in [k for k, v in _ventanas.items() if not v or ahora - v[-1] > ventana_s]:
                    _ventanas.pop(k, None)
            dq = _ventanas[clave] = deque()
        while dq and ahora - dq[0] > ventana_s:
            dq.popleft()
        if len(dq) >= limite:
            return True
        dq.append(ahora)
        return False


def _fallos_recientes(tabla, clave) -> int:
    ahora = time.time()
    with _lock:
        if len(tabla) > MAX_BUCKETS_MEMORIA:
            for k in [k for k, v in tabla.items() if not v or ahora - v[-1] >= LOGIN_VENTANA_S]:
                tabla.pop(k, None)
        v = [t for t in tabla.get(clave, []) if ahora - t < LOGIN_VENTANA_S]
        if v:
            tabla[clave] = v
        else:
            tabla.pop(clave, None)
        return len(v)


def _registrar_fallo(tabla, clave) -> None:
    ahora = time.time()
    with _lock:
        if len(tabla) > MAX_BUCKETS_MEMORIA and clave not in tabla:
            for k in [k for k, v in tabla.items() if not v or ahora - v[-1] >= LOGIN_VENTANA_S]:
                tabla.pop(k, None)
            if len(tabla) > MAX_BUCKETS_MEMORIA:
                tabla.pop(next(iter(tabla)), None)
        tabla.setdefault(clave, []).append(ahora)


def _sanear_y_validar(valor, clave="", profundidad=0):
    if profundidad > 12:
        return "Profundidad JSON excedida"
    if isinstance(valor, str):
        if clave in CAMPOS_BASE64_FOTO:
            s = valor.strip()
            if s:
                if len(s) > 450000:
                    return "La foto excede el límite permitido."
                if not _PHOTO_DATA_URL_RE.match(s):
                    return "Formato de foto inválido. Solo se permiten PNG, JPG o WEBP."
            return None
        if len(valor) > MAX_CAMPO_TEXTO:
            return f"El campo '{clave or 'texto'}' excede el límite de {MAX_CAMPO_TEXTO} caracteres."
        if clave in ("vacaciones_inicio", "vacaciones_fin") and valor.strip():
            ok, msg = validar_fecha_logica(valor)
            if not ok:
                return msg
        return None
    if isinstance(valor, list):
        if len(valor) > 2500:
            return "Lista demasiado grande en la solicitud."
        for i, v in enumerate(valor):
            if isinstance(v, str):
                if len(v) > MAX_CAMPO_TEXTO:
                    return f"Elemento de texto excede {MAX_CAMPO_TEXTO} caracteres."
                valor[i] = v.replace("<", "\u2039").replace(">", "\u203a")
            else:
                err = _sanear_y_validar(v, clave=clave, profundidad=profundidad + 1)
                if err:
                    return err
        return None
    if isinstance(valor, dict):
        if len(valor) > 500:
            return "Demasiados campos en el objeto JSON."
        for k in list(valor.keys()):
            v = valor[k]
            err = _sanear_y_validar(v, clave=str(k), profundidad=profundidad + 1)
            if err:
                return err
            if isinstance(v, str) and str(k) not in CAMPOS_BASE64_FOTO:
                valor[k] = v.replace("<", "\u2039").replace(">", "\u203a")
        return None
    return None


def _sesion_vigente():
    u = session.get("user")
    ts = session.get("login_ts")
    if not u or not isinstance(ts, (int, float)):
        return None
    if time.time() - ts > SESION_MAX_S or ts > time.time() + 60:
        session.clear()
        return None
    return u


class _SesionSegura(SecureCookieSessionInterface):
    def get_cookie_secure(self, app):
        try:
            return request.is_secure or request.headers.get("X-Forwarded-Proto") == "https"
        except RuntimeError:
            return False


def init_seguridad(app) -> None:
    app.secret_key = cargar_secret_key()
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
    app.session_interface = _SesionSegura()
    app.config["SESSION_COOKIE_HTTPONLY"] = True
    app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
    app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(seconds=SESION_MAX_S)
    app.config["SESSION_REFRESH_EACH_REQUEST"] = False
    app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

    raw_origins = os.environ.get("ORIGENES_PERMITIDOS", "") or os.environ.get("DASHBOARD_ORIGENES_PERMITIDOS", "")
    origenes = [o.strip() for o in raw_origins.split(",") if o.strip()]
    CORS(app, origins=origenes, supports_credentials=bool(origenes))

    @app.before_request
    def _control_acceso():
        p = request.path
        metodo = request.method
        ip = ip_cliente()

        # 1) Anti-CSRF por verificación de Origin/Referer en mutaciones
        if metodo in ("POST", "PUT", "PATCH", "DELETE"):
            origin = request.headers.get("Origin")
            referer = request.headers.get("Referer")
            if origin and origin != "null":
                if urlparse(origin).netloc != request.host and origin not in origenes:
                    registrar_auditoria("CSRF_ORIGEN_BLOQUEADO", ip=ip, ruta=p, metodo=metodo, detalle=f"Origin={origin}")
                    return jsonify({"error": "Origen no autorizado"}), 403
            elif referer:
                ref_host = urlparse(referer).netloc
                if ref_host and ref_host != request.host and not any(urlparse(o).netloc == ref_host for o in origenes):
                    registrar_auditoria("CSRF_REFERER_BLOQUEADO", ip=ip, ruta=p, metodo=metodo, detalle=f"Referer={referer}")
                    return jsonify({"error": "Origen de referencia no autorizado"}), 403

        libre = p in ("/login", "/healthz", "/favicon.ico", "/api/auth/me") or p.startswith("/static/")

        # 2) Rate limiting por IP para /api/
        if p.startswith("/api/") and _exceso(("ip", ip), API_MAX_POR_MIN_IP):
            resp = jsonify({"error": "Demasiadas solicitudes. Intenta en un momento."})
            resp.status_code = 429
            resp.headers["Retry-After"] = "30"
            return resp

        # 3) Login obligatorio (sesión máxima absoluta)
        usuario = _sesion_vigente()
        if not usuario:
            if libre:
                return None
            if p.startswith("/api/"):
                return jsonify({"error": "Sesión requerida o vencida", "login": "/login"}), 401
            destino = request.full_path.rstrip("?")
            return redirect(url_for("login", next=destino))

        # 4) Rate limiting por usuario autenticado
        uid = str(usuario.get("id"))
        uname = str(usuario.get("username") or uid)
        rol = str(usuario.get("rol") or "").lower()
        if p.startswith("/api/"):
            if _exceso(("u", uid), API_MAX_POR_MIN_USUARIO):
                return jsonify({"error": "Demasiadas solicitudes del usuario."}), 429
            if metodo != "GET" and _exceso(("w", uid), ESCRITURAS_MAX_POR_MIN_USUARIO):
                return jsonify({"error": "Demasiadas operaciones de escritura. Espera un momento."}), 429

        # 5) Control de roles en servidor
        for m, rx, roles in PERMISOS:
            if m == metodo and rx.match(p) and rol not in roles:
                registrar_auditoria("ACCESO_DENEGADO_403", usuario=uname, rol=rol, ip=ip, ruta=p, metodo=metodo, detalle="Rol insuficiente")
                return jsonify({"error": f"Tu rol ({rol or 'sin rol'}) no tiene permiso para realizar esta acción."}), 403

        # 6) Validación y saneamiento JSON
        data = None
        if metodo in ("POST", "PUT", "PATCH") and not (request.mimetype or "").startswith("multipart/"):
            cl = request.content_length or 0
            limit_bytes = JSON_MAX_BYTES_GRANDE if (p in RUTAS_JSON_GRANDES or p.startswith("/api/operadores")) else JSON_MAX_BYTES
            if cl and cl > limit_bytes:
                return jsonify({"error": "Solicitud demasiado grande."}), 413
            if cl and request.is_json:
                data = request.get_json(force=True, silent=True)
                if data is not None:
                    if not isinstance(data, dict):
                        return jsonify({"error": "El cuerpo debe ser un objeto JSON."}), 400
                    err_val = _sanear_y_validar(data)
                    if err_val:
                        return jsonify({"error": err_val}), 400

        # 7) Auditoría de acciones críticas (jamás registrar contraseñas ni binarios)
        if metodo in ("POST", "PATCH", "DELETE") and _RUTAS_AUDITABLES.match(p):
            resumen = ""
            if isinstance(data, dict):
                resumen = ", ".join(
                    f"{k}={str(v)[:100]}" for k, v in list(data.items())[:10] if k not in CAMPOS_NO_AUDITAR
                )
            registrar_auditoria("ACCION_CRITICA", usuario=uname, rol=rol, ip=ip, ruta=p, metodo=metodo, detalle=resumen)
        return None

    @app.after_request
    def _cabeceras_seguridad(response):
        response.headers.setdefault("X-Content-Type-Options", "nosniff")
        response.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
        response.headers.setdefault("Referrer-Policy", "same-origin")
        if not request.path.startswith("/static/") and request.path != "/api/stream":
            response.headers.setdefault("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
            response.headers.setdefault("Pragma", "no-cache")
        if request.is_secure or request.headers.get("X-Forwarded-Proto") == "https":
            response.headers.setdefault("Strict-Transport-Security", "max-age=31536000")
        return response

    @app.route("/login", methods=["GET", "POST"])
    def login():
        error = None
        next_url = next_url_seguro(request.form.get("next") or request.args.get("next") or "/")
        if request.method == "POST":
            ip = ip_cliente()
            username = request.form.get("username", "").strip()
            password = request.form.get("password", "").strip()
            if not username or not password:
                return render_template("login.html", error="Debes ingresar usuario y contraseña.", next_url=next_url), 400
            if len(username) > 64 or len(password) > 128:
                return render_template("login.html", error="Longitud de credenciales no permitida.", next_url=next_url), 400
            ukey = username.lower()[:64]
            if (
                _fallos_recientes(_login_fallos_ip, ip) >= LOGIN_MAX_FALLOS_IP
                or _fallos_recientes(_login_fallos_user, ukey) >= LOGIN_MAX_FALLOS_USUARIO
            ):
                registrar_auditoria("LOGIN_BLOCKED", usuario=ukey, ip=ip, ruta="/login", metodo="POST", detalle="Exceso de intentos fallidos")
                return render_template("login.html", error="Demasiados intentos fallidos. Espera unos minutos e intenta de nuevo.", next_url=next_url), 429

            user = verificar_usuario(username, password)
            if user:
                _login_fallos_ip.pop(ip, None)
                _login_fallos_user.pop(ukey, None)
                session.clear()
                session.permanent = True
                session["user"] = user
                session["login_ts"] = int(time.time())
                registrar_auditoria("LOGIN_OK", usuario=user.get("username"), rol=user.get("rol"), ip=ip, ruta="/login", metodo="POST", detalle=f"next={next_url}")
                return redirect(next_url)

            _registrar_fallo(_login_fallos_ip, ip)
            _registrar_fallo(_login_fallos_user, ukey)
            registrar_auditoria("LOGIN_FAIL", usuario=ukey, ip=ip, ruta="/login", metodo="POST", detalle="Credenciales inválidas")
            return render_template("login.html", error="Credenciales incorrectas. Verifica usuario y contraseña.", next_url=next_url), 401

        if _sesion_vigente():
            return redirect(next_url)
        sin_usuarios = False
        try:
            with get_db() as conn:
                init_auth_tables(conn)
                sin_usuarios = conn.execute("SELECT COUNT(*) AS c FROM usuarios").fetchone()["c"] == 0
        except Exception:
            pass
        return render_template("login.html", error=error, next_url=next_url, sin_usuarios=sin_usuarios)

    @app.route("/logout", methods=["GET", "POST"])
    def logout():
        u = session.get("user") or {}
        if u:
            registrar_auditoria("LOGOUT", usuario=u.get("username"), rol=u.get("rol"), ip=ip_cliente(), ruta="/logout", metodo=request.method, detalle="Cierre de sesión")
        session.clear()
        resp = redirect(url_for("login"))
        resp.delete_cookie(current_app.config.get("SESSION_COOKIE_NAME", "session"))
        resp.headers["Clear-Site-Data"] = '"cache", "storage"'
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        resp.headers["Pragma"] = "no-cache"
        resp.headers["Expires"] = "0"
        return resp

    @app.route("/api/auth/me", methods=["GET"])
    def api_auth_me():
        u = _sesion_vigente()
        if u:
            restante = max(0, int(SESION_MAX_S - (time.time() - session.get("login_ts", 0))))
            return jsonify({"authenticated": True, "user": u, "expira_en_s": restante})
        return jsonify({"authenticated": False, "user": None})

    @app.route("/api/auditoria", methods=["GET"])
    def api_auditoria():
        with get_db() as conn:
            init_auth_tables(conn)
            rows = conn.execute(
                "SELECT id, timestamp, evento, usuario, rol, ip, ruta, metodo, detalle FROM auditoria_seguridad ORDER BY id DESC LIMIT 120"
            ).fetchall()
        return jsonify({"eventos": [dict(r) for r in rows]})

    @app.route("/healthz", methods=["GET"])
    def healthz():
        return jsonify({"status": "ok"})
