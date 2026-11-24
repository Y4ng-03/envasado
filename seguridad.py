"""
seguridad.py — Endurecimiento integral OWASP/SAST (Minalba / Polar)

Incluye:
  1. Autenticación obligatoria en tablero, APIs y stream SSE; rutas legales públicas (/terminos, /privacidad, /cookies).
  2. Clave secreta desde entorno (.env / Render) o archivo local protegido; cero secretos en código/navegador.
  3. Sesión con expiración absoluta, timeout de inactividad, rotación periódica de SID,
     cookies HttpOnly/SameSite=Lax/Secure, y /logout con Clear-Site-Data completo ("cookies", "storage", "cache").
  4. Rate limiting anti-spoofing (ProxyFix solo cuando TRUST_PROXY_COUNT > 0).
  5. Anti-CSRF estricto: Fetch Metadata (Sec-Fetch-Site) + Origin/Referer + token CSRF de sesión.
  6. Control de roles en servidor (RBAC) + anti-IDOR en rutas de recursos y trabajos.
  7. Validación estricta de tipos, rangos, fechas/timestamps, subida de archivos por magic bytes y anti-SSRF.
  8. Cabeceras CSP (nonce por petición), HSTS, COOP, CORP, Permissions-Policy, Referrer-Policy, X-Content-Type-Options, X-Frame-Options.
"""

import base64
import binascii
import io
import ipaddress
import os
import re
import secrets
import socket
import threading
import time
import zipfile
from collections import deque
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urlparse

from flask import (
    current_app,
    g,
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
    ensure_auth_tables_ready,
    flush_auditoria,
    get_db,
    registrar_auditoria,
    verificar_usuario,
)
from observabilidad import record_timing

# ── Parámetros configurables mediante .env / Render Environment Variables ───────
SESION_MAX_S = max(300, int(os.environ.get("SESION_MAX_SEGUNDOS", 2 * 60 * 60)))
SESION_INACTIVA_S = max(300, min(SESION_MAX_S, int(os.environ.get("SESION_INACTIVA_SEGUNDOS", 60 * 60))))
SESION_ROTACION_S = max(120, int(os.environ.get("SESION_ROTACION_SEGUNDOS", 15 * 60)))

LOGIN_MAX_FALLOS_IP = int(os.environ.get("LOGIN_MAX_FALLOS_IP", 8))
LOGIN_MAX_FALLOS_USUARIO = int(os.environ.get("LOGIN_MAX_FALLOS_USUARIO", 15))
LOGIN_VENTANA_S = int(os.environ.get("LOGIN_VENTANA_SEGUNDOS", 300))
API_MAX_POR_MIN_IP = int(os.environ.get("API_MAX_POR_MIN_IP", 3000))
API_MAX_POR_MIN_USUARIO = int(os.environ.get("API_MAX_POR_MIN_USUARIO", 3000))
ESCRITURAS_MAX_POR_MIN_USUARIO = int(os.environ.get("ESCRITURAS_MAX_POR_MIN_USUARIO", 300))

TRUST_PROXY_COUNT = int(os.environ.get("TRUST_PROXY_COUNT", "1" if os.environ.get("RENDER") else "0"))
TRUST_CLOUDFLARE = os.environ.get("TRUST_CLOUDFLARE", "").strip().lower() in ("1", "true", "yes")

JSON_MAX_BYTES = 512 * 1024
JSON_MAX_BYTES_GRANDE = 10 * 1024 * 1024
MAX_UPLOAD_EXCEL_BYTES = 10 * 1024 * 1024
MAX_ZIP_UNCOMPRESSED_BYTES = 50 * 1024 * 1024
MAX_ZIP_ENTRIES = 500
MAX_FOTO_DECODED_BYTES = 350 * 1024
MAX_CAMPO_TEXTO = 4000
MAX_BUCKETS_MEMORIA = 2000
FECHA_MINIMA = date(2024, 1, 1)

RUTAS_JSON_GRANDES = ("/api/importar/preview", "/api/importar/aplicar", "/api/cuaderno/registros", "/api/cuaderno/importar")
CAMPOS_BASE64_FOTO = {"foto_url", "foto_etiqueta"}
CAMPOS_NO_AUDITAR = {"foto_url", "foto_etiqueta", "password", "csrf_token"}
_PHOTO_DATA_URL_RE = re.compile(r"^data:image/(png|jpeg|jpg|webp);base64,([A-Za-z0-9+/=]+)$")
_OP_ID_RE = re.compile(r"^[A-Za-z0-9_-]{1,64}$")

ROLES_ADMIN = {"admin", "gerencia"}
ROLES_LINEAS = {"admin", "gerencia", "supervisor"}
ROLES_SUPERVISOR = {"admin", "gerencia", "supervisor", "supervision", "mantenimiento", "coordinador"}
ROLES_CUADERNO_ADMIN = {"admin", "gerencia", "supervisor", "supervision", "coordinador"}
ROLES_CUADERNO_ESCRITURA = {"admin", "gerencia", "supervisor", "supervision", "coordinador", "escaner"}
ROLES_PIZARRA_LECTURA = {"admin", "gerencia", "supervisor", "supervision", "mantenimiento", "coordinador", "tv"}


def allowed_tipos_for_role(rol: str) -> set:
    r = str(rol or "").strip().lower()
    if r in ("admin", "gerencia", "supervisor"):
        return {"operador", "supervision", "mantenimiento"}
    if r == "coordinador":
        return {"supervision", "mantenimiento"}
    if r == "supervision":
        return {"supervision"}
    if r == "mantenimiento":
        return {"mantenimiento"}
    return set()


def role_allows_tipo(rol: str, tipo_personal: str) -> bool:
    return str(tipo_personal or "operador").strip().lower() in allowed_tipos_for_role(rol)


PERMISOS = [
    ("POST", r"^/api/tablero/vaciar$", ROLES_ADMIN),
    ("POST", r"^/api/operadores/deduplicar$", ROLES_ADMIN),
    ("GET", r"^/api/auditoria$", ROLES_ADMIN),
    ("GET", r"^/api/usuarios$", ROLES_ADMIN),
    ("POST", r"^/api/usuarios$", ROLES_ADMIN),
    ("PATCH", r"^/api/usuarios/.+$", ROLES_ADMIN),
    ("DELETE", r"^/api/usuarios/.+$", ROLES_ADMIN),
    ("DELETE", r"^/api/operadores/.+$", ROLES_ADMIN),
    ("GET", r"^/api/(tablero|operadores|historial)(/.*)?$", ROLES_PIZARRA_LECTURA),
    ("GET", r"^/api/cuaderno/exportar$", ROLES_CUADERNO_ADMIN),
    ("POST", r"^/api/cuaderno/importar$", ROLES_CUADERNO_ADMIN),
    ("POST", r"^/api/cuaderno/.+$", ROLES_CUADERNO_ESCRITURA),
    ("PATCH", r"^/api/cuaderno/.+$", ROLES_CUADERNO_ESCRITURA),
    ("DELETE", r"^/api/cuaderno/.+$", ROLES_CUADERNO_ESCRITURA),
    ("POST", r"^/api/importar/.+$", ROLES_LINEAS),
    ("POST", r"^/api/tablero/(limpiar_lineas|reset)$", ROLES_LINEAS),
    ("POST", r"^/api/.+$", ROLES_SUPERVISOR),
    ("PUT", r"^/api/.+$", ROLES_SUPERVISOR),
    ("PATCH", r"^/api/.+$", ROLES_SUPERVISOR),
    ("DELETE", r"^/api/.+$", ROLES_ADMIN),
]
PERMISOS = [(m, re.compile(rx), roles) for m, rx, roles in PERMISOS]

_RUTAS_AUDITABLES = re.compile(
    r"^/api/(asignaciones/|operadores|usuarios|cuaderno/|importar/aplicar|tablero/(limpiar_lineas|reset|vaciar))"
)

_RUTAS_PUBLICAS = {
    "/login",
    "/healthz",
    "/favicon.ico",
    "/api/auth/me",
    "/terminos",
    "/privacidad",
    "/cookies",
}

_lock = threading.Lock()
_login_fallos_ip: Dict[str, List[float]] = {}
_login_fallos_user: Dict[str, List[float]] = {}
_ventanas: Dict[Tuple[str, str], deque] = {}


def reset_security_state_for_tests() -> None:
    """Limpia contadores en memoria (exclusivo para pruebas automatizadas)."""
    with _lock:
        _login_fallos_ip.clear()
        _login_fallos_user.clear()
        _ventanas.clear()


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
            with open(ruta, "r", encoding="utf-8") as fh:
                val = fh.read().strip()
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
    """Devuelve IP válida del cliente sin permitir spoofing de cabeceras cuando no hay proxy confiable."""
    remota = request.remote_addr or "0.0.0.0"
    if TRUST_CLOUDFLARE and TRUST_PROXY_COUNT > 0 and request.headers.get("Cf-Ray"):
        cf = (request.headers.get("Cf-Connecting-Ip") or "").strip()
        if cf:
            try:
                return str(ipaddress.ip_address(cf))
            except ValueError:
                pass
    try:
        return str(ipaddress.ip_address(remota))
    except ValueError:
        return "0.0.0.0"


def next_url_seguro(destino: str) -> str:
    """Solo permite redirecciones relativas internas (anti open-redirect)."""
    d = (destino or "").strip()
    if d.startswith("/") and not d.startswith("//") and "\\" not in d and "@" not in d.split("?")[0]:
        return d
    return "/"


def is_valid_op_id(op_id: str) -> bool:
    return bool(_OP_ID_RE.match(str(op_id or "")))


def is_strict_int(val: Any) -> bool:
    return isinstance(val, int) and not isinstance(val, bool)


def validar_fecha_logica(fecha_val: Any) -> Tuple[bool, Optional[str]]:
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


def validar_imagen_data_url(raw_val: str) -> Tuple[bool, Optional[str]]:
    """Valida URL data:image/(png|jpeg|webp) decodificando base64 y verificando magic bytes reales."""
    s = (raw_val or "").strip()
    if not s:
        return True, None
    if len(s) > 480000:
        return False, "La foto excede el límite permitido."
    m = _PHOTO_DATA_URL_RE.match(s)
    if not m:
        return False, "Formato de foto inválido. Solo se permiten PNG, JPG o WEBP."
    fmt = m.group(1).lower()
    b64_data = m.group(2)
    try:
        raw_bytes = base64.b64decode(b64_data, validate=True)
    except (binascii.Error, ValueError):
        return False, "Datos base64 de imagen inválidos."
    if len(raw_bytes) == 0 or len(raw_bytes) > MAX_FOTO_DECODED_BYTES:
        return False, "La foto decodificada excede el tamaño máximo permitido (350 KB)."

    if fmt == "png":
        if not raw_bytes.startswith(b"\x89PNG\r\n\x1a\n"):
            return False, "Magic bytes inválidos para imagen PNG."
    elif fmt in ("jpeg", "jpg"):
        if not raw_bytes.startswith(b"\xff\xd8\xff"):
            return False, "Magic bytes inválidos para imagen JPEG."
    elif fmt == "webp":
        if len(raw_bytes) < 12 or raw_bytes[:4] != b"RIFF" or raw_bytes[8:12] != b"WEBP":
            return False, "Magic bytes inválidos para imagen WEBP."
    return True, None


_BINARY_FORBIDDEN_MAGICS = (
    b"PK\x03\x04",
    b"\x7fELF",
    b"MZ",
    b"%PDF-",
    b"\x89PNG\r\n\x1a\n",
    b"\xff\xd8\xff",
    b"\x1f\x8b\x08",
)


def validar_archivo_importacion(filename: str, raw_bytes: bytes) -> Tuple[bool, Optional[str]]:
    """Valida extensión, tamaño y magic bytes reales del archivo Excel/CSV antes de procesarlo."""
    fname = os.path.basename(filename or "").strip().lower()
    if not fname:
        return False, "Nombre de archivo inválido."
    if len(raw_bytes) == 0:
        return False, "El archivo enviado está vacío."
    if len(raw_bytes) > MAX_UPLOAD_EXCEL_BYTES:
        return False, "El archivo excede el límite de 10 MB."

    if fname.endswith((".xlsx", ".xlsm")):
        if not raw_bytes.startswith((b"PK\x03\x04", b"PK\x05\x06")):
            return False, "Archivo Excel inválido: la firma binaria (magic bytes) no corresponde a un .xlsx válido."
        try:
            with zipfile.ZipFile(io.BytesIO(raw_bytes), "r") as zf:
                infos = zf.infolist()
                if len(infos) > MAX_ZIP_ENTRIES:
                    return False, "Archivo Excel sospechoso: demasiadas entradas internas."
                total_uncompressed = sum(max(0, info.file_size) for info in infos)
                if total_uncompressed > MAX_ZIP_UNCOMPRESSED_BYTES:
                    return False, "Archivo Excel excede el tamaño descomprimido seguro (50 MB)."
                names = {info.filename for info in infos}
                if "[Content_Types].xml" not in names or not any(n.startswith("xl/") for n in names):
                    return False, "Estructura interna de libro Excel inválida."
        except zipfile.BadZipFile:
            return False, "El archivo .xlsx está corrupto o no es un paquete OOXML válido."
        return True, None

    if fname.endswith(".xls"):
        return False, "Formato .xls antiguo detectado. Guárdalo en Excel como .xlsx o .csv, o copia y pega las celdas en la pestaña 'Pegar desde Excel'."

    if fname.endswith((".csv", ".tsv", ".txt")):
        if b"\x00" in raw_bytes[:8192]:
            return False, "Archivo de texto inválido: contiene bytes nulos binarios."
        head = raw_bytes[:16]
        for sig in _BINARY_FORBIDDEN_MAGICS:
            if head.startswith(sig):
                return False, "El archivo declarado como CSV/texto contiene una cabecera binaria incompatible."
        return True, None

    return False, "Extensión no permitida. Usa .xlsx, .xlsm, .csv, .tsv o .txt."


def es_url_externa_segura(url: str) -> Tuple[bool, Optional[str]]:
    """Valida URLs externas contra SSRF bloqueando esquemas inseguros, metadatos e IPs privadas."""
    raw = (url or "").strip()
    if not raw or len(raw) > 2048:
        return False, "URL vacía o demasiado larga."
    parsed = urlparse(raw)
    if parsed.scheme.lower() not in ("http", "https"):
        return False, "Esquema de URL no permitido (solo http/https)."
    if parsed.username or parsed.password:
        return False, "Credenciales embebidas en URL no permitidas."
    host = (parsed.hostname or "").strip().lower()
    if not host:
        return False, "Host inválido en URL."
    if host in ("localhost", "metadata.google.internal", "metadata", "169.254.169.254"):
        return False, "Destino interno o de metadatos bloqueado (anti-SSRF)."

    def _ip_prohibida(ip_obj: Any) -> bool:
        return bool(
            ip_obj.is_private
            or ip_obj.is_loopback
            or ip_obj.is_link_local
            or ip_obj.is_multicast
            or ip_obj.is_reserved
            or ip_obj.is_unspecified
        )

    try:
        ip_lit = ipaddress.ip_address(host)
        if _ip_prohibida(ip_lit):
            return False, "Dirección IP privada o reservada bloqueada (anti-SSRF)."
        return True, None
    except ValueError:
        pass

    try:
        infos = socket.getaddrinfo(host, parsed.port or (443 if parsed.scheme == "https" else 80), type=socket.SOCK_STREAM)
        for _fam, _type, _proto, _canon, sockaddr in infos:
            ip_resolved = ipaddress.ip_address(sockaddr[0])
            if _ip_prohibida(ip_resolved):
                return False, "El dominio resuelve a una IP interna o reservada (anti-SSRF)."
    except OSError:
        return False, "No se pudo resolver el dominio de destino."
    return True, None


def _exceso(clave: Tuple[str, str], limite: int, ventana_s: int = 60) -> bool:
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


def _fallos_recientes(tabla: Dict[str, List[float]], clave: str) -> int:
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


def _registrar_fallo(tabla: Dict[str, List[float]], clave: str) -> None:
    ahora = time.time()
    with _lock:
        if len(tabla) > MAX_BUCKETS_MEMORIA and clave not in tabla:
            for k in [k for k, v in tabla.items() if not v or ahora - v[-1] >= LOGIN_VENTANA_S]:
                tabla.pop(k, None)
            if len(tabla) > MAX_BUCKETS_MEMORIA:
                tabla.pop(next(iter(tabla)), None)
        tabla.setdefault(clave, []).append(ahora)


def _sanear_y_validar(valor: Any, clave: str = "", profundidad: int = 0) -> Optional[str]:
    if profundidad > 12:
        return "Profundidad JSON excedida"
    if isinstance(valor, str):
        if clave in CAMPOS_BASE64_FOTO:
            ok_img, msg_img = validar_imagen_data_url(valor)
            if not ok_img:
                return msg_img
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


def get_csrf_token() -> str:
    """Devuelve (o inicializa) el token CSRF ligado a la sesión actual."""
    tok = session.get("csrf_token")
    if not tok or not isinstance(tok, str) or len(tok) < 16:
        tok = secrets.token_urlsafe(24)
        session["csrf_token"] = tok
    return tok


def _sesion_vigente() -> Optional[Dict[str, Any]]:
    u = session.get("user")
    ts = session.get("login_ts")
    if not u or not isinstance(ts, (int, float)):
        return None
    now = time.time()
    if now - ts > SESION_MAX_S or ts > now + 60:
        session.clear()
        return None
    last_seen = session.get("last_seen", ts)
    if isinstance(last_seen, (int, float)) and (now - last_seen > SESION_INACTIVA_S):
        session.clear()
        return None
    rotated_at = session.get("rotated_at", ts)
    if not isinstance(rotated_at, (int, float)) or (now - rotated_at >= SESION_ROTACION_S) or not session.get("sid"):
        session["sid"] = secrets.token_urlsafe(16)
        session["rotated_at"] = int(now)
        session["last_seen"] = int(now)
        session.modified = True
    elif not isinstance(last_seen, (int, float)) or (now - last_seen >= 60):
        session["last_seen"] = int(now)
        session.modified = True
    return u


class _SesionSegura(SecureCookieSessionInterface):
    def get_cookie_secure(self, app):
        try:
            return request.is_secure or request.headers.get("X-Forwarded-Proto") == "https"
        except RuntimeError:
            return False


def _is_https_request() -> bool:
    return bool(request.is_secure or request.headers.get("X-Forwarded-Proto") == "https")


def get_csp_nonce() -> str:
    nonce = getattr(g, "csp_nonce", "")
    if not nonce:
        nonce = secrets.token_urlsafe(16)
        g.csp_nonce = nonce
    return nonce


def init_seguridad(app) -> None:
    app.secret_key = cargar_secret_key()
    if TRUST_PROXY_COUNT > 0:
        app.wsgi_app = ProxyFix(
            app.wsgi_app,
            x_for=TRUST_PROXY_COUNT,
            x_proto=TRUST_PROXY_COUNT,
            x_host=TRUST_PROXY_COUNT,
        )
    app.session_interface = _SesionSegura()
    app.config["SESSION_COOKIE_NAME"] = "envasado_session"
    app.config["SESSION_COOKIE_HTTPONLY"] = True
    app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
    app.config["SESSION_COOKIE_PATH"] = "/"
    app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(seconds=SESION_MAX_S)
    app.config["SESSION_REFRESH_EACH_REQUEST"] = False
    app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

    raw_origins = os.environ.get("ORIGENES_PERMITIDOS", "") or os.environ.get("DASHBOARD_ORIGENES_PERMITIDOS", "")
    origenes = [o.strip() for o in raw_origins.split(",") if o.strip() and o.strip() != "*"]
    if origenes:
        CORS(
            app,
            origins=origenes,
            supports_credentials=True,
            methods=["GET", "POST", "PATCH", "DELETE", "OPTIONS"],
            allow_headers=["Content-Type", "X-CSRF-Token"],
        )

    @app.before_request
    def _control_acceso():
        auth_t0 = time.perf_counter()
        try:
            p = request.path
            metodo = request.method
            ip = ip_cliente()

            # 0) Fetch Metadata (Resource Isolation) antes de auth
            sec_site = (request.headers.get("Sec-Fetch-Site") or "").strip().lower()
            sec_mode = (request.headers.get("Sec-Fetch-Mode") or "").strip().lower()
            sec_dest = (request.headers.get("Sec-Fetch-Dest") or "").strip().lower()
            origin = request.headers.get("Origin")
            referer = request.headers.get("Referer")

            if sec_site == "cross-site":
                origin_allowed = bool(origin and origin in origenes)
                is_safe_nav = (
                    metodo in ("GET", "HEAD")
                    and sec_mode == "navigate"
                    and sec_dest not in ("object", "embed")
                    and not p.startswith("/api/")
                )
                if not origin_allowed and not is_safe_nav:
                    registrar_auditoria("CSRF_FETCH_META_BLOQUEADO", ip=ip, ruta=p, metodo=metodo, detalle=f"Sec-Fetch-Site={sec_site}")
                    return jsonify({"error": "Solicitud cross-site bloqueada por política de aislamiento."}), 403

            # 1) Anti-CSRF por verificación de Origin/Referer en mutaciones
            if metodo in ("POST", "PUT", "PATCH", "DELETE"):
                if origin:
                    if origin == "null" or (urlparse(origin).netloc != request.host and origin not in origenes):
                        registrar_auditoria("CSRF_ORIGEN_BLOQUEADO", ip=ip, ruta=p, metodo=metodo, detalle=f"Origin={origin}")
                        return jsonify({"error": "Origen no autorizado"}), 403
                elif referer:
                    ref_host = urlparse(referer).netloc
                    if ref_host and ref_host != request.host and not any(urlparse(o).netloc == ref_host for o in origenes):
                        registrar_auditoria("CSRF_REFERER_BLOQUEADO", ip=ip, ruta=p, metodo=metodo, detalle=f"Referer={referer}")
                        return jsonify({"error": "Origen de referencia no autorizado"}), 403

            libre = p in _RUTAS_PUBLICAS or p.startswith("/static/")

            # 2) Rate limiting por IP para /api/
            if p.startswith("/api/") and _exceso(("ip", ip), API_MAX_POR_MIN_IP):
                resp = jsonify({"error": "Demasiadas solicitudes. Intenta en un momento."})
                resp.status_code = 429
                resp.headers["Retry-After"] = "30"
                return resp

            # 3) Login obligatorio (sesión máxima absoluta + inactividad + rotación)
            usuario = _sesion_vigente()
            if not usuario:
                if libre:
                    return None
                if p.startswith("/api/"):
                    return jsonify({"error": "Sesión requerida o vencida", "login": "/login"}), 401
                destino = request.full_path.rstrip("?")
                return redirect(url_for("login", next=destino))

            # 3b) Verificación estricta de token CSRF en mutaciones autenticadas
            if metodo in ("POST", "PUT", "PATCH", "DELETE") and p != "/login":
                expected_csrf = get_csrf_token()
                supplied_csrf = (
                    request.headers.get("X-CSRF-Token")
                    or request.form.get("csrf_token")
                    or ""
                ).strip()
                if not supplied_csrf or not secrets.compare_digest(supplied_csrf, expected_csrf):
                    registrar_auditoria(
                        "CSRF_TOKEN_INVALIDO",
                        usuario=str(usuario.get("username") or ""),
                        rol=str(usuario.get("rol") or ""),
                        ip=ip,
                        ruta=p,
                        metodo=metodo,
                        detalle="Falta o no coincide X-CSRF-Token",
                    )
                    return jsonify({"error": "Token CSRF inválido o ausente."}), 403

            # 4) Rate limiting por usuario autenticado
            uid = str(usuario.get("id"))
            uname = str(usuario.get("username") or uid)
            rol = str(usuario.get("rol") or "").lower()
            if p.startswith("/api/"):
                if _exceso(("u", uid), API_MAX_POR_MIN_USUARIO):
                    return jsonify({"error": "Demasiadas solicitudes del usuario."}), 429
                if metodo != "GET" and _exceso(("w", uid), ESCRITURAS_MAX_POR_MIN_USUARIO):
                    return jsonify({"error": "Demasiadas operaciones de escritura. Espera un momento."}), 429

            # 5) Control de roles en servidor (RBAC: primera regla específica que coincida)
            for m, rx, roles in PERMISOS:
                if m == metodo and rx.match(p):
                    if rol not in roles:
                        registrar_auditoria("ACCESO_DENEGADO_403", usuario=uname, rol=rol, ip=ip, ruta=p, metodo=metodo, detalle="Rol insuficiente")
                        return jsonify({"error": f"Tu rol ({rol or 'sin rol'}) no tiene permiso para realizar esta acción."}), 403
                    break

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
        finally:
            record_timing("auth", (time.perf_counter() - auth_t0) * 1000.0)

    @app.after_request
    def _cabeceras_seguridad(response):
        response.headers.setdefault("X-Content-Type-Options", "nosniff")
        response.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
        response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
        response.headers.setdefault("Cross-Origin-Opener-Policy", "same-origin")
        response.headers.setdefault("Cross-Origin-Resource-Policy", "same-origin")
        response.headers.setdefault(
            "Permissions-Policy",
            "camera=(), geolocation=(), microphone=(), payment=(), usb=(), interest-cohort=(), focus-without-user-activation=()",
        )
        response.vary.add("Sec-Fetch-Dest")
        response.vary.add("Sec-Fetch-Mode")
        response.vary.add("Sec-Fetch-Site")

        if "Content-Security-Policy" not in response.headers:
            ctype = (response.mimetype or "").lower()
            if ctype == "text/html":
                nonce = get_csp_nonce()
                response.headers["Content-Security-Policy"] = (
                    f"default-src 'self'; "
                    f"script-src 'self' 'nonce-{nonce}'; "
                    f"style-src 'self' 'unsafe-inline'; "
                    f"img-src 'self' data: blob:; "
                    f"font-src 'self'; "
                    f"connect-src 'self'; "
                    f"object-src 'none'; "
                    f"base-uri 'none'; "
                    f"frame-ancestors 'self'; "
                    f"form-action 'self'"
                )
            elif ctype == "application/json":
                response.headers["Content-Security-Policy"] = "default-src 'none'; frame-ancestors 'none'"

        if not request.path.startswith("/static/") and request.path != "/api/stream" and "Cache-Control" not in response.headers:
            response.headers.setdefault("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
            response.headers.setdefault("Pragma", "no-cache")

        if _is_https_request():
            response.headers.setdefault("Strict-Transport-Security", "max-age=31536000; includeSubDomains")
        return response

    @app.route("/login", methods=["GET", "POST"])
    def login():
        ensure_auth_tables_ready()
        error = None
        next_url = next_url_seguro(request.form.get("next") or request.args.get("next") or "/")
        if request.method == "POST":
            ip = ip_cliente()
            expected_login_csrf = session.get("login_csrf_token")
            supplied_login_csrf = (request.form.get("csrf_token") or request.headers.get("X-CSRF-Token") or "").strip()
            if expected_login_csrf:
                if not supplied_login_csrf or not secrets.compare_digest(supplied_login_csrf, str(expected_login_csrf)):
                    registrar_auditoria("LOGIN_CSRF_FAIL", ip=ip, ruta="/login", metodo="POST", detalle="Token CSRF inválido en login")
                    login_csrf = secrets.token_urlsafe(24)
                    session["login_csrf_token"] = login_csrf
                    return (
                        render_template(
                            "login.html",
                            error="Sesión de formulario expirada. Intenta nuevamente.",
                            next_url=next_url,
                            csrf_token=login_csrf,
                            csp_nonce=get_csp_nonce(),
                        ),
                        403,
                    )

            username = request.form.get("username", "").strip()
            password = request.form.get("password", "").strip()
            if not username or not password:
                login_csrf = session.get("login_csrf_token") or secrets.token_urlsafe(24)
                session["login_csrf_token"] = login_csrf
                return (
                    render_template(
                        "login.html",
                        error="Debes ingresar usuario y contraseña.",
                        next_url=next_url,
                        csrf_token=login_csrf,
                        csp_nonce=get_csp_nonce(),
                    ),
                    400,
                )
            if len(username) > 64 or len(password) > 128:
                login_csrf = session.get("login_csrf_token") or secrets.token_urlsafe(24)
                session["login_csrf_token"] = login_csrf
                return (
                    render_template(
                        "login.html",
                        error="Longitud de credenciales no permitida.",
                        next_url=next_url,
                        csrf_token=login_csrf,
                        csp_nonce=get_csp_nonce(),
                    ),
                    400,
                )
            ukey = username.lower()[:64]
            if (
                _fallos_recientes(_login_fallos_ip, ip) >= LOGIN_MAX_FALLOS_IP
                or _fallos_recientes(_login_fallos_user, ukey) >= LOGIN_MAX_FALLOS_USUARIO
            ):
                registrar_auditoria("LOGIN_BLOCKED", usuario=ukey, ip=ip, ruta="/login", metodo="POST", detalle="Exceso de intentos fallidos")
                login_csrf = session.get("login_csrf_token") or secrets.token_urlsafe(24)
                session["login_csrf_token"] = login_csrf
                return (
                    render_template(
                        "login.html",
                        error="Demasiados intentos fallidos. Espera unos minutos e intenta de nuevo.",
                        next_url=next_url,
                        csrf_token=login_csrf,
                        csp_nonce=get_csp_nonce(),
                    ),
                    429,
                )

            user = verificar_usuario(username, password)
            if user:
                with _lock:
                    _login_fallos_ip.pop(ip, None)
                    _login_fallos_user.pop(ukey, None)
                now_ts = int(time.time())
                session.clear()
                session.permanent = True
                session["user"] = user
                session["sid"] = secrets.token_urlsafe(16)
                session["csrf_token"] = secrets.token_urlsafe(24)
                session["login_ts"] = now_ts
                session["last_seen"] = now_ts
                session["rotated_at"] = now_ts
                registrar_auditoria("LOGIN_OK", usuario=user.get("username"), rol=user.get("rol"), ip=ip, ruta="/login", metodo="POST", detalle=f"next={next_url}")
                return redirect(next_url)

            _registrar_fallo(_login_fallos_ip, ip)
            _registrar_fallo(_login_fallos_user, ukey)
            registrar_auditoria("LOGIN_FAIL", usuario=ukey, ip=ip, ruta="/login", metodo="POST", detalle="Credenciales inválidas")
            login_csrf = session.get("login_csrf_token") or secrets.token_urlsafe(24)
            session["login_csrf_token"] = login_csrf
            return (
                render_template(
                    "login.html",
                    error="Credenciales incorrectas. Verifica usuario y contraseña.",
                    next_url=next_url,
                    csrf_token=login_csrf,
                    csp_nonce=get_csp_nonce(),
                ),
                401,
            )

        if _sesion_vigente():
            return redirect(next_url)
        sin_usuarios = False
        try:
            with get_db() as conn:
                sin_usuarios = conn.execute("SELECT COUNT(*) AS c FROM usuarios").fetchone()["c"] == 0
        except Exception:
            pass
        login_csrf = secrets.token_urlsafe(24)
        session["login_csrf_token"] = login_csrf
        return render_template(
            "login.html",
            error=error,
            next_url=next_url,
            sin_usuarios=sin_usuarios,
            csrf_token=login_csrf,
            csp_nonce=get_csp_nonce(),
        )

    @app.route("/logout", methods=["GET", "POST"])
    def logout():
        u = session.get("user") or {}
        if u:
            registrar_auditoria("LOGOUT", usuario=u.get("username"), rol=u.get("rol"), ip=ip_cliente(), ruta="/logout", metodo=request.method, detalle="Cierre de sesión")
        session.clear()
        resp = redirect(url_for("login"))
        resp.delete_cookie(
            current_app.config.get("SESSION_COOKIE_NAME", "envasado_session"),
            path=current_app.config.get("SESSION_COOKIE_PATH", "/"),
            httponly=True,
            samesite="Lax",
            secure=_is_https_request(),
        )
        resp.headers["Clear-Site-Data"] = '"cookies", "storage", "cache"'
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        resp.headers["Pragma"] = "no-cache"
        resp.headers["Expires"] = "0"
        return resp

    @app.route("/api/auth/me", methods=["GET"])
    def api_auth_me():
        u = _sesion_vigente()
        if u:
            restante = max(0, int(SESION_MAX_S - (time.time() - session.get("login_ts", 0))))
            return jsonify({
                "authenticated": True,
                "user": u,
                "expira_en_s": restante,
                "csrf_token": get_csrf_token(),
            })
        return jsonify({"authenticated": False, "user": None})

    @app.route("/api/auditoria", methods=["GET"])
    def api_auditoria():
        flush_auditoria()
        try:
            page = max(1, int(request.args.get("page", 1)))
            per_page = max(1, min(100, int(request.args.get("per_page", 50))))
        except ValueError:
            return jsonify({"error": "Parámetros de paginación inválidos."}), 400
        evento_f = (request.args.get("evento") or "").strip()[:80]
        offset = (page - 1) * per_page

        t0 = time.perf_counter()
        with get_db() as conn:
            if evento_f:
                total = int(
                    conn.execute(
                        "SELECT COUNT(*) AS c FROM auditoria_seguridad WHERE evento = ?",
                        (evento_f,),
                    ).fetchone()["c"]
                )
                rows = conn.execute(
                    "SELECT id, timestamp, evento, usuario, rol, ip, ruta, metodo, detalle "
                    "FROM auditoria_seguridad WHERE evento = ? ORDER BY id DESC LIMIT ? OFFSET ?",
                    (evento_f, per_page, offset),
                ).fetchall()
            else:
                total = int(conn.execute("SELECT COUNT(*) AS c FROM auditoria_seguridad").fetchone()["c"])
                rows = conn.execute(
                    "SELECT id, timestamp, evento, usuario, rol, ip, ruta, metodo, detalle "
                    "FROM auditoria_seguridad ORDER BY id DESC LIMIT ? OFFSET ?",
                    (per_page, offset),
                ).fetchall()
        record_timing("db", (time.perf_counter() - t0) * 1000.0)
        total_pages = max(1, (total + per_page - 1) // per_page)
        items = [dict(r) for r in rows]
        return jsonify({
            "eventos": items,
            "items": items,
            "page": page,
            "per_page": per_page,
            "total": total,
            "total_pages": total_pages,
            "has_next": page < total_pages,
            "has_prev": page > 1,
        })

    @app.route("/healthz", methods=["GET"])
    def healthz():
        return jsonify({"status": "ok"})
