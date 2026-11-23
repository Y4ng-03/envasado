import base64
import csv
import gzip
import hashlib
import io
import json
import os
import queue
import re
import socket
import sys
import threading
import time
import unicodedata
import uuid
import webbrowser
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Optional, Set, Tuple

try:
    import brotli  # type: ignore
except ImportError:  # pragma: no cover
    brotli = None

from flask import Flask, Response, jsonify, render_template, request, send_file, send_from_directory, session
from openpyxl import Workbook, load_workbook

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

from db_compat import (
    BASE_DIR,
    DB_PATH,
    PLANT_TZ_ISO_SUFFIX,
    bg_jobs,
    get_db,
    get_sqlite_path,
    init_auth_tables,
    is_postgres,
    plant_now,
)
from observabilidad import (
    TTLCache,
    compute_etag,
    etag_matches,
    init_observabilidad,
    record_timing,
)
from seguridad import (
    ROLES_ADMIN,
    SESION_MAX_S,
    get_csp_nonce,
    get_csrf_token,
    init_seguridad,
    is_strict_int,
    is_valid_op_id,
    validar_archivo_importacion,
    validar_fecha_logica,
    validar_imagen_data_url,
)

STATIC_DIR = os.path.join(BASE_DIR, "static")

COLUMNAS_TABLERO = ("L1", "L2", "L3", "L4", "RELEVO", "SIAT", "MONTACARGAS")
COLUMNAS_VALIDAS = ("L1", "L2", "L3", "L4", "RELEVO", "SIAT", "MONTACARGAS", "SUPERVISION", "POOL")
LINEAS_CARGO_VALIDAS = ("TODAS", "L1", "L2", "L3", "L4", "SIAT", "MONTACARGAS", "RELEVO")
MAX_POR_COLUMNA: Dict[str, int] = {"SIAT": 2, "MONTACARGAS": 2}
ESTADOS_VALIDOS = ("normal", "relevo", "cambio_turno", "vacaciones", "ausente", "licencia")
ESTADOS_INACTIVOS = ("vacaciones", "ausente", "licencia")

SLOTS_DEFINICION: Dict[str, List[Dict[str, Any]]] = {
    "L1": [
        {"idx": 0, "puesto": "Pela paletas", "shared": "PELA_PALETAS"},
        {"idx": 1, "puesto": "Despaletizador"},
        {"idx": 2, "puesto": "Llenadora"},
        {"idx": 3, "puesto": "Ocme L1"},
        {"idx": 4, "puesto": "Mesas paletizadoras 1"},
        {"idx": 5, "puesto": "Mesas paletizadoras 2"},
    ],
    "L2": [
        {"idx": 0, "puesto": "Pela paletas", "shared": "PELA_PALETAS"},
        {"idx": 1, "puesto": "Despaletizador"},
        {"idx": 2, "puesto": "Posicionadora"},
        {"idx": 3, "puesto": "Llenadora"},
        {"idx": 4, "puesto": "Ocme L2"},
        {"idx": 5, "puesto": "Robot 1"},
        {"idx": 6, "puesto": "Robot 2"},
    ],
    "L3": [
        {"idx": 0, "puesto": "Despaletizador 1"},
        {"idx": 1, "puesto": "Despaletizador 2"},
        {"idx": 2, "puesto": "Llenadora"},
        {"idx": 3, "puesto": "Horno"},
        {"idx": 4, "puesto": "Paletizador 1"},
        {"idx": 5, "puesto": "Paletizador 2"},
        {"idx": 6, "puesto": "Paletizador 3"},
    ],
    "L4": [
        {"idx": 0, "puesto": "Despaletizador 1"},
        {"idx": 1, "puesto": "Despaletizador 2"},
        {"idx": 2, "puesto": "Llenadora"},
        {"idx": 3, "puesto": "SMI"},
        {"idx": 4, "puesto": "Etiquetadora"},
        {"idx": 5, "puesto": "Paletizador 1"},
        {"idx": 6, "puesto": "Paletizador 2"},
        {"idx": 7, "puesto": "Jarabero"},
    ],
    "SIAT": [
        {"idx": 0, "puesto": "Envolvedora Siat 1"},
        {"idx": 1, "puesto": "Envolvedora Siat 2"},
    ],
    "MONTACARGAS": [
        {"idx": 0, "puesto": "Montacarguista 1"},
        {"idx": 1, "puesto": "Montacarguista 2"},
    ],
    "RELEVO": [
        {"idx": 0, "puesto": "Relevo 1"},
        {"idx": 1, "puesto": "Relevo 2"},
        {"idx": 2, "puesto": "Relevo 3"},
        {"idx": 3, "puesto": "Relevo 4"},
        {"idx": 4, "puesto": "Relevo 5"},
        {"idx": 5, "puesto": "Relevo 6"},
        {"idx": 6, "puesto": "Relevo 7"},
        {"idx": 7, "puesto": "Relevo 8"},
    ],
}

# IDs heredados del demo inicial anterior (para migrar a tablero limpio sin datos fijos)
LEGACY_DEMO_IDS = {str(i) for i in range(1, 24)}

app = Flask(__name__, static_folder=STATIC_DIR, static_url_path="/static")
init_observabilidad(app)
init_seguridad(app)

db_lock = threading.RLock()
sse_lock = threading.Lock()
sse_clients: List[queue.Queue] = []
MAX_SSE_CLIENTS = 128
ACTIVE_PORT = 5000

_board_cache = TTLCache(default_ttl_s=15.0, max_entries=64)
_state_build_lock = threading.Lock()
_photo_cache_lock = threading.Lock()
_photo_cache: Dict[str, Tuple[str, str, bytes, str]] = {}
_FOTO_DATA_URL_RX = re.compile(r"^data:(image/(?:png|jpeg|jpg|webp));base64,([A-Za-z0-9+/=]+)$")
_INDEX_HTML_PATH = os.path.join(BASE_DIR, "index.html")
_index_html_cache_mtime: float = -1.0
_index_html_cache_text: str = ""
_index_html_lock = threading.Lock()


def invalidate_board_cache() -> None:
    _board_cache.invalidate()


def _cache_operator_photo(op_id: str, foto_url: str) -> str:
    raw = (foto_url or "").strip()
    if not raw or not raw.startswith("data:image/"):
        if op_id:
            with _photo_cache_lock:
                _photo_cache.pop(op_id, None)
        return ""
    ver_hash = hashlib.sha1(raw.encode("utf-8")).hexdigest()[:12]
    with _photo_cache_lock:
        existing = _photo_cache.get(op_id)
        if existing and existing[0] == ver_hash:
            return ver_hash
    m = _FOTO_DATA_URL_RX.match(raw)
    if not m:
        return ""
    mime = m.group(1).lower()
    if mime == "image/jpg":
        mime = "image/jpeg"
    try:
        img_bytes = base64.b64decode(m.group(2), validate=True)
    except Exception:
        return ""
    if not img_bytes:
        return ""
    etag = f'"foto-{op_id}-{ver_hash}"'
    with _photo_cache_lock:
        if len(_photo_cache) >= 512 and op_id not in _photo_cache:
            _photo_cache.pop(next(iter(_photo_cache)), None)
        _photo_cache[op_id] = (ver_hash, mime, img_bytes, etag)
    return ver_hash


def _slim_op_for_broadcast(op: Dict[str, Any]) -> Dict[str, Any]:
    foto = op.get("foto_url") or ""
    if foto.startswith("data:image/") and len(foto) > 128:
        ver_hash = op.get("foto_v") or _cache_operator_photo(str(op.get("id") or ""), foto)
        if ver_hash:
            slim = dict(op)
            slim["foto_url"] = f"/api/operadores/{op['id']}/foto?v={ver_hash}"
            slim["foto_v"] = ver_hash
            return slim
    return op


def _slim_state_for_broadcast(payload: Dict[str, Any]) -> Dict[str, Any]:
    slim_payload = dict(payload)
    tablero = payload.get("tablero") or {}
    slim_payload["tablero"] = {
        col: [_slim_op_for_broadcast(o) for o in (tablero.get(col) or [])]
        for col in COLUMNAS_TABLERO
    }
    slim_payload["supervision"] = [_slim_op_for_broadcast(o) for o in (payload.get("supervision") or [])]
    slim_payload["pool"] = [_slim_op_for_broadcast(o) for o in (payload.get("pool") or [])]
    return slim_payload


def _prime_board_cache(payload: Dict[str, Any]) -> None:
    invalidate_board_cache()
    ver = _board_cache.version
    payload["state_version"] = ver
    cache_key = f"tablero:{plant_now().strftime('%Y-%m-%d_%H:%M')}"
    raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    etag = compute_etag(raw, prefix=f"b{ver}")
    _board_cache.set_if_version(cache_key, payload, expected_version=ver, etag=etag, ttl_s=15.0)


def publish_state_payload(payload: Dict[str, Any], event_name: str = "tablero_actualizado") -> Dict[str, Any]:
    """Calienta caché en RAM con el nuevo estado y difunde por SSE sin fotos base64 pesadas."""
    _prime_board_cache(payload)
    slim = _slim_state_for_broadcast(payload)
    broadcast_sse(event_name, slim)
    return slim


def wants_slim_response() -> bool:
    return (request.headers.get("X-Board-Slim") or "").strip() == "1"


def _get_index_html_template() -> str:
    global _index_html_cache_mtime, _index_html_cache_text
    mtime = os.path.getmtime(_INDEX_HTML_PATH)
    with _index_html_lock:
        if mtime != _index_html_cache_mtime or not _index_html_cache_text:
            with open(_INDEX_HTML_PATH, "r", encoding="utf-8") as fh:
                _index_html_cache_text = fh.read()
            _index_html_cache_mtime = mtime
        return _index_html_cache_text


def now_iso() -> str:
    return plant_now().strftime("%Y-%m-%d %H:%M:%S")


def today_iso() -> str:
    return plant_now().strftime("%Y-%m-%d")


def parse_ymd(val: Any) -> Optional[date]:
    s = str(val or "").strip()
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        return None


def get_local_ip() -> str:
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.settimeout(0.5)
        sock.connect(("8.8.8.8", 80))
        ip = sock.getsockname()[0]
        sock.close()
        return ip
    except OSError:
        try:
            return socket.gethostbyname(socket.gethostname())
        except OSError:
            return "127.0.0.1"


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s).strip().lower()


def clean_cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return re.sub(r"\s+", " ", str(value)).strip()


SPANISH_ACCENT_MAP = {
    "jose": "José", "maria": "María", "angel": "Ángel", "angelo": "Ángelo",
    "saul": "Saúl", "ali": "Alí", "andres": "Andrés", "jesus": "Jesús",
    "ivan": "Iván", "moises": "Moisés", "ramon": "Ramón", "ruben": "Rubén",
    "victor": "Víctor", "anibal": "Aníbal", "edgar": "Édgar", "dario": "Darío",
    "josue": "Josué", "amilcar": "Amílcar", "erik": "Erick", "geroge": "Georgel",
    "richad": "Richard", "perez": "Pérez", "gonzalez": "González",
    "gonsalez": "González", "gonzales": "González", "rodriguez": "Rodríguez",
    "bolivar": "Bolívar", "lopez": "López", "lpez": "López",
    "martinez": "Martínez", "martines": "Martínez", "sanchez": "Sánchez",
    "hernandez": "Hernández", "melendez": "Meléndez", "benitez": "Benítez",
    "jimenez": "Jiménez", "marquez": "Márquez", "gomez": "Gómez",
    "diaz": "Díaz", "adrian": "Adrián", "uzcategui": "Uzcátegui",
    "marchan": "Marchán", "merchan": "Merchán", "rondon": "Rondón",
    "chacon": "Chacón", "calderon": "Calderón", "alarcon": "Alarcón",
    "alarkon": "Alarcón", "rincon": "Rincón", "moran": "Morán",
    "garcia": "García", "cordoba": "Córdoba", "vilchez": "Vílchez",
    "veliz": "Véliz", "colmenarez": "Colmenárez", "alvarez": "Álvarez",
    "berrios": "Berríos", "serrrano": "Serrano", "titrado": "Tirado",
    "goncavez": "Goncalves", "goncalvez": "Goncalves",
    "infante": "Infante", "blanco": "Blanco", "cabezo": "Cabezo",
    "cruz": "Cruz", "silva": "Silva", "da": "Da",
    "llnadora": "Llenadora", "llenadora": "Llenadora",
    "despaletizado": "Despaletizador", "paletizado": "Paletizador",
    "moviles": "Móviles",
}


def fix_spanish_accents(value: str) -> str:
    clean = clean_cell(value)
    if not clean:
        return ""

    def repl(match: re.Match) -> str:
        word = match.group(0)
        key = normalize_text(word)
        if key in SPANISH_ACCENT_MAP:
            return SPANISH_ACCENT_MAP[key]
        return word

    fixed = re.sub(r"[A-Za-zÁÉÍÓÚáéíóúÜüÑñ-]+", repl, clean)
    fixed = re.sub(r"\bSmi\b", "SMI", fixed)
    return fixed


def normalize_puesto_label(puesto: str) -> str:
    clean = clean_cell(puesto)
    if not clean:
        return ""
    norm = normalize_text(clean)
    if "preparador de paleta" in norm or "pela paleta" in norm:
        return "Pela paletas"
    if norm == "siat" or "envolvedora" in norm:
        return "Envolvedora Siat"
    return clean


_OPERADORES_ALTER_SQL: Dict[str, str] = {
    "puesto": "ALTER TABLE operadores ADD COLUMN puesto TEXT NOT NULL DEFAULT ''",
    "grupo": "ALTER TABLE operadores ADD COLUMN grupo TEXT NOT NULL DEFAULT ''",
    "supervisor_grupo": "ALTER TABLE operadores ADD COLUMN supervisor_grupo TEXT NOT NULL DEFAULT ''",
    "numero": "ALTER TABLE operadores ADD COLUMN numero TEXT NOT NULL DEFAULT ''",
    "vacaciones_inicio": "ALTER TABLE operadores ADD COLUMN vacaciones_inicio TEXT NOT NULL DEFAULT ''",
    "vacaciones_fin": "ALTER TABLE operadores ADD COLUMN vacaciones_fin TEXT NOT NULL DEFAULT ''",
    "columna_origen": "ALTER TABLE operadores ADD COLUMN columna_origen TEXT NOT NULL DEFAULT ''",
    "nota_estado": "ALTER TABLE operadores ADD COLUMN nota_estado TEXT NOT NULL DEFAULT ''",
    "foto_url": "ALTER TABLE operadores ADD COLUMN foto_url TEXT NOT NULL DEFAULT ''",
    "prevenido_hasta": "ALTER TABLE operadores ADD COLUMN prevenido_hasta TEXT NOT NULL DEFAULT ''",
    "prevenido_motivo": "ALTER TABLE operadores ADD COLUMN prevenido_motivo TEXT NOT NULL DEFAULT ''",
    "cambio_grupo": "ALTER TABLE operadores ADD COLUMN cambio_grupo TEXT NOT NULL DEFAULT ''",
    "tipo_personal": "ALTER TABLE operadores ADD COLUMN tipo_personal TEXT NOT NULL DEFAULT 'operador'",
    "lineas_cargo": "ALTER TABLE operadores ADD COLUMN lineas_cargo TEXT NOT NULL DEFAULT ''",
}

SUPERVISION_INITIAL_SEED: Tuple[Tuple[str, str], ...] = (
    ("José Useche", "Supervisor"),
    ("Juniors Delgado", "Supervisor"),
    ("Raynier Córdoba", "Supervisor"),
    ("Dany Blanco", "Supervisor"),
    ("Saúl Colmenares", "Supervisor"),
    ("Antonio Yanes", "Técnico de Producción"),
    ("Junior Padrino", "Técnico de Producción"),
    ("Iván Antón", "Especialista de Procesos"),
)

_SUPERVISION_SHORT_NAME_MIGRATION: Tuple[Tuple[str, str], ...] = (
    ("José Gabriel Useche Raga", "José Useche"),
    ("Juniors Arturo Delgado Muro", "Juniors Delgado"),
    ("Raynier Meyseir Córdoba Guerrero", "Raynier Córdoba"),
    ("Dany Antonio Blanco Cuello", "Dany Blanco"),
    ("Antonio José Yanes Marchena", "Antonio Yanes"),
    ("Junior David Padrino Cordovez", "Junior Padrino"),
    ("Iván Jesús Antón Cedeño", "Iván Antón"),
)


def infer_tipo_personal(puesto: str = "", columna: str = "", explicit_tipo: Any = None) -> str:
    raw = str(explicit_tipo or "").strip().lower()
    if raw in ("operador", "supervision"):
        return raw
    if str(columna or "").strip().upper() == "SUPERVISION":
        return "supervision"
    pn = normalize_text(puesto)
    if any(k in pn for k in ("supervisor", "supevisor", "tecnico de produccion", "especialista de procesos")):
        return "supervision"
    return "operador"


def supervision_rows_to_seed(existing_names: Set[str]) -> List[Tuple[str, str]]:
    return [
        (nombre, puesto)
        for nombre, puesto in SUPERVISION_INITIAL_SEED
        if normalize_text(nombre) not in existing_names
    ]


def sanitize_lineas_cargo(raw_val: Any) -> str:
    if raw_val is None:
        return ""
    if isinstance(raw_val, list):
        tokens = [str(x).strip().upper() for x in raw_val]
    else:
        tokens = [t.strip().upper() for t in str(raw_val).split(",")]
    cleaned: List[str] = []
    seen: Set[str] = set()
    for t in tokens:
        if t in LINEAS_CARGO_VALIDAS and t not in seen:
            seen.add(t)
            cleaned.append(t)
    if "TODAS" in cleaned:
        return "TODAS"
    return ",".join(cleaned)


def ensure_operadores_columns(conn: Any) -> None:
    cols = conn.table_columns("operadores")
    for col_name, alter_sql in _OPERADORES_ALTER_SQL.items():
        if col_name not in cols:
            conn.execute(alter_sql)


_db_initialized_keys: Set[str] = set()
_db_init_lock = threading.Lock()


def _current_db_key() -> str:
    return f"pg:{os.environ.get('DATABASE_URL', '').strip()}" if is_postgres() else f"sqlite:{os.path.abspath(get_sqlite_path())}"


def ensure_db_ready() -> None:
    key = _current_db_key()
    if key not in _db_initialized_keys:
        init_db()


def init_db() -> None:
    key = _current_db_key()
    with _db_init_lock:
        if key in _db_initialized_keys:
            return
        with db_lock:
            with get_db() as conn:
                init_auth_tables(conn)
                conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS operadores (
                        id TEXT PRIMARY KEY,
                        nombre TEXT NOT NULL,
                        columna TEXT NOT NULL,
                        estado TEXT NOT NULL DEFAULT 'normal',
                        puesto TEXT NOT NULL DEFAULT '',
                        grupo TEXT NOT NULL DEFAULT '',
                        supervisor_grupo TEXT NOT NULL DEFAULT '',
                        numero TEXT NOT NULL DEFAULT '',
                        posicion INTEGER NOT NULL DEFAULT 0,
                        updated_at TEXT NOT NULL,
                        updated_by TEXT NOT NULL DEFAULT 'Sistema'
                    )
                    """
                )
                ensure_operadores_columns(conn)

                conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS config_meta (
                        key TEXT PRIMARY KEY,
                        value TEXT NOT NULL DEFAULT ''
                    )
                    """
                )

                if conn.pg:
                    conn.execute(
                        """
                        CREATE TABLE IF NOT EXISTS historial (
                            id SERIAL PRIMARY KEY,
                            descripcion TEXT NOT NULL,
                            supervisor TEXT NOT NULL,
                            tipo TEXT NOT NULL,
                            created_at TEXT NOT NULL
                        )
                        """
                    )
                else:
                    conn.execute(
                        """
                        CREATE TABLE IF NOT EXISTS historial (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            descripcion TEXT NOT NULL,
                            supervisor TEXT NOT NULL,
                            tipo TEXT NOT NULL,
                            created_at TEXT NOT NULL
                        )
                        """
                    )

                conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS ultimo_import (
                        id INTEGER PRIMARY KEY CHECK (id = 1),
                        titulo TEXT NOT NULL DEFAULT '',
                        payload_json TEXT NOT NULL,
                        saved_at TEXT NOT NULL,
                        saved_by TEXT NOT NULL DEFAULT 'Sistema'
                    )
                    """
                )

                # Índices compuestos para consultas de lectura y filtrado sin escaneo completo
                conn.execute("CREATE INDEX IF NOT EXISTS idx_operadores_col_pos_nombre ON operadores (columna, posicion, nombre)")
                conn.execute("CREATE INDEX IF NOT EXISTS idx_operadores_estado_vac ON operadores (estado, vacaciones_inicio, vacaciones_fin)")
                conn.execute("CREATE INDEX IF NOT EXISTS idx_operadores_prevenido ON operadores (prevenido_hasta)")
                conn.execute("CREATE INDEX IF NOT EXISTS idx_operadores_grupo_col ON operadores (grupo, columna)")
                conn.execute("CREATE INDEX IF NOT EXISTS idx_historial_tipo_id ON historial (tipo, id DESC)")

            # Limpiar solo si la BD aún tenía exclusivamente los 23 registros demo heredados
            has_import = conn.execute("SELECT COUNT(*) AS c FROM ultimo_import").fetchone()["c"]
            if has_import == 0:
                rows = conn.execute("SELECT id, updated_by FROM operadores").fetchall()
                if rows and all(r["id"] in LEGACY_DEMO_IDS and r["updated_by"] == "Sistema" for r in rows):
                    conn.execute("DELETE FROM operadores")
                    conn.execute("DELETE FROM historial")
                    conn.execute(
                        """
                        INSERT INTO historial (descripcion, supervisor, tipo, created_at)
                        VALUES (?, ?, ?, ?)
                        """,
                        (
                            "Tablero listo para importar personal desde Excel o CSV.",
                            "Sistema",
                            "sistema",
                            now_iso(),
                        ),
                    )

            # Migrar operadores de Siat y Equipos Móviles/Montacargas fuera de RELEVO a sus cuadros propios
            migrated = False
            res_siat = conn.execute(
                """
                UPDATE operadores
                SET columna = 'SIAT', estado = CASE WHEN estado = 'relevo' THEN 'normal' ELSE estado END, puesto = 'Siat'
                WHERE columna = 'RELEVO' AND LOWER(puesto) LIKE '%siat%'
                """
            )
            if res_siat.rowcount:
                migrated = True

            res_mont = conn.execute(
                """
                UPDATE operadores
                SET columna = 'MONTACARGAS', estado = CASE WHEN estado = 'relevo' THEN 'normal' ELSE estado END, puesto = 'Montacargas'
                WHERE columna = 'RELEVO' AND (LOWER(puesto) LIKE '%movil%' OR LOWER(puesto) LIKE '%móvil%' OR LOWER(puesto) LIKE '%montacarga%')
                """
            )
            if res_mont.rowcount:
                migrated = True

            if migrated:
                for col_name in ("RELEVO", "SIAT", "MONTACARGAS"):
                    reindex_column(conn, col_name)

            # Actualizar también ultimo_import guardado para conservar separación en Restaurar
            imp_row = conn.execute("SELECT payload_json FROM ultimo_import WHERE id = 1").fetchone()
            if imp_row:
                try:
                    imp_list = json.loads(imp_row["payload_json"])
                    imp_changed = False
                    for item in imp_list:
                        p_norm = normalize_text(item.get("puesto", ""))
                        if item.get("columna") == "RELEVO" and "siat" in p_norm:
                            item["columna"] = "SIAT"
                            item["puesto"] = "Siat"
                            if item.get("estado") == "relevo":
                                item["estado"] = "normal"
                            imp_changed = True
                        elif item.get("columna") == "RELEVO" and ("movil" in p_norm or "montacarga" in p_norm):
                            item["columna"] = "MONTACARGAS"
                            item["puesto"] = "Montacargas"
                            if item.get("estado") == "relevo":
                                item["estado"] = "normal"
                            imp_changed = True
                    if imp_changed:
                        conn.execute(
                            "UPDATE ultimo_import SET payload_json = ? WHERE id = 1",
                            (json.dumps(imp_list, ensure_ascii=False),),
                        )
                except Exception:
                    pass

            # Migración única: iniciar líneas vacías en blanco (con símbolo '+') moviendo la nómina a POOL
            has_empty_migration = conn.execute(
                "SELECT value FROM config_meta WHERE key = 'migrated_empty_lines_v1'"
            ).fetchone()
            if not has_empty_migration:
                conn.execute(
                    """
                    UPDATE operadores
                    SET columna_origen = CASE
                            WHEN columna_origen = '' AND columna IN ('L1','L2','L3','L4','RELEVO','SIAT','MONTACARGAS') THEN columna
                            ELSE columna_origen
                        END,
                        columna = 'POOL',
                        estado = CASE WHEN estado = 'relevo' THEN 'normal' ELSE estado END
                    WHERE columna != 'POOL'
                    """
                )
                conn.execute(
                    """
                    INSERT INTO config_meta (key, value) VALUES ('migrated_empty_lines_v1', ?)
                    ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value
                    """,
                    (now_iso(),),
                )
                reindex_column(conn, "POOL")

            # Migración v2: normalizar 'Preparador de paleta' -> 'Pela paletas' y 'Siat' -> 'Envolvedora Siat'
            has_slots_v2 = conn.execute(
                "SELECT value FROM config_meta WHERE key = 'migrated_slots_v2'"
            ).fetchone()
            if not has_slots_v2:
                conn.execute(
                    "UPDATE operadores SET puesto = 'Pela paletas' WHERE LOWER(puesto) LIKE '%preparador de paleta%'"
                )
                conn.execute(
                    "UPDATE operadores SET puesto = 'Envolvedora Siat' WHERE LOWER(TRIM(puesto)) = 'siat'"
                )
                conn.execute(
                    """
                    INSERT INTO config_meta (key, value) VALUES ('migrated_slots_v2', ?)
                    ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value
                    """,
                    (now_iso(),),
                )

            # Migración única: limpiar nombres duplicados en operadores conservando el más relevante
            has_dedup_v1 = conn.execute(
                "SELECT value FROM config_meta WHERE key = 'dedup_operadores_v1'"
            ).fetchone()
            if not has_dedup_v1:
                removed = deduplicate_operadores(conn)
                if removed > 0:
                    log_event(
                        conn,
                        f"🧹 Limpieza automática: {removed} registro(s) duplicado(s) eliminados.",
                        "Sistema",
                        "sistema",
                    )
                conn.execute(
                    """
                    INSERT INTO config_meta (key, value) VALUES ('dedup_operadores_v1', ?)
                    ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value
                    """,
                    (now_iso(),),
                )

            # Carga inicial única de Supervisores, Técnicos de Producción y Especialista de Procesos (en POOL para asignar por grupo)
            has_sup_seed = conn.execute(
                "SELECT value FROM config_meta WHERE key = 'seed_supervision_v1'"
            ).fetchone()
            if not has_sup_seed:
                existing_names = {
                    normalize_text(r["nombre"])
                    for r in conn.execute("SELECT nombre FROM operadores").fetchall()
                }
                max_pool_row = conn.execute(
                    "SELECT COALESCE(MAX(posicion), -1) AS m FROM operadores WHERE columna = 'POOL'"
                ).fetchone()
                next_pool_pos = int(max_pool_row["m"]) + 1 if max_pool_row else 0
                seed_rows: List[Tuple[Any, ...]] = []
                ts_seed = now_iso()
                for sup_nombre, sup_puesto in supervision_rows_to_seed(existing_names):
                    seed_rows.append((
                        uuid.uuid4().hex[:10],
                        sup_nombre,
                        "POOL",
                        "normal",
                        sup_puesto,
                        "",
                        "supervision",
                        "",
                        next_pool_pos,
                        ts_seed,
                        "Sistema",
                    ))
                    next_pool_pos += 1
                if seed_rows:
                    conn.executemany(
                        """
                        INSERT INTO operadores (
                            id, nombre, columna, estado, puesto, grupo, tipo_personal, lineas_cargo, posicion, updated_at, updated_by
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        seed_rows,
                    )
                conn.execute(
                    """
                    INSERT INTO config_meta (key, value) VALUES ('seed_supervision_v1', ?)
                    ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value
                    """,
                    (ts_seed,),
                )

            has_sup_short = conn.execute(
                "SELECT value FROM config_meta WHERE key = 'seed_supervision_short_names_v1'"
            ).fetchone()
            if not has_sup_short:
                ts_short = now_iso()
                current_rows = conn.execute("SELECT id, nombre FROM operadores").fetchall()
                current_norms = {normalize_text(r["nombre"]) for r in current_rows}
                for full_name, short_name in _SUPERVISION_SHORT_NAME_MIGRATION:
                    if normalize_text(short_name) in current_norms:
                        continue
                    full_norm = normalize_text(full_name)
                    for row in current_rows:
                        if normalize_text(row["nombre"]) == full_norm:
                            conn.execute(
                                "UPDATE operadores SET nombre = ?, updated_at = ? WHERE id = ?",
                                (short_name, ts_short, row["id"]),
                            )
                            current_norms.add(normalize_text(short_name))
                conn.execute(
                    """
                    INSERT INTO config_meta (key, value) VALUES ('seed_supervision_short_names_v1', ?)
                    ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value
                    """,
                    (ts_short,),
                )
        _db_initialized_keys.add(key)
        invalidate_board_cache()


def deduplicate_operadores(conn: Any) -> int:
    """Elimina registros duplicados por nombre normalizado conservando la fila prioritaria."""
    rows = conn.execute(
        """
        SELECT id, nombre, columna, estado, foto_url, prevenido_hasta, vacaciones_inicio, updated_at
        FROM operadores
        ORDER BY updated_at DESC, id ASC
        """
    ).fetchall()
    groups: Dict[str, List[Any]] = {}
    for r in rows:
        key = normalize_text(r["nombre"])
        if not key:
            continue
        groups.setdefault(key, []).append(r)

    def _keep_rank(row: Any) -> Tuple[int, int, int, int, str]:
        return (
            1 if row["columna"] != "POOL" else 0,
            1 if (row["foto_url"] or "").strip() else 0,
            1 if ((row["prevenido_hasta"] or "").strip() or (row["vacaciones_inicio"] or "").strip() or row["estado"] != "normal") else 0,
            1 if (row["estado"] or "") != "normal" else 0,
            str(row["updated_at"] or ""),
        )

    removed_ids: List[str] = []
    for _key, items in groups.items():
        if len(items) <= 1:
            continue
        items_sorted = sorted(items, key=_keep_rank, reverse=True)
        for dup in items_sorted[1:]:
            removed_ids.append(dup["id"])

    if removed_ids:
        conn.executemany("DELETE FROM operadores WHERE id = ?", [(rid,) for rid in removed_ids])
        reindex_column(conn, "POOL")
    return len(removed_ids)


def get_actor_supervisor(data: Optional[Dict[str, Any]] = None, fallback: str = "Supervisor") -> str:
    """Prioriza siempre la identidad verificada en la sesión del servidor."""
    user = session.get("user") or {}
    session_name = clean_cell(user.get("nombre") or user.get("username"))
    if session_name:
        return session_name
    if isinstance(data, dict) and data.get("supervisor"):
        return clean_cell(data.get("supervisor")) or fallback
    return fallback


def log_events_batch(conn: Any, events: List[Tuple[str, str, str]]) -> None:
    if not events:
        return
    ts = now_iso()
    rows = [(desc, sup or "Supervisor", tipo, ts) for desc, sup, tipo in events]
    conn.executemany(
        """
        INSERT INTO historial (descripcion, supervisor, tipo, created_at)
        VALUES (?, ?, ?, ?)
        """,
        rows,
    )
    conn.execute(
        """
        DELETE FROM historial
        WHERE id NOT IN (
            SELECT id FROM historial ORDER BY id DESC LIMIT 200
        )
        """
    )


def log_event(conn: Any, descripcion: str, supervisor: str, tipo: str) -> None:
    log_events_batch(conn, [(descripcion, supervisor, tipo)])


_FIXED_CONTRACTUAL_HOLIDAYS: Tuple[Tuple[int, int, str], ...] = (
    (1, 1, "Año Nuevo"),
    (4, 19, "Declaración de la Independencia"),
    (5, 1, "Día del Trabajador"),
    (6, 24, "Batalla de Carabobo"),
    (6, 29, "San Pedro y San Pablo"),
    (7, 5, "Día de la Independencia"),
    (7, 24, "Natalicio de Simón Bolívar"),
    (10, 12, "Día de la Resistencia Indígena"),
    (10, 21, "Día de Los Teques"),
    (12, 24, "Nochebuena (Asueto)"),
    (12, 25, "Navidad"),
    (12, 31, "Fin de Año (Asueto)"),
)

_HOLIDAY_YEAR_CACHE: Dict[int, Dict[date, str]] = {}


def _easter_sunday(year: int) -> date:
    """Calcula el Domingo de Resurrección (algoritmo gregoriano) para Carnaval y Semana Santa."""
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)


def get_contractual_holidays(year: int) -> Dict[date, str]:
    """Devuelve el mapa de días feriados y de asueto contractual (SINTRAMAPEPC) para un año."""
    cached = _HOLIDAY_YEAR_CACHE.get(year)
    if cached is not None:
        return cached
    holidays: Dict[date, str] = {}
    for m, d, label in _FIXED_CONTRACTUAL_HOLIDAYS:
        holidays[date(year, m, d)] = label
    easter = _easter_sunday(year)
    movable = (
        (-48, "Lunes de Carnaval"),
        (-47, "Martes de Carnaval"),
        (-4, "Miércoles de Semana Santa"),
        (-3, "Jueves Santo"),
        (-2, "Viernes Santo"),
        (-1, "Sábado de Semana Santa"),
    )
    for offset, label in movable:
        holidays[easter + timedelta(days=offset)] = label
    _HOLIDAY_YEAR_CACHE[year] = holidays
    return holidays


def is_non_working_day(d: date) -> bool:
    return d.weekday() >= 5 or d in get_contractual_holidays(d.year)


def count_business_days_inclusive(ini_d: date, fin_d: date) -> int:
    """Cuenta días hábiles (lun-vie) excluyendo feriados y asuetos contractuales entre ini_d y fin_d."""
    if fin_d < ini_d:
        return 0
    total_cal = (fin_d - ini_d).days + 1
    full_weeks, rem = divmod(total_cal, 7)
    bdays = full_weeks * 5
    start_wd = ini_d.weekday()
    for offset in range(rem):
        if (start_wd + offset) % 7 < 5:
            bdays += 1

    # Descontar feriados/asuetos que caigan de lunes a viernes en el rango
    for yr in range(ini_d.year, fin_d.year + 1):
        for h_date in get_contractual_holidays(yr):
            if ini_d <= h_date <= fin_d and h_date.weekday() < 5:
                bdays -= 1
    return max(0, bdays)


def next_business_day(d: date) -> date:
    """Devuelve el siguiente día hábil posterior a d (saltando fines de semana y feriados)."""
    cand = d + timedelta(days=1)
    while is_non_working_day(cand):
        cand += timedelta(days=1)
    return cand


def op_vacation_metrics(op_row: Any, today_d: date) -> Dict[str, Any]:
    vac_ini = (op_row["vacaciones_inicio"] or "").strip()
    vac_fin = (op_row["vacaciones_fin"] or "").strip()
    estado = (op_row["estado"] or "normal").strip()

    ini_d = parse_ymd(vac_ini)
    fin_d = parse_ymd(vac_fin)

    en_vacaciones_hoy = (
        estado == "vacaciones"
        and (ini_d is None or ini_d <= today_d)
        and (fin_d is None or today_d <= fin_d)
    )
    vacaciones_programadas = bool(ini_d and ini_d > today_d and (fin_d is None or fin_d >= ini_d))
    inactivo = bool(en_vacaciones_hoy or estado in ESTADOS_INACTIVOS)

    dias_total = 0
    if ini_d and fin_d and fin_d >= ini_d:
        dias_total = count_business_days_inclusive(ini_d, fin_d)

    dias_restantes = 0
    if en_vacaciones_hoy and fin_d and fin_d >= today_d:
        dias_restantes = count_business_days_inclusive(today_d, fin_d)

    fecha_reactivacion = ""
    if fin_d:
        fecha_reactivacion = next_business_day(fin_d).strftime("%Y-%m-%d")

    return {
        "en_vacaciones_hoy": en_vacaciones_hoy,
        "vacaciones_programadas": vacaciones_programadas,
        "inactivo": inactivo,
        "dias_vacaciones_total": dias_total,
        "dias_vacaciones_restantes": dias_restantes,
        "fecha_reactivacion": fecha_reactivacion,
    }


def sync_expired_vacations(conn: Any) -> bool:
    """Activa vacaciones al llegar la fecha de inicio y reactiva al trabajador al día siguiente de terminar."""
    today_d = plant_now().date()
    now_str = now_iso()
    changed = False
    touched_cols: Set[str] = set()
    reactivate_params: List[Tuple[str, str]] = []
    activate_params: List[Tuple[str, str, str]] = []
    pending_logs: List[Tuple[str, str, str]] = []

    rows = conn.execute(
        """
        SELECT id, nombre, columna, columna_origen, estado, vacaciones_inicio, vacaciones_fin
        FROM operadores
        WHERE vacaciones_inicio != '' OR vacaciones_fin != '' OR estado = 'vacaciones'
        """
    ).fetchall()

    for r in rows:
        ini_d = parse_ymd(r["vacaciones_inicio"])
        fin_d = parse_ymd(r["vacaciones_fin"])

        if fin_d and today_d > fin_d:
            reactivate_params.append((now_str, r["id"]))
            pending_logs.append(
                (f"Fin de vacaciones: {r['nombre']} se reactivó automáticamente.", "Calendario", "estado")
            )
            changed = True
            continue

        in_range = (
            ini_d is not None
            and ini_d <= today_d
            and (fin_d is None or today_d <= fin_d)
        )
        if in_range and (r["estado"] != "vacaciones" or r["columna"] != "POOL"):
            col_origen = r["columna_origen"] or (r["columna"] if r["columna"] in COLUMNAS_TABLERO else "")
            if r["columna"] != "POOL":
                touched_cols.add(r["columna"])
                touched_cols.add("POOL")
            activate_params.append((col_origen, now_str, r["id"]))
            if r["estado"] != "vacaciones":
                pending_logs.append(
                    (f"Inicio de vacaciones: {r['nombre']} desactivado hasta {r['vacaciones_fin'] or 'nuevo aviso'}.", "Calendario", "estado")
                )
            changed = True

    if reactivate_params:
        conn.executemany(
            """
            UPDATE operadores
            SET estado = 'normal', vacaciones_inicio = '', vacaciones_fin = '', nota_estado = '', updated_at = ?
            WHERE id = ?
            """,
            reactivate_params,
        )
    if activate_params:
        conn.executemany(
            """
            UPDATE operadores
            SET estado = 'vacaciones', columna = 'POOL', columna_origen = ?, updated_at = ?
            WHERE id = ?
            """,
            activate_params,
        )

    for col in touched_cols:
        reindex_column(conn, col)

    prev_rows = conn.execute(
        """
        SELECT id, nombre, prevenido_hasta, prevenido_motivo
        FROM operadores
        WHERE prevenido_hasta != '' AND prevenido_hasta <= ?
        """,
        (now_str,),
    ).fetchall()
    if prev_rows:
        conn.executemany(
            """
            UPDATE operadores
            SET prevenido_hasta = '', prevenido_motivo = '', updated_at = ?
            WHERE id = ?
            """,
            [(now_str, pr["id"]) for pr in prev_rows],
        )
        for pr in prev_rows:
            pending_logs.append(
                (f"Tiempo de prevención cumplido: {pr['nombre']} ({pr['prevenido_motivo'] or 'Normalizado'}).", "Reloj", "estado")
            )
        changed = True

    if pending_logs:
        log_events_batch(conn, pending_logs)
    return changed


def canonical_slot_target(destino: str, slot_idx: int) -> Tuple[str, int]:
    """L1 slot 0 y L2 slot 0 comparten el mismo Pela paletas almacenado canónicamente en ('L1', 0)."""
    if destino in ("L1", "L2") and slot_idx == 0:
        return "L1", 0
    return destino, slot_idx


def slot_puesto_matches(slot_puesto: str, op_puesto: str) -> bool:
    sp = normalize_text(re.sub(r"\s+\d+$", "", slot_puesto or ""))
    op = normalize_text(op_puesto or "")
    if not sp or not op:
        return False
    if sp == op or sp in op or op in sp:
        return True
    if ("pela paleta" in sp or "preparador" in sp) and ("pela paleta" in op or "preparador" in op):
        return True
    if ("siat" in sp or "envolvedora" in sp) and ("siat" in op or "envolvedora" in op):
        return True
    if ("montacarguista" in sp or "montacarga" in sp) and ("montacarguista" in op or "montacarga" in op or "movil" in op):
        return True
    return False


def _row_val(row: Any, key: str, default: str = "") -> str:
    try:
        val = row[key]
        return default if val is None else str(val)
    except (KeyError, IndexError, TypeError):
        return default


def effective_op_group(row: Any) -> str:
    estado = _row_val(row, "estado", "normal") or "normal"
    cambio = _row_val(row, "cambio_grupo", "")
    grupo = _row_val(row, "grupo", "")
    if estado == "cambio_turno" and cambio:
        return cambio
    return grupo


def groups_collide_on_board(group_a: str, group_b: str) -> bool:
    ga = (group_a or "").strip()
    gb = (group_b or "").strip()
    if not ga and not gb:
        return True
    if not ga or not gb:
        return False
    return ga == gb


def find_free_slot(conn: Any, destino: str, op_puesto: str = "", target_group: str = "") -> Optional[int]:
    slots = SLOTS_DEFINICION.get(destino, [])
    rows = conn.execute(
        "SELECT posicion, grupo, estado, cambio_grupo FROM operadores WHERE columna = ?",
        (destino,),
    ).fetchall()
    occupied = {
        int(r["posicion"])
        for r in rows
        if groups_collide_on_board(effective_op_group(r), target_group)
    }

    # L2 slot 0 está ocupado si L1 slot 0 tiene Pela paletas del mismo grupo
    if destino == "L2":
        pelas = conn.execute(
            "SELECT id, grupo, estado, cambio_grupo FROM operadores WHERE columna = 'L1' AND posicion = 0"
        ).fetchall()
        if any(groups_collide_on_board(effective_op_group(p), target_group) for p in pelas):
            occupied.add(0)

    # 1. Buscar casilla vacía que coincida con el puesto del operador
    for s in slots:
        idx = int(s["idx"])
        if idx not in occupied and slot_puesto_matches(s["puesto"], op_puesto):
            return idx

    # 2. Primera casilla vacía disponible
    for s in slots:
        idx = int(s["idx"])
        if idx not in occupied:
            return idx

    # 3. RELEVO y SUPERVISION no tienen límite: buscan el primer índice libre
    if destino in ("RELEVO", "SUPERVISION"):
        idx = 0
        while idx in occupied:
            idx += 1
        return idx

    return None


def check_column_capacity(
    conn: Any, destino: str, exclude_op_id: Optional[str] = None, target_group: str = ""
) -> Optional[str]:
    max_cap = MAX_POR_COLUMNA.get(destino)
    if max_cap is None:
        return None
    rows = conn.execute(
        "SELECT id, grupo, estado, cambio_grupo FROM operadores WHERE columna = ?",
        (destino,),
    ).fetchall()
    current_count = sum(
        1
        for r in rows
        if r["id"] != exclude_op_id and groups_collide_on_board(effective_op_group(r), target_group)
    )
    if current_count >= max_cap:
        area_label = "Siat" if destino == "SIAT" else "Montacargas"
        group_note = f" en {target_group}" if target_group else ""
        return f"El área de {area_label}{group_note} solo permite máximo {max_cap} operadores ({current_count}/{max_cap})."
    return None


def build_state_payload(conn: Any, sync_vacations: bool = True) -> Dict[str, Any]:
    if sync_vacations:
        sync_expired_vacations(conn)
    now_dt = plant_now()
    today_d = now_dt.date()
    today_str = today_d.strftime("%Y-%m-%d")

    tablero: Dict[str, List[Dict[str, Any]]] = {col: [] for col in COLUMNAS_TABLERO}
    supervision: List[Dict[str, Any]] = []
    pool: List[Dict[str, Any]] = []

    rows = conn.execute(
        """
        SELECT id, nombre, columna, estado, puesto, grupo, supervisor_grupo, numero, posicion,
               vacaciones_inicio, vacaciones_fin, columna_origen, nota_estado, foto_url,
               prevenido_hasta, prevenido_motivo, cambio_grupo, tipo_personal, lineas_cargo,
               updated_at, updated_by
        FROM operadores
        ORDER BY columna ASC, posicion ASC, nombre ASC
        """
    ).fetchall()

    grupos_set = set()
    resumen = {
        "total_nomina": 0,
        "activos": 0,
        "en_lineas": 0,
        "supervision": 0,
        "disponibles_pool": 0,
        "vacaciones": 0,
        "ausentes": 0,
        "licencias": 0,
        "relevos_cambios": 0,
        "prevenidos": 0,
    }

    pela_paletas_ops: List[Dict[str, Any]] = []

    for r in rows:
        col = r["columna"] if r["columna"] in COLUMNAS_VALIDAS else "POOL"
        grupo = r["grupo"] or ""
        if grupo:
            grupos_set.add(grupo)

        metrics = op_vacation_metrics(r, today_d)
        estado = r["estado"] or "normal"
        prev_hasta = (r["prevenido_hasta"] or "").strip() if not metrics["inactivo"] else ""
        prev_motivo = (r["prevenido_motivo"] or "").strip() if prev_hasta else ""
        prev_activo = False
        prev_mins = 0
        if prev_hasta:
            try:
                dt_h = datetime.strptime(prev_hasta, "%Y-%m-%d %H:%M:%S")
                diff_s = (dt_h - now_dt).total_seconds()
                if diff_s > 0:
                    prev_activo = True
                    prev_mins = max(1, int((diff_s + 59) // 60))
            except ValueError:
                prev_hasta = ""
                prev_motivo = ""

        raw_foto = r["foto_url"] or ""
        foto_v = _cache_operator_photo(r["id"], raw_foto) if raw_foto.startswith("data:image/") else ""
        op_dict = {
            "id": r["id"],
            "nombre": r["nombre"],
            "columna": col,
            "columna_origen": r["columna_origen"] or "",
            "estado": estado,
            "puesto": r["puesto"] or "",
            "grupo": grupo,
            "cambio_grupo": (r["cambio_grupo"] or "") if estado == "cambio_turno" else "",
            "tipo_personal": infer_tipo_personal(r["puesto"] or "", col, r["tipo_personal"] if "tipo_personal" in r.keys() else ""),
            "lineas_cargo": sanitize_lineas_cargo(r["lineas_cargo"] if "lineas_cargo" in r.keys() else ""),
            "supervisor_grupo": r["supervisor_grupo"] or "",
            "numero": r["numero"] or "",
            "posicion": int(r["posicion"] or 0),
            "vacaciones_inicio": r["vacaciones_inicio"] or "",
            "vacaciones_fin": r["vacaciones_fin"] or "",
            "nota_estado": r["nota_estado"] or "",
            "foto_url": raw_foto,
            "foto_v": foto_v,
            "prevenido_activo": prev_activo,
            "prevenido_hasta": prev_hasta if prev_activo else "",
            "prevenido_hasta_iso": f"{prev_hasta.replace(' ', 'T')}{PLANT_TZ_ISO_SUFFIX}" if prev_activo else "",
            "prevenido_motivo": prev_motivo if prev_activo else "",
            "prevenido_minutos_restantes": prev_mins,
            "inactivo": metrics["inactivo"],
            "en_vacaciones_hoy": metrics["en_vacaciones_hoy"],
            "vacaciones_programadas": metrics["vacaciones_programadas"],
            "dias_vacaciones_total": metrics["dias_vacaciones_total"],
            "dias_vacaciones_restantes": metrics["dias_vacaciones_restantes"],
            "fecha_reactivacion": metrics["fecha_reactivacion"],
            "compartido_l1_l2": False,
            "updated_at": r["updated_at"],
            "updated_by": r["updated_by"],
        }

        resumen["total_nomina"] += 1
        if metrics["inactivo"]:
            if metrics["en_vacaciones_hoy"] or estado == "vacaciones":
                resumen["vacaciones"] += 1
            elif estado == "ausente":
                resumen["ausentes"] += 1
            elif estado == "licencia":
                resumen["licencias"] += 1
        else:
            resumen["activos"] += 1
            if prev_activo:
                resumen["prevenidos"] += 1
            if estado in ("relevo", "cambio_turno") or col == "RELEVO":
                resumen["relevos_cambios"] += 1

        if col == "POOL" or metrics["inactivo"]:
            op_dict["columna"] = "POOL"
            pool.append(op_dict)
            if not metrics["inactivo"]:
                resumen["disponibles_pool"] += 1
        elif col == "SUPERVISION":
            supervision.append(op_dict)
            resumen["supervision"] += 1
        else:
            resumen["en_lineas"] += 1
            if (col in ("L1", "L2")) and op_dict["posicion"] == 0:
                op_dict["compartido_l1_l2"] = True
                eff_g = effective_op_group(op_dict)
                existing_idx = next((i for i, p in enumerate(pela_paletas_ops) if groups_collide_on_board(effective_op_group(p), eff_g)), None)
                if existing_idx is None:
                    pela_paletas_ops.append(op_dict)
                elif col == "L1":
                    pela_paletas_ops[existing_idx] = op_dict
            else:
                tablero[col].append(op_dict)

    # Insertar Pela paletas compartido simultáneamente en L1 (slot 0) y L2 (slot 0) por cada grupo
    for idx_pp, pela_op in enumerate(pela_paletas_ops):
        l1_copy = dict(pela_op, columna="L1", posicion=0, compartido_l1_l2=True)
        l2_copy = dict(pela_op, columna="L2", posicion=0, compartido_l1_l2=True)
        tablero["L1"].insert(idx_pp, l1_copy)
        tablero["L2"].insert(idx_pp, l2_copy)

    hist_rows = conn.execute(
        """
        SELECT id, descripcion, supervisor, tipo, created_at
        FROM historial
        ORDER BY id DESC
        LIMIT 80
        """
    ).fetchall()
    historial = [dict(h) for h in hist_rows]

    imp_row = conn.execute("SELECT titulo, saved_at, saved_by FROM ultimo_import WHERE id = 1").fetchone()
    ultimo_import_info = dict(imp_row) if imp_row else None

    return {
        "tablero": tablero,
        "supervision": supervision,
        "pool": pool,
        "resumen": resumen,
        "historial": historial,
        "grupos": sorted(grupos_set),
        "ultimo_import": ultimo_import_info,
        "max_por_columna": MAX_POR_COLUMNA,
        "slots_definicion": SLOTS_DEFINICION,
        "server_time": now_iso(),
        "server_epoch_ms": int(time.time() * 1000),
        "plant_tz_suffix": PLANT_TZ_ISO_SUFFIX,
        "today": today_str,
    }


def broadcast_sse(event_name: str, payload: Dict[str, Any]) -> None:
    data = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    message = f"event: {event_name}\ndata: {data}\n\n"
    with sse_lock:
        dead_clients = []
        for client_queue in sse_clients:
            try:
                client_queue.put_nowait(message)
            except queue.Full:
                dead_clients.append(client_queue)
        for dead in dead_clients:
            if dead in sse_clients:
                sse_clients.remove(dead)


def reindex_column(conn: Any, columna: str) -> None:
    """Compacta solo POOL en lote para conservar intactos los índices fijos en L1..MONTACARGAS."""
    if columna != "POOL":
        return
    rows = conn.execute(
        "SELECT id, posicion FROM operadores WHERE columna = ? ORDER BY posicion ASC, updated_at ASC, nombre ASC",
        (columna,),
    ).fetchall()
    updates = [(idx, row["id"]) for idx, row in enumerate(rows) if int(row["posicion"] or 0) != idx]
    if updates:
        conn.executemany("UPDATE operadores SET posicion = ? WHERE id = ?", updates)


# ====================================================================
# Parser Inteligente de Excel / CSV de Envasado ("Rotacion por supervisor")
# ====================================================================
AREA_STOP_WORDS = (
    "actividades almacen",
    "reposo",
    "reposos",
    "licencia",
    "pesonal de acuerdo",
    "personal de acuerdo",
    "rotacion de relevo para comida",
)

NOT_A_PERSON_NAMES = {
    "linea 1",
    "linea 2",
    "linea 3",
    "linea 4",
    "relevo",
    "vacacionista",
    "siat",
    "equipos moviles",
    "sindicato",
    "delegado",
    "jarabero",
    "n#",
    "despaletizador",
    "despaletizado",
    "llenadora",
    "llnadora",
    "horno",
    "paletizador",
    "paletizado",
    "posicionadora",
    "smi",
    "etiquetadora",
    "robot",
    "ocme l1",
    "ocme l2",
    "mesas paletizadoras",
    "preparador de paleta",
    "transporte",
    "incorporados",
    "cambio de grupo",
}


def map_column_token(token: str) -> Optional[str]:
    norm = normalize_text(token)
    if not norm:
        return None
    if re.search(r"\blinea\s*1\b", norm) or norm == "l1":
        return "L1"
    if re.search(r"\blinea\s*2\b", norm) or norm == "l2":
        return "L2"
    if re.search(r"\blinea\s*3\b", norm) or norm == "l3":
        return "L3"
    if re.search(r"\blinea\s*4\b", norm) or norm == "l4":
        return "L4"
    if "siat" in norm:
        return "SIAT"
    if "equipos moviles" in norm or "montacarga" in norm:
        return "MONTACARGAS"
    if "relevo" in norm or "apoyo" in norm:
        return "RELEVO"
    return None


def looks_like_group_header(cell_val: str) -> Optional[Tuple[str, str]]:
    clean = clean_cell(cell_val)
    norm = normalize_text(clean)
    m = re.match(r"^grupo\s*([abc])\s*(?:\(([^)]*)\))?", norm)
    if not m:
        return None
    letter = m.group(1).upper()
    sup_match = re.search(r"\(([^)]*)\)", clean)
    supervisor_name = clean_cell(sup_match.group(1)) if sup_match else ""
    return f"Grupo {letter}", supervisor_name


def is_valid_operator_name(name: str) -> bool:
    clean = clean_cell(name)
    if not clean:
        return False
    if clean.isdigit():
        return False
    norm = normalize_text(clean)
    if norm in NOT_A_PERSON_NAMES:
        return False
    if looks_like_group_header(clean):
        return False
    if "rotacion por supervisor" in norm:
        return False
    if any(sw in norm for sw in AREA_STOP_WORDS):
        return False
    if len(clean) < 2:
        return False
    return True


def parse_envasado_matrix(rows: List[List[str]], sheet_name: str = "Hoja1") -> List[Dict[str, Any]]:
    """Analiza una matriz de celdas con formato multicapa 'Rotacion por supervisor'."""
    periodos: List[Dict[str, Any]] = []
    n_rows = len(rows)

    def get_cell(r: int, c: int) -> str:
        if 0 <= r < n_rows and 0 <= c < len(rows[r]):
            return clean_cell(rows[r][c])
        return ""

    # Encontrar todas las filas de inicio de sección ("Rotacion por supervisor ...")
    section_starts: List[Tuple[int, str]] = []
    for r_idx, row in enumerate(rows):
        for cell in row:
            norm = normalize_text(cell)
            if "rotacion por supervisor" in norm:
                section_starts.append((r_idx, clean_cell(cell)))
                break

    if not section_starts:
        section_starts = [(0, f"Rotación ({sheet_name})")]

    for s_idx, (start_r, titulo) in enumerate(section_starts):
        end_r = section_starts[s_idx + 1][0] if s_idx + 1 < len(section_starts) else n_rows

        # Estado activo de bloques por posición de columna de Excel
        active_blocks: Dict[int, Dict[str, Any]] = {}
        operadores: List[Dict[str, Any]] = []
        seen_keys = set()

        def flush_section():
            nonlocal active_blocks
            active_blocks = {}

        for r in range(start_r, end_r):
            row = rows[r]
            row_norms = [normalize_text(c) for c in row]

            # Detectar corte definitivo de metadatos inferiores de la sección
            joined_norm = " ".join(c for c in row_norms if c)
            if not joined_norm:
                for b in active_blocks.values():
                    b["had_blank"] = True
                    if b.get("linea") == "L3" and b.get("saw_l3_op"):
                        b["after_l3_blank"] = True
                continue

            if any(sw in joined_norm for sw in AREA_STOP_WORDS):
                break
            if any(c == "sindicato" for c in row_norms) and any(
                "alberto lopez" in c or "victor hernandez" in c or "delegado" in c for c in row_norms
            ):
                break
            # Evitar tabla resumen inferior del año 2020 (ej. "Linea 1" sin Grupo A/B arriba tras haber visto L4/Relevo)
            if (
                active_blocks
                and any(b.get("saw_l4_or_relevo") for b in active_blocks.values())
                and any(c == "linea 1" for c in row_norms)
                and not any(looks_like_group_header(c) for c in row)
            ):
                break

            # ¿Esta fila declara cabeceras de Grupo A / Grupo B / Grupo C?
            group_headers_in_row = []
            for c_idx, cell in enumerate(row):
                gh = looks_like_group_header(cell)
                if gh:
                    group_headers_in_row.append((c_idx, gh[0], gh[1]))

            if group_headers_in_row:
                # Actualizar o crear bloques por columna (tolerando +-1 columna entre sub-tablas)
                for col_idx, grupo_nombre, sup_nombre in group_headers_in_row:
                    matched_key = None
                    for existing_col in list(active_blocks.keys()):
                        if abs(existing_col - col_idx) <= 1:
                            matched_key = existing_col
                            break
                    if matched_key is not None and matched_key != col_idx:
                        block = active_blocks.pop(matched_key)
                    elif matched_key is not None:
                        block = active_blocks[matched_key]
                    else:
                        block = {
                            "grupo": grupo_nombre,
                            "supervisor": sup_nombre,
                            "linea": None,
                            "categoria": "linea",
                            "pending_l4": [],
                            "pending_relevo": [],
                            "saw_l3_op": False,
                            "after_l3_blank": False,
                            "had_blank": False,
                            "saw_l4_or_relevo": False,
                        }
                    block["grupo"] = grupo_nombre
                    if sup_nombre:
                        block["supervisor"] = sup_nombre
                    block["linea"] = None
                    block["categoria"] = "linea"
                    block["pending_l4"] = []
                    block["pending_relevo"] = []
                    block["saw_l3_op"] = False
                    block["after_l3_blank"] = False
                    block["had_blank"] = False
                    active_blocks[col_idx] = block
                continue

            if not active_blocks:
                continue

            # Ordenar columnas de bloque para saber hasta qué columna le pertenece a cada grupo
            sorted_cols = sorted(active_blocks.keys())
            for idx_pos, g_col in enumerate(sorted_cols):
                block = active_blocks[g_col]
                next_g_col = sorted_cols[idx_pos + 1] if idx_pos + 1 < len(sorted_cols) else len(row) + 5

                area_raw = get_cell(r, g_col)
                name_raw = get_cell(r, g_col + 1)
                puesto_raw = get_cell(r, g_col + 2)
                num_raw = get_cell(r, g_col + 3)
                if not num_raw and (g_col + 4) < next_g_col:
                    cand_num = get_cell(r, g_col + 4)
                    if cand_num.isdigit():
                        num_raw = cand_num

                area_norm = normalize_text(area_raw)
                name_norm = normalize_text(name_raw)
                puesto_norm = normalize_text(puesto_raw)

                # Fila completamente vacía para este grupo
                if not area_raw and not name_raw and not puesto_raw:
                    block["had_blank"] = True
                    if block["linea"] == "L3" and block["saw_l3_op"]:
                        block["after_l3_blank"] = True
                    continue

                # Actualizar estado de línea/área según etiqueta en columna g_col
                explicit_line = map_column_token(area_raw)
                if explicit_line == "L1":
                    block["linea"] = "L1"
                    block["categoria"] = "linea"
                    block["pending_l4"] = []
                    block["pending_relevo"] = []
                    block["saw_l3_op"] = False
                    block["after_l3_blank"] = False
                    block["had_blank"] = False
                elif explicit_line == "L2":
                    block["linea"] = "L2"
                    block["categoria"] = "linea"
                    block["pending_l4"] = []
                    block["pending_relevo"] = []
                    block["saw_l3_op"] = False
                    block["after_l3_blank"] = False
                    block["had_blank"] = False
                elif explicit_line == "L3":
                    block["linea"] = "L3"
                    block["categoria"] = "linea"
                    block["pending_l4"] = []
                    block["pending_relevo"] = []
                    block["saw_l3_op"] = False
                    block["after_l3_blank"] = False
                    block["had_blank"] = False
                elif explicit_line == "L4":
                    # Promover operadores de Despaletizador/Posicionadora que aparecen tras el espacio de L3
                    for prev_op in block["pending_l4"]:
                        prev_op["columna"] = "L4"
                        prev_op["categoria"] = "linea"
                        prev_op["estado"] = "normal"
                    block["pending_l4"] = []
                    block["pending_relevo"] = []
                    block["linea"] = "L4"
                    block["categoria"] = "linea"
                    block["saw_l4_or_relevo"] = True
                    block["after_l3_blank"] = False
                    block["had_blank"] = False
                elif explicit_line == "RELEVO":
                    # Promover operadores previos sin máquina inmediatamente antes de la etiqueta Relevo
                    for prev_op in block["pending_relevo"]:
                        prev_op["columna"] = "RELEVO"
                        prev_op["categoria"] = "relevo"
                        prev_op["estado"] = "relevo"
                    block["pending_relevo"] = []
                    block["pending_l4"] = []
                    block["linea"] = "RELEVO"
                    block["categoria"] = "relevo"
                    block["saw_l4_or_relevo"] = True
                    block["had_blank"] = False
                elif "vacacionista" in area_norm:
                    block["linea"] = "RELEVO"
                    block["categoria"] = "vacacionista"
                    block["pending_l4"] = []
                    block["pending_relevo"] = []
                    block["had_blank"] = False
                elif explicit_line == "MONTACARGAS" or "equipos moviles" in area_norm or "montacarga" in area_norm:
                    block["linea"] = "MONTACARGAS"
                    block["categoria"] = "equipos_moviles"
                    block["pending_l4"] = []
                    block["pending_relevo"] = []
                    block["had_blank"] = False
                elif explicit_line == "SIAT" or area_norm == "siat":
                    block["linea"] = "SIAT"
                    block["categoria"] = "siat"
                    block["pending_l4"] = []
                    block["pending_relevo"] = []
                    block["had_blank"] = False
                elif area_norm == "jarabero":
                    block["linea"] = "L4"
                    block["categoria"] = "linea"
                    block["pending_relevo"] = []
                    block["had_blank"] = False
                    if not puesto_raw:
                        puesto_raw = "Jarabero"
                        puesto_norm = "jarabero"
                    if not is_valid_operator_name(name_raw):
                        block["puesto_hint"] = "Jarabero"

                # Manejar inversión "Jarabero | Francisco Guevara"
                if name_norm == "jarabero" and is_valid_operator_name(puesto_raw):
                    name_raw, puesto_raw = puesto_raw, "Jarabero"
                    name_norm = normalize_text(name_raw)
                    puesto_norm = "jarabero"

                if not is_valid_operator_name(name_raw):
                    continue

                if block.get("puesto_hint") and not puesto_raw:
                    puesto_raw = block.pop("puesto_hint")
                    puesto_norm = normalize_text(puesto_raw)
                else:
                    block.pop("puesto_hint", None)

                if block["had_blank"]:
                    if block["linea"] == "L3" and block["saw_l3_op"]:
                        block["after_l3_blank"] = True
                    block["had_blank"] = False

                col_destino = block["linea"] or "L3"
                categoria = block["categoria"]

                # Detectar Siat y Montacargas / Equipos Móviles en sus cuadros propios
                if puesto_norm == "siat" or categoria == "siat":
                    col_destino = "SIAT"
                    categoria = "siat"
                    if not puesto_raw:
                        puesto_raw = "Siat"
                elif "movil" in puesto_norm or "montacarga" in puesto_norm or categoria == "equipos_moviles":
                    col_destino = "MONTACARGAS"
                    categoria = "equipos_moviles"
                    puesto_raw = "Montacargas"
                elif categoria == "vacacionista" and not puesto_raw:
                    puesto_raw = "Vacacionista"

                nombre_limpio = fix_spanish_accents(name_raw)
                puesto_limpio = normalize_puesto_label(fix_spanish_accents(puesto_raw))
                sup_limpio = fix_spanish_accents(block["supervisor"])
                puesto_norm = normalize_text(puesto_limpio)

                estado = "relevo" if col_destino == "RELEVO" else "normal"

                dedup_key = (block["grupo"], normalize_text(nombre_limpio), col_destino, puesto_norm)
                if dedup_key in seen_keys:
                    continue
                seen_keys.add(dedup_key)

                op_obj = {
                    "nombre": nombre_limpio,
                    "columna": col_destino,
                    "estado": estado,
                    "puesto": puesto_limpio,
                    "grupo": block["grupo"],
                    "supervisor_grupo": sup_limpio,
                    "numero": num_raw if num_raw.isdigit() else "",
                    "categoria": categoria,
                }
                operadores.append(op_obj)

                if col_destino == "L3" and categoria == "linea":
                    block["saw_l3_op"] = True
                    if block["after_l3_blank"]:
                        block["pending_l4"].append(op_obj)
                elif col_destino == "L4" and categoria == "linea":
                    if not puesto_norm:
                        block["pending_relevo"].append(op_obj)
                    else:
                        block["pending_relevo"] = []

        flush_section()
        if operadores:
            grupos_detectados = sorted({op["grupo"] for op in operadores if op["grupo"]})
            periodos.append(
                {
                    "id": f"{sheet_name}__{len(periodos)}",
                    "titulo": titulo,
                    "hoja": sheet_name,
                    "total": len(operadores),
                    "grupos": grupos_detectados,
                    "operadores": operadores,
                }
            )

    return periodos


def parse_tabular_rows(rows: List[List[str]], sheet_name: str = "Hoja1") -> List[Dict[str, Any]]:
    """Fallback para hojas simples tipo tabla: Nombre | Línea | Puesto | Grupo | Estado."""
    header_idx = None
    col_map: Dict[str, int] = {}

    for r_idx, row in enumerate(rows[:25]):
        norms = [normalize_text(c) for c in row]
        if any(c in ("nombre", "operador", "trabajador", "nombre y apellido") for c in norms):
            header_idx = r_idx
            for c_idx, h in enumerate(norms):
                if h in ("nombre", "operador", "trabajador", "nombre y apellido"):
                    col_map["nombre"] = c_idx
                elif h in ("linea", "linea/area", "columna", "area"):
                    col_map["linea"] = c_idx
                elif h in ("puesto", "cargo", "maquina", "equipo"):
                    col_map["puesto"] = c_idx
                elif h in ("grupo", "turno"):
                    col_map["grupo"] = c_idx
                elif h in ("estado", "condicion"):
                    col_map["estado"] = c_idx
                elif h in ("supervisor", "jefe"):
                    col_map["supervisor"] = c_idx
                elif h in ("n#", "numero", "num", "#"):
                    col_map["numero"] = c_idx
            break

    if header_idx is None or "nombre" not in col_map:
        return []

    operadores: List[Dict[str, Any]] = []
    for row in rows[header_idx + 1 :]:
        def cell_at(key: str) -> str:
            idx = col_map.get(key)
            if idx is None or idx >= len(row):
                return ""
            return clean_cell(row[idx])

        nombre = cell_at("nombre")
        if not is_valid_operator_name(nombre):
            continue

        linea_raw = cell_at("linea")
        columna = map_column_token(linea_raw) or "L3"
        puesto = cell_at("puesto")
        grupo = cell_at("grupo")
        sup = cell_at("supervisor")
        numero = cell_at("numero")
        estado_norm = normalize_text(cell_at("estado"))

        puesto_norm = normalize_text(puesto)
        categoria = "relevo" if columna == "RELEVO" else "linea"
        if puesto_norm == "siat" or columna == "SIAT":
            columna = "SIAT"
            categoria = "siat"
            if not puesto:
                puesto = "Siat"
        elif "vacacion" in puesto_norm:
            categoria = "vacacionista"
        elif "movil" in puesto_norm or "montacarga" in puesto_norm or columna == "MONTACARGAS":
            columna = "MONTACARGAS"
            categoria = "equipos_moviles"
            if not puesto or "movil" in puesto_norm:
                puesto = "Montacargas"

        if "cambio" in estado_norm:
            estado = "cambio_turno"
        elif "relevo" in estado_norm or columna == "RELEVO":
            estado = "relevo"
        else:
            estado = "normal"

        operadores.append(
            {
                "nombre": fix_spanish_accents(nombre),
                "columna": columna,
                "estado": estado,
                "puesto": normalize_puesto_label(fix_spanish_accents(puesto)),
                "grupo": grupo,
                "supervisor_grupo": fix_spanish_accents(sup),
                "numero": numero,
                "categoria": categoria,
            }
        )

    if not operadores:
        return []

    grupos_detectados = sorted({op["grupo"] for op in operadores if op["grupo"]})
    return [
        {
            "id": f"{sheet_name}__tabla",
            "titulo": f"Tabla de Operadores ({sheet_name})",
            "hoja": sheet_name,
            "total": len(operadores),
            "grupos": grupos_detectados,
            "operadores": operadores,
        }
    ]


def parse_rows_smart(rows: List[List[str]], sheet_name: str) -> List[Dict[str, Any]]:
    has_envasado_markers = any(
        "rotacion por supervisor" in normalize_text(cell) or looks_like_group_header(cell)
        for row in rows
        for cell in row
    )
    if has_envasado_markers:
        periodos = parse_envasado_matrix(rows, sheet_name)
        if periodos:
            return periodos
    return parse_tabular_rows(rows, sheet_name)


def decode_text_bytes(raw: bytes) -> str:
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def parse_delimited_text(text: str) -> List[List[str]]:
    sample = "\n".join(text.splitlines()[:30])
    delimiter = ","
    if "\t" in sample and sample.count("\t") >= sample.count(","):
        delimiter = "\t"
    elif sample.count(";") > sample.count(","):
        delimiter = ";"
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    return [[clean_cell(c) for c in row] for row in reader]


def choose_default_period_id(periodos: List[Dict[str, Any]]) -> Optional[str]:
    if not periodos:
        return None
    # Preferir la sección con año más alto en el título (p.ej. 2026-2027)
    best_id = periodos[0]["id"]
    best_score = (-1, -1)
    for idx, p in enumerate(periodos):
        years = [int(y) for y in re.findall(r"\b(20\d{2})\b", p["titulo"])]
        max_year = max(years) if years else 0
        score = (max_year, p["total"])
        if score > best_score:
            best_score = score
            best_id = p["id"]
    return best_id


# ====================================================================
# Rutas HTTP, Caché TTL/ETag, Compresión Brotli/Gzip, SSR y SSE
# ====================================================================
def get_cached_state_payload() -> Tuple[Dict[str, Any], str]:
    """Obtiene el estado del tablero usando caché TTL con invalidación automática por versión y singleflight."""
    ensure_db_ready()
    now_dt = plant_now()
    cache_key = f"tablero:{now_dt.strftime('%Y-%m-%d_%H:%M')}"
    hit, cached_payload, cached_etag = _board_cache.get(cache_key)
    if hit and isinstance(cached_payload, dict):
        return cached_payload, cached_etag

    with _state_build_lock:
        hit, cached_payload, cached_etag = _board_cache.get(cache_key)
        if hit and isinstance(cached_payload, dict):
            return cached_payload, cached_etag

        v_start = _board_cache.version
        t0 = time.perf_counter()
        with get_db() as conn:
            payload = build_state_payload(conn)
        payload["state_version"] = v_start
        record_timing("db", (time.perf_counter() - t0) * 1000.0)

        ser_t0 = time.perf_counter()
        raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        etag = compute_etag(raw, prefix=f"b{v_start}")
        record_timing("serialize", (time.perf_counter() - ser_t0) * 1000.0)

        _board_cache.set_if_version(cache_key, payload, expected_version=v_start, etag=etag, ttl_s=15.0)
        return payload, etag


@app.after_request
def add_performance_cache_headers(response):
    if request.path.startswith("/static/") or request.path == "/favicon.ico":
        response.headers["Cache-Control"] = "public, max-age=604800, stale-while-revalidate=86400"
    elif request.path == "/":
        response.headers["Cache-Control"] = "private, no-cache, must-revalidate"

    ae = request.headers.get("Accept-Encoding", "").lower()
    if (
        response.mimetype != "text/event-stream"
        and response.status_code == 200
        and "Content-Encoding" not in response.headers
        and ("br" in ae or "gzip" in ae)
        and any(
            response.mimetype.startswith(p)
            for p in ("text/", "application/json", "application/javascript", "image/svg+xml")
        )
    ):
        response.direct_passthrough = False
        raw = response.get_data()
        if len(raw) > 512:
            comp_t0 = time.perf_counter()
            if "br" in ae and brotli is not None:
                compressed = brotli.compress(raw, quality=4)
                encoding = "br"
            elif "gzip" in ae:
                compressed = gzip.compress(raw, compresslevel=5)
                encoding = "gzip"
            else:
                compressed = None
                encoding = ""
            if compressed is not None and len(compressed) < len(raw):
                response.set_data(compressed)
                response.headers["Content-Encoding"] = encoding
                response.headers["Content-Length"] = str(len(compressed))
                response.vary.add("Accept-Encoding")
                record_timing("compress", (time.perf_counter() - comp_t0) * 1000.0)
    return response


@app.route("/favicon.ico")
def serve_favicon():
    return send_from_directory(STATIC_DIR, "favicon.ico")


def _safe_ssr_json(obj: Any) -> str:
    """Serializa JSON de forma segura para incrustar en <script type='application/json'>."""
    return (
        json.dumps(obj, ensure_ascii=False, separators=(",", ":"))
        .replace("<", "\\u003c")
        .replace(">", "\\u003e")
        .replace("&", "\\u0026")
    )


@app.route("/")
def serve_index():
    payload, _etag = get_cached_state_payload()
    login_ts = session.get("login_ts", 0)
    expira_en_s = max(0, int(SESION_MAX_S - (time.time() - login_ts))) if login_ts else 0
    ssr_bundle = dict(payload)
    ssr_bundle["current_user"] = session.get("user")
    ssr_bundle["expira_en_s"] = expira_en_s
    ssr_bundle["csrf_token"] = get_csrf_token()

    nonce = get_csp_nonce()
    html = _get_index_html_template()
    html = html.replace("__CSP_NONCE__", nonce)
    html = html.replace("__CSRF_TOKEN__", get_csrf_token())
    html = html.replace(
        f'<script id="ssr-initial-state" type="application/json" nonce="{nonce}">{{}}</script>',
        f'<script id="ssr-initial-state" type="application/json" nonce="{nonce}">{_safe_ssr_json(ssr_bundle)}</script>',
        1,
    )
    resp = Response(html, mimetype="text/html; charset=utf-8")
    resp.headers["Cache-Control"] = "private, no-cache, must-revalidate"
    return resp


@app.route("/terminos", methods=["GET"])
def serve_terminos():
    return _render_legal_page("terminos")


@app.route("/privacidad", methods=["GET"])
def serve_privacidad():
    return _render_legal_page("privacidad")


@app.route("/cookies", methods=["GET"])
def serve_cookies():
    return _render_legal_page("cookies")


def _render_legal_page(section: str):
    titles = {
        "terminos": "Términos del Servicio y Uso Interno",
        "privacidad": "Política de Privacidad y Protección de Datos",
        "cookies": "Política de Cookies y Almacenamiento Local",
    }
    html = render_template(
        "legal.html",
        section=section,
        page_title=titles.get(section, "Información Legal"),
        csp_nonce=get_csp_nonce(),
    )
    raw = html.encode("utf-8")
    etag = compute_etag(raw, prefix=f"legal-{section}")
    if etag_matches(request.headers.get("If-None-Match"), etag):
        resp = Response(status=304)
        resp.headers["ETag"] = etag
        resp.headers["Cache-Control"] = "public, max-age=300"
        return resp
    resp = Response(html, mimetype="text/html; charset=utf-8")
    resp.headers["ETag"] = etag
    resp.headers["Cache-Control"] = "public, max-age=300"
    return resp


@app.route("/api/tablero", methods=["GET"])
def api_get_tablero():
    base_payload, base_etag = get_cached_state_payload()
    uid = str((session.get("user") or {}).get("id") or "")
    etag = f'{base_etag[:-1]}-u{uid}"'
    if etag_matches(request.headers.get("If-None-Match"), etag):
        resp = Response(status=304)
        resp.headers["ETag"] = etag
        resp.headers["Cache-Control"] = "private, no-cache, must-revalidate"
        return resp

    payload = dict(base_payload)
    payload["current_user"] = session.get("user")
    login_ts = session.get("login_ts", 0)
    payload["expira_en_s"] = max(0, int(SESION_MAX_S - (time.time() - login_ts))) if login_ts else 0
    payload["csrf_token"] = get_csrf_token()
    resp = jsonify(payload)
    resp.headers["ETag"] = etag
    resp.headers["Cache-Control"] = "private, no-cache, must-revalidate"
    return resp


@app.route("/api/operadores", methods=["GET"])
def api_list_operadores():
    ensure_db_ready()
    try:
        page = max(1, int(request.args.get("page", 1)))
        per_page = max(1, min(100, int(request.args.get("per_page", 25))))
    except ValueError:
        return jsonify({"error": "Parámetros de paginación inválidos."}), 400

    q = normalize_text(request.args.get("q", ""))
    col_f = (request.args.get("columna") or "").strip().upper()
    grupo_f = clean_cell(request.args.get("grupo", ""))
    estado_f = (request.args.get("estado") or "").strip().lower()

    payload, base_etag = get_cached_state_payload()
    all_ops: List[Dict[str, Any]] = []
    for col in COLUMNAS_TABLERO:
        all_ops.extend(payload["tablero"].get(col, []))
    all_ops.extend(payload.get("supervision", []))
    all_ops.extend(payload.get("pool", []))

    seen_ids: Set[str] = set()
    unique_ops: List[Dict[str, Any]] = []
    for op in all_ops:
        if op["id"] not in seen_ids:
            seen_ids.add(op["id"])
            unique_ops.append(op)

    filtered = unique_ops
    if col_f and col_f in COLUMNAS_VALIDAS:
        filtered = [o for o in filtered if o["columna"] == col_f]
    if grupo_f:
        filtered = [o for o in filtered if o.get("grupo", "") == grupo_f]
    if estado_f and estado_f in ESTADOS_VALIDOS:
        filtered = [o for o in filtered if o.get("estado", "") == estado_f]
    if q:
        filtered = [
            o
            for o in filtered
            if q in normalize_text(f"{o.get('nombre', '')} {o.get('puesto', '')} {o.get('grupo', '')}")
        ]

    total = len(filtered)
    total_pages = max(1, (total + per_page - 1) // per_page)
    start = (page - 1) * per_page
    items = filtered[start : start + per_page]

    etag = compute_etag(
        f"{base_etag}:{page}:{per_page}:{q}:{col_f}:{grupo_f}:{estado_f}".encode("utf-8"),
        prefix="ops",
    )
    if etag_matches(request.headers.get("If-None-Match"), etag):
        resp = Response(status=304)
        resp.headers["ETag"] = etag
        resp.headers["Cache-Control"] = "private, no-cache, must-revalidate"
        return resp

    resp = jsonify({
        "items": items,
        "page": page,
        "per_page": per_page,
        "total": total,
        "total_pages": total_pages,
        "has_next": page < total_pages,
        "has_prev": page > 1,
    })
    resp.headers["ETag"] = etag
    resp.headers["Cache-Control"] = "private, no-cache, must-revalidate"
    return resp


@app.route("/api/operadores/<operador_id>/foto", methods=["GET"])
def api_get_operador_foto(operador_id: str):
    if not is_valid_op_id(operador_id):
        return jsonify({"error": "ID de operador inválido."}), 400

    with _photo_cache_lock:
        cached = _photo_cache.get(operador_id)
    req_v = (request.args.get("v") or "").strip()

    if not cached or (req_v and cached[0] != req_v):
        ensure_db_ready()
        with get_db() as conn:
            row = conn.execute("SELECT foto_url FROM operadores WHERE id = ?", (operador_id,)).fetchone()
        if not row or not (row["foto_url"] or "").strip():
            return jsonify({"error": "Foto no disponible."}), 404
        ver_hash = _cache_operator_photo(operador_id, row["foto_url"])
        if not ver_hash:
            return jsonify({"error": "Formato de foto inválido."}), 404
        with _photo_cache_lock:
            cached = _photo_cache.get(operador_id)
        if not cached:
            return jsonify({"error": "Foto no disponible."}), 404

    ver_hash, mime, img_bytes, etag = cached
    cache_hdr = (
        "private, max-age=31536000, immutable"
        if (req_v and req_v == ver_hash)
        else "private, no-cache, must-revalidate"
    )
    if etag_matches(request.headers.get("If-None-Match"), etag):
        resp = Response(status=304)
        resp.headers["ETag"] = etag
        resp.headers["Cache-Control"] = cache_hdr
        return resp

    resp = Response(img_bytes, mimetype=mime)
    resp.headers["ETag"] = etag
    resp.headers["Cache-Control"] = cache_hdr
    return resp


@app.route("/api/historial", methods=["GET"])
def api_list_historial():
    ensure_db_ready()
    try:
        page = max(1, int(request.args.get("page", 1)))
        per_page = max(1, min(100, int(request.args.get("per_page", 25))))
    except ValueError:
        return jsonify({"error": "Parámetros de paginación inválidos."}), 400

    tipo_f = (request.args.get("tipo") or "").strip().lower()[:32]
    offset = (page - 1) * per_page

    t0 = time.perf_counter()
    with get_db() as conn:
        if tipo_f:
            total = int(conn.execute("SELECT COUNT(*) AS c FROM historial WHERE tipo = ?", (tipo_f,)).fetchone()["c"])
            rows = conn.execute(
                "SELECT id, descripcion, supervisor, tipo, created_at FROM historial WHERE tipo = ? ORDER BY id DESC LIMIT ? OFFSET ?",
                (tipo_f, per_page, offset),
            ).fetchall()
        else:
            total = int(conn.execute("SELECT COUNT(*) AS c FROM historial").fetchone()["c"])
            rows = conn.execute(
                "SELECT id, descripcion, supervisor, tipo, created_at FROM historial ORDER BY id DESC LIMIT ? OFFSET ?",
                (per_page, offset),
            ).fetchall()
    record_timing("db", (time.perf_counter() - t0) * 1000.0)

    items = [dict(r) for r in rows]
    total_pages = max(1, (total + per_page - 1) // per_page)
    etag = compute_etag(
        f"hist:{_board_cache.version}:{page}:{per_page}:{tipo_f}:{total}:{items[0]['id'] if items else 0}".encode("utf-8"),
        prefix="hist",
    )
    if etag_matches(request.headers.get("If-None-Match"), etag):
        resp = Response(status=304)
        resp.headers["ETag"] = etag
        resp.headers["Cache-Control"] = "private, no-cache, must-revalidate"
        return resp

    resp = jsonify({
        "items": items,
        "page": page,
        "per_page": per_page,
        "total": total,
        "total_pages": total_pages,
        "has_next": page < total_pages,
        "has_prev": page > 1,
    })
    resp.headers["ETag"] = etag
    resp.headers["Cache-Control"] = "private, no-cache, must-revalidate"
    return resp


@app.route("/api/jobs/<job_id>", methods=["GET"])
def api_get_job(job_id: str):
    if not is_valid_op_id(job_id):
        return jsonify({"error": "ID de trabajo inválido."}), 400
    rec = bg_jobs.get_job(job_id)
    if not rec:
        return jsonify({"error": "Trabajo no encontrado."}), 404
    user = session.get("user") or {}
    uname = str(user.get("username") or "").strip().lower()
    rol = str(user.get("rol") or "").strip().lower()
    if rec["owner"] != uname and rol not in ROLES_ADMIN:
        return jsonify({"error": "No tienes permiso para consultar este trabajo."}), 403
    return jsonify(rec)


ESTADO_ETIQUETAS = {
    "normal": "Turno Normal (Activo)",
    "relevo": "Relevo (Activo)",
    "cambio_turno": "Cambio de Turno (Activo)",
    "vacaciones": "Vacaciones (Apagado)",
    "ausente": "Ausente",
    "licencia": "Licencia / Reposo",
}


def apply_operator_update(
    conn: Any, op_row: Any, data: Dict[str, Any], supervisor: str
) -> Tuple[Optional[str], int]:
    """Aplica cambios de nombre, metadata, estado, vacaciones/calendario y columna con validación de cupo."""
    now_dt = plant_now()
    today_d = now_dt.date()
    op_id = op_row["id"]

    nombre = clean_cell(data.get("nombre", op_row["nombre"]))[:60] or op_row["nombre"]
    puesto = clean_cell(data.get("puesto", op_row["puesto"]))[:80] if "puesto" in data else op_row["puesto"]
    if "puesto" in data:
        puesto = normalize_puesto_label(puesto)
    grupo = clean_cell(data.get("grupo", op_row["grupo"]))[:40] if "grupo" in data else op_row["grupo"]
    nota_estado = (
        clean_cell(data.get("nota_estado", op_row["nota_estado"]))[:200]
        if "nota_estado" in data
        else op_row["nota_estado"]
    )
    foto_url = op_row["foto_url"] if "foto_url" in op_row.keys() else ""
    if "foto_url" in data:
        raw_foto = str(data.get("foto_url") or "").strip()
        ok_img, msg_img = validar_imagen_data_url(raw_foto)
        if not ok_img:
            return msg_img or "La imagen no es válida.", 400
        foto_url = raw_foto

    estado_req = str(data.get("estado", op_row["estado"])).strip()
    if estado_req not in ESTADOS_VALIDOS:
        estado_req = op_row["estado"]

    vac_ini = str(data.get("vacaciones_inicio", op_row["vacaciones_inicio"]) or "").strip()
    vac_fin = str(data.get("vacaciones_fin", op_row["vacaciones_fin"]) or "").strip()

    if estado_req == "vacaciones":
        ini_d = parse_ymd(vac_ini)
        fin_d = parse_ymd(vac_fin)
        if vac_ini and ini_d is None:
            return "Fecha de inicio de vacaciones inválida.", 400
        if vac_fin and fin_d is None:
            return "Fecha de fin de vacaciones inválida.", 400
        if ini_d and fin_d and fin_d < ini_d:
            return "La fecha final de vacaciones debe ser igual o posterior a la fecha inicial.", 400
        if fin_d and today_d > fin_d:
            return "La fecha de fin de vacaciones ya pasó. Elige una fecha desde hoy en adelante.", 400

        if ini_d and ini_d > today_d:
            # Vacación programada a futuro: permanece activo hasta que inicie la fecha
            estado_final = "normal" if op_row["estado"] in ESTADOS_INACTIVOS else op_row["estado"]
            is_inactive_now = False
        else:
            estado_final = "vacaciones"
            is_inactive_now = True
    elif estado_req in ("ausente", "licencia"):
        estado_final = estado_req
        vac_ini = ""
        vac_fin = ""
        is_inactive_now = True
    else:
        estado_final = estado_req
        vac_ini = ""
        vac_fin = ""
        if "nota_estado" not in data:
            nota_estado = ""
        is_inactive_now = False

    origen_col = op_row["columna"]
    target_col = str(data.get("columna", origen_col)).strip().upper()
    if target_col not in COLUMNAS_VALIDAS:
        target_col = origen_col

    col_origen_saved = op_row["columna_origen"] or (origen_col if origen_col in COLUMNAS_TABLERO else "")

    prev_hasta = (op_row["prevenido_hasta"] if "prevenido_hasta" in op_row.keys() else "") or ""
    prev_motivo = (op_row["prevenido_motivo"] if "prevenido_motivo" in op_row.keys() else "") or ""
    cambio_grupo = (op_row["cambio_grupo"] if "cambio_grupo" in op_row.keys() else "") or ""
    tipo_personal = infer_tipo_personal(
        puesto,
        str(data.get("columna", op_row["columna"])).strip().upper(),
        data.get("tipo_personal") if "tipo_personal" in data else (op_row["tipo_personal"] if "tipo_personal" in op_row.keys() else ""),
    )
    lineas_cargo = (
        sanitize_lineas_cargo(data.get("lineas_cargo"))
        if "lineas_cargo" in data
        else sanitize_lineas_cargo(op_row["lineas_cargo"] if "lineas_cargo" in op_row.keys() else "")
    )
    if "cambio_grupo" in data:
        cambio_grupo = clean_cell(data.get("cambio_grupo"))[:40]
    if estado_final != "cambio_turno":
        cambio_grupo = ""

    eff_target_group = (
        cambio_grupo
        if (estado_final == "cambio_turno" and cambio_grupo)
        else clean_cell(data.get("grupo", op_row["grupo"]))[:40]
    )

    if is_inactive_now:
        if "columna" in data and target_col in COLUMNAS_TABLERO and target_col != origen_col:
            return "Un trabajador en vacaciones, ausente o con licencia está desactivado y no puede asignarse a una línea.", 400
        target_col = "POOL"
        prev_hasta = ""
        prev_motivo = ""
    else:
        if "prevenido_activo" in data:
            if bool(data.get("prevenido_activo")):
                raw_mins = data.get("prevenido_minutos")
                raw_hasta = str(data.get("prevenido_hasta") or "").strip()
                if raw_mins is not None and str(raw_mins).strip() != "":
                    if isinstance(raw_mins, bool):
                        return "Los minutos de prevención deben ser un número entero.", 400
                    try:
                        mins = int(raw_mins)
                    except (TypeError, ValueError):
                        return "Los minutos de prevención deben ser un número entero.", 400
                    if mins < 1 or mins > 720:
                        return "La duración temporal de prevención debe estar entre 1 y 720 minutos.", 400
                    prev_hasta = (now_dt + timedelta(minutes=mins)).strftime("%Y-%m-%d %H:%M:%S")
                elif raw_hasta:
                    parsed_dt = None
                    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
                        try:
                            parsed_dt = datetime.strptime(raw_hasta, fmt)
                            break
                        except ValueError:
                            pass
                    if parsed_dt is None and re.match(r"^\d{2}:\d{2}$", raw_hasta):
                        hm = datetime.strptime(raw_hasta, "%H:%M").time()
                        parsed_dt = datetime.combine(today_d, hm)
                        if parsed_dt <= now_dt:
                            parsed_dt = parsed_dt + timedelta(days=1)
                    if parsed_dt is None or parsed_dt <= now_dt:
                        return "La hora límite de prevención debe ser posterior a la hora actual.", 400
                    prev_hasta = parsed_dt.strftime("%Y-%m-%d %H:%M:%S")
                else:
                    prev_hasta = (now_dt + timedelta(minutes=30)).strftime("%Y-%m-%d %H:%M:%S")
                prev_motivo = clean_cell(data.get("prevenido_motivo")) or "Prevenido"
            else:
                prev_hasta = ""
                prev_motivo = ""

        if target_col != origen_col:
            cap_err = check_column_capacity(conn, target_col, exclude_op_id=op_id, target_group=eff_target_group)
            if cap_err:
                return cap_err, 400
            if target_col == "RELEVO" and estado_final == "normal":
                estado_final = "relevo"

    new_pos = int(op_row["posicion"] or 0)
    if target_col != origen_col:
        if target_col in ("POOL", "SUPERVISION"):
            max_row = conn.execute(
                "SELECT COALESCE(MAX(posicion), -1) AS m FROM operadores WHERE columna = ?",
                (target_col,),
            ).fetchone()
            new_pos = int(max_row["m"]) + 1
        else:
            free_idx = find_free_slot(conn, target_col, puesto, target_group=eff_target_group)
            if free_idx is None:
                return f"Todas las casillas de {target_col} están ocupadas.", 400
            target_col, new_pos = canonical_slot_target(target_col, free_idx)

    conn.execute(
        """
        UPDATE operadores
        SET nombre = ?, puesto = ?, grupo = ?, estado = ?, columna = ?, columna_origen = ?,
            vacaciones_inicio = ?, vacaciones_fin = ?, nota_estado = ?, foto_url = ?,
            prevenido_hasta = ?, prevenido_motivo = ?, cambio_grupo = ?,
            tipo_personal = ?, lineas_cargo = ?, posicion = ?,
            updated_at = ?, updated_by = ?
        WHERE id = ?
        """,
        (
            nombre,
            puesto,
            grupo,
            estado_final,
            target_col,
            col_origen_saved,
            vac_ini,
            vac_fin,
            nota_estado,
            foto_url,
            prev_hasta,
            prev_motivo,
            cambio_grupo,
            tipo_personal,
            lineas_cargo,
            new_pos,
            now_iso(),
            supervisor,
            op_id,
        ),
    )

    if origen_col == "POOL" and target_col != "POOL":
        reindex_column(conn, "POOL")

    prev_was = (op_row["prevenido_hasta"] if "prevenido_hasta" in op_row.keys() else "") or ""
    prev_cambio_grupo = (op_row["cambio_grupo"] if "cambio_grupo" in op_row.keys() else "") or ""

    # Bitácora legible
    if estado_final == "vacaciones":
        rango = f"del {vac_ini or '?'} al {vac_fin or '?'}"
        log_event(conn, f"🏖️ {nombre} en vacaciones {rango} (trabajador apagado)", supervisor, "estado")
    elif vac_ini and estado_final == "normal":
        log_event(conn, f"🗓️ Vacaciones programadas para {nombre} del {vac_ini} al {vac_fin}", supervisor, "estado")
    elif estado_final in ("ausente", "licencia"):
        extra = f" — {nota_estado}" if nota_estado else ""
        log_event(conn, f"{nombre} marcado como {ESTADO_ETIQUETAS[estado_final]}{extra}", supervisor, "estado")
    elif prev_hasta and prev_hasta != prev_was:
        log_event(conn, f"⚠️ Prevenido: {nombre} ({prev_motivo}, hasta {prev_hasta[11:16]})", supervisor, "estado")
    elif prev_was and not prev_hasta:
        log_event(conn, f"✓ Prevención retirada de {nombre}", supervisor, "estado")
    elif nombre != op_row["nombre"]:
        log_event(conn, f"Nombre actualizado: {op_row['nombre']} → {nombre}", supervisor, "estado")
    elif estado_final == "cambio_turno" and (estado_final != op_row["estado"] or cambio_grupo != prev_cambio_grupo):
        cubre_lbl = f"cubre {cambio_grupo}" if cambio_grupo else "todos los grupos"
        log_event(conn, f"⇄ {nombre} en Cambio de Turno ({cubre_lbl})", supervisor, "estado")
    elif estado_final != op_row["estado"]:
        log_event(conn, f"{nombre} → {ESTADO_ETIQUETAS.get(estado_final, estado_final)}", supervisor, "estado")
    elif target_col != origen_col:
        log_event(conn, f"{nombre} movido de {origen_col} a {target_col}", supervisor, "mover")
    else:
        log_event(conn, f"Datos actualizados de {nombre}", supervisor, "estado")

    return None, 200


@app.route("/api/asignaciones/mover", methods=["POST"])
def api_mover_operador():
    data = request.get_json(silent=True) or {}
    operador_id = str(data.get("operador_id", "")).strip()
    destino_raw = str(data.get("destino", "")).strip().upper()
    supervisor = get_actor_supervisor(data)
    posicion_raw = data.get("posicion")
    req_estado = str(data.get("estado", "") or "").strip()
    req_cambio_grupo = clean_cell(data.get("cambio_grupo", ""))[:40] if "cambio_grupo" in data else None
    req_target_group = clean_cell(data.get("target_group", ""))[:40]
    if req_target_group.upper() == "ALL":
        req_target_group = ""

    if not is_valid_op_id(operador_id) or destino_raw not in COLUMNAS_VALIDAS:
        return jsonify({"error": "operador_id y destino válido son requeridos."}), 400
    max_pos_limit = 10000 if destino_raw in ("RELEVO", "SUPERVISION") else 200
    if posicion_raw is not None and (not is_strict_int(posicion_raw) or posicion_raw < 0 or posicion_raw > max_pos_limit):
        return jsonify({"error": "Posición de casilla inválida."}), 400
    if req_estado and req_estado not in ESTADOS_VALIDOS:
        return jsonify({"error": "Estado inválido."}), 400

    with db_lock:
        with get_db() as conn:
            sync_expired_vacations(conn)
            op = conn.execute(
                """
                SELECT id, nombre, columna, columna_origen, estado, puesto, grupo,
                       vacaciones_inicio, vacaciones_fin, nota_estado, posicion, cambio_grupo, tipo_personal
                FROM operadores WHERE id = ?
                """,
                (operador_id,),
            ).fetchone()
            if not op:
                return jsonify({"error": "Operador no encontrado."}), 404

            metrics = op_vacation_metrics(op, plant_now().date())
            if destino_raw != "POOL" and metrics["inactivo"]:
                return jsonify(
                    {
                        "error": f"{op['nombre']} está apagado ({ESTADO_ETIQUETAS.get(op['estado'], op['estado'])}). Reactívalo antes de asignarlo a una línea."
                    }
                ), 400

            origen = op["columna"]
            origen_pos = int(op["posicion"] or 0)
            col_origen_saved = op["columna_origen"] or (origen if origen in COLUMNAS_TABLERO else "")
            timestamp = now_iso()

            reemplazado_nombre = ""
            if destino_raw == "POOL":
                max_row = conn.execute(
                    "SELECT COALESCE(MAX(posicion), -1) AS m FROM operadores WHERE columna = 'POOL'"
                ).fetchone()
                next_pool_pos = int(max_row["m"]) + 1
                nuevo_estado = "normal" if op["estado"] == "relevo" else op["estado"]
                conn.execute(
                    """
                    UPDATE operadores
                    SET columna = 'POOL', columna_origen = ?, estado = ?, posicion = ?, updated_at = ?, updated_by = ?
                    WHERE id = ?
                    """,
                    (col_origen_saved, nuevo_estado, next_pool_pos, timestamp, supervisor, operador_id),
                )
                if origen != "POOL":
                    log_event(conn, f"{op['nombre']} retirado de {origen} a Nómina Disponible", supervisor, "mover")
                payload = build_state_payload(conn, sync_vacations=False)
            elif destino_raw == "SUPERVISION":
                max_sup = conn.execute(
                    "SELECT COALESCE(MAX(posicion), -1) AS m FROM operadores WHERE columna = 'SUPERVISION'"
                ).fetchone()
                next_sup_pos = int(max_sup["m"]) + 1 if origen != "SUPERVISION" else origen_pos
                base_estado = req_estado or op["estado"]
                nuevo_estado = "normal" if base_estado == "relevo" else base_estado
                nuevo_cambio_grupo = (
                    (req_cambio_grupo if req_cambio_grupo is not None else (op["cambio_grupo"] or ""))
                    if nuevo_estado == "cambio_turno"
                    else ""
                )
                op_grupo_saved = (op["grupo"] or "") or req_target_group
                tipo_pers = "supervision"
                conn.execute(
                    """
                    UPDATE operadores
                    SET columna = 'SUPERVISION', grupo = ?, estado = ?, cambio_grupo = ?, tipo_personal = ?,
                        posicion = ?, updated_at = ?, updated_by = ?
                    WHERE id = ?
                    """,
                    (op_grupo_saved, nuevo_estado, nuevo_cambio_grupo, tipo_pers, next_sup_pos, timestamp, supervisor, operador_id),
                )
                if origen == "POOL":
                    reindex_column(conn, "POOL")
                if origen != "SUPERVISION":
                    grp_note = f" ({nuevo_cambio_grupo or op_grupo_saved})" if (nuevo_cambio_grupo or op_grupo_saved) else ""
                    log_event(conn, f"🛡️ {op['nombre']} asignado a Supervisión de Turno{grp_note}", supervisor, "mover")
                payload = build_state_payload(conn, sync_vacations=False)
            else:
                op_grupo_saved = (op["grupo"] or "") or req_target_group
                tmp_cambio_grupo = (
                    (req_cambio_grupo if req_cambio_grupo is not None else (op["cambio_grupo"] or ""))
                    if (req_estado or op["estado"]) == "cambio_turno"
                    else ""
                )
                target_eff_group = (
                    tmp_cambio_grupo
                    if ((req_estado or op["estado"]) == "cambio_turno" and tmp_cambio_grupo)
                    else op_grupo_saved
                )

                if is_strict_int(posicion_raw) and posicion_raw >= 0:
                    max_cap = MAX_POR_COLUMNA.get(destino_raw)
                    if max_cap is not None and posicion_raw >= max_cap:
                        area_label = "Siat" if destino_raw == "SIAT" else "Montacargas"
                        return jsonify({"error": f"El área de {area_label} solo tiene {max_cap} casillas (máximo {max_cap})."}), 400
                    target_slot_idx = posicion_raw
                else:
                    free_idx = find_free_slot(conn, destino_raw, op["puesto"], target_group=target_eff_group)
                    if free_idx is None:
                        cap_err = check_column_capacity(conn, destino_raw, exclude_op_id=operador_id, target_group=target_eff_group)
                        return jsonify({"error": cap_err or f"Todas las casillas de {destino_raw} están ocupadas."}), 400
                    target_slot_idx = free_idx

                dest_col, dest_slot = canonical_slot_target(destino_raw, target_slot_idx)

                if origen == dest_col and origen_pos == dest_slot and op_grupo_saved == (op["grupo"] or "") and (req_estado is None or req_estado == op["estado"]):
                    payload = build_state_payload(conn, sync_vacations=False)
                    return jsonify(_slim_state_for_broadcast(payload) if wants_slim_response() else payload)

                slot_rows = conn.execute(
                    "SELECT id, nombre, estado, columna, columna_origen, grupo, cambio_grupo FROM operadores WHERE columna = ? AND posicion = ? AND id != ?",
                    (dest_col, dest_slot, operador_id),
                ).fetchall()
                occupant = next(
                    (r for r in slot_rows if groups_collide_on_board(effective_op_group(r), target_eff_group)),
                    None,
                )

                if not occupant and dest_col != origen:
                    cap_err = check_column_capacity(conn, dest_col, exclude_op_id=operador_id, target_group=target_eff_group)
                    if cap_err:
                        return jsonify({"error": cap_err}), 400

                if occupant:
                    if origen in COLUMNAS_TABLERO:
                        occ_estado = "relevo" if origen == "RELEVO" else ("normal" if occupant["estado"] == "relevo" and origen != "RELEVO" else occupant["estado"])
                        conn.execute(
                            """
                            UPDATE operadores
                            SET columna = ?, posicion = ?, estado = ?, updated_at = ?, updated_by = ?
                            WHERE id = ?
                            """,
                            (origen, origen_pos, occ_estado, timestamp, supervisor, occupant["id"]),
                        )
                    else:
                        max_pool = conn.execute("SELECT COALESCE(MAX(posicion), -1) AS m FROM operadores WHERE columna = 'POOL'").fetchone()
                        occ_estado = "normal" if occupant["estado"] == "relevo" else occupant["estado"]
                        occ_orig = occupant["columna_origen"] or dest_col
                        conn.execute(
                            """
                            UPDATE operadores
                            SET columna = 'POOL', columna_origen = ?, posicion = ?, estado = ?, updated_at = ?, updated_by = ?
                            WHERE id = ?
                            """,
                            (occ_orig, int(max_pool["m"]) + 1, occ_estado, timestamp, supervisor, occupant["id"]),
                        )

                base_estado = req_estado or op["estado"]
                if dest_col == "RELEVO" and base_estado == "normal":
                    nuevo_estado = "relevo"
                elif origen == "RELEVO" and dest_col != "RELEVO" and base_estado == "relevo":
                    nuevo_estado = "normal"
                else:
                    nuevo_estado = base_estado

                nuevo_cambio_grupo = (
                    (req_cambio_grupo if req_cambio_grupo is not None else (op["cambio_grupo"] or ""))
                    if nuevo_estado == "cambio_turno"
                    else ""
                )

                conn.execute(
                    """
                    UPDATE operadores
                    SET columna = ?, columna_origen = ?, grupo = ?, estado = ?, cambio_grupo = ?, posicion = ?, updated_at = ?, updated_by = ?
                    WHERE id = ?
                    """,
                    (dest_col, col_origen_saved, op_grupo_saved, nuevo_estado, nuevo_cambio_grupo, dest_slot, timestamp, supervisor, operador_id),
                )

                if origen == "POOL":
                    reindex_column(conn, "POOL")

                slot_defs = SLOTS_DEFINICION.get(destino_raw, [])
                slot_name = slot_defs[target_slot_idx]["puesto"] if target_slot_idx < len(slot_defs) else f"Casilla {target_slot_idx + 1}"
                if dest_col == "L1" and dest_slot == 0:
                    if occupant and origen == "POOL":
                        reemplazado_nombre = occupant["nombre"]
                        log_event(
                            conn,
                            f"Reemplazo en Pela paletas (L1 y L2): {op['nombre']} entró y {occupant['nombre']} pasó a Nómina Disponible",
                            supervisor,
                            "mover",
                        )
                    else:
                        log_event(conn, f"{op['nombre']} asignado a Pela paletas (compartido L1 y L2)", supervisor, "mover")
                elif occupant and origen in COLUMNAS_TABLERO:
                    log_event(conn, f"Intercambio: {op['nombre']} ↔ {occupant['nombre']} ({destino_raw} · {slot_name})", supervisor, "mover")
                elif occupant and origen == "POOL":
                    reemplazado_nombre = occupant["nombre"]
                    log_event(
                        conn,
                        f"Reemplazo: {op['nombre']} ubicado en {destino_raw} · {slot_name} ({occupant['nombre']} pasó a Nómina Disponible)",
                        supervisor,
                        "mover",
                    )
                else:
                    log_event(conn, f"{op['nombre']} ubicado en {destino_raw} · {slot_name}", supervisor, "mover")

                payload = build_state_payload(conn, sync_vacations=False)

    slim = publish_state_payload(payload)
    out = slim if wants_slim_response() else payload
    if reemplazado_nombre:
        resp_payload = dict(out)
        resp_payload["reemplazado"] = reemplazado_nombre
        return jsonify(resp_payload)
    return jsonify(out)


@app.route("/api/asignaciones/estado", methods=["POST"])
def api_cambiar_estado():
    data = request.get_json(silent=True) or {}
    operador_id = str(data.get("operador_id", "")).strip()
    nuevo_estado = str(data.get("estado", "")).strip()
    supervisor = get_actor_supervisor(data)

    if not is_valid_op_id(operador_id) or nuevo_estado not in ESTADOS_VALIDOS:
        return jsonify({"error": "operador_id y estado válido son requeridos."}), 400

    with db_lock:
        with get_db() as conn:
            sync_expired_vacations(conn)
            op = conn.execute("SELECT * FROM operadores WHERE id = ?", (operador_id,)).fetchone()
            if not op:
                return jsonify({"error": "Operador no encontrado."}), 404

            err, status = apply_operator_update(conn, op, data, supervisor)
            if err:
                return jsonify({"error": err}), status
            payload = build_state_payload(conn, sync_vacations=False)

    slim = publish_state_payload(payload)
    return jsonify(slim if wants_slim_response() else payload)


@app.route("/api/operadores", methods=["POST"])
def api_crear_operador():
    data = request.get_json(silent=True) or {}
    nombre = clean_cell(data.get("nombre", ""))[:60]
    columna_req = str(data.get("columna", "POOL")).strip().upper()
    estado = str(data.get("estado", "normal")).strip()
    puesto = normalize_puesto_label(clean_cell(data.get("puesto", ""))[:80])
    grupo = clean_cell(data.get("grupo", ""))[:40]
    cambio_grupo = clean_cell(data.get("cambio_grupo", ""))[:40] if estado == "cambio_turno" else ""
    tipo_personal = infer_tipo_personal(puesto, columna_req, data.get("tipo_personal"))
    lineas_cargo = sanitize_lineas_cargo(data.get("lineas_cargo"))
    vac_ini = str(data.get("vacaciones_inicio", "") or "").strip()
    vac_fin = str(data.get("vacaciones_fin", "") or "").strip()
    nota_estado = clean_cell(data.get("nota_estado", ""))[:200]
    foto_url = str(data.get("foto_url", "") or "").strip()
    posicion_raw = data.get("posicion")
    supervisor = get_actor_supervisor(data)

    if not nombre:
        return jsonify({"error": "El nombre del operador es obligatorio."}), 400
    ok_img, msg_img = validar_imagen_data_url(foto_url)
    if not ok_img:
        return jsonify({"error": msg_img or "Imagen inválida."}), 400
    if columna_req not in COLUMNAS_VALIDAS:
        return jsonify({"error": "Columna inválida."}), 400
    max_pos_limit = 10000 if columna_req in ("RELEVO", "SUPERVISION") else 200
    if posicion_raw is not None and (not is_strict_int(posicion_raw) or posicion_raw < 0 or posicion_raw > max_pos_limit):
        return jsonify({"error": "Posición de casilla inválida."}), 400
    if estado not in ESTADOS_VALIDOS:
        return jsonify({"error": "Estado inválido."}), 400

    if estado == "vacaciones":
        ok_i, msg_i = validar_fecha_logica(vac_ini)
        ok_f, msg_f = validar_fecha_logica(vac_fin)
        if not ok_i:
            return jsonify({"error": msg_i}), 400
        if not ok_f:
            return jsonify({"error": msg_f}), 400
        ini_d = parse_ymd(vac_ini)
        fin_d = parse_ymd(vac_fin)
        if ini_d and fin_d and fin_d < ini_d:
            return jsonify({"error": "La fecha final de vacaciones debe ser igual o posterior a la inicial."}), 400
        if fin_d and plant_now().date() > fin_d:
            return jsonify({"error": "La fecha de fin de vacaciones ya pasó."}), 400
        if ini_d and ini_d > plant_now().date():
            estado = "normal"
    else:
        vac_ini = ""
        vac_fin = ""

    if estado in ESTADOS_INACTIVOS:
        columna_req = "POOL"
    if columna_req == "RELEVO" and estado == "normal":
        estado = "relevo"

    nuevo_id = uuid.uuid4().hex[:10]

    with db_lock:
        with get_db() as conn:
            norm_new = normalize_text(nombre)
            existing_rows = conn.execute("SELECT id, nombre FROM operadores").fetchall()
            for er in existing_rows:
                if normalize_text(er["nombre"]) == norm_new:
                    return jsonify({"error": f"Ya existe un trabajador registrado como '{er['nombre']}'."}), 409

            columna = columna_req
            if columna in ("POOL", "SUPERVISION"):
                row = conn.execute(
                    "SELECT COALESCE(MAX(posicion), -1) AS max_pos FROM operadores WHERE columna = ?",
                    (columna,),
                ).fetchone()
                next_pos = int(row["max_pos"]) + 1
                col_origen = ""
            else:
                eff_new_group = cambio_grupo if (estado == "cambio_turno" and cambio_grupo) else grupo
                if columna in MAX_POR_COLUMNA:
                    cap_err = check_column_capacity(conn, columna, target_group=eff_new_group)
                    if cap_err:
                        return jsonify({"error": cap_err}), 400
                if is_strict_int(posicion_raw) and posicion_raw >= 0:
                    slot_idx = posicion_raw
                else:
                    slot_idx = find_free_slot(conn, columna, puesto, target_group=eff_new_group)
                    if slot_idx is None:
                        return jsonify({"error": f"Todas las casillas de {columna} están ocupadas."}), 400
                columna, next_pos = canonical_slot_target(columna, slot_idx)
                col_origen = columna_req

            conn.execute(
                """
                INSERT INTO operadores (
                    id, nombre, columna, estado, puesto, grupo, supervisor_grupo, numero, posicion,
                    vacaciones_inicio, vacaciones_fin, columna_origen, nota_estado, foto_url, cambio_grupo,
                    tipo_personal, lineas_cargo, updated_at, updated_by
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    nuevo_id,
                    nombre,
                    columna,
                    estado,
                    puesto,
                    grupo,
                    "",
                    "",
                    next_pos,
                    vac_ini,
                    vac_fin,
                    col_origen,
                    nota_estado,
                    foto_url,
                    cambio_grupo,
                    tipo_personal,
                    lineas_cargo,
                    now_iso(),
                    supervisor,
                ),
            )
            destino_label = "Nómina Disponible" if columna == "POOL" else columna
            log_event(conn, f"Nuevo operador {nombre} agregado en {destino_label}", supervisor, "crear")
            payload = build_state_payload(conn)

    slim = publish_state_payload(payload)
    return jsonify(slim if wants_slim_response() else payload), 201


@app.route("/api/operadores/<operador_id>", methods=["PATCH"])
def api_editar_operador(operador_id: str):
    if not is_valid_op_id(operador_id):
        return jsonify({"error": "ID de operador inválido."}), 400
    data = request.get_json(silent=True) or {}
    supervisor = get_actor_supervisor(data)

    with db_lock:
        with get_db() as conn:
            sync_expired_vacations(conn)
            op = conn.execute("SELECT * FROM operadores WHERE id = ?", (operador_id,)).fetchone()
            if not op:
                return jsonify({"error": "Operador no encontrado."}), 404

            err, status = apply_operator_update(conn, op, data, supervisor)
            if err:
                return jsonify({"error": err}), status
            payload = build_state_payload(conn, sync_vacations=False)

    slim = publish_state_payload(payload)
    return jsonify(slim if wants_slim_response() else payload)


@app.route("/api/operadores/<operador_id>", methods=["DELETE"])
def api_eliminar_operador(operador_id: str):
    if not is_valid_op_id(operador_id):
        return jsonify({"error": "ID de operador inválido."}), 400
    data = request.get_json(silent=True) or {}
    supervisor = get_actor_supervisor(data)

    with db_lock:
        with get_db() as conn:
            op = conn.execute("SELECT id, nombre, columna FROM operadores WHERE id = ?", (operador_id,)).fetchone()
            if not op:
                return jsonify({"error": "Operador no encontrado."}), 404

            conn.execute("DELETE FROM operadores WHERE id = ?", (operador_id,))
            with _photo_cache_lock:
                _photo_cache.pop(operador_id, None)
            reindex_column(conn, op["columna"])
            log_event(conn, f"Operador {op['nombre']} retirado del registro ({op['columna']})", supervisor, "eliminar")
            payload = build_state_payload(conn)

    slim = publish_state_payload(payload)
    return jsonify(slim if wants_slim_response() else payload)


@app.route("/api/tablero/limpiar_lineas", methods=["POST"])
def api_limpiar_lineas():
    """Devuelve el personal de las líneas a la Nómina Disponible (POOL); si se indica grupo, solo ese grupo."""
    data = request.get_json(silent=True) or {}
    supervisor = get_actor_supervisor(data)
    grupo_filtro = clean_cell(data.get("grupo", ""))[:40]
    if grupo_filtro.upper() == "ALL":
        grupo_filtro = ""

    with db_lock:
        with get_db() as conn:
            ts = now_iso()
            if grupo_filtro:
                rows_lin = conn.execute(
                    """
                    SELECT id, grupo, estado, cambio_grupo
                    FROM operadores
                    WHERE columna IN ('L1','L2','L3','L4','RELEVO','SIAT','MONTACARGAS')
                    """
                ).fetchall()
                target_ids = [
                    r["id"] for r in rows_lin if effective_op_group(r) == grupo_filtro
                ]
                if target_ids:
                    conn.executemany(
                        """
                        UPDATE operadores
                        SET columna_origen = columna,
                            columna = 'POOL',
                            estado = CASE WHEN estado = 'relevo' THEN 'normal' ELSE estado END,
                            updated_at = ?,
                            updated_by = ?
                        WHERE id = ?
                        """,
                        [(ts, supervisor, oid) for oid in target_ids],
                    )
                reindex_column(conn, "POOL")
                log_event(conn, f"Líneas vaciadas en {grupo_filtro} (personal devuelto a Nómina Disponible).", supervisor, "sistema")
            else:
                conn.execute(
                    """
                    UPDATE operadores
                    SET columna_origen = CASE
                            WHEN columna IN ('L1','L2','L3','L4','RELEVO','SIAT','MONTACARGAS') THEN columna
                            ELSE columna_origen
                        END,
                        columna = 'POOL',
                        estado = CASE WHEN estado = 'relevo' THEN 'normal' ELSE estado END,
                        updated_at = ?,
                        updated_by = ?
                    WHERE columna IN ('L1','L2','L3','L4','RELEVO','SIAT','MONTACARGAS')
                    """,
                    (ts, supervisor),
                )
                reindex_column(conn, "POOL")
                log_event(conn, "Líneas vaciadas (todo el personal quedó en Nómina Disponible).", supervisor, "sistema")
            payload = build_state_payload(conn)

    slim = publish_state_payload(payload)
    return jsonify(slim if wants_slim_response() else payload)


@app.route("/api/tablero/reset", methods=["POST"])
def api_reset_tablero():
    data = request.get_json(silent=True) or {}
    supervisor = get_actor_supervisor(data)
    destino_carga = str(data.get("destino_carga", "pool")).strip().lower()
    if destino_carga not in ("pool", "lineas"):
        destino_carga = "pool"

    with db_lock:
        with get_db() as conn:
            imp = conn.execute("SELECT titulo, payload_json FROM ultimo_import WHERE id = 1").fetchone()
            conn.execute("DELETE FROM operadores WHERE tipo_personal != 'supervision'")
            if imp:
                lista = json.loads(imp["payload_json"])
                pos_counters = {c: 0 for c in COLUMNAS_VALIDAS}
                max_pool_sup = conn.execute(
                    "SELECT COALESCE(MAX(posicion), -1) AS m FROM operadores WHERE columna = 'POOL'"
                ).fetchone()
                pos_counters["POOL"] = int(max_pool_sup["m"]) + 1 if max_pool_sup else 0
                ts = now_iso()
                insert_rows: List[Tuple[Any, ...]] = []
                for op in lista:
                    orig_col = op.get("columna", "L3")
                    if orig_col not in COLUMNAS_TABLERO:
                        orig_col = "L3"
                    if destino_carga == "lineas":
                        col = orig_col
                        max_cap = MAX_POR_COLUMNA.get(col)
                        if max_cap is not None and pos_counters[col] >= max_cap:
                            col = "POOL"
                    else:
                        col = "POOL"
                    estado = op.get("estado", "normal")
                    if col == "POOL" and estado == "relevo":
                        estado = "normal"
                    elif col == "RELEVO" and estado == "normal":
                        estado = "relevo"
                    insert_rows.append((
                        uuid.uuid4().hex[:10],
                        op["nombre"],
                        col,
                        orig_col,
                        estado,
                        op.get("puesto", ""),
                        op.get("grupo", ""),
                        op.get("supervisor_grupo", ""),
                        op.get("numero", ""),
                        pos_counters[col],
                        ts,
                        supervisor,
                    ))
                    pos_counters[col] += 1
                if insert_rows:
                    conn.executemany(
                        """
                        INSERT INTO operadores (
                            id, nombre, columna, columna_origen, estado, puesto, grupo, supervisor_grupo, numero, posicion, updated_at, updated_by
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        insert_rows,
                    )
                log_event(
                    conn,
                    f"Nómina restaurada desde la última carga ({imp['titulo'] or 'Excel'}).",
                    supervisor,
                    "sistema",
                )
            else:
                log_event(conn, "Tablero limpiado (sin archivo base cargado).", supervisor, "sistema")
            payload = build_state_payload(conn)

    publish_state_payload(payload)
    return jsonify(payload)


@app.route("/api/tablero/vaciar", methods=["POST"])
def api_vaciar_tablero():
    data = request.get_json(silent=True) or {}
    supervisor = get_actor_supervisor(data)

    with db_lock:
        with get_db() as conn:
            conn.execute("DELETE FROM operadores")
            with _photo_cache_lock:
                _photo_cache.clear()
            log_event(conn, "Nómina y tablero vaciados completamente.", supervisor, "sistema")
            payload = build_state_payload(conn)

    publish_state_payload(payload)
    return jsonify(payload)


@app.route("/api/importar/preview", methods=["POST"])
def api_importar_preview():
    periodos: List[Dict[str, Any]] = []

    if "file" in request.files and request.files["file"].filename:
        file_obj = request.files["file"]
        filename_orig = file_obj.filename or ""
        filename = filename_orig.lower()
        raw_bytes = file_obj.read()

        ok_file, err_file = validar_archivo_importacion(filename_orig, raw_bytes)
        if not ok_file:
            return jsonify({"error": err_file or "Archivo inválido."}), 400

        if filename.endswith((".xlsx", ".xlsm")):
            try:
                wb = load_workbook(io.BytesIO(raw_bytes), data_only=True, read_only=True)
                for sheet in wb.worksheets[:20]:
                    matrix = [[clean_cell(c) for c in row] for row in sheet.iter_rows(max_row=2000, max_col=60, values_only=True)]
                    periodos.extend(parse_rows_smart(matrix, sheet.title))
                wb.close()
            except Exception as exc:
                return jsonify({"error": f"No se pudo leer el archivo Excel: {exc}"}), 400
        else:
            text = decode_text_bytes(raw_bytes)
            matrix = parse_delimited_text(text)
            periodos.extend(parse_rows_smart(matrix, os.path.basename(filename_orig)))
    else:
        data = request.get_json(silent=True) or {}
        raw_text = str(data.get("raw_text", ""))
        if not raw_text.strip():
            return jsonify({"error": "Envía un archivo (.xlsx, .csv) o pega las celdas del Excel."}), 400
        if len(raw_text) > 1000000:
            return jsonify({"error": "El texto pegado excede el tamaño permitido."}), 400
        matrix = parse_delimited_text(raw_text)
        periodos.extend(parse_rows_smart(matrix, "Texto pegado"))

    if not periodos:
        return jsonify(
            {
                "error": "No se detectaron operadores en el archivo. Verifica que contenga bloques 'Rotacion por supervisor' o columnas 'Nombre' y 'Linea'."
            }
        ), 400

    return jsonify(
        {
            "periodos": periodos,
            "periodo_default": choose_default_period_id(periodos),
        }
    )


def _execute_import_job(
    clean_ops: List[Dict[str, str]],
    modo: str,
    destino_carga: str,
    titulo: str,
    supervisor: str,
) -> Dict[str, Any]:
    ts = now_iso()
    with db_lock:
        with get_db() as conn:
            pos_counters = {c: 0 for c in COLUMNAS_VALIDAS}
            occupied_by_col: Dict[str, Set[int]] = {c: set() for c in COLUMNAS_TABLERO}
            existing_by_name: Dict[str, str] = {}

            if modo == "reemplazar":
                conn.execute("DELETE FROM operadores WHERE tipo_personal != 'supervision'")
                max_pool_sup = conn.execute(
                    "SELECT COALESCE(MAX(posicion), -1) AS m FROM operadores WHERE columna = 'POOL'"
                ).fetchone()
                pos_counters["POOL"] = int(max_pool_sup["m"]) + 1 if max_pool_sup else 0
                conn.execute(
                    """
                    INSERT INTO ultimo_import (id, titulo, payload_json, saved_at, saved_by)
                    VALUES (1, ?, ?, ?, ?)
                    ON CONFLICT(id) DO UPDATE SET
                        titulo = excluded.titulo,
                        payload_json = excluded.payload_json,
                        saved_at = excluded.saved_at,
                        saved_by = excluded.saved_by
                    """,
                    (titulo, json.dumps(clean_ops, ensure_ascii=False), ts, supervisor),
                )
            else:
                deduplicate_operadores(conn)
                cur_rows = conn.execute("SELECT id, nombre, columna, posicion FROM operadores").fetchall()
                for ex in cur_rows:
                    nk = normalize_text(ex["nombre"])
                    if nk and nk not in existing_by_name:
                        existing_by_name[nk] = ex["id"]
                    col_ex = ex["columna"]
                    pos_ex = int(ex["posicion"] or 0)
                    if col_ex in occupied_by_col:
                        occupied_by_col[col_ex].add(pos_ex)
                    if col_ex == "POOL":
                        pos_counters["POOL"] = max(pos_counters["POOL"], pos_ex + 1)

            def _alloc_slot_in_mem(dest: str, op_puesto: str) -> Optional[int]:
                max_cap = MAX_POR_COLUMNA.get(dest)
                if max_cap is not None and len(occupied_by_col[dest]) >= max_cap:
                    return None
                occ = set(occupied_by_col[dest])
                if dest == "L2" and 0 in occupied_by_col["L1"]:
                    occ.add(0)
                slots = SLOTS_DEFINICION.get(dest, [])
                for s in slots:
                    idx = int(s["idx"])
                    if idx not in occ and slot_puesto_matches(s["puesto"], op_puesto):
                        return idx
                for s in slots:
                    idx = int(s["idx"])
                    if idx not in occ:
                        return idx
                if dest == "RELEVO":
                    idx = 0
                    while idx in occ:
                        idx += 1
                    return idx
                return None

            update_rows: List[Tuple[Any, ...]] = []
            insert_rows: List[Tuple[Any, ...]] = []

            for op in clean_ops:
                nk = normalize_text(op["nombre"])
                if modo != "reemplazar" and nk in existing_by_name:
                    update_rows.append((
                        op["puesto"],
                        op["grupo"],
                        op["supervisor_grupo"],
                        op["numero"],
                        op["columna"],
                        ts,
                        supervisor,
                        existing_by_name[nk],
                    ))
                    continue

                orig_col = op["columna"]
                if destino_carga == "lineas":
                    free_idx = _alloc_slot_in_mem(orig_col, op["puesto"])
                    if free_idx is not None:
                        c, slot_pos = canonical_slot_target(orig_col, free_idx)
                        occupied_by_col[c].add(slot_pos)
                    else:
                        c = "POOL"
                        slot_pos = pos_counters["POOL"]
                        pos_counters["POOL"] += 1
                else:
                    c = "POOL"
                    slot_pos = pos_counters["POOL"]
                    pos_counters["POOL"] += 1

                op_estado = "normal" if (c == "POOL" and op["estado"] == "relevo") else op["estado"]
                insert_rows.append((
                    uuid.uuid4().hex[:10],
                    op["nombre"],
                    c,
                    orig_col,
                    op_estado,
                    op["puesto"],
                    op["grupo"],
                    op["supervisor_grupo"],
                    op["numero"],
                    slot_pos,
                    ts,
                    supervisor,
                ))

            if update_rows:
                conn.executemany(
                    """
                    UPDATE operadores
                    SET puesto = ?, grupo = ?, supervisor_grupo = ?, numero = ?,
                        columna_origen = CASE WHEN columna = 'POOL' THEN ? ELSE columna_origen END,
                        updated_at = ?, updated_by = ?
                    WHERE id = ?
                    """,
                    update_rows,
                )
            if insert_rows:
                conn.executemany(
                    """
                    INSERT INTO operadores (
                        id, nombre, columna, columna_origen, estado, puesto, grupo, supervisor_grupo, numero, posicion, updated_at, updated_by
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    insert_rows,
                )

            inserted_count = len(insert_rows)
            updated_count = len(update_rows)
            accion = "cargados a Nómina Disponible" if destino_carga == "pool" else (
                "cargados (reemplazo)" if modo == "reemplazar" else f"procesados ({inserted_count} nuevos, {updated_count} actualizados sin duplicar)"
            )
            log_event(
                conn,
                f"Excel '{titulo}': {len(clean_ops)} operadores {accion}.",
                supervisor,
                "sistema",
            )
            payload = build_state_payload(conn)

    publish_state_payload(payload)
    return payload


@app.route("/api/importar/aplicar", methods=["POST"])
def api_importar_aplicar():
    data = request.get_json(silent=True) or {}
    operadores_in = data.get("operadores") or []
    modo = str(data.get("modo", "reemplazar")).strip().lower()
    if modo not in ("reemplazar", "agregar"):
        return jsonify({"error": "Modo de importación inválido."}), 400
    destino_carga = str(data.get("destino_carga", "pool")).strip().lower()
    if destino_carga not in ("pool", "lineas"):
        return jsonify({"error": "Destino de carga inválido."}), 400
    titulo = (clean_cell(data.get("titulo", "Archivo Excel"))[:120] or "Archivo Excel")
    supervisor = get_actor_supervisor(data)

    if not isinstance(operadores_in, list) or len(operadores_in) == 0:
        return jsonify({"error": "La lista de operadores a cargar está vacía."}), 400

    clean_ops: List[Dict[str, str]] = []
    seen_import_names: set = set()
    for item in operadores_in:
        if not isinstance(item, dict):
            continue
        nombre = clean_cell(item.get("nombre"))[:60]
        if not nombre:
            continue
        norm_key = normalize_text(nombre)
        if not norm_key or norm_key in seen_import_names:
            continue
        seen_import_names.add(norm_key)
        col = str(item.get("columna", "L3")).strip().upper()
        if col not in COLUMNAS_TABLERO:
            col = "L3"
        estado = str(item.get("estado", "normal")).strip()
        if estado not in ESTADOS_VALIDOS:
            estado = "relevo" if col == "RELEVO" else "normal"
        if col == "RELEVO" and estado == "normal":
            estado = "relevo"
        clean_ops.append(
            {
                "nombre": nombre,
                "columna": col,
                "estado": estado,
                "puesto": normalize_puesto_label(clean_cell(item.get("puesto"))[:80]),
                "grupo": clean_cell(item.get("grupo"))[:40],
                "supervisor_grupo": clean_cell(item.get("supervisor_grupo"))[:60],
                "numero": clean_cell(item.get("numero"))[:20],
            }
        )

    if not clean_ops:
        return jsonify({"error": "No quedaron operadores válidos para importar."}), 400

    run_async = request.args.get("async") == "1" or bool(data.get("async"))
    if run_async:
        owner = str((session.get("user") or {}).get("username") or "").strip().lower()

        def _bg_run():
            res_payload = _execute_import_job(clean_ops, modo, destino_carga, titulo, supervisor)
            return {"imported": len(clean_ops), "total_nomina": res_payload.get("resumen", {}).get("total_nomina", 0)}

        job_id = bg_jobs.enqueue("importar_aplicar", owner, _bg_run)
        return jsonify({"status": "queued", "job_id": job_id}), 202

    payload = _execute_import_job(clean_ops, modo, destino_carga, titulo, supervisor)
    return jsonify(payload)


@app.route("/api/operadores/deduplicar", methods=["POST"])
def api_deduplicar_operadores():
    data = request.get_json(silent=True) or {}
    supervisor = get_actor_supervisor(data)
    with db_lock:
        with get_db() as conn:
            removed = deduplicate_operadores(conn)
            if removed > 0:
                log_event(conn, f"{removed} trabajador(es) duplicado(s) eliminados.", supervisor, "sistema")
            payload = build_state_payload(conn)
            payload["eliminados"] = removed
    publish_state_payload(payload)
    return jsonify(payload)


@app.route("/api/plantilla/excel", methods=["GET"])
def api_plantilla_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Plantilla Simple"
    ws.append(["Nombre", "Linea", "Puesto", "Grupo", "Estado"])
    ws.append(["Ejemplo Operador 1", "Linea 1", "Despaletizador", "Grupo A", "Normal"])
    ws.append(["Ejemplo Operador 2", "Linea 2", "Llenadora", "Grupo A", "Normal"])
    ws.append(["Ejemplo Operador 3", "Linea 3", "Horno", "Grupo B", "Normal"])
    ws.append(["Ejemplo Operador 4", "Linea 4", "SMI", "Grupo B", "Cambio de Turno"])
    ws.append(["Ejemplo Relevo 1", "Relevo", "Siat", "Grupo A", "Relevo"])
    for col_letter, w in (("A", 26), ("B", 14), ("C", 22), ("D", 14), ("E", 18)):
        ws.column_dimensions[col_letter].width = w
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    raw = out.getvalue()
    etag = compute_etag(raw, prefix="tpl")
    if etag_matches(request.headers.get("If-None-Match"), etag):
        resp = Response(status=304)
        resp.headers["ETag"] = etag
        resp.headers["Cache-Control"] = "private, max-age=3600"
        return resp
    resp = send_file(
        io.BytesIO(raw),
        as_attachment=True,
        download_name="Plantilla_Rotacion_Envasado.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    resp.headers["ETag"] = etag
    resp.headers["Cache-Control"] = "private, max-age=3600"
    return resp


@app.route("/api/stream", methods=["GET"])
def api_stream():
    skip_init = request.args.get("skip_init") == "1"
    client_queue: queue.Queue = queue.Queue(maxsize=32)
    with sse_lock:
        while len(sse_clients) >= MAX_SSE_CLIENTS:
            old_q = sse_clients.pop(0)
            try:
                old_q.put_nowait(None)
            except queue.Full:
                pass
        sse_clients.append(client_queue)

    def event_generator():
        try:
            if not skip_init:
                initial, _etag = get_cached_state_payload()
                slim_init = _slim_state_for_broadcast(initial)
                yield f"event: tablero_actualizado\ndata: {json.dumps(slim_init, ensure_ascii=False, separators=(',', ':'))}\n\n"
            else:
                yield ": connected\n\n"

            while True:
                try:
                    msg = client_queue.get(timeout=20)
                    if msg is None:
                        break
                    yield msg
                except queue.Empty:
                    yield ": keepalive\n\n"
        except GeneratorExit:
            pass
        finally:
            with sse_lock:
                if client_queue in sse_clients:
                    sse_clients.remove(client_queue)

    headers = {
        "Cache-Control": "no-cache",
        "Connection": "keep-alive",
        "X-Accel-Buffering": "no",
    }
    return Response(event_generator(), mimetype="text/event-stream", headers=headers)


def pick_available_port(preferred_ports: Optional[List[int]] = None) -> int:
    env_port = os.environ.get("PORT")
    if env_port and env_port.isdigit():
        return int(env_port)
    candidates = preferred_ports or [5000, 5050, 8080, 8000]
    for port in candidates:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            sock.bind(("0.0.0.0", port))
            sock.close()
            return port
        except OSError:
            sock.close()
    return 5050


# Inicializar esquema y bootstrap de entorno también bajo Gunicorn/WSGI
try:
    init_db()
except Exception as exc:
    print("[DB INIT ADVERTENCIA]", exc)


if __name__ == "__main__":
    ACTIVE_PORT = pick_available_port()
    local_ip = get_local_ip()
    db_desc = "PostgreSQL (DATABASE_URL)" if is_postgres() else f"SQLite local ({DB_PATH})"

    print("=" * 64)
    print("  PIZARRA DIGITAL DE ENVASADO — SERVIDOR EN TIEMPO REAL")
    print("=" * 64)
    print(f"  Local (Esta PC / TV):   http://localhost:{ACTIVE_PORT}")
    print(f"  Red Local (Supervis.):  http://{local_ip}:{ACTIVE_PORT}")
    print(f"  Base de datos:          {db_desc}")
    print("  Presiona Ctrl+C para detener el servidor.")
    print("=" * 64)

    if "--open" in sys.argv:
        threading.Timer(1.2, lambda: webbrowser.open(f"http://localhost:{ACTIVE_PORT}")).start()

    app.run(host="0.0.0.0", port=ACTIVE_PORT, debug=False, threaded=True, use_reloader=False)
