import csv
import gzip
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
from typing import Any, Dict, List, Optional, Tuple

from flask import Flask, Response, jsonify, request, send_file, send_from_directory, session
from openpyxl import Workbook, load_workbook

from db_compat import (
    BASE_DIR,
    DB_PATH,
    get_db,
    init_auth_tables,
    is_postgres,
)
from seguridad import SESION_MAX_S, init_seguridad

STATIC_DIR = os.path.join(BASE_DIR, "static")

COLUMNAS_TABLERO = ("L1", "L2", "L3", "L4", "RELEVO", "SIAT", "MONTACARGAS")
COLUMNAS_VALIDAS = ("L1", "L2", "L3", "L4", "RELEVO", "SIAT", "MONTACARGAS", "POOL")
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
    ],
}

# IDs heredados del demo inicial anterior (para migrar a tablero limpio sin datos fijos)
LEGACY_DEMO_IDS = {str(i) for i in range(1, 24)}

app = Flask(__name__, static_folder=STATIC_DIR, static_url_path="/static")
init_seguridad(app)

db_lock = threading.RLock()
sse_lock = threading.Lock()
sse_clients: List[queue.Queue] = []
ACTIVE_PORT = 5000


def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def today_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d")


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


def ensure_column(conn: Any, table: str, col_name: str, col_def: str) -> None:
    cols = conn.table_columns(table)
    if col_name not in cols:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {col_name} {col_def}")


_db_initialized = False


def ensure_db_ready() -> None:
    if not _db_initialized:
        init_db()


def init_db() -> None:
    global _db_initialized
    with db_lock:
        with get_db() as conn:
            init_auth_tables(conn)
            pk_auto = "SERIAL PRIMARY KEY" if conn.pg else "INTEGER PRIMARY KEY AUTOINCREMENT"
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
            ensure_column(conn, "operadores", "puesto", "TEXT NOT NULL DEFAULT ''")
            ensure_column(conn, "operadores", "grupo", "TEXT NOT NULL DEFAULT ''")
            ensure_column(conn, "operadores", "supervisor_grupo", "TEXT NOT NULL DEFAULT ''")
            ensure_column(conn, "operadores", "numero", "TEXT NOT NULL DEFAULT ''")
            ensure_column(conn, "operadores", "vacaciones_inicio", "TEXT NOT NULL DEFAULT ''")
            ensure_column(conn, "operadores", "vacaciones_fin", "TEXT NOT NULL DEFAULT ''")
            ensure_column(conn, "operadores", "columna_origen", "TEXT NOT NULL DEFAULT ''")
            ensure_column(conn, "operadores", "nota_estado", "TEXT NOT NULL DEFAULT ''")
            ensure_column(conn, "operadores", "foto_url", "TEXT NOT NULL DEFAULT ''")
            ensure_column(conn, "operadores", "prevenido_hasta", "TEXT NOT NULL DEFAULT ''")
            ensure_column(conn, "operadores", "prevenido_motivo", "TEXT NOT NULL DEFAULT ''")

            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS config_meta (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL DEFAULT ''
                )
                """
            )

            conn.execute(
                f"""
                CREATE TABLE IF NOT EXISTS historial (
                    id {pk_auto},
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
    _db_initialized = True


def get_actor_supervisor(data: Optional[Dict[str, Any]] = None, fallback: str = "Supervisor") -> str:
    """Prioriza siempre la identidad verificada en la sesión del servidor."""
    user = session.get("user") or {}
    session_name = clean_cell(user.get("nombre") or user.get("username"))
    if session_name:
        return session_name
    if isinstance(data, dict) and data.get("supervisor"):
        return clean_cell(data.get("supervisor")) or fallback
    return fallback


def log_event(conn: Any, descripcion: str, supervisor: str, tipo: str) -> None:
    conn.execute(
        """
        INSERT INTO historial (descripcion, supervisor, tipo, created_at)
        VALUES (?, ?, ?, ?)
        """,
        (descripcion, supervisor or "Supervisor", tipo, now_iso()),
    )
    conn.execute(
        """
        DELETE FROM historial
        WHERE id NOT IN (
            SELECT id FROM historial ORDER BY id DESC LIMIT 200
        )
        """
    )


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
        dias_total = (fin_d - ini_d).days + 1

    dias_restantes = 0
    if en_vacaciones_hoy and fin_d and fin_d >= today_d:
        dias_restantes = (fin_d - today_d).days + 1

    fecha_reactivacion = ""
    if fin_d:
        fecha_reactivacion = (fin_d + timedelta(days=1)).strftime("%Y-%m-%d")

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
    today_d = datetime.now().date()
    today_str = today_d.strftime("%Y-%m-%d")
    changed = False
    touched_cols = set()

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

        # Día siguiente o posterior al fin de vacaciones -> reactivar automáticamente
        if fin_d and today_d > fin_d:
            conn.execute(
                """
                UPDATE operadores
                SET estado = 'normal', vacaciones_inicio = '', vacaciones_fin = '', nota_estado = '', updated_at = ?
                WHERE id = ?
                """,
                (now_iso(), r["id"]),
            )
            log_event(
                conn,
                f"🏖️ Fin de vacaciones: {r['nombre']} se reactivó automáticamente.",
                "Calendario",
                "estado",
            )
            changed = True
            continue

        # Vacaciones vigentes hoy -> asegurarse de estar inactivo y fuera de línea activa
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
            conn.execute(
                """
                UPDATE operadores
                SET estado = 'vacaciones', columna = 'POOL', columna_origen = ?, updated_at = ?
                WHERE id = ?
                """,
                (col_origen, now_iso(), r["id"]),
            )
            if r["estado"] != "vacaciones":
                log_event(
                    conn,
                    f"🏖️ Inicio de vacaciones: {r['nombre']} desactivado hasta {r['vacaciones_fin'] or 'nuevo aviso'}.",
                    "Calendario",
                    "estado",
                )
            changed = True

    for col in touched_cols:
        reindex_column(conn, col)

    # Limpiar estados temporales de 'Prevenido' cuyo tiempo ya venció
    now_str = now_iso()
    prev_rows = conn.execute(
        """
        SELECT id, nombre, prevenido_hasta, prevenido_motivo
        FROM operadores
        WHERE prevenido_hasta != ''
        """
    ).fetchall()
    for pr in prev_rows:
        if pr["prevenido_hasta"] <= now_str:
            conn.execute(
                """
                UPDATE operadores
                SET prevenido_hasta = '', prevenido_motivo = '', updated_at = ?
                WHERE id = ?
                """,
                (now_str, pr["id"]),
            )
            log_event(
                conn,
                f"⏱️ Tiempo de prevención cumplido: {pr['nombre']} ({pr['prevenido_motivo'] or 'Normalizado'}).",
                "Reloj",
                "estado",
            )
            changed = True
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


def find_free_slot(conn: Any, destino: str, op_puesto: str = "") -> Optional[int]:
    slots = SLOTS_DEFINICION.get(destino, [])
    rows = conn.execute("SELECT posicion FROM operadores WHERE columna = ?", (destino,)).fetchall()
    occupied = {int(r["posicion"]) for r in rows}

    # L2 slot 0 está ocupado si L1 slot 0 tiene Pela paletas
    if destino == "L2":
        pela = conn.execute("SELECT id FROM operadores WHERE columna = 'L1' AND posicion = 0").fetchone()
        if pela:
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

    # 3. RELEVO permite casillas adicionales más allá de las 4 iniciales
    if destino == "RELEVO":
        return (max(occupied) + 1) if occupied else len(slots)

    return None


def check_column_capacity(conn: Any, destino: str, exclude_op_id: Optional[str] = None) -> Optional[str]:
    max_cap = MAX_POR_COLUMNA.get(destino)
    if max_cap is None:
        return None
    if exclude_op_id:
        count_row = conn.execute(
            "SELECT COUNT(*) AS c FROM operadores WHERE columna = ? AND id != ?",
            (destino, exclude_op_id),
        ).fetchone()
    else:
        count_row = conn.execute(
            "SELECT COUNT(*) AS c FROM operadores WHERE columna = ?",
            (destino,),
        ).fetchone()
    current_count = int(count_row["c"]) if count_row else 0
    if current_count >= max_cap:
        area_label = "Siat" if destino == "SIAT" else "Montacargas"
        return f"El área de {area_label} solo permite máximo {max_cap} operadores ({current_count}/{max_cap})."
    return None


def build_state_payload(conn: Any) -> Dict[str, Any]:
    sync_expired_vacations(conn)
    now_dt = datetime.now()
    today_d = now_dt.date()
    today_str = today_d.strftime("%Y-%m-%d")

    tablero: Dict[str, List[Dict[str, Any]]] = {col: [] for col in COLUMNAS_TABLERO}
    pool: List[Dict[str, Any]] = []

    rows = conn.execute(
        """
        SELECT id, nombre, columna, estado, puesto, grupo, supervisor_grupo, numero, posicion,
               vacaciones_inicio, vacaciones_fin, columna_origen, nota_estado, foto_url,
               prevenido_hasta, prevenido_motivo, updated_at, updated_by
        FROM operadores
        ORDER BY columna ASC, posicion ASC, nombre ASC
        """
    ).fetchall()

    grupos_set = set()
    resumen = {
        "total_nomina": 0,
        "activos": 0,
        "en_lineas": 0,
        "disponibles_pool": 0,
        "vacaciones": 0,
        "ausentes": 0,
        "licencias": 0,
        "relevos_cambios": 0,
        "prevenidos": 0,
    }

    pela_paletas_op: Optional[Dict[str, Any]] = None

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

        op_dict = {
            "id": r["id"],
            "nombre": r["nombre"],
            "columna": col,
            "columna_origen": r["columna_origen"] or "",
            "estado": estado,
            "puesto": r["puesto"] or "",
            "grupo": grupo,
            "supervisor_grupo": r["supervisor_grupo"] or "",
            "numero": r["numero"] or "",
            "posicion": int(r["posicion"] or 0),
            "vacaciones_inicio": r["vacaciones_inicio"] or "",
            "vacaciones_fin": r["vacaciones_fin"] or "",
            "nota_estado": r["nota_estado"] or "",
            "foto_url": r["foto_url"] or "",
            "prevenido_activo": prev_activo,
            "prevenido_hasta": prev_hasta if prev_activo else "",
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
        else:
            resumen["en_lineas"] += 1
            if (col in ("L1", "L2")) and op_dict["posicion"] == 0:
                op_dict["compartido_l1_l2"] = True
                if pela_paletas_op is None or col == "L1":
                    pela_paletas_op = op_dict
            else:
                tablero[col].append(op_dict)

    # Insertar Pela paletas compartido simultáneamente en L1 (slot 0) y L2 (slot 0)
    if pela_paletas_op is not None:
        l1_copy = dict(pela_paletas_op, columna="L1", posicion=0, compartido_l1_l2=True)
        l2_copy = dict(pela_paletas_op, columna="L2", posicion=0, compartido_l1_l2=True)
        tablero["L1"].insert(0, l1_copy)
        tablero["L2"].insert(0, l2_copy)

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
        "pool": pool,
        "resumen": resumen,
        "historial": historial,
        "grupos": sorted(grupos_set),
        "ultimo_import": ultimo_import_info,
        "max_por_columna": MAX_POR_COLUMNA,
        "slots_definicion": SLOTS_DEFINICION,
        "server_time": now_iso(),
        "today": today_str,
    }


def broadcast_sse(event_name: str, payload: Dict[str, Any]) -> None:
    data = json.dumps(payload, ensure_ascii=False)
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
    """Compacta solo POOL para conservar intactos los índices fijos de casillas en L1..MONTACARGAS."""
    if columna != "POOL":
        return
    rows = conn.execute(
        "SELECT id FROM operadores WHERE columna = ? ORDER BY posicion ASC, updated_at ASC, nombre ASC",
        (columna,),
    ).fetchall()
    for idx, row in enumerate(rows):
        conn.execute("UPDATE operadores SET posicion = ? WHERE id = ?", (idx, row["id"]))


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
# Rutas HTTP y SSE
# ====================================================================
@app.after_request
def add_performance_cache_headers(response):
    if request.path.startswith("/static/") or request.path == "/favicon.ico":
        response.headers["Cache-Control"] = "public, max-age=604800"
    elif request.path == "/":
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"

    if (
        response.mimetype != "text/event-stream"
        and response.status_code == 200
        and "gzip" in request.headers.get("Accept-Encoding", "").lower()
        and "Content-Encoding" not in response.headers
        and any(
            response.mimetype.startswith(p)
            for p in ("text/", "application/json", "application/javascript", "image/svg+xml")
        )
    ):
        response.direct_passthrough = False
        raw = response.get_data()
        if len(raw) > 1024:
            compressed = gzip.compress(raw, compresslevel=5)
            response.set_data(compressed)
            response.headers["Content-Encoding"] = "gzip"
            response.headers["Content-Length"] = str(len(compressed))
            response.vary.add("Accept-Encoding")
    return response


@app.route("/favicon.ico")
def serve_favicon():
    return send_from_directory(STATIC_DIR, "favicon.ico")


@app.route("/")
def serve_index():
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/api/tablero", methods=["GET"])
def api_get_tablero():
    ensure_db_ready()
    with db_lock:
        with get_db() as conn:
            payload = build_state_payload(conn)
    payload["current_user"] = session.get("user")
    login_ts = session.get("login_ts", 0)
    payload["expira_en_s"] = max(0, int(SESION_MAX_S - (time.time() - login_ts))) if login_ts else 0
    return jsonify(payload)


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
    today_d = datetime.now().date()
    op_id = op_row["id"]

    nombre = clean_cell(data.get("nombre", op_row["nombre"])) or op_row["nombre"]
    puesto = clean_cell(data.get("puesto", op_row["puesto"])) if "puesto" in data else op_row["puesto"]
    if "puesto" in data:
        puesto = normalize_puesto_label(puesto)
    grupo = clean_cell(data.get("grupo", op_row["grupo"])) if "grupo" in data else op_row["grupo"]
    nota_estado = (
        clean_cell(data.get("nota_estado", op_row["nota_estado"]))
        if "nota_estado" in data
        else op_row["nota_estado"]
    )
    foto_url = op_row["foto_url"] if "foto_url" in op_row.keys() else ""
    if "foto_url" in data:
        raw_foto = str(data.get("foto_url") or "").strip()
        if len(raw_foto) > 450000:
            return "La imagen es demasiado grande. Usa una foto más liviana.", 400
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
                    try:
                        mins = int(raw_mins)
                    except (TypeError, ValueError):
                        return "Los minutos de prevención deben ser un número entero.", 400
                    if mins < 1 or mins > 720:
                        return "La duración temporal de prevención debe estar entre 1 y 720 minutos.", 400
                    prev_hasta = (datetime.now() + timedelta(minutes=mins)).strftime("%Y-%m-%d %H:%M:%S")
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
                        if parsed_dt <= datetime.now():
                            parsed_dt = parsed_dt + timedelta(days=1)
                    if parsed_dt is None or parsed_dt <= datetime.now():
                        return "La hora límite de prevención debe ser posterior a la hora actual.", 400
                    prev_hasta = parsed_dt.strftime("%Y-%m-%d %H:%M:%S")
                else:
                    prev_hasta = (datetime.now() + timedelta(minutes=30)).strftime("%Y-%m-%d %H:%M:%S")
                prev_motivo = clean_cell(data.get("prevenido_motivo")) or "Prevenido"
            else:
                prev_hasta = ""
                prev_motivo = ""

        if target_col != origen_col:
            cap_err = check_column_capacity(conn, target_col, exclude_op_id=op_id)
            if cap_err:
                return cap_err, 400
            if target_col == "RELEVO" and estado_final == "normal":
                estado_final = "relevo"

    new_pos = int(op_row["posicion"] or 0)
    if target_col != origen_col:
        if target_col == "POOL":
            max_row = conn.execute(
                "SELECT COALESCE(MAX(posicion), -1) AS m FROM operadores WHERE columna = 'POOL'"
            ).fetchone()
            new_pos = int(max_row["m"]) + 1
        else:
            free_idx = find_free_slot(conn, target_col, puesto)
            if free_idx is None:
                return f"Todas las casillas de {target_col} están ocupadas.", 400
            target_col, new_pos = canonical_slot_target(target_col, free_idx)

    conn.execute(
        """
        UPDATE operadores
        SET nombre = ?, puesto = ?, grupo = ?, estado = ?, columna = ?, columna_origen = ?,
            vacaciones_inicio = ?, vacaciones_fin = ?, nota_estado = ?, foto_url = ?,
            prevenido_hasta = ?, prevenido_motivo = ?, posicion = ?,
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
            new_pos,
            now_iso(),
            supervisor,
            op_id,
        ),
    )

    if origen_col == "POOL" and target_col != "POOL":
        reindex_column(conn, "POOL")

    prev_was = (op_row["prevenido_hasta"] if "prevenido_hasta" in op_row.keys() else "") or ""

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

    if not operador_id or destino_raw not in COLUMNAS_VALIDAS:
        return jsonify({"error": "operador_id y destino válido son requeridos."}), 400

    with db_lock:
        with get_db() as conn:
            sync_expired_vacations(conn)
            op = conn.execute(
                """
                SELECT id, nombre, columna, columna_origen, estado, puesto, grupo,
                       vacaciones_inicio, vacaciones_fin, nota_estado, posicion
                FROM operadores WHERE id = ?
                """,
                (operador_id,),
            ).fetchone()
            if not op:
                return jsonify({"error": "Operador no encontrado."}), 404

            metrics = op_vacation_metrics(op, datetime.now().date())
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
                payload = build_state_payload(conn)
                broadcast_sse("tablero_actualizado", payload)
                return jsonify(payload)

            # Destino en tablero (casilla fija)
            if isinstance(posicion_raw, int) and posicion_raw >= 0:
                max_cap = MAX_POR_COLUMNA.get(destino_raw)
                if max_cap is not None and posicion_raw >= max_cap:
                    area_label = "Siat" if destino_raw == "SIAT" else "Montacargas"
                    return jsonify({"error": f"El área de {area_label} solo tiene {max_cap} casillas (máximo {max_cap})."}), 400
                target_slot_idx = posicion_raw
            else:
                free_idx = find_free_slot(conn, destino_raw, op["puesto"])
                if free_idx is None:
                    cap_err = check_column_capacity(conn, destino_raw, exclude_op_id=operador_id)
                    return jsonify({"error": cap_err or f"Todas las casillas de {destino_raw} están ocupadas."}), 400
                target_slot_idx = free_idx

            dest_col, dest_slot = canonical_slot_target(destino_raw, target_slot_idx)

            # Sin cambio real (p. ej. arrastrar Pela paletas entre L1 slot 0 y L2 slot 0)
            if origen == dest_col and origen_pos == dest_slot:
                payload = build_state_payload(conn)
                return jsonify(payload)

            # Verificar si la casilla destino ya tiene ocupante
            occupant = conn.execute(
                "SELECT id, nombre, estado, columna, columna_origen FROM operadores WHERE columna = ? AND posicion = ? AND id != ?",
                (dest_col, dest_slot, operador_id),
            ).fetchone()

            if not occupant and dest_col != origen:
                cap_err = check_column_capacity(conn, dest_col, exclude_op_id=operador_id)
                if cap_err:
                    return jsonify({"error": cap_err}), 400

            if occupant:
                if origen in COLUMNAS_TABLERO:
                    # Intercambio limpio entre dos casillas del tablero
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
                    # El ocupante anterior vuelve a POOL
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

            if dest_col == "RELEVO":
                nuevo_estado = "relevo"
            elif origen == "RELEVO" and dest_col != "RELEVO" and op["estado"] == "relevo":
                nuevo_estado = "normal"
            else:
                nuevo_estado = op["estado"]

            conn.execute(
                """
                UPDATE operadores
                SET columna = ?, columna_origen = ?, estado = ?, posicion = ?, updated_at = ?, updated_by = ?
                WHERE id = ?
                """,
                (dest_col, col_origen_saved, nuevo_estado, dest_slot, timestamp, supervisor, operador_id),
            )

            if origen == "POOL":
                reindex_column(conn, "POOL")

            slot_defs = SLOTS_DEFINICION.get(destino_raw, [])
            slot_name = slot_defs[target_slot_idx]["puesto"] if target_slot_idx < len(slot_defs) else f"Casilla {target_slot_idx + 1}"
            if dest_col == "L1" and dest_slot == 0:
                log_event(conn, f"{op['nombre']} asignado a Pela paletas (compartido L1 y L2)", supervisor, "mover")
            elif occupant and origen in COLUMNAS_TABLERO:
                log_event(conn, f"Intercambio: {op['nombre']} ↔ {occupant['nombre']} ({destino_raw} · {slot_name})", supervisor, "mover")
            else:
                log_event(conn, f"{op['nombre']} ubicado en {destino_raw} · {slot_name}", supervisor, "mover")

            payload = build_state_payload(conn)

    broadcast_sse("tablero_actualizado", payload)
    return jsonify(payload)


@app.route("/api/asignaciones/estado", methods=["POST"])
def api_cambiar_estado():
    data = request.get_json(silent=True) or {}
    operador_id = str(data.get("operador_id", "")).strip()
    nuevo_estado = str(data.get("estado", "")).strip()
    supervisor = get_actor_supervisor(data)

    if not operador_id or nuevo_estado not in ESTADOS_VALIDOS:
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
            payload = build_state_payload(conn)

    broadcast_sse("tablero_actualizado", payload)
    return jsonify(payload)


@app.route("/api/operadores", methods=["POST"])
def api_crear_operador():
    data = request.get_json(silent=True) or {}
    nombre = clean_cell(data.get("nombre", ""))
    columna_req = str(data.get("columna", "POOL")).strip().upper()
    estado = str(data.get("estado", "normal")).strip()
    puesto = normalize_puesto_label(clean_cell(data.get("puesto", "")))
    grupo = clean_cell(data.get("grupo", ""))
    vac_ini = str(data.get("vacaciones_inicio", "") or "").strip()
    vac_fin = str(data.get("vacaciones_fin", "") or "").strip()
    nota_estado = clean_cell(data.get("nota_estado", ""))
    foto_url = str(data.get("foto_url", "") or "").strip()
    posicion_raw = data.get("posicion")
    supervisor = get_actor_supervisor(data)

    if not nombre:
        return jsonify({"error": "El nombre del operador es obligatorio."}), 400
    if len(foto_url) > 450000:
        return jsonify({"error": "La imagen es demasiado grande."}), 400
    if columna_req not in COLUMNAS_VALIDAS:
        return jsonify({"error": "Columna inválida."}), 400
    if estado not in ESTADOS_VALIDOS:
        estado = "relevo" if columna_req == "RELEVO" else "normal"
    if estado in ESTADOS_INACTIVOS:
        columna_req = "POOL"
    if columna_req == "RELEVO" and estado == "normal":
        estado = "relevo"

    nuevo_id = uuid.uuid4().hex[:10]

    with db_lock:
        with get_db() as conn:
            columna = columna_req
            if columna == "POOL":
                row = conn.execute(
                    "SELECT COALESCE(MAX(posicion), -1) AS max_pos FROM operadores WHERE columna = 'POOL'"
                ).fetchone()
                next_pos = int(row["max_pos"]) + 1
                col_origen = ""
            else:
                if columna in MAX_POR_COLUMNA:
                    cap_err = check_column_capacity(conn, columna)
                    if cap_err:
                        return jsonify({"error": cap_err}), 400
                if isinstance(posicion_raw, int) and posicion_raw >= 0:
                    slot_idx = posicion_raw
                else:
                    slot_idx = find_free_slot(conn, columna, puesto)
                    if slot_idx is None:
                        return jsonify({"error": f"Todas las casillas de {columna} están ocupadas."}), 400
                columna, next_pos = canonical_slot_target(columna, slot_idx)
                col_origen = columna_req

            conn.execute(
                """
                INSERT INTO operadores (
                    id, nombre, columna, estado, puesto, grupo, supervisor_grupo, numero, posicion,
                    vacaciones_inicio, vacaciones_fin, columna_origen, nota_estado, foto_url, updated_at, updated_by
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                    now_iso(),
                    supervisor,
                ),
            )
            destino_label = "Nómina Disponible" if columna == "POOL" else columna
            log_event(conn, f"Nuevo operador {nombre} agregado en {destino_label}", supervisor, "crear")
            payload = build_state_payload(conn)

    broadcast_sse("tablero_actualizado", payload)
    return jsonify(payload), 201


@app.route("/api/operadores/<operador_id>", methods=["PATCH"])
def api_editar_operador(operador_id: str):
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
            payload = build_state_payload(conn)

    broadcast_sse("tablero_actualizado", payload)
    return jsonify(payload)


@app.route("/api/operadores/<operador_id>", methods=["DELETE"])
def api_eliminar_operador(operador_id: str):
    data = request.get_json(silent=True) or {}
    supervisor = get_actor_supervisor(data)

    with db_lock:
        with get_db() as conn:
            op = conn.execute("SELECT id, nombre, columna FROM operadores WHERE id = ?", (operador_id,)).fetchone()
            if not op:
                return jsonify({"error": "Operador no encontrado."}), 404

            conn.execute("DELETE FROM operadores WHERE id = ?", (operador_id,))
            reindex_column(conn, op["columna"])
            log_event(conn, f"Operador {op['nombre']} retirado del registro ({op['columna']})", supervisor, "eliminar")
            payload = build_state_payload(conn)

    broadcast_sse("tablero_actualizado", payload)
    return jsonify(payload)


@app.route("/api/tablero/limpiar_lineas", methods=["POST"])
def api_limpiar_lineas():
    """Devuelve todo el personal de las líneas a la Nómina Disponible (POOL) dejando las cajas en blanco con '+'."""
    data = request.get_json(silent=True) or {}
    supervisor = get_actor_supervisor(data)

    with db_lock:
        with get_db() as conn:
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
                WHERE columna != 'POOL'
                """,
                (now_iso(), supervisor),
            )
            reindex_column(conn, "POOL")
            log_event(conn, "Líneas vaciadas (todo el personal quedó en Nómina Disponible).", supervisor, "sistema")
            payload = build_state_payload(conn)

    broadcast_sse("tablero_actualizado", payload)
    return jsonify(payload)


@app.route("/api/tablero/reset", methods=["POST"])
def api_reset_tablero():
    data = request.get_json(silent=True) or {}
    supervisor = get_actor_supervisor(data)
    destino_carga = str(data.get("destino_carga", "pool")).strip().lower()

    with db_lock:
        with get_db() as conn:
            imp = conn.execute("SELECT titulo, payload_json FROM ultimo_import WHERE id = 1").fetchone()
            conn.execute("DELETE FROM operadores")
            if imp:
                lista = json.loads(imp["payload_json"])
                pos_counters = {c: 0 for c in COLUMNAS_VALIDAS}
                ts = now_iso()
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
                    conn.execute(
                        """
                        INSERT INTO operadores (
                            id, nombre, columna, columna_origen, estado, puesto, grupo, supervisor_grupo, numero, posicion, updated_at, updated_by
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
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
                        ),
                    )
                    pos_counters[col] += 1
                log_event(
                    conn,
                    f"Nómina restaurada desde la última carga ({imp['titulo'] or 'Excel'}).",
                    supervisor,
                    "sistema",
                )
            else:
                log_event(conn, "Tablero limpiado (sin archivo base cargado).", supervisor, "sistema")
            payload = build_state_payload(conn)

    broadcast_sse("tablero_actualizado", payload)
    return jsonify(payload)


@app.route("/api/tablero/vaciar", methods=["POST"])
def api_vaciar_tablero():
    data = request.get_json(silent=True) or {}
    supervisor = get_actor_supervisor(data)

    with db_lock:
        with get_db() as conn:
            conn.execute("DELETE FROM operadores")
            log_event(conn, "Nómina y tablero vaciados completamente.", supervisor, "sistema")
            payload = build_state_payload(conn)

    broadcast_sse("tablero_actualizado", payload)
    return jsonify(payload)


@app.route("/api/importar/preview", methods=["POST"])
def api_importar_preview():
    periodos: List[Dict[str, Any]] = []

    if "file" in request.files and request.files["file"].filename:
        file_obj = request.files["file"]
        filename = file_obj.filename.lower()
        raw_bytes = file_obj.read()

        if filename.endswith((".xlsx", ".xlsm")):
            try:
                wb = load_workbook(io.BytesIO(raw_bytes), data_only=True, read_only=True)
                for sheet in wb.worksheets:
                    matrix = [[clean_cell(c) for c in row] for row in sheet.iter_rows(values_only=True)]
                    periodos.extend(parse_rows_smart(matrix, sheet.title))
                wb.close()
            except Exception as exc:
                return jsonify({"error": f"No se pudo leer el archivo Excel: {exc}"}), 400
        elif filename.endswith(".xls"):
            return jsonify(
                {
                    "error": "Formato .xls antiguo detectado. Guárdalo en Excel como .xlsx o .csv, o copia y pega las celdas en la pestaña 'Pegar desde Excel'."
                }
            ), 400
        else:
            text = decode_text_bytes(raw_bytes)
            matrix = parse_delimited_text(text)
            periodos.extend(parse_rows_smart(matrix, file_obj.filename))
    else:
        data = request.get_json(silent=True) or {}
        raw_text = str(data.get("raw_text", ""))
        if not raw_text.strip():
            return jsonify({"error": "Envía un archivo (.xlsx, .csv) o pega las celdas del Excel."}), 400
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


@app.route("/api/importar/aplicar", methods=["POST"])
def api_importar_aplicar():
    data = request.get_json(silent=True) or {}
    operadores_in = data.get("operadores") or []
    modo = str(data.get("modo", "reemplazar")).strip().lower()
    destino_carga = str(data.get("destino_carga", "pool")).strip().lower()
    titulo = str(data.get("titulo", "Archivo Excel")).strip() or "Archivo Excel"
    supervisor = get_actor_supervisor(data)

    if not isinstance(operadores_in, list) or len(operadores_in) == 0:
        return jsonify({"error": "La lista de operadores a cargar está vacía."}), 400

    clean_ops: List[Dict[str, str]] = []
    for item in operadores_in:
        nombre = clean_cell(item.get("nombre"))
        if not nombre:
            continue
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
                "puesto": normalize_puesto_label(clean_cell(item.get("puesto"))),
                "grupo": clean_cell(item.get("grupo")),
                "supervisor_grupo": clean_cell(item.get("supervisor_grupo")),
                "numero": clean_cell(item.get("numero")),
            }
        )

    if not clean_ops:
        return jsonify({"error": "No quedaron operadores válidos para importar."}), 400

    ts = now_iso()
    with db_lock:
        with get_db() as conn:
            pos_counters = {c: 0 for c in COLUMNAS_VALIDAS}
            if modo == "reemplazar":
                conn.execute("DELETE FROM operadores")
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
                for c in COLUMNAS_VALIDAS:
                    r = conn.execute(
                        "SELECT COALESCE(MAX(posicion), -1) AS m FROM operadores WHERE columna = ?", (c,)
                    ).fetchone()
                    pos_counters[c] = int(r["m"]) + 1

            for op in clean_ops:
                orig_col = op["columna"]
                if destino_carga == "lineas":
                    max_cap = MAX_POR_COLUMNA.get(orig_col)
                    if max_cap is not None:
                        in_col = conn.execute(
                            "SELECT COUNT(*) AS c FROM operadores WHERE columna = ?", (orig_col,)
                        ).fetchone()["c"]
                        free_idx = None if in_col >= max_cap else find_free_slot(conn, orig_col, op["puesto"])
                    else:
                        free_idx = find_free_slot(conn, orig_col, op["puesto"])

                    if free_idx is not None:
                        c, slot_pos = canonical_slot_target(orig_col, free_idx)
                    else:
                        c = "POOL"
                        slot_pos = pos_counters["POOL"]
                        pos_counters["POOL"] += 1
                else:
                    c = "POOL"
                    slot_pos = pos_counters["POOL"]
                    pos_counters["POOL"] += 1

                op_estado = "normal" if (c == "POOL" and op["estado"] == "relevo") else op["estado"]
                conn.execute(
                    """
                    INSERT INTO operadores (
                        id, nombre, columna, columna_origen, estado, puesto, grupo, supervisor_grupo, numero, posicion, updated_at, updated_by
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
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
                    ),
                )

            accion = "cargados a Nómina Disponible" if destino_carga == "pool" else (
                "cargados (reemplazo)" if modo == "reemplazar" else "agregados"
            )
            log_event(
                conn,
                f"Excel '{titulo}': {len(clean_ops)} operadores {accion}.",
                supervisor,
                "sistema",
            )
            payload = build_state_payload(conn)

    broadcast_sse("tablero_actualizado", payload)
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
    return send_file(
        out,
        as_attachment=True,
        download_name="Plantilla_Rotacion_Envasado.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/stream", methods=["GET"])
def api_stream():
    skip_init = request.args.get("skip_init") == "1"
    client_queue: queue.Queue = queue.Queue(maxsize=32)
    with sse_lock:
        sse_clients.append(client_queue)

    def event_generator():
        try:
            if not skip_init:
                with db_lock:
                    with get_db() as conn:
                        initial = build_state_payload(conn)
                yield f"event: tablero_actualizado\ndata: {json.dumps(initial, ensure_ascii=False)}\n\n"
            else:
                yield ": connected\n\n"

            while True:
                try:
                    msg = client_queue.get(timeout=20)
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
