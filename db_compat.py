"""
db_compat.py — Adaptador unificado SQLite (con pool de conexiones) / PostgreSQL (ThreadedConnectionPool),
gestión segura de usuarios (scrypt), cola de trabajos en segundo plano y auditoría sin DDL en rutas calientes.
"""

import atexit
import os
import queue
import sqlite3
import threading
import time
import uuid
from collections import OrderedDict
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from typing import Any, Callable, Dict, List, Optional, Sequence, Set, Tuple

from dotenv import load_dotenv
from werkzeug.security import check_password_hash, generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"), override=False)

_pg_pool = None
_pg_pool_url = ""
_pg_pool_lock = threading.Lock()

_sqlite_pools: Dict[str, "_SQLitePool"] = {}
_sqlite_pools_lock = threading.Lock()

_auth_tables_ready: Set[str] = set()
_auth_init_lock = threading.Lock()

_audit_queue: "queue.Queue[Tuple[str, str, str, str, str, str, str, str]]" = queue.Queue(maxsize=2000)
_audit_worker_started = False
_audit_worker_lock = threading.Lock()

DB_PATH = os.environ.get("SQLITE_DB_PATH", "").strip() or os.path.join(BASE_DIR, "produccion.db")

_ALLOWED_PRAGMA_SQL: Dict[str, str] = {
    "operadores": "PRAGMA table_info(operadores)",
    "usuarios": "PRAGMA table_info(usuarios)",
    "auditoria_seguridad": "PRAGMA table_info(auditoria_seguridad)",
    "historial": "PRAGMA table_info(historial)",
    "config_meta": "PRAGMA table_info(config_meta)",
    "ultimo_import": "PRAGMA table_info(ultimo_import)",
}

PASSWORD_HASH_METHOD = "scrypt:32768:8:1"


def get_database_url() -> str:
    raw = os.environ.get("DATABASE_URL", "").strip()
    if raw.startswith("postgres://"):
        raw = "postgresql://" + raw[len("postgres://") :]
    return raw


def is_postgres() -> bool:
    return bool(get_database_url())


def get_sqlite_path() -> str:
    return os.environ.get("SQLITE_DB_PATH", "").strip() or os.path.join(BASE_DIR, "produccion.db")


_PLANT_OFFSET_H = int(os.environ.get("PLANT_TZ_OFFSET_HOURS", "-4"))
PLANT_TZ = timezone(timedelta(hours=_PLANT_OFFSET_H))
PLANT_TZ_ISO_SUFFIX = f"{'+' if _PLANT_OFFSET_H >= 0 else '-'}{abs(_PLANT_OFFSET_H):02d}:00"


def plant_now() -> datetime:
    """Hora oficial de planta (Venezuela UTC-4 por defecto), independiente del reloj UTC de Render."""
    return datetime.now(PLANT_TZ).replace(tzinfo=None)


def _now_iso() -> str:
    return plant_now().strftime("%Y-%m-%d %H:%M:%S")


def _convert_qmarks_to_percent(sql: str) -> str:
    """Escapa '%' literales a '%%' y convierte placeholders '?' a '%s' fuera de literales."""
    out: List[str] = []
    in_single = False
    in_double = False
    i = 0
    n = len(sql)
    while i < n:
        ch = sql[i]
        if ch == "'" and not in_double:
            out.append(ch)
            if in_single and i + 1 < n and sql[i + 1] == "'":
                out.append("'")
                i += 2
                continue
            in_single = not in_single
        elif ch == '"' and not in_single:
            out.append(ch)
            in_double = not in_double
        elif ch == "%":
            out.append("%%")
        elif ch == "?" and not in_single and not in_double:
            out.append("%s")
        else:
            out.append(ch)
        i += 1
    return "".join(out)


class CompatCursor:
    """Envuelve un cursor de psycopg2 o sqlite3 con interfaz homogénea."""

    def __init__(self, cursor: Any, rowcount_override: Optional[int] = None) -> None:
        self._cur = cursor
        self._rowcount_override = rowcount_override

    @property
    def rowcount(self) -> int:
        if self._rowcount_override is not None:
            return self._rowcount_override
        return getattr(self._cur, "rowcount", -1)

    def fetchone(self) -> Any:
        return self._cur.fetchone()

    def fetchall(self) -> List[Any]:
        return self._cur.fetchall()


class CompatConnection:
    """Conexión uniforme sobre SQLite o PostgreSQL."""

    def __init__(self, raw_conn: Any, pg: bool = False) -> None:
        self._conn = raw_conn
        self.pg = pg

    def execute(self, sql: str, params: Optional[Sequence[Any]] = None) -> CompatCursor:
        if self.pg:
            cur = self._conn.cursor()
            if params:
                pg_sql = _convert_qmarks_to_percent(sql)
                cur.execute(pg_sql, tuple(params))
            else:
                cur.execute(sql)
            return CompatCursor(cur)
        cur = self._conn.execute(sql, tuple(params) if params is not None else ())
        return CompatCursor(cur)

    def executemany(self, sql: str, seq_of_params: Sequence[Sequence[Any]]) -> CompatCursor:
        if self.pg:
            pg_sql = _convert_qmarks_to_percent(sql)
            cur = self._conn.cursor()
            cur.executemany(pg_sql, [tuple(p) for p in seq_of_params])
            return CompatCursor(cur)
        cur = self._conn.executemany(sql, [tuple(p) for p in seq_of_params])
        return CompatCursor(cur)

    def table_columns(self, table: str) -> Set[str]:
        if self.pg:
            rows = self.execute(
                "SELECT column_name AS name FROM information_schema.columns WHERE table_name = ?",
                (table,),
            ).fetchall()
            return {r["name"] for r in rows}
        pragma_sql = _ALLOWED_PRAGMA_SQL.get(table)
        if not pragma_sql:
            raise ValueError(f"Tabla no permitida para inspección: {table}")
        rows = self.execute(pragma_sql).fetchall()
        return {r["name"] for r in rows}

    def commit(self) -> None:
        self._conn.commit()

    def rollback(self) -> None:
        self._conn.rollback()

    def close(self) -> None:
        self._conn.close()


class _SQLitePool:
    """Pool thread-safe de conexiones SQLite reutilizables con modo WAL."""

    def __init__(self, db_file: str, max_size: int = 8) -> None:
        self.db_file = db_file
        self.max_size = max(2, max_size)
        self._pool: "queue.LifoQueue[sqlite3.Connection]" = queue.LifoQueue(maxsize=self.max_size)
        self._created = 0
        self._lock = threading.Lock()

    def _new_conn(self) -> sqlite3.Connection:
        parent = os.path.dirname(os.path.abspath(self.db_file))
        if parent:
            os.makedirs(parent, exist_ok=True)
        conn = sqlite3.connect(self.db_file, timeout=15, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA synchronous=NORMAL;")
        conn.execute("PRAGMA busy_timeout=15000;")
        conn.execute("PRAGMA temp_store=MEMORY;")
        conn.execute("PRAGMA cache_size=-16000;")
        conn.execute("PRAGMA mmap_size=67108864;")
        conn.execute("PRAGMA foreign_keys=ON;")
        return conn

    def acquire(self) -> sqlite3.Connection:
        try:
            conn = self._pool.get_nowait()
            try:
                conn.execute("SELECT 1")
                return conn
            except sqlite3.Error:
                try:
                    conn.close()
                except Exception:
                    pass
        except queue.Empty:
            pass

        with self._lock:
            if self._created < self.max_size:
                self._created += 1
                return self._new_conn()

        try:
            conn = self._pool.get(timeout=10)
            try:
                conn.execute("SELECT 1")
                return conn
            except sqlite3.Error:
                try:
                    conn.close()
                except Exception:
                    pass
                return self._new_conn()
        except queue.Empty:
            return self._new_conn()

    def release(self, conn: sqlite3.Connection, broken: bool = False) -> None:
        if broken:
            try:
                conn.close()
            except Exception:
                pass
            with self._lock:
                self._created = max(0, self._created - 1)
            return
        try:
            self._pool.put_nowait(conn)
        except queue.Full:
            try:
                conn.close()
            except Exception:
                pass

    def close_all(self) -> None:
        while True:
            try:
                c = self._pool.get_nowait()
                c.close()
            except queue.Empty:
                break
        with self._lock:
            self._created = 0


def _get_sqlite_pool(db_file: str) -> _SQLitePool:
    norm_path = os.path.abspath(db_file)
    pool = _sqlite_pools.get(norm_path)
    if pool is not None:
        return pool
    with _sqlite_pools_lock:
        pool = _sqlite_pools.get(norm_path)
        if pool is None:
            pool_size = int(os.environ.get("DB_POOL_MAX", "20"))
            pool = _SQLitePool(norm_path, max_size=pool_size)
            _sqlite_pools[norm_path] = pool
    return pool


def close_all_pools() -> None:
    """Cierra los pools de conexiones activos (útil al apagar o cambiar BD en tests)."""
    with _sqlite_pools_lock:
        for p in _sqlite_pools.values():
            p.close_all()
        _sqlite_pools.clear()
    global _pg_pool, _pg_pool_url
    with _pg_pool_lock:
        if _pg_pool is not None:
            try:
                _pg_pool.closeall()
            except Exception:
                pass
            _pg_pool = None
            _pg_pool_url = ""


atexit.register(close_all_pools)


def _get_pg_pool(db_url: str):
    global _pg_pool, _pg_pool_url
    if _pg_pool is not None and _pg_pool_url == db_url:
        return _pg_pool
    with _pg_pool_lock:
        if _pg_pool is None or _pg_pool_url != db_url:
            from psycopg2.extras import RealDictCursor
            from psycopg2.pool import ThreadedConnectionPool

            min_c = max(1, int(os.environ.get("DB_POOL_MIN", "2")))
            max_c = max(min_c, int(os.environ.get("DB_POOL_MAX", "32")))
            _pg_pool = ThreadedConnectionPool(
                min_c,
                max_c,
                dsn=db_url,
                cursor_factory=RealDictCursor,
                connect_timeout=10,
                keepalives=1,
                keepalives_idle=30,
                keepalives_interval=10,
                keepalives_count=5,
            )
            _pg_pool_url = db_url
    return _pg_pool


def _acquire_healthy_pg_conn(pool: Any) -> Any:
    """Obtiene una conexión PostgreSQL viva del pool descartando sockets cerrados por idle/SSL EOF."""
    from psycopg2.pool import PoolError

    last_exc: Optional[Exception] = None
    for attempt in range(4):
        raw_conn = None
        for wait_try in range(25):
            try:
                raw_conn = pool.getconn()
                break
            except PoolError as exc:
                last_exc = exc
                if wait_try == 24:
                    raise
                time.sleep(0.1)
        if raw_conn is None:
            continue
        if getattr(raw_conn, "closed", 0):
            try:
                pool.putconn(raw_conn, close=True)
            except Exception:
                pass
            continue
        try:
            if not getattr(raw_conn, "autocommit", False):
                raw_conn.rollback()
            with raw_conn.cursor() as ping_cur:
                ping_cur.execute("SELECT 1")
            if not getattr(raw_conn, "autocommit", False):
                raw_conn.rollback()
            return raw_conn
        except Exception as exc:
            last_exc = exc
            try:
                pool.putconn(raw_conn, close=True)
            except Exception:
                pass
            if attempt == 3:
                raise
    if last_exc is not None:
        raise last_exc
    return pool.getconn()


@contextmanager
def get_db():
    db_url = get_database_url()
    if db_url:
        pool = _get_pg_pool(db_url)
        raw_conn = _acquire_healthy_pg_conn(pool)
        conn = CompatConnection(raw_conn, pg=True)
        broken = False
        try:
            yield conn
            conn.commit()
        except Exception:
            broken = True
            try:
                conn.rollback()
            except Exception:
                pass
            raise
        finally:
            pool.putconn(raw_conn, close=broken)
    else:
        db_file = get_sqlite_path()
        pool = _get_sqlite_pool(db_file)
        raw_conn = pool.acquire()
        conn = CompatConnection(raw_conn, pg=False)
        broken = False
        try:
            yield conn
            conn.commit()
        except Exception:
            broken = True
            try:
                conn.rollback()
            except Exception:
                pass
            raise
        finally:
            pool.release(raw_conn, broken=broken)


ROLES_VALIDOS_USER = ("admin", "gerencia", "supervisor", "supervision", "mantenimiento", "coordinador", "tv")


def _db_cache_key(pg: bool) -> str:
    return f"pg:{get_database_url()}" if pg else f"sqlite:{os.path.abspath(get_sqlite_path())}"


def init_auth_tables(conn: CompatConnection) -> None:
    """Ejecuta DDL de autenticación e índices compuestos (solo en arranque/migración)."""
    cache_key = _db_cache_key(conn.pg)
    if cache_key in _auth_tables_ready:
        return
    with _auth_init_lock:
        if cache_key in _auth_tables_ready:
            return
        if conn.pg:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS usuarios (
                    id SERIAL PRIMARY KEY,
                    username TEXT NOT NULL UNIQUE,
                    nombre TEXT NOT NULL,
                    password_hash TEXT NOT NULL,
                    rol TEXT NOT NULL DEFAULT 'supervisor',
                    cargo TEXT NOT NULL DEFAULT 'Supervisor de Envasado',
                    activo INTEGER NOT NULL DEFAULT 1,
                    created_at TEXT NOT NULL DEFAULT '',
                    ultimo_login TEXT NOT NULL DEFAULT ''
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS auditoria_seguridad (
                    id SERIAL PRIMARY KEY,
                    timestamp TEXT NOT NULL,
                    evento TEXT NOT NULL,
                    usuario TEXT NOT NULL DEFAULT '',
                    rol TEXT NOT NULL DEFAULT '',
                    ip TEXT NOT NULL DEFAULT '',
                    ruta TEXT NOT NULL DEFAULT '',
                    metodo TEXT NOT NULL DEFAULT '',
                    detalle TEXT NOT NULL DEFAULT ''
                )
                """
            )
        else:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS usuarios (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT NOT NULL UNIQUE,
                    nombre TEXT NOT NULL,
                    password_hash TEXT NOT NULL,
                    rol TEXT NOT NULL DEFAULT 'supervisor',
                    cargo TEXT NOT NULL DEFAULT 'Supervisor de Envasado',
                    activo INTEGER NOT NULL DEFAULT 1,
                    created_at TEXT NOT NULL DEFAULT '',
                    ultimo_login TEXT NOT NULL DEFAULT ''
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS auditoria_seguridad (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp TEXT NOT NULL,
                    evento TEXT NOT NULL,
                    usuario TEXT NOT NULL DEFAULT '',
                    rol TEXT NOT NULL DEFAULT '',
                    ip TEXT NOT NULL DEFAULT '',
                    ruta TEXT NOT NULL DEFAULT '',
                    metodo TEXT NOT NULL DEFAULT '',
                    detalle TEXT NOT NULL DEFAULT ''
                )
                """
            )
        conn.execute("CREATE INDEX IF NOT EXISTS idx_usuarios_activo_username ON usuarios (activo, username)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_auditoria_evento_id ON auditoria_seguridad (evento, id DESC)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_auditoria_usuario_id ON auditoria_seguridad (usuario, id DESC)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_auditoria_ts ON auditoria_seguridad (timestamp)")
        _auth_tables_ready.add(cache_key)


def ensure_auth_tables_ready() -> None:
    """Garantiza tablas de auth fuera de rutas calientes cuando aún no se han inicializado para la BD actual."""
    key = _db_cache_key(is_postgres())
    if key in _auth_tables_ready:
        return
    with get_db() as conn:
        init_auth_tables(conn)


def upsert_user(
    conn: CompatConnection,
    username: str,
    password: str,
    nombre: str,
    rol: str = "supervisor",
    cargo: str = "Supervisor de Envasado",
    activo: int = 1,
) -> None:
    uname = str(username or "").strip().lower()
    if not uname or len(uname) > 64:
        raise ValueError("Usuario inválido.")
    if not password or len(password) < 8 or len(password) > 128:
        raise ValueError("La contraseña debe tener entre 8 y 128 caracteres.")
    rol_clean = str(rol or "supervisor").strip().lower()
    if rol_clean not in ROLES_VALIDOS_USER:
        rol_clean = "supervisor"
    nombre_clean = str(nombre or uname).strip()[:120] or uname
    cargo_clean = str(cargo or "").strip()[:120]
    pw_hash = generate_password_hash(password, method=PASSWORD_HASH_METHOD)
    existing = conn.execute("SELECT id FROM usuarios WHERE LOWER(username) = ?", (uname,)).fetchone()
    if existing:
        conn.execute(
            """
            UPDATE usuarios
            SET nombre = ?, password_hash = ?, rol = ?, cargo = ?, activo = ?
            WHERE id = ?
            """,
            (nombre_clean, pw_hash, rol_clean, cargo_clean, int(bool(activo)), existing["id"]),
        )
    else:
        conn.execute(
            """
            INSERT INTO usuarios (username, nombre, password_hash, rol, cargo, activo, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (uname, nombre_clean, pw_hash, rol_clean, cargo_clean, int(bool(activo)), _now_iso()),
        )


def list_users_safe(conn: CompatConnection) -> List[Dict[str, Any]]:
    """Lista usuarios sin exponer password_hash."""
    rows = conn.execute(
        "SELECT id, username, nombre, rol, cargo, activo, created_at, ultimo_login FROM usuarios ORDER BY id ASC"
    ).fetchall()
    return [dict(r) for r in rows]


def verificar_usuario(username: str, password: str) -> Optional[Dict[str, Any]]:
    """Valida credenciales contra la tabla usuarios (acepta usuario o nombre visible único)."""
    if not username or not password:
        return None
    uname = str(username).strip().lower()
    try:
        with get_db() as conn:
            cur = conn.execute(
                "SELECT id, username, nombre, password_hash, rol, cargo, activo FROM usuarios WHERE LOWER(username) = ? AND activo = 1",
                (uname,),
            ).fetchone()
            if not cur:
                matches = conn.execute(
                    "SELECT id, username, nombre, password_hash, rol, cargo, activo FROM usuarios WHERE LOWER(nombre) = ? AND activo = 1",
                    (uname,),
                ).fetchall()
                if len(matches) == 1:
                    cur = matches[0]
            if not cur:
                return None
            u = dict(cur)
            if check_password_hash(u["password_hash"], password):
                if not str(u["password_hash"]).startswith("scrypt:"):
                    upgraded = generate_password_hash(password, method=PASSWORD_HASH_METHOD)
                    conn.execute(
                        "UPDATE usuarios SET ultimo_login = ?, password_hash = ? WHERE id = ?",
                        (_now_iso(), upgraded, u["id"]),
                    )
                else:
                    conn.execute(
                        "UPDATE usuarios SET ultimo_login = ? WHERE id = ?",
                        (_now_iso(), u["id"]),
                    )
                return {
                    "id": u["id"],
                    "username": u["username"],
                    "nombre": u["nombre"],
                    "rol": u["rol"],
                    "cargo": u.get("cargo", ""),
                }
    except Exception as exc:
        print("[AUTH ERROR] Error verificando usuario:", exc)
    return None


def _write_audit_batch(rows: List[Tuple[str, str, str, str, str, str, str, str]]) -> None:
    if not rows:
        return
    try:
        with get_db() as conn:
            conn.executemany(
                """
                INSERT INTO auditoria_seguridad (timestamp, evento, usuario, rol, ip, ruta, metodo, detalle)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                rows,
            )
    except Exception as exc:
        print("[AUDITORIA ERROR]", exc)


def _audit_worker_loop() -> None:
    while True:
        item = _audit_queue.get()
        batch = [item]
        try:
            while len(batch) < 25:
                batch.append(_audit_queue.get_nowait())
        except queue.Empty:
            pass
        _write_audit_batch(batch)
        for _ in batch:
            _audit_queue.task_done()


def _ensure_audit_worker() -> None:
    global _audit_worker_started
    if _audit_worker_started:
        return
    with _audit_worker_lock:
        if not _audit_worker_started:
            t = threading.Thread(target=_audit_worker_loop, daemon=True, name="audit-writer")
            t.start()
            _audit_worker_started = True


def flush_auditoria() -> None:
    """Espera a que la cola de auditoría se persista antes de leer reportes."""
    if _audit_worker_started:
        _audit_queue.join()


def registrar_auditoria(
    evento: str,
    usuario: str = "",
    rol: str = "",
    ip: str = "",
    ruta: str = "",
    metodo: str = "",
    detalle: str = "",
) -> None:
    """Encola eventos de seguridad sin bloquear el tiempo de respuesta del usuario."""
    _ensure_audit_worker()
    row = (
        _now_iso(),
        str(evento or "")[:80],
        str(usuario or "")[:120],
        str(rol or "")[:60],
        str(ip or "")[:64],
        str(ruta or "")[:240],
        str(metodo or "")[:16],
        str(detalle or "")[:8000],
    )
    try:
        _audit_queue.put_nowait(row)
    except queue.Full:
        _write_audit_batch([row])


# ── Cola general de trabajos en segundo plano ─────────────────────────────────
class BackgroundJobQueue:
    """Cola de trabajos asíncronos con trazabilidad de estado y control de propietario (anti-IDOR)."""

    def __init__(self, max_jobs: int = 200) -> None:
        self._queue: "queue.Queue[Tuple[str, Callable[..., Any], Tuple[Any, ...], Dict[str, Any]]]" = queue.Queue(maxsize=256)
        self._jobs: "OrderedDict[str, Dict[str, Any]]" = OrderedDict()
        self._lock = threading.Lock()
        self._max_jobs = max_jobs
        self._worker_started = False

    def _ensure_worker(self) -> None:
        if self._worker_started:
            return
        with self._lock:
            if not self._worker_started:
                t = threading.Thread(target=self._worker_loop, daemon=True, name="bg-job-worker")
                t.start()
                self._worker_started = True

    def _worker_loop(self) -> None:
        while True:
            job_id, fn, args, kwargs = self._queue.get()
            with self._lock:
                rec = self._jobs.get(job_id)
                if rec:
                    rec["status"] = "running"
                    rec["started_at"] = _now_iso()
            try:
                result = fn(*args, **kwargs)
                with self._lock:
                    rec = self._jobs.get(job_id)
                    if rec:
                        rec["status"] = "completed"
                        rec["finished_at"] = _now_iso()
                        rec["result"] = result if isinstance(result, (dict, list, str, int, float, bool)) else str(result)
            except Exception as exc:
                with self._lock:
                    rec = self._jobs.get(job_id)
                    if rec:
                        rec["status"] = "failed"
                        rec["finished_at"] = _now_iso()
                        rec["error"] = str(exc)[:300]
            finally:
                self._queue.task_done()

    def enqueue(
        self,
        job_type: str,
        owner_username: str,
        fn: Callable[..., Any],
        *args: Any,
        **kwargs: Any,
    ) -> str:
        self._ensure_worker()
        job_id = uuid.uuid4().hex[:16]
        record = {
            "id": job_id,
            "type": str(job_type)[:64],
            "owner": str(owner_username or "").strip().lower(),
            "status": "queued",
            "created_at": _now_iso(),
            "started_at": "",
            "finished_at": "",
            "result": None,
            "error": "",
        }
        with self._lock:
            while len(self._jobs) >= self._max_jobs:
                self._jobs.popitem(last=False)
            self._jobs[job_id] = record
        self._queue.put((job_id, fn, args, kwargs))
        return job_id

    def get_job(self, job_id: str) -> Optional[Dict[str, Any]]:
        with self._lock:
            rec = self._jobs.get(job_id)
            return dict(rec) if rec else None

    def flush(self, timeout_s: float = 5.0) -> None:
        if not self._worker_started:
            return
        deadline = time.monotonic() + timeout_s
        while self._queue.unfinished_tasks > 0 and time.monotonic() < deadline:
            time.sleep(0.02)


bg_jobs = BackgroundJobQueue()
