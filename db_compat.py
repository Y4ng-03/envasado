"""
db_compat.py — Adaptador unificado SQLite (local) / PostgreSQL (Render + Neon/Supabase)
y utilidades de usuarios y auditoría de seguridad (sin claves quemadas en código).
"""

import os
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from typing import Any, Dict, List, Optional, Sequence, Set, Tuple

from dotenv import load_dotenv
from werkzeug.security import check_password_hash, generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"), override=False)

DB_PATH = os.environ.get("SQLITE_DB_PATH", "").strip() or os.path.join(BASE_DIR, "produccion.db")


def get_database_url() -> str:
    raw = os.environ.get("DATABASE_URL", "").strip()
    if raw.startswith("postgres://"):
        raw = "postgresql://" + raw[len("postgres://"):]
    return raw


def is_postgres() -> bool:
    return bool(get_database_url())


def _now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _convert_qmarks_to_percent(sql: str) -> str:
    """Convierte placeholders '?' a '%s' fuera de literales de texto."""
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
            pg_sql = _convert_qmarks_to_percent(sql)
            cur = self._conn.cursor()
            cur.execute(pg_sql, tuple(params) if params is not None else ())
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
        rows = self.execute(f"PRAGMA table_info({table})").fetchall()
        return {r["name"] for r in rows}

    def commit(self) -> None:
        self._conn.commit()

    def rollback(self) -> None:
        self._conn.rollback()

    def close(self) -> None:
        self._conn.close()


@contextmanager
def get_db():
    db_url = get_database_url()
    if db_url:
        import psycopg2
        from psycopg2.extras import RealDictCursor

        raw_conn = psycopg2.connect(db_url, cursor_factory=RealDictCursor)
        conn = CompatConnection(raw_conn, pg=True)
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()
    else:
        raw_conn = sqlite3.connect(DB_PATH, timeout=15, check_same_thread=False)
        raw_conn.row_factory = sqlite3.Row
        conn = CompatConnection(raw_conn, pg=False)
        try:
            conn.execute("PRAGMA journal_mode=WAL;")
            conn.execute("PRAGMA synchronous=NORMAL;")
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()


ROLES_VALIDOS_USER = ("admin", "gerencia", "supervisor", "tv")


def init_auth_tables(conn: CompatConnection) -> None:
    pk_auto = "SERIAL PRIMARY KEY" if conn.pg else "INTEGER PRIMARY KEY AUTOINCREMENT"
    conn.execute(
        f"""
        CREATE TABLE IF NOT EXISTS usuarios (
            id {pk_auto},
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
        f"""
        CREATE TABLE IF NOT EXISTS auditoria_seguridad (
            id {pk_auto},
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
    if not password or len(password) < 8:
        raise ValueError("La contraseña debe tener al menos 8 caracteres.")
    rol_clean = str(rol or "supervisor").strip().lower()
    if rol_clean not in ROLES_VALIDOS_USER:
        rol_clean = "supervisor"
    nombre_clean = str(nombre or uname).strip()[:120] or uname
    cargo_clean = str(cargo or "").strip()[:120]
    pw_hash = generate_password_hash(password)
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
            init_auth_tables(conn)
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


def registrar_auditoria(
    evento: str,
    usuario: str = "",
    rol: str = "",
    ip: str = "",
    ruta: str = "",
    metodo: str = "",
    detalle: str = "",
) -> None:
    """Registra eventos de seguridad y acciones críticas en BD."""
    try:
        with get_db() as conn:
            init_auth_tables(conn)
            conn.execute(
                """
                INSERT INTO auditoria_seguridad (timestamp, evento, usuario, rol, ip, ruta, metodo, detalle)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    _now_iso(),
                    str(evento or "")[:80],
                    str(usuario or "")[:120],
                    str(rol or "")[:60],
                    str(ip or "")[:64],
                    str(ruta or "")[:240],
                    str(metodo or "")[:16],
                    str(detalle or "")[:8000],
                ),
            )
    except Exception as exc:
        print("[AUDITORIA ERROR]", exc)
