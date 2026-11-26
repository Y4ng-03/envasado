"""
cuaderno_envases.py — Cuaderno Digital de Consumo de Envases en Producción (FO003 - RP06 - PRO).
100% Modo Claro, sincronizado en tiempo real (SSE), casado automáticamente con el puesto de
Despaletizador (Grupo A / Grupo B) de la Pizarra Digital de Envasado, soporte multi-proveedor
(AMCOR y OTRO: LT / VL / otros prefijos) y exportación/importación Excel oficial (.xlsx)
para todas las líneas (L1, L2, L3, L4) y sus presentaciones.
"""

import base64
import csv
import hashlib
import io
import json
import re
import threading
import time
import uuid
from datetime import date as dt_date, datetime, time as dt_time
from typing import Any, Callable, Dict, List, Optional, Set, Tuple

from flask import Blueprint, Response, jsonify, render_template, request, send_file, session
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from db_compat import get_db, plant_now
from observabilidad import compute_etag, etag_matches
from seguridad import (
    ROLES_CUADERNO_ADMIN,
    SESION_MAX_S,
    get_csp_nonce,
    get_csrf_token,
    is_valid_op_id,
    validar_archivo_importacion,
    validar_fecha_logica,
    validar_imagen_data_url,
)

LINEAS_CUADERNO = ("L1", "L2", "L3", "L4")
TURNOS_CUADERNO = ("T1", "T2", "T3")
PREFIJOS_VALIDOS = ("LT", "VL", "OT")
_PREFIJO_CUSTOM_RE = re.compile(r"^[A-Z0-9]{1,4}$")

LINEA_DEFAULTS: Dict[str, Dict[str, Any]] = {
    "L1": {
        "nombre": "Línea 1",
        "codigo": "Cod.B",
        "presentaciones": [
            {"label": "355 ml", "cod": "32"},
        ],
        "default_pres": "355 ml",
        "tiene_sabor": False,
        "guia": {"lote": "26/10/18", "hora_desp": "11:00:00", "pref": "LT/VL", "bull": "83505072", "fecha_prod": "12/16/18", "hora_p": "05:37:00"},
    },
    "L2": {
        "nombre": "Línea 2",
        "codigo": "Cod.B",
        "presentaciones": [
            {"label": "600 ml", "cod": "80"},
            {"label": "1500 ml", "cod": "59"},
        ],
        "default_pres": "600 ml",
        "tiene_sabor": False,
        "guia": {"lote": "26/10/18", "hora_desp": "11:00:00", "pref": "LT/VL", "bull": "83505072", "fecha_prod": "12/16/18", "hora_p": "05:37:00"},
    },
    "L3": {
        "nombre": "Línea 3",
        "codigo": "Cod.B",
        "presentaciones": [
            {"label": "5 Litros (5L)", "cod": "390"},
        ],
        "default_pres": "5 Litros (5L)",
        "tiene_sabor": False,
        "guia": {"lote": "26/10/18", "hora_desp": "11:00:00", "pref": "LT/VL", "bull": "G3228973", "fecha_prod": "11/18/26", "hora_p": "09:23:00"},
    },
    "L4": {
        "nombre": "Línea 4",
        "codigo": "Cod.B",
        "presentaciones": [
            {"label": "500 ml", "cod": "18"},
            {"label": "1000 ml", "cod": ""},
        ],
        "default_pres": "500 ml",
        "tiene_sabor": True,
        "guia": {"lote": "26/10/18", "hora_desp": "11:00:00", "pref": "LT/VL", "bull": "91916156", "fecha_prod": "2019-10-17", "hora_p": "07:49:00"},
    },
}

_CUADERNO_HOJAS_ALTER_SQL: Dict[str, str] = {
    "operador_auto": "ALTER TABLE cuaderno_hojas ADD COLUMN operador_auto INTEGER NOT NULL DEFAULT 1",
    "analista_calidad": "ALTER TABLE cuaderno_hojas ADD COLUMN analista_calidad TEXT NOT NULL DEFAULT ''",
    "facilitador": "ALTER TABLE cuaderno_hojas ADD COLUMN facilitador TEXT NOT NULL DEFAULT ''",
}

# Soporta etiquetas AMCOR (LTG3228973, LTG2976017, 011563228973) y otros proveedores (VL91916156, VL62477447, prefijos 2 letras o numéricos)
_BULL_TOKEN_RE = re.compile(r"\b(VL|LT|OT)\s*[-:]?\s*([A-Z0-9]{6,12})\b", re.IGNORECASE)
_OTHER_PREFIX_BULL_RE = re.compile(r"\b(?!(?:NO|ID|OK|SAP)\b)([A-Z]{2})\s*[-:]?\s*((?=[A-Z0-9]*\d{4})[A-Z0-9]{6,12})\b", re.IGNORECASE)
_SAP_MATERIAL_RE = re.compile(r"\b(500\d{5})\b")
_AMCOR_UPPER_BAR_RE = re.compile(r"\b0?115\d(\d{7})\b")
_DIGITS_8_RE = re.compile(r"\b([A-Z]?\d{7,8})\b", re.IGNORECASE)
_TIME_HHMM_RE = re.compile(r"\b([01]?\d|2[0-3]):([0-5]\d)(?::([0-5]\d))?\b")
_DATE_PRINT_RE = re.compile(r"\b(\d{2}[/-]\d{2}[/-]\d{2,4}|\d{4}-\d{2}-\d{2})\b")

_tag_photo_lock = threading.Lock()
_tag_photo_cache: Dict[str, Tuple[str, str, bytes, str]] = {}
_FOTO_RX = re.compile(r"^data:(image/(?:png|jpeg|jpg|webp));base64,([A-Za-z0-9+/=]+)$")


def _now_iso() -> str:
    return plant_now().strftime("%Y-%m-%d %H:%M:%S")


def _today_ymd() -> str:
    return plant_now().strftime("%Y-%m-%d")


def _infer_current_turno(now_dt: Optional[datetime] = None) -> str:
    """T1: 06:00–14:30, T2: 14:30–22:30, T3: 22:30–06:00."""
    dt = now_dt or plant_now()
    mins = dt.hour * 60 + dt.minute
    if 6 * 60 <= mins < 14 * 60 + 30:
        return "T1"
    if 14 * 60 + 30 <= mins < 22 * 60 + 30:
        return "T2"
    return "T3"


def _clean_str(val: Any, max_len: int = 120) -> str:
    if val is None:
        return ""
    if isinstance(val, datetime):
        if val.hour == 0 and val.minute == 0 and val.second == 0:
            s = val.strftime("%d/%m/%y")
        else:
            s = val.strftime("%d/%m/%y %H:%M")
    elif isinstance(val, dt_date):
        s = val.strftime("%d/%m/%y")
    elif isinstance(val, dt_time):
        s = val.strftime("%H:%M:%S") if val.second else val.strftime("%H:%M")
    elif isinstance(val, float) and val.is_integer():
        s = str(int(val))
    else:
        s = str(val)
    return re.sub(r"\s+", " ", s).strip()[:max_len]


def _normalize_prefix(raw_pref: Any, default: str = "LT") -> str:
    s = re.sub(r"[^A-Z0-9]", "", str(raw_pref or "").upper())[:4]
    if s in PREFIJOS_VALIDOS or _PREFIJO_CUSTOM_RE.match(s):
        return s
    return default


def _clean_bull_code(raw_bull: Any) -> str:
    s = re.sub(r"[^A-Z0-9]", "", str(raw_bull or "").upper())
    if s.startswith(("VL", "LT", "OT")) and len(s) >= 8:
        s = s[2:]
    return s[:8] if len(s) > 8 else s


def _effective_group(row: Any) -> str:
    estado = str(row["estado"] or "normal").strip()
    cambio = str(row["cambio_grupo"] or "").strip()
    grupo = str(row["grupo"] or "").strip()
    if estado == "cambio_turno" and cambio:
        return cambio
    return grupo


def init_cuaderno_tables(conn: Any) -> None:
    """Crea y migra tablas e índices del Cuaderno Digital FO003-RP06-PRO en SQLite o PostgreSQL."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS cuaderno_hojas (
            id TEXT PRIMARY KEY,
            fecha TEXT NOT NULL,
            linea TEXT NOT NULL,
            turno TEXT NOT NULL,
            grupo TEXT NOT NULL DEFAULT '',
            codigo_linea TEXT NOT NULL DEFAULT 'Cod.B',
            operador TEXT NOT NULL DEFAULT '',
            operador_auto INTEGER NOT NULL DEFAULT 1,
            relevo TEXT NOT NULL DEFAULT '',
            proveedor_envase TEXT NOT NULL DEFAULT 'AMCOR',
            proveedor_otro TEXT NOT NULL DEFAULT '',
            presentacion TEXT NOT NULL DEFAULT '500 ml',
            sabor TEXT NOT NULL DEFAULT '',
            material_sap TEXT NOT NULL DEFAULT '',
            analista_calidad TEXT NOT NULL DEFAULT '',
            facilitador TEXT NOT NULL DEFAULT '',
            updated_at TEXT NOT NULL,
            updated_by TEXT NOT NULL DEFAULT ''
        )
        """
    )
    existing_cols = conn.table_columns("cuaderno_hojas")
    for col_name, alter_sql in _CUADERNO_HOJAS_ALTER_SQL.items():
        if col_name not in existing_cols:
            conn.execute(alter_sql)

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS cuaderno_registros (
            id TEXT PRIMARY KEY,
            hoja_id TEXT NOT NULL,
            num_item INTEGER NOT NULL DEFAULT 1,
            lote_sap TEXT NOT NULL DEFAULT '',
            hora_desp TEXT NOT NULL DEFAULT '',
            prefijo_proveedor TEXT NOT NULL DEFAULT 'LT',
            codigo_bull TEXT NOT NULL DEFAULT '',
            fecha_prod TEXT NOT NULL DEFAULT '',
            hora_prod TEXT NOT NULL DEFAULT '',
            observacion TEXT NOT NULL DEFAULT '',
            raw_barcode TEXT NOT NULL DEFAULT '',
            foto_etiqueta TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL,
            created_by TEXT NOT NULL DEFAULT ''
        )
        """
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_cuaderno_hojas_fecha_linea ON cuaderno_hojas (fecha DESC, linea, turno)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_cuaderno_reg_hoja_num ON cuaderno_registros (hoja_id, num_item ASC)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_cuaderno_reg_bull ON cuaderno_registros (hoja_id, codigo_bull)")


def make_hoja_id(fecha: str, linea: str, turno: str) -> str:
    return f"{fecha}_{linea}_{turno}"


def parse_bull_barcode(raw_text: str, default_pref: str = "LT") -> Dict[str, str]:
    """
    Extrae datos de etiquetas AMCOR y de otros proveedores al escanear:
    - Ej. 'LTG3228973' / 'LTG2976017' -> prefijo='LT', codigo_bull='G3228973' (8 casillas)
    - Ej. 'VL91916156' / 'VL62477447' -> prefijo='VL', codigo_bull='91916156'
    - Ej. 'PL88776655' / '62477447' (OTRO proveedor) -> conserva prefijo o usa default_pref
    - Ej. '50034396 ... 11/18/26 09:23 LTG3228973' -> extrae SAP, Fecha Prod, Hora P y Bull.
    """
    s = _clean_str(raw_text, 400)
    fallback_pref = _normalize_prefix(default_pref, "LT")
    out = {
        "prefijo_proveedor": fallback_pref,
        "codigo_bull": "",
        "fecha_prod": "",
        "hora_prod": "",
        "material_sap": "",
    }
    if not s:
        return out

    m_sap = _SAP_MATERIAL_RE.search(s)
    rest = s
    if m_sap:
        out["material_sap"] = m_sap.group(1)
        rest = rest.replace(m_sap.group(1), " ", 1)

    m_bull = _BULL_TOKEN_RE.search(rest)
    if m_bull:
        out["prefijo_proveedor"] = m_bull.group(1).upper()
        tail = re.sub(r"[^A-Z0-9]", "", m_bull.group(2).upper())
        out["codigo_bull"] = tail[-8:] if len(tail) > 8 else tail
        rest = rest[: m_bull.start()] + " " + rest[m_bull.end() :]
    else:
        m_other = _OTHER_PREFIX_BULL_RE.search(rest)
        if m_other:
            out["prefijo_proveedor"] = m_other.group(1).upper()
            tail = re.sub(r"[^A-Z0-9]", "", m_other.group(2).upper())
            out["codigo_bull"] = tail[-8:] if len(tail) > 8 else tail
            rest = rest[: m_other.start()] + " " + rest[m_other.end() :]
        else:
            m_upper = _AMCOR_UPPER_BAR_RE.search(rest)
            if m_upper:
                out["prefijo_proveedor"] = "LT"
                out["codigo_bull"] = m_upper.group(1)
                rest = rest.replace(m_upper.group(0), " ", 1)
            else:
                m_any = _DIGITS_8_RE.findall(rest)
                if m_any:
                    cand = m_any[0].upper()
                    out["prefijo_proveedor"] = "LT" if cand[0].isalpha() else ("VL" if fallback_pref == "LT" else fallback_pref)
                    out["codigo_bull"] = cand[-8:]
                else:
                    alnum = re.sub(r"[^A-Z0-9]", "", rest.upper())
                    if len(alnum) >= 6:
                        out["codigo_bull"] = alnum[-8:]

    m_time = _TIME_HHMM_RE.search(rest)
    if m_time:
        out["hora_prod"] = f"{int(m_time.group(1)):02d}:{m_time.group(2)}"

    m_date = _DATE_PRINT_RE.search(rest)
    if m_date:
        out["fecha_prod"] = m_date.group(1)

    return out


def _cache_tag_photo(reg_id: str, data_url: str) -> str:
    raw = (data_url or "").strip()
    if not raw or not raw.startswith("data:image/"):
        if reg_id:
            with _tag_photo_lock:
                _tag_photo_cache.pop(reg_id, None)
        return ""
    ver = hashlib.sha1(raw.encode("utf-8")).hexdigest()[:12]
    with _tag_photo_lock:
        ex = _tag_photo_cache.get(reg_id)
        if ex and ex[0] == ver:
            return ver
    m = _FOTO_RX.match(raw)
    if not m:
        return ""
    mime = "image/jpeg" if m.group(1).lower() == "image/jpg" else m.group(1).lower()
    try:
        img_bytes = base64.b64decode(m.group(2), validate=True)
    except Exception:
        return ""
    if not img_bytes:
        return ""
    etag = f'"bullfoto-{reg_id}-{ver}"'
    with _tag_photo_lock:
        if len(_tag_photo_cache) >= 512 and reg_id not in _tag_photo_cache:
            _tag_photo_cache.pop(next(iter(_tag_photo_cache)), None)
        _tag_photo_cache[reg_id] = (ver, mime, img_bytes, etag)
    return ver


def resolve_pizarra_link(conn: Any, linea: str, grupo_pref: str = "") -> Dict[str, Any]:
    """
    Obtiene en tiempo real desde la Pizarra quién está en el puesto de Despaletizador
    (por Grupo A, Grupo B, etc.) y quién está en Relevo / Pela paletas.
    """
    rows = conn.execute(
        """
        SELECT id, nombre, columna, puesto, grupo, estado, cambio_grupo, posicion
        FROM operadores
        WHERE columna IN ('L1', 'L2', 'L3', 'L4', 'RELEVO')
          AND estado NOT IN ('vacaciones', 'ausente', 'licencia')
        ORDER BY posicion ASC, nombre ASC
        """
    ).fetchall()

    desp_slots = {
        "L1": {1},
        "L2": {1},
        "L3": {0, 1},
        "L4": {0, 1},
    }.get(linea, {0, 1})

    grupos_set: Set[str] = {"Grupo A", "Grupo B"}
    por_grupo_desp: Dict[str, List[str]] = {}
    for r in rows:
        eg = _effective_group(r)
        if eg:
            grupos_set.add(eg)
        if r["columna"] == linea:
            pos = int(r["posicion"] or 0)
            p_low = (r["puesto"] or "").lower()
            if pos in desp_slots or "despaletiz" in p_low:
                por_grupo_desp.setdefault(eg or "General", []).append(r["nombre"])
        elif r["columna"] == "L1" and int(r["posicion"] or 0) == 0 and linea in ("L1", "L2"):
            por_grupo_desp.setdefault(eg or "General", [])

    target_group = _clean_str(grupo_pref, 40)
    if not target_group:
        for cand_g in ("Grupo A", "Grupo B", "Grupo C"):
            if por_grupo_desp.get(cand_g):
                target_group = cand_g
                break
        if not target_group and por_grupo_desp:
            first_g = next(iter(por_grupo_desp.keys()))
            target_group = "" if first_g == "General" else first_g

    def _matches_grp(r: Any) -> bool:
        if not target_group:
            return True
        eg = _effective_group(r)
        return (not eg) or eg == target_group

    despaletizadores: List[Dict[str, Any]] = []
    pela_paletas: List[Dict[str, Any]] = []
    relevos: List[Dict[str, Any]] = []

    for r in rows:
        if not _matches_grp(r):
            continue
        col = r["columna"]
        pos = int(r["posicion"] or 0)
        eg = _effective_group(r)
        item = {
            "id": r["id"],
            "nombre": r["nombre"],
            "puesto": r["puesto"] or "",
            "grupo": eg,
            "posicion": pos,
        }
        if col == "L1" and pos == 0 and linea in ("L1", "L2"):
            pela_paletas.append(item)
        elif col == linea:
            p_low = (r["puesto"] or "").lower()
            if pos in desp_slots or "despaletiz" in p_low:
                despaletizadores.append(item)
        elif col == "RELEVO":
            relevos.append(item)

    sugerido_op = ""
    sugerido_rol = ""
    if despaletizadores:
        sugerido_op = " / ".join(d["nombre"] for d in despaletizadores[:2])
        sugerido_rol = "Despaletizador"
    elif pela_paletas:
        sugerido_op = pela_paletas[0]["nombre"]
        sugerido_rol = "Pela paletas"

    sugerido_relevo = relevos[0]["nombre"] if relevos else ""

    return {
        "linea": linea,
        "grupo_activo": target_group or "Grupo A",
        "grupos_disponibles": sorted(grupos_set),
        "despaletizador_por_grupo": {k: " / ".join(v[:2]) for k, v in por_grupo_desp.items() if v},
        "operador_sugerido": sugerido_op,
        "puesto_sugerido": sugerido_rol,
        "relevo_sugerido": sugerido_relevo,
        "despaletizadores": despaletizadores,
        "pela_paletas": pela_paletas,
        "relevos": relevos,
    }


def _ensure_or_load_hoja(conn: Any, fecha: str, linea: str, turno: str, actor: str = "Sistema") -> Tuple[Dict[str, Any], Dict[str, Any]]:
    hoja_id = make_hoja_id(fecha, linea, turno)
    row = conn.execute("SELECT * FROM cuaderno_hojas WHERE id = ?", (hoja_id,)).fetchone()
    pizarra_info = resolve_pizarra_link(conn, linea, row["grupo"] if row else "")

    if not row:
        cfg = LINEA_DEFAULTS.get(linea, LINEA_DEFAULTS["L1"])
        now_s = _now_iso()
        grupo_ini = pizarra_info["grupo_activo"] or "Grupo A"
        op_ini = pizarra_info["operador_sugerido"] or ""
        rel_ini = pizarra_info["relevo_sugerido"] or ""
        conn.execute(
            """
            INSERT INTO cuaderno_hojas (
                id, fecha, linea, turno, grupo, codigo_linea, operador, operador_auto, relevo,
                proveedor_envase, proveedor_otro, presentacion, sabor, material_sap,
                analista_calidad, facilitador, updated_at, updated_by
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, 1, ?, 'AMCOR', '', ?, '', '', '', '', ?, ?)
            ON CONFLICT (id) DO NOTHING
            """,
            (
                hoja_id,
                fecha,
                linea,
                turno,
                grupo_ini,
                cfg["codigo"],
                op_ini,
                rel_ini,
                cfg["default_pres"],
                now_s,
                actor,
            ),
        )
        row = conn.execute("SELECT * FROM cuaderno_hojas WHERE id = ?", (hoja_id,)).fetchone()
    else:
        is_auto = int(row["operador_auto"] if "operador_auto" in row.keys() and row["operador_auto"] is not None else 1) == 1
        # Si es hoja del día y está casada automáticamente (operador_auto=1), reflejar siempre al Despaletizador actual de la Pizarra
        if fecha == _today_ymd() and is_auto:
            live_op = pizarra_info["operador_sugerido"]
            new_grp = row["grupo"] or pizarra_info["grupo_activo"] or "Grupo A"
            if live_op != (row["operador"] or "") or new_grp != (row["grupo"] or ""):
                conn.execute(
                    """
                    UPDATE cuaderno_hojas
                    SET operador = ?, grupo = ?, updated_at = ?
                    WHERE id = ?
                    """,
                    (live_op, new_grp, _now_iso(), hoja_id),
                )
                row = conn.execute("SELECT * FROM cuaderno_hojas WHERE id = ?", (hoja_id,)).fetchone()

    return dict(row), pizarra_info


def sync_cuaderno_with_pizarra(db_lock: threading.RLock, broadcast_fn: Callable[[str, Dict[str, Any]], None]) -> None:
    """
    Se invoca cada vez que cambia la Pizarra de Envasado (Grupo A / Grupo B).
    Actualiza automáticamente el Despaletizador en las hojas activas de hoy (operador_auto = 1)
    y difunde por SSE al instante a teléfonos y PCs conectados.
    """
    today = _today_ymd()
    updated_payloads: List[Dict[str, Any]] = []
    try:
        with db_lock:
            with get_db() as conn:
                hojas_hoy = conn.execute(
                    "SELECT id, fecha, linea, turno, grupo, operador, operador_auto FROM cuaderno_hojas WHERE fecha = ?",
                    (today,),
                ).fetchall()
                for h in hojas_hoy:
                    auto_flag = int(h["operador_auto"] if h["operador_auto"] is not None else 1)
                    if auto_flag != 1:
                        continue
                    piz = resolve_pizarra_link(conn, h["linea"], h["grupo"] or "")
                    nuevo_op = piz["operador_sugerido"]
                    nuevo_grp = h["grupo"] or piz["grupo_activo"] or "Grupo A"
                    if nuevo_op != (h["operador"] or "") or nuevo_grp != (h["grupo"] or ""):
                        conn.execute(
                            "UPDATE cuaderno_hojas SET operador = ?, grupo = ?, updated_at = ? WHERE id = ?",
                            (nuevo_op, nuevo_grp, _now_iso(), h["id"]),
                        )
                        updated_payloads.append(build_cuaderno_payload(conn, h["fecha"], h["linea"], h["turno"], "Pizarra"))
        for p in updated_payloads:
            broadcast_fn(
                "cuaderno_actualizado",
                {
                    "hoja": p["hoja"],
                    "registros": p["registros"],
                    "pizarra": p["pizarra"],
                    "resumen": p["resumen"],
                    "server_time": p["server_time"],
                    "hora_planta": p["hora_planta"],
                },
            )
    except Exception:
        pass


def _reindex_hoja_items(conn: Any, hoja_id: str) -> None:
    rows = conn.execute(
        "SELECT id, num_item FROM cuaderno_registros WHERE hoja_id = ? ORDER BY num_item ASC, created_at ASC, id ASC",
        (hoja_id,),
    ).fetchall()
    updates = [(idx + 1, r["id"]) for idx, r in enumerate(rows) if int(r["num_item"] or 0) != idx + 1]
    if updates:
        conn.executemany("UPDATE cuaderno_registros SET num_item = ? WHERE id = ?", updates)


def build_cuaderno_payload(conn: Any, fecha: str, linea: str, turno: str, actor: str = "Sistema") -> Dict[str, Any]:
    hoja, pizarra_info = _ensure_or_load_hoja(conn, fecha, linea, turno, actor)
    reg_rows = conn.execute(
        """
        SELECT id, hoja_id, num_item, lote_sap, hora_desp, prefijo_proveedor, codigo_bull,
               fecha_prod, hora_prod, observacion, raw_barcode, foto_etiqueta, created_at, created_by
        FROM cuaderno_registros
        WHERE hoja_id = ?
        ORDER BY num_item ASC, created_at ASC
        """,
        (hoja["id"],),
    ).fetchall()

    registros: List[Dict[str, Any]] = []
    lotes_vistos: Set[str] = set()
    for r in reg_rows:
        raw_foto = r["foto_etiqueta"] or ""
        ver = _cache_tag_photo(r["id"], raw_foto) if raw_foto.startswith("data:image/") else ""
        if r["lote_sap"]:
            lotes_vistos.add(r["lote_sap"])
        registros.append(
            {
                "id": r["id"],
                "hoja_id": r["hoja_id"],
                "num_item": int(r["num_item"] or 1),
                "lote_sap": r["lote_sap"] or "",
                "hora_desp": r["hora_desp"] or "",
                "prefijo_proveedor": r["prefijo_proveedor"] or "LT",
                "codigo_bull": r["codigo_bull"] or "",
                "fecha_prod": r["fecha_prod"] or "",
                "hora_prod": r["hora_prod"] or "",
                "observacion": r["observacion"] or "",
                "raw_barcode": r["raw_barcode"] or "",
                "tiene_foto": bool(ver),
                "foto_url": f"/api/cuaderno/registros/{r['id']}/foto?v={ver}" if ver else "",
                "created_at": r["created_at"] or "",
                "created_by": r["created_by"] or "",
            }
        )

    resumen_dia_rows = conn.execute(
        """
        SELECT h.linea, h.turno, h.operador, COUNT(r.id) AS total_paletas
        FROM cuaderno_hojas h
        LEFT JOIN cuaderno_registros r ON r.hoja_id = h.id
        WHERE h.fecha = ?
        GROUP BY h.linea, h.turno, h.operador
        ORDER BY h.linea ASC, h.turno ASC
        """,
        (fecha,),
    ).fetchall()

    resumen_lineas: Dict[str, int] = {ln: 0 for ln in LINEAS_CUADERNO}
    hojas_dia: List[Dict[str, Any]] = []
    for rr in resumen_dia_rows:
        cnt = int(rr["total_paletas"] or 0)
        if rr["turno"] == turno and rr["linea"] in resumen_lineas:
            resumen_lineas[rr["linea"]] = cnt
        hojas_dia.append(
            {
                "linea": rr["linea"],
                "turno": rr["turno"],
                "operador": rr["operador"] or "",
                "total_paletas": cnt,
            }
        )

    ultimo_lote = registros[-1]["lote_sap"] if registros else ""
    ultima_hora = registros[-1]["hora_desp"] if registros else ""

    return {
        "hoja": hoja,
        "registros": registros,
        "pizarra": pizarra_info,
        "resumen": {
            "total_paletas": len(registros),
            "total_lotes_sap": len(lotes_vistos),
            "ultimo_lote_sap": ultimo_lote,
            "ultima_hora_desp": ultima_hora,
            "por_linea_turno": resumen_lineas,
            "hojas_dia": hojas_dia,
        },
        "catalogo_lineas": LINEA_DEFAULTS,
        "server_time": _now_iso(),
        "hora_planta": plant_now().strftime("%H:%M:%S"),
        "today": _today_ymd(),
    }


def _format_dmy_short(ymd: str) -> str:
    try:
        dt = datetime.strptime(ymd, "%Y-%m-%d")
        return dt.strftime("%d/%m/%y")
    except ValueError:
        return ymd


def _populate_fo003_worksheet(ws: Any, hoja: Dict[str, Any], registros: List[Dict[str, Any]]) -> None:
    """Genera hoja Excel idéntica al formato oficial en blanco FO003-RP06-PRO (100% modo claro)."""
    thin = Side(border_style="thin", color="334155")
    med = Side(border_style="medium", color="0F172A")
    b_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
    b_hdr = Border(left=thin, right=thin, top=med, bottom=med)
    b_bottom = Border(bottom=thin)

    fill_white = PatternFill("solid", fgColor="FFFFFF")
    fill_hdr = PatternFill("solid", fgColor="F1F5F9")
    fill_guide = PatternFill("solid", fgColor="F8FAFC")

    f_title = Font(name="Calibri", size=13, bold=True, color="0F172A")
    f_meta = Font(name="Calibri", size=9, bold=False, color="334155")
    f_lbl = Font(name="Calibri", size=10, bold=True, color="0F172A")
    f_val = Font(name="Calibri", size=10, bold=False, color="0F172A")
    f_th = Font(name="Calibri", size=9, bold=True, color="0F172A")
    f_td = Font(name="Calibri", size=10, bold=False, color="0F172A")
    f_mono = Font(name="Consolas", size=10, bold=True, color="0F172A")

    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells("A1:C2")
    ws["A1"] = "EMPRESAS POLAR\nPlanta San Pedro"
    ws["A1"].font = Font(name="Calibri", size=10, bold=True, color="0F172A")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.merge_cells("D1:L2")
    ws["D1"] = "REPORTE DE CONSUMO DE ENVASES EN PRODUCCIÓN"
    ws["D1"].font = f_title
    ws["D1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.merge_cells("M1:O1")
    ws["M1"] = "CÓDIGO: FO003 - RP06 - PRO · Versión: 1"
    ws["M1"].font = f_meta
    ws["M1"].alignment = Alignment(horizontal="right", vertical="center")

    ws.merge_cells("M2:O2")
    ws["M2"] = "Fecha actualización: 20/01/2025 · Planta San Pedro"
    ws["M2"].font = f_meta
    ws["M2"].alignment = Alignment(horizontal="right", vertical="center")

    for r_idx in (1, 2):
        for c_idx in range(1, 16):
            ws.cell(row=r_idx, column=c_idx).border = b_thin
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 8

    cfg = LINEA_DEFAULTS.get(hoja["linea"], LINEA_DEFAULTS["L1"])
    linea_lbl = cfg["nombre"]
    pres_actual = (hoja.get("presentacion") or "").strip()
    pres_items = list(cfg["presentaciones"])
    if pres_actual and not any(p["label"] == pres_actual for p in pres_items):
        pres_items.append({"label": pres_actual, "cod": ""})
    pres_list_txt = "   ".join(
        f"{'☑' if pres_actual == p['label'] else '☐'} {p['label']}{(' (' + p['cod'] + ')') if p['cod'] else ''}"
        for p in pres_items
    )

    prov_env = (hoja.get("proveedor_envase") or "AMCOR").upper()
    prov_amcor = "☑ AMCOR" if prov_env == "AMCOR" else "☐ AMCOR"
    prov_otro_val = (hoja.get("proveedor_otro") or "").strip()
    prov_otro = f"☑ OTRO: {prov_otro_val or '________'}" if prov_env == "OTRO" else "☐ OTRO"
    t1 = "☑ T1" if hoja["turno"] == "T1" else "☐ T1"
    t2 = "☑ T2" if hoja["turno"] == "T2" else "☐ T2"
    t3 = "☑ T3" if hoja["turno"] == "T3" else "☐ T3"

    ws["A4"] = "FECHA:"
    ws["A4"].font = f_lbl
    ws.merge_cells("B4:C4")
    ws["B4"] = _format_dmy_short(hoja["fecha"])
    ws["B4"].font = f_val
    ws["B4"].alignment = Alignment(horizontal="center", vertical="center")
    ws["B4"].border = b_bottom

    ws["D4"] = "OPERADOR:"
    ws["D4"].font = f_lbl
    ws.merge_cells("E4:H4")
    ws["E4"] = hoja.get("operador") or ""
    ws["E4"].font = f_val
    ws["E4"].border = b_bottom

    ws["I4"] = "RELEVO:"
    ws["I4"].font = f_lbl
    ws.merge_cells("J4:L4")
    ws["J4"] = hoja.get("relevo") or ""
    ws["J4"].font = f_val
    ws["J4"].border = b_bottom

    ws.merge_cells("M4:O4")
    ws["M4"] = f"{linea_lbl}   {hoja.get('codigo_linea') or 'Cod.B'}   ({hoja.get('grupo') or ''})"
    ws["M4"].font = f_lbl
    ws["M4"].alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[4].height = 22

    ws["A5"] = "PROVEEDOR:"
    ws["A5"].font = f_lbl
    ws.merge_cells("B5:D5")
    ws["B5"] = prov_amcor
    ws["B5"].font = f_val
    ws.merge_cells("E5:G5")
    ws["E5"] = prov_otro
    ws["E5"].font = f_val

    ws.merge_cells("H5:K5")
    ws["H5"] = f"{t1}     {t2}     {t3}"
    ws["H5"].font = f_lbl
    ws["H5"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("L5:O5")
    ws["L5"] = pres_list_txt
    ws["L5"].font = f_val
    ws["L5"].alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
    ws.row_dimensions[5].height = 26

    ws["A6"] = "MATERIAL SAP:"
    ws["A6"].font = f_lbl
    ws.merge_cells("B6:C6")
    ws["B6"] = hoja.get("material_sap") or ""
    ws["B6"].font = f_val
    ws["B6"].border = b_bottom

    ws["D6"] = "PALETAS:"
    ws["D6"].font = f_lbl
    ws["E6"] = len(registros)
    ws["E6"].font = f_lbl
    ws["E6"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("F6:K6")
    ws["F6"] = f"Línea: {linea_lbl} · {hoja.get('codigo_linea') or 'Cod.B'}"
    ws["F6"].font = f_meta
    ws["F6"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("L6:O6")
    sabor_txt = f"Sabor: {hoja.get('sabor') or '________'}" if cfg.get("tiene_sabor") or hoja.get("sabor") else ""
    ws["L6"] = sabor_txt
    ws["L6"].font = f_val
    ws["L6"].alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[6].height = 20

    # Cabecera de tabla oficial (Filas 7 y 8)
    ws.merge_cells("A7:A8")
    ws["A7"] = "N°"
    ws.merge_cells("B7:B8")
    ws["B7"] = "LOTE SAP"
    ws.merge_cells("C7:C8")
    ws["C7"] = "HORA DESP"
    ws.merge_cells("D7:L7")
    ws["D7"] = "PROVEEDOR"
    ws["D8"] = "LT / VL"

    guia = cfg["guia"]
    guia_chars = list(guia["bull"].ljust(8)[:8])
    for i in range(8):
        cell = ws.cell(row=8, column=5 + i, value=f"C{i + 1}")
        cell.font = f_th
        cell.fill = fill_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = b_thin

    ws.merge_cells("M7:M8")
    ws["M7"] = "FECHA PROD"
    ws.merge_cells("N7:N8")
    ws["N7"] = "HORA P"
    ws.merge_cells("O7:O8")
    ws["O7"] = "OBSERVACIÓN"

    for r_idx in (7, 8):
        ws.row_dimensions[r_idx].height = 20
        for c_idx in range(1, 16):
            c = ws.cell(row=r_idx, column=c_idx)
            c.font = f_th
            c.fill = fill_hdr
            c.border = b_hdr if r_idx == 7 else b_thin
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Fila 9: fila guía N° 0 exacta del formato en blanco
    guide_vals = [
        0,
        guia["lote"],
        guia["hora_desp"],
        guia["pref"],
        guia_chars[0],
        guia_chars[1],
        guia_chars[2],
        guia_chars[3],
        guia_chars[4],
        guia_chars[5],
        guia_chars[6],
        guia_chars[7],
        guia["fecha_prod"],
        guia["hora_p"],
        "Falla Eléctrica Externa",
    ]
    ws.row_dimensions[9].height = 19
    for c_idx, val in enumerate(guide_vals, start=1):
        cell = ws.cell(row=9, column=c_idx, value=val)
        cell.font = Font(name="Calibri", size=9, italic=True, color="475569")
        cell.fill = fill_guide
        cell.border = b_thin
        cell.alignment = Alignment(horizontal="center" if c_idx < 15 else "left", vertical="center")

    total_rows = max(38, len(registros))
    for idx in range(total_rows):
        excel_row = 10 + idx
        reg = registros[idx] if idx < len(registros) else None
        num_no = reg["num_item"] if reg else (idx + 1)

        bull_str = (reg["codigo_bull"] if reg else "") or ""
        bull_padded = bull_str[:8].ljust(8)
        row_cells = [
            num_no,
            reg["lote_sap"] if reg else "",
            reg["hora_desp"] if reg else "",
            reg["prefijo_proveedor"] if reg else "",
            bull_padded[0].strip(),
            bull_padded[1].strip(),
            bull_padded[2].strip(),
            bull_padded[3].strip(),
            bull_padded[4].strip(),
            bull_padded[5].strip(),
            bull_padded[6].strip(),
            bull_padded[7].strip(),
            reg["fecha_prod"] if reg else "",
            reg["hora_prod"] if reg else "",
            reg["observacion"] if reg else "",
        ]
        for c_idx, val in enumerate(row_cells, start=1):
            cell = ws.cell(row=excel_row, column=c_idx, value=val)
            cell.fill = fill_white
            cell.border = b_thin
            if 5 <= c_idx <= 12:
                cell.font = f_mono
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx == 15:
                cell.font = f_td
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.font = f_td
                cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[excel_row].height = 19

    # Pie oficial FO003: Analista de Calidad y Facilitador (T1 / T2 / T3)
    foot_r1 = 10 + total_rows + 2
    foot_r2 = foot_r1 + 2
    analista_val = hoja.get("analista_calidad") or ""
    facilitador_val = hoja.get("facilitador") or ""

    ws.merge_cells(start_row=foot_r1, start_column=1, end_row=foot_r1, end_column=3)
    ws.cell(row=foot_r1, column=1, value="Analista de Calidad:").font = f_lbl
    ws.merge_cells(start_row=foot_r1, start_column=4, end_row=foot_r1, end_column=7)
    c_at1 = ws.cell(row=foot_r1, column=4, value=f"T1: {analista_val if hoja['turno'] == 'T1' else ''}")
    c_at1.font = f_val
    c_at1.border = b_bottom
    ws.merge_cells(start_row=foot_r1, start_column=8, end_row=foot_r1, end_column=11)
    c_at2 = ws.cell(row=foot_r1, column=8, value=f"T2: {analista_val if hoja['turno'] == 'T2' else ''}")
    c_at2.font = f_val
    c_at2.border = b_bottom
    ws.merge_cells(start_row=foot_r1, start_column=12, end_row=foot_r1, end_column=15)
    c_at3 = ws.cell(row=foot_r1, column=12, value=f"T3: {analista_val if hoja['turno'] == 'T3' else ''}")
    c_at3.font = f_val
    c_at3.border = b_bottom

    ws.merge_cells(start_row=foot_r2, start_column=1, end_row=foot_r2, end_column=3)
    ws.cell(row=foot_r2, column=1, value="Facilitador:").font = f_lbl
    ws.merge_cells(start_row=foot_r2, start_column=4, end_row=foot_r2, end_column=7)
    c_ft1 = ws.cell(row=foot_r2, column=4, value=f"T1: {facilitador_val if hoja['turno'] == 'T1' else ''}")
    c_ft1.font = f_val
    c_ft1.border = b_bottom
    ws.merge_cells(start_row=foot_r2, start_column=8, end_row=foot_r2, end_column=11)
    c_ft2 = ws.cell(row=foot_r2, column=8, value=f"T2: {facilitador_val if hoja['turno'] == 'T2' else ''}")
    c_ft2.font = f_val
    c_ft2.border = b_bottom
    ws.merge_cells(start_row=foot_r2, start_column=12, end_row=foot_r2, end_column=15)
    c_ft3 = ws.cell(row=foot_r2, column=12, value=f"T3: {facilitador_val if hoja['turno'] == 'T3' else ''}")
    c_ft3.font = f_val
    c_ft3.border = b_bottom

    col_widths = {
        "A": 5.5,
        "B": 14.0,
        "C": 11.5,
        "D": 9.5,
        "E": 3.8,
        "F": 3.8,
        "G": 3.8,
        "H": 3.8,
        "I": 3.8,
        "J": 3.8,
        "K": 3.8,
        "L": 3.8,
        "M": 13.0,
        "N": 10.5,
        "O": 28.0,
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def generate_cuaderno_excel_bytes(conn: Any, fecha: str, linea: str, turno: str, scope: str = "hoja") -> Tuple[bytes, str]:
    wb = Workbook()
    default_ws = wb.active

    if scope == "dia":
        hojas_rows = conn.execute(
            "SELECT * FROM cuaderno_hojas WHERE fecha = ? ORDER BY linea ASC, turno ASC",
            (fecha,),
        ).fetchall()
        hojas_map: Dict[Tuple[str, str], Tuple[Dict[str, Any], List[Dict[str, Any]]]] = {}
        for h in hojas_rows:
            h_dict = dict(h)
            regs = [
                dict(r)
                for r in conn.execute(
                    "SELECT * FROM cuaderno_registros WHERE hoja_id = ? ORDER BY num_item ASC, created_at ASC",
                    (h_dict["id"],),
                ).fetchall()
            ]
            hojas_map[(h_dict["linea"], h_dict["turno"])] = (h_dict, regs)

        # Garantizar que todas las líneas (L1, L2, L3, L4) y sus presentaciones estén presentes en el Excel del día
        for ln in LINEAS_CUADERNO:
            if (ln, turno) not in hojas_map:
                cfg = LINEA_DEFAULTS[ln]
                piz = resolve_pizarra_link(conn, ln, "")
                hojas_map[(ln, turno)] = (
                    {
                        "id": make_hoja_id(fecha, ln, turno),
                        "fecha": fecha,
                        "linea": ln,
                        "turno": turno,
                        "grupo": piz.get("grupo_activo") or "Grupo A",
                        "codigo_linea": cfg["codigo"],
                        "operador": piz.get("operador_sugerido") or "",
                        "operador_auto": 1,
                        "relevo": piz.get("relevo_sugerido") or "",
                        "proveedor_envase": "AMCOR",
                        "proveedor_otro": "",
                        "presentacion": cfg["default_pres"],
                        "sabor": "",
                        "material_sap": "",
                        "analista_calidad": "",
                        "facilitador": "",
                    },
                    [],
                )

        ordered_keys = sorted(
            hojas_map.keys(),
            key=lambda k: (
                LINEAS_CUADERNO.index(k[0]) if k[0] in LINEAS_CUADERNO else 99,
                TURNOS_CUADERNO.index(k[1]) if k[1] in TURNOS_CUADERNO else 99,
            ),
        )
        first = True
        for key in ordered_keys:
            h_dict, regs = hojas_map[key]
            ws = default_ws if first else wb.create_sheet()
            first = False
            ws.title = f"{h_dict['linea']}_{h_dict['turno']}"
            _populate_fo003_worksheet(ws, h_dict, regs)
        fname = f"FO003_Consumo_Envases_{fecha}_Completo.xlsx"
    else:
        payload = build_cuaderno_payload(conn, fecha, linea, turno)
        ws = default_ws
        ws.title = f"{linea}_{turno}"
        _populate_fo003_worksheet(ws, payload["hoja"], payload["registros"])
        fname = f"FO003_Consumo_Envases_{fecha}_{linea}_{turno}.xlsx"

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), fname


def parse_imported_cuaderno_rows(filename: str, raw_bytes: bytes, target_sheet: str = "") -> List[Dict[str, str]]:
    fname = (filename or "").lower()
    matrix: List[List[str]] = []
    if fname.endswith((".xlsx", ".xlsm")):
        wb = load_workbook(io.BytesIO(raw_bytes), data_only=True, read_only=True)
        try:
            ws = wb[target_sheet] if target_sheet and target_sheet in wb.sheetnames else wb.active
            for row in ws.iter_rows(max_row=1000, max_col=25, values_only=True):
                matrix.append([_clean_str(c, 160) for c in row])
        finally:
            wb.close()
    else:
        try:
            text = raw_bytes.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = raw_bytes.decode("latin-1", errors="replace")
        sample = text[:2048]
        delim = "\t" if sample.count("\t") > sample.count(",") and sample.count("\t") >= sample.count(";") else (";" if sample.count(";") > sample.count(",") else ",")
        for row in csv.reader(io.StringIO(text), delimiter=delim):
            matrix.append([_clean_str(c, 160) for c in row])

    parsed: List[Dict[str, str]] = []
    header_idx = -1
    col_map: Dict[str, int] = {}

    for r_i, row in enumerate(matrix[:30]):
        joined = " ".join(row).lower()
        if ("lote" in joined or "bull" in joined or "lt" in joined or "vl" in joined) and (
            "hora" in joined or "proveedor" in joined or "bull" in joined or "desp" in joined
        ):
            header_idx = r_i
            for c_i, cell in enumerate(row):
                cl = cell.lower()
                if ("lote sap" in cl or ("lote" in cl and "bull" not in cl and "no." not in cl)) and "lote_sap" not in col_map:
                    col_map["lote_sap"] = c_i
                elif ("hora desp" in cl or "desp" in cl) and "hora_desp" not in col_map:
                    col_map["hora_desp"] = c_i
                elif ("lt / vl" in cl or "lt/vl" in cl or "proveedor" in cl or "bull" in cl or "no. de lote" in cl) and "proveedor" not in col_map:
                    col_map["proveedor"] = c_i
                elif "fecha prod" in cl and "fecha_prod" not in col_map:
                    col_map["fecha_prod"] = c_i
                elif ("hora p" in cl or "hora prod" in cl) and "hora_prod" not in col_map:
                    col_map["hora_prod"] = c_i
                elif "observ" in cl and "observacion" not in col_map:
                    col_map["observacion"] = c_i
            break

    is_fo003_15col = (
        header_idx >= 0
        and col_map.get("lote_sap") == 1
        and col_map.get("hora_desp") == 2
        and col_map.get("proveedor") == 3
        and col_map.get("fecha_prod") == 12
    )

    start_row = header_idx + 1 if header_idx >= 0 else 0
    for row in matrix[start_row:]:
        if not any(row):
            continue
        first_col = (row[0] if len(row) > 0 else "").strip()
        first_low = first_col.lower()
        if first_low.startswith(("analista", "facilitador", "nota")):
            break
        col3_up = (row[3] if len(row) > 3 else "").strip().upper()
        col4_up = (row[4] if len(row) > 4 else "").strip().upper()
        if first_col in ("0", "N°", "Nº", "LT / VL", "D1", "C1") or col3_up in ("LT / VL", "LT/VL", "PROVEEDOR") or col4_up in ("D1", "C1"):
            continue
        if is_fo003_15col and len(row) >= 12:
            lote = row[1]
            h_desp = row[2]
            pref = _normalize_prefix(row[3], "LT")
            bull = "".join(ch for ch in "".join(row[4:12]).upper() if ch.isalnum())[:8]
            f_prod = row[12] if len(row) > 12 else ""
            h_prod = row[13] if len(row) > 13 else ""
            obs = row[14] if len(row) > 14 else ""
        elif col_map:
            lote = row[col_map["lote_sap"]] if "lote_sap" in col_map and col_map["lote_sap"] < len(row) else ""
            h_desp = row[col_map["hora_desp"]] if "hora_desp" in col_map and col_map["hora_desp"] < len(row) else ""
            raw_prov = row[col_map["proveedor"]] if "proveedor" in col_map and col_map["proveedor"] < len(row) else ""
            info = parse_bull_barcode(raw_prov)
            pref = info["prefijo_proveedor"]
            bull = info["codigo_bull"] or _clean_bull_code(raw_prov)
            f_prod = row[col_map["fecha_prod"]] if "fecha_prod" in col_map and col_map["fecha_prod"] < len(row) else ""
            h_prod = row[col_map["hora_prod"]] if "hora_prod" in col_map and col_map["hora_prod"] < len(row) else ""
            obs = row[col_map["observacion"]] if "observacion" in col_map and col_map["observacion"] < len(row) else ""
        else:
            joined_row = " ".join(row)
            info = parse_bull_barcode(joined_row)
            if not info["codigo_bull"]:
                continue
            lote = row[1] if len(row) > 1 else ""
            h_desp = row[2] if len(row) > 2 else ""
            pref = info["prefijo_proveedor"]
            bull = info["codigo_bull"]
            f_prod = info["fecha_prod"]
            h_prod = info["hora_prod"]
            obs = row[-1] if len(row) > 4 else ""

        if not bull and not lote and not obs:
            continue
        if bull in ("91916156", "83505072", "G3228973") and "falla" in obs.lower() and first_col == "0":
            continue
        parsed.append(
            {
                "lote_sap": lote[:40],
                "hora_desp": h_desp[:16] or plant_now().strftime("%H:%M:%S"),
                "prefijo_proveedor": _normalize_prefix(pref, "LT"),
                "codigo_bull": bull[:8],
                "fecha_prod": f_prod[:20],
                "hora_prod": h_prod[:16],
                "observacion": obs[:180],
            }
        )
    return parsed


def create_cuaderno_blueprint(
    db_lock: threading.RLock,
    ensure_ready_fn: Callable[[], None],
    broadcast_fn: Callable[[str, Dict[str, Any]], None],
    actor_fn: Callable[[Optional[Dict[str, Any]]], str],
) -> Blueprint:
    bp = Blueprint("cuaderno", __name__)

    def _resolve_params(source: Optional[Dict[str, Any]] = None) -> Tuple[Optional[Tuple[str, str, str]], Optional[str]]:
        src = source or {}
        fecha = str(src.get("fecha") or request.args.get("fecha") or _today_ymd()).strip()
        ok_f, err_f = validar_fecha_logica(fecha)
        if not ok_f:
            return None, err_f or "Fecha inválida."
        linea = str(src.get("linea") or request.args.get("linea") or "L1").strip().upper()
        if linea not in LINEAS_CUADERNO:
            return None, "Línea inválida. Usa L1, L2, L3 o L4."
        turno = str(src.get("turno") or request.args.get("turno") or _infer_current_turno()).strip().upper()
        if turno not in TURNOS_CUADERNO:
            return None, "Turno inválido. Usa T1, T2 o T3."
        return (fecha, linea, turno), None

    def _broadcast_sheet(payload: Dict[str, Any]) -> None:
        broadcast_fn(
            "cuaderno_actualizado",
            {
                "hoja": payload["hoja"],
                "registros": payload["registros"],
                "pizarra": payload["pizarra"],
                "resumen": payload["resumen"],
                "server_time": payload["server_time"],
                "hora_planta": payload["hora_planta"],
            },
        )

    @bp.route("/cuaderno", methods=["GET"])
    def serve_cuaderno():
        ensure_ready_fn()
        params, err = _resolve_params()
        if err or not params:
            fecha, linea, turno = _today_ymd(), "L1", _infer_current_turno()
        else:
            fecha, linea, turno = params

        actor = actor_fn(None)
        with get_db() as conn:
            initial_payload = build_cuaderno_payload(conn, fecha, linea, turno, actor)

        login_ts = session.get("login_ts", 0)
        expira_en_s = max(0, int(SESION_MAX_S - (time.time() - login_ts))) if login_ts else 0
        user = session.get("user") or {}
        rol = str(user.get("rol") or "supervisor").strip().lower()

        can_admin = rol in ROLES_CUADERNO_ADMIN
        can_view_pizarra = rol != "escaner"
        is_read_only = rol == "tv"

        initial_payload["current_user"] = user
        initial_payload["actor_name"] = actor
        initial_payload["can_admin_cuaderno"] = can_admin
        initial_payload["can_view_pizarra"] = can_view_pizarra
        initial_payload["is_read_only"] = is_read_only
        initial_payload["expira_en_s"] = expira_en_s
        initial_payload["csrf_token"] = get_csrf_token()

        ssr_json = (
            json.dumps(initial_payload, ensure_ascii=False, separators=(",", ":"))
            .replace("<", "\\u003c")
            .replace(">", "\\u003e")
            .replace("&", "\\u0026")
        )
        nonce = get_csp_nonce()
        html = render_template(
            "cuaderno.html",
            csp_nonce=nonce,
            csrf_token=get_csrf_token(),
            can_admin=can_admin,
            can_view_pizarra=can_view_pizarra,
            is_read_only=is_read_only,
        )
        html = html.replace(
            f'<script id="ssr-cuaderno-state" type="application/json" nonce="{nonce}">{{}}</script>',
            f'<script id="ssr-cuaderno-state" type="application/json" nonce="{nonce}">{ssr_json}</script>',
            1,
        )
        resp = Response(html, mimetype="text/html; charset=utf-8")
        resp.headers["Cache-Control"] = "private, no-cache, must-revalidate"
        resp.headers["Permissions-Policy"] = (
            "camera=(self), geolocation=(), microphone=(), payment=(), usb=(), interest-cohort=()"
        )
        return resp

    @bp.route("/api/cuaderno/hoja", methods=["GET"])
    def api_get_hoja():
        ensure_ready_fn()
        params, err = _resolve_params()
        if err or not params:
            return jsonify({"error": err}), 400
        fecha, linea, turno = params
        actor = actor_fn(None)
        with get_db() as conn:
            payload = build_cuaderno_payload(conn, fecha, linea, turno, actor)
        return jsonify(payload)

    @bp.route("/api/cuaderno/hoja", methods=["POST", "PATCH"])
    def api_save_hoja():
        ensure_ready_fn()
        data = request.get_json(silent=True) or {}
        params, err = _resolve_params(data)
        if err or not params:
            return jsonify({"error": err}), 400
        fecha, linea, turno = params
        actor = actor_fn(data)
        sync_from_pizarra = bool(data.get("sincronizar_pizarra"))

        with db_lock:
            with get_db() as conn:
                hoja, pizarra_info = _ensure_or_load_hoja(conn, fecha, linea, turno, actor)
                grupo_prev = hoja["grupo"] or ""
                grupo = _clean_str(data.get("grupo", grupo_prev), 40)
                grupo_changed = ("grupo" in data) and (grupo != grupo_prev)

                if sync_from_pizarra or grupo_changed:
                    pizarra_info = resolve_pizarra_link(conn, linea, grupo)
                    operador = pizarra_info["operador_sugerido"]
                    relevo = pizarra_info["relevo_sugerido"] or _clean_str(data.get("relevo", hoja["relevo"]), 80)
                    operador_auto = 1
                    if not grupo and pizarra_info["grupo_activo"]:
                        grupo = pizarra_info["grupo_activo"]
                else:
                    operador = _clean_str(data.get("operador", hoja["operador"]), 80)
                    relevo = _clean_str(data.get("relevo", hoja["relevo"]), 80)
                    if "operador_auto" in data:
                        operador_auto = 1 if bool(data.get("operador_auto")) else 0
                    elif "operador" in data and operador != (hoja["operador"] or ""):
                        pizarra_info = resolve_pizarra_link(conn, linea, grupo)
                        operador_auto = 1 if operador == pizarra_info["operador_sugerido"] else 0
                    else:
                        operador_auto = int(hoja.get("operador_auto", 1))

                codigo_linea = _clean_str(data.get("codigo_linea", hoja["codigo_linea"]), 24)
                prov = _clean_str(data.get("proveedor_envase", hoja["proveedor_envase"]), 20).upper()
                if prov not in ("AMCOR", "OTRO"):
                    prov = "AMCOR"
                prov_otro = _clean_str(data.get("proveedor_otro", hoja["proveedor_otro"]), 60) if prov == "OTRO" else ""
                presentacion = _clean_str(data.get("presentacion", hoja["presentacion"]), 30)
                sabor = _clean_str(data.get("sabor", hoja["sabor"]), 40)
                material_sap = _clean_str(data.get("material_sap", hoja["material_sap"]), 30)
                analista_calidad = _clean_str(data.get("analista_calidad", hoja.get("analista_calidad", "")), 80)
                facilitador = _clean_str(data.get("facilitador", hoja.get("facilitador", "")), 80)

                conn.execute(
                    """
                    UPDATE cuaderno_hojas
                    SET grupo = ?, codigo_linea = ?, operador = ?, operador_auto = ?, relevo = ?,
                        proveedor_envase = ?, proveedor_otro = ?, presentacion = ?,
                        sabor = ?, material_sap = ?, analista_calidad = ?, facilitador = ?,
                        updated_at = ?, updated_by = ?
                    WHERE id = ?
                    """,
                    (
                        grupo,
                        codigo_linea,
                        operador,
                        operador_auto,
                        relevo,
                        prov,
                        prov_otro,
                        presentacion,
                        sabor,
                        material_sap,
                        analista_calidad,
                        facilitador,
                        _now_iso(),
                        actor,
                        hoja["id"],
                    ),
                )
                payload = build_cuaderno_payload(conn, fecha, linea, turno, actor)

        _broadcast_sheet(payload)
        return jsonify(payload)

    @bp.route("/api/cuaderno/registros", methods=["POST"])
    def api_add_registro():
        ensure_ready_fn()
        data = request.get_json(silent=True) or {}
        params, err = _resolve_params(data)
        if err or not params:
            return jsonify({"error": err}), 400
        fecha, linea, turno = params
        actor = actor_fn(data)

        raw_barcode = _clean_str(data.get("raw_barcode", ""), 400)
        lote_sap = _clean_str(data.get("lote_sap", ""), 40)
        hora_desp = _clean_str(data.get("hora_desp", ""), 16) or plant_now().strftime("%H:%M:%S")
        observacion = _clean_str(data.get("observacion", ""), 180)
        foto_etiqueta = str(data.get("foto_etiqueta") or "").strip()
        allow_dup = bool(data.get("permitir_duplicado"))

        ok_img, msg_img = validar_imagen_data_url(foto_etiqueta)
        if not ok_img:
            return jsonify({"error": msg_img or "Foto de etiqueta inválida."}), 400

        reg_id = uuid.uuid4().hex[:12]
        with db_lock:
            with get_db() as conn:
                hoja, _ = _ensure_or_load_hoja(conn, fecha, linea, turno, actor)
                hoja_id = hoja["id"]
                default_pref = "VL" if (hoja.get("proveedor_envase") or "").upper() == "OTRO" else "LT"
                parsed = parse_bull_barcode(raw_barcode, default_pref=default_pref) if raw_barcode else {
                    "prefijo_proveedor": default_pref,
                    "codigo_bull": "",
                    "fecha_prod": "",
                    "hora_prod": "",
                    "material_sap": "",
                }

                pref = _normalize_prefix(data.get("prefijo_proveedor") or parsed["prefijo_proveedor"], default_pref)

                bull_explicit = _clean_str(data.get("codigo_bull", ""), 20)
                if bull_explicit:
                    if bull_explicit.upper().startswith(("LT", "VL", "OT")) and len(bull_explicit) >= 8:
                        pref = bull_explicit[:2].upper()
                    codigo_bull = _clean_bull_code(bull_explicit)
                else:
                    codigo_bull = _clean_bull_code(parsed["codigo_bull"])

                if not codigo_bull:
                    return jsonify({"error": "Escanea o ingresa el No. de Lote / Bull de la etiqueta (ej. LTG3228973, VL62477447 u OTRO)."}), 400

                fecha_prod = _clean_str(data.get("fecha_prod") or parsed["fecha_prod"], 20)
                hora_prod = _clean_str(data.get("hora_prod") or parsed["hora_prod"], 16)

                if not allow_dup:
                    dup = conn.execute(
                        "SELECT num_item, hora_desp FROM cuaderno_registros WHERE hoja_id = ? AND codigo_bull = ? LIMIT 1",
                        (hoja_id, codigo_bull),
                    ).fetchone()
                    if dup:
                        return jsonify(
                            {
                                "error": f"La paleta {pref}-{codigo_bull} ya fue escaneada en la fila N° {dup['num_item']} ({dup['hora_desp']}).",
                                "duplicado": True,
                            }
                        ), 409

                max_row = conn.execute(
                    "SELECT COALESCE(MAX(num_item), 0) AS m FROM cuaderno_registros WHERE hoja_id = ?",
                    (hoja_id,),
                ).fetchone()
                next_num = int((max_row["m"] if max_row and max_row["m"] is not None else 0)) + 1

                if not lote_sap:
                    prev_lote_row = conn.execute(
                        "SELECT lote_sap FROM cuaderno_registros WHERE hoja_id = ? AND lote_sap != '' ORDER BY num_item DESC LIMIT 1",
                        (hoja_id,),
                    ).fetchone()
                    if prev_lote_row:
                        lote_sap = prev_lote_row["lote_sap"] or ""

                now_s = _now_iso()
                conn.execute(
                    """
                    INSERT INTO cuaderno_registros (
                        id, hoja_id, num_item, lote_sap, hora_desp, prefijo_proveedor, codigo_bull,
                        fecha_prod, hora_prod, observacion, raw_barcode, foto_etiqueta, created_at, created_by
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        reg_id,
                        hoja_id,
                        next_num,
                        lote_sap,
                        hora_desp,
                        pref,
                        codigo_bull,
                        fecha_prod,
                        hora_prod,
                        observacion,
                        raw_barcode,
                        foto_etiqueta,
                        now_s,
                        actor,
                    ),
                )
                if parsed["material_sap"] and not hoja.get("material_sap"):
                    conn.execute(
                        "UPDATE cuaderno_hojas SET material_sap = ?, updated_at = ?, updated_by = ? WHERE id = ?",
                        (parsed["material_sap"], now_s, actor, hoja_id),
                    )
                else:
                    conn.execute(
                        "UPDATE cuaderno_hojas SET updated_at = ?, updated_by = ? WHERE id = ?",
                        (now_s, actor, hoja_id),
                    )
                payload = build_cuaderno_payload(conn, fecha, linea, turno, actor)

        _broadcast_sheet(payload)
        payload["registro_creado_id"] = reg_id
        return jsonify(payload), 201

    @bp.route("/api/cuaderno/registros/<reg_id>", methods=["PATCH"])
    def api_edit_registro(reg_id: str):
        ensure_ready_fn()
        if not is_valid_op_id(reg_id):
            return jsonify({"error": "ID de registro inválido."}), 400
        data = request.get_json(silent=True) or {}
        actor = actor_fn(data)
        rol = str(((session.get("user") or {}).get("rol")) or "").strip().lower()

        with db_lock:
            with get_db() as conn:
                reg = conn.execute("SELECT * FROM cuaderno_registros WHERE id = ?", (reg_id,)).fetchone()
                if not reg:
                    return jsonify({"error": "Registro de paleta no encontrado."}), 404
                if rol == "escaner" and reg["created_by"] and reg["created_by"] != actor:
                    return jsonify({"error": "Solo un supervisor puede editar registros cargados por otro usuario."}), 403

                hoja = conn.execute("SELECT * FROM cuaderno_hojas WHERE id = ?", (reg["hoja_id"],)).fetchone()
                if not hoja:
                    return jsonify({"error": "Hoja asociada no encontrada."}), 404

                pref = _normalize_prefix(data.get("prefijo_proveedor", reg["prefijo_proveedor"]), "LT")
                bull_raw = _clean_str(data.get("codigo_bull", reg["codigo_bull"]), 20)
                codigo_bull = _clean_bull_code(bull_raw)
                if not codigo_bull:
                    return jsonify({"error": "El código Bull no puede quedar vacío."}), 400

                lote_sap = _clean_str(data.get("lote_sap", reg["lote_sap"]), 40)
                hora_desp = _clean_str(data.get("hora_desp", reg["hora_desp"]), 16)
                fecha_prod = _clean_str(data.get("fecha_prod", reg["fecha_prod"]), 20)
                hora_prod = _clean_str(data.get("hora_prod", reg["hora_prod"]), 16)
                observacion = _clean_str(data.get("observacion", reg["observacion"]), 180)

                conn.execute(
                    """
                    UPDATE cuaderno_registros
                    SET lote_sap = ?, hora_desp = ?, prefijo_proveedor = ?, codigo_bull = ?,
                        fecha_prod = ?, hora_prod = ?, observacion = ?
                    WHERE id = ?
                    """,
                    (lote_sap, hora_desp, pref, codigo_bull, fecha_prod, hora_prod, observacion, reg_id),
                )
                conn.execute(
                    "UPDATE cuaderno_hojas SET updated_at = ?, updated_by = ? WHERE id = ?",
                    (_now_iso(), actor, hoja["id"]),
                )
                payload = build_cuaderno_payload(conn, hoja["fecha"], hoja["linea"], hoja["turno"], actor)

        _broadcast_sheet(payload)
        return jsonify(payload)

    @bp.route("/api/cuaderno/registros/<reg_id>", methods=["DELETE"])
    def api_delete_registro(reg_id: str):
        ensure_ready_fn()
        if not is_valid_op_id(reg_id):
            return jsonify({"error": "ID de registro inválido."}), 400
        actor = actor_fn(None)
        rol = str(((session.get("user") or {}).get("rol")) or "").strip().lower()

        with db_lock:
            with get_db() as conn:
                reg = conn.execute("SELECT * FROM cuaderno_registros WHERE id = ?", (reg_id,)).fetchone()
                if not reg:
                    return jsonify({"error": "Registro no encontrado."}), 404
                if rol == "escaner" and reg["created_by"] and reg["created_by"] != actor:
                    return jsonify({"error": "Solo un supervisor puede eliminar paletas registradas por otro operador."}), 403

                hoja = conn.execute("SELECT * FROM cuaderno_hojas WHERE id = ?", (reg["hoja_id"],)).fetchone()
                conn.execute("DELETE FROM cuaderno_registros WHERE id = ?", (reg_id,))
                with _tag_photo_lock:
                    _tag_photo_cache.pop(reg_id, None)
                _reindex_hoja_items(conn, reg["hoja_id"])
                if not hoja:
                    return jsonify({"ok": True})
                conn.execute(
                    "UPDATE cuaderno_hojas SET updated_at = ?, updated_by = ? WHERE id = ?",
                    (_now_iso(), actor, hoja["id"]),
                )
                payload = build_cuaderno_payload(conn, hoja["fecha"], hoja["linea"], hoja["turno"], actor)

        _broadcast_sheet(payload)
        return jsonify(payload)

    @bp.route("/api/cuaderno/registros/<reg_id>/foto", methods=["GET"])
    def api_get_registro_foto(reg_id: str):
        if not is_valid_op_id(reg_id):
            return jsonify({"error": "ID inválido."}), 400
        req_v = (request.args.get("v") or "").strip()
        with _tag_photo_lock:
            cached = _tag_photo_cache.get(reg_id)
        if not cached or (req_v and cached[0] != req_v):
            ensure_ready_fn()
            with get_db() as conn:
                row = conn.execute("SELECT foto_etiqueta FROM cuaderno_registros WHERE id = ?", (reg_id,)).fetchone()
            if not row or not (row["foto_etiqueta"] or "").strip():
                return jsonify({"error": "Foto no disponible."}), 404
            ver = _cache_tag_photo(reg_id, row["foto_etiqueta"])
            if not ver:
                return jsonify({"error": "Foto inválida."}), 404
            with _tag_photo_lock:
                cached = _tag_photo_cache.get(reg_id)
        if not cached:
            return jsonify({"error": "Foto no disponible."}), 404
        ver_hash, mime, img_bytes, etag = cached
        if etag_matches(request.headers.get("If-None-Match"), etag):
            resp = Response(status=304)
            resp.headers["ETag"] = etag
            return resp
        resp = Response(img_bytes, mimetype=mime)
        resp.headers["ETag"] = etag
        resp.headers["Cache-Control"] = "private, max-age=31536000, immutable" if req_v == ver_hash else "private, no-cache"
        return resp

    @bp.route("/api/cuaderno/exportar", methods=["GET"])
    def api_exportar_excel():
        ensure_ready_fn()
        params, err = _resolve_params()
        if err or not params:
            return jsonify({"error": err}), 400
        fecha, linea, turno = params
        scope = (request.args.get("scope") or "hoja").strip().lower()
        with get_db() as conn:
            raw_xlsx, fname = generate_cuaderno_excel_bytes(conn, fecha, linea, turno, scope=scope)

        etag = compute_etag(raw_xlsx, prefix="fo003")
        if etag_matches(request.headers.get("If-None-Match"), etag):
            resp = Response(status=304)
            resp.headers["ETag"] = etag
            return resp
        resp = send_file(
            io.BytesIO(raw_xlsx),
            as_attachment=True,
            download_name=fname,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        resp.headers["ETag"] = etag
        resp.headers["Cache-Control"] = "private, no-cache, must-revalidate"
        return resp

    @bp.route("/api/cuaderno/importar", methods=["POST"])
    def api_importar_excel():
        ensure_ready_fn()
        modo = str(request.form.get("modo") or request.args.get("modo") or "agregar").strip().lower()
        if modo not in ("agregar", "reemplazar"):
            modo = "agregar"
        form_dict = {
            "fecha": request.form.get("fecha"),
            "linea": request.form.get("linea"),
            "turno": request.form.get("turno"),
        }
        params, err = _resolve_params(form_dict)
        if err or not params:
            return jsonify({"error": err}), 400
        fecha, linea, turno = params
        actor = actor_fn(None)

        if "file" not in request.files or not request.files["file"].filename:
            return jsonify({"error": "Selecciona un archivo Excel (.xlsx) o CSV para importar."}), 400
        f_obj = request.files["file"]
        fname = f_obj.filename or ""
        raw_bytes = f_obj.read()
        ok_f, err_f = validar_archivo_importacion(fname, raw_bytes)
        if not ok_f:
            return jsonify({"error": err_f or "Archivo inválido."}), 400

        try:
            rows_in = parse_imported_cuaderno_rows(fname, raw_bytes, target_sheet=f"{linea}_{turno}")
        except Exception as exc:
            return jsonify({"error": f"No se pudo leer el archivo: {exc}"}), 400

        if not rows_in:
            return jsonify({"error": "No se encontraron filas válidas de consumo de envases en el archivo."}), 400

        with db_lock:
            with get_db() as conn:
                hoja, _ = _ensure_or_load_hoja(conn, fecha, linea, turno, actor)
                hoja_id = hoja["id"]
                if modo == "reemplazar":
                    conn.execute("DELETE FROM cuaderno_registros WHERE hoja_id = ?", (hoja_id,))
                    next_num = 1
                else:
                    m_row = conn.execute(
                        "SELECT COALESCE(MAX(num_item), 0) AS m FROM cuaderno_registros WHERE hoja_id = ?",
                        (hoja_id,),
                    ).fetchone()
                    next_num = int((m_row["m"] if m_row and m_row["m"] is not None else 0)) + 1

                now_s = _now_iso()
                batch_rows = []
                for item in rows_in[:500]:
                    batch_rows.append(
                        (
                            uuid.uuid4().hex[:12],
                            hoja_id,
                            next_num,
                            item["lote_sap"],
                            item["hora_desp"],
                            item["prefijo_proveedor"],
                            item["codigo_bull"],
                            item["fecha_prod"],
                            item["hora_prod"],
                            item["observacion"],
                            "",
                            "",
                            now_s,
                            actor,
                        )
                    )
                    next_num += 1

                if batch_rows:
                    conn.executemany(
                        """
                        INSERT INTO cuaderno_registros (
                            id, hoja_id, num_item, lote_sap, hora_desp, prefijo_proveedor, codigo_bull,
                            fecha_prod, hora_prod, observacion, raw_barcode, foto_etiqueta, created_at, created_by
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        batch_rows,
                    )
                conn.execute(
                    "UPDATE cuaderno_hojas SET updated_at = ?, updated_by = ? WHERE id = ?",
                    (now_s, actor, hoja_id),
                )
                payload = build_cuaderno_payload(conn, fecha, linea, turno, actor)

        _broadcast_sheet(payload)
        payload["importados"] = len(batch_rows)
        return jsonify(payload)

    return bp
