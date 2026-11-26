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
import os
import re
import threading
import time
import uuid
import xml.etree.ElementTree as ET
import zipfile
from copy import copy
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

SAP_MATERIAL_CATALOG: Dict[str, Dict[str, str]] = {
    "50034396": {"linea": "L3", "presentacion": "5 Litros (5L)", "proveedor_envase": "AMCOR", "proveedor_otro": ""},
    "50033037": {"linea": "L2", "presentacion": "1500 ml", "proveedor_envase": "AMCOR", "proveedor_otro": ""},
}

_PRES_DETECT_RULES: Tuple[Tuple[re.Pattern, str, str], ...] = (
    (re.compile(r"(?:^|[^0-9])(?:1[.,]5\s*(?:LT|L)|1500\s*ML)(?=$|[^A-Z0-9])", re.IGNORECASE), "L2", "1500 ml"),
    (re.compile(r"(?:^|[^0-9])(?:5\s*(?:LT|L|LITROS?)|5000\s*ML)(?=$|[^A-Z0-9])", re.IGNORECASE), "L3", "5 Litros (5L)"),
    (re.compile(r"(?:^|[^0-9])(?:1(?:[.,]0)?\s*(?:LT|L|LITRO)|1000\s*ML)(?=$|[^A-Z0-9])", re.IGNORECASE), "L4", "1000 ml"),
    (re.compile(r"(?:^|[^0-9])(?:600\s*ML|0[.,]6\s*(?:LT|L))(?=$|[^A-Z0-9])", re.IGNORECASE), "L2", "600 ml"),
    (re.compile(r"(?:^|[^0-9])(?:500\s*ML|0[.,]5\s*(?:LT|L))(?=$|[^A-Z0-9])", re.IGNORECASE), "L4", "500 ml"),
    (re.compile(r"(?:^|[^0-9])(?:355\s*ML|0[.,]355\s*(?:LT|L))(?=$|[^A-Z0-9])", re.IGNORECASE), "L1", "355 ml"),
)
_SUPPLIER_AMCOR_RE = re.compile(r"\bAMCOR\b", re.IGNORECASE)
_SUPPLIER_OTRO_RE = re.compile(r"\b(VENVIDRIO|PLASTIPAK|ENVASES(?:\s+\w+)?|SMURFIT|ALPLA|CROWN|OWENS)\b", re.IGNORECASE)

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
    conn.execute("CREATE INDEX IF NOT EXISTS idx_cuaderno_reg_bull_global ON cuaderno_registros (codigo_bull)")


def _find_duplicate_bull_error(conn: Any, codigo_bull: str, fecha: str, linea: str, turno: str, exclude_reg_id: str = "") -> str:
    """Devuelve mensaje de error si el codigo Bull/paleta ya fue escaneado en cualquier hoja."""
    if not codigo_bull:
        return ""
    if exclude_reg_id:
        dup = conn.execute(
            """
            SELECT r.num_item, r.hora_desp, r.prefijo_proveedor, h.fecha, h.linea, h.turno
            FROM cuaderno_registros r
            JOIN cuaderno_hojas h ON h.id = r.hoja_id
            WHERE r.codigo_bull = ? AND r.id != ?
            LIMIT 1
            """,
            (codigo_bull, exclude_reg_id),
        ).fetchone()
    else:
        dup = conn.execute(
            """
            SELECT r.num_item, r.hora_desp, r.prefijo_proveedor, h.fecha, h.linea, h.turno
            FROM cuaderno_registros r
            JOIN cuaderno_hojas h ON h.id = r.hoja_id
            WHERE r.codigo_bull = ?
            LIMIT 1
            """,
            (codigo_bull,),
        ).fetchone()
    if not dup:
        return ""
    pref_dup = dup["prefijo_proveedor"] or "LT"
    if dup["fecha"] == fecha and dup["linea"] == linea and dup["turno"] == turno:
        return f"La paleta {pref_dup}-{codigo_bull} ya fue escaneada en la fila N° {dup['num_item']} ({dup['hora_desp']})."
    return (
        f"La paleta {pref_dup}-{codigo_bull} ya fue escaneada en {dup['linea']} · {dup['turno']} "
        f"({dup['fecha']}, fila N° {dup['num_item']} · {dup['hora_desp']})."
    )


def make_hoja_id(fecha: str, linea: str, turno: str) -> str:
    return f"{fecha}_{linea}_{turno}"


def parse_bull_barcode(raw_text: str, default_pref: str = "LT") -> Dict[str, str]:
    """
    Extrae datos de etiquetas AMCOR y de otros proveedores al escanear:
    - Ej. 'LTG3228973' / 'LTG2976017' -> prefijo='LT', codigo_bull='G3228973' (8 casillas)
    - Ej. 'VL91916156' / 'VL62477447' -> prefijo='VL', codigo_bull='91916156'
    - Ej. '50033037 F,WT,1.5LT,1260,30G,MINALBA AMCOR 10/24/26 22:33 LTG2976017' -> detecta L2, 1500 ml, AMCOR, SAP, Fecha, Hora y Bull.
    """
    s = _clean_str(raw_text, 400)
    fallback_pref = _normalize_prefix(default_pref, "LT")
    out = {
        "prefijo_proveedor": fallback_pref,
        "codigo_bull": "",
        "fecha_prod": "",
        "hora_prod": "",
        "material_sap": "",
        "linea_detectada": "",
        "presentacion_detectada": "",
        "proveedor_detectado": "",
        "proveedor_otro_detectado": "",
    }
    if not s:
        return out

    m_sap = _SAP_MATERIAL_RE.search(s)
    rest = s
    if m_sap:
        out["material_sap"] = m_sap.group(1)
        rest = rest.replace(m_sap.group(1), " ", 1)

    for rx_pres, ln_code, pres_lbl in _PRES_DETECT_RULES:
        if rx_pres.search(s):
            out["linea_detectada"] = ln_code
            out["presentacion_detectada"] = pres_lbl
            break

    if _SUPPLIER_AMCOR_RE.search(s):
        out["proveedor_detectado"] = "AMCOR"
    else:
        m_sup_otro = _SUPPLIER_OTRO_RE.search(s)
        if m_sup_otro:
            out["proveedor_detectado"] = "OTRO"
            out["proveedor_otro_detectado"] = m_sup_otro.group(1).upper()

    if out["material_sap"] in SAP_MATERIAL_CATALOG:
        cat = SAP_MATERIAL_CATALOG[out["material_sap"]]
        if not out["linea_detectada"]:
            out["linea_detectada"] = cat["linea"]
        if not out["presentacion_detectada"]:
            out["presentacion_detectada"] = cat["presentacion"]
        if not out["proveedor_detectado"]:
            out["proveedor_detectado"] = cat["proveedor_envase"]

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


def _build_sap_catalogo(conn: Any) -> Dict[str, Dict[str, str]]:
    cat: Dict[str, Dict[str, str]] = {k: dict(v) for k, v in SAP_MATERIAL_CATALOG.items()}
    rows = conn.execute(
        """
        SELECT material_sap, linea, presentacion, proveedor_envase, proveedor_otro
        FROM cuaderno_hojas
        WHERE material_sap != ''
        ORDER BY updated_at ASC
        """
    ).fetchall()
    for r in rows:
        sap = _clean_str(r["material_sap"], 30)
        if not sap:
            continue
        cat[sap] = {
            "linea": r["linea"] or "",
            "presentacion": r["presentacion"] or "",
            "proveedor_envase": r["proveedor_envase"] or "AMCOR",
            "proveedor_otro": r["proveedor_otro"] or "",
        }
    return cat


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
        "sap_catalogo": _build_sap_catalogo(conn),
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


_OFFICIAL_TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates", "fo003_base_oficial.xlsx")
_XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_A14_NS = "http://schemas.microsoft.com/office/drawing/2010/main"
_A16_NS = "http://schemas.microsoft.com/office/drawing/2014/main"

ET.register_namespace("xdr", _XDR_NS)
ET.register_namespace("a", _A_NS)
ET.register_namespace("r", _R_NS)
ET.register_namespace("a14", _A14_NS)
ET.register_namespace("a16", _A16_NS)


def _load_official_drawing_assets() -> Tuple[bytes, bytes, Dict[str, bytes]]:
    with zipfile.ZipFile(_OFFICIAL_TEMPLATE_PATH, "r") as zf:
        d_xml = zf.read("xl/drawings/drawing1.xml")
        d_rels = zf.read("xl/drawings/_rels/drawing1.xml.rels")
        media = {
            name: zf.read(name)
            for name in zf.namelist()
            if name.startswith("xl/media/")
        }
    return d_xml, d_rels, media


def _set_shape_check_mark(anchor_el: ET.Element, checked: bool) -> None:
    sp = anchor_el.find(f"{{{_XDR_NS}}}sp")
    if sp is None:
        return
    tx_body = sp.find(f"{{{_XDR_NS}}}txBody")
    if tx_body is None:
        return
    body_pr = tx_body.find(f"{{{_A_NS}}}bodyPr")
    if body_pr is not None:
        body_pr.set("anchor", "ctr")
        body_pr.set("lIns", "0")
        body_pr.set("rIns", "0")
        body_pr.set("tIns", "0")
        body_pr.set("bIns", "0")
    p = tx_body.find(f"{{{_A_NS}}}p")
    if p is None:
        p = ET.SubElement(tx_body, f"{{{_A_NS}}}p")
    for child in list(p):
        p.remove(child)
    p_pr = ET.SubElement(p, f"{{{_A_NS}}}pPr")
    p_pr.set("algn", "ctr")
    if checked:
        r = ET.SubElement(p, f"{{{_A_NS}}}r")
        r_pr = ET.SubElement(r, f"{{{_A_NS}}}rPr")
        r_pr.set("lang", "es-VE")
        r_pr.set("sz", "900")
        r_pr.set("b", "1")
        fill = ET.SubElement(r_pr, f"{{{_A_NS}}}solidFill")
        clr = ET.SubElement(fill, f"{{{_A_NS}}}srgbClr")
        clr.set("val", "0F172A")
        t = ET.SubElement(r, f"{{{_A_NS}}}t")
        t.text = "✓"
    end_pr = ET.SubElement(p, f"{{{_A_NS}}}endParaRPr")
    end_pr.set("lang", "es-VE")
    end_pr.set("sz", "900")


def _customize_drawing_xml(base_drawing_xml: bytes, meta: Dict[str, Any]) -> bytes:
    root = ET.fromstring(base_drawing_xml)
    prov_env = str(meta.get("proveedor_envase") or "AMCOR").upper()
    turno = str(meta.get("turno") or "T1").upper()
    pres_idx = int(meta.get("pres_idx") if meta.get("pres_idx") is not None else -1)
    pres_count = int(meta.get("pres_count") or 1)
    extra_rows = int(meta.get("extra_rows") or 0)

    to_remove: List[ET.Element] = []
    for anc in list(root):
        nv = anc.find(f".//{{{_XDR_NS}}}cNvPr")
        name = nv.attrib.get("name", "") if nv is not None else ""
        if name in ("Rectángulo redondeado 55", "Rectángulo redondeado 56"):
            to_remove.append(anc)
            continue
        if name == "Rectángulo redondeado 15" and pres_count < 2:
            to_remove.append(anc)
            continue
        if name == "Rectángulo redondeado 24":
            _set_shape_check_mark(anc, prov_env == "AMCOR")
        elif name == "Rectángulo redondeado 25":
            _set_shape_check_mark(anc, turno == "T1")
        elif name == "Rectángulo redondeado 26":
            _set_shape_check_mark(anc, turno == "T2")
        elif name == "Rectángulo redondeado 27":
            _set_shape_check_mark(anc, turno == "T3")
        elif name == "Rectángulo redondeado 14":
            _set_shape_check_mark(anc, pres_idx == 0)
        elif name == "Rectángulo redondeado 15":
            _set_shape_check_mark(anc, pres_idx == 1)
        elif extra_rows > 0 and name in ("Imagen 3", "Imagen 23", "Imagen 43"):
            frm = anc.find(f"{{{_XDR_NS}}}from")
            if frm is not None:
                row_el = frm.find(f"{{{_XDR_NS}}}row")
                if row_el is not None and row_el.text and row_el.text.isdigit():
                    row_el.text = str(int(row_el.text) + extra_rows)

    for anc in to_remove:
        root.remove(anc)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _inject_official_drawings(raw_xlsx_bytes: bytes, sheet_metas: List[Dict[str, Any]]) -> bytes:
    base_drawing_xml, base_drawing_rels, base_media = _load_official_drawing_assets()
    in_zf = zipfile.ZipFile(io.BytesIO(raw_xlsx_bytes), "r")
    out_buf = io.BytesIO()
    num_sheets = len(sheet_metas)

    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as out_zf:
        for item in in_zf.infolist():
            fn = item.filename
            if fn.startswith("xl/media/") or fn.startswith("xl/drawings/") or fn.startswith("xl/worksheets/_rels/"):
                continue
            data = in_zf.read(fn)
            if fn.startswith("xl/worksheets/sheet") and fn.endswith(".xml"):
                s_txt = data.decode("utf-8")
                s_txt = re.sub(r"<drawing\b[^>]*/>", "", s_txt)
                s_txt = s_txt.replace(
                    "</worksheet>",
                    '<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdDrawing1"/></worksheet>',
                )
                data = s_txt.encode("utf-8")
            elif fn == "[Content_Types].xml":
                ct = data.decode("utf-8")
                ct = re.sub(r'<Override\s+PartName="/xl/drawings/drawing\d+\.xml"[^>]*/>', "", ct)
                if 'Extension="jpeg"' not in ct:
                    ct = ct.replace("</Types>", '<Default Extension="jpeg" ContentType="image/jpeg"/></Types>')
                if 'Extension="png"' not in ct:
                    ct = ct.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>')
                overrides = "".join(
                    f'<Override PartName="/xl/drawings/drawing{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
                    for i in range(1, num_sheets + 1)
                )
                ct = ct.replace("</Types>", f"{overrides}</Types>")
                data = ct.encode("utf-8")
            out_zf.writestr(fn, data)

        for media_name, media_bytes in base_media.items():
            out_zf.writestr(media_name, media_bytes)

        for idx, meta in enumerate(sheet_metas, start=1):
            custom_drawing = _customize_drawing_xml(base_drawing_xml, meta)
            out_zf.writestr(f"xl/drawings/drawing{idx}.xml", custom_drawing)
            out_zf.writestr(f"xl/drawings/_rels/drawing{idx}.xml.rels", base_drawing_rels)

            orig_rels_path = f"xl/worksheets/_rels/sheet{idx}.xml.rels"
            if orig_rels_path in in_zf.namelist():
                rels_txt = in_zf.read(orig_rels_path).decode("utf-8")
                rels_txt = re.sub(r'<Relationship[^>]*relationships/drawing[^>]*/>', "", rels_txt)
                rels_txt = rels_txt.replace(
                    "</Relationships>",
                    f'<Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{idx}.xml"/></Relationships>',
                )
            else:
                rels_txt = (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    f'<Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{idx}.xml"/>'
                    "</Relationships>"
                )
            out_zf.writestr(orig_rels_path, rels_txt.encode("utf-8"))

    in_zf.close()
    return out_buf.getvalue()


def _copy_row_style(ws: Any, src_row: int, dst_row: int, max_col: int = 48) -> None:
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height or 13.5
    for c in range(1, max_col + 1):
        sc = ws.cell(src_row, c)
        dc = ws.cell(dst_row, c)
        if sc.has_style:
            dc.font = copy(sc.font)
            dc.border = copy(sc.border)
            dc.fill = copy(sc.fill)
            dc.number_format = sc.number_format
            dc.alignment = copy(sc.alignment)


def _ensure_merged_range(ws: Any, cell_range: str) -> None:
    if cell_range not in {str(r) for r in ws.merged_cells.ranges}:
        ws.merge_cells(cell_range)


def _populate_fo003_worksheet(ws: Any, hoja: Dict[str, Any], registros: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Rellena la plantilla oficial exacta FO003-RP06-PRO con logos, imágenes, flechas y casillas para cada línea."""
    cfg = LINEA_DEFAULTS.get(hoja["linea"], LINEA_DEFAULTS["L1"])
    f_val = Font(name="Calibri", size=11, bold=False, color="0F172A")
    f_lbl = Font(name="Calibri", size=11, bold=True, color="0F172A")
    f_mono = Font(name="Calibri", size=11, bold=True, color="0F172A")
    b_bottom = Border(bottom=Side(border_style="thin", color="0F172A"))

    # Encabezado: Fecha, Operador, Relevo, Proveedor, Turno
    ws["A7"] = f"FECHA:   {_format_dmy_short(hoja['fecha'])}"

    _ensure_merged_range(ws, "G9:K9")
    ws["G9"] = hoja.get("operador") or ""
    ws["G9"].font = f_val
    ws["G9"].alignment = Alignment(horizontal="left", vertical="center")
    for c_idx in range(7, 12):
        ws.cell(9, c_idx).border = b_bottom

    _ensure_merged_range(ws, "Q9:U9")
    ws["Q9"] = hoja.get("relevo") or ""
    ws["Q9"].font = f_val
    ws["Q9"].alignment = Alignment(horizontal="left", vertical="center")
    for c_idx in range(17, 22):
        ws.cell(9, c_idx).border = b_bottom

    prov_env = (hoja.get("proveedor_envase") or "AMCOR").upper()
    prov_otro_val = (hoja.get("proveedor_otro") or "").strip()
    ws["G11"] = "AMCOR"
    ws["M11"] = "☑ OTRO" if prov_env == "OTRO" else "OTRO"
    _ensure_merged_range(ws, "Q11:U11")
    ws["Q11"] = prov_otro_val if prov_env == "OTRO" else ""
    ws["Q11"].font = f_val
    ws["Q11"].alignment = Alignment(horizontal="left", vertical="center")
    for c_idx in range(17, 22):
        ws.cell(11, c_idx).border = b_bottom

    # Bloque derecho: Línea, Cod.B, Presentaciones y Sabor/SAP
    linea_num = str(hoja["linea"]).replace("L", "")
    ws["AQ7"] = linea_num
    ws["AQ7"].font = f_lbl
    ws["AR7"] = hoja.get("codigo_linea") or cfg["codigo"]
    ws["AR7"].font = f_lbl

    pres_actual = (hoja.get("presentacion") or "").strip()
    pres_items = list(cfg["presentaciones"])
    if pres_actual and not any(p["label"] == pres_actual for p in pres_items):
        pres_items.append({"label": pres_actual, "cod": ""})

    pres_idx = next((idx for idx, p in enumerate(pres_items) if p["label"] == pres_actual), 0)

    ws["AN8"] = pres_items[0]["label"] if len(pres_items) >= 1 else ""
    ws["AR8"] = int(pres_items[0]["cod"]) if (len(pres_items) >= 1 and str(pres_items[0]["cod"]).isdigit()) else (pres_items[0]["cod"] if len(pres_items) >= 1 else "")

    ws["AN9"] = pres_items[1]["label"] if len(pres_items) >= 2 else ""
    ws["AR9"] = int(pres_items[1]["cod"]) if (len(pres_items) >= 2 and str(pres_items[1]["cod"]).isdigit()) else (pres_items[1]["cod"] if len(pres_items) >= 2 else "")

    if cfg.get("tiene_sabor") or hoja.get("sabor"):
        _ensure_merged_range(ws, "AN10:AP10")
        ws["AN10"] = "Sabor:"
        ws["AN10"].font = Font(name="Calibri", size=10, bold=True, color="0F172A")
        _ensure_merged_range(ws, "AR10:AT10")
        ws["AR10"] = hoja.get("sabor") or ""
        ws["AR10"].font = Font(name="Calibri", size=10, bold=False, color="0F172A")
    else:
        ws["AN10"] = ""
        ws["AR10"] = ""

    if hoja.get("material_sap"):
        _ensure_merged_range(ws, "AN11:AP11")
        ws["AN11"] = "SAP:"
        ws["AN11"].font = Font(name="Calibri", size=10, bold=True, color="0F172A")
        _ensure_merged_range(ws, "AR11:AT11")
        ws["AR11"] = hoja.get("material_sap") or ""
        ws["AR11"].font = Font(name="Calibri", size=10, bold=False, color="0F172A")
    else:
        ws["AN11"] = ""
        ws["AR11"] = ""

    # Tabla: filas 18..55 (1..38) o más si hay > 38 registros
    extra_rows = max(0, len(registros) - 38)
    if extra_rows > 0:
        for merged_rng in [str(m) for m in ws.merged_cells.ranges if m.min_row >= 56]:
            ws.unmerge_cells(merged_rng)
        for r_idx in range(56, 61):
            for c_idx in range(1, 49):
                ws.cell(r_idx, c_idx).value = None
                ws.cell(r_idx, c_idx).border = Border()
        for off in range(extra_rows):
            new_r = 56 + off
            _copy_row_style(ws, 55, new_r)
            for rng in (
                f"A{new_r}:B{new_r}",
                f"C{new_r}:G{new_r}",
                f"H{new_r}:L{new_r}",
                f"M{new_r}:N{new_r}",
                f"W{new_r}:AB{new_r}",
                f"AC{new_r}:AF{new_r}",
                f"AG{new_r}:AR{new_r}",
                f"AT{new_r}:AV{new_r}",
            ):
                _ensure_merged_range(ws, rng)

    total_rows = 38 + extra_rows
    for idx in range(total_rows):
        excel_row = 18 + idx
        reg = registros[idx] if idx < len(registros) else None
        ws.cell(excel_row, 1, reg["num_item"] if reg else (idx + 1))
        ws.cell(excel_row, 3, reg["lote_sap"] if reg else "")
        ws.cell(excel_row, 8, reg["hora_desp"] if reg else "")
        ws.cell(excel_row, 13, reg["prefijo_proveedor"] if reg else "")

        bull_str = (reg["codigo_bull"] if reg else "") or ""
        bull_padded = bull_str[:8].ljust(8)
        for d_i in range(8):
            c_digit = ws.cell(excel_row, 15 + d_i, bull_padded[d_i].strip())
            c_digit.font = f_mono
            c_digit.alignment = Alignment(horizontal="center", vertical="center")

        ws.cell(excel_row, 23, reg["fecha_prod"] if reg else "")
        ws.cell(excel_row, 29, reg["hora_prod"] if reg else "")
        ws.cell(excel_row, 33, reg["observacion"] if reg else "")

    # Pie oficial: Analista de Calidad y Facilitador (T1 / T2 / T3)
    foot_r1 = 57 + extra_rows
    foot_r2 = 59 + extra_rows
    if extra_rows > 0:
        for fr, label in ((foot_r1, "Analista de Calidad"), (foot_r2, "Facilitador")):
            ws.cell(fr, 1, label).font = f_val
            ws.cell(fr, 10, "T 1").font = f_lbl
            _ensure_merged_range(ws, f"J{fr}:K{fr}")
            ws.cell(fr, 17, "T 2").font = f_lbl
            _ensure_merged_range(ws, f"Q{fr}:R{fr}")
            ws.cell(fr, 24, "T 3").font = f_lbl
            _ensure_merged_range(ws, f"X{fr}:Y{fr}")
            for c_idx in list(range(12, 17)) + list(range(19, 24)) + list(range(26, 32)):
                ws.cell(fr, c_idx).border = b_bottom

    analista_val = hoja.get("analista_calidad") or ""
    facilitador_val = hoja.get("facilitador") or ""
    turno_col_map = {"T1": (12, 16), "T2": (19, 23), "T3": (26, 31)}
    for t_key, (c_start, c_end) in turno_col_map.items():
        c1_ltr = ws.cell(foot_r1, c_start).column_letter
        c2_ltr = ws.cell(foot_r1, c_end).column_letter
        _ensure_merged_range(ws, f"{c1_ltr}{foot_r1}:{c2_ltr}{foot_r1}")
        _ensure_merged_range(ws, f"{c1_ltr}{foot_r2}:{c2_ltr}{foot_r2}")
        ws.cell(foot_r1, c_start, analista_val if hoja["turno"] == t_key else "").font = f_val
        ws.cell(foot_r2, c_start, facilitador_val if hoja["turno"] == t_key else "").font = f_val

    return {
        "proveedor_envase": prov_env,
        "turno": hoja["turno"],
        "pres_idx": pres_idx,
        "pres_count": len(pres_items),
        "extra_rows": extra_rows,
    }


def generate_cuaderno_excel_bytes(conn: Any, fecha: str, linea: str, turno: str, scope: str = "hoja") -> Tuple[bytes, str]:
    wb = load_workbook(_OFFICIAL_TEMPLATE_PATH)
    base_ws = wb.active
    sheet_metas: List[Dict[str, Any]] = []

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
            ws = base_ws if first else wb.copy_worksheet(base_ws)
            first = False
            ws.title = f"{h_dict['linea']}_{h_dict['turno']}"
            sheet_metas.append(_populate_fo003_worksheet(ws, h_dict, regs))
        fname = f"FO003_Consumo_Envases_{fecha}_Completo.xlsx"
    else:
        payload = build_cuaderno_payload(conn, fecha, linea, turno)
        ws = base_ws
        ws.title = f"{linea}_{turno}"
        sheet_metas.append(_populate_fo003_worksheet(ws, payload["hoja"], payload["registros"]))
        fname = f"FO003_Consumo_Envases_{fecha}_{linea}_{turno}.xlsx"

    out = io.BytesIO()
    wb.save(out)
    final_bytes = _inject_official_drawings(out.getvalue(), sheet_metas)
    return final_bytes, fname


def parse_imported_cuaderno_rows(filename: str, raw_bytes: bytes, target_sheet: str = "") -> List[Dict[str, str]]:
    fname = (filename or "").lower()
    matrix: List[List[str]] = []
    if fname.endswith((".xlsx", ".xlsm")):
        wb = load_workbook(io.BytesIO(raw_bytes), data_only=True, read_only=True)
        try:
            ws = wb[target_sheet] if target_sheet and target_sheet in wb.sheetnames else wb.active
            for row in ws.iter_rows(max_row=1000, max_col=50, values_only=True):
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
    is_fo003_48col = (
        header_idx >= 0
        and col_map.get("lote_sap") == 2
        and col_map.get("hora_desp") == 7
        and col_map.get("proveedor") == 12
        and col_map.get("fecha_prod") == 22
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
        col12_up = (row[12] if len(row) > 12 else "").strip().upper()
        if (
            first_col in ("0", "N°", "Nº", "LT / VL", "D1", "C1")
            or col3_up in ("LT / VL", "LT/VL", "PROVEEDOR")
            or col4_up in ("D1", "C1")
            or col12_up in ("LT / VL", "LT/VL", "PROVEEDOR")
        ):
            continue
        if is_fo003_48col and len(row) >= 22:
            lote = row[2]
            h_desp = row[7]
            pref = _normalize_prefix(row[12], "LT")
            bull = "".join(ch for ch in "".join(row[14:22]).upper() if ch.isalnum())[:8]
            f_prod = row[22] if len(row) > 22 else ""
            h_prod = row[28] if len(row) > 28 else ""
            obs = row[32] if len(row) > 32 else ""
        elif is_fo003_15col and len(row) >= 12:
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
        auto_detect_linea = bool(data.get("auto_detect_linea"))

        ok_img, msg_img = validar_imagen_data_url(foto_etiqueta)
        if not ok_img:
            return jsonify({"error": msg_img or "Foto de etiqueta inválida."}), 400

        reg_id = uuid.uuid4().hex[:12]
        with db_lock:
            with get_db() as conn:
                parsed_pre = parse_bull_barcode(raw_barcode) if raw_barcode else {}
                sap_cat = _build_sap_catalogo(conn)
                sap_code = _clean_str(parsed_pre.get("material_sap") or "", 30)
                sap_entry = sap_cat.get(sap_code, {}) if sap_code else {}
                if auto_detect_linea:
                    det_ln = parsed_pre.get("linea_detectada") or sap_entry.get("linea") or ""
                    if det_ln in LINEAS_CUADERNO:
                        linea = det_ln

                bull_explicit_pre = _clean_str(data.get("codigo_bull", ""), 20)
                pre_bull = _clean_bull_code(bull_explicit_pre or parsed_pre.get("codigo_bull") or "")
                if not pre_bull and not raw_barcode:
                    return jsonify({"error": "Escanea o ingresa el No. de Lote / Bull de la etiqueta (ej. LTG3228973, VL62477447 u OTRO)."}), 400
                if pre_bull:
                    early_dup = _find_duplicate_bull_error(conn, pre_bull, fecha, linea, turno)
                    if early_dup:
                        return jsonify({"error": early_dup, "duplicado": True}), 409

                hoja, _ = _ensure_or_load_hoja(conn, fecha, linea, turno, actor)
                hoja_id = hoja["id"]
                default_pref = "VL" if (hoja.get("proveedor_envase") or "").upper() == "OTRO" else "LT"
                parsed = parse_bull_barcode(raw_barcode, default_pref=default_pref) if raw_barcode else {
                    "prefijo_proveedor": default_pref,
                    "codigo_bull": "",
                    "fecha_prod": "",
                    "hora_prod": "",
                    "material_sap": "",
                    "linea_detectada": "",
                    "presentacion_detectada": "",
                    "proveedor_detectado": "",
                    "proveedor_otro_detectado": "",
                }

                if auto_detect_linea:
                    det_pres = parsed.get("presentacion_detectada") or sap_entry.get("presentacion") or ""
                    det_prov = parsed.get("proveedor_detectado") or sap_entry.get("proveedor_envase") or ""
                    det_prov_otro = parsed.get("proveedor_otro_detectado") or sap_entry.get("proveedor_otro") or ""
                    upd_pres = det_pres or hoja.get("presentacion") or ""
                    upd_prov = det_prov if det_prov in ("AMCOR", "OTRO") else (hoja.get("proveedor_envase") or "AMCOR")
                    upd_otro = det_prov_otro if upd_prov == "OTRO" else (hoja.get("proveedor_otro") if upd_prov == "OTRO" else "")
                    upd_sap = sap_code or hoja.get("material_sap") or ""
                    if (
                        upd_pres != (hoja.get("presentacion") or "")
                        or upd_prov != (hoja.get("proveedor_envase") or "")
                        or upd_otro != (hoja.get("proveedor_otro") or "")
                        or upd_sap != (hoja.get("material_sap") or "")
                    ):
                        conn.execute(
                            """
                            UPDATE cuaderno_hojas
                            SET presentacion = ?, proveedor_envase = ?, proveedor_otro = ?, material_sap = ?, updated_at = ?, updated_by = ?
                            WHERE id = ?
                            """,
                            (upd_pres, upd_prov, upd_otro or "", upd_sap, _now_iso(), actor, hoja_id),
                        )
                        hoja, _ = _ensure_or_load_hoja(conn, fecha, linea, turno, actor)

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

                dup_msg = _find_duplicate_bull_error(conn, codigo_bull, fecha, linea, turno)
                if dup_msg:
                    return jsonify({"error": dup_msg, "duplicado": True}), 409

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
                new_sap = parsed["material_sap"] if (parsed["material_sap"] and not hoja.get("material_sap")) else (hoja.get("material_sap") or "")
                new_fac = hoja.get("facilitador") or (actor if actor and actor != "Sistema" else "")
                conn.execute(
                    "UPDATE cuaderno_hojas SET material_sap = ?, facilitador = ?, updated_at = ?, updated_by = ? WHERE id = ?",
                    (new_sap, new_fac, now_s, actor, hoja_id),
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
                dup_msg = _find_duplicate_bull_error(conn, codigo_bull, hoja["fecha"], hoja["linea"], hoja["turno"], exclude_reg_id=reg_id)
                if dup_msg:
                    return jsonify({"error": dup_msg, "duplicado": True}), 409

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
