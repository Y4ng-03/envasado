"""
Pruebas de calendario festivo contractual (SINTRAMAPEPC), caché versionado y reemplazo en líneas.
"""

import os
import pathlib
import shutil
import sys
import tempfile
import unittest
from datetime import date

ROOT = pathlib.Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

_TMP_DIR = tempfile.mkdtemp(prefix="envasado_holidays_")
os.environ.setdefault("SQLITE_DB_PATH", os.path.join(_TMP_DIR, "test_holidays.db"))
os.environ.setdefault("FLASK_SECRET_KEY", "test-secret-key-at-least-64-chars-long-0123456789abcdef0123456789")

import app as app_module  # noqa: E402
import db_compat  # noqa: E402
import observabilidad  # noqa: E402
import seguridad  # noqa: E402


class HolidaysAndMovesTestCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._prev_db_path = os.environ.get("SQLITE_DB_PATH")
        os.environ["SQLITE_DB_PATH"] = os.path.join(_TMP_DIR, "test_holidays.db")
        db_compat.close_all_pools()
        app_module._db_initialized_keys.clear()
        db_compat._auth_tables_ready.clear()
        app_module.init_db()
        with db_compat.get_db() as conn:
            db_compat.upsert_user(conn, "sup_hol", "SuperPass#2026", "Supervisor Feriados", rol="supervisor")

    @classmethod
    def tearDownClass(cls) -> None:
        db_compat.close_all_pools()
        app_module._db_initialized_keys.clear()
        db_compat._auth_tables_ready.clear()
        if cls._prev_db_path is None:
            os.environ.pop("SQLITE_DB_PATH", None)
        else:
            os.environ["SQLITE_DB_PATH"] = cls._prev_db_path
        shutil.rmtree(_TMP_DIR, ignore_errors=True)

    def setUp(self) -> None:
        seguridad.reset_security_state_for_tests()
        app_module.invalidate_board_cache()
        self.client = app_module.app.test_client()

    def test_contractual_holidays_2026_and_2027(self):
        h2026 = app_module.get_contractual_holidays(2026)
        # Pascua 2026: 5 de abril -> Carnaval 16-17 feb, Semana Santa 1-4 abr
        expected_2026 = {
            date(2026, 1, 1),
            date(2026, 2, 16),
            date(2026, 2, 17),
            date(2026, 4, 1),
            date(2026, 4, 2),
            date(2026, 4, 3),
            date(2026, 4, 4),
            date(2026, 4, 19),
            date(2026, 5, 1),
            date(2026, 6, 24),
            date(2026, 6, 29),
            date(2026, 7, 5),
            date(2026, 7, 24),
            date(2026, 10, 12),
            date(2026, 10, 21),
            date(2026, 12, 24),
            date(2026, 12, 25),
            date(2026, 12, 31),
        }
        self.assertEqual(set(h2026.keys()), expected_2026)

        # Semana Santa 2026 (lun 30 mar - dom 5 abr): solo lunes 30 y martes 31 son hábiles
        self.assertEqual(app_module.count_business_days_inclusive(date(2026, 3, 30), date(2026, 4, 5)), 2)
        # Reingreso después del 31 de marzo (en plena Semana Santa) salta 1-4 abr y dom 5 -> lun 6 abr
        self.assertEqual(app_module.next_business_day(date(2026, 3, 31)), date(2026, 4, 6))

        # Navidad y fin de año 2026: jue 24, vie 25, jue 31 dic + vie 1 ene 2027
        # Rango 23 dic 2026 (mié) al 3 ene 2027 (dom): hábiles son 23, 28, 29, 30 dic = 4 días
        self.assertEqual(app_module.count_business_days_inclusive(date(2026, 12, 23), date(2027, 1, 3)), 4)
        self.assertEqual(app_module.next_business_day(date(2026, 12, 30)), date(2027, 1, 4))

        # 21 oct 2026 (miércoles, Día de Los Teques) no es hábil
        self.assertEqual(app_module.count_business_days_inclusive(date(2026, 10, 19), date(2026, 10, 23)), 4)

    def test_ttl_cache_set_if_version_rejects_stale_write(self):
        cache = observabilidad.TTLCache(default_ttl_s=15.0)
        v1 = cache.version
        cache.invalidate()
        self.assertFalse(cache.set_if_version("tablero:k", {"stale": True}, expected_version=v1))
        hit, _, _ = cache.get("tablero:k")
        self.assertFalse(hit)
        v2 = cache.version
        self.assertTrue(cache.set_if_version("tablero:k", {"fresh": True}, expected_version=v2))
        hit, val, _ = cache.get("tablero:k")
        self.assertTrue(hit)
        self.assertEqual(val, {"fresh": True})

    def test_mover_replacement_logs_and_returns_reemplazado(self):
        self.client.get("/login")
        import re
        r_get = self.client.get("/login")
        csrf_login = re.search(r'name="csrf_token"\s+value="([^"]+)"', r_get.get_data(as_text=True)).group(1)
        self.client.post("/login", data={"username": "sup_hol", "password": "SuperPass#2026", "csrf_token": csrf_login})
        csrf = self.client.get("/api/auth/me").get_json()["csrf_token"]

        r1 = self.client.post("/api/operadores", json={"nombre": "Operador Uno", "columna": "L1"}, headers={"X-CSRF-Token": csrf})
        self.assertEqual(r1.status_code, 201)
        op1_id = r1.get_json()["tablero"]["L1"][0]["id"]

        r2 = self.client.post("/api/operadores", json={"nombre": "Operador Dos", "columna": "POOL"}, headers={"X-CSRF-Token": csrf})
        self.assertEqual(r2.status_code, 201)
        op2_id = next(o["id"] for o in r2.get_json()["pool"] if o["nombre"] == "Operador Dos")

        r_rep = self.client.post(
            "/api/asignaciones/mover",
            json={"operador_id": op2_id, "destino": "L1", "posicion": 0},
            headers={"X-CSRF-Token": csrf},
        )
        self.assertEqual(r_rep.status_code, 200)
        data = r_rep.get_json()
        self.assertEqual(data.get("reemplazado"), "Operador Uno")
        self.assertTrue(any("Operador Uno" in h["descripcion"] and "Nómina Disponible" in h["descripcion"] for h in data["historial"]))
        self.assertTrue(any(o["id"] == op1_id for o in data["pool"]))


if __name__ == "__main__":
    unittest.main(verbosity=2)
