"""
Pruebas de calendario festivo contractual (SINTRAMAPEPC), caché versionado y reemplazo en líneas.
"""

import os
import pathlib
import shutil
import sys
import tempfile
import threading
import unittest
from concurrent.futures import ThreadPoolExecutor
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

    def test_cambio_turno_and_cambio_grupo_across_groups(self):
        import re
        r_get = self.client.get("/login")
        csrf_login = re.search(r'name="csrf_token"\s+value="([^"]+)"', r_get.get_data(as_text=True)).group(1)
        self.client.post("/login", data={"username": "sup_hol", "password": "SuperPass#2026", "csrf_token": csrf_login})
        csrf = self.client.get("/api/auth/me").get_json()["csrf_token"]

        # Operador del Grupo B entra a cubrir en Grupo A como cambio_turno
        r_create = self.client.post(
            "/api/operadores",
            json={"nombre": "Operador Cambio B", "columna": "POOL", "grupo": "Grupo B"},
            headers={"X-CSRF-Token": csrf},
        )
        self.assertEqual(r_create.status_code, 201)
        op_id = next(o["id"] for o in r_create.get_json()["pool"] if o["nombre"] == "Operador Cambio B")

        r_move = self.client.post(
            "/api/asignaciones/mover",
            json={
                "operador_id": op_id,
                "destino": "RELEVO",
                "estado": "cambio_turno",
                "cambio_grupo": "Grupo A",
            },
            headers={"X-CSRF-Token": csrf},
        )
        self.assertEqual(r_move.status_code, 200)
        rel_op = next(o for o in r_move.get_json()["tablero"]["RELEVO"] if o["id"] == op_id)
        self.assertEqual(rel_op["estado"], "cambio_turno")
        self.assertEqual(rel_op["grupo"], "Grupo B")
        self.assertEqual(rel_op["cambio_grupo"], "Grupo A")

        # Cambiar cambio_grupo a Grupo C vía PATCH
        r_patch = self.client.patch(
            f"/api/operadores/{op_id}",
            json={"estado": "cambio_turno", "cambio_grupo": "Grupo C"},
            headers={"X-CSRF-Token": csrf},
        )
        self.assertEqual(r_patch.status_code, 200)
        rel_op2 = next(o for o in r_patch.get_json()["tablero"]["RELEVO"] if o["id"] == op_id)
        self.assertEqual(rel_op2["cambio_grupo"], "Grupo C")

        # Volver a estado normal limpia cambio_grupo
        r_norm = self.client.patch(
            f"/api/operadores/{op_id}",
            json={"estado": "normal"},
            headers={"X-CSRF-Token": csrf},
        )
        self.assertEqual(r_norm.status_code, 200)
        rel_op3 = next(o for o in r_norm.get_json()["tablero"]["RELEVO"] if o["id"] == op_id)
        self.assertEqual(rel_op3["estado"], "normal")
        self.assertEqual(rel_op3["cambio_grupo"], "")

    def test_relevo_unlimited_slots(self):
        import re
        r_get = self.client.get("/login")
        csrf_login = re.search(r'name="csrf_token"\s+value="([^"]+)"', r_get.get_data(as_text=True)).group(1)
        self.client.post("/login", data={"username": "sup_hol", "password": "SuperPass#2026", "csrf_token": csrf_login})
        csrf = self.client.get("/api/auth/me").get_json()["csrf_token"]

        self.assertGreaterEqual(len(app_module.SLOTS_DEFINICION["RELEVO"]), 8)

        for i in range(10):
            r = self.client.post(
                "/api/operadores",
                json={"nombre": f"Relevo Extra {i + 1}", "columna": "RELEVO"},
                headers={"X-CSRF-Token": csrf},
            )
            self.assertEqual(r.status_code, 201)

        r_high = self.client.post(
            "/api/operadores",
            json={"nombre": "Relevo Sin Limite 250", "columna": "RELEVO", "posicion": 250},
            headers={"X-CSRF-Token": csrf},
        )
        self.assertEqual(r_high.status_code, 201)
        relevos = r_high.get_json()["tablero"]["RELEVO"]
        self.assertTrue(any(o["nombre"] == "Relevo Sin Limite 250" and o["posicion"] == 250 for o in relevos))

    def test_supervision_seed_move_lineas_cargo_and_import_preserve(self):
        import re
        r_get = self.client.get("/login")
        csrf_login = re.search(r'name="csrf_token"\s+value="([^"]+)"', r_get.get_data(as_text=True)).group(1)
        self.client.post("/login", data={"username": "sup_hol", "password": "SuperPass#2026", "csrf_token": csrf_login})
        csrf = self.client.get("/api/auth/me").get_json()["csrf_token"]

        board = self.client.get("/api/tablero").get_json()
        pool_sup = [o for o in board["pool"] if o.get("tipo_personal") == "supervision"]
        self.assertEqual(len(pool_sup), 8)
        raynier = next(o for o in pool_sup if "Raynier" in o["nombre"])
        self.assertEqual(raynier["nombre"], "Raynier Córdoba")
        self.assertEqual(raynier["puesto"], "Supervisor")

        # Mover 4 personas de supervisión/técnico/especialista a SUPERVISION (sin límite ni reemplazos)
        selected_ids = [o["id"] for o in pool_sup[:4]]
        for idx, op_id in enumerate(selected_ids):
            r_mv = self.client.post(
                "/api/asignaciones/mover",
                json={"operador_id": op_id, "destino": "SUPERVISION"},
                headers={"X-CSRF-Token": csrf},
            )
            self.assertEqual(r_mv.status_code, 200)
            payload = r_mv.get_json()
            self.assertEqual(len(payload["supervision"]), idx + 1)
            self.assertEqual(payload["resumen"]["supervision"], idx + 1)

        # Asignar líneas a cargo y foto a un supervisor en turno
        sup_op_id = selected_ids[0]
        r_patch = self.client.patch(
            f"/api/operadores/{sup_op_id}",
            json={
                "grupo": "Grupo A",
                "tipo_personal": "supervision",
                "lineas_cargo": ["L1", "L2"],
                "foto_url": "data:image/webp;base64,UklGRh4AAABXRUJQVlA4TBEAAAAvAAAAAAfQ//73v/+BiOh/AAA=",
            },
            headers={"X-CSRF-Token": csrf},
        )
        self.assertEqual(r_patch.status_code, 200)
        sup_updated = next(o for o in r_patch.get_json()["supervision"] if o["id"] == sup_op_id)
        self.assertEqual(sup_updated["grupo"], "Grupo A")
        self.assertEqual(sup_updated["lineas_cargo"], "L1,L2")
        self.assertTrue(sup_updated["foto_url"].startswith("data:image/webp"))

        # Limpiar líneas NO quita a supervisión ni reemplazo de importación borra supervisión
        r_clear = self.client.post("/api/tablero/limpiar_lineas", json={}, headers={"X-CSRF-Token": csrf})
        self.assertEqual(r_clear.status_code, 200)
        self.assertEqual(len(r_clear.get_json()["supervision"]), 4)

        r_imp = self.client.post(
            "/api/importar/aplicar",
            json={
                "operadores": [{"nombre": "Nuevo Operador Excel", "columna": "L1", "puesto": "Maquinista", "grupo": "Grupo A"}],
                "modo": "reemplazar",
                "destino_carga": "pool",
                "titulo": "Carga Prueba",
            },
            headers={"X-CSRF-Token": csrf},
        )
        self.assertEqual(r_imp.status_code, 200)
        after_imp = self.client.get("/api/tablero").get_json()
        self.assertEqual(len(after_imp["supervision"]), 4)
        all_sup = after_imp["supervision"] + [o for o in after_imp["pool"] if o.get("tipo_personal") == "supervision"]
        self.assertEqual(len(all_sup), 8)
        preserved = next(o for o in after_imp["supervision"] if o["id"] == sup_op_id)
        self.assertEqual(preserved["lineas_cargo"], "L1,L2")
        self.assertTrue(preserved["foto_url"].startswith("data:image/webp"))

    def test_slim_response_photo_cache_and_concurrent_multiuser_ops(self):
        import re
        r_get = self.client.get("/login")
        csrf_login = re.search(r'name="csrf_token"\s+value="([^"]+)"', r_get.get_data(as_text=True)).group(1)
        self.client.post("/login", data={"username": "sup_hol", "password": "SuperPass#2026", "csrf_token": csrf_login})
        csrf = self.client.get("/api/auth/me").get_json()["csrf_token"]

        # Crear un operador con foto WebP válida (>128 bytes para probar slim mode)
        tiny_webp = "data:image/webp;base64," + ("UklGRh4AAABXRUJQVlA4TBEAAAAvAAAAAAfQ//73v/+BiOh/AAA=" + "A" * 160)
        # Usar payload WebP real generado por PIL si está disponible, si no PNG válido >=130 chars
        import base64, io
        from PIL import Image
        buf = io.BytesIO()
        Image.new("RGB", (48, 48), (0, 120, 215)).save(buf, format="WEBP", quality=80)
        real_webp = "data:image/webp;base64," + base64.b64encode(buf.getvalue()).decode("ascii")

        r_c = self.client.post(
            "/api/operadores",
            json={"nombre": "Foto Veloz", "columna": "POOL", "foto_url": real_webp},
            headers={"X-CSRF-Token": csrf, "X-Board-Slim": "1"},
        )
        self.assertEqual(r_c.status_code, 201)
        data_c = r_c.get_json()
        self.assertGreater(data_c.get("state_version", 0), 0)
        op_slim = next(o for o in data_c["pool"] if o["nombre"] == "Foto Veloz")
        op_id = op_slim["id"]
        self.assertTrue(op_slim["foto_url"].startswith(f"/api/operadores/{op_id}/foto?v="))
        foto_v = op_slim["foto_v"]
        self.assertTrue(foto_v)

        # Endpoint de foto con caché ETag / immutable + 304
        r_f1 = self.client.get(op_slim["foto_url"])
        self.assertEqual(r_f1.status_code, 200)
        self.assertEqual(r_f1.mimetype, "image/webp")
        self.assertIn("immutable", r_f1.headers.get("Cache-Control", ""))
        etag = r_f1.headers.get("ETag")
        self.assertTrue(etag)
        r_f304 = self.client.get(op_slim["foto_url"], headers={"If-None-Match": etag})
        self.assertEqual(r_f304.status_code, 304)

        # Caché caliente inmediatamente tras mutación (version coincide)
        cached_payload, _ = app_module.get_cached_state_payload()
        self.assertEqual(cached_payload["state_version"], data_c["state_version"])

        # Crear 6 operadores para prueba concurrente multipersona (lecturas + movimientos paralelos)
        op_ids = [op_id]
        for i in range(6):
            r_op = self.client.post(
                "/api/operadores",
                json={"nombre": f"Concurrente {i}", "columna": "POOL"},
                headers={"X-CSRF-Token": csrf, "X-Board-Slim": "1"},
            )
            self.assertEqual(r_op.status_code, 201)
            op_ids.append(next(o["id"] for o in r_op.get_json()["pool"] if o["nombre"] == f"Concurrente {i}"))

        def worker(idx: int):
            c = app_module.app.test_client()
            rg = c.get("/login")
            clogin = re.search(r'name="csrf_token"\s+value="([^"]+)"', rg.get_data(as_text=True)).group(1)
            c.post("/login", data={"username": "sup_hol", "password": "SuperPass#2026", "csrf_token": clogin})
            tok = c.get("/api/auth/me").get_json()["csrf_token"]
            rtab = c.get("/api/tablero")
            if rtab.status_code != 200:
                return rtab.status_code
            rmv = c.post(
                "/api/asignaciones/mover",
                json={"operador_id": op_ids[idx % len(op_ids)], "destino": "RELEVO"},
                headers={"X-CSRF-Token": tok, "X-Board-Slim": "1"},
            )
            return rmv.status_code

        with ThreadPoolExecutor(max_workers=8) as ex:
            statuses = list(ex.map(worker, range(12)))
        self.assertTrue(all(st == 200 for st in statuses), f"Estados inesperados: {statuses}")

    def test_independent_lines_and_capacity_per_group_a_and_b(self):
        import re
        r_get = self.client.get("/login")
        csrf_login = re.search(r'name="csrf_token"\s+value="([^"]+)"', r_get.get_data(as_text=True)).group(1)
        self.client.post("/login", data={"username": "sup_hol", "password": "SuperPass#2026", "csrf_token": csrf_login})
        csrf = self.client.get("/api/auth/me").get_json()["csrf_token"]

        def create_pool(nombre, grupo, puesto="Pela paletas"):
            r = self.client.post(
                "/api/operadores",
                json={"nombre": nombre, "columna": "POOL", "grupo": grupo, "puesto": puesto},
                headers={"X-CSRF-Token": csrf},
            )
            self.assertEqual(r.status_code, 201)
            return next(o["id"] for o in r.get_json()["pool"] if o["nombre"] == nombre)

        id_a1 = create_pool("Operador GA L1", "Grupo A")
        id_b1 = create_pool("Operador GB L1", "Grupo B")

        # Ambos grupos pueden ocupar L1 casilla 0 (Pela paletas compartido L1+L2) sin desplazarse entre sí
        r_a = self.client.post(
            "/api/asignaciones/mover",
            json={"operador_id": id_a1, "destino": "L1", "posicion": 0, "target_group": "Grupo A"},
            headers={"X-CSRF-Token": csrf},
        )
        self.assertEqual(r_a.status_code, 200)
        r_b = self.client.post(
            "/api/asignaciones/mover",
            json={"operador_id": id_b1, "destino": "L1", "posicion": 0, "target_group": "Grupo B"},
            headers={"X-CSRF-Token": csrf},
        )
        self.assertEqual(r_b.status_code, 200)
        data_b = r_b.get_json()
        l1_ids = {o["id"] for o in data_b["tablero"]["L1"]}
        l2_ids = {o["id"] for o in data_b["tablero"]["L2"]}
        self.assertIn(id_a1, l1_ids)
        self.assertIn(id_b1, l1_ids)
        self.assertIn(id_a1, l2_ids)
        self.assertIn(id_b1, l2_ids)

        # Capacidad de SIAT (máx 2) también es independiente por grupo
        siat_a = [create_pool(f"Siat GA {i}", "Grupo A", "Siat") for i in range(2)]
        siat_b = [create_pool(f"Siat GB {i}", "Grupo B", "Siat") for i in range(2)]
        for sid, grp in [(siat_a[0], "Grupo A"), (siat_a[1], "Grupo A"), (siat_b[0], "Grupo B"), (siat_b[1], "Grupo B")]:
            rm = self.client.post(
                "/api/asignaciones/mover",
                json={"operador_id": sid, "destino": "SIAT", "target_group": grp},
                headers={"X-CSRF-Token": csrf},
            )
            self.assertEqual(rm.status_code, 200)

        # Un 3er SIAT en Grupo A sí debe fallar por cupo 2/2 de Grupo A
        extra_a = create_pool("Siat GA Extra", "Grupo A", "Siat")
        r_full = self.client.post(
            "/api/asignaciones/mover",
            json={"operador_id": extra_a, "destino": "SIAT", "target_group": "Grupo A"},
            headers={"X-CSRF-Token": csrf},
        )
        self.assertEqual(r_full.status_code, 400)

        # Limpiar líneas de Grupo A deja intacto a Grupo B
        r_clean_a = self.client.post(
            "/api/tablero/limpiar_lineas",
            json={"grupo": "Grupo A"},
            headers={"X-CSRF-Token": csrf},
        )
        self.assertEqual(r_clean_a.status_code, 200)
        after_clean = r_clean_a.get_json()
        self.assertFalse(any(o["id"] == id_a1 for o in after_clean["tablero"]["L1"]))
        self.assertTrue(any(o["id"] == id_b1 for o in after_clean["tablero"]["L1"]))
        self.assertEqual(len([o for o in after_clean["tablero"]["SIAT"] if o["grupo"] == "Grupo B"]), 2)

        # Limpiar Grupo B al finalizar para no dejar ocupantes en tests siguientes
        self.client.post("/api/tablero/limpiar_lineas", json={"grupo": "Grupo B"}, headers={"X-CSRF-Token": csrf})

    def test_supervision_stays_in_selected_group_a_or_b_and_rescues_ungrouped(self):
        import re
        r_get = self.client.get("/login")
        csrf_login = re.search(r'name="csrf_token"\s+value="([^"]+)"', r_get.get_data(as_text=True)).group(1)
        self.client.post("/login", data={"username": "sup_hol", "password": "SuperPass#2026", "csrf_token": csrf_login})
        csrf = self.client.get("/api/auth/me").get_json()["csrf_token"]

        board = self.client.get("/api/tablero").get_json()
        all_sup = board["supervision"] + [o for o in board["pool"] if o.get("tipo_personal") == "supervision"]
        sup_target = next(o for o in all_sup if "José" in o["nombre"])

        # Asignar supervisor mientras se está en Grupo A -> debe quedar en Grupo A (no en Todos sin grupo)
        r_a = self.client.post(
            "/api/asignaciones/mover",
            json={"operador_id": sup_target["id"], "destino": "SUPERVISION", "target_group": "Grupo A"},
            headers={"X-CSRF-Token": csrf},
        )
        self.assertEqual(r_a.status_code, 200)
        assigned_a = next(o for o in r_a.get_json()["supervision"] if o["id"] == sup_target["id"])
        self.assertEqual(assigned_a["grupo"], "Grupo A")

        # Si luego se selecciona ese supervisor estando en Grupo B -> pasa a Grupo B (no se queda pegado en Grupo A ni en Todos)
        r_b = self.client.post(
            "/api/asignaciones/mover",
            json={"operador_id": sup_target["id"], "destino": "SUPERVISION", "target_group": "Grupo B"},
            headers={"X-CSRF-Token": csrf},
        )
        self.assertEqual(r_b.status_code, 200)
        assigned_b = next(o for o in r_b.get_json()["supervision"] if o["id"] == sup_target["id"])
        self.assertEqual(assigned_b["grupo"], "Grupo B")

        # Limpiar supervisión sin grupo si queda atrapada cuando ya existen grupos nombrados (rescate en init_db)
        self.client.patch(f"/api/operadores/{sup_target['id']}", json={"grupo": ""}, headers={"X-CSRF-Token": csrf})
        r_chk = self.client.get("/api/tablero").get_json()
        self.assertTrue(any(o["id"] == sup_target["id"] and o["columna"] == "SUPERVISION" for o in r_chk["supervision"]))
        with app_module.get_db() as conn:
            conn.execute("DELETE FROM config_meta WHERE key = 'fix_ungrouped_supervision_v1'")
        app_module._db_initialized_keys.clear()
        app_module.invalidate_board_cache()
        app_module.init_db()
        r_rescued = self.client.get("/api/tablero").get_json()
        self.assertFalse(any(o["id"] == sup_target["id"] for o in r_rescued["supervision"]))
        self.assertTrue(any(o["id"] == sup_target["id"] and o["columna"] == "POOL" for o in r_rescued["pool"]))


if __name__ == "__main__":
    unittest.main(verbosity=2)
