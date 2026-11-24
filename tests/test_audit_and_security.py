"""
Suite integral de pruebas automatizadas:
- SAST / AST (cero SQL dinámicos)
- Validación de HTML/JS en crudo y accesibilidad
- Seguridad OWASP (autenticación, scrypt, sesiones/rotación/logout, RBAC, anti-IDOR, anti-CSRF, cabeceras, magic bytes, anti-SSRF)
- Rendimiento (SSR, ETag/304 + invalidación, Brotli/Gzip, paginación, cola de trabajos, OpenTelemetry)
"""

import ast
import base64
import brotli
import gzip
import io
import json
import os
import pathlib
import py_compile
import re
import shutil
import subprocess
import sys
import tempfile
import time
import unittest
from html.parser import HTMLParser

from openpyxl import Workbook
from PIL import Image

ROOT = pathlib.Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# Configurar BD temporal aislada antes de importar app
_TEST_TMP_DIR = tempfile.mkdtemp(prefix="envasado_test_")
os.environ["SQLITE_DB_PATH"] = os.path.join(_TEST_TMP_DIR, "test_produccion.db")
os.environ["FLASK_SECRET_KEY"] = "test-secret-key-at-least-64-chars-long-0123456789abcdef0123456789"
os.environ["LOGIN_MAX_FALLOS_IP"] = "5"
os.environ["LOGIN_MAX_FALLOS_USUARIO"] = "5"

import app as app_module  # noqa: E402
import db_compat  # noqa: E402
import seguridad  # noqa: E402


class _HTMLAuditParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.mains = 0
        self.h1s = 0
        self.imgs_without_alt = []
        self.controls = []
        self.labels_for = set()
        self.ids = {}
        self.scripts = []
        self._in_script = False
        self._script_type = ""
        self._buf = []

    def handle_starttag(self, tag, attrs):
        d = dict(attrs)
        if "id" in d:
            self.ids[d["id"]] = self.ids.get(d["id"], 0) + 1
        if tag == "main":
            self.mains += 1
        if tag == "h1":
            self.h1s += 1
        if tag == "img":
            if "alt" not in d or d["alt"] is None or d["alt"].strip() == "":
                self.imgs_without_alt.append(d.get("id") or d.get("src"))
        if tag == "label" and "for" in d:
            self.labels_for.add(d["for"])
        if tag in ("input", "select", "textarea") and d.get("type") != "hidden":
            self.controls.append((tag, d.get("id", ""), d.get("aria-label", "")))
        if tag == "script" and "src" not in d:
            self._in_script = True
            self._script_type = d.get("type", "")
            self._buf = []

    def handle_data(self, data):
        if self._in_script:
            self._buf.append(data)

    def handle_endtag(self, tag):
        if tag == "script" and self._in_script:
            self._in_script = False
            self.scripts.append((self._script_type, "".join(self._buf)))


def _make_png_data_url() -> str:
    img = Image.new("RGB", (12, 12), color=(0, 158, 227))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("ascii")


def _make_xlsx_bytes() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.append(["Nombre", "Linea", "Puesto", "Grupo", "Estado"])
    ws.append(["Ana Pérez", "Linea 1", "Llenadora", "Grupo A", "Normal"])
    ws.append(["Luis Gómez", "Linea 2", "Despaletizador", "Grupo A", "Normal"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class EnvasadoIntegralTestSuite(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        app_module.init_db()
        with db_compat.get_db() as conn:
            db_compat.upsert_user(conn, "admin_test", "AdminPass#2026", "Admin Prueba", rol="admin")
            db_compat.upsert_user(conn, "sup_test", "SuperPass#2026", "Supervisor Prueba", rol="supervisor")
            db_compat.upsert_user(conn, "tv_test", "TvPass#2026", "Pantalla TV", rol="tv")

    @classmethod
    def tearDownClass(cls) -> None:
        db_compat.close_all_pools()
        shutil.rmtree(_TEST_TMP_DIR, ignore_errors=True)

    def setUp(self) -> None:
        seguridad.reset_security_state_for_tests()
        app_module.invalidate_board_cache()
        self.client = app_module.app.test_client()

    def _login(self, username: str, password: str) -> str:
        with self.client.session_transaction() as sess:
            sess.clear()
        r_get = self.client.get("/login")
        self.assertEqual(r_get.status_code, 200)
        m = re.search(r'name="csrf_token"\s+value="([^"]+)"', r_get.get_data(as_text=True))
        self.assertIsNotNone(m, "El formulario /login debe incluir csrf_token")
        login_csrf = m.group(1)
        r_post = self.client.post(
            "/login",
            data={"username": username, "password": password, "csrf_token": login_csrf},
            follow_redirects=False,
        )
        self.assertIn(r_post.status_code, (302, 200))
        me = self.client.get("/api/auth/me").get_json()
        self.assertTrue(me["authenticated"])
        return me["csrf_token"]

    # ── 1. SAST / AST / Compilación / HTML & JS en crudo ────────────────────
    def test_ast_zero_dynamic_sql(self):
        violations = []
        for py_path in ROOT.glob("*.py"):
            tree = ast.parse(py_path.read_text(encoding="utf-8"), filename=str(py_path))
            for node in ast.walk(tree):
                if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute):
                    if node.func.attr in ("execute", "executemany") and node.args:
                        arg0 = node.args[0]
                        if isinstance(arg0, (ast.JoinedStr, ast.BinOp, ast.Call)):
                            violations.append(f"{py_path.name}:{node.lineno}")
        self.assertEqual(violations, [], f"SQL dinámico detectado: {violations}")

    def test_py_compile_all_modules(self):
        for py_path in ROOT.glob("*.py"):
            py_compile.compile(str(py_path), doraise=True)

    def test_raw_html_accessibility_and_js_syntax(self):
        has_node = shutil.which("node") is not None
        for rel in ("index.html", "templates/login.html", "templates/legal.html", "templates/cuaderno.html"):
            raw = (ROOT / rel).read_text(encoding="utf-8")
            clean = re.sub(r"\{%.*?%\}|\{\{.*?\}\}", "", raw) if rel != "index.html" else raw
            parser = _HTMLAuditParser()
            parser.feed(clean)
            self.assertEqual(parser.mains, 1, f"{rel} debe tener exactamente un <main>")
            self.assertEqual(parser.imgs_without_alt, [], f"{rel} tiene imágenes sin alt")
            unlabeled = [c for c in parser.controls if not (c[1] in parser.labels_for or c[2])]
            self.assertEqual(unlabeled, [], f"{rel} tiene controles sin label/aria-label: {unlabeled}")
            dups = [k for k, v in parser.ids.items() if v > 1]
            self.assertEqual(dups, [], f"{rel} tiene IDs duplicados: {dups}")
            for stype, code in parser.scripts:
                if stype == "application/json":
                    json.loads(code)
                elif has_node:
                    with tempfile.NamedTemporaryFile("w", suffix=".js", delete=False, encoding="utf-8") as tf:
                        tf.write(code)
                        tname = tf.name
                    try:
                        proc = subprocess.run(["node", "--check", tname], capture_output=True, text=True)
                        self.assertEqual(proc.returncode, 0, f"Error JS en {rel}: {proc.stderr}")
                    finally:
                        pathlib.Path(tname).unlink(missing_ok=True)

    def test_no_hardcoded_secrets(self):
        secret_patterns = [
            re.compile(r"sk_live_[0-9a-zA-Z]{16,}"),
            re.compile(r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----"),
            re.compile(r"AIza[0-9A-Za-z\-_]{35}"),
        ]
        for path in list(ROOT.glob("*.py")) + [ROOT / "index.html", ROOT / "templates" / "login.html"]:
            content = path.read_text(encoding="utf-8")
            for rx in secret_patterns:
                self.assertIsNone(rx.search(content), f"Posible secreto en {path.name}")

    # ── 2. Seguridad OWASP ──────────────────────────────────────────────────
    def test_mandatory_auth_and_public_legal_routes(self):
        r_idx = self.client.get("/", follow_redirects=False)
        self.assertEqual(r_idx.status_code, 302)
        self.assertIn("/login", r_idx.headers["Location"])

        r_api = self.client.get("/api/tablero")
        self.assertEqual(r_api.status_code, 401)

        for legal_route in ("/terminos", "/privacidad", "/cookies"):
            rl = self.client.get(legal_route)
            self.assertEqual(rl.status_code, 200)
            html = rl.get_data(as_text=True)
            self.assertEqual(html.count("<main"), 1)

    def test_strong_password_hash_and_rate_limit_antispoof(self):
        with db_compat.get_db() as conn:
            row = conn.execute("SELECT password_hash FROM usuarios WHERE username = 'admin_test'").fetchone()
            self.assertTrue(str(row["password_hash"]).startswith("scrypt:"))

        # Intentos fallidos con X-Forwarded-For falso no deben evadir el rate limit por IP
        r_get = self.client.get("/login")
        m = re.search(r'name="csrf_token"\s+value="([^"]+)"', r_get.get_data(as_text=True))
        login_csrf = m.group(1)

        # Token CSRF inválido en /login debe devolver 403 y rotar el token
        r_bad_csrf = self.client.post(
            "/login",
            data={"username": "admin_test", "password": "wrong-password", "csrf_token": "invalid-token"},
        )
        self.assertEqual(r_bad_csrf.status_code, 403)
        m2 = re.search(r'name="csrf_token"\s+value="([^"]+)"', r_bad_csrf.get_data(as_text=True))
        login_csrf = m2.group(1)

        for idx in range(5):
            r = self.client.post(
                "/login",
                data={"username": "admin_test", "password": "wrong-password", "csrf_token": login_csrf},
                headers={"X-Forwarded-For": f"203.0.113.{idx + 1}"},
            )
            self.assertEqual(r.status_code, 401)
        r_blocked = self.client.post(
            "/login",
            data={"username": "admin_test", "password": "wrong-password", "csrf_token": login_csrf},
            headers={"X-Forwarded-For": "198.51.100.99"},
        )
        self.assertEqual(r_blocked.status_code, 429)

    def test_session_rotation_expiry_and_logout_clear_site_data(self):
        self._login("admin_test", "AdminPass#2026")
        with self.client.session_transaction() as sess:
            old_sid = sess["sid"]
            sess["rotated_at"] = int(time.time()) - seguridad.SESION_ROTACION_S - 10

        r_me = self.client.get("/api/auth/me")
        self.assertEqual(r_me.status_code, 200)
        with self.client.session_transaction() as sess:
            self.assertNotEqual(sess["sid"], old_sid)

        # Expiración por inactividad
        with self.client.session_transaction() as sess:
            sess["last_seen"] = int(time.time()) - seguridad.SESION_INACTIVA_S - 10
        self.assertEqual(self.client.get("/api/tablero").status_code, 401)

        # Re-login y logout con Clear-Site-Data
        csrf = self._login("admin_test", "AdminPass#2026")
        r_out = self.client.post("/logout", data={"csrf_token": csrf}, follow_redirects=False)
        self.assertEqual(r_out.status_code, 302)
        csd = r_out.headers.get("Clear-Site-Data", "")
        self.assertIn('"cookies"', csd)
        self.assertIn('"storage"', csd)
        self.assertIn('"cache"', csd)
        self.assertEqual(self.client.get("/api/tablero").status_code, 401)

    def test_rbac_csrf_headers_and_anti_idor(self):
        # TV solo lectura
        tv_csrf = self._login("tv_test", "TvPass#2026")
        self.assertEqual(self.client.get("/api/tablero").status_code, 200)
        r_tv_post = self.client.post(
            "/api/operadores",
            json={"nombre": "Intento TV"},
            headers={"X-CSRF-Token": tv_csrf},
        )
        self.assertEqual(r_tv_post.status_code, 403)
        self.assertEqual(self.client.get("/api/auditoria").status_code, 403)

        # Supervisor: puede crear pero no vaciar ni ver auditoría
        sup_csrf = self._login("sup_test", "SuperPass#2026")
        self.assertEqual(self.client.get("/api/auditoria").status_code, 403)
        self.assertEqual(
            self.client.post("/api/tablero/vaciar", json={}, headers={"X-CSRF-Token": sup_csrf}).status_code,
            403,
        )

        # Anti-CSRF: sin X-CSRF-Token, Origin externo, Sec-Fetch-Site cross-site
        self.assertEqual(self.client.post("/api/operadores", json={"nombre": "Sin CSRF"}).status_code, 403)
        self.assertEqual(
            self.client.post(
                "/api/operadores",
                json={"nombre": "Origen Malo"},
                headers={"Origin": "https://evil.example", "X-CSRF-Token": sup_csrf},
            ).status_code,
            403,
        )
        self.assertEqual(
            self.client.get("/api/tablero", headers={"Sec-Fetch-Site": "cross-site"}).status_code,
            403,
        )

        # Cabeceras de seguridad en HTML y HTTPS
        r_html = self.client.get("/", headers={"X-Forwarded-Proto": "https"})
        self.assertEqual(r_html.status_code, 200)
        self.assertIn("object-src 'none'", r_html.headers.get("Content-Security-Policy", ""))
        self.assertNotIn("'unsafe-inline'", r_html.headers.get("Content-Security-Policy", "").split("style-src")[0])
        self.assertEqual(r_html.headers.get("X-Content-Type-Options"), "nosniff")
        self.assertEqual(r_html.headers.get("X-Frame-Options"), "SAMEORIGIN")
        self.assertEqual(r_html.headers.get("Cross-Origin-Opener-Policy"), "same-origin")
        self.assertEqual(r_html.headers.get("Cross-Origin-Resource-Policy"), "same-origin")
        self.assertIn("camera=()", r_html.headers.get("Permissions-Policy", ""))
        self.assertIn("max-age=31536000", r_html.headers.get("Strict-Transport-Security", ""))

        # Anti-IDOR en trabajos en segundo plano
        admin_csrf = self._login("admin_test", "AdminPass#2026")
        r_job = self.client.post(
            "/api/importar/aplicar?async=1",
            json={"operadores": [{"nombre": "Carlos Ruiz", "columna": "L1", "puesto": "Llenadora"}]},
            headers={"X-CSRF-Token": admin_csrf},
        )
        self.assertEqual(r_job.status_code, 202)
        job_id = r_job.get_json()["job_id"]
        db_compat.bg_jobs.flush()
        self.assertEqual(self.client.get(f"/api/jobs/{job_id}").status_code, 200)

        # Otro usuario (supervisor) no debe poder consultar el trabajo de admin (anti-IDOR)
        self._login("sup_test", "SuperPass#2026")
        self.assertEqual(self.client.get(f"/api/jobs/{job_id}").status_code, 403)

    def test_validation_magic_bytes_xss_and_ssrf(self):
        csrf = self._login("admin_test", "AdminPass#2026")

        # Rechazar bool como entero en posición y fechas pasadas
        self.assertEqual(
            self.client.post(
                "/api/operadores",
                json={"nombre": "Bool Test", "columna": "L1", "posicion": True},
                headers={"X-CSRF-Token": csrf},
            ).status_code,
            400,
        )
        self.assertEqual(
            self.client.post(
                "/api/operadores",
                json={"nombre": "Fecha Mala", "estado": "vacaciones", "vacaciones_inicio": "2020-01-01", "vacaciones_fin": "2020-01-05"},
                headers={"X-CSRF-Token": csrf},
            ).status_code,
            400,
        )

        # Rechazar imagen falsa (magic bytes inválidos) y aceptar PNG real
        fake_b64 = base64.b64encode(b"<script>alert(1)</script>").decode("ascii")
        self.assertEqual(
            self.client.post(
                "/api/operadores",
                json={"nombre": "Fake Img", "foto_url": f"data:image/png;base64,{fake_b64}"},
                headers={"X-CSRF-Token": csrf},
            ).status_code,
            400,
        )

        # XSS neutralizado + PNG válido
        r_ok = self.client.post(
            "/api/operadores",
            json={"nombre": "Pedro <script>X</script>", "foto_url": _make_png_data_url()},
            headers={"X-CSRF-Token": csrf},
        )
        self.assertEqual(r_ok.status_code, 201)
        pool_names = [o["nombre"] for o in r_ok.get_json()["pool"]]
        self.assertTrue(any("<script>" not in n and "Pedro" in n for n in pool_names))

        # Magic bytes en subida de Excel
        r_bad_xlsx = self.client.post(
            "/api/importar/preview",
            data={"file": (io.BytesIO(b"not-a-real-xlsx-zip"), "malo.xlsx")},
            headers={"X-CSRF-Token": csrf},
        )
        self.assertEqual(r_bad_xlsx.status_code, 400)

        r_good_xlsx = self.client.post(
            "/api/importar/preview",
            data={"file": (io.BytesIO(_make_xlsx_bytes()), "nomina.xlsx")},
            headers={"X-CSRF-Token": csrf},
        )
        self.assertEqual(r_good_xlsx.status_code, 200)
        self.assertGreaterEqual(len(r_good_xlsx.get_json()["periodos"]), 1)

        # Anti-SSRF
        for bad_url in (
            "http://127.0.0.1:5000/api/tablero",
            "http://169.254.169.254/latest/meta-data/",
            "http://10.0.0.5/admin",
            "file:///etc/passwd",
        ):
            ok, _ = seguridad.es_url_externa_segura(bad_url)
            self.assertFalse(ok, f"Anti-SSRF debió bloquear {bad_url}")

    # ── 3. Rendimiento, Caché ETag/304, Brotli/Gzip, Paginación, OTel ───────
    def test_perf_ssr_etag_compression_pagination_and_otel(self):
        csrf = self._login("admin_test", "AdminPass#2026")

        # SSR en / + cabeceras OTel
        r_ssr = self.client.get("/")
        self.assertEqual(r_ssr.status_code, 200)
        html = r_ssr.get_data(as_text=True)
        self.assertIn('id="ssr-initial-state"', html)
        self.assertRegex(r_ssr.headers.get("X-Trace-Id", ""), r"^[0-9a-f]{32}$")
        self.assertIn("total;dur=", r_ssr.headers.get("Server-Timing", ""))

        # ETag + 304 Not Modified + invalidación tras mutación
        r1 = self.client.get("/api/tablero")
        self.assertEqual(r1.status_code, 200)
        etag1 = r1.headers.get("ETag")
        self.assertTrue(etag1)

        r304 = self.client.get("/api/tablero", headers={"If-None-Match": etag1})
        self.assertEqual(r304.status_code, 304)

        self.client.post(
            "/api/operadores",
            json={"nombre": "Operador Cache Inv"},
            headers={"X-CSRF-Token": csrf},
        )
        r2 = self.client.get("/api/tablero", headers={"If-None-Match": etag1})
        self.assertEqual(r2.status_code, 200)
        self.assertNotEqual(r2.headers.get("ETag"), etag1)

        # Compresión Brotli y Gzip
        r_br = self.client.get("/api/tablero", headers={"Accept-Encoding": "br, gzip"})
        self.assertEqual(r_br.headers.get("Content-Encoding"), "br")
        self.assertIn("tablero", json.loads(brotli.decompress(r_br.data).decode("utf-8")))

        r_gz = self.client.get("/api/tablero", headers={"Accept-Encoding": "gzip"})
        self.assertEqual(r_gz.headers.get("Content-Encoding"), "gzip")
        self.assertIn("tablero", json.loads(gzip.decompress(r_gz.data).decode("utf-8")))

        # Paginación en APIs
        ops_page = self.client.get("/api/operadores?page=1&per_page=1").get_json()
        self.assertEqual(ops_page["page"], 1)
        self.assertEqual(ops_page["per_page"], 1)
        self.assertIn("total_pages", ops_page)

        hist_page = self.client.get("/api/historial?page=1&per_page=2").get_json()
        self.assertEqual(hist_page["page"], 1)
        self.assertLessEqual(len(hist_page["items"]), 2)

        aud_page = self.client.get("/api/auditoria?page=1&per_page=5").get_json()
        self.assertEqual(aud_page["page"], 1)
        self.assertLessEqual(len(aud_page["items"]), 5)


if __name__ == "__main__":
    unittest.main(verbosity=2)
