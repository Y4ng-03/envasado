"""
observabilidad.py — Trazabilidad OpenTelemetry (X-Trace-Id, Server-Timing),
caché TTL thread-safe con invalidación por versión, y utilidades ETag.
"""

import hashlib
import os
import threading
import time
from typing import Any, Dict, Optional, Tuple

from flask import Response, g, request
from opentelemetry import trace
from opentelemetry.sdk.resources import Resource
from opentelemetry.sdk.trace import TracerProvider
from opentelemetry.trace import SpanContext, TraceFlags

_provider_lock = threading.Lock()
_provider_ready = False


def init_tracer() -> trace.Tracer:
    global _provider_ready
    if not _provider_ready:
        with _provider_lock:
            if not _provider_ready:
                service_name = os.environ.get("OTEL_SERVICE_NAME", "pizarra-envasado-rp06")
                provider = TracerProvider(resource=Resource.create({"service.name": service_name}))
                trace.set_tracer_provider(provider)
                _provider_ready = True
    return trace.get_tracer("pizarra-envasado")


def record_timing(name: str, duration_ms: float) -> None:
    """Registra una métrica de tiempo para la cabecera Server-Timing."""
    try:
        timings = getattr(g, "server_timings", None)
        if timings is None:
            timings = {}
            g.server_timings = timings
        timings[name] = round(timings.get(name, 0.0) + max(0.0, duration_ms), 2)
    except RuntimeError:
        pass


def init_observabilidad(app: Any) -> None:
    tracer = init_tracer()

    @app.before_request
    def _otel_start_request():
        g.req_start_perf = time.perf_counter()
        g.server_timings = {}

        parent_ctx = None
        tp = (request.headers.get("traceparent") or "").strip()
        if tp:
            parts = tp.split("-")
            if len(parts) == 4 and len(parts[1]) == 32 and len(parts[2]) == 16:
                try:
                    tid = int(parts[1], 16)
                    sid = int(parts[2], 16)
                    if tid and sid:
                        sc = SpanContext(
                            trace_id=tid,
                            span_id=sid,
                            is_remote=True,
                            trace_flags=TraceFlags(TraceFlags.SAMPLED),
                        )
                        parent_ctx = trace.set_span_in_context(trace.NonRecordingSpan(sc))
                except ValueError:
                    parent_ctx = None

        span = tracer.start_span(
            f"{request.method} {request.path}",
            context=parent_ctx,
            attributes={
                "http.method": request.method,
                "http.target": request.path,
            },
        )
        g.otel_span = span
        ctx = span.get_span_context()
        g.trace_id = f"{ctx.trace_id:032x}" if ctx and ctx.trace_id else ""

    @app.after_request
    def _otel_finish_request(response: Response) -> Response:
        start_perf = getattr(g, "req_start_perf", None)
        total_ms = round((time.perf_counter() - start_perf) * 1000.0, 2) if start_perf is not None else 0.0
        timings: Dict[str, float] = getattr(g, "server_timings", {}) or {}
        timings["total"] = total_ms

        parts = [f"{k};dur={v:.2f}" for k, v in timings.items()]
        if parts:
            existing = response.headers.get("Server-Timing")
            header_val = ", ".join(parts)
            response.headers["Server-Timing"] = f"{existing}, {header_val}" if existing else header_val

        trace_id = getattr(g, "trace_id", "")
        if trace_id:
            response.headers["X-Trace-Id"] = trace_id

        span = getattr(g, "otel_span", None)
        if span is not None:
            try:
                span.set_attribute("http.status_code", response.status_code)
                span.end()
            except Exception:
                pass
            g.otel_span = None
        return response


class TTLCache:
    """Caché TTL thread-safe en memoria con versionado e invalidación inmediata."""

    def __init__(self, default_ttl_s: float = 15.0, max_entries: int = 256) -> None:
        self._default_ttl_s = default_ttl_s
        self._max_entries = max_entries
        self._lock = threading.RLock()
        self._store: Dict[str, Tuple[float, Any, str]] = {}
        self._version = 1

    @property
    def version(self) -> int:
        with self._lock:
            return self._version

    def get(self, key: str) -> Tuple[bool, Any, str]:
        now = time.monotonic()
        with self._lock:
            item = self._store.get(key)
            if not item:
                return False, None, ""
            exp, value, etag = item
            if exp <= now:
                self._store.pop(key, None)
                return False, None, ""
            return True, value, etag

    def set(self, key: str, value: Any, etag: str = "", ttl_s: Optional[float] = None) -> str:
        ttl = self._default_ttl_s if ttl_s is None else ttl_s
        now = time.monotonic()
        with self._lock:
            if len(self._store) >= self._max_entries and key not in self._store:
                expired = [k for k, (e, _, _) in self._store.items() if e <= now]
                for k in expired:
                    self._store.pop(k, None)
                if len(self._store) >= self._max_entries:
                    self._store.pop(next(iter(self._store)), None)
            if not etag:
                etag = f'W/"v{self._version}-{abs(hash(key)) & 0xFFFFFFFF:08x}"'
            self._store[key] = (now + ttl, value, etag)
            return etag

    def set_if_version(
        self,
        key: str,
        value: Any,
        expected_version: int,
        etag: str = "",
        ttl_s: Optional[float] = None,
    ) -> bool:
        """Guarda en caché solo si no ocurrió una invalidación concurrente desde expected_version."""
        ttl = self._default_ttl_s if ttl_s is None else ttl_s
        now = time.monotonic()
        with self._lock:
            if self._version != expected_version:
                return False
            if len(self._store) >= self._max_entries and key not in self._store:
                expired = [k for k, (e, _, _) in self._store.items() if e <= now]
                for k in expired:
                    self._store.pop(k, None)
                if len(self._store) >= self._max_entries:
                    self._store.pop(next(iter(self._store)), None)
            if not etag:
                etag = f'W/"v{self._version}-{abs(hash(key)) & 0xFFFFFFFF:08x}"'
            self._store[key] = (now + ttl, value, etag)
            return True

    def invalidate(self, prefix: Optional[str] = None) -> int:
        with self._lock:
            self._version += 1
            if prefix is None:
                self._store.clear()
            else:
                for k in [k for k in self._store if k.startswith(prefix)]:
                    self._store.pop(k, None)
            return self._version


def compute_etag(raw_bytes: bytes, prefix: str = "r") -> str:
    digest = hashlib.sha256(raw_bytes).hexdigest()[:16]
    return f'W/"{prefix}-{digest}"'


def etag_matches(if_none_match: Optional[str], current_etag: str) -> bool:
    if not if_none_match or not current_etag:
        return False
    raw = if_none_match.strip()
    if raw == "*":
        return True
    norm_current = current_etag.removeprefix("W/").strip()
    for tag in raw.split(","):
        t = tag.strip()
        if t == current_etag or t.removeprefix("W/").strip() == norm_current:
            return True
    return False
