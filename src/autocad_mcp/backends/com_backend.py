"""COM/ActiveX backend for full AutoCAD on Windows.

Drives a real AutoCAD session through the ``AutoCAD.Application`` COM
server. Unlike the file-IPC backend this needs no LISP dispatcher loaded
into AutoCAD, and unlike the ezdxf backend it operates on the live
drawing. AutoCAD is launched automatically on first use unless
``AUTOCAD_MCP_COM_AUTOLAUNCH=0``.

All COM calls run on one dedicated STA thread; the asyncio side only
awaits their results, so a slow AutoCAD never blocks the event loop.
"""

from __future__ import annotations

import asyncio
import concurrent.futures
import math
import queue
import threading
import time
from typing import Any, Callable

import structlog

from autocad_mcp.backends.base import AutoCADBackend, BackendCapabilities, CommandResult
from autocad_mcp.screenshot import NullScreenshotProvider

log = structlog.get_logger()

# HRESULTs worth special handling (normalized to unsigned 32-bit).
BUSY_HRESULTS = {
    0x80010001,  # RPC_E_CALL_REJECTED — AutoCAD busy, call never ran
    0x8001010A,  # RPC_E_SERVERCALL_RETRYLATER — modal dialog / command active
}
DISCONNECTED_HRESULTS = {
    0x80010108,  # RPC_E_DISCONNECTED
    0x800401FD,  # CO_E_OBJNOTCONNECTED
    0x800706BA,  # RPC_S_SERVER_UNAVAILABLE
}

ACI_COLORS = {
    "red": 1,
    "yellow": 2,
    "green": 3,
    "cyan": 4,
    "blue": 5,
    "magenta": 6,
    "white": 7,
    "grey": 8,
    "gray": 8,
}

_TYPE_MAP = {
    "AcDbLine": "LINE",
    "AcDbCircle": "CIRCLE",
    "AcDbArc": "ARC",
    "AcDbEllipse": "ELLIPSE",
    "AcDb2dPolyline": "POLYLINE",
    "AcDb3dPolyline": "POLYLINE",
    "AcDbPolyline": "LWPOLYLINE",
    "AcDbText": "TEXT",
    "AcDbMText": "MTEXT",
    "AcDbHatch": "HATCH",
    "AcDbLeader": "LEADER",
    "AcDbBlockReference": "INSERT",
    "AcDbRotatedDimension": "DIMENSION",
    "AcDbAlignedDimension": "DIMENSION",
    "AcDbRadialDimension": "DIMENSION",
    "AcDbDiametricDimension": "DIMENSION",
    "AcDbAngularDimension2": "DIMENSION",
}

PID_LAYERS = (
    ("PID-EQUIPMENT", 6),
    ("PID-PROCESS-PIPING", 4),
    ("PID-UTILITY-PIPING", 3),
    ("PID-INSTRUMENTS", 5),
    ("PID-ELECTRICAL", 1),
    ("PID-ANNOTATION", 7),
    ("PID-VALVES", 2),
)

# ---------------------------------------------------------------------------
# COM loading (lazy so the module imports on any platform and in tests)
# ---------------------------------------------------------------------------

_COM: dict[str, Any] = {}


def _load_com() -> dict[str, Any]:
    """Import pywin32 lazily and cache the module handles."""
    if not _COM:
        import pythoncom
        import pywintypes
        import win32com.client as win32_client

        _COM.update(pythoncom=pythoncom, pywintypes=pywintypes, client=win32_client)
    return _COM


def _r8(values: list[float]):
    """VARIANT array of doubles (point / coordinate arrays)."""
    com = _load_com()
    return com["client"].VARIANT(com["pythoncom"].VT_ARRAY | com["pythoncom"].VT_R8, list(values))


def _dispatch_array(objects: tuple):
    """VARIANT array of COM objects (hatch boundary loops)."""
    com = _load_com()
    return com["client"].VARIANT(
        com["pythoncom"].VT_ARRAY | com["pythoncom"].VT_DISPATCH, tuple(objects)
    )


def _hresult(ex: BaseException) -> int | None:
    try:
        code = ex.args[0]
    except Exception:
        return None
    if not isinstance(code, int):
        return None
    return code & 0xFFFFFFFF


class ComTimeout(Exception):
    """A COM call did not return in time."""


class _Disconnected(Exception):
    """AutoCAD went away."""


def _friendly(ex: BaseException, context: str, hint: str | None = None) -> str:
    msg = str(ex).strip() or ex.__class__.__name__
    parts = [f"[{context}] {msg}" if context else msg]
    hr = _hresult(ex)
    if hr is not None:
        parts.append(f"HRESULT=0x{hr:08X}")
    if hint:
        parts.append(hint)
    return " | ".join(parts)


def _aci_color(color: str | int) -> int:
    if isinstance(color, int):
        return color
    text = str(color).strip().lower()
    if text.isdigit():
        return int(text)
    return ACI_COLORS.get(text, 7)


def _type_name(entity) -> str:
    raw = str(getattr(entity, "EntityName", "") or "")
    if raw in _TYPE_MAP:
        return _TYPE_MAP[raw]
    if raw.startswith("AcDb"):
        raw = raw[4:]
    return raw.upper()


def _safe_str(obj, attr: str) -> str | None:
    try:
        value = getattr(obj, attr)
    except Exception:
        return None
    return str(value) if value else None


def _flat_2d(points) -> list[float]:
    flat: list[float] = []
    for p in points:
        flat.extend((float(p[0]), float(p[1])))
    return flat


def _flat_3d(points) -> list[float]:
    flat: list[float] = []
    for p in points:
        z = float(p[2]) if len(p) > 2 else 0.0
        flat.extend((float(p[0]), float(p[1]), z))
    return flat


# ---------------------------------------------------------------------------
# STA worker thread
# ---------------------------------------------------------------------------


class _ComWorker:
    """Single thread that owns the COM apartment and runs jobs serially."""

    def __init__(self) -> None:
        self._queue: queue.Queue = queue.Queue()
        self._thread: threading.Thread | None = None
        self._init_error: BaseException | None = None
        self._lock = threading.Lock()

    def start(self) -> None:
        with self._lock:
            if self._thread is not None and self._thread.is_alive():
                return
            self._thread = threading.Thread(
                target=self._loop, name="autocad-com", daemon=True
            )
            self._thread.start()

    def _loop(self) -> None:
        pythoncom = None
        try:
            pythoncom = _load_com()["pythoncom"]
            pythoncom.CoInitialize()
        except BaseException as ex:  # pragma: no cover - platform dependent
            self._init_error = ex

        try:
            while True:
                item = self._queue.get()
                if item is None:
                    break
                fn, future = item
                if future.cancelled():
                    continue
                if self._init_error is not None:
                    future.set_exception(self._init_error)
                    continue
                try:
                    result = fn()
                except BaseException as ex:
                    self._set_exception(future, ex)
                else:
                    self._set_result(future, result)
        finally:
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    @staticmethod
    def _set_result(future: concurrent.futures.Future, value: Any) -> None:
        # A timed-out caller may have cancelled the future while the COM
        # call was still running; discarding the late result is expected.
        if future.cancelled():
            return
        try:
            future.set_result(value)
        except concurrent.futures.InvalidStateError:
            pass

    @staticmethod
    def _set_exception(future: concurrent.futures.Future, error: BaseException) -> None:
        if future.cancelled():
            return
        try:
            future.set_exception(error)
        except concurrent.futures.InvalidStateError:
            pass

    def submit(self, fn: Callable[[], Any]) -> concurrent.futures.Future:
        self.start()
        future: concurrent.futures.Future = concurrent.futures.Future()
        self._queue.put((fn, future))
        return future

    def stop(self) -> None:
        thread = self._thread
        if thread is not None and thread.is_alive():
            self._queue.put(None)

    @property
    def alive(self) -> bool:
        thread = self._thread
        return bool(thread is not None and thread.is_alive())


# ---------------------------------------------------------------------------
# Backend
# ---------------------------------------------------------------------------


class ComBackend(AutoCADBackend):
    """AutoCAD ActiveX backend."""

    def __init__(self, progid: str | None = None, autolaunch: bool | None = None):
        from autocad_mcp import config

        self._progid = progid or config.COM_PROGID
        self._autolaunch = config.COM_AUTOLAUNCH if autolaunch is None else autolaunch
        self._connect_timeout = config.COM_CONNECT_TIMEOUT
        self._call_timeout = config.COM_CALL_TIMEOUT
        self._max_retries = config.COM_RETRIES

        self._worker = _ComWorker()
        self._app = None
        self._hwnd = 0
        self._filedia_prev = None
        self._screenshot = NullScreenshotProvider()
        self._closed = False
        self._last_error: str | None = None
        self._retry_count = 0
        self._reconnect_count = 0

    @property
    def name(self) -> str:
        return "com"

    @property
    def capabilities(self) -> BackendCapabilities:
        return BackendCapabilities(
            can_read_drawing=True,
            can_modify_entities=True,
            can_create_entities=True,
            can_screenshot=True,
            can_save=True,
            can_plot_pdf=True,
            can_zoom=True,
            can_query_entities=True,
            can_file_operations=True,
            can_undo=True,
        )

    # -- lifecycle ---------------------------------------------------------

    async def initialize(self) -> CommandResult:
        try:
            info = await self._call(self._attach, timeout=self._connect_timeout, retries=0)
        except Exception as ex:
            self._last_error = _friendly(ex, "initialize")
            self._worker.stop()
            return CommandResult(ok=False, error=self._last_error)

        self._screenshot = self._make_screenshot_provider()
        return CommandResult(ok=True, payload={"backend": "com", **info})

    async def status(self) -> CommandResult:
        payload: dict[str, Any] = {
            "backend": "com",
            "progid": self._progid,
            "connected": bool(self._app is not None and not self._closed),
            "autolaunch": self._autolaunch,
            "worker_alive": self._worker.alive,
            "retry_count": self._retry_count,
            "reconnect_count": self._reconnect_count,
            "capabilities": dict(self.capabilities.__dict__),
        }
        if self._app is not None and not self._closed:
            try:
                payload.update(
                    await self._call(self._job_status, timeout=min(self._call_timeout, 15.0), retries=0)
                )
            except Exception as ex:
                payload["connected"] = False
                payload["last_error"] = _friendly(ex, "status")
        elif self._last_error:
            payload["last_error"] = self._last_error
        return CommandResult(ok=True, payload=payload)

    async def close(self) -> CommandResult:
        if self._app is not None and self._filedia_prev is not None:
            try:
                await self._call(self._job_restore_filedia, timeout=5.0, retries=0)
            except Exception:
                pass
        self._closed = True
        self._worker.stop()
        self._app = None
        return CommandResult(ok=True, payload={"closed": True})

    # -- call plumbing -----------------------------------------------------

    async def _call(self, fn: Callable[[], Any], timeout: float | None = None, retries: int | None = None):
        if self._closed:
            raise RuntimeError("COM backend is closed")

        timeout = self._call_timeout if timeout is None else timeout
        max_retries = self._max_retries if retries is None else retries
        attempt = 0

        while True:
            future = self._worker.submit(fn)
            try:
                return await asyncio.wait_for(asyncio.wrap_future(future), timeout=timeout)
            except asyncio.TimeoutError as ex:
                raise ComTimeout(
                    f"AutoCAD call did not return within {timeout:.0f}s "
                    "(AutoCAD may be in a modal dialog — press ESC in AutoCAD and retry)"
                ) from ex
            except Exception as ex:
                hr = _hresult(ex)
                if hr in BUSY_HRESULTS and attempt < max_retries:
                    attempt += 1
                    self._retry_count += 1
                    await asyncio.sleep(min(1.0, 0.1 * (2 ** attempt)))
                    continue
                if hr in DISCONNECTED_HRESULTS:
                    raise _Disconnected(str(ex)) from ex
                raise

    async def _exec(self, context: str, fn: Callable[[], Any], timeout: float | None = None) -> CommandResult:
        try:
            payload = await self._call(fn, timeout=timeout)
            return CommandResult(ok=True, payload=payload)
        except _Disconnected:
            try:
                await self._call(self._attach, timeout=self._connect_timeout, retries=0)
                self._reconnect_count += 1
                payload = await self._call(fn, timeout=timeout)
                return CommandResult(ok=True, payload=payload)
            except Exception as ex:
                return CommandResult(
                    ok=False,
                    error=_friendly(
                        ex,
                        context,
                        "AutoCAD was closed or stopped responding. Reopen it, or call system init.",
                    ),
                )
        except ComTimeout as ex:
            return CommandResult(ok=False, error=_friendly(ex, context))
        except Exception as ex:
            return CommandResult(ok=False, error=_friendly(ex, context))

    # -- connection --------------------------------------------------------

    def _attach(self) -> dict[str, Any]:
        com = _load_com()
        client = com["client"]

        app = None
        try:
            app = client.GetActiveObject(self._progid)
        except Exception as ex:
            if not self._autolaunch:
                raise RuntimeError(
                    f"AutoCAD is not running (ProgID {self._progid}). "
                    "Set AUTOCAD_MCP_COM_AUTOLAUNCH=1 to let the server start it."
                ) from ex
            try:
                app = client.Dispatch(self._progid)
            except Exception as ex2:
                raise RuntimeError(f"Failed to start AutoCAD via COM: {ex2}") from ex2

        deadline = time.time() + self._connect_timeout
        delay = 0.5
        version = None
        while True:
            try:
                version = str(app.Version)
                break
            except Exception as ex:
                if time.time() >= deadline:
                    raise RuntimeError(
                        f"AutoCAD did not become ready within {self._connect_timeout:.0f}s: {ex}"
                    ) from ex
                time.sleep(delay)
                delay = min(4.0, delay * 2.0)

        try:
            app.Visible = True
        except Exception:
            pass

        self._app = app
        self._hwnd = int(getattr(app, "HWND", 0) or 0)
        doc = self._ensure_doc_internal()
        try:
            self._filedia_prev = doc.GetVariable("FILEDIA")
            doc.SetVariable("FILEDIA", 0)
        except Exception:
            self._filedia_prev = None

        return {
            "version": version,
            "hwnd": self._hwnd,
            "document": _safe_str(doc, "Name"),
            "entity_count": self._count_internal(doc),
        }

    def _job_restore_filedia(self) -> dict[str, Any]:
        doc = self._doc_internal()
        doc.SetVariable("FILEDIA", self._filedia_prev)
        return {"filedia": self._filedia_prev}

    def _ensure_doc_internal(self):
        docs = self._app.Documents
        try:
            count = int(docs.Count)
        except Exception:
            count = 0
        if count <= 0:
            return docs.Add()
        try:
            return self._app.ActiveDocument
        except Exception as ex:
            if _hresult(ex) in DISCONNECTED_HRESULTS:
                raise
            return docs.Item(0)

    def _doc_internal(self):
        if self._app is None:
            raise RuntimeError("AutoCAD is not connected")
        try:
            return self._app.ActiveDocument
        except Exception as ex:
            if _hresult(ex) in DISCONNECTED_HRESULTS:
                raise
            return self._ensure_doc_internal()

    def _make_screenshot_provider(self):
        if not self._hwnd:
            return NullScreenshotProvider()
        try:
            from autocad_mcp.screenshot import Win32ScreenshotProvider

            return Win32ScreenshotProvider(self._hwnd)
        except Exception as ex:  # pragma: no cover - Windows only
            log.warning("screenshot_provider_failed", error=str(ex))
            return NullScreenshotProvider()

    # -- internal helpers --------------------------------------------------

    def _job_status(self) -> dict[str, Any]:
        doc = self._doc_internal()
        full_name = _safe_str(doc, "FullName")
        return {
            "autocad_version": str(self._app.Version),
            "hwnd": int(self._hwnd),
            "document_count": int(self._app.Documents.Count),
            "active_document": {
                "name": _safe_str(doc, "Name"),
                "full_name": full_name,
                "saved": bool(full_name),
            },
            "entity_count": self._count_internal(doc),
        }

    @staticmethod
    def _count_internal(doc) -> int:
        try:
            return int(doc.ModelSpace.Count)
        except Exception:
            return 0

    @staticmethod
    def _ensure_layer_internal(doc, name: str):
        try:
            return doc.Layers.Item(name)
        except Exception:
            return doc.Layers.Add(name)

    def _apply_layer(self, doc, entity, layer: str | None) -> None:
        if not layer:
            return
        self._ensure_layer_internal(doc, layer)
        entity.Layer = layer

    @staticmethod
    def _entity_internal(doc, handle: str):
        try:
            return doc.HandleToObject(str(handle).strip().upper())
        except Exception as ex:
            raise LookupError(
                f"Entity '{handle}' not found in the active drawing "
                "(it may have been deleted, or belong to another drawing)"
            ) from ex

    def _insert_block_internal(self, name: str, x: float, y: float, scale: float, rotation: float):
        doc = self._doc_internal()
        entity = doc.ModelSpace.InsertBlock(
            _r8([float(x), float(y), 0.0]),
            name,
            float(scale),
            float(scale),
            float(scale),
            math.radians(float(rotation)),
        )
        return doc, entity

    # -- drawing management ------------------------------------------------

    async def drawing_info(self) -> CommandResult:
        return await self._exec("drawing_info", self._job_drawing_info)

    def _job_drawing_info(self) -> dict[str, Any]:
        doc = self._doc_internal()
        layers = [str(layer.Name) for layer in doc.Layers]
        blocks = [str(block.Name) for block in doc.Blocks if not str(block.Name).startswith("*")]
        full_name = _safe_str(doc, "FullName")
        return {
            "document_name": _safe_str(doc, "Name"),
            "save_path": full_name,
            "entity_count": self._count_internal(doc),
            "layers": layers,
            "blocks": blocks,
            "autocad_version": str(self._app.Version),
        }

    async def drawing_save(self, path: str | None = None) -> CommandResult:
        return await self._exec("drawing_save", lambda: self._job_drawing_save(path))

    def _job_drawing_save(self, path: str | None) -> dict[str, Any]:
        doc = self._doc_internal()
        if path:
            doc.SaveAs(str(path))
            return {"path": str(path)}
        full_name = _safe_str(doc, "FullName")
        if not full_name:
            raise RuntimeError("Drawing has never been saved; provide a path")
        doc.Save()
        return {"path": full_name}

    async def drawing_save_as_dxf(self, path: str) -> CommandResult:
        return await self._exec("drawing_save_as_dxf", lambda: self._job_drawing_save(path))

    async def drawing_create(self, name: str | None = None) -> CommandResult:
        return await self._exec("drawing_create", lambda: self._job_drawing_create(name))

    def _job_drawing_create(self, name: str | None) -> dict[str, Any]:
        doc = self._app.Documents.Add()
        saved = False
        if name and ("/" in str(name) or "\\" in str(name) or str(name).lower().endswith(".dwg")):
            doc.SaveAs(str(name))
            saved = True
        return {
            "name": _safe_str(doc, "Name"),
            "saved": saved,
            "path": _safe_str(doc, "FullName"),
        }

    async def drawing_open(self, path: str) -> CommandResult:
        return await self._exec("drawing_open", lambda: self._job_drawing_open(path))

    def _job_drawing_open(self, path: str) -> dict[str, Any]:
        doc = self._app.Documents.Open(str(path))
        return {"path": str(path), "name": _safe_str(doc, "Name")}

    async def drawing_purge(self) -> CommandResult:
        return await self._exec("drawing_purge", self._job_drawing_purge)

    def _job_drawing_purge(self) -> dict[str, Any]:
        self._doc_internal().PurgeAll()
        return {"purged": True}

    async def drawing_plot_pdf(self, path: str) -> CommandResult:
        return await self._exec(
            "drawing_plot_pdf", lambda: self._job_drawing_plot_pdf(path), timeout=max(self._call_timeout, 120.0)
        )

    def _job_drawing_plot_pdf(self, path: str) -> dict[str, Any]:
        doc = self._doc_internal()
        plot = doc.Plot
        try:
            plot.QuietErrorMode = True
        except Exception:
            pass
        try:
            plot.PlotToFile(str(path))
        except Exception as ex:
            raise RuntimeError(
                f"Plot failed: {ex}. Make sure a page setup / plotter is configured in AutoCAD"
            ) from ex
        return {"path": str(path)}

    async def drawing_get_variables(self, names: list[str] | None = None) -> CommandResult:
        return await self._exec("drawing_get_variables", lambda: self._job_get_variables(names))

    def _job_get_variables(self, names: list[str] | None) -> dict[str, Any]:
        doc = self._doc_internal()
        result: dict[str, Any] = {}
        for name in names or []:
            key = str(name).lstrip("$")
            try:
                result[str(name)] = str(doc.GetVariable(key))
            except Exception:
                result[str(name)] = None
        return result

    # -- undo / redo / lisp ------------------------------------------------

    async def undo(self) -> CommandResult:
        return await self._exec("undo", lambda: self._job_send_command("_.UNDO 1 "))

    async def redo(self) -> CommandResult:
        return await self._exec("redo", lambda: self._job_send_command("_.REDO "))

    async def execute_lisp(self, code: str) -> CommandResult:
        return await self._exec("execute_lisp", lambda: self._job_send_command(str(code)))

    def _job_send_command(self, command: str) -> dict[str, Any]:
        doc = self._doc_internal()
        doc.SendCommand(command if command.endswith("\n") else command + "\n")
        return {"queued": True, "note": "SendCommand is asynchronous; no return value is captured"}

    # -- entity creation ---------------------------------------------------

    async def create_line(self, x1, y1, x2, y2, layer=None) -> CommandResult:
        return await self._exec(
            "create_line", lambda: self._job_create_line(x1, y1, x2, y2, layer)
        )

    def _job_create_line(self, x1, y1, x2, y2, layer) -> dict[str, Any]:
        doc = self._doc_internal()
        entity = doc.ModelSpace.AddLine(_r8([float(x1), float(y1), 0.0]), _r8([float(x2), float(y2), 0.0]))
        self._apply_layer(doc, entity, layer)
        return {"entity_type": "LINE", "handle": entity.Handle}

    async def create_circle(self, cx, cy, radius, layer=None) -> CommandResult:
        return await self._exec(
            "create_circle", lambda: self._job_create_circle(cx, cy, radius, layer)
        )

    def _job_create_circle(self, cx, cy, radius, layer) -> dict[str, Any]:
        doc = self._doc_internal()
        entity = doc.ModelSpace.AddCircle(_r8([float(cx), float(cy), 0.0]), float(radius))
        self._apply_layer(doc, entity, layer)
        return {"entity_type": "CIRCLE", "handle": entity.Handle}

    async def create_polyline(self, points, closed=False, layer=None) -> CommandResult:
        return await self._exec(
            "create_polyline", lambda: self._job_create_polyline(points, closed, layer)
        )

    def _job_create_polyline(self, points, closed, layer) -> dict[str, Any]:
        if not points:
            raise ValueError("points must not be empty")
        doc = self._doc_internal()
        entity = doc.ModelSpace.AddLightWeightPolyline(_r8(_flat_2d(points)))
        entity.Closed = bool(closed)
        self._apply_layer(doc, entity, layer)
        return {"entity_type": "LWPOLYLINE", "handle": entity.Handle}

    async def create_rectangle(self, x1, y1, x2, y2, layer=None) -> CommandResult:
        pts = [[x1, y1], [x2, y1], [x2, y2], [x1, y2]]
        return await self._exec(
            "create_rectangle", lambda: self._job_create_polyline(pts, True, layer)
        )

    async def create_arc(self, cx, cy, radius, start_angle, end_angle, layer=None) -> CommandResult:
        return await self._exec(
            "create_arc",
            lambda: self._job_create_arc(cx, cy, radius, start_angle, end_angle, layer),
        )

    def _job_create_arc(self, cx, cy, radius, start_angle, end_angle, layer) -> dict[str, Any]:
        doc = self._doc_internal()
        entity = doc.ModelSpace.AddArc(
            _r8([float(cx), float(cy), 0.0]),
            float(radius),
            math.radians(float(start_angle)),
            math.radians(float(end_angle)),
        )
        self._apply_layer(doc, entity, layer)
        return {"entity_type": "ARC", "handle": entity.Handle}

    async def create_ellipse(self, cx, cy, major_x, major_y, ratio, layer=None) -> CommandResult:
        return await self._exec(
            "create_ellipse",
            lambda: self._job_create_ellipse(cx, cy, major_x, major_y, ratio, layer),
        )

    def _job_create_ellipse(self, cx, cy, major_x, major_y, ratio, layer) -> dict[str, Any]:
        doc = self._doc_internal()
        entity = doc.ModelSpace.AddEllipse(
            _r8([float(cx), float(cy), 0.0]),
            _r8([float(major_x) - float(cx), float(major_y) - float(cy), 0.0]),
            float(ratio),
        )
        self._apply_layer(doc, entity, layer)
        return {"entity_type": "ELLIPSE", "handle": entity.Handle}

    async def create_mtext(self, x, y, width, text, height=2.5, layer=None) -> CommandResult:
        return await self._exec(
            "create_mtext", lambda: self._job_create_mtext(x, y, width, text, height, layer)
        )

    def _job_create_mtext(self, x, y, width, text, height, layer) -> dict[str, Any]:
        doc = self._doc_internal()
        entity = doc.ModelSpace.AddMText(_r8([float(x), float(y), 0.0]), float(width), str(text))
        try:
            entity.Height = float(height)
        except Exception:
            pass
        self._apply_layer(doc, entity, layer)
        return {"entity_type": "MTEXT", "handle": entity.Handle}

    async def create_hatch(self, entity_id: str, pattern: str = "ANSI31") -> CommandResult:
        return await self._exec(
            "create_hatch", lambda: self._job_create_hatch(entity_id, pattern)
        )

    def _job_create_hatch(self, entity_id: str, pattern: str) -> dict[str, Any]:
        doc = self._doc_internal()
        boundary = self._entity_internal(doc, entity_id)
        hatch = doc.ModelSpace.AddHatch(0, str(pattern), True)
        hatch.AppendOuterLoop(_dispatch_array((boundary,)))
        hatch.Evaluate()
        return {"entity_type": "HATCH", "handle": hatch.Handle}

    # -- entity query / modify --------------------------------------------

    async def entity_list(self, layer=None) -> CommandResult:
        return await self._exec("entity_list", lambda: self._job_entity_list(layer))

    def _job_entity_list(self, layer) -> dict[str, Any]:
        doc = self._doc_internal()
        entities = []
        for entity in doc.ModelSpace:
            entity_layer = str(entity.Layer)
            if layer and entity_layer != str(layer):
                continue
            entities.append(
                {"type": _type_name(entity), "handle": entity.Handle, "layer": entity_layer}
            )
        return {"entities": entities, "count": len(entities)}

    async def entity_count(self, layer=None) -> CommandResult:
        return await self._exec("entity_count", lambda: self._job_entity_count(layer))

    def _job_entity_count(self, layer) -> dict[str, Any]:
        doc = self._doc_internal()
        count = 0
        for entity in doc.ModelSpace:
            if layer and str(entity.Layer) != str(layer):
                continue
            count += 1
        return {"count": count}

    async def entity_get(self, entity_id: str) -> CommandResult:
        return await self._exec("entity_get", lambda: self._job_entity_get(entity_id))

    def _job_entity_get(self, entity_id: str) -> dict[str, Any]:
        doc = self._doc_internal()
        entity = self._entity_internal(doc, entity_id)
        kind = _type_name(entity)
        info: dict[str, Any] = {"type": kind, "handle": entity.Handle, "layer": str(entity.Layer)}
        if kind == "LINE":
            info["start"] = list(entity.StartPoint)[:2]
            info["end"] = list(entity.EndPoint)[:2]
        elif kind == "CIRCLE":
            info["center"] = list(entity.Center)[:2]
            info["radius"] = float(entity.Radius)
        return info

    async def entity_erase(self, entity_id: str) -> CommandResult:
        return await self._exec("entity_erase", lambda: self._job_entity_erase(entity_id))

    def _job_entity_erase(self, entity_id: str) -> dict[str, Any]:
        doc = self._doc_internal()
        self._entity_internal(doc, entity_id).Delete()
        return {"erased": str(entity_id)}

    async def entity_copy(self, entity_id: str, dx, dy) -> CommandResult:
        return await self._exec(
            "entity_copy", lambda: self._job_entity_copy(entity_id, dx, dy)
        )

    def _job_entity_copy(self, entity_id: str, dx, dy) -> dict[str, Any]:
        doc = self._doc_internal()
        clone = self._entity_internal(doc, entity_id).Copy()
        doc.ModelSpace.AppendEntity(clone)
        clone.Move(_r8([0.0, 0.0, 0.0]), _r8([float(dx), float(dy), 0.0]))
        return {"entity_type": _type_name(clone), "handle": clone.Handle, "source": str(entity_id)}

    async def entity_move(self, entity_id: str, dx, dy) -> CommandResult:
        return await self._exec(
            "entity_move", lambda: self._job_entity_move(entity_id, dx, dy)
        )

    def _job_entity_move(self, entity_id: str, dx, dy) -> dict[str, Any]:
        doc = self._doc_internal()
        self._entity_internal(doc, entity_id).Move(
            _r8([0.0, 0.0, 0.0]), _r8([float(dx), float(dy), 0.0])
        )
        return {"moved": str(entity_id), "dx": dx, "dy": dy}

    async def entity_rotate(self, entity_id: str, cx, cy, angle) -> CommandResult:
        return await self._exec(
            "entity_rotate", lambda: self._job_entity_rotate(entity_id, cx, cy, angle)
        )

    def _job_entity_rotate(self, entity_id: str, cx, cy, angle) -> dict[str, Any]:
        doc = self._doc_internal()
        self._entity_internal(doc, entity_id).Rotate(
            _r8([float(cx), float(cy), 0.0]), math.radians(float(angle))
        )
        return {"rotated": str(entity_id), "angle": angle}

    async def entity_scale(self, entity_id: str, cx, cy, factor) -> CommandResult:
        return await self._exec(
            "entity_scale", lambda: self._job_entity_scale(entity_id, cx, cy, factor)
        )

    def _job_entity_scale(self, entity_id: str, cx, cy, factor) -> dict[str, Any]:
        doc = self._doc_internal()
        self._entity_internal(doc, entity_id).ScaleEntity(
            _r8([float(cx), float(cy), 0.0]), float(factor)
        )
        return {"scaled": str(entity_id), "factor": factor}

    async def entity_mirror(self, entity_id: str, x1, y1, x2, y2) -> CommandResult:
        return await self._exec(
            "entity_mirror", lambda: self._job_entity_mirror(entity_id, x1, y1, x2, y2)
        )

    def _job_entity_mirror(self, entity_id: str, x1, y1, x2, y2) -> dict[str, Any]:
        doc = self._doc_internal()
        clone = self._entity_internal(doc, entity_id).Mirror(
            _r8([float(x1), float(y1), 0.0]), _r8([float(x2), float(y2), 0.0])
        )
        return {"entity_type": _type_name(clone), "handle": clone.Handle, "source": str(entity_id)}

    async def entity_offset(self, entity_id: str, distance) -> CommandResult:
        return await self._exec(
            "entity_offset", lambda: self._job_entity_offset(entity_id, distance)
        )

    def _job_entity_offset(self, entity_id: str, distance) -> dict[str, Any]:
        doc = self._doc_internal()
        result = self._entity_internal(doc, entity_id).Offset(float(distance))
        try:
            items = list(result)
        except TypeError:
            items = [result]
        handles = []
        for item in items:
            try:
                doc.ModelSpace.AppendEntity(item)
            except Exception:
                pass
            handles.append(item.Handle)
        return {"handles": handles, "count": len(handles)}

    async def entity_array(self, entity_id: str, rows, cols, row_dist, col_dist) -> CommandResult:
        return await self._exec(
            "entity_array",
            lambda: self._job_entity_array(entity_id, rows, cols, row_dist, col_dist),
        )

    def _job_entity_array(self, entity_id: str, rows, cols, row_dist, col_dist) -> dict[str, Any]:
        doc = self._doc_internal()
        self._entity_internal(doc, entity_id).ArrayRectangular(
            int(rows), int(cols), 1, float(row_dist), float(col_dist), 0.0
        )
        return {"arrayed": str(entity_id), "rows": rows, "cols": cols}

    async def entity_fillet(self, entity_id1: str, entity_id2: str, radius) -> CommandResult:
        return CommandResult(
            ok=False,
            error=(
                "FILLET is not exposed by the AutoCAD ActiveX API. "
                "Run the FILLET command in AutoCAD, or use the file_ipc backend."
            ),
        )

    async def entity_chamfer(self, entity_id1: str, entity_id2: str, dist1, dist2) -> CommandResult:
        return CommandResult(
            ok=False,
            error=(
                "CHAMFER is not exposed by the AutoCAD ActiveX API. "
                "Run the CHAMFER command in AutoCAD, or use the file_ipc backend."
            ),
        )

    # -- layers ------------------------------------------------------------

    async def layer_list(self) -> CommandResult:
        return await self._exec("layer_list", self._job_layer_list)

    def _job_layer_list(self) -> dict[str, Any]:
        doc = self._doc_internal()
        layers = []
        for layer in doc.Layers:
            frozen = getattr(layer, "Freeze", None)
            locked = getattr(layer, "Lock", None)
            layers.append(
                {
                    "name": str(layer.Name),
                    "color": int(layer.Color),
                    "linetype": str(layer.Linetype),
                    "is_frozen": bool(frozen) if isinstance(frozen, bool) else None,
                    "is_locked": bool(locked) if isinstance(locked, bool) else None,
                }
            )
        return {"layers": layers}

    async def layer_create(self, name, color="white", linetype="CONTINUOUS") -> CommandResult:
        return await self._exec(
            "layer_create", lambda: self._job_layer_create(name, color, linetype)
        )

    def _job_layer_create(self, name, color, linetype) -> dict[str, Any]:
        doc = self._doc_internal()
        existed = True
        try:
            layer = doc.Layers.Item(str(name))
        except Exception:
            layer = doc.Layers.Add(str(name))
            existed = False
        aci = _aci_color(color)
        if not existed:
            try:
                layer.Color = aci
            except Exception:
                pass
            try:
                layer.Linetype = str(linetype)
            except Exception:
                pass
        return {"name": str(name), "color": int(layer.Color), "existed": existed}

    async def layer_set_current(self, name) -> CommandResult:
        return await self._exec("layer_set_current", lambda: self._job_layer_set_current(name))

    def _job_layer_set_current(self, name) -> dict[str, Any]:
        doc = self._doc_internal()
        doc.ActiveLayer = doc.Layers.Item(str(name))
        return {"current_layer": str(name)}

    async def layer_set_properties(self, name, color=None, linetype=None, lineweight=None) -> CommandResult:
        return await self._exec(
            "layer_set_properties",
            lambda: self._job_layer_set_properties(name, color, linetype, lineweight),
        )

    def _job_layer_set_properties(self, name, color, linetype, lineweight) -> dict[str, Any]:
        doc = self._doc_internal()
        layer = doc.Layers.Item(str(name))
        if color is not None:
            layer.Color = _aci_color(color)
        if linetype is not None:
            layer.Linetype = str(linetype)
        if lineweight is not None:
            layer.Lineweight = int(round(float(lineweight) * 100))
        return {"name": str(name)}

    async def layer_freeze(self, name) -> CommandResult:
        return await self._exec("layer_freeze", lambda: self._job_layer_freeze(name, True))

    async def layer_thaw(self, name) -> CommandResult:
        return await self._exec("layer_thaw", lambda: self._job_layer_freeze(name, False))

    def _job_layer_freeze(self, name, frozen: bool) -> dict[str, Any]:
        doc = self._doc_internal()
        layer = doc.Layers.Item(str(name))
        try:
            layer.Freeze = frozen
            via = "Freeze"
        except Exception:
            layer.LayerOn = not frozen
            via = "LayerOn"
        return {"name": str(name), "frozen": frozen, "via": via}

    async def layer_lock(self, name) -> CommandResult:
        return await self._exec("layer_lock", lambda: self._job_layer_lock(name, True))

    async def layer_unlock(self, name) -> CommandResult:
        return await self._exec("layer_unlock", lambda: self._job_layer_lock(name, False))

    def _job_layer_lock(self, name, locked: bool) -> dict[str, Any]:
        doc = self._doc_internal()
        layer = doc.Layers.Item(str(name))
        layer.Lock = locked
        return {"name": str(name), "locked": locked}

    # -- blocks ------------------------------------------------------------

    async def block_list(self) -> CommandResult:
        return await self._exec("block_list", self._job_block_list)

    def _job_block_list(self) -> dict[str, Any]:
        doc = self._doc_internal()
        blocks = []
        for block in doc.Blocks:
            name = str(block.Name)
            if name.startswith("*"):
                continue
            blocks.append({"name": name})
        return {"blocks": blocks, "count": len(blocks)}

    async def block_insert(self, name, x, y, scale=1.0, rotation=0.0, block_id=None) -> CommandResult:
        return await self._exec(
            "block_insert", lambda: self._job_block_insert(name, x, y, scale, rotation)
        )

    def _job_block_insert(self, name, x, y, scale, rotation) -> dict[str, Any]:
        _, entity = self._insert_block_internal(name, x, y, scale, rotation)
        return {"entity_type": "INSERT", "handle": entity.Handle, "block": str(name)}

    async def block_insert_with_attributes(self, name, x, y, scale=1.0, rotation=0.0, attributes=None) -> CommandResult:
        return await self._exec(
            "block_insert_with_attributes",
            lambda: self._job_block_insert_attributes(name, x, y, scale, rotation, attributes),
        )

    def _job_block_insert_attributes(self, name, x, y, scale, rotation, attributes) -> dict[str, Any]:
        _, entity = self._insert_block_internal(name, x, y, scale, rotation)
        applied = []
        if attributes:
            for attr in entity.GetAttributes():
                tag = str(attr.TagString)
                if tag in attributes:
                    attr.TextString = str(attributes[tag])
                    applied.append(tag)
        return {"entity_type": "INSERT", "handle": entity.Handle, "attributes": applied}

    async def block_get_attributes(self, entity_id: str) -> CommandResult:
        return await self._exec(
            "block_get_attributes", lambda: self._job_block_get_attributes(entity_id)
        )

    def _job_block_get_attributes(self, entity_id: str) -> dict[str, Any]:
        doc = self._doc_internal()
        entity = self._entity_internal(doc, entity_id)
        attributes = {str(attr.TagString): str(attr.TextString) for attr in entity.GetAttributes()}
        return {"handle": str(entity_id), "attributes": attributes}

    async def block_update_attribute(self, entity_id: str, tag: str, value: str) -> CommandResult:
        return await self._exec(
            "block_update_attribute",
            lambda: self._job_block_update_attribute(entity_id, tag, value),
        )

    def _job_block_update_attribute(self, entity_id: str, tag: str, value: str) -> dict[str, Any]:
        doc = self._doc_internal()
        entity = self._entity_internal(doc, entity_id)
        for attr in entity.GetAttributes():
            if str(attr.TagString) == str(tag):
                attr.TextString = str(value)
                return {"handle": str(entity_id), "tag": str(tag), "value": str(value)}
        raise LookupError(f"Block attribute '{tag}' not found on entity {entity_id}")

    async def block_define(self, name: str, entities: list[dict]) -> CommandResult:
        return await self._exec("block_define", lambda: self._job_block_define(name, entities))

    def _job_block_define(self, name: str, entities: list[dict]) -> dict[str, Any]:
        doc = self._doc_internal()
        block = doc.Blocks.Add(_r8([0.0, 0.0, 0.0]), str(name))
        created = 0
        for spec in entities or []:
            kind = str(spec.get("type", "")).lower()
            if kind == "line":
                block.AddLine(
                    _r8([float(spec["x1"]), float(spec["y1"]), 0.0]),
                    _r8([float(spec["x2"]), float(spec["y2"]), 0.0]),
                )
            elif kind == "circle":
                block.AddCircle(
                    _r8([float(spec["cx"]), float(spec["cy"]), 0.0]), float(spec["radius"])
                )
            elif kind == "text":
                block.AddText(
                    str(spec["text"]),
                    _r8([float(spec["x"]), float(spec["y"]), 0.0]),
                    float(spec.get("height", 2.5)),
                )
            else:
                raise ValueError(
                    f"block_define does not support entity type '{kind}' on the COM backend "
                    "(supported: line, circle, text)"
                )
            created += 1
        return {"block": str(name), "entities": created}

    # -- annotation --------------------------------------------------------

    async def create_text(self, x, y, text, height=2.5, rotation=0.0, layer=None) -> CommandResult:
        return await self._exec(
            "create_text", lambda: self._job_create_text(x, y, text, height, rotation, layer)
        )

    def _job_create_text(self, x, y, text, height, rotation, layer) -> dict[str, Any]:
        doc = self._doc_internal()
        entity = doc.ModelSpace.AddText(
            str(text), _r8([float(x), float(y), 0.0]), float(height)
        )
        entity.Rotation = math.radians(float(rotation))
        self._apply_layer(doc, entity, layer)
        return {"entity_type": "TEXT", "handle": entity.Handle}

    async def create_dimension_linear(self, x1, y1, x2, y2, dim_x, dim_y) -> CommandResult:
        return await self._exec(
            "create_dimension_linear",
            lambda: self._job_dim_linear(x1, y1, x2, y2, dim_x, dim_y),
        )

    def _job_dim_linear(self, x1, y1, x2, y2, dim_x, dim_y) -> dict[str, Any]:
        doc = self._doc_internal()
        mid_x = (float(x1) + float(x2)) / 2.0
        mid_y = (float(y1) + float(y2)) / 2.0
        # DIMLINEAR picks horizontal/vertical from where the dimension line sits.
        rotation = (
            math.pi / 2.0
            if abs(float(dim_x) - mid_x) > abs(float(dim_y) - mid_y)
            else 0.0
        )
        try:
            # AutoCAD 2027 (R26) does not expose AddDimLinear over ActiveX;
            # AddDimRotated is the equivalent call.
            entity = doc.ModelSpace.AddDimRotated(
                _r8([float(x1), float(y1), 0.0]),
                _r8([float(x2), float(y2), 0.0]),
                _r8([float(dim_x), float(dim_y), 0.0]),
                rotation,
            )
            return {
                "entity_type": "DIMENSION",
                "handle": entity.Handle,
                "rotation": math.degrees(rotation),
            }
        except Exception as ex:
            if _hresult(ex) in DISCONNECTED_HRESULTS:
                raise
            # Last resort: drive the DIMLINEAR command directly.
            doc.SendCommand(
                "_.DIMLINEAR non {:.6f},{:.6f} non {:.6f},{:.6f} non {:.6f},{:.6f}\n".format(
                    float(x1), float(y1), float(x2), float(y2), float(dim_x), float(dim_y)
                )
            )
            return {
                "entity_type": "DIMENSION",
                "via": "sendcommand",
                "queued": True,
                "note": "AddDimRotated failed; DIMLINEAR was queued in AutoCAD",
            }

    async def create_dimension_aligned(self, x1, y1, x2, y2, offset) -> CommandResult:
        return await self._exec(
            "create_dimension_aligned",
            lambda: self._job_dim_aligned(x1, y1, x2, y2, offset),
        )

    def _job_dim_aligned(self, x1, y1, x2, y2, offset) -> dict[str, Any]:
        doc = self._doc_internal()
        dx, dy = float(x2) - float(x1), float(y2) - float(y1)
        length = math.hypot(dx, dy)
        if length == 0:
            raise ValueError("Dimension endpoints must not coincide")
        nx, ny = dy / length, -dx / length
        mx, my = (float(x1) + float(x2)) / 2.0, (float(y1) + float(y2)) / 2.0
        location = _r8([mx + nx * float(offset), my + ny * float(offset), 0.0])
        entity = doc.ModelSpace.AddDimAligned(
            _r8([float(x1), float(y1), 0.0]),
            _r8([float(x2), float(y2), 0.0]),
            location,
        )
        return {"entity_type": "DIMENSION", "handle": entity.Handle}

    async def create_dimension_angular(self, cx, cy, x1, y1, x2, y2) -> CommandResult:
        return await self._exec(
            "create_dimension_angular",
            lambda: self._job_dim_angular(cx, cy, x1, y1, x2, y2),
        )

    def _job_dim_angular(self, cx, cy, x1, y1, x2, y2) -> dict[str, Any]:
        doc = self._doc_internal()
        a1 = math.atan2(float(y1) - float(cy), float(x1) - float(cx))
        a2 = math.atan2(float(y2) - float(cy), float(x2) - float(cx))
        amid = (a1 + a2) / 2.0
        radius = max(
            math.hypot(float(x1) - float(cx), float(y1) - float(cy)),
            math.hypot(float(x2) - float(cx), float(y2) - float(cy)),
        ) * 0.7
        tx = float(cx) + radius * math.cos(amid)
        ty = float(cy) + radius * math.sin(amid)
        entity = doc.ModelSpace.AddDimAngular(
            _r8([float(cx), float(cy), 0.0]),
            _r8([float(x1), float(y1), 0.0]),
            _r8([float(x2), float(y2), 0.0]),
            _r8([tx, ty, 0.0]),
        )
        return {"entity_type": "DIMENSION", "handle": entity.Handle}

    async def create_dimension_radius(self, cx, cy, radius, angle) -> CommandResult:
        return await self._exec(
            "create_dimension_radius",
            lambda: self._job_dim_radius(cx, cy, radius, angle),
        )

    def _job_dim_radius(self, cx, cy, radius, angle) -> dict[str, Any]:
        doc = self._doc_internal()
        rad = math.radians(float(angle))
        px = float(cx) + float(radius) * math.cos(rad)
        py = float(cy) + float(radius) * math.sin(rad)
        entity = doc.ModelSpace.AddDimRadial(
            _r8([float(cx), float(cy), 0.0]), _r8([px, py, 0.0]), 0.0
        )
        return {"entity_type": "DIMENSION", "handle": entity.Handle}

    async def create_dimension_diameter(self, cx, cy, radius, angle=0.0) -> CommandResult:
        return await self._exec(
            "create_dimension_diameter",
            lambda: self._job_dim_diameter(cx, cy, radius, angle),
        )

    def _job_dim_diameter(self, cx, cy, radius, angle) -> dict[str, Any]:
        doc = self._doc_internal()
        rad = math.radians(float(angle))
        p1 = [float(cx) + float(radius) * math.cos(rad), float(cy) + float(radius) * math.sin(rad), 0.0]
        p2 = [float(cx) - float(radius) * math.cos(rad), float(cy) - float(radius) * math.sin(rad), 0.0]
        entity = doc.ModelSpace.AddDimDiametric(_r8(p1), _r8(p2), 0.0)
        return {"entity_type": "DIMENSION", "handle": entity.Handle}

    async def create_leader(self, points, text) -> CommandResult:
        return await self._exec("create_leader", lambda: self._job_create_leader(points, text))

    def _job_create_leader(self, points, text) -> dict[str, Any]:
        if not points:
            raise ValueError("points must not be empty")
        doc = self._doc_internal()
        leader = doc.ModelSpace.AddLeader(_r8(_flat_3d(points)))
        last = points[-1]
        doc.ModelSpace.AddMText(
            _r8([float(last[0]) + 2.0, float(last[1]), 0.0]), 30.0, str(text)
        )
        return {"entity_type": "LEADER", "handle": leader.Handle}

    # -- P&ID --------------------------------------------------------------

    async def pid_setup_layers(self) -> CommandResult:
        return await self._exec("pid_setup_layers", self._job_pid_setup_layers)

    def _job_pid_setup_layers(self) -> dict[str, Any]:
        doc = self._doc_internal()
        for name, color in PID_LAYERS:
            try:
                layer = doc.Layers.Item(name)
            except Exception:
                layer = doc.Layers.Add(name)
                try:
                    layer.Color = color
                except Exception:
                    pass
        return {"layers_created": len(PID_LAYERS)}

    async def pid_list_symbols(self, category) -> CommandResult:
        try:
            from autocad_mcp.pid.cto_library import list_symbols

            symbols = list_symbols(str(category))
            return CommandResult(
                ok=True,
                payload={"category": str(category), "symbols": symbols, "count": len(symbols)},
            )
        except Exception as ex:
            return CommandResult(ok=False, error=_friendly(ex, "pid_list_symbols"))

    async def pid_insert_symbol(self, category, symbol, x, y, scale=1.0, rotation=0.0) -> CommandResult:
        return await self._exec(
            "pid_insert_symbol",
            lambda: self._job_pid_insert_symbol(symbol, x, y, scale),
        )

    def _job_pid_insert_symbol(self, symbol, x, y, scale) -> dict[str, Any]:
        doc = self._doc_internal()
        msp = doc.ModelSpace
        half = 5.0 * float(scale)
        points = [
            [float(x) - half, float(y) - half],
            [float(x) + half, float(y) - half],
            [float(x) + half, float(y) + half],
            [float(x) - half, float(y) + half],
        ]
        entity = msp.AddLightWeightPolyline(_r8(_flat_2d(points)))
        entity.Closed = True
        self._ensure_layer_internal(doc, "PID-EQUIPMENT")
        entity.Layer = "PID-EQUIPMENT"
        self._ensure_layer_internal(doc, "PID-ANNOTATION")
        text = msp.AddText(str(symbol), _r8([float(x), float(y), 0.0]), 1.5 * float(scale))
        text.Layer = "PID-ANNOTATION"
        return {"symbol": str(symbol), "handle": entity.Handle}

    async def pid_draw_process_line(self, x1, y1, x2, y2) -> CommandResult:
        return await self._exec(
            "pid_draw_process_line",
            lambda: self._job_pid_line(x1, y1, x2, y2, "PID-PROCESS-PIPING"),
        )

    async def pid_connect_equipment(self, x1, y1, x2, y2) -> CommandResult:
        return await self._exec(
            "pid_connect_equipment",
            lambda: self._job_pid_connect(x1, y1, x2, y2),
        )

    def _job_pid_line(self, x1, y1, x2, y2, layer) -> dict[str, Any]:
        doc = self._doc_internal()
        self._ensure_layer_internal(doc, layer)
        entity = doc.ModelSpace.AddLine(_r8([float(x1), float(y1), 0.0]), _r8([float(x2), float(y2), 0.0]))
        entity.Layer = layer
        return {"entity_type": "LINE", "handle": entity.Handle}

    def _job_pid_connect(self, x1, y1, x2, y2) -> dict[str, Any]:
        doc = self._doc_internal()
        mid_x = (float(x1) + float(x2)) / 2.0
        points = [[float(x1), float(y1)], [mid_x, float(y1)], [mid_x, float(y2)], [float(x2), float(y2)]]
        entity = doc.ModelSpace.AddLightWeightPolyline(_r8(_flat_2d(points)))
        self._ensure_layer_internal(doc, "PID-PROCESS-PIPING")
        entity.Layer = "PID-PROCESS-PIPING"
        return {"entity_type": "LWPOLYLINE", "handle": entity.Handle}

    async def pid_add_flow_arrow(self, x, y, rotation=0.0) -> CommandResult:
        return await self._exec(
            "pid_add_flow_arrow", lambda: self._job_pid_arrow(x, y, rotation)
        )

    def _job_pid_arrow(self, x, y, rotation) -> dict[str, Any]:
        doc = self._doc_internal()
        rad = math.radians(float(rotation))
        size = 2.0
        points = [
            [float(x) + size * math.cos(rad), float(y) + size * math.sin(rad)],
            [float(x) + size * 0.5 * math.cos(rad + 2.4), float(y) + size * 0.5 * math.sin(rad + 2.4)],
            [float(x) + size * 0.5 * math.cos(rad - 2.4), float(y) + size * 0.5 * math.sin(rad - 2.4)],
        ]
        entity = doc.ModelSpace.AddLightWeightPolyline(_r8(_flat_2d(points)))
        entity.Closed = True
        self._ensure_layer_internal(doc, "PID-ANNOTATION")
        entity.Layer = "PID-ANNOTATION"
        return {"entity_type": "LWPOLYLINE", "handle": entity.Handle}

    async def pid_add_equipment_tag(self, x, y, tag, description="") -> CommandResult:
        return await self._exec(
            "pid_add_equipment_tag",
            lambda: self._job_pid_equipment_tag(x, y, tag, description),
        )

    def _job_pid_equipment_tag(self, x, y, tag, description) -> dict[str, Any]:
        doc = self._doc_internal()
        msp = doc.ModelSpace
        self._ensure_layer_internal(doc, "PID-ANNOTATION")
        entity = msp.AddText(str(tag), _r8([float(x), float(y), 0.0]), 2.5)
        entity.Layer = "PID-ANNOTATION"
        payload = {"entity_type": "TEXT", "handle": entity.Handle, "tag": str(tag)}
        if description:
            extra = msp.AddText(str(description), _r8([float(x), float(y) - 3.5, 0.0]), 1.8)
            extra.Layer = "PID-ANNOTATION"
            payload["description_handle"] = extra.Handle
        return payload

    async def pid_add_line_number(self, x, y, line_num, spec) -> CommandResult:
        return await self._exec(
            "pid_add_line_number",
            lambda: self._job_pid_line_number(x, y, line_num, spec),
        )

    def _job_pid_line_number(self, x, y, line_num, spec) -> dict[str, Any]:
        doc = self._doc_internal()
        self._ensure_layer_internal(doc, "PID-ANNOTATION")
        entity = doc.ModelSpace.AddText(
            f"{line_num}-{spec}", _r8([float(x), float(y), 0.0]), 2.0
        )
        entity.Layer = "PID-ANNOTATION"
        return {"entity_type": "TEXT", "handle": entity.Handle}

    async def pid_insert_valve(self, x, y, valve_type, rotation=0.0, attributes=None) -> CommandResult:
        return await self._exec(
            "pid_insert_valve", lambda: self._job_pid_valve(x, y, valve_type)
        )

    def _job_pid_valve(self, x, y, valve_type) -> dict[str, Any]:
        doc = self._doc_internal()
        msp = doc.ModelSpace
        size = 3.0
        points = [
            [float(x) - size, float(y)],
            [float(x), float(y) + size],
            [float(x) + size, float(y)],
            [float(x), float(y) - size],
        ]
        entity = msp.AddLightWeightPolyline(_r8(_flat_2d(points)))
        entity.Closed = True
        self._ensure_layer_internal(doc, "PID-VALVES")
        entity.Layer = "PID-VALVES"
        self._ensure_layer_internal(doc, "PID-ANNOTATION")
        text = msp.AddText(str(valve_type), _r8([float(x), float(y) - size - 2.0, 0.0]), 1.5)
        text.Layer = "PID-ANNOTATION"
        return {"valve_type": str(valve_type), "handle": entity.Handle}

    async def pid_insert_instrument(self, x, y, instrument_type, rotation=0.0, tag_id="", range_value="") -> CommandResult:
        return await self._exec(
            "pid_insert_instrument",
            lambda: self._job_pid_instrument(x, y, instrument_type, tag_id),
        )

    def _job_pid_instrument(self, x, y, instrument_type, tag_id) -> dict[str, Any]:
        doc = self._doc_internal()
        msp = doc.ModelSpace
        self._ensure_layer_internal(doc, "PID-INSTRUMENTS")
        entity = msp.AddCircle(_r8([float(x), float(y), 0.0]), 4.0)
        entity.Layer = "PID-INSTRUMENTS"
        line = msp.AddLine(
            _r8([float(x) - 4.0, float(y), 0.0]), _r8([float(x) + 4.0, float(y), 0.0])
        )
        line.Layer = "PID-INSTRUMENTS"
        label = tag_id or instrument_type
        self._ensure_layer_internal(doc, "PID-ANNOTATION")
        text = msp.AddText(str(label), _r8([float(x), float(y) - 6.0, 0.0]), 1.5)
        text.Layer = "PID-ANNOTATION"
        return {"instrument_type": str(instrument_type), "handle": entity.Handle}

    async def pid_insert_pump(self, x, y, pump_type, rotation=0.0, attributes=None) -> CommandResult:
        return await self._exec(
            "pid_insert_pump", lambda: self._job_pid_pump(x, y, pump_type, rotation)
        )

    def _job_pid_pump(self, x, y, pump_type, rotation) -> dict[str, Any]:
        doc = self._doc_internal()
        msp = doc.ModelSpace
        self._ensure_layer_internal(doc, "PID-EQUIPMENT")
        entity = msp.AddCircle(_r8([float(x), float(y), 0.0]), 6.0)
        entity.Layer = "PID-EQUIPMENT"
        rad = math.radians(float(rotation))
        points = [
            [float(x) + 6.0 * math.cos(rad + 0.5), float(y) + 6.0 * math.sin(rad + 0.5)],
            [float(x) + 8.0 * math.cos(rad), float(y) + 8.0 * math.sin(rad)],
            [float(x) + 6.0 * math.cos(rad - 0.5), float(y) + 6.0 * math.sin(rad - 0.5)],
        ]
        triangle = msp.AddLightWeightPolyline(_r8(_flat_2d(points)))
        triangle.Closed = True
        triangle.Layer = "PID-EQUIPMENT"
        self._ensure_layer_internal(doc, "PID-ANNOTATION")
        text = msp.AddText(str(pump_type), _r8([float(x), float(y) - 8.0, 0.0]), 1.5)
        text.Layer = "PID-ANNOTATION"
        return {"pump_type": str(pump_type), "handle": entity.Handle}

    async def pid_insert_tank(self, x, y, tank_type, scale=1.0, attributes=None) -> CommandResult:
        return await self._exec(
            "pid_insert_tank", lambda: self._job_pid_tank(x, y, tank_type, scale)
        )

    def _job_pid_tank(self, x, y, tank_type, scale) -> dict[str, Any]:
        doc = self._doc_internal()
        msp = doc.ModelSpace
        w = 10.0 * float(scale)
        h = 15.0 * float(scale)
        points = [
            [float(x) - w, float(y)],
            [float(x) + w, float(y)],
            [float(x) + w, float(y) + h],
            [float(x) - w, float(y) + h],
        ]
        entity = msp.AddLightWeightPolyline(_r8(_flat_2d(points)))
        entity.Closed = True
        self._ensure_layer_internal(doc, "PID-EQUIPMENT")
        entity.Layer = "PID-EQUIPMENT"
        self._ensure_layer_internal(doc, "PID-ANNOTATION")
        text = msp.AddText(str(tank_type), _r8([float(x), float(y) + h + 2.0, 0.0]), 2.0 * float(scale))
        text.Layer = "PID-ANNOTATION"
        return {"tank_type": str(tank_type), "handle": entity.Handle}

    # -- view --------------------------------------------------------------

    async def zoom_extents(self) -> CommandResult:
        return await self._exec("zoom_extents", self._job_zoom_extents)

    def _job_zoom_extents(self) -> dict[str, Any]:
        doc = self._doc_internal()
        self._app.ZoomExtents()
        return {"zoomed": "extents", "document": _safe_str(doc, "Name")}

    async def zoom_window(self, x1, y1, x2, y2) -> CommandResult:
        return await self._exec(
            "zoom_window", lambda: self._job_zoom_window(x1, y1, x2, y2)
        )

    def _job_zoom_window(self, x1, y1, x2, y2) -> dict[str, Any]:
        self._app.ZoomWindow(
            _r8([float(x1), float(y1), 0.0]), _r8([float(x2), float(y2), 0.0])
        )
        return {"zoomed": "window", "window": [x1, y1, x2, y2]}

    async def get_screenshot(self) -> CommandResult:
        provider = self._screenshot
        if provider is None or isinstance(provider, NullScreenshotProvider):
            return CommandResult(ok=False, error="Screenshot capture is not available for this session")
        try:
            loop = asyncio.get_running_loop()
            data = await loop.run_in_executor(None, provider.capture)
        except Exception as ex:
            return CommandResult(ok=False, error=_friendly(ex, "get_screenshot"))
        if data:
            return CommandResult(ok=True, payload=data)
        return CommandResult(ok=False, error="Screenshot capture failed")
