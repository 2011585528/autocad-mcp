"""Backend detection and environment configuration."""

from __future__ import annotations

import os
import sys
from pathlib import Path

import structlog

log = structlog.get_logger()

# Paths
LISP_DIR = Path(__file__).resolve().parent.parent.parent / "lisp-code"
IPC_DIR = Path(os.environ.get("AUTOCAD_MCP_IPC_DIR", "C:/temp"))

# Backend selection
BACKEND_DEFAULT = "auto"  # auto | file_ipc | ezdxf

# IPC timeout (seconds), clamped to [1, 300]
IPC_TIMEOUT = max(1.0, min(300.0, float(os.environ.get("AUTOCAD_MCP_IPC_TIMEOUT", "10.0"))))

# Screenshot
ONLY_TEXT_FEEDBACK = os.environ.get("AUTOCAD_MCP_ONLY_TEXT", "").lower() in ("1", "true", "yes")

# Win32 availability
WIN32_AVAILABLE = sys.platform == "win32"


def _env_float(name: str, default: float) -> float:
    try:
        return float(os.environ.get(name, default))
    except (TypeError, ValueError):
        return default


def _env_int(name: str, default: int) -> int:
    try:
        return int(os.environ.get(name, default))
    except (TypeError, ValueError):
        return default


def _env_flag(name: str, default: bool) -> bool:
    raw = os.environ.get(name)
    if raw is None:
        return default
    return raw.strip().lower() in ("1", "true", "yes", "on")


# COM backend (full AutoCAD on Windows). COM is unavailable in AutoCAD LT.
COM_PROGID = os.environ.get("AUTOCAD_MCP_COM_PROGID", "AutoCAD.Application").strip() or "AutoCAD.Application"
COM_AUTOLAUNCH = _env_flag("AUTOCAD_MCP_COM_AUTOLAUNCH", True)
COM_CONNECT_TIMEOUT = max(10.0, min(300.0, _env_float("AUTOCAD_MCP_COM_CONNECT_TIMEOUT", 90.0)))
COM_CALL_TIMEOUT = max(1.0, min(600.0, _env_float("AUTOCAD_MCP_COM_CALL_TIMEOUT", 60.0)))
COM_RETRIES = max(0, min(10, _env_int("AUTOCAD_MCP_COM_RETRIES", 3)))


def _com_available() -> bool:
    """True when the AutoCAD COM server is registered and pywin32 imports.

    Registry-only probe: never starts AutoCAD.
    """
    if not WIN32_AVAILABLE:
        return False
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, COM_PROGID + r"\CurVer") as key:
            winreg.QueryValueEx(key, "")
    except Exception:
        return False
    try:
        import win32com.client  # noqa: F401
    except Exception:
        return False
    return True


def _current_backend_env() -> str:
    """Read backend selection from env with normalization."""
    return os.environ.get("AUTOCAD_MCP_BACKEND", BACKEND_DEFAULT).strip().lower()


def _is_wsl() -> bool:
    """Detect WSL Linux runtime."""
    if os.environ.get("WSL_INTEROP"):
        return True
    try:
        return "microsoft" in os.uname().release.lower()
    except AttributeError:
        return False


def _write_debug_snapshot(backend_env: str):
    """Optionally write backend detection debug information.

    Set AUTOCAD_MCP_DEBUG_DETECT_FILE to enable.
    """
    debug_file = os.environ.get("AUTOCAD_MCP_DEBUG_DETECT_FILE", "").strip()
    if not debug_file:
        return

    try:
        debug_path = Path(debug_file)
        debug_path.parent.mkdir(parents=True, exist_ok=True)
        with debug_path.open("w", encoding="utf-8") as f:
            f.write(f"sys.platform={sys.platform}\n")
            f.write(f"WIN32_AVAILABLE={WIN32_AVAILABLE}\n")
            f.write(f"BACKEND_ENV={backend_env}\n")
            f.write(f"python={sys.executable}\n")
    except Exception:
        # Best-effort only; never fail backend detection due debug writes.
        pass


def detect_backend() -> str:
    """Return the backend name to use: 'com', 'file_ipc' or 'ezdxf'.

    Raises RuntimeError with an actionable message when an explicitly
    requested backend is unavailable. Only 'auto' may fall back.
    """
    backend_env = _current_backend_env()
    _write_debug_snapshot(backend_env)

    if backend_env == "ezdxf":
        log.info("backend_selected", backend="ezdxf", reason="explicit")
        return "ezdxf"

    if backend_env == "com":
        if not WIN32_AVAILABLE:
            raise RuntimeError("AUTOCAD_MCP_BACKEND=com requires Windows.")
        if not _com_available():
            raise RuntimeError(
                f"AUTOCAD_MCP_BACKEND=com but COM ProgID '{COM_PROGID}' is not registered "
                "(requires full AutoCAD, not AutoCAD LT) or pywin32 is missing."
            )
        log.info("backend_selected", backend="com", reason="explicit")
        return "com"

    if backend_env == "file_ipc":
        if not WIN32_AVAILABLE:
            raise RuntimeError(
                "AUTOCAD_MCP_BACKEND=file_ipc requires Windows. "
                "Use AUTOCAD_MCP_BACKEND=ezdxf for headless mode."
            )
        try:
            from autocad_mcp.backends.file_ipc import find_autocad_window
        except ImportError as ex:
            raise RuntimeError(
                "AUTOCAD_MCP_BACKEND=file_ipc requires pywin32. "
                "Install with: pip install pywin32"
            ) from ex
        if not find_autocad_window():
            raise RuntimeError(
                "AUTOCAD_MCP_BACKEND=file_ipc but no AutoCAD window found. "
                "Start AutoCAD, open a .dwg file and load mcp_dispatch.lsp."
            )
        log.info("backend_selected", backend="file_ipc", reason="explicit")
        return "file_ipc"

    if backend_env != "auto":
        raise RuntimeError(
            f"Unknown AUTOCAD_MCP_BACKEND={backend_env!r}; expected auto, com, file_ipc or ezdxf."
        )

    # auto: COM first (works even with no drawing open, and can launch AutoCAD),
    # then the file-IPC window probe, then the offline fallback.
    if _com_available():
        log.info("backend_selected", backend="com", reason="auto: COM server registered")
        return "com"

    if WIN32_AVAILABLE:
        try:
            from autocad_mcp.backends.file_ipc import find_autocad_window

            hwnd = find_autocad_window()
            if hwnd:
                log.info("backend_selected", backend="file_ipc", reason="auto: window found", hwnd=hwnd)
                return "file_ipc"
        except ImportError:
            log.info("win32_deps_missing_fallback_ezdxf")
        except Exception as ex:
            log.warning("file_ipc_probe_failed", error=str(ex))
    elif _is_wsl():
        log.info(
            "wsl_linux_python_fallback_ezdxf",
            platform=sys.platform,
            python=sys.executable,
            hint="Launch MCP with Windows python.exe for COM / File IPC backends.",
        )

    log.info("backend_selected", backend="ezdxf", reason="auto: no AutoCAD available")
    return "ezdxf"
