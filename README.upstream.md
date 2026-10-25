# AutoCAD MCP Server

MCP server for AutoCAD automation and headless DXF generation.

Three backends, one API:

| Backend | Runtime | Requires AutoCAD? | Screenshot |
|---------|---------|-------------------|------------|
| **COM** | Windows Python | Yes — full AutoCAD (not LT) | Win32 PrintWindow |
| **File IPC** | Windows Python | Yes — AutoCAD LT 2024+ (Windows) | Win32 PrintWindow |
| **ezdxf** | Any platform | No (headless) | matplotlib render |

`auto` prefers **COM** when full AutoCAD's COM server is registered, then File IPC if an AutoCAD window is open, then ezdxf. The COM backend can launch AutoCAD by itself and needs no LISP dispatcher, so there is nothing to load manually.

The server exposes **8 consolidated tools** (`drawing`, `entity`, `layer`, `block`, `annotation`, `pid`, `view`, `system`) over the MCP stdio transport. Any MCP client (Claude Code, Codex, Cursor, Windsurf, Zed, …) can connect and drive AutoCAD through natural-language requests.

## Prerequisites

- **Windows 10/11** for the COM and File IPC backends (they use Win32 APIs)
- **Full AutoCAD** for the COM backend (AutoCAD LT has no ActiveX/COM automation). `AUTOCAD_MCP_COM_AUTOLAUNCH=1` (default) lets the server start AutoCAD when it is not running.
- **AutoCAD LT 2024 or newer** for the File IPC backend — AutoLISP support was added in LT 2024 for Windows. AutoCAD LT for Mac exists but does **not** support AutoLISP.
- **Python 3.10+** (Windows native — not WSL Python)

> The ezdxf headless backend works on any platform (Linux, macOS, WSL) for offline DXF generation without AutoCAD installed.

## Quick Start

### 1. Clone and install

```powershell
git clone https://github.com/puran-water/autocad-mcp.git
cd autocad-mcp
uv sync
```

Or install it as a standalone command for any MCP client:

```powershell
pipx install .            # or: uvx --from . autocad-mcp
autocad-mcp               # stdio server on stdin/stdout
```

### 2. Load the LISP dispatcher in AutoCAD LT (File IPC backend only)

**COM backend users can skip this step** — COM talks to AutoCAD directly and does not need `mcp_dispatch.lsp`.

Open AutoCAD LT and load `mcp_dispatch.lsp` using **APPLOAD**:

1. Type `APPLOAD` in the AutoCAD command line
2. Browse to `<repo>/lisp-code/mcp_dispatch.lsp`
3. Click **Load**
4. You should see: `=== MCP Dispatch v3.1 loaded ===` and `Ready for commands via (c:mcp-dispatch)`

> **Tip:** Add the file to your AutoCAD Startup Suite (in the APPLOAD dialog) so it loads automatically with every drawing.

### 3. Configure your MCP client

Using the installed console script (works for any MCP client):

```json
{
  "mcpServers": {
    "autocad": {
      "command": "autocad-mcp",
      "args": [],
      "env": { "AUTOCAD_MCP_BACKEND": "auto" }
    }
  }
}
```

Or point at the project venv directly:

```json
{
  "mcpServers": {
    "autocad": {
      "command": "C:\\path\\to\\autocad-mcp\\.venv\\Scripts\\python.exe",
      "args": ["-m", "autocad_mcp"],
      "env": { "AUTOCAD_MCP_BACKEND": "auto" }
    }
  }
}
```

Claude Code:

```bash
claude mcp add autocad -s user -- autocad-mcp
```

Codex CLI (`~/.codex/config.toml`; check your client's docs for the exact schema):

```toml
[mcp_servers.autocad]
command = "autocad-mcp"
args = []
env = { AUTOCAD_MCP_BACKEND = "auto" }
```

**Key points:**

- If the entry point is not on PATH, point `command` at the **Windows Python** inside the project venv (not WSL python), or use the `cmd.exe` form below.
- `AUTOCAD_MCP_BACKEND` can be `auto` (default — COM when available, then File IPC, then ezdxf), `com` (full AutoCAD via ActiveX; fails loudly when unavailable), `file_ipc` (requires AutoCAD + dispatcher), or `ezdxf` (headless only).

#### Running from WSL

If your MCP client runs in WSL (e.g. Claude Code), launch the server through `cmd.exe` so it runs as a native Windows process:

```json
{
  "mcpServers": {
    "autocad-mcp": {
      "type": "stdio",
      "command": "cmd.exe",
      "args": ["/d", "/s", "/c", "cd /d C:\\path\\to\\autocad-mcp && .venv\\Scripts\\python.exe -m autocad_mcp"],
      "env": { "AUTOCAD_MCP_BACKEND": "auto" }
    }
  }
}
```

### 4. Verify

From your MCP client, call:

```
system(operation="status")
```

You should see `backend: "file_ipc"` if AutoCAD is running, or `backend: "ezdxf"` for headless mode.

## Tools

### `drawing` — File/drawing management

| Operation | Description | File IPC | ezdxf |
|-----------|-------------|----------|-------|
| `create` | Reset to clean drawing (erase all + purge) | Yes | Yes |
| `open` | Open an existing drawing | Yes | Yes (DXF) |
| `info` | Get entity count and layers | Yes | Yes |
| `save` | Save current drawing (to path if given) | Yes | Yes |
| `save_as_dxf` | Export as DXF | Yes | Yes |
| `plot_pdf` | Plot to PDF | Yes | No |
| `purge` | Purge unused objects | Yes | Yes |
| `get_variables` | Get system variables by name | Yes | Yes |
| `undo` | Undo last operation | Yes | No |
| `redo` | Redo last undone operation | Yes | No |

### `entity` — Entity CRUD + modification

**Create:** `create_line`, `create_circle`, `create_polyline`, `create_rectangle`, `create_arc`, `create_ellipse`, `create_mtext`, `create_hatch`

**Read:** `list`, `count`, `get`

**Modify:** `copy`, `move`, `rotate`, `scale`, `mirror`, `offset`\*, `array`, `fillet`\*, `chamfer`\*, `erase`

> \* `offset`, `fillet`, `chamfer` are File IPC only (not supported in ezdxf headless backend).

### `layer` — Layer management

`list`, `create`, `set_current`, `set_properties`, `freeze`, `thaw`, `lock`, `unlock`

### `block` — Block operations

| Operation | File IPC | ezdxf |
|-----------|----------|-------|
| `list` | Yes | Yes |
| `insert` | Yes | Yes |
| `insert_with_attributes` | Yes | Yes |
| `get_attributes` | Yes | Yes |
| `update_attribute` | Yes | Yes |
| `define` | No | Yes |

### `annotation` — Text, dimensions, leaders

`create_text`, `create_dimension_linear`, `create_dimension_aligned`, `create_dimension_angular`, `create_dimension_radius`, `create_leader`

### `pid` — P&ID operations (CTO symbol library)

`setup_layers`, `insert_symbol`, `list_symbols`, `draw_process_line`, `connect_equipment`, `add_flow_arrow`, `add_equipment_tag`, `add_line_number`, `insert_valve`, `insert_instrument`, `insert_pump`, `insert_tank`

> P&ID symbol insertion requires the [CAD Tools Online](https://www.cadtoolsonline.com/) (CTO) P&ID Symbol Library installed at `C:\PIDv4-CTO\`. The ezdxf backend has built-in CTO library support. For the File IPC backend, some P&ID operations require additional LISP helpers — see the P&ID section in the wiki for setup details.

### `view` — Viewport and screenshot

| Operation | Description |
|-----------|-------------|
| `zoom_extents` | Zoom to show all entities |
| `zoom_window` | Zoom to a specified window |
| `get_screenshot` | Capture current AutoCAD view as PNG |

Screenshots use `PrintWindow` (Win32) for the File IPC backend — works even when AutoCAD is minimized or in the background. The ezdxf backend renders via matplotlib.

### `system` — Server management

`status`, `health`, `get_backend`, `runtime`, `init`, `execute_lisp`

> `execute_lisp` runs arbitrary AutoLISP code (COM and File IPC backends). Pass `data: {code: "(+ 1 2)"}`. On COM this is queued through `Document.SendCommand`; on File IPC it runs via the dispatcher. This turns the server into an extensible automation platform — any valid AutoLISP expression can be executed. **Security note:** it can run arbitrary code inside your AutoCAD session; only connect trusted MCP clients.

## Architecture

```
MCP Client (Claude Code, Codex, Cursor, …)
    │  stdio (JSON-RPC)
    ▼
Python MCP Server (autocad_mcp)
    │
    ├── COM Backend ──► AutoCAD.Application (ActiveX) ──► live drawing
    │   Single STA worker thread, serialized calls, auto-launch
    │
    ├── File IPC Backend ──► C:/temp/*.json ──► mcp_dispatch.lsp (AutoCAD LT)
    │   PostMessageW(WM_CHAR) to MDIClient — no focus steal
    │
    └── ezdxf Backend ──► in-memory DXF (headless, no AutoCAD needed)
```

The COM backend talks to full AutoCAD through its out-of-process ActiveX server (`acad.exe /Automation`). All COM calls run on one dedicated STA thread and are serialized; the asyncio loop only awaits results, so slow AutoCAD work never blocks the MCP server. `GetActiveObject` attaches to a running instance, otherwise `Dispatch` starts one.

The File IPC backend sends keystrokes to AutoCAD's MDIClient window via `PostMessageW(WM_CHAR)`, triggering the `(c:mcp-dispatch)` AutoLISP command. This approach does **not** steal window focus — you can continue working in other applications while automation runs.

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `AUTOCAD_MCP_BACKEND` | `auto` | Backend selection: `auto`, `com`, `file_ipc`, `ezdxf` |
| `AUTOCAD_MCP_COM_PROGID` | `AutoCAD.Application` | COM ProgID to attach to / launch |
| `AUTOCAD_MCP_COM_AUTOLAUNCH` | `1` | Let the server start AutoCAD when it is not running (`0` disables) |
| `AUTOCAD_MCP_COM_CONNECT_TIMEOUT` | `90` | Seconds to wait for AutoCAD to become ready (10–300) |
| `AUTOCAD_MCP_COM_CALL_TIMEOUT` | `60` | Timeout per COM call in seconds (1–600) |
| `AUTOCAD_MCP_COM_RETRIES` | `3` | Retries when AutoCAD is busy (`RPC_E_CALL_REJECTED` / `RETRYLATER`) |
| `AUTOCAD_MCP_IPC_DIR` | `C:/temp` | Directory for IPC command/result JSON files (must match on both Python and LISP sides) |
| `AUTOCAD_MCP_IPC_TIMEOUT` | `10.0` | IPC command timeout in seconds (1-300) |
| `AUTOCAD_MCP_ONLY_TEXT` | `false` | Disable screenshot capture (text feedback only) |

> **Note:** If you change `AUTOCAD_MCP_IPC_DIR`, you must also update the `*mcp-ipc-dir*` variable in `mcp_dispatch.lsp` to match.

## Development

```powershell
uv sync
uv run pytest tests/ -v
```

## AutoCAD LT AutoLISP Compatibility

AutoLISP was added to AutoCAD LT in the **2024 release (Windows only)**. AutoCAD LT for Mac does not support AutoLISP.

| Supported (LT 2024+ Windows) | Not Supported |
|-------------------------------|---------------|
| `.lsp` / `.fas` / `.vlx` / `.dcl` | VLIDE (Visual LISP IDE) |
| All `vl-*` utility functions | `vlax-*` (ActiveX/COM) |
| File I/O (`open`, `read-line`, etc.) | Express Tools |
| Entity access (`entget`, `entmod`, etc.) | 3D operations |
| Selection sets | AutoLISP on Mac |

The `mcp_dispatch.lsp` dispatcher is fully compatible with LT 2024+.

## What's New in v3.1

- **COM backend** — drives full AutoCAD through ActiveX (`AutoCAD.Application`). Starts AutoCAD automatically, runs every call on one dedicated STA thread, retries when AutoCAD is busy, and re-attaches once if AutoCAD is closed. No LISP dispatcher is needed.
- **Diameter dimensions** — new `annotation` operation `create_dimension_diameter` (`data: {cx, cy, radius, angle?}`), implemented with `AddDimDiametric`.
- **Explicit backend selection** — unknown `AUTOCAD_MCP_BACKEND` values now raise instead of silently falling back to ezdxf; `system init` disposes the old backend and re-runs detection, so you can reconnect without restarting the server.
- **Packaging** — `autocad-mcp` console entry point, version single-sourced from `src/autocad_mcp/__init__.py`.
- **`execute_lisp`** — Run arbitrary AutoLISP code via temp file pattern. Turns the server from a fixed command set into an extensible automation platform.
- **Undo / Redo** — Single-step undo and redo via `drawing` tool.
- **Drawing open** — Open existing `.dwg` files programmatically (FILEDIA suppressed).
- **Drawing create** — Now resets current drawing (erase all + purge) instead of `_.NEW`, preserving the LISP dispatcher namespace.
- **Drawing save with path** — `save` with a `path` parameter uses SAVEAS; without path uses QSAVE.
- **`get_variables` fix** — Respects the `names` parameter; returns requested variables with proper type handling.
- **Polyline/leader fix** — Point arrays properly encoded via semicolon-delimited format.
- **ESC prefix** — Sends 2x ESC before each dispatch to cancel stale pending commands from prior timeouts.
- **UTF-8/cp1252 fallback** — Handles non-ASCII characters in LISP result files (AutoCAD writes Windows-1252).
- **Configurable IPC timeout** — `AUTOCAD_MCP_IPC_TIMEOUT` env var (1–300 seconds, default 10).
- **Thread-safe backend init** — `asyncio.Lock` prevents parallel initialization races.

## License

MIT
