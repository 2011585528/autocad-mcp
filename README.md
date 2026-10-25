# AutoCAD MCP Server — COM / LT / Headless

**English** | [简体中文](README.zh-CN.md)

An [MCP](https://modelcontextprotocol.io) server that lets an AI agent draw and edit **real AutoCAD drawings**. This fork adds a native **COM (ActiveX) backend** for full AutoCAD, alongside the upstream File IPC (AutoCAD LT) and ezdxf (headless) backends.

> Forked from [puran-water/autocad-mcp](https://github.com/puran-water/autocad-mcp) (MIT).
> The original README is kept as [`README.upstream.md`](README.upstream.md).

## 简体中文简介

这是 [puran-water/autocad-mcp](https://github.com/puran-water/autocad-mcp) 的 fork，主要新增：

- **COM 后端**：直接驱动完整版 AutoCAD（LT 不支持 COM），可自动启动 CAD，无需加载任何 LISP。
- **直径标注**：新增 `create_dimension_diameter`。
- **竖直线性标注修复**：AutoCAD 2027 的 ActiveX 已没有 `AddDimLinear`，改用 `AddDimRotated` 并保留 `DIMLINEAR` 命令兜底。
- **无需重启的重连**：`system init` 会释放旧后端并重新探测。
- **可打包分发**：提供 `autocad-mcp` 命令行入口，wheel / sdist 均可安装。
- **测试**：145 个用例全部通过，其中 22 个是新增的 COM 后端测试（使用假 COM 对象，不需要真实 CAD）。

---

## Backends

| Backend | AutoCAD required | Works with | Needs LISP | Typical use |
|---|---|---|---|---|
| **`com`** | Yes — launches it automatically | Full AutoCAD (not LT) | No | Live editing in a real drawing |
| **`file_ipc`** | Yes — must already be running | AutoCAD LT 2024+ (Windows) | Yes | AutoCAD LT users |
| **`ezdxf`** | No | Any platform | No | Offline DXF generation |

With `AUTOCAD_MCP_BACKEND=auto` (the default), the server picks **COM → File IPC → ezdxf**, in that order.

> **Important:** AutoCAD LT has no ActiveX/COM server, so it cannot use the `com` backend. On LT, select `file_ipc` explicitly and load `lisp-code/mcp_dispatch.lsp`.

## Requirements

- Windows 10/11 for the `com` and `file_ipc` backends
- Python 3.10+
- **Full AutoCAD** for the `com` backend; AutoCAD LT 2024+ for `file_ipc`
- The `ezdxf` backend works anywhere (Linux, macOS, WSL)
- Network access for the first dependency install

Tested against **AutoCAD 2027 (R26, `26.0s`)** on Windows 11. Other full AutoCAD versions exposing the `AutoCAD.Application` ProgID should work but are untested.

## Install

### From a wheel

```powershell
pipx install https://github.com/2011585528/autocad-mcp/releases/download/v3.1.0/autocad_mcp-3.1.0-py3-none-any.whl
```

### From source

```powershell
git clone https://github.com/2011585528/autocad-mcp.git
cd autocad-mcp
uv sync
```

or with venv + pip:

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -e ".[dev]"
```

The package installs a console script:

```
autocad-mcp
```

## Configure your MCP client

Using the console script (works with any MCP client):

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

Pointing at a project venv instead:

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

Restart the client after changing its configuration. On the first tool call, the COM backend attaches to a running AutoCAD instance or starts one.

## Quick start

Ask the agent something like:

> Draw a 100 × 60 rectangle at (0,0), a Ø20 hole at its center, and four Ø8 holes 10 mm from the adjacent edges. Add length and width dimensions, dimension the center hole diameter, then zoom to extents.

The server exposes 8 consolidated tools with operation-based dispatch:

| Tool | Examples |
|---|---|
| `drawing` | `create`, `open`, `info`, `save`, `save_as_dxf`, `plot_pdf`, `undo`, `redo` |
| `entity` | `create_line`, `create_circle`, `create_rectangle`, `create_polyline`, `move`, `copy`, `rotate`, `mirror`, `offset`, `array`, `list`, `get`, `erase` |
| `layer` | `list`, `create`, `set_current`, `set_properties`, `freeze`, `thaw`, `lock`, `unlock` |
| `block` | `list`, `insert`, `insert_with_attributes`, `get_attributes`, `update_attribute`, `define` |
| `annotation` | `create_text`, `create_dimension_linear`, `create_dimension_aligned`, `create_dimension_angular`, `create_dimension_radius`, **`create_dimension_diameter`**, `create_leader` |
| `pid` | `setup_layers`, `insert_symbol`, `draw_process_line`, `connect_equipment`, `insert_valve`, `insert_pump`, `insert_tank` |
| `view` | `zoom_extents`, `zoom_window`, `get_screenshot` |
| `system` | `status`, `health`, `get_backend`, `runtime`, `init`, `execute_lisp` |

## What this fork adds

### COM backend for full AutoCAD

- Talks to `AutoCAD.Application` through `win32com`, so no LISP dispatcher is needed.
- **Auto-launch**: if AutoCAD is not running, the server starts it and polls until it is ready.
- **Single STA worker thread**: every COM call runs on one dedicated thread; the asyncio loop only awaits results, so slow AutoCAD work never blocks the server.
- **Serialized calls** with exponential backoff on `RPC_E_CALL_REJECTED` / `RPC_E_SERVERCALL_RETRYLATER`.
- **Reconnect**: if AutoCAD is closed mid-session, the backend re-attaches once; `system init` forces a full reset.
- **Entity IDs** are AutoCAD `Handle` strings, resolved with `Document.HandleToObject`.
- **`FILEDIA` suppression** while connected, restored on close, so `SAVEAS` / `OPEN` do not block on dialogs.
- New drawings go to a real new document (`Documents.Add()`), not an erase-all of the current one.

### Diameter dimensions

`annotation create_dimension_diameter` with `data: {cx, cy, radius, angle?}`, implemented with `AddDimDiametric`.

### Linear dimension fix for AutoCAD 2027

AutoCAD R26 no longer exposes `AddDimLinear` over ActiveX (`DISP_E_UNKNOWNNAME`). This fork uses `AddDimRotated` with the orientation derived from the dimension-line location, and falls back to the native `DIMLINEAR` command if the COM call fails.

### Explicit backend selection

- Unknown `AUTOCAD_MCP_BACKEND` values now raise instead of silently using ezdxf.
- `system init` disposes the previous backend (including its worker thread) before re-running detection.

### Packaging

- `autocad-mcp` console entry point.
- Version single-sourced from `src/autocad_mcp/__init__.py`.
- `dev` dependencies moved to `[project.optional-dependencies]`.
- `python -m build` produces a clean sdist (backups and caches excluded).

### Tests

145 tests pass, including **22 COM backend tests** that use fake COM objects — no AutoCAD and no `win32com` required:

```bash
python -m pytest tests -q
# 145 passed
```

## How the COM backend works

```
MCP client
   │ stdio (JSON-RPC)
   ▼
autocad_mcp server (asyncio)
   │
   └── ComBackend
         │  submit(job)
         ▼
   dedicated STA thread  ──►  AutoCAD.Application (COM)
   CoInitialize / queue        GetActiveObject → Dispatch
```

- All COM calls are queued to one thread, which keeps STA affinity and serializes access.
- A timed-out caller does not kill the worker; late results are discarded.
- Disconnect HRESULTs trigger one re-attach; busy HRESULTs trigger bounded retries.

## Environment variables

| Variable | Default | Description |
|---|---|---|
| `AUTOCAD_MCP_BACKEND` | `auto` | `auto` / `com` / `file_ipc` / `ezdxf` |
| `AUTOCAD_MCP_COM_PROGID` | `AutoCAD.Application` | COM ProgID to attach to / launch |
| `AUTOCAD_MCP_COM_AUTOLAUNCH` | `1` | Start AutoCAD when it is not running (`0` disables) |
| `AUTOCAD_MCP_COM_CONNECT_TIMEOUT` | `90` | Seconds to wait for AutoCAD to become ready (10–300) |
| `AUTOCAD_MCP_COM_CALL_TIMEOUT` | `60` | Timeout per COM call in seconds (1–600) |
| `AUTOCAD_MCP_COM_RETRIES` | `3` | Retries when AutoCAD is busy |
| `AUTOCAD_MCP_IPC_DIR` | `C:/temp` | File IPC directory (must match the LISP dispatcher) |
| `AUTOCAD_MCP_IPC_TIMEOUT` | `10.0` | File IPC timeout in seconds (1–300) |
| `AUTOCAD_MCP_ONLY_TEXT` | `false` | Disable screenshots, text feedback only |

## Verify the connection

Call `system status`. With AutoCAD attached you should see:

```json
{
  "backend": "com",
  "connected": true,
  "autocad_version": "26.0s (LMS Tech)",
  "hwnd": 459884,
  "active_document": { "name": "Drawing1.dwg", "saved": false },
  "capabilities": { "can_zoom": true, "can_undo": true, "can_plot_pdf": true }
}
```

If `backend` is `ezdxf`, the server is running headless and is **not** connected to AutoCAD.

## AutoCAD LT compatibility

AutoCAD LT has no COM/ActiveX automation. On LT:

```toml
AUTOCAD_MCP_BACKEND = "file_ipc"
```

Then load the dispatcher in AutoCAD LT:

```
(load "C:/path/to/autocad-mcp/lisp-code/mcp_dispatch.lsp")
```

The dispatcher relies on AutoLISP, which is available in AutoCAD LT 2024+ on Windows (not on Mac).

## Security

`system execute_lisp` and the COM backend's `SendCommand` path can execute arbitrary AutoLISP inside your AutoCAD session. Only connect MCP clients you trust, and review agent actions before approving destructive operations.

## Attribution & License

This project is a fork of [puran-water/autocad-mcp](https://github.com/puran-water/autocad-mcp).

- Upstream: MIT License, Copyright (c) 2024 AutoCAD MCP Server Contributors.
- Modifications in this fork: © 2026 Bai Xiangyuan.

The original [`LICENSE`](LICENSE) is retained. Upstream documentation is preserved as [`README.upstream.md`](README.upstream.md).
