# AutoCAD MCP 服务器 —— COM / LT / 离线

一个基于 [MCP](https://modelcontextprotocol.io) 的服务器，让 AI agent 能够绘制和编辑**真实的 AutoCAD 图纸**。本 fork 在原有 File IPC（AutoCAD LT）和 ezdxf（离线）后端的基础上，新增了原生 **COM（ActiveX）后端**，可直接驱动完整版 AutoCAD。

> 本项目 fork 自 [puran-water/autocad-mcp](https://github.com/puran-water/autocad-mcp)（MIT 许可）。
> 原版 README 保留在 [`README.upstream.md`](README.upstream.md)，英文版说明见 [`README.md`](README.md)。

---

## 三种后端

| 后端 | 是否需要 AutoCAD | 适用对象 | 是否需要 LISP | 典型用途 |
|---|---|---|---|---|
| **`com`** | 是（可自动启动） | 完整版 AutoCAD（非 LT） | 否 | 在真实图纸中实时绘图 |
| **`file_ipc`** | 是（必须已运行） | AutoCAD LT 2024+（Windows） | 是 | LT 用户 |
| **`ezdxf`** | 否 | 任意平台 | 否 | 离线生成 DXF |

默认 `AUTOCAD_MCP_BACKEND=auto`，选择顺序为 **COM → File IPC → ezdxf**。

> **重要：** AutoCAD LT 不提供 ActiveX/COM 服务，因此**无法使用 `com` 后端**。LT 用户请显式设置 `file_ipc`，并加载 `lisp-code/mcp_dispatch.lsp`。

## 环境要求

- `com` 和 `file_ipc` 后端需要 Windows 10/11
- Python 3.10 及以上
- `com` 后端需要**完整版 AutoCAD**；`file_ipc` 需要 AutoCAD LT 2024+
- `ezdxf` 后端可在任意平台使用（Linux、macOS、WSL）
- 首次安装依赖需要联网

已在 **AutoCAD 2027（R26，`26.0s`）+ Windows 11** 上实测通过。其他注册了 `AutoCAD.Application` ProgID 的完整版 AutoCAD 理论上也可用，但未实测。

## 安装

### 使用 wheel

```powershell
pipx install autocad_mcp-3.1.0-py3-none-any.whl
```

### 从源码安装

```powershell
git clone https://github.com/2011585528/autocad-mcp.git
cd autocad-mcp
uv sync
```

或者使用 venv + pip：

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -e ".[dev]"
```

安装后会生成命令行入口：

```
autocad-mcp
```

## 配置 MCP 客户端

使用命令行入口（适用于任何 MCP 客户端）：

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

也可以直接指向项目的虚拟环境：

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

Claude Code：

```bash
claude mcp add autocad -s user -- autocad-mcp
```

Codex CLI（`~/.codex/config.toml`，字段以客户端文档为准）：

```toml
[mcp_servers.autocad]
command = "autocad-mcp"
args = []
env = { AUTOCAD_MCP_BACKEND = "auto" }
```

修改配置后重启客户端。首次调用工具时，COM 后端会附加到正在运行的 AutoCAD，或自动启动一个实例。

## 快速开始

可以直接对 agent 说：

> 在 (0,0) 画一个 100 × 60 的矩形，中心画一个 Ø20 的孔，四角各画一个 Ø8 的孔，孔心距相邻两边均为 10。给长度和宽度加尺寸标注，标注中心孔直径，最后缩放到图形范围。

服务器提供 8 个按操作分发的工具：

| 工具 | 常用操作 |
|---|---|
| `drawing` | `create`、`open`、`info`、`save`、`save_as_dxf`、`plot_pdf`、`undo`、`redo` |
| `entity` | `create_line`、`create_circle`、`create_rectangle`、`create_polyline`、`move`、`copy`、`rotate`、`mirror`、`offset`、`array`、`list`、`get`、`erase` |
| `layer` | `list`、`create`、`set_current`、`set_properties`、`freeze`、`thaw`、`lock`、`unlock` |
| `block` | `list`、`insert`、`insert_with_attributes`、`get_attributes`、`update_attribute`、`define` |
| `annotation` | `create_text`、`create_dimension_linear`、`create_dimension_aligned`、`create_dimension_angular`、`create_dimension_radius`、**`create_dimension_diameter`**、`create_leader` |
| `pid` | `setup_layers`、`insert_symbol`、`draw_process_line`、`connect_equipment`、`insert_valve`、`insert_pump`、`insert_tank` |
| `view` | `zoom_extents`、`zoom_window`、`get_screenshot` |
| `system` | `status`、`health`、`get_backend`、`runtime`、`init`、`execute_lisp` |

## 本 fork 新增的内容

### 完整版 AutoCAD 的 COM 后端

- 通过 `win32com` 连接 `AutoCAD.Application`，**不需要加载任何 LISP 调度器**。
- **自动启动**：AutoCAD 未运行时自动拉起，并轮询等待其就绪。
- **专用 STA 工作线程**：所有 COM 调用都在同一个线程上执行，asyncio 事件循环只等待结果，因此 AutoCAD 的慢操作不会阻塞服务器。
- **串行调用**：遇到 `RPC_E_CALL_REJECTED` / `RPC_E_SERVERCALL_RETRYLATER` 时按指数退避重试。
- **断线重连**：AutoCAD 中途关闭时会尝试重新附加一次；`system init` 可强制完整重置。
- **实体 ID**：对外统一使用 AutoCAD 的 `Handle` 字符串，内部通过 `Document.HandleToObject` 解析。
- **抑制 FILEDIA**：连接期间关闭文件对话框，避免 `SAVEAS` / `OPEN` 阻塞；关闭时恢复原值。
- **新建图纸**：使用 `Documents.Add()` 真正创建新文档，而不是清空当前图纸。

### 直径标注

新增 `annotation create_dimension_diameter`，参数为 `data: {cx, cy, radius, angle?}`，底层调用 `AddDimDiametric`。

### AutoCAD 2027 线性标注修复

AutoCAD R26 的 ActiveX 已不再提供 `AddDimLinear`（返回 `DISP_E_UNKNOWNNAME`）。本 fork 改用 `AddDimRotated`，根据标注线位置自动判断水平/垂直方向；如果 COM 调用失败，则回退到原生 `DIMLINEAR` 命令。

### 显式后端选择

- 未知的 `AUTOCAD_MCP_BACKEND` 值现在会直接报错，不再静默使用 ezdxf。
- `system init` 会先释放旧后端（包括其工作线程），再重新探测后端。

### 打包

- 提供 `autocad-mcp` 命令行入口。
- 版本号单一来源：`src/autocad_mcp/__init__.py`。
- `dev` 依赖已移入 `[project.optional-dependencies]`。
- `python -m build` 产出的 sdist 已排除备份目录和缓存。

### 测试

145 个测试全部通过，其中 **22 个是新增的 COM 后端测试**（使用假 COM 对象，不需要真实 AutoCAD，也不需要 `win32com`）：

```bash
python -m pytest tests -q
# 145 passed
```

## COM 后端工作原理

```
MCP 客户端
   │ stdio（JSON-RPC）
   ▼
autocad_mcp 服务器（asyncio）
   │
   └── ComBackend
         │  submit(job)
         ▼
   专用 STA 线程  ──►  AutoCAD.Application（COM）
   CoInitialize / 队列   GetActiveObject → Dispatch
```

- 所有 COM 调用都排入同一个线程，保证 STA 亲和性并实现串行访问。
- 调用超时不会杀死工作线程；迟到的结果会被丢弃。
- 断线类 HRESULT 触发一次重新附加；忙碌类 HRESULT 触发有限次重试。

## 环境变量

| 变量 | 默认值 | 说明 |
|---|---|---|
| `AUTOCAD_MCP_BACKEND` | `auto` | `auto` / `com` / `file_ipc` / `ezdxf` |
| `AUTOCAD_MCP_COM_PROGID` | `AutoCAD.Application` | 要连接或启动的 COM ProgID |
| `AUTOCAD_MCP_COM_AUTOLAUNCH` | `1` | AutoCAD 未运行时是否自动启动（`0` 关闭） |
| `AUTOCAD_MCP_COM_CONNECT_TIMEOUT` | `90` | 等待 AutoCAD 就绪的秒数（10–300） |
| `AUTOCAD_MCP_COM_CALL_TIMEOUT` | `60` | 单次 COM 调用超时秒数（1–600） |
| `AUTOCAD_MCP_COM_RETRIES` | `3` | AutoCAD 忙碌时的重试次数 |
| `AUTOCAD_MCP_IPC_DIR` | `C:/temp` | File IPC 目录（必须与 LISP 调度器一致） |
| `AUTOCAD_MCP_IPC_TIMEOUT` | `10.0` | File IPC 超时秒数（1–300） |
| `AUTOCAD_MCP_ONLY_TEXT` | `false` | 关闭截图，仅返回文本 |

## 验证连接

调用 `system status`。连接成功时应看到：

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

如果 `backend` 显示为 `ezdxf`，说明服务器运行在离线模式，**并没有连接 AutoCAD**。

## AutoCAD LT 兼容性

AutoCAD LT 不支持 COM/ActiveX 自动化。在 LT 上请使用：

```toml
AUTOCAD_MCP_BACKEND = "file_ipc"
```

然后在 AutoCAD LT 命令行加载调度器：

```
(load "C:/path/to/autocad-mcp/lisp-code/mcp_dispatch.lsp")
```

该调度器依赖 AutoLISP，只有 AutoCAD LT 2024+（Windows）支持；Mac 版 LT 不支持 AutoLISP。

## 安全提示

`system execute_lisp` 以及 COM 后端的 `SendCommand` 通道可以在你的 AutoCAD 会话中执行任意 AutoLISP 代码。请只连接你信任的 MCP 客户端，并在执行删除等破坏性操作前确认 agent 的行为。

## 归属与许可

本项目 fork 自 [puran-water/autocad-mcp](https://github.com/puran-water/autocad-mcp)。

- 上游项目：MIT 许可，Copyright (c) 2024 AutoCAD MCP Server Contributors。
- 本 fork 的修改：© 2026 Bai Xiangyuan。

原始 [`LICENSE`](LICENSE) 完整保留；上游文档保留在 [`README.upstream.md`](README.upstream.md)。
