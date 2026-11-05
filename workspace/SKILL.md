---
name: autoofficemcp
description: 当任务涉及远端 Windows Office 主机上的 Excel / Word / PowerPoint，且当前工具列表缺少所需能力、工具名称或参数未知、或需要调用未直接暴露的工具时使用。说明如何通过 HTTP 文档路由、MCP tools/list、office_batch / office_batch_files 发现并调用 AutoOfficeMCP 的全部已注册工具。
MCP_BASE_URL: http://172.168.10.1:8100
---

# AutoOfficeMCP 远端能力发现与调用

本 SKILL 面向“工具列表里没有，但服务端实际支持”的场景。**不要凭印象猜工具名**：先发现，再调用。本文不维护工具清单（清单会随描述符变化），工具名、参数和示例一律以服务端文档页为准。

术语约定：

- **服务端**：运行 AutoOfficeMCP 的 Windows Office 主机。
- **调用端**：Agent / 脚本所在机器，可以是任意能发 HTTP 请求的环境。
- 除非特别说明，所有路径参数都指**服务端文件系统**，不是调用端路径。

## 0. 一句话模型

AutoOfficeMCP 是部署在 Windows Office 主机上的远端 HTTP MCP server，操作该主机上**正在运行的** Excel / Word / PowerPoint。所有工具都挂在同一个 `POST /mcp` 上；`/xlsx`、`/word`、`/ppt` 只是文档页，不是调用端点。

推荐流程：

```text
GET /health
  → 确认 Base URL 与服务状态
  → GET /<路由> 找到候选工具
  → GET /<路由>/<工具名> 查看参数与示例
  → POST /mcp tools/call（或 office_batch / office_batch_files）
  → 读回结果验证 → 按需保存 / 导出
```

## 1. 连接信息

| 项 | 值 | 说明 |
| :--- | :--- | :--- |
| Base URL | 由部署方提供 | 形如 `<MCP_BASE_URL>`；可能是反向代理或 HTTPS 地址。不要写死、不要猜地址或端口 |
| MCP 端点 | `POST /mcp` | JSON-RPC 2.0，一次请求一个 JSON 对象 |
| 文档路由 | `GET /xlsx`、`GET /word`、`GET /ppt` | 返回 Markdown 工具总览 |
| 单工具文档 | `GET /<路由>/<工具名>` | 返回该工具的参数表与说明 |
| 健康检查 | `GET /health` | 返回路由、工作目录、会话数量与并行池状态 |
| 认证 | 可选 | 若服务端启用了认证，需加 `Authorization: Bearer <token>` |

路由与应用对应关系：

| 路由 | appType | 工具名前缀 |
| :--- | :--- | :--- |
| `/xlsx` | Excel | `excel_*` |
| `/word` | Word | `word_*` |
| `/ppt` | PowerPoint | `ppt_*` |
| 任意路由 | 通用 | `office_*` |

MCP 的 `tools/list` / `tools/call` 使用**全局注册表**，不受文档路由限制；路由只影响文档页展示哪些工具。`office_*` 通用工具在三个路由的文档里都会出现。

示例使用 PowerShell 与 curl；调用端可按自身环境替换，HTTP 协议本身与语言无关。若服务端启用了认证，所有请求都要加 `Authorization: Bearer <token>`；下面的示例为简洁起见省略该头。

先确认服务可达：

```powershell
$BaseUrl = "<MCP_BASE_URL>"  
Invoke-RestMethod "$BaseUrl/health"
```

连接失败或超时，通常表示 Base URL 与部署信息不一致、网络/代理不通，或服务端未启动；先向部署方确认 Base URL。

### 1.1 不会出现在 `tools/list` 里的 HTTP 能力

这些是服务端的“非工具”能力，Agent 需要直接发 HTTP 请求：

| 端点 | 用途 |
| :--- | :--- |
| `GET /health` | 健康检查、路由、工作目录、会话数量、并行池状态 |
| `GET /xlsx` / `/word` / `/ppt` | 应用工具总览（Markdown） |
| `GET /<路由>/<工具名>` | 单工具参数页（Markdown） |
| `GET /<路由>?refresh=true` | 按当前注册表重建总览文件 |
| `POST /files/upload?name=...` | 上传文件到服务端工作目录；已存在需 `overwrite=true` |
| `GET /files/download?name=...` | 从服务端工作目录下载文件 |
| `POST /mcp` | 唯一的 MCP 调用端点（`tools/list` / `tools/call`） |
| `DELETE /mcp` | 删除当前 `Mcp-Session-Id` 会话 |
| `GET /sse` + `POST /messages?sessionId=...` | 旧版 HTTP+SSE 传输 |
| `POST /shutdown` | 停止服务端（仅在用户明确要求时使用） |

## 2. 找到工具（核心）

### 2.2 HTTP 总览与单工具页

```powershell
$BaseUrl = "<MCP_BASE_URL>"   

# 三个应用的完整工具目录（Markdown）
(Invoke-WebRequest "$BaseUrl/xlsx").Content
(Invoke-WebRequest "$BaseUrl/word").Content
(Invoke-WebRequest "$BaseUrl/ppt").Content

# 单个工具：参数、必填项、风险级别、示例
(Invoke-WebRequest "$BaseUrl/xlsx/excel_read_range").Content
```

按关键词过滤总览：

```powershell
$BaseUrl = "<MCP_BASE_URL>"
$md = (Invoke-WebRequest "$BaseUrl/ppt").Content
$md -split "`n" | Select-String -Pattern "动画|切换"
```

### 2.3 MCP `tools/list`（机器可读）

需要结构化 `inputSchema` 时使用：

```powershell
$BaseUrl = "<MCP_BASE_URL>"
$body = '{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}'
$list = Invoke-RestMethod -Method Post -Uri "$BaseUrl/mcp" `
  -ContentType "application/json; charset=utf-8" -Body $body

# 按名称或描述搜索
$list.result.tools |
  Where-Object { $_.name -match "pivot|chart" -or $_.description -match "透视表|图表" } |
  Select-Object name, description
```

### 3.3 通过 `office_batch_files` 并行处理文件

多个文件需要并行处理时使用。每项显式给文件，不依赖服务端“当前活动文档”。

```json
{
  "jsonrpc": "2.0",
  "id": 4,
  "method": "tools/call",
  "params": {
    "name": "office_batch_files",
    "arguments": {
      "confirm": true,
      "items": [
        {
          "app": "excel",
          "tool": "excel_read_range",
          "file": "report.xlsx",
          "arguments": { "range": "Data!A1:C10" },
          "savePolicy": "none"
        },
        {
          "app": "word",
          "tool": "word_document_stats",
          "file": "说明.docx",
          "savePolicy": "none"
        }
      ]
    }
  }
}
```

## 4. 工具 schema 里看不到、但运行时会用到的参数

这是“不在工具列表里的能力”的关键部分。Excel / Word 的文件级工具由服务端执行，**`file` 参数没有写进 `inputSchema`，但直接调用时必须显式提供**。

| 参数 | 适用场景 | 说明 |
| :--- | :--- | :--- |
| `file` | Excel / Word 直连工具 | 服务端工作目录中的纯文件名，或服务端主机上的绝对路径；缺省会返回“需要显式 file 参数”。也接受 `filePath`。 |
| `savePolicy` | Excel / Word 直连工具 | `none` 不写回；`save` 成功后写回。写类工具默认 `save`，只读类默认 `none`；不想改源文件时显式传 `none`。 |
| `confirm` | 所有高危工具 | 必须为 `true`，否则返回 `CONFIRMATION_REQUIRED`（可重试）。 |
| `timeoutSec` | `office_batch_files` 的 item | 单作业超时；直连工具的超时由服务端控制。 |

补充：

- 某些服务端配置下只允许显式带 `file` 的无状态调用；没有 `file` 会被拒绝，批量请用 `office_batch_files`。
- `excel_profile_files` / `excel_query_files` 在多文件只读场景可传 `backend:"openxml"`。
- PPT 直连工具和 `office_*` 通用工具依赖当前活动的演示文稿/窗口；执行前先用文档类工具确认当前文档。

## 5. 文件与工作目录

服务端提供一个工作目录；工具中传纯文件名时相对于该目录解析：

- 工具的文件参数只接受两种形式：**纯文件名**（如 `report.xlsx`）或**服务端主机上的绝对路径**（如 `D:\docs\report.xlsx`）。
- 相对子路径（如 `sub\report.xlsx`）会被拒绝。
- 文档参数会按 `appType` 校验扩展名：Excel 只接受 xlsx/xlsm/xlsb/xls/csv；Word 只接受 docx/docm/doc/rtf/txt/pdf；PPT 只接受 pptx/pptm/ppsx/pps/potx/potm/ppt。

**重要**：`file` / `path` / `saveAs` 等路径参数都指服务端主机的文件系统。调用端的文件不会被自动传输；需要先把文件上传到服务端工作目录，或把文件放到服务端可访问的绝对路径。

HTTP 文件传输（不经过工具列表）：

```powershell
$BaseUrl = "<MCP_BASE_URL>"

# 上传调用端文件到服务端工作目录；已存在时需 overwrite=true
# Windows 用 curl.exe；Linux/macOS 用 curl
curl.exe -X POST --data-binary "@./report.xlsx" `
  "$BaseUrl/files/upload?name=report.xlsx&overwrite=true"

# 从服务端工作目录下载到调用端当前目录
curl.exe -o report.xlsx "$BaseUrl/files/download?name=report.xlsx"
```

限制：上传/下载只允许纯文件名，禁止目录穿越；大小与扩展名受服务端限制（如可执行文件不允许上传）。

## 6. 安全门与调用语义

高危工具（描述符 `riskLevel=high`）必须显式确认：

```json
{ "confirm": true }
```

未确认时返回：

```json
{ "errorCode": "CONFIRMATION_REQUIRED", "recoverable": true }
```

部分高危能力（如 VBA 执行、删除工作表/行列等结构变更、关闭有未保存改动的文档、依赖剪贴板的工具）由服务端安全门控制。被拒绝时错误信息会说明原因（如 `CLIPBOARD_DISABLED`）。**不要尝试用参数绕过**；需要放行时联系部署方处理。

写操作信任合同：

- MCP 响应的 `isError=false`（内部 `Success=true`）才代表操作真实生效；失败结果的 JSON 载荷中会带 `errorCode` / `recoverable`。
- `PARTIAL_APPLY` 表示批处理中途失败，已完成的步骤不会自动回滚。
- 可用 `office_undo_last` / `office_undo` / `office_undo_list` 撤销 MCP 自己登记的节点；PPT 使用 PowerPoint 原生撤销组。结构变更、图表、条件格式等可能无法撤销，结果会给出 `undoReason`。

## 7. 读结果与验证

MCP 响应形状：

```json
{
  "result": {
    "content": [
      { "type": "text", "text": "中文说明\n{...紧凑JSON数据...}" }
    ],
    "isError": false
  }
}
```

- `isError` 才是成功/失败判据，不要只看 HTTP 200。
- 文本通常由“中文说明 + 一行紧凑 JSON”组成；需要结构化字段时解析最后一行 JSON。
- 截图、缩略图等工具会额外返回 `{ "type": "image", "data": "<base64>", "mimeType": "image/png" }`。
- JSON-RPC 层面的错误（未知方法、未知工具、参数不是对象）返回 `error` 对象；工具执行失败返回 `result.isError=true`。

写后验证建议：

1. 写入类工具成功后，用对应的读取工具回读目标区域/段落/形状。
2. 需要落盘时显式保存或导出；不要假设内存改动已经写回。
3. 批量操作后检查 `completed` / `failed` / `rollbackNote` 字段。

## 8. 服务端诊断能力

这些能力不一定在你的工具列表里，但可以通过文档页或 `office_batch` 调用：

- `GET /health`：服务是否正常、路由、工作目录、主会话与并行池状态。
- `office_pool_status`：队列深度、在跑作业、并发与文件锁状态。
- `office_list_documents`：服务端当前已打开的 Excel/Word/PPT 文档、活动文档、未保存状态。
- `office_get_window_info`：活动窗口、文档、缩放、视图等。
- `office_screenshot`：截取服务端 Office 窗口，返回 PNG image content。
- `office_undo_list`：当前服务端进程维护的撤销栈。

`POST /shutdown` 会停止服务端；除非用户明确要求，不要调用。

## 9. 常见错误与处理

| 现象 | 原因 | 处理 |
| :--- | :--- | :--- |
| 连接失败/超时 | Base URL 不对、网络/代理不通，或服务端未启动 | 向部署方确认 Base URL；访问 `$BaseUrl/health`；确认服务端已启动 |
| `INVALID_PARAMS: 未知或已禁用的工具` | 工具名拼错，或被服务端禁用 | 用文档页/`tools/list` 核对；必要时联系部署方 |
| `CONFIRMATION_REQUIRED` | 高危工具未确认 | 加 `confirm:true` 后重试 |
| `CLIPBOARD_DISABLED` | 服务端未启用剪贴板能力 | 联系部署方处理，或改用不依赖剪贴板的工具 |
| 提示 VBA / 删除结构 / 关闭未保存文档等安全门未放行 | 安全门限制 | 联系部署方处理；不要绕过 |
| “只支持纯文件名或绝对路径” | 传了相对子路径 | 改成纯文件名或服务端主机绝对路径 |
| “文件扩展名不符合 … 文档要求” | 文件类型与 app 不匹配 | 换正确应用/文件，或先转换格式 |
| “需要显式 file 参数” | Excel/Word 直连工具缺少隐式 `file` | 补 `file`；或改用 `office_batch` / `office_batch_files` |
| 无 `file` 的无状态调用被拒绝 | 并发模式限制 | 补 `file`，或改用 `office_batch_files` |
| `404 SESSION_NOT_FOUND` | `Mcp-Session-Id` 过期/服务端重启 | 丢弃旧会话，重新 `initialize` |
| `PARTIAL_APPLY` | 批处理中途失败 | 检查 `completed` / `failed`；必要时用 undo 回滚已完成项 |
| PPT 作业卡住/并发不生效 | PowerPoint 单实例限制 | 减少 PPT 并发；不要把 `ppt_execute_vba` 放进批处理 |

## 10. 行为约定

- **先发现再调用**：始终通过 `/xlsx`、`/word`、`/ppt`、`tools/list` 确认工具名和参数，不猜。
- **先读后写**：写之前先读取当前上下文；写之后读回验证。
- **默认不破坏源文件**：文件级直连工具显式传 `savePolicy:"none"`，确认无误后再保存。
- **不要用 `/xlsx`、`/word`、`/ppt` 执行工具**：它们只支持 `GET`，执行统一走 `POST /mcp`。
- **不要修改用户未要求修改的文档**：批处理只关闭自己打开的文档；不要对用户已打开的文档调用关闭类工具。
- **路径参数指服务端**：调用端文件需要先上传到服务端工作目录，或使用服务端可访问的绝对路径。
- **找不到能力时如实说明**：给出已搜索的路径和相近工具，不要编造工具名或参数。