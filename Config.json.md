# config.json 配置说明

本文档说明 AutoOfficeMCP 的 `config.json`：每个配置项的作用、默认值、取值范围、生效范围，以及 HTTP 认证与 userId 的行为。

> 适用对象：当前仓库源码（.NET Framework 4.8，`AutoOfficeMCP.exe`）。
> 传输固定为 HTTP/SSE；worker 固定为独立进程模式。配置修改后需要**重启 server**才会生效。

---

## 1. 文件位置与加载规则

| 项目 | 说明 |
| :--- | :--- |
| 固定路径 | `AppDomain.CurrentDomain.BaseDirectory\config.json`，即 **exe 同目录**。工程根目录的 `config.json` 会在构建时复制到输出目录（当前为 `F:\c#2\AutoOfficeMCP\config.json`，见 `AutoOfficeMCP.csproj`）。 |
| 覆盖方式 | **不读取环境变量，不支持命令行参数**；只能修改文件。 |
| 加载时机 | 进程启动时由 `AppConfig.Current` 读取一次并缓存；没有文件监视。 |
| 缺失/损坏 | 文件不存在、内容为空、JSON 解析失败：全部使用默认值，并向 stderr 输出 `[config] ...`，server 继续启动。 |
| 字段容错 | 字段缺失、类型不匹配、枚举值未知：回退默认值；部分数值有下限/上限钳制（见各表）。 |

---

## 2. HTTP 认证与 userId

### 2.1 认证方式：只保留 `transport.tokens`

`transport.tokens` 是 `token → userId` 映射，也是当前唯一的 HTTP 认证方式：

```json
"transport": {
  "tokens": {
    "token-a": "alice",
    "token-b": "bob"
  }
}
```

- `tokens` 非空：所有 HTTP 业务端点都要求 `Authorization: Bearer <token>`。
  - 缺失或不是 `Bearer` 格式：401。
  - token 不在映射中：403。
  - 命中后 `userId` 使用映射值；映射值为空时使用 `"user"`。
- `tokens` 为空或未配置：**不做认证**，所有请求直接放行；`userId` 自动为 `"anonymous"`。
- 空白 token 键、空白映射值会被忽略。
- token 比较使用固定时间比较（`FixedTimeEquals`）。

> 旧的 `transport.authToken`（单 token）与 `transport.allowAnonymous` 已删除。需要单 token 时，在 `transport.tokens` 中配置一条映射即可。

### 2.2 userId 的用途

`userId` 不是登录账号，只用于 worker 路由和日志：

- `balanced` 模式下，`WorkerProcessPool` 以 `userId + app` 作为 worker 的键；同一用户同一应用串行，不同用户/应用可并行。
- `throughput` 模式下，userId 会被替换为 `throughput-<app>-<slot>` 槽位 ID。
- 未配置 `tokens` 时自动使用 `"anonymous"`，不需要手动分配。

### 2.3 认证判定顺序

```text
请求进入
  │
  ├─ transport.tokens 非空？
  │    ├─ 是：要求 Bearer token，命中映射后 userId = 映射值
  │    └─ 否：直接放行，userId = "anonymous"
  │
  └─ 认证失败：缺失/格式错误 401；token 不匹配 403
```

### 2.4 生效范围

所有 HTTP 业务端点统一认证：`/mcp`、`/xlsx/mcp`、`/word/mcp`、`/ppt/mcp`、`/sse`、`/messages`、`/health`、`/shutdown`、`/files/upload`、`/files/download` 及路由级文件接口。

> 不再提供 stdio 传输。server 启动后只监听 HTTP/SSE，不占用 stdin/stdout。

---

## 3. 完整配置结构

当前仓库的 `config.json`：

```json
{
  "transport": {
    "bind": "127.0.0.1",
    "port": 8765,
    "maxBodyBytes": 8388608
  },
  "server": {
    "maxInflightPerSession": 1
  },
  "worker": {
    "concurrencyMode": "balanced",
    "maxWarmWorkers": 8,
    "maxWorkersPerUser": 1,
    "idleTimeoutSec": 600,
    "startupTimeoutSec": 30,
    "jobTimeoutSec": 300,
    "memoryLimitMb": 0,
    "pipeNamePrefix": "AutoOfficeMCP.Worker",
    "heartbeatIntervalSec": 10,
    "heartbeatTimeoutSec": 30,
    "saveVerifyMaxHashMb": 64,
    "fileLockTtlSec": 3600,
    "fileLockProvider": "file",
    "fileLockRoot": "",
    "sqlLockConnectionString": ""
  },
  "office": {
    "excel": { "autoLaunch": true, "visible": true },
    "word":  { "autoLaunch": true, "visible": true },
    "ppt":   { "autoLaunch": true, "visible": true }
  },
  "session": { "timeoutSec": 120 },
  "security": {
    "allowVba": false,
    "allowDeleteSheet": false,
    "allowCloseWithoutSave": false,
    "allowClipboard": false,
    "disableAddins": true,
    "wordTemplatePath": "Templates\\AutoOfficeMCP.dotx"
  },
  "tools": { "disabled": [] },
  "workspace": {
    "root": "workspace",
    "maxUploadBytes": 268435456
  },
  "parallel": {
    "excel": { "maxConcurrency": 2 },
    "word":  { "maxConcurrency": 1 },
    "ppt":   { "maxConcurrency": 2 },
    "itemTimeoutSec": 300,
    "maxQueue": 200,
    "idempotencyTtlSec": 300
  }
}
```

> `transport.tokens` 是可选认证配置，当前文件未写入。需要认证或多用户时手动添加：
>
> ```json
> "tokens": {
>   "token-a": "alice",
>   "token-b": "bob"
> }
> ```

---

## 4. 配置项速查

### 4.1 transport（HTTP/SSE 入口）

| 配置项 | 默认值 | 说明 |
| :--- | :--- | :--- |
| `bind` | `"127.0.0.1"` | HTTP 监听地址。`"*"` 或 `"0.0.0.0"` 会转换为 `"+"`（监听所有网卡）。非 loopback 绑定通常需要 Windows URL ACL 或管理员权限；不要把端口直接暴露到公网。 |
| `port` | `8765` | HTTP 监听端口。 |
| `maxBodyBytes` | `8388608` | 单个 HTTP 请求体上限（字节），最小 1024。`Content-Length` 或实际 UTF-8 字节数超限返回 413。 |
| `tokens` | 未配置（空映射） | `token → userId` 映射。非空时所有 HTTP 端点要求 Bearer token；为空时不做认证，userId 自动为 `"anonymous"`。 |

### 4.2 server（会话并发）

| 配置项 | 默认值 | 说明 |
| :--- | :--- | :--- |
| `maxInflightPerSession` | `1` | 单个 MCP 会话（HTTP/SSE）的最大并发请求数，最小 1。超出的请求会在该会话的 `SemaphoreSlim` 上排队等待，不直接拒绝。 |

### 4.3 worker（固定进程 worker 池）

worker 固定使用 `WorkerProcessPool`：每个 `userId + app` 一个子进程，进程内一个 STA 线程 + 一个常驻 `Application`。不再提供 `worker.mode` / `processPerUser` 配置，也没有 inProcess 回退。

| 配置项 | 默认值 | 说明 |
| :--- | :--- | :--- |
| `concurrencyMode` | `"balanced"` | `balanced`：按真实 userId 分配 worker；`throughput`：按 `parallel.<app>.maxConcurrency` 把 worker 分片为 `throughput-<app>-<slot>` 槽位轮转。未知值回退 `balanced`。 |
| `maxWarmWorkers` | `8` | worker 总数上限，最小 1。池满且没有空闲 worker 时返回 `OFFICE_BUSY`；有空闲 worker 时优先回收。 |
| `maxWorkersPerUser` | `1` | 同一 `userId + app` 的 worker 数上限，最小 1。 |
| `idleTimeoutSec` | `600` | worker 空闲回收秒数，最小 30；空闲扫描周期约 30 秒。 |
| `startupTimeoutSec` | `30` | worker 命名管道连接与 ready 握手超时；进程池内最小按 5 处理。 |
| `jobTimeoutSec` | `300` | 单个作业的执行超时，最小按 5 处理；item 自带 `TimeoutSec` 时优先使用 item 的值。 |
| `memoryLimitMb` | `0` | 通过 Job Object 限制单个 worker 进程内存（MB）；`0` 表示不限制。 |
| `pipeNamePrefix` | `"AutoOfficeMCP.Worker"` | worker 命名管道前缀，实际管道名为“前缀 + GUID”。 |
| `heartbeatIntervalSec` | `10` | worker → Gateway 心跳间隔，最小 2。 |
| `heartbeatTimeoutSec` | `30` | 超过该时间未收到心跳视为失联并重建；实际下限为 `max(6, 心跳间隔 × 2, 配置值)`。 |
| `saveVerifyMaxHashMb` | `64` | 保存后文件完整性校验的 SHA256 阈值（MB）。`0` 表示只比较长度 + mtime；文件超过阈值时同样退回长度 + mtime。 |
| `fileLockTtlSec` | `3600` | 跨进程锁文件空闲 TTL，最小 60；只清理确认未被持有的陈旧锁。 |
| `fileLockProvider` | `"file"` | `file`：锁文件；`sql`：SQL Server 分布式锁；未知值回退 `file`。`sql` 且连接字符串非空时使用 SQL 锁。 |
| `fileLockRoot` | `""` | 锁文件根目录；留空使用 `%TEMP%\AutoOfficeMCP\locks`，可填 UNC 共享路径实现多机互斥。仅 `fileLockProvider="file"` 时生效。 |
| `sqlLockConnectionString` | `""` | SQL Server 连接字符串；`fileLockProvider="sql"` 时必须填写，否则记录警告并回退锁文件。 |

**文件参数与 `throughput` 说明**

- 直连 Excel/Word 工具必须显式传 `file`，否则返回 `INVALID_PARAMS`。
- `file` 接受纯文件名或工作目录内的相对子路径（从 `workspace.root` 解析），也接受完整本机绝对路径；相对子路径禁止 `..` 与盘符。
- `throughput` 模式下 worker 按应用槽位共享，适合无状态文件任务；批量任务建议使用 `office_batch_files`。

### 4.4 office（Office 启动与可见性）

| 配置项 | 默认值 | 说明 |
| :--- | :--- | :--- |
| `excel.autoLaunch` | `true` | 主 STA 会话未找到运行中的 Excel 时，是否由 server 启动；`false` 时返回 `APP_NOT_RUNNING`。 |
| `excel.visible` | `true` | 主 STA 会话/Common 文档工具自动启动 Excel、以及 Excel worker 创建实例时是否显示窗口；附着到已有实例时不改变其可见性。 |
| `word.autoLaunch` | `true` | 主 STA 会话未找到运行中的 Word 时，是否由 server 启动。 |
| `word.visible` | `false` | 主 STA 会话自动启动 Word、以及 Word worker 创建实例时是否显示窗口；附着到已有实例时不改变其可见性。 |
| `ppt.autoLaunch` | `true` | 主 STA 会话未找到运行中的 PowerPoint 时，是否由 server 启动。 |
| `ppt.visible` | `true` | 主 STA 会话自动启动 PowerPoint、以及 PowerPoint worker 需要新建实例时是否显示窗口。附着到用户已运行的实例时不改变其可见性。 |

> `autoLaunch` 只作用于主 STA 会话的“附着或启动”逻辑；Excel/Word worker 由 `WorkerHost` 直接创建实例，不读取 `autoLaunch`。`visible` 同时作用于主 STA 会话和 worker：Excel/Word 直连工具在固定进程模式下会转到 worker，worker 按 `excel.visible` / `word.visible` 设置窗口可见性；`office_batch_files` 的 PPT 作业由 worker 的 `PowerPointApplicationLease` 处理，需要新建实例时遵循 `ppt.visible`（附着已有实例时保持原可见性）。因此把 `visible` 设为 `true` 时，批量/并行作业也会显示对应的 Office 窗口。PowerPoint 是单实例应用，`parallel.ppt.maxConcurrency` 控制调度槽位，真正的隔离并发受 worker/Windows 会话限制。

### 4.5 session（主会话锁）

| 配置项 | 默认值 | 说明 |
| :--- | :--- | :--- |
| `timeoutSec` | `120` | 主 STA 会话锁 `SessionGate` 的等待超时（秒）。小于等于 0 时按 120 处理；超时返回 `SERVER_BUSY`。该值不等于 worker 作业超时，worker 超时见 `worker.jobTimeoutSec` / `parallel.itemTimeoutSec`。 |

### 4.6 security（高危操作与兼容开关）

| 配置项 | 默认值 | 说明 |
| :--- | :--- | :--- |
| `allowVba` | `false` | 是否允许 `excel_run_vba` / `ppt_execute_vba`；`false` 返回 `VBA_DISABLED`。还要满足 Office 信任中心的“信任对 VBA 工程对象模型的访问”。 |
| `allowDeleteSheet` | `false` | 是否允许 `excel_delete_sheet`；`false` 返回 `CANCELLED`。 |
| `allowCloseWithoutSave` | `false` | 是否允许关闭有未保存改动的文档；`false` 返回 `CANCELLED`。 |
| `allowClipboard` | `false` | 是否允许依赖 Windows 剪贴板的工具：`excel_transform_data`、`excel_export`、`excel_camera_picture`、`word_image_layout`、`word_insert_equation`；`false` 返回 `CLIPBOARD_DISABLED`。 |
| `disableAddins` | `true` | worker 启动时是否断开 COM/Office 加载项（`COMAddIns` / `AddIns`）。 |
| `wordTemplatePath` | `"Templates\\AutoOfficeMCP.dotx"` | Word 新文档使用的专用模板。相对路径以 exe 同目录为基准；模板缺失时返回 `FILE_NOT_FOUND`，不会静默退回 `Normal.dotm`。 |

> 高危工具还需要协议层 `confirm=true` 放行；`security.*` 只是第二道门禁。

### 4.7 tools（工具开关）

| 配置项 | 默认值 | 说明 |
| :--- | :--- | :--- |
| `disabled` | `[]` | 精确工具名数组，大小写不敏感。命中的工具不会出现在 `tools/list`，也无法被调用；未知名称会被忽略。 |

### 4.8 workspace（工作目录）

| 配置项 | 默认值 | 说明 |
| :--- | :--- | :--- |
| `root` | `"workspace"` | 工作目录。相对路径以 exe 同目录为基准，支持环境变量展开；目录不存在时自动创建。工具文件参数的纯文件名/相对子路径、HTTP 上传/下载根目录、`xlsx.md` / `word.md` / `ppt.md` 输出目录都以它为根。写入/导出类工具**仅在本目录内**自动创建缺失的父目录；工作目录外的目录必须预先存在。 |
| `maxUploadBytes` | `268435456` | HTTP 单次上传大小上限（字节），最小 1024；只限制上传，不限制工具本地读写。 |

### 4.9 parallel（批量并行调度）

| 配置项 | 默认值 | 说明 |
| :--- | :--- | :--- |
| `excel.maxConcurrency` | `2` | 并行池 Excel 作业并发槽位数，最小 1。 |
| `word.maxConcurrency` | `1` | 并行池 Word 作业并发槽位数，最小 1。 |
| `ppt.maxConcurrency` | `2` | 并行池 PowerPoint 作业并发槽位数，最小 1；PowerPoint 为单实例 COM 应用，真正隔离并发仍受 worker/Windows 会话限制。 |
| `itemTimeoutSec` | `300` | 单个并行 item 从提交起算的总超时（含排队、等待应用槽、等待文件锁、执行），最小 1；item 自带 `TimeoutSec` 时优先使用 item 的值。 |
| `maxQueue` | `200` | 未完成并行作业数量上限，最小 1；超过后返回 `OFFICE_BUSY`。`office_batch_files.items` 数量也受此限制。 |
| `idempotencyTtlSec` | `300` | `idempotencyKey` 完成结果缓存秒数；`0` 表示只做在飞去重，不缓存完成结果。 |

---

## 5. 配置示例

### 5.1 本机无认证（当前配置）

```json
"transport": {
  "bind": "127.0.0.1",
  "port": 8765,
  "maxBodyBytes": 8388608
}
```

客户端不需要 `Authorization` 头；只能从本机访问。

### 5.2 多用户 token

```json
"transport": {
  "bind": "127.0.0.1",
  "port": 8765,
  "tokens": {
    "token-a": "alice",
    "token-b": "bob"
  }
}
```

客户端：

```text
Authorization: Bearer token-a
```

### 5.3 单 token（用一条映射实现）

```json
"transport": {
  "bind": "0.0.0.0",
  "port": 8765,
  "tokens": {
    "replace-with-a-strong-token": "default"
  }
}
```

---

## 6. 已移除配置项

旧配置中如仍存在以下键会被忽略：

| 配置项 | 说明 |
| :--- | :--- |
| `transport.mode` | 传输固定为 HTTP/SSE，不再读取该键。 |
| `transport.stdioCompat` | stdio 传输已删除。 |
| `transport.authToken` | 单 token 认证已删除；改用 `transport.tokens`。 |
| `transport.allowAnonymous` | 匿名认证开关已删除；未配置 `tokens` 时直接放行。 |
| `server.maxInflightRequests` | 该限流项已删除；实际使用 `server.maxInflightPerSession`。 |
| `worker.mode` | worker 固定为进程模式。 |
| `worker.processPerUser` | 进程 worker 本身按 `userId + app` 隔离，该布尔项已删除。 |

---

## 7. 修改与排错

1. **修改位置**：改 exe 同目录的 `config.json`（构建输出目录），不要只改工程根目录后直接运行旧 exe。
2. **重启生效**：没有文件监视/热加载；修改后重启 server。
3. **启动确认**：stderr 的 `[NOTICE] HTTP 监听 ... auth=...` 会显示监听地址、端口和认证模式（`tokens(n)` 或 `none`）。
4. **配置写坏**：JSON 解析失败会回退全部默认值并打印 `[config] 读取失败`；server 仍会启动，但行为可能与预期不同。
5. **认证失败**：仅在配置了 `transport.tokens` 时出现。401 表示缺少/格式错误的 `Authorization`；403 表示 token 不匹配。
6. **端口占用**：`HttpListener` 启动失败会写 stderr；换 `transport.port` 或释放端口。
7. **非本机访问失败**：默认 `bind=127.0.0.1` 只允许本机；改为对外监听前先配置 `transport.tokens`，并建议通过 Tailscale/VPN/反向代理 + TLS 访问。

---

## 8. 相关实现位置

| 文件 | 职责 |
| :--- | :--- |
| `Office/Config/AppConfig.cs` | 读取配置、类型转换、默认值与数值钳制。 |
| `Transport/HttpTransportHost.cs` | HTTP 监听、`transport.tokens` 认证、请求体限制、会话并发。 |
| `Program.cs` | 启动 HTTP/SSE 入口。 |
| `Office/Parallel/ParallelDispatcher.cs` | `parallel.*`、`worker.concurrencyMode`、固定进程 worker 池、锁与超时。 |
| `Office/Parallel/WorkerProcessPool.cs` | 进程 worker 池、心跳、内存限制、作业超时。 |
| `Office/OfficeSession.cs` | 主 STA 会话、`office.*.autoLaunch/visible`、`session.timeoutSec`。 |
| `Office/SecurityGate.cs` | `security.allowVba` / `allowDeleteSheet` / `allowCloseWithoutSave`。 |
| `Office/Security/ClipboardGuard.cs` | `security.allowClipboard`。 |
| `Worker/WorkerHost.cs` | worker 进程启动、`security.disableAddins`、心跳。 |
| `Office/Config/WorkspacePathResolver.cs` | `workspace.root` 解析与自动创建。 |
| `Tools/ToolRegistry.cs` | `tools.disabled` 过滤。 |
