# 全工具冒烟测试（tests/all-tools）

覆盖 `tools/list` 中的全部 **215** 个工具（Excel 74 / Word 72 / PPT 54 / Common 15），默认连接：

```text
http://127.0.0.1:8010
```

脚本**不会**启动/停止 server，也**不会**修改 `config.json`；server 需要提前运行，并建议使用最新编译的 exe。

## 快速开始

```powershell
# 首次运行：生成夹具（Office COM；只需一次）
pwsh -NoProfile -ExecutionPolicy Bypass -File tests\all-tools\prepare-fixtures.ps1

# 并发运行全部工具（Excel/Word/PPT 各两份 + Common 一份）
pwsh -NoProfile -ExecutionPolicy Bypass -File tests\all-tools\run-all.ps1 -BaseUrl http://127.0.0.1:8010
```

也可以单独执行某个分片：

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File tests\all-tools\excel-part1.ps1 -BaseUrl http://127.0.0.1:8010
pwsh -NoProfile -ExecutionPolicy Bypass -File tests\all-tools\excel-part2.ps1 -BaseUrl http://127.0.0.1:8010
pwsh -NoProfile -ExecutionPolicy Bypass -File tests\all-tools\word-part1.ps1  -BaseUrl http://127.0.0.1:8010
pwsh -NoProfile -ExecutionPolicy Bypass -File tests\all-tools\word-part2.ps1  -BaseUrl http://127.0.0.1:8010
pwsh -NoProfile -ExecutionPolicy Bypass -File tests\all-tools\ppt-part1.ps1   -BaseUrl http://127.0.0.1:8010
pwsh -NoProfile -ExecutionPolicy Bypass -File tests\all-tools\ppt-part2.ps1   -BaseUrl http://127.0.0.1:8010
pwsh -NoProfile -ExecutionPolicy Bypass -File tests\all-tools\common.ps1      -BaseUrl http://127.0.0.1:8010
```

常用参数：

| 参数 | 说明 |
| :--- | :--- |
| `-BaseUrl` | server 地址，默认 `http://127.0.0.1:8010`。 |
| `-Token` | 配置了 `transport.tokens` 时传 Bearer token。 |
| `-Limit N` | 只执行当前分片排序后的前 N 个工具，便于调试。 |
| `-Sequential` | 仅 `run-all.ps1`：串行执行各分片，便于排查并发问题。 |
| `-SkipPrepare` | 仅 `run-all.ps1`：不自动生成缺失夹具。 |

## 分片方式

每个分片都统一从 `POST /mcp` 的 `tools/list` 拉取工具，按名称排序后按 `index % Parts` 取模分配：

| 脚本 | 端点 | 工具前缀 | 说明 |
| :--- | :--- | :--- | :--- |
| `excel-part1.ps1` / `excel-part2.ps1` | `/mcp` | `excel_` | Excel 74 个工具各 37 个。 |
| `word-part1.ps1` / `word-part2.ps1` | `/mcp` | `word_` | Word 72 个工具各 36 个。 |
| `ppt-part1.ps1` / `ppt-part2.ps1` | `/mcp` | `ppt_` | PPT 54 个工具各 27 个。 |
| `common.ps1` | `/mcp` | `office_` | Common 15 个工具。 |

`run-all.ps1` 最后会重新拉取 `/mcp` 的全量工具，按前缀核对覆盖率，确保 215 个工具全部被分配到。

## 夹具与隔离

- `prepare-fixtures.ps1` 在 `tests/all-tools/fixtures/` 生成：
  - `base.xlsx` / `base-copy.xlsx`：多工作表、公式、表格、命名区域、图表、合并单元格、批注。
  - `base.docx` / `base-copy.docx`：标题、段落、表格、批注、页眉页脚、图片、书签、修订、文本框。
  - `base.pptx` / `base-copy.pptx`：两页、标题/正文、表格、图表、图片、备注、可分组矩形与已分组形状。
  - `image.png`、`data.csv`。
- 每个分片把自己的夹具复制到 **server `/health` 返回的 workspace**，使用独立的 `all-tools-<app>-<part>.*` 文件名，避免并发冲突。
- 每次工具调用前都会重置夹具，因此删除/清空/覆盖类工具的副作用不会影响后续用例。
- Common 分片使用 PPT 夹具，因为 `office_*` 活动文档类工具在固定进程模式下由主 STA 会话处理。

## 结果判定

每次调用输出一行：

```text
[PASS]  excel excel_read_range
[GATED] excel excel_delete_sheet CANCELLED
[FAIL]  excel some_tool INTERNAL_ERROR
```

| 状态 | 含义 |
| :--- | :--- |
| `PASS` | 工具返回 `isError=false`。 |
| `GATED` | 工具返回预期的参数/安全/状态门禁错误，例如 `INVALID_PARAMS`、`CANCELLED`、`VBA_DISABLED`、`CLIPBOARD_DISABLED`、`FILE_NOT_FOUND`、`APP_NOT_RUNNING`、`OFFICE_BUSY`、`EXECUTOR_NOT_AVAILABLE`。 |
| `FAIL` | `INTERNAL_ERROR`、协议失败、脚本异常等需要修复的问题。 |

`GATED` 也可能是工具的正常业务拒绝（例如当前文档没有动画、没有修订、主题文件缺失）。详细错误码和完整返回文本在：

```text
tests/all-tools/results/<app>-part<N>.json
```

`run-all.ps1` 的退出码：`0` 表示无 FAIL 且覆盖率完整；`1` 表示存在 FAIL/分片失败；`2` 表示夹具或 server 准备失败。

## 最近一次实测结果

在最新编译版本 + 本机 Office 环境下：

```text
COVERAGE excel 74/74
COVERAGE word 72/72
COVERAGE ppt 54/54
COVERAGE common 15/15

TOTAL total=215 pass=198 gated=17 fail=0
RUNALL_EXIT=0
```

剩余 17 个 `GATED` 主要是：

- `security.allowClipboard=false`：`excel_camera_picture`、`excel_export`、`excel_transform_data`、`word_image_layout`、`word_insert_equation`。
- `security.allowVba=false`：`excel_run_vba`、`ppt_execute_vba`。
- `security.allowDeleteSheet=false`：`excel_delete_sheet`、`excel_delete_row_col`、`excel_batch_rename_sheets`、`word_table_delete_row_col`。
- 状态/夹具限制：`office_stub_probe`（预期 `EXECUTOR_NOT_AVAILABLE`）、`office_undo*`（撤销栈归属其他演示文稿）、`ppt_delete_animation`（夹具无动画）、`ppt_apply_theme`（缺少 `.thmx`）、`word_compare_documents`（并发 Word COM 瞬态断开）。

这些不是 `FAIL`。如果 server 配置放开了对应安全开关，这些工具会实际执行；夹具每次重置，不会污染后续用例。

## 注意

- 需要 server 正在运行；脚本通过 `/health` 获取 workspace。
- 需要本机 Office 与 PowerShell 7（`pwsh`）。
- PPT 直连工具在主 STA 会话执行，脚本会在每次调用前打开夹具、调用后关闭。
- Excel/Word 直连工具走进程 worker，隐式携带 `file`（纯文件名从 workspace 解析）。
- 不要同时运行多个 `run-all.ps1`；同一分片重复运行会互相覆盖 `all-tools-<app>-<part>` 夹具。
