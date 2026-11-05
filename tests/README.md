# 测试与质量门禁

本目录是 AutoOfficeMCP 当前的质量门禁。所有脚本优先使用 PowerShell 7（`pwsh`），并且**不设置环境变量**；传输端口等通过临时改写 exe 同目录 `config.json` 注入，脚本结束自动恢复。

## 快速开始

```powershell
# 7 个静态门禁 + 7 组“基线绿 → 变异红”探针
pwsh -NoProfile -ExecutionPolicy Bypass -File tests\static\run-all.ps1

# Word TODO 真机验证（页/段映射、批注、样式、域、带格式替换）
pwsh -NoProfile -ExecutionPolicy Bypass -File tests\transport\test-word-todo.ps1

# Excel TODO 真机验证（重算、分页、审计、动作、撤销、保护、图表等）
pwsh -NoProfile -ExecutionPolicy Bypass -File tests\transport\test-excel-todo.ps1

# 进程 worker 文件参数约定（无 file / 纯文件名 / 绝对路径 / 相对子路径）
pwsh -NoProfile -ExecutionPolicy Bypass -File tests\transport\test-file-params-process-mode.ps1

# PPT 真机验证（进程 worker 模式；自动启动/附着 PowerPoint）
pwsh -NoProfile -ExecutionPolicy Bypass -File tests\transport\test-ppt-tools.ps1

# PPT 进程 worker 验证（直连工具 + office_batch_files）
pwsh -NoProfile -ExecutionPolicy Bypass -File tests\transport\test-ppt-process-mode.ps1

# PPT 第二轮查漏补缺（层级、动画编排、表格、组合/对齐、图片编辑、跨文件复制、版面自检、Excel 直通）
pwsh -NoProfile -ExecutionPolicy Bypass -File tests\transport\test-ppt-gaps.ps1

# 全工具冒烟（215 个；需要 server 已在 127.0.0.1:8010 运行）
pwsh -NoProfile -ExecutionPolicy Bypass -File tests\all-tools\prepare-fixtures.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File tests\all-tools\run-all.ps1 -BaseUrl http://127.0.0.1:8010
```

全工具冒烟的分片、结果判定与注意事项见 [all-tools/README.md](all-tools/README.md)。

> `test-word-todo.ps1` / `test-excel-todo.ps1` 仍按已移除的 inProcess 活动文档语义编写，尚未迁移到固定进程 worker 的显式 `file` 模式；文件参数规则请用 `test-file-params-process-mode.ps1` 验证。

退出码：`0` 通过；`1` 门禁红/断言失败；`2` 被阻塞（例如缺少 exe 或 Office），**不算通过**。

## 静态门禁

| 门禁 | 检查事实 |
| :--- | :--- |
| `check-stdout-purity.ps1` | HTTP/SSE 模式不得打开标准输出；Program.cs/Protocol 不得出现 `Console.WriteLine`/`Console.Out`/`Console.OpenStandardOutput` |
| `check-utf8-nobom.ps1` | `.cs` / `.ps1` / `.psm1` / `.json` 为有效 UTF-8 且无 BOM |
| `check-no-var.ps1` | C# 源码剥离注释/字符串后不得出现 `var` |
| `check-compile-sync.ps1` | 4 个 csproj 的 `<Compile Include>` 与磁盘 `.cs` 双向一致 |
| `check-descriptor-schema.ps1` | 描述符必填字段、参数类型白名单、name/id 全局唯一、appType/riskLevel 白名单 |
| `check-registry-drift.ps1` | 描述符 name 集合与执行器静态工具名集合双向一致；动态匹配 fail-closed |
| `check-error-codes.ps1` | `OfficeErrorCode` 每个枚举成员都有非空中文消息 |

`run-all.ps1` 自动发现 `check-*.ps1`，发现数少于 6 视为发现机制损坏；随后为 7 个门禁分别植入违规，要求门禁变红且报错命中预期。

## Word / Excel 真机测试

- `test-word-todo.ps1`：页/段映射与跳页、批注完成/删除、样式清单/定义修改、域清单、`word_replace_text` 带格式替换与旧行为兼容。
- `test-excel-todo.ps1`：重算与模式恢复、`excel_read_range` 三元组/分页/旧字段兼容、引用审计、精度扫描、选择性粘贴/公式固化、格式复制、多键排序、超链接、合并单元格、建簿/另存、保护、图表增改删、条件格式 list/delete、`office_undo`/`office_undo_list`、函数目录版本/fallback/模板。

## PPT 真机测试

`test-ppt-tools.ps1` 覆盖 37 个基础 PPT 工具与 PPT 相关 common 工具（新增工具由 `test-ppt-gaps.ps1` 覆盖），包括：
读取、插入文本/形状/表格/图片/图表、形状格式、文字格式、项目符号、段落、隐藏/复制/移动/重排/清空/页面尺寸、备注、动画、切换、美化、导出图片/PDF、缩略图、单形状删除后原生撤销、VBA 默认安全门禁、`office_open_document`/`list`/`activate`/`save`/`screenshot`/`close`、`office_batch_files` PPT。

`test-ppt-process-mode.ps1` 在固定进程 worker 模式下验证：
- PPT 直连工具由 Gateway 主 STA 会话执行，不再返回 `INVALID_PARAMS`；
- `office_batch_files` 的 `app=ppt` 作业通过 `worker.settings.json` 启动 worker 执行。

`test-ppt-gaps.ps1` 覆盖第二轮 17 个新增工具与契约修正：形状层级、动画编排/回读/删除、切换回读、表格读改格式化、组合/解组/对齐、图片裁切/替换/边框阴影/圆角、跨文件复制页、版面自检、Excel 表格/图表直通、`ppt_insert_slide position=index`、`includeFormat` 回读、`office_pool_status` PPT 分项、批次拒绝 `ppt_execute_vba`。

## 已知限制

- PowerPoint 单实例：worker 先附着用户已运行的 PowerPoint；非自建实例绝不 Quit，避免关闭用户文档。PPT 文件作业由 `parallel.ppt.maxConcurrency` 控制槽位，单实例下真正隔离并发受 worker/Windows 会话限制。
- `office_undo_last` 对 PPT 使用原生撤销组 + 状态签名校验。单形状删除已真机验证；批量 `clear_slide` 等复合操作若被 PowerPoint 拆成多组，会尝试多次 Undo，仍无法恢复时如实返回失败并保留节点。
- `ppt_execute_vba` 的 `timeoutSec` 只做参数记录，PowerPoint COM 无法中断已同步执行的 VBA；`Debug.Print` 输出无法回读；`office_batch_files` 直接拒绝该工具。
- `ppt_set_image_format` 的 `cornerRadius` 会把图片转换为带图片填充的圆角矩形（返回新 `shapeId`），原图片的动画与替代文字不保留。
- `ppt_validate_layout` 的字体检查基于本机 `InstalledFontCollection`；图片分辨率按 PowerPoint 导出的形状 PNG 估算，不等同于源文件像素元数据。表格边框设置按 COM Borders 索引尽力执行，个别主题下可能部分边框不生效。
- `ppt_insert_table_from_excel` / `ppt_insert_chart_from_excel` 会启动一个隐藏的独立 Excel 实例读取文件；读取完成后关闭工作簿并退出该实例，不修改源文件。
