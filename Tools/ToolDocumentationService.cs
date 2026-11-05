using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Office.Config;

namespace AutoOfficeMCP.Tools
{
    /// <summary>
    /// 路由工具说明（ToolMD）。说明文件直接写入 config.json workspace.root 根目录，
    /// 文件名为 xlsx.md / word.md / ppt.md，不设独立代理目录。
    /// </summary>
    public static class ToolDocumentationService
    {
        /// <summary>构建指定路由的 Markdown 说明（app 专属工具 + common 工具）。</summary>
        public static string BuildMarkdown(ToolRoute route, ToolRegistry registry)
        {
            if (route == null)
            {
                throw new ArgumentNullException("route");
            }

            List<ToolDescriptor> primary = new List<ToolDescriptor>();
            List<ToolDescriptor> common = new List<ToolDescriptor>();
            if (registry != null)
            {
                foreach (ToolDescriptor descriptor in registry.Tools)
                {
                    if (string.Equals(descriptor.AppType, route.AppType, StringComparison.OrdinalIgnoreCase))
                    {
                        primary.Add(descriptor);
                    }
                    else if (string.Equals(descriptor.AppType, "common", StringComparison.OrdinalIgnoreCase))
                    {
                        common.Add(descriptor);
                    }
                }
            }

            StringBuilder builder = new StringBuilder();
            builder.AppendLine("# /" + route.Key + " 工具说明");
            builder.AppendLine();
            builder.AppendLine("> 由 AutoOfficeMCP 根据工具描述符自动生成。本文件直接位于 config.json " +
                               "`workspace.root` 指定的工作目录根目录。");
            builder.AppendLine("> 启动时仅在文件缺失时生成，不会覆盖手工补充内容；" +
                               "需要按最新描述符重建时访问 `/" + route.Key + "?refresh=true`；" +
                               "单工具说明页见 `/" + route.Key + "/<工具名称>`。");
            builder.AppendLine("> 路由：" + route.Key + "；应用：" + route.DisplayName +
                               "；工具总数：" + (primary.Count + common.Count).ToString() + "。");
            builder.AppendLine("> 文件参数：纯文件名或工作目录内的相对子路径从 `workspace.root` 解析，" +
                               "也可传完整本机绝对路径；写入类工具仅在工作目录内自动创建缺失的父目录。");
            builder.AppendLine();

            if (primary.Count + common.Count == 0)
            {
                builder.AppendLine("当前路由没有注册任何工具。");
                builder.AppendLine();
                return builder.ToString();
            }

            AppendToolTable(builder, route, route.DisplayName + " 专属工具", primary);
            AppendToolTable(builder, route, "通用工具（跨应用）", common);

            builder.AppendLine("## 文件参数约定");
            builder.AppendLine();
            builder.AppendLine("- 文件参数接受三种形式：**纯文件名**（如 `report.xlsx`）、**工作目录内的相对子路径**（如 `sub\\report.xlsx`）或**本机绝对路径**（如 `D:\\docs\\report.xlsx`）。");
            builder.AppendLine("- 纯文件名与相对子路径从工作目录解析；相对路径禁止 `..` 与盘符。写入类工具仅在工作目录内自动创建缺失的父目录，工作目录外需预先存在。");
            builder.AppendLine("- 文档类参数（`file` / `files` 等）必须与路由应用类型匹配，例如 /" + route.Key +
                               " 只接受 `" +
                               string.Join("` / `", WorkspacePathResolver.GetDocumentExtensions(route.AppType)) + "`。");
            builder.AppendLine("- 上传/下载接口只允许工作目录内的纯文件名，禁止路径穿越。");
            builder.AppendLine();
            return builder.ToString();
        }

        /// <summary>
        /// 启动时补齐缺失的路由说明文件；已存在的文件不覆盖，保留人工补充内容。
        /// </summary>
        public static void GenerateAll(ToolRegistry registry)
        {
            StringBuilder errors = new StringBuilder();
            foreach (ToolRoute route in ToolRouteCatalog.Routes)
            {
                ToolRegistry scoped = registry == null
                    ? null
                    : registry.CreateScoped(route.AppType, true);
                string path;
                string error;
                GetMarkdown(route, scoped, false, out path, out error);
                if (!string.IsNullOrWhiteSpace(error))
                {
                    errors.AppendLine(error);
                }
            }

            if (errors.Length > 0)
            {
                Console.Error.WriteLine("生成路由工具说明失败：" + Environment.NewLine +
                                        errors.ToString().TrimEnd());
            }
        }

        /// <summary>
        /// 读取路由说明：文件存在时直接返回（保留人工补充）；缺失或 refresh=true 时按描述符重建。
        /// </summary>
        public static string GetMarkdown(
            ToolRoute route,
            ToolRegistry registry,
            bool refresh,
            out string path,
            out string error)
        {
            path = string.Empty;
            error = string.Empty;

            string root;
            string rootError;
            if (!WorkspacePathResolver.TryGetWorkspaceRoot(out root, out rootError))
            {
                error = rootError;
                return BuildMarkdown(route, registry);
            }

            path = Path.Combine(root, route.Key + ".md");
            if (!refresh && File.Exists(path))
            {
                try
                {
                    return File.ReadAllText(path, Encoding.UTF8);
                }
                catch (Exception ex)
                {
                    error = "读取工具说明失败：" + path + "；" + ex.Message;
                }
            }

            string markdown = BuildMarkdown(route, registry);
            string writePath;
            string writeError;
            if (!TryWriteDocument(route, markdown, out writePath, out writeError))
            {
                if (string.IsNullOrWhiteSpace(error))
                {
                    error = writeError;
                }
            }
            else
            {
                path = writePath;
            }
            return markdown;
        }

        /// <summary>
        /// 构建单工具 Markdown 说明（GET /xlsx/&lt;工具名称&gt; 文档页）。
        /// 动态生成、不落盘；common 工具在任意路由下均可查看。
        /// </summary>
        public static string BuildToolMarkdown(ToolRoute route, ToolDescriptor descriptor)
        {
            if (route == null)
            {
                throw new ArgumentNullException("route");
            }
            if (descriptor == null)
            {
                throw new ArgumentNullException("descriptor");
            }

            StringBuilder builder = new StringBuilder();
            builder.AppendLine("# " + descriptor.ToolName);
            builder.AppendLine();
            builder.AppendLine("> 路由：/" + route.Key + "；应用：" + route.DisplayName +
                               "；类别：" + descriptor.Category + "；风险级别：" + descriptor.RiskLevel + "。");
            builder.AppendLine("> 调用方式：`POST /mcp`，`tools/call(name=\"" + descriptor.ToolName + "\")`。");
            builder.AppendLine("> 返回总览：[" + route.DisplayName + " 工具说明](/" + route.Key + ")。");
            builder.AppendLine();

            builder.AppendLine("## 说明");
            builder.AppendLine();
            builder.AppendLine(string.IsNullOrWhiteSpace(descriptor.Description)
                ? "（无描述。）"
                : descriptor.Description);
            builder.AppendLine();

            builder.AppendLine("## 参数");
            builder.AppendLine();
            if (descriptor.Parameters == null || descriptor.Parameters.Count == 0)
            {
                builder.AppendLine("无参数。");
            }
            else
            {
                builder.AppendLine("| 参数 | 类型 | 必填 | 说明 |");
                builder.AppendLine("|:---:|:---:|:---:|:---|");
                foreach (ToolParameter parameter in descriptor.Parameters)
                {
                    builder.AppendLine("| `" + EscapeCell(parameter.Name) + "` | " +
                                       EscapeCell(parameter.Type) + " | " +
                                       (parameter.Required ? "是" : "否") + " | " +
                                       EscapeCell(parameter.Description) + " |");
                }
            }
            builder.AppendLine();
            return builder.ToString();
        }

        public static bool TryWriteDocument(
            ToolRoute route,
            string markdown,
            out string path,
            out string error)
        {
            path = string.Empty;
            error = string.Empty;

            string root;
            string rootError;
            if (!WorkspacePathResolver.TryGetWorkspaceRoot(out root, out rootError))
            {
                error = rootError;
                return false;
            }

            try
            {
                path = Path.Combine(root, route.Key + ".md");
                File.WriteAllText(path, markdown ?? string.Empty, new UTF8Encoding(false));
                return true;
            }
            catch (Exception ex)
            {
                error = "写入工具说明失败：" + path + "；" + ex.Message;
                return false;
            }
        }

        private static void AppendToolTable(
            StringBuilder builder,
            ToolRoute route,
            string title,
            List<ToolDescriptor> tools)
        {
            if (tools.Count == 0)
            {
                return;
            }

            List<ToolDescriptor> sorted = new List<ToolDescriptor>(tools);
            sorted.Sort(delegate(ToolDescriptor left, ToolDescriptor right)
            {
                return string.Compare(left.ToolName, right.ToolName, StringComparison.OrdinalIgnoreCase);
            });

            builder.AppendLine("## " + title + "（" + sorted.Count.ToString() + "）");
            builder.AppendLine();
            builder.AppendLine("| 工具名称 | 接口 | 参数说明 |");
            builder.AppendLine("|:---:|:---:|:---:|");
            foreach (ToolDescriptor descriptor in sorted)
            {
                string interfaceCell = "`POST /mcp`<br>`tools/call(name=\"" +
                                       EscapeCell(descriptor.ToolName) + "\")`<br>[详情](/" +
                                       route.Key + "/" + descriptor.ToolName + ")";
                if (!string.IsNullOrWhiteSpace(descriptor.Description))
                {
                    interfaceCell = interfaceCell + "<br>" + EscapeCell(descriptor.Description);
                }
                builder.AppendLine("| `" + EscapeCell(descriptor.ToolName) + "` | " +
                                   interfaceCell + " | " +
                                   BuildParameterCell(descriptor) + " |");
            }
            builder.AppendLine();
        }

        private static string BuildParameterCell(ToolDescriptor descriptor)
        {
            if (descriptor.Parameters == null || descriptor.Parameters.Count == 0)
            {
                return "无参数。";
            }

            StringBuilder builder = new StringBuilder();
            for (int index = 0; index < descriptor.Parameters.Count; index++)
            {
                if (index > 0)
                {
                    builder.Append("<br>");
                }
                ToolParameter parameter = descriptor.Parameters[index];
                builder.Append("`").Append(EscapeCell(parameter.Name)).Append("`（")
                       .Append(EscapeCell(parameter.Type)).Append("，")
                       .Append(parameter.Required ? "必填" : "可选").Append("）：")
                       .Append(EscapeCell(parameter.Description));
            }
            return builder.ToString();
        }

        private static string EscapeCell(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }
            return value.Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ").Trim();
        }
    }
}
