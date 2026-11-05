using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Newtonsoft.Json.Linq;
using Office.Config;

namespace AutoOfficeMCP.Tools
{
    /// <summary>
    /// 启动时扫描 exe 同目录 Tools\**\*.json，生成 MCP tools/list 并支持按名查找。
    /// config.json 的 tools.disabled 名单在这里统一过滤。
    /// </summary>
    public sealed class ToolRegistry
    {
        private readonly List<ToolDescriptor> _tools;
        private readonly Dictionary<string, ToolDescriptor> _byName;

        private ToolRegistry(List<ToolDescriptor> tools, Dictionary<string, ToolDescriptor> byName)
        {
            _tools = tools;
            _byName = byName;
        }

        public IReadOnlyList<ToolDescriptor> Tools
        {
            get { return _tools; }
        }

        public static ToolRegistry Load()
        {
            string toolsRoot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Tools");
            List<ToolDescriptor> loaded = new List<ToolDescriptor>();
            Dictionary<string, ToolDescriptor> byName = new Dictionary<string, ToolDescriptor>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> disabled = AppConfig.Current.DisabledTools;

            if (!Directory.Exists(toolsRoot))
            {
                Console.Error.WriteLine("工具描述符目录不存在：" + toolsRoot);
                return new ToolRegistry(loaded, byName);
            }

            string[] files;
            try
            {
                files = Directory.GetFiles(toolsRoot, "*.json", SearchOption.AllDirectories);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("扫描工具描述符失败：" + toolsRoot +
                                        Environment.NewLine + ex);
                return new ToolRegistry(loaded, byName);
            }

            Array.Sort(files, StringComparer.OrdinalIgnoreCase);
            foreach (string file in files)
            {
                // Tools\**\resources\** 是执行器资源（如 Excel 函数目录），不是工具描述符。
                // 扫描时直接跳过，避免启动日志出现“无效描述符”噪音，也不影响工具总数。
                if (IsResourceJson(file, toolsRoot))
                {
                    continue;
                }

                try
                {
                    string text = File.ReadAllText(file, Encoding.UTF8);
                    JObject root = JObject.Parse(text);
                    string error;
                    ToolDescriptor descriptor = ToolDescriptor.FromJson(root, file, out error);
                    if (descriptor == null)
                    {
                        Console.Error.WriteLine("跳过无效工具描述符 " + file + "：" + error);
                        continue;
                    }

                    if (disabled.Contains(descriptor.ToolName))
                    {
                        continue;
                    }

                    if (byName.ContainsKey(descriptor.ToolName))
                    {
                        Console.Error.WriteLine("工具名重复，保留先加载项：" + descriptor.ToolName +
                                                "（" + file + "）");
                        continue;
                    }

                    loaded.Add(descriptor);
                    byName.Add(descriptor.ToolName, descriptor);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("加载工具描述符失败：" + file + Environment.NewLine + ex);
                }
            }

            loaded.Sort(delegate(ToolDescriptor left, ToolDescriptor right)
            {
                return string.Compare(left.ToolName, right.ToolName, StringComparison.OrdinalIgnoreCase);
            });

            return new ToolRegistry(loaded, byName);
        }

        /// <summary>判断文件是否位于 Tools 下的 resources 子目录（任意层级）。</summary>
        private static bool IsResourceJson(string file, string toolsRoot)
        {
            if (string.IsNullOrWhiteSpace(file) || string.IsNullOrWhiteSpace(toolsRoot))
            {
                return false;
            }
            string relative = file.Substring(toolsRoot.Length).TrimStart('\\', '/');
            string[] segments = relative.Split(new char[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries);
            // 最后一段是文件名；只检查目录段，避免误伤名为 resources.json 的工具描述符。
            for (int index = 0; index < segments.Length - 1; index++)
            {
                if (string.Equals(segments[index], "resources", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 生成按 appType 过滤的注册表视图（/xlsx、/word、/ppt 路由使用）。
        /// includeCommon=true 时一并暴露 common 工具（打开/保存/批处理等跨应用能力）。
        /// </summary>
        public ToolRegistry CreateScoped(string appType, bool includeCommon)
        {
            List<ToolDescriptor> scoped = new List<ToolDescriptor>();
            Dictionary<string, ToolDescriptor> byName =
                new Dictionary<string, ToolDescriptor>(StringComparer.OrdinalIgnoreCase);
            foreach (ToolDescriptor descriptor in _tools)
            {
                bool match = string.Equals(descriptor.AppType, appType, StringComparison.OrdinalIgnoreCase);
                if (!match && includeCommon)
                {
                    match = string.Equals(descriptor.AppType, "common", StringComparison.OrdinalIgnoreCase);
                }
                if (!match || byName.ContainsKey(descriptor.ToolName))
                {
                    continue;
                }
                scoped.Add(descriptor);
                byName.Add(descriptor.ToolName, descriptor);
            }
            return new ToolRegistry(scoped, byName);
        }

        /// <summary>按工具名查找；已禁用与不存在的工具统一返回 null。</summary>
        public ToolDescriptor Find(string toolName)
        {
            if (string.IsNullOrWhiteSpace(toolName))
            {
                return null;
            }

            ToolDescriptor descriptor;
            if (_byName.TryGetValue(toolName, out descriptor))
            {
                return descriptor;
            }
            return null;
        }

        /// <summary>tools/list 的 result 对象。</summary>
        public JObject CreateToolsListResult()
        {
            JArray tools = new JArray();
            foreach (ToolDescriptor descriptor in _tools)
            {
                tools.Add(descriptor.BuildToolListItem());
            }

            JObject result = new JObject();
            result["tools"] = tools;
            return result;
        }
    }
}
