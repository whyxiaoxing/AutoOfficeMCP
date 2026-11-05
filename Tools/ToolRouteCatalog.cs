using System;
using System.Collections.Generic;

namespace AutoOfficeMCP.Tools
{
    /// <summary>
    /// HTTP 工具路由定义：/xlsx → excel，/word → word，/ppt → ppt。
    /// 路由只负责筛选工具与生成说明，不复制执行器逻辑。
    /// </summary>
    public sealed class ToolRoute
    {
        public string Key { get; private set; }

        public string AppType { get; private set; }

        public string DisplayName { get; private set; }

        public ToolRoute(string key, string appType, string displayName)
        {
            Key = key;
            AppType = appType;
            DisplayName = displayName;
        }
    }

    public static class ToolRouteCatalog
    {
        private static readonly ToolRoute[] RoutesArray =
        {
            new ToolRoute("xlsx", "excel", "Excel"),
            new ToolRoute("word", "word", "Word"),
            new ToolRoute("ppt", "ppt", "PowerPoint")
        };

        public static IReadOnlyList<ToolRoute> Routes
        {
            get { return RoutesArray; }
        }

        public static ToolRoute Find(string key)
        {
            if (string.IsNullOrWhiteSpace(key))
            {
                return null;
            }
            foreach (ToolRoute route in RoutesArray)
            {
                if (string.Equals(route.Key, key.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    return route;
                }
            }
            return null;
        }
    }
}
