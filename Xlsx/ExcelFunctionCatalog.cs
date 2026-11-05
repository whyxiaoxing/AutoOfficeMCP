using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using Newtonsoft.Json.Linq;
using Office;

namespace Xlsx
{
    /// <summary>
    /// Excel 函数静态目录检索。COM 无法枚举 Excel 内置函数，本目录随版本维护；
    /// 未列出的函数仍可通过 excel_apply_formula / excel_evaluate 直接调用。
    /// </summary>
    internal static class ExcelFunctionCatalog
    {
        private const string ResourceName = "Xlsx.Resources.excel_functions.json";
        private static readonly object SyncRoot = new object();
        private static JArray _catalog;

        public static OperationResult Query(JObject arguments)
        {
            JArray functions;
            try
            {
                functions = LoadCatalog();
            }
            catch (Exception ex)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "无法加载 Excel 函数目录资源：" + ex.Message);
            }

            string mode = ExcelArgumentReader.GetTrimmedString(arguments, "mode").ToLowerInvariant();
            if (mode.Length == 0)
            {
                mode = "catalog";
            }
            if (mode == "fallback" || mode == "fallbacks")
            {
                return ExcelFunctionMetadata.QueryFallback(arguments);
            }
            if (mode == "template" || mode == "templates")
            {
                return ExcelFunctionMetadata.QueryTemplates(arguments);
            }
            if (mode != "catalog" && mode != "list")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "mode 仅支持 catalog/fallback/templates。");
            }

            string keyword = ExcelArgumentReader.GetTrimmedString(arguments, "keyword", "query", "search", "q");
            string category = ExcelArgumentReader.GetTrimmedString(arguments, "category");
            string version = ExcelArgumentReader.GetTrimmedString(arguments, "version", "excelVersion");
            bool includeTemplates = ExcelArgumentReader.GetBool(arguments, false, "includeTemplates");
            int limit = ExcelArgumentReader.GetInt(arguments, 50, "limit");
            limit = Math.Max(1, Math.Min(limit, 500));
            string[] terms = keyword.Split(
                new char[] { ' ', '\t', '\r', '\n', ',', '，' },
                StringSplitOptions.RemoveEmptyEntries);

            JArray matches = new JArray();
            int totalMatches = 0;
            foreach (JToken item in functions)
            {
                JObject function = item as JObject;
                if (function == null)
                {
                    continue;
                }
                if (category.Length > 0)
                {
                    string functionCategory = function.Value<string>("category") ?? string.Empty;
                    if (functionCategory.IndexOf(category, StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        continue;
                    }
                }
                if (version.Length > 0 &&
                    !ExcelFunctionMetadata.IsAvailableInVersion(function, version))
                {
                    continue;
                }
                if (!Matches(function, terms))
                {
                    continue;
                }

                totalMatches++;
                if (matches.Count < limit)
                {
                    matches.Add(ExcelFunctionMetadata.Enrich(function));
                }
            }

            JObject data = new JObject();
            data["mode"] = "catalog";
            data["keyword"] = keyword;
            data["category"] = category;
            data["version"] = version;
            data["totalFunctions"] = functions.Count;
            data["totalMatches"] = totalMatches;
            data["returned"] = matches.Count;
            data["truncated"] = totalMatches > matches.Count;
            data["functions"] = matches;
            data["note"] = "Excel COM 不提供内置函数枚举接口，本目录为静态维护；未列出的函数仍可使用 excel_apply_formula 或 excel_evaluate 调用。minVersion/availability 为静态兼容性标注。";
            if (includeTemplates)
            {
                OperationResult templateResult = ExcelFunctionMetadata.QueryTemplates(arguments);
                data["templates"] = templateResult.Data == null
                    ? new JArray()
                    : (templateResult.Data["templates"] == null
                        ? new JArray()
                        : templateResult.Data["templates"].DeepClone());
            }

            string message = keyword.Length == 0
                ? "函数目录共 " + functions.Count + " 条，返回 " + matches.Count + " 条。"
                : "函数目录匹配 " + totalMatches + " 条，返回 " + matches.Count + " 条。";
            return OperationResult.Succeed(message, data);
        }

        private static bool Matches(JObject function, string[] terms)
        {
            if (terms == null || terms.Length == 0)
            {
                return true;
            }
            string haystack = string.Join(
                " ",
                new string[]
                {
                    function.Value<string>("name") ?? string.Empty,
                    function.Value<string>("chineseName") ?? string.Empty,
                    function.Value<string>("category") ?? string.Empty,
                    function.Value<string>("signature") ?? string.Empty,
                    function.Value<string>("description") ?? string.Empty,
                    function.Value<string>("example") ?? string.Empty,
                    function["keywords"] == null ? string.Empty : function["keywords"].ToString()
                });
            foreach (string term in terms)
            {
                if (haystack.IndexOf(term, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    return false;
                }
            }
            return true;
        }

        private static JArray LoadCatalog()
        {
            lock (SyncRoot)
            {
                if (_catalog != null)
                {
                    return _catalog;
                }

                string text = null;
                Assembly assembly = Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream(ResourceName))
                {
                    if (stream != null)
                    {
                        using (StreamReader reader = new StreamReader(stream, Encoding.UTF8, true))
                        {
                            text = reader.ReadToEnd();
                        }
                    }
                }

                if (string.IsNullOrWhiteSpace(text))
                {
                    string[] candidates = new string[]
                    {
                        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Tools", "excel", "resources", "excel_functions.json"),
                        Path.Combine(
                            Path.GetDirectoryName(assembly.Location) ?? AppDomain.CurrentDomain.BaseDirectory,
                            "Tools", "excel", "resources", "excel_functions.json")
                    };
                    foreach (string candidate in candidates)
                    {
                        if (File.Exists(candidate))
                        {
                            text = File.ReadAllText(candidate, Encoding.UTF8);
                            break;
                        }
                    }
                }

                if (string.IsNullOrWhiteSpace(text))
                {
                    throw new FileNotFoundException("未找到 Excel 函数目录资源 excel_functions.json。");
                }

                JObject root = JObject.Parse(text);
                JArray functions = root["functions"] as JArray;
                if (functions == null)
                {
                    throw new InvalidDataException("excel_functions.json 缺少 functions 数组。");
                }
                _catalog = functions;
                return _catalog;
            }
        }
    }
}
