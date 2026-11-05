using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;

namespace Xlsx
{
    /// <summary>
    /// MCP 工具参数读取器：统一处理大小写、params 嵌套、别名和基础类型转换。
    /// 所有方法只读参数，不修改调用方对象。
    /// </summary>
    internal static class ExcelArgumentReader
    {
        /// <summary>
        /// 将 {"params": {...}, "range": ...} 扁平化为一个 JObject。
        /// params 中的同名字段不覆盖顶层字段，符合"显式顶层优先"的直觉。
        /// </summary>
        public static JObject Normalize(JObject arguments)
        {
            JObject result = arguments == null ? new JObject() : (JObject)arguments.DeepClone();
            MergeNestedObject(result, "params");
            MergeNestedObject(result, "arguments");
            return result;
        }

        private static void MergeNestedObject(JObject root, string propertyName)
        {
            JObject nested = Find(root, propertyName) as JObject;
            if (nested == null)
            {
                return;
            }
            foreach (JProperty property in nested.Properties())
            {
                if (!Has(root, property.Name))
                {
                    root[property.Name] = property.Value.DeepClone();
                }
            }
        }

        public static bool Has(JObject arguments, string name)
        {
            return Find(arguments, name) != null;
        }

        public static JToken Find(JObject arguments, string name)
        {
            if (arguments == null || string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            JToken direct = arguments[name];
            if (direct != null)
            {
                return direct;
            }

            foreach (JProperty property in arguments.Properties())
            {
                if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return property.Value;
                }
            }
            return null;
        }

        public static JToken First(JObject arguments, params string[] names)
        {
            if (names == null)
            {
                return null;
            }
            foreach (string name in names)
            {
                JToken token = Find(arguments, name);
                if (token != null && token.Type != JTokenType.Null)
                {
                    return token;
                }
            }
            return null;
        }

        public static string GetString(JObject arguments, params string[] names)
        {
            JToken token = First(arguments, names);
            if (token == null || token.Type == JTokenType.Null)
            {
                return string.Empty;
            }
            if (token.Type == JTokenType.String)
            {
                return token.Value<string>() ?? string.Empty;
            }
            return token.ToString();
        }

        public static string GetTrimmedString(JObject arguments, params string[] names)
        {
            return GetString(arguments, names).Trim();
        }

        public static bool GetBool(JObject arguments, bool defaultValue, params string[] names)
        {
            JToken token = First(arguments, names);
            if (token == null || token.Type == JTokenType.Null)
            {
                return defaultValue;
            }
            if (token.Type == JTokenType.Boolean)
            {
                return token.Value<bool>();
            }

            string text = token.ToString().Trim();
            if (string.Equals(text, "true", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(text, "1", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(text, "yes", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(text, "是", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (string.Equals(text, "false", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(text, "0", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(text, "no", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(text, "否", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
            return defaultValue;
        }

        public static bool TryGetInt(JObject arguments, out int value, params string[] names)
        {
            value = 0;
            JToken token = First(arguments, names);
            if (token == null || token.Type == JTokenType.Null)
            {
                return false;
            }
            if (token.Type == JTokenType.Integer)
            {
                value = token.Value<int>();
                return true;
            }
            if (token.Type == JTokenType.Float)
            {
                double number = token.Value<double>();
                if (number >= int.MinValue && number <= int.MaxValue)
                {
                    value = (int)number;
                    return true;
                }
            }
            return int.TryParse(token.ToString().Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
        }

        public static int GetInt(JObject arguments, int defaultValue, params string[] names)
        {
            int value;
            return TryGetInt(arguments, out value, names) ? value : defaultValue;
        }

        public static bool TryGetDouble(JObject arguments, out double value, params string[] names)
        {
            value = 0;
            JToken token = First(arguments, names);
            if (token == null || token.Type == JTokenType.Null)
            {
                return false;
            }
            if (token.Type == JTokenType.Integer || token.Type == JTokenType.Float)
            {
                value = token.Value<double>();
                return true;
            }
            return double.TryParse(token.ToString().Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        public static JArray GetArray(JObject arguments, params string[] names)
        {
            JToken token = First(arguments, names);
            return token as JArray;
        }

        /// <summary>数组参数兼容单个字符串/数字写法，例如 rowFields:"地区"。</summary>
        public static JArray GetArrayOrSingle(JObject arguments, params string[] names)
        {
            JToken token = First(arguments, names);
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }
            JArray array = token as JArray;
            if (array != null)
            {
                return array;
            }
            JArray single = new JArray();
            single.Add(token.DeepClone());
            return single;
        }

        public static List<string> GetStringList(JObject arguments, params string[] names)
        {
            List<string> values = new List<string>();
            JToken token = First(arguments, names);
            if (token == null || token.Type == JTokenType.Null)
            {
                return values;
            }

            JArray array = token as JArray;
            if (array != null)
            {
                foreach (JToken item in array)
                {
                    if (item != null && item.Type != JTokenType.Null)
                    {
                        string text = item.ToString().Trim();
                        if (text.Length > 0)
                        {
                            values.Add(text);
                        }
                    }
                }
                return values;
            }

            string single = token.ToString().Trim();
            if (single.Length > 0)
            {
                values.Add(single);
            }
            return values;
        }

        public static int[] GetIntList(JObject arguments, params string[] names)
        {
            List<int> values = new List<int>();
            JToken token = First(arguments, names);
            if (token == null || token.Type == JTokenType.Null)
            {
                return values.ToArray();
            }

            JArray array = token as JArray;
            if (array != null)
            {
                foreach (JToken item in array)
                {
                    int parsed;
                    if (int.TryParse(item.ToString().Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed))
                    {
                        values.Add(parsed);
                    }
                }
                return values.ToArray();
            }

            string text = token.ToString().Trim();
            string[] parts = text.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts)
            {
                int parsed;
                if (int.TryParse(part.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed))
                {
                    values.Add(parsed);
                }
            }
            return values.ToArray();
        }

        /// <summary>读取范围地址，并把 targetSheet/sheetName 拼接到不含 ! 的地址前。</summary>
        public static string GetRangeAddress(JObject arguments, params string[] names)
        {
            string address = GetTrimmedString(arguments, names);
            string targetSheet = GetTrimmedString(arguments, "targetSheet", "sheetName");
            if (address.Length > 0 && targetSheet.Length > 0 && address.IndexOf('!') < 0)
            {
                return ExcelComSupport.QuoteSheetName(targetSheet) + "!" + address;
            }
            return address;
        }

        public static string GetSheetName(JObject arguments)
        {
            return GetTrimmedString(arguments, "sheetName", "targetSheet", "sheet");
        }

        public static JObject CloneObject(JToken token)
        {
            JObject obj = token as JObject;
            return obj == null ? null : (JObject)obj.DeepClone();
        }
    }
}
