using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Office;

namespace PowerPointe
{
    /// <summary>
    /// PowerPoint 工具参数读取器。只做类型归一与必填校验，不执行 COM。
    /// 所有失败都抛 OfficeComException(INVALID_PARAMS)，由执行器统一转 ToolResult。
    /// </summary>
    internal static class PowerPointArgumentReader
    {
        public static string GetString(JObject arguments, params string[] names)
        {
            if (arguments == null || names == null)
            {
                return string.Empty;
            }
            foreach (string name in names)
            {
                JToken token = arguments[name];
                if (token == null || token.Type == JTokenType.Null)
                {
                    continue;
                }
                if (token.Type == JTokenType.String)
                {
                    string value = token.Value<string>();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        return value;
                    }
                }
                else
                {
                    string text = token.ToString();
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        return text;
                    }
                }
            }
            return string.Empty;
        }

        public static string GetTrimmedString(JObject arguments, params string[] names)
        {
            return GetString(arguments, names).Trim();
        }

        public static string RequireString(JObject arguments, string name, params string[] aliases)
        {
            List<string> names = new List<string>();
            names.Add(name);
            if (aliases != null)
            {
                names.AddRange(aliases);
            }
            string value = GetTrimmedString(arguments, names.ToArray());
            if (value.Length == 0)
            {
                throw Invalid("缺少必填参数 " + name + "。");
            }
            return value;
        }

        public static int GetInt(JObject arguments, int defaultValue, params string[] names)
        {
            if (arguments == null || names == null)
            {
                return defaultValue;
            }
            foreach (string name in names)
            {
                JToken token = arguments[name];
                if (token == null || token.Type == JTokenType.Null)
                {
                    continue;
                }
                try
                {
                    return token.Value<int>();
                }
                catch
                {
                    throw Invalid("参数 " + name + " 必须是整数。");
                }
            }
            return defaultValue;
        }

        public static float GetFloat(JObject arguments, float defaultValue, params string[] names)
        {
            if (arguments == null || names == null)
            {
                return defaultValue;
            }
            foreach (string name in names)
            {
                JToken token = arguments[name];
                if (token == null || token.Type == JTokenType.Null)
                {
                    continue;
                }
                try
                {
                    return token.Value<float>();
                }
                catch
                {
                    throw Invalid("参数 " + name + " 必须是数字。");
                }
            }
            return defaultValue;
        }

        public static bool GetBool(JObject arguments, bool defaultValue, params string[] names)
        {
            if (arguments == null || names == null)
            {
                return defaultValue;
            }
            foreach (string name in names)
            {
                JToken token = arguments[name];
                if (token == null || token.Type == JTokenType.Null)
                {
                    continue;
                }
                try
                {
                    return token.Value<bool>();
                }
                catch
                {
                    throw Invalid("参数 " + name + " 必须是布尔值。");
                }
            }
            return defaultValue;
        }

        public static JArray GetArray(JObject arguments, params string[] names)
        {
            if (arguments == null || names == null)
            {
                return null;
            }
            foreach (string name in names)
            {
                JArray array = arguments[name] as JArray;
                if (array != null)
                {
                    return array;
                }
            }
            return null;
        }

        public static JObject GetObject(JObject arguments, params string[] names)
        {
            if (arguments == null || names == null)
            {
                return null;
            }
            foreach (string name in names)
            {
                JObject value = arguments[name] as JObject;
                if (value != null)
                {
                    return value;
                }
            }
            return null;
        }

        public static OfficeComException Invalid(string message)
        {
            return new OfficeComException(OfficeErrorCode.INVALID_PARAMS, message);
        }
    }
}
