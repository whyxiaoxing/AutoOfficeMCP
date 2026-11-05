using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;

namespace Office.WordTools
{
    /// <summary>
    /// Word 工具参数读取与校验。所有失败都抛 OfficeComException(INVALID_PARAMS)，
    /// 消息里必须说明"收到了什么、支持什么"，让调用方可以自行修复。
    /// </summary>
    internal sealed class WordArguments
    {
        private readonly JObject _root;

        public WordArguments(JObject arguments)
        {
            _root = arguments == null ? new JObject() : arguments;
        }

        public JObject Root
        {
            get { return _root; }
        }

        public bool Has(string name)
        {
            JToken token = _root[name];
            return token != null && token.Type != JTokenType.Null;
        }

        public bool HasAny(params string[] names)
        {
            if (names == null)
            {
                return false;
            }
            foreach (string name in names)
            {
                if (Has(name))
                {
                    return true;
                }
            }
            return false;
        }

        public void RequireAny(params string[] names)
        {
            if (HasAny(names))
            {
                return;
            }
            throw Invalid("缺少可用参数。收到 keys=" + DescribeKeys() + "；支持 " + string.Join("/", names));
        }

        public string RequireString(string name)
        {
            string value = GetOptionalString(name, null);
            if (string.IsNullOrWhiteSpace(value))
            {
                throw Invalid("缺少参数 " + name + "（非空字符串）。收到 keys=" + DescribeKeys());
            }
            return value.Trim();
        }

        public string RequireText(string name)
        {
            string value = GetOptionalString(name, null);
            if (value == null || value.Length == 0)
            {
                throw Invalid("缺少参数 " + name + "（非空文本，允许换行）。收到 keys=" + DescribeKeys());
            }
            return value;
        }

        public string GetOptionalString(string name, string defaultValue)
        {
            JToken token = _root[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return defaultValue;
            }
            if (token.Type != JTokenType.String)
            {
                throw Invalid("参数 " + name + " 必须是字符串，实际为 " + token.Type + "（值为 " + PreviewToken(token) + "）");
            }
            return token.Value<string>();
        }

        public int? GetOptionalInt(string name)
        {
            JToken token = _root[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }
            try
            {
                if (token.Type == JTokenType.Integer || token.Type == JTokenType.Float)
                {
                    double raw = token.Value<double>();
                    int value = Convert.ToInt32(raw, CultureInfo.InvariantCulture);
                    if (Math.Abs(raw - value) > 0.0001)
                    {
                        throw Invalid("参数 " + name + " 必须是整数，实际为 " + raw.ToString(CultureInfo.InvariantCulture));
                    }
                    return value;
                }
            }
            catch (OfficeComException)
            {
                throw;
            }
            catch (Exception)
            {
                throw Invalid("参数 " + name + " 必须是整数，实际为 " + PreviewToken(token));
            }
            throw Invalid("参数 " + name + " 必须是整数，实际为 " + PreviewToken(token));
        }

        public int GetOptionalInt(string name, int defaultValue)
        {
            int? value = GetOptionalInt(name);
            return value.HasValue ? value.Value : defaultValue;
        }

        public int RequireInt(string name, int minValue, int maxValue)
        {
            int? value = GetOptionalInt(name);
            if (!value.HasValue)
            {
                throw Invalid("缺少参数 " + name + "（整数，范围 " + minValue + "-" + maxValue + "）");
            }
            if (value.Value < minValue || value.Value > maxValue)
            {
                throw Invalid("参数 " + name + " 超出范围：收到 " + value.Value +
                              "，允许 " + minValue + "-" + maxValue);
            }
            return value.Value;
        }

        public bool GetOptionalBool(string name, bool defaultValue)
        {
            JToken token = _root[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return defaultValue;
            }
            if (token.Type == JTokenType.Boolean)
            {
                return token.Value<bool>();
            }
            throw Invalid("参数 " + name + " 必须是布尔值 true/false，实际为 " + PreviewToken(token));
        }

        public double? GetOptionalDouble(string name)
        {
            JToken token = _root[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }
            if (token.Type == JTokenType.Integer || token.Type == JTokenType.Float)
            {
                return token.Value<double>();
            }
            throw Invalid("参数 " + name + " 必须是数字，实际为 " + PreviewToken(token));
        }

        public double RequireDouble(string name)
        {
            double? value = GetOptionalDouble(name);
            if (!value.HasValue)
            {
                throw Invalid("缺少参数 " + name + "（数字）");
            }
            return value.Value;
        }

        public JArray GetOptionalArray(string name)
        {
            JToken token = _root[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }
            JArray array = token as JArray;
            if (array == null)
            {
                throw Invalid("参数 " + name + " 必须是数组，实际为 " + token.Type);
            }
            return array;
        }

        public JArray RequireArray(string name)
        {
            JArray array = GetOptionalArray(name);
            if (array == null)
            {
                throw Invalid("缺少参数 " + name + "（数组）");
            }
            return array;
        }

        public JObject GetOptionalObject(string name)
        {
            JToken token = _root[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }
            JObject obj = token as JObject;
            if (obj == null)
            {
                throw Invalid("参数 " + name + " 必须是对象，实际为 " + token.Type);
            }
            return obj;
        }

        public List<int> GetOptionalIntList(string name)
        {
            JArray array = GetOptionalArray(name);
            if (array == null)
            {
                return null;
            }
            List<int> values = new List<int>();
            for (int index = 0; index < array.Count; index++)
            {
                JToken token = array[index];
                try
                {
                    values.Add(token.Value<int>());
                }
                catch (Exception)
                {
                    throw Invalid("参数 " + name + "[" + index + "] 必须是整数，实际为 " + PreviewToken(token));
                }
            }
            return values;
        }

        public List<string> GetOptionalStringList(string name)
        {
            JArray array = GetOptionalArray(name);
            if (array == null)
            {
                return null;
            }
            List<string> values = new List<string>();
            for (int index = 0; index < array.Count; index++)
            {
                JToken token = array[index];
                if (token == null || token.Type == JTokenType.Null)
                {
                    values.Add(string.Empty);
                    continue;
                }
                if (token.Type == JTokenType.String)
                {
                    values.Add(token.Value<string>());
                    continue;
                }
                values.Add(token.ToString());
            }
            return values;
        }

        public string DescribeKeys()
        {
            if (_root == null || _root.Count == 0)
            {
                return "空参数";
            }
            List<string> names = new List<string>();
            foreach (JProperty property in _root.Properties())
            {
                names.Add(property.Name);
            }
            return string.Join("/", names);
        }

        public static OfficeComException Invalid(string message)
        {
            return new OfficeComException(OfficeErrorCode.INVALID_PARAMS, message);
        }

        private static string PreviewToken(JToken token)
        {
            if (token == null)
            {
                return "null";
            }
            string text = token.ToString();
            if (text.Length > 80)
            {
                text = text.Substring(0, 80) + "...";
            }
            return text;
        }
    }
}
