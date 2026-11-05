using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace Office.WordTools
{
    /// <summary>word_beautify 的单个样式主题设置。</summary>
    internal sealed class WordThemeStyle
    {
        public string FontName { get; set; }

        public double? Size { get; set; }

        public int? Color { get; set; }

        public bool? Bold { get; set; }

        public bool? Italic { get; set; }

        public double? LineSpacing { get; set; }

        public bool HasAny
        {
            get
            {
                return !string.IsNullOrWhiteSpace(FontName) || Size.HasValue || Color.HasValue ||
                       Bold.HasValue || Italic.HasValue || LineSpacing.HasValue;
            }
        }
    }

    /// <summary>word_beautify 的 theme 参数解析器。</summary>
    internal static class WordThemeSettings
    {
        public static Dictionary<string, WordThemeStyle> ParseTheme(JObject theme)
        {
            Dictionary<string, WordThemeStyle> result =
                new Dictionary<string, WordThemeStyle>(StringComparer.OrdinalIgnoreCase);
            if (theme == null)
            {
                return result;
            }

            string[] keys = new string[] { "h1", "h2", "h3", "body" };
            foreach (string key in keys)
            {
                JObject styleObject = theme[key] as JObject;
                if (styleObject == null)
                {
                    continue;
                }
                WordThemeStyle style = new WordThemeStyle();
                JToken fontToken = styleObject["font"];
                if (fontToken != null && fontToken.Type != JTokenType.Null)
                {
                    style.FontName = fontToken.ToString();
                }
                style.Size = ReadDouble(styleObject, "size");
                JToken colorToken = styleObject["color"];
                if (colorToken != null && colorToken.Type != JTokenType.Null)
                {
                    style.Color = WordFormatUtil.ParseColor(colorToken.ToString());
                }
                style.Bold = ReadBool(styleObject, "bold");
                style.Italic = ReadBool(styleObject, "italic");
                style.LineSpacing = ReadDouble(styleObject, "lineSpacing");
                if (style.HasAny)
                {
                    result[key] = style;
                }
            }
            return result;
        }

        public static double? ReadMargin(JObject margins, string name, string unit)
        {
            if (margins == null || margins[name] == null || margins[name].Type == JTokenType.Null)
            {
                return null;
            }
            double value;
            try
            {
                value = margins[name].Value<double>();
            }
            catch (Exception)
            {
                throw WordArguments.Invalid("margins." + name + " 必须是数字。");
            }
            if (value < 0)
            {
                throw WordArguments.Invalid("margins." + name + " 不能为负数。");
            }
            if (string.Equals(unit, "pt", StringComparison.OrdinalIgnoreCase))
            {
                return value;
            }
            return value * WordFormatUtil.PointsPerCentimeter;
        }

        private static double? ReadDouble(JObject root, string name)
        {
            JToken token = root[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }
            try
            {
                return token.Value<double>();
            }
            catch (Exception)
            {
                throw WordArguments.Invalid("theme." + name + " 必须是数字。");
            }
        }

        private static bool? ReadBool(JObject root, string name)
        {
            JToken token = root[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }
            if (token.Type == JTokenType.Boolean)
            {
                return token.Value<bool>();
            }
            throw WordArguments.Invalid("theme." + name + " 必须是布尔值。");
        }
    }
}
