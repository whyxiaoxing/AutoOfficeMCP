using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    /// <summary>
    /// Word 内置样式名映射。中文 Word 的本地化名称是「标题 1」，英文是「Heading 1」；
    /// 直接拿 Heading1 去 doc.Styles[...] 会找不到样式，因此必须先走候选名解析。
    /// </summary>
    internal static class WordStyleCatalog
    {
        private static readonly Dictionary<string, string[]> StyleMap = BuildStyleMap();

        public static string ResolveStyleName(Word.Document document, string requested)
        {
            if (document == null)
            {
                throw new OfficeComException(OfficeErrorCode.DOC_NOT_OPEN, "当前没有活动文档，无法解析样式。");
            }
            if (string.IsNullOrWhiteSpace(requested))
            {
                throw WordArguments.Invalid("styleName 不能为空。支持「标题1」「标题 1」「Heading1」「正文」「Normal」等。");
            }

            string trimmed = requested.Trim();
            string resolved;
            if (TryGetStyleName(document, trimmed, out resolved))
            {
                return resolved;
            }

            string key = NormalizeKey(trimmed);
            string[] candidates;
            if (StyleMap.TryGetValue(key, out candidates))
            {
                foreach (string candidate in candidates)
                {
                    if (TryGetStyleName(document, candidate, out resolved))
                    {
                        return resolved;
                    }
                }
            }

            string available = DescribeAvailableStyles(document, 20);
            throw WordArguments.Invalid(
                "样式「" + trimmed + "」不存在。已尝试内置映射（标题1/标题 1/Heading1→标题 1 等）。" +
                "当前文档可用样式示例：" + available +
                "。请先用「标题1」或「Heading1」等标准名称重试。");
        }

        public static bool TryGetStyleName(Word.Document document, string styleName, out string resolvedName)
        {
            resolvedName = string.Empty;
            if (document == null || string.IsNullOrWhiteSpace(styleName))
            {
                return false;
            }

            Word.Styles styles = null;
            object styleObject = null;
            Word.Style style = null;
            try
            {
                styles = document.Styles;
                styleObject = styles[styleName];
                if (styleObject == null)
                {
                    return false;
                }
                style = styleObject as Word.Style;
                resolvedName = style != null ? (style.NameLocal ?? styleName) : styleName;
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                ComHelper.Release(style, styleObject, styles);
            }
        }

        public static string DescribeAvailableStyles(Word.Document document, int maxCount)
        {
            if (document == null)
            {
                return "无活动文档";
            }

            Word.Styles styles = null;
            List<string> names = new List<string>();
            try
            {
                styles = document.Styles;
                int count = styles == null ? 0 : styles.Count;
                int limit = Math.Min(count, Math.Max(1, maxCount));
                for (int index = 1; index <= limit; index++)
                {
                    Word.Style style = null;
                    try
                    {
                        style = styles[index];
                        if (style != null && !string.IsNullOrWhiteSpace(style.NameLocal))
                        {
                            names.Add(style.NameLocal);
                        }
                    }
                    finally
                    {
                        ComHelper.Release(style);
                    }
                }
            }
            catch (Exception ex)
            {
                names.Add("（枚举样式失败：" + ex.Message + "）");
            }
            finally
            {
                ComHelper.Release(styles);
            }
            return names.Count == 0 ? "无" : string.Join("、", names.ToArray());
        }

        private static Dictionary<string, string[]> BuildStyleMap()
        {
            Dictionary<string, string[]> map =
                new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);

            map["标题"] = new string[] { "标题 1", "Heading 1", "标题" };
            map["title"] = new string[] { "标题", "Title" };
            map["副标题"] = new string[] { "副标题", "Subtitle" };
            map["subtitle"] = new string[] { "副标题", "Subtitle" };
            map["正文"] = new string[] { "正文", "Normal" };
            map["normal"] = new string[] { "Normal", "正文" };
            map["题注"] = new string[] { "题注", "Caption" };
            map["caption"] = new string[] { "题注", "Caption" };
            map["引用"] = new string[] { "引用", "Quote" };
            map["quote"] = new string[] { "引用", "Quote" };

            for (int level = 1; level <= 9; level++)
            {
                string number = level.ToString(CultureInfo.InvariantCulture);
                string[] candidates = new string[]
                {
                    "标题 " + number,
                    "Heading " + number,
                    "标题" + number,
                    "Heading" + number
                };
                map[NormalizeKey("标题" + number)] = candidates;
                map[NormalizeKey("标题 " + number)] = candidates;
                map[NormalizeKey("heading" + number)] = candidates;
                map[NormalizeKey("heading " + number)] = candidates;
            }
            return map;
        }

        private static string NormalizeKey(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }
            StringBuilder builder = new StringBuilder(value.Length);
            foreach (char character in value)
            {
                if (!char.IsWhiteSpace(character))
                {
                    builder.Append(char.ToLowerInvariant(character));
                }
            }
            return builder.ToString();
        }
    }
}
