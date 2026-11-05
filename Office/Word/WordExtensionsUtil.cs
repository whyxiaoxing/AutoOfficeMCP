using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Office.Config;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    /// <summary>
    /// T06 Word 扩展工具的共享辅助：段落范围解析、段落文本读写、手工标题识别、输出路径守卫等。
    /// 只放与具体工具无关的通用能力，避免把业务逻辑散落到各处理器。
    /// </summary>
    internal static class WordExtensionsUtil
    {
        private static readonly Regex WhitespaceRegex = new Regex(@"\s+", RegexOptions.Compiled);
        private static readonly Regex HeadingStyleRegex =
            new Regex(@"^(?:标题|Heading)\s*([1-9])$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ChapterHeadingRegex =
            new Regex(@"^\s*第[一二三四五六七八九十百千零〇0-9]+[章节篇部]\s*", RegexOptions.Compiled);
        private static readonly Regex ChineseNumberHeadingRegex =
            new Regex(@"^\s*[一二三四五六七八九十]+[、.．]\s*", RegexOptions.Compiled);
        private static readonly Regex ParenthesizedChineseRegex =
            new Regex(@"^\s*[（(][一二三四五六七八九十百零〇]+[）)]\s*", RegexOptions.Compiled);
        private static readonly Regex ParenthesizedArabicRegex =
            new Regex(@"^\s*[（(][0-9]{1,2}[）)]\s*", RegexOptions.Compiled);
        private static readonly Regex MultiLevelNumberRegex =
            new Regex(@"^\s*([0-9]+(?:[.．][0-9]+){1,8})[.．、]?\s+(?=\S)", RegexOptions.Compiled);
        private static readonly Regex SingleNumberPunctuationRegex =
            new Regex(@"^\s*([0-9]{1,2})[.．、]\s*(?=\S)", RegexOptions.Compiled);
        private static readonly Regex SingleNumberSpaceRegex =
            new Regex(@"^\s*([0-9]{1,2})\s+(?=\S)", RegexOptions.Compiled);
        private static readonly Regex MeasureWordRegex =
            new Regex(@"^(?:个|种|项|次|条|年|月|日|元|%|％|人|页|章|节|部分|方面)",
                RegexOptions.Compiled);

        /// <summary>
        /// 解析段落范围：paragraphIndexes 优先，其次 range=all/selection。
        /// 返回的段号均为 1 基、去重，并已用 Word 实际段数校验；失败消息包含实际段数。
        /// </summary>
        public static List<int> ResolveParagraphIndexes(
            Word.Application application,
            Word.Document document,
            WordArguments args,
            out string scopeLabel)
        {
            int total = WordComUtil.GetParagraphCount(document);
            List<int> explicitIndexes = args.GetOptionalIntList("paragraphIndexes");
            if (explicitIndexes != null && explicitIndexes.Count > 0)
            {
                List<int> selected = new List<int>();
                HashSet<int> seen = new HashSet<int>();
                foreach (int index in explicitIndexes)
                {
                    WordComUtil.ValidateParagraphIndex(document, index);
                    if (seen.Add(index))
                    {
                        selected.Add(index);
                    }
                }
                scopeLabel = "paragraphIndexes（" + selected.Count + " 段）";
                return selected;
            }

            string range = (args.GetOptionalString("range", "all") ?? "all")
                .Trim().ToLowerInvariant();
            List<int> result = new List<int>();
            if (range == "all" || range == "document" || range == "全文")
            {
                for (int index = 1; index <= total; index++)
                {
                    result.Add(index);
                }
                scopeLabel = "全文（" + total + " 段）";
                return result;
            }
            if (range != "selection" && range != "选区")
            {
                throw WordArguments.Invalid(
                    "range 只支持 all/selection，收到「" + range +
                    "」；也可用 paragraphIndexes 指定 1 基段号数组。当前文档共 " + total + " 段。");
            }

            Word.Selection selection = null;
            Word.Range selectedRange = null;
            try
            {
                selection = application.Selection;
                if (selection == null)
                {
                    throw WordArguments.Invalid("当前没有有效 Word 选区，无法按 range=selection 处理。");
                }
                selectedRange = selection.Range;
                int selectionStart = selectedRange.Start;
                int selectionEnd = selectedRange.End;
                if (selectionStart == selectionEnd)
                {
                    int collapsedIndex = WordComUtil.FindParagraphIndexAt(document, selectionStart);
                    if (collapsedIndex > 0)
                    {
                        result.Add(collapsedIndex);
                    }
                    scopeLabel = "选区（光标所在第 " + collapsedIndex + " 段）";
                    return result;
                }

                Word.Paragraphs paragraphs = null;
                try
                {
                    paragraphs = document.Paragraphs;
                    int count = paragraphs == null ? 0 : paragraphs.Count;
                    for (int index = 1; index <= count; index++)
                    {
                        Word.Paragraph paragraph = null;
                        Word.Range paragraphRange = null;
                        try
                        {
                            paragraph = paragraphs[index];
                            paragraphRange = paragraph.Range;
                            if (paragraphRange.Start <= selectionEnd &&
                                paragraphRange.End >= selectionStart)
                            {
                                result.Add(index);
                            }
                        }
                        finally
                        {
                            ComHelper.Release(paragraphRange, paragraph);
                        }
                    }
                }
                finally
                {
                    ComHelper.Release(paragraphs);
                }
                scopeLabel = "选区（" + result.Count + " 段相交）";
                return result;
            }
            finally
            {
                ComHelper.Release(selectedRange, selection);
            }
        }

        /// <summary>读取清理后的段落文本（不含段落标记）。</summary>
        public static string ReadParagraphText(Word.Paragraph paragraph)
        {
            if (paragraph == null)
            {
                return string.Empty;
            }
            Word.Range range = null;
            try
            {
                range = paragraph.Range;
                return WordComUtil.CleanParagraphText(range.Text);
            }
            finally
            {
                ComHelper.Release(range);
            }
        }

        /// <summary>整段替换文本并保留段落标记；返回 Word 回读文本。</summary>
        public static string SetParagraphText(Word.Paragraph paragraph, string newText)
        {
            if (paragraph == null)
            {
                throw WordArguments.Invalid("段落对象为空，无法写入文本。");
            }
            Word.Range range = null;
            try
            {
                range = paragraph.Range;
                range.Text = (newText ?? string.Empty) + "\r";
                return WordComUtil.CleanParagraphText(range.Text);
            }
            finally
            {
                ComHelper.Release(range);
            }
        }

        /// <summary>
        /// 只对“段落正文”范围（排除结尾段落标记）应用段落样式。
        /// Word 的 Paragraph.Range.End 位于下一段起点，直接 set_Style 会把样式泄漏到下一段；
        /// 这里把 end 回退 1 个字符，只覆盖当前段正文。可为 0 长度段落工作。
        /// 返回应用后的段落样式名（通过完整段落范围回读）。
        /// </summary>
        public static string ApplyParagraphStyle(
            Word.Document document,
            Word.Paragraph paragraph,
            string styleName,
            bool resetFont,
            bool resetParagraphFormat)
        {
            if (document == null || paragraph == null)
            {
                return string.Empty;
            }
            Word.Range paragraphRange = null;
            Word.Range bodyRange = null;
            Word.Font font = null;
            Word.ParagraphFormat paragraphFormat = null;
            try
            {
                paragraphRange = paragraph.Range;
                int bodyEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
                bodyRange = WordComUtil.CreateRange(
                    document, paragraphRange.Start, bodyEnd);
                if (resetFont)
                {
                    font = bodyRange.Font;
                    font.Reset();
                }
                if (resetParagraphFormat)
                {
                    paragraphFormat = bodyRange.ParagraphFormat;
                    paragraphFormat.Reset();
                }
                if (!string.IsNullOrWhiteSpace(styleName))
                {
                    // 段落级 set_Style 精确作用于该 Paragraph；对 Range 设置样式在 Word 中
                    // 可能把样式泄漏到下一段，因此这里始终用 Paragraph.set_Style。
                    object styleObject = styleName;
                    paragraph.set_Style(ref styleObject);
                }
                return WordComUtil.ReadStyleName(paragraphRange);
            }
            finally
            {
                ComHelper.Release(paragraphFormat, font, bodyRange, paragraphRange);
            }
        }

        /// <summary>段落是否只含普通文字（跳过表格/图片/域/公式，避免破坏对象）。</summary>
        public static bool IsPlainTextParagraph(Word.Paragraph paragraph)
        {
            if (paragraph == null)
            {
                return false;
            }
            Word.Range range = null;
            try
            {
                range = paragraph.Range;
                return WordComUtil.IsPlainTextRange(range);
            }
            catch
            {
                return false;
            }
            finally
            {
                ComHelper.Release(range);
            }
        }

        /// <summary>返回 0 表示不是标题样式；否则返回 1-9 级。</summary>
        public static int TryGetHeadingLevel(string styleName)
        {
            if (string.IsNullOrWhiteSpace(styleName))
            {
                return 0;
            }
            Match match = HeadingStyleRegex.Match(styleName.Trim());
            if (!match.Success)
            {
                return 0;
            }
            int level;
            return int.TryParse(match.Groups[1].Value, NumberStyles.Integer,
                CultureInfo.InvariantCulture, out level) ? level : 0;
        }

        /// <summary>按中文/英文候选名解析标题样式；找不到返回 null（不抛出）。</summary>
        public static string ResolveHeadingStyleName(Word.Document document, int level)
        {
            if (document == null || level < 1 || level > 9)
            {
                return null;
            }
            string number = level.ToString(CultureInfo.InvariantCulture);
            string[] candidates = new string[]
            {
                "标题 " + number,
                "Heading " + number,
                "标题" + number,
                "Heading" + number
            };
            foreach (string candidate in candidates)
            {
                string resolved;
                if (WordStyleCatalog.TryGetStyleName(document, candidate, out resolved))
                {
                    return resolved;
                }
            }
            return null;
        }

        public static bool IsBodyLikeStyle(string styleName)
        {
            if (string.IsNullOrWhiteSpace(styleName))
            {
                return true;
            }
            string key = styleName.Replace(" ", string.Empty).Trim().ToLowerInvariant();
            return key == "正文" || key == "normal" || key == "bodytext" ||
                   key == "默认段落字体" || key == "plaintext" || key == "body";
        }

        public static bool IsCaptionOrTocStyle(string styleName)
        {
            if (string.IsNullOrWhiteSpace(styleName))
            {
                return false;
            }
            string key = styleName.ToLowerInvariant();
            return key.IndexOf("题注", StringComparison.Ordinal) >= 0 ||
                   key.IndexOf("caption", StringComparison.Ordinal) >= 0 ||
                   key.IndexOf("目录", StringComparison.Ordinal) >= 0 ||
                   key.IndexOf("toc", StringComparison.Ordinal) >= 0;
        }

        /// <summary>
        /// 识别“手工编号 + 标题文字”段落。成功时给出建议标题级别、匹配前缀长度与原文前缀。
        /// 识别刻意保守：数字后必须有分隔符或空白，且排除图/表题注与常见量词开头。
        /// </summary>
        public static bool TryMatchManualHeading(
            string text,
            out int level,
            out int prefixLength,
            out string prefixText)
        {
            level = 0;
            prefixLength = 0;
            prefixText = string.Empty;
            if (string.IsNullOrWhiteSpace(text) || text.Length > 200)
            {
                return false;
            }
            string trimmed = text.TrimStart();
            if (trimmed.StartsWith("图", StringComparison.Ordinal) ||
                trimmed.StartsWith("表", StringComparison.Ordinal) ||
                trimmed.StartsWith("Figure", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("Table", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            Match match = ChapterHeadingRegex.Match(text);
            if (match.Success)
            {
                return AcceptManualHeading(text, match, 1, out level, out prefixLength, out prefixText);
            }

            match = MultiLevelNumberRegex.Match(text);
            if (match.Success)
            {
                string numberText = match.Groups[1].Value.Replace('．', '.');
                int components = numberText.Split('.').Length;
                int detectedLevel = Math.Max(1, Math.Min(9, components));
                return AcceptManualHeading(
                    text, match, detectedLevel, out level, out prefixLength, out prefixText);
            }

            match = ParenthesizedChineseRegex.Match(text);
            if (match.Success)
            {
                return AcceptManualHeading(text, match, 2, out level, out prefixLength, out prefixText);
            }

            match = ParenthesizedArabicRegex.Match(text);
            if (match.Success)
            {
                return AcceptManualHeading(text, match, 2, out level, out prefixLength, out prefixText);
            }

            match = ChineseNumberHeadingRegex.Match(text);
            if (match.Success)
            {
                return AcceptManualHeading(text, match, 1, out level, out prefixLength, out prefixText);
            }

            match = SingleNumberPunctuationRegex.Match(text);
            if (match.Success)
            {
                return AcceptManualHeading(text, match, 1, out level, out prefixLength, out prefixText);
            }

            match = SingleNumberSpaceRegex.Match(text);
            if (match.Success)
            {
                string remainder = text.Substring(match.Length).TrimStart();
                if (remainder.Length > 0 && !MeasureWordRegex.IsMatch(remainder))
                {
                    return AcceptManualHeading(
                        text, match, 1, out level, out prefixLength, out prefixText);
                }
            }
            return false;
        }

        private static bool AcceptManualHeading(
            string text,
            Match match,
            int detectedLevel,
            out int level,
            out int prefixLength,
            out string prefixText)
        {
            level = 0;
            prefixLength = 0;
            prefixText = string.Empty;
            if (match == null || !match.Success || match.Length <= 0)
            {
                return false;
            }
            string remainder = text.Substring(match.Length).TrimStart();
            if (remainder.Length == 0)
            {
                return false;
            }
            level = detectedLevel;
            prefixLength = match.Length;
            prefixText = match.Value;
            return true;
        }

        public static string BuildNumberFormat(int level)
        {
            StringBuilder builder = new StringBuilder();
            for (int index = 1; index <= level; index++)
            {
                builder.Append("%");
                builder.Append(index.ToString(CultureInfo.InvariantCulture));
                if (index < level)
                {
                    builder.Append(".");
                }
            }
            builder.Append(".");
            return builder.ToString();
        }

        public static bool IsCjkChar(char value)
        {
            return (value >= '㐀' && value <= '䶿') ||
                   (value >= '一' && value <= '鿿') ||
                   (value >= '豈' && value <= '﫿') ||
                   (value >= '぀' && value <= 'ヿ') ||
                   value == '々' || value == '〇';
        }

        public static bool IsAsciiLetterOrDigit(char value)
        {
            return (value >= 'a' && value <= 'z') ||
                   (value >= 'A' && value <= 'Z') ||
                   (value >= '0' && value <= '9');
        }

        public static bool IsAsciiLetter(char value)
        {
            return (value >= 'a' && value <= 'z') || (value >= 'A' && value <= 'Z');
        }

        public static bool IsAsciiDigit(char value)
        {
            return value >= '0' && value <= '9';
        }

        public static bool IsAsciiPunctuation(char value)
        {
            return (value >= '!' && value <= '/') ||
                   (value >= ':' && value <= '@') ||
                   (value >= '[' && value <= '`') ||
                   (value >= '{' && value <= '~');
        }

        public static bool ContainsCjk(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }
            foreach (char value in text)
            {
                if (IsCjkChar(value))
                {
                    return true;
                }
            }
            return false;
        }

        public static char? FindNeighborChar(string text, int index, int direction)
        {
            if (text == null || index < 0 || index >= text.Length)
            {
                return null;
            }
            int cursor = index + direction;
            while (cursor >= 0 && cursor < text.Length)
            {
                char value = text[cursor];
                if (value == ' ' || value == '\t' || value == '　')
                {
                    cursor += direction;
                    continue;
                }
                return value;
            }
            return null;
        }

        public static string NormalizeWhitespace(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }
            return WhitespaceRegex.Replace(text, " ").Trim();
        }

        public static string Preview(string text, int maxLength)
        {
            return WordComUtil.NormalizePreview(text, maxLength);
        }

        public static void AddWarning(JArray warnings, string message)
        {
            if (warnings != null && !string.IsNullOrWhiteSpace(message))
            {
                warnings.Add(message.Trim());
            }
        }

        /// <summary>把输出文件路径规范化并做覆盖守卫；返回 null 表示可以写入。</summary>
        public static OperationResult PrepareOutputFile(
            string requestedPath,
            string defaultExtension,
            bool confirm,
            out string path,
            out bool existed)
        {
            path = string.Empty;
            existed = false;
            if (string.IsNullOrWhiteSpace(requestedPath))
            {
                throw WordArguments.Invalid("输出文件路径不能为空。");
            }
            string candidate = requestedPath.Trim();
            if (!string.IsNullOrWhiteSpace(defaultExtension) &&
                string.IsNullOrWhiteSpace(Path.GetExtension(candidate)))
            {
                candidate = candidate + defaultExtension;
            }
            string pathError;
            if (!WorkspacePathResolver.TryResolve(candidate, out path, out pathError))
            {
                throw WordArguments.Invalid("输出文件路径无效：" + requestedPath + "；" + pathError);
            }

            if (Directory.Exists(path))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "输出路径是目录而不是文件：" + path + "。请给出带文件名的完整路径。");
            }
            string directory = Path.GetDirectoryName(path);
            if (string.IsNullOrWhiteSpace(directory))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "输出路径缺少有效目录：" + path + "。");
            }
            string directoryError;
            if (!WorkspacePathResolver.TryEnsureOutputDirectory(directory, out directoryError))
            {
                return OperationResult.Failed(OfficeErrorCode.READ_ONLY, directoryError);
            }
            existed = File.Exists(path);
            if (existed && !confirm)
            {
                return OperationResult.Cancelled(
                    "输出文件已存在：" + path +
                    "。如需覆盖，请显式传 confirm=true；本次未写入任何文件。");
            }
            return null;
        }

        /// <summary>创建或校验输出目录；返回 null 表示可以写入。</summary>
        public static OperationResult PrepareOutputDirectory(string requestedPath, out string fullPath)
        {
            fullPath = string.Empty;
            if (string.IsNullOrWhiteSpace(requestedPath))
            {
                throw WordArguments.Invalid("输出目录不能为空。");
            }
            try
            {
                string resolveError;
                if (!WorkspacePathResolver.TryResolve(requestedPath, out fullPath, out resolveError))
                {
                    return OperationResult.Failed(OfficeErrorCode.INVALID_PARAMS, resolveError);
                }
                if (File.Exists(fullPath))
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.INVALID_PARAMS,
                        "输出目录路径被同名文件占用：" + fullPath + "。请改用目录路径。");
                }
                string directoryError;
                if (!WorkspacePathResolver.TryEnsureOutputDirectory(fullPath, out directoryError))
                {
                    return OperationResult.Failed(OfficeErrorCode.READ_ONLY, directoryError);
                }
                return null;
            }
            catch (Exception ex)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.READ_ONLY,
                    "无法创建或访问输出目录：" + requestedPath + "；" + ex.Message);
            }
        }

        public static void WriteUtf8NoBom(string path, string content)
        {
            File.WriteAllText(path, content ?? string.Empty, new UTF8Encoding(false));
        }

        /// <summary>表格/文本转换分隔符：tab/comma/paragraph/space 或任意自定义字符串。</summary>
        public static string ParseSeparator(string raw)
        {
            string value = string.IsNullOrWhiteSpace(raw) ? "tab" : raw.Trim();
            switch (value.ToLowerInvariant())
            {
                case "tab":
                case "\\t":
                case "制表符":
                    return "\t";
                case "comma":
                case ",":
                case "逗号":
                    return ",";
                case "paragraph":
                case "newline":
                case "\\r":
                case "\\n":
                case "段落":
                case "换行":
                    return "\r";
                case "space":
                case "空格":
                    return " ";
                default:
                    return value;
            }
        }

        public static string SafeBookmarkName(string raw)
        {
            StringBuilder builder = new StringBuilder();
            if (!string.IsNullOrEmpty(raw))
            {
                foreach (char value in raw)
                {
                    if (char.IsLetterOrDigit(value) || value == '_')
                    {
                        builder.Append(value);
                    }
                    else
                    {
                        builder.Append('_');
                    }
                }
            }
            string result = builder.ToString();
            if (result.Length == 0)
            {
                result = "Bookmark";
            }
            if (char.IsDigit(result[0]))
            {
                result = "B_" + result;
            }
            if (result.Length > 35)
            {
                result = result.Substring(0, 35);
            }
            return result;
        }

        public static string GetListIndent(int listLevel)
        {
            int safeLevel = Math.Max(1, Math.Min(9, listLevel));
            return new string(' ', (safeLevel - 1) * 2);
        }

        /// <summary>
        /// InlineShape.AlternativeText 在部分 Word 状态/图片上会抛 E_FAIL；
        /// Alt 文本只是元数据，读取失败时返回空串，不阻断图片导出/题注。
        /// </summary>
        public static string SafeReadAlternativeText(Word.InlineShape shape)
        {
            if (shape == null)
            {
                return string.Empty;
            }
            try
            {
                return shape.AlternativeText ?? string.Empty;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[word-ext] 读取 InlineShape.AlternativeText 失败：" + ex.Message);
                return string.Empty;
            }
        }

        /// <summary>浮动 Shape.AlternativeText 的安全读取，语义同上。</summary>
        public static string SafeReadAlternativeText(Word.Shape shape)
        {
            if (shape == null)
            {
                return string.Empty;
            }
            try
            {
                return shape.AlternativeText ?? string.Empty;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[word-ext] 读取 Shape.AlternativeText 失败：" + ex.Message);
                return string.Empty;
            }
        }
    }
}
