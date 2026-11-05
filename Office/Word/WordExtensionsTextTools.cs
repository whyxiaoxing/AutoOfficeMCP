using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    public sealed partial class WordExtensionsExecutor
    {
        private static readonly Regex CjkLetterBoundaryRegex = new Regex(
            @"(?<=[㐀-䶿一-鿿豈-﫿぀-ヿ])(?=[A-Za-z])|(?<=[A-Za-z])(?=[㐀-䶿一-鿿豈-﫿぀-ヿ])",
            RegexOptions.Compiled);

        private static readonly Regex CjkDigitBoundaryRegex = new Regex(
            @"(?<=[㐀-䶿一-鿿豈-﫿぀-ヿ])(?=[0-9])|(?<=[0-9])(?=[㐀-䶿一-鿿豈-﫿぀-ヿ])",
            RegexOptions.Compiled);

        private static readonly Regex SentenceSplitRegex = new Regex(
            @"[。！？!?；;\r\n]+|\.(?=\s|$)", RegexOptions.Compiled);

        private static readonly Regex EnglishWordRegex = new Regex(
            @"[A-Za-z]+(?:['\-][A-Za-z]+)*", RegexOptions.Compiled);

        private static readonly Dictionary<char, char> HalfToFullPunctuation = BuildHalfToFullPunctuation();
        private static readonly Dictionary<char, char> FullToHalfPunctuation = BuildFullToHalfPunctuation();
        private static readonly HashSet<char> AutoHalfPunctuation = new HashSet<char>(
            new char[] { ',', '.', ';', ':', '?', '!', '(', ')', '%' });
        private static readonly HashSet<char> AutoFullPunctuation = new HashSet<char>(
            new char[] { '，', '。', '；', '：', '？', '！', '（', '）', '％' });

        private void RegisterTextExtensionHandlers(
            Dictionary<string, Func<Word.Application, JObject, OperationResult>> handlers)
        {
            handlers[WordExtensionsToolNames.PunctuationNormalize] = HandlePunctuationNormalize;
            handlers[WordExtensionsToolNames.CjkSpacing] = HandleCjkSpacing;
            handlers[WordExtensionsToolNames.NormalizeHeadings] = HandleNormalizeHeadings;
            handlers[WordExtensionsToolNames.RepeatCheck] = HandleRepeatCheck;
            handlers[WordExtensionsToolNames.HiddenTextReport] = HandleHiddenTextReport;
            handlers[WordExtensionsToolNames.ReadingTime] = HandleReadingTime;
        }

        // ------------------------------------------------------------------
        // word_punctuation_normalize
        // ------------------------------------------------------------------

        private OperationResult HandlePunctuationNormalize(
            Word.Application application,
            JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string mode = (args.GetOptionalString("mode", "auto") ?? "auto")
                .Trim().ToLowerInvariant();
            if (mode == "全角")
            {
                mode = "fullwidth";
            }
            else if (mode == "半角")
            {
                mode = "halfwidth";
            }
            if (mode != "auto" && mode != "fullwidth" && mode != "halfwidth")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "mode 只支持 auto/fullwidth/halfwidth（全角/半角），收到「" + mode + "」。");
            }
            bool normalizeQuotes = args.GetOptionalBool("normalizeQuotes", false);
            int maxReport = args.GetOptionalInt("maxReport", 300);
            if (maxReport < 1 || maxReport > 2000)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "maxReport 允许 1-2000，收到 " + maxReport + "。");
            }

            return RunWriteTool(
                application,
                WordExtensionsToolNames.PunctuationNormalize,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.None("批量段落标点", "Word:Paragraphs");
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    List<int> indexes = WordExtensionsUtil.ResolveParagraphIndexes(
                        application, document, args, out string scopeLabel);
                    JArray items = new JArray();
                    JArray skipped = new JArray();
                    int totalReplacements = 0;
                    int changedParagraphs = 0;
                    int failedParagraphs = 0;
                    int reportedChanges = 0;
                    int scannedParagraphs = 0;

                    List<int> ordered = new List<int>(indexes);
                    ordered.Sort();
                    ordered.Reverse();

                    foreach (int paragraphIndex in ordered)
                    {
                        Word.Paragraph paragraph = null;
                        try
                        {
                            paragraph = WordComUtil.GetParagraph(document, paragraphIndex);
                            string original = WordExtensionsUtil.ReadParagraphText(paragraph);
                            if (original.Length == 0)
                            {
                                continue;
                            }
                            scannedParagraphs++;

                            List<PunctuationChange> changes = BuildPunctuationChanges(
                                original, mode, normalizeQuotes);
                            if (changes.Count == 0)
                            {
                                continue;
                            }
                            if (!WordExtensionsUtil.IsPlainTextParagraph(paragraph))
                            {
                                skipped.Add(new JObject
                                {
                                    { "paragraphIndex", paragraphIndex },
                                    { "reason", "段落含表格/图片/域/公式等非纯文本内容，已跳过以免破坏对象" }
                                });
                                continue;
                            }

                            string expected = ApplyPunctuationChanges(original, changes);
                            Word.Range paragraphRange = null;
                            int paragraphStart;
                            try
                            {
                                paragraphRange = paragraph.Range;
                                paragraphStart = paragraphRange.Start;
                            }
                            finally
                            {
                                ComHelper.Release(paragraphRange);
                            }

                            for (int index = changes.Count - 1; index >= 0; index--)
                            {
                                PunctuationChange change = changes[index];
                                Word.Range charRange = null;
                                try
                                {
                                    charRange = WordComUtil.CreateRange(
                                        document,
                                        paragraphStart + change.RelativeIndex,
                                        paragraphStart + change.RelativeIndex + 1);
                                    charRange.Text = change.Replacement.ToString();
                                }
                                finally
                                {
                                    ComHelper.Release(charRange);
                                }
                            }

                            string actual = WordExtensionsUtil.ReadParagraphText(paragraph);
                            totalReplacements += changes.Count;
                            if (!string.Equals(actual, expected, StringComparison.Ordinal))
                            {
                                failedParagraphs++;
                                items.Add(new JObject
                                {
                                    { "paragraphIndex", paragraphIndex },
                                    { "verified", false },
                                    { "beforePreview", WordExtensionsUtil.Preview(original, 120) },
                                    { "expectedPreview", WordExtensionsUtil.Preview(expected, 120) },
                                    { "actualPreview", WordExtensionsUtil.Preview(actual, 120) },
                                    { "changes", BuildPunctuationChangeArray(
                                        changes, original, maxReport, ref reportedChanges) }
                                });
                                continue;
                            }

                            changedParagraphs++;
                            items.Add(new JObject
                            {
                                { "paragraphIndex", paragraphIndex },
                                { "verified", true },
                                { "beforePreview", WordExtensionsUtil.Preview(original, 120) },
                                { "afterPreview", WordExtensionsUtil.Preview(actual, 120) },
                                { "changeCount", changes.Count },
                                { "changes", BuildPunctuationChangeArray(
                                    changes, original, maxReport, ref reportedChanges) }
                            });
                        }
                        catch (Exception ex)
                        {
                            failedParagraphs++;
                            skipped.Add(new JObject
                            {
                                { "paragraphIndex", paragraphIndex },
                                { "reason", OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) }
                            });
                        }
                        finally
                        {
                            ComHelper.Release(paragraph);
                        }
                    }

                    JObject data = new JObject();
                    data["scope"] = scopeLabel;
                    data["mode"] = mode;
                    data["normalizeQuotes"] = normalizeQuotes;
                    data["scannedParagraphs"] = scannedParagraphs;
                    data["changedParagraphs"] = changedParagraphs;
                    data["failedParagraphs"] = failedParagraphs;
                    data["totalReplacements"] = totalReplacements;
                    data["items"] = items;
                    data["skipped"] = skipped;
                    data["reportLimit"] = maxReport;
                    data["reportTruncated"] = reportedChanges < totalReplacements;
                    data["note"] = "仅做 1:1 标点替换，保留原有字符格式与段落标记；含图片/域/公式/表格的段落跳过。" +
                                   "normalizeQuotes=true 时按段落内的配对顺序转换直引号/弯引号，auto 模式不改引号。";

                    if (totalReplacements == 0)
                    {
                        return WordWriteResult.Ok(
                            "未发现需要规范化的中英文标点（范围：" + scopeLabel + "）；文档未修改。",
                            data,
                            false);
                    }
                    if (failedParagraphs > 0 && changedParagraphs == 0)
                    {
                        return WordWriteResult.Fail(
                            OfficeErrorCode.INTERNAL_ERROR,
                            "标点规范化后回读校验全部失败（" + failedParagraphs +
                            " 段），文档内容无法确认。请检查文档是否受保护/启用修订。",
                            data,
                            false);
                    }
                    if (failedParagraphs > 0)
                    {
                        return WordWriteResult.Fail(
                            OfficeErrorCode.PARTIAL_APPLY,
                            "标点规范化部分完成：成功 " + changedParagraphs + " 段，失败/跳过 " +
                            failedParagraphs + " 段；共替换 " + totalReplacements +
                            " 处。失败段未自动回滚，请对照 items/skipped 人工确认。",
                            data,
                            true);
                    }

                    return WordWriteResult.Ok(
                        "已规范化 " + changedParagraphs + " 段的 " + totalReplacements +
                        " 处标点（范围：" + scopeLabel + "，mode=" + mode + "），全部回读校验通过。",
                        data,
                        true).AddTargetRef("Word:Paragraphs");
                });
        }

        /// <summary>计算段落内 1:1 标点替换；返回按位置升序的变更列表。</summary>
        private static List<PunctuationChange> BuildPunctuationChanges(
            string text,
            string mode,
            bool normalizeQuotes)
        {
            List<PunctuationChange> changes = new List<PunctuationChange>();
            bool doubleQuoteOpen = true;
            bool singleQuoteOpen = true;
            for (int index = 0; index < text.Length; index++)
            {
                char original = text[index];
                char replacement = original;
                bool numericContext = IsNumericPunctuationContext(text, index);
                if (normalizeQuotes && original == '"')
                {
                    if (mode == "fullwidth")
                    {
                        replacement = doubleQuoteOpen ? '“' : '”';
                        doubleQuoteOpen = !doubleQuoteOpen;
                    }
                    else if (mode == "halfwidth" && (original == '“' || original == '”'))
                    {
                        replacement = '"';
                    }
                }
                else if (normalizeQuotes && original == '\'')
                {
                    char? left = WordExtensionsUtil.FindNeighborChar(text, index, -1);
                    char? right = WordExtensionsUtil.FindNeighborChar(text, index, 1);
                    bool betweenAsciiLetters =
                        left.HasValue && right.HasValue &&
                        WordExtensionsUtil.IsAsciiLetter(left.Value) &&
                        WordExtensionsUtil.IsAsciiLetter(right.Value);
                    if (mode == "fullwidth" && !betweenAsciiLetters)
                    {
                        replacement = singleQuoteOpen ? '‘' : '’';
                        singleQuoteOpen = !singleQuoteOpen;
                    }
                }
                else if (mode == "halfwidth" && (original == '“' || original == '”'))
                {
                    if (normalizeQuotes)
                    {
                        replacement = '"';
                    }
                }
                else if (mode == "halfwidth" && (original == '‘' || original == '’'))
                {
                    if (normalizeQuotes)
                    {
                        replacement = '\'';
                    }
                }
                else if (mode == "fullwidth")
                {
                    if (!numericContext || !IsAutoPunctuation(original))
                    {
                        char mapped;
                        if (HalfToFullPunctuation.TryGetValue(original, out mapped))
                        {
                            replacement = mapped;
                        }
                    }
                }
                else if (mode == "halfwidth")
                {
                    char mapped;
                    if (FullToHalfPunctuation.TryGetValue(original, out mapped))
                    {
                        replacement = mapped;
                    }
                }
                else
                {
                    replacement = ResolveAutoPunctuationReplacement(text, index, original, numericContext);
                }

                if (replacement != original)
                {
                    changes.Add(new PunctuationChange(index, original, replacement));
                }
            }
            return changes;
        }

        private static char ResolveAutoPunctuationReplacement(
            string text,
            int index,
            char original,
            bool numericContext)
        {
            if (numericContext || (!AutoHalfPunctuation.Contains(original) &&
                                   !AutoFullPunctuation.Contains(original)))
            {
                return original;
            }
            char? left = WordExtensionsUtil.FindNeighborChar(text, index, -1);
            char? right = WordExtensionsUtil.FindNeighborChar(text, index, 1);
            bool cjkContext = (left.HasValue && WordExtensionsUtil.IsCjkChar(left.Value)) ||
                              (right.HasValue && WordExtensionsUtil.IsCjkChar(right.Value));
            bool asciiContext = (left.HasValue &&
                                 WordExtensionsUtil.IsAsciiLetterOrDigit(left.Value)) ||
                                (right.HasValue &&
                                 WordExtensionsUtil.IsAsciiLetterOrDigit(right.Value));
            if (cjkContext)
            {
                char mapped;
                return HalfToFullPunctuation.TryGetValue(original, out mapped) ? mapped : original;
            }
            if (asciiContext)
            {
                char mapped;
                return FullToHalfPunctuation.TryGetValue(original, out mapped) ? mapped : original;
            }
            return original;
        }

        private static bool IsNumericPunctuationContext(string text, int index)
        {
            if (text == null || index < 0 || index >= text.Length)
            {
                return false;
            }
            char original = text[index];
            if (original != '.' && original != '．' && original != ',' && original != '，')
            {
                return false;
            }
            char? left = WordExtensionsUtil.FindNeighborChar(text, index, -1);
            char? right = WordExtensionsUtil.FindNeighborChar(text, index, 1);
            if (left.HasValue && WordExtensionsUtil.IsAsciiDigit(left.Value))
            {
                // 1.2 / 1,000 等数字上下文。
                if (right.HasValue && WordExtensionsUtil.IsAsciiDigit(right.Value))
                {
                    return true;
                }
                // "1. 标题" 等手工编号/列表序号：句点后紧跟空白时不转全角，
                // 否则会把编号变成“1。 标题”，破坏 word_normalize_headings 等下游工具。
                if ((original == '.' || original == '．') &&
                    index + 1 < text.Length &&
                    (text[index + 1] == ' ' || text[index + 1] == '\t' ||
                     text[index + 1] == '　'))
                {
                    return true;
                }
                if ((original == '.' || original == '．') && index + 1 >= text.Length)
                {
                    return true;
                }
            }
            return false;
        }

        private static bool IsAutoPunctuation(char value)
        {
            return AutoHalfPunctuation.Contains(value) || AutoFullPunctuation.Contains(value);
        }

        private static string ApplyPunctuationChanges(
            string original,
            List<PunctuationChange> changes)
        {
            if (string.IsNullOrEmpty(original) || changes == null || changes.Count == 0)
            {
                return original ?? string.Empty;
            }
            char[] buffer = original.ToCharArray();
            foreach (PunctuationChange change in changes)
            {
                if (change.RelativeIndex >= 0 && change.RelativeIndex < buffer.Length)
                {
                    buffer[change.RelativeIndex] = change.Replacement;
                }
            }
            return new string(buffer);
        }

        private static JArray BuildPunctuationChangeArray(
            List<PunctuationChange> changes,
            string original,
            int maxReport,
            ref int reportedChanges)
        {
            JArray array = new JArray();
            if (changes == null)
            {
                return array;
            }
            foreach (PunctuationChange change in changes)
            {
                if (reportedChanges >= maxReport)
                {
                    break;
                }
                array.Add(new JObject
                {
                    { "offset", change.RelativeIndex },
                    { "from", change.Original.ToString() },
                    { "to", change.Replacement.ToString() },
                    { "context", BuildContext(original, change.RelativeIndex, 12) }
                });
                reportedChanges++;
            }
            return array;
        }

        private static string BuildContext(string text, int index, int radius)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }
            int start = Math.Max(0, index - radius);
            int end = Math.Min(text.Length, index + radius + 1);
            return text.Substring(start, end - start);
        }

        private static Dictionary<char, char> BuildHalfToFullPunctuation()
        {
            Dictionary<char, char> map = new Dictionary<char, char>();
            map[','] = '，';
            map['.'] = '。';
            map[';'] = '；';
            map[':'] = '：';
            map['?'] = '？';
            map['!'] = '！';
            map['('] = '（';
            map[')'] = '）';
            map['%'] = '％';
            map['#'] = '＃';
            map['&'] = '＆';
            map['*'] = '＊';
            map['+'] = '＋';
            map['-'] = '－';
            map['/'] = '／';
            map['@'] = '＠';
            map['='] = '＝';
            map['$'] = '＄';
            map['~'] = '～';
            map['<'] = '＜';
            map['>'] = '＞';
            return map;
        }

        private static Dictionary<char, char> BuildFullToHalfPunctuation()
        {
            Dictionary<char, char> map = new Dictionary<char, char>();
            foreach (KeyValuePair<char, char> pair in HalfToFullPunctuation)
            {
                map[pair.Value] = pair.Key;
            }
            map['、'] = ',';
            return map;
        }

        private sealed class PunctuationChange
        {
            public int RelativeIndex { get; private set; }

            public char Original { get; private set; }

            public char Replacement { get; private set; }

            public PunctuationChange(int relativeIndex, char original, char replacement)
            {
                RelativeIndex = relativeIndex;
                Original = original;
                Replacement = replacement;
            }
        }

        // ------------------------------------------------------------------
        // word_cjk_spacing
        // ------------------------------------------------------------------

        private OperationResult HandleCjkSpacing(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            bool spaceWithLetter = args.GetOptionalBool("spaceBetweenCjkAndLetter", true);
            bool spaceWithDigit = args.GetOptionalBool("spaceBetweenCjkAndDigit", true);
            int maxReport = args.GetOptionalInt("maxReport", 300);
            if (!spaceWithLetter && !spaceWithDigit)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "spaceBetweenCjkAndLetter 与 spaceBetweenCjkAndDigit 不能同时为 false。");
            }
            if (maxReport < 1 || maxReport > 2000)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "maxReport 允许 1-2000，收到 " + maxReport + "。");
            }

            return RunWriteTool(
                application,
                WordExtensionsToolNames.CjkSpacing,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.None("批量段落中西文间距", "Word:Paragraphs");
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    List<int> indexes = WordExtensionsUtil.ResolveParagraphIndexes(
                        application, document, args, out string scopeLabel);
                    JArray items = new JArray();
                    JArray skipped = new JArray();
                    int changedParagraphs = 0;
                    int failedParagraphs = 0;
                    int totalAdded = 0;
                    int reportedItems = 0;
                    int scannedParagraphs = 0;

                    List<int> ordered = new List<int>(indexes);
                    ordered.Sort();
                    ordered.Reverse();

                    foreach (int paragraphIndex in ordered)
                    {
                        Word.Paragraph paragraph = null;
                        try
                        {
                            paragraph = WordComUtil.GetParagraph(document, paragraphIndex);
                            string original = WordExtensionsUtil.ReadParagraphText(paragraph);
                            if (original.Length == 0)
                            {
                                continue;
                            }
                            scannedParagraphs++;
                            List<int> positions = new List<int>();
                            if (spaceWithLetter)
                            {
                                foreach (Match match in CjkLetterBoundaryRegex.Matches(original))
                                {
                                    positions.Add(match.Index);
                                }
                            }
                            if (spaceWithDigit)
                            {
                                foreach (Match match in CjkDigitBoundaryRegex.Matches(original))
                                {
                                    positions.Add(match.Index);
                                }
                            }
                            if (positions.Count == 0)
                            {
                                continue;
                            }
                            positions.Sort();
                            List<int> uniquePositions = new List<int>();
                            int lastPosition = -1;
                            foreach (int position in positions)
                            {
                                if (position != lastPosition)
                                {
                                    uniquePositions.Add(position);
                                    lastPosition = position;
                                }
                            }

                            if (!WordExtensionsUtil.IsPlainTextParagraph(paragraph))
                            {
                                skipped.Add(new JObject
                                {
                                    { "paragraphIndex", paragraphIndex },
                                    { "reason", "段落含表格/图片/域/公式等非纯文本内容，已跳过以免破坏对象" }
                                });
                                continue;
                            }

                            string expected = InsertSpacesAt(original, uniquePositions);
                            Word.Range paragraphRange = null;
                            int paragraphStart;
                            try
                            {
                                paragraphRange = paragraph.Range;
                                paragraphStart = paragraphRange.Start;
                            }
                            finally
                            {
                                ComHelper.Release(paragraphRange);
                            }

                            for (int index = uniquePositions.Count - 1; index >= 0; index--)
                            {
                                int position = uniquePositions[index];
                                Word.Range insertion = null;
                                try
                                {
                                    insertion = WordComUtil.CreateRange(
                                        document,
                                        paragraphStart + position,
                                        paragraphStart + position);
                                    insertion.InsertAfter(" ");
                                }
                                finally
                                {
                                    ComHelper.Release(insertion);
                                }
                            }

                            string actual = WordExtensionsUtil.ReadParagraphText(paragraph);
                            totalAdded += uniquePositions.Count;
                            if (!string.Equals(actual, expected, StringComparison.Ordinal))
                            {
                                failedParagraphs++;
                                items.Add(new JObject
                                {
                                    { "paragraphIndex", paragraphIndex },
                                    { "verified", false },
                                    { "addedSpaces", uniquePositions.Count },
                                    { "beforePreview", WordExtensionsUtil.Preview(original, 120) },
                                    { "expectedPreview", WordExtensionsUtil.Preview(expected, 120) },
                                    { "actualPreview", WordExtensionsUtil.Preview(actual, 120) }
                                });
                                continue;
                            }

                            changedParagraphs++;
                            JObject item = new JObject
                            {
                                { "paragraphIndex", paragraphIndex },
                                { "verified", true },
                                { "addedSpaces", uniquePositions.Count },
                                { "beforePreview", WordExtensionsUtil.Preview(original, 120) },
                                { "afterPreview", WordExtensionsUtil.Preview(actual, 120) }
                            };
                            JArray samples = new JArray();
                            for (int index = 0;
                                 index < uniquePositions.Count && index < 10 && reportedItems < maxReport;
                                 index++)
                            {
                                samples.Add(new JObject
                                {
                                    { "offset", uniquePositions[index] },
                                    { "context", BuildContext(original, uniquePositions[index], 10) }
                                });
                                reportedItems++;
                            }
                            item["samples"] = samples;
                            items.Add(item);
                        }
                        catch (Exception ex)
                        {
                            failedParagraphs++;
                            skipped.Add(new JObject
                            {
                                { "paragraphIndex", paragraphIndex },
                                { "reason", OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) }
                            });
                        }
                        finally
                        {
                            ComHelper.Release(paragraph);
                        }
                    }

                    JObject data = new JObject();
                    data["scope"] = scopeLabel;
                    data["spaceBetweenCjkAndLetter"] = spaceWithLetter;
                    data["spaceBetweenCjkAndDigit"] = spaceWithDigit;
                    data["scannedParagraphs"] = scannedParagraphs;
                    data["changedParagraphs"] = changedParagraphs;
                    data["failedParagraphs"] = failedParagraphs;
                    data["totalAddedSpaces"] = totalAdded;
                    data["items"] = items;
                    data["skipped"] = skipped;
                    data["reportLimit"] = maxReport;
                    data["note"] = "盘古之白规则：只在汉字与拉丁字母/数字边界插入一个空格；不删除既有空格，" +
                                   "不处理汉字与汉字之间；含图片/域/公式/表格的段落跳过。";

                    if (totalAdded == 0)
                    {
                        return WordWriteResult.Ok(
                            "未发现需要补充中西文空格的位置（范围：" + scopeLabel + "）；文档未修改。",
                            data,
                            false);
                    }
                    if (failedParagraphs > 0 && changedParagraphs == 0)
                    {
                        return WordWriteResult.Fail(
                            OfficeErrorCode.INTERNAL_ERROR,
                            "中西文空格写入后回读校验全部失败（" + failedParagraphs +
                            " 段），文档内容无法确认。请检查文档是否受保护/启用修订。",
                            data,
                            false);
                    }
                    if (failedParagraphs > 0)
                    {
                        return WordWriteResult.Fail(
                            OfficeErrorCode.PARTIAL_APPLY,
                            "中西文空格部分完成：成功 " + changedParagraphs + " 段，失败/跳过 " +
                            failedParagraphs + " 段，共插入 " + totalAdded +
                            " 个空格。失败段未自动回滚，请对照 items/skipped 人工确认。",
                            data,
                            true);
                    }

                    return WordWriteResult.Ok(
                        "已在 " + changedParagraphs + " 段的 " + totalAdded +
                        " 个中西文边界插入空格（范围：" + scopeLabel + "），全部回读校验通过。",
                        data,
                        true).AddTargetRef("Word:Paragraphs");
                });
        }

        private static string InsertSpacesAt(string original, List<int> positions)
        {
            if (string.IsNullOrEmpty(original) || positions == null || positions.Count == 0)
            {
                return original ?? string.Empty;
            }
            StringBuilder builder = new StringBuilder(original);
            for (int index = positions.Count - 1; index >= 0; index--)
            {
                int position = positions[index];
                if (position >= 0 && position <= builder.Length)
                {
                    builder.Insert(position, ' ');
                }
            }
            return builder.ToString();
        }

        // ------------------------------------------------------------------
        // word_normalize_headings
        // ------------------------------------------------------------------

        private OperationResult HandleNormalizeHeadings(
            Word.Application application,
            JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            bool addNumbering = args.GetOptionalBool("addNumbering", true);
            int maxLevel = args.GetOptionalInt("maxLevel", 3);
            int maxReport = args.GetOptionalInt("maxReport", 300);
            if (maxLevel < 1 || maxLevel > 6)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "maxLevel 允许 1-6，收到 " + maxLevel + "。");
            }
            if (maxReport < 1 || maxReport > 2000)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "maxReport 允许 1-2000，收到 " + maxReport + "。");
            }

            return RunWriteTool(
                application,
                WordExtensionsToolNames.NormalizeHeadings,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.None("手工编号标题转换", "Word:Paragraphs");
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    List<int> indexes = WordExtensionsUtil.ResolveParagraphIndexes(
                        application, document, args, out string scopeLabel);
                    List<HeadingCandidate> candidates = new List<HeadingCandidate>();
                    JArray skipped = new JArray();
                    int scannedParagraphs = 0;

                    foreach (int paragraphIndex in indexes)
                    {
                        Word.Paragraph paragraph = null;
                        try
                        {
                            paragraph = WordComUtil.GetParagraph(document, paragraphIndex);
                            string text = WordExtensionsUtil.ReadParagraphText(paragraph);
                            if (text.Length == 0)
                            {
                                continue;
                            }
                            scannedParagraphs++;

                            int detectedLevel;
                            int prefixLength;
                            string prefixText;
                            if (!WordExtensionsUtil.TryMatchManualHeading(
                                text, out detectedLevel, out prefixLength, out prefixText))
                            {
                                continue;
                            }
                            if (!WordExtensionsUtil.IsPlainTextParagraph(paragraph))
                            {
                                skipped.Add(new JObject
                                {
                                    { "paragraphIndex", paragraphIndex },
                                    { "reason", "段落含表格/图片/域/公式等非纯文本内容，已跳过" }
                                });
                                continue;
                            }

                            int targetLevel = Math.Min(detectedLevel, maxLevel);
                            string styleName = WordExtensionsUtil.ResolveHeadingStyleName(
                                document, targetLevel);
                            if (string.IsNullOrWhiteSpace(styleName))
                            {
                                skipped.Add(new JObject
                                {
                                    { "paragraphIndex", paragraphIndex },
                                    { "reason", "文档中找不到层级 " + targetLevel +
                                                " 的标题样式（已尝试「标题 " + targetLevel +
                                                "」/「Heading " + targetLevel + "」）" }
                                });
                                continue;
                            }

                            HeadingCandidate candidate = new HeadingCandidate();
                            candidate.ParagraphIndex = paragraphIndex;
                            candidate.Level = targetLevel;
                            candidate.DetectedLevel = detectedLevel;
                            candidate.PrefixLength = prefixLength;
                            candidate.PrefixText = prefixText;
                            candidate.OriginalText = text;
                            candidate.StyleName = styleName;
                            candidates.Add(candidate);
                        }
                        finally
                        {
                            ComHelper.Release(paragraph);
                        }
                    }

                    if (candidates.Count == 0)
                    {
                        JObject emptyData = new JObject();
                        emptyData["scope"] = scopeLabel;
                        emptyData["scannedParagraphs"] = scannedParagraphs;
                        emptyData["skipped"] = skipped;
                        return WordWriteResult.Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "未发现可转换的手工编号标题段（范围：" + scopeLabel + "，扫描 " +
                            scannedParagraphs + " 段）。支持：第X章/第X节、一、二、、（一）、" +
                            "1. / 1.2 / 1.2.3 + 空格、1 空格 标题。请先用 word_style_audit 检查候选段落。",
                            emptyData,
                            false);
                    }

                    List<int> writeOrder = new List<int>();
                    foreach (HeadingCandidate candidate in candidates)
                    {
                        writeOrder.Add(candidate.ParagraphIndex);
                    }
                    writeOrder.Sort();
                    writeOrder.Reverse();

                    Dictionary<int, HeadingCandidate> byIndex =
                        new Dictionary<int, HeadingCandidate>();
                    foreach (HeadingCandidate candidate in candidates)
                    {
                        byIndex[candidate.ParagraphIndex] = candidate;
                    }

                    JArray converted = new JArray();
                    JArray failures = new JArray();
                    HashSet<int> usedLevels = new HashSet<int>();
                    int convertedCount = 0;
                    bool convertedReportTruncated = false;

                    foreach (int paragraphIndex in writeOrder)
                    {
                        HeadingCandidate candidate = byIndex[paragraphIndex];
                        Word.Paragraph paragraph = null;
                        Word.Range range = null;
                        Word.Font font = null;
                        try
                        {
                            paragraph = WordComUtil.GetParagraph(document, paragraphIndex);
                            string currentText = WordExtensionsUtil.ReadParagraphText(paragraph);
                            if (!currentText.StartsWith(
                                candidate.PrefixText, StringComparison.Ordinal))
                            {
                                failures.Add(new JObject
                                {
                                    { "paragraphIndex", paragraphIndex },
                                    { "reason", "写入前段落文本已变化，前缀不再是「" +
                                                WordExtensionsUtil.Preview(candidate.PrefixText, 40) +
                                                "」；当前：" +
                                                WordExtensionsUtil.Preview(currentText, 80) }
                                });
                                continue;
                            }

                            string remaining = currentText.Substring(candidate.PrefixLength).TrimStart();
                            if (remaining.Length == 0)
                            {
                                failures.Add(new JObject
                                {
                                    { "paragraphIndex", paragraphIndex },
                                    { "reason", "去掉手工编号后没有标题文字，已跳过以免制造空标题" }
                                });
                                continue;
                            }

                            string actualText = WordExtensionsUtil.SetParagraphText(paragraph, remaining);
                            if (!string.Equals(actualText, remaining, StringComparison.Ordinal))
                            {
                                failures.Add(new JObject
                                {
                                    { "paragraphIndex", paragraphIndex },
                                    { "reason", "去除手工编号后回读文本不一致；预期「" +
                                                WordExtensionsUtil.Preview(remaining, 80) +
                                                "」，实际「" +
                                                WordExtensionsUtil.Preview(actualText, 80) + "」" }
                                });
                                continue;
                            }

                            // Range.Text 整段替换可能让原 Paragraph COM 对象指向下一段；
                            // 必须按原段号重新取一次，避免把标题样式套到下一段。
                            ComHelper.Release(paragraph);
                            paragraph = WordComUtil.GetParagraph(document, paragraphIndex);
                            string actualStyle = WordExtensionsUtil.ApplyParagraphStyle(
                                document, paragraph, candidate.StyleName, true, false);
                            if (!string.Equals(actualStyle, candidate.StyleName,
                                StringComparison.OrdinalIgnoreCase))
                            {
                                failures.Add(new JObject
                                {
                                    { "paragraphIndex", paragraphIndex },
                                    { "reason", "应用标题样式后回读为「" + actualStyle +
                                                "」，期望「" + candidate.StyleName + "」" }
                                });
                                continue;
                            }

                            candidate.RemainingText = remaining;
                            candidate.ActualStyle = actualStyle;
                            candidate.TextChanged = true;
                            convertedCount++;
                            usedLevels.Add(candidate.Level);
                            if (converted.Count < maxReport)
                            {
                                converted.Add(new JObject
                                {
                                    { "paragraphIndex", paragraphIndex },
                                    { "level", candidate.Level },
                                    { "detectedLevel", candidate.DetectedLevel },
                                    { "removedPrefix", candidate.PrefixText.Trim() },
                                    { "styleApplied", actualStyle },
                                    { "textPreview", WordExtensionsUtil.Preview(remaining, 100) }
                                });
                            }
                            else
                            {
                                convertedReportTruncated = true;
                            }
                        }
                        catch (Exception ex)
                        {
                            failures.Add(new JObject
                            {
                                { "paragraphIndex", paragraphIndex },
                                { "reason", OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) }
                            });
                        }
                        finally
                        {
                            ComHelper.Release(font, range, paragraph);
                        }
                    }

                    JArray warnings = new JArray();
                    string numberingNote = "未请求自动编号（addNumbering=false），只转换标题样式。";
                    if (addNumbering && convertedCount > 0)
                    {
                        numberingNote = ApplyHeadingNumbering(
                            document, candidates, usedLevels, warnings);
                    }

                    JArray numberingItems = new JArray();
                    int missingListString = 0;
                    int styleLostAfterNumbering = 0;
                    if (addNumbering)
                    {
                        foreach (HeadingCandidate candidate in candidates)
                        {
                            if (!candidate.TextChanged)
                            {
                                continue;
                            }
                            Word.Paragraph paragraph = null;
                            Word.Range range = null;
                            Word.ListFormat listFormat = null;
                            try
                            {
                                paragraph = WordComUtil.GetParagraph(document, candidate.ParagraphIndex);
                                range = paragraph.Range;
                                listFormat = range.ListFormat;
                                string listString = listFormat.ListString ?? string.Empty;
                                string styleAfterNumbering = WordComUtil.ReadStyleName(range);
                                bool styleOk = string.Equals(
                                    styleAfterNumbering, candidate.StyleName,
                                    StringComparison.OrdinalIgnoreCase);
                                if (string.IsNullOrWhiteSpace(listString))
                                {
                                    missingListString++;
                                }
                                if (!styleOk)
                                {
                                    styleLostAfterNumbering++;
                                    WordExtensionsUtil.AddWarning(
                                        warnings,
                                        "第 " + candidate.ParagraphIndex + " 段添加编号后样式变为「" +
                                        styleAfterNumbering + "」，期望「" + candidate.StyleName +
                                        "」。");
                                }
                                numberingItems.Add(new JObject
                                {
                                    { "paragraphIndex", candidate.ParagraphIndex },
                                    { "level", candidate.Level },
                                    { "listString", listString },
                                    { "styleAfterNumbering", styleAfterNumbering },
                                    { "styleOk", styleOk }
                                });
                            }
                            catch (Exception ex)
                            {
                                missingListString++;
                                WordExtensionsUtil.AddWarning(
                                    warnings,
                                    "第 " + candidate.ParagraphIndex + " 段读取编号失败：" +
                                    OfficeErrorMapper.DescribeException(ex, OfficeApp.Word));
                            }
                            finally
                            {
                                ComHelper.Release(listFormat, range, paragraph);
                            }
                        }
                    }

                    JObject data = new JObject();
                    data["scope"] = scopeLabel;
                    data["scannedParagraphs"] = scannedParagraphs;
                    data["convertedCount"] = convertedCount;
                    data["failureCount"] = failures.Count;
                    data["skippedCount"] = skipped.Count;
                    data["addNumbering"] = addNumbering;
                    data["maxLevel"] = maxLevel;
                    data["converted"] = converted;
                    data["failures"] = failures;
                    data["skipped"] = skipped;
                    data["warnings"] = warnings;
                    data["numberingNote"] = numberingNote;
                    data["numbering"] = numberingItems;
                    data["styleLostAfterNumbering"] = styleLostAfterNumbering;
                    data["reportLimit"] = maxReport;
                    data["reportTruncated"] = convertedReportTruncated;
                    data["note"] = "转换时先去掉手工编号文本，再清除该段直接字体格式并套用标题样式；" +
                                   "addNumbering=true 时尝试把多级编号链接到标题样式（编号进样式），" +
                                   "失败时回退为段落级编号并在 warnings 说明。";

                    if (convertedCount == 0)
                    {
                        return WordWriteResult.Fail(
                            OfficeErrorCode.INTERNAL_ERROR,
                            "检测到 " + candidates.Count + " 个手工编号标题候选，但全部转换失败；" +
                            "首个失败：" + (failures.Count > 0
                                ? failures[0]["reason"].ToString()
                                : "未知原因") + "。请确认文档未受保护且段落未被锁定。",
                            data,
                            false);
                    }
                    if (failures.Count > 0 || (addNumbering && missingListString > 0) ||
                        styleLostAfterNumbering > 0)
                    {
                        string detail;
                        if (failures.Count > 0)
                        {
                            detail = "首个失败：" + failures[0]["reason"].ToString();
                        }
                        else if (missingListString > 0)
                        {
                            detail = "有 " + missingListString +
                                     " 段未读到自动编号（ListString 为空）。";
                        }
                        else
                        {
                            detail = "有 " + styleLostAfterNumbering +
                                     " 段添加编号后丢失标题样式。";
                        }
                        return WordWriteResult.Fail(
                            OfficeErrorCode.PARTIAL_APPLY,
                            "手工编号标题转换部分完成：成功 " + convertedCount + " 段，失败 " +
                            failures.Count + " 段，" + detail +
                            " 已完成的改动不会自动回滚，请对照 converted/failures/numbering 人工确认。",
                            data,
                            true);
                    }

                    return WordWriteResult.Ok(
                        "已将 " + convertedCount + " 个手工编号段落转换为标题样式（范围：" +
                        scopeLabel + "，" + (addNumbering
                            ? "并尝试把多级编号链接到标题样式"
                            : "未添加自动编号") + "）。",
                        data,
                        true).AddTargetRef("Word:Paragraphs");
                });
        }

        private static string ApplyHeadingNumbering(
            Word.Document document,
            List<HeadingCandidate> candidates,
            HashSet<int> usedLevels,
            JArray warnings)
        {
            Word.ListTemplates templates = null;
            Word.ListTemplate template = null;
            Word.ListLevels levels = null;
            try
            {
                templates = document.ListTemplates;
                object outlineNumbered = true;
                object templateName = "AutoOfficeMCPHeadings";
                template = templates.Add(ref outlineNumbered, ref templateName);
                levels = template.ListLevels;
                int maxLevel = 1;
                foreach (int level in usedLevels)
                {
                    maxLevel = Math.Max(maxLevel, level);
                }
                for (int level = 1; level <= maxLevel; level++)
                {
                    Word.ListLevel listLevel = null;
                    try
                    {
                        listLevel = levels[level];
                        listLevel.NumberFormat = WordExtensionsUtil.BuildNumberFormat(level);
                        listLevel.NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
                        listLevel.StartAt = 1;
                        listLevel.NumberPosition = (level - 1) * 21f;
                        listLevel.TextPosition = (level - 1) * 21f + 21f;
                        listLevel.ResetOnHigher = level - 1;
                        listLevel.TrailingCharacter = Word.WdTrailingCharacter.wdTrailingSpace;
                    }
                    finally
                    {
                        ComHelper.Release(listLevel);
                    }
                }

                bool anyStyleLinked = false;
                Dictionary<string, int> styleLevels = new Dictionary<string, int>(
                    StringComparer.OrdinalIgnoreCase);
                foreach (HeadingCandidate candidate in candidates)
                {
                    if (candidate.TextChanged && !styleLevels.ContainsKey(candidate.StyleName))
                    {
                        styleLevels[candidate.StyleName] = candidate.Level;
                    }
                }
                foreach (KeyValuePair<string, int> pair in styleLevels)
                {
                    Word.Style style = null;
                    object styleObject = null;
                    try
                    {
                        object rawStyle = document.Styles[pair.Key];
                        style = rawStyle as Word.Style;
                        if (style == null)
                        {
                            WordExtensionsUtil.AddWarning(
                                warnings, "样式「" + pair.Key + "」不是可链接的段落样式，改用段落级编号。");
                            continue;
                        }
                        object linkLevel = pair.Value;
                        style.LinkToListTemplate(template, ref linkLevel);
                        anyStyleLinked = true;
                    }
                    catch (Exception ex)
                    {
                        WordExtensionsUtil.AddWarning(
                            warnings,
                            "样式「" + pair.Key + "」链接多级编号失败：" +
                            OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) +
                            "；将回退为段落级编号。");
                    }
                    finally
                    {
                        ComHelper.Release(style, styleObject);
                    }
                }

                int fallbackCount = 0;
                foreach (HeadingCandidate candidate in candidates)
                {
                    if (!candidate.TextChanged)
                    {
                        continue;
                    }
                    Word.Paragraph paragraph = null;
                    Word.Range range = null;
                    Word.ListFormat listFormat = null;
                    try
                    {
                        paragraph = WordComUtil.GetParagraph(document, candidate.ParagraphIndex);
                        range = paragraph.Range;
                        if (anyStyleLinked)
                        {
                            // 样式链接编号是“样式定义”变更，已存在段落需要重新套一次样式才会刷新 ListString。
                            WordExtensionsUtil.ApplyParagraphStyle(
                                document, paragraph, candidate.StyleName, false, false);
                        }
                        listFormat = range.ListFormat;
                        string existing = listFormat.ListString ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(existing))
                        {
                            continue;
                        }
                        object continuePrevious = true;
                        object applyTo = Word.WdListApplyTo.wdListApplyToWholeList;
                        object behavior = Word.WdDefaultListBehavior.wdWord9ListBehavior;
                        object applyLevel = candidate.Level;
                        listFormat.ApplyListTemplateWithLevel(
                            template, ref continuePrevious, ref applyTo, ref behavior, ref applyLevel);
                        ComHelper.Release(listFormat);
                        listFormat = null;
                        // ApplyListTemplate 可能把段落样式改成“列表段落”；恢复标题样式并核对编号仍在。
                        WordExtensionsUtil.ApplyParagraphStyle(
                            document, paragraph, candidate.StyleName, false, false);
                        listFormat = range.ListFormat;
                        string afterFallback = listFormat.ListString ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(afterFallback))
                        {
                            WordExtensionsUtil.AddWarning(
                                warnings,
                                "第 " + candidate.ParagraphIndex +
                                " 段回退编号后恢复标题样式，但 ListString 仍为空；样式已恢复为「" +
                                candidate.StyleName + "」，编号缺失。");
                        }
                        fallbackCount++;
                    }
                    catch (Exception ex)
                    {
                        WordExtensionsUtil.AddWarning(
                            warnings,
                            "第 " + candidate.ParagraphIndex + " 段回退编号失败：" +
                            OfficeErrorMapper.DescribeException(ex, OfficeApp.Word));
                    }
                    finally
                    {
                        ComHelper.Release(listFormat, range, paragraph);
                    }
                }

                if (anyStyleLinked && fallbackCount == 0)
                {
                    return "已创建多级编号模板并把编号链接到标题样式（样式级编号，编号随标题样式更新）。";
                }
                if (anyStyleLinked)
                {
                    return "已把编号链接到部分标题样式；另有 " + fallbackCount +
                           " 段回退为段落级编号（详见 warnings）。";
                }
                return fallbackCount > 0
                    ? "未能把编号链接到标题样式，已对 " + fallbackCount +
                      " 段应用段落级多级编号（详见 warnings）。"
                    : "未能建立多级编号（详见 warnings），标题样式已转换但未添加自动编号。";
            }
            catch (Exception ex)
            {
                WordExtensionsUtil.AddWarning(
                    warnings,
                    "创建/应用多级编号失败：" +
                    OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) +
                    "；标题样式已转换，但编号未写入。");
                return "多级编号创建失败，标题样式已转换但未添加自动编号（详见 warnings）。";
            }
            finally
            {
                ComHelper.Release(levels, template, templates);
            }
        }

        private sealed class HeadingCandidate
        {
            public int ParagraphIndex { get; set; }

            public int Level { get; set; }

            public int DetectedLevel { get; set; }

            public int PrefixLength { get; set; }

            public string PrefixText { get; set; }

            public string OriginalText { get; set; }

            public string RemainingText { get; set; }

            public string StyleName { get; set; }

            public string ActualStyle { get; set; }

            public bool TextChanged { get; set; }

            public HeadingCandidate()
            {
                PrefixText = string.Empty;
                OriginalText = string.Empty;
                RemainingText = string.Empty;
                StyleName = string.Empty;
                ActualStyle = string.Empty;
            }
        }

        // ------------------------------------------------------------------
        // word_repeat_check（只读）
        // ------------------------------------------------------------------

        private OperationResult HandleRepeatCheck(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int minSentenceLength = args.GetOptionalInt("minSentenceLength", 10);
            int ngramSize = args.GetOptionalInt("ngramSize", 4);
            int topN = args.GetOptionalInt("topN", 20);
            int minRepeatCount = args.GetOptionalInt("minRepeatCount", 2);
            if (minSentenceLength < 2 || minSentenceLength > 500)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "minSentenceLength 允许 2-500，收到 " + minSentenceLength + "。");
            }
            if (ngramSize < 2 || ngramSize > 10)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "ngramSize 允许 2-10，收到 " + ngramSize + "。");
            }
            if (topN < 1 || topN > 100)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "topN 允许 1-100，收到 " + topN + "。");
            }
            if (minRepeatCount < 2 || minRepeatCount > 100)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "minRepeatCount 允许 2-100，收到 " + minRepeatCount + "。");
            }

            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    List<int> indexes = WordExtensionsUtil.ResolveParagraphIndexes(
                        application, document, args, out string scopeLabel);
                    Dictionary<string, int> sentenceCounts =
                        new Dictionary<string, int>(StringComparer.Ordinal);
                    Dictionary<string, int> paragraphCounts =
                        new Dictionary<string, int>(StringComparer.Ordinal);
                    Dictionary<string, int> ngramCounts =
                        new Dictionary<string, int>(StringComparer.Ordinal);
                    Dictionary<string, int> wordCounts =
                        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    int totalSentences = 0;
                    int charactersScanned = 0;
                    int paragraphsScanned = 0;

                    foreach (int paragraphIndex in indexes)
                    {
                        Word.Paragraph paragraph = null;
                        try
                        {
                            paragraph = WordComUtil.GetParagraph(document, paragraphIndex);
                            string text = WordExtensionsUtil.ReadParagraphText(paragraph);
                            if (text.Length == 0)
                            {
                                continue;
                            }
                            paragraphsScanned++;
                            charactersScanned += text.Length;

                            string normalizedParagraph =
                                WordExtensionsUtil.NormalizeWhitespace(text);
                            if (normalizedParagraph.Length >= minSentenceLength)
                            {
                                Increment(paragraphCounts, normalizedParagraph);
                            }

                            string[] sentences = SentenceSplitRegex.Split(text);
                            foreach (string rawSentence in sentences)
                            {
                                string sentence =
                                    WordExtensionsUtil.NormalizeWhitespace(rawSentence);
                                if (sentence.Length < minSentenceLength)
                                {
                                    continue;
                                }
                                totalSentences++;
                                Increment(sentenceCounts, sentence);
                            }

                            CollectNgrams(text, ngramSize, ngramCounts);
                            foreach (Match match in EnglishWordRegex.Matches(text))
                            {
                                if (match.Value.Length >= 3)
                                {
                                    Increment(wordCounts, match.Value.ToLowerInvariant());
                                }
                            }
                        }
                        finally
                        {
                            ComHelper.Release(paragraph);
                        }
                    }

                    JArray repeatedSentences = BuildTopRepeats(
                        sentenceCounts, topN, minRepeatCount);
                    JArray repeatedParagraphs = BuildTopRepeats(
                        paragraphCounts, topN, minRepeatCount);
                    JArray repeatedNgrams = BuildTopRepeats(
                        ngramCounts, topN, minRepeatCount);
                    JArray repeatedWords = BuildTopRepeats(
                        wordCounts, topN, minRepeatCount);

                    JObject data = new JObject();
                    data["scope"] = scopeLabel;
                    data["paragraphsScanned"] = paragraphsScanned;
                    data["charactersScanned"] = charactersScanned;
                    data["totalSentences"] = totalSentences;
                    data["uniqueSentences"] = sentenceCounts.Count;
                    data["repeatedSentenceCount"] = CountRepeated(
                        sentenceCounts, minRepeatCount);
                    data["ngramSize"] = ngramSize;
                    data["minSentenceLength"] = minSentenceLength;
                    data["minRepeatCount"] = minRepeatCount;
                    data["topN"] = topN;
                    data["repeatedSentences"] = repeatedSentences;
                    data["repeatedParagraphs"] = repeatedParagraphs;
                    data["repeatedNgrams"] = repeatedNgrams;
                    data["repeatedWords"] = repeatedWords;
                    data["note"] = "纯本地统计，不做语义判断。句级重复按中文/英文句末标点切分并规范化空白；" +
                                    "中文使用连续汉字 n-gram（默认 4 字），英文使用长度≥3 的单词；" +
                                    "仅返回出现次数≥minRepeatCount 的条目。";

                    return ReadSucceed(
                        "查重统计完成：扫描 " + paragraphsScanned + " 段、" + charactersScanned +
                        " 字符、" + totalSentences + " 句；重复句 " +
                        CountRepeated(sentenceCounts, minRepeatCount) + " 条，重复 " +
                        ngramSize + "-gram " + CountRepeated(ngramCounts, minRepeatCount) + " 条。",
                        data, "Word:Document");
                });
        }

        private static void CollectNgrams(
            string text,
            int ngramSize,
            Dictionary<string, int> counts)
        {
            if (string.IsNullOrEmpty(text) || ngramSize <= 0)
            {
                return;
            }
            int start = -1;
            for (int index = 0; index <= text.Length; index++)
            {
                bool isCjk = index < text.Length && WordExtensionsUtil.IsCjkChar(text[index]);
                if (isCjk && start < 0)
                {
                    start = index;
                }
                else if (!isCjk && start >= 0)
                {
                    AddCjkNgrams(text, start, index, ngramSize, counts);
                    start = -1;
                }
            }
        }

        private static void AddCjkNgrams(
            string text,
            int start,
            int end,
            int ngramSize,
            Dictionary<string, int> counts)
        {
            int length = end - start;
            if (length < ngramSize)
            {
                return;
            }
            for (int index = start; index + ngramSize <= end; index++)
            {
                string gram = text.Substring(index, ngramSize);
                if (WordExtensionsUtil.ContainsCjk(gram))
                {
                    Increment(counts, gram);
                }
            }
        }

        private static void Increment(Dictionary<string, int> counts, string key)
        {
            if (counts == null || string.IsNullOrEmpty(key))
            {
                return;
            }
            int current;
            counts.TryGetValue(key, out current);
            counts[key] = current + 1;
        }

        private static int CountRepeated(Dictionary<string, int> counts, int minCount)
        {
            int repeated = 0;
            foreach (KeyValuePair<string, int> pair in counts)
            {
                if (pair.Value >= minCount)
                {
                    repeated++;
                }
            }
            return repeated;
        }

        private static JArray BuildTopRepeats(
            Dictionary<string, int> counts,
            int topN,
            int minCount)
        {
            List<KeyValuePair<string, int>> filtered =
                new List<KeyValuePair<string, int>>();
            foreach (KeyValuePair<string, int> pair in counts)
            {
                if (pair.Value >= minCount)
                {
                    filtered.Add(pair);
                }
            }
            filtered.Sort(delegate(KeyValuePair<string, int> left, KeyValuePair<string, int> right)
            {
                int countCompare = right.Value.CompareTo(left.Value);
                return countCompare != 0
                    ? countCompare
                    : string.Compare(left.Key, right.Key, StringComparison.Ordinal);
            });

            JArray result = new JArray();
            int limit = Math.Min(topN, filtered.Count);
            for (int index = 0; index < limit; index++)
            {
                result.Add(new JObject
                {
                    { "text", WordExtensionsUtil.Preview(filtered[index].Key, 200) },
                    { "count", filtered[index].Value }
                });
            }
            return result;
        }

        // ------------------------------------------------------------------
        // word_hidden_text_report（只读）
        // ------------------------------------------------------------------

        private OperationResult HandleHiddenTextReport(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int maxItems = args.GetOptionalInt("maxItems", 200);
            if (maxItems < 1 || maxItems > 2000)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "maxItems 允许 1-2000，收到 " + maxItems + "。");
            }

            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    Word.Range searchRange = null;
                    Word.Find find = null;
                    JArray items = new JArray();
                    int hiddenRangeCount = 0;
                    int hiddenCharacterCount = 0;
                    bool truncated = false;
                    int guard = 0;
                    int lastEnd = -1;
                    try
                    {
                        searchRange = document.Content;
                        find = searchRange.Find;
                        find.ClearFormatting();
                        find.Text = string.Empty;
                        find.Format = true;
                        find.Forward = true;
                        find.Wrap = Word.WdFindWrap.wdFindStop;
                        find.MatchWildcards = false;
                        find.Font.Hidden = -1;

                        while (guard < 10000 && ExecuteFind(find))
                        {
                            guard++;
                            int start = searchRange.Start;
                            int end = searchRange.End;
                            if (end <= start || (end <= lastEnd && start >= lastEnd))
                            {
                                break;
                            }
                            string text = searchRange.Text ?? string.Empty;
                            hiddenRangeCount++;
                            hiddenCharacterCount += text.Length;
                            if (items.Count < maxItems)
                            {
                                items.Add(new JObject
                                {
                                    { "ordinal", hiddenRangeCount },
                                    { "paragraphIndex", WordComUtil.FindParagraphIndexAt(document, start) },
                                    { "start", start },
                                    { "end", end },
                                    { "length", text.Length },
                                    { "textPreview", WordExtensionsUtil.Preview(text, 160) }
                                });
                            }
                            else
                            {
                                truncated = true;
                            }
                            lastEnd = end;
                            searchRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        }
                    }
                    finally
                    {
                        ComHelper.Release(find, searchRange);
                    }

                    JObject data = new JObject();
                    data["hiddenRangeCount"] = hiddenRangeCount;
                    data["hiddenCharacterCount"] = hiddenCharacterCount;
                    data["items"] = items;
                    data["maxItems"] = maxItems;
                    data["truncated"] = truncated;
                    data["scanGuard"] = guard;
                    data["note"] = "通过 Find.Format=true + Font.Hidden=-1 扫描正文故事；" +
                                    "页眉页脚、文本框、脚注中的隐藏文字不在本次扫描范围。";

                    return ReadSucceed(
                        hiddenRangeCount == 0
                            ? "未在正文中发现隐藏文字。"
                            : "发现 " + hiddenRangeCount + " 处隐藏文字，共 " +
                              hiddenCharacterCount + " 字符" +
                              (truncated ? "（仅返回前 " + maxItems + " 条）" : "") + "。",
                        data, "Word:Document");
                });
        }

        private static bool ExecuteFind(Word.Find find)
        {
            object missing = Type.Missing;
            return find.Execute(
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing);
        }

        // ------------------------------------------------------------------
        // word_reading_time（只读）
        // ------------------------------------------------------------------

        private OperationResult HandleReadingTime(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int cjkCharsPerMinute = args.GetOptionalInt("cjkCharsPerMinute", 400);
            int englishWordsPerMinute = args.GetOptionalInt("englishWordsPerMinute", 200);
            if (cjkCharsPerMinute < 50 || cjkCharsPerMinute > 3000)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "cjkCharsPerMinute 允许 50-3000，收到 " + cjkCharsPerMinute + "。");
            }
            if (englishWordsPerMinute < 30 || englishWordsPerMinute > 2000)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "englishWordsPerMinute 允许 30-2000，收到 " + englishWordsPerMinute + "。");
            }

            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    Word.Range content = null;
                    string text = string.Empty;
                    try
                    {
                        content = document.Content;
                        text = content == null ? string.Empty : (content.Text ?? string.Empty);
                    }
                    finally
                    {
                        ComHelper.Release(content);
                    }

                    int cjkCharacters = 0;
                    int punctuationCharacters = 0;
                    int latinLetters = 0;
                    foreach (char value in text)
                    {
                        if (WordExtensionsUtil.IsCjkChar(value))
                        {
                            cjkCharacters++;
                        }
                        else if (WordExtensionsUtil.IsAsciiLetter(value))
                        {
                            latinLetters++;
                        }
                        else if (IsPunctuationCharacter(value))
                        {
                            punctuationCharacters++;
                        }
                    }
                    int englishWords = EnglishWordRegex.Matches(text).Count;
                    int digitGroups = 0;
                    bool inDigits = false;
                    foreach (char value in text)
                    {
                        bool isDigit = value >= '0' && value <= '9';
                        if (isDigit && !inDigits)
                        {
                            digitGroups++;
                        }
                        inDigits = isDigit;
                    }

                    double minutes = (double)cjkCharacters / cjkCharsPerMinute +
                                     (double)englishWords / englishWordsPerMinute;
                    double roundedMinutes = Math.Round(minutes, 1, MidpointRounding.AwayFromZero);
                    if (roundedMinutes <= 0 && text.Trim().Length > 0)
                    {
                        roundedMinutes = 0.1;
                    }
                    int totalParagraphs = WordComUtil.GetParagraphCount(document);
                    int headingCount = 0;
                    Word.Paragraphs paragraphs = null;
                    try
                    {
                        paragraphs = document.Paragraphs;
                        int count = paragraphs == null ? 0 : paragraphs.Count;
                        for (int index = 1; index <= count; index++)
                        {
                            Word.Paragraph paragraph = null;
                            Word.Range range = null;
                            try
                            {
                                paragraph = paragraphs[index];
                                range = paragraph.Range;
                                if (WordExtensionsUtil.TryGetHeadingLevel(
                                    WordComUtil.ReadStyleName(range)) > 0)
                                {
                                    headingCount++;
                                }
                            }
                            finally
                            {
                                ComHelper.Release(range, paragraph);
                            }
                        }
                    }
                    finally
                    {
                        ComHelper.Release(paragraphs);
                    }

                    JObject data = new JObject();
                    data["cjkCharacters"] = cjkCharacters;
                    data["latinLetters"] = latinLetters;
                    data["englishWords"] = englishWords;
                    data["digitGroups"] = digitGroups;
                    data["punctuationCharacters"] = punctuationCharacters;
                    data["totalCharacters"] = text.Length;
                    data["paragraphs"] = totalParagraphs;
                    data["headings"] = headingCount;
                    data["tables"] = CountDocumentCollection(document, "tables");
                    data["inlineImages"] = CountDocumentCollection(document, "inlineShapes");
                    data["floatingShapes"] = CountDocumentCollection(document, "shapes");
                    data["cjkCharsPerMinute"] = cjkCharsPerMinute;
                    data["englishWordsPerMinute"] = englishWordsPerMinute;
                    data["readingMinutes"] = roundedMinutes;
                    data["readingSeconds"] = (int)Math.Round(
                        roundedMinutes * 60d, MidpointRounding.AwayFromZero);
                    data["note"] = "阅读时长 = 汉字数/汉字速度 + 英文单词数/英文速度；" +
                                    "数字串按组统计但不计入阅读时长，标点与图片不计入。";

                    return ReadSucceed(
                        "预计阅读 " + roundedMinutes.ToString("0.0", CultureInfo.InvariantCulture) +
                        " 分钟（汉字 " + cjkCharacters + "，英文词 " + englishWords + "）。",
                        data, "Word:Document");
                });
        }

        private static int CountDocumentCollection(Word.Document document, string collectionName)
        {
            object collection = null;
            try
            {
                if (collectionName == "tables")
                {
                    collection = document.Tables;
                    Word.Tables tables = collection as Word.Tables;
                    return tables == null ? 0 : tables.Count;
                }
                if (collectionName == "inlineShapes")
                {
                    collection = document.InlineShapes;
                    Word.InlineShapes shapes = collection as Word.InlineShapes;
                    return shapes == null ? 0 : shapes.Count;
                }
                collection = document.Shapes;
                Word.Shapes floating = collection as Word.Shapes;
                return floating == null ? 0 : floating.Count;
            }
            catch
            {
                return 0;
            }
            finally
            {
                ComHelper.Release(collection);
            }
        }

        private static bool IsPunctuationCharacter(char value)
        {
            return char.IsPunctuation(value) || char.IsSymbol(value);
        }
    }
}
