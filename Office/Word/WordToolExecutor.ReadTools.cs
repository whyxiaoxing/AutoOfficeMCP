using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    public sealed partial class WordToolExecutor
    {
        private void RegisterReadHandlers(
            Dictionary<string, Func<Word.Application, JObject, OperationResult>> handlers)
        {
            handlers[WordToolNames.GetContext] = HandleGetContext;
            handlers[WordToolNames.ListParagraphs] = HandleListParagraphs;
            handlers[WordToolNames.GetParagraphInfo] = HandleGetParagraphInfo;
            handlers[WordToolNames.FindAndSelect] = HandleFindAndSelect;
            handlers[WordToolNames.ReadSelection] = HandleReadSelection;
            handlers[WordToolNames.DocumentStats] = HandleDocumentStats;
        }

        private OperationResult HandleGetContext(Word.Application application, JObject arguments)
        {
            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    JObject selection = WordComUtil.GetSelectionSnapshot(application, document);
                    int currentIndex = selection.Value<int?>("paragraphIndex") ?? 0;
                    Word.Paragraphs paragraphs = null;
                    try
                    {
                        paragraphs = document.Paragraphs;
                        int total = paragraphs == null ? 0 : paragraphs.Count;
                        JArray headings = new JArray();
                        for (int index = 1; index <= total && headings.Count < 10; index++)
                        {
                            Word.Paragraph paragraph = null;
                            Word.Range range = null;
                            try
                            {
                                paragraph = paragraphs[index];
                                range = paragraph.Range;
                                string style = WordComUtil.ReadStyleName(range);
                                if (IsHeadingStyle(style))
                                {
                                    headings.Add(new JObject
                                    {
                                        { "index", index },
                                        { "style", style },
                                        { "textPreview", WordComUtil.NormalizePreview(range.Text, 50) }
                                    });
                                }
                            }
                            finally
                            {
                                ComHelper.Release(range, paragraph);
                            }
                        }

                        JArray context = new JArray();
                        if (currentIndex > 0 && total > 0)
                        {
                            int first = Math.Max(1, currentIndex - 2);
                            int last = Math.Min(total, currentIndex + 2);
                            for (int index = first; index <= last; index++)
                            {
                                Word.Paragraph paragraph = null;
                                Word.Range range = null;
                                try
                                {
                                    paragraph = paragraphs[index];
                                    range = paragraph.Range;
                                    context.Add(new JObject
                                    {
                                        { "index", index },
                                        { "textPreview", WordComUtil.NormalizePreview(range.Text, 40) },
                                        { "style", WordComUtil.ReadStyleName(range) },
                                        { "isCurrent", index == currentIndex }
                                    });
                                }
                                finally
                                {
                                    ComHelper.Release(range, paragraph);
                                }
                            }
                        }

                        JObject structure = new JObject();
                        structure["totalParagraphs"] = total;
                        structure["headings"] = headings;
                        structure["headingLimitNote"] = "最多返回 10 个标题样式段落；完整列表请用 word_list_paragraphs。";
                        structure["currentParagraphIndex"] = currentIndex;

                        JObject data = new JObject();
                        data["document"] = new JObject
                        {
                            { "name", document.Name },
                            { "path", document.FullName },
                            { "readOnly", document.ReadOnly }
                        };
                        data["selection"] = selection;
                        data["structure"] = structure;
                        data["contextParagraphs"] = context;
                        return ReadSucceed(
                            "已读取 Word 上下文：选区 " +
                            (selection.Value<bool?>("exists") == true
                                ? selection.Value<int?>("textLength").ToString() + " 字符"
                                : "无效") +
                            "，文档共 " + total + " 段，标题 " + headings.Count + " 个。",
                            data, "Word:Selection", "Word:Document");
                    }
                    finally
                    {
                        ComHelper.Release(paragraphs);
                    }
                });
        }

        private OperationResult HandleListParagraphs(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int maxCount = args.GetOptionalInt("maxCount", 50);
            if (maxCount < 1 || maxCount > 2000)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "maxCount 允许 1-2000，收到 " + maxCount + "。");
            }

            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    Word.Paragraphs paragraphs = null;
                    JArray items = new JArray();
                    try
                    {
                        paragraphs = document.Paragraphs;
                        int total = paragraphs == null ? 0 : paragraphs.Count;
                        int returned = Math.Min(total, maxCount);
                        for (int index = 1; index <= returned; index++)
                        {
                            Word.Paragraph paragraph = null;
                            Word.Range range = null;
                            Word.Font font = null;
                            try
                            {
                                paragraph = paragraphs[index];
                                range = paragraph.Range;
                                font = range.Font;
                                string text = WordComUtil.CleanParagraphText(range.Text);
                                items.Add(new JObject
                                {
                                    { "index", index },
                                    { "summary", WordComUtil.NormalizePreview(text, 60) },
                                    { "length", text.Length },
                                    { "style", WordComUtil.ReadStyleName(range) },
                                    { "fontSize", WordFormatUtil.IsUndefined(font.Size)
                                        ? JValue.CreateNull()
                                        : new JValue((double)font.Size) }
                                });
                            }
                            finally
                            {
                                ComHelper.Release(font, range, paragraph);
                            }
                        }
                        JObject data = new JObject();
                        data["items"] = items;
                        data["total"] = total;
                        data["returned"] = returned;
                        data["truncated"] = returned < total;
                        data["maxCount"] = maxCount;
                        return ReadSucceed(
                            "读取 " + returned + "/" + total + " 个段落" +
                            (returned < total ? "（已截断）" : "") + "。",
                            data, "Word:Document");
                    }
                    finally
                    {
                        ComHelper.Release(paragraphs);
                    }
                });
        }

        private OperationResult HandleGetParagraphInfo(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int paragraphIndex = args.RequireInt("paragraphIndex", 1, int.MaxValue);

            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    WordComUtil.ValidateParagraphIndex(document, paragraphIndex);
                    Word.Paragraph paragraph = null;
                    Word.Range range = null;
                    Word.Font font = null;
                    Word.ParagraphFormat format = null;
                    Word.ListFormat listFormat = null;
                    try
                    {
                        paragraph = WordComUtil.GetParagraph(document, paragraphIndex);
                        range = paragraph.Range;
                        font = range.Font;
                        format = range.ParagraphFormat;
                        listFormat = range.ListFormat;
                        JObject info = new JObject();
                        info["index"] = paragraphIndex;
                        info["text"] = WordComUtil.CleanParagraphText(range.Text);
                        info["style"] = WordComUtil.ReadStyleName(range);
                        info["outlineLevel"] = paragraph.OutlineLevel.ToString();
                        info["fontSize"] = WordFormatUtil.IsUndefined(font.Size)
                            ? JValue.CreateNull()
                            : new JValue((double)font.Size);
                        info["fontName"] = font.Name ?? string.Empty;
                        info["fontNameFarEast"] = font.NameFarEast ?? string.Empty;
                        info["bold"] = WordFormatUtil.TriStateToken(font.Bold);
                        info["italic"] = WordFormatUtil.TriStateToken(font.Italic);
                        info["underline"] = font.Underline != Word.WdUnderline.wdUnderlineNone;
                        info["underlineStyle"] = WordFormatUtil.GetUnderlineStyleName(font.Underline);
                        info["color"] = (int)font.Color;
                        info["colorHex"] = WordFormatUtil.ColorToHex((int)font.Color);
                        info["subscript"] = WordFormatUtil.TriStateToken(font.Subscript);
                        info["superscript"] = WordFormatUtil.TriStateToken(font.Superscript);
                        info["alignment"] = WordFormatUtil.GetAlignmentName(format.Alignment);
                        info["firstLineIndentPt"] = (double)format.FirstLineIndent;
                        info["characterUnitFirstLineIndent"] = (double)format.CharacterUnitFirstLineIndent;
                        info["spaceBeforePt"] = (double)format.SpaceBefore;
                        info["spaceAfterPt"] = (double)format.SpaceAfter;
                        info["lineSpacingRule"] = format.LineSpacingRule.ToString();
                        info["lineSpacingPt"] = (double)format.LineSpacing;
                        info["listType"] = listFormat == null ? string.Empty : listFormat.ListType.ToString();
                        info["listString"] = listFormat == null ? string.Empty : (listFormat.ListString ?? string.Empty);
                        info["listValue"] = listFormat == null ? 0 : listFormat.ListValue;
                        JObject data = new JObject();
                        data["paragraph"] = info;
                        return ReadSucceed(
                            "读取第 " + paragraphIndex + " 段信息。", data,
                            "Word:Paragraph:" + paragraphIndex);
                    }
                    finally
                    {
                        ComHelper.Release(listFormat, format, font, range, paragraph);
                    }
                });
        }

        private OperationResult HandleFindAndSelect(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string query = args.RequireString("query");

            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    Match explicitParagraph = Regex.Match(query ?? string.Empty, @"第\s*(\d+)\s*段");
                    if (explicitParagraph.Success)
                    {
                        int requestedIndex;
                        if (int.TryParse(explicitParagraph.Groups[1].Value,
                            NumberStyles.Integer, CultureInfo.InvariantCulture, out requestedIndex))
                        {
                            int total = WordComUtil.GetParagraphCount(document);
                            if (requestedIndex < 1 || requestedIndex > total)
                            {
                                return OperationResult.Failed(
                                    OfficeErrorCode.INVALID_PARAMS,
                                    "第 " + requestedIndex + " 段不存在（当前共 " + total +
                                    " 段）。请先用 word_list_paragraphs 读取真实段号。");
                            }
                        }
                    }

                    WordLocateResult located = Locate(document, application, query);
                    if (located == null)
                    {
                        return OperationResult.Failed(
                            OfficeErrorCode.INVALID_PARAMS,
                            "未找到匹配「" + WordComUtil.Truncate(query, 60) + "」的内容。" +
                            "可用 query：标题1、第3段、下一段、上一段、包含“关键词”的段落，或直接输入要查找的文本。");
                    }

                    Word.Range foundRange = null;
                    Word.Selection selection = null;
                    try
                    {
                        foundRange = WordComUtil.CreateRange(document, located.Start, located.End);
                        foundRange.Select();
                        selection = application.Selection;
                        JObject selectionSnapshot = WordComUtil.GetSelectionSnapshot(application, document);
                        JObject data = new JObject();
                        data["query"] = query;
                        data["matchKind"] = located.Kind;
                        data["paragraphIndex"] = located.ParagraphIndex > 0
                            ? new JValue(located.ParagraphIndex)
                            : JValue.CreateNull();
                        data["rangeStart"] = located.Start;
                        data["rangeEnd"] = located.End;
                        data["textPreview"] = located.Preview;
                        data["selection"] = selectionSnapshot;
                        return ReadSucceed(
                            "已选中" + located.Kind + "（第 " + located.ParagraphIndex + " 段，" +
                            (located.End - located.Start) + " 字符）。",
                            data, "Word:Selection");
                    }
                    finally
                    {
                        ComHelper.Release(selection, foundRange);
                    }
                });
        }

        private OperationResult HandleReadSelection(Word.Application application, JObject arguments)
        {
            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    JObject selectionSnapshot = WordComUtil.GetSelectionSnapshot(application, document);
                    Word.Selection selection = null;
                    Word.Range range = null;
                    try
                    {
                        selection = application.Selection;
                        range = selection == null ? null : selection.Range;
                        string text = range == null ? string.Empty : (range.Text ?? string.Empty);
                        int start = range == null ? 0 : range.Start;
                        int end = range == null ? 0 : range.End;
                        int paragraphIndex = WordComUtil.FindParagraphIndexAt(document, start);
                        JObject data = new JObject();
                        data["text"] = WordComUtil.CleanParagraphText(text);
                        data["start"] = start;
                        data["end"] = end;
                        data["length"] = text.Length;
                        data["paragraphIndex"] = paragraphIndex;
                        data["boundaryDescription"] = "文档字符 " + start + "-" + end +
                                                      "，第 " + paragraphIndex + " 段";
                        data["format"] = selectionSnapshot["format"];
                        return ReadSucceed(
                            "读取当前选区：" + selectionSnapshot.Value<int?>("textLength") +
                            " 字符，第 " + paragraphIndex + " 段。",
                            data, "Word:Selection");
                    }
                    finally
                    {
                        ComHelper.Release(range, selection);
                    }
                });
        }

        private OperationResult HandleDocumentStats(Word.Application application, JObject arguments)
        {
            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    Word.Tables tables = null;
                    Word.InlineShapes inlineShapes = null;
                    Word.Shapes shapes = null;
                    Word.Comments comments = null;
                    Word.Revisions revisions = null;
                    Word.Fields fields = null;
                    Word.OMaths equations = null;
                    Word.Sections sections = null;
                    try
                    {
                        tables = document.Tables;
                        inlineShapes = document.InlineShapes;
                        shapes = document.Shapes;
                        comments = document.Comments;
                        revisions = document.Revisions;
                        fields = document.Fields;
                        equations = document.OMaths;
                        sections = document.Sections;
                        JObject data = new JObject();
                        data["documentName"] = document.Name;
                        data["path"] = document.FullName;
                        data["words"] = WordComUtil.ComputeStatistics(document, Word.WdStatistic.wdStatisticWords);
                        data["characters"] = WordComUtil.ComputeStatistics(document, Word.WdStatistic.wdStatisticCharacters);
                        data["charactersWithSpaces"] = WordComUtil.ComputeStatistics(
                            document, Word.WdStatistic.wdStatisticCharactersWithSpaces);
                        data["pages"] = WordComUtil.ComputeStatistics(document, Word.WdStatistic.wdStatisticPages);
                        data["paragraphs"] = WordComUtil.ComputeStatistics(document, Word.WdStatistic.wdStatisticParagraphs);
                        data["tables"] = tables == null ? 0 : tables.Count;
                        data["inlineShapes"] = inlineShapes == null ? 0 : inlineShapes.Count;
                        data["floatingShapes"] = shapes == null ? 0 : shapes.Count;
                        data["imagesApprox"] = (inlineShapes == null ? 0 : inlineShapes.Count) +
                                               (shapes == null ? 0 : shapes.Count);
                        data["comments"] = comments == null ? 0 : comments.Count;
                        data["revisions"] = revisions == null ? 0 : revisions.Count;
                        data["fields"] = fields == null ? 0 : fields.Count;
                        data["equations"] = equations == null ? 0 : equations.Count;
                        data["sections"] = sections == null ? 0 : sections.Count;
                        data["imageNote"] = "imagesApprox=InlineShapes+Shapes；Shapes 可能含非图片对象，精确类型请用 word_images_list。";
                        return ReadSucceed(
                            "文档统计：" + data["words"] + " 词，" + data["paragraphs"] + " 段，" +
                            data["tables"] + " 个表格，" + data["imagesApprox"] + " 个图片/形状对象。",
                            data, "Word:Document");
                    }
                    finally
                    {
                        ComHelper.Release(sections, equations, fields, revisions, comments,
                            shapes, inlineShapes, tables);
                    }
                });
        }

        private sealed class WordLocateResult
        {
            public int Start;
            public int End;
            public int ParagraphIndex;
            public string Kind;
            public string Preview;
        }

        private static WordLocateResult Locate(
            Word.Document document,
            Word.Application application,
            string query)
        {
            string lower = (query ?? string.Empty).Trim().ToLowerInvariant();

            Match paragraphMatch = Regex.Match(lower, @"第\s*(\d+)\s*段");
            bool headingIntent = lower.StartsWith("标题", StringComparison.Ordinal) ||
                                 lower.StartsWith("heading", StringComparison.Ordinal) ||
                                 Regex.IsMatch(lower, @"^(找|查找|定位)\s*(标题|heading)");
            if (headingIntent)
            {
                int level = 0;
                Match levelMatch = Regex.Match(lower, @"标题\s*(\d)");
                if (!levelMatch.Success)
                {
                    levelMatch = Regex.Match(lower, @"heading\s*(\d)");
                }
                if (levelMatch.Success)
                {
                    level = int.Parse(levelMatch.Groups[1].Value, CultureInfo.InvariantCulture);
                }
                WordLocateResult heading = LocateHeading(document, level);
                if (heading != null)
                {
                    return heading;
                }
            }

            if (paragraphMatch.Success)
            {
                int paragraphIndex = int.Parse(paragraphMatch.Groups[1].Value, CultureInfo.InvariantCulture);
                if (paragraphIndex < 1 || paragraphIndex > WordComUtil.GetParagraphCount(document))
                {
                    return null;
                }
                Word.Paragraph paragraph = null;
                Word.Range range = null;
                try
                {
                    paragraph = WordComUtil.GetParagraph(document, paragraphIndex);
                    range = paragraph.Range;
                    WordLocateResult result = new WordLocateResult();
                    result.Start = range.Start;
                    result.End = range.End;
                    result.ParagraphIndex = paragraphIndex;
                    result.Kind = "第 " + paragraphIndex + " 段";
                    result.Preview = WordComUtil.NormalizePreview(range.Text, 100);
                    return result;
                }
                finally
                {
                    ComHelper.Release(range, paragraph);
                }
            }

            if (lower.IndexOf("下一段", StringComparison.Ordinal) >= 0 ||
                lower.IndexOf("下一个段落", StringComparison.Ordinal) >= 0)
            {
                return LocateRelative(document, application, 1);
            }
            if (lower.IndexOf("上一段", StringComparison.Ordinal) >= 0 ||
                lower.IndexOf("前一段", StringComparison.Ordinal) >= 0 ||
                lower.IndexOf("前一个段落", StringComparison.Ordinal) >= 0)
            {
                return LocateRelative(document, application, -1);
            }

            Match containsMatch = Regex.Match(
                query ?? string.Empty, @"包含\s*[\""'“”‘’](.+?)[\""'“”‘’]");
            if (containsMatch.Success)
            {
                return LocateContaining(document, containsMatch.Groups[1].Value);
            }
            if (lower.StartsWith("包含", StringComparison.Ordinal) &&
                lower.EndsWith("的段落", StringComparison.Ordinal))
            {
                string text = (query ?? string.Empty).Substring(2, (query ?? string.Empty).Length - 5);
                return LocateContaining(document, text);
            }

            string searchText = ExtractQuotedText(query);
            if (string.IsNullOrWhiteSpace(searchText))
            {
                searchText = query;
            }
            return LocateText(document, searchText);
        }

        private static WordLocateResult LocateHeading(Word.Document document, int level)
        {
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
                        string style = WordComUtil.ReadStyleName(range);
                        if (!IsHeadingStyle(style))
                        {
                            continue;
                        }
                        string expectedChinese = "标题 " + level;
                        string expectedEnglish = "Heading " + level;
                        if (level > 0 &&
                            style.IndexOf(expectedChinese, StringComparison.OrdinalIgnoreCase) < 0 &&
                            style.IndexOf(expectedEnglish, StringComparison.OrdinalIgnoreCase) < 0 &&
                            style.IndexOf("标题" + level, StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            continue;
                        }
                        WordLocateResult result = new WordLocateResult();
                        result.Start = range.Start;
                        result.End = range.End;
                        result.ParagraphIndex = index;
                        result.Kind = "标题样式段落（" + style + "）";
                        result.Preview = WordComUtil.NormalizePreview(range.Text, 100);
                        return result;
                    }
                    finally
                    {
                        ComHelper.Release(range, paragraph);
                    }
                }
                return null;
            }
            finally
            {
                ComHelper.Release(paragraphs);
            }
        }

        private static WordLocateResult LocateRelative(
            Word.Document document,
            Word.Application application,
            int offset)
        {
            Word.Selection selection = null;
            Word.Range range = null;
            try
            {
                selection = application.Selection;
                range = selection == null ? null : selection.Range;
                if (range == null)
                {
                    return null;
                }
                int currentIndex = WordComUtil.FindParagraphIndexAt(document, range.Start);
                int targetIndex = currentIndex + offset;
                int count = WordComUtil.GetParagraphCount(document);
                if (targetIndex < 1 || targetIndex > count)
                {
                    return null;
                }
                Word.Paragraph paragraph = null;
                Word.Range paragraphRange = null;
                try
                {
                    paragraph = WordComUtil.GetParagraph(document, targetIndex);
                    paragraphRange = paragraph.Range;
                    WordLocateResult result = new WordLocateResult();
                    result.Start = paragraphRange.Start;
                    result.End = paragraphRange.End;
                    result.ParagraphIndex = targetIndex;
                    result.Kind = offset > 0 ? "下一段" : "上一段";
                    result.Preview = WordComUtil.NormalizePreview(paragraphRange.Text, 100);
                    return result;
                }
                finally
                {
                    ComHelper.Release(paragraphRange, paragraph);
                }
            }
            finally
            {
                ComHelper.Release(range, selection);
            }
        }

        private static WordLocateResult LocateContaining(Word.Document document, string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }
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
                        string paragraphText = range.Text ?? string.Empty;
                        if (paragraphText.IndexOf(text, StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            continue;
                        }
                        WordLocateResult result = new WordLocateResult();
                        result.Start = range.Start;
                        result.End = range.End;
                        result.ParagraphIndex = index;
                        result.Kind = "包含「" + WordComUtil.Truncate(text, 30) + "」的段落";
                        result.Preview = WordComUtil.NormalizePreview(paragraphText, 100);
                        return result;
                    }
                    finally
                    {
                        ComHelper.Release(range, paragraph);
                    }
                }
                return null;
            }
            finally
            {
                ComHelper.Release(paragraphs);
            }
        }

        private static WordLocateResult LocateText(Word.Document document, string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }
            Word.Range range = null;
            Word.Find find = null;
            try
            {
                range = document.Content;
                find = range.Find;
                if (!WordTextWriter.FindFirst(find, text))
                {
                    return null;
                }
                WordLocateResult result = new WordLocateResult();
                result.Start = range.Start;
                result.End = range.End;
                result.ParagraphIndex = WordComUtil.FindParagraphIndexAt(document, range.Start);
                result.Kind = "普通文本";
                result.Preview = WordComUtil.NormalizePreview(range.Text, 100);
                return result;
            }
            finally
            {
                ComHelper.Release(find, range);
            }
        }

        private static string ExtractQuotedText(string query)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                return string.Empty;
            }
            Match match = Regex.Match(query, @"[\""'“”‘’](.+?)[\""'“”‘’]");
            return match.Success ? match.Groups[1].Value : string.Empty;
        }

        private static bool IsHeadingStyle(string style)
        {
            if (string.IsNullOrWhiteSpace(style))
            {
                return false;
            }
            return style.IndexOf("标题", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   style.IndexOf("Heading", StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
