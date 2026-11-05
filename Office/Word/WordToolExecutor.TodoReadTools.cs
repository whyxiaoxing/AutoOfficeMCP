using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Office.Contracts;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    /// <summary>
    /// Word TODO 只读能力：页/段映射、跳页、样式清单、域清单。
    /// 所有页码与样式使用量都以 Word 实际回读为准；无法证明的状态只做保守标记并附 reason。
    /// </summary>
    public sealed partial class WordToolExecutor
    {
        private const int MaxCharacterStyleRunCount = 10000;

        private sealed class ParagraphPageSpan
        {
            public int ParagraphIndex { get; set; }

            public int StartPage { get; set; }

            public int EndPage { get; set; }
        }

        private void RegisterTodoReadHandlers(
            Dictionary<string, Func<Word.Application, JObject, OperationResult>> handlers)
        {
            handlers[WordToolNames.PageMap] = HandlePageMap;
            handlers[WordToolNames.GotoPage] = HandleGotoPage;
            handlers[WordToolNames.StylesList] = HandleStylesList;
            handlers[WordToolNames.FieldsList] = HandleFieldsList;
        }

        private OperationResult HandlePageMap(Word.Application application, JObject arguments)
        {
            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    int totalPages;
                    JArray pages = BuildPageMap(document, out totalPages);
                    int emptyPages = 0;
                    foreach (JToken item in pages)
                    {
                        if (item.Value<int?>("startParagraph") == 0)
                        {
                            emptyPages++;
                        }
                    }

                    JObject data = new JObject();
                    data["totalPages"] = totalPages;
                    data["pages"] = pages;
                    data["emptyPages"] = emptyPages;
                    data["pageMap"] = pages.DeepClone();
                    data["note"] = "页码由 Word 分页后的 wdActiveEndPageNumber 回读；startParagraph/endParagraph 为 1 基段号。" +
                                    "空页没有可归属段落时返回 0。";
                    return ReadSucceed(
                        "已读取页/段映射：共 " + totalPages + " 页，" +
                        (emptyPages == 0 ? "每页均有段落。" : emptyPages + " 页没有可归属段落。"),
                        data,
                        "Word:Document");
                });
        }

        private OperationResult HandleGotoPage(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int requestedPage = args.RequireInt("page", 1, int.MaxValue);

            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    int totalPages;
                    JArray pages = BuildPageMap(document, out totalPages);
                    if (requestedPage > totalPages)
                    {
                        JObject invalidData = new JObject();
                        invalidData["requestedPage"] = requestedPage;
                        invalidData["totalPages"] = totalPages;
                        invalidData["pages"] = pages.DeepClone();
                        return OperationResult.Failed(
                            OfficeErrorCode.INVALID_PARAMS,
                            "页码越界：请求第 " + requestedPage + " 页，当前文档实际共 " +
                            totalPages + " 页。",
                            null,
                            invalidData,
                            null);
                    }

                    JObject pageEntry = null;
                    foreach (JToken item in pages)
                    {
                        if (item.Value<int?>("page") == requestedPage)
                        {
                            pageEntry = item as JObject;
                            break;
                        }
                    }
                    int startParagraph = pageEntry == null
                        ? 0
                        : (pageEntry.Value<int?>("startParagraph") ?? 0);
                    int endParagraph = pageEntry == null
                        ? 0
                        : (pageEntry.Value<int?>("endParagraph") ?? 0);
                    if (startParagraph <= 0)
                    {
                        JObject emptyData = new JObject();
                        emptyData["page"] = requestedPage;
                        emptyData["totalPages"] = totalPages;
                        emptyData["startParagraph"] = 0;
                        emptyData["endParagraph"] = 0;
                        return OperationResult.Failed(
                            OfficeErrorCode.INVALID_PARAMS,
                            "第 " + requestedPage + " 页没有可定位的段落（可能是空白页）。" +
                            "当前文档共 " + totalPages + " 页。",
                            null,
                            emptyData,
                            null);
                    }

                    Word.Paragraph paragraph = null;
                    Word.Range paragraphRange = null;
                    Word.Selection selection = null;
                    Word.Range selectedRange = null;
                    try
                    {
                        paragraph = WordComUtil.GetParagraph(document, startParagraph);
                        paragraphRange = paragraph.Range;
                        paragraphRange.Select();
                        selection = application.Selection;
                        selectedRange = selection == null ? null : selection.Range;
                        int selectedParagraph = selectedRange == null
                            ? 0
                            : WordComUtil.FindParagraphIndexAt(document, selectedRange.Start);
                        JObject selectionSnapshot = WordComUtil.GetSelectionSnapshot(application, document);

                        JObject data = new JObject();
                        data["page"] = requestedPage;
                        data["totalPages"] = totalPages;
                        data["startParagraph"] = startParagraph;
                        data["endParagraph"] = endParagraph;
                        data["selectedParagraphIndex"] = selectedParagraph;
                        data["selectionStart"] = selectedRange == null ? 0 : selectedRange.Start;
                        data["selectionEnd"] = selectedRange == null ? 0 : selectedRange.End;
                        data["selectionTextPreview"] = selectedRange == null
                            ? string.Empty
                            : WordComUtil.NormalizePreview(selectedRange.Text, 120);
                        data["selection"] = selectionSnapshot;
                        return ReadSucceed(
                            "已选中第 " + requestedPage + " 页首段（第 " + startParagraph +
                            " 段，共 " + totalPages + " 页）。",
                            data,
                            "Word:Paragraph:" + startParagraph,
                            "Word:Selection");
                    }
                    finally
                    {
                        ComHelper.Release(selectedRange, selection, paragraphRange, paragraph);
                    }
                });
        }

        /// <summary>
        /// 构建页 ↔ 段映射。页码来自 Word 分页后的 Range.Information；
        /// 同时保留页统计值，避免仅凭段落位置猜测总页数。
        /// </summary>
        private static JArray BuildPageMap(Word.Document document, out int totalPages)
        {
            try
            {
                document.Repaginate();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[word] 分页刷新失败，继续使用当前分页：" + ex.Message);
            }

            totalPages = 0;
            try
            {
                totalPages = WordComUtil.ComputeStatistics(
                    document, Word.WdStatistic.wdStatisticPages);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[word] 读取总页数失败：" + ex.Message);
            }

            List<ParagraphPageSpan> spans = new List<ParagraphPageSpan>();
            Word.Paragraphs paragraphs = null;
            int maxObservedPage = 0;
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
                        int startPage = GetPageAtPosition(document, range.Start);
                        int endPage = GetPageAtPosition(
                            document, Math.Max(range.Start, range.End - 1));
                        if (startPage <= 0)
                        {
                            startPage = endPage;
                        }
                        if (endPage <= 0)
                        {
                            endPage = startPage;
                        }
                        if (startPage <= 0 && endPage <= 0)
                        {
                            continue;
                        }
                        if (startPage <= 0)
                        {
                            startPage = 1;
                        }
                        if (endPage < startPage)
                        {
                            endPage = startPage;
                        }
                        maxObservedPage = Math.Max(maxObservedPage, Math.Max(startPage, endPage));
                        ParagraphPageSpan span = new ParagraphPageSpan();
                        span.ParagraphIndex = index;
                        span.StartPage = startPage;
                        span.EndPage = endPage;
                        spans.Add(span);
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[word] 第 " + index + " 段页码读取失败：" + ex.Message);
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

            totalPages = Math.Max(Math.Max(totalPages, maxObservedPage), 1);
            int[] startParagraphs = new int[totalPages + 1];
            int[] endParagraphs = new int[totalPages + 1];
            foreach (ParagraphPageSpan span in spans)
            {
                int first = Math.Max(1, Math.Min(totalPages, span.StartPage));
                int last = Math.Max(1, Math.Min(totalPages, span.EndPage));
                for (int page = first; page <= last; page++)
                {
                    if (startParagraphs[page] == 0)
                    {
                        startParagraphs[page] = span.ParagraphIndex;
                    }
                    endParagraphs[page] = span.ParagraphIndex;
                }
            }

            JArray pages = new JArray();
            for (int page = 1; page <= totalPages; page++)
            {
                pages.Add(new JObject
                {
                    { "page", page },
                    { "startParagraph", startParagraphs[page] },
                    { "endParagraph", endParagraphs[page] }
                });
            }
            return pages;
        }

        private static int GetPageAtPosition(Word.Document document, int position)
        {
            Word.Range range = null;
            try
            {
                range = WordComUtil.CreateRange(document, position, position);
                object page = range.get_Information(Word.WdInformation.wdActiveEndPageNumber);
                return page == null
                    ? 0
                    : Convert.ToInt32(page, CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[word] 字符位置 " + position + " 的页码读取失败：" + ex.Message);
                return 0;
            }
            finally
            {
                ComHelper.Release(range);
            }
        }

        private OperationResult HandleStylesList(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string requestedType = (args.GetOptionalString("type", "all") ?? "all")
                .Trim()
                .ToLowerInvariant();
            if (requestedType == string.Empty)
            {
                requestedType = "all";
            }
            if (requestedType == "段落")
            {
                requestedType = "paragraph";
            }
            else if (requestedType == "字符")
            {
                requestedType = "character";
            }
            else if (requestedType == "表格")
            {
                requestedType = "table";
            }
            if (requestedType != "all" && requestedType != "paragraph" &&
                requestedType != "character" && requestedType != "table")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "type 只支持 all/paragraph/character/table，收到「" + requestedType + "」。");
            }

            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    Dictionary<string, int> paragraphUsage =
                        CountParagraphStyleUsage(document);
                    Dictionary<string, int> tableUsage = CountTableStyleUsage(document);
                    Word.Styles styles = null;
                    JArray items = new JArray();
                    JArray warnings = new JArray();
                    int allStyleCount = 0;
                    bool characterCountCapped = false;
                    try
                    {
                        styles = document.Styles;
                        allStyleCount = styles == null ? 0 : styles.Count;
                        for (int index = 1; index <= allStyleCount; index++)
                        {
                            object rawStyle = null;
                            Word.Style style = null;
                            try
                            {
                                rawStyle = styles[index];
                                style = rawStyle as Word.Style;
                                if (style == null)
                                {
                                    continue;
                                }
                                string styleType = WordStyleTypeName(style.Type);
                                if (requestedType != "all" && requestedType != styleType)
                                {
                                    continue;
                                }

                                string name = style.NameLocal ?? string.Empty;
                                bool inUse = false;
                                try
                                {
                                    inUse = style.InUse;
                                }
                                catch (Exception ex)
                                {
                                    warnings.Add("样式「" + name + "」InUse 读取失败：" + ex.Message);
                                }

                                int inUseCount = 0;
                                string inUseCountKind = "none";
                                if (styleType == "paragraph")
                                {
                                    if (!string.IsNullOrWhiteSpace(name))
                                    {
                                        paragraphUsage.TryGetValue(name, out inUseCount);
                                    }
                                    inUseCountKind = "paragraphs";
                                }
                                else if (styleType == "table")
                                {
                                    if (!string.IsNullOrWhiteSpace(name))
                                    {
                                        tableUsage.TryGetValue(name, out inUseCount);
                                    }
                                    inUseCountKind = "tables";
                                }
                                else if (styleType == "character" && inUse)
                                {
                                    bool capped;
                                    inUseCount = CountCharacterStyleRuns(
                                        document, style, out capped);
                                    characterCountCapped = characterCountCapped || capped;
                                    inUseCountKind = capped ? "characterRunsCapped" : "characterRuns";
                                }

                                JObject item = new JObject();
                                item["name"] = name;
                                item["type"] = styleType;
                                item["typeCode"] = (int)style.Type;
                                item["inUse"] = inUse;
                                item["inUseCount"] = inUseCount;
                                item["inUseCountKind"] = inUseCountKind;
                                item["basedOn"] = ReadBaseStyleName(style);
                                item["builtIn"] = ReadStyleBuiltIn(style);
                                item["hidden"] = ReadStyleHidden(style);
                                items.Add(item);
                            }
                            catch (Exception ex)
                            {
                                warnings.Add("样式 #" + index + " 读取失败：" + ex.Message);
                            }
                            finally
                            {
                                if (style != null)
                                {
                                    ComHelper.Release(style);
                                }
                                else
                                {
                                    ComHelper.Release(rawStyle);
                                }
                            }
                        }
                    }
                    finally
                    {
                        ComHelper.Release(styles);
                    }

                    JObject data = new JObject();
                    data["total"] = items.Count;
                    data["returned"] = items.Count;
                    data["allStyleCount"] = allStyleCount;
                    data["requestedType"] = requestedType;
                    data["styles"] = items;
                    data["items"] = items.DeepClone();
                    data["warnings"] = warnings;
                    data["characterCountCapped"] = characterCountCapped;
                    data["inUseCountNote"] = "paragraph 样式按实际段落数统计，table 样式按实际表格数统计，" +
                                             "character 样式按 Word Find 命中的连续字符样式运行段统计；" +
                                             "list/其他样式不提供使用段计数。";
                    return ReadSucceed(
                        "读取 " + items.Count + " 个样式（文档共 " + allStyleCount +
                        " 个，过滤 " + requestedType + "）。",
                        data,
                        "Word:Styles");
                });
        }

        private static Dictionary<string, int> CountParagraphStyleUsage(Word.Document document)
        {
            Dictionary<string, int> counts =
                new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
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
                        string name = WordComUtil.ReadStyleName(range);
                        if (string.IsNullOrWhiteSpace(name))
                        {
                            continue;
                        }
                        int current;
                        counts.TryGetValue(name, out current);
                        counts[name] = current + 1;
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[word] 第 " + index + " 段样式统计失败：" + ex.Message);
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
            return counts;
        }

        private static Dictionary<string, int> CountTableStyleUsage(Word.Document document)
        {
            Dictionary<string, int> counts =
                new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            Word.Tables tables = null;
            try
            {
                tables = document.Tables;
                int count = tables == null ? 0 : tables.Count;
                for (int index = 1; index <= count; index++)
                {
                    Word.Table table = null;
                    try
                    {
                        table = tables[index];
                        string name = WordComUtil.ReadTableStyleName(table);
                        if (string.IsNullOrWhiteSpace(name))
                        {
                            continue;
                        }
                        int current;
                        counts.TryGetValue(name, out current);
                        counts[name] = current + 1;
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[word] 第 " + index + " 个表格样式统计失败：" + ex.Message);
                    }
                    finally
                    {
                        ComHelper.Release(table);
                    }
                }
            }
            finally
            {
                ComHelper.Release(tables);
            }
            return counts;
        }

        /// <summary>
        /// 用 Word Find 按样式查找文本运行段，统计字符样式实际使用次数。
        /// 每次从上一个命中末尾继续，避免同一段被重复计数；上限防止损坏文档导致无限查找。
        /// </summary>
        private static int CountCharacterStyleRuns(
            Word.Document document,
            Word.Style style,
            out bool capped)
        {
            capped = false;
            int count = 0;
            Word.Range content = null;
            try
            {
                content = document.Content;
                int contentEnd = content.End;
                int position = content.Start;
                while (position < contentEnd && count < MaxCharacterStyleRunCount)
                {
                    Word.Range searchRange = null;
                    Word.Find find = null;
                    try
                    {
                        searchRange = WordComUtil.CreateRange(document, position, contentEnd);
                        find = searchRange.Find;
                        find.ClearFormatting();
                        object styleObject = style;
                        find.set_Style(ref styleObject);
                        find.Text = string.Empty;
                        find.Format = true;
                        find.Forward = true;
                        find.Wrap = Word.WdFindWrap.wdFindStop;
                        find.MatchCase = false;
                        find.MatchWholeWord = false;
                        find.MatchWildcards = false;
                        object missing = Type.Missing;
                        bool found = find.Execute(
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing);
                        if (!found)
                        {
                            break;
                        }
                        count++;
                        int nextPosition = searchRange.End;
                        if (nextPosition <= position)
                        {
                            break;
                        }
                        position = nextPosition;
                    }
                    finally
                    {
                        ComHelper.Release(find, searchRange);
                    }
                }
                if (count >= MaxCharacterStyleRunCount)
                {
                    capped = true;
                }
                return count;
            }
            finally
            {
                ComHelper.Release(content);
            }
        }

        private static string WordStyleTypeName(Word.WdStyleType type)
        {
            switch (type)
            {
                case Word.WdStyleType.wdStyleTypeParagraph:
                    return "paragraph";
                case Word.WdStyleType.wdStyleTypeCharacter:
                    return "character";
                case Word.WdStyleType.wdStyleTypeTable:
                    return "table";
                case Word.WdStyleType.wdStyleTypeList:
                    return "list";
                default:
                    return "other";
            }
        }

        private static string ReadBaseStyleName(Word.Style style)
        {
            if (style == null)
            {
                return string.Empty;
            }
            object rawBase = null;
            Word.Style baseStyle = null;
            try
            {
                rawBase = style.get_BaseStyle();
                if (rawBase == null)
                {
                    return string.Empty;
                }
                baseStyle = rawBase as Word.Style;
                if (baseStyle != null)
                {
                    return baseStyle.NameLocal ?? string.Empty;
                }
                return rawBase.ToString();
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                if (baseStyle != null)
                {
                    ComHelper.Release(baseStyle);
                }
                else
                {
                    ComHelper.Release(rawBase);
                }
            }
        }

        private static bool ReadStyleBuiltIn(Word.Style style)
        {
            if (style == null)
            {
                return false;
            }
            try
            {
                return style.BuiltIn;
            }
            catch
            {
                return false;
            }
        }

        private static bool ReadStyleHidden(Word.Style style)
        {
            if (style == null)
            {
                return false;
            }
            try
            {
                return style.Hidden;
            }
            catch
            {
                return false;
            }
        }

        private OperationResult HandleFieldsList(Word.Application application, JObject arguments)
        {
            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    int totalPages = 0;
                    try
                    {
                        totalPages = WordComUtil.ComputeStatistics(
                            document, Word.WdStatistic.wdStatisticPages);
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[word] 域清单读取总页数失败：" + ex.Message);
                    }

                    JArray fields = new JArray();
                    JArray warnings = new JArray();
                    CollectDocumentFields(document, totalPages, fields, warnings);
                    JObject data = new JObject();
                    data["total"] = fields.Count;
                    data["returned"] = fields.Count;
                    data["totalPages"] = totalPages;
                    data["fields"] = fields;
                    data["items"] = fields.DeepClone();
                    data["warnings"] = warnings;
                    data["outdatedNote"] = "outdated 只在结果可证明异常时标记：错误占位文本、" +
                                           "页码/总页数不是有效整数或越界、正文 PAGE 域结果与所在页不一致。" +
                                           "页眉页脚中的 PAGE 域无法在不逐页实例化的情况下证明过期，只检查结果格式与范围。";
                    return ReadSucceed(
                        "读取 " + fields.Count + " 个域（含正文/页眉页脚等故事范围）。",
                        data,
                        "Word:Fields");
                });
        }

        private static void CollectDocumentFields(
            Word.Document document,
            int totalPages,
            JArray fields,
            JArray warnings)
        {
            Word.StoryRanges stories = null;
            try
            {
                stories = document.StoryRanges;
                if (stories == null)
                {
                    return;
                }
                foreach (Word.Range story in stories)
                {
                    Word.Range current = story;
                    while (current != null)
                    {
                        try
                        {
                            CollectStoryFields(document, current, totalPages, fields, warnings);
                        }
                        catch (Exception ex)
                        {
                            warnings.Add("故事范围 " + current.StoryType + " 的域读取失败：" + ex.Message);
                        }

                        Word.Range next = null;
                        try
                        {
                            next = current.NextStoryRange;
                        }
                        catch (Exception ex)
                        {
                            warnings.Add("故事范围 " + current.StoryType + " 的下一段读取失败：" + ex.Message);
                        }
                        ComHelper.Release(current);
                        current = next;
                    }
                }
            }
            finally
            {
                ComHelper.Release(stories);
            }
        }

        private static void CollectStoryFields(
            Word.Document document,
            Word.Range storyRange,
            int totalPages,
            JArray fields,
            JArray warnings)
        {
            Word.Fields storyFields = null;
            try
            {
                storyFields = storyRange.Fields;
                int count = storyFields == null ? 0 : storyFields.Count;
                for (int index = 1; index <= count; index++)
                {
                    Word.Field field = null;
                    Word.Range codeRange = null;
                    Word.Range resultRange = null;
                    try
                    {
                        field = storyFields[index];
                        codeRange = field.Code;
                        resultRange = field.Result;
                        string code = WordComUtil.CleanParagraphText(
                            codeRange == null ? string.Empty : codeRange.Text);
                        string result = WordComUtil.CleanParagraphText(
                            resultRange == null ? string.Empty : resultRange.Text);
                        Word.WdFieldType fieldType = field.Type;
                        Word.WdStoryType storyType = storyRange.StoryType;
                        int actualPage = 0;
                        if (fieldType == Word.WdFieldType.wdFieldPage &&
                            storyType == Word.WdStoryType.wdMainTextStory &&
                            resultRange != null)
                        {
                            actualPage = GetPageAtPosition(document, resultRange.Start);
                        }

                        bool outdated;
                        string reason;
                        DetermineFieldOutdated(
                            fieldType, storyType, result, totalPages, actualPage,
                            out outdated, out reason);

                        JObject item = new JObject();
                        item["index"] = fields.Count + 1;
                        item["type"] = FieldTypeName(fieldType);
                        item["fieldType"] = fieldType.ToString();
                        item["typeCode"] = (int)fieldType;
                        item["code"] = code;
                        item["result"] = result;
                        item["locked"] = field.Locked;
                        item["outdated"] = outdated;
                        item["outdatedReason"] = reason ?? string.Empty;
                        item["fieldIndex"] = field.Index;
                        item["storyType"] = StoryTypeName(storyType);
                        item["actualPage"] = actualPage > 0
                            ? new JValue(actualPage)
                            : JValue.CreateNull();
                        fields.Add(item);
                    }
                    catch (Exception ex)
                    {
                        warnings.Add("故事 " + storyRange.StoryType + " 的第 " + index +
                                     " 个域读取失败：" + ex.Message);
                    }
                    finally
                    {
                        ComHelper.Release(resultRange, codeRange, field);
                    }
                }
            }
            finally
            {
                ComHelper.Release(storyFields);
            }
        }

        /// <summary>
        /// outdated 判定依据（只标记可证明异常，不猜测 Word 是否需要刷新）：
        /// 1) 结果文本含 Word 已知错误占位（错误！/Error! Reference source not found 等）；
        /// 2) PAGE/NUMPAGES/PAGEREF 的结果不是正整数，或页码超出文档实际页数；
        /// 3) 正文中的 PAGE 域结果与域所在位置的实际页码不一致；
        /// 4) REF/PAGEREF/NOTEREF/TOC 等引用域结果为空时视为可疑（Word 未更新/引用源缺失）。
        /// 页眉页脚中的 PAGE 域会随页面实例变化，无法用单一结果与页码比较，故只做 1、2 类判定。
        /// </summary>
        private static void DetermineFieldOutdated(
            Word.WdFieldType fieldType,
            Word.WdStoryType storyType,
            string result,
            int totalPages,
            int actualPage,
            out bool outdated,
            out string reason)
        {
            outdated = false;
            reason = string.Empty;
            string text = (result ?? string.Empty).Trim();
            string lower = text.ToLowerInvariant();
            string[] errorTokens = new string[]
            {
                "错误！",
                "错误!",
                "error!",
                "reference source not found",
                "未找到引用源",
                "未定义书签",
                "bookmark not defined",
                "无法找到引用源"
            };
            foreach (string token in errorTokens)
            {
                if (lower.IndexOf(token, StringComparison.Ordinal) >= 0)
                {
                    outdated = true;
                    reason = "结果包含 Word 错误占位文本「" + token + "」。";
                    return;
                }
            }

            bool pageRelated = fieldType == Word.WdFieldType.wdFieldPage ||
                               fieldType == Word.WdFieldType.wdFieldNumPages ||
                               fieldType == Word.WdFieldType.wdFieldPageRef ||
                               fieldType == Word.WdFieldType.wdFieldSectionPages;
            int numericResult;
            bool hasNumber = int.TryParse(
                text, NumberStyles.Integer, CultureInfo.InvariantCulture, out numericResult);
            if (pageRelated && !hasNumber)
            {
                outdated = true;
                reason = "页码类域结果为「" + WordComUtil.Truncate(text, 40) + "」，不是有效整数。";
                return;
            }

            if (hasNumber && totalPages > 0 && (numericResult < 1 || numericResult > totalPages))
            {
                outdated = true;
                reason = "页码类域结果 " + numericResult + " 超出文档实际页数 1-" + totalPages + "。";
                return;
            }

            if (fieldType == Word.WdFieldType.wdFieldNumPages &&
                hasNumber && totalPages > 0 && numericResult != totalPages)
            {
                outdated = true;
                reason = "NUMPAGES 结果 " + numericResult + " 与实际总页数 " + totalPages + " 不一致。";
                return;
            }

            if (fieldType == Word.WdFieldType.wdFieldPage &&
                storyType == Word.WdStoryType.wdMainTextStory &&
                hasNumber && actualPage > 0 && numericResult != actualPage)
            {
                outdated = true;
                reason = "PAGE 结果 " + numericResult + " 与域所在实际页 " + actualPage + " 不一致。";
                return;
            }

            bool referenceField = fieldType == Word.WdFieldType.wdFieldRef ||
                                  fieldType == Word.WdFieldType.wdFieldPageRef ||
                                  fieldType == Word.WdFieldType.wdFieldNoteRef ||
                                  fieldType == Word.WdFieldType.wdFieldTOC;
            if (referenceField && text.Length == 0)
            {
                outdated = true;
                reason = "引用/目录域结果为空，Word 可能尚未更新或引用源缺失。";
            }
        }

        private static string FieldTypeName(Word.WdFieldType type)
        {
            switch (type)
            {
                case Word.WdFieldType.wdFieldPage:
                    return "PAGE";
                case Word.WdFieldType.wdFieldNumPages:
                    return "NUMPAGES";
                case Word.WdFieldType.wdFieldPageRef:
                    return "PAGEREF";
                case Word.WdFieldType.wdFieldSectionPages:
                    return "SECTIONPAGES";
                case Word.WdFieldType.wdFieldRef:
                    return "REF";
                case Word.WdFieldType.wdFieldNoteRef:
                    return "NOTEREF";
                case Word.WdFieldType.wdFieldTOC:
                    return "TOC";
                case Word.WdFieldType.wdFieldSequence:
                    return "SEQ";
                case Word.WdFieldType.wdFieldHyperlink:
                    return "HYPERLINK";
                case Word.WdFieldType.wdFieldDate:
                    return "DATE";
                case Word.WdFieldType.wdFieldTime:
                    return "TIME";
                case Word.WdFieldType.wdFieldFileName:
                    return "FILENAME";
                case Word.WdFieldType.wdFieldAuthor:
                    return "AUTHOR";
                case Word.WdFieldType.wdFieldTitle:
                    return "TITLE";
                case Word.WdFieldType.wdFieldSubject:
                    return "SUBJECT";
                case Word.WdFieldType.wdFieldFormula:
                    return "FORMULA";
                case Word.WdFieldType.wdFieldIf:
                    return "IF";
                case Word.WdFieldType.wdFieldIncludeText:
                    return "INCLUDETEXT";
                default:
                    return type.ToString().Replace("wdField", string.Empty).ToUpperInvariant();
            }
        }

        private static string StoryTypeName(Word.WdStoryType storyType)
        {
            switch (storyType)
            {
                case Word.WdStoryType.wdMainTextStory:
                    return "mainText";
                case Word.WdStoryType.wdPrimaryHeaderStory:
                    return "primaryHeader";
                case Word.WdStoryType.wdPrimaryFooterStory:
                    return "primaryFooter";
                case Word.WdStoryType.wdFirstPageHeaderStory:
                    return "firstPageHeader";
                case Word.WdStoryType.wdFirstPageFooterStory:
                    return "firstPageFooter";
                case Word.WdStoryType.wdEvenPagesHeaderStory:
                    return "evenPagesHeader";
                case Word.WdStoryType.wdEvenPagesFooterStory:
                    return "evenPagesFooter";
                case Word.WdStoryType.wdFootnotesStory:
                    return "footnotes";
                case Word.WdStoryType.wdEndnotesStory:
                    return "endnotes";
                case Word.WdStoryType.wdCommentsStory:
                    return "comments";
                case Word.WdStoryType.wdTextFrameStory:
                    return "textFrame";
                default:
                    return storyType.ToString().Replace("wd", string.Empty);
            }
        }
    }
}
