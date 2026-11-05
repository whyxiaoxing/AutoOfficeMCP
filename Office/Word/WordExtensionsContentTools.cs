using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    public sealed partial class WordExtensionsExecutor
    {
        private const string WatermarkShapePrefix = "AutoOfficeMCPWatermark_";

        private void RegisterContentExtensionHandlers(
            Dictionary<string, Func<Word.Application, JObject, OperationResult>> handlers)
        {
            handlers[WordExtensionsToolNames.CaptionXref] = HandleCaptionXref;
            handlers[WordExtensionsToolNames.Watermark] = HandleWatermark;
            handlers[WordExtensionsToolNames.FootnoteManage] = HandleFootnoteManage;
            handlers[WordExtensionsToolNames.ContentControls] = HandleContentControls;
            handlers[WordExtensionsToolNames.LineNumbering] = HandleLineNumbering;
        }

        // ------------------------------------------------------------------
        // word_caption_xref
        // ------------------------------------------------------------------

        private OperationResult HandleCaptionXref(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string action = (args.GetOptionalString("action", "add") ?? "add")
                .Trim().ToLowerInvariant();
            if (action != "add" && action != "xref")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "action 只支持 add（批量题注）/xref（插入交叉引用域），收到「" + action + "」。");
            }

            if (action == "xref")
            {
                string bookmarkName = args.RequireString("bookmarkName");
                string prefixText = args.GetOptionalString("prefixText", "见 ");
                return RunWriteTool(
                    application,
                    WordExtensionsToolNames.CaptionXref,
                    false,
                    delegate(Word.Document document)
                    {
                        Word.Bookmarks bookmarks = null;
                        try
                        {
                            bookmarks = document.Bookmarks;
                            if (bookmarks == null || !bookmarks.Exists(bookmarkName))
                            {
                                throw WordArguments.Invalid(
                                    "书签「" + bookmarkName + "」不存在。请先用 action=add 生成题注书签，" +
                                    "或传 bookmarkPrefix 对应的完整书签名。");
                            }
                        }
                        finally
                        {
                            ComHelper.Release(bookmarks);
                        }
                        return WordWriteTargets.Cursor(application);
                    },
                    delegate(Word.Document document, WordWriteTarget writeTarget)
                    {
                        Word.Selection selection = null;
                        Word.Range selectionRange = null;
                        Word.Range prefixRange = null;
                        Word.Range fieldRange = null;
                        Word.Fields fields = null;
                        try
                        {
                            selection = application.Selection;
                            if (selection == null)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INVALID_PARAMS,
                                    "当前没有有效选区，无法插入交叉引用域。请先把光标放到插入位置。",
                                    null,
                                    false);
                            }
                            selectionRange = selection.Range;
                            int beforeFields = CountFields(document);
                            prefixRange = WordComUtil.CreateRange(
                                document, selectionRange.Start, selectionRange.Start);
                            if (!string.IsNullOrEmpty(prefixText))
                            {
                                prefixRange.InsertAfter(prefixText);
                            }
                            fieldRange = WordComUtil.CreateRange(
                                document, prefixRange.End, prefixRange.End);
                            fields = document.Fields;
                            object fieldType = Word.WdFieldType.wdFieldRef;
                            object fieldText = bookmarkName + " \\h";
                            object preserveFormatting = false;
                            fields.Add(fieldRange, ref fieldType, ref fieldText,
                                ref preserveFormatting);
                            int afterFields = CountFields(document);
                            bool changed = afterFields > beforeFields;
                            JObject data = new JObject();
                            data["action"] = "xref";
                            data["bookmarkName"] = bookmarkName;
                            data["prefixText"] = prefixText ?? string.Empty;
                            data["fieldCountBefore"] = beforeFields;
                            data["fieldCountAfter"] = afterFields;
                            data["note"] = "已插入 REF 交叉引用域（\\h 超链接开关）；题注内容变化后可用 " +
                                            "word_field_update_all 更新域。";
                            if (!changed)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INTERNAL_ERROR,
                                    "Fields.Add 已返回，但域数量未增加（" + beforeFields + "→" +
                                    afterFields + "）。请确认插入位置不在受保护区域。",
                                    data,
                                    false);
                            }
                            WordWriteResult result = WordWriteResult.Ok(
                                "已在光标处插入指向书签「" + bookmarkName + "」的交叉引用域。",
                                data,
                                true);
                            result.AddTargetRef("Word:Field:REF:" + bookmarkName);
                            return result;
                        }
                        finally
                        {
                            ComHelper.Release(fields, fieldRange, prefixRange, selectionRange, selection);
                        }
                    });
            }

            string targetKind = (args.GetOptionalString("target", "tables") ?? "tables")
                .Trim().ToLowerInvariant();
            if (targetKind != "tables" && targetKind != "images" && targetKind != "both")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "target 只支持 tables/images/both，收到「" + targetKind + "」。");
            }
            string label = args.GetOptionalString("label", null);
            string captionText = args.GetOptionalString("captionText", string.Empty);
            string captionStyle = args.GetOptionalString("captionStyle", "题注");
            string bookmarkPrefix = args.GetOptionalString("bookmarkPrefix", "AO_Caption");
            bool addBookmark = args.GetOptionalBool("addBookmark", true);
            bool includeInline = args.GetOptionalBool("includeInline", true);
            bool includeFloating = args.GetOptionalBool("includeFloating", true);
            List<int> tableIndexes = args.GetOptionalIntList("tableIndexes");
            int maxItems = args.GetOptionalInt("maxItems", 200);
            if (maxItems < 1 || maxItems > 1000)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "maxItems 允许 1-1000，收到 " + maxItems + "。");
            }

            return RunWriteTool(
                application,
                WordExtensionsToolNames.CaptionXref,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.None("批量题注", "Word:Captions");
                },
                delegate(Word.Document document, WordWriteTarget writeTarget)
                {
                    List<CaptionTarget> targets = new List<CaptionTarget>();
                    JArray warnings = new JArray();
                    CollectCaptionTargets(
                        document, targetKind, tableIndexes, includeInline, includeFloating,
                        captionText,
                        string.IsNullOrWhiteSpace(label) ? null : label,
                        targets, warnings);
                    if (targets.Count == 0)
                    {
                        JObject emptyData = new JObject();
                        emptyData["action"] = "add";
                        emptyData["target"] = targetKind;
                        emptyData["captions"] = new JArray();
                        emptyData["captionCount"] = 0;
                        emptyData["warnings"] = warnings;
                        return WordWriteResult.Ok(
                            "未找到需要添加题注的" + (targetKind == "tables" ? "表格" : "图片") +
                            "；文档未修改。",
                            emptyData,
                            false);
                    }
                    if (targets.Count > maxItems)
                    {
                        targets = targets.GetRange(0, maxItems);
                        WordExtensionsUtil.AddWarning(
                            warnings, "目标数量超过 maxItems=" + maxItems + "，仅处理前 " +
                                      maxItems + " 个。");
                    }

                    Dictionary<string, int> existingCounts =
                        CountExistingCaptions(document);
                    targets.Sort(delegate(CaptionTarget left, CaptionTarget right)
                    {
                        return left.AnchorPosition.CompareTo(right.AnchorPosition);
                    });
                    Dictionary<string, int> runningCounts =
                        new Dictionary<string, int>(existingCounts, StringComparer.Ordinal);
                    foreach (CaptionTarget captionTarget in targets)
                    {
                        int running;
                        runningCounts.TryGetValue(captionTarget.Label, out running);
                        running++;
                        runningCounts[captionTarget.Label] = running;
                        captionTarget.Number = running;
                    }
                    targets.Sort(delegate(CaptionTarget left, CaptionTarget right)
                    {
                        return right.AnchorPosition.CompareTo(left.AnchorPosition);
                    });

                    string resolvedStyle = null;
                    if (WordStyleCatalog.TryGetStyleName(document, captionStyle, out resolvedStyle))
                    {
                        // 使用文档中真实存在的题注样式。
                    }
                    else
                    {
                        resolvedStyle = null;
                        WordExtensionsUtil.AddWarning(
                            warnings, "找不到题注样式「" + captionStyle + "」，题注段落不套用样式。");
                    }

                    JArray captions = new JArray();
                    JArray failures = new JArray();
                    int addedCount = 0;
                    foreach (CaptionTarget item in targets)
                    {
                        Word.Range anchorRange = null;
                        Word.Paragraph captionParagraph = null;
                        Word.Range captionRange = null;
                        Word.Range bookmarkRange = null;
                        Word.Bookmarks bookmarks = null;
                        try
                        {
                            anchorRange = ResolveCaptionAnchor(document, item);
                            if (anchorRange == null)
                            {
                                failures.Add(new JObject
                                {
                                    { "kind", item.Kind },
                                    { "index", item.Index },
                                    { "reason", "目标对象已不存在或锚点无效" }
                                });
                                continue;
                            }
                            int insertionPosition;
                            if (item.Kind == "table")
                            {
                                insertionPosition = anchorRange.End;
                            }
                            else
                            {
                                Word.Tables anchorTables = null;
                                try
                                {
                                    anchorTables = anchorRange.Tables;
                                    if (anchorTables != null && anchorTables.Count > 0)
                                    {
                                        failures.Add(new JObject
                                        {
                                            { "kind", item.Kind },
                                            { "index", item.Index },
                                            { "reason", "图片位于表格单元格内，暂不支持自动插入题注；" +
                                                        "请手动在图后添加题注段落" }
                                        });
                                        continue;
                                    }
                                }
                                finally
                                {
                                    ComHelper.Release(anchorTables);
                                }
                                int anchorParagraphIndex = WordComUtil.FindParagraphIndexAt(
                                    document, anchorRange.Start);
                                Word.Paragraph anchorParagraph = null;
                                Word.Range anchorParagraphRange = null;
                                try
                                {
                                    anchorParagraph = WordComUtil.GetParagraph(
                                        document, anchorParagraphIndex);
                                    anchorParagraphRange = anchorParagraph.Range;
                                    insertionPosition = anchorParagraphRange.End;
                                }
                                finally
                                {
                                    ComHelper.Release(anchorParagraphRange, anchorParagraph);
                                }
                            }

                            string fullCaption = item.Label + " " +
                                item.Number.ToString(CultureInfo.InvariantCulture);
                            if (!string.IsNullOrWhiteSpace(item.CaptionText))
                            {
                                fullCaption = fullCaption + " " + item.CaptionText.Trim();
                            }
                            int captionParagraphIndex;
                            string insertFailure;
                            if (!InsertCaptionParagraph(
                                document, insertionPosition, fullCaption, resolvedStyle,
                                out captionParagraphIndex, out insertFailure))
                            {
                                failures.Add(new JObject
                                {
                                    { "kind", item.Kind },
                                    { "index", item.Index },
                                    { "reason", insertFailure }
                                });
                                continue;
                            }

                            captionParagraph = WordComUtil.GetParagraph(
                                document, captionParagraphIndex);
                            captionRange = captionParagraph.Range;
                            string actual = WordExtensionsUtil.ReadParagraphText(captionParagraph);
                            if (!string.Equals(actual, fullCaption, StringComparison.Ordinal))
                            {
                                failures.Add(new JObject
                                {
                                    { "kind", item.Kind },
                                    { "index", item.Index },
                                    { "reason", "题注写入后回读不一致；期望「" +
                                                WordExtensionsUtil.Preview(fullCaption, 80) +
                                                "」，实际「" +
                                                WordExtensionsUtil.Preview(actual, 80) + "」" }
                                });
                                continue;
                            }

                            string bookmarkName = string.Empty;
                            if (addBookmark)
                            {
                                bookmarkName = WordExtensionsUtil.SafeBookmarkName(
                                    bookmarkPrefix + "_" + item.Label + "_" + item.Number);
                                bookmarkRange = WordComUtil.CreateRange(
                                    document,
                                    captionRange.Start,
                                    Math.Max(captionRange.Start, captionRange.End - 1));
                                bookmarks = document.Bookmarks;
                                object bookmarkRangeObject = bookmarkRange;
                                bookmarks.Add(bookmarkName, ref bookmarkRangeObject);
                            }

                            addedCount++;
                            JObject captionJson = new JObject();
                            captionJson["kind"] = item.Kind;
                            captionJson["index"] = item.Index;
                            captionJson["label"] = item.Label;
                            captionJson["number"] = item.Number;
                            captionJson["text"] = fullCaption;
                            captionJson["bookmark"] = bookmarkName;
                            captionJson["captionParagraphIndex"] = captionParagraphIndex;
                            captions.Add(captionJson);
                        }
                        catch (Exception ex)
                        {
                            failures.Add(new JObject
                            {
                                { "kind", item.Kind },
                                { "index", item.Index },
                                { "reason", OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) }
                            });
                        }
                        finally
                        {
                            ComHelper.Release(bookmarks, bookmarkRange, captionRange,
                                captionParagraph, anchorRange);
                        }
                    }

                    JObject data = new JObject();
                    data["action"] = "add";
                    data["target"] = targetKind;
                    data["label"] = label ?? string.Empty;
                    data["captionStyle"] = resolvedStyle ?? captionStyle;
                    data["captionCount"] = addedCount;
                    data["failureCount"] = failures.Count;
                    data["captions"] = captions;
                    data["failures"] = failures;
                    data["warnings"] = warnings;
                    data["note"] = "题注编号是插入时的快照计数（不是 SEQ 域），不会随文档后续增删自动更新；" +
                                    "addBookmark=true 时每个题注创建书签，可用 action=xref 插入 REF 交叉引用域。" +
                                    "表格题注默认插入到表格后的新段落，图片题注插入到锚定段落后的新段落。";

                    if (addedCount == 0)
                    {
                        return WordWriteResult.Fail(
                            OfficeErrorCode.INTERNAL_ERROR,
                            "批量题注全部失败（" + failures.Count + " 个目标）。首个原因：" +
                            (failures.Count > 0 ? failures[0]["reason"].ToString() : "未知") +
                            "。请确认文档未受保护、目标表格/图片仍存在。",
                            data,
                            false);
                    }
                    if (failures.Count > 0)
                    {
                        return WordWriteResult.Fail(
                            OfficeErrorCode.PARTIAL_APPLY,
                            "批量题注部分完成：成功 " + addedCount + " 个，失败 " +
                            failures.Count + " 个；首个失败：" +
                            failures[0]["reason"].ToString() + "。已添加题注不会自动回滚。",
                            data,
                            true);
                    }
                    return WordWriteResult.Ok(
                        "已为 " + addedCount + " 个目标添加题注（target=" + targetKind +
                        (addBookmark ? "，并创建 " + addedCount + " 个书签" : "") + "）。",
                        data,
                        true).AddTargetRef("Word:Captions");
                });
        }

        private static void CollectCaptionTargets(
            Word.Document document,
            string target,
            List<int> tableIndexes,
            bool includeInline,
            bool includeFloating,
            string captionText,
            string label,
            List<CaptionTarget> targets,
            JArray warnings)
        {
            if (target == "tables" || target == "both")
            {
                string tableLabel = string.IsNullOrWhiteSpace(label) ? "表" : label;
                Word.Tables tables = null;
                try
                {
                    tables = document.Tables;
                    int count = tables == null ? 0 : tables.Count;
                    List<int> indexes = new List<int>();
                    if (tableIndexes != null && tableIndexes.Count > 0)
                    {
                        foreach (int index in tableIndexes)
                        {
                            if (index < 1 || index > count)
                            {
                                WordExtensionsUtil.AddWarning(
                                    warnings,
                                    "tableIndexes 中的 " + index + " 越界（当前共 " + count +
                                    " 个表格），已跳过。");
                                continue;
                            }
                            indexes.Add(index);
                        }
                    }
                    else
                    {
                        for (int index = 1; index <= count; index++)
                        {
                            indexes.Add(index);
                        }
                    }
                    foreach (int index in indexes)
                    {
                        Word.Table table = null;
                        Word.Range range = null;
                        try
                        {
                            table = tables[index];
                            range = table.Range;
                            CaptionTarget item = new CaptionTarget();
                            item.Kind = "table";
                            item.Index = index;
                            item.Label = tableLabel;
                            item.CaptionText = captionText ?? string.Empty;
                            item.AnchorPosition = range.End;
                            targets.Add(item);
                        }
                        finally
                        {
                            ComHelper.Release(range, table);
                        }
                    }
                }
                finally
                {
                    ComHelper.Release(tables);
                }
            }

            if (target == "images" || target == "both")
            {
                string imageLabel = string.IsNullOrWhiteSpace(label) ? "图" : label;
                if (includeInline)
                {
                    Word.InlineShapes inlineShapes = null;
                    try
                    {
                        inlineShapes = document.InlineShapes;
                        int count = inlineShapes == null ? 0 : inlineShapes.Count;
                        for (int index = 1; index <= count; index++)
                        {
                            Word.InlineShape shape = null;
                            Word.Range range = null;
                            try
                            {
                                shape = inlineShapes[index];
                                bool isPicture =
                                    shape.Type == Word.WdInlineShapeType.wdInlineShapePicture ||
                                    shape.Type == Word.WdInlineShapeType.wdInlineShapeLinkedPicture;
                                if (!isPicture)
                                {
                                    continue;
                                }
                                range = shape.Range;
                                CaptionTarget item = new CaptionTarget();
                                item.Kind = "inline";
                                item.Index = index;
                                item.Label = imageLabel;
                                item.CaptionText = string.IsNullOrWhiteSpace(captionText)
                                    ? WordExtensionsUtil.SafeReadAlternativeText(shape)
                                    : captionText;
                                item.AnchorPosition = range == null ? index : range.Start;
                                targets.Add(item);
                            }
                            finally
                            {
                                ComHelper.Release(range, shape);
                            }
                        }
                    }
                    finally
                    {
                        ComHelper.Release(inlineShapes);
                    }
                }
                if (includeFloating)
                {
                    Word.Shapes shapes = null;
                    try
                    {
                        shapes = document.Shapes;
                        int count = shapes == null ? 0 : shapes.Count;
                        for (int index = 1; index <= count; index++)
                        {
                            Word.Shape shape = null;
                            Word.Range anchor = null;
                            try
                            {
                                shape = shapes[index];
                                int typeCode = WordImageUtil.ReadFloatingTypeCode(shape);
                                if (typeCode != 13 && typeCode != 11 && typeCode != 28 &&
                                    typeCode != 29)
                                {
                                    continue;
                                }
                                anchor = shape.Anchor;
                                CaptionTarget item = new CaptionTarget();
                                item.Kind = "floating";
                                item.Index = index;
                                item.Label = imageLabel;
                                item.CaptionText = string.IsNullOrWhiteSpace(captionText)
                                    ? WordExtensionsUtil.SafeReadAlternativeText(shape)
                                    : captionText;
                                item.AnchorPosition = anchor == null ? index : anchor.Start;
                                targets.Add(item);
                            }
                            finally
                            {
                                ComHelper.Release(anchor, shape);
                            }
                        }
                    }
                    finally
                    {
                        ComHelper.Release(shapes);
                    }
                }
            }
        }

        /// <summary>
        /// 在指定字符位置之后插入一个题注段落。优先在“位置之后的第一段”之前插入新段落，
        /// 这样表格后的题注会落在表格之后、后续正文之前；没有可插入段时在位置后补齐段落。
        /// 返回 false 时 failure 说明具体原因（实际插入位置/回读文本）。
        /// </summary>
        private static bool InsertCaptionParagraph(
            Word.Document document,
            int position,
            string captionText,
            string styleName,
            out int paragraphIndex,
            out string failure)
        {
            paragraphIndex = 0;
            failure = string.Empty;
            Word.Paragraphs paragraphs = null;
            Word.Paragraph targetParagraph = null;
            Word.Range targetRange = null;
            Word.Range insertionRange = null;
            try
            {
                paragraphs = document.Paragraphs;
                int count = paragraphs == null ? 0 : paragraphs.Count;
                int targetIndex = 0;
                for (int index = 1; index <= count; index++)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range range = null;
                    try
                    {
                        paragraph = paragraphs[index];
                        range = paragraph.Range;
                        if (range.Start >= position)
                        {
                            targetIndex = index;
                            break;
                        }
                    }
                    finally
                    {
                        ComHelper.Release(range, paragraph);
                    }
                }

                if (targetIndex > 0)
                {
                    targetParagraph = WordComUtil.GetParagraph(document, targetIndex);
                    targetRange = targetParagraph.Range;
                    targetRange.InsertBefore(captionText + "\r");
                }
                else
                {
                    int contentEnd = Math.Max(0, WordComUtil.GetContentEnd(document) - 1);
                    int safePosition = Math.Max(0, Math.Min(position, contentEnd));
                    insertionRange = WordComUtil.CreateRange(
                        document, safePosition, safePosition);
                    insertionRange.InsertAfter("\r" + captionText + "\r");
                }

                paragraphIndex = FindCaptionParagraphIndex(
                    document, captionText, targetIndex);
                if (paragraphIndex <= 0)
                {
                    failure = "题注段落插入后未找到文本「" +
                              WordExtensionsUtil.Preview(captionText, 80) +
                              "」；插入位置 " + position + "，原有 " + count +
                              " 段。请确认锚点不在表格单元格/受保护区域。";
                    return false;
                }

                Word.Paragraph captionParagraph = null;
                Word.Range captionRange = null;
                try
                {
                    captionParagraph = WordComUtil.GetParagraph(document, paragraphIndex);
                    captionRange = captionParagraph.Range;
                    if (!string.IsNullOrWhiteSpace(styleName))
                    {
                        WordExtensionsUtil.ApplyParagraphStyle(
                            document, captionParagraph, styleName, false, false);
                    }
                }
                finally
                {
                    ComHelper.Release(captionRange, captionParagraph);
                }
                return true;
            }
            catch (Exception ex)
            {
                failure = "题注段落插入失败：" +
                          OfficeErrorMapper.DescribeException(ex, OfficeApp.Word);
                return false;
            }
            finally
            {
                ComHelper.Release(insertionRange, targetRange, targetParagraph, paragraphs);
            }
        }

        private static int FindCaptionParagraphIndex(
            Word.Document document,
            string captionText,
            int hintIndex)
        {
            Word.Paragraphs paragraphs = null;
            try
            {
                paragraphs = document.Paragraphs;
                int count = paragraphs == null ? 0 : paragraphs.Count;
                if (count <= 0)
                {
                    return 0;
                }
                int start = hintIndex > 0 ? Math.Max(1, hintIndex - 2) : 1;
                int end = hintIndex > 0 ? Math.Min(count, hintIndex + 3) : count;
                for (int index = start; index <= end; index++)
                {
                    if (ParagraphTextEquals(document, paragraphs, index, captionText))
                    {
                        return index;
                    }
                }
                for (int index = count; index >= 1; index--)
                {
                    if (ParagraphTextEquals(document, paragraphs, index, captionText))
                    {
                        return index;
                    }
                }
                return 0;
            }
            finally
            {
                ComHelper.Release(paragraphs);
            }
        }

        private static bool ParagraphTextEquals(
            Word.Document document,
            Word.Paragraphs paragraphs,
            int index,
            string expected)
        {
            Word.Paragraph paragraph = null;
            Word.Range range = null;
            try
            {
                paragraph = paragraphs[index];
                range = paragraph.Range;
                return string.Equals(
                    WordExtensionsUtil.ReadParagraphText(paragraph),
                    expected,
                    StringComparison.Ordinal);
            }
            finally
            {
                ComHelper.Release(range, paragraph);
            }
        }

        private static Word.Range ResolveCaptionAnchor(Word.Document document, CaptionTarget target)
        {
            if (target.Kind == "table")
            {
                Word.Table table = null;
                try
                {
                    table = WordComUtil.GetTable(document, target.Index);
                    return table.Range;
                }
                finally
                {
                    ComHelper.Release(table);
                }
            }
            if (target.Kind == "inline")
            {
                Word.InlineShapes shapes = null;
                try
                {
                    shapes = document.InlineShapes;
                    if (shapes == null || target.Index < 1 || target.Index > shapes.Count)
                    {
                        return null;
                    }
                    Word.InlineShape shape = null;
                    try
                    {
                        shape = shapes[target.Index];
                        return shape.Range;
                    }
                    finally
                    {
                        ComHelper.Release(shape);
                    }
                }
                finally
                {
                    ComHelper.Release(shapes);
                }
            }
            Word.Shapes floatingShapes = null;
            try
            {
                floatingShapes = document.Shapes;
                if (floatingShapes == null || target.Index < 1 ||
                    target.Index > floatingShapes.Count)
                {
                    return null;
                }
                Word.Shape shape = null;
                try
                {
                    shape = floatingShapes[target.Index];
                    return shape.Anchor;
                }
                finally
                {
                    ComHelper.Release(shape);
                }
            }
            finally
            {
                ComHelper.Release(floatingShapes);
            }
        }

        private static Dictionary<string, int> CountExistingCaptions(Word.Document document)
        {
            Dictionary<string, int> counts = new Dictionary<string, int>(StringComparer.Ordinal);
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
                        string text = WordExtensionsUtil.ReadParagraphText(paragraph);
                        if (string.IsNullOrWhiteSpace(text))
                        {
                            continue;
                        }
                        string trimmed = text.TrimStart();
                        string label;
                        if (trimmed.StartsWith("表", StringComparison.Ordinal))
                        {
                            label = "表";
                        }
                        else if (trimmed.StartsWith("图", StringComparison.Ordinal))
                        {
                            label = "图";
                        }
                        else
                        {
                            continue;
                        }
                        int cursor = label.Length;
                        while (cursor < trimmed.Length && trimmed[cursor] == ' ')
                        {
                            cursor++;
                        }
                        int digits = 0;
                        while (cursor < trimmed.Length && char.IsDigit(trimmed[cursor]))
                        {
                            cursor++;
                            digits++;
                        }
                        if (digits > 0)
                        {
                            int current;
                            counts.TryGetValue(label, out current);
                            counts[label] = current + 1;
                        }
                    }
                    finally
                    {
                        ComHelper.Release(range, paragraph);
                    }
                }
                return counts;
            }
            finally
            {
                ComHelper.Release(paragraphs);
            }
        }

        private static int CountFields(Word.Document document)
        {
            Word.Fields fields = null;
            try
            {
                fields = document.Fields;
                return fields == null ? 0 : fields.Count;
            }
            finally
            {
                ComHelper.Release(fields);
            }
        }

        private sealed class CaptionTarget
        {
            public string Kind { get; set; }

            public int Index { get; set; }

            public string Label { get; set; }

            public string CaptionText { get; set; }

            public int AnchorPosition { get; set; }

            public int Number { get; set; }

            public CaptionTarget()
            {
                Kind = string.Empty;
                Label = string.Empty;
                CaptionText = string.Empty;
            }
        }

        // ------------------------------------------------------------------
        // word_watermark
        // ------------------------------------------------------------------

        private OperationResult HandleWatermark(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string action = (args.GetOptionalString("action", "add") ?? "add")
                .Trim().ToLowerInvariant();
            if (action == "list")
            {
                return RunReadTool(
                    application,
                    delegate(Word.Document document)
                    {
                        JArray items = new JArray();
                        HashSet<string> seenNames = new HashSet<string>(StringComparer.Ordinal);
                        int sectionCount = 0;
                        Word.Sections sections = null;
                        try
                        {
                            sections = document.Sections;
                            sectionCount = sections == null ? 0 : sections.Count;
                            for (int index = 1; index <= sectionCount; index++)
                            {
                                CollectWatermarkShapes(
                                    document, sections, index, seenNames, items);
                            }
                        }
                        finally
                        {
                            ComHelper.Release(sections);
                        }
                        JObject data = new JObject();
                        data["action"] = "list";
                        data["sectionCount"] = sectionCount;
                        data["watermarkCount"] = items.Count;
                        data["watermarks"] = items;
                        return ReadSucceed(
                            "已读取 " + items.Count + " 个 AutoOfficeMCP 水印形状。",
                            data, "Word:Headers");
                    });
            }
            if (action != "add" && action != "remove")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "action 只支持 add/remove/list，收到「" + action + "」。");
            }
            int sectionIndex = args.GetOptionalInt("sectionIndex", 0);

            if (action == "remove")
            {
                return RunWriteTool(
                    application,
                    WordExtensionsToolNames.Watermark,
                    false,
                    delegate(Word.Document document)
                    {
                        if (sectionIndex > 0)
                        {
                            WordComUtil.ValidateSectionIndex(document, sectionIndex);
                        }
                        return WordWriteTargets.None("页眉水印", "Word:Headers");
                    },
                    delegate(Word.Document document, WordWriteTarget target)
                    {
                        JArray warnings = new JArray();
                        int removed = RemoveWatermarks(
                            document, sectionIndex, warnings);
                        JObject data = new JObject();
                        data["action"] = "remove";
                        data["sectionIndex"] = sectionIndex;
                        data["removedCount"] = removed;
                        data["warnings"] = warnings;
                        data["note"] = "只删除名称以 " + WatermarkShapePrefix +
                                        " 开头的水印形状，不改动用户自行插入的页眉形状。";
                        return WordWriteResult.Ok(
                            removed > 0
                                ? "已删除 " + removed + " 个 AutoOfficeMCP 水印形状。"
                                : "未找到 AutoOfficeMCP 水印形状，文档未修改。",
                            data,
                            removed > 0).AddTargetRef("Word:Headers");
                    });
            }

            string text = args.GetOptionalString("text", "内部资料");
            if (string.IsNullOrWhiteSpace(text))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "水印文字 text 不能为空。");
            }
            string fontName = args.GetOptionalString("fontName", "宋体");
            double fontSize = args.GetOptionalDouble("fontSize").HasValue
                ? args.GetOptionalDouble("fontSize").Value : 54d;
            if (fontSize < 8d || fontSize > 200d)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "fontSize 允许 8-200，收到 " + fontSize + "。");
            }
            string colorText = args.GetOptionalString("color", "#C0C0C0");
            int color = WordFormatUtil.ParseColor(colorText);
            double rotation = args.GetOptionalDouble("rotation").HasValue
                ? args.GetOptionalDouble("rotation").Value : 315d;
            if (rotation < 0d || rotation > 360d)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "rotation 允许 0-360，收到 " + rotation + "。");
            }
            double transparency = args.GetOptionalDouble("transparency").HasValue
                ? args.GetOptionalDouble("transparency").Value : 0.5d;
            if (transparency < 0d || transparency > 1d)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "transparency 允许 0-1，收到 " + transparency + "。");
            }

            return RunWriteTool(
                application,
                WordExtensionsToolNames.Watermark,
                false,
                delegate(Word.Document document)
                {
                    if (sectionIndex > 0)
                    {
                        WordComUtil.ValidateSectionIndex(document, sectionIndex);
                    }
                    return WordWriteTargets.None("页眉水印", "Word:Headers");
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    JArray warnings = new JArray();
                    JArray created = new JArray();
                    int added = AddWatermarks(
                        document, sectionIndex, text, fontName, (float)fontSize,
                        color, (float)rotation, (float)transparency,
                        created, warnings);
                    JObject data = new JObject();
                    data["action"] = "add";
                    data["text"] = text;
                    data["fontName"] = fontName;
                    data["fontSize"] = fontSize;
                    data["color"] = colorText;
                    data["rotation"] = rotation;
                    data["transparency"] = transparency;
                    data["sectionIndex"] = sectionIndex;
                    data["addedCount"] = added;
                    data["created"] = created;
                    data["warnings"] = warnings;
                    data["note"] = "水印形状添加到节的页眉 Shapes（名称前缀 " +
                                    WatermarkShapePrefix + "），衬于文字下方并半透明；" +
                                    "链接到上一节的页眉不会重复添加。";
                    if (added == 0)
                    {
                        return WordWriteResult.Fail(
                            OfficeErrorCode.INTERNAL_ERROR,
                            "水印添加失败：未创建任何形状。请检查节/页眉是否存在且文档未受保护。" +
                            (warnings.Count > 0 ? " 首个警告：" + warnings[0].ToString() : ""),
                            data,
                            false);
                    }
                    return WordWriteResult.Ok(
                        "已添加 " + added + " 个页眉水印「" + text + "」。",
                        data,
                        true).AddTargetRef("Word:Headers");
                });
        }

        private static int AddWatermarks(
            Word.Document document,
            int sectionIndex,
            string text,
            string fontName,
            float fontSize,
            int color,
            float rotation,
            float transparency,
            JArray created,
            JArray warnings)
        {
            int added = 0;
            Word.Sections sections = null;
            try
            {
                sections = document.Sections;
                int count = sections == null ? 0 : sections.Count;
                if (sectionIndex > 0 && sectionIndex > count)
                {
                    WordExtensionsUtil.AddWarning(
                        warnings,
                        "sectionIndex " + sectionIndex + " 不存在（当前共 " + count + " 节）。");
                    return 0;
                }
                List<int> indexes = new List<int>();
                if (sectionIndex > 0)
                {
                    indexes.Add(sectionIndex);
                }
                else
                {
                    for (int index = 1; index <= count; index++)
                    {
                        indexes.Add(index);
                    }
                }
                foreach (int index in indexes)
                {
                    Word.Section section = null;
                    Word.PageSetup pageSetup = null;
                    Word.HeadersFooters headers = null;
                    try
                    {
                        section = sections[index];
                        pageSetup = section.PageSetup;
                        headers = section.Headers;
                        AddWatermarkToHeader(
                            document, section, pageSetup, headers,
                            Word.WdHeaderFooterIndex.wdHeaderFooterPrimary,
                            index, 0, text, fontName, fontSize, color, rotation,
                            transparency, created, warnings, ref added);
                        if (pageSetup.DifferentFirstPageHeaderFooter != 0)
                        {
                            AddWatermarkToHeader(
                                document, section, pageSetup, headers,
                                Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage,
                                index, 1, text, fontName, fontSize, color, rotation,
                                transparency, created, warnings, ref added);
                        }
                        if (pageSetup.OddAndEvenPagesHeaderFooter != 0)
                        {
                            AddWatermarkToHeader(
                                document, section, pageSetup, headers,
                                Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages,
                                index, 2, text, fontName, fontSize, color, rotation,
                                transparency, created, warnings, ref added);
                        }
                    }
                    catch (Exception ex)
                    {
                        WordExtensionsUtil.AddWarning(
                            warnings,
                            "第 " + index + " 节水印添加失败：" +
                            OfficeErrorMapper.DescribeException(ex, OfficeApp.Word));
                    }
                    finally
                    {
                        ComHelper.Release(headers, pageSetup, section);
                    }
                }
                return added;
            }
            finally
            {
                ComHelper.Release(sections);
            }
        }

        private static void AddWatermarkToHeader(
            Word.Document document,
            Word.Section section,
            Word.PageSetup pageSetup,
            Word.HeadersFooters headers,
            Word.WdHeaderFooterIndex headerIndex,
            int sectionIndex,
            int headerOrdinal,
            string text,
            string fontName,
            float fontSize,
            int color,
            float rotation,
            float transparency,
            JArray created,
            JArray warnings,
            ref int added)
        {
            Word.HeaderFooter header = null;
            Word.Range headerRange = null;
            Word.Shapes shapes = null;
            Word.Shape shape = null;
            Word.TextFrame textFrame = null;
            Word.Range textRange = null;
            Word.Font font = null;
            Word.ParagraphFormat paragraphFormat = null;
            Word.WrapFormat wrapFormat = null;
            Word.LineFormat lineFormat = null;
            Word.FillFormat fillFormat = null;
            Word.ColorFormat fillColor = null;
            try
            {
                header = headers[headerIndex];
                if (header == null)
                {
                    return;
                }
                if (sectionIndex > 1 && header.LinkToPrevious)
                {
                    return;
                }
                headerRange = header.Range;
                object anchorObject = headerRange;
                shapes = header.Shapes;
                float pageWidth = pageSetup == null ? 595f : pageSetup.PageWidth;
                float pageHeight = pageSetup == null ? 842f : pageSetup.PageHeight;
                float width = pageWidth * 0.8f;
                float height = 120f;
                float left = (pageWidth - width) / 2f;
                float top = (pageHeight - height) / 2f;
                shape = shapes.AddShape(1, left, top, width, height, ref anchorObject);
                if (shape == null)
                {
                    return;
                }
                string shapeName = WatermarkShapePrefix + sectionIndex + "_" + headerOrdinal;
                shape.Name = shapeName;
                shape.Rotation = rotation;
                wrapFormat = shape.WrapFormat;
                wrapFormat.Type = Word.WdWrapType.wdWrapBehind;
                shape.RelativeHorizontalPosition =
                    Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                shape.RelativeVerticalPosition =
                    Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                shape.Left = left;
                shape.Top = top;
                shape.Width = width;
                shape.Height = height;
                fillFormat = shape.Fill;
                fillFormat.Transparency = transparency;
                fillColor = fillFormat.ForeColor;
                fillColor.RGB = color;
                lineFormat = shape.Line;
                TrySetComProperty(lineFormat, "Visible", 0);
                textFrame = shape.TextFrame;
                textRange = textFrame.TextRange;
                textRange.Text = text;
                font = textRange.Font;
                font.Name = fontName;
                font.NameFarEast = fontName;
                font.Size = fontSize;
                font.Bold = -1;
                font.Color = (Word.WdColor)color;
                paragraphFormat = textRange.ParagraphFormat;
                paragraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                added++;
                created.Add(new JObject
                {
                    { "sectionIndex", sectionIndex },
                    { "header", headerIndex.ToString() },
                    { "shapeName", shapeName },
                    { "text", text }
                });
            }
            catch (Exception ex)
            {
                WordExtensionsUtil.AddWarning(
                    warnings,
                    "第 " + sectionIndex + " 节页眉 " + headerIndex +
                    " 水印形状创建失败：" +
                    OfficeErrorMapper.DescribeException(ex, OfficeApp.Word));
            }
            finally
            {
                ComHelper.Release(paragraphFormat, font, textRange, textFrame,
                    lineFormat, wrapFormat, fillColor, fillFormat, shape, shapes,
                    headerRange, header);
            }
        }

        private static int RemoveWatermarks(
            Word.Document document,
            int sectionIndex,
            JArray warnings)
        {
            int removed = 0;
            Word.Sections sections = null;
            try
            {
                sections = document.Sections;
                int count = sections == null ? 0 : sections.Count;
                List<int> indexes = new List<int>();
                if (sectionIndex > 0)
                {
                    indexes.Add(sectionIndex);
                }
                else
                {
                    for (int index = 1; index <= count; index++)
                    {
                        indexes.Add(index);
                    }
                }
                foreach (int index in indexes)
                {
                    Word.Section section = null;
                    Word.HeadersFooters headers = null;
                    try
                    {
                        section = sections[index];
                        headers = section.Headers;
                        for (int headerOrdinal = 1; headerOrdinal <= 3; headerOrdinal++)
                        {
                            Word.HeaderFooter header = null;
                            Word.Shapes shapes = null;
                            try
                            {
                                header = headers[(Word.WdHeaderFooterIndex)headerOrdinal];
                                if (header == null)
                                {
                                    continue;
                                }
                                shapes = header.Shapes;
                                int shapeCount = shapes == null ? 0 : shapes.Count;
                                for (int shapeIndex = shapeCount; shapeIndex >= 1; shapeIndex--)
                                {
                                    Word.Shape shape = null;
                                    try
                                    {
                                        shape = shapes[shapeIndex];
                                        string name = shape.Name ?? string.Empty;
                                        if (name.StartsWith(
                                            WatermarkShapePrefix, StringComparison.Ordinal))
                                        {
                                            shape.Delete();
                                            removed++;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        WordExtensionsUtil.AddWarning(
                                            warnings,
                                            "第 " + index + " 节页眉 " + headerOrdinal +
                                            " 删除水印失败：" + ex.Message);
                                    }
                                    finally
                                    {
                                        ComHelper.Release(shape);
                                    }
                                }
                            }
                            finally
                            {
                                ComHelper.Release(shapes, header);
                            }
                        }
                    }
                    finally
                    {
                        ComHelper.Release(headers, section);
                    }
                }
                return removed;
            }
            finally
            {
                ComHelper.Release(sections);
            }
        }

        private static void CollectWatermarkShapes(
            Word.Document document,
            Word.Sections sections,
            int sectionIndex,
            HashSet<string> seenNames,
            JArray items)
        {
            Word.Section section = null;
            Word.HeadersFooters headers = null;
            try
            {
                section = sections[sectionIndex];
                headers = section.Headers;
                for (int headerOrdinal = 1; headerOrdinal <= 3; headerOrdinal++)
                {
                    Word.HeaderFooter header = null;
                    Word.Shapes shapes = null;
                    try
                    {
                        header = headers[(Word.WdHeaderFooterIndex)headerOrdinal];
                        if (header == null)
                        {
                            continue;
                        }
                        shapes = header.Shapes;
                        int count = shapes == null ? 0 : shapes.Count;
                        for (int index = 1; index <= count; index++)
                        {
                            Word.Shape shape = null;
                            Word.TextFrame textFrame = null;
                            Word.Range textRange = null;
                            Word.FillFormat fillFormat = null;
                            try
                            {
                                shape = shapes[index];
                                string name = shape.Name ?? string.Empty;
                                if (!name.StartsWith(
                                    WatermarkShapePrefix, StringComparison.Ordinal))
                                {
                                    continue;
                                }
                                if (seenNames != null && !seenNames.Add(name))
                                {
                                    // 链接到上一节的页眉会在每个节/页眉类型下暴露同一形状，按名称去重。
                                    continue;
                                }
                                textFrame = shape.TextFrame;
                                textRange = textFrame.TextRange;
                                fillFormat = shape.Fill;
                                items.Add(new JObject
                                {
                                    { "sectionIndex", sectionIndex },
                                    { "header", ((Word.WdHeaderFooterIndex)headerOrdinal).ToString() },
                                    { "shapeIndex", index },
                                    { "shapeName", name },
                                    { "text", textRange == null ? string.Empty : (textRange.Text ?? string.Empty) },
                                    { "rotation", (double)shape.Rotation },
                                    { "transparency", (double)fillFormat.Transparency }
                                });
                            }
                            finally
                            {
                                ComHelper.Release(fillFormat, textRange, textFrame, shape);
                            }
                        }
                    }
                    finally
                    {
                        ComHelper.Release(shapes, header);
                    }
                }
            }
            finally
            {
                ComHelper.Release(headers, section);
            }
        }

        private static void TrySetComProperty(object target, string propertyName, object value)
        {
            if (target == null)
            {
                return;
            }
            try
            {
                target.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.SetProperty,
                    null,
                    target,
                    new object[] { value });
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[word-ext] 设置 COM 属性 " + propertyName +
                                        " 失败：" + ex.Message);
            }
        }

        // ------------------------------------------------------------------
        // word_footnote_manage
        // ------------------------------------------------------------------

        private OperationResult HandleFootnoteManage(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string action = (args.GetOptionalString("action", "list") ?? "list")
                .Trim().ToLowerInvariant();
            string noteType = (args.GetOptionalString("noteType", "footnote") ?? "footnote")
                .Trim().ToLowerInvariant();
            if (action != "list" && action != "add" &&
                action != "delete" && action != "format")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "action 只支持 list/add/delete/format，收到「" + action + "」。");
            }
            if (noteType != "footnote" && noteType != "endnote")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "noteType 只支持 footnote/endnote，收到「" + noteType + "」。");
            }

            if (action == "list")
            {
                return RunReadTool(
                    application,
                    delegate(Word.Document document)
                    {
                        JArray items = new JArray();
                        int count;
                        if (noteType == "footnote")
                        {
                            count = CollectFootnotes(document, items);
                        }
                        else
                        {
                            count = CollectEndnotes(document, items);
                        }
                        JObject data = new JObject();
                        data["action"] = "list";
                        data["noteType"] = noteType;
                        data["count"] = count;
                        data["notes"] = items;
                        return ReadSucceed(
                            "已读取 " + count + " 个" +
                            (noteType == "footnote" ? "脚注" : "尾注") + "。",
                            data, "Word:Document");
                    });
            }

            int paragraphIndex = args.GetOptionalInt("paragraphIndex", 0);
            string text = args.GetOptionalString("text", string.Empty);
            List<int> indexes = args.GetOptionalIntList("indexes");
            bool confirm = args.GetOptionalBool("confirm", false);
            string fontName = args.GetOptionalString("fontName", null);
            double? fontSize = args.GetOptionalDouble("fontSize");
            if (fontSize.HasValue && (fontSize.Value < 6d || fontSize.Value > 72d))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "fontSize 允许 6-72，收到 " + fontSize.Value + "。");
            }
            bool? bold = args.Has("bold") ? args.GetOptionalBool("bold", false) : (bool?)null;
            bool? italic = args.Has("italic") ? args.GetOptionalBool("italic", false) : (bool?)null;

            if (action == "add")
            {
                if (string.IsNullOrWhiteSpace(text))
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.INVALID_PARAMS,
                        "action=add 时 text 不能为空（脚注/尾注正文）。");
                }
                return RunWriteTool(
                    application,
                    WordExtensionsToolNames.FootnoteManage,
                    false,
                    delegate(Word.Document document)
                    {
                        if (paragraphIndex > 0)
                        {
                            WordComUtil.ValidateParagraphIndex(document, paragraphIndex);
                            return WordWriteTargets.Paragraph(
                                document, paragraphIndex,
                                "Word:Paragraph:" + paragraphIndex, true);
                        }
                        return WordWriteTargets.Cursor(application);
                    },
                    delegate(Word.Document document, WordWriteTarget target)
                    {
                        Word.Selection selection = null;
                        Word.Range insertionRange = null;
                        Word.Paragraph paragraph = null;
                        Word.Range paragraphRange = null;
                        Word.Footnotes footnotes = null;
                        Word.Endnotes endnotes = null;
                        Word.Footnote footnote = null;
                        Word.Endnote endnote = null;
                        try
                        {
                            if (paragraphIndex > 0)
                            {
                                paragraph = WordComUtil.GetParagraph(document, paragraphIndex);
                                paragraphRange = paragraph.Range;
                                int position = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
                                insertionRange = WordComUtil.CreateRange(document, position, position);
                            }
                            else
                            {
                                selection = application.Selection;
                                if (selection == null)
                                {
                                    return WordWriteResult.Fail(
                                        OfficeErrorCode.INVALID_PARAMS,
                                        "当前没有有效选区，无法添加" +
                                        (noteType == "footnote" ? "脚注" : "尾注") +
                                        "。请把光标放到插入位置或传 paragraphIndex。",
                                        null,
                                        false);
                                }
                                insertionRange = selection.Range;
                            }

                            int beforeCount;
                            int afterCount;
                            object reference = Type.Missing;
                            object noteText = text;
                            if (noteType == "footnote")
                            {
                                footnotes = document.Footnotes;
                                beforeCount = footnotes == null ? 0 : footnotes.Count;
                                footnote = footnotes.Add(
                                    insertionRange, ref reference, ref noteText);
                                afterCount = footnotes == null ? 0 : footnotes.Count;
                            }
                            else
                            {
                                endnotes = document.Endnotes;
                                beforeCount = endnotes == null ? 0 : endnotes.Count;
                                endnote = endnotes.Add(
                                    insertionRange, ref reference, ref noteText);
                                afterCount = endnotes == null ? 0 : endnotes.Count;
                            }

                            JObject data = new JObject();
                            data["action"] = "add";
                            data["noteType"] = noteType;
                            data["paragraphIndex"] = paragraphIndex;
                            data["text"] = text;
                            data["countBefore"] = beforeCount;
                            data["countAfter"] = afterCount;
                            data["note"] = "脚注/尾注引用插入到指定段落末尾（paragraphIndex）或当前选区；" +
                                            "引用编号由 Word 自动维护。";
                            if (afterCount <= beforeCount)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INTERNAL_ERROR,
                                    (noteType == "footnote" ? "Footnotes" : "Endnotes") +
                                    ".Add 已返回，但数量未增加（" + beforeCount + "→" +
                                    afterCount + "）。",
                                    data,
                                    false);
                            }
                            return WordWriteResult.Ok(
                                "已添加 " + (noteType == "footnote" ? "脚注" : "尾注") +
                                "（当前共 " + afterCount + " 个）。",
                                data,
                                true).AddTargetRef("Word:Document");
                        }
                        finally
                        {
                            ComHelper.Release(endnote, footnote, endnotes, footnotes,
                                paragraphRange, paragraph, insertionRange, selection);
                        }
                    });
            }

            if (action == "delete")
            {
                if (!confirm)
                {
                    return OperationResult.Cancelled(
                        "删除" + (noteType == "footnote" ? "脚注" : "尾注") +
                        "需要 confirm=true；本次未修改文档。");
                }
                if (indexes == null || indexes.Count == 0)
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.INVALID_PARAMS,
                        "action=delete 必须提供 indexes 数组（1 基），避免误删全部注释。" +
                        "若要删除全部，请显式列出全部索引。");
                }
                return RunWriteTool(
                    application,
                    WordExtensionsToolNames.FootnoteManage,
                    false,
                    delegate(Word.Document document)
                    {
                        return WordWriteTargets.None("脚注/尾注删除", "Word:Document");
                    },
                    delegate(Word.Document document, WordWriteTarget target)
                    {
                        JArray deleted = new JArray();
                        JArray failures = new JArray();
                        Word.Footnotes footnotes = null;
                        Word.Endnotes endnotes = null;
                        try
                        {
                            int count = noteType == "footnote"
                                ? CountFootnotes(document)
                                : CountEndnotes(document);
                            List<int> sorted = new List<int>(indexes);
                            sorted.Sort();
                            sorted.Reverse();
                            foreach (int index in sorted)
                            {
                                if (index < 1 || index > count)
                                {
                                    failures.Add(new JObject
                                    {
                                        { "index", index },
                                        { "reason", "索引越界（当前共 " + count + " 个）" }
                                    });
                                    continue;
                                }
                                try
                                {
                                    if (noteType == "footnote")
                                    {
                                        footnotes = document.Footnotes;
                                        Word.Footnote note = null;
                                        try
                                        {
                                            note = footnotes[index];
                                            note.Delete();
                                        }
                                        finally
                                        {
                                            ComHelper.Release(note);
                                        }
                                    }
                                    else
                                    {
                                        endnotes = document.Endnotes;
                                        Word.Endnote note = null;
                                        try
                                        {
                                            note = endnotes[index];
                                            note.Delete();
                                        }
                                        finally
                                        {
                                            ComHelper.Release(note);
                                        }
                                    }
                                    deleted.Add(index);
                                }
                                catch (Exception ex)
                                {
                                    failures.Add(new JObject
                                    {
                                        { "index", index },
                                        { "reason", OfficeErrorMapper.DescribeException(
                                            ex, OfficeApp.Word) }
                                    });
                                }
                            }
                            int afterCount = noteType == "footnote"
                                ? CountFootnotes(document)
                                : CountEndnotes(document);
                            JObject data = new JObject();
                            data["action"] = "delete";
                            data["noteType"] = noteType;
                            data["countBefore"] = count;
                            data["countAfter"] = afterCount;
                            data["deleted"] = deleted;
                            data["failures"] = failures;
                            data["note"] = "按索引倒序删除，避免索引位移；删除后编号由 Word 自动重排。";
                            if (deleted.Count == 0)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INVALID_PARAMS,
                                    "没有任何" + (noteType == "footnote" ? "脚注" : "尾注") +
                                    "被删除。首个原因：" +
                                    (failures.Count > 0 ? failures[0]["reason"].ToString() : "未知") +
                                    "。当前共 " + count + " 个。",
                                    data,
                                    false);
                            }
                            if (failures.Count > 0)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.PARTIAL_APPLY,
                                    "删除部分完成：删除 " + deleted.Count + " 个，失败 " +
                                    failures.Count + " 个；不会自动回滚。",
                                    data,
                                    true);
                            }
                            return WordWriteResult.Ok(
                                "已删除 " + deleted.Count + " 个" +
                                (noteType == "footnote" ? "脚注" : "尾注") +
                                "（" + count + "→" + afterCount + "）。",
                                data,
                                true).AddTargetRef("Word:Document");
                        }
                        finally
                        {
                            ComHelper.Release(endnotes, footnotes);
                        }
                    });
            }

            // action == format
            if (!fontSize.HasValue && string.IsNullOrWhiteSpace(fontName) &&
                !bold.HasValue && !italic.HasValue)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "action=format 至少需要 fontName/fontSize/bold/italic 之一。");
            }
            return RunWriteTool(
                application,
                WordExtensionsToolNames.FootnoteManage,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.None("脚注/尾注格式", "Word:Document");
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    JArray formatted = new JArray();
                    JArray failures = new JArray();
                    Word.Footnotes footnotes = null;
                    Word.Endnotes endnotes = null;
                    try
                    {
                        int count = noteType == "footnote"
                            ? CountFootnotes(document)
                            : CountEndnotes(document);
                        List<int> requested = new List<int>();
                        if (indexes != null && indexes.Count > 0)
                        {
                            requested.AddRange(indexes);
                        }
                        else
                        {
                            for (int index = 1; index <= count; index++)
                            {
                                requested.Add(index);
                            }
                        }
                        requested.Sort();
                        requested.Reverse();
                        foreach (int index in requested)
                        {
                            if (index < 1 || index > count)
                            {
                                failures.Add(new JObject
                                {
                                    { "index", index },
                                    { "reason", "索引越界（当前共 " + count + " 个）" }
                                });
                                continue;
                            }
                            Word.Range noteRange = null;
                            Word.Font font = null;
                            try
                            {
                                if (noteType == "footnote")
                                {
                                    footnotes = document.Footnotes;
                                    Word.Footnote note = null;
                                    try
                                    {
                                        note = footnotes[index];
                                        noteRange = note.Range;
                                    }
                                    finally
                                    {
                                        ComHelper.Release(note);
                                    }
                                }
                                else
                                {
                                    endnotes = document.Endnotes;
                                    Word.Endnote note = null;
                                    try
                                    {
                                        note = endnotes[index];
                                        noteRange = note.Range;
                                    }
                                    finally
                                    {
                                        ComHelper.Release(note);
                                    }
                                }
                                font = noteRange.Font;
                                if (!string.IsNullOrWhiteSpace(fontName))
                                {
                                    font.Name = fontName;
                                    font.NameFarEast = fontName;
                                }
                                if (fontSize.HasValue)
                                {
                                    font.Size = (float)fontSize.Value;
                                }
                                if (bold.HasValue)
                                {
                                    font.Bold = bold.Value ? -1 : 0;
                                }
                                if (italic.HasValue)
                                {
                                    font.Italic = italic.Value ? -1 : 0;
                                }
                                formatted.Add(index);
                            }
                            catch (Exception ex)
                            {
                                failures.Add(new JObject
                                {
                                    { "index", index },
                                    { "reason", OfficeErrorMapper.DescribeException(
                                        ex, OfficeApp.Word) }
                                });
                            }
                            finally
                            {
                                ComHelper.Release(font, noteRange);
                            }
                        }
                        JObject data = new JObject();
                        data["action"] = "format";
                        data["noteType"] = noteType;
                        data["count"] = count;
                        data["formatted"] = formatted;
                        data["failures"] = failures;
                        if (formatted.Count == 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "没有" + (noteType == "footnote" ? "脚注" : "尾注") +
                                "被格式化。首个原因：" +
                                (failures.Count > 0 ? failures[0]["reason"].ToString() : "未知") +
                                "。",
                                data,
                                false);
                        }
                        if (failures.Count > 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.PARTIAL_APPLY,
                                "格式设置部分完成：成功 " + formatted.Count + " 个，失败 " +
                                failures.Count + " 个；不会自动回滚。",
                                data,
                                true);
                        }
                        return WordWriteResult.Ok(
                            "已格式化 " + formatted.Count + " 个" +
                            (noteType == "footnote" ? "脚注" : "尾注") + "。",
                            data,
                            true).AddTargetRef("Word:Document");
                    }
                    finally
                    {
                        ComHelper.Release(endnotes, footnotes);
                    }
                });
        }

        private static int CountFootnotes(Word.Document document)
        {
            Word.Footnotes footnotes = null;
            try
            {
                footnotes = document.Footnotes;
                return footnotes == null ? 0 : footnotes.Count;
            }
            finally
            {
                ComHelper.Release(footnotes);
            }
        }

        private static int CountEndnotes(Word.Document document)
        {
            Word.Endnotes endnotes = null;
            try
            {
                endnotes = document.Endnotes;
                return endnotes == null ? 0 : endnotes.Count;
            }
            finally
            {
                ComHelper.Release(endnotes);
            }
        }

        private static int CollectFootnotes(Word.Document document, JArray items)
        {
            Word.Footnotes footnotes = null;
            try
            {
                footnotes = document.Footnotes;
                int count = footnotes == null ? 0 : footnotes.Count;
                for (int index = 1; index <= count; index++)
                {
                    Word.Footnote note = null;
                    Word.Range range = null;
                    Word.Range reference = null;
                    try
                    {
                        note = footnotes[index];
                        range = note.Range;
                        reference = note.Reference;
                        items.Add(new JObject
                        {
                            { "index", index },
                            { "text", WordComUtil.CleanParagraphText(
                                range == null ? string.Empty : range.Text) },
                            { "reference", WordComUtil.CleanParagraphText(
                                reference == null ? string.Empty : reference.Text) },
                            { "rangeStart", range == null ? 0 : range.Start },
                            { "paragraphIndex", range == null ? 0
                                : WordComUtil.FindParagraphIndexAt(document, range.Start) }
                        });
                    }
                    finally
                    {
                        ComHelper.Release(reference, range, note);
                    }
                }
                return count;
            }
            finally
            {
                ComHelper.Release(footnotes);
            }
        }

        private static int CollectEndnotes(Word.Document document, JArray items)
        {
            Word.Endnotes endnotes = null;
            try
            {
                endnotes = document.Endnotes;
                int count = endnotes == null ? 0 : endnotes.Count;
                for (int index = 1; index <= count; index++)
                {
                    Word.Endnote note = null;
                    Word.Range range = null;
                    Word.Range reference = null;
                    try
                    {
                        note = endnotes[index];
                        range = note.Range;
                        reference = note.Reference;
                        items.Add(new JObject
                        {
                            { "index", index },
                            { "text", WordComUtil.CleanParagraphText(
                                range == null ? string.Empty : range.Text) },
                            { "reference", WordComUtil.CleanParagraphText(
                                reference == null ? string.Empty : reference.Text) },
                            { "rangeStart", range == null ? 0 : range.Start },
                            { "paragraphIndex", range == null ? 0
                                : WordComUtil.FindParagraphIndexAt(document, range.Start) }
                        });
                    }
                    finally
                    {
                        ComHelper.Release(reference, range, note);
                    }
                }
                return count;
            }
            finally
            {
                ComHelper.Release(endnotes);
            }
        }

        // ------------------------------------------------------------------
        // word_content_controls
        // ------------------------------------------------------------------

        private OperationResult HandleContentControls(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string action = (args.GetOptionalString("action", "list") ?? "list")
                .Trim().ToLowerInvariant();
            if (action != "list" && action != "add" &&
                action != "remove" && action != "set")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "action 只支持 list/add/remove/set，收到「" + action + "」。");
            }

            if (action == "list")
            {
                return RunReadTool(
                    application,
                    delegate(Word.Document document)
                    {
                        JArray items = new JArray();
                        Word.ContentControls controls = null;
                        try
                        {
                            controls = document.ContentControls;
                            int count = controls == null ? 0 : controls.Count;
                            for (int controlIndex = 1; controlIndex <= count; controlIndex++)
                            {
                                Word.ContentControl control = null;
                                Word.Range range = null;
                                try
                                {
                                    control = controls[controlIndex];
                                    range = control.Range;
                                    items.Add(new JObject
                                    {
                                        { "index", controlIndex },
                                        { "type", control.Type.ToString() },
                                        { "title", control.Title ?? string.Empty },
                                        { "tag", control.Tag ?? string.Empty },
                                        { "text", WordComUtil.CleanParagraphText(
                                            range == null ? string.Empty : range.Text) },
                                        { "rangeStart", range == null ? 0 : range.Start },
                                        { "rangeEnd", range == null ? 0 : range.End },
                                        { "lockContents", control.LockContents },
                                        { "showingPlaceholderText", control.ShowingPlaceholderText }
                                    });
                                }
                                finally
                                {
                                    ComHelper.Release(range, control);
                                }
                            }
                            JObject data = new JObject();
                            data["action"] = "list";
                            data["count"] = count;
                            data["contentControls"] = items;
                            return ReadSucceed(
                                "已读取 " + count + " 个内容控件。", data, "Word:Document");
                        }
                        finally
                        {
                            ComHelper.Release(controls);
                        }
                    });
            }

            if (action == "add")
            {
                string typeText = (args.GetOptionalString("type", "text") ?? "text")
                    .Trim().ToLowerInvariant();
                Word.WdContentControlType controlType;
                if (!TryParseContentControlType(typeText, out controlType))
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.INVALID_PARAMS,
                        "type 只支持 text/richText/checkbox/dropdown/combobox/date/picture，" +
                        "收到「" + typeText + "」。");
                }
                string title = args.GetOptionalString("title", "AutoOfficeMCP");
                string tag = args.GetOptionalString("tag", "AutoOfficeMCP");
                string text = args.GetOptionalString("text", string.Empty);
                int paragraphIndex = args.GetOptionalInt("paragraphIndex", 0);
                int beforeCount = 0;
                return RunWriteTool(
                    application,
                    WordExtensionsToolNames.ContentControls,
                    false,
                    delegate(Word.Document document)
                    {
                        if (paragraphIndex > 0)
                        {
                            WordComUtil.ValidateParagraphIndex(document, paragraphIndex);
                            return WordWriteTargets.Paragraph(
                                document, paragraphIndex,
                                "Word:Paragraph:" + paragraphIndex, true);
                        }
                        return WordWriteTargets.Cursor(application);
                    },
                    delegate(Word.Document document, WordWriteTarget target)
                    {
                        Word.Selection selection = null;
                        Word.Range range = null;
                        Word.Paragraph paragraph = null;
                        Word.Range paragraphRange = null;
                        Word.ContentControls controls = null;
                        Word.ContentControl control = null;
                        try
                        {
                            beforeCount = CountContentControls(document);
                            if (paragraphIndex > 0)
                            {
                                paragraph = WordComUtil.GetParagraph(document, paragraphIndex);
                                paragraphRange = paragraph.Range;
                                range = WordComUtil.CreateRange(
                                    document, paragraphRange.Start, paragraphRange.Start);
                            }
                            else
                            {
                                selection = application.Selection;
                                if (selection == null)
                                {
                                    return WordWriteResult.Fail(
                                        OfficeErrorCode.INVALID_PARAMS,
                                        "当前没有有效选区，无法创建内容控件。",
                                        null,
                                        false);
                                }
                                range = selection.Range;
                            }
                            controls = range.ContentControls;
                            object rangeObject = range;
                            control = controls.Add(controlType, ref rangeObject);
                            if (control == null)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INTERNAL_ERROR,
                                    "ContentControls.Add 返回 null，内容控件未创建。",
                                    null,
                                    false);
                            }
                            control.Title = title ?? string.Empty;
                            control.Tag = tag ?? string.Empty;
                            if (!string.IsNullOrEmpty(text) &&
                                (controlType == Word.WdContentControlType.wdContentControlText ||
                                 controlType == Word.WdContentControlType.wdContentControlRichText))
                            {
                                Word.Range controlRange = control.Range;
                                try
                                {
                                    controlRange.Text = text;
                                }
                                finally
                                {
                                    ComHelper.Release(controlRange);
                                }
                            }
                            int afterCount = CountContentControls(document);
                            JObject data = new JObject();
                            data["action"] = "add";
                            data["type"] = controlType.ToString();
                            data["title"] = title ?? string.Empty;
                            data["tag"] = tag ?? string.Empty;
                            data["text"] = text ?? string.Empty;
                            data["countBefore"] = beforeCount;
                            data["countAfter"] = afterCount;
                            data["note"] = "内容控件创建于光标/段落起点；text 仅对 text/richText 类型写入。";
                            if (afterCount <= beforeCount)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INTERNAL_ERROR,
                                    "内容控件创建后数量未增加（" + beforeCount + "→" +
                                    afterCount + "）。请确认位置有效。",
                                    data,
                                    false);
                            }
                            return WordWriteResult.Ok(
                                "已创建内容控件（type=" + controlType + "，title=" + title +
                                "，当前共 " + afterCount + " 个）。",
                                data,
                                true).AddTargetRef("Word:ContentControls");
                        }
                        finally
                        {
                            ComHelper.Release(control, controls, paragraphRange,
                                paragraph, range, selection);
                        }
                    });
            }

            int index = args.RequireInt("index", 1, int.MaxValue);
            if (action == "remove")
            {
                bool confirm = args.GetOptionalBool("confirm", false);
                if (!confirm)
                {
                    return OperationResult.Cancelled(
                        "删除内容控件需要 confirm=true；本次未修改文档。");
                }
                bool deleteContents = args.GetOptionalBool("deleteContents", true);
                return RunWriteTool(
                    application,
                    WordExtensionsToolNames.ContentControls,
                    false,
                    delegate(Word.Document document)
                    {
                        return WordWriteTargets.None("内容控件删除", "Word:ContentControls");
                    },
                    delegate(Word.Document document, WordWriteTarget target)
                    {
                        int beforeCount = CountContentControls(document);
                        if (index > beforeCount)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "内容控件索引 " + index + " 越界（当前共 " + beforeCount +
                                " 个，1 基）。",
                                null,
                                false);
                        }
                        Word.ContentControls controls = null;
                        Word.ContentControl control = null;
                        try
                        {
                            controls = document.ContentControls;
                            control = controls[index];
                            control.Delete(deleteContents);
                            int afterCount = CountContentControls(document);
                            JObject data = new JObject();
                            data["action"] = "remove";
                            data["index"] = index;
                            data["deleteContents"] = deleteContents;
                            data["countBefore"] = beforeCount;
                            data["countAfter"] = afterCount;
                            if (afterCount >= beforeCount)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INTERNAL_ERROR,
                                    "ContentControl.Delete 已返回，但控件数量未减少（" +
                                    beforeCount + "→" + afterCount + "）。",
                                    data,
                                    false);
                            }
                            return WordWriteResult.Ok(
                                "已删除第 " + index + " 个内容控件（" + beforeCount + "→" +
                                afterCount + "）。",
                                data,
                                true).AddTargetRef("Word:ContentControls");
                        }
                        finally
                        {
                            ComHelper.Release(control, controls);
                        }
                    });
            }

            // action == set
            string newTitle = args.Has("title")
                ? args.GetOptionalString("title", string.Empty) : null;
            string newTag = args.Has("tag") ? args.GetOptionalString("tag", string.Empty) : null;
            string newText = args.Has("text")
                ? args.GetOptionalString("text", string.Empty) : null;
            if (newTitle == null && newTag == null && newText == null)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "action=set 至少需要 title/tag/text 之一。");
            }
            return RunWriteTool(
                application,
                WordExtensionsToolNames.ContentControls,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.None("内容控件设置", "Word:ContentControls");
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.ContentControls controls = null;
                    Word.ContentControl control = null;
                    Word.Range range = null;
                    try
                    {
                        controls = document.ContentControls;
                        int count = controls == null ? 0 : controls.Count;
                        if (index > count)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "内容控件索引 " + index + " 越界（当前共 " + count +
                                " 个，1 基）。",
                                null,
                                false);
                        }
                        control = controls[index];
                        JObject before = new JObject
                        {
                            { "title", control.Title ?? string.Empty },
                            { "tag", control.Tag ?? string.Empty }
                        };
                        if (newTitle != null)
                        {
                            control.Title = newTitle;
                        }
                        if (newTag != null)
                        {
                            control.Tag = newTag;
                        }
                        if (newText != null)
                        {
                            range = control.Range;
                            range.Text = newText;
                        }
                        JObject after = new JObject
                        {
                            { "title", control.Title ?? string.Empty },
                            { "tag", control.Tag ?? string.Empty }
                        };
                        if (newText != null)
                        {
                            if (range == null)
                            {
                                range = control.Range;
                            }
                            after["text"] = WordComUtil.CleanParagraphText(
                                range == null ? string.Empty : range.Text);
                        }
                        bool titleChanged = !string.Equals(
                            before.Value<string>("title"), after.Value<string>("title"),
                            StringComparison.Ordinal);
                        bool tagChanged = !string.Equals(
                            before.Value<string>("tag"), after.Value<string>("tag"),
                            StringComparison.Ordinal);
                        bool textChanged = newText != null && !string.Equals(
                            newText, after.Value<string>("text"), StringComparison.Ordinal);
                        bool changed = titleChanged || tagChanged || textChanged;
                        JObject data = new JObject();
                        data["action"] = "set";
                        data["index"] = index;
                        data["before"] = before;
                        data["after"] = after;
                        data["changedTitle"] = titleChanged;
                        data["changedTag"] = tagChanged;
                        data["changedText"] = textChanged;
                        return WordWriteResult.Ok(
                            "内容控件设置完成（title " + (titleChanged ? "已改" : "未变") +
                            "，tag " + (tagChanged ? "已改" : "未变") +
                            "，text " + (textChanged ? "已改" : "未变") + "）。",
                            data,
                            changed).AddTargetRef("Word:ContentControls");
                    }
                    finally
                    {
                        ComHelper.Release(range, control, controls);
                    }
                });
        }

        private static bool TryParseContentControlType(
            string text,
            out Word.WdContentControlType controlType)
        {
            controlType = Word.WdContentControlType.wdContentControlText;
            switch ((text ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "text":
                case "文本":
                    controlType = Word.WdContentControlType.wdContentControlText;
                    return true;
                case "richtext":
                case "rich":
                case "富文本":
                    controlType = Word.WdContentControlType.wdContentControlRichText;
                    return true;
                case "checkbox":
                case "check":
                case "复选框":
                    controlType = Word.WdContentControlType.wdContentControlCheckBox;
                    return true;
                case "dropdown":
                case "dropdownlist":
                case "下拉":
                    controlType = Word.WdContentControlType.wdContentControlDropdownList;
                    return true;
                case "combobox":
                case "组合框":
                    controlType = Word.WdContentControlType.wdContentControlComboBox;
                    return true;
                case "date":
                case "日期":
                    controlType = Word.WdContentControlType.wdContentControlDate;
                    return true;
                case "picture":
                case "图片":
                    controlType = Word.WdContentControlType.wdContentControlPicture;
                    return true;
                default:
                    return false;
            }
        }

        private static int CountContentControls(Word.Document document)
        {
            Word.ContentControls controls = null;
            try
            {
                controls = document.ContentControls;
                return controls == null ? 0 : controls.Count;
            }
            finally
            {
                ComHelper.Release(controls);
            }
        }

        // ------------------------------------------------------------------
        // word_line_numbering
        // ------------------------------------------------------------------

        private OperationResult HandleLineNumbering(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string action = (args.GetOptionalString("action", "get") ?? "get")
                .Trim().ToLowerInvariant();
            if (action != "get" && action != "set")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "action 只支持 get/set，收到「" + action + "」。");
            }
            int sectionIndex = args.GetOptionalInt("sectionIndex", 0);

            if (action == "get")
            {
                return RunReadTool(
                    application,
                    delegate(Word.Document document)
                    {
                        JArray items = ReadLineNumbering(document, sectionIndex);
                        JObject data = new JObject();
                        data["action"] = "get";
                        data["sectionIndex"] = sectionIndex;
                        data["sections"] = items;
                        return ReadSucceed(
                            "已读取 " + items.Count + " 个节的行号设置。",
                            data, "Word:Document");
                    });
            }

            bool? active = args.Has("active")
                ? args.GetOptionalBool("active", true) : (bool?)null;
            int? countBy = args.Has("countBy") ? args.GetOptionalInt("countBy") : (int?)null;
            int? startingNumber = args.Has("startingNumber")
                ? args.GetOptionalInt("startingNumber") : (int?)null;
            double? distanceFromText = args.Has("distanceFromTextPt")
                ? args.GetOptionalDouble("distanceFromTextPt") : (double?)null;
            Word.WdNumberingRule? restartMode = null;
            if (args.Has("restartMode"))
            {
                string modeText = (args.GetOptionalString("restartMode", null) ?? string.Empty)
                    .Trim().ToLowerInvariant();
                if (modeText == "continuous" || modeText == "连续")
                {
                    restartMode = Word.WdNumberingRule.wdRestartContinuous;
                }
                else if (modeText == "section" || modeText == "每节")
                {
                    restartMode = Word.WdNumberingRule.wdRestartSection;
                }
                else if (modeText == "page" || modeText == "每页")
                {
                    restartMode = Word.WdNumberingRule.wdRestartPage;
                }
                else
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.INVALID_PARAMS,
                        "restartMode 只支持 continuous/section/page，收到「" + modeText + "」。");
                }
            }
            if (!active.HasValue && !countBy.HasValue && !startingNumber.HasValue &&
                !distanceFromText.HasValue && !restartMode.HasValue)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "action=set 至少需要 active/countBy/startingNumber/distanceFromTextPt/restartMode 之一。");
            }
            if (countBy.HasValue && (countBy.Value < 1 || countBy.Value > 100))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "countBy 允许 1-100，收到 " + countBy.Value + "。");
            }
            if (startingNumber.HasValue && startingNumber.Value < 0)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "startingNumber 不能为负数，收到 " + startingNumber.Value + "。");
            }
            if (distanceFromText.HasValue && distanceFromText.Value < 0)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "distanceFromTextPt 不能为负数，收到 " + distanceFromText.Value + "。");
            }

            return RunWriteTool(
                application,
                WordExtensionsToolNames.LineNumbering,
                false,
                delegate(Word.Document document)
                {
                    if (sectionIndex > 0)
                    {
                        WordComUtil.ValidateSectionIndex(document, sectionIndex);
                    }
                    return WordWriteTargets.None("行号设置", "Word:Sections");
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Sections sections = null;
                    JArray items = new JArray();
                    JArray failures = new JArray();
                    int changedCount = 0;
                    try
                    {
                        sections = document.Sections;
                        int count = sections == null ? 0 : sections.Count;
                        List<int> indexes = new List<int>();
                        if (sectionIndex > 0)
                        {
                            indexes.Add(sectionIndex);
                        }
                        else
                        {
                            for (int index = 1; index <= count; index++)
                            {
                                indexes.Add(index);
                            }
                        }
                        indexes.Sort();
                        indexes.Reverse();
                        foreach (int index in indexes)
                        {
                            Word.Section section = null;
                            Word.PageSetup pageSetup = null;
                            Word.LineNumbering lineNumbering = null;
                            try
                            {
                                section = sections[index];
                                pageSetup = section.PageSetup;
                                lineNumbering = pageSetup.LineNumbering;
                                JObject before = CaptureLineNumbering(lineNumbering);
                                if (active.HasValue)
                                {
                                    lineNumbering.Active = active.Value ? -1 : 0;
                                }
                                if (countBy.HasValue)
                                {
                                    lineNumbering.CountBy = countBy.Value;
                                }
                                if (startingNumber.HasValue)
                                {
                                    lineNumbering.StartingNumber = startingNumber.Value;
                                }
                                if (distanceFromText.HasValue)
                                {
                                    lineNumbering.DistanceFromText =
                                        (float)distanceFromText.Value;
                                }
                                if (restartMode.HasValue)
                                {
                                    lineNumbering.RestartMode = restartMode.Value;
                                }
                                JObject after = CaptureLineNumbering(lineNumbering);
                                bool changed = !JToken.DeepEquals(before, after);
                                if (changed)
                                {
                                    changedCount++;
                                }
                                items.Add(new JObject
                                {
                                    { "sectionIndex", index },
                                    { "before", before },
                                    { "after", after },
                                    { "changed", changed }
                                });
                            }
                            catch (Exception ex)
                            {
                                failures.Add(new JObject
                                {
                                    { "sectionIndex", index },
                                    { "reason", OfficeErrorMapper.DescribeException(
                                        ex, OfficeApp.Word) }
                                });
                            }
                            finally
                            {
                                ComHelper.Release(lineNumbering, pageSetup, section);
                            }
                        }
                        JObject data = new JObject();
                        data["action"] = "set";
                        data["sectionIndex"] = sectionIndex;
                        data["changedSections"] = changedCount;
                        data["sections"] = items;
                        data["failures"] = failures;
                        data["note"] = "行号属于页面设置，T02 快照引擎不覆盖；Word 自定义撤销记录已尝试开启。";
                        if (failures.Count > 0 && changedCount == 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INTERNAL_ERROR,
                                "行号设置失败：首个原因 " +
                                failures[0]["reason"].ToString() + "。",
                                data,
                                false);
                        }
                        if (failures.Count > 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.PARTIAL_APPLY,
                                "行号设置部分完成：成功 " + changedCount + " 节，失败 " +
                                failures.Count + " 节；不会自动回滚。",
                                data,
                                true);
                        }
                        return WordWriteResult.Ok(
                            "已完成 " + changedCount + "/" + count + " 节的行号设置" +
                            (changedCount == 0 ? "（设置与当前值相同）" : "") + "。",
                            data,
                            changedCount > 0).AddTargetRef("Word:Sections");
                    }
                    finally
                    {
                        ComHelper.Release(sections);
                    }
                });
        }

        private static JArray ReadLineNumbering(Word.Document document, int sectionIndex)
        {
            JArray items = new JArray();
            Word.Sections sections = null;
            try
            {
                sections = document.Sections;
                int count = sections == null ? 0 : sections.Count;
                if (sectionIndex > count)
                {
                    throw WordArguments.Invalid(
                        "sectionIndex " + sectionIndex + " 不存在（当前共 " + count +
                        " 节，1 基）。");
                }
                List<int> indexes = new List<int>();
                if (sectionIndex > 0)
                {
                    indexes.Add(sectionIndex);
                }
                else
                {
                    for (int index = 1; index <= count; index++)
                    {
                        indexes.Add(index);
                    }
                }
                foreach (int index in indexes)
                {
                    Word.Section section = null;
                    Word.PageSetup pageSetup = null;
                    Word.LineNumbering lineNumbering = null;
                    try
                    {
                        section = sections[index];
                        pageSetup = section.PageSetup;
                        lineNumbering = pageSetup.LineNumbering;
                        JObject item = CaptureLineNumbering(lineNumbering);
                        item["sectionIndex"] = index;
                        items.Add(item);
                    }
                    finally
                    {
                        ComHelper.Release(lineNumbering, pageSetup, section);
                    }
                }
                return items;
            }
            finally
            {
                ComHelper.Release(sections);
            }
        }

        private static JObject CaptureLineNumbering(Word.LineNumbering lineNumbering)
        {
            JObject json = new JObject();
            if (lineNumbering == null)
            {
                json["available"] = false;
                return json;
            }
            json["active"] = lineNumbering.Active != 0;
            json["activeCode"] = lineNumbering.Active;
            json["countBy"] = lineNumbering.CountBy;
            json["startingNumber"] = lineNumbering.StartingNumber;
            json["distanceFromTextPt"] = (double)lineNumbering.DistanceFromText;
            json["restartMode"] = lineNumbering.RestartMode.ToString();
            return json;
        }
    }
}
