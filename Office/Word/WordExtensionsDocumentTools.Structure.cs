using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    public sealed partial class WordExtensionsExecutor
    {
        private static readonly Regex MarkdownHeadingRegex =
            new Regex(@"^(#{1,6})\s+(.*)$", RegexOptions.Compiled);
        private static readonly Regex MarkdownBulletRegex =
            new Regex(@"^(\s*)[-*+]\s+(.*)$", RegexOptions.Compiled);
        private static readonly Regex MarkdownOrderedRegex =
            new Regex(@"^(\s*)[0-9]+[.)]\s+(.*)$", RegexOptions.Compiled);
        private static readonly Regex MarkdownTableSeparatorRegex =
            new Regex(@"^:?-{3,}:?$", RegexOptions.Compiled);

        private void RegisterStructureExtensionHandlers(
            Dictionary<string, Func<Word.Application, JObject, OperationResult>> handlers)
        {
            handlers[WordExtensionsToolNames.MdExport] = HandleMdExport;
            handlers[WordExtensionsToolNames.MdImport] = HandleMdImport;
            handlers[WordExtensionsToolNames.TableConvert] = HandleTableConvert;
            handlers[WordExtensionsToolNames.SectionManage] = HandleSectionManage;
        }

        // ------------------------------------------------------------------
        // word_md_export（只读文档；可选写 .md 文件）
        // ------------------------------------------------------------------

        private OperationResult HandleMdExport(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            bool includeTables = args.GetOptionalBool("includeTables", true);
            bool includeLists = args.GetOptionalBool("includeLists", true);
            bool includeImages = args.GetOptionalBool("includeImages", true);
            bool confirm = args.GetOptionalBool("confirm", false);
            int maxCharacters = args.GetOptionalInt("maxCharacters", 200000);
            if (maxCharacters < 1000 || maxCharacters > 2000000)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "maxCharacters 允许 1000-2000000，收到 " + maxCharacters + "。");
            }

            string outputPath = null;
            bool outputExisted = false;
            if (args.Has("outputPath"))
            {
                string requested = args.GetOptionalString("outputPath", null);
                OperationResult pathFailure = WordExtensionsUtil.PrepareOutputFile(
                    requested, ".md", confirm, out outputPath, out outputExisted);
                if (pathFailure != null)
                {
                    return pathFailure;
                }
            }

            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    MarkdownExportResult export = BuildMarkdown(
                        document, includeTables, includeLists, includeImages, maxCharacters);
                    JObject data = new JObject();
                    data["documentName"] = document.Name;
                    data["markdown"] = export.Markdown;
                    data["markdownLength"] = export.Markdown.Length;
                    data["truncated"] = export.Truncated;
                    data["maxCharacters"] = maxCharacters;
                    data["tableCount"] = export.TableCount;
                    data["paragraphCount"] = export.ParagraphCount;
                    data["headingCount"] = export.HeadingCount;
                    data["listCount"] = export.ListCount;
                    data["imagePlaceholderCount"] = export.ImagePlaceholderCount;
                    data["warnings"] = export.Warnings;
                    data["documentModified"] = false;

                    string message;
                    if (!string.IsNullOrWhiteSpace(outputPath))
                    {
                        WordExtensionsUtil.WriteUtf8NoBom(outputPath, export.Markdown);
                        FileInfo info = new FileInfo(outputPath);
                        data["outputPath"] = outputPath;
                        data["bytes"] = info.Length;
                        data["overwroteExisting"] = outputExisted;
                        message = "已导出 Markdown：" + outputPath + "（" + export.Markdown.Length +
                                  " 字符，" + info.Length + " 字节；文档未修改）。";
                    }
                    else
                    {
                        message = "已生成 Markdown（" + export.Markdown.Length + " 字符，" +
                                  export.TableCount + " 个表格、" + export.HeadingCount +
                                  " 个标题；文档未修改）。";
                    }
                    if (export.Truncated)
                    {
                        message = message + " 内容超过 maxCharacters=" + maxCharacters + "，已截断。";
                    }

                    if (string.IsNullOrWhiteSpace(outputPath))
                    {
                        return ReadSucceed(message, data, "Word:Document");
                    }
                    return ReadSucceed(message, data, "Word:Document", "File:" + outputPath);
                });
        }

        private static MarkdownExportResult BuildMarkdown(
            Word.Document document,
            bool includeTables,
            bool includeLists,
            bool includeImages,
            int maxCharacters)
        {
            MarkdownExportResult result = new MarkdownExportResult();
            List<MarkdownEvent> events = new List<MarkdownEvent>();
            List<TableRangeInfo> tableRanges = new List<TableRangeInfo>();
            Word.Tables tables = null;
            Word.Paragraphs paragraphs = null;
            Dictionary<int, List<int>> floatingImagesByParagraph =
                CollectFloatingImageAnchors(document, result.Warnings);
            try
            {
                tables = document.Tables;
                int tableCount = tables == null ? 0 : tables.Count;
                for (int index = 1; index <= tableCount; index++)
                {
                    Word.Table table = null;
                    Word.Range range = null;
                    try
                    {
                        table = tables[index];
                        range = table.Range;
                        TableRangeInfo rangeInfo = new TableRangeInfo();
                        rangeInfo.Start = range.Start;
                        rangeInfo.End = range.End;
                        tableRanges.Add(rangeInfo);
                        if (includeTables)
                        {
                            string tableMarkdown = BuildTableMarkdown(
                                document, table, index, result.Warnings);
                            events.Add(new MarkdownEvent(range.Start, tableMarkdown));
                            result.TableCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        WordExtensionsUtil.AddWarning(
                            result.Warnings,
                            "第 " + index + " 个表格导出失败：" +
                            OfficeErrorMapper.DescribeException(ex, OfficeApp.Word));
                    }
                    finally
                    {
                        ComHelper.Release(range, table);
                    }
                }

                paragraphs = document.Paragraphs;
                int paragraphCount = paragraphs == null ? 0 : paragraphs.Count;
                for (int index = 1; index <= paragraphCount; index++)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range range = null;
                    Word.ListFormat listFormat = null;
                    Word.InlineShapes inlineShapes = null;
                    try
                    {
                        paragraph = paragraphs[index];
                        range = paragraph.Range;
                        if (IsInsideTableRange(range, tableRanges))
                        {
                            continue;
                        }
                        string text = WordExtensionsUtil.ReadParagraphText(paragraph);
                        string style = WordComUtil.ReadStyleName(range);
                        int headingLevel = WordExtensionsUtil.TryGetHeadingLevel(style);
                        int listLevel = 0;
                        bool isList = false;
                        bool orderedList = false;
                        string listString = string.Empty;
                        if (includeLists)
                        {
                            listFormat = range.ListFormat;
                            if (listFormat != null &&
                                listFormat.ListType != Word.WdListType.wdListNoNumbering)
                            {
                                isList = true;
                                listLevel = Math.Max(1, listFormat.ListLevelNumber);
                                listString = listFormat.ListString ?? string.Empty;
                                orderedList = listFormat.ListType != Word.WdListType.wdListBullet ||
                                              LooksOrderedListString(listString);
                            }
                        }

                        StringBuilder builder = new StringBuilder();
                        if (headingLevel > 0)
                        {
                            builder.Append(new string('#', headingLevel));
                            builder.Append(' ');
                            result.HeadingCount++;
                        }
                        else if (isList)
                        {
                            builder.Append(new string(' ', (listLevel - 1) * 2));
                            builder.Append(orderedList ? "1. " : "- ");
                            result.ListCount++;
                        }
                        builder.Append(text);

                        int imageCount = 0;
                        if (includeImages)
                        {
                            inlineShapes = range.InlineShapes;
                            int inlineCount = inlineShapes == null ? 0 : inlineShapes.Count;
                            for (int inlineIndex = 1; inlineIndex <= inlineCount; inlineIndex++)
                            {
                                Word.InlineShape shape = null;
                                try
                                {
                                    shape = inlineShapes[inlineIndex];
                                    bool isPicture =
                                        shape.Type == Word.WdInlineShapeType.wdInlineShapePicture ||
                                        shape.Type == Word.WdInlineShapeType.wdInlineShapeLinkedPicture;
                                    if (isPicture)
                                    {
                                        builder.Append(" ![inline ").Append(inlineIndex).Append("](word-inline:")
                                               .Append(inlineIndex).Append(')');
                                        imageCount++;
                                    }
                                }
                                finally
                                {
                                    ComHelper.Release(shape);
                                }
                            }
                            List<int> floatingIndexes;
                            if (floatingImagesByParagraph.TryGetValue(index, out floatingIndexes))
                            {
                                foreach (int shapeIndex in floatingIndexes)
                                {
                                    builder.Append(" ![floating ").Append(shapeIndex)
                                           .Append("](word-floating:").Append(shapeIndex).Append(')');
                                    imageCount++;
                                }
                            }
                        }
                        result.ImagePlaceholderCount += imageCount;
                        result.ParagraphCount++;
                        if (text.Length > 0 || imageCount > 0)
                        {
                            events.Add(new MarkdownEvent(range.Start, builder.ToString()));
                        }
                    }
                    catch (Exception ex)
                    {
                        WordExtensionsUtil.AddWarning(
                            result.Warnings,
                            "第 " + index + " 段导出失败：" +
                            OfficeErrorMapper.DescribeException(ex, OfficeApp.Word));
                    }
                    finally
                    {
                        ComHelper.Release(inlineShapes, listFormat, range, paragraph);
                    }
                }

                events.Sort(delegate(MarkdownEvent left, MarkdownEvent right)
                {
                    return left.Position.CompareTo(right.Position);
                });
                StringBuilder markdown = new StringBuilder();
                bool first = true;
                foreach (MarkdownEvent item in events)
                {
                    if (string.IsNullOrWhiteSpace(item.Text))
                    {
                        continue;
                    }
                    if (!first)
                    {
                        markdown.Append("\r\n\r\n");
                    }
                    markdown.Append(item.Text.TrimEnd());
                    first = false;
                    if (markdown.Length >= maxCharacters)
                    {
                        result.Truncated = true;
                        break;
                    }
                }
                if (markdown.Length > maxCharacters)
                {
                    markdown.Length = maxCharacters;
                    result.Truncated = true;
                }
                result.Markdown = markdown.ToString();
                return result;
            }
            finally
            {
                ComHelper.Release(paragraphs, tables);
            }
        }

        private static Dictionary<int, List<int>> CollectFloatingImageAnchors(
            Word.Document document,
            JArray warnings)
        {
            Dictionary<int, List<int>> map = new Dictionary<int, List<int>>();
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
                        if (typeCode != 13 && typeCode != 11 && typeCode != 28 && typeCode != 29)
                        {
                            continue;
                        }
                        anchor = shape.Anchor;
                        int paragraphIndex = WordComUtil.FindParagraphIndexAt(
                            document, anchor == null ? 0 : anchor.Start);
                        List<int> list;
                        if (!map.TryGetValue(paragraphIndex, out list))
                        {
                            list = new List<int>();
                            map[paragraphIndex] = list;
                        }
                        list.Add(index);
                    }
                    catch (Exception ex)
                    {
                        WordExtensionsUtil.AddWarning(
                            warnings,
                            "第 " + index + " 个浮动对象锚点读取失败，已跳过：" +
                            OfficeErrorMapper.DescribeException(ex, OfficeApp.Word));
                    }
                    finally
                    {
                        ComHelper.Release(anchor, shape);
                    }
                }
                return map;
            }
            finally
            {
                ComHelper.Release(shapes);
            }
        }

        private static bool IsInsideTableRange(Word.Range range, List<TableRangeInfo> tableRanges)
        {
            if (range == null || tableRanges == null)
            {
                return false;
            }
            foreach (TableRangeInfo tableRange in tableRanges)
            {
                if (range.Start >= tableRange.Start && range.Start < tableRange.End)
                {
                    return true;
                }
            }
            return false;
        }

        private static bool LooksOrderedListString(string listString)
        {
            if (string.IsNullOrWhiteSpace(listString))
            {
                return false;
            }
            char first = listString.Trim()[0];
            return char.IsDigit(first) || char.IsLetter(first);
        }

        private static string BuildTableMarkdown(
            Word.Document document,
            Word.Table table,
            int tableIndex,
            JArray warnings)
        {
            Word.Rows rows = null;
            Word.Columns columns = null;
            try
            {
                rows = table.Rows;
                columns = table.Columns;
                int rowCount = rows == null ? 0 : rows.Count;
                int columnCount = columns == null ? 0 : columns.Count;
                if (rowCount <= 0 || columnCount <= 0)
                {
                    return string.Empty;
                }
                StringBuilder builder = new StringBuilder();
                for (int row = 1; row <= rowCount; row++)
                {
                    List<string> cells = new List<string>();
                    for (int column = 1; column <= columnCount; column++)
                    {
                        string cellText = string.Empty;
                        Word.Cell cell = null;
                        Word.Range cellRange = null;
                        try
                        {
                            cell = table.Cell(row, column);
                            cellRange = cell.Range;
                            cellText = WordComUtil.CleanCellText(cellRange.Text);
                        }
                        catch (Exception ex)
                        {
                            WordExtensionsUtil.AddWarning(
                                warnings,
                                "表格 " + tableIndex + " 第 " + row + " 行第 " + column +
                                " 列读取失败（可能为合并单元格）：" + ex.Message);
                        }
                        finally
                        {
                            ComHelper.Release(cellRange, cell);
                        }
                        cellText = cellText.Replace("\r", " ").Replace("\n", " ")
                                             .Replace("|", "\\|").Trim();
                        cells.Add(cellText);
                    }
                    builder.Append("| ");
                    builder.Append(string.Join(" | ", cells.ToArray()));
                    builder.Append(" |");
                    if (row == 1)
                    {
                        builder.Append("\r\n|");
                        for (int column = 0; column < columnCount; column++)
                        {
                            builder.Append(" --- |");
                        }
                    }
                    if (row < rowCount)
                    {
                        builder.Append("\r\n");
                    }
                }
                return builder.ToString();
            }
            finally
            {
                ComHelper.Release(columns, rows);
            }
        }

        private sealed class MarkdownExportResult
        {
            public string Markdown { get; set; }

            public bool Truncated { get; set; }

            public int TableCount { get; set; }

            public int ParagraphCount { get; set; }

            public int HeadingCount { get; set; }

            public int ListCount { get; set; }

            public int ImagePlaceholderCount { get; set; }

            public JArray Warnings { get; private set; }

            public MarkdownExportResult()
            {
                Markdown = string.Empty;
                Warnings = new JArray();
            }
        }

        private sealed class MarkdownEvent
        {
            public int Position { get; private set; }

            public string Text { get; private set; }

            public MarkdownEvent(int position, string text)
            {
                Position = position;
                Text = text ?? string.Empty;
            }
        }

        private sealed class TableRangeInfo
        {
            public int Start { get; set; }

            public int End { get; set; }
        }

        // ------------------------------------------------------------------
        // word_md_import
        // ------------------------------------------------------------------

        private OperationResult HandleMdImport(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string markdown = null;
            bool hasInlineMarkdown = args.Has("markdown");
            if (hasInlineMarkdown)
            {
                markdown = args.GetOptionalString("markdown", string.Empty);
            }
            string inputPath = null;
            if (args.Has("inputPath"))
            {
                inputPath = args.GetOptionalString("inputPath", null);
                if (!File.Exists(inputPath))
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.FILE_NOT_FOUND,
                        "Markdown 输入文件不存在：" + inputPath + "。请检查路径后重试。");
                }
                if (!hasInlineMarkdown)
                {
                    try
                    {
                        markdown = File.ReadAllText(inputPath, Encoding.UTF8);
                    }
                    catch (Exception ex)
                    {
                        return OperationResult.Failed(
                            OfficeErrorCode.READ_ONLY,
                            "无法读取 Markdown 文件：" + inputPath + "；" + ex.Message);
                    }
                }
            }
            if (string.IsNullOrWhiteSpace(markdown))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "缺少 markdown 文本或 inputPath。请传 markdown 字符串，或传存在的 inputPath 文件。");
            }
            if (markdown.Length > 2000000)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "Markdown 内容过长（" + markdown.Length + " 字符），上限 2000000 字符。");
            }

            string position = (args.GetOptionalString("position", "cursor") ?? "cursor")
                .Trim().ToLowerInvariant();
            if (position != "start" && position != "end" &&
                position != "cursor" && position != "paragraph")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "position 只支持 start/end/cursor/paragraph，收到「" + position + "」。");
            }
            int paragraphIndex = args.GetOptionalInt("paragraphIndex", 0);
            if (position == "paragraph" && paragraphIndex <= 0)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "position=paragraph 时必须提供 paragraphIndex（1 基）。收到 " + paragraphIndex + "。");
            }
            int headingOffset = args.GetOptionalInt("headingOffset", 0);
            if (headingOffset < -5 || headingOffset > 5)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "headingOffset 允许 -5 到 5，收到 " + headingOffset + "。");
            }
            bool applyStyles = args.GetOptionalBool("applyStyles", true);

            JArray warnings = new JArray();
            List<MarkdownParagraph> blocks = ParseMarkdown(markdown, warnings);
            if (blocks.Count == 0)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "Markdown 解析后没有可导入的段落。支持 # 标题、- 列表、1. 列表、普通段落；" +
                    "表格行会按制表符文本导入。");
            }

            return RunWriteTool(
                application,
                WordExtensionsToolNames.MdImport,
                false,
                delegate(Word.Document document)
                {
                    if (position == "start")
                    {
                        return WordWriteTargets.DocumentStart();
                    }
                    if (position == "end")
                    {
                        return WordWriteTargets.DocumentEnd(document);
                    }
                    if (position == "paragraph")
                    {
                        WordComUtil.ValidateParagraphIndex(document, paragraphIndex);
                        return WordWriteTargets.Paragraph(
                            document, paragraphIndex, "Word:Paragraph:" + paragraphIndex, true);
                    }
                    return WordWriteTargets.Cursor(application);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Range insertion = null;
                    try
                    {
                        int insertionStart = target.Start;
                        bool needsLeadingBreak = position == "cursor" || position == "paragraph";
                        if (needsLeadingBreak)
                        {
                            insertion = WordComUtil.CreateRange(
                                document, insertionStart, insertionStart);
                            insertion.InsertAfter("\r");
                            insertionStart++;
                            ComHelper.Release(insertion);
                            insertion = null;
                        }

                        StringBuilder textBuilder = new StringBuilder();
                        foreach (MarkdownParagraph block in blocks)
                        {
                            textBuilder.Append(block.PlainText);
                            textBuilder.Append('\r');
                        }
                        string insertedText = textBuilder.ToString();
                        insertion = WordComUtil.CreateRange(
                            document, insertionStart, insertionStart);
                        insertion.InsertAfter(insertedText);
                        int insertedEnd = insertionStart + insertedText.Length;

                        List<int> insertedParagraphIndexes = FindInsertedParagraphIndexes(
                            document, insertionStart, insertedEnd);
                        JArray appliedHeadings = new JArray();
                        JArray appliedLists = new JArray();
                        JArray failures = new JArray();
                        int headingsApplied = 0;
                        int listsApplied = 0;
                        int paragraphsVerified = 0;

                        int count = Math.Min(insertedParagraphIndexes.Count, blocks.Count);
                        for (int index = 0; index < count; index++)
                        {
                            MarkdownParagraph block = blocks[index];
                            int paragraphNumber = insertedParagraphIndexes[index];
                            Word.Paragraph paragraph = null;
                            Word.Range range = null;
                            Word.ListFormat listFormat = null;
                            try
                            {
                                paragraph = WordComUtil.GetParagraph(document, paragraphNumber);
                                range = paragraph.Range;
                                string actualText = WordExtensionsUtil.ReadParagraphText(paragraph);
                                if (string.Equals(actualText, block.PlainText, StringComparison.Ordinal))
                                {
                                    paragraphsVerified++;
                                }
                                else
                                {
                                    failures.Add(new JObject
                                    {
                                        { "blockIndex", index + 1 },
                                        { "paragraphIndex", paragraphNumber },
                                        { "reason", "插入后回读文本不一致；预期「" +
                                                    WordExtensionsUtil.Preview(block.PlainText, 80) +
                                                    "」，实际「" +
                                                    WordExtensionsUtil.Preview(actualText, 80) + "」" }
                                    });
                                }

                                if (applyStyles)
                                {
                                    if (block.Kind == "heading")
                                    {
                                        int level = Math.Max(1, Math.Min(9,
                                            block.Level + headingOffset));
                                        string styleName = WordExtensionsUtil.ResolveHeadingStyleName(
                                            document, level);
                                        if (!string.IsNullOrWhiteSpace(styleName))
                                        {
                                            WordExtensionsUtil.ApplyParagraphStyle(
                                                document, paragraph, styleName, false, false);
                                            headingsApplied++;
                                            appliedHeadings.Add(new JObject
                                            {
                                                { "blockIndex", index + 1 },
                                                { "paragraphIndex", paragraphNumber },
                                                { "level", level },
                                                { "style", styleName }
                                            });
                                        }
                                        else
                                        {
                                            WordExtensionsUtil.AddWarning(
                                                warnings,
                                                "第 " + (index + 1) +
                                                " 个块需要「标题 " + level +
                                                "」样式，但当前文档不存在该样式，已保留正文格式。");
                                        }
                                    }
                                    else if (block.Kind == "bullet" || block.Kind == "ordered")
                                    {
                                        listFormat = range.ListFormat;
                                        string listError = ApplyMarkdownListFormat(
                                            application, range, listFormat,
                                            block.Kind == "ordered");
                                        if (string.IsNullOrWhiteSpace(listError))
                                        {
                                            listsApplied++;
                                            appliedLists.Add(new JObject
                                            {
                                                { "blockIndex", index + 1 },
                                                { "paragraphIndex", paragraphNumber },
                                                { "kind", block.Kind },
                                                { "level", block.Level },
                                                { "listString", listFormat.ListString ?? string.Empty }
                                            });
                                        }
                                        else
                                        {
                                            WordExtensionsUtil.AddWarning(warnings, listError);
                                        }
                                    }
                                }

                                ApplyInlineRuns(document, paragraph, block, failures, index + 1);
                            }
                            catch (Exception ex)
                            {
                                failures.Add(new JObject
                                {
                                    { "blockIndex", index + 1 },
                                    { "paragraphIndex", paragraphNumber },
                                    { "reason", OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) }
                                });
                            }
                            finally
                            {
                                ComHelper.Release(listFormat, range, paragraph);
                            }
                        }

                        if (insertedParagraphIndexes.Count < blocks.Count)
                        {
                            WordExtensionsUtil.AddWarning(
                                warnings,
                                "Markdown 解析出 " + blocks.Count + " 个块，但只找到 " +
                                insertedParagraphIndexes.Count + " 个插入段落；后 " +
                                (blocks.Count - insertedParagraphIndexes.Count) +
                                " 个块未应用样式。");
                        }

                        JObject data = new JObject();
                        data["position"] = position;
                        data["paragraphIndex"] = paragraphIndex > 0 ? paragraphIndex : 0;
                        data["blockCount"] = blocks.Count;
                        data["insertedParagraphs"] = insertedParagraphIndexes.Count;
                        data["paragraphsVerified"] = paragraphsVerified;
                        data["headingsApplied"] = headingsApplied;
                        data["listsApplied"] = listsApplied;
                        data["failureCount"] = failures.Count;
                        data["headings"] = appliedHeadings;
                        data["lists"] = appliedLists;
                        data["failures"] = failures;
                        data["warnings"] = warnings;
                        data["applyStyles"] = applyStyles;
                        data["headingOffset"] = headingOffset;
                        data["tablesImportedAsText"] = CountBlocks(blocks, "tableRow");
                        data["note"] = "Markdown 子集导入：标题 (#)、无序/有序列表、普通段落、表格行（按制表符文本）；" +
                                        "支持 **粗体**、*斜体*、__粗体__、_斜体_ 的行内格式；" +
                                        "导入内容插入到指定位置，不覆盖原文档其他部分。";

                        if (paragraphsVerified == 0 && failures.Count > 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INTERNAL_ERROR,
                                "Markdown 插入后回读校验全部失败；首个失败：" +
                                failures[0]["reason"].ToString() +
                                "。请检查文档是否受保护或位置是否有效。",
                                data,
                                false);
                        }
                        if (failures.Count > 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.PARTIAL_APPLY,
                                "Markdown 导入部分完成：解析 " + blocks.Count + " 块，成功应用样式 " +
                                paragraphsVerified + " 段，失败 " + failures.Count +
                                " 项；首个失败：" + failures[0]["reason"].ToString() +
                                "。已插入内容不会自动回滚。",
                                data,
                                true);
                        }
                        return WordWriteResult.Ok(
                            "已从 Markdown 导入 " + blocks.Count + " 个块（插入 " +
                            insertedParagraphIndexes.Count + " 段，标题样式 " + headingsApplied +
                            " 个，列表 " + listsApplied + " 个；" +
                            CountBlocks(blocks, "tableRow") + " 个表格行按文本导入）。",
                            data,
                            true).AddTargetRef("Word:Document");
                    }
                    finally
                    {
                        ComHelper.Release(insertion);
                    }
                });
        }

        private static List<int> FindInsertedParagraphIndexes(
            Word.Document document,
            int insertedStart,
            int insertedEnd)
        {
            List<int> result = new List<int>();
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
                        if (range.Start < insertedStart || range.Start >= insertedEnd + 1)
                        {
                            continue;
                        }
                        string text = WordExtensionsUtil.ReadParagraphText(paragraph);
                        if (text.Length > 0)
                        {
                            result.Add(index);
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
            return result;
        }

        private static string ApplyMarkdownListFormat(
            Word.Application application,
            Word.Range range,
            Word.ListFormat listFormat,
            bool ordered)
        {
            Word.ListGalleries galleries = null;
            Word.ListGallery gallery = null;
            Word.ListTemplates templates = null;
            Word.ListTemplate template = null;
            try
            {
                galleries = application.ListGalleries;
                Word.WdListGalleryType galleryType = ordered
                    ? Word.WdListGalleryType.wdNumberGallery
                    : Word.WdListGalleryType.wdBulletGallery;
                gallery = galleries[galleryType];
                templates = gallery == null ? null : gallery.ListTemplates;
                if (templates == null || templates.Count == 0)
                {
                    return "Word 列表库为空，无法应用" + (ordered ? "有序" : "无序") + "列表。";
                }
                template = templates[1];
                object continuePrevious = true;
                object applyTo = Word.WdListApplyTo.wdListApplyToWholeList;
                object behavior = Word.WdDefaultListBehavior.wdWord9ListBehavior;
                listFormat.ApplyListTemplate(
                    template, ref continuePrevious, ref applyTo, ref behavior);
                return string.Empty;
            }
            catch (Exception ex)
            {
                return "应用" + (ordered ? "有序" : "无序") + "列表格式失败：" +
                       OfficeErrorMapper.DescribeException(ex, OfficeApp.Word);
            }
            finally
            {
                ComHelper.Release(template, templates, gallery, galleries);
            }
        }

        private static void ApplyInlineRuns(
            Word.Document document,
            Word.Paragraph paragraph,
            MarkdownParagraph block,
            JArray failures,
            int blockNumber)
        {
            if (block.Runs == null || block.Runs.Count == 0)
            {
                return;
            }
            bool hasFormatting = false;
            foreach (MarkdownInlineRun run in block.Runs)
            {
                if (run.Bold || run.Italic)
                {
                    hasFormatting = true;
                    break;
                }
            }
            if (!hasFormatting)
            {
                return;
            }

            Word.Range paragraphRange = null;
            Word.Font font = null;
            try
            {
                paragraphRange = paragraph.Range;
                string actualText = WordExtensionsUtil.ReadParagraphText(paragraph);
                int searchStart = 0;
                foreach (MarkdownInlineRun run in block.Runs)
                {
                    if (!run.Bold && !run.Italic)
                    {
                        searchStart += run.Text.Length;
                        continue;
                    }
                    int found = actualText.IndexOf(
                        run.Text, Math.Min(searchStart, actualText.Length),
                        StringComparison.Ordinal);
                    if (found < 0)
                    {
                        failures.Add(new JObject
                        {
                            { "blockIndex", blockNumber },
                            { "reason", "行内格式文本未找到：「" +
                                        WordExtensionsUtil.Preview(run.Text, 60) + "」" }
                        });
                        searchStart += run.Text.Length;
                        continue;
                    }
                    Word.Range runRange = null;
                    try
                    {
                        runRange = WordComUtil.CreateRange(
                            document,
                            paragraphRange.Start + found,
                            paragraphRange.Start + found + run.Text.Length);
                        font = runRange.Font;
                        if (run.Bold)
                        {
                            font.Bold = -1;
                        }
                        if (run.Italic)
                        {
                            font.Italic = -1;
                        }
                        ComHelper.Release(font);
                        font = null;
                    }
                    finally
                    {
                        ComHelper.Release(runRange);
                    }
                    searchStart = found + run.Text.Length;
                }
            }
            finally
            {
                ComHelper.Release(font, paragraphRange);
            }
        }

        private static int CountBlocks(List<MarkdownParagraph> blocks, string kind)
        {
            int count = 0;
            foreach (MarkdownParagraph block in blocks)
            {
                if (string.Equals(block.Kind, kind, StringComparison.Ordinal))
                {
                    count++;
                }
            }
            return count;
        }

        private static List<MarkdownParagraph> ParseMarkdown(string markdown, JArray warnings)
        {
            List<MarkdownParagraph> blocks = new List<MarkdownParagraph>();
            string[] lines = markdown.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            StringBuilder paragraphBuffer = new StringBuilder();
            bool inCodeFence = false;
            bool tableWarningAdded = false;

            for (int index = 0; index < lines.Length; index++)
            {
                string line = lines[index];
                string trimmed = line.Trim();
                if (trimmed.StartsWith("```", StringComparison.Ordinal))
                {
                    FlushMarkdownParagraph(blocks, paragraphBuffer);
                    inCodeFence = !inCodeFence;
                    continue;
                }
                if (inCodeFence)
                {
                    if (paragraphBuffer.Length > 0)
                    {
                        paragraphBuffer.Append(' ');
                    }
                    paragraphBuffer.Append(line.TrimEnd());
                    continue;
                }
                if (trimmed.Length == 0)
                {
                    FlushMarkdownParagraph(blocks, paragraphBuffer);
                    continue;
                }

                if (trimmed.StartsWith("|", StringComparison.Ordinal) &&
                    trimmed.IndexOf('|', 1) >= 0)
                {
                    FlushMarkdownParagraph(blocks, paragraphBuffer);
                    List<string> tableLines = new List<string>();
                    int cursor = index;
                    while (cursor < lines.Length &&
                           lines[cursor].Trim().StartsWith("|", StringComparison.Ordinal))
                    {
                        tableLines.Add(lines[cursor]);
                        cursor++;
                    }
                    index = cursor - 1;
                    foreach (string tableLine in tableLines)
                    {
                        List<string> cells = SplitMarkdownTableRow(tableLine);
                        if (cells.Count == 0 || IsMarkdownTableSeparator(cells))
                        {
                            continue;
                        }
                        MarkdownParagraph row = new MarkdownParagraph();
                        row.Kind = "tableRow";
                        List<MarkdownInlineRun> runs = new List<MarkdownInlineRun>();
                        StringBuilder rowText = new StringBuilder();
                        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++)
                        {
                            if (cellIndex > 0)
                            {
                                rowText.Append('\t');
                            }
                            List<MarkdownInlineRun> cellRuns;
                            string cellText = ParseInline(cells[cellIndex], out cellRuns);
                            rowText.Append(cellText);
                            foreach (MarkdownInlineRun run in cellRuns)
                            {
                                runs.Add(run);
                            }
                        }
                        row.PlainText = rowText.ToString();
                        row.Runs = runs;
                        blocks.Add(row);
                    }
                    if (!tableWarningAdded)
                    {
                        tableWarningAdded = true;
                        WordExtensionsUtil.AddWarning(
                            warnings,
                            "Markdown 表格按制表符分隔文本导入，未创建 Word 表格对象。");
                    }
                    continue;
                }

                Match heading = MarkdownHeadingRegex.Match(line);
                if (heading.Success)
                {
                    FlushMarkdownParagraph(blocks, paragraphBuffer);
                    MarkdownParagraph block = new MarkdownParagraph();
                    block.Kind = "heading";
                    block.Level = heading.Groups[1].Value.Length;
                    List<MarkdownInlineRun> runs;
                    block.PlainText = ParseInline(heading.Groups[2].Value.Trim(), out runs);
                    block.Runs = runs;
                    blocks.Add(block);
                    continue;
                }

                Match bullet = MarkdownBulletRegex.Match(line);
                if (bullet.Success)
                {
                    FlushMarkdownParagraph(blocks, paragraphBuffer);
                    MarkdownParagraph block = new MarkdownParagraph();
                    block.Kind = "bullet";
                    block.Level = Math.Max(1, bullet.Groups[1].Value.Length / 2 + 1);
                    List<MarkdownInlineRun> runs;
                    block.PlainText = ParseInline(bullet.Groups[2].Value.Trim(), out runs);
                    block.Runs = runs;
                    blocks.Add(block);
                    continue;
                }

                Match ordered = MarkdownOrderedRegex.Match(line);
                if (ordered.Success)
                {
                    FlushMarkdownParagraph(blocks, paragraphBuffer);
                    MarkdownParagraph block = new MarkdownParagraph();
                    block.Kind = "ordered";
                    block.Level = Math.Max(1, ordered.Groups[1].Value.Length / 2 + 1);
                    List<MarkdownInlineRun> runs;
                    block.PlainText = ParseInline(ordered.Groups[2].Value.Trim(), out runs);
                    block.Runs = runs;
                    blocks.Add(block);
                    continue;
                }

                if (trimmed.StartsWith(">", StringComparison.Ordinal))
                {
                    trimmed = trimmed.TrimStart('>').Trim();
                }
                if (paragraphBuffer.Length > 0)
                {
                    paragraphBuffer.Append(' ');
                }
                paragraphBuffer.Append(trimmed);
            }
            FlushMarkdownParagraph(blocks, paragraphBuffer);
            return blocks;
        }

        private static void FlushMarkdownParagraph(
            List<MarkdownParagraph> blocks,
            StringBuilder buffer)
        {
            if (blocks == null || buffer == null || buffer.Length == 0)
            {
                if (buffer != null)
                {
                    buffer.Length = 0;
                }
                return;
            }
            string rawText = buffer.ToString();
            buffer.Length = 0;
            if (string.IsNullOrWhiteSpace(rawText))
            {
                return;
            }
            MarkdownParagraph block = new MarkdownParagraph();
            block.Kind = "paragraph";
            List<MarkdownInlineRun> runs;
            block.PlainText = ParseInline(rawText.Trim(), out runs);
            block.Runs = runs;
            blocks.Add(block);
        }

        private static List<string> SplitMarkdownTableRow(string line)
        {
            List<string> cells = new List<string>();
            string trimmed = line.Trim();
            if (trimmed.StartsWith("|", StringComparison.Ordinal))
            {
                trimmed = trimmed.Substring(1);
            }
            if (trimmed.EndsWith("|", StringComparison.Ordinal))
            {
                trimmed = trimmed.Substring(0, trimmed.Length - 1);
            }
            string[] parts = trimmed.Split('|');
            foreach (string part in parts)
            {
                cells.Add(part.Trim());
            }
            return cells;
        }

        private static bool IsMarkdownTableSeparator(List<string> cells)
        {
            if (cells == null || cells.Count == 0)
            {
                return false;
            }
            foreach (string cell in cells)
            {
                if (!MarkdownTableSeparatorRegex.IsMatch(cell.Trim()))
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>解析 **粗体**、*斜体*、__粗体__、_斜体_；返回纯文本并输出格式片段。</summary>
        private static string ParseInline(string raw, out List<MarkdownInlineRun> runs)
        {
            runs = new List<MarkdownInlineRun>();
            if (string.IsNullOrEmpty(raw))
            {
                return string.Empty;
            }
            StringBuilder plain = new StringBuilder();
            StringBuilder current = new StringBuilder();
            bool bold = false;
            bool italic = false;
            int index = 0;
            while (index < raw.Length)
            {
                if (index + 1 < raw.Length && raw[index] == '*' && raw[index + 1] == '*')
                {
                    FlushInlineRun(runs, current, bold, italic);
                    bold = !bold;
                    index += 2;
                    continue;
                }
                if (index + 1 < raw.Length && raw[index] == '_' && raw[index + 1] == '_')
                {
                    FlushInlineRun(runs, current, bold, italic);
                    bold = !bold;
                    index += 2;
                    continue;
                }
                if (raw[index] == '*')
                {
                    FlushInlineRun(runs, current, bold, italic);
                    italic = !italic;
                    index++;
                    continue;
                }
                if (raw[index] == '_')
                {
                    bool leftWord = index > 0 && char.IsLetterOrDigit(raw[index - 1]);
                    bool rightWord = index + 1 < raw.Length && char.IsLetterOrDigit(raw[index + 1]);
                    if (!leftWord && !rightWord)
                    {
                        FlushInlineRun(runs, current, bold, italic);
                        italic = !italic;
                        index++;
                        continue;
                    }
                }
                if (raw[index] == '`')
                {
                    index++;
                    continue;
                }
                if (raw[index] == '\\' && index + 1 < raw.Length)
                {
                    current.Append(raw[index + 1]);
                    plain.Append(raw[index + 1]);
                    index += 2;
                    continue;
                }
                current.Append(raw[index]);
                plain.Append(raw[index]);
                index++;
            }
            FlushInlineRun(runs, current, bold, italic);
            return plain.ToString();
        }

        private static void FlushInlineRun(
            List<MarkdownInlineRun> runs,
            StringBuilder buffer,
            bool bold,
            bool italic)
        {
            if (runs == null || buffer == null || buffer.Length == 0)
            {
                if (buffer != null)
                {
                    buffer.Length = 0;
                }
                return;
            }
            MarkdownInlineRun run = new MarkdownInlineRun();
            run.Text = buffer.ToString();
            run.Bold = bold;
            run.Italic = italic;
            runs.Add(run);
            buffer.Length = 0;
        }

        private sealed class MarkdownParagraph
        {
            public string Kind { get; set; }

            public int Level { get; set; }

            public string PlainText { get; set; }

            public List<MarkdownInlineRun> Runs { get; set; }

            public MarkdownParagraph()
            {
                Kind = "paragraph";
                Level = 1;
                PlainText = string.Empty;
                Runs = new List<MarkdownInlineRun>();
            }
        }

        private sealed class MarkdownInlineRun
        {
            public string Text { get; set; }

            public bool Bold { get; set; }

            public bool Italic { get; set; }

            public MarkdownInlineRun()
            {
                Text = string.Empty;
            }
        }

        // ------------------------------------------------------------------
        // word_table_convert
        // ------------------------------------------------------------------

        private OperationResult HandleTableConvert(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string requestedDirection = (args.GetOptionalString("direction", "tableToText") ??
                                         "tableToText").Trim();
            string directionKey = requestedDirection.ToLowerInvariant();
            if (directionKey == "表格转文本")
            {
                directionKey = "tabletotext";
            }
            else if (directionKey == "文本转表格")
            {
                directionKey = "texttotable";
            }
            if (directionKey != "tabletotext" && directionKey != "texttotable")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "direction 只支持 tableToText/textToTable（表格转文本/文本转表格），收到「" +
                    requestedDirection + "」。");
            }
            bool tableToText = directionKey == "tabletotext";
            string direction = tableToText ? "tableToText" : "textToTable";
            string separator = WordExtensionsUtil.ParseSeparator(
                args.GetOptionalString("separator", "tab"));
            bool confirm = args.GetOptionalBool("confirm", false);

            if (tableToText)
            {
                if (!confirm)
                {
                    return OperationResult.Cancelled(
                        "tableToText 会移除表格结构（单元格文本保留）。请显式传 confirm=true 后再执行；" +
                        "本次未修改文档。");
                }
                int tableIndex = args.RequireInt("tableIndex", 1, int.MaxValue);
                return RunWriteTool(
                    application,
                    WordExtensionsToolNames.TableConvert,
                    false,
                    delegate(Word.Document document)
                    {
                        Word.Table probe = null;
                        try
                        {
                            probe = WordComUtil.GetTable(document, tableIndex);
                        }
                        finally
                        {
                            ComHelper.Release(probe);
                        }
                        return WordWriteTargets.Table(document, tableIndex);
                    },
                    delegate(Word.Document document, WordWriteTarget target)
                    {
                        Word.Table table = null;
                        Word.Range convertedRange = null;
                        try
                        {
                            int beforeCount = WordDocumentStateCaptureHelper.CountTables(document);
                            table = WordComUtil.GetTable(document, tableIndex);
                            object separatorObject = separator;
                            object nestedTables = false;
                            convertedRange = table.ConvertToText(
                                ref separatorObject, ref nestedTables);
                            int afterCount = WordDocumentStateCaptureHelper.CountTables(document);
                            string convertedText = convertedRange == null
                                ? string.Empty : (convertedRange.Text ?? string.Empty);
                            JObject data = new JObject();
                            data["direction"] = direction;
                            data["tableIndex"] = tableIndex;
                            data["separator"] = DisplaySeparator(separator);
                            data["tableCountBefore"] = beforeCount;
                            data["tableCountAfter"] = afterCount;
                            data["convertedTextPreview"] =
                                WordExtensionsUtil.Preview(convertedText, 300);
                            data["convertedCharacters"] = convertedText.Length;
                            data["note"] = "Word.Table.ConvertToText 已执行；表格结构移除，单元格文本按分隔符保留。";
                            if (afterCount >= beforeCount)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INTERNAL_ERROR,
                                    "ConvertToText 已返回，但表格数未减少（" + beforeCount + "→" +
                                    afterCount + "）。请确认 tableIndex=" + tableIndex +
                                    " 与 Word 实际表格集合一致。",
                                    data,
                                    false);
                            }
                            return WordWriteResult.Ok(
                                "已把第 " + tableIndex + " 个表格转换为文本（分隔符 " +
                                DisplaySeparator(separator) + "），表格数 " + beforeCount +
                                "→" + afterCount + "。",
                                data,
                                true).AddTargetRef("Word:Table:" + tableIndex);
                        }
                        finally
                        {
                            ComHelper.Release(convertedRange, table);
                        }
                    });
            }

            string rangeMode = (args.GetOptionalString("range", "selection") ?? "selection")
                .Trim().ToLowerInvariant();
            int paragraphIndex = args.GetOptionalInt("paragraphIndex", 0);
            return RunWriteTool(
                application,
                WordExtensionsToolNames.TableConvert,
                false,
                delegate(Word.Document document)
                {
                    if (rangeMode == "all" || rangeMode == "document" || rangeMode == "全文")
                    {
                        return WordWriteTargets.Document(
                            document, "全文", "Word:Document", false, true);
                    }
                    if (rangeMode == "paragraph")
                    {
                        if (paragraphIndex <= 0)
                        {
                            throw WordArguments.Invalid(
                                "range=paragraph 时必须提供 paragraphIndex（1 基），收到 " +
                                paragraphIndex + "。当前文档共 " +
                                WordComUtil.GetParagraphCount(document) + " 段。");
                        }
                        WordComUtil.ValidateParagraphIndex(document, paragraphIndex);
                        return WordWriteTargets.Paragraph(
                            document, paragraphIndex, "Word:Paragraph:" + paragraphIndex, true);
                    }
                    if (rangeMode != "selection" && rangeMode != "选区")
                    {
                        throw WordArguments.Invalid(
                            "range 只支持 selection/paragraph/all，收到「" + rangeMode + "」。");
                    }
                    return WordWriteTargets.Selection(
                        application, "当前选区", "Word:Selection", false, true);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Range range = null;
                    Word.Table table = null;
                    try
                    {
                        int beforeCount = WordDocumentStateCaptureHelper.CountTables(document);
                        if (rangeMode == "all" || rangeMode == "document" || rangeMode == "全文")
                        {
                            range = document.Content;
                        }
                        else if (rangeMode == "paragraph")
                        {
                            Word.Paragraph paragraph = null;
                            try
                            {
                                paragraph = WordComUtil.GetParagraph(document, paragraphIndex);
                                range = paragraph.Range;
                            }
                            finally
                            {
                                ComHelper.Release(paragraph);
                            }
                        }
                        else
                        {
                            Word.Selection selection = null;
                            try
                            {
                                selection = application.Selection;
                                if (selection == null)
                                {
                                    return WordWriteResult.Fail(
                                        OfficeErrorCode.INVALID_PARAMS,
                                        "当前没有有效选区，无法执行 textToTable；可改用 range=paragraph/all。",
                                        null,
                                        false);
                                }
                                range = selection.Range;
                            }
                            finally
                            {
                                ComHelper.Release(selection);
                            }
                        }
                        if (range == null || string.IsNullOrWhiteSpace(range.Text))
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "转换范围为空，无法执行 textToTable。请先选中含分隔符的文本。",
                                null,
                                false);
                        }

                        string sourceText = range.Text ?? string.Empty;
                        object separatorObject = separator;
                        object missing = Type.Missing;
                        table = range.ConvertToTable(
                            ref separatorObject,
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing);
                        int afterCount = WordDocumentStateCaptureHelper.CountTables(document);
                        JObject tableSummary = WordComUtil.TableSummary(table);
                        JObject data = new JObject();
                        data["direction"] = direction;
                        data["range"] = rangeMode;
                        data["paragraphIndex"] = paragraphIndex;
                        data["separator"] = DisplaySeparator(separator);
                        data["tableCountBefore"] = beforeCount;
                        data["tableCountAfter"] = afterCount;
                        data["table"] = tableSummary;
                        data["sourceTextPreview"] = WordExtensionsUtil.Preview(sourceText, 300);
                        data["note"] = "Word.Range.ConvertToTable 已执行；分隔符按 separator 解析。";

                        if (afterCount <= beforeCount)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "ConvertToTable 已返回，但表格数未增加（" + beforeCount + "→" +
                                afterCount + "）。请确认选区文本包含分隔符「" +
                                DisplaySeparator(separator) + "」。",
                                data,
                                false);
                        }
                        return WordWriteResult.Ok(
                            "已把文本转换为表格（分隔符 " + DisplaySeparator(separator) +
                            "，表格数 " + beforeCount + "→" + afterCount + "）。",
                            data,
                            true).AddTargetRef("Word:Document");
                    }
                    finally
                    {
                        ComHelper.Release(table, range);
                    }
                });
        }

        private static string DisplaySeparator(string separator)
        {
            if (separator == "\t")
            {
                return "tab(\\t)";
            }
            if (separator == "\r")
            {
                return "paragraph(\\r)";
            }
            if (separator == " ")
            {
                return "space";
            }
            return separator;
        }

        // ------------------------------------------------------------------
        // word_section_manage
        // ------------------------------------------------------------------

        private OperationResult HandleSectionManage(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string action = (args.GetOptionalString("action", "list") ?? "list")
                .Trim().ToLowerInvariant();
            if (action != "list" && action != "insert" && action != "format")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "action 只支持 list/insert/format，收到「" + action + "」。");
            }

            if (action == "list")
            {
                return RunReadTool(
                    application,
                    delegate(Word.Document document)
                    {
                        JArray items = new JArray();
                        Word.Sections sections = null;
                        try
                        {
                            sections = document.Sections;
                            int count = sections == null ? 0 : sections.Count;
                            for (int index = 1; index <= count; index++)
                            {
                                Word.Section section = null;
                                Word.Range range = null;
                                Word.PageSetup pageSetup = null;
                                Word.HeadersFooters headers = null;
                                Word.HeadersFooters footers = null;
                                try
                                {
                                    section = sections[index];
                                    range = section.Range;
                                    pageSetup = section.PageSetup;
                                    headers = section.Headers;
                                    footers = section.Footers;
                                    bool headerExists = false;
                                    bool footerExists = false;
                                    Word.HeaderFooter header = null;
                                    Word.HeaderFooter footer = null;
                                    try
                                    {
                                        header = headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                                        footer = footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                                        headerExists = header != null && header.Exists;
                                        footerExists = footer != null && footer.Exists;
                                    }
                                    finally
                                    {
                                        ComHelper.Release(footer, header);
                                    }
                                    items.Add(new JObject
                                    {
                                        { "sectionIndex", index },
                                        { "start", range == null ? 0 : range.Start },
                                        { "end", range == null ? 0 : range.End },
                                        { "textPreview", WordExtensionsUtil.Preview(
                                            range == null ? string.Empty : range.Text, 120) },
                                        { "pageSetup", WordComUtil.ReadPageSetup(pageSetup) },
                                        { "sectionDirection", pageSetup == null
                                            ? string.Empty : pageSetup.SectionDirection.ToString() },
                                        { "primaryHeaderExists", headerExists },
                                        { "primaryFooterExists", footerExists }
                                    });
                                }
                                finally
                                {
                                    ComHelper.Release(footers, headers, pageSetup, range, section);
                                }
                            }
                            JObject data = new JObject();
                            data["sectionCount"] = count;
                            data["sections"] = items;
                            return ReadSucceed(
                                "已读取 " + count + " 个节的信息。", data, "Word:Document");
                        }
                        finally
                        {
                            ComHelper.Release(sections);
                        }
                    });
            }

            if (action == "insert")
            {
                string position = (args.GetOptionalString("position", "cursor") ?? "cursor")
                    .Trim().ToLowerInvariant();
                int paragraphIndex = args.GetOptionalInt("paragraphIndex", 0);
                if (position != "start" && position != "end" &&
                    position != "cursor" && position != "paragraph")
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.INVALID_PARAMS,
                        "action=insert 时 position 只支持 start/end/cursor/paragraph，收到「" +
                        position + "」。");
                }
                if (position == "paragraph" && paragraphIndex <= 0)
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.INVALID_PARAMS,
                        "position=paragraph 时必须提供 paragraphIndex（1 基），收到 " +
                        paragraphIndex + "。");
                }
                return RunWriteTool(
                    application,
                    WordExtensionsToolNames.SectionManage,
                    false,
                    delegate(Word.Document document)
                    {
                        if (position == "start")
                        {
                            return WordWriteTargets.DocumentStart();
                        }
                        if (position == "end")
                        {
                            return WordWriteTargets.DocumentEnd(document);
                        }
                        if (position == "paragraph")
                        {
                            WordComUtil.ValidateParagraphIndex(document, paragraphIndex);
                            return WordWriteTargets.Paragraph(
                                document, paragraphIndex, "Word:Paragraph:" + paragraphIndex, true);
                        }
                        return WordWriteTargets.Cursor(application);
                    },
                    delegate(Word.Document document, WordWriteTarget target)
                    {
                        Word.Range range = null;
                        Word.Sections sections = null;
                        try
                        {
                            int beforeCount = WordExtensionsUtilCountSections(document);
                            range = WordComUtil.CreateRange(document, target.Start, target.Start);
                            object breakType = Word.WdBreakType.wdSectionBreakNextPage;
                            range.InsertBreak(ref breakType);
                            int afterCount = WordExtensionsUtilCountSections(document);
                            JObject data = new JObject();
                            data["action"] = "insert";
                            data["position"] = position;
                            data["paragraphIndex"] = paragraphIndex;
                            data["breakType"] = "wdSectionBreakNextPage";
                            data["sectionCountBefore"] = beforeCount;
                            data["sectionCountAfter"] = afterCount;
                            data["note"] = "已插入“下一页”分节符；页码/页眉页脚继承关系需用 word_section_manage list 复核。";
                            if (afterCount <= beforeCount)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INTERNAL_ERROR,
                                    "InsertBreak 已返回，但节数未增加（" + beforeCount + "→" +
                                    afterCount + "）。请确认插入位置有效且文档未受保护。",
                                    data,
                                    false);
                            }
                            return WordWriteResult.Ok(
                                "已在 " + target.Label + " 插入下一页分节符（节数 " +
                                beforeCount + "→" + afterCount + "）。",
                                data,
                                true).AddTargetRef("Word:Section:" + afterCount);
                        }
                        finally
                        {
                            ComHelper.Release(sections, range);
                        }
                    });
            }

            // action == format
            int sectionIndex = args.GetOptionalInt("sectionIndex", 0);
            double? topMarginPt = ResolveMarginPt(args, "topMarginPt", "topMarginCm");
            double? bottomMarginPt = ResolveMarginPt(args, "bottomMarginPt", "bottomMarginCm");
            double? leftMarginPt = ResolveMarginPt(args, "leftMarginPt", "leftMarginCm");
            double? rightMarginPt = ResolveMarginPt(args, "rightMarginPt", "rightMarginCm");
            Word.WdPaperSize? paperSize = null;
            if (args.Has("paperSize"))
            {
                paperSize = WordComUtil.ParsePaperSize(args.GetOptionalString("paperSize", null));
            }
            Word.WdOrientation? orientation = null;
            if (args.Has("orientation"))
            {
                string orientationText = (args.GetOptionalString("orientation", null) ?? string.Empty)
                    .Trim().ToLowerInvariant();
                if (orientationText == "portrait" || orientationText == "纵向")
                {
                    orientation = Word.WdOrientation.wdOrientPortrait;
                }
                else if (orientationText == "landscape" || orientationText == "横向")
                {
                    orientation = Word.WdOrientation.wdOrientLandscape;
                }
                else
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.INVALID_PARAMS,
                        "orientation 只支持 portrait/landscape（纵向/横向），收到「" +
                        orientationText + "」。");
                }
            }
            Word.WdVerticalAlignment? verticalAlignment = null;
            if (args.Has("verticalAlignment"))
            {
                string alignmentText = (args.GetOptionalString("verticalAlignment", null) ??
                                        string.Empty).Trim().ToLowerInvariant();
                if (alignmentText == "top" || alignmentText == "顶端")
                {
                    verticalAlignment = Word.WdVerticalAlignment.wdAlignVerticalTop;
                }
                else if (alignmentText == "center" || alignmentText == "居中")
                {
                    verticalAlignment = Word.WdVerticalAlignment.wdAlignVerticalCenter;
                }
                else if (alignmentText == "bottom" || alignmentText == "底端")
                {
                    verticalAlignment = Word.WdVerticalAlignment.wdAlignVerticalBottom;
                }
                else if (alignmentText == "justify" || alignmentText == "两端")
                {
                    verticalAlignment = Word.WdVerticalAlignment.wdAlignVerticalJustify;
                }
                else
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.INVALID_PARAMS,
                        "verticalAlignment 只支持 top/center/bottom/justify，收到「" +
                        alignmentText + "」。");
                }
            }
            Word.WdSectionDirection? sectionDirection = null;
            if (args.Has("sectionDirection"))
            {
                string directionText = (args.GetOptionalString("sectionDirection", null) ??
                                        string.Empty).Trim().ToLowerInvariant();
                if (directionText == "ltr" || directionText == "从左到右")
                {
                    sectionDirection = Word.WdSectionDirection.wdSectionDirectionLtr;
                }
                else if (directionText == "rtl" || directionText == "从右到左")
                {
                    sectionDirection = Word.WdSectionDirection.wdSectionDirectionRtl;
                }
                else
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.INVALID_PARAMS,
                        "sectionDirection 只支持 ltr/rtl，收到「" + directionText +
                        "」。注意该属性控制左右书写方向，不是竖排文字。");
                }
            }
            if (!topMarginPt.HasValue && !bottomMarginPt.HasValue &&
                !leftMarginPt.HasValue && !rightMarginPt.HasValue &&
                !paperSize.HasValue && !orientation.HasValue &&
                !verticalAlignment.HasValue && !sectionDirection.HasValue)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "action=format 至少需要一个设置：topMarginPt/topMarginCm、" +
                    "bottomMarginPt/bottomMarginCm、leftMarginPt/leftMarginCm、" +
                    "rightMarginPt/rightMarginCm、paperSize、orientation、verticalAlignment、sectionDirection。");
            }

            return RunWriteTool(
                application,
                WordExtensionsToolNames.SectionManage,
                false,
                delegate(Word.Document document)
                {
                    if (sectionIndex > 0)
                    {
                        WordComUtil.ValidateSectionIndex(document, sectionIndex);
                    }
                    return WordWriteTargets.None("节页面设置", "Word:Sections");
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
                        if (sectionIndex > count)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "sectionIndex " + sectionIndex + " 不存在（当前共 " + count +
                                " 节，1 基）。",
                                null,
                                false);
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
                        indexes.Sort();
                        indexes.Reverse();

                        foreach (int index in indexes)
                        {
                            Word.Section section = null;
                            Word.PageSetup pageSetup = null;
                            try
                            {
                                section = sections[index];
                                pageSetup = section.PageSetup;
                                JObject before = WordComUtil.ReadPageSetup(pageSetup);
                                before["verticalAlignment"] = pageSetup.VerticalAlignment.ToString();
                                before["sectionDirection"] = pageSetup.SectionDirection.ToString();
                                if (topMarginPt.HasValue)
                                {
                                    pageSetup.TopMargin = (float)topMarginPt.Value;
                                }
                                if (bottomMarginPt.HasValue)
                                {
                                    pageSetup.BottomMargin = (float)bottomMarginPt.Value;
                                }
                                if (leftMarginPt.HasValue)
                                {
                                    pageSetup.LeftMargin = (float)leftMarginPt.Value;
                                }
                                if (rightMarginPt.HasValue)
                                {
                                    pageSetup.RightMargin = (float)rightMarginPt.Value;
                                }
                                if (paperSize.HasValue)
                                {
                                    pageSetup.PaperSize = paperSize.Value;
                                }
                                if (orientation.HasValue)
                                {
                                    pageSetup.Orientation = orientation.Value;
                                }
                                if (verticalAlignment.HasValue)
                                {
                                    pageSetup.VerticalAlignment = verticalAlignment.Value;
                                }
                                if (sectionDirection.HasValue)
                                {
                                    pageSetup.SectionDirection = sectionDirection.Value;
                                }
                                JObject after = WordComUtil.ReadPageSetup(pageSetup);
                                after["verticalAlignment"] = pageSetup.VerticalAlignment.ToString();
                                after["sectionDirection"] = pageSetup.SectionDirection.ToString();
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
                                    { "reason", OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) }
                                });
                            }
                            finally
                            {
                                ComHelper.Release(pageSetup, section);
                            }
                        }

                        JObject data = new JObject();
                        data["action"] = "format";
                        data["sectionIndex"] = sectionIndex;
                        data["sectionCount"] = count;
                        data["changedSections"] = changedCount;
                        data["sections"] = items;
                        data["failures"] = failures;
                        data["note"] = "页面设置/方向变更会改变文档结构，T02 快照引擎不覆盖，未登记可回放撤销节点；" +
                                        "Word 自定义撤销记录已尝试开启。";

                        if (failures.Count > 0 && changedCount == 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INTERNAL_ERROR,
                                "节格式设置失败：首个原因 " +
                                failures[0]["reason"].ToString() +
                                "。请确认文档未受保护且 sectionIndex 正确。",
                                data,
                                false);
                        }
                        if (failures.Count > 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.PARTIAL_APPLY,
                                "节格式部分完成：成功修改 " + changedCount + " 节，失败 " +
                                failures.Count + " 节；首个原因：" +
                                failures[0]["reason"].ToString() + "。不会自动回滚。",
                                data,
                                changedCount > 0);
                        }
                        return WordWriteResult.Ok(
                            "已完成 " + changedCount + "/" + count + " 节的页面设置" +
                            (changedCount == 0 ? "（设置与当前值相同，未检测到变化）" : "") + "。",
                            data,
                            changedCount > 0).AddTargetRef("Word:Sections");
                    }
                    finally
                    {
                        ComHelper.Release(sections);
                    }
                });
        }

        private static double? ResolveMarginPt(
            WordArguments args,
            string pointsName,
            string centimetersName)
        {
            double? points = args.GetOptionalDouble(pointsName);
            if (points.HasValue)
            {
                if (points.Value < 0)
                {
                    throw WordArguments.Invalid(pointsName + " 不能为负数，收到 " + points.Value + "。");
                }
                return points.Value;
            }
            double? centimeters = args.GetOptionalDouble(centimetersName);
            if (centimeters.HasValue)
            {
                if (centimeters.Value < 0)
                {
                    throw WordArguments.Invalid(
                        centimetersName + " 不能为负数，收到 " + centimeters.Value + "。");
                }
                return centimeters.Value * WordFormatUtil.PointsPerCentimeter;
            }
            return null;
        }

        private static int WordExtensionsUtilCountSections(Word.Document document)
        {
            Word.Sections sections = null;
            try
            {
                sections = document.Sections;
                return sections == null ? 0 : sections.Count;
            }
            finally
            {
                ComHelper.Release(sections);
            }
        }

        /// <summary>局部计数助手：避免把表格计数逻辑散落到多个处理器。</summary>
        private static class WordDocumentStateCaptureHelper
        {
            public static int CountTables(Word.Document document)
            {
                Word.Tables tables = null;
                try
                {
                    tables = document.Tables;
                    return tables == null ? 0 : tables.Count;
                }
                finally
                {
                    ComHelper.Release(tables);
                }
            }
        }
    }
}
