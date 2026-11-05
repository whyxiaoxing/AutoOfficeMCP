using System;
using System.Collections.Generic;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    /// <summary>
    /// Word COM 通用操作：范围创建、段落定位、文本清洗、样式/表格/图片索引与轻量快照。
    /// 所有返回 COM 对象的方法都要求调用方 finally Release；本类内部对象一律自行释放。
    /// </summary>
    internal static class WordComUtil
    {
        private static readonly Regex WhitespaceRegex = new Regex(@"\s+", RegexOptions.Compiled);

        public static Word.Range CreateRange(Word.Document document, int start, int end)
        {
            if (document == null)
            {
                throw WordArguments.Invalid("当前没有活动文档，无法创建 Word 范围。");
            }
            object startObject = start;
            object endObject = end;
            return document.Range(ref startObject, ref endObject);
        }

        public static int GetContentEnd(Word.Document document)
        {
            if (document == null)
            {
                return 0;
            }
            Word.Range content = null;
            try
            {
                content = document.Content;
                return content.End;
            }
            finally
            {
                ComHelper.Release(content);
            }
        }

        public static int GetParagraphCount(Word.Document document)
        {
            if (document == null)
            {
                return 0;
            }
            Word.Paragraphs paragraphs = null;
            try
            {
                paragraphs = document.Paragraphs;
                return paragraphs == null ? 0 : paragraphs.Count;
            }
            finally
            {
                ComHelper.Release(paragraphs);
            }
        }

        public static int ValidateParagraphIndex(Word.Document document, int paragraphIndex)
        {
            if (paragraphIndex < 1)
            {
                throw WordArguments.Invalid(
                    "paragraphIndex 必须从 1 开始（1 基，指向文档当前真实段落序号）；收到 " +
                    paragraphIndex.ToString(CultureInfo.InvariantCulture) + "。");
            }

            int count = GetParagraphCount(document);
            if (paragraphIndex > count)
            {
                throw WordArguments.Invalid(
                    "第 " + paragraphIndex.ToString(CultureInfo.InvariantCulture) +
                    " 段不存在（当前共 " + count.ToString(CultureInfo.InvariantCulture) +
                    " 段）。请先用 word_list_paragraphs 或 word_get_context 读取真实段号。");
            }
            return count;
        }

        public static Word.Paragraph GetParagraph(Word.Document document, int paragraphIndex)
        {
            ValidateParagraphIndex(document, paragraphIndex);
            Word.Paragraphs paragraphs = null;
            try
            {
                paragraphs = document.Paragraphs;
                return paragraphs[paragraphIndex];
            }
            finally
            {
                ComHelper.Release(paragraphs);
            }
        }

        public static string GetParagraphText(Word.Paragraph paragraph)
        {
            if (paragraph == null)
            {
                return string.Empty;
            }
            Word.Range range = null;
            try
            {
                range = paragraph.Range;
                return CleanParagraphText(range.Text);
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                ComHelper.Release(range);
            }
        }

        public static int FindParagraphIndexAt(Word.Document document, int position)
        {
            if (document == null)
            {
                return 0;
            }
            Word.Paragraphs paragraphs = null;
            try
            {
                paragraphs = document.Paragraphs;
                int count = paragraphs.Count;
                if (count <= 0)
                {
                    return 0;
                }

                // 段落的 Start 单调递增；二分查找“最后一个 Start <= position 的段落”。
                // 500 段文档从逐段 COM 往返（约 20 秒）降到约 9 次，语义与原线性扫描一致：
                // 边界点归属下一段，文档末尾归属最后一段。
                int low = 1;
                int high = count;
                int candidate = 1;
                while (low <= high)
                {
                    int middle = low + (high - low) / 2;
                    Word.Paragraph paragraph = null;
                    Word.Range range = null;
                    try
                    {
                        paragraph = paragraphs[middle];
                        range = paragraph.Range;
                        if (range.Start <= position)
                        {
                            candidate = middle;
                            low = middle + 1;
                        }
                        else
                        {
                            high = middle - 1;
                        }
                    }
                    finally
                    {
                        ComHelper.Release(range, paragraph);
                    }
                }
                return candidate;
            }
            finally
            {
                ComHelper.Release(paragraphs);
            }
        }

        public static string ReadStyleName(Word.Range range)
        {
            if (range == null)
            {
                return string.Empty;
            }
            object styleObject = null;
            Word.Style style = null;
            try
            {
                styleObject = range.get_Style();
                if (styleObject == null)
                {
                    return string.Empty;
                }
                style = styleObject as Word.Style;
                if (style != null)
                {
                    return style.NameLocal ?? string.Empty;
                }
                return styleObject.ToString();
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                ComHelper.Release(styleObject, style);
            }
        }

        public static string ReadTableStyleName(Word.Table table)
        {
            if (table == null)
            {
                return string.Empty;
            }
            object styleObject = null;
            Word.Style style = null;
            try
            {
                styleObject = table.get_Style();
                if (styleObject == null)
                {
                    return string.Empty;
                }
                style = styleObject as Word.Style;
                return style != null ? (style.NameLocal ?? string.Empty) : styleObject.ToString();
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                ComHelper.Release(styleObject, style);
            }
        }

        public static bool StyleExists(Word.Document document, string styleName)
        {
            if (document == null || string.IsNullOrWhiteSpace(styleName))
            {
                return false;
            }
            Word.Styles styles = null;
            object styleObject = null;
            try
            {
                styles = document.Styles;
                styleObject = styles[styleName];
                return styleObject != null;
            }
            catch
            {
                return false;
            }
            finally
            {
                ComHelper.Release(styleObject, styles);
            }
        }

        public static string CleanParagraphText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }
            string cleaned = text.Replace("\r\a", string.Empty)
                                 .Replace("\a", string.Empty)
                                 .Replace("\r", "\n")
                                 .Replace('\v', '\n')
                                 .Replace('\f', '\n');
            return cleaned.TrimEnd('\n');
        }

        public static string CleanCellText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }
            string cleaned = text.Replace("\r\a", string.Empty)
                                 .Replace("\a", string.Empty)
                                 .Replace("\r", "\n")
                                 .Replace('\v', '\n')
                                 .Replace('\f', '\n');
            return cleaned.TrimEnd('\n');
        }

        public static string NormalizePreview(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }
            string normalized = WhitespaceRegex.Replace(text, " ").Trim();
            return Truncate(normalized, maxLength);
        }

        public static string Truncate(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }
            if (maxLength <= 0 || text.Length <= maxLength)
            {
                return text;
            }
            return text.Substring(0, maxLength) + "...";
        }

        public static string HashText(string text)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(text ?? string.Empty);
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(bytes);
                StringBuilder builder = new StringBuilder(hash.Length * 2);
                foreach (byte item in hash)
                {
                    builder.Append(item.ToString("x2", CultureInfo.InvariantCulture));
                }
                return builder.ToString();
            }
        }

        public static int CountOccurrences(string text, string find, bool matchCase, bool matchWholeWord)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(find))
            {
                return 0;
            }

            StringComparison comparison = matchCase
                ? StringComparison.Ordinal
                : StringComparison.OrdinalIgnoreCase;
            int count = 0;
            int position = 0;
            while (position <= text.Length - find.Length)
            {
                int found = text.IndexOf(find, position, comparison);
                if (found < 0)
                {
                    break;
                }
                if (!matchWholeWord || IsWholeWordMatch(text, found, find.Length))
                {
                    count++;
                }
                position = found + Math.Max(1, find.Length);
            }
            return count;
        }

        public static JObject GetSelectionSnapshot(Word.Application application, Word.Document document)
        {
            JObject snapshot = new JObject();
            Word.Selection selection = null;
            Word.Range range = null;
            try
            {
                selection = application.Selection;
                if (selection == null)
                {
                    snapshot["exists"] = false;
                    snapshot["error"] = "当前没有有效选区。";
                    return snapshot;
                }
                range = selection.Range;
                string text = range.Text ?? string.Empty;
                snapshot["exists"] = true;
                snapshot["type"] = selection.Type.ToString();
                snapshot["start"] = range.Start;
                snapshot["end"] = range.End;
                snapshot["textLength"] = text.Length;
                snapshot["textHash"] = HashText(text);
                snapshot["preview"] = NormalizePreview(text, 120);
                snapshot["paragraphIndex"] = FindParagraphIndexAt(document, range.Start);
                snapshot["format"] = WordFormatUtil.CaptureRangeFormat(range);
                return snapshot;
            }
            catch (Exception ex)
            {
                snapshot["exists"] = false;
                snapshot["error"] = OfficeErrorMapper.DescribeException(ex, OfficeApp.Word);
                return snapshot;
            }
            finally
            {
                ComHelper.Release(range, selection);
            }
        }

        public static JObject TableSummary(Word.Table table)
        {
            JObject summary = new JObject();
            if (table == null)
            {
                return summary;
            }

            Word.Rows rows = null;
            Word.Columns columns = null;
            try
            {
                rows = table.Rows;
                columns = table.Columns;
                summary["rows"] = rows == null ? 0 : rows.Count;
                summary["columns"] = columns == null ? 0 : columns.Count;
                summary["style"] = ReadTableStyleName(table);
                summary["uniform"] = table.Uniform;
                summary["nestingLevel"] = table.NestingLevel;
                return summary;
            }
            catch (Exception ex)
            {
                summary["error"] = OfficeErrorMapper.DescribeException(ex, OfficeApp.Word);
                return summary;
            }
            finally
            {
                ComHelper.Release(rows, columns);
            }
        }

        public static void ValidateSectionIndex(Word.Document document, int sectionIndex)
        {
            if (sectionIndex < 1)
            {
                throw WordArguments.Invalid("sectionIndex 必须从 1 开始，收到 " + sectionIndex + "。");
            }
            Word.Sections sections = null;
            try
            {
                sections = document.Sections;
                int count = sections == null ? 0 : sections.Count;
                if (sectionIndex > count)
                {
                    throw WordArguments.Invalid(
                        "节索引 " + sectionIndex + " 不存在（当前共 " + count + " 节，1 基）。");
                }
            }
            finally
            {
                ComHelper.Release(sections);
            }
        }

        public static Word.Table GetTable(Word.Document document, int tableIndex)
        {
            if (document == null)
            {
                throw WordArguments.Invalid("当前没有活动文档，无法读取表格。");
            }
            Word.Tables tables = null;
            try
            {
                tables = document.Tables;
                int count = tables == null ? 0 : tables.Count;
                if (tableIndex < 1 || tableIndex > count)
                {
                    throw WordArguments.Invalid(
                        "表格索引 " + tableIndex.ToString(CultureInfo.InvariantCulture) +
                        " 不存在（当前共 " + count.ToString(CultureInfo.InvariantCulture) +
                        " 个表格，1 基）。注意页眉页脚与文本框内的表格不在 doc.Tables 集合。");
                }
                return tables[tableIndex];
            }
            finally
            {
                ComHelper.Release(tables);
            }
        }

        public static void ValidateCellIndex(Word.Table table, int row, int column)
        {
            if (table == null)
            {
                throw WordArguments.Invalid("表格对象为空，无法定位单元格。");
            }
            Word.Rows rows = null;
            Word.Columns columns = null;
            try
            {
                rows = table.Rows;
                columns = table.Columns;
                int rowCount = rows == null ? 0 : rows.Count;
                int columnCount = columns == null ? 0 : columns.Count;
                if (row < 1 || row > rowCount)
                {
                    throw WordArguments.Invalid(
                        "行号 " + row.ToString(CultureInfo.InvariantCulture) +
                        " 越界（当前表格共 " + rowCount.ToString(CultureInfo.InvariantCulture) + " 行，1 基）。");
                }
                if (column < 1 || column > columnCount)
                {
                    throw WordArguments.Invalid(
                        "列号 " + column.ToString(CultureInfo.InvariantCulture) +
                        " 越界（当前表格共 " + columnCount.ToString(CultureInfo.InvariantCulture) + " 列，1 基）。");
                }
            }
            finally
            {
                ComHelper.Release(rows, columns);
            }
        }

        public static Word.Cell GetCell(Word.Table table, int row, int column)
        {
            ValidateCellIndex(table, row, column);
            try
            {
                return table.Cell(row, column);
            }
            catch (Exception ex)
            {
                throw WordArguments.Invalid(
                    "无法读取表格第 " + row.ToString(CultureInfo.InvariantCulture) + " 行第 " +
                    column.ToString(CultureInfo.InvariantCulture) + " 列单元格（表格可能含合并单元格）：" +
                    OfficeErrorMapper.DescribeException(ex, OfficeApp.Word));
            }
        }

        public static int FindTableIndex(Word.Document document, Word.Table table)
        {
            if (document == null || table == null)
            {
                return 0;
            }
            Word.Tables tables = null;
            Word.Range targetRange = null;
            try
            {
                targetRange = table.Range;
                int targetStart = targetRange.Start;
                int targetEnd = targetRange.End;
                tables = document.Tables;
                int count = tables == null ? 0 : tables.Count;
                for (int index = 1; index <= count; index++)
                {
                    Word.Table candidate = null;
                    Word.Range candidateRange = null;
                    try
                    {
                        candidate = tables[index];
                        candidateRange = candidate.Range;
                        if (candidateRange.Start == targetStart && candidateRange.End == targetEnd)
                        {
                            return index;
                        }
                    }
                    finally
                    {
                        ComHelper.Release(candidateRange, candidate);
                    }
                }
                return count;
            }
            finally
            {
                ComHelper.Release(targetRange, tables);
            }
        }

        public static int FindInlineShapeIndex(Word.Document document, Word.InlineShape shape)
        {
            if (document == null || shape == null)
            {
                return 0;
            }
            Word.InlineShapes shapes = null;
            Word.Range targetRange = null;
            try
            {
                targetRange = shape.Range;
                int targetStart = targetRange.Start;
                int targetEnd = targetRange.End;
                shapes = document.InlineShapes;
                int count = shapes == null ? 0 : shapes.Count;
                for (int index = 1; index <= count; index++)
                {
                    Word.InlineShape candidate = null;
                    Word.Range candidateRange = null;
                    try
                    {
                        candidate = shapes[index];
                        candidateRange = candidate.Range;
                        if (candidateRange.Start == targetStart && candidateRange.End == targetEnd)
                        {
                            return index;
                        }
                    }
                    finally
                    {
                        ComHelper.Release(candidateRange, candidate);
                    }
                }
                return count;
            }
            finally
            {
                ComHelper.Release(targetRange, shapes);
            }
        }

        public static int FindShapeIndex(Word.Document document, Word.Shape shape)
        {
            if (document == null || shape == null)
            {
                return 0;
            }
            Word.Shapes shapes = null;
            try
            {
                int targetId = shape.ID;
                shapes = document.Shapes;
                int count = shapes == null ? 0 : shapes.Count;
                for (int index = 1; index <= count; index++)
                {
                    Word.Shape candidate = null;
                    try
                    {
                        candidate = shapes[index];
                        if (candidate.ID == targetId)
                        {
                            return index;
                        }
                    }
                    finally
                    {
                        ComHelper.Release(candidate);
                    }
                }
                return count;
            }
            finally
            {
                ComHelper.Release(shapes);
            }
        }

        public static bool IsPlainTextRange(Word.Range range)
        {
            if (range == null)
            {
                return false;
            }

            Word.Tables tables = null;
            Word.InlineShapes inlineShapes = null;
            Word.ShapeRange shapeRange = null;
            Word.Fields fields = null;
            Word.OMaths equations = null;
            try
            {
                tables = range.Tables;
                inlineShapes = range.InlineShapes;
                shapeRange = range.ShapeRange;
                fields = range.Fields;
                equations = range.OMaths;
                return (tables == null || tables.Count == 0) &&
                       (inlineShapes == null || inlineShapes.Count == 0) &&
                       (shapeRange == null || shapeRange.Count == 0) &&
                       (fields == null || fields.Count == 0) &&
                       (equations == null || equations.Count == 0);
            }
            catch
            {
                return false;
            }
            finally
            {
                ComHelper.Release(equations, fields, shapeRange, inlineShapes, tables);
            }
        }

        public static void CheckWritable(Word.Document document, string toolName)
        {
            if (document == null)
            {
                throw new OfficeComException(OfficeErrorCode.DOC_NOT_OPEN, "当前没有活动 Word 文档。");
            }
            if (document.ReadOnly)
            {
                throw new OfficeComException(
                    OfficeErrorCode.READ_ONLY,
                    "文档「" + document.Name + "」处于只读状态，工具「" + toolName +
                    "」无法写入。请另存为可写副本，或检查文件属性/是否被其他程序占用。");
            }
            if (document.ProtectionType != Word.WdProtectionType.wdNoProtection)
            {
                throw new OfficeComException(
                    OfficeErrorCode.DOC_PROTECTED,
                    "文档「" + document.Name + "」受保护（ProtectionType=" + document.ProtectionType +
                    "），工具「" + toolName + "」无法写入。请先解除限制编辑/保护后重试。");
            }
        }

        public static int ComputeStatistics(Word.Document document, Word.WdStatistic statistic)
        {
            object missing = Type.Missing;
            return document.ComputeStatistics(statistic, ref missing);
        }

        public static JObject ReadPageSetup(Word.PageSetup pageSetup)
        {
            JObject json = new JObject();
            if (pageSetup == null)
            {
                return json;
            }
            json["topMarginPt"] = (double)pageSetup.TopMargin;
            json["bottomMarginPt"] = (double)pageSetup.BottomMargin;
            json["leftMarginPt"] = (double)pageSetup.LeftMargin;
            json["rightMarginPt"] = (double)pageSetup.RightMargin;
            json["paperSize"] = PaperSizeName(pageSetup.PaperSize);
            json["paperSizeCode"] = (int)pageSetup.PaperSize;
            json["orientation"] = pageSetup.Orientation.ToString();
            json["pageWidthPt"] = (double)pageSetup.PageWidth;
            json["pageHeightPt"] = (double)pageSetup.PageHeight;
            return json;
        }

        public static string PaperSizeName(Word.WdPaperSize paperSize)
        {
            switch (paperSize)
            {
                case Word.WdPaperSize.wdPaperA4:
                    return "A4";
                case Word.WdPaperSize.wdPaperA3:
                    return "A3";
                case Word.WdPaperSize.wdPaperA5:
                    return "A5";
                case Word.WdPaperSize.wdPaperLetter:
                    return "Letter";
                case Word.WdPaperSize.wdPaperLegal:
                    return "Legal";
                case Word.WdPaperSize.wdPaperB4:
                    return "B4";
                case Word.WdPaperSize.wdPaperB5:
                    return "B5";
                case Word.WdPaperSize.wdPaperExecutive:
                    return "Executive";
                default:
                    return paperSize.ToString();
            }
        }

        public static Word.WdPaperSize ParsePaperSize(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                throw WordArguments.Invalid("paperSize 不能为空；支持 A4/A3/A5/Letter/Legal/B4/B5/Executive");
            }
            switch (raw.Trim().ToLowerInvariant())
            {
                case "a4":
                    return Word.WdPaperSize.wdPaperA4;
                case "a3":
                    return Word.WdPaperSize.wdPaperA3;
                case "a5":
                    return Word.WdPaperSize.wdPaperA5;
                case "letter":
                    return Word.WdPaperSize.wdPaperLetter;
                case "legal":
                    return Word.WdPaperSize.wdPaperLegal;
                case "b4":
                    return Word.WdPaperSize.wdPaperB4;
                case "b5":
                    return Word.WdPaperSize.wdPaperB5;
                case "executive":
                    return Word.WdPaperSize.wdPaperExecutive;
                default:
                    throw WordArguments.Invalid("不支持的 paperSize=" + raw +
                                                "；支持 A4/A3/A5/Letter/Legal/B4/B5/Executive");
            }
        }

        public static bool TriStateIsTrue(int value)
        {
            return value != 0;
        }

        public static void Release(params object[] objects)
        {
            ComHelper.Release(objects);
        }

        private static bool IsWholeWordMatch(string text, int start, int length)
        {
            bool leftOk = start == 0 || !IsWordCharacter(text[start - 1]);
            int end = start + length;
            bool rightOk = end >= text.Length || !IsWordCharacter(text[end]);
            return leftOk && rightOk;
        }

        private static bool IsWordCharacter(char value)
        {
            return char.IsLetterOrDigit(value) || value == '_';
        }
    }
}
