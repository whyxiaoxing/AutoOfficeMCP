using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    public sealed partial class WordToolExecutor
    {
        private void RegisterTableHandlers(
            Dictionary<string, Func<Word.Application, JObject, OperationResult>> handlers)
        {
            handlers[WordToolNames.InsertTable] = HandleInsertTable;
            handlers[WordToolNames.TablesList] = HandleTablesList;
            handlers[WordToolNames.TableRead] = HandleTableRead;
            handlers[WordToolNames.TableSetCell] = HandleTableSetCell;
            handlers[WordToolNames.TableSetRow] = HandleTableSetRow;
            handlers[WordToolNames.TableSetCol] = HandleTableSetCol;
            handlers[WordToolNames.TableAddRowCol] = HandleTableAddRowCol;
            handlers[WordToolNames.TableDeleteRowCol] = HandleTableDeleteRowCol;
            handlers[WordToolNames.TableMergeCells] = HandleTableMergeCells;
            handlers[WordToolNames.TableFormat] = HandleTableFormat;
            handlers[WordToolNames.TableFormula] = HandleTableFormula;
        }

        private OperationResult HandleInsertTable(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int rows = args.RequireInt("rows", 1, 500);
            int cols = args.RequireInt("cols", 1, 63);
            JArray data = args.GetOptionalArray("data");
            string style = args.Has("style") ? args.GetOptionalString("style", null) : null;
            List<double> columnWidths = ReadDoubleList(args, "columnWidths");
            bool caption = args.GetOptionalBool("caption", false);
            string position = (args.GetOptionalString("position", "cursor") ?? "cursor")
                .Trim().ToLowerInvariant();
            if (position != "start" && position != "cursor" && position != "end")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "position 只支持 start/cursor/end，收到 " + position + "。");
            }
            string dataError = ValidateTableDataShape(data, rows, cols);
            if (dataError != null)
            {
                return OperationResult.Failed(OfficeErrorCode.INVALID_PARAMS, dataError);
            }
            if (columnWidths != null && columnWidths.Count > cols)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "columnWidths 数量 " + columnWidths.Count + " 超过列数 " + cols + "。");
            }
            if (columnWidths != null)
            {
                foreach (double width in columnWidths)
                {
                    if (width <= 0)
                    {
                        return OperationResult.Failed(
                            OfficeErrorCode.INVALID_PARAMS,
                            "columnWidths 每一项必须是正数磅值，收到 " + width + "。");
                    }
                }
            }

            return RunWriteTool(
                application,
                WordToolNames.InsertTable,
                false,
                delegate(Word.Document document)
                {
                    return SelectTableInsertionTarget(application, document, position);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    if (!string.IsNullOrWhiteSpace(style) && !WordComUtil.StyleExists(document, style))
                    {
                        return WordWriteResult.Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "表格样式「" + style + "」不存在。请省略 style 使用默认样式，或先用 Word 的表格样式名称。",
                            null,
                            false);
                    }

                    Word.Range insertionRange = null;
                    Word.Tables tables = null;
                    Word.Table table = null;
                    Word.Rows headingRows = null;
                    Word.Row headingRow = null;
                    Word.Range headingRange = null;
                    Word.Font headingFont = null;
                    Word.Columns columns = null;
                    JArray columnWidthResults = new JArray();
                    JArray warnings = new JArray();
                    try
                    {
                        bool trailingParagraphAdded = false;
                        int insertPosition = target.Start;
                        tables = document.Tables;
                        int existingTableCount = tables == null ? 0 : tables.Count;
                        if (position == "end" && existingTableCount > 0)
                        {
                            // 文档末尾已有表格时，直接在该位置 Add 会落进最后一个单元格，
                            // 导致 Tables.Count 不增加。先补一个表后段落，再在段落处插入新表。
                            Word.Range content = null;
                            try
                            {
                                content = document.Content;
                                content.InsertAfter("\r");
                                insertPosition = Math.Max(0, content.End - 1);
                                trailingParagraphAdded = true;
                            }
                            finally
                            {
                                ComHelper.Release(content);
                            }
                        }
                        insertionRange = WordComUtil.CreateRange(document, insertPosition, insertPosition);
                        int beforeCount = tables == null ? 0 : tables.Count;
                        table = tables.Add(insertionRange, rows, cols);
                        int filledCells = 0;
                        if (data != null)
                        {
                            for (int rowIndex = Math.Min(rows, data.Count) - 1; rowIndex >= 0; rowIndex--)
                            {
                                JArray rowData = data[rowIndex] as JArray;
                                if (rowData == null)
                                {
                                    continue;
                                }
                                for (int columnIndex = Math.Min(cols, rowData.Count) - 1; columnIndex >= 0; columnIndex--)
                                {
                                    Word.Cell cell = null;
                                    Word.Range cellRange = null;
                                    try
                                    {
                                        cell = table.Cell(rowIndex + 1, columnIndex + 1);
                                        cellRange = cell.Range;
                                        cellRange.Text = TableCellValue(rowData[columnIndex]);
                                        filledCells++;
                                    }
                                    finally
                                    {
                                        ComHelper.Release(cellRange, cell);
                                    }
                                }
                            }
                        }

                        if (!string.IsNullOrWhiteSpace(style))
                        {
                            object styleObject = style;
                            table.set_Style(ref styleObject);
                        }
                        table.AllowAutoFit = columnWidths == null;

                        headingRows = table.Rows;
                        headingRow = headingRows[1];
                        headingRow.HeadingFormat = -1;
                        headingRange = headingRow.Range;
                        headingFont = headingRange.Font;
                        headingFont.Bold = -1;

                        if (columnWidths != null)
                        {
                            columns = table.Columns;
                            for (int index = columnWidths.Count - 1; index >= 0; index--)
                            {
                                Word.Column column = null;
                                try
                                {
                                    column = columns[index + 1];
                                    column.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
                                    column.PreferredWidth = (float)columnWidths[index];
                                    columnWidthResults.Add(new JObject
                                    {
                                        { "column", index + 1 },
                                        { "requestedPt", columnWidths[index] },
                                        { "actualPt", (double)column.PreferredWidth }
                                    });
                                }
                                finally
                                {
                                    ComHelper.Release(column);
                                }
                            }
                        }

                        bool captionInserted = false;
                        if (caption)
                        {
                            captionInserted = TryInsertTableCaption(table, warnings);
                        }

                        int afterCount = tables.Count;
                        int tableIndex = WordComUtil.FindTableIndex(document, table);
                        if (tableIndex <= 0)
                        {
                            tableIndex = afterCount;
                        }

                        JObject finalSummary = WordComUtil.TableSummary(table);
                        JObject resultData = new JObject();
                        resultData["tableIndex"] = tableIndex;
                        resultData["rows"] = rows;
                        resultData["columns"] = cols;
                        resultData["filledCells"] = filledCells;
                        resultData["style"] = style ?? string.Empty;
                        resultData["styleActual"] = finalSummary.Value<string>("style");
                        resultData["columnWidths"] = columnWidthResults;
                        resultData["captionRequested"] = caption;
                        resultData["captionInserted"] = captionInserted;
                        resultData["trailingParagraphAdded"] = trailingParagraphAdded;
                        resultData["warnings"] = warnings;
                        resultData["table"] = finalSummary;

                        if (afterCount <= beforeCount)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INTERNAL_ERROR,
                                "Tables.Add 已调用，但表格数量未增加（前 " + beforeCount + "，后 " +
                                afterCount + "）。",
                                resultData,
                                false);
                        }
                        if (caption && !captionInserted)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.PARTIAL_APPLY,
                                "表格已插入，但题注域插入失败（表样式/题注标签不可用）。请检查 Word 题注标签「表」是否存在。",
                                resultData,
                                true);
                        }

                        WordWriteResult result = WordWriteResult.Ok(
                            "已插入 " + rows + "×" + cols + " 表格（索引 " + tableIndex + "），首行已设为标题行并加粗" +
                            (captionInserted ? "，已插入表题注域" : "") + "。",
                            resultData,
                            true);
                        result.AddTargetRef("Word:Table:" + tableIndex);
                        return result;
                    }
                    finally
                    {
                        ComHelper.Release(columns, headingFont, headingRange, headingRow, headingRows,
                            table, tables, insertionRange);
                    }
                });
        }

        private OperationResult HandleTablesList(Word.Application application, JObject arguments)
        {
            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    Word.Tables tables = null;
                    JArray items = new JArray();
                    try
                    {
                        tables = document.Tables;
                        int count = tables == null ? 0 : tables.Count;
                        for (int index = 1; index <= count; index++)
                        {
                            Word.Table table = null;
                            Word.Rows rows = null;
                            Word.Columns columns = null;
                            try
                            {
                                table = tables[index];
                                rows = table.Rows;
                                columns = table.Columns;
                                int rowCount = rows == null ? 0 : rows.Count;
                                int columnCount = columns == null ? 0 : columns.Count;
                                JArray firstRow = new JArray();
                                int sampleColumns = Math.Min(columnCount, 4);
                                for (int columnIndex = 1; columnIndex <= sampleColumns; columnIndex++)
                                {
                                    Word.Cell cell = null;
                                    Word.Range cellRange = null;
                                    try
                                    {
                                        cell = table.Cell(1, columnIndex);
                                        cellRange = cell.Range;
                                        firstRow.Add(WordComUtil.NormalizePreview(
                                            WordComUtil.CleanCellText(cellRange.Text), 50));
                                    }
                                    catch
                                    {
                                        firstRow.Add("[合并/不可读]");
                                    }
                                    finally
                                    {
                                        ComHelper.Release(cellRange, cell);
                                    }
                                }
                                items.Add(new JObject
                                {
                                    { "index", index },
                                    { "rows", rowCount },
                                    { "columns", columnCount },
                                    { "nestingLevel", table.NestingLevel },
                                    { "uniform", table.Uniform },
                                    { "style", WordComUtil.ReadTableStyleName(table) },
                                    { "firstRowPreview", firstRow }
                                });
                            }
                            finally
                            {
                                ComHelper.Release(columns, rows, table);
                            }
                        }
                        JObject data = new JObject();
                        data["items"] = items;
                        data["total"] = count;
                        data["limitation"] =
                            "只枚举 doc.Tables（主文档正文）。页眉页脚、文本框内的表格不在该集合，需在 Word 内单独查看。";
                        return ReadSucceed("共 " + count + " 个正文表格。", data, "Word:Tables");
                    }
                    finally
                    {
                        ComHelper.Release(tables);
                    }
                });
        }

        private OperationResult HandleTableRead(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int tableIndex = args.RequireInt("tableIndex", 1, int.MaxValue);
            int? startRow = args.GetOptionalInt("startRow");
            int? endRow = args.GetOptionalInt("endRow");
            int? startCol = args.GetOptionalInt("startCol");
            int? endCol = args.GetOptionalInt("endCol");

            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    Word.Table table = null;
                    Word.Rows tableRows = null;
                    Word.Columns tableColumns = null;
                    try
                    {
                        table = WordComUtil.GetTable(document, tableIndex);
                        tableRows = table.Rows;
                        tableColumns = table.Columns;
                        int rowCount = tableRows.Count;
                        int columnCount = tableColumns.Count;
                        int firstRow = startRow.HasValue ? startRow.Value : 1;
                        int lastRow = endRow.HasValue ? endRow.Value : rowCount;
                        int firstColumn = startCol.HasValue ? startCol.Value : 1;
                        int lastColumn = endCol.HasValue ? endCol.Value : columnCount;
                        if (firstRow < 1 || lastRow > rowCount || firstRow > lastRow)
                        {
                            return OperationResult.Failed(
                                OfficeErrorCode.INVALID_PARAMS,
                                "表格行范围无效：收到 " + firstRow + "-" + lastRow +
                                "，当前表共 " + rowCount + " 行（1 基）。");
                        }
                        if (firstColumn < 1 || lastColumn > columnCount || firstColumn > lastColumn)
                        {
                            return OperationResult.Failed(
                                OfficeErrorCode.INVALID_PARAMS,
                                "表格列范围无效：收到 " + firstColumn + "-" + lastColumn +
                                "，当前表共 " + columnCount + " 列（1 基）。");
                        }

                        JArray matrix = new JArray();
                        int unreadable = 0;
                        for (int row = firstRow; row <= lastRow; row++)
                        {
                            JArray rowValues = new JArray();
                            for (int column = firstColumn; column <= lastColumn; column++)
                            {
                                Word.Cell cell = null;
                                Word.Range cellRange = null;
                                try
                                {
                                    cell = table.Cell(row, column);
                                    cellRange = cell.Range;
                                    rowValues.Add(WordComUtil.CleanCellText(cellRange.Text));
                                }
                                catch
                                {
                                    unreadable++;
                                    rowValues.Add(JValue.CreateNull());
                                }
                                finally
                                {
                                    ComHelper.Release(cellRange, cell);
                                }
                            }
                            matrix.Add(rowValues);
                        }

                        JObject data = new JObject();
                        data["tableIndex"] = tableIndex;
                        data["totalRows"] = rowCount;
                        data["totalColumns"] = columnCount;
                        data["startRow"] = firstRow;
                        data["endRow"] = lastRow;
                        data["startColumn"] = firstColumn;
                        data["endColumn"] = lastColumn;
                        data["rows"] = matrix;
                        data["uniform"] = table.Uniform;
                        data["unreadableCellCount"] = unreadable;
                        data["mergedCellsSuspected"] = unreadable > 0 || !table.Uniform;
                        data["note"] = "单元格值已剥离 Word 单元格结束标记 \\r\\a；null 表示该索引处单元格不存在（通常是合并区域）。";
                        return ReadSucceed(
                            "已读取表格 " + tableIndex + " 的 " + (lastRow - firstRow + 1) + "×" +
                            (lastColumn - firstColumn + 1) + " 单元格。",
                            data,
                            "Word:Table:" + tableIndex);
                    }
                    finally
                    {
                        ComHelper.Release(tableColumns, tableRows, table);
                    }
                });
        }

        private OperationResult HandleTableSetCell(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int tableIndex = args.RequireInt("tableIndex", 1, int.MaxValue);
            int row = args.RequireInt("row", 1, int.MaxValue);
            int column = args.RequireInt("col", 1, int.MaxValue);
            if (!args.Has("text"))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "缺少 text 参数（允许空字符串）。");
            }
            string text = args.GetOptionalString("text", string.Empty);
            JObject format = args.GetOptionalObject("format");
            string formatError;
            ValidateCellFormatArguments(format, out formatError);
            if (formatError != null)
            {
                return OperationResult.Failed(OfficeErrorCode.INVALID_PARAMS, formatError);
            }

            return RunWriteTool(
                application,
                WordToolNames.TableSetCell,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.Table(document, tableIndex);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Table table = null;
                    Word.Cell cell = null;
                    Word.Range cellRange = null;
                    Word.Font font = null;
                    try
                    {
                        table = WordComUtil.GetTable(document, tableIndex);
                        WordComUtil.ValidateCellIndex(table, row, column);
                        cell = WordComUtil.GetCell(table, row, column);
                        cellRange = cell.Range;
                        string beforeText = WordComUtil.CleanCellText(cellRange.Text);
                        font = cellRange.Font;
                        JObject beforeFormat = WordFormatUtil.CaptureFont(font);
                        cellRange.Text = WordFormatUtil.NormalizeForWord(text);
                        JArray formatVerification = new JArray();
                        if (format != null)
                        {
                            ApplyCellFormat(cellRange, format, formatVerification);
                        }
                        string afterText = WordComUtil.CleanCellText(cellRange.Text);
                        JObject afterFormat = WordFormatUtil.CaptureFont(font);
                        bool textMatched = string.Equals(
                            WordFormatUtil.NormalizeForCompare(afterText),
                            WordFormatUtil.NormalizeForCompare(WordFormatUtil.NormalizeForWord(text)),
                            StringComparison.Ordinal);
                        bool formatMatched = AllVerificationsMatched(formatVerification);
                        bool changed = !string.Equals(beforeText, afterText, StringComparison.Ordinal) ||
                                       !JToken.DeepEquals(beforeFormat, afterFormat);

                        JObject data = new JObject();
                        data["tableIndex"] = tableIndex;
                        data["row"] = row;
                        data["col"] = column;
                        data["text"] = afterText;
                        data["textMatched"] = textMatched;
                        data["formatVerification"] = formatVerification;
                        data["formatMatched"] = formatMatched;
                        data["table"] = WordComUtil.TableSummary(table);

                        if (!textMatched)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "单元格写入后回读不一致：期望「" +
                                WordComUtil.NormalizePreview(text, 80) + "」，实际「" +
                                WordComUtil.NormalizePreview(afterText, 80) + "」。",
                                data,
                                changed);
                        }
                        if (!formatMatched)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "单元格文本已写入，但字体格式回读不一致；请查看 formatVerification。",
                                data,
                                changed);
                        }
                        return WordWriteResult.Ok(
                            "已写入表格 " + tableIndex + " 第 " + row + " 行第 " + column + " 列。",
                            data,
                            changed);
                    }
                    finally
                    {
                        ComHelper.Release(font, cellRange, cell, table);
                    }
                });
        }

        private OperationResult HandleTableSetRow(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int tableIndex = args.RequireInt("tableIndex", 1, int.MaxValue);
            int row = args.RequireInt("row", 1, int.MaxValue);
            List<string> values = args.GetOptionalStringList("values");
            if (values == null || values.Count == 0)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "缺少 values 数组（按列顺序写入整行）。");
            }
            JObject format = args.GetOptionalObject("format");
            string formatError;
            ValidateCellFormatArguments(format, out formatError);
            if (formatError != null)
            {
                return OperationResult.Failed(OfficeErrorCode.INVALID_PARAMS, formatError);
            }

            return RunWriteTool(
                application,
                WordToolNames.TableSetRow,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.Table(document, tableIndex);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Table table = null;
                    Word.Row tableRow = null;
                    Word.Range rowRange = null;
                    Word.Columns columns = null;
                    try
                    {
                        table = WordComUtil.GetTable(document, tableIndex);
                        WordComUtil.ValidateCellIndex(table, row, 1);
                        columns = table.Columns;
                        if (values.Count > columns.Count)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "values 数量 " + values.Count + " 超过表格列数 " + columns.Count + "。",
                                null,
                                false);
                        }
                        tableRow = table.Rows[row];
                        rowRange = tableRow.Range;
                        JArray mismatches = new JArray();
                        for (int column = values.Count; column >= 1; column--)
                        {
                            Word.Cell cell = null;
                            Word.Range cellRange = null;
                            try
                            {
                                cell = table.Cell(row, column);
                                cellRange = cell.Range;
                                cellRange.Text = WordFormatUtil.NormalizeForWord(values[column - 1]);
                                string actual = WordComUtil.CleanCellText(cellRange.Text);
                                if (!string.Equals(
                                    WordFormatUtil.NormalizeForCompare(actual),
                                    WordFormatUtil.NormalizeForCompare(
                                        WordFormatUtil.NormalizeForWord(values[column - 1])),
                                    StringComparison.Ordinal))
                                {
                                    mismatches.Add(new JObject
                                    {
                                        { "col", column },
                                        { "expected", values[column - 1] },
                                        { "actual", actual }
                                    });
                                }
                            }
                            finally
                            {
                                ComHelper.Release(cellRange, cell);
                            }
                        }
                        JArray formatVerification = new JArray();
                        if (format != null)
                        {
                            ApplyCellFormat(rowRange, format, formatVerification);
                        }

                        JObject data = new JObject();
                        data["tableIndex"] = tableIndex;
                        data["row"] = row;
                        data["values"] = new JArray(values.ToArray());
                        data["mismatches"] = mismatches;
                        data["formatVerification"] = formatVerification;
                        data["table"] = WordComUtil.TableSummary(table);
                        if (mismatches.Count > 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "整行写入后回读不一致：" + mismatches[0].ToString(),
                                data,
                                true);
                        }
                        if (!AllVerificationsMatched(formatVerification))
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "整行文本已写入，但字体格式回读不一致。",
                                data,
                                true);
                        }
                        return WordWriteResult.Ok(
                            "已写入表格 " + tableIndex + " 第 " + row + " 行 " + values.Count + " 个单元格。",
                            data,
                            true);
                    }
                    finally
                    {
                        ComHelper.Release(columns, rowRange, tableRow, table);
                    }
                });
        }

        private OperationResult HandleTableSetCol(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int tableIndex = args.RequireInt("tableIndex", 1, int.MaxValue);
            int column = args.RequireInt("col", 1, int.MaxValue);
            List<string> values = args.GetOptionalStringList("values");
            if (values == null || values.Count == 0)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "缺少 values 数组（按行顺序写入整列）。");
            }
            JObject format = args.GetOptionalObject("format");
            string formatError;
            ValidateCellFormatArguments(format, out formatError);
            if (formatError != null)
            {
                return OperationResult.Failed(OfficeErrorCode.INVALID_PARAMS, formatError);
            }

            return RunWriteTool(
                application,
                WordToolNames.TableSetCol,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.Table(document, tableIndex);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Table table = null;
                    Word.Rows rows = null;
                    try
                    {
                        table = WordComUtil.GetTable(document, tableIndex);
                        WordComUtil.ValidateCellIndex(table, 1, column);
                        rows = table.Rows;
                        if (values.Count > rows.Count)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "values 数量 " + values.Count + " 超过表格行数 " + rows.Count + "。",
                                null,
                                false);
                        }
                        JArray mismatches = new JArray();
                        JArray formatVerification = new JArray();
                        for (int row = values.Count; row >= 1; row--)
                        {
                            Word.Cell cell = null;
                            Word.Range cellRange = null;
                            try
                            {
                                cell = table.Cell(row, column);
                                cellRange = cell.Range;
                                cellRange.Text = WordFormatUtil.NormalizeForWord(values[row - 1]);
                                if (format != null)
                                {
                                    ApplyCellFormat(cellRange, format, formatVerification);
                                }
                                string actual = WordComUtil.CleanCellText(cellRange.Text);
                                if (!string.Equals(
                                    WordFormatUtil.NormalizeForCompare(actual),
                                    WordFormatUtil.NormalizeForCompare(
                                        WordFormatUtil.NormalizeForWord(values[row - 1])),
                                    StringComparison.Ordinal))
                                {
                                    mismatches.Add(new JObject
                                    {
                                        { "row", row },
                                        { "expected", values[row - 1] },
                                        { "actual", actual }
                                    });
                                }
                            }
                            finally
                            {
                                ComHelper.Release(cellRange, cell);
                            }
                        }

                        JObject data = new JObject();
                        data["tableIndex"] = tableIndex;
                        data["col"] = column;
                        data["values"] = new JArray(values.ToArray());
                        data["mismatches"] = mismatches;
                        data["formatVerification"] = formatVerification;
                        data["table"] = WordComUtil.TableSummary(table);
                        if (mismatches.Count > 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "整列写入后回读不一致：" + mismatches[0].ToString(),
                                data,
                                true);
                        }
                        if (!AllVerificationsMatched(formatVerification))
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "整列文本已写入，但字体格式回读不一致。",
                                data,
                                true);
                        }
                        return WordWriteResult.Ok(
                            "已写入表格 " + tableIndex + " 第 " + column + " 列 " + values.Count + " 个单元格。",
                            data,
                            true);
                    }
                    finally
                    {
                        ComHelper.Release(rows, table);
                    }
                });
        }

        private OperationResult HandleTableAddRowCol(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int tableIndex = args.RequireInt("tableIndex", 1, int.MaxValue);
            int rowCount = args.GetOptionalInt("rowCount", 0);
            int colCount = args.GetOptionalInt("colCount", 0);
            if (rowCount < 0 || colCount < 0 || rowCount > 200 || colCount > 63)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "rowCount 允许 0-200、colCount 允许 0-63，收到 " + rowCount + "/" + colCount + "。");
            }
            if (rowCount == 0 && colCount == 0)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "rowCount 与 colCount 至少有一个大于 0。");
            }

            return RunWriteTool(
                application,
                WordToolNames.TableAddRowCol,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.Table(document, tableIndex);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Table table = null;
                    Word.Rows rows = null;
                    Word.Columns columns = null;
                    try
                    {
                        table = WordComUtil.GetTable(document, tableIndex);
                        rows = table.Rows;
                        columns = table.Columns;
                        int beforeRows = rows.Count;
                        int beforeColumns = columns.Count;
                        ComHelper.Release(rows);
                        ComHelper.Release(columns);
                        rows = null;
                        columns = null;
                        JArray failures = new JArray();

                        object missing = Type.Missing;
                        if (rowCount > 0)
                        {
                            rows = table.Rows;
                            for (int index = rowCount; index > 0; index--)
                            {
                                Word.Row newRow = null;
                                try
                                {
                                    newRow = rows.Add(ref missing);
                                }
                                catch (Exception ex)
                                {
                                    failures.Add(new JObject
                                    {
                                        { "action", "addRow" },
                                        { "message", OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) }
                                    });
                                    break;
                                }
                                finally
                                {
                                    ComHelper.Release(newRow);
                                }
                            }
                            ComHelper.Release(rows);
                            rows = null;
                        }
                        if (colCount > 0)
                        {
                            // 增加行后 Columns 集合可能失效，必须重新获取。
                            ComHelper.Release(table);
                            table = WordComUtil.GetTable(document, tableIndex);
                            columns = table.Columns;
                            for (int index = colCount; index > 0; index--)
                            {
                                Word.Column newColumn = null;
                                try
                                {
                                    newColumn = columns.Add(ref missing);
                                }
                                catch (Exception ex)
                                {
                                    failures.Add(new JObject
                                    {
                                        { "action", "addColumn" },
                                        { "message", OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) }
                                    });
                                    break;
                                }
                                finally
                                {
                                    ComHelper.Release(newColumn);
                                }
                            }
                            ComHelper.Release(columns);
                            columns = null;
                        }

                        ComHelper.Release(table);
                        table = WordComUtil.GetTable(document, tableIndex);
                        rows = table.Rows;
                        columns = table.Columns;
                        int afterRows = rows.Count;
                        int afterColumns = columns.Count;
                        bool changed = afterRows != beforeRows || afterColumns != beforeColumns;
                        JObject data = new JObject();
                        data["tableIndex"] = tableIndex;
                        data["beforeRows"] = beforeRows;
                        data["beforeColumns"] = beforeColumns;
                        data["afterRows"] = afterRows;
                        data["afterColumns"] = afterColumns;
                        data["failures"] = failures;
                        data["table"] = WordComUtil.TableSummary(table);

                        if (failures.Count > 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "表格行列增加部分失败：" + failures[0]["message"].ToString() +
                                "。含合并单元格的表格可能拒绝 Rows.Add/Columns.Add；请先在 Word 中处理合并。",
                                data,
                                changed);
                        }
                        if (afterRows - beforeRows != rowCount || afterColumns - beforeColumns != colCount)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "行列数量变化与请求不一致：请求 +" + rowCount + " 行/+" + colCount +
                                " 列，实际 +" + (afterRows - beforeRows) + " 行/+" +
                                (afterColumns - beforeColumns) + " 列。",
                                data,
                                changed);
                        }
                        return WordWriteResult.Ok(
                            "已为表格 " + tableIndex + " 增加 " + rowCount + " 行、" + colCount + " 列。",
                            data,
                            changed);
                    }
                    finally
                    {
                        ComHelper.Release(columns, rows, table);
                    }
                });
        }

        private OperationResult HandleTableDeleteRowCol(Word.Application application, JObject arguments)
        {
            SecurityDecision decision = SecurityGate.Evaluate(
                SecurityRisk.DeleteSheet, WordToolNames.TableDeleteRowCol);
            if (!decision.Allowed)
            {
                return OperationResult.Failed(
                    decision.ErrorCode,
                    "删除表格行/列被安全策略取消（当前复用 security.allowDeleteSheet 配置）：" +
                    decision.Message);
            }

            WordArguments args = new WordArguments(arguments);
            int tableIndex = args.RequireInt("tableIndex", 1, int.MaxValue);
            string targetKind = args.RequireString("target").Trim().ToLowerInvariant();
            bool deleteRow = targetKind == "row" || targetKind == "行";
            bool deleteColumn = targetKind == "column" || targetKind == "col" || targetKind == "列";
            if (!deleteRow && !deleteColumn)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "target 只支持 row/column，收到 " + targetKind + "。");
            }
            int index = args.RequireInt("index", 1, int.MaxValue);
            int count = args.GetOptionalInt("count", 1);
            if (count < 1 || count > 200)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "count 允许 1-200，收到 " + count + "。");
            }

            return RunWriteTool(
                application,
                WordToolNames.TableDeleteRowCol,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.Table(document, tableIndex);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Table table = null;
                    Word.Rows rows = null;
                    Word.Columns columns = null;
                    try
                    {
                        table = WordComUtil.GetTable(document, tableIndex);
                        rows = table.Rows;
                        columns = table.Columns;
                        int beforeRows = rows.Count;
                        int beforeColumns = columns.Count;
                        int available = deleteRow ? beforeRows : beforeColumns;
                        if (index + count - 1 > available)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "删除范围越界：" + (deleteRow ? "行" : "列") + " " + index + "-" +
                                (index + count - 1) + "，当前共 " + available + " " +
                                (deleteRow ? "行" : "列") + "。",
                                null,
                                false);
                        }

                        JArray failures = new JArray();
                        for (int current = index + count - 1; current >= index; current--)
                        {
                            try
                            {
                                if (deleteRow)
                                {
                                    Word.Row row = null;
                                    try
                                    {
                                        row = rows[current];
                                        row.Delete();
                                    }
                                    finally
                                    {
                                        ComHelper.Release(row);
                                    }
                                }
                                else
                                {
                                    Word.Column column = null;
                                    try
                                    {
                                        column = columns[current];
                                        column.Delete();
                                    }
                                    finally
                                    {
                                        ComHelper.Release(column);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                failures.Add(new JObject
                                {
                                    { "index", current },
                                    { "message", OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) }
                                });
                            }
                        }

                        ComHelper.Release(rows);
                        ComHelper.Release(columns);
                        rows = table.Rows;
                        columns = table.Columns;
                        int afterRows = rows.Count;
                        int afterColumns = columns.Count;
                        bool changed = afterRows != beforeRows || afterColumns != beforeColumns;
                        JObject data = new JObject();
                        data["tableIndex"] = tableIndex;
                        data["target"] = deleteRow ? "row" : "column";
                        data["index"] = index;
                        data["count"] = count;
                        data["beforeRows"] = beforeRows;
                        data["beforeColumns"] = beforeColumns;
                        data["afterRows"] = afterRows;
                        data["afterColumns"] = afterColumns;
                        data["failures"] = failures;
                        data["table"] = WordComUtil.TableSummary(table);

                        if (failures.Count > 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "删除表格" + (deleteRow ? "行" : "列") + "部分失败：" +
                                failures[0]["message"].ToString() +
                                "。含合并单元格时请先在 Word 中拆分/调整。",
                                data,
                                changed);
                        }
                        if (!changed)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "删除命令已执行，但表格行列数未变化（可能是合并单元格或索引指向同一合并区域）。",
                                data,
                                false);
                        }
                        return WordWriteResult.Ok(
                            "已删除表格 " + tableIndex + " 的 " + (deleteRow ? "行" : "列") + " " +
                            index + "-" + (index + count - 1) + "。",
                            data,
                            changed);
                    }
                    finally
                    {
                        ComHelper.Release(columns, rows, table);
                    }
                });
        }

        private OperationResult HandleTableMergeCells(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int tableIndex = args.RequireInt("tableIndex", 1, int.MaxValue);
            int startRow = args.RequireInt("startRow", 1, int.MaxValue);
            int startCol = args.RequireInt("startCol", 1, int.MaxValue);
            int endRow = args.RequireInt("endRow", 1, int.MaxValue);
            int endCol = args.RequireInt("endCol", 1, int.MaxValue);
            if (endRow < startRow || endCol < startCol)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "合并区域无效：endRow/endCol 必须大于等于 startRow/startCol。");
            }

            return RunWriteTool(
                application,
                WordToolNames.TableMergeCells,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.Table(document, tableIndex);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Table table = null;
                    Word.Cell startCell = null;
                    Word.Cell endCell = null;
                    try
                    {
                        table = WordComUtil.GetTable(document, tableIndex);
                        WordComUtil.ValidateCellIndex(table, endRow, endCol);
                        startCell = WordComUtil.GetCell(table, startRow, startCol);
                        endCell = WordComUtil.GetCell(table, endRow, endCol);
                        startCell.Merge(endCell);
                        JObject summary = WordComUtil.TableSummary(table);
                        JObject data = new JObject();
                        data["tableIndex"] = tableIndex;
                        data["startRow"] = startRow;
                        data["startCol"] = startCol;
                        data["endRow"] = endRow;
                        data["endCol"] = endCol;
                        data["table"] = summary;
                        data["note"] = "合并后表格行列索引会变化；请重新调用 word_table_read 读取实际结构。";
                        return WordWriteResult.Ok(
                            "已合并表格 " + tableIndex + " 的 (" + startRow + "," + startCol + ")-(" +
                            endRow + "," + endCol + ")。合并后行列索引可能变化。",
                            data,
                            true);
                    }
                    finally
                    {
                        ComHelper.Release(endCell, startCell, table);
                    }
                });
        }

        private OperationResult HandleTableFormat(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int tableIndex = args.RequireInt("tableIndex", 1, int.MaxValue);
            string borderStyle = args.Has("borderStyle") ? args.GetOptionalString("borderStyle", null) : null;
            int? borderColor = args.Has("borderColor")
                ? WordFormatUtil.ParseColor(args.GetOptionalString("borderColor", null))
                : (int?)null;
            double? borderWidthPt = args.GetOptionalDouble("borderWidthPt");
            int? shadingColor = args.Has("shadingColor")
                ? WordFormatUtil.ParseColor(args.GetOptionalString("shadingColor", null))
                : (int?)null;
            List<double> columnWidths = ReadDoubleList(args, "columnWidths");
            string autofit = args.Has("autofit") ? args.GetOptionalString("autofit", null) : null;
            string alignment = args.Has("alignment") ? args.GetOptionalString("alignment", null) : null;
            bool? repeatHeaderRow = ReadNullableBool(args, "repeatHeaderRow");
            string verticalAlignment = args.Has("verticalAlignment")
                ? args.GetOptionalString("verticalAlignment", null)
                : null;

            if (string.IsNullOrWhiteSpace(borderStyle) && !borderColor.HasValue &&
                !borderWidthPt.HasValue && !shadingColor.HasValue &&
                (columnWidths == null || columnWidths.Count == 0) &&
                string.IsNullOrWhiteSpace(autofit) && string.IsNullOrWhiteSpace(alignment) &&
                !repeatHeaderRow.HasValue && string.IsNullOrWhiteSpace(verticalAlignment))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "word_table_format 至少需要一个设置：borderStyle/borderColor/borderWidthPt/" +
                    "shadingColor/columnWidths/autofit/alignment/repeatHeaderRow/verticalAlignment。");
            }

            Word.WdLineStyle? lineStyle = null;
            if (!string.IsNullOrWhiteSpace(borderStyle))
            {
                lineStyle = ParseBorderStyle(borderStyle);
            }
            Word.WdLineWidth? lineWidth = borderWidthPt.HasValue
                ? MapBorderWidth(borderWidthPt.Value)
                : (Word.WdLineWidth?)null;
            Word.WdAutoFitBehavior? autoFitBehavior = null;
            if (!string.IsNullOrWhiteSpace(autofit))
            {
                switch (autofit.Trim().ToLowerInvariant())
                {
                    case "content":
                        autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;
                        break;
                    case "window":
                        autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;
                        break;
                    case "fixed":
                        autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitFixed;
                        break;
                    default:
                        return OperationResult.Failed(
                            OfficeErrorCode.INVALID_PARAMS,
                            "autofit 只支持 content/window/fixed，收到 " + autofit + "。");
                }
            }
            Word.WdRowAlignment? rowAlignment = null;
            if (!string.IsNullOrWhiteSpace(alignment))
            {
                switch (alignment.Trim().ToLowerInvariant())
                {
                    case "left":
                        rowAlignment = Word.WdRowAlignment.wdAlignRowLeft;
                        break;
                    case "center":
                        rowAlignment = Word.WdRowAlignment.wdAlignRowCenter;
                        break;
                    case "right":
                        rowAlignment = Word.WdRowAlignment.wdAlignRowRight;
                        break;
                    default:
                        return OperationResult.Failed(
                            OfficeErrorCode.INVALID_PARAMS,
                            "alignment 只支持 left/center/right，收到 " + alignment + "。");
                }
            }
            Word.WdCellVerticalAlignment? cellAlignment = null;
            if (!string.IsNullOrWhiteSpace(verticalAlignment))
            {
                switch (verticalAlignment.Trim().ToLowerInvariant())
                {
                    case "top":
                        cellAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                        break;
                    case "center":
                        cellAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        break;
                    case "bottom":
                        cellAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom;
                        break;
                    default:
                        return OperationResult.Failed(
                            OfficeErrorCode.INVALID_PARAMS,
                            "verticalAlignment 只支持 top/center/bottom，收到 " + verticalAlignment + "。");
                }
            }

            return RunWriteTool(
                application,
                WordToolNames.TableFormat,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.Table(document, tableIndex);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Table table = null;
                    Word.Borders borders = null;
                    Word.Shading shading = null;
                    Word.Columns columns = null;
                    Word.Rows rows = null;
                    Word.Cells cells = null;
                    try
                    {
                        table = WordComUtil.GetTable(document, tableIndex);
                        JObject before = CaptureTableFormatState(table);
                        JArray failures = new JArray();
                        JArray widthResults = new JArray();

                        if (autoFitBehavior.HasValue)
                        {
                            table.AutoFitBehavior(autoFitBehavior.Value);
                        }
                        if (lineStyle.HasValue || borderColor.HasValue || lineWidth.HasValue)
                        {
                            borders = table.Borders;
                            if (lineStyle.HasValue)
                            {
                                borders.OutsideLineStyle = lineStyle.Value;
                                borders.InsideLineStyle = lineStyle.Value;
                            }
                            if (borderColor.HasValue)
                            {
                                borders.OutsideColor = (Word.WdColor)borderColor.Value;
                                borders.InsideColor = (Word.WdColor)borderColor.Value;
                            }
                            if (lineWidth.HasValue)
                            {
                                borders.OutsideLineWidth = lineWidth.Value;
                                borders.InsideLineWidth = lineWidth.Value;
                            }
                        }
                        if (shadingColor.HasValue)
                        {
                            shading = table.Shading;
                            shading.BackgroundPatternColor = (Word.WdColor)shadingColor.Value;
                        }
                        if (columnWidths != null && columnWidths.Count > 0)
                        {
                            columns = table.Columns;
                            if (columnWidths.Count > columns.Count)
                            {
                                failures.Add(new JObject
                                {
                                    { "target", "columnWidths" },
                                    { "message", "columnWidths 数量 " + columnWidths.Count +
                                                 " 超过表格列数 " + columns.Count }
                                });
                            }
                            else
                            {
                                table.AllowAutoFit = false;
                                for (int index = columnWidths.Count - 1; index >= 0; index--)
                                {
                                    Word.Column column = null;
                                    try
                                    {
                                        column = columns[index + 1];
                                        column.PreferredWidthType =
                                            Word.WdPreferredWidthType.wdPreferredWidthPoints;
                                        column.PreferredWidth = (float)columnWidths[index];
                                        widthResults.Add(new JObject
                                        {
                                            { "column", index + 1 },
                                            { "requestedPt", columnWidths[index] },
                                            { "actualPt", (double)column.PreferredWidth }
                                        });
                                    }
                                    finally
                                    {
                                        ComHelper.Release(column);
                                    }
                                }
                            }
                        }
                        if (rowAlignment.HasValue)
                        {
                            rows = table.Rows;
                            rows.Alignment = rowAlignment.Value;
                        }
                        if (repeatHeaderRow.HasValue)
                        {
                            if (rows == null)
                            {
                                rows = table.Rows;
                            }
                            Word.Row firstRow = null;
                            try
                            {
                                firstRow = rows[1];
                                firstRow.HeadingFormat = repeatHeaderRow.Value ? -1 : 0;
                            }
                            finally
                            {
                                ComHelper.Release(firstRow);
                            }
                        }
                        if (cellAlignment.HasValue)
                        {
                            Word.Range tableRange = null;
                            try
                            {
                                tableRange = table.Range;
                                cells = tableRange.Cells;
                                cells.VerticalAlignment = cellAlignment.Value;
                            }
                            finally
                            {
                                ComHelper.Release(tableRange);
                            }
                        }

                        JObject after = CaptureTableFormatState(table);
                        bool changed = !JToken.DeepEquals(before, after);
                        JObject data = new JObject();
                        data["tableIndex"] = tableIndex;
                        data["before"] = before;
                        data["after"] = after;
                        data["columnWidths"] = widthResults;
                        data["failures"] = failures;
                        data["table"] = WordComUtil.TableSummary(table);

                        if (failures.Count > 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "表格格式部分失败：" + failures[0]["message"].ToString(),
                                data,
                                changed);
                        }
                        return WordWriteResult.Ok(
                            "表格 " + tableIndex + " 格式已应用（边框/底纹/列宽/对齐/标题行按请求处理）。",
                            data,
                            changed);
                    }
                    finally
                    {
                        ComHelper.Release(cells, rows, columns, shading, borders, table);
                    }
                });
        }

        private OperationResult HandleTableFormula(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int tableIndex = args.RequireInt("tableIndex", 1, int.MaxValue);
            int row = args.RequireInt("row", 1, int.MaxValue);
            int column = args.RequireInt("col", 1, int.MaxValue);
            string formula = args.RequireString("formula");
            string numberFormat = args.Has("numberFormat")
                ? args.GetOptionalString("numberFormat", null)
                : null;
            if (formula.IndexOf("=", StringComparison.Ordinal) != 0)
            {
                formula = "=" + formula;
            }
            if (formula.IndexOf("\\#", StringComparison.Ordinal) < 0 &&
                !string.IsNullOrWhiteSpace(numberFormat))
            {
                formula = formula + " \\# \"" + numberFormat.Replace("\"", string.Empty) + "\"";
            }

            return RunWriteTool(
                application,
                WordToolNames.TableFormula,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.Table(document, tableIndex);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Table table = null;
                    Word.Cell cell = null;
                    Word.Range cellRange = null;
                    Word.Range fieldRange = null;
                    Word.Fields fields = null;
                    Word.Field field = null;
                    Word.Range codeRange = null;
                    Word.Range resultRange = null;
                    try
                    {
                        table = WordComUtil.GetTable(document, tableIndex);
                        WordComUtil.ValidateCellIndex(table, row, column);
                        cell = WordComUtil.GetCell(table, row, column);
                        cellRange = cell.Range;
                        int fieldPosition = cellRange.Start;
                        cellRange.Text = string.Empty;
                        fieldRange = WordComUtil.CreateRange(document, fieldPosition, fieldPosition);
                        fields = fieldRange.Fields;
                        object fieldType = Word.WdFieldType.wdFieldEmpty;
                        object fieldText = formula;
                        object preserve = false;
                        field = fields.Add(fieldRange, ref fieldType, ref fieldText, ref preserve);
                        if (field == null)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "表格公式域创建失败（Fields.Add 返回 null）。请检查公式语法。",
                                null,
                                false);
                        }
                        bool updated = field.Update();
                        codeRange = field.Code;
                        resultRange = field.Result;
                        string codeText = codeRange == null ? string.Empty : (codeRange.Text ?? string.Empty);
                        string resultText = resultRange == null ? string.Empty : (resultRange.Text ?? string.Empty);
                        string cellText = WordComUtil.CleanCellText(cellRange.Text);
                        bool formulaError = resultText.IndexOf("!") >= 0 ||
                                            resultText.IndexOf("错误", StringComparison.Ordinal) >= 0 ||
                                            resultText.IndexOf("Error", StringComparison.OrdinalIgnoreCase) >= 0;

                        JObject data = new JObject();
                        data["tableIndex"] = tableIndex;
                        data["row"] = row;
                        data["col"] = column;
                        data["formula"] = formula;
                        data["fieldCode"] = WordComUtil.NormalizePreview(codeText, 160);
                        data["fieldResult"] = WordComUtil.NormalizePreview(resultText, 160);
                        data["cellText"] = cellText;
                        data["fieldUpdated"] = updated;
                        data["numberFormat"] = numberFormat ?? string.Empty;

                        if (formulaError)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "表格公式已写入但 Word 返回错误结果「" +
                                WordComUtil.NormalizePreview(resultText, 80) +
                                "」。请检查 =SUM(ABOVE)/=AVERAGE(LEFT) 的方向与相邻单元格是否有数值。",
                                data,
                                true);
                        }
                        return WordWriteResult.Ok(
                            "已向表格 " + tableIndex + " 的 (" + row + "," + column + ") 写入公式域 " +
                            formula + "，结果「" + WordComUtil.NormalizePreview(resultText, 80) + "」。",
                            data,
                            true);
                    }
                    finally
                    {
                        ComHelper.Release(resultRange, codeRange, field, fields, fieldRange, cellRange, cell, table);
                    }
                });
        }

        private static WordWriteTarget SelectTableInsertionTarget(
            Word.Application application,
            Word.Document document,
            string position)
        {
            if (position == "start")
            {
                return new WordWriteTarget(0, 0, "文档开头", "Word:Table", false, true);
            }
            if (position == "end")
            {
                int end = Math.Max(0, WordComUtil.GetContentEnd(document) - 1);
                return new WordWriteTarget(end, end, "文档末尾", "Word:Table", false, true);
            }
            Word.Selection selection = null;
            Word.Range range = null;
            try
            {
                selection = application.Selection;
                range = selection == null ? null : selection.Range;
                if (range == null)
                {
                    throw WordArguments.Invalid("position=cursor 时没有有效选区，无法插入表格。");
                }
                return new WordWriteTarget(range.Start, range.Start, "当前光标位置", "Word:Table", false, true);
            }
            finally
            {
                ComHelper.Release(range, selection);
            }
        }

        private static string ValidateTableDataShape(JArray data, int rows, int cols)
        {
            if (data == null)
            {
                return null;
            }
            if (data.Count > rows)
            {
                return "data 行数 " + data.Count + " 超过 rows=" + rows + "。";
            }
            for (int rowIndex = 0; rowIndex < data.Count; rowIndex++)
            {
                JArray rowData = data[rowIndex] as JArray;
                if (rowData == null)
                {
                    return "data[" + rowIndex + "] 必须是数组。";
                }
                if (rowData.Count > cols)
                {
                    return "data[" + rowIndex + "] 有 " + rowData.Count + " 列，超过 cols=" + cols + "。";
                }
            }
            return null;
        }

        private static string TableCellValue(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                return string.Empty;
            }
            return token.Type == JTokenType.String ? token.Value<string>() : token.ToString();
        }

        private static bool TryInsertTableCaption(Word.Table table, JArray warnings)
        {
            if (table == null)
            {
                return false;
            }
            Word.Range tableRange = null;
            try
            {
                tableRange = table.Range;
                string[] labels = new string[] { "表", "Table" };
                foreach (string label in labels)
                {
                    object labelObject = label;
                    object titleObject = string.Empty;
                    object autoText = Type.Missing;
                    object positionObject = Word.WdCaptionPosition.wdCaptionPositionAbove;
                    object excludeLabel = false;
                    try
                    {
                        tableRange.InsertCaption(
                            ref labelObject, ref titleObject, ref autoText,
                            ref positionObject, ref excludeLabel);
                        return true;
                    }
                    catch (Exception ex)
                    {
                        warnings.Add("题注标签「" + label + "」插入失败：" + ex.Message);
                    }
                }
                return false;
            }
            finally
            {
                ComHelper.Release(tableRange);
            }
        }

        private static List<double> ReadDoubleList(WordArguments args, string name)
        {
            JArray array = args.GetOptionalArray(name);
            if (array == null)
            {
                return null;
            }
            List<double> values = new List<double>();
            for (int index = 0; index < array.Count; index++)
            {
                try
                {
                    values.Add(array[index].Value<double>());
                }
                catch (Exception)
                {
                    throw WordArguments.Invalid(name + "[" + index + "] 必须是数字。");
                }
            }
            return values;
        }

        private static void ValidateCellFormatArguments(JObject format, out string error)
        {
            error = null;
            if (format == null)
            {
                return;
            }
            string[] supported = new string[]
            {
                "bold", "italic", "fontSize", "fontName", "color"
            };
            bool any = false;
            foreach (JProperty property in format.Properties())
            {
                bool found = false;
                foreach (string name in supported)
                {
                    if (string.Equals(name, property.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    error = "format 含不支持的键 " + property.Name +
                            "；支持 bold/italic/fontSize/fontName/color。收到 keys=" +
                            string.Join("/", new List<string>(GetPropertyNames(format)).ToArray());
                    return;
                }
                any = true;
            }
            if (!any)
            {
                error = "format 为空对象；请至少提供 bold/italic/fontSize/fontName/color 之一。";
            }
        }

        private static IEnumerable<string> GetPropertyNames(JObject obj)
        {
            foreach (JProperty property in obj.Properties())
            {
                yield return property.Name;
            }
        }

        private static void ApplyCellFormat(Word.Range range, JObject format, JArray verification)
        {
            if (range == null || format == null)
            {
                return;
            }
            Word.Font font = null;
            try
            {
                font = range.Font;
                JObject before = WordFormatUtil.CaptureFont(font);
                if (format["bold"] != null && format["bold"].Type != JTokenType.Null)
                {
                    font.Bold = format["bold"].Value<bool>() ? -1 : 0;
                }
                if (format["italic"] != null && format["italic"].Type != JTokenType.Null)
                {
                    font.Italic = format["italic"].Value<bool>() ? -1 : 0;
                }
                if (format["fontSize"] != null && format["fontSize"].Type != JTokenType.Null)
                {
                    font.Size = (float)format["fontSize"].Value<double>();
                }
                if (format["fontName"] != null && format["fontName"].Type != JTokenType.Null)
                {
                    string fontName = format["fontName"].ToString();
                    font.Name = fontName;
                    try
                    {
                        font.NameFarEast = fontName;
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[word] 单元格 NameFarEast 设置失败：" + ex.Message);
                    }
                }
                if (format["color"] != null && format["color"].Type != JTokenType.Null)
                {
                    font.Color = (Word.WdColor)WordFormatUtil.ParseColor(format["color"].ToString());
                }
                JObject after = WordFormatUtil.CaptureFont(font);
                VerifyCellFormat(verification, before, after, format);
            }
            finally
            {
                ComHelper.Release(font);
            }
        }

        private static void VerifyCellFormat(
            JArray verification,
            JObject before,
            JObject after,
            JObject format)
        {
            foreach (JProperty property in format.Properties())
            {
                string matchingKey = property.Name;
                if (string.Equals(matchingKey, "strikethrough", StringComparison.OrdinalIgnoreCase))
                {
                    matchingKey = "strikeThrough";
                }
                JToken actual = after[matchingKey];
                bool matched = true;
                if (property.Name == "bold" || property.Name == "italic")
                {
                    matched = actual != null && actual.Type == JTokenType.Boolean &&
                              actual.Value<bool>() == property.Value.Value<bool>();
                }
                else if (property.Name == "fontSize")
                {
                    matched = actual != null && actual.Type != JTokenType.Null &&
                              WordFormatUtil.CloseEnough(
                                  actual.Value<double>(), property.Value.Value<double>(), 0.2);
                }
                else if (property.Name == "fontName")
                {
                    string expected = property.Value.ToString();
                    string name = after.Value<string>("fontName");
                    string farEast = after.Value<string>("fontNameFarEast");
                    matched = string.Equals(name, expected, StringComparison.OrdinalIgnoreCase) ||
                              string.Equals(farEast, expected, StringComparison.OrdinalIgnoreCase);
                    actual = new JObject { { "fontName", name }, { "fontNameFarEast", farEast } };
                }
                else if (property.Name == "color")
                {
                    string expectedHex = WordFormatUtil.ColorToHex(
                        WordFormatUtil.ParseColor(property.Value.ToString()));
                    string actualHex = after.Value<string>("colorHex");
                    actual = new JValue(actualHex);
                    matched = string.Equals(expectedHex, actualHex, StringComparison.OrdinalIgnoreCase);
                }
                JObject item = new JObject();
                item["property"] = property.Name;
                item["expected"] = property.Value == null ? JValue.CreateNull() : property.Value.DeepClone();
                item["actual"] = actual == null ? JValue.CreateNull() : actual.DeepClone();
                item["matched"] = matched;
                verification.Add(item);
            }
        }

        private static bool AllVerificationsMatched(JArray verification)
        {
            if (verification == null)
            {
                return true;
            }
            foreach (JToken item in verification)
            {
                if (item["matched"] != null && !item["matched"].Value<bool>())
                {
                    return false;
                }
            }
            return true;
        }

        private static Word.WdLineStyle ParseBorderStyle(string raw)
        {
            switch ((raw ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "none":
                case "无":
                    return Word.WdLineStyle.wdLineStyleNone;
                case "single":
                case "单线":
                    return Word.WdLineStyle.wdLineStyleSingle;
                case "double":
                case "双线":
                    return Word.WdLineStyle.wdLineStyleDouble;
                case "dashed":
                case "虚线":
                    return Word.WdLineStyle.wdLineStyleDashLargeGap;
                case "dotted":
                case "点线":
                    return Word.WdLineStyle.wdLineStyleDot;
                default:
                    throw WordArguments.Invalid(
                        "borderStyle 只支持 none/single/double/dashed/dotted，收到 " + raw + "。");
            }
        }

        private static Word.WdLineWidth MapBorderWidth(double points)
        {
            double[] widths = new double[] { 0.25, 0.5, 0.75, 1.0, 1.5, 2.25, 3.0, 4.5, 6.0 };
            Word.WdLineWidth[] values = new Word.WdLineWidth[]
            {
                Word.WdLineWidth.wdLineWidth025pt,
                Word.WdLineWidth.wdLineWidth050pt,
                Word.WdLineWidth.wdLineWidth075pt,
                Word.WdLineWidth.wdLineWidth100pt,
                Word.WdLineWidth.wdLineWidth150pt,
                Word.WdLineWidth.wdLineWidth225pt,
                Word.WdLineWidth.wdLineWidth300pt,
                Word.WdLineWidth.wdLineWidth450pt,
                Word.WdLineWidth.wdLineWidth600pt
            };
            int nearest = 0;
            double distance = double.MaxValue;
            for (int index = 0; index < widths.Length; index++)
            {
                double current = Math.Abs(widths[index] - points);
                if (current < distance)
                {
                    distance = current;
                    nearest = index;
                }
            }
            return values[nearest];
        }

        private static JObject CaptureTableFormatState(Word.Table table)
        {
            JObject state = new JObject();
            Word.Borders borders = null;
            Word.Shading shading = null;
            Word.Rows rows = null;
            Word.Row firstRow = null;
            Word.Range tableRange = null;
            Word.Cells cells = null;
            Word.Columns columns = null;
            try
            {
                state["uniform"] = table.Uniform;
                state["allowAutoFit"] = table.AllowAutoFit;
                state["style"] = WordComUtil.ReadTableStyleName(table);
                borders = table.Borders;
                state["outsideLineStyle"] = borders.OutsideLineStyle.ToString();
                state["insideLineStyle"] = borders.InsideLineStyle.ToString();
                state["outsideColor"] = (int)borders.OutsideColor;
                state["insideColor"] = (int)borders.InsideColor;
                state["outsideLineWidth"] = (int)borders.OutsideLineWidth;
                state["insideLineWidth"] = (int)borders.InsideLineWidth;
                shading = table.Shading;
                state["shadingColor"] = (int)shading.BackgroundPatternColor;
                rows = table.Rows;
                state["rowAlignment"] = rows.Alignment.ToString();
                firstRow = rows[1];
                state["headingFormat"] = firstRow.HeadingFormat;
                tableRange = table.Range;
                cells = tableRange.Cells;
                state["cellVerticalAlignment"] = cells.VerticalAlignment.ToString();
                columns = table.Columns;
                JArray widths = new JArray();
                for (int index = 1; index <= columns.Count; index++)
                {
                    Word.Column column = null;
                    try
                    {
                        column = columns[index];
                        widths.Add((double)column.PreferredWidth);
                    }
                    finally
                    {
                        ComHelper.Release(column);
                    }
                }
                state["columnPreferredWidths"] = widths;
                return state;
            }
            finally
            {
                ComHelper.Release(columns, cells, tableRange, firstRow, rows, shading, borders);
            }
        }

        private static double? ReadNullableDouble(WordArguments args, string name)
        {
            return args.GetOptionalDouble(name);
        }
    }
}
