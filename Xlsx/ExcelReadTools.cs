using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>只读观察工具：上下文、范围、选区、工作表和画像。</summary>
    public sealed partial class ExcelToolExecutor
    {
        private static OperationResult ExecuteGetContext(Excel.Application application, JObject arguments)
        {
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Range usedRange = null;
            object selection = null;
            Excel.Range selectedRange = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                worksheet = ExcelComSupport.RequireActiveWorksheet(application);
                usedRange = worksheet.UsedRange;

                int rows = usedRange.Rows.Count;
                int columns = usedRange.Columns.Count;
                string[] headers = RangeResolver.GetHeaders(worksheet, usedRange, Math.Min(columns, 20));
                int formulaCount = CountSpecialCells(usedRange, Excel.XlCellType.xlCellTypeFormulas, Type.Missing);
                int errorCount = CountSpecialCells(usedRange, Excel.XlCellType.xlCellTypeFormulas, Excel.XlSpecialCellsValue.xlErrors)
                                 + CountSpecialCells(usedRange, Excel.XlCellType.xlCellTypeConstants, Excel.XlSpecialCellsValue.xlErrors);
                long nonBlank = CountNonBlank(application, usedRange);
                long total = (long)rows * columns;
                long blank = Math.Max(0, total - nonBlank);
                int duplicateRows;
                int sampledRows;
                int sampledColumns;
                CountDuplicateRows(usedRange, out duplicateRows, out sampledRows, out sampledColumns);

                JObject usedRangeInfo = new JObject();
                usedRangeInfo["address"] = usedRange.AddressLocal();
                usedRangeInfo["rows"] = rows;
                usedRangeInfo["columns"] = columns;
                usedRangeInfo["cells"] = total;

                JObject formulaInfo = new JObject();
                formulaInfo["formulaCells"] = formulaCount;
                formulaInfo["errorCells"] = errorCount;

                JObject quality = new JObject();
                quality["blankCells"] = blank;
                quality["nonBlankCells"] = nonBlank;
                quality["duplicateRows"] = duplicateRows;
                quality["sampledRows"] = sampledRows;
                quality["sampledColumns"] = sampledColumns;

                string recommended = ExcelComSupport.FullAddress(
                    worksheet.Name,
                    ExcelComSupport.ColumnLetter(usedRange.Column + columns + 2) +
                    usedRange.Row.ToString(CultureInfo.InvariantCulture));

                JObject data = new JObject();
                data["workbook"] = workbook.Name;
                data["workbookFullName"] = workbook.FullName;
                data["saved"] = workbook.Saved;
                data["activeSheet"] = worksheet.Name;
                data["usedRange"] = usedRangeInfo;
                data["headers"] = new JArray(headers);
                data["formulas"] = formulaInfo;
                data["quality"] = quality;
                data["recommendedOutput"] = recommended;
                data["sheets"] = BuildSheetSummaries(workbook);

                selection = application.Selection;
                selectedRange = selection as Excel.Range;
                if (selectedRange != null)
                {
                    JObject selectionInfo = new JObject();
                    selectionInfo["address"] = ExcelComSupport.FullAddress(
                        worksheet.Name,
                        selectedRange.AddressLocal());
                    selectionInfo["itemCount"] = Convert.ToInt64(selectedRange.CountLarge, CultureInfo.InvariantCulture);
                    selectionInfo["dataType"] = InferSelectionType(selectedRange);
                    data["selection"] = selectionInfo;
                }

                string contextRef = ExcelComSupport.FullAddress(
                    worksheet.Name, usedRange.AddressLocal());
                Observation observation = Observation.Create(
                    "read_context",
                    "已读取 Excel 上下文：工作簿「" + workbook.Name + "」，活动工作表「" + worksheet.Name + "」。",
                    false);
                observation.AddTargetRef(contextRef);
                return OperationResult.Succeed(
                    "已读取 Excel 上下文：工作簿「" + workbook.Name + "」，活动工作表「" + worksheet.Name + "」。",
                    data,
                    observation);
            }
            finally
            {
                ComHelper.Release(selectedRange, selection, usedRange, worksheet, workbook);
            }
        }

        private static OperationResult ExecuteReadRange(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange", "address");
            if (string.IsNullOrWhiteSpace(address))
            {
                address = "selection";
            }

            using (ExcelRangeContext context = RangeResolver.Resolve(application, address, false, "excel_read_range"))
            {
                int totalRows = context.Range.Rows.Count;
                int totalColumns = context.Range.Columns.Count;
                int maxRows = ExcelArgumentReader.GetInt(arguments, 100, "maxRows", "limitRows");
                int maxColumns = ExcelArgumentReader.GetInt(arguments, 26, "maxColumns", "limitColumns");
                maxRows = Math.Max(1, Math.Min(maxRows, 5000));
                maxColumns = Math.Max(1, Math.Min(maxColumns, 100));

                int startRow = ExcelArgumentReader.GetInt(arguments, 0, "startRow", "fromRow");
                if (startRow <= 0)
                {
                    string cursor = ExcelArgumentReader.GetTrimmedString(arguments, "cursor", "nextCursor");
                    if (cursor.Length > 0)
                    {
                        TryDecodeReadCursor(cursor, out startRow);
                    }
                }
                if (startRow <= 0)
                {
                    startRow = 1;
                }

                bool includeFormula = ExcelArgumentReader.GetBool(
                    arguments,
                    false,
                    "includeFormula",
                    "includeFormulas");
                bool includeFormat = ExcelArgumentReader.GetBool(
                    arguments,
                    false,
                    "includeFormat",
                    "includeFormats");

                JObject data = new JObject();
                data["sheet"] = context.SheetName;
                data["rows"] = totalRows;
                data["columns"] = totalColumns;
                data["totalRows"] = totalRows;
                data["totalEstimate"] = totalRows;
                data["startRow"] = startRow;
                data["maxRows"] = maxRows;
                data["maxColumns"] = maxColumns;

                if (startRow > totalRows)
                {
                    data["address"] = context.Range.AddressLocal();
                    data["endRow"] = JValue.CreateNull();
                    data["returnedRows"] = 0;
                    data["returnedColumns"] = 0;
                    data["truncated"] = false;
                    data["hasMoreRows"] = false;
                    data["nextCursor"] = JValue.CreateNull();
                    data["values"] = new JArray();
                    data["cells"] = new JArray();
                    if (includeFormula)
                    {
                        data["formulas"] = new JArray();
                    }
                    if (includeFormat)
                    {
                        data["numberFormats"] = new JArray();
                    }
                    data["note"] = "startRow 超出区域总行数，返回空页。";
                    Observation emptyObservation = Observation.Create(
                        "read_range",
                        "startRow 超出区域总行数，未读取单元格。",
                        false);
                    emptyObservation.AddTargetRef(
                        ExcelComSupport.FullAddress(context.SheetName, context.Range.AddressLocal()));
                    return OperationResult.Succeed("startRow 超出范围，返回空页。", data, emptyObservation);
                }

                int rowsToRead = Math.Min(maxRows, totalRows - startRow + 1);
                int columnsToRead = Math.Min(maxColumns, totalColumns);
                Excel.Range startCell = null;
                Excel.Range readRange = null;
                try
                {
                    startCell = (Excel.Range)context.Range.Cells[startRow, 1];
                    readRange = startCell.Resize[rowsToRead, columnsToRead];
                    object[,] values = ExcelComSupport.ReadMatrix(readRange, rowsToRead, columnsToRead);
                    object[,] formulas = ExcelComSupport.NormalizeMatrix(
                        readRange.Formula,
                        rowsToRead,
                        columnsToRead);

                    bool hasMoreRows = startRow + rowsToRead - 1 < totalRows;
                    bool columnsTruncated = columnsToRead < totalColumns;
                    data["address"] = readRange.AddressLocal();
                    data["endRow"] = startRow + rowsToRead - 1;
                    data["returnedRows"] = rowsToRead;
                    data["returnedColumns"] = columnsToRead;
                    data["truncated"] = hasMoreRows || columnsTruncated;
                    data["hasMoreRows"] = hasMoreRows;
                    data["columnsTruncated"] = columnsTruncated;
                    data["nextCursor"] = hasMoreRows
                        ? new JValue((startRow + rowsToRead).ToString(CultureInfo.InvariantCulture))
                        : JValue.CreateNull();
                    data["values"] = ExcelComSupport.MatrixToJArray(values);
                    data["cells"] = BuildCellTriplets(values, formulas, rowsToRead, columnsToRead);
                    if (includeFormula)
                    {
                        data["formulas"] = ExcelComSupport.MatrixToJArray(formulas);
                    }
                    if (includeFormat)
                    {
                        bool formatsComplete;
                        object[,] formats = ReadNumberFormatGrid(
                            readRange,
                            values,
                            rowsToRead,
                            columnsToRead,
                            out formatsComplete);
                        data["numberFormats"] = ExcelComSupport.MatrixToJArray(formats);
                        data["numberFormatsComplete"] = formatsComplete;
                    }

                    string readRef = ExcelComSupport.FullAddress(
                        context.SheetName, readRange.AddressLocal());
                    Observation observation = Observation.Create(
                        "read_range",
                        "已读取 " + readRef + "（第 " + startRow + " 行起，" +
                        rowsToRead + " 行 × " + columnsToRead + " 列）。",
                        false);
                    observation.AddTargetRef(readRef);
                    return OperationResult.Succeed(
                        "已读取 " + readRef + "。",
                        data,
                        observation);
                }
                finally
                {
                    ComHelper.Release(readRange, startCell);
                }
            }
        }

        /// <summary>构造单元格三元组矩阵：formula 只对真公式返回字符串，常量返回 null。</summary>
        private static JArray BuildCellTriplets(
            object[,] values,
            object[,] formulas,
            int rows,
            int columns)
        {
            JArray result = new JArray();
            for (int row = 0; row < rows; row++)
            {
                JArray rowArray = new JArray();
                for (int column = 0; column < columns; column++)
                {
                    object value = values == null ? null : values[row, column];
                    object rawFormula = formulas == null ? null : formulas[row, column];
                    string formulaText = rawFormula == null || rawFormula is DBNull
                        ? string.Empty
                        : Convert.ToString(rawFormula, CultureInfo.InvariantCulture) ?? string.Empty;
                    JObject cell = new JObject();
                    cell["formula"] = formulaText.Length > 0 && formulaText[0] == '='
                        ? new JValue(formulaText)
                        : JValue.CreateNull();
                    cell["value"] = ExcelComSupport.ToJToken(value);
                    string errorType = ExcelExtensionSupport.ErrorTextFromValue(value);
                    cell["errorType"] = errorType.Length > 0 ? new JValue(errorType) : JValue.CreateNull();
                    rowArray.Add(cell);
                }
                result.Add(rowArray);
            }
            return result;
        }

        /// <summary>游标兼容纯数字、row=N/startRow=N 与 Base64 JSON 三种形式。</summary>
        private static bool TryDecodeReadCursor(string cursor, out int startRow)
        {
            startRow = 0;
            if (string.IsNullOrWhiteSpace(cursor))
            {
                return false;
            }
            string text = cursor.Trim();
            int parsed;
            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) && parsed > 0)
            {
                startRow = parsed;
                return true;
            }
            string lower = text.ToLowerInvariant();
            int separator = lower.IndexOf(':');
            if (separator < 0)
            {
                separator = lower.IndexOf('=');
            }
            if (separator > 0)
            {
                string name = lower.Substring(0, separator);
                string value = text.Substring(separator + 1).Trim();
                if ((name == "row" || name == "startrow") &&
                    int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) &&
                    parsed > 0)
                {
                    startRow = parsed;
                    return true;
                }
            }
            try
            {
                byte[] bytes = Convert.FromBase64String(text);
                string json = System.Text.Encoding.UTF8.GetString(bytes);
                JObject root = JObject.Parse(json);
                int? row = root.Value<int?>("startRow");
                if (row.HasValue && row.Value > 0)
                {
                    startRow = row.Value;
                    return true;
                }
            }
            catch (Exception)
            {
            }
            return false;
        }

        private static OperationResult ExecuteGetSelection(Excel.Application application, JObject arguments)
        {
            object selected = application.Selection;
            Excel.Range range = selected as Excel.Range;
            if (range == null)
            {
                ComHelper.Release(selected);
                return Fail(OfficeErrorCode.INVALID_RANGE, "当前选择不是单元格区域。");
            }

            Excel.Worksheet worksheet = null;
            Excel.Range readRange = null;
            try
            {
                worksheet = range.Worksheet as Excel.Worksheet;
                int rows = range.Rows.Count;
                int columns = range.Columns.Count;
                int previewRows = Math.Min(rows, 8);
                int previewColumns = Math.Min(columns, 8);
                readRange = range.Resize[previewRows, previewColumns];
                object[,] values = ExcelComSupport.ReadMatrix(readRange, previewRows, previewColumns);

                JObject data = new JObject();
                data["address"] = ExcelComSupport.FullAddress(
                    worksheet == null ? string.Empty : worksheet.Name,
                    range.AddressLocal());
                data["rows"] = rows;
                data["columns"] = columns;
                data["itemCount"] = Convert.ToInt64(range.CountLarge, CultureInfo.InvariantCulture);
                data["dataType"] = InferSelectionType(range);
                data["preview"] = ExcelComSupport.MatrixToJArray(values);
                data["previewRows"] = previewRows;
                data["previewColumns"] = previewColumns;
                data["truncated"] = previewRows < rows || previewColumns < columns;

                return OperationResult.Succeed("已读取当前选区。", data);
            }
            finally
            {
                ComHelper.Release(readRange, worksheet, range, selected);
            }
        }

        private static OperationResult ExecuteListSheets(Excel.Application application, JObject arguments)
        {
            Excel.Workbook workbook = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                JArray sheets = BuildSheetSummaries(workbook);
                JObject data = new JObject();
                data["workbook"] = workbook.Name;
                data["count"] = sheets.Count;
                data["sheets"] = sheets;
                return OperationResult.Succeed(
                    "工作簿「" + workbook.Name + "」共 " + sheets.Count + " 个工作表。",
                    data);
            }
            finally
            {
                ComHelper.Release(workbook);
            }
        }

        private static OperationResult ExecuteRangeProfile(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange", "address");
            if (string.IsNullOrWhiteSpace(address))
            {
                address = "usedrange";
            }
            int maxRows = ExcelArgumentReader.GetInt(arguments, 30, "maxRows", "sampleRows");
            int maxColumns = ExcelArgumentReader.GetInt(arguments, 8, "maxColumns", "sampleColumns");
            maxRows = Math.Max(1, Math.Min(maxRows, 100));
            maxColumns = Math.Max(1, Math.Min(maxColumns, 26));

            using (ExcelRangeContext context = RangeResolver.Resolve(application, address, false, "excel_range_profile"))
            {
                JObject data = ExcelRangeProfileService.Build(
                    context.Worksheet,
                    context.Range,
                    maxRows,
                    maxColumns);
                data["workbook"] = context.Workbook.Name;
                return OperationResult.Succeed("范围画像已生成：" + context.FullAddress + "。", data);
            }
        }

        private static JArray BuildSheetSummaries(Excel.Workbook workbook)
        {
            JArray sheets = new JArray();
            Excel.Sheets collection = null;
            try
            {
                collection = workbook.Worksheets;
                int count = collection.Count;
                for (int index = 1; index <= count; index++)
                {
                    object item = collection[index];
                    Excel.Worksheet worksheet = item as Excel.Worksheet;
                    if (worksheet == null)
                    {
                        ComHelper.Release(item);
                        continue;
                    }
                    try
                    {
                        List<JObject> one = new List<JObject>();
                        ExcelComSupport.AddSheetSummary(worksheet, one);
                        if (one.Count > 0)
                        {
                            sheets.Add(one[0]);
                        }
                    }
                    finally
                    {
                        ComHelper.Release(worksheet);
                    }
                }
            }
            finally
            {
                ComHelper.Release(collection);
            }
            return sheets;
        }

        private static int CountSpecialCells(Excel.Range range, Excel.XlCellType type, object value)
        {
            if (range == null)
            {
                return 0;
            }
            Excel.Range special = null;
            try
            {
                special = range.SpecialCells(type, value);
                if (special == null)
                {
                    return 0;
                }
                return Convert.ToInt32(special.CountLarge, CultureInfo.InvariantCulture);
            }
            catch (COMException)
            {
                return 0;
            }
            finally
            {
                ComHelper.Release(special);
            }
        }

        private static long CountNonBlank(Excel.Application application, Excel.Range range)
        {
            Excel.WorksheetFunction functions = null;
            try
            {
                functions = application.WorksheetFunction;
                object result = functions.CountA(range);
                return Convert.ToInt64(result, CultureInfo.InvariantCulture);
            }
            finally
            {
                ComHelper.Release(functions);
            }
        }

        private static void CountDuplicateRows(
            Excel.Range usedRange,
            out int duplicateRows,
            out int sampledRows,
            out int sampledColumns)
        {
            duplicateRows = 0;
            int rows = usedRange.Rows.Count;
            int columns = usedRange.Columns.Count;
            sampledRows = Math.Min(rows, 200);
            sampledColumns = Math.Min(columns, 30);
            Excel.Range sample = null;
            try
            {
                sample = usedRange.Resize[sampledRows, sampledColumns];
                object[,] values = ExcelComSupport.ReadMatrix(sample, sampledRows, sampledColumns);
                HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int row = 0; row < sampledRows; row++)
                {
                    List<string> parts = new List<string>();
                    bool allEmpty = true;
                    for (int column = 0; column < sampledColumns; column++)
                    {
                        object value = values[row, column];
                        string text = value == null || value is DBNull
                            ? string.Empty
                            : Convert.ToString(value, CultureInfo.InvariantCulture);
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            allEmpty = false;
                        }
                        parts.Add(text);
                    }
                    if (allEmpty)
                    {
                        continue;
                    }
                    string key = string.Join("\u001f", parts.ToArray());
                    if (!seen.Add(key))
                    {
                        duplicateRows++;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[excel] 重复行统计失败：" + ex.Message);
            }
            finally
            {
                ComHelper.Release(sample);
            }
        }

        private static string InferSelectionType(Excel.Range range)
        {
            int rows = range.Rows.Count;
            int columns = range.Columns.Count;
            if (rows == 1 && columns == 1)
            {
                return "singleCell";
            }
            if (rows == 1)
            {
                return "row";
            }
            if (columns == 1)
            {
                return "column";
            }
            return "table";
        }
    }
}
