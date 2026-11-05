using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json.Linq;
using Office;

namespace Xlsx
{
    /// <summary>
    /// T09 P2：xlsx 只读 OpenXML 后端。不启动 Excel、不占用 worker，直接读文件；
    /// 用于 excel_profile_files / excel_query_files 的 backend=openxml 路径。
    /// 公式只返回公式文本，不计算公式结果；日期按单元格样式尽力转换。
    /// </summary>
    internal static class ExcelOpenXmlReadService
    {
        private sealed class SheetData
        {
            public string Name;
            public int FirstRow;
            public int FirstColumn;
            public int RowCount;
            public int ColumnCount;
            public object[,] Values;
            public string[,] Formulas;
        }

        private sealed class CellValue
        {
            public int Row;
            public int Column;
            public object Value;
            public string Formula;
        }

        private sealed class FileProfileResult
        {
            public int Index;
            public JArray Profiles = new JArray();
            public JArray Errors = new JArray();
            public int Processed;
        }

        private sealed class FileQueryResult
        {
            public int Index;
            public JArray Rows = new JArray();
            public JArray Errors = new JArray();
            public JArray Columns = new JArray();
            public int ScannedRows;
            public int MatchedRows;
            public bool Truncated;
            public bool Processed;
        }

        public static OperationResult ProfileFiles(JObject arguments)
        {
            List<string> files = ExcelArgumentReader.GetStringList(arguments, "files", "filePaths");
            if (files.Count == 0)
            {
                return ExcelToolExecutor.Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 files 数组。");
            }
            int maxFiles = Clamp(ExcelArgumentReader.GetInt(arguments, 10, "maxFiles"), 1, 50);
            bool filesTruncated = files.Count > maxFiles;
            int maxRows = Clamp(ExcelArgumentReader.GetInt(arguments, 30, "maxRows", "sampleRows"), 1, 100);
            int maxColumns = Clamp(ExcelArgumentReader.GetInt(arguments, 8, "maxColumns", "sampleColumns"), 1, 26);
            List<string> requestedSheets = ExcelArgumentReader.GetStringList(arguments, "sheets", "sheetNames");
            int count = Math.Min(files.Count, maxFiles);
            int concurrency = Math.Max(1, Math.Min(8, Office.Config.AppConfig.Current.ExcelMaxConcurrency));

            ConcurrentBag<FileProfileResult> bag = new ConcurrentBag<FileProfileResult>();
            ParallelOptions options = new ParallelOptions();
            options.MaxDegreeOfParallelism = concurrency;
            Parallel.For(0, count, options, delegate(int fileIndex)
            {
                FileProfileResult result = new FileProfileResult();
                result.Index = fileIndex;
                string requestedPath = files[fileIndex];
                string fullPath;
                if (!TryResolveFile(requestedPath, out fullPath))
                {
                    result.Errors.Add(FileError(requestedPath, "文件不存在或路径不可访问。"));
                    bag.Add(result);
                    return;
                }
                try
                {
                    List<SheetData> sheets = ReadWorkbook(fullPath, requestedSheets, result.Errors);
                    foreach (SheetData sheet in sheets)
                    {
                        JObject profile = BuildProfile(sheet, maxRows, maxColumns);
                        profile["file"] = fullPath;
                        profile["fileName"] = Path.GetFileName(fullPath);
                        result.Profiles.Add(profile);
                    }
                    result.Processed = 1;
                }
                catch (Exception ex)
                {
                    result.Errors.Add(FileError(fullPath, "OpenXML 读取失败：" + ex.Message));
                }
                bag.Add(result);
            });

            List<FileProfileResult> ordered = new List<FileProfileResult>(bag);
            ordered.Sort(delegate(FileProfileResult left, FileProfileResult right)
            {
                return left.Index.CompareTo(right.Index);
            });
            JArray profiles = new JArray();
            JArray errors = new JArray();
            int processed = 0;
            foreach (FileProfileResult result in ordered)
            {
                foreach (JToken profile in result.Profiles) { profiles.Add(profile.DeepClone()); }
                foreach (JToken error in result.Errors) { errors.Add(error.DeepClone()); }
                processed += result.Processed;
            }

            JObject data = new JObject();
            data["backend"] = "openxml";
            data["filesRequested"] = files.Count;
            data["filesProcessed"] = processed;
            data["filesTruncated"] = filesTruncated;
            data["parallelism"] = concurrency;
            data["readOnly"] = true;
            data["saved"] = false;
            data["profiles"] = profiles;
            data["errors"] = errors;

            if (errors.Count > 0)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.FILE_NOT_FOUND,
                    "OpenXML 画像部分失败：" + errors.Count + " 个错误。已返回成功读取的文件画像。",
                    null,
                    data,
                    "本次为只读 OpenXML 画像，未修改或保存任何源文件。",
                    true);
            }
            return OperationResult.Succeed(
                "OpenXML 只读画像 " + profiles.Count + " 个工作表，源文件未保存。",
                data);
        }

        public static OperationResult QueryFiles(JObject arguments)
        {
            List<string> files = ExcelArgumentReader.GetStringList(arguments, "files", "filePaths");
            if (files.Count == 0)
            {
                return ExcelToolExecutor.Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 files 数组。");
            }
            int maxFiles = Clamp(ExcelArgumentReader.GetInt(arguments, 10, "maxFiles"), 1, 50);
            bool filesTruncated = files.Count > maxFiles;
            int limit = Clamp(ExcelArgumentReader.GetInt(arguments, 1000, "limit"), 1, 5000);
            bool hasHeader = ExcelArgumentReader.GetBool(arguments, true, "hasHeader");
            bool caseSensitive = ExcelArgumentReader.GetBool(arguments, false, "caseSensitive");
            List<string> requestedSheets = ExcelArgumentReader.GetStringList(arguments, "sheets", "sheetNames");
            List<string> returnColumnNames = ExcelArgumentReader.GetStringList(arguments, "returnColumns", "returnFields");
            JArray filters = ExcelArgumentReader.GetArray(arguments, "filters");
            JObject lookup = ExcelArgumentReader.First(arguments, "lookup") as JObject;

            string keyColumnText = string.Empty;
            List<string> keys = new List<string>();
            string matchMode = "exact";
            if (lookup != null)
            {
                keyColumnText = ExcelArgumentReader.GetTrimmedString(lookup, "keyColumn", "key", "column");
                keys = ExcelArgumentReader.GetStringList(lookup, "keys", "values");
                string configuredMode = ExcelArgumentReader.GetTrimmedString(lookup, "matchMode", "mode").ToLowerInvariant();
                if (configuredMode.Length > 0) { matchMode = configuredMode; }
            }
            if (matchMode != "exact" && matchMode != "fuzzy" && matchMode != "regex")
            {
                return ExcelToolExecutor.Fail(OfficeErrorCode.INVALID_PARAMS, "lookup.matchMode 仅支持 exact/fuzzy/regex。");
            }
            if (keys.Count > 0 && keyColumnText.Length == 0)
            {
                return ExcelToolExecutor.Fail(OfficeErrorCode.INVALID_PARAMS, "lookup.keys 已提供，但缺少 lookup.keyColumn。");
            }
            string writeSetting = ExcelArgumentReader.GetString(arguments, "writeResultToSheet");
            if (writeSetting.Length > 0 && !string.Equals(writeSetting, "false", StringComparison.OrdinalIgnoreCase))
            {
                return ExcelToolExecutor.Fail(
                    OfficeErrorCode.INVALID_PARAMS,
                    "OpenXML 后端为纯只读，不支持 writeResultToSheet；请改用 backend=com。");
            }

            int count = Math.Min(files.Count, maxFiles);
            int concurrency = Math.Max(1, Math.Min(8, Office.Config.AppConfig.Current.ExcelMaxConcurrency));
            ConcurrentBag<FileQueryResult> bag = new ConcurrentBag<FileQueryResult>();
            ParallelOptions options = new ParallelOptions();
            options.MaxDegreeOfParallelism = concurrency;
            Parallel.For(0, count, options, delegate(int fileIndex)
            {
                FileQueryResult result = new FileQueryResult();
                result.Index = fileIndex;
                string requestedPath = files[fileIndex];
                string fullPath;
                if (!TryResolveFile(requestedPath, out fullPath))
                {
                    result.Errors.Add(FileError(requestedPath, "文件不存在或路径不可访问。"));
                    bag.Add(result);
                    return;
                }
                try
                {
                    List<SheetData> sheets = ReadWorkbook(fullPath, requestedSheets, result.Errors);
                    result.Processed = true;
                    QueryWorkbook(
                        sheets,
                        fullPath,
                        hasHeader,
                        caseSensitive,
                        keyColumnText,
                        keys,
                        matchMode,
                        filters,
                        returnColumnNames,
                        limit,
                        result);
                }
                catch (Exception ex)
                {
                    result.Errors.Add(FileError(fullPath, "OpenXML 查询失败：" + ex.Message));
                }
                bag.Add(result);
            });

            List<FileQueryResult> ordered = new List<FileQueryResult>(bag);
            ordered.Sort(delegate(FileQueryResult left, FileQueryResult right)
            {
                return left.Index.CompareTo(right.Index);
            });

            JArray rows = new JArray();
            JArray errors = new JArray();
            JArray resultColumns = new JArray();
            int scannedRows = 0;
            int matchedRows = 0;
            int filesProcessed = 0;
            bool truncated = false;
            foreach (FileQueryResult result in ordered)
            {
                foreach (JToken error in result.Errors) { errors.Add(error.DeepClone()); }
                scannedRows += result.ScannedRows;
                matchedRows += result.MatchedRows;
                if (result.Processed) { filesProcessed++; }
                if (result.Truncated) { truncated = true; }
                foreach (JToken column in result.Columns) { AddUniqueString(resultColumns, column.ToString()); }
                foreach (JToken row in result.Rows)
                {
                    if (rows.Count >= limit)
                    {
                        truncated = true;
                        break;
                    }
                    rows.Add(row.DeepClone());
                }
            }

            JObject data = new JObject();
            data["backend"] = "openxml";
            data["filesRequested"] = files.Count;
            data["filesProcessed"] = filesProcessed;
            data["filesTruncated"] = filesTruncated;
            data["limit"] = limit;
            data["returned"] = rows.Count;
            data["matched"] = matchedRows;
            data["scannedRows"] = scannedRows;
            data["truncated"] = truncated;
            data["parallelism"] = concurrency;
            data["readOnly"] = true;
            data["saved"] = false;
            data["columns"] = resultColumns;
            data["rows"] = rows;
            data["errors"] = errors;

            if (errors.Count > 0)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.FILE_NOT_FOUND,
                    "OpenXML 查询有 " + errors.Count + " 个文件/工作表错误；已返回可读取部分的结果。",
                    null,
                    data,
                    "本次以只读 OpenXML 读取源文件，未修改或保存。",
                    true);
            }
            return OperationResult.Succeed(
                "OpenXML 查询完成：扫描 " + scannedRows + " 行，匹配 " + matchedRows + " 行，返回 " + rows.Count + " 行。",
                data);
        }

        // ------------------------------------------------------------------
        // 工作簿读取
        // ------------------------------------------------------------------

        private static List<SheetData> ReadWorkbook(string fullPath, List<string> requestedSheets, JArray errors)
        {
            List<SheetData> result = new List<SheetData>();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fullPath, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                if (workbookPart == null || workbookPart.Workbook == null)
                {
                    throw new InvalidOperationException("xlsx 缺少 WorkbookPart。");
                }
                SharedStringTable sharedStrings = null;
                if (workbookPart.SharedStringTablePart != null)
                {
                    sharedStrings = workbookPart.SharedStringTablePart.SharedStringTable;
                }
                Stylesheet stylesheet = null;
                if (workbookPart.WorkbookStylesPart != null)
                {
                    stylesheet = workbookPart.WorkbookStylesPart.Stylesheet;
                }

                Sheets sheets = workbookPart.Workbook.Sheets;
                if (sheets == null)
                {
                    return result;
                }
                foreach (Sheet sheet in sheets.Elements<Sheet>())
                {
                    string sheetName = sheet.Name == null ? string.Empty : sheet.Name.Value;
                    if (requestedSheets != null && requestedSheets.Count > 0
                        && !ContainsIgnoreCase(requestedSheets, sheetName))
                    {
                        continue;
                    }
                    try
                    {
                        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
                        SheetData read = ReadSheet(sheetName, worksheetPart, sharedStrings, stylesheet);
                        if (read != null) { result.Add(read); }
                    }
                    catch (Exception ex)
                    {
                        errors.Add(FileError(fullPath + "!" + sheetName, "读取工作表失败：" + ex.Message));
                    }
                }
            }
            return result;
        }

        private static SheetData ReadSheet(
            string sheetName,
            WorksheetPart worksheetPart,
            SharedStringTable sharedStrings,
            Stylesheet stylesheet)
        {
            SheetData data = new SheetData();
            data.Name = sheetName;
            DocumentFormat.OpenXml.Spreadsheet.SheetData sheetData =
                worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
            if (sheetData == null)
            {
                data.FirstRow = 1;
                data.FirstColumn = 1;
                data.RowCount = 0;
                data.ColumnCount = 0;
                data.Values = new object[0, 0];
                data.Formulas = null;
                return data;
            }

            List<CellValue> cells = new List<CellValue>();
            int minRow = int.MaxValue;
            int minColumn = int.MaxValue;
            int maxRow = 0;
            int maxColumn = 0;
            int currentRow = 0;
            foreach (Row row in sheetData.Elements<Row>())
            {
                currentRow = row.RowIndex == null ? currentRow + 1 : (int)row.RowIndex.Value;
                int currentColumn = 0;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    currentColumn = cell.CellReference == null
                        ? currentColumn + 1
                        : ColumnIndexFromReference(cell.CellReference.Value);
                    object value = ReadCellValue(cell, sharedStrings, stylesheet);
                    string formula = cell.CellFormula == null ? string.Empty : "=" + cell.CellFormula.Text;
                    if (value == null && formula.Length == 0)
                    {
                        continue;
                    }
                    CellValue item = new CellValue();
                    item.Row = currentRow;
                    item.Column = currentColumn;
                    item.Value = value;
                    item.Formula = formula;
                    cells.Add(item);
                    if (currentRow < minRow) { minRow = currentRow; }
                    if (currentColumn < minColumn) { minColumn = currentColumn; }
                    if (currentRow > maxRow) { maxRow = currentRow; }
                    if (currentColumn > maxColumn) { maxColumn = currentColumn; }
                }
            }

            if (cells.Count == 0)
            {
                data.FirstRow = 1;
                data.FirstColumn = 1;
                data.RowCount = 0;
                data.ColumnCount = 0;
                data.Values = new object[0, 0];
                data.Formulas = null;
                return data;
            }

            data.FirstRow = minRow;
            data.FirstColumn = minColumn;
            data.RowCount = maxRow - minRow + 1;
            data.ColumnCount = maxColumn - minColumn + 1;
            data.Values = new object[data.RowCount, data.ColumnCount];
            data.Formulas = new string[data.RowCount, data.ColumnCount];
            foreach (CellValue cell in cells)
            {
                data.Values[cell.Row - minRow, cell.Column - minColumn] = cell.Value;
                data.Formulas[cell.Row - minRow, cell.Column - minColumn] = cell.Formula;
            }
            return data;
        }

        private static object ReadCellValue(Cell cell, SharedStringTable sharedStrings, Stylesheet stylesheet)
        {
            string raw = cell.CellValue == null ? string.Empty : (cell.CellValue.Text ?? string.Empty);
            CellValues dataType = cell.DataType == null ? CellValues.Number : cell.DataType.Value;
            if (dataType == CellValues.SharedString)
            {
                int index;
                if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out index)
                    && sharedStrings != null
                    && index >= 0
                    && index < sharedStrings.ChildElements.Count)
                {
                    return sharedStrings.ChildElements[index].InnerText;
                }
                return raw;
            }
            if (dataType == CellValues.InlineString)
            {
                return cell.InlineString == null ? raw : cell.InlineString.InnerText;
            }
            if (dataType == CellValues.Boolean)
            {
                return string.Equals(raw, "1", StringComparison.Ordinal);
            }
            if (dataType == CellValues.Date)
            {
                DateTime date;
                if (DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out date))
                {
                    return date;
                }
                return raw;
            }
            if (dataType == CellValues.Error || dataType == CellValues.String)
            {
                return raw;
            }

            double number;
            if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
            {
                if (IsDateStyle(stylesheet, cell.StyleIndex))
                {
                    try { return DateTime.FromOADate(number); }
                    catch { return number; }
                }
                return number;
            }
            return raw.Length == 0 ? null : raw;
        }

        private static bool IsDateStyle(Stylesheet stylesheet, UInt32Value styleIndex)
        {
            if (stylesheet == null || styleIndex == null || stylesheet.CellFormats == null)
            {
                return false;
            }
            CellFormats formats = stylesheet.CellFormats;
            if (formats.Count == null || styleIndex.Value >= formats.Count.Value)
            {
                return false;
            }
            CellFormat format = formats.ChildElements[(int)styleIndex.Value] as CellFormat;
            if (format == null || format.NumberFormatId == null)
            {
                return false;
            }
            uint id = format.NumberFormatId.Value;
            if ((id >= 14 && id <= 22) || (id >= 27 && id <= 36)
                || (id >= 45 && id <= 47) || (id >= 50 && id <= 58))
            {
                return true;
            }
            if (id >= 164 && stylesheet.NumberingFormats != null)
            {
                foreach (NumberingFormat numbering in stylesheet.NumberingFormats.Elements<NumberingFormat>())
                {
                    if (numbering.NumberFormatId != null
                        && numbering.NumberFormatId.Value == id
                        && numbering.FormatCode != null)
                    {
                        return Regex.IsMatch(numbering.FormatCode.Value, "[ymdhs]", RegexOptions.IgnoreCase);
                    }
                }
            }
            return false;
        }

        // ------------------------------------------------------------------
        // 画像
        // ------------------------------------------------------------------

        private static JObject BuildProfile(SheetData sheet, int maxRows, int maxColumns)
        {
            int totalRows = sheet.RowCount;
            int totalColumns = sheet.ColumnCount;
            int rowsToRead = Math.Min(totalRows, Math.Max(1, maxRows));
            int columnsToRead = Math.Min(totalColumns, Math.Max(1, maxColumns));
            string[] headers = new string[columnsToRead];
            for (int column = 0; column < columnsToRead; column++)
            {
                headers[column] = totalRows > 0
                    ? (Convert.ToString(sheet.Values[0, column], CultureInfo.InvariantCulture) ?? string.Empty)
                    : string.Empty;
            }
            bool hasHeader = rowsToRead > 1 && HasAnyText(headers);

            JArray samples = new JArray();
            for (int row = 0; row < rowsToRead; row++)
            {
                JArray sampleRow = new JArray();
                for (int column = 0; column < columnsToRead; column++)
                {
                    sampleRow.Add(ToJToken(sheet.Values[row, column]));
                }
                samples.Add(sampleRow);
            }

            JArray columnProfiles = new JArray();
            for (int column = 0; column < columnsToRead; column++)
            {
                int numeric = 0, date = 0, text = 0, empty = 0, formulaCount = 0;
                int firstDataRow = hasHeader ? 1 : 0;
                for (int row = firstDataRow; row < rowsToRead; row++)
                {
                    string formula = sheet.Formulas == null
                        ? string.Empty
                        : (sheet.Formulas[row, column] ?? string.Empty);
                    if (formula.StartsWith("=", StringComparison.Ordinal)) { formulaCount++; }
                    object value = sheet.Values[row, column];
                    if (value == null || (value is string && ((string)value).Length == 0))
                    {
                        empty++;
                    }
                    else if (value is DateTime)
                    {
                        date++;
                    }
                    else
                    {
                        double number;
                        if (IsNumeric(value, out number)) { numeric++; }
                        else { text++; }
                    }
                }
                string kind = "empty";
                if (numeric + date + text > 0)
                {
                    if (numeric >= text && numeric >= date) { kind = "numeric"; }
                    else if (date >= numeric && date >= text) { kind = "date"; }
                    else { kind = "text"; }
                }
                JObject profile = new JObject();
                profile["column"] = ExcelComSupport.ColumnLetter(sheet.FirstColumn + column);
                profile["header"] = headers[column];
                profile["type"] = formulaCount > 0 ? kind + "+formula" : kind;
                profile["numeric"] = numeric;
                profile["date"] = date;
                profile["text"] = text;
                profile["empty"] = empty;
                profile["formula"] = formulaCount;
                columnProfiles.Add(profile);
            }

            JObject result = new JObject();
            result["sheet"] = sheet.Name;
            result["address"] = RangeAddress(sheet, totalRows, totalColumns);
            result["totalRows"] = totalRows;
            result["totalColumns"] = totalColumns;
            result["sampledRows"] = rowsToRead;
            result["sampledColumns"] = columnsToRead;
            result["truncated"] = rowsToRead < totalRows || columnsToRead < totalColumns;
            result["hasHeader"] = hasHeader;
            result["headers"] = new JArray(headers);
            result["columns"] = columnProfiles;
            result["sample"] = samples;
            return result;
        }

        // ------------------------------------------------------------------
        // 查询
        // ------------------------------------------------------------------

        private static void QueryWorkbook(
            List<SheetData> sheets,
            string fullPath,
            bool hasHeader,
            bool caseSensitive,
            string keyColumnText,
            List<string> keys,
            string matchMode,
            JArray filters,
            List<string> returnColumnNames,
            int limit,
            FileQueryResult result)
        {
            foreach (SheetData sheet in sheets)
            {
                if (result.Rows.Count >= limit)
                {
                    result.Truncated = true;
                    break;
                }
                QuerySheet(
                    sheet,
                    fullPath,
                    hasHeader,
                    caseSensitive,
                    keyColumnText,
                    keys,
                    matchMode,
                    filters,
                    returnColumnNames,
                    limit,
                    result);
            }
        }

        private static void QuerySheet(
            SheetData sheet,
            string fullPath,
            bool hasHeader,
            bool caseSensitive,
            string keyColumnText,
            List<string> keys,
            string matchMode,
            JArray filters,
            List<string> returnColumnNames,
            int limit,
            FileQueryResult result)
        {
            int totalRows = sheet.RowCount;
            int totalColumns = sheet.ColumnCount;
            if (totalRows < 1 || totalColumns < 1) { return; }

            string[] rawHeaders = new string[totalColumns];
            for (int column = 0; column < totalColumns; column++)
            {
                rawHeaders[column] = hasHeader
                    ? (Convert.ToString(sheet.Values[0, column], CultureInfo.InvariantCulture) ?? string.Empty)
                    : string.Empty;
            }
            string[] headers = BuildUniqueHeaders(rawHeaders, totalColumns);

            int keyColumn = 0;
            if (!string.IsNullOrWhiteSpace(keyColumnText))
            {
                keyColumn = ResolveColumnRef(keyColumnText, headers, hasHeader, totalColumns);
                if (keyColumn <= 0)
                {
                    result.Errors.Add(FileError(fullPath + "!" + sheet.Name,
                        "lookup.keyColumn「" + keyColumnText + "」不存在。"));
                    return;
                }
            }

            int[] selectedColumns = ResolveReturnColumns(
                returnColumnNames, headers, hasHeader, totalColumns,
                fullPath, sheet.Name, result.Errors, result.Columns);
            if (selectedColumns == null) { return; }

            List<FilterSpec> specs = BuildFilters(
                filters, headers, hasHeader, totalColumns,
                fullPath, sheet.Name, result.Errors);
            if (specs == null) { return; }

            int dataStart = hasHeader ? 1 : 0;
            for (int row = dataStart; row < totalRows; row++)
            {
                result.ScannedRows++;
                if (!PassesFilters(sheet.Values, row, specs, caseSensitive)) { continue; }
                if (keyColumn > 0 && keys.Count > 0
                    && !MatchesLookup(sheet.Values[row, keyColumn - 1], keys, matchMode, caseSensitive))
                {
                    continue;
                }
                result.MatchedRows++;
                if (result.Rows.Count >= limit)
                {
                    result.Truncated = true;
                    break;
                }
                JObject item = new JObject();
                for (int selected = 0; selected < selectedColumns.Length; selected++)
                {
                    int column = selectedColumns[selected];
                    item[headers[column - 1]] = ToJToken(sheet.Values[row, column - 1]);
                }
                item["file"] = fullPath;
                item["sheet"] = sheet.Name;
                item["rowNumber"] = sheet.FirstRow + row;
                result.Rows.Add(item);
            }
        }

        private static bool PassesFilters(object[,] values, int row, List<FilterSpec> filters, bool caseSensitive)
        {
            foreach (FilterSpec filter in filters)
            {
                if (!MatchesFilter(values[row, filter.Column - 1], filter, caseSensitive)) { return false; }
            }
            return true;
        }

        private static bool MatchesFilter(object cellValue, FilterSpec filter, bool caseSensitive)
        {
            string text = Convert.ToString(cellValue, CultureInfo.InvariantCulture) ?? string.Empty;
            string value = filter.Value ?? string.Empty;
            StringComparison comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            switch (filter.Operator)
            {
                case "contains": return text.IndexOf(value, comparison) >= 0;
                case "startswith": return text.StartsWith(value, comparison);
                case "regex":
                    return Regex.IsMatch(text, value, caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase);
                case "=":
                case "==": return CompareEquality(text, value, comparison);
                case "!=":
                case "<>": return !CompareEquality(text, value, comparison);
                default: return CompareNumeric(text, value, filter.Operator);
            }
        }

        private static bool CompareEquality(string left, string right, StringComparison comparison)
        {
            double leftNumber, rightNumber;
            if (TryParseNumber(left, out leftNumber) && TryParseNumber(right, out rightNumber))
            {
                return Math.Abs(leftNumber - rightNumber) < 0.0000000001;
            }
            return string.Equals(left, right, comparison);
        }

        private static bool CompareNumeric(string left, string right, string op)
        {
            double leftNumber, rightNumber;
            if (!TryParseNumber(left, out leftNumber) || !TryParseNumber(right, out rightNumber)) { return false; }
            switch (op)
            {
                case ">": return leftNumber > rightNumber;
                case "<": return leftNumber < rightNumber;
                case ">=": return leftNumber >= rightNumber;
                case "<=": return leftNumber <= rightNumber;
                default: return false;
            }
        }

        private static bool MatchesLookup(object cellValue, List<string> keys, string matchMode, bool caseSensitive)
        {
            string text = Convert.ToString(cellValue, CultureInfo.InvariantCulture) ?? string.Empty;
            StringComparison comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            foreach (string key in keys)
            {
                if (matchMode == "exact")
                {
                    if (string.Equals(text.Trim(), key.Trim(), comparison)) { return true; }
                }
                else if (matchMode == "fuzzy")
                {
                    if (text.IndexOf(key, comparison) >= 0) { return true; }
                }
                else if (matchMode == "regex")
                {
                    if (Regex.IsMatch(text, key, caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase)) { return true; }
                }
            }
            return false;
        }

        private static int[] ResolveReturnColumns(
            List<string> names, string[] headers, bool hasHeader, int totalColumns,
            string file, string sheet, JArray errors, JArray resultColumns)
        {
            if (names == null || names.Count == 0)
            {
                int[] all = new int[totalColumns];
                for (int index = 0; index < totalColumns; index++)
                {
                    all[index] = index + 1;
                    AddUniqueString(resultColumns, headers[index]);
                }
                return all;
            }
            List<int> selected = new List<int>();
            foreach (string name in names)
            {
                int column = ResolveColumnRef(name, headers, hasHeader, totalColumns);
                if (column <= 0)
                {
                    errors.Add(FileError(file + "!" + sheet, "returnColumns 中的「" + name + "」不存在。"));
                    return null;
                }
                if (!selected.Contains(column))
                {
                    selected.Add(column);
                    AddUniqueString(resultColumns, headers[column - 1]);
                }
            }
            return selected.ToArray();
        }

        private static List<FilterSpec> BuildFilters(
            JArray filters, string[] headers, bool hasHeader, int totalColumns,
            string file, string sheet, JArray errors)
        {
            List<FilterSpec> specs = new List<FilterSpec>();
            if (filters == null) { return specs; }
            foreach (JToken token in filters)
            {
                JObject filter = token as JObject;
                if (filter == null) { continue; }
                string columnText = ExcelArgumentReader.GetTrimmedString(filter, "column", "field", "col");
                string op = ExcelArgumentReader.GetTrimmedString(filter, "op", "operator").ToLowerInvariant();
                JToken valueToken = ExcelArgumentReader.First(filter, "value", "criteria");
                if (columnText.Length == 0 || op.Length == 0 || valueToken == null)
                {
                    errors.Add(FileError(file + "!" + sheet, "filters 项需要 column/op/value 三个字段。"));
                    return null;
                }
                int column = ResolveColumnRef(columnText, headers, hasHeader, totalColumns);
                if (column <= 0)
                {
                    errors.Add(FileError(file + "!" + sheet, "filters 列「" + columnText + "」不存在。"));
                    return null;
                }
                if (!IsSupportedFilterOperator(op))
                {
                    errors.Add(FileError(file + "!" + sheet, "不支持的 filter op：「" + op + "」。"));
                    return null;
                }
                FilterSpec spec = new FilterSpec();
                spec.Column = column;
                spec.Operator = op;
                spec.Value = valueToken.ToString();
                specs.Add(spec);
            }
            return specs;
        }

        private static bool IsSupportedFilterOperator(string op)
        {
            switch (op)
            {
                case "=":
                case "==":
                case "!=":
                case "<>":
                case ">":
                case "<":
                case ">=":
                case "<=":
                case "contains":
                case "startswith":
                case "regex":
                    return true;
                default:
                    return false;
            }
        }

        private static int ResolveColumnRef(string name, string[] headers, bool hasHeader, int totalColumns)
        {
            int numeric;
            if (int.TryParse(name, NumberStyles.Integer, CultureInfo.InvariantCulture, out numeric))
            {
                return numeric >= 1 && numeric <= totalColumns ? numeric : 0;
            }
            if (hasHeader)
            {
                for (int index = 0; index < headers.Length; index++)
                {
                    if (string.Equals(headers[index], name, StringComparison.OrdinalIgnoreCase)) { return index + 1; }
                }
            }
            int column = ExcelComSupport.ColumnIndex(name);
            return column >= 1 && column <= totalColumns ? column : 0;
        }

        private static string[] BuildUniqueHeaders(string[] rawHeaders, int totalColumns)
        {
            string[] headers = new string[totalColumns];
            HashSet<string> used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < totalColumns; index++)
            {
                string baseName = rawHeaders[index] == null ? string.Empty : rawHeaders[index].Trim();
                if (baseName.Length == 0) { baseName = "Column" + ExcelComSupport.ColumnLetter(index + 1); }
                string name = baseName;
                int suffix = 2;
                while (!used.Add(name))
                {
                    name = baseName + "_" + suffix.ToString(CultureInfo.InvariantCulture);
                    suffix++;
                }
                headers[index] = name;
            }
            return headers;
        }

        // ------------------------------------------------------------------
        // 辅助
        // ------------------------------------------------------------------

        private static int ColumnIndexFromReference(string reference)
        {
            if (string.IsNullOrWhiteSpace(reference)) { return 1; }
            int result = 0;
            for (int index = 0; index < reference.Length; index++)
            {
                char character = reference[index];
                if (character == '$') { continue; }
                if (character >= 'A' && character <= 'Z')
                {
                    result = result * 26 + (character - 'A' + 1);
                }
                else if (character >= 'a' && character <= 'z')
                {
                    result = result * 26 + (character - 'a' + 1);
                }
                else
                {
                    break;
                }
            }
            return result <= 0 ? 1 : result;
        }

        private static string RangeAddress(SheetData sheet, int totalRows, int totalColumns)
        {
            if (totalRows < 1 || totalColumns < 1) { return "A1"; }
            string start = ExcelComSupport.ColumnLetter(sheet.FirstColumn) +
                           sheet.FirstRow.ToString(CultureInfo.InvariantCulture);
            string end = ExcelComSupport.ColumnLetter(sheet.FirstColumn + totalColumns - 1) +
                         (sheet.FirstRow + totalRows - 1).ToString(CultureInfo.InvariantCulture);
            return start + ":" + end;
        }

        private static bool HasAnyText(string[] values)
        {
            if (values == null) { return false; }
            foreach (string value in values)
            {
                if (!string.IsNullOrWhiteSpace(value)) { return true; }
            }
            return false;
        }

        private static bool IsNumeric(object value, out double number)
        {
            number = 0;
            if (value == null || value is DateTime) { return false; }
            if (value is double || value is float || value is decimal || value is int || value is long)
            {
                number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                return true;
            }
            return TryParseNumber(Convert.ToString(value, CultureInfo.InvariantCulture), out number);
        }

        private static bool TryParseNumber(string text, out double number)
        {
            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) { return true; }
            return double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out number);
        }

        private static JToken ToJToken(object value)
        {
            if (value == null) { return JValue.CreateNull(); }
            if (value is DateTime) { return new JValue(((DateTime)value).ToString("o", CultureInfo.InvariantCulture)); }
            if (value is bool) { return new JValue((bool)value); }
            if (value is double) { return new JValue((double)value); }
            if (value is int) { return new JValue((int)value); }
            if (value is long) { return new JValue((long)value); }
            return new JValue(Convert.ToString(value, CultureInfo.InvariantCulture));
        }

        private static bool TryResolveFile(string requestedPath, out string fullPath)
        {
            fullPath = string.Empty;
            if (string.IsNullOrWhiteSpace(requestedPath)) { return false; }
            string pathError;
            if (!Office.Config.WorkspacePathResolver.TryResolveDocument(
                    "excel", requestedPath, out fullPath, out pathError))
            {
                return false;
            }
            return File.Exists(fullPath);
        }

        private static bool ContainsIgnoreCase(List<string> values, string value)
        {
            foreach (string item in values)
            {
                if (string.Equals(item, value, StringComparison.OrdinalIgnoreCase)) { return true; }
            }
            return false;
        }

        private static void AddUniqueString(JArray array, string value)
        {
            if (array == null || string.IsNullOrWhiteSpace(value)) { return; }
            foreach (JToken item in array)
            {
                if (string.Equals(item.ToString(), value, StringComparison.OrdinalIgnoreCase)) { return; }
            }
            array.Add(value);
        }

        private static JObject FileError(string file, string message)
        {
            JObject error = new JObject();
            error["file"] = file ?? string.Empty;
            error["error"] = message ?? string.Empty;
            return error;
        }

        private static int Clamp(int value, int min, int max)
        {
            if (value < min) { return min; }
            if (value > max) { return max; }
            return value;
        }

        private sealed class FilterSpec
        {
            public int Column { get; set; }
            public string Operator { get; set; }
            public string Value { get; set; }
        }
    }
}
