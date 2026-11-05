using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Office;
using Office.Contracts;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>
    /// 多文件只读画像/查询。使用专用隐藏 Excel 实例打开源文件：
    /// 不注册进 OfficeSession、不触碰当前会话、不保存源文件；
    /// UsedRange 一次读入内存后完成筛选/匹配。
    /// </summary>
    internal static class ExcelExternalWorkbookService
    {
        public static OperationResult ProfileFiles(JObject arguments)
        {
            string backend = ExcelArgumentReader.GetTrimmedString(arguments, "backend", "readBackend");
            if (string.Equals(backend, "openxml", StringComparison.OrdinalIgnoreCase))
            {
                return ExcelOpenXmlReadService.ProfileFiles(arguments);
            }
            ComHelper.EnsureSta("excel_profile_files");
            List<string> files = ExcelArgumentReader.GetStringList(arguments, "files", "filePaths");
            if (files.Count == 0)
            {
                return ExcelToolExecutor.Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 files 数组。");
            }

            int maxFiles = ExcelArgumentReader.GetInt(arguments, 10, "maxFiles");
            maxFiles = Math.Max(1, Math.Min(maxFiles, 50));
            bool filesTruncated = files.Count > maxFiles;
            int maxRows = ExcelArgumentReader.GetInt(arguments, 30, "maxRows", "sampleRows");
            int maxColumns = ExcelArgumentReader.GetInt(arguments, 8, "maxColumns", "sampleColumns");
            maxRows = Math.Max(1, Math.Min(maxRows, 100));
            maxColumns = Math.Max(1, Math.Min(maxColumns, 26));
            List<string> requestedSheets = ExcelArgumentReader.GetStringList(arguments, "sheets", "sheetNames");

            JArray profiles = new JArray();
            JArray errors = new JArray();
            int processed = 0;
            using (HiddenExcelScope scope = new HiddenExcelScope())
            {
                int count = Math.Min(files.Count, maxFiles);
                for (int fileIndex = 0; fileIndex < count; fileIndex++)
                {
                    string requestedPath = files[fileIndex];
                    string fullPath;
                    if (!TryResolveFile(requestedPath, out fullPath))
                    {
                        errors.Add(FileError(requestedPath, "文件不存在或路径不可访问。"));
                        continue;
                    }

                    Excel.Workbook workbook = null;
                    try
                    {
                        workbook = OpenReadOnly(scope, fullPath);
                        processed++;
                        List<Excel.Worksheet> sheets = GetTargetSheets(workbook, requestedSheets, fullPath, errors);
                        try
                        {
                            foreach (Excel.Worksheet worksheet in sheets)
                            {
                                Excel.Range usedRange = null;
                                try
                                {
                                    usedRange = worksheet.UsedRange;
                                    JObject profile = ExcelRangeProfileService.Build(
                                        worksheet,
                                        usedRange,
                                        maxRows,
                                        maxColumns);
                                    profile["file"] = fullPath;
                                    profile["fileName"] = Path.GetFileName(fullPath);
                                    profiles.Add(profile);
                                }
                                finally
                                {
                                    ComHelper.Release(usedRange);
                                }
                            }
                        }
                        finally
                        {
                            foreach (Excel.Worksheet worksheet in sheets)
                            {
                                ComHelper.Release(worksheet);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        errors.Add(FileError(fullPath, OfficeErrorMapper.DescribeException(ex, OfficeApp.Excel)));
                    }
                    finally
                    {
                        CloseWorkbook(workbook);
                    }
                }
            }

            JObject data = new JObject();
            data["filesRequested"] = files.Count;
            data["filesProcessed"] = processed;
            data["filesTruncated"] = filesTruncated;
            data["readOnly"] = true;
            data["saved"] = false;
            data["profiles"] = profiles;
            data["errors"] = errors;

            if (errors.Count > 0)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.FILE_NOT_FOUND,
                    "部分文件无法只读打开：" + errors.Count + " 个错误。已返回成功读取的文件画像。",
                    null,
                    data,
                    "本次为只读画像，未修改或保存任何源文件。",
                    true);
            }

            return OperationResult.Succeed(
                "已只读画像 " + profiles.Count + " 个工作表，源文件未保存。",
                data);
        }

        public static OperationResult QueryFiles(JObject arguments)
        {
            string backend = ExcelArgumentReader.GetTrimmedString(arguments, "backend", "readBackend");
            if (string.Equals(backend, "openxml", StringComparison.OrdinalIgnoreCase))
            {
                return ExcelOpenXmlReadService.QueryFiles(arguments);
            }
            ComHelper.EnsureSta("excel_query_files");
            List<string> files = ExcelArgumentReader.GetStringList(arguments, "files", "filePaths");
            if (files.Count == 0)
            {
                return ExcelToolExecutor.Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 files 数组。");
            }

            int maxFiles = ExcelArgumentReader.GetInt(arguments, 10, "maxFiles");
            maxFiles = Math.Max(1, Math.Min(maxFiles, 50));
            bool filesTruncated = files.Count > maxFiles;
            int limit = ExcelArgumentReader.GetInt(arguments, 1000, "limit");
            limit = Math.Max(1, Math.Min(limit, 5000));
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
                if (configuredMode.Length > 0)
                {
                    matchMode = configuredMode;
                }
            }
            if (matchMode != "exact" && matchMode != "fuzzy" && matchMode != "regex")
            {
                return ExcelToolExecutor.Fail(
                    OfficeErrorCode.INVALID_PARAMS,
                    "lookup.matchMode 仅支持 exact/fuzzy/regex。");
            }
            if (keys.Count > 0 && keyColumnText.Length == 0)
            {
                return ExcelToolExecutor.Fail(
                    OfficeErrorCode.INVALID_PARAMS,
                    "lookup.keys 已提供，但缺少 lookup.keyColumn。");
            }

            JArray rows = new JArray();
            JArray errors = new JArray();
            JArray resultColumns = new JArray();
            int scannedRows = 0;
            int matchedRows = 0;
            int filesProcessed = 0;
            bool truncated = false;

            using (HiddenExcelScope scope = new HiddenExcelScope())
            {
                int count = Math.Min(files.Count, maxFiles);
                for (int fileIndex = 0; fileIndex < count && !truncated; fileIndex++)
                {
                    string requestedPath = files[fileIndex];
                    string fullPath;
                    if (!TryResolveFile(requestedPath, out fullPath))
                    {
                        errors.Add(FileError(requestedPath, "文件不存在或路径不可访问。"));
                        continue;
                    }

                    Excel.Workbook workbook = null;
                    try
                    {
                        workbook = OpenReadOnly(scope, fullPath);
                        filesProcessed++;
                        List<Excel.Worksheet> sheets = GetTargetSheets(workbook, requestedSheets, fullPath, errors);
                        try
                        {
                            foreach (Excel.Worksheet worksheet in sheets)
                            {
                                if (truncated)
                                {
                                    break;
                                }
                                QueryWorksheet(
                                    worksheet,
                                    fullPath,
                                    hasHeader,
                                    caseSensitive,
                                    keyColumnText,
                                    keys,
                                    matchMode,
                                    filters,
                                    returnColumnNames,
                                    limit,
                                    rows,
                                    resultColumns,
                                    ref scannedRows,
                                    ref matchedRows,
                                    ref truncated,
                                    errors);
                            }
                        }
                        finally
                        {
                            foreach (Excel.Worksheet worksheet in sheets)
                            {
                                ComHelper.Release(worksheet);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        errors.Add(FileError(fullPath, OfficeErrorMapper.DescribeException(ex, OfficeApp.Excel)));
                    }
                    finally
                    {
                        CloseWorkbook(workbook);
                    }
                }
            }

            JObject data = new JObject();
            data["filesRequested"] = files.Count;
            data["filesProcessed"] = filesProcessed;
            data["filesTruncated"] = filesTruncated;
            data["limit"] = limit;
            data["returned"] = rows.Count;
            data["matched"] = matchedRows;
            data["scannedRows"] = scannedRows;
            data["truncated"] = truncated;
            data["readOnly"] = true;
            data["saved"] = false;
            data["columns"] = resultColumns;
            data["rows"] = rows;
            data["errors"] = errors;

            if (errors.Count > 0)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.FILE_NOT_FOUND,
                    "查询过程中有 " + errors.Count + " 个文件/工作表错误；已返回可读取部分的结果。",
                    null,
                    data,
                    "本次查询以只读方式打开源文件，未修改或保存；源文件 SHA256 不应变化。",
                    true);
            }

            string writeSetting = ExcelArgumentReader.GetString(arguments, "writeResultToSheet");
            bool writeRequested = ExcelArgumentReader.GetBool(arguments, false, "writeResultToSheet");
            if (!writeRequested && writeSetting.Length > 0 &&
                !string.Equals(writeSetting, "false", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(writeSetting, "0", StringComparison.OrdinalIgnoreCase))
            {
                writeRequested = true;
            }

            if (!writeRequested)
            {
                return OperationResult.Succeed(
                    "查询完成：扫描 " + scannedRows + " 行，匹配 " + matchedRows + " 行，返回 " + rows.Count + " 行。",
                    data);
            }

            string writeTarget = writeSetting.Length > 0 &&
                                 !string.Equals(writeSetting, "true", StringComparison.OrdinalIgnoreCase) &&
                                 !string.Equals(writeSetting, "1", StringComparison.OrdinalIgnoreCase)
                ? writeSetting
                : string.Empty;
            OperationResult writeResult = OfficeSession.RunExcel(
                "excel_query_files:writeResult",
                delegate(Excel.Application application)
                {
                    return WriteQueryResultToCurrentWorkbook(application, data, writeTarget);
                });
            if (!writeResult.Success)
            {
                OfficeErrorCode writeCode;
                if (!OfficeErrorMapper.TryParseCode(writeResult.ErrorCode, out writeCode))
                {
                    writeCode = OfficeErrorCode.INTERNAL_ERROR;
                }
                return OperationResult.Failed(
                    writeCode,
                    "查询已完成，但写入当前工作簿失败：" + writeResult.UserMessage,
                    null,
                    data,
                    "查询源文件未修改；写入失败仅影响当前工作簿结果区域。",
                    writeResult.Recoverable);
            }

            data["writeResult"] = writeResult.Data;
            Observation queryObservation = writeResult.Observation;
            return OperationResult.Succeed(
                "查询完成：返回 " + rows.Count + " 行，并写入当前活动工作簿。",
                data,
                queryObservation);
        }

        private static void QueryWorksheet(
            Excel.Worksheet worksheet,
            string fullPath,
            bool hasHeader,
            bool caseSensitive,
            string keyColumnText,
            List<string> keys,
            string matchMode,
            JArray filters,
            List<string> returnColumnNames,
            int limit,
            JArray rows,
            JArray resultColumns,
            ref int scannedRows,
            ref int matchedRows,
            ref bool truncated,
            JArray errors)
        {
            Excel.Range usedRange = null;
            try
            {
                usedRange = worksheet.UsedRange;
                int totalRows = usedRange.Rows.Count;
                int totalColumns = usedRange.Columns.Count;
                if (totalRows < 1 || totalColumns < 1)
                {
                    return;
                }

                object[,] values = ExcelComSupport.ReadMatrix(usedRange, totalRows, totalColumns);
                string[] rawHeaders = new string[totalColumns];
                for (int column = 0; column < totalColumns; column++)
                {
                    rawHeaders[column] = hasHeader
                        ? Convert.ToString(values[0, column], CultureInfo.InvariantCulture) ?? string.Empty
                        : string.Empty;
                }
                string[] headers = BuildUniqueHeaders(rawHeaders, totalColumns);

                int keyColumn = 0;
                if (!string.IsNullOrWhiteSpace(keyColumnText))
                {
                    keyColumn = ResolveColumnRef(keyColumnText, headers, hasHeader, totalColumns);
                    if (keyColumn <= 0)
                    {
                        errors.Add(FileError(
                            fullPath + "!" + worksheet.Name,
                            "lookup.keyColumn「" + keyColumnText + "」不存在。"));
                        return;
                    }
                }

                int[] selectedColumns = ResolveReturnColumns(
                    returnColumnNames,
                    headers,
                    hasHeader,
                    totalColumns,
                    fullPath,
                    worksheet.Name,
                    errors,
                    resultColumns);
                if (selectedColumns == null)
                {
                    return;
                }

                List<FilterSpec> filterSpecs = BuildFilters(
                    filters,
                    headers,
                    hasHeader,
                    totalColumns,
                    fullPath,
                    worksheet.Name,
                    errors);
                if (filterSpecs == null)
                {
                    return;
                }

                int dataStart = hasHeader ? 1 : 0;
                for (int row = dataStart; row < totalRows; row++)
                {
                    scannedRows++;
                    if (!PassesFilters(values, row, filterSpecs, caseSensitive))
                    {
                        continue;
                    }
                    if (keyColumn > 0 && keys.Count > 0 &&
                        !MatchesLookup(values[row, keyColumn - 1], keys, matchMode, caseSensitive))
                    {
                        continue;
                    }

                    matchedRows++;
                    if (rows.Count >= limit)
                    {
                        truncated = true;
                        break;
                    }

                    JObject result = new JObject();
                    for (int selectedIndex = 0; selectedIndex < selectedColumns.Length; selectedIndex++)
                    {
                        int column = selectedColumns[selectedIndex];
                        string columnName = headers[column - 1];
                        result[columnName] = ExcelComSupport.ToJToken(values[row, column - 1]);
                    }
                    result["file"] = fullPath;
                    result["sheet"] = worksheet.Name;
                    result["rowNumber"] = usedRange.Row + row;
                    rows.Add(result);
                }
            }
            finally
            {
                ComHelper.Release(usedRange);
            }
        }

        private static bool PassesFilters(
            object[,] values,
            int row,
            List<FilterSpec> filters,
            bool caseSensitive)
        {
            foreach (FilterSpec filter in filters)
            {
                object cellValue = values[row, filter.Column - 1];
                if (!MatchesFilter(cellValue, filter, caseSensitive))
                {
                    return false;
                }
            }
            return true;
        }

        private static bool MatchesFilter(object cellValue, FilterSpec filter, bool caseSensitive)
        {
            string cellText = Convert.ToString(cellValue, CultureInfo.InvariantCulture) ?? string.Empty;
            string filterValue = filter.Value ?? string.Empty;
            StringComparison comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            switch (filter.Operator)
            {
                case "contains":
                    return cellText.IndexOf(filterValue, comparison) >= 0;
                case "startswith":
                    return cellText.StartsWith(filterValue, comparison);
                case "regex":
                    return Regex.IsMatch(
                        cellText,
                        filterValue,
                        caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase);
                case "=":
                case "==":
                    return CompareEquality(cellText, filterValue, comparison);
                case "!=":
                case "<>":
                    return !CompareEquality(cellText, filterValue, comparison);
                default:
                    return CompareNumeric(cellText, filterValue, filter.Operator);
            }
        }

        private static bool CompareEquality(string left, string right, StringComparison comparison)
        {
            double leftNumber;
            double rightNumber;
            if (TryParseNumber(left, out leftNumber) && TryParseNumber(right, out rightNumber))
            {
                return Math.Abs(leftNumber - rightNumber) < 0.0000000001;
            }
            return string.Equals(left, right, comparison);
        }

        private static bool CompareNumeric(string left, string right, string op)
        {
            double leftNumber;
            double rightNumber;
            if (!TryParseNumber(left, out leftNumber) || !TryParseNumber(right, out rightNumber))
            {
                return false;
            }
            switch (op)
            {
                case ">":
                    return leftNumber > rightNumber;
                case "<":
                    return leftNumber < rightNumber;
                case ">=":
                    return leftNumber >= rightNumber;
                case "<=":
                    return leftNumber <= rightNumber;
                default:
                    return false;
            }
        }

        private static bool TryParseNumber(string text, out double number)
        {
            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
            {
                return true;
            }
            return double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out number);
        }

        private static bool MatchesLookup(
            object cellValue,
            List<string> keys,
            string matchMode,
            bool caseSensitive)
        {
            string text = Convert.ToString(cellValue, CultureInfo.InvariantCulture) ?? string.Empty;
            StringComparison comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            foreach (string key in keys)
            {
                if (matchMode == "exact")
                {
                    if (string.Equals(text.Trim(), key.Trim(), comparison))
                    {
                        return true;
                    }
                }
                else if (matchMode == "fuzzy")
                {
                    if (text.IndexOf(key, comparison) >= 0)
                    {
                        return true;
                    }
                }
                else if (matchMode == "regex")
                {
                    if (Regex.IsMatch(text, key, caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private static int[] ResolveReturnColumns(
            List<string> returnColumnNames,
            string[] headers,
            bool hasHeader,
            int totalColumns,
            string file,
            string sheet,
            JArray errors,
            JArray resultColumns)
        {
            if (returnColumnNames == null || returnColumnNames.Count == 0)
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
            foreach (string name in returnColumnNames)
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
            JArray filters,
            string[] headers,
            bool hasHeader,
            int totalColumns,
            string file,
            string sheet,
            JArray errors)
        {
            List<FilterSpec> specs = new List<FilterSpec>();
            if (filters == null)
            {
                return specs;
            }
            foreach (JToken token in filters)
            {
                JObject filter = token as JObject;
                if (filter == null)
                {
                    continue;
                }
                string columnText = ExcelArgumentReader.GetTrimmedString(filter, "column", "field", "col");
                string op = ExcelArgumentReader.GetTrimmedString(filter, "op", "operator").ToLowerInvariant();
                JToken valueToken = ExcelArgumentReader.First(filter, "value", "criteria");
                if (columnText.Length == 0 || op.Length == 0 || valueToken == null)
                {
                    errors.Add(FileError(
                        file + "!" + sheet,
                        "filters 项需要 column/op/value 三个字段。"));
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
                    errors.Add(FileError(
                        file + "!" + sheet,
                        "不支持的 filter op：「" + op + "」。"));
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
            if (hasHeader && headers != null)
            {
                for (int index = 0; index < headers.Length; index++)
                {
                    if (string.Equals(headers[index], name, StringComparison.OrdinalIgnoreCase))
                    {
                        return index + 1;
                    }
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
                if (baseName.Length == 0)
                {
                    baseName = "Column" + ExcelComSupport.ColumnLetter(index + 1);
                }
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

        private static List<Excel.Worksheet> GetTargetSheets(
            Excel.Workbook workbook,
            List<string> requestedSheets,
            string fullPath,
            JArray errors)
        {
            List<Excel.Worksheet> sheets = new List<Excel.Worksheet>();
            if (requestedSheets == null || requestedSheets.Count == 0)
            {
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
                        sheets.Add(worksheet);
                    }
                }
                finally
                {
                    ComHelper.Release(collection);
                }
                return sheets;
            }

            foreach (string sheetName in requestedSheets)
            {
                Excel.Worksheet worksheet = ExcelComSupport.FindWorksheet(workbook, sheetName);
                if (worksheet == null)
                {
                    errors.Add(FileError(fullPath, "工作表「" + sheetName + "」不存在。"));
                    continue;
                }
                sheets.Add(worksheet);
            }
            return sheets;
        }

        private static OperationResult WriteQueryResultToCurrentWorkbook(
            Excel.Application application,
            JObject queryData,
            string writeTarget)
        {
            Excel.Workbook workbook = null;
            Excel.Worksheet createdSheet = null;
            bool created = false;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                JArray resultColumns = queryData["columns"] as JArray ?? new JArray();
                JArray resultRows = queryData["rows"] as JArray ?? new JArray();
                string targetAddress = writeTarget;
                if (string.IsNullOrWhiteSpace(targetAddress))
                {
                    string sheetName = ExcelPivotService.UniqueSheetName(
                        workbook,
                        "查询结果_" + DateTime.Now.ToString("HHmmss", CultureInfo.InvariantCulture));
                    createdSheet = RangeResolver.CreateWorksheet(workbook, sheetName);
                    created = true;
                    targetAddress = ExcelComSupport.FullAddress(sheetName, "A1");
                }

                using (ExcelRangeContext target = RangeResolver.Resolve(
                           application,
                           targetAddress,
                           false,
                           "excel_query_files"))
                {
                    ExcelComSupport.EnsureWritable(application, target.Workbook, target.Worksheet, "excel_query_files");
                    int columns = 3 + resultColumns.Count;
                    int rows = resultRows.Count + 1;
                    Excel.Range writeRange = null;
                    try
                    {
                        writeRange = target.Range.Resize[rows, columns];
                        string address = writeRange.AddressLocal();
                        ExcelUndoScope undo = created
                            ? ExcelUndoScope.Failed("查询结果写入新建工作表，执行前不存在目标区域快照。")
                            : ExcelUndoScope.Capture(target.Workbook, target.Worksheet, writeRange, address);

                        object[,] matrix = new object[rows, columns];
                        matrix[0, 0] = "file";
                        matrix[0, 1] = "sheet";
                        matrix[0, 2] = "rowNumber";
                        for (int column = 3; column < columns; column++)
                        {
                            matrix[0, column] = resultColumns[column - 3].ToString();
                        }
                        for (int row = 1; row < rows; row++)
                        {
                            JObject resultRow = resultRows[row - 1] as JObject;
                            matrix[row, 0] = resultRow == null ? null : resultRow.Value<string>("file");
                            matrix[row, 1] = resultRow == null ? null : resultRow.Value<string>("sheet");
                            matrix[row, 2] = resultRow == null ? null : resultRow.Value<object>("rowNumber");
                            for (int column = 3; column < columns; column++)
                            {
                                string columnName = resultColumns[column - 3].ToString();
                                JToken value = resultRow == null ? null : resultRow[columnName];
                                matrix[row, column] = value == null ? null : ExcelComSupport.JTokenToValue(value);
                            }
                        }
                        writeRange.Value2 = matrix;

                        string targetRef = ExcelComSupport.FullAddress(target.SheetName, address);
                        Observation observation = Observation.Create(
                            "query_files_write",
                            "查询结果已写入当前活动工作簿 " + targetRef + "。",
                            true);
                        observation.AddTargetRef(targetRef);
                        JObject data = new JObject();
                        data["sheet"] = target.SheetName;
                        data["address"] = address;
                        data["rowsWritten"] = resultRows.Count;
                        data["columns"] = columns;
                        data["createdSheet"] = created;

                        undo.Complete(
                            target.Workbook,
                            target.Worksheet,
                            writeRange,
                            address,
                            false,
                            "写入查询结果 " + targetRef,
                            observation,
                            data,
                            created
                                ? "写入新建查询结果表，新增工作表本身不在区域快照可撤销范围。"
                                : null);

                        return OperationResult.Succeed(
                            "查询结果已写入 " + targetRef + "。",
                            data,
                            observation);
                    }
                    finally
                    {
                        ComHelper.Release(writeRange);
                    }
                }
            }
            catch
            {
                if (created && createdSheet != null)
                {
                    ExcelPivotService.TryDeleteWorksheet(application, createdSheet);
                }
                throw;
            }
            finally
            {
                ComHelper.Release(createdSheet, workbook);
            }
        }

        private static Excel.Workbook OpenReadOnly(HiddenExcelScope scope, string fullPath)
        {
            return scope.Workbooks.Open(
                Filename: fullPath,
                UpdateLinks: 0,
                ReadOnly: true,
                IgnoreReadOnlyRecommended: true,
                AddToMru: false,
                Notify: false);
        }

        private static void CloseWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }
            try
            {
                workbook.Close(false);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[excel] 关闭只读查询工作簿失败：" + ex.Message);
            }
            finally
            {
                ComHelper.Release(workbook);
            }
        }

        private static bool TryResolveFile(string requestedPath, out string fullPath)
        {
            fullPath = string.Empty;
            if (string.IsNullOrWhiteSpace(requestedPath))
            {
                return false;
            }
            string pathError;
            if (!Office.Config.WorkspacePathResolver.TryResolveDocument(
                    "excel", requestedPath, out fullPath, out pathError))
            {
                return false;
            }
            return File.Exists(fullPath);
        }

        private static void AddUniqueString(JArray array, string value)
        {
            if (array == null || string.IsNullOrWhiteSpace(value))
            {
                return;
            }
            foreach (JToken item in array)
            {
                if (string.Equals(item.ToString(), value, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
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

        private sealed class FilterSpec
        {
            public int Column { get; set; }

            public string Operator { get; set; }

            public string Value { get; set; }
        }

        /// <summary>
        /// 专用隐藏 Excel 实例。只用于只读查询；不进入 OfficeSession 单例。
        /// </summary>
        private sealed class HiddenExcelScope : IDisposable
        {
            public Excel.Application Application { get; private set; }

            public Excel.Workbooks Workbooks { get; private set; }

            public HiddenExcelScope()
            {
                ComHelper.EnsureSta("excel_remote_read");
                Excel.Application application = new Excel.Application();
                try
                {
                    application.Visible = false;
                    application.DisplayAlerts = false;
                    application.ScreenUpdating = false;
                    application.EnableEvents = false;
                    application.AskToUpdateLinks = false;
                    try
                    {
                        dynamic dynamicApplication = application;
                        dynamicApplication.AutomationSecurity = 3; // msoAutomationSecurityForceDisable
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[excel] 设置隐藏实例 AutomationSecurity 失败：" + ex.Message);
                    }
                    Application = application;
                    Workbooks = application.Workbooks;
                }
                catch
                {
                    TryQuit(application);
                    throw;
                }
            }

            public void Dispose()
            {
                if (Workbooks != null)
                {
                    try
                    {
                        while (Workbooks.Count > 0)
                        {
                            Excel.Workbook workbook = Workbooks[1];
                            try
                            {
                                workbook.Close(false);
                            }
                            finally
                            {
                                ComHelper.Release(workbook);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[excel] 清理隐藏实例工作簿失败：" + ex.Message);
                    }
                }

                TryQuit(Application);
                ComHelper.Release(Workbooks, Application);
                Workbooks = null;
                Application = null;
            }

            private static void TryQuit(Excel.Application application)
            {
                if (application == null)
                {
                    return;
                }
                try
                {
                    application.DisplayAlerts = false;
                    application.Quit();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[excel] 隐藏 Excel 实例退出失败：" + ex.Message);
                }
            }
        }
    }
}
