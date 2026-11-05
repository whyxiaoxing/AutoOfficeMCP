using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>P1 只读类扩展：全文搜索、错误扫描、工作表差异。</summary>
    public sealed partial class ExcelExtensionsExecutor
    {
        private const int SearchMaxScanCells = 500000;

        private static OperationResult ExecuteSearchAll(Excel.Application application, JObject arguments)
        {
            string query = ExcelArgumentReader.GetString(arguments, "query", "text", "findText", "keyword");
            if (query.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 query/text/findText 参数。");
            }

            bool useRegex = ExcelArgumentReader.GetBool(arguments, false, "useRegex", "regex");
            bool matchCase = ExcelArgumentReader.GetBool(arguments, false, "matchCase");
            bool matchEntireCell = ExcelArgumentReader.GetBool(arguments, false, "matchEntireCell", "entireCell");
            bool includeHidden = ExcelArgumentReader.GetBool(arguments, false, "includeHidden");
            bool writeToSheet = ExcelArgumentReader.GetBool(arguments, false, "writeToSheet", "write");
            int maxResults = ClampInt(
                ExcelArgumentReader.GetInt(arguments, 200, "maxResults", "limit"),
                1,
                5000);
            int contextLength = ClampInt(
                ExcelArgumentReader.GetInt(arguments, 30, "contextLength", "context"),
                0,
                200);
            string resultSheetName = ExcelArgumentReader.GetTrimmedString(
                arguments,
                "resultSheet",
                "outputSheet");
            if (resultSheetName.Length == 0)
            {
                resultSheetName = "搜索结果";
            }

            Regex regularExpression = null;
            if (useRegex)
            {
                try
                {
                    RegexOptions options = RegexOptions.CultureInvariant;
                    if (!matchCase)
                    {
                        options = options | RegexOptions.IgnoreCase;
                    }
                    regularExpression = new Regex(query, options);
                }
                catch (ArgumentException ex)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, ExcelExtensionSupport.ToRegexPatternError(ex));
                }
            }

            List<string> requestedSheets = ExcelArgumentReader.GetStringList(arguments, "sheets");
            Excel.Workbook workbook = null;
            List<Excel.Worksheet> worksheets = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                if (writeToSheet && workbook.ReadOnly)
                {
                    return Fail(OfficeErrorCode.READ_ONLY, "工作簿以只读方式打开，无法写入搜索结果表。");
                }

                worksheets = ExcelExtensionSupport.CollectWorksheets(
                    workbook,
                    includeHidden,
                    writeToSheet ? resultSheetName : string.Empty);
                FilterWorksheetsByName(worksheets, requestedSheets);
                if (worksheets.Count == 0)
                {
                    return Fail(OfficeErrorCode.SHEET_NOT_FOUND, "没有找到可搜索的工作表。");
                }

                JArray results = new JArray();
                bool truncated = false;
                int sheetsScanned = 0;
                StringComparison comparison = matchCase
                    ? StringComparison.Ordinal
                    : StringComparison.OrdinalIgnoreCase;

                foreach (Excel.Worksheet worksheet in worksheets)
                {
                    Excel.Range usedRange = null;
                    try
                    {
                        usedRange = worksheet.UsedRange;
                        int totalRows = usedRange.Rows.Count;
                        int totalColumns = usedRange.Columns.Count;
                        if (totalRows < 1 || totalColumns < 1)
                        {
                            continue;
                        }

                        int scanRows = totalRows;
                        int scanColumns = totalColumns;
                        if ((long)totalRows * totalColumns > SearchMaxScanCells)
                        {
                            scanColumns = Math.Min(totalColumns, 256);
                            scanRows = Math.Max(1, Math.Min(totalRows, SearchMaxScanCells / scanColumns));
                            truncated = true;
                        }

                        object[,] values = ExcelExtensionSupport.ReadMatrix(usedRange, scanRows, scanColumns);
                        sheetsScanned++;
                        bool stop = false;
                        for (int row = 0; row < scanRows && !stop; row++)
                        {
                            for (int column = 0; column < scanColumns; column++)
                            {
                                string text = ExcelExtensionSupport.CellText(values[row, column]);
                                if (text.Length == 0)
                                {
                                    continue;
                                }

                                string matchText;
                                int matchIndex;
                                int matchLength;
                                if (regularExpression != null)
                                {
                                    Match match = regularExpression.Match(text);
                                    if (!match.Success)
                                    {
                                        continue;
                                    }
                                    if (matchEntireCell && !string.Equals(match.Value, text, StringComparison.Ordinal))
                                    {
                                        continue;
                                    }
                                    matchText = match.Value;
                                    matchIndex = match.Index;
                                    matchLength = match.Length;
                                }
                                else
                                {
                                    int index = text.IndexOf(query, comparison);
                                    if (index < 0)
                                    {
                                        continue;
                                    }
                                    if (matchEntireCell && !string.Equals(text, query, comparison))
                                    {
                                        continue;
                                    }
                                    matchIndex = index;
                                    matchLength = query.Length;
                                    matchText = text.Substring(index, query.Length);
                                }

                                int cellRow = usedRange.Row + row;
                                int cellColumn = usedRange.Column + column;
                                JObject item = new JObject();
                                item["sheet"] = worksheet.Name;
                                item["address"] = ExcelExtensionSupport.FullCellAddress(
                                    worksheet.Name,
                                    cellRow,
                                    cellColumn);
                                item["row"] = cellRow;
                                item["column"] = cellColumn;
                                item["columnLetter"] = ExcelComSupport.ColumnLetter(cellColumn);
                                item["value"] = text;
                                item["match"] = matchText;
                                item["context"] = ExcelExtensionSupport.ContextAround(
                                    text,
                                    matchIndex,
                                    matchLength,
                                    contextLength);
                                results.Add(item);

                                if (results.Count >= maxResults)
                                {
                                    truncated = true;
                                    stop = true;
                                    break;
                                }
                            }
                        }
                    }
                    finally
                    {
                        ComHelper.Release(usedRange);
                    }

                    if (results.Count >= maxResults)
                    {
                        truncated = true;
                        break;
                    }
                }

                JObject data = new JObject();
                data["query"] = query;
                data["useRegex"] = useRegex;
                data["matchCase"] = matchCase;
                data["matchEntireCell"] = matchEntireCell;
                data["sheetsScanned"] = sheetsScanned;
                data["matchCount"] = results.Count;
                data["truncated"] = truncated;
                data["written"] = false;
                data["results"] = results;
                if (writeToSheet)
                {
                    return WriteSearchResults(
                        application,
                        workbook,
                        resultSheetName,
                        results,
                        data,
                        truncated);
                }

                string summary = truncated
                    ? "搜索完成（结果已截断）：返回前 " + results.Count + " 条匹配。"
                    : "搜索完成：共 " + results.Count + " 条匹配。";
                Observation observation = ReadOnlyObservation("search_all", summary);
                return OperationResult.Succeed(summary, data, observation);
            }
            finally
            {
                ExcelExtensionSupport.ReleaseWorksheets(worksheets);
                ComHelper.Release(workbook);
            }
        }

        private static OperationResult WriteSearchResults(
            Excel.Application application,
            Excel.Workbook workbook,
            string resultSheetName,
            JArray results,
            JObject data,
            bool truncated)
        {
            Excel.Worksheet target = null;
            Excel.Range beforeRange = null;
            Excel.Range destination = null;
            Excel.Range startCell = null;
            try
            {
                bool created = false;
                target = ExcelExtensionSupport.FindWorksheet(workbook, resultSheetName);
                ExcelUndoScope undo = null;
                string beforeAddress = string.Empty;
                if (target == null)
                {
                    target = RangeResolver.CreateWorksheet(workbook, resultSheetName);
                    created = true;
                }
                else
                {
                    ExcelComSupport.EnsureWritable(application, workbook, target, "excel_search_all");
                    beforeRange = target.UsedRange;
                    beforeAddress = beforeRange.AddressLocal();
                    undo = ExcelUndoScope.Capture(workbook, target, beforeRange, beforeAddress);
                    Excel.Range clearRange = beforeRange;
                    application.Guard(delegate
                    {
                        clearRange.ClearContents();
                        clearRange.ClearHyperlinks();
                    });
                }

                int rowCount = Math.Max(1, results.Count + 1);
                const int ColumnCount = 6;
                object[,] matrix = new object[rowCount, ColumnCount];
                string[] headers = { "工作表", "地址", "匹配内容", "上下文", "行", "列" };
                for (int column = 0; column < ColumnCount; column++)
                {
                    matrix[0, column] = headers[column];
                }
                for (int index = 0; index < results.Count; index++)
                {
                    JObject item = (JObject)results[index];
                    matrix[index + 1, 0] = item.Value<string>("sheet");
                    matrix[index + 1, 1] = item.Value<string>("address");
                    matrix[index + 1, 2] = item.Value<string>("match");
                    matrix[index + 1, 3] = item.Value<string>("context");
                    matrix[index + 1, 4] = item.Value<long>("row");
                    matrix[index + 1, 5] = item.Value<long>("column");
                }

                startCell = (Excel.Range)target.Cells[1, 1];
                destination = startCell.Resize[rowCount, ColumnCount];
                string destinationAddress = destination.AddressLocal();
                application.Guard(delegate
                {
                    destination.NumberFormat = "@";
                    destination.Value2 = matrix;
                });

                string targetRef = ExcelComSupport.FullAddress(target.Name, destinationAddress);
                string summary = "已将 " + results.Count + " 条搜索结果写入工作表「" + resultSheetName + "」。";
                Observation observation = Observation.Create("search_all", summary, true);
                observation.AddTargetRef(targetRef);

                data["written"] = true;
                data["resultSheet"] = resultSheetName;
                data["resultAddress"] = destinationAddress;
                data["truncated"] = truncated;

                if (created)
                {
                    ExcelUndoScope.RecordUnavailable(
                        data,
                        "搜索结果表为新建工作表，属于结构变更，区域快照无法撤销新增工作表。");
                }
                else
                {
                    undo.Complete(
                        workbook,
                        target,
                        destination,
                        destinationAddress,
                        true,
                        "写入搜索结果 " + targetRef,
                        observation,
                        data,
                        "旧结果中的超链接不在区域快照中，自动撤销只恢复单元格值/公式/数字格式。");
                }
                return OperationResult.Succeed(summary, data, observation);
            }
            finally
            {
                ComHelper.Release(startCell, destination, beforeRange, target);
            }
        }

        private static void FilterWorksheetsByName(List<Excel.Worksheet> worksheets, List<string> requestedNames)
        {
            if (requestedNames == null || requestedNames.Count == 0)
            {
                return;
            }
            for (int index = worksheets.Count - 1; index >= 0; index--)
            {
                Excel.Worksheet worksheet = worksheets[index];
                bool found = false;
                foreach (string name in requestedNames)
                {
                    if (string.Equals(worksheet.Name, name, StringComparison.OrdinalIgnoreCase))
                    {
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    ComHelper.Release(worksheet);
                    worksheets.RemoveAt(index);
                }
            }
        }

        private static OperationResult ExecuteErrorScan(Excel.Application application, JObject arguments)
        {
            string rangeAddress = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange");
            string sheetName = ExcelArgumentReader.GetSheetName(arguments);
            bool includeHidden = ExcelArgumentReader.GetBool(arguments, true, "includeHidden");
            int maxResults = ClampInt(
                ExcelArgumentReader.GetInt(arguments, 500, "maxResults", "limit"),
                1,
                10000);

            Excel.Workbook workbook = null;
            List<Excel.Worksheet> worksheets = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                JArray entries = new JArray();
                JObject typeCounts = new JObject();
                HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                bool truncated = false;
                int sheetsScanned = 0;
                JArray sheets = new JArray();

                if (!string.IsNullOrWhiteSpace(rangeAddress))
                {
                    using (ExcelRangeContext context = RangeResolver.Resolve(
                               application,
                               rangeAddress,
                               false,
                               "excel_error_scan"))
                    {
                        sheetsScanned = 1;
                        sheets.Add(context.SheetName);
                        ScanErrorRange(
                            context.Worksheet,
                            context.Range,
                            entries,
                            typeCounts,
                            seen,
                            maxResults,
                            ref truncated);
                    }
                }
                else
                {
                    worksheets = ExcelExtensionSupport.CollectWorksheets(workbook, includeHidden, string.Empty);
                    if (sheetName.Length > 0)
                    {
                        List<string> requested = new List<string>();
                        requested.Add(sheetName);
                        FilterWorksheetsByName(worksheets, requested);
                    }
                    if (worksheets.Count == 0)
                    {
                        return Fail(OfficeErrorCode.SHEET_NOT_FOUND, "没有找到可扫描的工作表。");
                    }

                    foreach (Excel.Worksheet worksheet in worksheets)
                    {
                        Excel.Range usedRange = null;
                        try
                        {
                            usedRange = worksheet.UsedRange;
                            sheetsScanned++;
                            sheets.Add(worksheet.Name);
                            ScanErrorRange(
                                worksheet,
                                usedRange,
                                entries,
                                typeCounts,
                                seen,
                                maxResults,
                                ref truncated);
                        }
                        finally
                        {
                            ComHelper.Release(usedRange);
                        }
                        if (entries.Count >= maxResults)
                        {
                            truncated = true;
                            break;
                        }
                    }
                }

                JObject data = new JObject();
                data["sheetsScanned"] = sheetsScanned;
                data["sheets"] = sheets;
                data["errorCount"] = entries.Count;
                data["errorTypes"] = typeCounts;
                data["truncated"] = truncated;
                data["errors"] = entries;

                string summary = entries.Count == 0
                    ? "扫描完成：未发现错误单元格。"
                    : "扫描完成：发现 " + entries.Count + " 个错误单元格" + (truncated ? "（结果已截断）。" : "。");
                Observation observation = ReadOnlyObservation("error_scan", summary);
                return OperationResult.Succeed(summary, data, observation);
            }
            finally
            {
                ExcelExtensionSupport.ReleaseWorksheets(worksheets);
                ComHelper.Release(workbook);
            }
        }

        private static void ScanErrorRange(
            Excel.Worksheet worksheet,
            Excel.Range range,
            JArray entries,
            JObject typeCounts,
            HashSet<string> seen,
            int maxResults,
            ref bool truncated)
        {
            if (worksheet == null || range == null || entries.Count >= maxResults)
            {
                return;
            }

            Excel.Range formulaErrors = null;
            Excel.Range constantErrors = null;
            try
            {
                try
                {
                    formulaErrors = range.SpecialCells(
                        Excel.XlCellType.xlCellTypeFormulas,
                        Excel.XlSpecialCellsValue.xlErrors);
                }
                catch (COMException)
                {
                    formulaErrors = null;
                }
                try
                {
                    constantErrors = range.SpecialCells(
                        Excel.XlCellType.xlCellTypeConstants,
                        Excel.XlSpecialCellsValue.xlErrors);
                }
                catch (COMException)
                {
                    constantErrors = null;
                }

                CollectErrorCells(worksheet, formulaErrors, entries, typeCounts, seen, maxResults, ref truncated);
                if (entries.Count < maxResults)
                {
                    CollectErrorCells(worksheet, constantErrors, entries, typeCounts, seen, maxResults, ref truncated);
                }
            }
            finally
            {
                ComHelper.Release(formulaErrors, constantErrors);
            }
        }

        private static void CollectErrorCells(
            Excel.Worksheet worksheet,
            Excel.Range specialRange,
            JArray entries,
            JObject typeCounts,
            HashSet<string> seen,
            int maxResults,
            ref bool truncated)
        {
            if (specialRange == null || entries.Count >= maxResults)
            {
                return;
            }

            Excel.Areas areas = null;
            try
            {
                areas = specialRange.Areas;
                int areaCount = areas.Count;
                for (int areaIndex = 1; areaIndex <= areaCount; areaIndex++)
                {
                    Excel.Range area = null;
                    try
                    {
                        area = (Excel.Range)areas[areaIndex];
                        int rows = area.Rows.Count;
                        int columns = area.Columns.Count;
                        for (int row = 1; row <= rows; row++)
                        {
                            for (int column = 1; column <= columns; column++)
                            {
                                Excel.Range cell = null;
                                try
                                {
                                    cell = (Excel.Range)area.Cells[row, column];
                                    string errorText = ExcelExtensionSupport.ErrorTextFromValue(cell.Value2);
                                    if (errorText.Length == 0)
                                    {
                                        string cellText = ExcelExtensionSupport.CellText(cell.Value2);
                                        if (cellText.StartsWith("#", StringComparison.Ordinal))
                                        {
                                            errorText = cellText;
                                        }
                                    }
                                    if (errorText.Length == 0)
                                    {
                                        continue;
                                    }

                                    string localAddress = cell.AddressLocal();
                                    string fullAddress = ExcelComSupport.FullAddress(
                                        worksheet.Name,
                                        localAddress);
                                    if (!seen.Add(fullAddress))
                                    {
                                        continue;
                                    }

                                    string formula = string.Empty;
                                    try
                                    {
                                        if (Convert.ToBoolean(cell.HasFormula, CultureInfo.InvariantCulture))
                                        {
                                            formula = Convert.ToString(cell.Formula, CultureInfo.InvariantCulture) ?? string.Empty;
                                        }
                                        else
                                        {
                                            formula = ExcelExtensionSupport.CellText(cell.Value2);
                                        }
                                    }
                                    catch (COMException)
                                    {
                                        formula = string.Empty;
                                    }

                                    JObject item = new JObject();
                                    item["sheet"] = worksheet.Name;
                                    item["address"] = fullAddress;
                                    item["addressLocal"] = localAddress;
                                    item["row"] = cell.Row;
                                    item["column"] = cell.Column;
                                    item["error"] = errorText;
                                    item["formula"] = formula;
                                    item["hasFormula"] = Convert.ToBoolean(cell.HasFormula, CultureInfo.InvariantCulture);
                                    entries.Add(item);

                                    JToken existingCount = typeCounts[errorText];
                                    int count = existingCount == null ? 0 : existingCount.Value<int>();
                                    typeCounts[errorText] = count + 1;

                                    if (entries.Count >= maxResults)
                                    {
                                        truncated = true;
                                        return;
                                    }
                                }
                                finally
                                {
                                    ComHelper.Release(cell);
                                }
                            }
                        }
                    }
                    finally
                    {
                        ComHelper.Release(area);
                    }
                }
            }
            finally
            {
                ComHelper.Release(areas);
            }
        }

        private static OperationResult ExecuteDiffSheets(Excel.Application application, JObject arguments)
        {
            string leftName = ExcelArgumentReader.GetTrimmedString(
                arguments,
                "sheet1",
                "leftSheet",
                "sourceSheet");
            string rightName = ExcelArgumentReader.GetTrimmedString(
                arguments,
                "sheet2",
                "rightSheet",
                "targetSheet");
            if (leftName.Length == 0 || rightName.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 sheet1 与 sheet2 参数。");
            }
            if (string.Equals(leftName, rightName, StringComparison.OrdinalIgnoreCase))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "sheet1 与 sheet2 不能是同一个工作表。");
            }

            string compareBy = ExcelArgumentReader.GetTrimmedString(arguments, "compareBy", "compare")
                .ToLowerInvariant();
            if (compareBy.Length == 0)
            {
                compareBy = "value";
            }
            if (compareBy != "value" && compareBy != "formula")
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "compareBy 仅支持 value/formula。");
            }

            string leftRangeAddress = ExcelArgumentReader.GetTrimmedString(arguments, "range1", "leftRange", "sourceRange");
            string rightRangeAddress = ExcelArgumentReader.GetTrimmedString(arguments, "range2", "rightRange", "targetRange");
            bool ignoreCase = ExcelArgumentReader.GetBool(arguments, false, "ignoreCase");
            bool trimWhitespace = ExcelArgumentReader.GetBool(arguments, false, "trimWhitespace");
            bool includeEqual = ExcelArgumentReader.GetBool(arguments, false, "includeEqual");
            bool highlight = ExcelArgumentReader.GetBool(arguments, false, "highlight");
            int maxDiffs = ClampInt(
                ExcelArgumentReader.GetInt(arguments, 200, "maxDiffs", "maxResults", "limit"),
                1,
                5000);
            int maxCells = ClampInt(
                ExcelArgumentReader.GetInt(arguments, 50000, "maxCells"),
                1,
                200000);
            int highlightColor = ExcelComSupport.ParseColor(
                ExcelArgumentReader.GetString(arguments, "highlightColor", "color"),
                ExcelComSupport.ToOleColor(255, 199, 206));

            Excel.Workbook workbook = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                using (ExcelRangeContext left = RangeResolver.ResolveOnSheet(
                           application,
                           leftName,
                           string.IsNullOrWhiteSpace(leftRangeAddress) ? "usedrange" : leftRangeAddress,
                           false,
                           "excel_diff_sheets"))
                using (ExcelRangeContext right = RangeResolver.ResolveOnSheet(
                           application,
                           rightName,
                           string.IsNullOrWhiteSpace(rightRangeAddress) ? "usedrange" : rightRangeAddress,
                           false,
                           "excel_diff_sheets"))
                {
                    int leftRows = left.Range.Rows.Count;
                    int leftColumns = left.Range.Columns.Count;
                    int rightRows = right.Range.Rows.Count;
                    int rightColumns = right.Range.Columns.Count;
                    int rows = Math.Max(leftRows, rightRows);
                    int columns = Math.Max(leftColumns, rightColumns);
                    long cells = (long)rows * columns;
                    if (cells > maxCells)
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "比较区域共 " + cells + " 个单元格，超过 maxCells=" + maxCells +
                            "。请缩小范围或提高 maxCells（上限 200000）。");
                    }

                    object[,] leftValues = ExcelExtensionSupport.ReadMatrix(left.Range, leftRows, leftColumns);
                    object[,] rightValues = ExcelExtensionSupport.ReadMatrix(right.Range, rightRows, rightColumns);
                    object[,] leftFormulas = null;
                    object[,] rightFormulas = null;
                    if (compareBy == "formula")
                    {
                        leftFormulas = ExcelExtensionSupport.ReadFormulaMatrix(left.Range, leftRows, leftColumns);
                        rightFormulas = ExcelExtensionSupport.ReadFormulaMatrix(right.Range, rightRows, rightColumns);
                    }

                    bool highlightChanged = false;
                    if (highlight)
                    {
                        ExcelComSupport.EnsureWritable(application, workbook, left.Worksheet, "excel_diff_sheets");
                        ExcelComSupport.EnsureWritable(application, workbook, right.Worksheet, "excel_diff_sheets");
                    }

                    JArray diffs = new JArray();
                    int diffCount = 0;
                    int highlighted = 0;
                    bool truncated = false;
                    for (int row = 0; row < rows; row++)
                    {
                        for (int column = 0; column < columns; column++)
                        {
                            object leftValue;
                            object rightValue;
                            if (compareBy == "formula")
                            {
                                leftValue = row < leftRows && column < leftColumns
                                    ? leftFormulas[row, column]
                                    : null;
                                rightValue = row < rightRows && column < rightColumns
                                    ? rightFormulas[row, column]
                                    : null;
                            }
                            else
                            {
                                leftValue = row < leftRows && column < leftColumns
                                    ? leftValues[row, column]
                                    : null;
                                rightValue = row < rightRows && column < rightColumns
                                    ? rightValues[row, column]
                                    : null;
                            }

                            string leftText = ExcelExtensionSupport.CellText(leftValue);
                            string rightText = ExcelExtensionSupport.CellText(rightValue);
                            bool equal = ComparisonEquals(leftText, rightText, ignoreCase, trimWhitespace);
                            if (equal && !includeEqual)
                            {
                                continue;
                            }
                            if (!equal)
                            {
                                diffCount++;
                            }
                            if (diffs.Count >= maxDiffs)
                            {
                                truncated = true;
                                continue;
                            }

                            int leftRow = left.Range.Row + row;
                            int leftColumn = left.Range.Column + column;
                            int rightRow = right.Range.Row + row;
                            int rightColumn = right.Range.Column + column;
                            JObject entry = new JObject();
                            entry["row"] = row + 1;
                            entry["column"] = column + 1;
                            entry["address1"] = row < leftRows && column < leftColumns
                                ? ExcelExtensionSupport.FullCellAddress(left.SheetName, leftRow, leftColumn)
                                : string.Empty;
                            entry["address2"] = row < rightRows && column < rightColumns
                                ? ExcelExtensionSupport.FullCellAddress(right.SheetName, rightRow, rightColumn)
                                : string.Empty;
                            entry["left"] = leftText;
                            entry["right"] = rightText;
                            entry["equal"] = equal;

                            if (!equal && highlight)
                            {
                                if (row < leftRows && column < leftColumns)
                                {
                                    SetCellColor(left.Worksheet, leftRow, leftColumn, highlightColor);
                                }
                                if (row < rightRows && column < rightColumns)
                                {
                                    SetCellColor(right.Worksheet, rightRow, rightColumn, highlightColor);
                                }
                                highlighted++;
                            }
                            diffs.Add(entry);
                        }
                    }

                    highlightChanged = highlighted > 0;
                    JObject data = new JObject();
                    data["sheet1"] = left.SheetName;
                    data["sheet2"] = right.SheetName;
                    data["range1"] = left.Range.AddressLocal();
                    data["range2"] = right.Range.AddressLocal();
                    data["compareBy"] = compareBy;
                    data["comparedCells"] = cells;
                    data["diffCount"] = diffCount;
                    data["returnedDiffs"] = diffs.Count;
                    data["truncated"] = truncated;
                    data["highlight"] = highlight;
                    data["highlightedCells"] = highlighted;
                    data["diffs"] = diffs;
                    if (highlight)
                    {
                        ExcelUndoScope.RecordUnavailable(
                            data,
                            "高亮使用单元格填充色，区域快照不包含视觉格式，未登记自动撤销。");
                    }

                    Observation observation = Observation.Create(
                        "diff_sheets",
                        "工作表差异比较完成：" + left.SheetName + " vs " + right.SheetName +
                        "，差异 " + diffCount + " 处。",
                        highlightChanged);
                    if (highlightChanged)
                    {
                        int addedRefs = 0;
                        foreach (JToken token in diffs)
                        {
                            JObject entry = token as JObject;
                            if (entry == null || entry.Value<bool>("equal"))
                            {
                                continue;
                            }
                            observation.AddTargetRef(entry.Value<string>("address1"));
                            observation.AddTargetRef(entry.Value<string>("address2"));
                            addedRefs++;
                            if (addedRefs >= 20)
                            {
                                break;
                            }
                        }
                    }

                    string summary = highlightChanged
                        ? "比较完成：差异 " + diffCount + " 处，已高亮 " + highlighted + " 个差异单元格。"
                        : "比较完成：差异 " + diffCount + " 处。";
                    return OperationResult.Succeed(summary, data, observation);
                }
            }
            finally
            {
                ComHelper.Release(workbook);
            }
        }

        private static bool ComparisonEquals(
            string left,
            string right,
            bool ignoreCase,
            bool trimWhitespace)
        {
            string leftText = left ?? string.Empty;
            string rightText = right ?? string.Empty;
            if (trimWhitespace)
            {
                leftText = leftText.Trim();
                rightText = rightText.Trim();
            }
            return string.Equals(
                leftText,
                rightText,
                ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
        }

        private static void SetCellColor(Excel.Worksheet worksheet, int row, int column, int color)
        {
            Excel.Range cell = null;
            Excel.Interior interior = null;
            try
            {
                cell = (Excel.Range)worksheet.Cells[row, column];
                interior = cell.Interior;
                interior.Color = color;
            }
            finally
            {
                ComHelper.Release(interior, cell);
            }
        }
    }
}
