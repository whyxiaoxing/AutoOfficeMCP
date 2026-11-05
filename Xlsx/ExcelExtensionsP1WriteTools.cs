using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>P1 数据加工类扩展：正则提取、人民币大写、文本型数字修复。</summary>
    public sealed partial class ExcelExtensionsExecutor
    {
        private static OperationResult ExecuteRegexExtract(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(
                arguments,
                "range",
                "sourceRange",
                "targetRange",
                "address");
            if (string.IsNullOrWhiteSpace(address))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 range/sourceRange 参数。");
            }

            string pattern = ExcelArgumentReader.GetString(arguments, "pattern", "regex");
            if (pattern.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 pattern 参数。");
            }

            bool matchCase = ExcelArgumentReader.GetBool(arguments, false, "matchCase");
            bool multiMatch = ExcelArgumentReader.GetBool(arguments, false, "multiMatch", "allMatches");
            bool write = ExcelArgumentReader.GetBool(arguments, true, "write", "writeToSheet");
            string joinSeparator = ExcelArgumentReader.GetString(arguments, "joinSeparator", "separator");
            if (joinSeparator.Length == 0)
            {
                joinSeparator = ", ";
            }
            int previewLimit = ClampInt(
                ExcelArgumentReader.GetInt(arguments, 200, "maxResults", "previewLimit"),
                1,
                5000);

            Regex regex;
            try
            {
                RegexOptions options = RegexOptions.CultureInvariant;
                if (!matchCase)
                {
                    options = options | RegexOptions.IgnoreCase;
                }
                regex = new Regex(pattern, options);
            }
            catch (ArgumentException ex)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, ExcelExtensionSupport.ToRegexPatternError(ex));
            }

            List<RegexGroupSelector> selectors;
            string selectorError;
            if (!BuildGroupSelectors(arguments, regex, out selectors, out selectorError))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, selectorError);
            }

            string outputSheetName = ExcelArgumentReader.GetTrimmedString(arguments, "outputSheet", "targetSheet");
            using (ExcelRangeContext source = RangeResolver.Resolve(
                       application,
                       address,
                       false,
                       "excel_regex_extract"))
            {
                int rows = source.Range.Rows.Count;
                int columns = source.Range.Columns.Count;
                if (write)
                {
                    ExcelComSupport.EnsureWritable(application, source.Workbook, source.Worksheet, "excel_regex_extract");
                }

                object[,] values = ExcelExtensionSupport.ReadMatrix(source.Range, rows, columns);
                object[,] output = new object[rows, selectors.Count];
                JArray preview = new JArray();
                int matchedCells = 0;

                for (int row = 0; row < rows; row++)
                {
                    for (int column = 0; column < columns; column++)
                    {
                        string text = ExcelExtensionSupport.CellText(values[row, column]);
                        if (text.Length == 0)
                        {
                            continue;
                        }

                        bool anyMatch = false;
                        List<string> extracted = new List<string>();
                        if (multiMatch)
                        {
                            MatchCollection matches = regex.Matches(text);
                            foreach (RegexGroupSelector selector in selectors)
                            {
                                List<string> parts = new List<string>();
                                foreach (Match match in matches)
                                {
                                    Group group = ResolveGroup(match, selector);
                                    if (group != null && group.Success)
                                    {
                                        parts.Add(group.Value);
                                    }
                                }
                                string joined = string.Join(joinSeparator, parts.ToArray());
                                extracted.Add(joined);
                                if (joined.Length > 0)
                                {
                                    anyMatch = true;
                                }
                            }
                        }
                        else
                        {
                            Match match = regex.Match(text);
                            foreach (RegexGroupSelector selector in selectors)
                            {
                                Group group = match.Success ? ResolveGroup(match, selector) : null;
                                string value = group != null && group.Success ? group.Value : string.Empty;
                                extracted.Add(value);
                                if (value.Length > 0)
                                {
                                    anyMatch = true;
                                }
                            }
                        }

                        if (!anyMatch)
                        {
                            continue;
                        }

                        matchedCells++;
                        int cellRow = source.Range.Row + row;
                        int cellColumn = source.Range.Column + column;
                        for (int selectorIndex = 0; selectorIndex < selectors.Count; selectorIndex++)
                        {
                            output[row, selectorIndex] = extracted[selectorIndex];
                        }

                        if (preview.Count < previewLimit)
                        {
                            JObject item = new JObject();
                            item["address"] = ExcelExtensionSupport.FullCellAddress(
                                source.SheetName,
                                cellRow,
                                cellColumn);
                            item["source"] = text;
                            JArray valuesArray = new JArray();
                            foreach (string value in extracted)
                            {
                                valuesArray.Add(value);
                            }
                            item["extracted"] = valuesArray;
                            preview.Add(item);
                        }
                    }
                }

                JObject data = new JObject();
                data["pattern"] = pattern;
                data["source"] = source.FullAddress;
                data["matchedCells"] = matchedCells;
                data["selectors"] = BuildSelectorJson(selectors);
                data["write"] = write;
                data["preview"] = preview;

                if (!write)
                {
                    Observation readObservation = ReadOnlyObservation(
                        "regex_extract",
                        "正则提取完成：匹配 " + matchedCells + " 个单元格，未写入工作表。");
                    return OperationResult.Succeed(
                        "正则提取完成：匹配 " + matchedCells + " 个单元格。",
                        data,
                        readObservation);
                }

                if (matchedCells == 0)
                {
                    Observation unchanged = ReadOnlyObservation(
                        "regex_extract",
                        "正则提取完成：没有匹配项，未写入任何单元格。");
                    data["written"] = false;
                    return OperationResult.Succeed("没有匹配项，未写入任何单元格。", data, unchanged);
                }

                Excel.Worksheet outputWorksheet = source.Worksheet;
                bool ownsOutputWorksheet = false;
                Excel.Range startCell = null;
                Excel.Range destination = null;
                try
                {
                    if (!string.IsNullOrWhiteSpace(outputSheetName) &&
                        !string.Equals(outputSheetName, source.SheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        outputWorksheet = RangeResolver.ResolveWorksheet(
                            application,
                            source.Workbook,
                            outputSheetName,
                            false);
                        ownsOutputWorksheet = true;
                        ExcelComSupport.EnsureWritable(
                            application,
                            source.Workbook,
                            outputWorksheet,
                            "excel_regex_extract");
                    }

                    int outputColumn = ExcelExtensionSupport.ResolveOutputColumn(arguments, source.Range, 0);
                    ExcelExtensionSupport.EnsureRangeCapacity(
                        source.Range.Row,
                        outputColumn,
                        rows,
                        selectors.Count);

                    startCell = (Excel.Range)outputWorksheet.Cells[source.Range.Row, outputColumn];
                    destination = startCell.Resize[rows, selectors.Count];
                    string outputAddress = destination.AddressLocal();
                    ExcelUndoScope undo = ExcelUndoScope.Capture(
                        source.Workbook,
                        outputWorksheet,
                        destination,
                        outputAddress);

                    application.Guard(delegate
                    {
                        destination.NumberFormat = "@";
                        destination.Value2 = output;
                    });

                    string targetRef = ExcelComSupport.FullAddress(outputWorksheet.Name, outputAddress);
                    Observation observation = Observation.Create(
                        "regex_extract",
                        "正则提取完成：匹配 " + matchedCells + " 个单元格，已写入 " + targetRef + "。",
                        true);
                    observation.AddTargetRef(targetRef);

                    data["written"] = true;
                    data["outputSheet"] = outputWorksheet.Name;
                    data["outputAddress"] = outputAddress;
                    data["outputRows"] = rows;
                    data["outputColumns"] = selectors.Count;

                    undo.Complete(
                        source.Workbook,
                        outputWorksheet,
                        destination,
                        outputAddress,
                        true,
                        "正则提取写入 " + targetRef,
                        observation,
                        data,
                        null);

                    return OperationResult.Succeed(
                        "正则提取完成：匹配 " + matchedCells + " 个单元格，已写入 " + targetRef + "。",
                        data,
                        observation);
                }
                finally
                {
                    ComHelper.Release(startCell, destination);
                    if (ownsOutputWorksheet)
                    {
                        ComHelper.Release(outputWorksheet);
                    }
                }
            }
        }

        private sealed class RegexGroupSelector
        {
            public int Index { get; set; }

            public string Name { get; set; }

            public string Header { get; set; }
        }

        private static bool BuildGroupSelectors(
            JObject arguments,
            Regex regex,
            out List<RegexGroupSelector> selectors,
            out string error)
        {
            selectors = new List<RegexGroupSelector>();
            error = string.Empty;

            JArray groups = ExcelArgumentReader.GetArray(arguments, "groups");
            if (groups != null && groups.Count > 0)
            {
                foreach (JToken token in groups)
                {
                    RegexGroupSelector selector = ParseGroupSelector(token, regex, out error);
                    if (selector == null)
                    {
                        selectors.Clear();
                        return false;
                    }
                    selectors.Add(selector);
                }
                return true;
            }

            JToken groupToken = ExcelArgumentReader.First(arguments, "group", "groupIndex");
            string groupName = ExcelArgumentReader.GetTrimmedString(arguments, "groupName");
            RegexGroupSelector single = null;
            if (!string.IsNullOrWhiteSpace(groupName))
            {
                single = CreateNamedSelector(groupName, regex, out error);
            }
            else if (groupToken != null && groupToken.Type != JTokenType.Null)
            {
                single = ParseGroupSelector(groupToken, regex, out error);
            }
            else
            {
                single = new RegexGroupSelector();
                single.Index = 0;
                single.Name = string.Empty;
                single.Header = "group0";
            }
            if (single == null)
            {
                return false;
            }
            selectors.Add(single);
            return true;
        }

        private static RegexGroupSelector ParseGroupSelector(
            JToken token,
            Regex regex,
            out string error)
        {
            error = string.Empty;
            if (token == null)
            {
                error = "groups 中存在空项。";
                return null;
            }

            JObject config = token as JObject;
            if (config != null)
            {
                string header = ExcelArgumentReader.GetTrimmedString(config, "header", "name");
                string name = ExcelArgumentReader.GetTrimmedString(config, "groupName", "name");
                int index = ExcelArgumentReader.GetInt(config, 0, "index", "group");
                if (!string.IsNullOrWhiteSpace(name) &&
                    string.Equals(header, name, StringComparison.Ordinal))
                {
                    RegexGroupSelector named = CreateNamedSelector(name, regex, out error);
                    if (named != null)
                    {
                        named.Header = header;
                    }
                    return named;
                }
                RegexGroupSelector byIndex = CreateIndexedSelector(index, regex, out error);
                if (byIndex != null)
                {
                    byIndex.Header = string.IsNullOrWhiteSpace(header) ? "group" + index : header;
                }
                return byIndex;
            }

            string text = token.ToString().Trim();
            int parsed;
            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed))
            {
                return CreateIndexedSelector(parsed, regex, out error);
            }
            return CreateNamedSelector(text, regex, out error);
        }

        private static RegexGroupSelector CreateIndexedSelector(
            int index,
            Regex regex,
            out string error)
        {
            error = string.Empty;
            if (index < 0)
            {
                error = "group 索引不能为负数。";
                return null;
            }
            bool exists = false;
            foreach (int groupNumber in regex.GetGroupNumbers())
            {
                if (groupNumber == index)
                {
                    exists = true;
                    break;
                }
            }
            if (!exists)
            {
                error = "正则中没有编号为 " + index + " 的捕获组。";
                return null;
            }
            RegexGroupSelector selector = new RegexGroupSelector();
            selector.Index = index;
            selector.Name = string.Empty;
            selector.Header = "group" + index;
            return selector;
        }

        private static RegexGroupSelector CreateNamedSelector(
            string groupName,
            Regex regex,
            out string error)
        {
            error = string.Empty;
            int index = regex.GroupNumberFromName(groupName);
            if (index < 0)
            {
                error = "正则中没有名为「" + groupName + "」的捕获组。";
                return null;
            }
            RegexGroupSelector selector = new RegexGroupSelector();
            selector.Index = index;
            selector.Name = groupName;
            selector.Header = groupName;
            return selector;
        }

        private static Group ResolveGroup(Match match, RegexGroupSelector selector)
        {
            if (!string.IsNullOrWhiteSpace(selector.Name))
            {
                return match.Groups[selector.Name];
            }
            return match.Groups[selector.Index];
        }

        private static JArray BuildSelectorJson(List<RegexGroupSelector> selectors)
        {
            JArray result = new JArray();
            foreach (RegexGroupSelector selector in selectors)
            {
                JObject item = new JObject();
                item["index"] = selector.Index;
                item["name"] = selector.Name ?? string.Empty;
                item["header"] = selector.Header ?? string.Empty;
                result.Add(item);
            }
            return result;
        }

        private static OperationResult ExecuteRmbUpper(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(
                arguments,
                "range",
                "sourceRange",
                "targetRange",
                "address");
            if (string.IsNullOrWhiteSpace(address))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 range/sourceRange 参数。");
            }

            bool write = ExcelArgumentReader.GetBool(arguments, true, "write", "writeToSheet");
            string outputSheetName = ExcelArgumentReader.GetTrimmedString(arguments, "outputSheet", "targetSheet");

            using (ExcelRangeContext source = RangeResolver.Resolve(
                       application,
                       address,
                       false,
                       "excel_rmb_upper"))
            {
                int rows = source.Range.Rows.Count;
                int columns = source.Range.Columns.Count;
                if (write)
                {
                    ExcelComSupport.EnsureWritable(application, source.Workbook, source.Worksheet, "excel_rmb_upper");
                }

                object[,] values = ExcelExtensionSupport.ReadMatrix(source.Range, rows, columns);
                object[,] output = new object[rows, columns];
                JArray samples = new JArray();
                JArray invalid = new JArray();
                int convertedCount = 0;
                int blankCount = 0;
                int invalidTotal = 0;

                for (int row = 0; row < rows; row++)
                {
                    for (int column = 0; column < columns; column++)
                    {
                        object value = values[row, column];
                        if (ExcelExtensionSupport.IsBlankValue(value))
                        {
                            output[row, column] = null;
                            blankCount++;
                            continue;
                        }

                        string upper;
                        string reason;
                        if (!ExcelExtensionSupport.TryConvertRmb(value, out upper, out reason))
                        {
                            output[row, column] = null;
                            invalidTotal++;
                            if (invalid.Count < 50)
                            {
                                JObject item = new JObject();
                                item["address"] = ExcelExtensionSupport.FullCellAddress(
                                    source.SheetName,
                                    source.Range.Row + row,
                                    source.Range.Column + column);
                                item["value"] = ExcelExtensionSupport.CellText(value);
                                item["reason"] = reason;
                                invalid.Add(item);
                            }
                            continue;
                        }

                        output[row, column] = upper;
                        convertedCount++;
                        if (samples.Count < 20)
                        {
                            JObject item = new JObject();
                            item["address"] = ExcelExtensionSupport.FullCellAddress(
                                source.SheetName,
                                source.Range.Row + row,
                                source.Range.Column + column);
                            item["value"] = ExcelExtensionSupport.CellText(value);
                            item["upper"] = upper;
                            samples.Add(item);
                        }
                    }
                }

                JObject data = new JObject();
                data["source"] = source.FullAddress;
                data["convertedCount"] = convertedCount;
                data["blankCount"] = blankCount;
                data["invalidCount"] = invalidTotal;
                data["invalid"] = invalid;
                data["samples"] = samples;
                data["write"] = write;

                if (!write)
                {
                    Observation readObservation = ReadOnlyObservation(
                        "rmb_upper",
                        "人民币大写转换完成：成功 " + convertedCount + " 个，失败 " + invalid.Count + " 个（未写入）。");
                    return OperationResult.Succeed(
                        "人民币大写转换完成：成功 " + convertedCount + " 个。",
                        data,
                        readObservation);
                }

                if (convertedCount == 0)
                {
                    Observation unchanged = ReadOnlyObservation(
                        "rmb_upper",
                        "人民币大写转换完成：没有可转换的金额，未写入任何单元格。");
                    data["written"] = false;
                    return OperationResult.Succeed("没有可转换的金额，未写入任何单元格。", data, unchanged);
                }

                Excel.Worksheet outputWorksheet = source.Worksheet;
                bool ownsOutputWorksheet = false;
                Excel.Range startCell = null;
                Excel.Range destination = null;
                try
                {
                    if (!string.IsNullOrWhiteSpace(outputSheetName) &&
                        !string.Equals(outputSheetName, source.SheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        outputWorksheet = RangeResolver.ResolveWorksheet(
                            application,
                            source.Workbook,
                            outputSheetName,
                            false);
                        ownsOutputWorksheet = true;
                        ExcelComSupport.EnsureWritable(
                            application,
                            source.Workbook,
                            outputWorksheet,
                            "excel_rmb_upper");
                    }

                    int outputColumn = ExcelExtensionSupport.ResolveOutputColumn(arguments, source.Range, 0);
                    ExcelExtensionSupport.EnsureRangeCapacity(
                        source.Range.Row,
                        outputColumn,
                        rows,
                        columns);

                    startCell = (Excel.Range)outputWorksheet.Cells[source.Range.Row, outputColumn];
                    destination = startCell.Resize[rows, columns];
                    string outputAddress = destination.AddressLocal();
                    ExcelUndoScope undo = ExcelUndoScope.Capture(
                        source.Workbook,
                        outputWorksheet,
                        destination,
                        outputAddress);

                    application.Guard(delegate
                    {
                        destination.NumberFormat = "@";
                        destination.Value2 = output;
                    });

                    string targetRef = ExcelComSupport.FullAddress(outputWorksheet.Name, outputAddress);
                    Observation observation = Observation.Create(
                        "rmb_upper",
                        "人民币大写转换完成：成功 " + convertedCount + " 个，已写入 " + targetRef + "。",
                        true);
                    observation.AddTargetRef(targetRef);
                    data["written"] = true;
                    data["outputSheet"] = outputWorksheet.Name;
                    data["outputAddress"] = outputAddress;

                    undo.Complete(
                        source.Workbook,
                        outputWorksheet,
                        destination,
                        outputAddress,
                        true,
                        "人民币大写写入 " + targetRef,
                        observation,
                        data,
                        null);

                    return OperationResult.Succeed(
                        "人民币大写转换完成：成功 " + convertedCount + " 个，已写入 " + targetRef + "。",
                        data,
                        observation);
                }
                finally
                {
                    ComHelper.Release(startCell, destination);
                    if (ownsOutputWorksheet)
                    {
                        ComHelper.Release(outputWorksheet);
                    }
                }
            }
        }

        private static OperationResult ExecuteNumbersAsTextFix(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange", "address");
            if (string.IsNullOrWhiteSpace(address))
            {
                address = "usedrange";
            }

            bool fix = ExcelArgumentReader.GetBool(arguments, true, "fix", "apply");
            bool preserveLeadingZeros = ExcelArgumentReader.GetBool(arguments, true, "preserveLeadingZeros");
            string numberFormat = ExcelArgumentReader.GetString(arguments, "numberFormat");
            if (string.IsNullOrWhiteSpace(numberFormat))
            {
                numberFormat = "General";
            }
            int maxCells = ClampInt(
                ExcelArgumentReader.GetInt(arguments, 100000, "maxCells"),
                1,
                500000);

            Regex skipRegex = null;
            string skipPattern = ExcelArgumentReader.GetString(arguments, "skipPattern");
            if (!string.IsNullOrWhiteSpace(skipPattern))
            {
                try
                {
                    skipRegex = new Regex(skipPattern, RegexOptions.CultureInvariant);
                }
                catch (ArgumentException ex)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, ExcelExtensionSupport.ToRegexPatternError(ex));
                }
            }

            using (ExcelRangeContext context = RangeResolver.Resolve(
                       application,
                       address,
                       false,
                       "excel_numbers_as_text_fix"))
            {
                long totalCells = (long)context.Range.Rows.Count * context.Range.Columns.Count;
                if (totalCells > maxCells)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "目标区域共 " + totalCells + " 个单元格，超过 maxCells=" + maxCells +
                        "。请缩小范围或提高 maxCells（上限 500000）。");
                }
                if (fix)
                {
                    ExcelComSupport.EnsureWritable(
                        application,
                        context.Workbook,
                        context.Worksheet,
                        "excel_numbers_as_text_fix");
                }

                int rows = context.Range.Rows.Count;
                int columns = context.Range.Columns.Count;
                object[,] values = ExcelExtensionSupport.ReadMatrix(context.Range, rows, columns);
                object[,] formats = ExcelComSupport.NormalizeMatrix(
                    context.Range.NumberFormat,
                    rows,
                    columns);
                string rangeAddress = context.Range.AddressLocal();

                JArray samples = new JArray();
                int detectedCount = 0;
                int fixedCount = 0;
                int skippedCount = 0;
                ExcelUndoScope undo = null;
                if (fix)
                {
                    undo = ExcelUndoScope.Capture(
                        context.Workbook,
                        context.Worksheet,
                        context.Range,
                        rangeAddress);
                }

                for (int row = 0; row < rows; row++)
                {
                    for (int column = 0; column < columns; column++)
                    {
                        object value = values[row, column];
                        string cellFormat = ExcelExtensionSupport.CellText(formats[row, column]);
                        bool isTextFormat = cellFormat.IndexOf('@') >= 0;
                        decimal parsed = 0m;
                        bool isInteger = false;
                        string reason = string.Empty;
                        bool detected = false;

                        if (value is string)
                        {
                            string text = (string)value;
                            if (text.Trim().Length == 0)
                            {
                                continue;
                            }
                            if (skipRegex != null && skipRegex.IsMatch(text))
                            {
                                skippedCount++;
                                continue;
                            }
                            if (preserveLeadingZeros && ExcelExtensionSupport.LooksLikeLeadingZero(text))
                            {
                                skippedCount++;
                                continue;
                            }
                            if (!ExcelExtensionSupport.TryParseTextNumber(text, out parsed, out isInteger))
                            {
                                continue;
                            }
                            detected = true;
                            reason = isTextFormat ? "textNumberWithTextFormat" : "textNumber";
                        }
                        else if (isTextFormat &&
                                 (value is double || value is int || value is long || value is decimal))
                        {
                            if (ExcelExtensionSupport.TryParseAmount(value, out parsed))
                            {
                                isInteger = parsed == decimal.Truncate(parsed);
                                detected = true;
                                reason = "numberWithTextFormat";
                            }
                        }

                        if (!detected)
                        {
                            continue;
                        }

                        detectedCount++;
                        if (samples.Count < 50)
                        {
                            JObject item = new JObject();
                            item["address"] = ExcelExtensionSupport.FullCellAddress(
                                context.SheetName,
                                context.Range.Row + row,
                                context.Range.Column + column);
                            item["value"] = ExcelExtensionSupport.CellText(value);
                            item["numberFormat"] = cellFormat;
                            item["reason"] = reason;
                            samples.Add(item);
                        }

                        if (!fix)
                        {
                            continue;
                        }

                        Excel.Range cell = null;
                        try
                        {
                            cell = (Excel.Range)context.Range.Cells[row + 1, column + 1];
                            if (value is string)
                            {
                                cell.Value2 = ExcelExtensionSupport.ToExcelNumber(parsed, isInteger);
                                if (isTextFormat)
                                {
                                    cell.NumberFormat = numberFormat;
                                }
                            }
                            else
                            {
                                cell.NumberFormat = numberFormat;
                            }
                            fixedCount++;
                        }
                        finally
                        {
                            ComHelper.Release(cell);
                        }
                    }
                }

                JObject data = new JObject();
                data["sheet"] = context.SheetName;
                data["address"] = rangeAddress;
                data["detectedCount"] = detectedCount;
                data["fixedCount"] = fixedCount;
                data["skippedCount"] = skippedCount;
                data["fix"] = fix;
                data["preserveLeadingZeros"] = preserveLeadingZeros;
                data["samples"] = samples;

                if (!fix)
                {
                    Observation readObservation = ReadOnlyObservation(
                        "numbers_as_text_fix",
                        "检测完成：发现 " + detectedCount + " 个文本型数字，未修改。");
                    return OperationResult.Succeed(
                        "检测完成：发现 " + detectedCount + " 个文本型数字。",
                        data,
                        readObservation);
                }

                string targetRef = ExcelComSupport.FullAddress(context.SheetName, rangeAddress);
                Observation observation = Observation.Create(
                    "numbers_as_text_fix",
                    "文本型数字修复完成：检测 " + detectedCount + " 个，已修复 " + fixedCount + " 个。",
                    fixedCount > 0);
                if (fixedCount > 0)
                {
                    observation.AddTargetRef(targetRef);
                }

                if (undo != null)
                {
                    undo.Complete(
                        context.Workbook,
                        context.Worksheet,
                        context.Range,
                        rangeAddress,
                        true,
                        "修复文本型数字 " + targetRef,
                        observation,
                        data,
                        fixedCount == 0 ? "没有发生修改，无需撤销。" : null);
                }

                return OperationResult.Succeed(
                    "文本型数字修复完成：检测 " + detectedCount + " 个，已修复 " + fixedCount + " 个。",
                    data,
                    observation);
            }
        }
    }
}
