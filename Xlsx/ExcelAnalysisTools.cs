using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>数据转换、分析、报表与透视表入口。</summary>
    public sealed partial class ExcelToolExecutor
    {
        private static OperationResult ExecuteCreatePivotTable(Excel.Application application, JObject arguments)
        {
            string sourceAddress = ExcelArgumentReader.GetRangeAddress(arguments, "sourceRange", "range", "targetRange");
            if (string.IsNullOrWhiteSpace(sourceAddress))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 sourceRange 参数。");
            }

            using (ExcelRangeContext source = RangeResolver.Resolve(
                       application, sourceAddress, false, "excel_create_pivot_table"))
            {
                ExcelComSupport.EnsureWritable(application, source.Workbook, source.Worksheet, "excel_create_pivot_table");
                JArray rowFields = ExcelArgumentReader.GetArrayOrSingle(arguments, "rowFields");
                JArray columnFields = ExcelArgumentReader.GetArrayOrSingle(arguments, "columnFields");
                JArray valueFields = ExcelArgumentReader.GetArrayOrSingle(arguments, "valueFields");
                if (valueFields == null || valueFields.Count == 0)
                {
                    valueFields = DetectNumericValueFields(source);
                    if (valueFields == null)
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "缺少 valueFields，且数据源中未识别到数值列。");
                    }
                }

                string targetCell = ExcelArgumentReader.GetTrimmedString(arguments, "targetCell", "targetRange", "target");
                if (targetCell.Length == 0)
                {
                    targetCell = "A3";
                }
                string targetSheet = ExcelArgumentReader.GetSheetName(arguments);
                string tableName = ExcelArgumentReader.GetTrimmedString(arguments, "tableName", "name");

                return ExcelPivotService.Create(
                    application,
                    source,
                    targetCell,
                    targetSheet,
                    rowFields,
                    columnFields,
                    valueFields,
                    tableName);
            }
        }

        private static JArray DetectNumericValueFields(ExcelRangeContext source)
        {
            int rows = source.Range.Rows.Count;
            int columns = source.Range.Columns.Count;
            if (rows < 1 || columns < 1)
            {
                return null;
            }
            int sampleRows = Math.Min(rows, 50);
            object[,] values = ExcelComSupport.ReadMatrix(source.Range, sampleRows, columns);
            string[] headers = RangeResolver.GetHeaders(source.Worksheet, source.Range, columns);
            JArray fields = new JArray();
            for (int column = 0; column < columns; column++)
            {
                bool numeric = false;
                for (int row = 1; row < sampleRows; row++)
                {
                    double number;
                    if (TryNumber(values[row, column], out number))
                    {
                        numeric = true;
                        break;
                    }
                }
                if (numeric)
                {
                    string header = headers.Length > column && !string.IsNullOrWhiteSpace(headers[column])
                        ? headers[column]
                        : "列" + ExcelComSupport.ColumnLetter(column + 1);
                    fields.Add(header);
                }
            }
            return fields.Count == 0 ? null : fields;
        }

        private static OperationResult ExecuteTransformData(Excel.Application application, JObject arguments)
        {
            string sourceAddress = ExcelArgumentReader.GetRangeAddress(arguments, "sourceRange", "range", "source");
            if (string.IsNullOrWhiteSpace(sourceAddress))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 sourceRange 参数。");
            }
            string operation = ExcelArgumentReader.GetString(arguments, "operation", "type").Trim().ToLowerInvariant();
            if (operation.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 operation 参数（transpose/split/merge）。");
            }

            using (ExcelRangeContext source = RangeResolver.Resolve(application, sourceAddress, false, "excel_transform_data"))
            {
                if (operation == "transpose")
                {
                    return TransformTranspose(application, source, arguments);
                }
                if (operation == "split")
                {
                    return TransformSplit(application, source, arguments);
                }
                if (operation == "merge")
                {
                    return TransformMerge(application, source, arguments);
                }
                return Fail(OfficeErrorCode.INVALID_PARAMS, "不支持的转换操作：" + operation + "。");
            }
        }

        private static OperationResult TransformTranspose(
            Excel.Application application,
            ExcelRangeContext source,
            JObject arguments)
        {
            int sourceRows = source.Range.Rows.Count;
            int sourceColumns = source.Range.Columns.Count;
            string targetAddress = ExcelArgumentReader.GetTrimmedString(arguments, "targetRange", "target");
            using (ExcelRangeContext target = RangeResolver.ResolveTargetOrDefault(
                       application,
                       targetAddress,
                       source,
                       sourceColumns + 2,
                       false,
                       "excel_transform_data"))
            {
                ExcelComSupport.EnsureWritable(application, target.Workbook, target.Worksheet, "excel_transform_data");
                Excel.Range destination = null;
                try
                {
                    destination = target.Range.Resize[sourceColumns, sourceRows];
                    string targetAddressLocal = destination.AddressLocal();
                    ExcelUndoScope undo = ExcelUndoScope.Capture(
                        target.Workbook,
                        target.Worksheet,
                        destination,
                        targetAddressLocal);

                    application.Guard(delegate
                    {
                        source.Range.Copy();
                        try
                        {
                            destination.PasteSpecial(
                                Paste: Excel.XlPasteType.xlPasteAll,
                                Operation: Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                                SkipBlanks: false,
                                Transpose: true);
                        }
                        finally
                        {
                            application.CutCopyMode = (Excel.XlCutCopyMode)0;
                        }
                    });

                    string targetRef = ExcelComSupport.FullAddress(target.SheetName, targetAddressLocal);
                    Observation observation = Observation.Create(
                        "transform_data",
                        "已将 " + source.FullAddress + " 转置到 " + targetRef + "。",
                        true);
                    observation.AddTargetRef(targetRef);
                    JObject data = new JObject();
                    data["operation"] = "transpose";
                    data["source"] = source.FullAddress;
                    data["sheet"] = target.SheetName;
                    data["address"] = targetAddressLocal;
                    data["rows"] = sourceColumns;
                    data["columns"] = sourceRows;

                    undo.Complete(
                        target.Workbook,
                        target.Worksheet,
                        destination,
                        targetAddressLocal,
                        true,
                        "转置数据到 " + targetRef,
                        observation,
                        data,
                        "转置会粘贴源区域格式；自动撤销恢复值/公式/数字格式，不恢复字体、填充等视觉格式。");

                    return OperationResult.Succeed("数据转置完成：" + targetRef + "。", data, observation);
                }
                finally
                {
                    ComHelper.Release(destination);
                }
            }
        }

        private static OperationResult TransformSplit(
            Excel.Application application,
            ExcelRangeContext source,
            JObject arguments)
        {
            string delimiter = ExcelArgumentReader.GetString(arguments, "delimiter");
            if (string.IsNullOrEmpty(delimiter))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "split 操作缺少 delimiter 参数。");
            }

            int sourceRows = source.Range.Rows.Count;
            int sourceColumns = source.Range.Columns.Count;
            object[,] values = ExcelComSupport.ReadMatrix(source.Range, sourceRows, sourceColumns);
            int maxParts = 1;
            for (int row = 0; row < sourceRows; row++)
            {
                for (int column = 0; column < sourceColumns; column++)
                {
                    string text = Convert.ToString(values[row, column], CultureInfo.InvariantCulture);
                    if (string.IsNullOrEmpty(text))
                    {
                        continue;
                    }
                    int parts = text.Split(new string[] { delimiter }, StringSplitOptions.None).Length;
                    if (parts > maxParts)
                    {
                        maxParts = parts;
                    }
                }
            }

            string targetAddress = ExcelArgumentReader.GetTrimmedString(arguments, "targetRange", "target");
            using (ExcelRangeContext target = string.IsNullOrWhiteSpace(targetAddress)
                       ? RangeResolver.Resolve(application, source.FullAddress, false, "excel_transform_data")
                       : RangeResolver.ResolveTargetOrDefault(
                           application, targetAddress, source, sourceColumns + 2, false, "excel_transform_data"))
            {
                ExcelComSupport.EnsureWritable(application, target.Workbook, target.Worksheet, "excel_transform_data");
                Excel.Range destination = null;
                Excel.Range destinationCell = null;
                try
                {
                    destination = target.Range.Resize[sourceRows, Math.Max(sourceColumns, maxParts)];
                    destinationCell = (Excel.Range)target.Range.Cells[1, 1];
                    string targetAddressLocal = destination.AddressLocal();
                    ExcelUndoScope undo = ExcelUndoScope.Capture(
                        target.Workbook,
                        target.Worksheet,
                        destination,
                        targetAddressLocal);

                    bool isTab = delimiter == "\t";
                    bool isComma = delimiter == ",";
                    bool isSemicolon = delimiter == ";";
                    bool useOther = !isTab && !isComma && !isSemicolon;
                    application.Guard(delegate
                    {
                        source.Range.TextToColumns(
                            Destination: destinationCell,
                            DataType: Excel.XlTextParsingType.xlDelimited,
                            TextQualifier: Excel.XlTextQualifier.xlTextQualifierDoubleQuote,
                            ConsecutiveDelimiter: false,
                            Tab: isTab,
                            Semicolon: isSemicolon,
                            Comma: isComma,
                            Space: false,
                            Other: useOther,
                            OtherChar: useOther ? delimiter : Type.Missing);
                    });

                    string targetRef = ExcelComSupport.FullAddress(target.SheetName, targetAddressLocal);
                    Observation observation = Observation.Create(
                        "transform_data",
                        "已按分隔符「" + delimiter + "」拆分 " + source.FullAddress + " 到 " + targetRef + "。",
                        true);
                    observation.AddTargetRef(targetRef);
                    JObject data = new JObject();
                    data["operation"] = "split";
                    data["source"] = source.FullAddress;
                    data["sheet"] = target.SheetName;
                    data["address"] = targetAddressLocal;
                    data["delimiter"] = delimiter;
                    data["maxParts"] = maxParts;

                    undo.Complete(
                        target.Workbook,
                        target.Worksheet,
                        destination,
                        targetAddressLocal,
                        true,
                        "拆分列 " + targetRef,
                        observation,
                        data,
                        null);

                    return OperationResult.Succeed("分列完成：" + targetRef + "。", data, observation);
                }
                finally
                {
                    ComHelper.Release(destinationCell, destination);
                }
            }
        }

        private static OperationResult TransformMerge(
            Excel.Application application,
            ExcelRangeContext source,
            JObject arguments)
        {
            int sourceRows = source.Range.Rows.Count;
            int sourceColumns = source.Range.Columns.Count;
            string delimiter = ExcelArgumentReader.GetString(arguments, "delimiter");
            if (delimiter.Length == 0)
            {
                delimiter = " ";
            }
            string targetAddress = ExcelArgumentReader.GetTrimmedString(arguments, "targetRange", "target");
            using (ExcelRangeContext target = RangeResolver.ResolveTargetOrDefault(
                       application,
                       targetAddress,
                       source,
                       sourceColumns + 1,
                       false,
                       "excel_transform_data"))
            {
                ExcelComSupport.EnsureWritable(application, target.Workbook, target.Worksheet, "excel_transform_data");
                Excel.Range destination = null;
                try
                {
                    destination = target.Range.Resize[sourceRows, 1];
                    string targetAddressLocal = destination.AddressLocal();
                    ExcelUndoScope undo = ExcelUndoScope.Capture(
                        target.Workbook,
                        target.Worksheet,
                        destination,
                        targetAddressLocal);

                    object[,] formulas = new object[sourceRows, 1];
                    string escapedDelimiter = delimiter.Replace("\"", "\"\"");
                    int sourceStartRow = source.Range.Row;
                    int sourceStartColumn = source.Range.Column;
                    for (int row = 0; row < sourceRows; row++)
                    {
                        StringBuilder formula = new StringBuilder("=CONCAT(");
                        for (int column = 0; column < sourceColumns; column++)
                        {
                            if (column > 0)
                            {
                                formula.Append(",\"").Append(escapedDelimiter).Append("\",");
                            }
                            string cellAddress = ExcelComSupport.ColumnLetter(sourceStartColumn + column)
                                                 + (sourceStartRow + row).ToString(CultureInfo.InvariantCulture);
                            formula.Append(source.SheetName == target.SheetName
                                ? cellAddress
                                : ExcelComSupport.FullAddress(source.SheetName, cellAddress));
                        }
                        formula.Append(")");
                        formulas[row, 0] = formula.ToString();
                    }

                    string engine;
                    ExcelComSupport.SetFormula(destination, formulas, out engine);

                    string targetRef = ExcelComSupport.FullAddress(target.SheetName, targetAddressLocal);
                    Observation observation = Observation.Create(
                        "transform_data",
                        "已使用 CONCAT 合并 " + source.FullAddress + " 到 " + targetRef + "。",
                        true);
                    observation.AddTargetRef(targetRef);
                    JObject data = new JObject();
                    data["operation"] = "merge";
                    data["source"] = source.FullAddress;
                    data["sheet"] = target.SheetName;
                    data["address"] = targetAddressLocal;
                    data["delimiter"] = delimiter;
                    data["formulaEngine"] = engine;

                    undo.Complete(
                        target.Workbook,
                        target.Worksheet,
                        destination,
                        targetAddressLocal,
                        true,
                        "合并列到 " + targetRef,
                        observation,
                        data,
                        null);

                    return OperationResult.Succeed("列合并完成：" + targetRef + "。", data, observation);
                }
                finally
                {
                    ComHelper.Release(destination);
                }
            }
        }

        private static OperationResult ExecuteAnalyzeData(Excel.Application application, JObject arguments)
        {
            string sourceAddress = ExcelArgumentReader.GetRangeAddress(arguments, "sourceRange", "range", "source");
            if (string.IsNullOrWhiteSpace(sourceAddress))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 sourceRange 参数。");
            }
            string type = ExcelArgumentReader.GetString(arguments, "type", "analysisType").Trim().ToLowerInvariant();
            if (type.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 type 参数（summary/pivot/groupby/ranking）。");
            }

            using (ExcelRangeContext source = RangeResolver.Resolve(application, sourceAddress, false, "excel_analyze_data"))
            {
                if (type == "pivot")
                {
                    string targetCell = ExcelArgumentReader.GetTrimmedString(arguments, "targetRange", "targetCell");
                    if (targetCell.Length == 0)
                    {
                        targetCell = "A3";
                    }
                    string targetSheet = ExcelArgumentReader.GetSheetName(arguments);
                    return ExcelPivotService.Create(
                        application,
                        source,
                        targetCell,
                        targetSheet,
                        ExcelArgumentReader.GetArrayOrSingle(arguments, "rowFields"),
                        ExcelArgumentReader.GetArrayOrSingle(arguments, "columnFields"),
                        ExcelArgumentReader.GetArrayOrSingle(arguments, "valueFields"),
                        ExcelArgumentReader.GetTrimmedString(arguments, "tableName"));
                }

                return AnalyzeWithMemory(application, source, arguments, type);
            }
        }

        private static OperationResult AnalyzeWithMemory(
            Excel.Application application,
            ExcelRangeContext source,
            JObject arguments,
            string type)
        {
            int rows = source.Range.Rows.Count;
            int columns = source.Range.Columns.Count;
            object[,] values = ExcelComSupport.ReadMatrix(source.Range, rows, columns);
            bool hasHeader = ExcelArgumentReader.GetBool(arguments, true, "hasHeader");
            string[] headers = RangeResolver.GetHeaders(source.Worksheet, source.Range, columns);
            int dataStart = hasHeader ? 1 : 0;
            string targetAddress = ExcelArgumentReader.GetTrimmedString(arguments, "targetRange", "targetCell");

            if (type == "summary")
            {
                List<object[]> output = new List<object[]>();
                output.Add(new object[] { "数据摘要" });
                output.Add(new object[] { "行数", rows });
                output.Add(new object[] { "列数", columns });
                output.Add(new object[] { "数据行数", Math.Max(0, rows - dataStart) });
                output.Add(new object[] { "指标", "合计", "平均", "数值个数" });
                JArray numericColumns = new JArray();

                for (int column = 0; column < columns; column++)
                {
                    double sum = 0;
                    int count = 0;
                    for (int row = dataStart; row < rows; row++)
                    {
                        double number;
                        if (TryNumber(values[row, column], out number))
                        {
                            sum += number;
                            count++;
                        }
                    }
                    if (count == 0)
                    {
                        continue;
                    }
                    string header = HeaderName(headers, column);
                    output.Add(new object[] { header + " 合计", sum });
                    output.Add(new object[] { header + " 平均", sum / count });
                    output.Add(new object[] { header + " 数值个数", count });
                    JObject stat = new JObject();
                    stat["column"] = ExcelComSupport.ColumnLetter(source.Range.Column + column);
                    stat["header"] = header;
                    stat["sum"] = sum;
                    stat["average"] = sum / count;
                    stat["count"] = count;
                    numericColumns.Add(stat);
                }

                JObject data = new JObject();
                data["type"] = "summary";
                data["source"] = source.FullAddress;
                data["rowCount"] = rows;
                data["columnCount"] = columns;
                data["numericColumns"] = numericColumns;
                return WriteAnalysisOutput(
                    application,
                    source,
                    targetAddress,
                    output,
                    "analyze_data",
                    "已生成数据摘要，源区域 " + source.FullAddress + "。",
                    data);
            }

            if (type == "groupby")
            {
                JToken groupToken = ExcelArgumentReader.First(arguments, "groupBy", "groupField");
                JToken valueToken = ExcelArgumentReader.First(arguments, "valueField", "valueBy");
                if (groupToken == null || valueToken == null)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "groupby 需要 groupBy/groupField 与 valueField。");
                }
                int groupColumn = ExcelComSupport.ResolveColumnIndex(arguments, groupToken, headers, hasHeader, 0);
                int valueColumn = ExcelComSupport.ResolveColumnIndex(arguments, valueToken, headers, hasHeader, 0);
                if (groupColumn < 1 || groupColumn > columns || valueColumn < 1 || valueColumn > columns)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "groupBy/valueField 列不存在。");
                }

                string aggregate = ExcelArgumentReader.GetString(arguments, "aggregate").Trim().ToLowerInvariant();
                if (aggregate.Length == 0)
                {
                    aggregate = "sum";
                }
                if (aggregate != "sum" && aggregate != "count" && aggregate != "avg" && aggregate != "average")
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "aggregate 仅支持 sum/count/avg。");
                }

                Dictionary<string, double> sums = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
                Dictionary<string, int> counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                List<string> order = new List<string>();
                for (int row = dataStart; row < rows; row++)
                {
                    string key = Convert.ToString(values[row, groupColumn - 1], CultureInfo.InvariantCulture) ?? string.Empty;
                    if (key.Length == 0)
                    {
                        key = "(空)";
                    }
                    double number;
                    TryNumber(values[row, valueColumn - 1], out number);
                    if (!sums.ContainsKey(key))
                    {
                        sums[key] = 0;
                        counts[key] = 0;
                        order.Add(key);
                    }
                    sums[key] += number;
                    counts[key]++;
                }

                List<object[]> output = new List<object[]>();
                output.Add(new object[] { "分组汇总" });
                output.Add(new object[] { HeaderName(headers, groupColumn - 1), aggregate, "数量" });
                foreach (string key in order)
                {
                    object aggregateValue;
                    if (aggregate == "count")
                    {
                        aggregateValue = counts[key];
                    }
                    else if (aggregate == "avg" || aggregate == "average")
                    {
                        aggregateValue = counts[key] == 0 ? 0 : sums[key] / counts[key];
                    }
                    else
                    {
                        aggregateValue = sums[key];
                    }
                    output.Add(new object[] { key, aggregateValue, counts[key] });
                }

                JObject data = new JObject();
                data["type"] = "groupby";
                data["source"] = source.FullAddress;
                data["groupBy"] = groupToken.ToString();
                data["valueField"] = valueToken.ToString();
                data["aggregate"] = aggregate;
                data["groupCount"] = order.Count;
                return WriteAnalysisOutput(
                    application,
                    source,
                    targetAddress,
                    output,
                    "analyze_data",
                    "已生成分组汇总，共 " + order.Count + " 组。",
                    data);
            }

            if (type == "ranking")
            {
                JToken labelToken = ExcelArgumentReader.First(arguments, "labelField", "labelBy");
                JToken valueToken = ExcelArgumentReader.First(arguments, "rankBy", "valueField");
                if (labelToken == null || valueToken == null)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "ranking 需要 labelField 与 rankBy/valueField。");
                }
                int labelColumn = ExcelComSupport.ResolveColumnIndex(arguments, labelToken, headers, hasHeader, 0);
                int valueColumn = ExcelComSupport.ResolveColumnIndex(arguments, valueToken, headers, hasHeader, 0);
                if (labelColumn < 1 || labelColumn > columns || valueColumn < 1 || valueColumn > columns)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "labelField/rankBy 列不存在。");
                }
                bool descending = ExcelArgumentReader.GetBool(arguments, true, "descending");
                int topN = ExcelArgumentReader.GetInt(arguments, 0, "topN", "limit");

                List<RankingItem> items = new List<RankingItem>();
                for (int row = dataStart; row < rows; row++)
                {
                    double number;
                    if (!TryNumber(values[row, valueColumn - 1], out number))
                    {
                        continue;
                    }
                    string label = Convert.ToString(values[row, labelColumn - 1], CultureInfo.InvariantCulture) ?? string.Empty;
                    if (label.Length == 0)
                    {
                        label = "第" + (row + 1).ToString(CultureInfo.InvariantCulture) + "行";
                    }
                    RankingItem item = new RankingItem();
                    item.Label = label;
                    item.Value = number;
                    items.Add(item);
                }

                items.Sort(delegate(RankingItem left, RankingItem right)
                {
                    return descending ? right.Value.CompareTo(left.Value) : left.Value.CompareTo(right.Value);
                });
                int take = topN > 0 ? Math.Min(topN, items.Count) : items.Count;

                List<object[]> output = new List<object[]>();
                output.Add(new object[] { "排名分析" });
                output.Add(new object[] { "排名", HeaderName(headers, labelColumn - 1), HeaderName(headers, valueColumn - 1) });
                for (int index = 0; index < take; index++)
                {
                    output.Add(new object[] { index + 1, items[index].Label, items[index].Value });
                }

                JObject data = new JObject();
                data["type"] = "ranking";
                data["source"] = source.FullAddress;
                data["labelField"] = labelToken.ToString();
                data["rankBy"] = valueToken.ToString();
                data["descending"] = descending;
                data["topN"] = topN;
                data["returned"] = take;
                return WriteAnalysisOutput(
                    application,
                    source,
                    targetAddress,
                    output,
                    "analyze_data",
                    "已生成排名分析，共 " + take + " 条。",
                    data);
            }

            return Fail(OfficeErrorCode.INVALID_PARAMS, "不支持的 type：" + type + "。");
        }

        private static OperationResult WriteAnalysisOutput(
            Excel.Application application,
            ExcelRangeContext source,
            string targetAddress,
            List<object[]> output,
            string kind,
            string summary,
            JObject data)
        {
            if (output == null || output.Count == 0)
            {
                return Fail(OfficeErrorCode.INTERNAL_ERROR, "分析结果为空。");
            }
            int outputColumns = 0;
            foreach (object[] row in output)
            {
                if (row != null && row.Length > outputColumns)
                {
                    outputColumns = row.Length;
                }
            }

            using (ExcelRangeContext target = RangeResolver.ResolveTargetOrDefault(
                       application,
                       targetAddress,
                       source,
                       source.Range.Columns.Count + 2,
                       false,
                       "excel_analyze_data"))
            {
                ExcelComSupport.EnsureWritable(application, target.Workbook, target.Worksheet, "excel_analyze_data");
                Excel.Range writeRange = null;
                Excel.Range titleCell = null;
                try
                {
                    writeRange = target.Range.Resize[output.Count, outputColumns];
                    string targetAddressLocal = writeRange.AddressLocal();
                    ExcelUndoScope undo = ExcelUndoScope.Capture(
                        target.Workbook,
                        target.Worksheet,
                        writeRange,
                        targetAddressLocal);

                    object[,] matrix = new object[output.Count, outputColumns];
                    for (int row = 0; row < output.Count; row++)
                    {
                        object[] sourceRow = output[row];
                        for (int column = 0; column < outputColumns; column++)
                        {
                            matrix[row, column] = sourceRow != null && column < sourceRow.Length
                                ? sourceRow[column]
                                : null;
                        }
                    }
                    writeRange.Value2 = matrix;

                    titleCell = (Excel.Range)writeRange.Cells[1, 1];
                    Excel.Font titleFont = null;
                    try
                    {
                        titleFont = titleCell.Font;
                        titleFont.Bold = true;
                    }
                    finally
                    {
                        ComHelper.Release(titleFont);
                    }

                    string targetRef = ExcelComSupport.FullAddress(target.SheetName, targetAddressLocal);
                    Observation observation = Observation.Create(kind, summary + " 输出到 " + targetRef + "。", true);
                    observation.AddTargetRef(targetRef);
                    data["sheet"] = target.SheetName;
                    data["address"] = targetAddressLocal;
                    data["outputRows"] = output.Count;
                    data["outputColumns"] = outputColumns;

                    undo.Complete(
                        target.Workbook,
                        target.Worksheet,
                        writeRange,
                        targetAddressLocal,
                        true,
                        summary + " 输出 " + targetRef,
                        observation,
                        data,
                        null);

                    return OperationResult.Succeed(summary + " 输出到 " + targetRef + "。", data, observation);
                }
                finally
                {
                    ComHelper.Release(titleCell, writeRange);
                }
            }
        }

        private static OperationResult ExecuteGenerateReport(Excel.Application application, JObject arguments)
        {
            string sourceAddress = ExcelArgumentReader.GetRangeAddress(arguments, "sourceRange", "range", "source");
            if (string.IsNullOrWhiteSpace(sourceAddress))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 sourceRange 参数。");
            }
            string title = ExcelArgumentReader.GetString(arguments, "title");
            bool includeChart = ExcelArgumentReader.GetBool(arguments, false, "includeChart");
            string requestedSheet = ExcelArgumentReader.GetTrimmedString(arguments, "targetSheet", "sheetName");

            using (ExcelRangeContext source = RangeResolver.Resolve(application, sourceAddress, false, "excel_generate_report"))
            {
                Excel.Worksheet targetSheet = null;
                bool createdSheet = false;
                Excel.Range usedRange = null;
                Excel.Range copyDestination = null;
                try
                {
                    string reportSheetName = requestedSheet;
                    if (reportSheetName.Length > 0)
                    {
                        targetSheet = ExcelComSupport.FindWorksheet(source.Workbook, reportSheetName);
                        if (targetSheet == null)
                        {
                            targetSheet = RangeResolver.CreateWorksheet(source.Workbook, reportSheetName);
                            createdSheet = true;
                        }
                    }
                    else
                    {
                        reportSheetName = ExcelPivotService.UniqueSheetName(
                            source.Workbook,
                            "报表_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture));
                        targetSheet = RangeResolver.CreateWorksheet(source.Workbook, reportSheetName);
                        createdSheet = true;
                    }

                    ExcelComSupport.EnsureWritable(application, source.Workbook, targetSheet, "excel_generate_report");
                    int sourceRows = source.Range.Rows.Count;
                    int sourceColumns = source.Range.Columns.Count;
                    int dataStartRow = string.IsNullOrWhiteSpace(title) ? 1 : 3;
                    int endRow = dataStartRow + sourceRows - 1;
                    string lastColumn = ExcelComSupport.ColumnLetter(sourceColumns);

                    usedRange = targetSheet.Range["A1"].Resize[endRow, sourceColumns];
                    string usedAddress = usedRange.AddressLocal();
                    ExcelUndoScope undo = createdSheet
                        ? ExcelUndoScope.Failed("报表写入新建工作表，执行前不存在可快照区域。")
                        : ExcelUndoScope.Capture(source.Workbook, targetSheet, usedRange, usedAddress);

                    copyDestination = targetSheet.Range["A" + dataStartRow.ToString(CultureInfo.InvariantCulture)];
                    application.Guard(delegate
                    {
                        if (!string.IsNullOrWhiteSpace(title))
                        {
                            Excel.Range titleRange = null;
                            Excel.Font titleFont = null;
                            try
                            {
                                titleRange = targetSheet.Range["A1"];
                                titleRange.Value2 = title;
                                titleFont = titleRange.Font;
                                titleFont.Size = 16;
                                titleFont.Bold = true;
                            }
                            finally
                            {
                                ComHelper.Release(titleFont, titleRange);
                            }
                        }
                        source.Range.Copy(copyDestination);
                        application.CutCopyMode = (Excel.XlCutCopyMode)0;

                        Excel.Range headerRange = null;
                        Excel.Range dataRange = null;
                        Excel.Range targetColumns = null;
                        try
                        {
                            headerRange = targetSheet.Range[
                                "A" + dataStartRow.ToString(CultureInfo.InvariantCulture) + ":" +
                                lastColumn + dataStartRow.ToString(CultureInfo.InvariantCulture)];
                            Excel.Font headerFont = null;
                            Excel.Interior headerInterior = null;
                            try
                            {
                                headerFont = headerRange.Font;
                                headerFont.Bold = true;
                                headerFont.Color = ExcelComSupport.ToOleColor(255, 255, 255);
                                headerInterior = headerRange.Interior;
                                headerInterior.Color = ExcelComSupport.ToOleColor(68, 114, 196);
                            }
                            finally
                            {
                                ComHelper.Release(headerFont, headerInterior);
                            }

                            dataRange = targetSheet.Range[
                                "A" + dataStartRow.ToString(CultureInfo.InvariantCulture) + ":" +
                                lastColumn + endRow.ToString(CultureInfo.InvariantCulture)];
                            ExcelComSupport.ApplyBorder(dataRange, "all");
                            targetColumns = usedRange.Columns;
                            targetColumns.AutoFit();
                        }
                        finally
                        {
                            ComHelper.Release(targetColumns, dataRange, headerRange);
                        }
                        targetSheet.Activate();
                    });

                    JObject data = new JObject();
                    data["source"] = source.FullAddress;
                    data["sheet"] = targetSheet.Name;
                    data["dataRange"] = ExcelComSupport.FullAddress(
                        targetSheet.Name,
                        "A" + dataStartRow.ToString(CultureInfo.InvariantCulture) + ":" + lastColumn + endRow.ToString(CultureInfo.InvariantCulture));
                    data["title"] = title;
                    data["includeChart"] = includeChart;
                    data["createdSheet"] = createdSheet;

                    Observation observation = Observation.Create(
                        "generate_report",
                        "已生成报表工作表「" + targetSheet.Name + "」。",
                        true);
                    observation.AddTargetRef(ExcelComSupport.FullAddress(targetSheet.Name, usedAddress));

                    if (includeChart)
                    {
                        string chartRange = ExcelComSupport.FullAddress(
                            targetSheet.Name,
                            "A" + dataStartRow.ToString(CultureInfo.InvariantCulture) + ":" + lastColumn + endRow.ToString(CultureInfo.InvariantCulture));
                        JObject chartArguments = new JObject();
                        chartArguments["type"] = ExcelArgumentReader.GetString(arguments, "chartType", "type");
                        if (string.IsNullOrWhiteSpace(chartArguments.Value<string>("type")))
                        {
                            chartArguments["type"] = "column";
                        }
                        chartArguments["title"] = string.IsNullOrWhiteSpace(title) ? "报表图表" : title;
                        chartArguments["position"] = "A" + (endRow + 2).ToString(CultureInfo.InvariantCulture);
                        JObject chartInfo = CreateChartCore(application, chartRange, chartArguments);
                        data["chart"] = chartInfo;
                        observation.AddTargetRef(
                            ExcelComSupport.FullAddress(
                                chartInfo.Value<string>("sheet"),
                                chartInfo.Value<string>("chartName")));
                    }

                    string undoSummary = "生成报表 " + targetSheet.Name;
                    undo.Complete(
                        source.Workbook,
                        targetSheet,
                        usedRange,
                        usedAddress,
                        false,
                        undoSummary,
                        observation,
                        data,
                        createdSheet
                            ? "报表写入新建工作表，无法通过区域快照撤销整个新表。"
                            : "报表复制了源区域格式；自动撤销仅恢复值/公式/数字格式，不恢复视觉格式。");

                    return OperationResult.Succeed(
                        "报表已生成：" + targetSheet.Name + "。",
                        data,
                        observation);
                }
                catch
                {
                    if (createdSheet && targetSheet != null)
                    {
                        ExcelPivotService.TryDeleteWorksheet(application, targetSheet);
                    }
                    throw;
                }
                finally
                {
                    ComHelper.Release(copyDestination, usedRange, targetSheet);
                }
            }
        }

        private static string HeaderName(string[] headers, int zeroBasedColumn)
        {
            if (headers != null && zeroBasedColumn >= 0 && zeroBasedColumn < headers.Length &&
                !string.IsNullOrWhiteSpace(headers[zeroBasedColumn]))
            {
                return headers[zeroBasedColumn];
            }
            return "列" + ExcelComSupport.ColumnLetter(zeroBasedColumn + 1);
        }

        private static bool TryNumber(object value, out double number)
        {
            number = 0;
            if (value == null || value is DBNull)
            {
                return false;
            }
            if (value is double)
            {
                number = (double)value;
                return true;
            }
            if (value is int)
            {
                number = (int)value;
                return true;
            }
            if (value is long)
            {
                number = (long)value;
                return true;
            }
            if (value is decimal)
            {
                number = (double)(decimal)value;
                return true;
            }
            if (value is DateTime)
            {
                return false;
            }
            return double.TryParse(
                Convert.ToString(value, CultureInfo.InvariantCulture),
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out number);
        }

        private sealed class RankingItem
        {
            public string Label { get; set; }

            public double Value { get; set; }
        }
    }
}
