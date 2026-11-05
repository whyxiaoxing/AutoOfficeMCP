using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Office;
using Office.Undo;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>基础写入、公式、清洗与清除。</summary>
    public sealed partial class ExcelToolExecutor
    {
        private static OperationResult ExecuteWriteData(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(
                arguments,
                "targetRange", "startCell", "range", "address", "cell");
            if (string.IsNullOrWhiteSpace(address))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 targetRange/startCell/range 参数。");
            }

            JToken dataToken = ExcelArgumentReader.First(arguments, "data", "targetData", "values");
            if (dataToken == null)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 data/targetData 参数。");
            }

            bool createSheet = ExcelArgumentReader.GetBool(arguments, false, "createSheet", "autoCreateSheet");
            string numberFormat = ExcelArgumentReader.GetString(arguments, "numberFormat", "format");
            int rows;
            int columns;
            bool scalar;
            object[,] matrix;
            try
            {
                matrix = ExcelComSupport.JsonToMatrix(dataToken, out rows, out columns, out scalar);
            }
            catch (OfficeComException ex)
            {
                return Fail(ex.Code, ex.Message);
            }

            using (ExcelRangeContext context = RangeResolver.Resolve(
                       application, address, createSheet, "excel_write_data"))
            {
                ExcelComSupport.EnsureWritable(application, context.Workbook, context.Worksheet, "excel_write_data");

                Excel.Range writeRange = context.Range;
                bool ownsWriteRange = false;
                bool merged = IsMerged(context.Range);
                try
                {
                    if (merged)
                    {
                        if (!scalar && (rows > 1 || columns > 1))
                        {
                            return Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "目标区域命中合并单元格；二维数组不能写入合并区。请改用单值写入合并区，或先取消合并。");
                        }
                        writeRange = context.Range.MergeArea;
                        ownsWriteRange = true;
                    }
                    else if (!scalar)
                    {
                        writeRange = context.Range.Resize[rows, columns];
                        ownsWriteRange = true;
                    }

                    string writeAddress = writeRange.AddressLocal();
                    ExcelUndoScope undo = ExcelUndoScope.Capture(
                        context.Workbook,
                        context.Worksheet,
                        writeRange,
                        writeAddress);

                    if (!string.IsNullOrWhiteSpace(numberFormat))
                    {
                        writeRange.NumberFormat = numberFormat;
                    }

                    object writeValue = matrix[0, 0];
                    if (scalar)
                    {
                        if (merged)
                        {
                            Excel.Range firstCell = null;
                            try
                            {
                                firstCell = (Excel.Range)writeRange.Cells[1, 1];
                                firstCell.Value2 = writeValue;
                            }
                            finally
                            {
                                ComHelper.Release(firstCell);
                            }
                        }
                        else
                        {
                            writeRange.Value2 = writeValue;
                        }
                    }
                    else
                    {
                        writeRange.Value2 = matrix;
                    }

                    string targetRef = ExcelComSupport.FullAddress(context.SheetName, writeAddress);
                    Observation observation = Observation.Create(
                        "write_data",
                        "已向 " + targetRef + " 写入 " + rows.ToString(CultureInfo.InvariantCulture) + " 行 × " +
                        columns.ToString(CultureInfo.InvariantCulture) + " 列数据。",
                        true);
                    observation.AddTargetRef(targetRef);

                    JObject data = new JObject();
                    data["sheet"] = context.SheetName;
                    data["address"] = writeAddress;
                    data["rows"] = rows;
                    data["columns"] = columns;
                    data["merged"] = merged;
                    data["numberFormat"] = string.IsNullOrWhiteSpace(numberFormat) ? string.Empty : numberFormat;

                    undo.Complete(
                        context.Workbook,
                        context.Worksheet,
                        writeRange,
                        writeAddress,
                        true,
                        "写入数据 " + targetRef,
                        observation,
                        data,
                        null);

                    return OperationResult.Succeed(
                        "已写入 " + targetRef + "（" + rows + " 行 × " + columns + " 列）。",
                        data,
                        observation);
                }
                finally
                {
                    if (ownsWriteRange)
                    {
                        ComHelper.Release(writeRange);
                    }
                }
            }
        }

        private static OperationResult ExecuteApplyFormula(Excel.Application application, JObject arguments)
        {
            string targetAddress = ExcelArgumentReader.GetRangeAddress(
                arguments,
                "targetRange", "range", "target", "startCell");
            string formula = ExcelArgumentReader.GetString(arguments, "formula");
            bool fillDown = ExcelArgumentReader.GetBool(arguments, false, "fillDown", "autoFill");
            JArray operations = ExcelArgumentReader.GetArray(arguments, "operations");

            if (operations == null && string.IsNullOrWhiteSpace(formula))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 formula 或 operations 参数。");
            }

            string baseSheet = ResolveBaseSheetName(application, targetAddress);
            if (string.IsNullOrWhiteSpace(baseSheet))
            {
                return Fail(OfficeErrorCode.DOC_NOT_OPEN, "当前没有活动工作表。");
            }

            if (operations == null || operations.Count == 0)
            {
                JObject single = new JObject();
                single["range"] = targetAddress;
                single["formula"] = formula;
                single["fillDown"] = fillDown;
                return ApplyFormulaBatch(application, baseSheet, new JArray(single), fillDown);
            }

            return ApplyFormulaBatch(application, baseSheet, operations, fillDown);
        }

        private static OperationResult ApplyFormulaBatch(
            Excel.Application application,
            string baseSheet,
            JArray operations,
            bool defaultFillDown)
        {
            Observation observation = Observation.Create("apply_formula", "已应用 Excel 公式。", false);
            JObject data = new JObject();
            JArray applied = new JArray();
            int succeeded = 0;
            string lastError = string.Empty;
            OfficeErrorCode failureCode = OfficeErrorCode.INVALID_PARAMS;

            for (int index = 0; index < operations.Count; index++)
            {
                JObject operation = operations[index] as JObject;
                if (operation == null)
                {
                    lastError = "operations[" + index + "] 不是对象。";
                    break;
                }

                string rangeAddress = ExcelArgumentReader.GetTrimmedString(operation, "range", "targetRange", "target", "startCell");
                string formulaText = ExcelArgumentReader.GetString(operation, "formula");
                bool fillDown = ExcelArgumentReader.GetBool(operation, defaultFillDown, "fillDown", "autoFill");
                if (string.IsNullOrWhiteSpace(formulaText))
                {
                    lastError = "operations[" + index + "] 缺少 formula。";
                    break;
                }

                try
                {
                    FormulaApplication result = ApplySingleFormula(
                        application,
                        baseSheet,
                        rangeAddress,
                        formulaText,
                        fillDown,
                        observation);
                    applied.Add(result.ToJObject());
                    succeeded++;
                }
                catch (Exception ex)
                {
                    failureCode = OfficeErrorMapper.MapException(ex, OfficeApp.Excel);
                    lastError = OfficeErrorMapper.DescribeException(ex, OfficeApp.Excel);
                    break;
                }
            }

            data["operations"] = applied;
            data["succeededCount"] = succeeded;
            data["fillDown"] = defaultFillDown;
            observation.Changed = succeeded > 0;
            observation.Summary = "已应用 " + succeeded.ToString(CultureInfo.InvariantCulture) + " 个公式操作。";

            if (succeeded == operations.Count)
            {
                return OperationResult.Succeed(
                    "已应用 " + succeeded.ToString(CultureInfo.InvariantCulture) + " 个公式操作。",
                    data,
                    observation);
            }

            if (succeeded == 0)
            {
                return OperationResult.Failed(
                    failureCode,
                    "公式应用失败：" + lastError,
                    observation,
                    data,
                    "失败发生在第一个操作，未产生已知改动。");
            }

            return OperationResult.PartialApplied(
                succeeded,
                "前 " + succeeded + " 个公式已生效，后续操作失败：" + lastError +
                " 请使用返回的 undoUnits 撤销已执行部分，或人工核对后重试。",
                observation,
                data);
        }

        private static FormulaApplication ApplySingleFormula(
            Excel.Application application,
            string baseSheet,
            string targetAddress,
            string formulaText,
            bool fillDown,
            Observation observation)
        {
            string normalizedFormula = formulaText.Trim();
            string resolvedTarget = targetAddress;
            if (string.IsNullOrWhiteSpace(resolvedTarget) && normalizedFormula.IndexOf('=') > 0)
            {
                int equals = normalizedFormula.IndexOf('=');
                string left = normalizedFormula.Substring(0, equals).Trim();
                if (Regex.IsMatch(left, @"^(?:[A-Za-z]{1,3}\d+)(?::[A-Za-z]{1,3}\d+)?$"))
                {
                    resolvedTarget = left;
                    normalizedFormula = normalizedFormula.Substring(equals + 1).Trim();
                }
            }
            if (string.IsNullOrWhiteSpace(resolvedTarget))
            {
                resolvedTarget = "C1";
            }
            if (!normalizedFormula.StartsWith("=", StringComparison.Ordinal))
            {
                normalizedFormula = "=" + normalizedFormula;
            }

            using (ExcelRangeContext context = RangeResolver.ResolveOnSheet(
                       application,
                       baseSheet,
                       resolvedTarget,
                       false,
                       "excel_apply_formula"))
            {
                ExcelComSupport.EnsureWritable(application, context.Workbook, context.Worksheet, "excel_apply_formula");
                string address = context.Range.AddressLocal();
                ExcelUndoScope undo = ExcelUndoScope.Capture(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    address);

                string engine;
                int rowCount = context.Range.Rows.Count;
                if (fillDown && rowCount > 1)
                {
                    Excel.Range firstCell = null;
                    try
                    {
                        firstCell = (Excel.Range)context.Range.Cells[1, 1];
                        ExcelComSupport.SetFormula(firstCell, normalizedFormula, out engine);
                        firstCell.AutoFill(context.Range, Excel.XlAutoFillType.xlFillDefault);
                    }
                    finally
                    {
                        ComHelper.Release(firstCell);
                    }
                }
                else
                {
                    ExcelComSupport.SetFormula(context.Range, normalizedFormula, out engine);
                }

                string targetRef = ExcelComSupport.FullAddress(context.SheetName, address);
                observation.AddTargetRef(targetRef);
                JObject undoData = new JObject();
                undo.Complete(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    address,
                    true,
                    "应用公式 " + targetRef,
                    observation,
                    undoData,
                    null);

                FormulaApplication result = new FormulaApplication();
                result.Range = targetRef;
                result.Address = address;
                result.Formula = normalizedFormula;
                result.FormulaEngine = engine;
                result.FillDown = fillDown;
                result.Rows = rowCount;
                result.UndoRestorable = undoData.Value<bool?>("undoRestorable") ?? false;
                result.UndoReason = undoData.Value<string>("undoReason");
                return result;
            }
        }

        private static OperationResult ExecuteCleanData(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange");
            if (string.IsNullOrWhiteSpace(address))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 range 参数。");
            }

            string operation = ExcelArgumentReader.GetTrimmedString(arguments, "operation", "type").ToLowerInvariant();
            if (operation.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 operation 参数（removeduplicates/fillempty/trim/replace）。");
            }

            bool createSheet = ExcelArgumentReader.GetBool(arguments, false, "createSheet");
            using (ExcelRangeContext context = RangeResolver.Resolve(
                       application, address, createSheet, "excel_clean_data"))
            {
                ExcelComSupport.EnsureWritable(application, context.Workbook, context.Worksheet, "excel_clean_data");
                string rangeAddress = context.Range.AddressLocal();
                ExcelUndoScope undo = ExcelUndoScope.Capture(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress);

                int affected = 0;
                bool structureStable = true;
                string effectiveOperation = operation;
                if (operation == "removeduplicates" || operation == "remove_duplicates")
                {
                    effectiveOperation = "removeduplicates";
                    object columnsObject = BuildAllColumnIndexes(context.Range.Columns.Count);
                    bool hasHeader = ExcelArgumentReader.GetBool(arguments, true, "hasHeader");
                    context.Range.RemoveDuplicates(
                        columnsObject,
                        hasHeader ? Excel.XlYesNoGuess.xlYes : Excel.XlYesNoGuess.xlNo);
                    structureStable = false;
                }
                else if (operation == "fillempty" || operation == "fill_empty" || operation == "fillblank")
                {
                    effectiveOperation = "fillempty";
                    string fillValue = ExcelArgumentReader.GetString(arguments, "fillValue");
                    if (fillValue.Length == 0)
                    {
                        fillValue = "0";
                    }
                    affected = FillEmptyCells(context.Range, fillValue);
                }
                else if (operation == "trim")
                {
                    affected = TrimTextCells(context.Range);
                }
                else if (operation == "replace")
                {
                    string findText = ExcelArgumentReader.GetString(arguments, "findText", "find");
                    string replaceText = ExcelArgumentReader.GetString(arguments, "replaceText", "replace");
                    if (findText.Length == 0)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "replace 操作缺少 findText。");
                    }
                    bool matchCase = ExcelArgumentReader.GetBool(arguments, false, "matchCase");
                    int rows = context.Range.Rows.Count;
                    int columns = context.Range.Columns.Count;
                    object[,] values = ExcelComSupport.ReadMatrix(context.Range, rows, columns);
                    affected = ExcelComSupport.CountMatches(values, findText, matchCase, false);
                    if (affected > 0)
                    {
                        context.Range.Replace(
                            What: findText,
                            Replacement: replaceText,
                            LookAt: Excel.XlLookAt.xlPart,
                            MatchCase: matchCase);
                    }
                }
                else
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "不支持的清洗操作：" + operation + "。可用 removeduplicates/fillempty/trim/replace。");
                }

                string targetRef = ExcelComSupport.FullAddress(context.SheetName, rangeAddress);
                Observation observation = Observation.Create(
                    "clean_data",
                    "数据清洗 " + effectiveOperation + " 已完成" +
                    (affected > 0 ? "，影响 " + affected + " 个单元格。" : "。"),
                    affected > 0 || effectiveOperation == "removeduplicates");
                observation.AddTargetRef(targetRef);

                JObject data = new JObject();
                data["sheet"] = context.SheetName;
                data["address"] = rangeAddress;
                data["operation"] = effectiveOperation;
                data["affectedCells"] = affected;

                undo.Complete(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress,
                    structureStable,
                    "清洗数据 " + targetRef,
                    observation,
                    data,
                    structureStable
                        ? null
                        : "删除重复行会改变工作表结构，区域快照无法完整逆操作，未登记自动撤销。");

                return OperationResult.Succeed(
                    "数据清洗完成：" + effectiveOperation + "，目标 " + targetRef + "。",
                    data,
                    observation);
            }
        }

        private static OperationResult ExecuteClearRange(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange");
            if (string.IsNullOrWhiteSpace(address))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 range 参数。");
            }
            string scope = ExcelArgumentReader.GetTrimmedString(arguments, "scope").ToLowerInvariant();
            if (scope.Length == 0)
            {
                scope = "contents";
            }

            using (ExcelRangeContext context = RangeResolver.Resolve(application, address, false, "excel_clear_range"))
            {
                ExcelComSupport.EnsureWritable(application, context.Workbook, context.Worksheet, "excel_clear_range");
                string rangeAddress = context.Range.AddressLocal();
                long cellCount = ExcelComSupport.CellCount(context.Range);
                ExcelUndoScope undo = ExcelUndoScope.Capture(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress);

                if (scope == "contents" || scope == "content" || scope == "values")
                {
                    context.Range.ClearContents();
                }
                else if (scope == "formats" || scope == "format")
                {
                    context.Range.ClearFormats();
                }
                else if (scope == "all")
                {
                    context.Range.Clear();
                }
                else if (scope == "comments")
                {
                    context.Range.ClearComments();
                }
                else if (scope == "hyperlinks")
                {
                    context.Range.ClearHyperlinks();
                }
                else
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "scope 仅支持 contents/formats/all/comments/hyperlinks。");
                }

                string targetRef = ExcelComSupport.FullAddress(context.SheetName, rangeAddress);
                Observation observation = Observation.Create(
                    "clear_range",
                    "已清除 " + targetRef + " 的 " + scope + "，共 " + cellCount + " 个单元格。",
                    true);
                observation.AddTargetRef(targetRef);

                JObject data = new JObject();
                data["sheet"] = context.SheetName;
                data["address"] = rangeAddress;
                data["scope"] = scope;
                data["clearedCells"] = cellCount;

                string fallbackReason = null;
                if (scope == "formats" || scope == "all")
                {
                    fallbackReason = "区域快照只包含值、公式和数字格式，字体/填充/边框等视觉格式不会被自动撤销。";
                }
                else if (scope == "comments" || scope == "hyperlinks")
                {
                    fallbackReason = "批注/超链接状态不在区域快照中，未登记自动撤销。";
                }

                undo.Complete(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress,
                    true,
                    "清除 " + targetRef + " 的 " + scope,
                    observation,
                    data,
                    fallbackReason);

                return OperationResult.Succeed(
                    "已清除 " + targetRef + " 的 " + scope + "。",
                    data,
                    observation);
            }
        }

        private static string ResolveBaseSheetName(Excel.Application application, string targetAddress)
        {
            string sheetName;
            string local;
            RangeResolver.SplitAddress(targetAddress ?? string.Empty, out sheetName, out local);
            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                return sheetName;
            }
            object active = application.ActiveSheet;
            Excel.Worksheet worksheet = active as Excel.Worksheet;
            if (worksheet == null)
            {
                ComHelper.Release(active);
                return string.Empty;
            }
            try
            {
                return worksheet.Name;
            }
            finally
            {
                ComHelper.Release(worksheet);
            }
        }

        private static bool IsMerged(Excel.Range range)
        {
            object merged = range.MergeCells;
            return merged is bool && (bool)merged;
        }

        private static object BuildAllColumnIndexes(int columns)
        {
            object[] indexes = new object[columns];
            for (int index = 0; index < columns; index++)
            {
                indexes[index] = index + 1;
            }
            return indexes;
        }

        private static int FillEmptyCells(Excel.Range range, string fillValue)
        {
            int affected = 0;
            int rows = range.Rows.Count;
            int columns = range.Columns.Count;
            for (int row = 1; row <= rows; row++)
            {
                for (int column = 1; column <= columns; column++)
                {
                    Excel.Range cell = null;
                    try
                    {
                        cell = (Excel.Range)range.Cells[row, column];
                        bool merged = IsMerged(cell);
                        if (merged)
                        {
                            Excel.Range mergeArea = null;
                            Excel.Range first = null;
                            try
                            {
                                mergeArea = cell.MergeArea;
                                first = (Excel.Range)mergeArea.Cells[1, 1];
                                string firstAddress = first.AddressLocal();
                                string cellAddress = cell.AddressLocal();
                                if (!string.Equals(firstAddress, cellAddress, StringComparison.Ordinal))
                                {
                                    continue;
                                }
                            }
                            finally
                            {
                                ComHelper.Release(first, mergeArea);
                            }
                        }

                        object value = cell.Value2;
                        if (value == null || value is DBNull || string.IsNullOrEmpty(value.ToString()))
                        {
                            cell.Value2 = fillValue;
                            affected++;
                        }
                    }
                    finally
                    {
                        ComHelper.Release(cell);
                    }
                }
            }
            return affected;
        }

        private static int TrimTextCells(Excel.Range range)
        {
            int affected = 0;
            int rows = range.Rows.Count;
            int columns = range.Columns.Count;
            for (int row = 1; row <= rows; row++)
            {
                for (int column = 1; column <= columns; column++)
                {
                    Excel.Range cell = null;
                    try
                    {
                        cell = (Excel.Range)range.Cells[row, column];
                        object value = cell.Value2;
                        string text = value as string;
                        if (text == null)
                        {
                            continue;
                        }
                        string trimmed = text.Trim();
                        if (!string.Equals(text, trimmed, StringComparison.Ordinal))
                        {
                            cell.Value2 = trimmed;
                            affected++;
                        }
                    }
                    finally
                    {
                        ComHelper.Release(cell);
                    }
                }
            }
            return affected;
        }

        private sealed class FormulaApplication
        {
            public string Range { get; set; }

            public string Address { get; set; }

            public string Formula { get; set; }

            public string FormulaEngine { get; set; }

            public bool FillDown { get; set; }

            public int Rows { get; set; }

            public bool UndoRestorable { get; set; }

            public string UndoReason { get; set; }

            public JObject ToJObject()
            {
                JObject json = new JObject();
                json["range"] = Range;
                json["address"] = Address;
                json["formula"] = Formula;
                json["formulaEngine"] = FormulaEngine;
                json["fillDown"] = FillDown;
                json["rows"] = Rows;
                json["undoRestorable"] = UndoRestorable;
                if (!string.IsNullOrWhiteSpace(UndoReason))
                {
                    json["undoReason"] = UndoReason;
                }
                return json;
            }
        }
    }
}
