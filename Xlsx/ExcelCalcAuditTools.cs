using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>P0 求值闭环：强制重算、引用平移审计与金额精度扫描。</summary>
    public sealed partial class ExcelToolExecutor
    {
        private sealed class FormulaStats
        {
            public int FormulaCells { get; set; }

            public int ErrorCells { get; set; }

            public bool ReadSucceeded { get; set; }
        }

        private static OperationResult ExecuteRecalculate(Excel.Application application, JObject arguments)
        {
            string scope = ExcelArgumentReader.GetTrimmedString(arguments, "scope").ToLowerInvariant();
            if (scope.Length == 0)
            {
                scope = "workbook";
            }
            if (scope != "workbook" && scope != "sheet" && scope != "range")
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "scope 仅支持 workbook/sheet/range。");
            }
            bool fullRebuild = ExcelArgumentReader.GetBool(arguments, false, "fullRebuild", "rebuild");
            string calcMode = ExcelArgumentReader.GetTrimmedString(arguments, "calcMode").ToLowerInvariant();
            if (calcMode.Length == 0)
            {
                calcMode = "auto";
            }
            if (calcMode != "auto" && calcMode != "manual" && calcMode != "iterative")
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "calcMode 仅支持 auto/manual/iterative。");
            }

            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange", "address");
            if (scope == "range" && string.IsNullOrWhiteSpace(address))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "scope=range 时必须提供 range/targetRange。");
            }

            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Range targetRange = null;
            ExcelRangeContext rangeContext = null;
            FormulaStats before = new FormulaStats();
            FormulaStats after = new FormulaStats();
            object[,] valuesBefore = null;
            object[,] valuesAfter = null;
            int rows = 0;
            int columns = 0;
            int changedCells = 0;
            int verificationRows = 0;
            int verificationColumns = 0;
            bool valueSnapshotTaken = false;
            string snapshotNote = string.Empty;

            int originalCalculation = Convert.ToInt32(
                application.Calculation,
                CultureInfo.InvariantCulture);
            bool originalIteration = false;
            int originalMaxIterations = 100;
            double originalMaxChange = 0.001;
            try
            {
                originalIteration = application.Iteration;
                originalMaxIterations = application.MaxIterations;
                originalMaxChange = application.MaxChange;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[excel_recalculate] 读取迭代设置失败：" + ex.Message);
            }

            string restoreError = string.Empty;
            OperationResult outcome = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                if (scope == "workbook")
                {
                    before = CollectWorkbookFormulaStats(workbook);
                }
                else if (scope == "sheet")
                {
                    string sheetName = ExcelArgumentReader.GetSheetName(arguments);
                    worksheet = RangeResolver.ResolveWorksheet(application, workbook, sheetName, false);
                    targetRange = worksheet.UsedRange;
                    rows = targetRange.Rows.Count;
                    columns = targetRange.Columns.Count;
                    before = CollectRangeFormulaStats(targetRange);
                }
                else
                {
                    rangeContext = RangeResolver.Resolve(
                        application,
                        address,
                        false,
                        "excel_recalculate");
                    workbook = rangeContext.Workbook;
                    worksheet = rangeContext.Worksheet;
                    targetRange = rangeContext.Range;
                    rows = targetRange.Rows.Count;
                    columns = targetRange.Columns.Count;
                    before = CollectRangeFormulaStats(targetRange);
                }

                if (scope != "workbook" && targetRange != null &&
                    (long)rows * (long)columns <= 200000)
                {
                    try
                    {
                        valuesBefore = ExcelComSupport.ReadMatrix(targetRange, rows, columns);
                        verificationRows = rows;
                        verificationColumns = columns;
                        valueSnapshotTaken = true;
                    }
                    catch (Exception ex)
                    {
                        snapshotNote = "重算前回读快照失败：" + ex.Message;
                    }
                }
                else if (scope != "workbook")
                {
                    snapshotNote = "区域超过 200000 单元格，未采集逐单元格变化快照，只核对公式/错误计数。";
                }

                ApplyCalculationMode(application, calcMode, arguments);

                if (fullRebuild)
                {
                    application.CalculateFullRebuild();
                }
                else if (scope == "workbook")
                {
                    application.Calculate();
                }
                else if (scope == "sheet")
                {
                    worksheet.Calculate();
                }
                else
                {
                    targetRange.Calculate();
                }

                if (workbook != null)
                {
                    if (scope == "workbook")
                    {
                        after = CollectWorkbookFormulaStats(workbook);
                    }
                    else if (targetRange != null)
                    {
                        after = CollectRangeFormulaStats(targetRange);
                        if (valueSnapshotTaken)
                        {
                            valuesAfter = ExcelComSupport.ReadMatrix(targetRange, rows, columns);
                            changedCells = CountChangedValues(valuesBefore, valuesAfter);
                        }
                    }
                }

                bool verified;
                string verificationNote;
                if (scope == "workbook")
                {
                    verified = before.ReadSucceeded && after.ReadSucceeded;
                    verificationNote = verified
                        ? "已回读全工作簿公式统计：公式 " + after.FormulaCells.ToString(CultureInfo.InvariantCulture) +
                          " 个，错误值 " + after.ErrorCells.ToString(CultureInfo.InvariantCulture) + " 个。"
                        : "全工作簿公式统计回读失败，无法完成验证。";
                }
                else
                {
                    verified = before.ReadSucceeded && after.ReadSucceeded;
                    verificationNote = verified
                        ? "已回读目标范围：公式 " + after.FormulaCells.ToString(CultureInfo.InvariantCulture) +
                          " 个，错误值 " + after.ErrorCells.ToString(CultureInfo.InvariantCulture) +
                          " 个，值变化 " + changedCells.ToString(CultureInfo.InvariantCulture) + " 个单元格。"
                        : "目标范围公式统计回读失败，无法完成验证。";
                    if (!string.IsNullOrWhiteSpace(snapshotNote))
                    {
                        verificationNote = verificationNote + " " + snapshotNote;
                    }
                }

                JObject data = new JObject();
                data["scope"] = scope;
                data["sheet"] = worksheet == null ? string.Empty : worksheet.Name;
                data["address"] = targetRange == null ? string.Empty : targetRange.AddressLocal();
                data["fullRebuild"] = fullRebuild;
                data["calcMode"] = calcMode;
                data["originalCalcMode"] = DescribeCalculationMode(originalCalculation);
                data["restoredCalcMode"] = DescribeCalculationMode(originalCalculation);
                data["iteration"] = calcMode == "iterative";
                data["formulaCellsBefore"] = before.FormulaCells;
                data["formulaCellsAfter"] = after.FormulaCells;
                data["errorCellsBefore"] = before.ErrorCells;
                data["errorCellsAfter"] = after.ErrorCells;
                data["changedCells"] = changedCells;
                data["rows"] = rows;
                data["columns"] = columns;
                data["verified"] = verified;
                data["verificationNote"] = verificationNote;
                data["restored"] = false;
                ExcelUndoScope.RecordUnavailable(
                    data,
                    "重算只刷新计算缓存，不改变单元格存储内容，未登记自动撤销。");

                Observation observation = Observation.Create(
                    "recalculate",
                    "已重算 " + DescribeScope(scope, worksheet, targetRange) + "，回读验证" +
                    (verified ? "完成" : "未完成") + "。",
                    changedCells > 0);
                if (targetRange != null)
                {
                    observation.AddTargetRef(
                        ExcelComSupport.FullAddress(worksheet.Name, targetRange.AddressLocal()));
                }

                string message = "重算完成（scope=" + scope + "，calcMode=" + calcMode +
                                 (fullRebuild ? "，fullRebuild=true" : string.Empty) + "）；" +
                                 verificationNote + " 原计算模式已恢复为 " +
                                 DescribeCalculationMode(originalCalculation) + "。";
                if (!verified)
                {
                    message = message + " 注意：回读验证失败，请检查目标区域是否仍存在错误值。";
                }
                outcome = OperationResult.Succeed(message, data, observation);
            }
            finally
            {
                try
                {
                    application.Calculation = (Excel.XlCalculation)originalCalculation;
                }
                catch (Exception ex)
                {
                    restoreError = "恢复 Calculation 失败：" + ex.Message;
                    Console.Error.WriteLine("[excel_recalculate] " + restoreError);
                }
                try
                {
                    application.Iteration = originalIteration;
                    application.MaxIterations = originalMaxIterations;
                    application.MaxChange = originalMaxChange;
                }
                catch (Exception ex)
                {
                    restoreError = restoreError.Length == 0
                        ? "恢复迭代设置失败：" + ex.Message
                        : restoreError + "；恢复迭代设置失败：" + ex.Message;
                    Console.Error.WriteLine("[excel_recalculate] 恢复迭代设置失败：" + ex.Message);
                }
                if (rangeContext != null)
                {
                    rangeContext.Dispose();
                }
                else
                {
                    ComHelper.Release(targetRange, worksheet, workbook);
                }
            }

            if (outcome == null)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "重算未返回结果。" + (restoreError.Length == 0 ? string.Empty : " " + restoreError));
            }

            if (outcome.Data != null)
            {
                outcome.Data["restored"] = restoreError.Length == 0;
                if (restoreError.Length > 0)
                {
                    outcome.Data["restoreError"] = restoreError;
                }
            }

            if (outcome.Success && restoreError.Length > 0)
            {
                string rollbackNote =
                    "重算已执行，但 Excel 计算模式/迭代设置未能完整恢复：" + restoreError +
                    "。请手动在 Excel 公式选项中恢复原计算模式。";
                return OperationResult.PartialApplied(1, rollbackNote, outcome.Observation, outcome.Data);
            }

            return outcome;
        }

        private static void ApplyCalculationMode(Excel.Application application, string calcMode, JObject arguments)
        {
            if (calcMode == "auto")
            {
                application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            }
            else if (calcMode == "manual")
            {
                application.Calculation = Excel.XlCalculation.xlCalculationManual;
            }
            else
            {
                application.Calculation = Excel.XlCalculation.xlCalculationManual;
                application.Iteration = true;
                int maxIterations = ExcelArgumentReader.GetInt(arguments, application.MaxIterations, "maxIterations", "iterations");
                double maxChange;
                if (!ExcelArgumentReader.TryGetDouble(arguments, out maxChange, "maxChange"))
                {
                    maxChange = application.MaxChange;
                }
                application.MaxIterations = Math.Max(1, Math.Min(maxIterations, 10000));
                application.MaxChange = Math.Max(0.0, maxChange);
            }
        }

        private static string DescribeCalculationMode(int calculation)
        {
            if (calculation == (int)Excel.XlCalculation.xlCalculationAutomatic)
            {
                return "auto";
            }
            if (calculation == (int)Excel.XlCalculation.xlCalculationManual)
            {
                return "manual";
            }
            if (calculation == (int)Excel.XlCalculation.xlCalculationSemiautomatic)
            {
                return "semiautomatic";
            }
            return calculation.ToString(CultureInfo.InvariantCulture);
        }

        private static string DescribeScope(string scope, Excel.Worksheet worksheet, Excel.Range range)
        {
            if (scope == "workbook")
            {
                return "工作簿";
            }
            if (scope == "sheet")
            {
                return worksheet == null ? "工作表" : "工作表 " + worksheet.Name;
            }
            return range == null ? "目标区域" : "区域 " + range.AddressLocal();
        }

        private static FormulaStats CollectWorkbookFormulaStats(Excel.Workbook workbook)
        {
            FormulaStats stats = new FormulaStats();
            stats.ReadSucceeded = true;
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                int count = sheets.Count;
                for (int index = 1; index <= count; index++)
                {
                    object item = sheets[index];
                    Excel.Worksheet worksheet = item as Excel.Worksheet;
                    if (worksheet == null)
                    {
                        ComHelper.Release(item);
                        continue;
                    }
                    Excel.Range usedRange = null;
                    try
                    {
                        usedRange = worksheet.UsedRange;
                        FormulaStats one = CollectRangeFormulaStats(usedRange);
                        stats.FormulaCells += one.FormulaCells;
                        stats.ErrorCells += one.ErrorCells;
                        if (!one.ReadSucceeded)
                        {
                            stats.ReadSucceeded = false;
                        }
                    }
                    catch (Exception ex)
                    {
                        stats.ReadSucceeded = false;
                        Console.Error.WriteLine("[excel_recalculate] 统计工作表「" + worksheet.Name + "」失败：" + ex.Message);
                    }
                    finally
                    {
                        ComHelper.Release(usedRange, worksheet);
                    }
                }
            }
            catch (Exception ex)
            {
                stats.ReadSucceeded = false;
                Console.Error.WriteLine("[excel_recalculate] 遍历工作表失败：" + ex.Message);
            }
            finally
            {
                ComHelper.Release(sheets);
            }
            return stats;
        }

        private static FormulaStats CollectRangeFormulaStats(Excel.Range range)
        {
            FormulaStats stats = new FormulaStats();
            stats.ReadSucceeded = true;
            try
            {
                stats.FormulaCells = CountSpecialCells(
                    range,
                    Excel.XlCellType.xlCellTypeFormulas,
                    Type.Missing);
                stats.ErrorCells = CountSpecialCells(
                                       range,
                                       Excel.XlCellType.xlCellTypeFormulas,
                                       Excel.XlSpecialCellsValue.xlErrors) +
                                   CountSpecialCells(
                                       range,
                                       Excel.XlCellType.xlCellTypeConstants,
                                       Excel.XlSpecialCellsValue.xlErrors);
            }
            catch (Exception ex)
            {
                stats.ReadSucceeded = false;
                Console.Error.WriteLine("[excel_recalculate] 读取区域公式统计失败：" + ex.Message);
            }
            return stats;
        }

        private static int CountChangedValues(object[,] before, object[,] after)
        {
            if (before == null || after == null)
            {
                return 0;
            }
            int rows = Math.Min(before.GetLength(0), after.GetLength(0));
            int columns = Math.Min(before.GetLength(1), after.GetLength(1));
            int changed = 0;
            for (int row = 0; row < rows; row++)
            {
                for (int column = 0; column < columns; column++)
                {
                    if (!CellValuesEqual(before[row, column], after[row, column]))
                    {
                        changed++;
                    }
                }
            }
            return changed;
        }

        private static bool CellValuesEqual(object left, object right)
        {
            bool leftEmpty = left == null || left is DBNull;
            bool rightEmpty = right == null || right is DBNull;
            if (leftEmpty || rightEmpty)
            {
                return leftEmpty && rightEmpty;
            }
            if (IsNumeric(left) && IsNumeric(right))
            {
                double leftNumber = Convert.ToDouble(left, CultureInfo.InvariantCulture);
                double rightNumber = Convert.ToDouble(right, CultureInfo.InvariantCulture);
                if (double.IsNaN(leftNumber) && double.IsNaN(rightNumber))
                {
                    return true;
                }
                return Math.Abs(leftNumber - rightNumber) <= 1e-12;
            }
            return string.Equals(
                Convert.ToString(left, CultureInfo.InvariantCulture),
                Convert.ToString(right, CultureInfo.InvariantCulture),
                StringComparison.Ordinal);
        }

        private static bool IsNumeric(object value)
        {
            return value is double || value is float || value is int || value is long ||
                   value is short || value is decimal || value is byte;
        }

        // ------------------------------------------------------------------
        // 引用平移审计
        // ------------------------------------------------------------------

        private static OperationResult ExecuteFormulaReferenceAudit(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "targetRange", "range", "address");
            if (string.IsNullOrWhiteSpace(address))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 targetRange/range 参数。");
            }
            string expectedMode = ExcelArgumentReader.GetTrimmedString(arguments, "expectedMode", "mode").ToLowerInvariant();
            if (expectedMode.Length == 0)
            {
                expectedMode = "auto";
            }
            if (expectedMode != "auto" && expectedMode != "mixed" && expectedMode != "row" &&
                expectedMode != "column" && expectedMode != "both" && expectedMode != "relative" &&
                expectedMode != "absolute" && expectedMode != "none")
            {
                return Fail(
                    OfficeErrorCode.INVALID_PARAMS,
                    "expectedMode 仅支持 auto/mixed/row/column/both/relative/absolute/none。");
            }
            int sampleRows = ExcelArgumentReader.GetInt(arguments, 3, "sampleRows", "samples");
            sampleRows = Math.Max(1, Math.Min(sampleRows, 20));
            int maxResults = ExcelArgumentReader.GetInt(arguments, 100, "maxResults", "limit");
            maxResults = Math.Max(1, Math.Min(maxResults, 1000));

            using (ExcelRangeContext context = RangeResolver.Resolve(
                       application,
                       address,
                       false,
                       "excel_formula_reference_audit"))
            {
                int rows = context.Range.Rows.Count;
                int columns = context.Range.Columns.Count;
                object[,] formulas = ExcelComSupport.NormalizeMatrix(context.Range.Formula, rows, columns);
                object[,] values = ExcelComSupport.ReadMatrix(context.Range, rows, columns);

                int templateRow = -1;
                int templateColumn = -1;
                string templateFormula = string.Empty;
                for (int row = 0; row < rows && templateFormula.Length == 0; row++)
                {
                    for (int column = 0; column < columns; column++)
                    {
                        string formula = ReadFormulaCell(formulas, row, column);
                        if (formula.Length > 0 && formula[0] == '=')
                        {
                            templateRow = row;
                            templateColumn = column;
                            templateFormula = formula;
                            break;
                        }
                    }
                }

                JObject data = new JObject();
                data["sheet"] = context.SheetName;
                data["address"] = context.Range.AddressLocal();
                data["rows"] = rows;
                data["columns"] = columns;
                data["expectedMode"] = expectedMode;

                if (templateFormula.Length == 0)
                {
                    data["formulaCells"] = 0;
                    data["issueCount"] = 0;
                    data["issues"] = new JArray();
                    data["samples"] = new JArray();
                    data["passed"] = true;
                    ExcelUndoScope.RecordUnavailable(data, "引用审计为只读操作，无需撤销。");
                    return OperationResult.Succeed(
                        "目标区域没有公式，引用平移审计无事可做。",
                        data,
                        Observation.Create("formula_reference_audit", "目标区域没有公式。", false));
                }

                bool shiftRows = expectedMode != "column" && expectedMode != "absolute" && expectedMode != "none";
                bool shiftColumns = expectedMode != "row" && expectedMode != "absolute" && expectedMode != "none";

                JArray issues = new JArray();
                int issueCount = 0;
                int formulaCells = 0;
                int referenceCount = 0;
                int rowAbsoluteCount = 0;
                int columnAbsoluteCount = 0;
                int relativeCount = 0;
                List<int[]> sampleCandidates = new List<int[]>();

                for (int row = 0; row < rows; row++)
                {
                    for (int column = 0; column < columns; column++)
                    {
                        string formula = ReadFormulaCell(formulas, row, column);
                        if (formula.Length == 0 || formula[0] != '=')
                        {
                            continue;
                        }
                        formulaCells++;
                        List<ExcelFormulaReference> references = ExcelFormulaReferenceEngine.EnumerateReferences(formula);
                        foreach (ExcelFormulaReference reference in references)
                        {
                            referenceCount++;
                            if (reference.RowAbsolute)
                            {
                                rowAbsoluteCount++;
                            }
                            if (reference.ColumnAbsolute)
                            {
                                columnAbsoluteCount++;
                            }
                            if (!reference.RowAbsolute && !reference.ColumnAbsolute)
                            {
                                relativeCount++;
                            }
                        }

                        int rowDelta = shiftRows ? row - templateRow : 0;
                        int columnDelta = shiftColumns ? column - templateColumn : 0;
                        string expected = ExcelFormulaReferenceEngine.ShiftFormula(templateFormula, rowDelta, columnDelta);
                        if (!string.Equals(
                                ExcelFormulaReferenceEngine.NormalizeFormula(expected),
                                ExcelFormulaReferenceEngine.NormalizeFormula(formula),
                                StringComparison.Ordinal))
                        {
                            issueCount++;
                            if (issues.Count < maxResults)
                            {
                                issues.Add(BuildReferenceIssue(
                                    context,
                                    row,
                                    column,
                                    "referenceShiftMismatch",
                                    expected,
                                    formula,
                                    rowDelta,
                                    columnDelta));
                            }
                        }

                        bool modeViolation = false;
                        string violationKind = string.Empty;
                        if (expectedMode == "absolute")
                        {
                            foreach (ExcelFormulaReference reference in references)
                            {
                                if (!reference.RowAbsolute || !reference.ColumnAbsolute)
                                {
                                    modeViolation = true;
                                    violationKind = "expectedAbsoluteReference";
                                    break;
                                }
                            }
                        }
                        else if (expectedMode == "relative")
                        {
                            foreach (ExcelFormulaReference reference in references)
                            {
                                if (reference.RowAbsolute || reference.ColumnAbsolute)
                                {
                                    modeViolation = true;
                                    violationKind = "expectedRelativeReference";
                                    break;
                                }
                            }
                        }
                        else if (expectedMode == "row")
                        {
                            foreach (ExcelFormulaReference reference in references)
                            {
                                if (reference.RowAbsolute)
                                {
                                    modeViolation = true;
                                    violationKind = "expectedRelativeRowReference";
                                    break;
                                }
                            }
                        }
                        else if (expectedMode == "column")
                        {
                            foreach (ExcelFormulaReference reference in references)
                            {
                                if (reference.ColumnAbsolute)
                                {
                                    modeViolation = true;
                                    violationKind = "expectedRelativeColumnReference";
                                    break;
                                }
                            }
                        }

                        if (modeViolation)
                        {
                            issueCount++;
                            if (issues.Count < maxResults)
                            {
                                issues.Add(BuildReferenceIssue(
                                    context,
                                    row,
                                    column,
                                    violationKind,
                                    "按 expectedMode=" + expectedMode + " 预期的引用模式",
                                    formula,
                                    rowDelta,
                                    columnDelta));
                            }
                        }

                        int absoluteRow = row + 1;
                        bool alreadySampled = false;
                        foreach (int[] candidate in sampleCandidates)
                        {
                            if (candidate[0] == absoluteRow)
                            {
                                alreadySampled = true;
                                break;
                            }
                        }
                        if (!alreadySampled && sampleCandidates.Count < sampleRows + 2)
                        {
                            sampleCandidates.Add(new int[] { absoluteRow, column + 1, row, column });
                        }
                    }
                }

                JArray samples = BuildReferenceSamples(
                    application,
                    context,
                    formulas,
                    values,
                    sampleCandidates,
                    sampleRows);
                bool samplesPassed = true;
                foreach (JToken sampleToken in samples)
                {
                    JObject sample = sampleToken as JObject;
                    if (sample != null && !(sample.Value<bool?>("match") ?? false))
                    {
                        samplesPassed = false;
                    }
                }

                data["formulaCells"] = formulaCells;
                data["issueCount"] = issueCount;
                data["truncated"] = issueCount > issues.Count;
                data["issues"] = issues;
                data["references"] = BuildReferenceSummary(
                    referenceCount,
                    rowAbsoluteCount,
                    columnAbsoluteCount,
                    relativeCount);
                data["templateFormula"] = templateFormula;
                data["templateCell"] = ExcelExtensionSupport.CellAddress(
                    context.Range.Row + templateRow,
                    context.Range.Column + templateColumn);
                data["expectedShifts"] = new JObject
                {
                    { "rows", shiftRows },
                    { "columns", shiftColumns }
                };
                data["samples"] = samples;
                data["samplesPassed"] = samplesPassed;
                data["passed"] = issueCount == 0 && samplesPassed;
                ExcelUndoScope.RecordUnavailable(data, "引用审计为只读操作，无需撤销。");

                string message = "引用平移审计完成：" + formulaCells.ToString(CultureInfo.InvariantCulture) +
                                 " 个公式，发现 " + issueCount.ToString(CultureInfo.InvariantCulture) +
                                 " 个引用问题，抽样 " + samples.Count.ToString(CultureInfo.InvariantCulture) +
                                 " 个单元格独立重算" + (samplesPassed ? "全部一致" : "存在不一致") + "。";
                Observation observation = Observation.Create(
                    "formula_reference_audit",
                    message,
                    false);
                observation.AddTargetRef(ExcelComSupport.FullAddress(context.SheetName, context.Range.AddressLocal()));
                return OperationResult.Succeed(message, data, observation);
            }
        }

        private static JObject BuildReferenceIssue(
            ExcelRangeContext context,
            int row,
            int column,
            string issue,
            string expected,
            string actual,
            int rowDelta,
            int columnDelta)
        {
            JObject item = new JObject();
            item["cell"] = ExcelExtensionSupport.CellAddress(
                context.Range.Row + row,
                context.Range.Column + column);
            item["issue"] = issue;
            item["rowDelta"] = rowDelta;
            item["columnDelta"] = columnDelta;
            item["expectedFormula"] = expected;
            item["actualFormula"] = actual;
            return item;
        }

        private static JObject BuildReferenceSummary(
            int total,
            int rowAbsolute,
            int columnAbsolute,
            int relative)
        {
            JObject references = new JObject();
            references["total"] = total;
            references["rowAbsolute"] = rowAbsolute;
            references["columnAbsolute"] = columnAbsolute;
            references["fullyRelative"] = relative;
            return references;
        }

        private static JArray BuildReferenceSamples(
            Excel.Application application,
            ExcelRangeContext context,
            object[,] formulas,
            object[,] values,
            List<int[]> candidates,
            int sampleRows)
        {
            JArray samples = new JArray();
            if (candidates.Count == 0)
            {
                return samples;
            }

            List<int> indexes = new List<int>();
            if (sampleRows <= 1 || candidates.Count == 1)
            {
                indexes.Add(0);
            }
            else
            {
                indexes.Add(0);
                int middle = candidates.Count / 2;
                if (middle != 0 && middle != candidates.Count - 1)
                {
                    indexes.Add(middle);
                }
                indexes.Add(candidates.Count - 1);
            }

            int added = 0;
            foreach (int index in indexes)
            {
                if (added >= sampleRows)
                {
                    break;
                }
                int[] candidate = candidates[index];
                int row = candidate[2];
                int column = candidate[3];
                JObject sample = new JObject();
                sample["cell"] = ExcelExtensionSupport.CellAddress(
                    context.Range.Row + row,
                    context.Range.Column + column);
                string formula = ReadFormulaCell(formulas, row, column);
                sample["formula"] = formula;
                object cached = values[row, column];
                sample["cachedValue"] = ExcelComSupport.ToJToken(cached);
                bool matched = false;
                string note = string.Empty;
                string qualified;
                if (!ExcelFormulaReferenceEngine.TryQualifyForEvaluate(formula, context.SheetName, out qualified))
                {
                    note = "公式包含跨表/结构化引用或整行整列引用，未做独立求值；已用回读值替代。";
                    matched = true;
                }
                else
                {
                    try
                    {
                        object evaluated = application.Evaluate(qualified);
                        sample["independentValue"] = ExcelComSupport.ToJToken(evaluated);
                        matched = EvaluateValuesEqual(cached, evaluated);
                    }
                    catch (Exception ex)
                    {
                        note = "独立求值失败：" + ex.Message;
                        matched = false;
                    }
                }
                sample["match"] = matched;
                if (note.Length > 0)
                {
                    sample["note"] = note;
                }
                samples.Add(sample);
                added++;
            }
            return samples;
        }

        private static bool EvaluateValuesEqual(object cached, object independent)
        {
            bool cachedEmpty = cached == null || cached is DBNull;
            bool independentEmpty = independent == null || independent is DBNull;
            if (cachedEmpty || independentEmpty)
            {
                return cachedEmpty && independentEmpty;
            }
            if (IsNumeric(cached) && IsNumeric(independent))
            {
                double left = Convert.ToDouble(cached, CultureInfo.InvariantCulture);
                double right = Convert.ToDouble(independent, CultureInfo.InvariantCulture);
                if (double.IsNaN(left) && double.IsNaN(right))
                {
                    return true;
                }
                return Math.Abs(left - right) <= 1e-9 * Math.Max(1.0, Math.Max(Math.Abs(left), Math.Abs(right)));
            }
            if (cached is bool && independent is bool)
            {
                return (bool)cached == (bool)independent;
            }
            return string.Equals(
                Convert.ToString(cached, CultureInfo.InvariantCulture),
                Convert.ToString(independent, CultureInfo.InvariantCulture),
                StringComparison.Ordinal);
        }

        private static string ReadFormulaCell(object[,] formulas, int row, int column)
        {
            if (formulas == null)
            {
                return string.Empty;
            }
            object value = formulas[row, column];
            if (value == null || value is DBNull)
            {
                return string.Empty;
            }
            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        // ------------------------------------------------------------------
        // 金额精度扫描
        // ------------------------------------------------------------------

        private static OperationResult ExecutePrecisionCheck(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange", "address");
            if (string.IsNullOrWhiteSpace(address))
            {
                string sheetName = ExcelArgumentReader.GetSheetName(arguments);
                address = string.IsNullOrWhiteSpace(sheetName)
                    ? "usedrange"
                    : ExcelComSupport.QuoteSheetName(sheetName) + "!usedrange";
            }
            double tolerance;
            if (!ExcelArgumentReader.TryGetDouble(arguments, out tolerance, "tolerance", "threshold"))
            {
                tolerance = 0.005;
            }
            if (tolerance < 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "tolerance 不能为负数。");
            }
            int maxResults = ExcelArgumentReader.GetInt(arguments, 200, "maxResults", "limit");
            maxResults = Math.Max(1, Math.Min(maxResults, 5000));

            using (ExcelRangeContext context = RangeResolver.Resolve(
                       application,
                       address,
                       false,
                       "excel_precision_check"))
            {
                int rows = context.Range.Rows.Count;
                int columns = context.Range.Columns.Count;
                if ((long)rows * (long)columns > 500000)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_RANGE,
                        "精度扫描区域超过 500000 单元格，请缩小 range 后重试。");
                }

                object[,] values = ExcelComSupport.ReadMatrix(context.Range, rows, columns);
                object[,] formulas = ExcelComSupport.NormalizeMatrix(context.Range.Formula, rows, columns);
                bool formatsComplete;
                object[,] formats = ReadNumberFormatGrid(
                    context.Range,
                    values,
                    rows,
                    columns,
                    out formatsComplete);

                JArray issues = new JArray();
                int issueCount = 0;
                int textNumberCount = 0;
                int scientificCount = 0;
                int roundedCount = 0;
                int numericCells = 0;
                int textCells = 0;
                int formulaCells = 0;
                double doubleSum = 0.0;
                decimal decimalSum = 0.0m;
                bool decimalSumValid = true;

                for (int row = 0; row < rows; row++)
                {
                    for (int column = 0; column < columns; column++)
                    {
                        object value = values[row, column];
                        if (value == null || value is DBNull)
                        {
                            continue;
                        }
                        string numberFormat = Convert.ToString(formats[row, column], CultureInfo.InvariantCulture) ?? string.Empty;
                        string formula = ReadFormulaCell(formulas, row, column);
                        if (formula.Length > 0 && formula[0] == '=')
                        {
                            formulaCells++;
                        }

                        if (IsNumeric(value))
                        {
                            numericCells++;
                            double number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                            doubleSum += number;
                            if (decimalSumValid)
                            {
                                try
                                {
                                    decimalSum += Convert.ToDecimal(number, CultureInfo.InvariantCulture);
                                }
                                catch (OverflowException)
                                {
                                    decimalSumValid = false;
                                }
                            }

                            if (LooksScientificFormat(numberFormat))
                            {
                                scientificCount++;
                                issueCount++;
                                AddPrecisionIssue(
                                    issues,
                                    maxResults,
                                    context,
                                    row,
                                    column,
                                    "scientificNotation",
                                    "info",
                                    value,
                                    null,
                                    0.0,
                                    numberFormat,
                                    "数值使用了科学计数格式，导出/核对时可能丢失可读性；建议改成常规或货币格式。");
                            }

                            int decimals = ReadFormatDecimals(numberFormat);
                            if (decimals >= 0 && decimals <= 12)
                            {
                                double rounded = Math.Round(number, decimals, MidpointRounding.AwayFromZero);
                                double delta = Math.Abs(number - rounded);
                                if (delta > tolerance)
                                {
                                    roundedCount++;
                                    issueCount++;
                                    AddPrecisionIssue(
                                        issues,
                                        maxResults,
                                        context,
                                        row,
                                        column,
                                        "roundedValue",
                                        "warning",
                                        value,
                                        rounded,
                                        delta,
                                        numberFormat,
                                        "该值按当前显示格式会被舍入 " + decimals.ToString(CultureInfo.InvariantCulture) +
                                        " 位小数，与底层存储值相差超过 tolerance。");
                                }
                            }
                        }
                        else if (value is string)
                        {
                            textCells++;
                            string text = ((string)value).Trim();
                            decimal parsed;
                            if (TryParseAmount(text, out parsed))
                            {
                                textNumberCount++;
                                issueCount++;
                                AddPrecisionIssue(
                                    issues,
                                    maxResults,
                                    context,
                                    row,
                                    column,
                                    "numberStoredAsText",
                                    "warning",
                                    value,
                                    parsed,
                                    0.0,
                                    numberFormat,
                                    "金额/数值以文本存储，SUM 等聚合会漏算；建议改用数值并设置数字格式。");
                            }
                            if (Regex.IsMatch(text, @"^[+-]?\d+(?:\.\d+)?[Ee][+-]?\d+$"))
                            {
                                scientificCount++;
                                issueCount++;
                                AddPrecisionIssue(
                                    issues,
                                    maxResults,
                                    context,
                                    row,
                                    column,
                                    "scientificNotation",
                                    "info",
                                    value,
                                    null,
                                    0.0,
                                    numberFormat,
                                    "文本形式的科学计数数值，解析时可能被误读。");
                            }
                        }
                    }
                }

                double floatingDelta = 0.0;
                if (decimalSumValid)
                {
                    floatingDelta = Math.Abs(doubleSum - Convert.ToDouble(decimalSum, CultureInfo.InvariantCulture));
                    if (floatingDelta > tolerance)
                    {
                        issueCount++;
                        JObject floatingIssue = new JObject();
                        floatingIssue["cell"] = context.Range.AddressLocal();
                        floatingIssue["issue"] = "floatingAccumulation";
                        floatingIssue["severity"] = "warning";
                        floatingIssue["doubleSum"] = doubleSum;
                        floatingIssue["decimalSum"] = Convert.ToDouble(decimalSum, CultureInfo.InvariantCulture);
                        floatingIssue["delta"] = floatingDelta;
                        floatingIssue["suggestion"] =
                            "浮点累计与十进制累计相差超过 tolerance；建议用 ROUND(SUM(...),2) 固定金额精度，或改为分单位整数计算。";
                        if (issues.Count < maxResults)
                        {
                            issues.Add(floatingIssue);
                        }
                    }
                }

                JObject data = new JObject();
                data["sheet"] = context.SheetName;
                data["address"] = context.Range.AddressLocal();
                data["rows"] = rows;
                data["columns"] = columns;
                data["tolerance"] = tolerance;
                data["scannedCells"] = (long)rows * (long)columns;
                data["numericCells"] = numericCells;
                data["textCells"] = textCells;
                data["formulaCells"] = formulaCells;
                data["issueCount"] = issueCount;
                data["truncated"] = issueCount > issues.Count;
                data["issues"] = issues;
                data["counts"] = new JObject
                {
                    { "numberStoredAsText", textNumberCount },
                    { "scientificNotation", scientificCount },
                    { "roundedValue", roundedCount },
                    { "floatingAccumulation", floatingDelta > tolerance ? 1 : 0 }
                };
                JObject floatingPoint = new JObject();
                floatingPoint["withinTolerance"] = floatingDelta <= tolerance;
                floatingPoint["delta"] = floatingDelta;
                floatingPoint["doubleSum"] = doubleSum;
                floatingPoint["decimalSumValid"] = decimalSumValid;
                if (decimalSumValid)
                {
                    floatingPoint["decimalSum"] = Convert.ToDouble(decimalSum, CultureInfo.InvariantCulture);
                }
                data["floatingPoint"] = floatingPoint;
                data["numberFormatsComplete"] = formatsComplete;
                if (!formatsComplete)
                {
                    data["note"] = "区域过大，数字格式仅完整读取前 100000 个非空单元格；舍入/科学计数结论可能不完整。";
                }
                ExcelUndoScope.RecordUnavailable(data, "精度扫描为只读操作，无需撤销。");

                string message = "精度扫描完成：" + issueCount.ToString(CultureInfo.InvariantCulture) +
                                 " 个问题（文本数值 " + textNumberCount.ToString(CultureInfo.InvariantCulture) +
                                 "，科学计数 " + scientificCount.ToString(CultureInfo.InvariantCulture) +
                                 "，舍入 " + roundedCount.ToString(CultureInfo.InvariantCulture) +
                                 "，浮点累计差 " + (floatingDelta > tolerance ? "有" : "无") + "）。";
                Observation observation = Observation.Create("precision_check", message, false);
                observation.AddTargetRef(ExcelComSupport.FullAddress(context.SheetName, context.Range.AddressLocal()));
                return OperationResult.Succeed(message, data, observation);
            }
        }

        /// <summary>
        /// 读取数字格式矩阵。Excel Range.NumberFormat 在混合格式时可能返回 null、统一格式时返回单个字符串，
        /// 这里对 null 回退到逐单元格读取（限制 100000 个非空单元格，避免大区域拖死会话）。
        /// </summary>
        private static object[,] ReadNumberFormatGrid(
            Excel.Range range,
            object[,] values,
            int rows,
            int columns,
            out bool complete)
        {
            complete = true;
            object[,] formats = new object[rows, columns];
            object raw = range.NumberFormat;
            if (raw is object[,])
            {
                return ExcelComSupport.NormalizeMatrix(raw, rows, columns);
            }
            string uniform = raw as string;
            if (!string.IsNullOrEmpty(uniform))
            {
                for (int row = 0; row < rows; row++)
                {
                    for (int column = 0; column < columns; column++)
                    {
                        formats[row, column] = uniform;
                    }
                }
                return formats;
            }

            int reads = 0;
            const int MaxFormatReads = 100000;
            for (int row = 0; row < rows; row++)
            {
                for (int column = 0; column < columns; column++)
                {
                    object value = values == null ? null : values[row, column];
                    if (value == null || value is DBNull)
                    {
                        continue;
                    }
                    if (reads >= MaxFormatReads)
                    {
                        complete = false;
                        return formats;
                    }
                    Excel.Range cell = null;
                    try
                    {
                        cell = (Excel.Range)range.Cells[row + 1, column + 1];
                        formats[row, column] = cell.NumberFormat;
                        reads++;
                    }
                    catch (Exception ex)
                    {
                        complete = false;
                        Console.Error.WriteLine("[excel_precision_check] 读取单元格数字格式失败：" + ex.Message);
                    }
                    finally
                    {
                        ComHelper.Release(cell);
                    }
                }
            }
            return formats;
        }

        private static void AddPrecisionIssue(
            JArray issues,
            int maxResults,
            ExcelRangeContext context,
            int row,
            int column,
            string issue,
            string severity,
            object value,
            object expected,
            double delta,
            string numberFormat,
            string suggestion)
        {
            if (issues.Count >= maxResults)
            {
                return;
            }
            JObject item = new JObject();
            item["cell"] = ExcelExtensionSupport.CellAddress(
                context.Range.Row + row,
                context.Range.Column + column);
            item["issue"] = issue;
            item["severity"] = severity;
            item["value"] = ExcelComSupport.ToJToken(value);
            if (expected != null)
            {
                item["expected"] = ExcelComSupport.ToJToken(expected);
            }
            item["delta"] = delta;
            item["numberFormat"] = numberFormat;
            item["suggestion"] = suggestion;
            issues.Add(item);
        }

        private static bool TryParseAmount(string text, out decimal value)
        {
            value = 0m;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }
            string normalized = text.Trim();
            if (normalized.StartsWith("'", StringComparison.Ordinal))
            {
                normalized = normalized.Substring(1);
            }
            normalized = normalized.Replace(",", string.Empty).Replace("，", string.Empty);
            if (normalized.EndsWith("%", StringComparison.Ordinal))
            {
                return false;
            }
            return decimal.TryParse(
                normalized,
                NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent,
                CultureInfo.InvariantCulture,
                out value);
        }

        private static bool LooksScientificFormat(string numberFormat)
        {
            if (string.IsNullOrWhiteSpace(numberFormat))
            {
                return false;
            }
            return numberFormat.IndexOf("E+", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   numberFormat.IndexOf("E-", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static int ReadFormatDecimals(string numberFormat)
        {
            if (string.IsNullOrWhiteSpace(numberFormat))
            {
                return -1;
            }
            if (numberFormat.IndexOf("General", StringComparison.OrdinalIgnoreCase) >= 0 ||
                numberFormat.IndexOf("@", StringComparison.Ordinal) >= 0 ||
                numberFormat.IndexOf("%", StringComparison.Ordinal) >= 0 ||
                LooksScientificFormat(numberFormat))
            {
                return -1;
            }
            string section = numberFormat.Split(';')[0];
            int dot = section.IndexOf('.');
            if (dot < 0)
            {
                return 0;
            }
            int decimals = 0;
            for (int index = dot + 1; index < section.Length; index++)
            {
                char character = section[index];
                if (character == '0' || character == '#')
                {
                    decimals++;
                    continue;
                }
                break;
            }
            return decimals;
        }
    }
}
