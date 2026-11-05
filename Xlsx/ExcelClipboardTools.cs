using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>P1 动作类：选择性粘贴、公式固化、格式复制、超链接与合并单元格。</summary>
    public sealed partial class ExcelToolExecutor
    {
        private static OperationResult ExecutePasteSpecial(Excel.Application application, JObject arguments)
        {
            string sourceAddress = ExcelArgumentReader.GetRangeAddress(
                arguments,
                "sourceRange",
                "source",
                "fromRange");
            string targetAddress = ExcelArgumentReader.GetRangeAddress(
                arguments,
                "targetRange",
                "target",
                "toRange");
            if (string.IsNullOrWhiteSpace(sourceAddress) || string.IsNullOrWhiteSpace(targetAddress))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 sourceRange 或 targetRange 参数。");
            }

            string mode = ExcelArgumentReader.GetTrimmedString(arguments, "mode", "pasteMode").ToLowerInvariant();
            if (mode.Length == 0)
            {
                mode = "values";
            }
            if (mode == "valuesandformats")
            {
                mode = "valuesAndFormats";
            }
            if (mode != "values" && mode != "formulas" && mode != "formats" &&
                mode != "transpose" && mode != "valuesAndFormats")
            {
                return Fail(
                    OfficeErrorCode.INVALID_PARAMS,
                    "mode 仅支持 values/formulas/formats/transpose/valuesAndFormats。");
            }
            bool transpose = mode == "transpose" ||
                             ExcelArgumentReader.GetBool(arguments, false, "transpose");
            if (mode == "transpose")
            {
                mode = "values";
            }

            string operate = ExcelArgumentReader.GetTrimmedString(arguments, "operate", "operation").ToLowerInvariant();
            if (operate.Length == 0)
            {
                operate = "none";
            }
            if (operate != "none" && operate != "add" && operate != "subtract" &&
                operate != "multiply" && operate != "divide")
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "operate 仅支持 none/add/subtract/multiply/divide。");
            }
            if (operate != "none" && mode != "values" && mode != "valuesAndFormats")
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "operate 只支持 values/valuesAndFormats 模式。");
            }
            if (transpose && mode == "formulas")
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "formulas 模式暂不支持 transpose，请先转置值再写入公式。");
            }

            bool skipBlanks = ExcelArgumentReader.GetBool(arguments, false, "skipBlanks");
            bool includeFormats = mode == "valuesAndFormats" ||
                                  ExcelArgumentReader.GetBool(arguments, false, "includeFormats");

            using (ExcelRangeContext source = RangeResolver.Resolve(
                       application,
                       sourceAddress,
                       false,
                       "excel_paste_special"))
            {
                using (ExcelRangeContext target = RangeResolver.Resolve(
                           application,
                           targetAddress,
                           false,
                           "excel_paste_special"))
                {
                    ExcelComSupport.EnsureWritable(
                        application,
                        target.Workbook,
                        target.Worksheet,
                        "excel_paste_special");

                    int sourceRows = source.Range.Rows.Count;
                    int sourceColumns = source.Range.Columns.Count;
                    int pastedRows = transpose ? sourceColumns : sourceRows;
                    int pastedColumns = transpose ? sourceRows : sourceColumns;
                    int targetRows = target.Range.Rows.Count;
                    int targetColumns = target.Range.Columns.Count;

                    Excel.Range destination = null;
                    bool ownsDestination = false;
                    int appliedCells = 0;
                    int formattedCells = 0;
                    int widthsCopied = 0;
                    int heightsCopied = 0;
                    int conditionalFormatsCopied = 0;
                    int conditionalFormatsSkipped = 0;
                    ExcelUndoScope pasteUndo = null;
                    try
                    {
                        string sizeError = ResolvePasteDestination(
                            target.Range,
                            pastedRows,
                            pastedColumns,
                            out destination,
                            out ownsDestination);
                        if (sizeError.Length > 0)
                        {
                            return Fail(OfficeErrorCode.INVALID_PARAMS, sizeError);
                        }

                        int destinationRows = destination.Rows.Count;
                        int destinationColumns = destination.Columns.Count;
                        if (mode == "values" || mode == "valuesAndFormats")
                        {
                            if ((long)destinationRows * (long)destinationColumns > 200000)
                            {
                                return Fail(
                                    OfficeErrorCode.INVALID_PARAMS,
                                    "目标区域超过 200000 单元格，请缩小 targetRange 后分批粘贴。");
                            }
                            object[,] sourceValues = ExcelComSupport.ReadMatrix(
                                source.Range,
                                sourceRows,
                                sourceColumns);
                            object[,] targetValues = null;
                            if (skipBlanks || operate != "none")
                            {
                                targetValues = ExcelComSupport.ReadMatrix(
                                    destination,
                                    destinationRows,
                                    destinationColumns);
                            }
                            object[,] result = new object[destinationRows, destinationColumns];
                            for (int row = 0; row < destinationRows; row++)
                            {
                                for (int column = 0; column < destinationColumns; column++)
                                {
                                    object existing = targetValues == null ? null : targetValues[row, column];
                                    object raw = ReadTiledValue(
                                        sourceValues,
                                        sourceRows,
                                        sourceColumns,
                                        row,
                                        column,
                                        transpose);
                                    if (skipBlanks && (raw == null || raw is DBNull))
                                    {
                                        result[row, column] = existing;
                                        continue;
                                    }
                                    if (operate == "none")
                                    {
                                        result[row, column] = raw;
                                        continue;
                                    }
                                    double sourceNumber;
                                    double targetNumber;
                                    if (!TryReadNumber(raw, out sourceNumber))
                                    {
                                        return Fail(
                                            OfficeErrorCode.INVALID_PARAMS,
                                            "operate=" + operate + " 要求源单元格为数值，但 " +
                                            ExcelExtensionSupport.CellAddress(
                                                destination.Row + row,
                                                destination.Column + column) +
                                            " 的值是「" + ExcelExtensionSupport.CellText(raw) + "」。");
                                    }
                                    if (!TryReadNumber(existing, out targetNumber))
                                    {
                                        return Fail(
                                            OfficeErrorCode.INVALID_PARAMS,
                                            "operate=" + operate + " 要求目标单元格为数值，但 " +
                                            ExcelExtensionSupport.CellAddress(
                                                destination.Row + row,
                                                destination.Column + column) +
                                            " 的值是「" + ExcelExtensionSupport.CellText(existing) + "」。");
                                    }
                                    result[row, column] = ApplyArithmetic(targetNumber, sourceNumber, operate);
                                }
                            }

                            pasteUndo = ExcelUndoScope.Capture(
                                target.Workbook,
                                target.Worksheet,
                                destination,
                                destination.AddressLocal());
                            application.Guard(delegate { destination.Value2 = result; });
                            appliedCells = destinationRows * destinationColumns;

                            if (includeFormats)
                            {
                                string copyError = CopyFormatCore(
                                    source.Range,
                                    destination,
                                    transpose,
                                    ExcelArgumentReader.GetBool(arguments, true, "includeColumnWidths"),
                                    ExcelArgumentReader.GetBool(arguments, true, "includeRowHeights"),
                                    ExcelArgumentReader.GetBool(arguments, false, "includeConditionalFormats"),
                                    out formattedCells,
                                    out widthsCopied,
                                    out heightsCopied,
                                    out conditionalFormatsCopied,
                                    out conditionalFormatsSkipped);
                                if (copyError.Length > 0)
                                {
                                    return PartialApplyPaste(
                                        appliedCells,
                                        copyError,
                                        target,
                                        destination,
                                        mode,
                                        operate,
                                        transpose,
                                        appliedCells,
                                        formattedCells,
                                        widthsCopied,
                                        heightsCopied,
                                        conditionalFormatsCopied,
                                        conditionalFormatsSkipped);
                                }
                            }
                        }
                        else if (mode == "formulas")
                        {
                            if ((long)destinationRows * (long)destinationColumns > 200000)
                            {
                                return Fail(
                                    OfficeErrorCode.INVALID_PARAMS,
                                    "目标区域超过 200000 单元格，请缩小 targetRange 后分批粘贴。");
                            }
                            object[,] sourceFormulas = ExcelComSupport.NormalizeMatrix(
                                source.Range.Formula,
                                sourceRows,
                                sourceColumns);
                            object[,] targetFormulas = null;
                            if (skipBlanks)
                            {
                                targetFormulas = ExcelComSupport.NormalizeMatrix(
                                    destination.Formula,
                                    destinationRows,
                                    destinationColumns);
                            }
                            object[,] result = new object[destinationRows, destinationColumns];
                            for (int row = 0; row < destinationRows; row++)
                            {
                                for (int column = 0; column < destinationColumns; column++)
                                {
                                    int sourceRow = row % sourceRows;
                                    int sourceColumn = column % sourceColumns;
                                    object rawFormula = sourceFormulas[sourceRow, sourceColumn];
                                    if (skipBlanks && (rawFormula == null || rawFormula is DBNull ||
                                                       ExcelExtensionSupport.CellText(rawFormula).Length == 0))
                                    {
                                        result[row, column] = targetFormulas[row, column];
                                        continue;
                                    }
                                    string formulaText = Convert.ToString(rawFormula, CultureInfo.InvariantCulture) ?? string.Empty;
                                    if (formulaText.Length > 0 && formulaText[0] == '=')
                                    {
                                        int rowShift = destination.Row + row -
                                                       (source.Range.Row + sourceRow);
                                        int columnShift = destination.Column + column -
                                                          (source.Range.Column + sourceColumn);
                                        result[row, column] = ExcelFormulaReferenceEngine.ShiftFormula(
                                            formulaText,
                                            rowShift,
                                            columnShift);
                                    }
                                    else
                                    {
                                        result[row, column] = rawFormula;
                                    }
                                }
                            }

                            pasteUndo = ExcelUndoScope.Capture(
                                target.Workbook,
                                target.Worksheet,
                                destination,
                                destination.AddressLocal());
                            application.Guard(delegate { destination.Formula = result; });
                            appliedCells = destinationRows * destinationColumns;
                        }
                        else
                        {
                            if ((long)destinationRows * (long)destinationColumns > 20000)
                            {
                                return Fail(
                                    OfficeErrorCode.INVALID_PARAMS,
                                    "格式逐单元格复制上限为 20000 单元格，请缩小 targetRange。");
                            }
                            pasteUndo = ExcelUndoScope.Capture(
                                target.Workbook,
                                target.Worksheet,
                                destination,
                                destination.AddressLocal());
                            string copyError = CopyFormatCore(
                                source.Range,
                                destination,
                                transpose,
                                ExcelArgumentReader.GetBool(arguments, true, "includeColumnWidths"),
                                ExcelArgumentReader.GetBool(arguments, true, "includeRowHeights"),
                                ExcelArgumentReader.GetBool(arguments, false, "includeConditionalFormats"),
                                out formattedCells,
                                out widthsCopied,
                                out heightsCopied,
                                out conditionalFormatsCopied,
                                out conditionalFormatsSkipped);
                            if (copyError.Length > 0)
                            {
                                return PartialApplyPaste(
                                    Math.Max(formattedCells, widthsCopied + heightsCopied),
                                    copyError,
                                    target,
                                    destination,
                                    mode,
                                    operate,
                                    transpose,
                                    0,
                                    formattedCells,
                                    widthsCopied,
                                    heightsCopied,
                                    conditionalFormatsCopied,
                                    conditionalFormatsSkipped);
                            }
                        }

                        string targetRef = ExcelComSupport.FullAddress(
                            target.SheetName,
                            destination.AddressLocal());
                        Observation observation = Observation.Create(
                            "paste_special",
                            "选择性粘贴完成：" + targetRef + "（mode=" + mode +
                            (operate == "none" ? string.Empty : "，operate=" + operate) + "）。",
                            appliedCells > 0 || formattedCells > 0 || widthsCopied > 0 || heightsCopied > 0);
                        observation.AddTargetRef(targetRef);

                        JObject data = BuildPasteData(
                            target,
                            destination,
                            mode,
                            operate,
                            transpose,
                            appliedCells,
                            formattedCells,
                            widthsCopied,
                            heightsCopied,
                            conditionalFormatsCopied,
                            conditionalFormatsSkipped);
                        if (pasteUndo != null)
                        {
                            pasteUndo.Complete(
                                target.Workbook,
                                target.Worksheet,
                                destination,
                                destination.AddressLocal(),
                                true,
                                "选择性粘贴 " + mode + " " + targetRef,
                                observation,
                                data,
                                "选择性粘贴只登记目标区域值/公式/数字格式快照；字体、填充、边框、行高列宽与条件格式不在撤销范围内。");
                        }
                        else
                        {
                            ExcelUndoScope.RecordUnavailable(
                                data,
                                "选择性粘贴未发生可快照的内容变化，未登记自动撤销。");
                        }

                        return OperationResult.Succeed(
                            "选择性粘贴完成：" + targetRef + "（" + mode + "）。",
                            data,
                            observation);
                    }
                    finally
                    {
                        if (ownsDestination)
                        {
                            ComHelper.Release(destination);
                        }
                    }
                }
            }
        }

        private static OperationResult PartialApplyPaste(
            int succeededCells,
            string error,
            ExcelRangeContext target,
            Excel.Range destination,
            string mode,
            string operate,
            bool transpose,
            int appliedCells,
            int formattedCells,
            int widthsCopied,
            int heightsCopied,
            int conditionalFormatsCopied,
            int conditionalFormatsSkipped)
        {
            JObject data = BuildPasteData(
                target,
                destination,
                mode,
                operate,
                transpose,
                appliedCells,
                formattedCells,
                widthsCopied,
                heightsCopied,
                conditionalFormatsCopied,
                conditionalFormatsSkipped);
            Observation observation = Observation.Create(
                "paste_special",
                "选择性粘贴部分完成：" + error,
                succeededCells > 0 || formattedCells > 0);
            observation.AddTargetRef(
                ExcelComSupport.FullAddress(target.SheetName, destination.AddressLocal()));
            ExcelUndoScope.RecordUnavailable(
                data,
                "已应用的部分格式/值不会自动回滚；请对照 formattedCells/appliedCells 人工检查。");
            return OperationResult.PartialApplied(
                succeededCells,
                "已应用的 " + succeededCells.ToString(CultureInfo.InvariantCulture) +
                " 个单元格不会自动回滚；" + error,
                observation,
                data);
        }

        private static JObject BuildPasteData(
            ExcelRangeContext target,
            Excel.Range destination,
            string mode,
            string operate,
            bool transpose,
            int appliedCells,
            int formattedCells,
            int widthsCopied,
            int heightsCopied,
            int conditionalFormatsCopied,
            int conditionalFormatsSkipped)
        {
            JObject data = new JObject();
            data["sheet"] = target.SheetName;
            data["targetAddress"] = destination.AddressLocal();
            data["mode"] = mode;
            data["operate"] = operate;
            data["transpose"] = transpose;
            data["appliedCells"] = appliedCells;
            data["formattedCells"] = formattedCells;
            data["columnWidthsCopied"] = widthsCopied;
            data["rowHeightsCopied"] = heightsCopied;
            data["conditionalFormatsCopied"] = conditionalFormatsCopied;
            data["conditionalFormatsSkipped"] = conditionalFormatsSkipped;
            return data;
        }

        private static string ResolvePasteDestination(
            Excel.Range targetRange,
            int sourceRows,
            int sourceColumns,
            out Excel.Range destination,
            out bool ownsDestination)
        {
            destination = null;
            ownsDestination = false;
            int targetRows = targetRange.Rows.Count;
            int targetColumns = targetRange.Columns.Count;
            if (targetRows == 1 && targetColumns == 1)
            {
                destination = targetRange.Resize[sourceRows, sourceColumns];
                ownsDestination = true;
                return string.Empty;
            }
            if (targetRows == sourceRows && targetColumns == sourceColumns)
            {
                destination = targetRange;
                return string.Empty;
            }
            if (targetRows % sourceRows == 0 && targetColumns % sourceColumns == 0)
            {
                destination = targetRange;
                return string.Empty;
            }
            return "目标区域 " + targetRange.AddressLocal() + "（" + targetRows + "×" + targetColumns +
                   "）与源区域（" + sourceRows + "×" + sourceColumns +
                   "）尺寸不匹配；请使用单个起始单元格，或使用源区域尺寸的整数倍。";
        }

        private static object ReadTiledValue(
            object[,] source,
            int sourceRows,
            int sourceColumns,
            int row,
            int column,
            bool transpose)
        {
            int sourceRow = transpose ? column % sourceRows : row % sourceRows;
            int sourceColumn = transpose ? row % sourceColumns : column % sourceColumns;
            return source[sourceRow, sourceColumn];
        }

        private static bool TryReadNumber(object value, out double number)
        {
            number = 0.0;
            if (value == null || value is DBNull)
            {
                return false;
            }
            if (value is double || value is float || value is int || value is long ||
                value is short || value is decimal || value is byte)
            {
                number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                return true;
            }
            return false;
        }

        private static double ApplyArithmetic(double target, double source, string operate)
        {
            if (operate == "add")
            {
                return target + source;
            }
            if (operate == "subtract")
            {
                return target - source;
            }
            if (operate == "multiply")
            {
                return target * source;
            }
            if (operate == "divide")
            {
                if (Math.Abs(source) < double.Epsilon)
                {
                    throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "operate=divide 时源单元格不能为 0。");
                }
                return target / source;
            }
            return source;
        }

        private static string CopyFormatCore(
            Excel.Range source,
            Excel.Range destination,
            bool transpose,
            bool includeColumnWidths,
            bool includeRowHeights,
            bool includeConditionalFormats,
            out int formattedCells,
            out int widthsCopied,
            out int heightsCopied,
            out int conditionalFormatsCopied,
            out int conditionalFormatsSkipped)
        {
            formattedCells = 0;
            widthsCopied = 0;
            heightsCopied = 0;
            conditionalFormatsCopied = 0;
            conditionalFormatsSkipped = 0;

            int sourceRows = source.Rows.Count;
            int sourceColumns = source.Columns.Count;
            int destinationRows = destination.Rows.Count;
            int destinationColumns = destination.Columns.Count;
            for (int row = 0; row < destinationRows; row++)
            {
                for (int column = 0; column < destinationColumns; column++)
                {
                    int sourceRow = transpose ? column % sourceRows : row % sourceRows;
                    int sourceColumn = transpose ? row % sourceColumns : column % sourceColumns;
                    Excel.Range sourceCell = null;
                    Excel.Range destinationCell = null;
                    try
                    {
                        sourceCell = (Excel.Range)source.Cells[sourceRow + 1, sourceColumn + 1];
                        destinationCell = (Excel.Range)destination.Cells[row + 1, column + 1];
                        CopyCellFormat(sourceCell, destinationCell);
                        formattedCells++;
                    }
                    catch (Exception ex)
                    {
                        return "复制单元格格式在 " + ExcelExtensionSupport.CellAddress(
                                   destination.Row + row,
                                   destination.Column + column) + " 失败：" + ex.Message;
                    }
                    finally
                    {
                        ComHelper.Release(destinationCell, sourceCell);
                    }
                }
            }

            if (includeColumnWidths)
            {
                for (int column = 0; column < destinationColumns; column++)
                {
                    int sourceColumn = column % sourceColumns;
                    Excel.Range sourceColumnRange = null;
                    Excel.Range destinationColumnRange = null;
                    try
                    {
                        sourceColumnRange = (Excel.Range)source.Columns[sourceColumn + 1];
                        destinationColumnRange = (Excel.Range)destination.Columns[column + 1];
                        destinationColumnRange.ColumnWidth = sourceColumnRange.ColumnWidth;
                        widthsCopied++;
                    }
                    catch (Exception ex)
                    {
                        return "复制列宽失败：" + ex.Message;
                    }
                    finally
                    {
                        ComHelper.Release(destinationColumnRange, sourceColumnRange);
                    }
                }
            }

            if (includeRowHeights)
            {
                for (int row = 0; row < destinationRows; row++)
                {
                    int sourceRow = row % sourceRows;
                    Excel.Range sourceRowRange = null;
                    Excel.Range destinationRowRange = null;
                    try
                    {
                        sourceRowRange = (Excel.Range)source.Rows[sourceRow + 1];
                        destinationRowRange = (Excel.Range)destination.Rows[row + 1];
                        destinationRowRange.RowHeight = sourceRowRange.RowHeight;
                        heightsCopied++;
                    }
                    catch (Exception ex)
                    {
                        return "复制行高失败：" + ex.Message;
                    }
                    finally
                    {
                        ComHelper.Release(destinationRowRange, sourceRowRange);
                    }
                }
            }

            if (includeConditionalFormats)
            {
                Excel.FormatConditions sourceConditions = null;
                Excel.FormatConditions destinationConditions = null;
                try
                {
                    sourceConditions = source.FormatConditions;
                    destinationConditions = destination.FormatConditions;
                    int count = sourceConditions.Count;
                    for (int index = 1; index <= count; index++)
                    {
                        object conditionObject = null;
                        object addedCondition = null;
                        try
                        {
                            conditionObject = sourceConditions.Item(index);
                            if (!CloneCondition(conditionObject, destinationConditions, out addedCondition))
                            {
                                conditionalFormatsSkipped++;
                            }
                            else
                            {
                                conditionalFormatsCopied++;
                            }
                        }
                        catch (Exception ex)
                        {
                            conditionalFormatsSkipped++;
                            Console.Error.WriteLine("[excel_copy_format] 复制条件格式失败：" + ex.Message);
                        }
                        finally
                        {
                            ComHelper.Release(addedCondition, conditionObject);
                        }
                    }
                }
                finally
                {
                    ComHelper.Release(destinationConditions, sourceConditions);
                }
            }

            return string.Empty;
        }

        private static bool CloneCondition(
            object conditionObject,
            Excel.FormatConditions destinationConditions,
            out object addedCondition)
        {
            addedCondition = null;
            Excel.FormatCondition condition = conditionObject as Excel.FormatCondition;
            if (condition != null)
            {
                object formula1 = condition.Formula1;
                object formula2 = condition.Formula2;
                if (formula2 != null && formula2 is string && ((string)formula2).Length == 0)
                {
                    formula2 = Type.Missing;
                }
                Excel.FormatCondition added = (Excel.FormatCondition)destinationConditions.Add(
                    Type: (Excel.XlFormatConditionType)Convert.ToInt32(
                        condition.Type,
                        CultureInfo.InvariantCulture),
                    Operator: (Excel.XlFormatConditionOperator)Convert.ToInt32(
                        condition.Operator,
                        CultureInfo.InvariantCulture),
                    Formula1: formula1,
                    Formula2: formula2);
                addedCondition = added;
                try
                {
                    Excel.Interior sourceInterior = null;
                    Excel.Interior targetInterior = null;
                    Excel.Font sourceFont = null;
                    Excel.Font targetFont = null;
                    try
                    {
                        sourceInterior = condition.Interior;
                        targetInterior = added.Interior;
                        targetInterior.Color = sourceInterior.Color;
                    }
                    finally
                    {
                        ComHelper.Release(targetInterior, sourceInterior);
                    }
                    try
                    {
                        sourceFont = condition.Font;
                        targetFont = added.Font;
                        targetFont.Color = sourceFont.Color;
                        targetFont.Bold = sourceFont.Bold;
                        targetFont.Italic = sourceFont.Italic;
                    }
                    finally
                    {
                        ComHelper.Release(targetFont, sourceFont);
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[excel_copy_format] 复制条件格式样式失败：" + ex.Message);
                }
                return true;
            }

            Excel.ColorScale colorScale = conditionObject as Excel.ColorScale;
            if (colorScale != null)
            {
                addedCondition = destinationConditions.AddColorScale(3);
                return true;
            }

            Excel.Databar databar = conditionObject as Excel.Databar;
            if (databar != null)
            {
                addedCondition = destinationConditions.AddDatabar();
                return true;
            }

            Excel.IconSetCondition iconSet = conditionObject as Excel.IconSetCondition;
            if (iconSet != null)
            {
                addedCondition = destinationConditions.AddIconSetCondition();
                return true;
            }
            return false;
        }

        private static void CopyCellFormat(Excel.Range source, Excel.Range destination)
        {
            destination.NumberFormat = source.NumberFormat;
            destination.Locked = source.Locked;
            destination.FormulaHidden = source.FormulaHidden;
            destination.HorizontalAlignment = source.HorizontalAlignment;
            destination.VerticalAlignment = source.VerticalAlignment;
            destination.WrapText = source.WrapText;
            destination.Orientation = source.Orientation;
            destination.IndentLevel = source.IndentLevel;
            destination.ShrinkToFit = source.ShrinkToFit;
            destination.ReadingOrder = source.ReadingOrder;

            Excel.Font sourceFont = null;
            Excel.Font destinationFont = null;
            try
            {
                sourceFont = source.Font;
                destinationFont = destination.Font;
                destinationFont.Name = sourceFont.Name;
                destinationFont.Size = sourceFont.Size;
                destinationFont.Bold = sourceFont.Bold;
                destinationFont.Italic = sourceFont.Italic;
                destinationFont.Underline = sourceFont.Underline;
                destinationFont.Strikethrough = sourceFont.Strikethrough;
                destinationFont.Color = sourceFont.Color;
                destinationFont.Superscript = sourceFont.Superscript;
                destinationFont.Subscript = sourceFont.Subscript;
            }
            finally
            {
                ComHelper.Release(destinationFont, sourceFont);
            }

            Excel.Interior sourceInterior = null;
            Excel.Interior destinationInterior = null;
            try
            {
                sourceInterior = source.Interior;
                destinationInterior = destination.Interior;
                destinationInterior.Color = sourceInterior.Color;
                destinationInterior.Pattern = sourceInterior.Pattern;
                destinationInterior.PatternColor = sourceInterior.PatternColor;
            }
            finally
            {
                ComHelper.Release(destinationInterior, sourceInterior);
            }

            int[] borderIndexes = new int[] { 7, 8, 9, 10, 11, 12 };
            foreach (int borderIndex in borderIndexes)
            {
                Excel.Border sourceBorder = null;
                Excel.Border destinationBorder = null;
                try
                {
                    sourceBorder = source.Borders[(Excel.XlBordersIndex)borderIndex];
                    destinationBorder = destination.Borders[(Excel.XlBordersIndex)borderIndex];
                    destinationBorder.LineStyle = sourceBorder.LineStyle;
                    destinationBorder.Weight = sourceBorder.Weight;
                    destinationBorder.Color = sourceBorder.Color;
                }
                finally
                {
                    ComHelper.Release(destinationBorder, sourceBorder);
                }
            }
        }

        private static OperationResult ExecuteFormulaToValue(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange", "address");
            if (string.IsNullOrWhiteSpace(address))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 range/targetRange 参数。");
            }

            using (ExcelRangeContext context = RangeResolver.Resolve(
                       application,
                       address,
                       false,
                       "excel_formula_to_value"))
            {
                ExcelComSupport.EnsureWritable(
                    application,
                    context.Workbook,
                    context.Worksheet,
                    "excel_formula_to_value");
                int rows = context.Range.Rows.Count;
                int columns = context.Range.Columns.Count;
                if ((long)rows * (long)columns > 200000)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_RANGE,
                        "公式固化区域超过 200000 单元格，请缩小 range 后分批处理。");
                }

                object[,] formulas = ExcelComSupport.NormalizeMatrix(
                    context.Range.Formula,
                    rows,
                    columns);
                int formulaCells = 0;
                for (int row = 0; row < rows; row++)
                {
                    for (int column = 0; column < columns; column++)
                    {
                        object value = formulas[row, column];
                        if (value is string && ((string)value).Length > 0 && ((string)value)[0] == '=')
                        {
                            formulaCells++;
                        }
                    }
                }

                JObject data = new JObject();
                data["sheet"] = context.SheetName;
                data["address"] = context.Range.AddressLocal();
                data["formulaCells"] = formulaCells;
                data["convertedCells"] = 0;
                if (formulaCells == 0)
                {
                    ExcelUndoScope.RecordUnavailable(data, "没有公式，无需撤销。");
                    Observation unchanged = Observation.Create(
                        "formula_to_value",
                        "目标区域没有公式，未做修改。",
                        false);
                    unchanged.AddTargetRef(
                        ExcelComSupport.FullAddress(context.SheetName, context.Range.AddressLocal()));
                    return OperationResult.Succeed("目标区域没有公式，未做修改。", data, unchanged);
                }

                object[,] values = ExcelComSupport.ReadMatrix(context.Range, rows, columns);
                string rangeAddress = context.Range.AddressLocal();
                ExcelUndoScope undo = ExcelUndoScope.Capture(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress);
                application.Guard(delegate { context.Range.Value2 = values; });

                object[,] afterFormulas = ExcelComSupport.NormalizeMatrix(
                    context.Range.Formula,
                    rows,
                    columns);
                int remainingFormulas = 0;
                for (int row = 0; row < rows; row++)
                {
                    for (int column = 0; column < columns; column++)
                    {
                        object value = afterFormulas[row, column];
                        if (value is string && ((string)value).Length > 0 && ((string)value)[0] == '=')
                        {
                            remainingFormulas++;
                        }
                    }
                }

                string targetRef = ExcelComSupport.FullAddress(context.SheetName, rangeAddress);
                data["convertedCells"] = Math.Max(0, formulaCells - remainingFormulas);
                data["remainingFormulaCells"] = remainingFormulas;

                Observation observation = Observation.Create(
                    "formula_to_value",
                    "已把 " + data.Value<int>("convertedCells") + " 个公式固化为值：" + targetRef + "。",
                    data.Value<int>("convertedCells") > 0);
                observation.AddTargetRef(targetRef);

                undo.Complete(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress,
                    true,
                    "公式固化为值 " + targetRef,
                    observation,
                    data,
                    "公式固化会删除公式文本；自动撤销仅恢复值/公式/数字格式，不恢复依赖此公式的图表/数据验证引用。");

                if (remainingFormulas > 0)
                {
                    return OperationResult.PartialApplied(
                        data.Value<int>("convertedCells"),
                        "仍有 " + remainingFormulas.ToString(CultureInfo.InvariantCulture) +
                        " 个公式未固化（可能是数组公式或受保护区域）；请检查后手动处理。",
                        observation,
                        data);
                }

                return OperationResult.Succeed(
                    "已把 " + formulaCells.ToString(CultureInfo.InvariantCulture) +
                    " 个公式固化为值：" + targetRef + "。",
                    data,
                    observation);
            }
        }

        private static OperationResult ExecuteCopyFormat(Excel.Application application, JObject arguments)
        {
            string sourceAddress = ExcelArgumentReader.GetRangeAddress(
                arguments,
                "sourceRange",
                "source",
                "fromRange");
            string targetAddress = ExcelArgumentReader.GetRangeAddress(
                arguments,
                "targetRange",
                "target",
                "toRange");
            if (string.IsNullOrWhiteSpace(sourceAddress) || string.IsNullOrWhiteSpace(targetAddress))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 sourceRange 或 targetRange 参数。");
            }
            bool transpose = ExcelArgumentReader.GetBool(arguments, false, "transpose");

            using (ExcelRangeContext source = RangeResolver.Resolve(
                       application,
                       sourceAddress,
                       false,
                       "excel_copy_format"))
            {
                using (ExcelRangeContext target = RangeResolver.Resolve(
                           application,
                           targetAddress,
                           false,
                           "excel_copy_format"))
                {
                    ExcelComSupport.EnsureWritable(
                        application,
                        target.Workbook,
                        target.Worksheet,
                        "excel_copy_format");
                    int sourceRows = source.Range.Rows.Count;
                    int sourceColumns = source.Range.Columns.Count;
                    int pastedRows = transpose ? sourceColumns : sourceRows;
                    int pastedColumns = transpose ? sourceRows : sourceColumns;

                    Excel.Range destination = null;
                    bool ownsDestination = false;
                    try
                    {
                        string sizeError = ResolvePasteDestination(
                            target.Range,
                            pastedRows,
                            pastedColumns,
                            out destination,
                            out ownsDestination);
                        if (sizeError.Length > 0)
                        {
                            return Fail(OfficeErrorCode.INVALID_PARAMS, sizeError);
                        }
                        int destinationRows = destination.Rows.Count;
                        int destinationColumns = destination.Columns.Count;
                        if ((long)destinationRows * (long)destinationColumns > 20000)
                        {
                            return Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "格式逐单元格复制上限为 20000 单元格，请缩小 targetRange。");
                        }

                        ExcelUndoScope undo = ExcelUndoScope.Capture(
                            target.Workbook,
                            target.Worksheet,
                            destination,
                            destination.AddressLocal());
                        int formattedCells;
                        int widthsCopied;
                        int heightsCopied;
                        int conditionalFormatsCopied;
                        int conditionalFormatsSkipped;
                        string copyError = CopyFormatCore(
                            source.Range,
                            destination,
                            transpose,
                            ExcelArgumentReader.GetBool(arguments, true, "includeColumnWidths"),
                            ExcelArgumentReader.GetBool(arguments, true, "includeRowHeights"),
                            ExcelArgumentReader.GetBool(arguments, false, "includeConditionalFormats"),
                            out formattedCells,
                            out widthsCopied,
                            out heightsCopied,
                            out conditionalFormatsCopied,
                            out conditionalFormatsSkipped);

                        string targetRef = ExcelComSupport.FullAddress(
                            target.SheetName,
                            destination.AddressLocal());
                        JObject data = new JObject();
                        data["sheet"] = target.SheetName;
                        data["targetAddress"] = destination.AddressLocal();
                        data["transpose"] = transpose;
                        data["formattedCells"] = formattedCells;
                        data["columnWidthsCopied"] = widthsCopied;
                        data["rowHeightsCopied"] = heightsCopied;
                        data["conditionalFormatsCopied"] = conditionalFormatsCopied;
                        data["conditionalFormatsSkipped"] = conditionalFormatsSkipped;

                        if (copyError.Length > 0)
                        {
                            ExcelUndoScope.RecordUnavailable(
                                data,
                                "已复制的部分格式不会自动回滚；请对照 formattedCells 人工检查。");
                            Observation partialObservation = Observation.Create(
                                "copy_format",
                                "格式复制部分完成：" + copyError,
                                formattedCells > 0);
                            partialObservation.AddTargetRef(targetRef);
                            return OperationResult.PartialApplied(
                                formattedCells,
                                "已复制的 " + formattedCells.ToString(CultureInfo.InvariantCulture) +
                                " 个单元格格式不会自动回滚；" + copyError,
                                partialObservation,
                                data);
                        }

                        ExcelUndoScope.RecordUnavailable(
                            data,
                            "逐单元格复制字体/填充/边框/行高列宽不在区域快照中，未登记完整自动撤销；数字格式已随快照恢复。");
                        Observation observation = Observation.Create(
                            "copy_format",
                            "已复制格式到 " + targetRef + "（" + formattedCells + " 个单元格）。",
                            formattedCells > 0 || widthsCopied > 0 || heightsCopied > 0 ||
                            conditionalFormatsCopied > 0);
                        observation.AddTargetRef(targetRef);
                        return OperationResult.Succeed(
                            "格式复制完成：" + targetRef + "。",
                            data,
                            observation);
                    }
                    finally
                    {
                        if (ownsDestination)
                        {
                            ComHelper.Release(destination);
                        }
                    }
                }
            }
        }

        private static OperationResult ExecuteHyperlinkManage(Excel.Application application, JObject arguments)
        {
            string action = ExcelArgumentReader.GetTrimmedString(arguments, "action").ToLowerInvariant();
            if (action.Length == 0)
            {
                action = "list";
            }
            if (action != "list" && action != "set" && action != "clear")
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "action 仅支持 list/set/clear。");
            }

            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange", "address");
            if (string.IsNullOrWhiteSpace(address))
            {
                string sheetName = ExcelArgumentReader.GetSheetName(arguments);
                address = string.IsNullOrWhiteSpace(sheetName)
                    ? "usedrange"
                    : ExcelComSupport.QuoteSheetName(sheetName) + "!usedrange";
            }

            using (ExcelRangeContext context = RangeResolver.Resolve(
                       application,
                       address,
                       false,
                       "excel_hyperlink_manage"))
            {
                if (action == "list")
                {
                    return ListHyperlinks(context);
                }
                ExcelComSupport.EnsureWritable(
                    application,
                    context.Workbook,
                    context.Worksheet,
                    "excel_hyperlink_manage");

                if (action == "clear")
                {
                    int before = ReadHyperlinkCount(context.Range);
                    application.Guard(delegate { context.Range.ClearHyperlinks(); });
                    int after = ReadHyperlinkCount(context.Range);
                    JObject data = new JObject();
                    data["sheet"] = context.SheetName;
                    data["address"] = context.Range.AddressLocal();
                    data["removedCount"] = Math.Max(0, before - after);
                    data["remainingCount"] = after;
                    ExcelUndoScope.RecordUnavailable(
                        data,
                        "超链接不在区域快照中，未登记自动撤销。");
                    Observation observation = Observation.Create(
                        "hyperlink_manage",
                        "已清除 " + data.Value<int>("removedCount") + " 个超链接。",
                        before != after);
                    observation.AddTargetRef(
                        ExcelComSupport.FullAddress(context.SheetName, context.Range.AddressLocal()));
                    return OperationResult.Succeed(
                        "已清除 " + data.Value<int>("removedCount") + " 个超链接。",
                        data,
                        observation);
                }

                string linkType = ExcelArgumentReader.GetTrimmedString(arguments, "type", "linkType").ToLowerInvariant();
                if (linkType.Length == 0)
                {
                    linkType = "auto";
                }
                if (linkType != "auto" && linkType != "internal" && linkType != "external")
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "type 仅支持 internal/external/auto。");
                }
                string linkAddress = ExcelArgumentReader.GetTrimmedString(
                    arguments,
                    "linkAddress",
                    "url",
                    "address");
                string subAddress = ExcelArgumentReader.GetTrimmedString(
                    arguments,
                    "subAddress",
                    "subaddress",
                    "target");
                if (linkType == "internal" && subAddress.Length == 0)
                {
                    subAddress = linkAddress;
                    linkAddress = string.Empty;
                }
                if (linkType == "external" && linkAddress.Length == 0)
                {
                    linkAddress = subAddress;
                }
                if (linkAddress.Length == 0 && subAddress.Length == 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "set 需要 address 或 subAddress。");
                }
                string displayText = ExcelArgumentReader.GetString(
                    arguments,
                    "displayText",
                    "text",
                    "textToDisplay");
                string screenTip = ExcelArgumentReader.GetString(arguments, "screenTip", "tooltip");

                Excel.Range anchor = null;
                Excel.Hyperlink hyperlink = null;
                try
                {
                    anchor = (Excel.Range)context.Range.Cells[1, 1];
                    ExcelUndoScope undo = ExcelUndoScope.Capture(
                        context.Workbook,
                        context.Worksheet,
                        anchor,
                        anchor.AddressLocal());
                    application.Guard(delegate
                    {
                        if (ReadHyperlinkCount(anchor) > 0)
                        {
                            anchor.Hyperlinks.Delete();
                        }
                        hyperlink = (Excel.Hyperlink)anchor.Hyperlinks.Add(
                            anchor,
                            linkAddress,
                            subAddress.Length == 0 ? Type.Missing : subAddress,
                            screenTip,
                            displayText);
                    });

                    JObject data = BuildHyperlinkData(
                        context.SheetName,
                        anchor,
                        hyperlink,
                        displayText);
                    data["action"] = "set";
                    Observation observation = Observation.Create(
                        "hyperlink_manage",
                        "已设置超链接：" + data.Value<string>("address") + "。",
                        true);
                    observation.AddTargetRef(
                        ExcelComSupport.FullAddress(context.SheetName, anchor.AddressLocal()));
                    undo.Complete(
                        context.Workbook,
                        context.Worksheet,
                        anchor,
                        anchor.AddressLocal(),
                        true,
                        "设置超链接 " + anchor.AddressLocal(),
                        observation,
                        data,
                        "超链接对象不在区域快照中，未登记完整自动撤销。");
                    return OperationResult.Succeed("超链接已设置。", data, observation);
                }
                finally
                {
                    ComHelper.Release(hyperlink, anchor);
                }
            }
        }

        private static OperationResult ListHyperlinks(ExcelRangeContext context)
        {
            Excel.Hyperlinks hyperlinks = null;
            JArray items = new JArray();
            try
            {
                hyperlinks = context.Range.Hyperlinks;
                int count = hyperlinks.Count;
                for (int index = 1; index <= count; index++)
                {
                    Excel.Hyperlink hyperlink = null;
                    try
                    {
                        hyperlink = (Excel.Hyperlink)hyperlinks[index];
                        JObject item = BuildHyperlinkData(
                            context.SheetName,
                            null,
                            hyperlink,
                            string.Empty);
                        items.Add(item);
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[excel_hyperlink_manage] 读取超链接失败：" + ex.Message);
                    }
                    finally
                    {
                        ComHelper.Release(hyperlink);
                    }
                }
            }
            finally
            {
                ComHelper.Release(hyperlinks);
            }

            JObject data = new JObject();
            data["sheet"] = context.SheetName;
            data["address"] = context.Range.AddressLocal();
            data["count"] = items.Count;
            data["hyperlinks"] = items;
            ExcelUndoScope.RecordUnavailable(data, "超链接列举为只读操作，无需撤销。");
            Observation observation = Observation.Create(
                "hyperlink_manage",
                "已列出 " + items.Count + " 个超链接。",
                false);
            observation.AddTargetRef(
                ExcelComSupport.FullAddress(context.SheetName, context.Range.AddressLocal()));
            return OperationResult.Succeed("已列出 " + items.Count + " 个超链接。", data, observation);
        }

        private static JObject BuildHyperlinkData(
            string sheetName,
            Excel.Range anchor,
            Excel.Hyperlink hyperlink,
            string displayTextOverride)
        {
            JObject item = new JObject();
            string address = string.Empty;
            string subAddress = string.Empty;
            string displayText = string.Empty;
            string screenTip = string.Empty;
            string anchorAddress = string.Empty;
            int typeValue = 0;
            if (hyperlink != null)
            {
                try { address = hyperlink.Address ?? string.Empty; } catch (Exception) { }
                try { subAddress = hyperlink.SubAddress ?? string.Empty; } catch (Exception) { }
                try { displayText = hyperlink.TextToDisplay ?? string.Empty; } catch (Exception) { }
                try { screenTip = hyperlink.ScreenTip ?? string.Empty; } catch (Exception) { }
                try
                {
                    object rawType = ReadComProperty(hyperlink, "Type");
                    if (rawType != null)
                    {
                        typeValue = Convert.ToInt32(rawType, CultureInfo.InvariantCulture);
                    }
                }
                catch (Exception)
                {
                }
                Excel.Range anchorRange = null;
                try
                {
                    anchorRange = hyperlink.Range;
                    if (anchorRange != null)
                    {
                        anchorAddress = anchorRange.AddressLocal();
                    }
                }
                catch (Exception)
                {
                }
                finally
                {
                    ComHelper.Release(anchorRange);
                }
            }
            if (anchorAddress.Length == 0 && anchor != null)
            {
                anchorAddress = anchor.AddressLocal();
            }
            if (displayText.Length == 0)
            {
                displayText = displayTextOverride;
            }
            string linkType = address.Length > 0 ? "external" : "internal";
            item["sheet"] = sheetName;
            item["cell"] = anchorAddress;
            item["address"] = address;
            item["subAddress"] = subAddress;
            item["displayText"] = displayText;
            item["screenTip"] = screenTip;
            item["type"] = linkType;
            item["typeValue"] = typeValue;
            return item;
        }

        private static object ReadComProperty(object target, string propertyName)
        {
            if (target == null)
            {
                return null;
            }
            try
            {
                return target.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.GetProperty,
                    null,
                    target,
                    new object[0]);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static int ReadHyperlinkCount(Excel.Range range)
        {
            if (range == null)
            {
                return 0;
            }
            Excel.Hyperlinks hyperlinks = null;
            try
            {
                hyperlinks = range.Hyperlinks;
                return hyperlinks.Count;
            }
            catch (Exception)
            {
                return 0;
            }
            finally
            {
                ComHelper.Release(hyperlinks);
            }
        }

        private static OperationResult ExecuteListMergedCells(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange", "address");
            if (string.IsNullOrWhiteSpace(address))
            {
                string sheetName = ExcelArgumentReader.GetSheetName(arguments);
                address = string.IsNullOrWhiteSpace(sheetName)
                    ? "usedrange"
                    : ExcelComSupport.QuoteSheetName(sheetName) + "!usedrange";
            }
            int maxResults = ExcelArgumentReader.GetInt(arguments, 500, "maxResults", "limit");
            maxResults = Math.Max(1, Math.Min(maxResults, 5000));

            using (ExcelRangeContext context = RangeResolver.Resolve(
                       application,
                       address,
                       false,
                       "excel_list_merged_cells"))
            {
                int rows = context.Range.Rows.Count;
                int columns = context.Range.Columns.Count;
                long scanned = 0;
                long scanLimit = 200000;
                bool truncated = false;
                JArray mergedCells = new JArray();
                HashSet<string> seenAreas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                for (int row = 1; row <= rows; row++)
                {
                    for (int column = 1; column <= columns; column++)
                    {
                        scanned++;
                        if (scanned > scanLimit)
                        {
                            truncated = true;
                            break;
                        }
                        Excel.Range cell = null;
                        try
                        {
                            cell = (Excel.Range)context.Range.Cells[row, column];
                            object mergedValue = cell.MergeCells;
                            if (!(mergedValue is bool) || !(bool)mergedValue)
                            {
                                continue;
                            }
                            Excel.Range mergeArea = null;
                            Excel.Range topLeft = null;
                            try
                            {
                                mergeArea = cell.MergeArea;
                                string areaAddress = mergeArea.AddressLocal();
                                if (!seenAreas.Add(areaAddress))
                                {
                                    continue;
                                }
                                if (mergedCells.Count >= maxResults)
                                {
                                    truncated = true;
                                    continue;
                                }
                                topLeft = (Excel.Range)mergeArea.Cells[1, 1];
                                JObject item = new JObject();
                                item["address"] = areaAddress;
                                item["rows"] = mergeArea.Rows.Count;
                                item["columns"] = mergeArea.Columns.Count;
                                item["value"] = ExcelComSupport.ToJToken(topLeft.Value2);
                                object formula = topLeft.Formula;
                                item["formula"] = formula == null
                                    ? JValue.CreateNull()
                                    : new JValue(Convert.ToString(formula, CultureInfo.InvariantCulture));
                                mergedCells.Add(item);
                            }
                            finally
                            {
                                ComHelper.Release(topLeft, mergeArea);
                            }
                        }
                        finally
                        {
                            ComHelper.Release(cell);
                        }
                    }
                    if (truncated)
                    {
                        break;
                    }
                }

                JObject data = new JObject();
                data["sheet"] = context.SheetName;
                data["address"] = context.Range.AddressLocal();
                data["count"] = mergedCells.Count;
                data["mergedCells"] = mergedCells;
                data["scannedCells"] = scanned;
                data["truncated"] = truncated;
                ExcelUndoScope.RecordUnavailable(data, "合并单元格列举为只读操作，无需撤销。");
                Observation observation = Observation.Create(
                    "list_merged_cells",
                    "已列出 " + mergedCells.Count + " 个合并区域。",
                    false);
                observation.AddTargetRef(
                    ExcelComSupport.FullAddress(context.SheetName, context.Range.AddressLocal()));
                return OperationResult.Succeed(
                    "已列出 " + mergedCells.Count + " 个合并区域" +
                    (truncated ? "（扫描/结果已截断）" : string.Empty) + "。",
                    data,
                    observation);
            }
        }
    }
}
