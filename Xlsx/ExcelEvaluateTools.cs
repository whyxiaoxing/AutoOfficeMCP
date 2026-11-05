using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>任意公式求值（Application.Evaluate + 临时单元格兜底）。</summary>
    public sealed partial class ExcelToolExecutor
    {
        private static OperationResult ExecuteEvaluate(Excel.Application application, JObject arguments)
        {
            string formula = ExcelArgumentReader.GetString(arguments, "formula", "expression");
            if (string.IsNullOrWhiteSpace(formula))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 formula 参数。");
            }
            if (!formula.TrimStart().StartsWith("=", StringComparison.Ordinal))
            {
                formula = "=" + formula.Trim();
            }

            object evaluated;
            try
            {
                evaluated = application.Evaluate(formula);
            }
            catch (COMException)
            {
                evaluated = EvaluateInTemporaryCell(application, formula);
            }

            JToken resultToken = NormalizeEvaluateResult(evaluated);
            string resultType = DescribeResultType(evaluated);

            JObject data = new JObject();
            data["formula"] = formula;
            data["result"] = resultToken;
            data["resultType"] = resultType;

            string writeTarget = ExcelArgumentReader.GetTrimmedString(arguments, "writeTarget", "targetRange");
            if (string.IsNullOrWhiteSpace(writeTarget))
            {
                data["written"] = false;
                return OperationResult.Succeed("公式求值完成：" + formula, data);
            }

            int rows;
            int columns;
            bool scalar;
            object[,] matrix = ExcelComSupport.JsonToMatrix(resultToken, out rows, out columns, out scalar);
            using (ExcelRangeContext target = RangeResolver.Resolve(application, writeTarget, false, "excel_evaluate"))
            {
                ExcelComSupport.EnsureWritable(application, target.Workbook, target.Worksheet, "excel_evaluate");
                Excel.Range writeRange = null;
                bool ownsWriteRange = false;
                try
                {
                    if (scalar)
                    {
                        writeRange = target.Range;
                    }
                    else
                    {
                        writeRange = target.Range.Resize[rows, columns];
                        ownsWriteRange = true;
                    }
                    string address = writeRange.AddressLocal();
                    ExcelUndoScope undo = ExcelUndoScope.Capture(
                        target.Workbook,
                        target.Worksheet,
                        writeRange,
                        address);

                    if (scalar)
                    {
                        writeRange.Value2 = matrix[0, 0];
                    }
                    else
                    {
                        writeRange.Value2 = matrix;
                    }

                    string targetRef = ExcelComSupport.FullAddress(target.SheetName, address);
                    Observation observation = Observation.Create(
                        "evaluate",
                        "公式结果已写入 " + targetRef + "。",
                        true);
                    observation.AddTargetRef(targetRef);
                    data["written"] = true;
                    data["writeTarget"] = targetRef;
                    data["writtenRows"] = rows;
                    data["writtenColumns"] = columns;

                    undo.Complete(
                        target.Workbook,
                        target.Worksheet,
                        writeRange,
                        address,
                        true,
                        "写入求值结果 " + targetRef,
                        observation,
                        data,
                        null);

                    return OperationResult.Succeed(
                        "公式求值完成，并写入 " + targetRef + "。",
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

        private static object EvaluateInTemporaryCell(Excel.Application application, string formula)
        {
            Excel.Workbook workbook = null;
            Excel.Worksheet tempSheet = null;
            Excel.Range cell = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                string sheetName = ExcelPivotService.UniqueSheetName(
                    workbook,
                    "_mcp_eval_" + DateTime.Now.ToString("HHmmssfff", CultureInfo.InvariantCulture));
                tempSheet = RangeResolver.CreateWorksheet(workbook, sheetName);
                cell = tempSheet.Range["A1"];
                string engine;
                ExcelComSupport.SetFormula(cell, formula, out engine);
                application.CalculateFull();
                return cell.Value2;
            }
            finally
            {
                if (tempSheet != null)
                {
                    ExcelPivotService.TryDeleteWorksheet(application, tempSheet);
                }
                ComHelper.Release(cell, tempSheet, workbook);
            }
        }

        private static JToken NormalizeEvaluateResult(object value)
        {
            if (value == null)
            {
                return JValue.CreateNull();
            }
            Excel.Range range = value as Excel.Range;
            if (range != null)
            {
                try
                {
                    return ExcelComSupport.ToJToken(range.Value2);
                }
                finally
                {
                    ComHelper.Release(range);
                }
            }
            return ExcelComSupport.ToJToken(value);
        }

        private static string DescribeResultType(object value)
        {
            if (value == null)
            {
                return "null";
            }
            if (value is Excel.Range)
            {
                return "range";
            }
            if (value is object[,])
            {
                return "array2d";
            }
            if (value is Array)
            {
                return "array";
            }
            if (value is double || value is int || value is long || value is decimal || value is float)
            {
                return "number";
            }
            if (value is bool)
            {
                return "boolean";
            }
            if (value is DateTime)
            {
                return "date";
            }
            return "string";
        }
    }
}
