using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>排序、筛选、去重、条件格式、合并、自动调整和查找替换。</summary>
    public sealed partial class ExcelToolExecutor
    {
        private static OperationResult ExecuteSortData(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange");
            if (string.IsNullOrWhiteSpace(address))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 range 参数。");
            }
            if (ExcelArgumentReader.GetArray(arguments, "sortKeys", "keys") != null)
            {
                return ExecuteSortDataMulti(application, arguments);
            }

            using (ExcelRangeContext context = RangeResolver.Resolve(application, address, false, "excel_sort_data"))
            {
                ExcelComSupport.EnsureWritable(application, context.Workbook, context.Worksheet, "excel_sort_data");
                bool hasHeader = ExcelArgumentReader.GetBool(arguments, true, "hasHeader");
                int columns = context.Range.Columns.Count;
                string[] headers = RangeResolver.GetHeaders(
                    context.Worksheet,
                    context.Range,
                    Math.Min(columns, 50));
                int sortColumn = ExcelComSupport.ResolveColumnIndex(
                    arguments,
                    ExcelArgumentReader.First(arguments, "sortColumn", "column", "sortBy"),
                    headers,
                    hasHeader,
                    1);
                if (sortColumn < 1 || sortColumn > columns)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "sortColumn 超出范围：范围共 " + columns + " 列，收到 " + sortColumn + "。");
                }

                string order = ExcelArgumentReader.GetString(arguments, "order", "sortOrder").Trim().ToLowerInvariant();
                Excel.XlSortOrder sortOrder = order == "desc" || order == "descending"
                    ? Excel.XlSortOrder.xlDescending
                    : Excel.XlSortOrder.xlAscending;
                string rangeAddress = context.Range.AddressLocal();
                ExcelUndoScope undo = ExcelUndoScope.Capture(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress);

                Excel.Range keyRange = null;
                try
                {
                    keyRange = (Excel.Range)context.Range.Columns[sortColumn];
                    context.Range.Sort(
                        Key1: keyRange,
                        Order1: sortOrder,
                        Header: hasHeader ? Excel.XlYesNoGuess.xlYes : Excel.XlYesNoGuess.xlNo,
                        Orientation: Excel.XlSortOrientation.xlSortColumns);
                }
                finally
                {
                    ComHelper.Release(keyRange);
                }

                string targetRef = ExcelComSupport.FullAddress(context.SheetName, rangeAddress);
                Observation observation = Observation.Create(
                    "sort_data",
                    "已按第 " + sortColumn + " 列" + (sortOrder == Excel.XlSortOrder.xlDescending ? "降序" : "升序") +
                    "排序 " + targetRef + "。",
                    true);
                observation.AddTargetRef(targetRef);

                JObject data = new JObject();
                data["sheet"] = context.SheetName;
                data["address"] = rangeAddress;
                data["sortColumn"] = sortColumn;
                data["order"] = sortOrder == Excel.XlSortOrder.xlDescending ? "desc" : "asc";
                data["hasHeader"] = hasHeader;

                undo.Complete(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress,
                    true,
                    "排序 " + targetRef,
                    observation,
                    data,
                    "排序会移动单元格格式；自动撤销仅恢复值/公式/数字格式，不恢复字体、填充等视觉格式。");

                return OperationResult.Succeed("排序完成：" + targetRef + "。", data, observation);
            }
        }

        private static OperationResult ExecuteFilterData(Excel.Application application, JObject arguments)
        {
            bool clearFilter = ExcelArgumentReader.GetBool(arguments, false, "clearFilter");
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange");
            if (string.IsNullOrWhiteSpace(address))
            {
                address = "usedrange";
            }

            using (ExcelRangeContext context = RangeResolver.Resolve(application, address, false, "excel_filter_data"))
            {
                ExcelComSupport.EnsureWritable(application, context.Workbook, context.Worksheet, "excel_filter_data");
                string rangeAddress = context.Range.AddressLocal();
                string targetRef = ExcelComSupport.FullAddress(context.SheetName, rangeAddress);

                if (clearFilter)
                {
                    bool hadFilter = context.Worksheet.AutoFilterMode;
                    if (!hadFilter)
                    {
                        Observation unchanged = Observation.Create(
                            "filter_data",
                            "工作表当前没有自动筛选，无需清除。",
                            false);
                        unchanged.AddTargetRef(targetRef);
                        JObject emptyData = new JObject();
                        emptyData["sheet"] = context.SheetName;
                        emptyData["address"] = rangeAddress;
                        emptyData["cleared"] = false;
                        ExcelUndoScope.RecordUnavailable(emptyData, "筛选状态不在区域快照范围内。");
                        return OperationResult.Succeed("当前没有自动筛选。", emptyData, unchanged);
                    }

                    application.Guard(delegate
                    {
                        if (context.Worksheet.FilterMode)
                        {
                            context.Worksheet.ShowAllData();
                        }
                        context.Worksheet.AutoFilterMode = false;
                    });

                    Observation clearedObservation = Observation.Create(
                        "filter_data",
                        "已清除 " + context.SheetName + " 上的自动筛选。",
                        true);
                    clearedObservation.AddTargetRef(targetRef);
                    JObject clearedData = new JObject();
                    clearedData["sheet"] = context.SheetName;
                    clearedData["address"] = rangeAddress;
                    clearedData["cleared"] = true;
                    ExcelUndoScope.RecordUnavailable(clearedData, "筛选状态不在区域快照范围内，未登记自动撤销。");
                    return OperationResult.Succeed("筛选已清除。", clearedData, clearedObservation);
                }

                int column = ExcelArgumentReader.GetInt(arguments, 0, "column", "field");
                if (column <= 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 column 参数（从 1 开始）。");
                }
                if (column > context.Range.Columns.Count)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "column 超出范围：范围共 " + context.Range.Columns.Count + " 列。");
                }
                JToken criteriaToken = ExcelArgumentReader.First(arguments, "criteria", "value");
                if (criteriaToken == null)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 criteria 参数。");
                }
                string criteria = criteriaToken.ToString();

                ExcelUndoScope undo = ExcelUndoScope.Capture(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress);
                context.Range.AutoFilter(Field: column, Criteria1: criteria);

                Observation observation = Observation.Create(
                    "filter_data",
                    "已对 " + targetRef + " 的第 " + column + " 列应用筛选：" + criteria,
                    true);
                observation.AddTargetRef(targetRef);
                JObject data = new JObject();
                data["sheet"] = context.SheetName;
                data["address"] = rangeAddress;
                data["column"] = column;
                data["criteria"] = criteria;
                data["cleared"] = false;

                undo.Complete(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress,
                    true,
                    "筛选 " + targetRef,
                    observation,
                    data,
                    "筛选只改变行的可见状态，区域快照不包含 AutoFilter 状态，未登记自动撤销。");

                return OperationResult.Succeed(
                    "筛选已应用：" + targetRef + " 第 " + column + " 列 " + criteria + "。",
                    data,
                    observation);
            }
        }

        private static OperationResult ExecuteRemoveDuplicates(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange");
            if (string.IsNullOrWhiteSpace(address))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 range 参数。");
            }

            using (ExcelRangeContext context = RangeResolver.Resolve(application, address, false, "excel_remove_duplicates"))
            {
                ExcelComSupport.EnsureWritable(application, context.Workbook, context.Worksheet, "excel_remove_duplicates");
                bool hasHeader = ExcelArgumentReader.GetBool(arguments, true, "hasHeader");
                int columns = context.Range.Columns.Count;
                int[] requested = ExcelArgumentReader.GetIntList(arguments, "columns");
                int[] selected = NormalizeColumns(requested, columns);
                if (selected.Length == 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "columns 参数为空或超出范围。");
                }

                string rangeAddress = context.Range.AddressLocal();
                int removed = CountDuplicateRows(
                    context.Range,
                    selected,
                    hasHeader);
                ExcelUndoScope undo = ExcelUndoScope.Capture(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress);

                object[] columnsObject = new object[selected.Length];
                for (int index = 0; index < selected.Length; index++)
                {
                    columnsObject[index] = selected[index];
                }
                context.Range.RemoveDuplicates(
                    Columns: columnsObject,
                    Header: hasHeader ? Excel.XlYesNoGuess.xlYes : Excel.XlYesNoGuess.xlNo);

                string targetRef = ExcelComSupport.FullAddress(context.SheetName, rangeAddress);
                Observation observation = Observation.Create(
                    "remove_duplicates",
                    "已删除 " + removed.ToString(CultureInfo.InvariantCulture) + " 行重复数据：" + targetRef,
                    removed > 0);
                observation.AddTargetRef(targetRef);

                JObject data = new JObject();
                data["sheet"] = context.SheetName;
                data["address"] = rangeAddress;
                data["removedRows"] = removed;
                data["columns"] = new JArray(selected);
                data["hasHeader"] = hasHeader;

                undo.Complete(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress,
                    false,
                    "删除重复行 " + targetRef,
                    observation,
                    data,
                    "删除重复行会改变工作表结构，区域快照无法完整逆操作，未登记自动撤销。");

                return OperationResult.Succeed(
                    "已删除 " + removed + " 行重复数据：" + targetRef + "。",
                    data,
                    observation);
            }
        }

        private static OperationResult ExecuteConditionalFormat(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange");
            string action = ExcelArgumentReader.GetTrimmedString(arguments, "action").ToLowerInvariant();
            if (action.Length == 0)
            {
                action = "add";
            }
            if (action != "add" && action != "list" && action != "delete")
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "action 仅支持 add/list/delete。");
            }
            if (action == "list" || action == "delete")
            {
                return ExecuteConditionalFormatMaintenance(application, arguments, action);
            }
            if (string.IsNullOrWhiteSpace(address))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 range 参数。");
            }
            string rule = ExcelArgumentReader.GetString(arguments, "rule").Trim().ToLowerInvariant();
            if (rule.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 rule 参数（highlight/databar/colorscale/iconset）。");
            }

            using (ExcelRangeContext context = RangeResolver.Resolve(application, address, false, "excel_conditional_format"))
            {
                ExcelComSupport.EnsureWritable(application, context.Workbook, context.Worksheet, "excel_conditional_format");
                string rangeAddress = context.Range.AddressLocal();

                Excel.FormatConditions conditions = null;
                object added = null;
                try
                {
                    conditions = context.Range.FormatConditions;
                    if (ExcelArgumentReader.GetBool(arguments, false, "clearExisting"))
                    {
                        conditions.Delete();
                    }

                    if (rule == "highlight")
                    {
                        added = AddHighlightCondition(conditions, arguments);
                    }
                    else if (rule == "databar")
                    {
                        added = conditions.AddDatabar();
                    }
                    else if (rule == "colorscale")
                    {
                        added = conditions.AddColorScale(3);
                    }
                    else if (rule == "iconset")
                    {
                        added = conditions.AddIconSetCondition();
                    }
                    else
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "不支持的 rule：" + rule + "。可用 highlight/databar/colorscale/iconset。");
                    }
                }
                finally
                {
                    ComHelper.Release(added);
                    ComHelper.Release(conditions);
                }

                string targetRef = ExcelComSupport.FullAddress(context.SheetName, rangeAddress);
                Observation observation = Observation.Create(
                    "conditional_format",
                    "已对 " + targetRef + " 应用条件格式：" + rule,
                    true);
                observation.AddTargetRef(targetRef);

                JObject data = new JObject();
                data["sheet"] = context.SheetName;
                data["address"] = rangeAddress;
                data["rule"] = rule;
                ExcelUndoScope.RecordUnavailable(data, "条件格式规则不在区域快照范围内，未登记自动撤销。");
                return OperationResult.Succeed(
                    "已应用条件格式：" + targetRef + "（" + rule + "）。",
                    data,
                    observation);
            }
        }

        private static OperationResult ExecuteMergeCells(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange");
            if (string.IsNullOrWhiteSpace(address))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 range 参数。");
            }
            bool unmerge = ExcelArgumentReader.GetBool(arguments, false, "unmerge");

            using (ExcelRangeContext context = RangeResolver.Resolve(application, address, false, "excel_merge_cells"))
            {
                ExcelComSupport.EnsureWritable(application, context.Workbook, context.Worksheet, "excel_merge_cells");
                string rangeAddress = context.Range.AddressLocal();
                ExcelUndoScope undo = ExcelUndoScope.Capture(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress);

                if (unmerge)
                {
                    context.Range.UnMerge();
                }
                else
                {
                    context.Range.Merge();
                }

                string targetRef = ExcelComSupport.FullAddress(context.SheetName, rangeAddress);
                Observation observation = Observation.Create(
                    "merge_cells",
                    (unmerge ? "已取消合并 " : "已合并 ") + targetRef,
                    true);
                observation.AddTargetRef(targetRef);

                JObject data = new JObject();
                data["sheet"] = context.SheetName;
                data["address"] = rangeAddress;
                data["unmerge"] = unmerge;

                undo.Complete(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress,
                    false,
                    (unmerge ? "取消合并 " : "合并 ") + targetRef,
                    observation,
                    data,
                    "合并/取消合并会改变工作表结构，区域快照无法可靠回放，未登记自动撤销。");

                return OperationResult.Succeed(
                    (unmerge ? "已取消合并 " : "已合并 ") + targetRef + "。",
                    data,
                    observation);
            }
        }

        private static OperationResult ExecuteAutoFit(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange");
            if (string.IsNullOrWhiteSpace(address))
            {
                address = "usedrange";
            }
            string type = ExcelArgumentReader.GetString(arguments, "type", "fitType").Trim().ToLowerInvariant();
            if (type.Length == 0)
            {
                type = "columns";
            }
            if (type != "columns" && type != "column" && type != "both" &&
                type != "rows" && type != "row")
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "type 仅支持 columns/rows/both。");
            }

            using (ExcelRangeContext context = RangeResolver.Resolve(application, address, false, "excel_autofit"))
            {
                ExcelComSupport.EnsureWritable(application, context.Workbook, context.Worksheet, "excel_autofit");
                string rangeAddress = context.Range.AddressLocal();
                ExcelUndoScope undo = ExcelUndoScope.Capture(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress);

                application.Guard(delegate
                {
                    if (type == "columns" || type == "column" || type == "both")
                    {
                        Excel.Range columns = null;
                        try
                        {
                            columns = context.Range.Columns;
                            columns.AutoFit();
                        }
                        finally
                        {
                            ComHelper.Release(columns);
                        }
                    }
                    if (type == "rows" || type == "row" || type == "both")
                    {
                        Excel.Range rows = null;
                        try
                        {
                            rows = context.Range.Rows;
                            rows.AutoFit();
                        }
                        finally
                        {
                            ComHelper.Release(rows);
                        }
                    }
                });

                string targetRef = ExcelComSupport.FullAddress(context.SheetName, rangeAddress);
                Observation observation = Observation.Create(
                    "autofit",
                    "已自动调整 " + targetRef + " 的 " + type + "。",
                    true);
                observation.AddTargetRef(targetRef);

                JObject data = new JObject();
                data["sheet"] = context.SheetName;
                data["address"] = rangeAddress;
                data["type"] = type;

                undo.Complete(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress,
                    true,
                    "自动调整 " + targetRef,
                    observation,
                    data,
                    "行高/列宽不在区域快照中，未登记自动撤销。");

                return OperationResult.Succeed("自动调整完成：" + targetRef + "。", data, observation);
            }
        }

        private static OperationResult ExecuteFindReplace(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange");
            if (string.IsNullOrWhiteSpace(address) || string.Equals(address, "all", StringComparison.OrdinalIgnoreCase))
            {
                address = "usedrange";
            }
            string findText = ExcelArgumentReader.GetString(arguments, "find", "findText");
            if (findText.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 find/findText 参数。");
            }
            JToken replacementToken = ExcelArgumentReader.First(arguments, "replace", "replaceText");
            string replaceText = replacementToken == null ? string.Empty : replacementToken.ToString();
            bool matchCase = ExcelArgumentReader.GetBool(arguments, false, "matchCase");
            bool matchEntireCell = ExcelArgumentReader.GetBool(arguments, false, "matchEntireCell");

            using (ExcelRangeContext context = RangeResolver.Resolve(application, address, false, "excel_find_replace"))
            {
                ExcelComSupport.EnsureWritable(application, context.Workbook, context.Worksheet, "excel_find_replace");
                string rangeAddress = context.Range.AddressLocal();
                int rows = context.Range.Rows.Count;
                int columns = context.Range.Columns.Count;
                object[,] values = ExcelComSupport.ReadMatrix(context.Range, rows, columns);
                int matches = ExcelComSupport.CountMatches(values, findText, matchCase, matchEntireCell);

                string targetRef = ExcelComSupport.FullAddress(context.SheetName, rangeAddress);
                if (matches == 0)
                {
                    Observation emptyObservation = Observation.Create(
                        "find_replace",
                        "在 " + targetRef + " 中未找到「" + findText + "」，未做修改。",
                        false);
                    emptyObservation.AddTargetRef(targetRef);
                    JObject emptyData = new JObject();
                    emptyData["sheet"] = context.SheetName;
                    emptyData["address"] = rangeAddress;
                    emptyData["replacements"] = 0;
                    emptyData["find"] = findText;
                    ExcelUndoScope.RecordUnavailable(emptyData, "没有发生修改，无需撤销。");
                    return OperationResult.Succeed("未找到匹配内容，未做修改。", emptyData, emptyObservation);
                }

                ExcelUndoScope undo = ExcelUndoScope.Capture(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress);
                context.Range.Replace(
                    What: findText,
                    Replacement: replaceText,
                    LookAt: matchEntireCell ? Excel.XlLookAt.xlWhole : Excel.XlLookAt.xlPart,
                    MatchCase: matchCase);

                Observation observation = Observation.Create(
                    "find_replace",
                    "已在 " + targetRef + " 中替换 " + matches + " 处「" + findText + "」。",
                    true);
                observation.AddTargetRef(targetRef);

                JObject data = new JObject();
                data["sheet"] = context.SheetName;
                data["address"] = rangeAddress;
                data["replacements"] = matches;
                data["find"] = findText;
                data["replace"] = replaceText;
                data["matchCase"] = matchCase;
                data["matchEntireCell"] = matchEntireCell;

                undo.Complete(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress,
                    true,
                    "查找替换 " + targetRef,
                    observation,
                    data,
                    null);

                return OperationResult.Succeed(
                    "已替换 " + matches + " 处，目标 " + targetRef + "。",
                    data,
                    observation);
            }
        }

        private static int[] NormalizeColumns(int[] requested, int totalColumns)
        {
            if (requested == null || requested.Length == 0)
            {
                int[] all = new int[totalColumns];
                for (int index = 0; index < totalColumns; index++)
                {
                    all[index] = index + 1;
                }
                return all;
            }

            List<int> valid = new List<int>();
            foreach (int column in requested)
            {
                if (column >= 1 && column <= totalColumns && !valid.Contains(column))
                {
                    valid.Add(column);
                }
            }
            return valid.ToArray();
        }

        private static int CountDuplicateRows(Excel.Range range, int[] columns, bool hasHeader)
        {
            int rows = range.Rows.Count;
            int totalColumns = range.Columns.Count;
            object[,] values = ExcelComSupport.ReadMatrix(range, rows, totalColumns);
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int duplicates = 0;
            int startRow = hasHeader ? 1 : 0;
            for (int row = startRow; row < rows; row++)
            {
                List<string> parts = new List<string>();
                foreach (int column in columns)
                {
                    object value = values[row, column - 1];
                    if (value == null || value is DBNull)
                    {
                        parts.Add("<NULL>");
                    }
                    else
                    {
                        parts.Add(value.GetType().FullName + ":" + Convert.ToString(value, CultureInfo.InvariantCulture));
                    }
                }
                string key = string.Join("\u001f", parts.ToArray());
                if (!seen.Add(key))
                {
                    duplicates++;
                }
            }
            return duplicates;
        }

        private static object AddHighlightCondition(Excel.FormatConditions conditions, JObject arguments)
        {
            string operatorText = ExcelArgumentReader.GetString(arguments, "operator", "conditionOperator")
                .Trim()
                .ToLowerInvariant();
            Excel.XlFormatConditionOperator conditionOperator = MapConditionOperator(operatorText);
            string formula1 = ExcelArgumentReader.GetString(arguments, "condition", "value");
            if (string.IsNullOrWhiteSpace(formula1))
            {
                formula1 = "0";
            }
            string formula2 = ExcelArgumentReader.GetString(arguments, "condition2", "value2");
            string color = ExcelArgumentReader.GetString(arguments, "color");
            if (string.IsNullOrWhiteSpace(color))
            {
                color = "#FFC7CE";
            }

            Excel.FormatCondition condition;
            if (conditionOperator == Excel.XlFormatConditionOperator.xlBetween ||
                conditionOperator == Excel.XlFormatConditionOperator.xlNotBetween)
            {
                condition = (Excel.FormatCondition)conditions.Add(
                    Type: Excel.XlFormatConditionType.xlCellValue,
                    Operator: conditionOperator,
                    Formula1: formula1,
                    Formula2: string.IsNullOrWhiteSpace(formula2) ? formula1 : formula2);
            }
            else
            {
                condition = (Excel.FormatCondition)conditions.Add(
                    Type: Excel.XlFormatConditionType.xlCellValue,
                    Operator: conditionOperator,
                    Formula1: formula1);
            }
            try
            {
                Excel.Interior interior = null;
                try
                {
                    interior = condition.Interior;
                    interior.Color = ExcelComSupport.ParseColor(color, ExcelComSupport.ToOleColor(255, 199, 206));
                }
                finally
                {
                    ComHelper.Release(interior);
                }
                return condition;
            }
            catch
            {
                ComHelper.Release(condition);
                throw;
            }
        }

        private static Excel.XlFormatConditionOperator MapConditionOperator(string operatorText)
        {
            switch (operatorText)
            {
                case "equal":
                case "=":
                case "eq":
                    return Excel.XlFormatConditionOperator.xlEqual;
                case "notequal":
                case "<>":
                case "ne":
                    return Excel.XlFormatConditionOperator.xlNotEqual;
                case "less":
                case "<":
                case "lt":
                    return Excel.XlFormatConditionOperator.xlLess;
                case "lessequal":
                case "<=":
                case "le":
                    return Excel.XlFormatConditionOperator.xlLessEqual;
                case "greaterequal":
                case ">=":
                case "ge":
                    return Excel.XlFormatConditionOperator.xlGreaterEqual;
                case "between":
                    return Excel.XlFormatConditionOperator.xlBetween;
                case "notbetween":
                    return Excel.XlFormatConditionOperator.xlNotBetween;
                default:
                    return Excel.XlFormatConditionOperator.xlGreater;
            }
        }
    }
}
