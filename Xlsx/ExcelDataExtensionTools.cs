using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>数据动作扩展：多键排序与条件格式维护。</summary>
    public sealed partial class ExcelToolExecutor
    {
        private sealed class SortKeySpec
        {
            public int Column { get; set; }

            public string Order { get; set; }

            public Excel.XlSortOrder SortOrder { get; set; }
        }

        private static OperationResult ExecuteSortDataMulti(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange");
            if (string.IsNullOrWhiteSpace(address))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 range 参数。");
            }
            JArray sortKeysArray = ExcelArgumentReader.GetArray(arguments, "sortKeys", "keys");
            if (sortKeysArray == null || sortKeysArray.Count == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "sortKeys 数组不能为空。");
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

                List<SortKeySpec> keys = new List<SortKeySpec>();
                for (int index = 0; index < sortKeysArray.Count; index++)
                {
                    JObject key = sortKeysArray[index] as JObject;
                    if (key == null)
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "sortKeys[" + index.ToString(CultureInfo.InvariantCulture) + "] 必须是对象。");
                    }
                    JToken columnToken = ExcelArgumentReader.First(key, "column", "sortColumn", "field", "by");
                    if (columnToken == null)
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "sortKeys[" + index.ToString(CultureInfo.InvariantCulture) + "] 缺少 column。");
                    }
                    int column = ExcelComSupport.ResolveColumnIndex(
                        arguments,
                        columnToken,
                        headers,
                        hasHeader,
                        0);
                    if (column < 1 || column > columns)
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "sortKeys[" + index.ToString(CultureInfo.InvariantCulture) + "].column=" +
                            column + " 超出范围（范围共 " + columns + " 列）。");
                    }
                    string order = ExcelArgumentReader.GetTrimmedString(key, "order", "direction").ToLowerInvariant();
                    if (order.Length == 0)
                    {
                        order = "asc";
                    }
                    if (order != "asc" && order != "ascending" &&
                        order != "desc" && order != "descending")
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "sortKeys[" + index.ToString(CultureInfo.InvariantCulture) + "].order 仅支持 asc/desc。");
                    }
                    SortKeySpec spec = new SortKeySpec();
                    spec.Column = column;
                    spec.Order = order == "desc" || order == "descending" ? "desc" : "asc";
                    spec.SortOrder = spec.Order == "desc"
                        ? Excel.XlSortOrder.xlDescending
                        : Excel.XlSortOrder.xlAscending;
                    keys.Add(spec);
                }

                string rangeAddress = context.Range.AddressLocal();
                ExcelUndoScope undo = ExcelUndoScope.Capture(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress);
                List<Excel.Range> keyRanges = new List<Excel.Range>();
                try
                {
                    if (keys.Count <= 3)
                    {
                        int[] keyIndexes = new int[] { keys[0].Column, 0, 0 };
                        Excel.XlSortOrder[] keyOrders = new Excel.XlSortOrder[]
                        {
                            keys[0].SortOrder,
                            Excel.XlSortOrder.xlAscending,
                            Excel.XlSortOrder.xlAscending
                        };
                        if (keys.Count > 1)
                        {
                            keyIndexes[1] = keys[1].Column;
                            keyOrders[1] = keys[1].SortOrder;
                        }
                        if (keys.Count > 2)
                        {
                            keyIndexes[2] = keys[2].Column;
                            keyOrders[2] = keys[2].SortOrder;
                        }
                        for (int index = 0; index < keys.Count; index++)
                        {
                            keyRanges.Add((Excel.Range)context.Range.Columns[keyIndexes[index]]);
                        }

                        object key2 = keys.Count > 1 ? (object)keyRanges[1] : Type.Missing;
                        Excel.XlSortOrder order2 = keys.Count > 1
                            ? keyOrders[1]
                            : Excel.XlSortOrder.xlAscending;
                        object key3 = keys.Count > 2 ? (object)keyRanges[2] : Type.Missing;
                        Excel.XlSortOrder order3 = keys.Count > 2
                            ? keyOrders[2]
                            : Excel.XlSortOrder.xlAscending;
                        context.Range.Sort(
                            Key1: keyRanges[0],
                            Order1: keyOrders[0],
                            Key2: key2,
                            Order2: order2,
                            Key3: key3,
                            Order3: order3,
                            Header: hasHeader ? Excel.XlYesNoGuess.xlYes : Excel.XlYesNoGuess.xlNo,
                            Orientation: Excel.XlSortOrientation.xlSortColumns);
                    }
                    else
                    {
                        // Excel Range.Sort 最多接受 3 个键；从最次要键开始逐轮排序，
                        // Excel 对相等键保持既有相对顺序，从而得到多键稳定排序。
                        for (int index = keys.Count - 1; index >= 0; index--)
                        {
                            Excel.Range keyRange = null;
                            try
                            {
                                keyRange = (Excel.Range)context.Range.Columns[keys[index].Column];
                                context.Range.Sort(
                                    Key1: keyRange,
                                    Order1: keys[index].SortOrder,
                                    Header: hasHeader ? Excel.XlYesNoGuess.xlYes : Excel.XlYesNoGuess.xlNo,
                                    Orientation: Excel.XlSortOrientation.xlSortColumns);
                            }
                            finally
                            {
                                ComHelper.Release(keyRange);
                            }
                        }
                    }
                }
                finally
                {
                    foreach (Excel.Range keyRange in keyRanges)
                    {
                        ComHelper.Release(keyRange);
                    }
                    keyRanges.Clear();
                }

                string targetRef = ExcelComSupport.FullAddress(context.SheetName, rangeAddress);
                JArray returnedKeys = new JArray();
                foreach (SortKeySpec spec in keys)
                {
                    JObject item = new JObject();
                    item["column"] = spec.Column;
                    item["order"] = spec.Order;
                    returnedKeys.Add(item);
                }

                JObject data = new JObject();
                data["sheet"] = context.SheetName;
                data["address"] = rangeAddress;
                data["sortKeys"] = returnedKeys;
                data["sortColumn"] = keys[0].Column;
                data["order"] = keys[0].Order;
                data["hasHeader"] = hasHeader;

                Observation observation = Observation.Create(
                    "sort_data",
                    "已按 " + keys.Count + " 个键排序 " + targetRef + "。",
                    context.Range.Rows.Count > 1);
                observation.AddTargetRef(targetRef);
                undo.Complete(
                    context.Workbook,
                    context.Worksheet,
                    context.Range,
                    rangeAddress,
                    true,
                    "多键排序 " + targetRef,
                    observation,
                    data,
                    "排序会移动单元格格式；自动撤销仅恢复值/公式/数字格式，不恢复字体、填充等视觉格式。");

                return OperationResult.Succeed(
                    "多键排序完成：" + targetRef + "（" + keys.Count + " 个键）。",
                    data,
                    observation);
            }
        }

        private static OperationResult ExecuteConditionalFormatMaintenance(
            Excel.Application application,
            JObject arguments,
            string action)
        {
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange");
            if (string.IsNullOrWhiteSpace(address))
            {
                string sheetName = ExcelArgumentReader.GetSheetName(arguments);
                address = string.IsNullOrWhiteSpace(sheetName)
                    ? "usedrange"
                    : ExcelComSupport.QuoteSheetName(sheetName) + "!usedrange";
            }
            int maxResults = ExcelArgumentReader.GetInt(arguments, 200, "maxResults", "limit");
            maxResults = Math.Max(1, Math.Min(maxResults, 2000));

            using (ExcelRangeContext context = RangeResolver.Resolve(
                       application,
                       address,
                       false,
                       "excel_conditional_format"))
            {
                Excel.FormatConditions conditions = null;
                try
                {
                    conditions = context.Range.FormatConditions;
                    if (action == "list")
                    {
                        JArray rules = new JArray();
                        int count = conditions.Count;
                        for (int index = 1; index <= Math.Min(count, maxResults); index++)
                        {
                            object conditionObject = null;
                            try
                            {
                                conditionObject = conditions.Item(index);
                                rules.Add(BuildConditionInfo(conditionObject, index));
                            }
                            finally
                            {
                                ComHelper.Release(conditionObject);
                            }
                        }
                        JObject data = new JObject();
                        data["action"] = "list";
                        data["sheet"] = context.SheetName;
                        data["address"] = context.Range.AddressLocal();
                        data["count"] = count;
                        data["returned"] = rules.Count;
                        data["truncated"] = count > rules.Count;
                        data["rules"] = rules;
                        ExcelUndoScope.RecordUnavailable(data, "条件格式列举为只读操作，无需撤销。");
                        Observation observation = Observation.Create(
                            "conditional_format",
                            "已列出 " + rules.Count + " 条条件格式规则。",
                            false);
                        observation.AddTargetRef(
                            ExcelComSupport.FullAddress(context.SheetName, context.Range.AddressLocal()));
                        return OperationResult.Succeed(
                            "已列出 " + rules.Count + " 条条件格式规则。",
                            data,
                            observation);
                    }

                    ExcelComSupport.EnsureWritable(
                        application,
                        context.Workbook,
                        context.Worksheet,
                        "excel_conditional_format");
                    int ruleIndex = ExcelArgumentReader.GetInt(arguments, 0, "ruleIndex", "index", "ruleId");
                    string ruleType = ExcelArgumentReader.GetTrimmedString(arguments, "rule", "ruleType").ToLowerInvariant();
                    bool deleteAll = ExcelArgumentReader.GetBool(arguments, false, "deleteAll", "clearExisting");
                    if (!deleteAll && ruleIndex <= 0 && ruleType.Length == 0)
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "action=delete 需要 ruleIndex、rule 或 deleteAll=true 之一。");
                    }

                    int beforeCount = conditions.Count;
                    int deleted = 0;
                    application.Guard(delegate
                    {
                        if (deleteAll)
                        {
                            deleted = beforeCount;
                            conditions.Delete();
                            return;
                        }
                        if (ruleIndex > 0)
                        {
                            if (ruleIndex > beforeCount)
                            {
                                throw new OfficeComException(
                                    OfficeErrorCode.INVALID_PARAMS,
                                    "ruleIndex=" + ruleIndex + " 超出范围，当前区域共 " +
                                    beforeCount + " 条条件格式规则。");
                            }
                            object conditionObject = conditions.Item(ruleIndex);
                            try
                            {
                                if (DeleteConditionObject(conditionObject))
                                {
                                    deleted++;
                                }
                            }
                            finally
                            {
                                ComHelper.Release(conditionObject);
                            }
                            return;
                        }

                        int expectedType = MapConditionTypeForDelete(ruleType);
                        for (int index = beforeCount; index >= 1; index--)
                        {
                            object conditionObject = conditions.Item(index);
                            try
                            {
                                object rawType = ReadComProperty(conditionObject, "Type");
                                int actualType = rawType == null
                                    ? 0
                                    : Convert.ToInt32(rawType, CultureInfo.InvariantCulture);
                                if ((expectedType == 0 || actualType == expectedType) &&
                                    DeleteConditionObject(conditionObject))
                                {
                                    deleted++;
                                }
                            }
                            finally
                            {
                                ComHelper.Release(conditionObject);
                            }
                        }
                    });

                    JObject resultData = new JObject();
                    resultData["action"] = "delete";
                    resultData["sheet"] = context.SheetName;
                    resultData["address"] = context.Range.AddressLocal();
                    resultData["beforeCount"] = beforeCount;
                    resultData["deletedCount"] = deleted;
                    resultData["remainingCount"] = conditions.Count;
                    ExcelUndoScope.RecordUnavailable(
                        resultData,
                        "条件格式规则不在区域快照中，未登记自动撤销。");
                    Observation resultObservation = Observation.Create(
                        "conditional_format",
                        "已删除 " + deleted + " 条条件格式规则。",
                        deleted > 0);
                    resultObservation.AddTargetRef(
                        ExcelComSupport.FullAddress(context.SheetName, context.Range.AddressLocal()));
                    return OperationResult.Succeed(
                        "已删除 " + deleted + " 条条件格式规则，剩余 " + conditions.Count + " 条。",
                        resultData,
                        resultObservation);
                }
                finally
                {
                    ComHelper.Release(conditions);
                }
            }
        }

        private static JObject BuildConditionInfo(object conditionObject, int index)
        {
            JObject item = new JObject();
            item["ruleIndex"] = index;
            Excel.FormatCondition condition = conditionObject as Excel.FormatCondition;
            if (condition != null)
            {
                int type = Convert.ToInt32(condition.Type, CultureInfo.InvariantCulture);
                item["type"] = type;
                item["typeName"] = ConditionTypeName(type);
                try { item["operator"] = Convert.ToInt32(condition.Operator, CultureInfo.InvariantCulture); }
                catch (Exception) { item["operator"] = 0; }
                try { item["formula1"] = condition.Formula1; }
                catch (Exception) { item["formula1"] = string.Empty; }
                try { item["formula2"] = condition.Formula2; }
                catch (Exception) { item["formula2"] = string.Empty; }
                try { item["priority"] = condition.Priority; }
                catch (Exception) { item["priority"] = 0; }
                try { item["stopIfTrue"] = condition.StopIfTrue; }
                catch (Exception) { item["stopIfTrue"] = false; }
                Excel.Range appliesTo = null;
                try
                {
                    appliesTo = condition.AppliesTo;
                    item["appliesTo"] = appliesTo == null ? string.Empty : appliesTo.AddressLocal();
                }
                catch (Exception)
                {
                    item["appliesTo"] = string.Empty;
                }
                finally
                {
                    ComHelper.Release(appliesTo);
                }
                Excel.Interior interior = null;
                Excel.Font font = null;
                try
                {
                    interior = condition.Interior;
                    item["fillColor"] = ExcelComSupport.ToJToken(interior.Color);
                }
                catch (Exception)
                {
                    item["fillColor"] = JValue.CreateNull();
                }
                finally
                {
                    ComHelper.Release(interior);
                }
                try
                {
                    font = condition.Font;
                    item["fontColor"] = ExcelComSupport.ToJToken(font.Color);
                    item["fontBold"] = ExcelComSupport.ToJToken(font.Bold);
                }
                catch (Exception)
                {
                    item["fontColor"] = JValue.CreateNull();
                    item["fontBold"] = JValue.CreateNull();
                }
                finally
                {
                    ComHelper.Release(font);
                }
                return item;
            }

            Excel.Databar databar = conditionObject as Excel.Databar;
            if (databar != null)
            {
                item["type"] = (int)Excel.XlFormatConditionType.xlDatabar;
                item["typeName"] = "databar";
                return item;
            }
            Excel.ColorScale colorScale = conditionObject as Excel.ColorScale;
            if (colorScale != null)
            {
                item["type"] = (int)Excel.XlFormatConditionType.xlColorScale;
                item["typeName"] = "colorscale";
                return item;
            }
            Excel.IconSetCondition iconSet = conditionObject as Excel.IconSetCondition;
            if (iconSet != null)
            {
                item["type"] = (int)Excel.XlFormatConditionType.xlIconSets;
                item["typeName"] = "iconset";
                return item;
            }
            item["type"] = 0;
            item["typeName"] = "unknown";
            return item;
        }

        private static string ConditionTypeName(int type)
        {
            if (type == (int)Excel.XlFormatConditionType.xlCellValue)
            {
                return "highlight";
            }
            if (type == (int)Excel.XlFormatConditionType.xlExpression)
            {
                return "expression";
            }
            if (type == (int)Excel.XlFormatConditionType.xlDatabar)
            {
                return "databar";
            }
            if (type == (int)Excel.XlFormatConditionType.xlColorScale)
            {
                return "colorscale";
            }
            if (type == (int)Excel.XlFormatConditionType.xlIconSets)
            {
                return "iconset";
            }
            if (type == (int)Excel.XlFormatConditionType.xlTop10)
            {
                return "top10";
            }
            return "type" + type.ToString(CultureInfo.InvariantCulture);
        }

        private static bool DeleteConditionObject(object conditionObject)
        {
            if (conditionObject == null)
            {
                return false;
            }
            Excel.FormatCondition formatCondition = conditionObject as Excel.FormatCondition;
            if (formatCondition != null)
            {
                formatCondition.Delete();
                return true;
            }
            try
            {
                dynamic dynamicCondition = conditionObject;
                dynamicCondition.Delete();
                return true;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[excel_conditional_format] 删除条件格式失败：" + ex.Message);
                return false;
            }
        }

        private static int MapConditionTypeForDelete(string ruleType)
        {
            switch (ruleType)
            {
                case "highlight":
                case "cellvalue":
                    return (int)Excel.XlFormatConditionType.xlCellValue;
                case "expression":
                case "formula":
                    return (int)Excel.XlFormatConditionType.xlExpression;
                case "databar":
                    return (int)Excel.XlFormatConditionType.xlDatabar;
                case "colorscale":
                    return (int)Excel.XlFormatConditionType.xlColorScale;
                case "iconset":
                    return (int)Excel.XlFormatConditionType.xlIconSets;
                default:
                    return 0;
            }
        }
    }
}
