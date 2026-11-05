using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>数据透视表创建服务，供 excel_create_pivot_table 与 analyze_data(pivot) 复用。</summary>
    internal static class ExcelPivotService
    {
        public static OperationResult Create(
            Excel.Application application,
            ExcelRangeContext source,
            string targetCell,
            string targetSheetName,
            JArray rowFields,
            JArray columnFields,
            JArray valueFields,
            string tableName)
        {
            if (source == null || source.Range == null)
            {
                return ExcelToolExecutor.Fail(OfficeErrorCode.INVALID_PARAMS, "缺少可用的数据源区域。");
            }
            if (valueFields == null || valueFields.Count == 0)
            {
                return ExcelToolExecutor.Fail(
                    OfficeErrorCode.INVALID_PARAMS,
                    "缺少 valueFields。请提供至少一个数值字段，例如 [\"销售额\"]。");
            }

            string destinationSheetName = targetSheetName;
            string destinationAddress = targetCell;
            string sheetPart;
            string localPart;
            RangeResolver.SplitAddress(targetCell ?? string.Empty, out sheetPart, out localPart);
            if (!string.IsNullOrWhiteSpace(sheetPart))
            {
                destinationSheetName = sheetPart;
            }
            if (!string.IsNullOrWhiteSpace(localPart))
            {
                destinationAddress = localPart;
            }
            if (string.IsNullOrWhiteSpace(destinationAddress))
            {
                destinationAddress = "A3";
            }

            bool createdSheet = false;
            Excel.Worksheet destinationSheet = null;
            Excel.Range destinationRange = null;
            Excel.PivotCaches caches = null;
            Excel.PivotCache cache = null;
            Excel.PivotTable pivotTable = null;
            try
            {
                if (string.IsNullOrWhiteSpace(destinationSheetName))
                {
                    destinationSheetName = UniqueSheetName(source.Workbook, "透视表_" + DateTime.Now.ToString("HHmmss", CultureInfo.InvariantCulture));
                    destinationSheet = RangeResolver.CreateWorksheet(source.Workbook, destinationSheetName);
                    createdSheet = true;
                }
                else
                {
                    destinationSheet = ExcelComSupport.FindWorksheet(source.Workbook, destinationSheetName);
                    if (destinationSheet == null)
                    {
                        destinationSheet = RangeResolver.CreateWorksheet(source.Workbook, destinationSheetName);
                        createdSheet = true;
                    }
                }

                ExcelComSupport.EnsureWritable(application, source.Workbook, destinationSheet, "excel_create_pivot_table");
                destinationRange = destinationSheet.Range[destinationAddress];
                string sourceAddress = ExcelComSupport.FullAddress(
                    source.SheetName,
                    source.Range.AddressLocal());

                application.Guard(delegate
                {
                    caches = source.Workbook.PivotCaches();
                    cache = caches.Create(
                        SourceType: Excel.XlPivotTableSourceType.xlDatabase,
                        SourceData: sourceAddress);
                    string finalTableName = string.IsNullOrWhiteSpace(tableName)
                        ? "PivotTable_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture)
                        : tableName.Trim();
                    pivotTable = cache.CreatePivotTable(
                        TableDestination: destinationRange,
                        TableName: finalTableName);

                    AddOrientationFields(pivotTable, rowFields, Excel.XlPivotFieldOrientation.xlRowField);
                    AddOrientationFields(pivotTable, columnFields, Excel.XlPivotFieldOrientation.xlColumnField);
                    AddValueFields(pivotTable, valueFields);

                    destinationSheet.Activate();
                    Excel.Range used = null;
                    Excel.Range columns = null;
                    try
                    {
                        used = destinationSheet.UsedRange;
                        columns = used.Columns;
                        columns.AutoFit();
                    }
                    finally
                    {
                        ComHelper.Release(columns, used);
                    }
                });

                string finalTable = pivotTable.Name;
                string destinationRef = ExcelComSupport.FullAddress(
                    destinationSheet.Name,
                    destinationAddress);
                Observation observation = Observation.Create(
                    "create_pivot_table",
                    "已在 " + destinationRef + " 创建透视表「" + finalTable + "」。",
                    true);
                observation.AddTargetRef(destinationRef);

                JObject data = new JObject();
                data["source"] = ExcelComSupport.FullAddress(
                    source.SheetName,
                    source.Range.AddressLocal());
                data["sheet"] = destinationSheet.Name;
                data["destination"] = destinationAddress;
                data["pivotTableName"] = finalTable;
                ExcelUndoScope.RecordUnavailable(
                    data,
                    "透视表与新增工作表不属于区域快照可逆范围，未登记自动撤销。");

                return OperationResult.Succeed(
                    "透视表已创建：" + destinationRef + "（" + finalTable + "）。",
                    data,
                    observation);
            }
            catch
            {
                if (createdSheet && destinationSheet != null)
                {
                    TryDeleteWorksheet(application, destinationSheet);
                }
                throw;
            }
            finally
            {
                ComHelper.Release(pivotTable, cache, caches, destinationRange, destinationSheet);
            }
        }

        private static void AddOrientationFields(
            Excel.PivotTable pivotTable,
            JArray fields,
            Excel.XlPivotFieldOrientation orientation)
        {
            if (fields == null)
            {
                return;
            }
            foreach (JToken token in fields)
            {
                string fieldName = token.ToString().Trim();
                if (fieldName.Length == 0)
                {
                    continue;
                }
                Excel.PivotField field = null;
                try
                {
                    field = (Excel.PivotField)pivotTable.PivotFields(fieldName);
                    field.Orientation = orientation;
                }
                finally
                {
                    ComHelper.Release(field);
                }
            }
        }

        private static void AddValueFields(Excel.PivotTable pivotTable, JArray fields)
        {
            foreach (JToken token in fields)
            {
                string fieldName;
                string caption;
                string aggregate;
                if (token is JObject)
                {
                    JObject config = (JObject)token;
                    fieldName = ExcelArgumentReader.GetTrimmedString(config, "field", "name");
                    caption = ExcelArgumentReader.GetTrimmedString(config, "caption");
                    aggregate = ExcelArgumentReader.GetTrimmedString(config, "function", "aggregate");
                }
                else
                {
                    fieldName = token.ToString().Trim();
                    caption = string.Empty;
                    aggregate = "sum";
                }

                if (fieldName.Length == 0)
                {
                    continue;
                }
                if (aggregate.Length == 0)
                {
                    aggregate = "sum";
                }

                Excel.PivotField sourceField = null;
                Excel.PivotField dataField = null;
                try
                {
                    sourceField = (Excel.PivotField)pivotTable.PivotFields(fieldName);
                    Excel.XlConsolidationFunction function = MapAggregate(aggregate);
                    if (string.IsNullOrWhiteSpace(caption))
                    {
                        dataField = pivotTable.AddDataField(sourceField, Type.Missing, function);
                    }
                    else
                    {
                        dataField = pivotTable.AddDataField(sourceField, caption, function);
                    }
                }
                finally
                {
                    ComHelper.Release(dataField, sourceField);
                }
            }
        }

        private static Excel.XlConsolidationFunction MapAggregate(string aggregate)
        {
            string value = (aggregate ?? string.Empty).Trim().ToLowerInvariant();
            switch (value)
            {
                case "count":
                    return Excel.XlConsolidationFunction.xlCount;
                case "avg":
                case "average":
                    return Excel.XlConsolidationFunction.xlAverage;
                case "max":
                    return Excel.XlConsolidationFunction.xlMax;
                case "min":
                    return Excel.XlConsolidationFunction.xlMin;
                case "product":
                    return Excel.XlConsolidationFunction.xlProduct;
                default:
                    return Excel.XlConsolidationFunction.xlSum;
            }
        }

        internal static string UniqueSheetName(Excel.Workbook workbook, string baseName)
        {
            string candidate = baseName;
            int suffix = 1;
            while (true)
            {
                Excel.Worksheet existing = ExcelComSupport.FindWorksheet(workbook, candidate);
                if (existing == null)
                {
                    return candidate;
                }
                ComHelper.Release(existing);
                candidate = baseName + "_" + suffix.ToString(CultureInfo.InvariantCulture);
                suffix++;
                if (candidate.Length > 31)
                {
                    candidate = baseName.Substring(0, Math.Max(1, 31 - suffix.ToString(CultureInfo.InvariantCulture).Length - 1))
                                + "_" + suffix.ToString(CultureInfo.InvariantCulture);
                }
            }
        }

        internal static void TryDeleteWorksheet(Excel.Application application, Excel.Worksheet worksheet)
        {
            if (application == null || worksheet == null)
            {
                return;
            }
            try
            {
                application.Guard(delegate { worksheet.Delete(); });
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[excel] 清理临时工作表失败：" + ex.Message);
            }
        }
    }
}
