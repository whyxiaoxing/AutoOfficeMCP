using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>P3 图表维护：列出、更新数据源/标题/类型/位置、删除。</summary>
    public sealed partial class ExcelToolExecutor
    {
        private sealed class ChartTarget
        {
            public Excel.ChartObject ChartObject { get; set; }

            public Excel.Chart Chart { get; set; }

            public Excel.Worksheet Worksheet { get; set; }

            public string SheetName { get; set; }

            public string ChartName { get; set; }

            public bool IsChartSheet { get; set; }
        }

        private static OperationResult ExecuteListCharts(Excel.Application application, JObject arguments)
        {
            string sheetName = ExcelArgumentReader.GetSheetName(arguments);
            bool includeChartSheets = ExcelArgumentReader.GetBool(arguments, true, "includeChartSheets");
            int maxResults = ExcelArgumentReader.GetInt(arguments, 200, "maxResults", "limit");
            maxResults = Math.Max(1, Math.Min(maxResults, 2000));

            Excel.Workbook workbook = null;
            Excel.Worksheet onlySheet = null;
            JArray charts = new JArray();
            bool truncated = false;
            bool chartSheetRequested = false;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                if (sheetName.Length > 0)
                {
                    onlySheet = ExcelComSupport.FindWorksheet(workbook, sheetName);
                    if (onlySheet == null)
                    {
                        chartSheetRequested = IsChartSheetName(workbook, sheetName);
                        if (!chartSheetRequested)
                        {
                            return Fail(OfficeErrorCode.SHEET_NOT_FOUND, "找不到工作表或图表页「" + sheetName + "」。");
                        }
                    }
                    else
                    {
                        CollectEmbeddedCharts(onlySheet, charts, maxResults, ref truncated);
                    }
                }
                else
                {
                    Excel.Sheets worksheets = null;
                    try
                    {
                        worksheets = workbook.Worksheets;
                        for (int index = 1; index <= worksheets.Count; index++)
                        {
                            object item = worksheets[index];
                            Excel.Worksheet worksheet = item as Excel.Worksheet;
                            if (worksheet == null)
                            {
                                ComHelper.Release(item);
                                continue;
                            }
                            try
                            {
                                CollectEmbeddedCharts(worksheet, charts, maxResults, ref truncated);
                            }
                            finally
                            {
                                ComHelper.Release(worksheet);
                            }
                        }
                    }
                    finally
                    {
                        ComHelper.Release(worksheets);
                    }
                }

                if (includeChartSheets && (sheetName.Length == 0 || chartSheetRequested))
                {
                    Excel.Sheets chartSheets = null;
                    try
                    {
                        chartSheets = workbook.Charts;
                        for (int index = 1; index <= chartSheets.Count; index++)
                        {
                            object item = chartSheets[index];
                            Excel.Chart chart = item as Excel.Chart;
                            if (chart == null)
                            {
                                ComHelper.Release(item);
                                continue;
                            }
                            try
                            {
                                string name = chart.Name;
                                if (sheetName.Length > 0 &&
                                    !string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase))
                                {
                                    continue;
                                }
                                if (charts.Count >= maxResults)
                                {
                                    truncated = true;
                                    continue;
                                }
                                charts.Add(BuildChartInfo(chart, null, name, name, true, index));
                            }
                            finally
                            {
                                ComHelper.Release(chart);
                            }
                        }
                    }
                    finally
                    {
                        ComHelper.Release(chartSheets);
                    }
                }

                JObject data = new JObject();
                data["workbook"] = workbook.Name;
                data["count"] = charts.Count;
                data["charts"] = charts;
                data["truncated"] = truncated;
                ExcelUndoScope.RecordUnavailable(data, "图表列举为只读操作，无需撤销。");
                Observation observation = Observation.Create(
                    "list_charts",
                    "已列出 " + charts.Count + " 个图表。",
                    false);
                observation.AddTargetRef(workbook.FullName);
                return OperationResult.Succeed(
                    "已列出 " + charts.Count + " 个图表" + (truncated ? "（已截断）" : string.Empty) + "。",
                    data,
                    observation);
            }
            finally
            {
                ComHelper.Release(onlySheet, workbook);
            }
        }

        private static bool IsChartSheetName(Excel.Workbook workbook, string sheetName)
        {
            if (workbook == null || string.IsNullOrWhiteSpace(sheetName))
            {
                return false;
            }
            Excel.Sheets chartSheets = null;
            try
            {
                chartSheets = workbook.Charts;
                for (int index = 1; index <= chartSheets.Count; index++)
                {
                    object item = chartSheets[index];
                    Excel.Chart chart = item as Excel.Chart;
                    if (chart == null)
                    {
                        ComHelper.Release(item);
                        continue;
                    }
                    try
                    {
                        if (string.Equals(chart.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
                    }
                    finally
                    {
                        ComHelper.Release(chart);
                    }
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                ComHelper.Release(chartSheets);
            }
        }

        private static void CollectEmbeddedCharts(
            Excel.Worksheet worksheet,
            JArray charts,
            int maxResults,
            ref bool truncated)
        {
            Excel.ChartObjects chartObjects = null;
            try
            {
                chartObjects = (Excel.ChartObjects)worksheet.ChartObjects();
                for (int index = 1; index <= chartObjects.Count; index++)
                {
                    Excel.ChartObject chartObject = null;
                    Excel.Chart chart = null;
                    try
                    {
                        chartObject = (Excel.ChartObject)chartObjects.Item(index);
                        chart = chartObject.Chart;
                        if (charts.Count >= maxResults)
                        {
                            truncated = true;
                            continue;
                        }
                        charts.Add(BuildChartInfo(
                            chart,
                            chartObject,
                            worksheet.Name + "!" + chartObject.Name,
                            worksheet.Name,
                            false,
                            index));
                    }
                    finally
                    {
                        ComHelper.Release(chart, chartObject);
                    }
                }
            }
            finally
            {
                ComHelper.Release(chartObjects);
            }
        }

        private static JObject BuildChartInfo(
            Excel.Chart chart,
            Excel.ChartObject chartObject,
            string chartName,
            string sheetName,
            bool isChartSheet,
            int index)
        {
            JObject item = new JObject();
            item["chartName"] = chartName;
            item["index"] = index;
            item["isChartSheet"] = isChartSheet;
            int chartType = 0;
            try { chartType = Convert.ToInt32(chart.ChartType, CultureInfo.InvariantCulture); }
            catch (Exception) { }
            item["chartType"] = chartType;
            item["chartTypeName"] = FriendlyChartType(chartType);
            string title = string.Empty;
            bool hasTitle = false;
            try
            {
                hasTitle = chart.HasTitle;
                if (hasTitle)
                {
                    Excel.ChartTitle chartTitle = null;
                    try
                    {
                        chartTitle = chart.ChartTitle;
                        title = chartTitle.Text;
                    }
                    finally
                    {
                        ComHelper.Release(chartTitle);
                    }
                }
            }
            catch (Exception) { }
            item["hasTitle"] = hasTitle;
            item["title"] = title;
            item["sourceData"] = ReadFirstSeriesFormula(chart);
            if (chartObject != null)
            {
                item["sheet"] = sheetName;
                item["left"] = chartObject.Left;
                item["top"] = chartObject.Top;
                item["width"] = chartObject.Width;
                item["height"] = chartObject.Height;
                item["visible"] = chartObject.Visible;
            }
            else
            {
                item["sheet"] = chartName;
                item["left"] = JValue.CreateNull();
                item["top"] = JValue.CreateNull();
                item["width"] = JValue.CreateNull();
                item["height"] = JValue.CreateNull();
                item["visible"] = true;
            }
            return item;
        }

        private static string ReadFirstSeriesFormula(Excel.Chart chart)
        {
            Excel.SeriesCollection seriesCollection = null;
            Excel.Series series = null;
            try
            {
                seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection();
                if (seriesCollection.Count == 0)
                {
                    return string.Empty;
                }
                series = seriesCollection.Item(1);
                return series.Formula ?? string.Empty;
            }
            catch (Exception)
            {
                return string.Empty;
            }
            finally
            {
                ComHelper.Release(series, seriesCollection);
            }
        }

        private static string FriendlyChartType(int chartType)
        {
            if (chartType == (int)Excel.XlChartType.xlColumnClustered)
            {
                return "column";
            }
            if (chartType == (int)Excel.XlChartType.xlLine)
            {
                return "line";
            }
            if (chartType == (int)Excel.XlChartType.xlPie)
            {
                return "pie";
            }
            if (chartType == (int)Excel.XlChartType.xlBarClustered)
            {
                return "bar";
            }
            if (chartType == (int)Excel.XlChartType.xlXYScatter)
            {
                return "scatter";
            }
            if (chartType == (int)Excel.XlChartType.xlArea)
            {
                return "area";
            }
            return "type" + chartType.ToString(CultureInfo.InvariantCulture);
        }

        private static OperationResult ExecuteUpdateChart(Excel.Application application, JObject arguments)
        {
            string chartName = ExcelArgumentReader.GetTrimmedString(arguments, "chartName", "name", "chart");
            if (chartName.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 chartName 参数。");
            }
            string sheetName = ExcelArgumentReader.GetSheetName(arguments);
            string dataRange = ExcelArgumentReader.GetTrimmedString(
                arguments,
                "dataRange",
                "sourceRange",
                "range");
            bool hasTitleArgument = ExcelArgumentReader.Has(arguments, "title");
            string title = ExcelArgumentReader.GetString(arguments, "title");
            bool removeTitle = ExcelArgumentReader.GetBool(arguments, false, "removeTitle");
            string chartTypeText = ExcelArgumentReader.GetTrimmedString(arguments, "type", "chartType");

            Excel.Workbook workbook = null;
            ChartTarget target = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                string findError;
                target = FindChart(workbook, chartName, sheetName, out findError);
                if (target == null)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        findError.Length == 0 ? "找不到图表「" + chartName + "」。" : findError);
                }
                ExcelComSupport.EnsureWritable(
                    application,
                    workbook,
                    target.Worksheet,
                    "excel_update_chart");

                int applied = 0;
                List<string> changedFields = new List<string>();
                try
                {
                    if (dataRange.Length > 0)
                    {
                        using (ExcelRangeContext source = RangeResolver.Resolve(
                                   application,
                                   dataRange,
                                   false,
                                   "excel_update_chart"))
                        {
                            string plotBy = ExcelArgumentReader.GetTrimmedString(arguments, "plotBy").ToLowerInvariant();
                            target.Chart.SetSourceData(
                                source.Range,
                                plotBy == "row" ? Excel.XlRowCol.xlRows : Excel.XlRowCol.xlColumns);
                        }
                        applied++;
                        changedFields.Add("dataRange");
                    }

                    if (removeTitle)
                    {
                        target.Chart.HasTitle = false;
                        applied++;
                        changedFields.Add("title");
                    }
                    else if (hasTitleArgument)
                    {
                        target.Chart.HasTitle = true;
                        Excel.ChartTitle chartTitle = null;
                        try
                        {
                            chartTitle = target.Chart.ChartTitle;
                            chartTitle.Text = title;
                        }
                        finally
                        {
                            ComHelper.Release(chartTitle);
                        }
                        applied++;
                        changedFields.Add("title");
                    }

                    if (chartTypeText.Length > 0)
                    {
                        target.Chart.ChartType = MapChartType(chartTypeText);
                        applied++;
                        changedFields.Add("type");
                    }

                    if (target.ChartObject != null)
                    {
                        bool widthProvided;
                        bool heightProvided;
                        double width = ExcelArgumentReader.GetInt(arguments, 0, "width");
                        double height = ExcelArgumentReader.GetInt(arguments, 0, "height");
                        widthProvided = width > 0;
                        heightProvided = height > 0;
                        if (widthProvided && width < 100)
                        {
                            return Fail(OfficeErrorCode.INVALID_PARAMS, "图表宽度必须不小于 100。");
                        }
                        if (heightProvided && height < 100)
                        {
                            return Fail(OfficeErrorCode.INVALID_PARAMS, "图表高度必须不小于 100。");
                        }
                        if (widthProvided)
                        {
                            target.ChartObject.Width = width;
                            applied++;
                            changedFields.Add("width");
                        }
                        if (heightProvided)
                        {
                            target.ChartObject.Height = height;
                            applied++;
                            changedFields.Add("height");
                        }

                        string position = ExcelArgumentReader.GetTrimmedString(arguments, "position", "targetCell");
                        if (position.Length > 0)
                        {
                            using (ExcelRangeContext positionContext = RangeResolver.Resolve(
                                       application,
                                       position,
                                       false,
                                       "excel_update_chart"))
                            {
                                target.ChartObject.Left = Convert.ToDouble(
                                    positionContext.Range.Left,
                                    CultureInfo.InvariantCulture);
                                target.ChartObject.Top = Convert.ToDouble(
                                    positionContext.Range.Top,
                                    CultureInfo.InvariantCulture);
                            }
                            applied++;
                            changedFields.Add("position");
                        }
                    }

                    if (applied == 0)
                    {
                        JObject unchanged = BuildChartInfo(
                            target.Chart,
                            target.ChartObject,
                            target.ChartName,
                            target.SheetName,
                            target.IsChartSheet,
                            0);
                        ExcelUndoScope.RecordUnavailable(unchanged, "未提供任何更新参数，图表未修改。");
                        return OperationResult.Succeed(
                            "未提供任何更新参数，图表未修改。",
                            unchanged,
                            Observation.Create("update_chart", "未提供更新参数，图表未修改。", false));
                    }
                }
                catch (Exception ex)
                {
                    JObject partialData = new JObject();
                    partialData["chartName"] = target.ChartName;
                    partialData["sheet"] = target.SheetName;
                    partialData["appliedCount"] = applied;
                    partialData["changedFields"] = new JArray(changedFields);
                    partialData["error"] = ex.Message;
                    ExcelUndoScope.RecordUnavailable(
                        partialData,
                        "已应用的图表属性不会自动回滚；请重新设置或重建图表。");
                    Observation partialObservation = Observation.Create(
                        "update_chart",
                        "图表更新部分完成（" + applied + " 项），后续失败：" + ex.Message,
                        applied > 0);
                    partialObservation.AddTargetRef(
                        ExcelComSupport.FullAddress(target.SheetName, target.ChartName));
                    return OperationResult.PartialApplied(
                        applied,
                        "已应用的图表属性不会自动回滚；请重新设置或重建图表。",
                        partialObservation,
                        partialData);
                }

                JObject data = BuildChartInfo(
                    target.Chart,
                    target.ChartObject,
                    target.ChartName,
                    target.SheetName,
                    target.IsChartSheet,
                    0);
                data["changedFields"] = new JArray(changedFields);
                ExcelUndoScope.RecordUnavailable(
                    data,
                    "图表对象状态不在单元格区域快照范围内，未登记自动撤销。");
                Observation observation = Observation.Create(
                    "update_chart",
                    "已更新图表「" + target.ChartName + "」的 " + changedFields.Count + " 项属性。",
                    applied > 0);
                observation.AddTargetRef(ExcelComSupport.FullAddress(target.SheetName, target.ChartName));
                return OperationResult.Succeed(
                    "图表已更新：" + target.ChartName + "（" + applied + " 项）。",
                    data,
                    observation);
            }
            finally
            {
                ReleaseChartTarget(target);
                ComHelper.Release(workbook);
            }
        }

        private static OperationResult ExecuteDeleteChart(Excel.Application application, JObject arguments)
        {
            string chartName = ExcelArgumentReader.GetTrimmedString(arguments, "chartName", "name", "chart");
            if (chartName.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 chartName 参数。");
            }
            string sheetName = ExcelArgumentReader.GetSheetName(arguments);

            Excel.Workbook workbook = null;
            ChartTarget target = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                string findError;
                target = FindChart(workbook, chartName, sheetName, out findError);
                if (target == null)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        findError.Length == 0 ? "找不到图表「" + chartName + "」。" : findError);
                }
                ExcelComSupport.EnsureWritable(
                    application,
                    workbook,
                    target.Worksheet,
                    "excel_delete_chart");

                string deletedName = target.ChartName;
                string deletedSheet = target.SheetName;
                bool isChartSheet = target.IsChartSheet;
                application.Guard(delegate
                {
                    if (target.IsChartSheet)
                    {
                        dynamic chartSheet = target.Chart;
                        chartSheet.Delete();
                    }
                    else
                    {
                        target.ChartObject.Delete();
                    }
                });

                JObject data = new JObject();
                data["chartName"] = deletedName;
                data["sheet"] = deletedSheet;
                data["isChartSheet"] = isChartSheet;
                data["deleted"] = true;
                ExcelUndoScope.RecordUnavailable(
                    data,
                    "图表删除属于对象级删除，单元格区域快照无法恢复，未登记自动撤销。");
                Observation observation = Observation.Create(
                    "delete_chart",
                    "已删除图表「" + deletedName + "」。",
                    true);
                observation.AddTargetRef(ExcelComSupport.FullAddress(deletedSheet, deletedName));
                return OperationResult.Succeed(
                    "图表已删除：" + deletedName + "。",
                    data,
                    observation);
            }
            finally
            {
                ReleaseChartTarget(target);
                ComHelper.Release(workbook);
            }
        }

        private static ChartTarget FindChart(
            Excel.Workbook workbook,
            string chartName,
            string sheetName,
            out string error)
        {
            error = string.Empty;
            List<Excel.Worksheet> sheets = new List<Excel.Worksheet>();
            Excel.Worksheet claimedWorksheet = null;
            try
            {
                if (!string.IsNullOrWhiteSpace(sheetName))
                {
                    Excel.Worksheet worksheet = ExcelComSupport.FindWorksheet(workbook, sheetName);
                    if (worksheet == null)
                    {
                        error = "找不到工作表「" + sheetName + "」。";
                        return null;
                    }
                    sheets.Add(worksheet);
                }
                else
                {
                    Excel.Sheets collection = workbook.Worksheets;
                    try
                    {
                        for (int index = 1; index <= collection.Count; index++)
                        {
                            object item = collection[index];
                            Excel.Worksheet worksheet = item as Excel.Worksheet;
                            if (worksheet == null)
                            {
                                ComHelper.Release(item);
                                continue;
                            }
                            sheets.Add(worksheet);
                        }
                    }
                    finally
                    {
                        ComHelper.Release(collection);
                    }
                }

                foreach (Excel.Worksheet worksheet in sheets)
                {
                    Excel.ChartObjects chartObjects = null;
                    try
                    {
                        chartObjects = (Excel.ChartObjects)worksheet.ChartObjects();
                        for (int index = 1; index <= chartObjects.Count; index++)
                        {
                            Excel.ChartObject chartObject = (Excel.ChartObject)chartObjects.Item(index);
                            string candidateName = chartObject.Name;
                            bool matches = string.Equals(candidateName, chartName, StringComparison.OrdinalIgnoreCase);
                            if (!matches && sheetName.Length == 0)
                            {
                                string fullName = worksheet.Name + "!" + candidateName;
                                matches = string.Equals(fullName, chartName, StringComparison.OrdinalIgnoreCase);
                            }
                            if (!matches)
                            {
                                ComHelper.Release(chartObject);
                                continue;
                            }
                            ChartTarget found = new ChartTarget();
                            found.ChartObject = chartObject;
                            found.Chart = chartObject.Chart;
                            found.Worksheet = worksheet;
                            found.SheetName = worksheet.Name;
                            found.ChartName = candidateName;
                            found.IsChartSheet = false;
                            claimedWorksheet = worksheet;
                            return found;
                        }
                    }
                    finally
                    {
                        ComHelper.Release(chartObjects);
                    }
                }

                Excel.Sheets chartSheets = null;
                try
                {
                    chartSheets = workbook.Charts;
                    for (int index = 1; index <= chartSheets.Count; index++)
                    {
                        object item = chartSheets[index];
                        Excel.Chart chart = item as Excel.Chart;
                        if (chart == null)
                        {
                            ComHelper.Release(item);
                            continue;
                        }
                        string candidateName = chart.Name;
                        if (!string.Equals(candidateName, chartName, StringComparison.OrdinalIgnoreCase))
                        {
                            ComHelper.Release(chart);
                            continue;
                        }
                        ChartTarget found = new ChartTarget();
                        found.Chart = chart;
                        found.Worksheet = null;
                        found.SheetName = candidateName;
                        found.ChartName = candidateName;
                        found.IsChartSheet = true;
                        return found;
                    }
                }
                finally
                {
                    ComHelper.Release(chartSheets);
                }

                error = string.IsNullOrWhiteSpace(sheetName)
                    ? "找不到图表「" + chartName + "」（支持内嵌图表名或 Sheet!图表名）。"
                    : "在工作表「" + sheetName + "」上找不到图表「" + chartName + "」。";
                return null;
            }
            finally
            {
                foreach (Excel.Worksheet worksheet in sheets)
                {
                    if (!object.ReferenceEquals(worksheet, claimedWorksheet))
                    {
                        ComHelper.Release(worksheet);
                    }
                }
                sheets.Clear();
            }
        }

        private static void ReleaseChartTarget(ChartTarget target)
        {
            if (target == null)
            {
                return;
            }
            ComHelper.Release(target.Chart, target.ChartObject);
            if (target.Worksheet != null)
            {
                ComHelper.Release(target.Worksheet);
            }
            target.Chart = null;
            target.ChartObject = null;
            target.Worksheet = null;
        }
    }
}
