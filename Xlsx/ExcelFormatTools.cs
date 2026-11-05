using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>单元格/图表格式化与图表创建。</summary>
    public sealed partial class ExcelToolExecutor
    {
        private static OperationResult ExecuteFormatRange(Excel.Application application, JObject arguments)
        {
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange", "target");
            if (string.IsNullOrWhiteSpace(address))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 range/targetRange 参数。");
            }

            string sheetName;
            string localAddress;
            RangeResolver.SplitAddress(address, out sheetName, out localAddress);

            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.ChartObjects chartObjects = null;
            Excel.ChartObject chartObject = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                if (!string.IsNullOrWhiteSpace(sheetName))
                {
                    worksheet = ExcelComSupport.FindWorksheet(workbook, sheetName);
                    if (worksheet == null)
                    {
                        return Fail(OfficeErrorCode.SHEET_NOT_FOUND, "找不到工作表「" + sheetName + "」。");
                    }
                }
                else
                {
                    object active = application.ActiveSheet;
                    worksheet = active as Excel.Worksheet;
                    if (worksheet == null)
                    {
                        ComHelper.Release(active);
                        return Fail(OfficeErrorCode.DOC_NOT_OPEN, "当前活动对象不是工作表。");
                    }
                }

                ExcelComSupport.EnsureWritable(application, workbook, worksheet, "excel_format_range");
                chartObjects = (Excel.ChartObjects)worksheet.ChartObjects();
                int chartCount = chartObjects.Count;
                for (int index = 1; index <= chartCount; index++)
                {
                    object chartItem = chartObjects.Item(index);
                    Excel.ChartObject candidate = chartItem as Excel.ChartObject;
                    if (candidate == null)
                    {
                        ComHelper.Release(chartItem);
                        continue;
                    }
                    string candidateName = candidate.Name;
                    if (string.Equals(candidateName, localAddress, StringComparison.OrdinalIgnoreCase))
                    {
                        chartObject = candidate;
                        break;
                    }
                    ComHelper.Release(candidate);
                }

                if (chartObject != null)
                {
                    string chartName = chartObject.Name;
                    FormatChartObject(chartObject, arguments);
                    JObject chartData = new JObject();
                    chartData["targetType"] = "chart";
                    chartData["chartName"] = chartName;
                    chartData["sheet"] = worksheet.Name;
                    Observation chartObservation = Observation.Create(
                        "format_chart",
                        "已格式化图表「" + chartName + "」。",
                        true);
                    chartObservation.AddTargetRef(ExcelComSupport.FullAddress(worksheet.Name, chartName));
                    ExcelUndoScope.RecordUnavailable(
                        chartData,
                        "图表对象属性不在单元格区域快照范围内，未登记自动撤销。");
                    return OperationResult.Succeed(
                        "已格式化图表「" + chartName + "」。",
                        chartData,
                        chartObservation);
                }

                Excel.Range targetRange = null;
                try
                {
                    targetRange = worksheet.Range[localAddress];
                }
                catch (COMException ex)
                {
                    return Fail(OfficeErrorCode.INVALID_RANGE, "无法解析范围「" + address + "」：" + ex.Message);
                }

                try
                {
                    string rangeAddress = targetRange.AddressLocal();
                    ExcelUndoScope undo = ExcelUndoScope.Capture(workbook, worksheet, targetRange, rangeAddress);
                    ApplyCellFormat(targetRange, arguments);

                    string targetRef = ExcelComSupport.FullAddress(worksheet.Name, rangeAddress);
                    Observation observation = Observation.Create(
                        "format_range",
                        "已格式化区域 " + targetRef + "。",
                        true);
                    observation.AddTargetRef(targetRef);

                    JObject data = new JObject();
                    data["targetType"] = "range";
                    data["sheet"] = worksheet.Name;
                    data["address"] = rangeAddress;
                    data["style"] = ExcelArgumentReader.GetString(arguments, "style");

                    undo.Complete(
                        workbook,
                        worksheet,
                        targetRange,
                        rangeAddress,
                        true,
                        "格式化区域 " + targetRef,
                        observation,
                        data,
                        "区域快照只保存值/公式/数字格式，字体、填充、边框等视觉格式不会被自动撤销。");

                    return OperationResult.Succeed(
                        "已格式化 " + targetRef + "。",
                        data,
                        observation);
                }
                finally
                {
                    ComHelper.Release(targetRange);
                }
            }
            finally
            {
                ComHelper.Release(chartObject, chartObjects, worksheet, workbook);
            }
        }

        private static void ApplyCellFormat(Excel.Range range, JObject arguments)
        {
            JObject styleObject = ExcelArgumentReader.Find(arguments, "style") as JObject;
            string preset = ExcelArgumentReader.GetString(arguments, "style").Trim().ToLowerInvariant();

            if (preset == "header")
            {
                Excel.Font headerFont = null;
                Excel.Interior headerInterior = null;
                try
                {
                    headerFont = range.Font;
                    headerFont.Bold = true;
                    headerFont.Color = ExcelComSupport.ToOleColor(255, 255, 255);
                    headerInterior = range.Interior;
                    headerInterior.Color = ExcelComSupport.ToOleColor(68, 114, 196);
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }
                finally
                {
                    ComHelper.Release(headerFont, headerInterior);
                }
            }
            else if (preset == "total")
            {
                Excel.Font totalFont = null;
                Excel.Border topBorder = null;
                try
                {
                    totalFont = range.Font;
                    totalFont.Bold = true;
                    topBorder = range.Borders[Excel.XlBordersIndex.xlEdgeTop];
                    topBorder.LineStyle = Excel.XlLineStyle.xlDouble;
                }
                finally
                {
                    ComHelper.Release(topBorder, totalFont);
                }
            }
            else if (preset == "data")
            {
                ExcelComSupport.ApplyBorder(range, "all");
            }

            JToken boldToken = StyleToken(arguments, styleObject, "bold");
            JToken italicToken = StyleToken(arguments, styleObject, "italic");
            JToken fontSizeToken = StyleToken(arguments, styleObject, "fontSize", "size");
            JToken fontNameToken = StyleToken(arguments, styleObject, "fontName");
            JToken fontColor = StyleToken(arguments, styleObject, "fontColor");
            JToken backgroundColor = StyleToken(arguments, styleObject, "backgroundColor", "bgColor");
            JToken borders = StyleToken(arguments, styleObject, "borders");
            JToken numberFormat = StyleToken(arguments, styleObject, "numberFormat");
            JToken wrapText = StyleToken(arguments, styleObject, "wrapText");

            Excel.Font font = null;
            try
            {
                font = range.Font;
                if (boldToken != null)
                {
                    font.Bold = ReadBoolean(boldToken, false);
                }
                if (italicToken != null)
                {
                    font.Italic = ReadBoolean(italicToken, false);
                }
                int fontSize;
                if (fontSizeToken != null && int.TryParse(fontSizeToken.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out fontSize) && fontSize > 0)
                {
                    font.Size = fontSize;
                }
                if (fontNameToken != null && !string.IsNullOrWhiteSpace(fontNameToken.ToString()))
                {
                    font.Name = fontNameToken.ToString();
                }
                if (fontColor != null && !string.IsNullOrWhiteSpace(fontColor.ToString()))
                {
                    font.Color = ExcelComSupport.ParseColor(fontColor.ToString(), ExcelComSupport.ToOleColor(0, 0, 0));
                }
            }
            finally
            {
                ComHelper.Release(font);
            }

            if (backgroundColor != null && !string.IsNullOrWhiteSpace(backgroundColor.ToString()))
            {
                ExcelComSupport.SetBackgroundColor(range, backgroundColor.ToString());
            }
            if (borders != null && !string.IsNullOrWhiteSpace(borders.ToString()))
            {
                ExcelComSupport.ApplyBorder(range, borders.ToString());
            }
            if (numberFormat != null && !string.IsNullOrWhiteSpace(numberFormat.ToString()))
            {
                range.NumberFormat = numberFormat.ToString();
            }
            if (wrapText != null)
            {
                range.WrapText = ReadBoolean(wrapText, false);
            }
        }

        private static void FormatChartObject(Excel.ChartObject chartObject, JObject arguments)
        {
            JObject styleObject = ExcelArgumentReader.Find(arguments, "style") as JObject;
            Excel.Chart chart = null;
            try
            {
                chart = chartObject.Chart;
                string title = ExcelArgumentReader.GetString(arguments, "title");
                if (!string.IsNullOrWhiteSpace(title))
                {
                    chart.HasTitle = true;
                    Excel.ChartTitle chartTitle = null;
                    try
                    {
                        chartTitle = chart.ChartTitle;
                        chartTitle.Text = title;
                    }
                    finally
                    {
                        ComHelper.Release(chartTitle);
                    }
                }

                JToken fontSizeToken = StyleToken(arguments, styleObject, "fontSize", "size");
                JToken fontColor = StyleToken(arguments, styleObject, "fontColor");
                JToken boldToken = StyleToken(arguments, styleObject, "bold");
                JToken backgroundColor = StyleToken(arguments, styleObject, "backgroundColor", "bgColor");
                JToken legendPosition = StyleToken(arguments, styleObject, "legendPosition");

                int fontSize = 0;
                if (fontSizeToken != null)
                {
                    int.TryParse(fontSizeToken.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out fontSize);
                }
                int color = -1;
                if (fontColor != null && !string.IsNullOrWhiteSpace(fontColor.ToString()))
                {
                    color = ExcelComSupport.ParseColor(fontColor.ToString(), ExcelComSupport.ToOleColor(0, 0, 0));
                }

                if (chart.HasTitle)
                {
                    Excel.ChartTitle titleObject = null;
                    Excel.Font titleFont = null;
                    try
                    {
                        titleObject = chart.ChartTitle;
                        titleFont = titleObject.Font;
                        ApplyFontOptions(titleFont, fontSize, color, boldToken);
                    }
                    finally
                    {
                        ComHelper.Release(titleFont, titleObject);
                    }
                }

                TryFormatChartAxes(chart, fontSize, color);

                if (chart.HasLegend)
                {
                    Excel.Legend legend = null;
                    Excel.Font legendFont = null;
                    try
                    {
                        legend = chart.Legend;
                        legendFont = legend.Font;
                        ApplyFontOptions(legendFont, fontSize, color, null);
                        if (legendPosition != null)
                        {
                            ApplyLegendPosition(legend, legendPosition.ToString());
                        }
                    }
                    finally
                    {
                        ComHelper.Release(legendFont, legend);
                    }
                }

                if (backgroundColor != null && !string.IsNullOrWhiteSpace(backgroundColor.ToString()))
                {
                    Excel.ChartArea area = null;
                    Excel.Interior interior = null;
                    try
                    {
                        area = chart.ChartArea;
                        interior = area.Interior;
                        interior.Color = ExcelComSupport.ParseColor(
                            backgroundColor.ToString(),
                            ExcelComSupport.ToOleColor(255, 255, 255));
                    }
                    finally
                    {
                        ComHelper.Release(interior, area);
                    }
                }
            }
            finally
            {
                ComHelper.Release(chart);
            }
        }

        private static void TryFormatChartAxes(Excel.Chart chart, int fontSize, int color)
        {
            Excel.Axes axes = null;
            try
            {
                axes = (Excel.Axes)chart.Axes();
                foreach (object item in axes)
                {
                    Excel.Axis axis = item as Excel.Axis;
                    if (axis == null)
                    {
                        ComHelper.Release(item);
                        continue;
                    }
                    Excel.TickLabels labels = null;
                    Excel.Font labelsFont = null;
                    try
                    {
                        labels = axis.TickLabels;
                        labelsFont = labels.Font;
                        ApplyFontOptions(labelsFont, fontSize, color, null);
                    }
                    finally
                    {
                        ComHelper.Release(labelsFont, labels, axis);
                    }
                }
            }
            catch (COMException)
            {
                // 饼图等无坐标轴图表会抛异常，格式化标题/图例仍然有效。
            }
            finally
            {
                ComHelper.Release(axes);
            }
        }

        private static void ApplyFontOptions(Excel.Font font, int fontSize, int color, JToken boldToken)
        {
            if (font == null)
            {
                return;
            }
            if (fontSize > 0)
            {
                font.Size = fontSize;
            }
            if (color >= 0)
            {
                font.Color = color;
            }
            if (boldToken != null)
            {
                font.Bold = ReadBoolean(boldToken, false);
            }
        }

        private static OperationResult ExecuteCreateChart(Excel.Application application, JObject arguments)
        {
            string dataRange = ExcelArgumentReader.GetTrimmedString(arguments, "dataRange", "sourceRange", "range");
            if (dataRange.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 dataRange 参数。");
            }

            JObject chartInfo = null;
            application.Guard(delegate
            {
                chartInfo = CreateChartCore(application, dataRange, arguments);
            });

            string chartName = chartInfo.Value<string>("chartName");
            string sheetName = chartInfo.Value<string>("sheet");
            string targetRef = ExcelComSupport.FullAddress(sheetName, chartName);
            Observation observation = Observation.Create(
                "create_chart",
                "已在 " + targetRef + " 创建图表。",
                true);
            observation.AddTargetRef(targetRef);
            ExcelUndoScope.RecordUnavailable(
                chartInfo,
                "图表对象状态不在单元格区域快照范围内，未登记自动撤销。");

            return OperationResult.Succeed("图表已创建：" + targetRef + "。", chartInfo, observation);
        }

        internal static JObject CreateChartCore(Excel.Application application, string dataRange, JObject arguments)
        {
            using (ExcelRangeContext source = RangeResolver.Resolve(
                       application, dataRange, false, "excel_create_chart"))
            {
                ExcelComSupport.EnsureWritable(application, source.Workbook, source.Worksheet, "excel_create_chart");

                string position = ExcelArgumentReader.GetTrimmedString(arguments, "position", "targetCell");
                string positionSheet;
                string positionAddress;
                RangeResolver.SplitAddress(position, out positionSheet, out positionAddress);

                Excel.Worksheet chartSheet = null;
                Excel.Range positionRange = null;
                bool ownsChartSheet = false;
                bool chartCreated = false;
                Excel.ChartObjects chartObjects = null;
                Excel.ChartObject chartObject = null;
                try
                {
                    if (!string.IsNullOrWhiteSpace(positionSheet))
                    {
                        chartSheet = ExcelComSupport.FindWorksheet(source.Workbook, positionSheet);
                        if (chartSheet == null)
                        {
                            throw new OfficeComException(OfficeErrorCode.SHEET_NOT_FOUND, "找不到图表定位工作表「" + positionSheet + "」。");
                        }
                        ownsChartSheet = true;
                        positionRange = chartSheet.Range[positionAddress];
                    }
                    else
                    {
                        chartSheet = source.Worksheet;
                        if (!string.IsNullOrWhiteSpace(positionAddress))
                        {
                            positionRange = chartSheet.Range[positionAddress];
                        }
                        else
                        {
                            int sourceColumns = source.Range.Columns.Count;
                            positionRange = source.Range.Offset[0, sourceColumns + 1];
                        }
                    }

                    double width = ExcelArgumentReader.GetInt(arguments, 450, "width");
                    double height = ExcelArgumentReader.GetInt(arguments, 320, "height");
                    if (width < 100 || height < 100)
                    {
                        throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "图表宽高必须不小于 100。");
                    }

                    chartObjects = (Excel.ChartObjects)chartSheet.ChartObjects();
                    chartObject = chartObjects.Add(
                        Convert.ToDouble(positionRange.Left, CultureInfo.InvariantCulture),
                        Convert.ToDouble(positionRange.Top, CultureInfo.InvariantCulture),
                        width,
                        height);
                    chartCreated = true;
                    Excel.Chart chart = null;
                    try
                    {
                        chart = chartObject.Chart;
                        Excel.XlChartType chartType = MapChartType(ExcelArgumentReader.GetString(arguments, "type"));
                        chart.ChartType = chartType;

                        string plotBy = ExcelArgumentReader.GetString(arguments, "plotBy").Trim().ToLowerInvariant();
                        chart.SetSourceData(
                            source.Range,
                            plotBy == "row" ? Excel.XlRowCol.xlRows : Excel.XlRowCol.xlColumns);

                        string title = ExcelArgumentReader.GetString(arguments, "title");
                        if (!string.IsNullOrWhiteSpace(title))
                        {
                            chart.HasTitle = true;
                            Excel.ChartTitle chartTitle = null;
                            try
                            {
                                chartTitle = chart.ChartTitle;
                                chartTitle.Text = title;
                            }
                            finally
                            {
                                ComHelper.Release(chartTitle);
                            }
                        }

                        ApplySeriesNames(application, chart, arguments, source);
                        ApplyCategoryAxis(application, chart, arguments, source);
                        ApplyLegend(chart, ExcelArgumentReader.GetString(arguments, "legendPosition"));
                    }
                    finally
                    {
                        ComHelper.Release(chart);
                    }

                    JObject info = new JObject();
                    info["chartName"] = chartObject.Name;
                    info["sheet"] = chartSheet.Name;
                    info["position"] = positionRange.AddressLocal();
                    info["width"] = width;
                    info["height"] = height;
                    info["type"] = ExcelArgumentReader.GetString(arguments, "type");
                    return info;
                }
                catch
                {
                    if (chartCreated && chartObject != null)
                    {
                        try
                        {
                            chartObject.Delete();
                        }
                        catch (Exception cleanupError)
                        {
                            Console.Error.WriteLine("[excel] 创建图表失败后清理图表对象失败：" + cleanupError.Message);
                        }
                    }
                    throw;
                }
                finally
                {
                    ComHelper.Release(chartObject, chartObjects, positionRange);
                    if (ownsChartSheet)
                    {
                        ComHelper.Release(chartSheet);
                    }
                }
            }
        }

        private static void ApplySeriesNames(
            Excel.Application application,
            Excel.Chart chart,
            JObject arguments,
            ExcelRangeContext source)
        {
            JToken namesToken = ExcelArgumentReader.First(arguments, "seriesNames");
            if (namesToken == null)
            {
                return;
            }

            Excel.SeriesCollection seriesCollection = null;
            try
            {
                seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection();
                int seriesCount = seriesCollection.Count;

                JArray namesArray = namesToken as JArray;
                if (namesArray != null)
                {
                    for (int index = 0; index < Math.Min(seriesCount, namesArray.Count); index++)
                    {
                        Excel.Series series = null;
                        try
                        {
                            series = seriesCollection.Item(index + 1);
                            series.Name = namesArray[index].ToString();
                        }
                        finally
                        {
                            ComHelper.Release(series);
                        }
                    }
                    return;
                }

                string namesAddress = namesToken.ToString().Trim();
                if (namesAddress.Length == 0)
                {
                    return;
                }
                using (ExcelRangeContext namesRange = RangeResolver.ResolveOnSheet(
                           application,
                           source.SheetName,
                           namesAddress,
                           false,
                           "excel_create_chart"))
                {
                    int rows = Math.Min(namesRange.Range.Rows.Count, seriesCount);
                    for (int index = 1; index <= rows; index++)
                    {
                        Excel.Range cell = null;
                        Excel.Series series = null;
                        try
                        {
                            cell = (Excel.Range)namesRange.Range.Cells[index, 1];
                            object raw = cell.Value2;
                            if (raw == null || raw is DBNull)
                            {
                                continue;
                            }
                            series = seriesCollection.Item(index);
                            series.Name = Convert.ToString(raw, CultureInfo.InvariantCulture);
                        }
                        finally
                        {
                            ComHelper.Release(series, cell);
                        }
                    }
                }
            }
            finally
            {
                ComHelper.Release(seriesCollection);
            }
        }

        private static void ApplyCategoryAxis(
            Excel.Application application,
            Excel.Chart chart,
            JObject arguments,
            ExcelRangeContext source)
        {
            JToken axisToken = ExcelArgumentReader.First(arguments, "categoryAxis");
            if (axisToken == null)
            {
                return;
            }
            Excel.SeriesCollection seriesCollection = null;
            Excel.Series firstSeries = null;
            try
            {
                seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection();
                if (seriesCollection.Count == 0)
                {
                    return;
                }
                firstSeries = seriesCollection.Item(1);

                JArray axisArray = axisToken as JArray;
                if (axisArray != null)
                {
                    object[] labels = new object[axisArray.Count];
                    for (int index = 0; index < axisArray.Count; index++)
                    {
                        labels[index] = ExcelComSupport.JTokenToValue(axisArray[index]);
                    }
                    firstSeries.XValues = labels;
                    return;
                }

                string axisAddress = axisToken.ToString().Trim();
                if (axisAddress.Length == 0)
                {
                    return;
                }
                using (ExcelRangeContext axisRange = RangeResolver.ResolveOnSheet(
                           application,
                           source.SheetName,
                           axisAddress,
                           false,
                           "excel_create_chart"))
                {
                    firstSeries.XValues = axisRange.Range;
                }
            }
            finally
            {
                ComHelper.Release(firstSeries, seriesCollection);
            }
        }

        private static void ApplyLegend(Excel.Chart chart, string legendPosition)
        {
            string position = (legendPosition ?? string.Empty).Trim().ToLowerInvariant();
            if (position == "none")
            {
                chart.HasLegend = false;
                return;
            }
            chart.HasLegend = true;
            if (position.Length == 0)
            {
                return;
            }
            Excel.Legend legend = null;
            try
            {
                legend = chart.Legend;
                ApplyLegendPosition(legend, position);
            }
            finally
            {
                ComHelper.Release(legend);
            }
        }

        private static void ApplyLegendPosition(Excel.Legend legend, string legendPosition)
        {
            if (legend == null)
            {
                return;
            }
            string position = (legendPosition ?? string.Empty).Trim().ToLowerInvariant();
            if (position == "right")
            {
                legend.Position = Excel.XlLegendPosition.xlLegendPositionRight;
            }
            else if (position == "left")
            {
                legend.Position = Excel.XlLegendPosition.xlLegendPositionLeft;
            }
            else if (position == "top")
            {
                legend.Position = Excel.XlLegendPosition.xlLegendPositionTop;
            }
            else if (position == "bottom")
            {
                legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom;
            }
            else if (position == "corner")
            {
                legend.Position = Excel.XlLegendPosition.xlLegendPositionCorner;
            }
        }

        private static Excel.XlChartType MapChartType(string type)
        {
            string value = (type ?? string.Empty).Trim().ToLowerInvariant();
            switch (value)
            {
                case "line":
                    return Excel.XlChartType.xlLine;
                case "pie":
                    return Excel.XlChartType.xlPie;
                case "bar":
                    return Excel.XlChartType.xlBarClustered;
                case "scatter":
                    return Excel.XlChartType.xlXYScatter;
                case "area":
                    return Excel.XlChartType.xlArea;
                default:
                    return Excel.XlChartType.xlColumnClustered;
            }
        }

        private static JToken StyleToken(JObject arguments, JObject styleObject, params string[] names)
        {
            if (styleObject != null)
            {
                JToken nested = ExcelArgumentReader.First(styleObject, names);
                if (nested != null)
                {
                    return nested;
                }
            }
            return ExcelArgumentReader.First(arguments, names);
        }

        private static bool ReadBoolean(JToken token, bool defaultValue)
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                return defaultValue;
            }
            if (token.Type == JTokenType.Boolean)
            {
                return token.Value<bool>();
            }
            bool parsed;
            return bool.TryParse(token.ToString(), out parsed) ? parsed : defaultValue;
        }
    }
}
