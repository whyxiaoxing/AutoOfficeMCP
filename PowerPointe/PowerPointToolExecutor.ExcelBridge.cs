using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Newtonsoft.Json.Linq;
using Office;
using Office.Config;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointe
{
    public sealed partial class PowerPointToolExecutor
    {
        private static OperationResult ExecuteInsertTableFromExcel(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.Shape tableShape = null;
            PowerPoint.Table table = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                string excelFile = PowerPointArgumentReader.RequireString(arguments, "excelFile");
                string rangeAddress = PowerPointArgumentReader.RequireString(
                    arguments, "range", "rangeAddress");
                string sheetName = PowerPointArgumentReader.GetTrimmedString(arguments, "sheetName");
                bool hasHeader = PowerPointArgumentReader.GetBool(arguments, true, "hasHeader");
                JArray data;
                OfficeErrorCode readCode;
                string readError;
                if (!TryReadExcelRange(
                        excelFile, sheetName, rangeAddress, out data, out readCode, out readError))
                {
                    return Fail(readCode, readError);
                }
                if (data == null || data.Count == 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "Excel 区域没有可读取的数据。");
                }
                data = NormalizeExcelTable(data, hasHeader);
                int rows = data.Count;
                int cols = MaxColumnCount(data);
                if (rows <= 0 || cols <= 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "Excel 区域为空或没有有效列。");
                }

                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                float x = PowerPointArgumentReader.GetFloat(arguments, 50f, "x");
                float y = PowerPointArgumentReader.GetFloat(arguments, 150f, "y");
                float width = PowerPointArgumentReader.GetFloat(arguments, 600f, "width");
                float height = PowerPointArgumentReader.GetFloat(arguments, 300f, "height");
                if (width <= 0f || height <= 0f)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "width/height 必须大于 0。");
                }

                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                tableShape = slide.Shapes.AddTable(rows, cols, x, y, width, height);
                table = tableShape.Table;
                int filled = FillTable(table, data, rows, cols);

                JObject result = new JObject();
                result["slideIndex"] = slide.SlideIndex;
                result["shapeId"] = PowerPointComUtil.SafeInt(delegate
                {
                    dynamic dynamicShape = tableShape;
                    return Convert.ToInt32((object)dynamicShape.Id);
                });
                result["shapeName"] = tableShape.Name;
                result["rows"] = rows;
                result["cols"] = cols;
                result["filledCells"] = filled;
                result["excelFile"] = excelFile;
                result["range"] = rangeAddress;
                result["hasHeader"] = hasHeader;
                return Success(
                    "已从 Excel 区域 " + rangeAddress + " 向第 " +
                    slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片插入 " + rows.ToString(CultureInfo.InvariantCulture) + "×" +
                    cols.ToString(CultureInfo.InvariantCulture) + " 表格。",
                    "write",
                    true,
                    SlideRef(presentation, slide) + "!" + tableShape.Name,
                    result);
            }
            finally
            {
                ComHelper.Release(table, tableShape, slide, presentation);
            }
        }

        private static OperationResult ExecuteInsertChartFromExcel(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                string excelFile = PowerPointArgumentReader.RequireString(arguments, "excelFile");
                string rangeAddress = PowerPointArgumentReader.RequireString(
                    arguments, "range", "rangeAddress");
                string sheetName = PowerPointArgumentReader.GetTrimmedString(arguments, "sheetName");
                bool hasHeader = PowerPointArgumentReader.GetBool(arguments, true, "hasHeader");
                string chartType = PowerPointArgumentReader.GetTrimmedString(arguments, "chartType");
                if (chartType.Length == 0)
                {
                    chartType = "column";
                }
                JArray data;
                OfficeErrorCode readCode;
                string readError;
                if (!TryReadExcelRange(
                        excelFile, sheetName, rangeAddress, out data, out readCode, out readError))
                {
                    return Fail(readCode, readError);
                }
                if (data == null || data.Count == 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "Excel 区域没有可读取的数据。");
                }
                data = NormalizeExcelChart(data, hasHeader);
                int rows = data.Count;
                int columns = MaxColumnCount(data);
                if (rows < 2 || columns < 2)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "图表数据至少需要两行两列（表头 + 一行数据，类别列 + 数值列）。");
                }

                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                float x = PowerPointArgumentReader.GetFloat(arguments, 50f, "x");
                float y = PowerPointArgumentReader.GetFloat(arguments, 150f, "y");
                float width = PowerPointArgumentReader.GetFloat(arguments, 600f, "width");
                float height = PowerPointArgumentReader.GetFloat(arguments, 350f, "height");
                if (width <= 0f || height <= 0f)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "width/height 必须大于 0。");
                }
                string title = PowerPointArgumentReader.GetTrimmedString(arguments, "title");
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                OperationResult result = InsertChartFromData(
                    slide, data, chartType, title, x, y, width, height);
                if (result == null)
                {
                    return Fail(OfficeErrorCode.INTERNAL_ERROR, "插入 Excel 图表失败。");
                }
                if (result.Data != null)
                {
                    result.Data["excelFile"] = excelFile;
                    result.Data["range"] = rangeAddress;
                    result.Data["hasHeader"] = hasHeader;
                }
                return result;
            }
            finally
            {
                ComHelper.Release(slide, presentation);
            }
        }

        private static JArray NormalizeExcelTable(JArray data, bool hasHeader)
        {
            if (hasHeader || data == null || data.Count == 0)
            {
                return data;
            }
            JArray normalized = new JArray();
            int columns = MaxColumnCount(data);
            JArray header = new JArray();
            for (int index = 0; index < columns; index++)
            {
                header.Add("Column" + (index + 1).ToString(CultureInfo.InvariantCulture));
            }
            normalized.Add(header);
            foreach (JToken row in data)
            {
                normalized.Add(row.DeepClone());
            }
            return normalized;
        }

        private static JArray NormalizeExcelChart(JArray data, bool hasHeader)
        {
            if (hasHeader || data == null || data.Count == 0)
            {
                return data;
            }
            JArray normalized = new JArray();
            int columns = MaxColumnCount(data);
            JArray header = new JArray();
            for (int index = 0; index < columns; index++)
            {
                header.Add(index == 0
                    ? "Category"
                    : "Series" + index.ToString(CultureInfo.InvariantCulture));
            }
            normalized.Add(header);
            foreach (JToken row in data)
            {
                normalized.Add(row.DeepClone());
            }
            return normalized;
        }

        private static int MaxColumnCount(JArray rows)
        {
            int columns = 0;
            if (rows == null)
            {
                return 0;
            }
            foreach (JToken rowToken in rows)
            {
                JArray row = rowToken as JArray;
                if (row != null)
                {
                    columns = Math.Max(columns, row.Count);
                }
            }
            return columns;
        }

        private static bool TryReadExcelRange(
            string excelFile,
            string sheetName,
            string rangeAddress,
            out JArray data,
            out OfficeErrorCode errorCode,
            out string error)
        {
            data = null;
            errorCode = OfficeErrorCode.INTERNAL_ERROR;
            error = string.Empty;
            string filePath;
            string pathError;
            if (!WorkspacePathResolver.TryResolve(excelFile, out filePath, out pathError))
            {
                errorCode = OfficeErrorCode.INVALID_PARAMS;
                error = pathError;
                return false;
            }
            if (!File.Exists(filePath))
            {
                errorCode = OfficeErrorCode.FILE_NOT_FOUND;
                error = "Excel 文件不存在：" + filePath;
                return false;
            }

            string effectiveSheet = sheetName;
            string effectiveRange = rangeAddress;
            int bang = rangeAddress.LastIndexOf('!');
            if (bang > 0)
            {
                effectiveSheet = rangeAddress.Substring(0, bang).Trim('\'', '"');
                effectiveRange = rangeAddress.Substring(bang + 1);
            }
            if (string.IsNullOrWhiteSpace(effectiveRange))
            {
                errorCode = OfficeErrorCode.INVALID_PARAMS;
                error = "range 不能为空，例如 Sheet1!A1:D10 或 A1:D10。";
                return false;
            }

            Excel.Application excel = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Range range = null;
            try
            {
                excel = new Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                workbook = excel.Workbooks.Open(filePath, 0, true);
                if (string.IsNullOrWhiteSpace(effectiveSheet))
                {
                    worksheet = workbook.ActiveSheet as Excel.Worksheet;
                }
                else
                {
                    worksheet = workbook.Worksheets[effectiveSheet] as Excel.Worksheet;
                }
                if (worksheet == null)
                {
                    errorCode = OfficeErrorCode.SHEET_NOT_FOUND;
                    error = "找不到工作表「" + effectiveSheet + "」。";
                    return false;
                }
                range = worksheet.Range[effectiveRange];
                object value = range.Value2;
                data = ConvertExcelValues(value);
                if (data == null || data.Count == 0)
                {
                    errorCode = OfficeErrorCode.INVALID_PARAMS;
                    error = "Excel 区域没有数据：" + effectiveRange;
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                if (ex is COMException && ((COMException)ex).HResult == OfficeErrorMapper.ExcelGeneric)
                {
                    errorCode = OfficeErrorCode.INVALID_RANGE;
                }
                else if (ex is COMException &&
                         ((COMException)ex).HResult == OfficeErrorMapper.ExcelSubscriptOutOfRange)
                {
                    errorCode = OfficeErrorCode.SHEET_NOT_FOUND;
                }
                else
                {
                    errorCode = OfficeErrorCode.INTERNAL_ERROR;
                }
                error = "读取 Excel 区域失败：" + ex.Message;
                return false;
            }
            finally
            {
                if (workbook != null)
                {
                    try
                    {
                        workbook.Close(false);
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[ppt] 关闭 Excel 工作簿失败：" + ex.Message);
                    }
                }
                if (excel != null)
                {
                    try
                    {
                        excel.Quit();
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[ppt] 退出 Excel 实例失败：" + ex.Message);
                    }
                }
                ComHelper.Release(range, worksheet, workbook, excel);
            }
        }

        private static JArray ConvertExcelValues(object value)
        {
            JArray rows = new JArray();
            if (value == null)
            {
                return rows;
            }
            object[,] matrix = value as object[,];
            if (matrix != null)
            {
                int rowCount = matrix.GetLength(0);
                int colCount = matrix.GetLength(1);
                for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++)
                {
                    JArray row = new JArray();
                    for (int colIndex = 1; colIndex <= colCount; colIndex++)
                    {
                        row.Add(ToJToken(matrix[rowIndex, colIndex]));
                    }
                    rows.Add(row);
                }
                return rows;
            }
            JArray single = new JArray();
            single.Add(ToJToken(value));
            rows.Add(single);
            return rows;
        }

        private static JToken ToJToken(object value)
        {
            if (value == null)
            {
                return JValue.CreateNull();
            }
            if (value is DateTime)
            {
                return new JValue(((DateTime)value).ToString("o", CultureInfo.InvariantCulture));
            }
            if (value is double)
            {
                return new JValue((double)value);
            }
            if (value is float)
            {
                return new JValue((float)value);
            }
            if (value is decimal)
            {
                return new JValue((decimal)value);
            }
            if (value is int)
            {
                return new JValue((int)value);
            }
            if (value is long)
            {
                return new JValue((long)value);
            }
            if (value is bool)
            {
                return new JValue((bool)value);
            }
            return new JValue(Convert.ToString(value, CultureInfo.InvariantCulture));
        }

        private static OperationResult InsertChartFromData(
            PowerPoint.Slide slide,
            JArray data,
            string chartType,
            string title,
            float x,
            float y,
            float width,
            float height)
        {
            PowerPoint.Shape shape = null;
            object chartDataObject = null;
            object workbookObject = null;
            object worksheetObject = null;
            object rangeObject = null;
            try
            {
                int rows = data.Count;
                int columns = MaxColumnCount(data);
                dynamic shapes = slide.Shapes;
                object chartShapeObject = shapes.AddChart2(
                    -1,
                    PowerPointComUtil.MapChartType(chartType),
                    x,
                    y,
                    width,
                    height,
                    0);
                shape = chartShapeObject as PowerPoint.Shape;
                if (shape == null)
                {
                    return Fail(OfficeErrorCode.INTERNAL_ERROR, "PowerPoint 未返回图表形状对象。");
                }
                PowerPoint.Chart chart = shape.Chart;
                PowerPoint.ChartData chartData = chart.ChartData;
                chartDataObject = chartData;
                chartData.Activate();
                workbookObject = TryGetChartWorkbook(chartDataObject, chartData);
                if (workbookObject == null)
                {
                    return Fail(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "无法打开 PowerPoint 图表数据工作簿（Excel 自动化不可用或图表初始化失败）。");
                }
                dynamic workbook = workbookObject;
                worksheetObject = workbook.Worksheets[1];
                dynamic worksheet = worksheetObject;
                dynamic cells = worksheet.Cells;
                cells.Clear();

                for (int rowIndex = 0; rowIndex < rows; rowIndex++)
                {
                    JArray row = data[rowIndex] as JArray;
                    for (int columnIndex = 0; columnIndex < columns; columnIndex++)
                    {
                        dynamic cell = worksheet.Cells[rowIndex + 1, columnIndex + 1];
                        JToken value = row != null && columnIndex < row.Count ? row[columnIndex] : null;
                        cell.Value2 = value == null || value.Type == JTokenType.Null
                            ? null
                            : value.ToObject<object>();
                        ComHelper.Release(cell);
                    }
                }

                rangeObject = worksheet.Range(
                    worksheet.Cells[1, 1],
                    worksheet.Cells[rows, columns]);
                if (title.Length > 0)
                {
                    dynamic dynamicChart = chart;
                    dynamicChart.HasTitle = -1;
                    dynamicChart.ChartTitle.Text = title;
                }
                try
                {
                    workbook.Close(false);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[ppt] 关闭图表数据工作簿失败：" + ex.Message);
                }

                JObject resultData = new JObject();
                resultData["slideIndex"] = slide.SlideIndex;
                resultData["shapeId"] = PowerPointComUtil.SafeInt(delegate
                {
                    dynamic dynamicShape = shape;
                    return Convert.ToInt32((object)dynamicShape.Id);
                });
                resultData["shapeName"] = shape.Name;
                resultData["chartType"] = chartType;
                resultData["rows"] = rows - 1;
                resultData["seriesCount"] = columns - 1;
                return Success(
                    "已在第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片插入 Excel 图表（" + (rows - 1).ToString(CultureInfo.InvariantCulture) +
                    " 行，" + (columns - 1).ToString(CultureInfo.InvariantCulture) + " 个系列）。",
                    "write",
                    true,
                    "ppt:Slide" + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    "!" + shape.Name,
                    resultData);
            }
            catch (Exception ex)
            {
                return OfficeErrorMapper.ToOperationResult(ex, OfficeApp.PowerPoint);
            }
            finally
            {
                ComHelper.Release(rangeObject, worksheetObject, workbookObject, chartDataObject, shape);
            }
        }
    }
}
