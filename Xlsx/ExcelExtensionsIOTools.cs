using System;
using System.Globalization;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>P2 导入导出：CSV / JSON / PDF / 选区图片（PNG）。</summary>
    public sealed partial class ExcelExtensionsExecutor
    {
        private static OperationResult ExecuteExport(Excel.Application application, JObject arguments)
        {
            string format = ExcelArgumentReader.GetTrimmedString(arguments, "format", "type").ToLowerInvariant();
            if (format.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 format 参数（csv/json/pdf/image）。");
            }
            if (format == "png")
            {
                format = "image";
            }
            if (format != "csv" && format != "json" && format != "pdf" && format != "image")
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "format 仅支持 csv/json/pdf/image。");
            }

            bool overwrite = ExcelArgumentReader.GetBool(arguments, false, "overwrite");
            string path = ExcelExtensionSupport.ResolveFilePath(arguments, false, overwrite);
            bool includeFormulas = ExcelArgumentReader.GetBool(arguments, false, "includeFormulas");
            string delimiter = ExcelExtensionSupport.NormalizeDelimiter(
                ExcelArgumentReader.GetString(arguments, "delimiter"));
            string encodingName = ExcelArgumentReader.GetString(arguments, "encoding");
            if (encodingName.Length == 0)
            {
                encodingName = "utf8bom";
            }
            string scope = ExcelArgumentReader.GetTrimmedString(arguments, "scope").ToLowerInvariant();
            if (scope.Length == 0)
            {
                scope = "range";
            }

            string rangeAddress = ExcelArgumentReader.GetRangeAddress(
                arguments,
                "range",
                "sourceRange",
                "targetRange",
                "address");
            if (string.IsNullOrWhiteSpace(rangeAddress))
            {
                string sheetName = ExcelArgumentReader.GetSheetName(arguments);
                rangeAddress = sheetName.Length == 0
                    ? "usedrange"
                    : ExcelComSupport.QuoteSheetName(sheetName) + "!usedrange";
            }

            Excel.Workbook workbook = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                using (ExcelRangeContext context = RangeResolver.Resolve(
                           application,
                           rangeAddress,
                           false,
                           "excel_export"))
                {
                    int rows = context.Range.Rows.Count;
                    int columns = context.Range.Columns.Count;
                    object[,] values = ExcelExtensionSupport.ReadMatrix(context.Range, rows, columns);
                    long bytesWritten = 0;
                    string layout = string.Empty;

                    if (format == "csv")
                    {
                        string csv = ExcelExtensionSupport.BuildCsv(values, delimiter);
                        Encoding encoding = ExcelExtensionSupport.ResolveEncoding(encodingName, true);
                        File.WriteAllText(path, csv, encoding);
                        bytesWritten = new FileInfo(path).Length;
                        layout = "csv";
                    }
                    else if (format == "json")
                    {
                        JObject root = new JObject();
                        root["sheet"] = context.SheetName;
                        root["address"] = context.Range.AddressLocal();
                        root["rows"] = rows;
                        root["columns"] = columns;
                        root["values"] = ExcelExtensionSupport.MatrixToJson(values);
                        if (includeFormulas)
                        {
                            object[,] formulas = ExcelExtensionSupport.ReadFormulaMatrix(
                                context.Range,
                                rows,
                                columns);
                            root["formulas"] = ExcelExtensionSupport.MatrixToJson(formulas);
                        }
                        string json = root.ToString(Formatting.Indented);
                        File.WriteAllText(path, json, new UTF8Encoding(false));
                        bytesWritten = new FileInfo(path).Length;
                        layout = "json";
                    }
                    else if (format == "pdf")
                    {
                        if (scope == "sheet")
                        {
                            context.Worksheet.ExportAsFixedFormat(
                                Excel.XlFixedFormatType.xlTypePDF,
                                path,
                                Excel.XlFixedFormatQuality.xlQualityStandard,
                                true,
                                false,
                                Type.Missing,
                                Type.Missing,
                                false,
                                Type.Missing);
                        }
                        else
                        {
                            context.Range.ExportAsFixedFormat(
                                Excel.XlFixedFormatType.xlTypePDF,
                                path,
                                Excel.XlFixedFormatQuality.xlQualityStandard,
                                true,
                                false,
                                Type.Missing,
                                Type.Missing,
                                false,
                                Type.Missing);
                        }
                        if (!File.Exists(path))
                        {
                            return Fail(OfficeErrorCode.INTERNAL_ERROR, "PDF 导出后文件不存在：" + path);
                        }
                        bytesWritten = new FileInfo(path).Length;
                        layout = scope;
                    }
                    else
                    {
                        bytesWritten = ExportRangeAsImage(application, context.Worksheet, context.Range, path);
                        layout = "image";
                    }

                    JObject data = new JObject();
                    data["format"] = format;
                    data["path"] = path;
                    data["sheet"] = context.SheetName;
                    data["address"] = context.Range.AddressLocal();
                    data["rows"] = rows;
                    data["columns"] = columns;
                    data["bytes"] = bytesWritten;
                    data["layout"] = layout;
                    Observation observation = ReadOnlyObservation(
                        "export",
                        "已导出 " + format + " 到 " + path + "（" + bytesWritten + " 字节）。");
                    return OperationResult.Succeed(
                        "已导出 " + format + "：" + path + "（" + bytesWritten + " 字节）。",
                        data,
                        observation);
                }
            }
            finally
            {
                ComHelper.Release(workbook);
            }
        }

        private static long ExportRangeAsImage(
            Excel.Application application,
            Excel.Worksheet worksheet,
            Excel.Range range,
            string path)
        {
            Excel.ChartObjects chartObjects = null;
            Excel.ChartObject chartObject = null;
            Excel.Chart chart = null;
            try
            {
                double width = Math.Max(120.0, Convert.ToDouble(range.Width, CultureInfo.InvariantCulture));
                double height = Math.Max(80.0, Convert.ToDouble(range.Height, CultureInfo.InvariantCulture));
                double left = Convert.ToDouble(range.Left, CultureInfo.InvariantCulture);
                double top = Convert.ToDouble(range.Top, CultureInfo.InvariantCulture);

                range.CopyPicture(
                    Excel.XlPictureAppearance.xlScreen,
                    Excel.XlCopyPictureFormat.xlBitmap);

                chartObjects = (Excel.ChartObjects)worksheet.ChartObjects();
                chartObject = chartObjects.Add(left, top, width, height);
                chart = chartObject.Chart;
                chart.Paste();
                chart.Export(path, "PNG", false);
                if (!File.Exists(path))
                {
                    throw new OfficeComException(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "图片导出后文件不存在：" + path);
                }
                return new FileInfo(path).Length;
            }
            finally
            {
                try
                {
                    application.CutCopyMode = (Excel.XlCutCopyMode)0;
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[excel_export] 清除剪贴板状态失败：" + ex.Message);
                }
                if (chartObject != null)
                {
                    try
                    {
                        chartObject.Delete();
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[excel_export] 删除临时图表失败：" + ex.Message);
                    }
                }
                ComHelper.Release(chartObject, chart, chartObjects);
            }
        }

        private static OperationResult ExecuteImport(Excel.Application application, JObject arguments)
        {
            string format = ExcelArgumentReader.GetTrimmedString(arguments, "format", "type").ToLowerInvariant();
            if (format.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 format 参数（csv/json）。");
            }
            if (format != "csv" && format != "json")
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "format 仅支持 csv/json。");
            }

            string path = ExcelExtensionSupport.ResolveFilePath(arguments, true, false);
            bool createSheet = ExcelArgumentReader.GetBool(arguments, false, "createSheet", "autoCreateSheet");
            string sheetName = ExcelArgumentReader.GetSheetName(arguments);
            string targetAddress = ExcelArgumentReader.GetRangeAddress(
                arguments,
                "targetRange",
                "range",
                "startCell",
                "cell");
            if (string.IsNullOrWhiteSpace(targetAddress))
            {
                targetAddress = "A1";
            }

            int rows;
            int columns;
            object[,] matrix;
            string layout;
            if (format == "csv")
            {
                string encodingName = ExcelArgumentReader.GetString(arguments, "encoding");
                string text = ExcelExtensionSupport.ReadTextFile(path, encodingName);
                string delimiter = ExcelExtensionSupport.NormalizeDelimiter(
                    ExcelArgumentReader.GetString(arguments, "delimiter"));
                bool inferTypes = ExcelArgumentReader.GetBool(arguments, true, "inferTypes");
                matrix = ExcelExtensionSupport.ParseCsv(text, delimiter, inferTypes, out rows, out columns);
                layout = "csv";
            }
            else
            {
                string text = ExcelExtensionSupport.ReadTextFile(path, "auto");
                matrix = ExcelExtensionSupport.ParseJsonMatrix(text, out rows, out columns, out layout);
            }

            Excel.Workbook workbook = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                if (workbook.ReadOnly)
                {
                    return Fail(OfficeErrorCode.READ_ONLY, "工作簿以只读方式打开，无法导入数据。");
                }

                ExcelRangeContext context;
                bool useSheetName = sheetName.Length > 0 && targetAddress.IndexOf('!') < 0;
                if (useSheetName)
                {
                    context = RangeResolver.ResolveOnSheet(
                        application,
                        sheetName,
                        targetAddress,
                        createSheet,
                        "excel_import");
                }
                else
                {
                    context = RangeResolver.Resolve(
                        application,
                        targetAddress,
                        createSheet,
                        "excel_import");
                }

                using (context)
                {
                    ExcelComSupport.EnsureWritable(
                        application,
                        context.Workbook,
                        context.Worksheet,
                        "excel_import");

                    Excel.Range startCell = null;
                    Excel.Range destination = null;
                    try
                    {
                        startCell = (Excel.Range)context.Range.Cells[1, 1];
                        int startRow = startCell.Row;
                        int startColumn = startCell.Column;
                        ExcelExtensionSupport.EnsureRangeCapacity(startRow, startColumn, rows, columns);
                        destination = startCell.Resize[rows, columns];
                        string destinationAddress = destination.AddressLocal();

                        bool createdSheet = context.WorksheetCreated;
                        ExcelUndoScope undo = null;
                        if (!createdSheet)
                        {
                            undo = ExcelUndoScope.Capture(
                                context.Workbook,
                                context.Worksheet,
                                destination,
                                destinationAddress);
                        }

                        application.Guard(delegate { destination.Value2 = matrix; });

                        string targetRef = ExcelComSupport.FullAddress(
                            context.SheetName,
                            destinationAddress);
                        Observation observation = Observation.Create(
                            "import",
                            "已从 " + Path.GetFileName(path) + " 导入 " + rows + " 行 × " +
                            columns + " 列到 " + targetRef + "。",
                            true);
                        observation.AddTargetRef(targetRef);

                        JObject data = new JObject();
                        data["path"] = path;
                        data["format"] = format;
                        data["layout"] = layout;
                        data["sheet"] = context.SheetName;
                        data["address"] = destinationAddress;
                        data["rows"] = rows;
                        data["columns"] = columns;
                        data["createdSheet"] = createdSheet;

                        if (createdSheet)
                        {
                            ExcelUndoScope.RecordUnavailable(
                                data,
                                "导入目标工作表为新建工作表，区域快照无法撤销新增工作表本身，未登记自动撤销。");
                        }
                        else
                        {
                            undo.Complete(
                                context.Workbook,
                                context.Worksheet,
                                destination,
                                destinationAddress,
                                true,
                                "导入 " + Path.GetFileName(path) + " 到 " + targetRef,
                                observation,
                                data,
                                null);
                        }

                        return OperationResult.Succeed(
                            "已导入 " + rows + " 行 × " + columns + " 列到 " + targetRef + "。",
                            data,
                            observation);
                    }
                    finally
                    {
                        ComHelper.Release(startCell, destination);
                    }
                }
            }
            finally
            {
                ComHelper.Release(workbook);
            }
        }
    }
}
