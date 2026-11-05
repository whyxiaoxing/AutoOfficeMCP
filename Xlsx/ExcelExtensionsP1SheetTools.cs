using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>P1 工作表层扩展：多表纵向汇总与目录页。</summary>
    public sealed partial class ExcelExtensionsExecutor
    {
        private sealed class ConsolidateBlock
        {
            public string SheetName { get; set; }

            public object[,] Values { get; set; }

            public int Rows { get; set; }

            public int Columns { get; set; }
        }

        private static OperationResult ExecuteCrossSheetConsolidate(Excel.Application application, JObject arguments)
        {
            string targetSheetName = ExcelArgumentReader.GetTrimmedString(
                arguments,
                "targetSheet",
                "outputSheet");
            if (targetSheetName.Length == 0)
            {
                targetSheetName = "汇总";
            }
            string outputAddress = ExcelArgumentReader.GetTrimmedString(
                arguments,
                "outputAddress",
                "targetCell",
                "startCell");
            if (outputAddress.Length == 0)
            {
                outputAddress = "A1";
            }
            string rangeToken = ExcelArgumentReader.GetTrimmedString(arguments, "range", "sourceRange");
            bool hasHeader = ExcelArgumentReader.GetBool(arguments, true, "hasHeader", "header");
            bool includeSourceColumn = ExcelArgumentReader.GetBool(
                arguments,
                false,
                "includeSourceColumn",
                "addSourceColumn");
            string sourceHeader = ExcelArgumentReader.GetString(arguments, "sourceColumnHeader");
            if (sourceHeader.Length == 0)
            {
                sourceHeader = "来源工作表";
            }
            bool includeHidden = ExcelArgumentReader.GetBool(arguments, false, "includeHidden");
            List<string> requestedSheets = ExcelArgumentReader.GetStringList(arguments, "sheets");

            Excel.Workbook workbook = null;
            List<Excel.Worksheet> worksheets = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                worksheets = ExcelExtensionSupport.CollectWorksheets(
                    workbook,
                    includeHidden,
                    targetSheetName);
                FilterWorksheetsByName(worksheets, requestedSheets);
                if (worksheets.Count == 0)
                {
                    return Fail(OfficeErrorCode.SHEET_NOT_FOUND, "没有找到可汇总的工作表。");
                }

                List<ConsolidateBlock> blocks = new List<ConsolidateBlock>();
                List<string> skipped = new List<string>();
                List<string> mismatched = new List<string>();
                int expectedColumns = -1;

                foreach (Excel.Worksheet worksheet in worksheets)
                {
                    Excel.Range sourceRange = null;
                    ExcelRangeContext explicitContext = null;
                    try
                    {
                        if (string.IsNullOrWhiteSpace(rangeToken))
                        {
                            sourceRange = worksheet.UsedRange;
                        }
                        else
                        {
                            string sheetPart;
                            string localAddress;
                            RangeResolver.SplitAddress(rangeToken, out sheetPart, out localAddress);
                            if (!string.IsNullOrWhiteSpace(sheetPart) &&
                                !string.Equals(sheetPart, worksheet.Name, StringComparison.OrdinalIgnoreCase))
                            {
                                mismatched.Add(worksheet.Name + "：range 指向了其他工作表「" + sheetPart + "」");
                                continue;
                            }
                            explicitContext = RangeResolver.ResolveOnSheet(
                                application,
                                worksheet.Name,
                                localAddress,
                                false,
                                "excel_cross_sheet_consolidate");
                            sourceRange = explicitContext.Range;
                        }

                        int rows = sourceRange.Rows.Count;
                        int columns = sourceRange.Columns.Count;
                        if (rows < 1 || columns < 1)
                        {
                            skipped.Add(worksheet.Name);
                            continue;
                        }

                        object[,] values = ExcelExtensionSupport.ReadMatrix(sourceRange, rows, columns);
                        if (!ExcelExtensionSupport.MatrixHasData(values))
                        {
                            skipped.Add(worksheet.Name);
                            continue;
                        }
                        if (expectedColumns < 0)
                        {
                            expectedColumns = columns;
                        }
                        else if (columns != expectedColumns)
                        {
                            mismatched.Add(
                                worksheet.Name + "：列数为 " + columns + "，期望 " + expectedColumns);
                            continue;
                        }

                        ConsolidateBlock block = new ConsolidateBlock();
                        block.SheetName = worksheet.Name;
                        block.Values = values;
                        block.Rows = rows;
                        block.Columns = columns;
                        blocks.Add(block);
                    }
                    finally
                    {
                        if (explicitContext != null)
                        {
                            explicitContext.Dispose();
                        }
                        else
                        {
                            ComHelper.Release(sourceRange);
                        }
                    }
                }

                if (mismatched.Count > 0)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "以下工作表结构不一致，未执行汇总：" + string.Join("；", mismatched.ToArray()));
                }
                if (blocks.Count == 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "没有可汇总的数据（所有工作表为空）。");
                }

                int dataRows = 0;
                int outputColumns = expectedColumns + (includeSourceColumn ? 1 : 0);
                foreach (ConsolidateBlock block in blocks)
                {
                    int blockDataRows = hasHeader ? block.Rows - 1 : block.Rows;
                    if (blockDataRows > 0)
                    {
                        dataRows += blockDataRows;
                    }
                }
                if (dataRows == 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "没有可汇总的数据行（来源只包含表头）。");
                }

                int outputRows = dataRows + (hasHeader ? 1 : 0);
                object[,] output = new object[outputRows, outputColumns];
                int outputRow = 0;
                if (hasHeader)
                {
                    int columnOffset = 0;
                    if (includeSourceColumn)
                    {
                        output[0, 0] = sourceHeader;
                        columnOffset = 1;
                    }
                    ConsolidateBlock first = blocks[0];
                    for (int column = 0; column < first.Columns; column++)
                    {
                        output[0, column + columnOffset] = first.Values[0, column];
                    }
                    outputRow = 1;
                }

                JArray sheetSummary = new JArray();
                foreach (ConsolidateBlock block in blocks)
                {
                    int startRow = hasHeader ? 1 : 0;
                    int blockDataRows = 0;
                    for (int row = startRow; row < block.Rows; row++)
                    {
                        int columnOffset = 0;
                        if (includeSourceColumn)
                        {
                            output[outputRow, 0] = block.SheetName;
                            columnOffset = 1;
                        }
                        for (int column = 0; column < block.Columns; column++)
                        {
                            output[outputRow, column + columnOffset] = block.Values[row, column];
                        }
                        outputRow++;
                        blockDataRows++;
                    }

                    JObject summary = new JObject();
                    summary["sheet"] = block.SheetName;
                    summary["dataRows"] = blockDataRows;
                    summary["columns"] = block.Columns;
                    sheetSummary.Add(summary);
                }

                Excel.Worksheet target = null;
                Excel.Range beforeRange = null;
                Excel.Range startCell = null;
                Excel.Range destination = null;
                try
                {
                    if (workbook.ReadOnly)
                    {
                        return Fail(OfficeErrorCode.READ_ONLY, "工作簿以只读方式打开，无法写入汇总表。");
                    }

                    bool created = false;
                    target = ExcelExtensionSupport.FindWorksheet(workbook, targetSheetName);
                    ExcelUndoScope undo = null;
                    if (target == null)
                    {
                        target = RangeResolver.CreateWorksheet(workbook, targetSheetName);
                        created = true;
                    }
                    else
                    {
                        ExcelComSupport.EnsureWritable(
                            application,
                            workbook,
                            target,
                            "excel_cross_sheet_consolidate");
                        beforeRange = target.UsedRange;
                        string beforeAddress = beforeRange.AddressLocal();
                        undo = ExcelUndoScope.Capture(workbook, target, beforeRange, beforeAddress);
                        Excel.Range clearRange = beforeRange;
                        application.Guard(delegate { clearRange.ClearContents(); });
                    }

                    string sheetPart;
                    string localAddress;
                    RangeResolver.SplitAddress(outputAddress, out sheetPart, out localAddress);
                    if (!string.IsNullOrWhiteSpace(sheetPart) &&
                        !string.Equals(sheetPart, target.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "outputAddress 指向了其他工作表「" + sheetPart + "」，请改为本地地址。");
                    }
                    if (string.IsNullOrWhiteSpace(localAddress))
                    {
                        localAddress = "A1";
                    }

                    startCell = (Excel.Range)target.Range[localAddress];
                    int startRow = startCell.Row;
                    int startColumn = startCell.Column;
                    ExcelExtensionSupport.EnsureRangeCapacity(startRow, startColumn, outputRows, outputColumns);
                    destination = startCell.Resize[outputRows, outputColumns];
                    string outputAddressLocal = destination.AddressLocal();

                    application.Guard(delegate { destination.Value2 = output; });

                    string targetRef = ExcelComSupport.FullAddress(target.Name, outputAddressLocal);
                    Observation observation = Observation.Create(
                        "cross_sheet_consolidate",
                        "已汇总 " + blocks.Count + " 个工作表到 " + targetRef +
                        "（" + dataRows + " 行数据）。",
                        true);
                    observation.AddTargetRef(targetRef);

                    JObject data = new JObject();
                    data["targetSheet"] = target.Name;
                    data["outputAddress"] = outputAddressLocal;
                    data["sheetCount"] = blocks.Count;
                    data["dataRows"] = dataRows;
                    data["outputRows"] = outputRows;
                    data["outputColumns"] = outputColumns;
                    data["hasHeader"] = hasHeader;
                    data["includeSourceColumn"] = includeSourceColumn;
                    data["sheets"] = sheetSummary;
                    data["skippedSheets"] = new JArray(skipped.ToArray());

                    if (created)
                    {
                        ExcelUndoScope.RecordUnavailable(
                            data,
                            "汇总目标工作表为新建工作表，属于结构变更，区域快照无法撤销新增工作表。");
                    }
                    else
                    {
                        undo.Complete(
                            workbook,
                            target,
                            destination,
                            outputAddressLocal,
                            true,
                            "多表汇总 " + targetRef,
                            observation,
                            data,
                            null);
                    }

                    return OperationResult.Succeed(
                        "已汇总 " + blocks.Count + " 个工作表共 " + dataRows + " 行数据到 " + targetRef + "。",
                        data,
                        observation);
                }
                finally
                {
                    ComHelper.Release(startCell, destination, beforeRange, target);
                }
            }
            finally
            {
                ExcelExtensionSupport.ReleaseWorksheets(worksheets);
                ComHelper.Release(workbook);
            }
        }

        private static OperationResult ExecuteSheetToc(Excel.Application application, JObject arguments)
        {
            string targetSheetName = ExcelArgumentReader.GetTrimmedString(
                arguments,
                "targetSheet",
                "outputSheet",
                "tocSheet");
            if (targetSheetName.Length == 0)
            {
                targetSheetName = "目录";
            }
            bool includeHidden = ExcelArgumentReader.GetBool(arguments, false, "includeHidden");
            bool includeVeryHidden = ExcelArgumentReader.GetBool(arguments, true, "includeVeryHidden");
            bool overwrite = ExcelArgumentReader.GetBool(arguments, true, "overwrite");
            string position = ExcelArgumentReader.GetTrimmedString(arguments, "position")
                .ToLowerInvariant();
            if (position.Length == 0)
            {
                position = "first";
            }

            Excel.Workbook workbook = null;
            List<Excel.Worksheet> allVisible = null;
            List<Excel.Worksheet> tocSheets = null;
            Excel.Worksheet target = null;
            Excel.Range beforeRange = null;
            Excel.Range startCell = null;
            Excel.Range destination = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                if (workbook.ReadOnly)
                {
                    return Fail(OfficeErrorCode.READ_ONLY, "工作簿以只读方式打开，无法生成目录页。");
                }

                target = ExcelExtensionSupport.FindWorksheet(workbook, targetSheetName);
                bool created = false;
                bool existingTarget = target != null;
                if (existingTarget && !overwrite)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "工作表「" + targetSheetName + "」已存在。如需重建目录请传 overwrite=true。");
                }

                allVisible = ExcelExtensionSupport.CollectWorksheets(workbook, true, targetSheetName);
                tocSheets = new List<Excel.Worksheet>();
                foreach (Excel.Worksheet worksheet in allVisible)
                {
                    bool include = worksheet.Visible == Excel.XlSheetVisibility.xlSheetVisible ||
                                   (worksheet.Visible == Excel.XlSheetVisibility.xlSheetHidden && includeHidden) ||
                                   (worksheet.Visible == Excel.XlSheetVisibility.xlSheetVeryHidden &&
                                    includeHidden && includeVeryHidden);
                    if (include)
                    {
                        tocSheets.Add(worksheet);
                    }
                    else
                    {
                        ComHelper.Release(worksheet);
                    }
                }
                allVisible.Clear();

                ExcelUndoScope undo = null;
                if (existingTarget)
                {
                    ExcelComSupport.EnsureWritable(application, workbook, target, "excel_sheet_toc");
                    beforeRange = target.UsedRange;
                    string beforeAddress = beforeRange.AddressLocal();
                    undo = ExcelUndoScope.Capture(workbook, target, beforeRange, beforeAddress);
                    Excel.Range clearRange = beforeRange;
                    application.Guard(delegate
                    {
                        clearRange.ClearContents();
                        clearRange.ClearHyperlinks();
                    });
                }
                else
                {
                    target = CreateTocWorksheet(workbook, targetSheetName, position);
                    created = true;
                }

                int rows = Math.Max(1, tocSheets.Count + 1);
                const int ColumnCount = 3;
                object[,] matrix = new object[rows, ColumnCount];
                matrix[0, 0] = "序号";
                matrix[0, 1] = "工作表名称";
                matrix[0, 2] = "跳转";
                for (int index = 0; index < tocSheets.Count; index++)
                {
                    Excel.Worksheet sheet = tocSheets[index];
                    matrix[index + 1, 0] = index + 1;
                    matrix[index + 1, 1] = sheet.Name;
                    matrix[index + 1, 2] = "打开";
                }

                startCell = (Excel.Range)target.Cells[1, 1];
                destination = startCell.Resize[rows, ColumnCount];
                string destinationAddress = destination.AddressLocal();
                application.Guard(delegate
                {
                    destination.Value2 = matrix;
                    destination.NumberFormat = "@";
                });

                int hyperlinkFailures = 0;
                for (int index = 0; index < tocSheets.Count; index++)
                {
                    Excel.Worksheet sheet = tocSheets[index];
                    Excel.Range anchor = null;
                    Excel.Hyperlink hyperlink = null;
                    try
                    {
                        anchor = (Excel.Range)target.Cells[index + 2, 2];
                        hyperlink = (Excel.Hyperlink)target.Hyperlinks.Add(
                            anchor,
                            string.Empty,
                            ExcelExtensionSupport.SheetInternalReference(sheet.Name, "A1"),
                            "跳转到工作表 " + sheet.Name,
                            sheet.Name);
                    }
                    catch (Exception ex)
                    {
                        hyperlinkFailures++;
                        Console.Error.WriteLine(
                            "[excel_sheet_toc] 为「" + sheet.Name + "」添加超链接失败：" + ex.Message);
                    }
                    finally
                    {
                        ComHelper.Release(hyperlink, anchor);
                    }
                }

                ApplyTocFormatting(target, rows);

                string targetRef = ExcelComSupport.FullAddress(target.Name, destinationAddress);
                Observation observation = Observation.Create(
                    "sheet_toc",
                    "已生成目录页 " + targetRef + "，收录 " + tocSheets.Count + " 个工作表。",
                    true);
                observation.AddTargetRef(targetRef);

                JObject data = new JObject();
                data["targetSheet"] = target.Name;
                data["address"] = destinationAddress;
                data["sheetCount"] = tocSheets.Count;
                data["createdSheet"] = created;
                data["hyperlinkFailures"] = hyperlinkFailures;
                JArray names = new JArray();
                foreach (Excel.Worksheet sheet in tocSheets)
                {
                    names.Add(sheet.Name);
                }
                data["sheets"] = names;

                if (created)
                {
                    ExcelUndoScope.RecordUnavailable(
                        data,
                        "目录工作表为新建工作表，属于结构变更，区域快照无法撤销新增工作表；超链接与格式也不在快照中。");
                }
                else
                {
                    undo.Complete(
                        workbook,
                        target,
                        destination,
                        destinationAddress,
                        true,
                        "生成目录页 " + targetRef,
                        observation,
                        data,
                        "超链接、列宽与表头格式不在区域快照中，自动撤销只恢复单元格值/公式/数字格式。");
                }

                return OperationResult.Succeed(
                    "已生成目录页：" + targetRef + "（" + tocSheets.Count + " 个工作表）。",
                    data,
                    observation);
            }
            finally
            {
                ComHelper.Release(startCell, destination, beforeRange);
                ExcelExtensionSupport.ReleaseWorksheets(tocSheets);
                ExcelExtensionSupport.ReleaseWorksheets(allVisible);
                ComHelper.Release(target, workbook);
            }
        }

        private static Excel.Worksheet CreateTocWorksheet(
            Excel.Workbook workbook,
            string sheetName,
            string position)
        {
            Excel.Sheets sheets = null;
            Excel.Worksheet reference = null;
            Excel.Worksheet created = null;
            try
            {
                sheets = workbook.Worksheets;
                int count = sheets.Count;
                object item = sheets[position == "last" || position == "end" || position == "after" ? count : 1];
                reference = item as Excel.Worksheet;
                if (reference == null)
                {
                    ComHelper.Release(item);
                    throw new OfficeComException(OfficeErrorCode.INTERNAL_ERROR, "无法取得目录页定位工作表。");
                }
                created = (Excel.Worksheet)(position == "last" || position == "end" || position == "after"
                    ? workbook.Worksheets.Add(After: reference)
                    : workbook.Worksheets.Add(Before: reference));
                created.Name = sheetName;
                return created;
            }
            finally
            {
                ComHelper.Release(reference);
                ComHelper.Release(sheets);
            }
        }

        private static void ApplyTocFormatting(Excel.Worksheet target, int rows)
        {
            Excel.Range header = null;
            Excel.Range columnA = null;
            Excel.Range columnB = null;
            Excel.Range columnC = null;
            Excel.Font font = null;
            try
            {
                header = target.Range["A1:C1"];
                font = header.Font;
                font.Bold = true;
                columnA = target.Range["A:A"];
                columnB = target.Range["B:B"];
                columnC = target.Range["C:C"];
                columnA.ColumnWidth = 7;
                columnB.ColumnWidth = 30;
                columnC.ColumnWidth = 10;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[excel_sheet_toc] 目录格式设置失败：" + ex.Message);
            }
            finally
            {
                ComHelper.Release(font, header, columnA, columnB, columnC);
            }
        }
    }
}
