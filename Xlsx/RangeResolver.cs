using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>
    /// 解析后的 Excel 区域。持有 Workbook / Worksheet / Range 三个独立 RCW，
    /// Dispose 时逐个释放；Application 由 OfficeSession 持有，禁止在此释放。
    /// </summary>
    internal sealed class ExcelRangeContext : IDisposable
    {
        public Excel.Workbook Workbook { get; private set; }

        public Excel.Worksheet Worksheet { get; private set; }

        public Excel.Range Range { get; private set; }

        public string SheetName { get; private set; }

        public string LocalAddress { get; private set; }

        public string FullAddress { get; private set; }

        public bool WorksheetCreated { get; private set; }

        public ExcelRangeContext(
            Excel.Workbook workbook,
            Excel.Worksheet worksheet,
            Excel.Range range,
            string sheetName,
            string localAddress,
            bool worksheetCreated)
        {
            Workbook = workbook;
            Worksheet = worksheet;
            Range = range;
            SheetName = sheetName ?? string.Empty;
            LocalAddress = localAddress ?? string.Empty;
            FullAddress = ExcelComSupport.FullAddress(SheetName, LocalAddress);
            WorksheetCreated = worksheetCreated;
        }

        public void Dispose()
        {
            ComHelper.Release(Range, Worksheet, Workbook);
            Range = null;
            Worksheet = null;
            Workbook = null;
        }
    }

    /// <summary>
    /// 区域地址解析器：支持 Sheet!Range、{lastRow}/{lastCol} 占位符、
    /// selection/usedrange/currenttable 关键字、列字母与偏移，以及"表不存在时创建"开关。
    /// </summary>
    internal static class RangeResolver
    {
        public static ExcelRangeContext Resolve(
            Excel.Application application,
            string address,
            bool createSheet,
            string toolName)
        {
            if (application == null)
            {
                throw new OfficeComException(OfficeErrorCode.APP_NOT_RUNNING, "Excel 应用不可用。");
            }
            if (string.IsNullOrWhiteSpace(address))
            {
                throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "缺少目标范围地址。");
            }

            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Range range = null;
            bool created = false;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);

                string sheetName;
                string localAddress;
                SplitAddress(address.Trim(), out sheetName, out localAddress);

                if (string.IsNullOrWhiteSpace(localAddress))
                {
                    throw new OfficeComException(OfficeErrorCode.INVALID_RANGE, "范围地址缺少单元格部分：" + address);
                }

                if (sheetName.Length > 0)
                {
                    worksheet = ExcelComSupport.FindWorksheet(workbook, sheetName);
                    if (worksheet == null)
                    {
                        if (!createSheet)
                        {
                            throw new OfficeComException(
                                OfficeErrorCode.SHEET_NOT_FOUND,
                                "找不到工作表「" + sheetName + "」。如果希望自动创建，请传入 createSheet=true。");
                        }
                        worksheet = CreateWorksheet(workbook, sheetName);
                        created = true;
                    }
                }
                else
                {
                    object active = application.ActiveSheet;
                    worksheet = active as Excel.Worksheet;
                    if (worksheet == null)
                    {
                        ComHelper.Release(active);
                        throw new OfficeComException(OfficeErrorCode.DOC_NOT_OPEN, "当前活动对象不是工作表。");
                    }
                }

                localAddress = ResolveKeyword(application, worksheet, localAddress);
                localAddress = ResolvePlaceholders(worksheet, localAddress);

                try
                {
                    range = worksheet.Range[localAddress];
                }
                catch (COMException ex)
                {
                    throw new OfficeComException(
                        OfficeErrorCode.INVALID_RANGE,
                        "无法解析范围「" + address + "」：" + ex.Message,
                        ex);
                }

                string concreteAddress = range.AddressLocal();
                string actualSheetName = worksheet.Name;
                ExcelRangeContext context = new ExcelRangeContext(
                    workbook,
                    worksheet,
                    range,
                    actualSheetName,
                    concreteAddress,
                    created);
                workbook = null;
                worksheet = null;
                range = null;
                return context;
            }
            finally
            {
                ComHelper.Release(range, worksheet, workbook);
            }
        }

        public static ExcelRangeContext ResolveOnSheet(
            Excel.Application application,
            string sheetName,
            string localAddress,
            bool createSheet,
            string toolName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return Resolve(application, localAddress, createSheet, toolName);
            }
            if (string.IsNullOrWhiteSpace(localAddress))
            {
                throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "缺少目标单元格地址。");
            }
            if (localAddress.IndexOf('!') >= 0)
            {
                return Resolve(application, localAddress, createSheet, toolName);
            }
            return Resolve(
                application,
                ExcelComSupport.FullAddress(sheetName, localAddress),
                createSheet,
                toolName);
        }

        public static ExcelRangeContext ResolveTargetOrDefault(
            Excel.Application application,
            string explicitTarget,
            ExcelRangeContext source,
            int columnOffset,
            bool createSheet,
            string toolName)
        {
            if (!string.IsNullOrWhiteSpace(explicitTarget))
            {
                if (explicitTarget.IndexOf('!') >= 0)
                {
                    return Resolve(application, explicitTarget, createSheet, toolName);
                }
                return ResolveOnSheet(application, source.SheetName, explicitTarget, createSheet, toolName);
            }

            string defaultAddress = ExcelComSupport.OffsetAddress(source.Range, 0, columnOffset);
            return ResolveOnSheet(application, source.SheetName, defaultAddress, createSheet, toolName);
        }

        public static Excel.Worksheet ResolveWorksheet(
            Excel.Application application,
            Excel.Workbook workbook,
            string sheetName,
            bool createSheet)
        {
            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                Excel.Worksheet found = ExcelComSupport.FindWorksheet(workbook, sheetName);
                if (found == null)
                {
                    if (!createSheet)
                    {
                        throw new OfficeComException(OfficeErrorCode.SHEET_NOT_FOUND, "找不到工作表「" + sheetName + "」。");
                    }
                    found = CreateWorksheet(workbook, sheetName);
                }
                return found;
            }

            object active = application.ActiveSheet;
            Excel.Worksheet worksheet = active as Excel.Worksheet;
            if (worksheet == null)
            {
                ComHelper.Release(active);
                throw new OfficeComException(OfficeErrorCode.DOC_NOT_OPEN, "当前活动对象不是工作表。");
            }
            return worksheet;
        }

        public static Excel.Worksheet CreateWorksheet(Excel.Workbook workbook, string sheetName)
        {
            if (workbook == null)
            {
                throw new OfficeComException(OfficeErrorCode.DOC_NOT_OPEN, "工作簿不可用。");
            }
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "工作表名称不能为空。");
            }
            if (sheetName.Length > 31)
            {
                throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "工作表名称不能超过 31 个字符。");
            }

            Excel.Worksheet existing = ExcelComSupport.FindWorksheet(workbook, sheetName);
            if (existing != null)
            {
                ComHelper.Release(existing);
                throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "工作表「" + sheetName + "」已存在。");
            }

            Excel.Worksheet created = null;
            try
            {
                created = (Excel.Worksheet)workbook.Worksheets.Add();
                created.Name = sheetName;
                return created;
            }
            catch (Exception ex)
            {
                ComHelper.Release(created);
                throw new OfficeComException(
                    OfficeErrorCode.INVALID_PARAMS,
                    "创建工作表「" + sheetName + "」失败：" + ex.Message,
                    ex);
            }
        }

        public static void SplitAddress(string address, out string sheetName, out string localAddress)
        {
            sheetName = string.Empty;
            localAddress = address ?? string.Empty;
            int bang = localAddress.LastIndexOf('!');
            if (bang < 0)
            {
                return;
            }

            sheetName = localAddress.Substring(0, bang).Trim();
            localAddress = localAddress.Substring(bang + 1).Trim();
            if (sheetName.Length >= 2 && sheetName.StartsWith("'", StringComparison.Ordinal) &&
                sheetName.EndsWith("'", StringComparison.Ordinal))
            {
                sheetName = sheetName.Substring(1, sheetName.Length - 2).Replace("''", "'");
            }
        }

        public static string[] GetHeaders(Excel.Worksheet worksheet, Excel.Range range, int maxColumns)
        {
            if (range == null)
            {
                return new string[0];
            }
            int columns = Math.Min(range.Columns.Count, Math.Max(1, maxColumns));
            Excel.Range headerRange = null;
            try
            {
                headerRange = range.Resize[1, columns];
                object[,] values = ExcelComSupport.ReadMatrix(headerRange, 1, columns);
                string[] headers = new string[columns];
                for (int column = 0; column < columns; column++)
                {
                    string header = Convert.ToString(values[0, column], CultureInfo.InvariantCulture) ?? string.Empty;
                    headers[column] = header.Trim();
                }
                return headers;
            }
            finally
            {
                ComHelper.Release(headerRange);
            }
        }

        public static void GetUsedBounds(
            Excel.Worksheet worksheet,
            out int firstRow,
            out int firstColumn,
            out int lastRow,
            out int lastColumn)
        {
            firstRow = 1;
            firstColumn = 1;
            lastRow = 1;
            lastColumn = 1;
            Excel.Range used = null;
            try
            {
                used = worksheet.UsedRange;
                firstRow = used.Row;
                firstColumn = used.Column;
                lastRow = used.Row + used.Rows.Count - 1;
                lastColumn = used.Column + used.Columns.Count - 1;
            }
            finally
            {
                ComHelper.Release(used);
            }
        }

        public static string CurrentRegionAddress(Excel.Application application)
        {
            Excel.Range activeCell = null;
            Excel.Range region = null;
            try
            {
                activeCell = application.ActiveCell;
                if (activeCell == null)
                {
                    throw new OfficeComException(OfficeErrorCode.INVALID_RANGE, "当前活动工作簿没有活动单元格。");
                }
                region = activeCell.CurrentRegion;
                return region.AddressLocal();
            }
            finally
            {
                ComHelper.Release(region, activeCell);
            }
        }

        private static string ResolveKeyword(Excel.Application application, Excel.Worksheet worksheet, string address)
        {
            string key = address.Trim().Trim('{', '}').ToLowerInvariant();
            if (key == "usedrange" || key == "used")
            {
                Excel.Range used = null;
                try
                {
                    used = worksheet.UsedRange;
                    return used.AddressLocal();
                }
                finally
                {
                    ComHelper.Release(used);
                }
            }
            if (key == "currenttable" || key == "currentregion")
            {
                object active = application.ActiveSheet;
                Excel.Worksheet activeWorksheet = active as Excel.Worksheet;
                bool isActive = activeWorksheet != null &&
                                string.Equals(activeWorksheet.Name, worksheet.Name, StringComparison.Ordinal);
                ComHelper.Release(activeWorksheet, active);
                if (isActive)
                {
                    return CurrentRegionAddress(application);
                }
                Excel.Range used = null;
                try
                {
                    used = worksheet.UsedRange;
                    return used.AddressLocal();
                }
                finally
                {
                    ComHelper.Release(used);
                }
            }
            if (key == "selection" || key == "selectedrange" || key == "currentselection")
            {
                object selection = application.Selection;
                Excel.Range selectedRange = selection as Excel.Range;
                if (selectedRange == null)
                {
                    ComHelper.Release(selection);
                    throw new OfficeComException(OfficeErrorCode.INVALID_RANGE, "当前选择不是单元格区域。");
                }
                try
                {
                    Excel.Worksheet selectedSheet = selectedRange.Worksheet as Excel.Worksheet;
                    try
                    {
                        if (selectedSheet == null || !string.Equals(selectedSheet.Name, worksheet.Name, StringComparison.Ordinal))
                        {
                            throw new OfficeComException(
                                OfficeErrorCode.INVALID_RANGE,
                                "当前选择不在目标工作表「" + worksheet.Name + "」上。");
                        }
                    }
                    finally
                    {
                        ComHelper.Release(selectedSheet);
                    }
                    return selectedRange.AddressLocal();
                }
                finally
                {
                    ComHelper.Release(selectedRange);
                }
            }
            return address;
        }

        private static string ResolvePlaceholders(Excel.Worksheet worksheet, string address)
        {
            if (address.IndexOf('{') < 0)
            {
                return address;
            }

            int firstRow;
            int firstColumn;
            int lastRow;
            int lastColumn;
            GetUsedBounds(worksheet, out firstRow, out firstColumn, out lastRow, out lastColumn);

            string result = Regex.Replace(
                address,
                @"\{lastrow\}",
                lastRow.ToString(CultureInfo.InvariantCulture),
                RegexOptions.IgnoreCase);
            result = Regex.Replace(
                result,
                @"\{lastcol(?:umn)?\}",
                lastColumn.ToString(CultureInfo.InvariantCulture),
                RegexOptions.IgnoreCase);
            result = Regex.Replace(
                result,
                @"\{firstrow\}",
                firstRow.ToString(CultureInfo.InvariantCulture),
                RegexOptions.IgnoreCase);
            result = Regex.Replace(
                result,
                @"\{firstcol(?:umn)?\}",
                firstColumn.ToString(CultureInfo.InvariantCulture),
                RegexOptions.IgnoreCase);
            return result;
        }
    }
}
