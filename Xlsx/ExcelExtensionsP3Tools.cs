using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>P3 只读清单：隐藏工作表/行列、VBA 组件清单。</summary>
    public sealed partial class ExcelExtensionsExecutor
    {
        private static OperationResult ExecuteHiddenInventory(Excel.Application application, JObject arguments)
        {
            int maxRows = ClampInt(
                ExcelArgumentReader.GetInt(arguments, 2000, "maxRowsPerSheet"),
                1,
                100000);
            int maxColumns = ClampInt(
                ExcelArgumentReader.GetInt(arguments, 256, "maxColumnsPerSheet"),
                1,
                16384);

            Excel.Workbook workbook = null;
            List<Excel.Worksheet> worksheets = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                worksheets = ExcelExtensionSupport.CollectWorksheets(workbook, true, string.Empty);

                JArray hiddenSheets = new JArray();
                JArray sheetReports = new JArray();
                int hiddenRowCount = 0;
                int hiddenColumnCount = 0;
                bool truncated = false;

                foreach (Excel.Worksheet worksheet in worksheets)
                {
                    string visibility = ExcelComSupport.VisibilityText(worksheet.Visible);
                    if (!string.Equals(visibility, "visible", StringComparison.OrdinalIgnoreCase))
                    {
                        JObject hiddenSheet = new JObject();
                        hiddenSheet["name"] = worksheet.Name;
                        hiddenSheet["visibility"] = visibility;
                        hiddenSheets.Add(hiddenSheet);
                    }

                    Excel.Range usedRange = null;
                    try
                    {
                        usedRange = worksheet.UsedRange;
                        int lastRow = usedRange.Row + usedRange.Rows.Count - 1;
                        int lastColumn = usedRange.Column + usedRange.Columns.Count - 1;

                        JArray hiddenRows = new JArray();
                        JArray hiddenColumns = new JArray();
                        int rowLimit = Math.Min(lastRow, maxRows);
                        int columnLimit = Math.Min(lastColumn, maxColumns);
                        if (lastRow > rowLimit || lastColumn > columnLimit)
                        {
                            truncated = true;
                        }

                        Excel.Range rows = null;
                        Excel.Range columns = null;
                        try
                        {
                            rows = worksheet.Rows;
                            for (int row = 1; row <= rowLimit; row++)
                            {
                                Excel.Range rowRange = null;
                                try
                                {
                                    rowRange = (Excel.Range)rows[row];
                                    if (Convert.ToBoolean(rowRange.Hidden, CultureInfo.InvariantCulture))
                                    {
                                        hiddenRows.Add(row);
                                        hiddenRowCount++;
                                    }
                                }
                                finally
                                {
                                    ComHelper.Release(rowRange);
                                }
                            }

                            columns = worksheet.Columns;
                            for (int column = 1; column <= columnLimit; column++)
                            {
                                Excel.Range columnRange = null;
                                try
                                {
                                    columnRange = (Excel.Range)columns[column];
                                    if (Convert.ToBoolean(columnRange.Hidden, CultureInfo.InvariantCulture))
                                    {
                                        hiddenColumns.Add(ExcelComSupport.ColumnLetter(column));
                                        hiddenColumnCount++;
                                    }
                                }
                                finally
                                {
                                    ComHelper.Release(columnRange);
                                }
                            }
                        }
                        finally
                        {
                            ComHelper.Release(rows, columns);
                        }

                        JObject report = new JObject();
                        report["name"] = worksheet.Name;
                        report["visibility"] = visibility;
                        report["usedRange"] = usedRange.AddressLocal();
                        report["scannedRows"] = rowLimit;
                        report["scannedColumns"] = columnLimit;
                        report["hiddenRows"] = hiddenRows;
                        report["hiddenColumns"] = hiddenColumns;
                        report["hiddenRowCount"] = hiddenRows.Count;
                        report["hiddenColumnCount"] = hiddenColumns.Count;
                        sheetReports.Add(report);
                    }
                    finally
                    {
                        ComHelper.Release(usedRange);
                    }
                }

                JObject data = new JObject();
                data["workbook"] = workbook.Name;
                data["sheetCount"] = worksheets.Count;
                data["hiddenSheetCount"] = hiddenSheets.Count;
                data["hiddenSheets"] = hiddenSheets;
                data["hiddenRowCount"] = hiddenRowCount;
                data["hiddenColumnCount"] = hiddenColumnCount;
                data["truncated"] = truncated;
                data["sheets"] = sheetReports;

                string summary = "隐藏清单：隐藏工作表 " + hiddenSheets.Count + " 个，隐藏行 " +
                                 hiddenRowCount + " 行，隐藏列 " + hiddenColumnCount + " 列" +
                                 (truncated ? "（已按上限截断）。" : "。");
                Observation observation = ReadOnlyObservation("hidden_inventory", summary);
                return OperationResult.Succeed(summary, data, observation);
            }
            finally
            {
                ExcelExtensionSupport.ReleaseWorksheets(worksheets);
                ComHelper.Release(workbook);
            }
        }

        private static OperationResult ExecuteMacroList(Excel.Application application, JObject arguments)
        {
            Excel.Workbook workbook = null;
            object projectObject = null;
            object componentsObject = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                bool hasVba = workbook.HasVBProject;
                JArray components = new JArray();
                JObject data = new JObject();
                data["workbook"] = workbook.Name;
                data["hasVBProject"] = hasVba;

                if (!hasVba)
                {
                    data["componentCount"] = 0;
                    data["components"] = components;
                    string emptySummary = "工作簿不包含 VBA 工程。";
                    Observation emptyObservation = ReadOnlyObservation("macro_list", emptySummary);
                    return OperationResult.Succeed(emptySummary, data, emptyObservation);
                }

                dynamic dynamicWorkbook = workbook;
                projectObject = (object)dynamicWorkbook.VBProject;
                if (projectObject == null)
                {
                    return Fail(
                        OfficeErrorCode.VBA_DISABLED,
                        "无法访问 VBA 工程。请在 Excel 信任中心启用「信任对 VBA 工程对象模型的访问」后重试。");
                }

                dynamic project = projectObject;
                componentsObject = (object)project.VBComponents;
                dynamic componentCollection = componentsObject;
                int count = Convert.ToInt32(componentCollection.Count, CultureInfo.InvariantCulture);
                for (int index = 1; index <= count; index++)
                {
                    object componentObject = null;
                    try
                    {
                        componentObject = (object)componentCollection[index];
                        dynamic component = componentObject;
                        JObject item = new JObject();
                        item["name"] = Convert.ToString(component.Name, CultureInfo.InvariantCulture) ?? string.Empty;
                        int typeValue = Convert.ToInt32(component.Type, CultureInfo.InvariantCulture);
                        item["type"] = VbaComponentTypeText(typeValue);

                        object codeModuleObject = null;
                        try
                        {
                            codeModuleObject = (object)component.CodeModule;
                            if (codeModuleObject != null)
                            {
                                dynamic codeModule = codeModuleObject;
                                item["lineCount"] = Convert.ToInt32(
                                    codeModule.CountOfLines,
                                    CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                item["lineCount"] = 0;
                            }
                        }
                        finally
                        {
                            ComHelper.Release(codeModuleObject);
                        }
                        components.Add(item);
                    }
                    finally
                    {
                        ComHelper.Release(componentObject);
                    }
                }

                data["componentCount"] = components.Count;
                data["components"] = components;
                string summary = "VBA 工程清单：共 " + components.Count + " 个组件。";
                Observation observation = ReadOnlyObservation("macro_list", summary);
                return OperationResult.Succeed(summary, data, observation);
            }
            catch (Exception ex)
            {
                OfficeErrorCode code = OfficeErrorMapper.MapException(ex, OfficeApp.Excel);
                if (code == OfficeErrorCode.VBA_DISABLED)
                {
                    return Fail(
                        OfficeErrorCode.VBA_DISABLED,
                        "无法访问 VBA 工程：请在 Excel 信任中心启用「信任对 VBA 工程对象模型的访问」后重试。" +
                        "原始信息：" + ex.Message);
                }
                return Fail(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "读取 VBA 工程清单失败：" + ex.Message);
            }
            finally
            {
                ComHelper.Release(componentsObject, projectObject, workbook);
            }
        }

        private static string VbaComponentTypeText(int typeValue)
        {
            switch (typeValue)
            {
                case 1:
                    return "standardModule";
                case 2:
                    return "classModule";
                case 3:
                    return "userForm";
                case 100:
                    return "documentModule";
                default:
                    return "type" + typeValue.ToString(CultureInfo.InvariantCulture);
            }
        }
    }
}
