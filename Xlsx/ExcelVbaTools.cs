using System;
using System.Globalization;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>
    /// VBA 执行入口。VBE 类型由 Excel PIA 引用 Microsoft.Vbe.Interop，
    /// 但当前 NuGet PIA 未携带该程序集，因此这里用 late binding 访问 VBE，
    /// 避免为一个可选高危工具引入无法还原的 COM 依赖。
    /// </summary>
    public sealed partial class ExcelToolExecutor
    {
        private const int VbextCtStdModule = 1;

        private static OperationResult ExecuteRunVba(Excel.Application application, JObject arguments)
        {
            if (!ExcelArgumentReader.GetBool(arguments, false, "confirm"))
            {
                return OperationResult.Cancelled(
                    "excel_run_vba 需要显式 confirm=true 才会执行；本次未执行任何 VBA 代码。");
            }

            OperationResult denied = SecurityGate.CreateDenial(SecurityRisk.Vba, "excel_run_vba");
            if (denied != null)
            {
                return denied;
            }

            string code = ExcelArgumentReader.GetString(arguments, "code", "vbaCode");
            if (string.IsNullOrWhiteSpace(code))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 code/vbaCode 参数。");
            }

            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Range usedRange = null;
            Excel.Range afterRange = null;
            object vbeObject = null;
            object projectObject = null;
            object componentsObject = null;
            object componentObject = null;
            ExcelUndoScope undo = null;
            string moduleName = "McpTemp_" + DateTime.Now.ToString("HHmmssfff", CultureInfo.InvariantCulture);
            bool moduleAdded = false;
            int sheetsBefore = 0;
            int sheetsAfter = 0;
            int changedCells = 0;
            bool shapeChanged = false;
            object runResult = null;

            try
            {
                try
                {
                    workbook = ExcelComSupport.RequireActiveWorkbook(application);
                    worksheet = ExcelComSupport.RequireActiveWorksheet(application);
                    usedRange = worksheet.UsedRange;
                    sheetsBefore = workbook.Worksheets.Count;
                    string usedAddress = usedRange.AddressLocal();
                    undo = ExcelUndoScope.Capture(workbook, worksheet, usedRange, usedAddress);

                    dynamic dynamicApplication = application;
                    vbeObject = dynamicApplication.VBE;
                    if (vbeObject == null)
                    {
                        throw new OfficeComException(
                            OfficeErrorCode.VBA_DISABLED,
                            "无法访问 Excel VBE。请在信任中心启用“信任对 VBA 工程对象模型的访问”。");
                    }
                    dynamic vbe = vbeObject;
                    projectObject = vbe.ActiveVBProject;
                    if (projectObject == null)
                    {
                        throw new OfficeComException(
                            OfficeErrorCode.VBA_DISABLED,
                            "当前工作簿没有可用的 VBA 工程。");
                    }

                    dynamic project = projectObject;
                    componentsObject = project.VBComponents;
                    dynamic components = componentsObject;
                    componentObject = components.Add(VbextCtStdModule);
                    moduleAdded = true;
                    dynamic component = componentObject;
                    component.Name = moduleName;

                    string procedureName;
                    string moduleCode = NormalizeVbaCode(code, out procedureName);
                    object codeModuleObject = component.CodeModule;
                    try
                    {
                        dynamic codeModule = codeModuleObject;
                        codeModule.AddFromString(moduleCode);
                    }
                    finally
                    {
                        ComHelper.Release(codeModuleObject);
                    }

                    application.Guard(delegate
                    {
                        dynamic app = application;
                        runResult = app.Run(moduleName + "." + procedureName);
                    });

                    sheetsAfter = workbook.Worksheets.Count;
                    changedCells = CountChangedCells(
                        worksheet,
                        usedRange,
                        sheetsBefore == sheetsAfter,
                        out shapeChanged);
                }
                catch (Exception ex)
                {
                    OfficeErrorCode errorCode = OfficeErrorMapper.MapException(ex, OfficeApp.Excel);
                    if (errorCode == OfficeErrorCode.VBA_DISABLED)
                    {
                        return OperationResult.Failed(
                            OfficeErrorCode.VBA_DISABLED,
                            "VBA 执行被拒绝：security.allowVba 已放行，但 Excel 拒绝访问 VBA 工程。" +
                            "请在“文件 > 选项 > 信任中心 > 信任中心设置 > 宏设置”中启用" +
                            "“信任对 VBA 工程对象模型的访问”，然后重启 Excel。原始信息：" + ex.Message);
                    }
                    return OperationResult.Failed(
                        errorCode,
                        "VBA 执行失败：" + OfficeErrorMapper.DescribeException(ex, OfficeApp.Excel));
                }

                sheetsAfter = workbook.Worksheets.Count;
                afterRange = worksheet.UsedRange;
                string afterAddress = afterRange.AddressLocal();
                if (undo == null)
                {
                    undo = ExcelUndoScope.Failed("VBA 执行前快照不可用，未登记可撤销节点。");
                }

                Observation observation = Observation.Create(
                    "run_vba",
                    "VBA 已执行；可观测到 " + changedCells + " 个单元格变化，工作表数 " + sheetsBefore + " -> " + sheetsAfter + "。",
                    changedCells > 0 || shapeChanged || sheetsAfter != sheetsBefore);
                observation.AddTargetRef(ExcelComSupport.FullAddress(worksheet.Name, afterAddress));

                JObject data = new JObject();
                data["result"] = ReadVbaResult(runResult);
                data["changedCells"] = changedCells;
                data["shapeChanged"] = shapeChanged;
                data["sheetsBefore"] = sheetsBefore;
                data["sheetsAfter"] = sheetsAfter;
                data["moduleName"] = moduleName;
                data["changeDetectionLimited"] = true;
                data["changeDetectionNote"] = "仅比较活动工作表 UsedRange 的值与工作表数量；其他工作表、格式、名称和 VBA 工程副作用无法完整统计。";

                undo.Complete(
                    workbook,
                    worksheet,
                    afterRange,
                    afterAddress,
                    false,
                    "执行 VBA " + moduleName,
                    observation,
                    data,
                    "VBA 可修改任意工作表/名称/工程，区域快照不能完整还原，未登记自动撤销。");

                return OperationResult.Succeed(
                    "VBA 执行成功；可观测变化：" + changedCells + " 个单元格。",
                    data,
                    observation);
            }
            finally
            {
                if (moduleAdded && componentsObject != null && componentObject != null)
                {
                    try
                    {
                        dynamic components = componentsObject;
                        components.Remove(componentObject);
                    }
                    catch (Exception cleanupError)
                    {
                        Console.Error.WriteLine("[excel] 删除 VBA 临时模块失败：" + cleanupError.Message);
                    }
                }
                ComHelper.Release(componentObject, componentsObject, projectObject, vbeObject);
                ComHelper.Release(afterRange, usedRange, worksheet, workbook);
            }
        }

        private static string NormalizeVbaCode(string code, out string procedureName)
        {
            string normalized = (code ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace("\r", "\n")
                .Replace("\n", "\r\n");
            Match match = Regex.Match(
                normalized,
                @"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)",
                RegexOptions.Multiline | RegexOptions.IgnoreCase);
            if (match.Success)
            {
                procedureName = match.Groups[1].Value;
                return normalized;
            }
            procedureName = "Auto_Run";
            return "Sub Auto_Run()\r\n" + normalized + "\r\nEnd Sub";
        }

        private static int CountChangedCells(
            Excel.Worksheet worksheet,
            Excel.Range previousRange,
            bool sameSheetCount,
            out bool shapeChanged)
        {
            shapeChanged = false;
            if (worksheet == null)
            {
                return 0;
            }
            Excel.Range current = null;
            try
            {
                current = worksheet.UsedRange;
                int oldRows = previousRange.Rows.Count;
                int oldColumns = previousRange.Columns.Count;
                int newRows = current.Rows.Count;
                int newColumns = current.Columns.Count;
                if (oldRows != newRows || oldColumns != newColumns || !sameSheetCount)
                {
                    shapeChanged = true;
                }

                int rows = Math.Min(oldRows, newRows);
                int columns = Math.Min(oldColumns, newColumns);
                if (rows <= 0 || columns <= 0)
                {
                    return 0;
                }
                object[,] before = ExcelComSupport.ReadMatrix(previousRange, oldRows, oldColumns);
                object[,] after = ExcelComSupport.ReadMatrix(current, newRows, newColumns);
                int changed = 0;
                for (int row = 0; row < rows; row++)
                {
                    for (int column = 0; column < columns; column++)
                    {
                        string left = Convert.ToString(before[row, column], CultureInfo.InvariantCulture) ?? string.Empty;
                        string right = Convert.ToString(after[row, column], CultureInfo.InvariantCulture) ?? string.Empty;
                        if (!string.Equals(left, right, StringComparison.Ordinal))
                        {
                            changed++;
                        }
                    }
                }
                return changed;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[excel] VBA 变更计数失败：" + ex.Message);
                return 0;
            }
            finally
            {
                ComHelper.Release(current);
            }
        }

        private static JToken ReadVbaResult(object value)
        {
            if (value == null)
            {
                return JValue.CreateNull();
            }
            Excel.Range range = value as Excel.Range;
            if (range != null)
            {
                try
                {
                    object raw = range.Value2;
                    return ExcelComSupport.ToJToken(raw);
                }
                finally
                {
                    ComHelper.Release(range);
                }
            }
            return ExcelComSupport.ToJToken(value);
        }
    }
}
