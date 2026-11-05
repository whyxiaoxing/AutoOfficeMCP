using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>工作表、行列与保护操作。</summary>
    public sealed partial class ExcelToolExecutor
    {
        private static OperationResult ExecuteCreateSheet(Excel.Application application, JObject arguments)
        {
            string name = ExcelArgumentReader.GetTrimmedString(arguments, "name", "sheetName", "newName");
            string error = ValidateSheetName(name);
            if (error.Length > 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, error);
            }

            Excel.Workbook workbook = null;
            Excel.Worksheet created = null;
            Excel.Worksheet reference = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                Excel.Worksheet existing = ExcelComSupport.FindWorksheet(workbook, name);
                if (existing != null)
                {
                    ComHelper.Release(existing);
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "工作表「" + name + "」已存在。");
                }

                string referenceName = ExcelArgumentReader.GetTrimmedString(arguments, "referenceSheet", "afterSheet", "beforeSheet");
                string position = ExcelArgumentReader.GetString(arguments, "position").Trim().ToLowerInvariant();
                if (!string.IsNullOrWhiteSpace(referenceName))
                {
                    reference = ExcelComSupport.FindWorksheet(workbook, referenceName);
                    if (reference == null)
                    {
                        return Fail(OfficeErrorCode.SHEET_NOT_FOUND, "找不到参考工作表「" + referenceName + "」。");
                    }
                }

                application.Guard(delegate
                {
                    if (reference == null)
                    {
                        created = (Excel.Worksheet)workbook.Worksheets.Add();
                    }
                    else if (position == "before")
                    {
                        created = (Excel.Worksheet)workbook.Worksheets.Add(Before: reference);
                    }
                    else
                    {
                        created = (Excel.Worksheet)workbook.Worksheets.Add(After: reference);
                    }
                    created.Name = name;
                });

                Observation observation = Observation.Create(
                    "create_sheet",
                    "已创建工作表「" + name + "」。",
                    true);
                observation.AddTargetRef(ExcelComSupport.FullAddress(name, "A1"));

                JObject data = new JObject();
                data["name"] = name;
                data["index"] = created.Index;
                data["position"] = string.IsNullOrWhiteSpace(position) ? "after" : position;
                ExcelUndoScope.RecordUnavailable(
                    data,
                    "新建工作表属于结构变更，T02 区域快照无法撤销，未登记自动撤销。");
                return OperationResult.Succeed("工作表已创建：" + name + "。", data, observation);
            }
            finally
            {
                ComHelper.Release(created, reference, workbook);
            }
        }

        private static OperationResult ExecuteDeleteSheet(Excel.Application application, JObject arguments)
        {
            OperationResult denied = SecurityGate.CreateDenial(SecurityRisk.DeleteSheet, "excel_delete_sheet");
            if (denied != null)
            {
                return denied;
            }

            string name = ExcelArgumentReader.GetTrimmedString(arguments, "name", "sheetName");
            if (name.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 name/sheetName 参数。");
            }

            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Range usedRange = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                worksheet = ExcelComSupport.FindWorksheet(workbook, name);
                if (worksheet == null)
                {
                    return Fail(OfficeErrorCode.SHEET_NOT_FOUND, "找不到工作表「" + name + "」。");
                }

                usedRange = worksheet.UsedRange;
                string usedAddress = usedRange.AddressLocal();
                ExcelUndoScope undo = ExcelUndoScope.Capture(workbook, worksheet, usedRange, usedAddress);
                int sheetIndex = worksheet.Index;

                application.Guard(delegate { worksheet.Delete(); });

                Observation observation = Observation.Create(
                    "delete_sheet",
                    "已删除工作表「" + name + "」。",
                    true);
                observation.AddTargetRef(ExcelComSupport.FullAddress(name, usedAddress));

                JObject data = new JObject();
                data["name"] = name;
                data["index"] = sheetIndex;
                undo.Complete(
                    workbook,
                    worksheet,
                    null,
                    null,
                    false,
                    "删除工作表 " + name,
                    observation,
                    data,
                    "删除工作表属于结构变更，区域快照无法恢复已删除的工作表，未登记自动撤销。");

                return OperationResult.Succeed("工作表已删除：" + name + "。", data, observation);
            }
            finally
            {
                ComHelper.Release(usedRange, worksheet, workbook);
            }
        }

        private static OperationResult ExecuteRenameSheet(Excel.Application application, JObject arguments)
        {
            string oldName = ExcelArgumentReader.GetTrimmedString(arguments, "oldName", "name", "sheetName");
            string newName = ExcelArgumentReader.GetTrimmedString(arguments, "newName");
            if (oldName.Length == 0 || newName.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 oldName/name 与 newName 参数。");
            }
            string nameError = ValidateSheetName(newName);
            if (nameError.Length > 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, nameError);
            }

            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Range usedRange = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                worksheet = ExcelComSupport.FindWorksheet(workbook, oldName);
                if (worksheet == null)
                {
                    return Fail(OfficeErrorCode.SHEET_NOT_FOUND, "找不到工作表「" + oldName + "」。");
                }
                if (string.Equals(oldName, newName, StringComparison.Ordinal))
                {
                    Observation unchanged = Observation.Create(
                        "rename_sheet",
                        "新旧工作表名相同，未做修改。",
                        false);
                    unchanged.AddTargetRef(ExcelComSupport.FullAddress(oldName, "A1"));
                    JObject unchangedData = new JObject();
                    unchangedData["oldName"] = oldName;
                    unchangedData["newName"] = newName;
                    ExcelUndoScope.RecordUnavailable(unchangedData, "没有发生修改，无需撤销。");
                    return OperationResult.Succeed("新旧工作表名相同，未做修改。", unchangedData, unchanged);
                }

                Excel.Worksheet duplicate = ExcelComSupport.FindWorksheet(workbook, newName);
                if (duplicate != null)
                {
                    ComHelper.Release(duplicate);
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "工作表「" + newName + "」已存在。");
                }

                ExcelComSupport.EnsureWritable(application, workbook, worksheet, "excel_rename_sheet");
                usedRange = worksheet.UsedRange;
                string usedAddress = usedRange.AddressLocal();
                ExcelUndoScope undo = ExcelUndoScope.Capture(workbook, worksheet, usedRange, usedAddress);

                application.Guard(delegate { worksheet.Name = newName; });

                Observation observation = Observation.Create(
                    "rename_sheet",
                    "工作表「" + oldName + "」已重命名为「" + newName + "」。",
                    true);
                observation.AddTargetRef(ExcelComSupport.FullAddress(newName, usedAddress));

                JObject data = new JObject();
                data["oldName"] = oldName;
                data["newName"] = newName;
                undo.Complete(
                    workbook,
                    worksheet,
                    usedRange,
                    usedAddress,
                    false,
                    "重命名工作表 " + oldName + " -> " + newName,
                    observation,
                    data,
                    "重命名属于结构变更，区域快照仍引用旧表名，无法自动回放，未登记自动撤销。");

                return OperationResult.Succeed("工作表已重命名：" + oldName + " -> " + newName + "。", data, observation);
            }
            finally
            {
                ComHelper.Release(usedRange, worksheet, workbook);
            }
        }

        private static OperationResult ExecuteCopySheet(Excel.Application application, JObject arguments)
        {
            string sourceName = ExcelArgumentReader.GetTrimmedString(arguments, "sourceName", "name", "sheetName");
            string newName = ExcelArgumentReader.GetTrimmedString(arguments, "newName");
            if (sourceName.Length == 0 || newName.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 sourceName/name 与 newName 参数。");
            }
            string nameError = ValidateSheetName(newName);
            if (nameError.Length > 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, nameError);
            }

            Excel.Workbook workbook = null;
            Excel.Worksheet source = null;
            Excel.Worksheet copied = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                source = ExcelComSupport.FindWorksheet(workbook, sourceName);
                if (source == null)
                {
                    return Fail(OfficeErrorCode.SHEET_NOT_FOUND, "找不到工作表「" + sourceName + "」。");
                }
                Excel.Worksheet duplicate = ExcelComSupport.FindWorksheet(workbook, newName);
                if (duplicate != null)
                {
                    ComHelper.Release(duplicate);
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "工作表「" + newName + "」已存在。");
                }

                ExcelComSupport.EnsureWritable(application, workbook, source, "excel_copy_sheet");
                application.Guard(delegate
                {
                    source.Copy(After: source);
                    object active = application.ActiveSheet;
                    copied = active as Excel.Worksheet;
                    if (copied == null)
                    {
                        ComHelper.Release(active);
                        throw new OfficeComException(OfficeErrorCode.INTERNAL_ERROR, "复制工作表后无法取得新工作表对象。");
                    }
                    copied.Name = newName;
                });

                Observation observation = Observation.Create(
                    "copy_sheet",
                    "已复制工作表「" + sourceName + "」为「" + newName + "」。",
                    true);
                observation.AddTargetRef(ExcelComSupport.FullAddress(newName, "A1"));

                JObject data = new JObject();
                data["sourceName"] = sourceName;
                data["newName"] = newName;
                data["index"] = copied.Index;
                ExcelUndoScope.RecordUnavailable(
                    data,
                    "复制工作表属于结构变更，区域快照无法撤销新增工作表，未登记自动撤销。");
                return OperationResult.Succeed(
                    "工作表已复制：" + sourceName + " -> " + newName + "。",
                    data,
                    observation);
            }
            catch
            {
                if (copied != null)
                {
                    ExcelPivotService.TryDeleteWorksheet(application, copied);
                }
                throw;
            }
            finally
            {
                ComHelper.Release(copied, source, workbook);
            }
        }

        private static OperationResult ExecuteInsertRowCol(Excel.Application application, JObject arguments)
        {
            return ExecuteRowColumnCommand(application, arguments, "insert");
        }

        private static OperationResult ExecuteDeleteRowCol(Excel.Application application, JObject arguments)
        {
            SecurityDecision decision = SecurityGate.Evaluate(SecurityRisk.DeleteSheet, "excel_delete_row_col");
            if (!decision.Allowed)
            {
                return OperationResult.Failed(
                    decision.ErrorCode,
                    "工具「excel_delete_row_col」被安全策略拒绝：删除行列属于不可逆高危操作，" +
                    "受 config.json 的 security.allowDeleteSheet 控制。请先备份，或由用户授权后把该配置项改为 true。");
            }
            return ExecuteRowColumnCommand(application, arguments, "delete");
        }

        private static OperationResult ExecuteHideRowCol(Excel.Application application, JObject arguments)
        {
            return ExecuteRowColumnCommand(application, arguments, "hide");
        }

        private static OperationResult ExecuteRowColumnCommand(
            Excel.Application application,
            JObject arguments,
            string action)
        {
            string sheetName = ExcelArgumentReader.GetSheetName(arguments);
            string type = ExcelArgumentReader.GetString(arguments, "type", "rowOrColumn").Trim().ToLowerInvariant();
            if (type == "rows")
            {
                type = "row";
            }
            else if (type == "columns")
            {
                type = "column";
            }
            if (type != "row" && type != "column")
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "type 仅支持 row/column。");
            }

            string position = ExcelArgumentReader.GetTrimmedString(arguments, "position", "start", "row", "column");
            int count = ExcelArgumentReader.GetInt(arguments, 1, "count");
            bool unhide = ExcelArgumentReader.GetBool(arguments, false, "unhide");
            if (position.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 position 参数。");
            }
            if (count < 1 || count > 100000)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "count 必须在 1 到 100000 之间。");
            }

            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Range usedRange = null;
            Excel.Range target = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                worksheet = RangeResolver.ResolveWorksheet(application, workbook, sheetName, false);
                ExcelComSupport.EnsureWritable(application, workbook, worksheet, "excel_" + action + "_row_col");

                string startLabel;
                string endLabel;
                if (type == "row")
                {
                    int startRow;
                    if (!int.TryParse(position, NumberStyles.Integer, CultureInfo.InvariantCulture, out startRow) || startRow < 1)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "行号必须是大于 0 的整数。");
                    }
                    startLabel = startRow.ToString(CultureInfo.InvariantCulture);
                    endLabel = (startRow + count - 1).ToString(CultureInfo.InvariantCulture);
                    target = (Excel.Range)worksheet.Rows[startLabel + ":" + endLabel];
                }
                else
                {
                    int startColumn = ExcelComSupport.ParseColumnToken(position, 0);
                    if (startColumn < 1)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "列位置必须是列字母（A、B、AA）或大于 0 的整数。");
                    }
                    startLabel = ExcelComSupport.ColumnLetter(startColumn);
                    endLabel = ExcelComSupport.ColumnLetter(startColumn + count - 1);
                    target = (Excel.Range)worksheet.Columns[startLabel + ":" + endLabel];
                }

                string targetAddress = target.AddressLocal();
                usedRange = worksheet.UsedRange;
                string usedAddress = usedRange.AddressLocal();
                ExcelUndoScope undo = ExcelUndoScope.Capture(workbook, worksheet, usedRange, usedAddress);

                application.Guard(delegate
                {
                    if (action == "insert")
                    {
                        target.Insert(type == "row" ? Excel.XlInsertShiftDirection.xlShiftDown : Excel.XlInsertShiftDirection.xlShiftToRight);
                    }
                    else if (action == "delete")
                    {
                        target.Delete(type == "row" ? Excel.XlDeleteShiftDirection.xlShiftUp : Excel.XlDeleteShiftDirection.xlShiftToLeft);
                    }
                    else
                    {
                        target.Hidden = !unhide;
                    }
                });

                string targetRef = ExcelComSupport.FullAddress(worksheet.Name, targetAddress);
                string verb = action == "insert" ? "插入" : action == "delete" ? "删除" : (unhide ? "显示" : "隐藏");
                Observation observation = Observation.Create(
                    action + "_row_col",
                    "已" + verb + " " + targetRef + "（" + count + (type == "row" ? " 行" : " 列") + "）。",
                    true);
                observation.AddTargetRef(targetRef);

                JObject data = new JObject();
                data["sheet"] = worksheet.Name;
                data["type"] = type;
                data["position"] = position;
                data["count"] = count;
                data["action"] = action;
                data["unhide"] = unhide;
                data["address"] = targetAddress;

                string reason = action == "hide"
                    ? "隐藏/显示状态不在区域快照中，未登记自动撤销。"
                    : "插入/删除行列会改变工作表结构，区域快照无法完整逆操作，未登记自动撤销。";
                undo.Complete(
                    workbook,
                    worksheet,
                    action == "hide" ? target : usedRange,
                    action == "hide" ? targetAddress : usedAddress,
                    false,
                    verb + "行列 " + targetRef,
                    observation,
                    data,
                    reason);

                return OperationResult.Succeed("已" + verb + "：" + targetRef + "。", data, observation);
            }
            finally
            {
                ComHelper.Release(target, usedRange, worksheet, workbook);
            }
        }

        private static OperationResult ExecuteProtectSheet(Excel.Application application, JObject arguments)
        {
            string sheetName = ExcelArgumentReader.GetSheetName(arguments);
            string password = ExcelArgumentReader.GetString(arguments, "password");
            bool unprotect = ExcelArgumentReader.GetBool(arguments, false, "unprotect");

            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                worksheet = RangeResolver.ResolveWorksheet(application, workbook, sheetName, false);
                bool wasProtected = worksheet.ProtectContents;

                if (unprotect)
                {
                    if (!wasProtected)
                    {
                        return ProtectResult(worksheet.Name, false, false, "工作表当前未受保护，无需取消。");
                    }
                    application.Guard(delegate { worksheet.Unprotect(password); });
                }
                else
                {
                    if (wasProtected)
                    {
                        return ProtectResult(worksheet.Name, true, false, "工作表已处于保护状态，未重复保护。");
                    }
                    application.Guard(delegate
                    {
                        worksheet.Protect(
                            Password: password,
                            DrawingObjects: true,
                            Contents: true,
                            Scenarios: false);
                    });
                }

                return ProtectResult(
                    worksheet.Name,
                    !unprotect,
                    true,
                    unprotect ? "工作表保护已取消。" : "工作表已保护。");
            }
            finally
            {
                ComHelper.Release(worksheet, workbook);
            }
        }

        private static OperationResult ProtectResult(string sheetName, bool protectedState, bool changed, string message)
        {
            Observation observation = Observation.Create(
                "protect_sheet",
                message + "（" + sheetName + "）",
                changed);
            observation.AddTargetRef(ExcelComSupport.FullAddress(sheetName, "A1"));
            JObject data = new JObject();
            data["sheet"] = sheetName;
            data["protected"] = protectedState;
            ExcelUndoScope.RecordUnavailable(
                data,
                "工作表保护状态不在区域快照中，未登记自动撤销。");
            return OperationResult.Succeed(message, data, observation);
        }

        private static string ValidateSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return "工作表名称不能为空。";
            }
            if (name.Length > 31)
            {
                return "工作表名称不能超过 31 个字符。";
            }
            if (name.Trim().Length != name.Length)
            {
                return "工作表名称首尾不能包含空格。";
            }
            if (Regex.IsMatch(name, @"[\[\]\:\*\?\/\\]"))
            {
                return "工作表名称不能包含 \\ / ? * [ ] : 等字符。";
            }
            return string.Empty;
        }
    }
}
