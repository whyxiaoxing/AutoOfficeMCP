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
    /// <summary>P2 管理工作簿对象：命名区域、数据验证、Excel 表（ListObject）。</summary>
    public sealed partial class ExcelExtensionsExecutor
    {
        private static OperationResult ExecuteNamedRangeManage(Excel.Application application, JObject arguments)
        {
            string action = ExcelArgumentReader.GetTrimmedString(arguments, "action", "operation")
                .ToLowerInvariant();
            if (action.Length == 0)
            {
                action = "list";
            }
            string name = ExcelArgumentReader.GetTrimmedString(arguments, "name", "rangeName");
            string newName = ExcelArgumentReader.GetTrimmedString(arguments, "newName");
            string refersTo = ExcelArgumentReader.GetTrimmedString(arguments, "refersTo", "reference", "formula");
            string scope = ExcelArgumentReader.GetTrimmedString(arguments, "scope").ToLowerInvariant();
            if (scope.Length == 0)
            {
                scope = "workbook";
            }
            string sheetName = ExcelArgumentReader.GetSheetName(arguments);
            bool visible = ExcelArgumentReader.GetBool(arguments, true, "visible");
            string comment = ExcelArgumentReader.GetString(arguments, "comment");

            Excel.Workbook workbook = null;
            Excel.Worksheet scopeWorksheet = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);

                if (action == "list")
                {
                    JArray names = ListDefinedNames(workbook);
                    int brokenCount = 0;
                    foreach (JToken token in names)
                    {
                        JObject item = token as JObject;
                        if (item != null && item.Value<bool>("broken"))
                        {
                            brokenCount++;
                        }
                    }
                    JObject data = new JObject();
                    data["count"] = names.Count;
                    data["brokenCount"] = brokenCount;
                    data["names"] = names;
                    string summary = "命名区域清单：共 " + names.Count + " 个，断链 " + brokenCount + " 个。";
                    Observation observation = ReadOnlyObservation("named_range_manage", summary);
                    return OperationResult.Succeed(summary, data, observation);
                }

                if (workbook.ReadOnly)
                {
                    return Fail(OfficeErrorCode.READ_ONLY, "工作簿以只读方式打开，无法修改命名区域。");
                }

                if (action == "cleanbroken" || action == "clean_broken")
                {
                    return CleanBrokenNames(workbook, application);
                }

                if (action != "add" && action != "create" && action != "update" &&
                    action != "delete" && action != "remove")
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "action 仅支持 list/add/update/delete/cleanBroken。");
                }

                if (name.Length == 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 name 参数。");
                }

                bool sheetScope = scope == "sheet" || scope == "worksheet" || scope == "local";
                if (sheetScope)
                {
                    if (sheetName.Length == 0)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "scope=sheet 时必须提供 sheetName。");
                    }
                    scopeWorksheet = RangeResolver.ResolveWorksheet(
                        application,
                        workbook,
                        sheetName,
                        false);
                }

                if (action == "delete" || action == "remove")
                {
                    Excel.Name target = FindDefinedName(workbook, scopeWorksheet, name, sheetScope);
                    if (target == null)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "找不到命名区域「" + name + "」。");
                    }
                    string oldRefersTo;
                    string oldScope;
                    try
                    {
                        oldRefersTo = Convert.ToString(target.RefersTo, CultureInfo.InvariantCulture) ?? string.Empty;
                        oldScope = target.Name ?? name;
                        target.Delete();
                    }
                    finally
                    {
                        ComHelper.Release(target);
                    }

                    Observation observation = Observation.Create(
                        "named_range_manage",
                        "已删除命名区域「" + oldScope + "」。",
                        true);
                    observation.AddTargetRef(oldScope);
                    JObject data = new JObject();
                    data["action"] = "delete";
                    data["name"] = name;
                    data["refersTo"] = oldRefersTo;
                    ExcelUndoScope.RecordUnavailable(
                        data,
                        "命名区域属于工作簿元数据，区域快照无法撤销名称的增删改。");
                    return OperationResult.Succeed("已删除命名区域「" + name + "」。", data, observation);
                }

                if (refersTo.Length == 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 refersTo 参数（例如 =Data!$A$1:$B$10）。");
                }
                string nameError = ValidateDefinedName(name);
                if (nameError.Length > 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, nameError);
                }

                string normalizedRefersTo = NormalizeRefersTo(refersTo, sheetScope ? sheetName : string.Empty);
                Excel.Name existing = FindDefinedName(workbook, scopeWorksheet, name, sheetScope);

                if (action == "add" || action == "create")
                {
                    if (existing != null)
                    {
                        ComHelper.Release(existing);
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "命名区域「" + name + "」已存在。如需修改请使用 action=update。");
                    }
                    Excel.Name created = null;
                    Excel.Names collection = null;
                    try
                    {
                        collection = sheetScope ? scopeWorksheet.Names : workbook.Names;
                        created = collection.Add(name, normalizedRefersTo);
                        if (created == null)
                        {
                            return Fail(OfficeErrorCode.INTERNAL_ERROR, "Excel 未返回新建的命名区域对象。");
                        }
                        created.Visible = visible;
                        if (comment.Length > 0)
                        {
                            created.Comment = comment;
                        }
                    }
                    finally
                    {
                        ComHelper.Release(created, collection);
                    }

                    Observation observation = Observation.Create(
                        "named_range_manage",
                        "已创建命名区域「" + name + "」→ " + normalizedRefersTo,
                        true);
                    observation.AddTargetRef(name);
                    JObject data = new JObject();
                    data["action"] = "add";
                    data["name"] = name;
                    data["refersTo"] = normalizedRefersTo;
                    data["scope"] = sheetScope ? "sheet:" + sheetName : "workbook";
                    data["visible"] = visible;
                    ExcelUndoScope.RecordUnavailable(
                        data,
                        "命名区域属于工作簿元数据，区域快照无法撤销名称的增删改。");
                    return OperationResult.Succeed("已创建命名区域「" + name + "」。", data, observation);
                }

                if (existing == null)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "找不到命名区域「" + name + "」。");
                }
                try
                {
                    if (newName.Length > 0 && !string.Equals(newName, name, StringComparison.Ordinal))
                    {
                        string renameError = ValidateDefinedName(newName);
                        if (renameError.Length > 0)
                        {
                            return Fail(OfficeErrorCode.INVALID_PARAMS, renameError);
                        }
                        Excel.Name duplicate = FindDefinedName(workbook, scopeWorksheet, newName, sheetScope);
                        if (duplicate != null)
                        {
                            ComHelper.Release(duplicate);
                            return Fail(OfficeErrorCode.INVALID_PARAMS, "命名区域「" + newName + "」已存在。");
                        }
                    }

                    existing.RefersTo = normalizedRefersTo;
                    existing.Visible = visible;
                    if (comment.Length > 0)
                    {
                        existing.Comment = comment;
                    }
                    if (newName.Length > 0)
                    {
                        existing.Name = newName;
                    }
                }
                finally
                {
                    ComHelper.Release(existing);
                }

                string finalName = newName.Length > 0 ? newName : name;
                Observation updatedObservation = Observation.Create(
                    "named_range_manage",
                    "已更新命名区域「" + finalName + "」→ " + normalizedRefersTo,
                    true);
                updatedObservation.AddTargetRef(finalName);
                JObject updatedData = new JObject();
                updatedData["action"] = "update";
                updatedData["name"] = finalName;
                updatedData["previousName"] = name;
                updatedData["refersTo"] = normalizedRefersTo;
                updatedData["scope"] = sheetScope ? "sheet:" + sheetName : "workbook";
                updatedData["visible"] = visible;
                ExcelUndoScope.RecordUnavailable(
                    updatedData,
                    "命名区域属于工作簿元数据，区域快照无法撤销名称的增删改。");
                return OperationResult.Succeed(
                    "已更新命名区域「" + finalName + "」。",
                    updatedData,
                    updatedObservation);
            }
            finally
            {
                ComHelper.Release(scopeWorksheet, workbook);
            }
        }

        private static JArray ListDefinedNames(Excel.Workbook workbook)
        {
            JArray result = new JArray();
            Excel.Names names = null;
            try
            {
                names = workbook.Names;
                int count = names.Count;
                for (int index = 1; index <= count; index++)
                {
                    Excel.Name nameObject = null;
                    try
                    {
                        object item = names.Item(index);
                        nameObject = item as Excel.Name;
                        if (nameObject == null)
                        {
                            ComHelper.Release(item);
                            continue;
                        }

                        string rawName = nameObject.Name ?? string.Empty;
                        string scopeName = "workbook";
                        string localName = rawName;
                        int bang = rawName.IndexOf('!');
                        if (bang > 0)
                        {
                            scopeName = rawName.Substring(0, bang).Trim('\'');
                            localName = rawName.Substring(bang + 1);
                        }

                        string refersToText = Convert.ToString(
                            nameObject.RefersTo,
                            CultureInfo.InvariantCulture) ?? string.Empty;
                        JObject entry = new JObject();
                        entry["name"] = localName;
                        entry["rawName"] = rawName;
                        entry["scope"] = scopeName;
                        entry["refersTo"] = refersToText;
                        entry["visible"] = nameObject.Visible;
                        entry["comment"] = nameObject.Comment ?? string.Empty;
                        entry["broken"] = refersToText.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0;

                        Excel.Range targetRange = null;
                        Excel.Worksheet targetSheet = null;
                        try
                        {
                            targetRange = nameObject.RefersToRange;
                            if (targetRange != null)
                            {
                                targetSheet = targetRange.Worksheet as Excel.Worksheet;
                                entry["address"] = ExcelComSupport.FullAddress(
                                    targetSheet == null ? string.Empty : targetSheet.Name,
                                    targetRange.AddressLocal());
                            }
                        }
                        catch (COMException)
                        {
                            entry["address"] = string.Empty;
                        }
                        finally
                        {
                            ComHelper.Release(targetRange, targetSheet);
                        }
                        result.Add(entry);
                    }
                    finally
                    {
                        ComHelper.Release(nameObject);
                    }
                }
            }
            finally
            {
                ComHelper.Release(names);
            }
            return result;
        }

        private static OperationResult CleanBrokenNames(Excel.Workbook workbook, Excel.Application application)
        {
            List<string> broken = new List<string>();
            Excel.Names names = null;
            try
            {
                names = workbook.Names;
                int count = names.Count;
                for (int index = 1; index <= count; index++)
                {
                    Excel.Name nameObject = null;
                    try
                    {
                        nameObject = names.Item(index) as Excel.Name;
                        if (nameObject == null)
                        {
                            continue;
                        }
                        string refersToText = Convert.ToString(
                            nameObject.RefersTo,
                            CultureInfo.InvariantCulture) ?? string.Empty;
                        if (refersToText.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            broken.Add(nameObject.Name);
                        }
                    }
                    finally
                    {
                        ComHelper.Release(nameObject);
                    }
                }
            }
            finally
            {
                ComHelper.Release(names);
            }

            List<string> deleted = new List<string>();
            List<string> failed = new List<string>();
            foreach (string rawName in broken)
            {
                Excel.Name nameObject = null;
                try
                {
                    nameObject = FindDefinedName(workbook, null, rawName, false);
                    if (nameObject == null)
                    {
                        failed.Add(rawName + "（已不存在）");
                        continue;
                    }
                    nameObject.Delete();
                    deleted.Add(rawName);
                }
                catch (Exception ex)
                {
                    failed.Add(rawName + "：" + ex.Message);
                }
                finally
                {
                    ComHelper.Release(nameObject);
                }
            }

            JObject data = new JObject();
            data["action"] = "cleanBroken";
            data["deletedCount"] = deleted.Count;
            data["deleted"] = new JArray(deleted.ToArray());
            data["failed"] = new JArray(failed.ToArray());
            ExcelUndoScope.RecordUnavailable(
                data,
                "命名区域属于工作簿元数据，区域快照无法撤销名称删除。");
            Observation observation = Observation.Create(
                "named_range_manage",
                "已清理断链命名区域 " + deleted.Count + " 个。",
                deleted.Count > 0);
            string summary = deleted.Count == 0 && failed.Count == 0
                ? "没有发现断链命名区域。"
                : "已清理断链命名区域 " + deleted.Count + " 个" +
                  (failed.Count > 0 ? "，失败 " + failed.Count + " 个。" : "。");
            return OperationResult.Succeed(summary, data, observation);
        }

        private static Excel.Name FindDefinedName(
            Excel.Workbook workbook,
            Excel.Worksheet worksheet,
            string name,
            bool sheetScope)
        {
            Excel.Names collection = null;
            try
            {
                collection = sheetScope && worksheet != null ? worksheet.Names : workbook.Names;
                object item = collection.Item(name);
                Excel.Name result = item as Excel.Name;
                if (result == null)
                {
                    ComHelper.Release(item);
                }
                return result;
            }
            catch (COMException)
            {
                return null;
            }
            finally
            {
                ComHelper.Release(collection);
            }
        }

        private static string ValidateDefinedName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return "命名区域名称不能为空。";
            }
            if (name.Length > 255)
            {
                return "命名区域名称不能超过 255 个字符。";
            }
            if (name.IndexOf('!') >= 0 || name.IndexOf(' ') >= 0)
            {
                return "命名区域名称不能包含空格或感叹号。";
            }
            if (Regex.IsMatch(name, @"^[A-Za-z]{1,3}\d+$"))
            {
                return "命名区域名称不能与单元格地址相同（例如 A1、XFD1048576）。";
            }
            if (Regex.IsMatch(name, @"^R\d+C\d+$", RegexOptions.IgnoreCase))
            {
                return "命名区域名称不能使用 R1C1 单元格引用格式。";
            }
            return string.Empty;
        }

        private static string NormalizeRefersTo(string refersTo, string sheetName)
        {
            string text = refersTo.Trim();
            if (text.StartsWith("=", StringComparison.Ordinal))
            {
                text = text.Substring(1).Trim();
            }
            if (!string.IsNullOrWhiteSpace(sheetName) && text.IndexOf('!') < 0)
            {
                text = ExcelComSupport.QuoteSheetName(sheetName) + "!" + text;
            }
            return "=" + text;
        }

        private static OperationResult ExecuteDataValidationManage(Excel.Application application, JObject arguments)
        {
            string action = ExcelArgumentReader.GetTrimmedString(arguments, "action", "operation")
                .ToLowerInvariant();
            if (action.Length == 0)
            {
                action = "list";
            }
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange", "address");
            if (address.Length == 0)
            {
                address = "usedrange";
            }
            string sheetName = ExcelArgumentReader.GetSheetName(arguments);
            if (address == "usedrange" && sheetName.Length > 0)
            {
                address = ExcelComSupport.QuoteSheetName(sheetName) + "!usedrange";
            }

            using (ExcelRangeContext context = RangeResolver.Resolve(
                       application,
                       address,
                       false,
                       "excel_data_validation_manage"))
            {
                if (action == "list")
                {
                    int maxResults = ClampInt(
                        ExcelArgumentReader.GetInt(arguments, 500, "maxResults", "limit"),
                        1,
                        5000);
                    JArray entries = new JArray();
                    bool truncated = false;
                    CollectValidationEntries(
                        context.Worksheet,
                        context.Range,
                        entries,
                        maxResults,
                        ref truncated);
                    JObject listData = new JObject();
                    listData["sheet"] = context.SheetName;
                    listData["address"] = context.Range.AddressLocal();
                    listData["count"] = entries.Count;
                    listData["truncated"] = truncated;
                    listData["validations"] = entries;
                    string summary = "数据验证清单：共 " + entries.Count + " 条规则。";
                    Observation listObservation = ReadOnlyObservation("data_validation_manage", summary);
                    return OperationResult.Succeed(summary, listData, listObservation);
                }

                ExcelComSupport.EnsureWritable(
                    application,
                    context.Workbook,
                    context.Worksheet,
                    "excel_data_validation_manage");

                if (action == "clear" || action == "delete" || action == "remove")
                {
                    Excel.Validation validation = null;
                    try
                    {
                        validation = context.Range.Validation;
                        validation.Delete();
                    }
                    finally
                    {
                        ComHelper.Release(validation);
                    }
                    Observation clearObservation = Observation.Create(
                        "data_validation_manage",
                        "已清除 " + context.FullAddress + " 的数据验证规则。",
                        true);
                    clearObservation.AddTargetRef(context.FullAddress);
                    JObject clearData = new JObject();
                    clearData["action"] = "clear";
                    clearData["sheet"] = context.SheetName;
                    clearData["address"] = context.Range.AddressLocal();
                    ExcelUndoScope.RecordUnavailable(
                        clearData,
                        "数据验证规则不在区域快照中，未登记自动撤销。");
                    return OperationResult.Succeed("数据验证已清除。", clearData, clearObservation);
                }

                if (action != "set" && action != "add" && action != "update")
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "action 仅支持 list/set/clear。");
                }

                string typeText = ExcelArgumentReader.GetTrimmedString(arguments, "type", "validationType");
                Excel.XlDVType validationType;
                if (!TryMapValidationType(typeText, out validationType))
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "type 仅支持 list/whole/decimal/date/time/textLength/custom。");
                }

                string formula1 = BuildValidationFormula(arguments, validationType);
                if (formula1.Length == 0)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "缺少 formula1 或 values 参数（list 可用 values 提供下拉选项）。");
                }
                string formula2 = ExcelArgumentReader.GetString(arguments, "formula2", "max");
                Excel.XlFormatConditionOperator validationOperator;
                if (!TryMapValidationOperator(
                        ExcelArgumentReader.GetTrimmedString(arguments, "operator"),
                        out validationOperator))
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "operator 仅支持 between/notBetween/equal/notEqual/greater/less/greaterEqual/lessEqual。");
                }

                bool ignoreBlank = ExcelArgumentReader.GetBool(arguments, true, "ignoreBlank");
                bool inCellDropdown = ExcelArgumentReader.GetBool(arguments, true, "inCellDropdown");
                bool showError = ExcelArgumentReader.GetBool(arguments, true, "showError");
                bool showInput = ExcelArgumentReader.GetBool(arguments, false, "showInput");
                string errorTitle = ExcelArgumentReader.GetString(arguments, "errorTitle", "title");
                string errorMessage = ExcelArgumentReader.GetString(arguments, "errorMessage", "message");
                string inputTitle = ExcelArgumentReader.GetString(arguments, "inputTitle");
                string inputMessage = ExcelArgumentReader.GetString(arguments, "inputMessage");

                object operatorValue = validationType == Excel.XlDVType.xlValidateList ||
                                       validationType == Excel.XlDVType.xlValidateCustom
                    ? Type.Missing
                    : (object)validationOperator;
                bool rangeOperator = validationOperator == Excel.XlFormatConditionOperator.xlBetween ||
                                     validationOperator == Excel.XlFormatConditionOperator.xlNotBetween;
                bool supportsRange = validationType != Excel.XlDVType.xlValidateList &&
                                     validationType != Excel.XlDVType.xlValidateCustom;
                bool needsFormula2 = rangeOperator && supportsRange;
                object formula2Value = needsFormula2 ? (object)formula2 : Type.Missing;
                if (needsFormula2 && formula2.Length == 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "between/notBetween 需要 formula2。");
                }

                Excel.Validation targetValidation = null;
                bool oldRuleDeleted = false;
                try
                {
                    targetValidation = context.Range.Validation;
                    targetValidation.Delete();
                    oldRuleDeleted = true;
                    targetValidation.Add(
                        validationType,
                        Excel.XlDVAlertStyle.xlValidAlertStop,
                        operatorValue,
                        formula1,
                        formula2Value);
                    targetValidation.IgnoreBlank = ignoreBlank;
                    targetValidation.InCellDropdown = inCellDropdown;
                    targetValidation.ShowError = showError;
                    targetValidation.ShowInput = showInput;
                    if (errorTitle.Length > 0)
                    {
                        targetValidation.ErrorTitle = errorTitle;
                    }
                    if (errorMessage.Length > 0)
                    {
                        targetValidation.ErrorMessage = errorMessage;
                    }
                    if (inputTitle.Length > 0)
                    {
                        targetValidation.InputTitle = inputTitle;
                    }
                    if (inputMessage.Length > 0)
                    {
                        targetValidation.InputMessage = inputMessage;
                    }
                }
                catch (Exception ex)
                {
                    if (!oldRuleDeleted)
                    {
                        throw;
                    }
                    JObject partialData = new JObject();
                    partialData["action"] = "set";
                    partialData["sheet"] = context.SheetName;
                    partialData["address"] = context.Range.AddressLocal();
                    partialData["type"] = typeText;
                    partialData["formula1"] = formula1;
                    partialData["formula2"] = formula2;
                    ExcelUndoScope.RecordUnavailable(
                        partialData,
                        "数据验证规则不在区域快照中，未登记自动撤销。");
                    Observation partialObservation = Observation.Create(
                        "data_validation_manage",
                        "原数据验证已删除，新规则设置失败。",
                        true);
                    return OperationResult.PartialApplied(
                        1,
                        "原数据验证规则已删除，但新规则设置失败：" + ex.Message +
                        "。请检查参数或重新设置；不会自动回放。",
                        partialObservation,
                        partialData);
                }
                finally
                {
                    ComHelper.Release(targetValidation);
                }

                Observation observation = Observation.Create(
                    "data_validation_manage",
                    "已为 " + context.FullAddress + " 设置数据验证（" + typeText + "）。",
                    true);
                observation.AddTargetRef(context.FullAddress);
                JObject data = new JObject();
                data["action"] = "set";
                data["sheet"] = context.SheetName;
                data["address"] = context.Range.AddressLocal();
                data["type"] = typeText;
                data["operator"] = validationOperator.ToString();
                data["formula1"] = formula1;
                data["formula2"] = formula2;
                data["inCellDropdown"] = inCellDropdown;
                ExcelUndoScope.RecordUnavailable(
                    data,
                    "数据验证规则不在区域快照中，未登记自动撤销。");
                return OperationResult.Succeed(
                    "数据验证已设置：" + context.FullAddress + "。",
                    data,
                    observation);
            }
        }

        private static void CollectValidationEntries(
            Excel.Worksheet worksheet,
            Excel.Range range,
            JArray entries,
            int maxResults,
            ref bool truncated)
        {
            Excel.Range special = null;
            try
            {
                special = range.SpecialCells(Excel.XlCellType.xlCellTypeAllValidation, Type.Missing);
            }
            catch (COMException)
            {
                special = null;
            }
            if (special == null)
            {
                return;
            }

            Excel.Areas areas = null;
            try
            {
                areas = special.Areas;
                int areaCount = areas.Count;
                for (int areaIndex = 1; areaIndex <= areaCount; areaIndex++)
                {
                    Excel.Range area = null;
                    Excel.Validation validation = null;
                    try
                    {
                        area = (Excel.Range)areas[areaIndex];
                        validation = area.Validation;
                        JObject entry = new JObject();
                        int type = validation.Type;
                        entry["address"] = ExcelComSupport.FullAddress(
                            worksheet.Name,
                            area.AddressLocal());
                        entry["type"] = ValidationTypeText((Excel.XlDVType)type);
                        entry["operator"] = ((Excel.XlFormatConditionOperator)validation.Operator).ToString();
                        entry["formula1"] = SafeValidationFormula(validation, true);
                        entry["formula2"] = SafeValidationFormula(validation, false);
                        entry["inCellDropdown"] = validation.InCellDropdown;
                        entry["ignoreBlank"] = validation.IgnoreBlank;
                        entry["showError"] = validation.ShowError;
                        entry["showInput"] = validation.ShowInput;
                        entry["errorTitle"] = validation.ErrorTitle ?? string.Empty;
                        entry["errorMessage"] = validation.ErrorMessage ?? string.Empty;
                        entry["inputTitle"] = validation.InputTitle ?? string.Empty;
                        entry["inputMessage"] = validation.InputMessage ?? string.Empty;
                        entries.Add(entry);
                        if (entries.Count >= maxResults)
                        {
                            truncated = true;
                            return;
                        }
                    }
                    catch (COMException)
                    {
                        // 某些区域没有可读的验证对象，跳过即可。
                    }
                    finally
                    {
                        ComHelper.Release(validation, area);
                    }
                }
            }
            finally
            {
                ComHelper.Release(areas, special);
            }
        }

        private static string SafeValidationFormula(Excel.Validation validation, bool first)
        {
            try
            {
                return first ? validation.Formula1 ?? string.Empty : validation.Formula2 ?? string.Empty;
            }
            catch (COMException)
            {
                return string.Empty;
            }
        }

        private static string BuildValidationFormula(JObject arguments, Excel.XlDVType type)
        {
            JArray values = ExcelArgumentReader.GetArrayOrSingle(arguments, "values", "options", "listValues");
            if (values != null && values.Count > 0)
            {
                List<string> parts = new List<string>();
                foreach (JToken token in values)
                {
                    string text = token.ToString().Trim();
                    if (text.Length > 0)
                    {
                        parts.Add(text.Replace(",", "，"));
                    }
                }
                if (parts.Count > 0)
                {
                    return string.Join(",", parts.ToArray());
                }
            }
            string formula = ExcelArgumentReader.GetString(arguments, "formula1", "source", "min");
            return formula.Trim();
        }

        private static bool TryMapValidationType(string text, out Excel.XlDVType type)
        {
            type = Excel.XlDVType.xlValidateInputOnly;
            switch ((text ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "list":
                    type = Excel.XlDVType.xlValidateList;
                    return true;
                case "whole":
                case "wholenumber":
                case "integer":
                    type = Excel.XlDVType.xlValidateWholeNumber;
                    return true;
                case "decimal":
                case "number":
                    type = Excel.XlDVType.xlValidateDecimal;
                    return true;
                case "date":
                    type = Excel.XlDVType.xlValidateDate;
                    return true;
                case "time":
                    type = Excel.XlDVType.xlValidateTime;
                    return true;
                case "textlength":
                case "text_length":
                    type = Excel.XlDVType.xlValidateTextLength;
                    return true;
                case "custom":
                case "formula":
                    type = Excel.XlDVType.xlValidateCustom;
                    return true;
                default:
                    return false;
            }
        }

        private static bool TryMapValidationOperator(
            string text,
            out Excel.XlFormatConditionOperator comparisonOperator)
        {
            comparisonOperator = Excel.XlFormatConditionOperator.xlBetween;
            switch ((text ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "":
                    return true;
                case "between":
                    comparisonOperator = Excel.XlFormatConditionOperator.xlBetween;
                    return true;
                case "notbetween":
                case "not_between":
                    comparisonOperator = Excel.XlFormatConditionOperator.xlNotBetween;
                    return true;
                case "equal":
                case "=":
                    comparisonOperator = Excel.XlFormatConditionOperator.xlEqual;
                    return true;
                case "notequal":
                case "not_equal":
                case "<>":
                    comparisonOperator = Excel.XlFormatConditionOperator.xlNotEqual;
                    return true;
                case "greater":
                case ">":
                    comparisonOperator = Excel.XlFormatConditionOperator.xlGreater;
                    return true;
                case "less":
                case "<":
                    comparisonOperator = Excel.XlFormatConditionOperator.xlLess;
                    return true;
                case "greaterequal":
                case "greater_equal":
                case ">=":
                    comparisonOperator = Excel.XlFormatConditionOperator.xlGreaterEqual;
                    return true;
                case "lessequal":
                case "less_equal":
                case "<=":
                    comparisonOperator = Excel.XlFormatConditionOperator.xlLessEqual;
                    return true;
                default:
                    return false;
            }
        }

        private static string ValidationTypeText(Excel.XlDVType type)
        {
            switch (type)
            {
                case Excel.XlDVType.xlValidateList:
                    return "list";
                case Excel.XlDVType.xlValidateWholeNumber:
                    return "whole";
                case Excel.XlDVType.xlValidateDecimal:
                    return "decimal";
                case Excel.XlDVType.xlValidateDate:
                    return "date";
                case Excel.XlDVType.xlValidateTime:
                    return "time";
                case Excel.XlDVType.xlValidateTextLength:
                    return "textLength";
                case Excel.XlDVType.xlValidateCustom:
                    return "custom";
                default:
                    return "inputOnly";
            }
        }

        private static OperationResult ExecuteTableManage(Excel.Application application, JObject arguments)
        {
            string action = ExcelArgumentReader.GetTrimmedString(arguments, "action", "operation")
                .ToLowerInvariant();
            if (action.Length == 0)
            {
                action = "list";
            }
            string sheetName = ExcelArgumentReader.GetSheetName(arguments);
            string tableName = ExcelArgumentReader.GetTrimmedString(arguments, "tableName", "name");
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange", "sourceRange");

            Excel.Workbook workbook = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);

                if (action == "list")
                {
                    JArray tables = ListTables(workbook, sheetName);
                    JObject data = new JObject();
                    data["count"] = tables.Count;
                    data["tables"] = tables;
                    string summary = "Excel 表清单：共 " + tables.Count + " 个。";
                    Observation observation = ReadOnlyObservation("table_manage", summary);
                    return OperationResult.Succeed(summary, data, observation);
                }

                if (workbook.ReadOnly)
                {
                    return Fail(OfficeErrorCode.READ_ONLY, "工作簿以只读方式打开，无法修改 Excel 表。");
                }

                if (action == "create" || action == "add")
                {
                    if (address.Length == 0)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "create 操作需要 range 参数。");
                    }
                    bool hasHeader = ExcelArgumentReader.GetBool(arguments, true, "hasHeader");
                    string styleName = ExcelArgumentReader.GetTrimmedString(arguments, "tableStyle", "style");
                    if (tableName.Length > 0)
                    {
                        string tableNameError = ValidateDefinedName(tableName);
                        if (tableNameError.Length > 0)
                        {
                            return Fail(OfficeErrorCode.INVALID_PARAMS, "tableName 无效：" + tableNameError);
                        }
                    }
                    using (ExcelRangeContext context = RangeResolver.Resolve(
                               application,
                               address,
                               false,
                               "excel_table_manage"))
                    {
                        if (context.Range.Rows.Count < 1 || context.Range.Columns.Count < 1)
                        {
                            return Fail(OfficeErrorCode.INVALID_PARAMS, "创建 Excel 表的区域不能为空。");
                        }
                        ExcelComSupport.EnsureWritable(
                            application,
                            context.Workbook,
                            context.Worksheet,
                            "excel_table_manage");

                        Excel.ListObject table = null;
                        string styleWarning = string.Empty;
                        try
                        {
                            table = context.Worksheet.ListObjects.Add(
                                Excel.XlListObjectSourceType.xlSrcRange,
                                context.Range,
                                Type.Missing,
                                hasHeader ? Excel.XlYesNoGuess.xlYes : Excel.XlYesNoGuess.xlNo,
                                Type.Missing);
                            if (!string.IsNullOrWhiteSpace(tableName))
                            {
                                table.Name = tableName;
                            }
                            if (!string.IsNullOrWhiteSpace(styleName))
                            {
                                try
                                {
                                    table.TableStyle = styleName;
                                }
                                catch (Exception ex)
                                {
                                    styleWarning = "表样式设置失败：" + ex.Message;
                                }
                            }

                            string finalName = table.Name;
                            JObject data = new JObject();
                            data["action"] = "create";
                            data["name"] = finalName;
                            data["sheet"] = context.SheetName;
                            data["address"] = context.Range.AddressLocal();
                            data["rows"] = context.Range.Rows.Count;
                            data["columns"] = context.Range.Columns.Count;
                            data["hasHeader"] = hasHeader;
                            data["tableStyle"] = styleName;
                            if (styleWarning.Length > 0)
                            {
                                data["styleWarning"] = styleWarning;
                            }
                            ExcelUndoScope.RecordUnavailable(
                                data,
                                "Excel 表属于结构/元数据变更，区域快照无法撤销表对象本身。");

                            Observation observation = Observation.Create(
                                "table_manage",
                                "已创建 Excel 表「" + finalName + "」：" +
                                ExcelComSupport.FullAddress(context.SheetName, context.Range.AddressLocal()),
                                true);
                            observation.AddTargetRef(finalName);
                            string message = "已创建 Excel 表「" + finalName + "」。";
                            if (styleWarning.Length > 0)
                            {
                                message = message + " " + styleWarning;
                            }
                            return OperationResult.Succeed(message, data, observation);
                        }
                        finally
                        {
                            ComHelper.Release(table);
                        }
                    }
                }

                if (tableName.Length == 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "该操作需要 tableName 参数。");
                }

                Excel.Worksheet tableSheet;
                Excel.ListObject target = FindTable(workbook, sheetName, tableName, out tableSheet);
                if (target == null)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "找不到 Excel 表「" + tableName + "」。");
                }
                try
                {
                    ExcelComSupport.EnsureWritable(
                        application,
                        workbook,
                        tableSheet,
                        "excel_table_manage");

                    if (action == "converttorange" || action == "convert_to_range" || action == "unlist")
                    {
                        string oldRange;
                        Excel.Range oldTableRange = null;
                        try
                        {
                            oldTableRange = target.Range;
                            oldRange = oldTableRange.AddressLocal();
                        }
                        finally
                        {
                            ComHelper.Release(oldTableRange);
                        }
                        target.Unlist();
                        JObject data = new JObject();
                        data["action"] = "convertToRange";
                        data["name"] = tableName;
                        data["sheet"] = tableSheet.Name;
                        data["range"] = oldRange;
                        ExcelUndoScope.RecordUnavailable(
                            data,
                            "Excel 表转普通区域属于结构变更，区域快照无法恢复表对象。");
                        Observation observation = Observation.Create(
                            "table_manage",
                            "已将 Excel 表「" + tableName + "」转换为普通区域。",
                            true);
                        observation.AddTargetRef(ExcelComSupport.FullAddress(tableSheet.Name, oldRange));
                        return OperationResult.Succeed(
                            "Excel 表「" + tableName + "」已转换为普通区域。",
                            data,
                            observation);
                    }

                    if (action == "totals" || action == "total")
                    {
                        string columnToken = ExcelArgumentReader.GetTrimmedString(
                            arguments,
                            "totalsColumn",
                            "column",
                            "columnName");
                        if (columnToken.Length == 0)
                        {
                            return Fail(OfficeErrorCode.INVALID_PARAMS, "totals 操作需要 totalsColumn。");
                        }
                        string calculationText = ExcelArgumentReader.GetTrimmedString(
                            arguments,
                            "totalsFunction",
                            "calculation",
                            "function");
                        Excel.XlTotalsCalculation calculation;
                        if (!TryMapTotalsCalculation(calculationText, out calculation))
                        {
                            return Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "totalsFunction 仅支持 sum/count/average/max/min/stddev/var/none。");
                        }

                        Excel.ListColumns columns = null;
                        Excel.ListColumn listColumn = null;
                        try
                        {
                            columns = target.ListColumns;
                            listColumn = FindListColumn(columns, columnToken);
                            if (listColumn == null)
                            {
                                return Fail(
                                    OfficeErrorCode.INVALID_PARAMS,
                                    "找不到汇总列「" + columnToken + "」。");
                            }
                            listColumn.TotalsCalculation = calculation;
                            target.ShowTotals = true;
                        }
                        finally
                        {
                            ComHelper.Release(listColumn, columns);
                        }

                        JObject data = new JObject();
                        data["action"] = "totals";
                        data["name"] = target.Name;
                        data["sheet"] = tableSheet.Name;
                        data["column"] = columnToken;
                        data["totalsCalculation"] = calculation.ToString();
                        ExcelUndoScope.RecordUnavailable(
                            data,
                            "Excel 表汇总行属于表元数据，区域快照无法完整撤销。");
                        Observation observation = Observation.Create(
                            "table_manage",
                            "已为 Excel 表「" + target.Name + "」设置汇总行（" +
                            columnToken + " = " + calculation + "）。",
                            true);
                        observation.AddTargetRef(target.Name);
                        return OperationResult.Succeed(
                            "已设置汇总行：" + columnToken + " = " + calculation + "。",
                            data,
                            observation);
                    }

                    if (action == "resize")
                    {
                        if (address.Length == 0)
                        {
                            return Fail(OfficeErrorCode.INVALID_PARAMS, "resize 操作需要 range 参数。");
                        }
                        using (ExcelRangeContext resizeContext = RangeResolver.ResolveOnSheet(
                                   application,
                                   tableSheet.Name,
                                   address,
                                   false,
                                   "excel_table_manage"))
                        {
                            target.Resize(resizeContext.Range);
                            JObject data = new JObject();
                            data["action"] = "resize";
                            data["name"] = target.Name;
                            data["sheet"] = tableSheet.Name;
                            data["range"] = resizeContext.Range.AddressLocal();
                            ExcelUndoScope.RecordUnavailable(
                                data,
                                "Excel 表调整范围属于结构变更，区域快照无法恢复表对象。");
                            Observation observation = Observation.Create(
                                "table_manage",
                                "已调整 Excel 表「" + target.Name + "」范围。",
                                true);
                            observation.AddTargetRef(
                                ExcelComSupport.FullAddress(tableSheet.Name, resizeContext.Range.AddressLocal()));
                            return OperationResult.Succeed(
                                "已调整 Excel 表「" + target.Name + "」范围。",
                                data,
                                observation);
                        }
                    }

                    if (action == "rename")
                    {
                        string newName = ExcelArgumentReader.GetTrimmedString(arguments, "newName");
                        if (newName.Length == 0)
                        {
                            return Fail(OfficeErrorCode.INVALID_PARAMS, "rename 操作需要 newName。");
                        }
                        string oldName = target.Name;
                        target.Name = newName;
                        JObject data = new JObject();
                        data["action"] = "rename";
                        data["oldName"] = oldName;
                        data["newName"] = newName;
                        data["sheet"] = tableSheet.Name;
                        ExcelUndoScope.RecordUnavailable(
                            data,
                            "Excel 表重命名属于元数据变更，区域快照无法撤销。");
                        Observation observation = Observation.Create(
                            "table_manage",
                            "Excel 表「" + oldName + "」已重命名为「" + newName + "」。",
                            true);
                        observation.AddTargetRef(newName);
                        return OperationResult.Succeed(
                            "Excel 表已重命名为「" + newName + "」。",
                            data,
                            observation);
                    }

                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "action 仅支持 list/create/convertToRange/totals/resize/rename。");
                }
                finally
                {
                    ComHelper.Release(target, tableSheet);
                }
            }
            finally
            {
                ComHelper.Release(workbook);
            }
        }

        private static JArray ListTables(Excel.Workbook workbook, string sheetName)
        {
            JArray result = new JArray();
            List<Excel.Worksheet> worksheets = null;
            try
            {
                if (!string.IsNullOrWhiteSpace(sheetName))
                {
                    Excel.Worksheet worksheet = RangeResolver.ResolveWorksheet(
                        null,
                        workbook,
                        sheetName,
                        false);
                    worksheets = new List<Excel.Worksheet>();
                    worksheets.Add(worksheet);
                }
                else
                {
                    worksheets = ExcelExtensionSupport.CollectWorksheets(workbook, true, string.Empty);
                }

                foreach (Excel.Worksheet worksheet in worksheets)
                {
                    Excel.ListObjects tables = null;
                    try
                    {
                        tables = worksheet.ListObjects;
                        int count = tables.Count;
                        for (int index = 1; index <= count; index++)
                        {
                            Excel.ListObject table = null;
                            Excel.Range tableRange = null;
                            try
                            {
                                table = tables[index] as Excel.ListObject;
                                if (table == null)
                                {
                                    continue;
                                }
                                tableRange = table.Range;
                                JObject item = new JObject();
                                item["name"] = table.Name;
                                item["sheet"] = worksheet.Name;
                                item["address"] = tableRange.AddressLocal();
                                item["rows"] = tableRange.Rows.Count;
                                item["columns"] = tableRange.Columns.Count;
                                item["showTotals"] = table.ShowTotals;
                                item["tableStyle"] = Convert.ToString(
                                    table.TableStyle,
                                    CultureInfo.InvariantCulture) ?? string.Empty;
                                result.Add(item);
                            }
                            finally
                            {
                                ComHelper.Release(tableRange, table);
                            }
                        }
                    }
                    finally
                    {
                        ComHelper.Release(tables);
                    }
                }
                return result;
            }
            finally
            {
                ExcelExtensionSupport.ReleaseWorksheets(worksheets);
            }
        }

        private static Excel.ListObject FindTable(
            Excel.Workbook workbook,
            string sheetName,
            string tableName,
            out Excel.Worksheet tableSheet)
        {
            tableSheet = null;
            List<Excel.Worksheet> worksheets = null;
            try
            {
                if (!string.IsNullOrWhiteSpace(sheetName))
                {
                    Excel.Worksheet worksheet = RangeResolver.ResolveWorksheet(
                        null,
                        workbook,
                        sheetName,
                        false);
                    worksheets = new List<Excel.Worksheet>();
                    worksheets.Add(worksheet);
                }
                else
                {
                    worksheets = ExcelExtensionSupport.CollectWorksheets(workbook, true, string.Empty);
                }

                for (int worksheetIndex = 0; worksheetIndex < worksheets.Count; worksheetIndex++)
                {
                    Excel.Worksheet worksheet = worksheets[worksheetIndex];
                    Excel.ListObjects tables = null;
                    try
                    {
                        tables = worksheet.ListObjects;
                        int count = tables.Count;
                        for (int index = 1; index <= count; index++)
                        {
                            Excel.ListObject table = null;
                            try
                            {
                                table = tables[index] as Excel.ListObject;
                                if (table == null)
                                {
                                    continue;
                                }
                                if (string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase))
                                {
                                    Excel.ListObject found = table;
                                    table = null;
                                    tableSheet = worksheet;
                                    worksheets.RemoveAt(worksheetIndex);
                                    return found;
                                }
                            }
                            finally
                            {
                                ComHelper.Release(table);
                            }
                        }
                    }
                    finally
                    {
                        ComHelper.Release(tables);
                    }
                }
                return null;
            }
            finally
            {
                ExcelExtensionSupport.ReleaseWorksheets(worksheets);
            }
        }

        private static Excel.ListColumn FindListColumn(Excel.ListColumns columns, string token)
        {
            int index;
            if (int.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out index))
            {
                if (index >= 1 && index <= columns.Count)
                {
                    return columns[index] as Excel.ListColumn;
                }
                return null;
            }

            int count = columns.Count;
            for (int columnIndex = 1; columnIndex <= count; columnIndex++)
            {
                Excel.ListColumn column = null;
                try
                {
                    column = columns[columnIndex] as Excel.ListColumn;
                    if (column != null &&
                        string.Equals(column.Name, token, StringComparison.OrdinalIgnoreCase))
                    {
                        Excel.ListColumn found = column;
                        column = null;
                        return found;
                    }
                }
                finally
                {
                    ComHelper.Release(column);
                }
            }
            return null;
        }

        private static bool TryMapTotalsCalculation(string text, out Excel.XlTotalsCalculation calculation)
        {
            calculation = Excel.XlTotalsCalculation.xlTotalsCalculationSum;
            switch ((text ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "":
                case "sum":
                    calculation = Excel.XlTotalsCalculation.xlTotalsCalculationSum;
                    return true;
                case "count":
                    calculation = Excel.XlTotalsCalculation.xlTotalsCalculationCount;
                    return true;
                case "countnums":
                case "countnumbers":
                    calculation = Excel.XlTotalsCalculation.xlTotalsCalculationCountNums;
                    return true;
                case "average":
                case "avg":
                    calculation = Excel.XlTotalsCalculation.xlTotalsCalculationAverage;
                    return true;
                case "max":
                    calculation = Excel.XlTotalsCalculation.xlTotalsCalculationMax;
                    return true;
                case "min":
                    calculation = Excel.XlTotalsCalculation.xlTotalsCalculationMin;
                    return true;
                case "stddev":
                    calculation = Excel.XlTotalsCalculation.xlTotalsCalculationStdDev;
                    return true;
                case "var":
                case "variance":
                    calculation = Excel.XlTotalsCalculation.xlTotalsCalculationVar;
                    return true;
                case "none":
                    calculation = Excel.XlTotalsCalculation.xlTotalsCalculationNone;
                    return true;
                default:
                    return false;
            }
        }
    }
}
