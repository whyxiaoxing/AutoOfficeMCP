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
    /// <summary>P2 公式审计、单变量求解、批量重命名工作表。</summary>
    public sealed partial class ExcelExtensionsExecutor
    {
        private static OperationResult ExecuteFormulaAudit(Excel.Application application, JObject arguments)
        {
            string mode = ExcelArgumentReader.GetTrimmedString(arguments, "mode", "action", "operation")
                .ToLowerInvariant();
            if (mode.Length == 0)
            {
                mode = "scan";
            }
            string sheetName = ExcelArgumentReader.GetSheetName(arguments);
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange", "address");
            bool includeHidden = ExcelArgumentReader.GetBool(arguments, true, "includeHidden");
            bool includeAllFormulas = ExcelArgumentReader.GetBool(arguments, false, "includeAllFormulas");
            int maxResults = ClampInt(
                ExcelArgumentReader.GetInt(arguments, 200, "maxResults", "limit"),
                1,
                5000);

            Excel.Workbook workbook = null;
            List<Excel.Worksheet> worksheets = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);

                if (mode == "external" || mode == "externallinks" || mode == "links")
                {
                    JObject data = BuildExternalLinkReport(workbook);
                    string summary = "外部引用清单：" + data.Value<int>("externalLinkCount") +
                                     " 个链接，断链引用 " + data.Value<int>("brokenRefCount") + " 处。";
                    Observation observation = ReadOnlyObservation("formula_audit", summary);
                    return OperationResult.Succeed(summary, data, observation);
                }

                if (mode == "precedents" || mode == "dependents")
                {
                    if (address.Length == 0)
                    {
                        string baseSheetName = sheetName.Length > 0 ? sheetName : "usedrange";
                        address = sheetName.Length > 0
                            ? ExcelComSupport.QuoteSheetName(sheetName) + "!usedrange"
                            : "usedrange";
                        if (baseSheetName == "usedrange")
                        {
                            address = "usedrange";
                        }
                    }
                    using (ExcelRangeContext context = RangeResolver.Resolve(
                               application,
                               address,
                               false,
                               "excel_formula_audit"))
                    {
                        JArray entries = new JArray();
                        bool truncated = false;
                        AuditPrecedentsOrDependents(
                            context.Range,
                            mode == "precedents",
                            entries,
                            maxResults,
                            ref truncated);
                        JObject data = new JObject();
                        data["mode"] = mode;
                        data["sheet"] = context.SheetName;
                        data["address"] = context.Range.AddressLocal();
                        data["count"] = entries.Count;
                        data["truncated"] = truncated;
                        data["entries"] = entries;
                        string summary = "公式" + (mode == "precedents" ? "前例" : "依赖") +
                                         "追踪：共 " + entries.Count + " 个公式单元格。";
                        Observation observation = ReadOnlyObservation("formula_audit", summary);
                        return OperationResult.Succeed(summary, data, observation);
                    }
                }

                if (mode != "scan" && mode != "issues")
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "mode 仅支持 scan/precedents/dependents/external。");
                }

                JArray issues = new JArray();
                JArray sheets = new JArray();
                HashSet<string> externalLinks = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int formulaCells = 0;
                int brokenRefCount = 0;
                int errorFormulaCount = 0;
                bool truncatedByLimit = false;

                if (address.Length > 0)
                {
                    using (ExcelRangeContext context = RangeResolver.Resolve(
                               application,
                               address,
                               false,
                               "excel_formula_audit"))
                    {
                        sheets.Add(context.SheetName);
                        ScanFormulaRange(
                            context.SheetName,
                            context.Range,
                            includeAllFormulas,
                            maxResults,
                            issues,
                            externalLinks,
                            ref formulaCells,
                            ref brokenRefCount,
                            ref errorFormulaCount,
                            ref truncatedByLimit);
                    }
                }
                else
                {
                    worksheets = ExcelExtensionSupport.CollectWorksheets(workbook, includeHidden, string.Empty);
                    if (sheetName.Length > 0)
                    {
                        List<string> requested = new List<string>();
                        requested.Add(sheetName);
                        FilterWorksheetsByName(worksheets, requested);
                    }
                    if (worksheets.Count == 0)
                    {
                        return Fail(OfficeErrorCode.SHEET_NOT_FOUND, "没有找到可审计的工作表。");
                    }
                    foreach (Excel.Worksheet worksheet in worksheets)
                    {
                        Excel.Range usedRange = null;
                        try
                        {
                            usedRange = worksheet.UsedRange;
                            sheets.Add(worksheet.Name);
                            ScanFormulaRange(
                                worksheet.Name,
                                usedRange,
                                includeAllFormulas,
                                maxResults,
                                issues,
                                externalLinks,
                                ref formulaCells,
                                ref brokenRefCount,
                                ref errorFormulaCount,
                                ref truncatedByLimit);
                        }
                        finally
                        {
                            ComHelper.Release(usedRange);
                        }
                        if (issues.Count >= maxResults)
                        {
                            truncatedByLimit = true;
                            break;
                        }
                    }
                }

                JArray linkArray = new JArray();
                foreach (string link in externalLinks)
                {
                    linkArray.Add(link);
                }
                JObject resultData = new JObject();
                resultData["mode"] = "scan";
                resultData["sheets"] = sheets;
                resultData["formulaCells"] = formulaCells;
                resultData["issueCount"] = issues.Count;
                resultData["brokenRefCount"] = brokenRefCount;
                resultData["errorFormulaCount"] = errorFormulaCount;
                resultData["externalLinkCount"] = linkArray.Count;
                resultData["externalLinks"] = linkArray;
                resultData["truncated"] = truncatedByLimit;
                resultData["issues"] = issues;

                string message = "公式审计完成：公式 " + formulaCells + " 个，断链引用 " +
                                 brokenRefCount + " 处，错误公式 " + errorFormulaCount +
                                 " 个，外部链接 " + linkArray.Count + " 个。";
                Observation auditObservation = ReadOnlyObservation("formula_audit", message);
                return OperationResult.Succeed(message, resultData, auditObservation);
            }
            finally
            {
                ExcelExtensionSupport.ReleaseWorksheets(worksheets);
                ComHelper.Release(workbook);
            }
        }

        private static void ScanFormulaRange(
            string sheetName,
            Excel.Range range,
            bool includeAllFormulas,
            int maxResults,
            JArray issues,
            HashSet<string> externalLinks,
            ref int formulaCells,
            ref int brokenRefCount,
            ref int errorFormulaCount,
            ref bool truncated)
        {
            if (range == null || issues.Count >= maxResults)
            {
                return;
            }
            int rows = range.Rows.Count;
            int columns = range.Columns.Count;
            long cells = (long)rows * columns;
            if (cells > 500000)
            {
                truncated = true;
                return;
            }

            object[,] formulas = ExcelExtensionSupport.ReadFormulaMatrix(range, rows, columns);
            object[,] values = ExcelExtensionSupport.ReadMatrix(range, rows, columns);
            Regex externalPattern = new Regex(@"\[[^\]]+\]", RegexOptions.CultureInvariant);

            for (int row = 0; row < rows; row++)
            {
                for (int column = 0; column < columns; column++)
                {
                    string formula = ExcelExtensionSupport.CellText(formulas[row, column]);
                    if (formula.Length == 0 || !formula.StartsWith("=", StringComparison.Ordinal))
                    {
                        continue;
                    }
                    formulaCells++;
                    bool broken = formula.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0;
                    if (broken)
                    {
                        brokenRefCount++;
                    }
                    Match externalMatch = externalPattern.Match(formula);
                    if (externalMatch.Success)
                    {
                        foreach (Match match in externalPattern.Matches(formula))
                        {
                            externalLinks.Add(match.Value);
                        }
                    }
                    string errorText = ExcelExtensionSupport.ErrorTextFromValue(values[row, column]);
                    if (errorText.Length > 0)
                    {
                        errorFormulaCount++;
                    }

                    bool issue = broken || externalMatch.Success || errorText.Length > 0;
                    if (!issue && !includeAllFormulas)
                    {
                        continue;
                    }
                    if (issues.Count >= maxResults)
                    {
                        truncated = true;
                        return;
                    }

                    JObject entry = new JObject();
                    entry["sheet"] = sheetName;
                    entry["address"] = ExcelExtensionSupport.FullCellAddress(
                        sheetName,
                        range.Row + row,
                        range.Column + column);
                    entry["formula"] = formula;
                    entry["brokenRef"] = broken;
                    entry["externalRef"] = externalMatch.Success ? externalMatch.Value : string.Empty;
                    entry["error"] = errorText;
                    issues.Add(entry);
                }
            }
        }

        private static JObject BuildExternalLinkReport(Excel.Workbook workbook)
        {
            JArray links = new JArray();
            object rawLinks = null;
            try
            {
                rawLinks = workbook.LinkSources(Excel.XlLink.xlExcelLinks);
                Array array = rawLinks as Array;
                if (array != null)
                {
                    foreach (object item in array)
                    {
                        if (item != null)
                        {
                            links.Add(Convert.ToString(item, CultureInfo.InvariantCulture));
                        }
                    }
                }
            }
            catch (COMException)
            {
                rawLinks = null;
            }
            finally
            {
                ComHelper.Release(rawLinks);
            }

            JArray brokenReferences = new JArray();
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
                        string refersTo = Convert.ToString(
                            nameObject.RefersTo,
                            CultureInfo.InvariantCulture) ?? string.Empty;
                        if (refersTo.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            JObject entry = new JObject();
                            entry["kind"] = "name";
                            entry["name"] = nameObject.Name;
                            entry["refersTo"] = refersTo;
                            brokenReferences.Add(entry);
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

            JObject data = new JObject();
            data["externalLinkCount"] = links.Count;
            data["externalLinks"] = links;
            data["brokenRefCount"] = brokenReferences.Count;
            data["brokenReferences"] = brokenReferences;
            return data;
        }

        private static void AuditPrecedentsOrDependents(
            Excel.Range range,
            bool precedents,
            JArray entries,
            int maxResults,
            ref bool truncated)
        {
            int rows = range.Rows.Count;
            int columns = range.Columns.Count;
            long cells = (long)rows * columns;
            if (cells > 500000)
            {
                truncated = true;
                return;
            }

            object[,] formulas = ExcelExtensionSupport.ReadFormulaMatrix(range, rows, columns);
            for (int row = 0; row < rows; row++)
            {
                for (int column = 0; column < columns; column++)
                {
                    string formula = ExcelExtensionSupport.CellText(formulas[row, column]);
                    if (formula.Length == 0 || !formula.StartsWith("=", StringComparison.Ordinal))
                    {
                        continue;
                    }
                    if (entries.Count >= maxResults)
                    {
                        truncated = true;
                        return;
                    }

                    Excel.Range cell = null;
                    Excel.Range references = null;
                    try
                    {
                        cell = (Excel.Range)range.Cells[row + 1, column + 1];
                        references = precedents ? cell.DirectPrecedents : cell.DirectDependents;
                        JArray referenceList = new JArray();
                        if (references != null)
                        {
                            Excel.Areas areas = null;
                            try
                            {
                                areas = references.Areas;
                                int areaCount = areas.Count;
                                for (int index = 1; index <= areaCount; index++)
                                {
                                    Excel.Range area = null;
                                    Excel.Worksheet areaSheet = null;
                                    try
                                    {
                                        area = (Excel.Range)areas[index];
                                        areaSheet = area.Worksheet as Excel.Worksheet;
                                        referenceList.Add(ExcelComSupport.FullAddress(
                                            areaSheet == null ? string.Empty : areaSheet.Name,
                                            area.AddressLocal()));
                                    }
                                    finally
                                    {
                                        ComHelper.Release(area, areaSheet);
                                    }
                                }
                            }
                            finally
                            {
                                ComHelper.Release(areas);
                            }
                        }

                        JObject entry = new JObject();
                        Excel.Worksheet cellSheet = null;
                        try
                        {
                            cellSheet = cell.Worksheet as Excel.Worksheet;
                            entry["address"] = ExcelComSupport.FullAddress(
                                cellSheet == null ? string.Empty : cellSheet.Name,
                                cell.AddressLocal());
                        }
                        finally
                        {
                            ComHelper.Release(cellSheet);
                        }
                        entry["formula"] = formula;
                        entry["references"] = referenceList;
                        entry["referenceCount"] = referenceList.Count;
                        entries.Add(entry);
                    }
                    catch (COMException)
                    {
                        JObject entry = new JObject();
                        entry["formula"] = formula;
                        entry["references"] = new JArray();
                        entry["referenceCount"] = 0;
                        entry["note"] = "Excel 未能返回该单元格的前例/依赖（可能跨工作簿或被保护）。";
                        entries.Add(entry);
                    }
                    finally
                    {
                        ComHelper.Release(references, cell);
                    }
                }
            }
        }

        private static OperationResult ExecuteGoalSeek(Excel.Application application, JObject arguments)
        {
            string targetAddress = ExcelArgumentReader.GetRangeAddress(
                arguments,
                "targetCell",
                "targetRange",
                "target");
            if (string.IsNullOrWhiteSpace(targetAddress))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 targetCell 参数。");
            }
            string changingAddress = ExcelArgumentReader.GetTrimmedString(
                arguments,
                "changingCell",
                "byChangingCell",
                "variableCell");
            if (changingAddress.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 changingCell 参数。");
            }
            double goalValue;
            if (!ExcelArgumentReader.TryGetDouble(arguments, out goalValue, "goalValue", "goal", "targetValue"))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 goalValue 参数（数值）。");
            }

            Excel.Workbook workbook = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                using (ExcelRangeContext target = RangeResolver.Resolve(
                           application,
                           targetAddress,
                           false,
                           "excel_goal_seek"))
                {
                    int bang = changingAddress.LastIndexOf('!');
                    ExcelRangeContext changing = null;
                    if (bang >= 0)
                    {
                        changing = RangeResolver.Resolve(
                            application,
                            changingAddress,
                            false,
                            "excel_goal_seek");
                    }
                    else
                    {
                        changing = RangeResolver.ResolveOnSheet(
                            application,
                            target.SheetName,
                            changingAddress,
                            false,
                            "excel_goal_seek");
                    }

                    using (changing)
                    {
                        if (Convert.ToInt64(
                                target.Range.Cells.CountLarge,
                                CultureInfo.InvariantCulture) > 1)
                        {
                            return Fail(OfficeErrorCode.INVALID_PARAMS, "targetCell 必须是单个单元格。");
                        }
                        if (Convert.ToInt64(
                                changing.Range.Cells.CountLarge,
                                CultureInfo.InvariantCulture) > 1)
                        {
                            return Fail(OfficeErrorCode.INVALID_PARAMS, "changingCell 必须是单个单元格。");
                        }
                        if (!string.Equals(
                                target.Workbook.FullName,
                                changing.Workbook.FullName,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            return Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "targetCell 与 changingCell 必须在同一工作簿中。");
                        }
                        ExcelComSupport.EnsureWritable(
                            application,
                            target.Workbook,
                            target.Worksheet,
                            "excel_goal_seek");
                        ExcelComSupport.EnsureWritable(
                            application,
                            changing.Workbook,
                            changing.Worksheet,
                            "excel_goal_seek");

                        object beforeValue = changing.Range.Value2;
                        string changingBefore = ExcelExtensionSupport.CellText(beforeValue);
                        ExcelUndoScope undo = ExcelUndoScope.Capture(
                            changing.Workbook,
                            changing.Worksheet,
                            changing.Range,
                            changing.Range.AddressLocal());

                        bool converged = target.Range.GoalSeek(goalValue, changing.Range);

                        object afterValue = changing.Range.Value2;
                        string changingAfter = ExcelExtensionSupport.CellText(afterValue);
                        object targetValue = target.Range.Value2;
                        string targetAfter = ExcelExtensionSupport.CellText(targetValue);
                        bool changed = !string.Equals(changingBefore, changingAfter, StringComparison.Ordinal);

                        string targetRef = target.FullAddress;
                        Observation observation = Observation.Create(
                            "goal_seek",
                            converged
                                ? "单变量求解成功：" + targetRef + " = " + goalValue + "。"
                                : "单变量求解未收敛。",
                            changed);
                        if (changed)
                        {
                            observation.AddTargetRef(changing.FullAddress);
                        }

                        JObject data = new JObject();
                        data["targetCell"] = targetRef;
                        data["goalValue"] = goalValue;
                        data["changingCell"] = changing.FullAddress;
                        data["converged"] = converged;
                        data["changingValueBefore"] = changingBefore;
                        data["changingValueAfter"] = changingAfter;
                        data["targetValueAfter"] = targetAfter;

                        undo.Complete(
                            changing.Workbook,
                            changing.Worksheet,
                            changing.Range,
                            changing.Range.AddressLocal(),
                            true,
                            "单变量求解 " + targetRef,
                            observation,
                            data,
                            converged ? null : "目标未收敛，变量单元格可能停留在最后迭代值。");

                        if (!converged)
                        {
                            return OperationResult.PartialApplied(
                                changed ? 1 : 0,
                                "单变量求解未收敛；变量单元格可能已改变。请检查目标公式/初始值，" +
                                "或使用返回的 undoUnits 恢复变量单元格。",
                                observation,
                                data);
                        }
                        return OperationResult.Succeed(
                            "单变量求解完成：" + targetRef + " 达到 " + targetAfter + "。",
                            data,
                            observation);
                    }
                }
            }
            finally
            {
                ComHelper.Release(workbook);
            }
        }

        private static OperationResult ExecuteBatchRenameSheets(Excel.Application application, JObject arguments)
        {
            OperationResult denied = SecurityGate.CreateDenial(
                SecurityRisk.DeleteSheet,
                "excel_batch_rename_sheets");
            if (denied != null)
            {
                return denied;
            }

            string mode = ExcelArgumentReader.GetTrimmedString(arguments, "mode").ToLowerInvariant();
            if (mode.Length == 0)
            {
                mode = "regex";
            }
            bool dryRun = ExcelArgumentReader.GetBool(arguments, false, "dryRun", "preview");
            List<string> requestedSheets = ExcelArgumentReader.GetStringList(arguments, "sheets");

            Excel.Workbook workbook = null;
            List<Excel.Worksheet> worksheets = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                worksheets = ExcelExtensionSupport.CollectWorksheets(workbook, true, string.Empty);
                FilterWorksheetsByName(worksheets, requestedSheets);
                if (worksheets.Count == 0)
                {
                    return Fail(OfficeErrorCode.SHEET_NOT_FOUND, "没有找到可重命名的工作表。");
                }

                List<SheetRenamePlan> plans = new List<SheetRenamePlan>();
                List<string> errors = new List<string>();
                if (mode == "mapping" || mode == "map")
                {
                    JArray mapping = ExcelArgumentReader.GetArray(arguments, "mapping", "renames");
                    if (mapping == null || mapping.Count == 0)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "mapping 模式需要 mapping 数组。");
                    }
                    foreach (JToken token in mapping)
                    {
                        JObject item = token as JObject;
                        if (item == null)
                        {
                            errors.Add("mapping 中存在非对象项。");
                            continue;
                        }
                        string oldName = ExcelArgumentReader.GetTrimmedString(item, "oldName", "name", "sheet");
                        string newName = ExcelArgumentReader.GetTrimmedString(item, "newName", "to");
                        Excel.Worksheet worksheet = FindWorksheetInList(worksheets, oldName);
                        if (worksheet == null)
                        {
                            errors.Add("找不到工作表「" + oldName + "」。");
                            continue;
                        }
                        SheetRenamePlan plan = new SheetRenamePlan();
                        plan.Worksheet = worksheet;
                        plan.OldName = worksheet.Name;
                        plan.NewName = newName;
                        plans.Add(plan);
                    }
                }
                else if (mode == "regex" || mode == "pattern")
                {
                    string pattern = ExcelArgumentReader.GetString(arguments, "pattern", "regex");
                    if (pattern.Length == 0)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "regex 模式需要 pattern 参数。");
                    }
                    string replacement = ExcelArgumentReader.GetString(arguments, "replacement", "replace");
                    Regex regex;
                    try
                    {
                        regex = new Regex(pattern, RegexOptions.CultureInvariant);
                    }
                    catch (ArgumentException ex)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, ExcelExtensionSupport.ToRegexPatternError(ex));
                    }
                    foreach (Excel.Worksheet worksheet in worksheets)
                    {
                        SheetRenamePlan plan = new SheetRenamePlan();
                        plan.Worksheet = worksheet;
                        plan.OldName = worksheet.Name;
                        plan.NewName = regex.Replace(worksheet.Name, replacement);
                        plans.Add(plan);
                    }
                }
                else
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "mode 仅支持 regex/mapping。");
                }

                foreach (SheetRenamePlan plan in plans)
                {
                    if (string.IsNullOrWhiteSpace(plan.NewName))
                    {
                        errors.Add("工作表「" + plan.OldName + "」的新名称为空。");
                        continue;
                    }
                    string nameError = ValidateSheetNameForRename(plan.NewName);
                    if (nameError.Length > 0)
                    {
                        errors.Add("工作表「" + plan.OldName + "」：" + nameError);
                    }
                }

                HashSet<string> finalNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (SheetRenamePlan plan in plans)
                {
                    if (string.IsNullOrWhiteSpace(plan.NewName))
                    {
                        continue;
                    }
                    if (!finalNames.Add(plan.NewName))
                    {
                        errors.Add("新名称「" + plan.NewName + "」重复。");
                    }
                }

                if (errors.Count == 0)
                {
                    foreach (SheetRenamePlan plan in plans)
                    {
                        if (string.Equals(plan.OldName, plan.NewName, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }
                        Excel.Worksheet conflict = FindWorksheetInList(worksheets, plan.NewName);
                        if (conflict != null &&
                            !string.Equals(conflict.Name, plan.OldName, StringComparison.OrdinalIgnoreCase))
                        {
                            errors.Add(
                                "新名称「" + plan.NewName + "」已被现有工作表「" +
                                conflict.Name + "」占用（不支持互换名称，请分两步执行）。");
                        }
                    }
                }

                if (errors.Count > 0)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "批量重命名校验失败，未修改任何工作表：" + string.Join("；", errors.ToArray()));
                }

                JArray proposed = new JArray();
                foreach (SheetRenamePlan plan in plans)
                {
                    JObject item = new JObject();
                    item["oldName"] = plan.OldName;
                    item["newName"] = plan.NewName;
                    item["changed"] = !string.Equals(
                        plan.OldName,
                        plan.NewName,
                        StringComparison.Ordinal);
                    proposed.Add(item);
                }

                if (dryRun)
                {
                    JObject dryData = new JObject();
                    dryData["mode"] = mode;
                    dryData["dryRun"] = true;
                    dryData["renames"] = proposed;
                    ExcelUndoScope.RecordUnavailable(dryData, "dryRun 预览未修改任何工作表，无需撤销。");
                    Observation dryObservation = ReadOnlyObservation(
                        "batch_rename_sheets",
                        "批量重命名预览：共 " + plans.Count + " 个工作表。");
                    return OperationResult.Succeed("批量重命名预览（未修改）。", dryData, dryObservation);
                }

                if (workbook.ReadOnly)
                {
                    return Fail(OfficeErrorCode.READ_ONLY, "工作簿以只读方式打开，无法重命名工作表。");
                }

                List<SheetRenamePlan> applied = new List<SheetRenamePlan>();
                try
                {
                    application.Guard(delegate
                    {
                        foreach (SheetRenamePlan plan in plans)
                        {
                            if (string.Equals(plan.OldName, plan.NewName, StringComparison.Ordinal))
                            {
                                continue;
                            }
                            plan.Worksheet.Name = plan.NewName;
                            applied.Add(plan);
                        }
                    });
                }
                catch (Exception ex)
                {
                    JObject partialData = new JObject();
                    partialData["mode"] = mode;
                    partialData["dryRun"] = false;
                    partialData["renames"] = proposed;
                    JArray appliedArray = new JArray();
                    foreach (SheetRenamePlan plan in applied)
                    {
                        JObject item = new JObject();
                        item["oldName"] = plan.OldName;
                        item["newName"] = plan.NewName;
                        appliedArray.Add(item);
                    }
                    partialData["applied"] = appliedArray;
                    ExcelUndoScope.RecordUnavailable(
                        partialData,
                        "工作表重命名属于结构变更，区域快照无法撤销。");
                    Observation partialObservation = Observation.Create(
                        "batch_rename_sheets",
                        "批量重命名中途失败：已完成 " + applied.Count + " 个。",
                        applied.Count > 0);
                    return OperationResult.PartialApplied(
                        applied.Count,
                        "重命名失败：" + ex.Message +
                        "。已重命名的工作表不会自动回滚，请按 applied 清单人工改回。",
                        partialObservation,
                        partialData);
                }

                JObject data = new JObject();
                data["mode"] = mode;
                data["dryRun"] = false;
                data["renamedCount"] = applied.Count;
                data["renames"] = proposed;
                ExcelUndoScope.RecordUnavailable(
                    data,
                    "工作表重命名属于结构变更，区域快照无法撤销。");
                Observation observation = Observation.Create(
                    "batch_rename_sheets",
                    "批量重命名完成：已重命名 " + applied.Count + " 个工作表。",
                    applied.Count > 0);
                foreach (SheetRenamePlan plan in applied)
                {
                    observation.AddTargetRef(plan.NewName);
                }
                return OperationResult.Succeed(
                    "批量重命名完成：已重命名 " + applied.Count + " 个工作表。",
                    data,
                    observation);
            }
            finally
            {
                ExcelExtensionSupport.ReleaseWorksheets(worksheets);
                ComHelper.Release(workbook);
            }
        }

        private sealed class SheetRenamePlan
        {
            public Excel.Worksheet Worksheet { get; set; }

            public string OldName { get; set; }

            public string NewName { get; set; }
        }

        private static Excel.Worksheet FindWorksheetInList(
            List<Excel.Worksheet> worksheets,
            string name)
        {
            foreach (Excel.Worksheet worksheet in worksheets)
            {
                if (string.Equals(worksheet.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return worksheet;
                }
            }
            return null;
        }

        private static string ValidateSheetNameForRename(string name)
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
