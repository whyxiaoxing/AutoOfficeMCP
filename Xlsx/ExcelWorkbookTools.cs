using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Newtonsoft.Json.Linq;
using Office;
using Office.Config;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>P2 工作簿生命周期：新建、副本另存与工作簿/工作表保护。</summary>
    public sealed partial class ExcelToolExecutor
    {
        private static readonly string[] ProtectionOptionNames = new string[]
        {
            "AllowFiltering",
            "AllowSorting",
            "AllowFormattingCells",
            "AllowFormattingColumns",
            "AllowFormattingRows",
            "AllowInsertingColumns",
            "AllowInsertingRows",
            "AllowInsertingHyperlinks",
            "AllowDeletingColumns",
            "AllowDeletingRows",
            "AllowUsingPivotTables"
        };

        private static OperationResult ExecuteCreateWorkbook(Excel.Application application, JObject arguments)
        {
            string rawPath = ExcelArgumentReader.GetTrimmedString(
                arguments,
                "path",
                "filePath",
                "saveAs",
                "outputPath");
            if (rawPath.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 path/filePath 参数。");
            }
            string fullPath;
            string pathError;
            if (!WorkspacePathResolver.TryResolve(rawPath, out fullPath, out pathError))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, pathError);
            }
            int fileFormat = WorkbookFileFormat(fullPath);
            if (fileFormat == 0)
            {
                return Fail(
                    OfficeErrorCode.INVALID_PARAMS,
                    "新建工作簿只支持 .xlsx/.xlsm/.xls/.csv，收到：" + fullPath);
            }

            bool overwrite = ExcelArgumentReader.GetBool(arguments, false, "overwrite");
            bool openAfterCreate = ExcelArgumentReader.GetBool(arguments, true, "open", "openAfterCreate");
            string templateRaw = ExcelArgumentReader.GetTrimmedString(arguments, "template", "templatePath");
            string templatePath = string.Empty;
            if (templateRaw.Length > 0)
            {
                string templateError;
                if (!WorkspacePathResolver.TryResolve(templateRaw, out templatePath, out templateError))
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "template 参数无效：" + templateError);
                }
                if (!File.Exists(templatePath))
                {
                    return Fail(OfficeErrorCode.FILE_NOT_FOUND, "模板文件不存在：" + templatePath);
                }
            }

            List<string> sheetNames = ExcelArgumentReader.GetStringList(arguments, "sheets", "sheetNames", "initialSheets");
            for (int index = 0; index < sheetNames.Count; index++)
            {
                string nameError = ValidateSheetName(sheetNames[index]);
                if (nameError.Length > 0)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "sheets[" + index.ToString(CultureInfo.InvariantCulture) + "]：" + nameError);
                }
                for (int other = 0; other < index; other++)
                {
                    if (string.Equals(sheetNames[other], sheetNames[index], StringComparison.OrdinalIgnoreCase))
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "sheets 中存在重复名称：" + sheetNames[index]);
                    }
                }
            }

            if (File.Exists(fullPath) && !overwrite)
            {
                Excel.Workbook existing = FindOpenWorkbook(application, fullPath);
                if (existing != null)
                {
                    JObject existingData = BuildCreatedWorkbookData(existing, fullPath, false, true, templatePath);
                    ComHelper.Release(existing);
                    Observation existingObservation = Observation.Create(
                        "create_workbook",
                        "目标工作簿已打开，未重复创建：" + fullPath,
                        false);
                    existingObservation.AddTargetRef(fullPath);
                    ExcelUndoScope.RecordUnavailable(existingData, "文件已存在且已打开，未做修改。");
                    return OperationResult.Succeed(
                        "目标工作簿已打开，未重复创建：" + fullPath,
                        existingData,
                        existingObservation);
                }
                if (!openAfterCreate)
                {
                    JObject closedData = new JObject();
                    closedData["path"] = fullPath;
                    closedData["created"] = false;
                    closedData["opened"] = false;
                    closedData["note"] = "文件已存在；overwrite=false 且 open=false，未打开也未修改。";
                    ExcelUndoScope.RecordUnavailable(closedData, "未做修改。");
                    return OperationResult.Succeed(
                        "目标文件已存在，按参数未创建/打开。",
                        closedData,
                        Observation.Create("create_workbook", "目标文件已存在，未修改。", false));
                }
                Excel.Workbook opened = null;
                try
                {
                    application.Guard(delegate { opened = application.Workbooks.Open(fullPath); });
                    JObject openedData = BuildCreatedWorkbookData(opened, fullPath, false, true, templatePath);
                    Observation openedObservation = Observation.Create(
                        "create_workbook",
                        "目标文件已存在，已打开现有工作簿：" + fullPath,
                        false);
                    openedObservation.AddTargetRef(fullPath);
                    ExcelUndoScope.RecordUnavailable(openedData, "打开现有工作簿属于只读生命周期操作，未修改内容。");
                    return OperationResult.Succeed(
                        "目标文件已存在，已打开现有工作簿：" + fullPath,
                        openedData,
                        openedObservation);
                }
                finally
                {
                    ComHelper.Release(opened);
                }
            }

            if (File.Exists(fullPath) && overwrite)
            {
                Excel.Workbook openConflict = FindOpenWorkbook(application, fullPath);
                if (openConflict != null)
                {
                    ComHelper.Release(openConflict);
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "目标工作簿当前已在 Excel 中打开，覆盖前请先关闭：" + fullPath);
                }
            }

            string directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                string directoryError;
                if (!WorkspacePathResolver.TryEnsureOutputDirectory(directory, out directoryError))
                {
                    return Fail(OfficeErrorCode.READ_ONLY, directoryError);
                }
            }

            Excel.Workbook created = null;
            bool saved = false;
            List<string> createdSheetNames = null;
            try
            {
                application.Guard(delegate
                {
                    created = templatePath.Length > 0
                        ? application.Workbooks.Add(templatePath)
                        : application.Workbooks.Add(Type.Missing);
                    if (sheetNames.Count > 0)
                    {
                        ApplyInitialSheets(created, sheetNames, templatePath.Length > 0);
                    }
                    created.SaveAs(fullPath, fileFormat);
                    saved = true;
                    createdSheetNames = GetWorksheetNames(created, sheetNames, true);
                    if (!openAfterCreate)
                    {
                        created.Close(false);
                        created = null;
                    }
                });

                JObject data = new JObject();
                data["path"] = fullPath;
                data["created"] = true;
                data["opened"] = openAfterCreate;
                data["template"] = templatePath;
                data["fileFormat"] = fileFormat;
                data["sheets"] = new JArray(createdSheetNames == null ? sheetNames : createdSheetNames);
                data["bytes"] = File.Exists(fullPath) ? new FileInfo(fullPath).Length : 0;
                data["note"] = fileFormat == 6
                    ? "CSV 保存只导出活动工作表；工作簿对象保留在内存中。"
                    : string.Empty;
                ExcelUndoScope.RecordUnavailable(
                    data,
                    "新建工作簿属于文件生命周期操作，不在区域快照撤销范围内；请直接关闭/删除文件。");

                Observation observation = Observation.Create(
                    "create_workbook",
                    "已新建工作簿：" + fullPath + (openAfterCreate ? "（已打开）" : "（已保存未打开）"),
                    true);
                observation.AddTargetRef(fullPath);
                return OperationResult.Succeed(
                    "已新建工作簿：" + fullPath + (openAfterCreate ? "。" : "（已保存，未保持打开）。"),
                    data,
                    observation);
            }
            catch
            {
                if (created != null && !saved)
                {
                    try
                    {
                        created.Close(false);
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[excel_create_workbook] 创建失败后关闭工作簿失败：" + ex.Message);
                    }
                }
                throw;
            }
            finally
            {
                ComHelper.Release(created);
            }
        }

        private static OperationResult ExecuteSaveAs(Excel.Application application, JObject arguments)
        {
            string rawPath = ExcelArgumentReader.GetTrimmedString(
                arguments,
                "path",
                "filePath",
                "saveAs",
                "targetPath");
            if (rawPath.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 path/filePath 参数。");
            }
            string fullPath;
            string pathError;
            if (!WorkspacePathResolver.TryResolve(rawPath, out fullPath, out pathError))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, pathError);
            }
            int fileFormat = WorkbookFileFormat(fullPath);
            if (fileFormat == 0)
            {
                return Fail(
                    OfficeErrorCode.INVALID_PARAMS,
                    "save_as 只支持 .xlsx/.xlsm/.xls/.csv，收到：" + fullPath);
            }

            Excel.Workbook source = null;
            try
            {
                source = ExcelComSupport.RequireActiveWorkbook(application);
                string sourceFullName = source.FullName;
                if (string.Equals(sourceFullName, fullPath, StringComparison.OrdinalIgnoreCase))
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "save_as 目标与当前工作簿相同，请使用 office_save_document 保存原文件。");
                }
                bool overwrite = ExcelArgumentReader.GetBool(arguments, false, "overwrite");
                bool existedBefore = File.Exists(fullPath);
                if (existedBefore && !overwrite)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "目标文件已存在：" + fullPath + "。如需覆盖请传 overwrite=true。");
                }
                string directory = Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    string directoryError;
                    if (!WorkspacePathResolver.TryEnsureOutputDirectory(directory, out directoryError))
                    {
                        return Fail(OfficeErrorCode.READ_ONLY, directoryError);
                    }
                }

                string method = string.Empty;
                string saveError = string.Empty;
                string sourceExtension = Path.GetExtension(sourceFullName);
                string targetExtension = Path.GetExtension(fullPath);
                bool sameFormat = string.Equals(
                    sourceExtension,
                    targetExtension,
                    StringComparison.OrdinalIgnoreCase);
                if (sameFormat && fileFormat != 6)
                {
                    try
                    {
                        application.Guard(delegate { source.SaveCopyAs(fullPath); });
                        method = "SaveCopyAs";
                    }
                    catch (Exception ex)
                    {
                        saveError = ex.Message;
                        Console.Error.WriteLine("[excel_save_as] SaveCopyAs 失败，改用副本 SaveAs：" + ex.Message);
                    }
                }

                if (method.Length == 0)
                {
                    Excel.Workbook copy = null;
                    try
                    {
                        string activeSheetName = string.Empty;
                        object activeObject = source.ActiveSheet;
                        Excel.Worksheet activeSheet = activeObject as Excel.Worksheet;
                        if (activeSheet != null)
                        {
                            activeSheetName = activeSheet.Name;
                        }
                        ComHelper.Release(activeSheet, activeObject);

                        application.Guard(delegate
                        {
                            copy = application.Workbooks.Add(Type.Missing);
                            Excel.Sheets sourceSheets = null;
                            Excel.Sheets copySheets = null;
                            Excel.Worksheet defaultSheet = null;
                            Excel.Worksheet activeCopySheet = null;
                            try
                            {
                                sourceSheets = source.Sheets;
                                copySheets = copy.Worksheets;
                                sourceSheets.Copy(After: copySheets[copySheets.Count]);
                                defaultSheet = (Excel.Worksheet)copy.Worksheets[1];
                                defaultSheet.Delete();
                                if (fileFormat == 6 && activeSheetName.Length > 0)
                                {
                                    activeCopySheet = ExcelComSupport.FindWorksheet(copy, activeSheetName);
                                    if (activeCopySheet != null)
                                    {
                                        activeCopySheet.Activate();
                                    }
                                }
                                copy.SaveAs(fullPath, fileFormat);
                            }
                            finally
                            {
                                ComHelper.Release(activeCopySheet, defaultSheet, copySheets, sourceSheets);
                            }
                            string copyName = copy.Name;
                            copy.Close(false);
                            copy = null;
                            source.Activate();
                            Console.Error.WriteLine("[excel_save_as] 已通过副本 SaveAs 生成 " + copyName);
                        });
                        method = "CopyThenSaveAs";
                    }
                    catch
                    {
                        if (copy != null)
                        {
                            try { copy.Close(false); } catch (Exception) { }
                        }
                        throw;
                    }
                    finally
                    {
                        ComHelper.Release(copy);
                    }
                }

                if (!File.Exists(fullPath))
                {
                    return Fail(OfficeErrorCode.INTERNAL_ERROR, "另存后目标文件不存在：" + fullPath);
                }
                if (!string.Equals(source.FullName, sourceFullName, StringComparison.OrdinalIgnoreCase))
                {
                    return Fail(
                        OfficeErrorCode.PARTIAL_APPLY,
                        "另存副本已生成，但原工作簿路径发生了变化（" + sourceFullName + " -> " +
                        source.FullName + "）；请人工检查。");
                }

                source.Activate();
                long bytes = new FileInfo(fullPath).Length;
                JObject data = new JObject();
                data["sourceFullName"] = sourceFullName;
                data["path"] = fullPath;
                data["format"] = DescribeWorkbookFormat(fileFormat, fullPath);
                data["method"] = method;
                data["bytes"] = bytes;
                data["originalUnchanged"] = true;
                data["overwritten"] = existedBefore;
                if (saveError.Length > 0)
                {
                    data["saveCopyAsError"] = saveError;
                }
                if (fileFormat == 6)
                {
                    data["note"] = "CSV 另存只导出当前活动工作表。";
                }
                ExcelUndoScope.RecordUnavailable(
                    data,
                    "另存副本不修改源工作簿内容/路径，不登记自动撤销；如需清理请删除目标文件。");

                Observation observation = Observation.Create(
                    "save_as",
                    "已另存副本：" + fullPath + "（原文件保持不变）。",
                    false);
                observation.AddTargetRef(fullPath);
                return OperationResult.Succeed(
                    "已另存副本：" + fullPath + "，原工作簿未被修改。",
                    data,
                    observation);
            }
            finally
            {
                ComHelper.Release(source);
            }
        }

        private static OperationResult ExecuteGetProtectionStatus(Excel.Application application, JObject arguments)
        {
            Excel.Workbook workbook = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                string sheetName = ExcelArgumentReader.GetSheetName(arguments);
                JObject data = BuildWorkbookProtectionData(application, workbook, sheetName);

                Observation observation = Observation.Create(
                    "get_protection_status",
                    "已读取 " + workbook.Name + " 的保护状态。",
                    false);
                observation.AddTargetRef(workbook.FullName);
                return OperationResult.Succeed(
                    "已读取工作簿「" + workbook.Name + "」的保护状态。",
                    data,
                    observation);
            }
            finally
            {
                ComHelper.Release(workbook);
            }
        }

        private static OperationResult ExecuteProtectWorkbook(Excel.Application application, JObject arguments)
        {
            string sheetName = ExcelArgumentReader.GetSheetName(arguments);
            bool unprotect = ExcelArgumentReader.GetBool(arguments, false, "unprotect");
            bool protect = ExcelArgumentReader.GetBool(arguments, true, "protect");
            if (unprotect)
            {
                protect = false;
            }
            string password = ExcelArgumentReader.GetString(arguments, "password");

            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                if (sheetName.Length > 0)
                {
                    worksheet = ExcelComSupport.FindWorksheet(workbook, sheetName);
                    if (worksheet == null)
                    {
                        return Fail(OfficeErrorCode.SHEET_NOT_FOUND, "找不到工作表「" + sheetName + "」。");
                    }
                    return ProtectSingleWorksheet(application, workbook, worksheet, arguments, protect, password);
                }

                bool structure = ExcelArgumentReader.GetBool(arguments, true, "structure", "protectStructure");
                bool windows = ExcelArgumentReader.GetBool(arguments, false, "windows", "protectWindows");
                bool wasStructureProtected = workbook.ProtectStructure;
                bool wasWindowsProtected = workbook.ProtectWindows;
                if (!protect)
                {
                    if (!wasStructureProtected && !wasWindowsProtected)
                    {
                        JObject unchanged = BuildWorkbookProtectionData(application, workbook, string.Empty);
                        ExcelUndoScope.RecordUnavailable(unchanged, "工作簿当前未受保护，无需取消。");
                        return OperationResult.Succeed(
                            "工作簿当前未受结构/窗口保护，未做修改。",
                            unchanged,
                            Observation.Create("protect_workbook", "工作簿当前未受保护，未修改。", false));
                    }
                    application.Guard(delegate { workbook.Unprotect(password); });
                }
                else
                {
                    if (!structure && !windows)
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "工作簿保护至少要启用 structure 或 windows 之一。");
                    }
                    if ((structure && wasStructureProtected) && (!windows || wasWindowsProtected))
                    {
                        JObject unchanged = BuildWorkbookProtectionData(application, workbook, string.Empty);
                        ExcelUndoScope.RecordUnavailable(unchanged, "请求的保护状态已存在，未重复保护。");
                        return OperationResult.Succeed(
                            "工作簿已处于请求的保护状态，未重复保护。",
                            unchanged,
                            Observation.Create("protect_workbook", "工作簿已处于请求的保护状态。", false));
                    }
                    application.Guard(delegate { workbook.Protect(password, structure, windows); });
                }

                JObject data = BuildWorkbookProtectionData(application, workbook, string.Empty);
                data["action"] = protect ? "protect" : "unprotect";
                data["changedStructure"] = wasStructureProtected != workbook.ProtectStructure;
                data["changedWindows"] = wasWindowsProtected != workbook.ProtectWindows;
                ExcelUndoScope.RecordUnavailable(
                    data,
                    "工作簿保护状态不在区域快照中，未登记自动撤销。");
                Observation observation = Observation.Create(
                    "protect_workbook",
                    protect ? "已保护工作簿（结构=" + structure + "，窗口=" + windows + "）。" : "已取消工作簿保护。",
                    protect || wasStructureProtected || wasWindowsProtected);
                observation.AddTargetRef(workbook.FullName);
                return OperationResult.Succeed(
                    protect ? "工作簿保护已启用。" : "工作簿保护已取消。",
                    data,
                    observation);
            }
            finally
            {
                ComHelper.Release(worksheet, workbook);
            }
        }

        private static OperationResult ProtectSingleWorksheet(
            Excel.Application application,
            Excel.Workbook workbook,
            Excel.Worksheet worksheet,
            JObject arguments,
            bool protect,
            string password)
        {
            bool wasProtected = worksheet.ProtectContents ||
                                worksheet.ProtectDrawingObjects ||
                                worksheet.ProtectScenarios;
            if (!protect)
            {
                if (!wasProtected)
                {
                    JObject unchanged = BuildWorkbookProtectionData(application, workbook, worksheet.Name);
                    ExcelUndoScope.RecordUnavailable(unchanged, "工作表当前未受保护，无需取消。");
                    return OperationResult.Succeed(
                        "工作表当前未受保护，未做修改。",
                        unchanged,
                        Observation.Create("protect_workbook", "工作表当前未受保护。", false));
                }
                application.Guard(delegate { worksheet.Unprotect(password); });
            }
            else
            {
                if (wasProtected)
                {
                    JObject unchanged = BuildWorkbookProtectionData(application, workbook, worksheet.Name);
                    ExcelUndoScope.RecordUnavailable(unchanged, "工作表已受保护，未重复保护；如需调整选项请先取消保护。");
                    return OperationResult.Succeed(
                        "工作表已受保护，未重复保护；如需调整细粒度选项请先取消保护。",
                        unchanged,
                        Observation.Create("protect_workbook", "工作表已受保护，未修改。", false));
                }

                bool drawingObjects = ExcelArgumentReader.GetBool(arguments, true, "drawingObjects");
                bool scenarios = ExcelArgumentReader.GetBool(arguments, false, "scenarios");
                bool allowFiltering = ExcelArgumentReader.GetBool(arguments, false, "allowFiltering");
                bool allowSorting = ExcelArgumentReader.GetBool(arguments, false, "allowSorting");
                bool allowFormattingCells = ExcelArgumentReader.GetBool(arguments, false, "allowFormattingCells");
                bool allowFormattingColumns = ExcelArgumentReader.GetBool(arguments, false, "allowFormattingColumns");
                bool allowFormattingRows = ExcelArgumentReader.GetBool(arguments, false, "allowFormattingRows");
                bool allowInsertingColumns = ExcelArgumentReader.GetBool(arguments, false, "allowInsertingColumns");
                bool allowInsertingRows = ExcelArgumentReader.GetBool(arguments, false, "allowInsertingRows");
                bool allowInsertingHyperlinks = ExcelArgumentReader.GetBool(arguments, false, "allowInsertingHyperlinks");
                bool allowDeletingColumns = ExcelArgumentReader.GetBool(arguments, false, "allowDeletingColumns");
                bool allowDeletingRows = ExcelArgumentReader.GetBool(arguments, false, "allowDeletingRows");
                bool allowUsingPivotTables = ExcelArgumentReader.GetBool(arguments, false, "allowUsingPivotTables");

                application.Guard(delegate
                {
                    worksheet.Protect(
                        Password: password,
                        DrawingObjects: drawingObjects,
                        Contents: true,
                        Scenarios: scenarios,
                        UserInterfaceOnly: false,
                        AllowFormattingCells: allowFormattingCells,
                        AllowFormattingColumns: allowFormattingColumns,
                        AllowFormattingRows: allowFormattingRows,
                        AllowInsertingColumns: allowInsertingColumns,
                        AllowInsertingRows: allowInsertingRows,
                        AllowInsertingHyperlinks: allowInsertingHyperlinks,
                        AllowDeletingColumns: allowDeletingColumns,
                        AllowDeletingRows: allowDeletingRows,
                        AllowSorting: allowSorting,
                        AllowFiltering: allowFiltering,
                        AllowUsingPivotTables: allowUsingPivotTables);
                });
            }

            JObject data = BuildWorkbookProtectionData(application, workbook, worksheet.Name);
            data["action"] = protect ? "protect" : "unprotect";
            data["changed"] = wasProtected != protect;
            ExcelUndoScope.RecordUnavailable(
                data,
                "工作表保护状态与细粒度选项不在区域快照中，未登记自动撤销。");
            Observation observation = Observation.Create(
                "protect_workbook",
                (protect ? "已保护工作表「" : "已取消工作表「") + worksheet.Name + "」的保护。",
                wasProtected != protect);
            observation.AddTargetRef(ExcelComSupport.FullAddress(worksheet.Name, "A1"));
            return OperationResult.Succeed(
                protect ? "工作表保护已启用（含细粒度选项）。" : "工作表保护已取消。",
                data,
                observation);
        }

        private static JObject BuildWorkbookProtectionData(
            Excel.Application application,
            Excel.Workbook workbook,
            string sheetName)
        {
            JObject data = new JObject();
            data["workbook"] = workbook.Name;
            data["fullName"] = workbook.FullName;
            JObject workbookProtection = new JObject();
            workbookProtection["structureProtected"] = workbook.ProtectStructure;
            workbookProtection["windowsProtected"] = workbook.ProtectWindows;
            workbookProtection["hasPassword"] = workbook.HasPassword;
            data["workbookProtection"] = workbookProtection;

            JArray sheets = new JArray();
            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                Excel.Worksheet worksheet = ExcelComSupport.FindWorksheet(workbook, sheetName);
                if (worksheet != null)
                {
                    try
                    {
                        sheets.Add(BuildSheetProtectionData(worksheet));
                    }
                    finally
                    {
                        ComHelper.Release(worksheet);
                    }
                }
            }
            else
            {
                Excel.Sheets collection = null;
                try
                {
                    collection = workbook.Worksheets;
                    int count = collection.Count;
                    for (int index = 1; index <= count; index++)
                    {
                        object item = collection[index];
                        Excel.Worksheet worksheet = item as Excel.Worksheet;
                        if (worksheet == null)
                        {
                            ComHelper.Release(item);
                            continue;
                        }
                        try
                        {
                            sheets.Add(BuildSheetProtectionData(worksheet));
                        }
                        finally
                        {
                            ComHelper.Release(worksheet);
                        }
                    }
                }
                finally
                {
                    ComHelper.Release(collection);
                }
            }
            data["sheets"] = sheets;
            data["sheetCount"] = sheets.Count;
            return data;
        }

        private static JObject BuildSheetProtectionData(Excel.Worksheet worksheet)
        {
            JObject sheet = new JObject();
            sheet["name"] = worksheet.Name;
            sheet["index"] = worksheet.Index;
            sheet["visible"] = ExcelComSupport.VisibilityText(worksheet.Visible);
            sheet["protected"] = worksheet.ProtectContents;
            sheet["drawingObjectsProtected"] = worksheet.ProtectDrawingObjects;
            sheet["scenariosProtected"] = worksheet.ProtectScenarios;

            Excel.Protection protection = null;
            try
            {
                protection = worksheet.Protection;
                sheet["protectionMode"] = worksheet.ProtectionMode;
                JObject options = new JObject();
                foreach (string optionName in ProtectionOptionNames)
                {
                    options[optionName] = ReadProtectionFlag(protection, optionName);
                }
                sheet["allow"] = options;
            }
            catch (Exception ex)
            {
                sheet["protectionError"] = ex.Message;
            }
            finally
            {
                ComHelper.Release(protection);
            }
            return sheet;
        }

        private static bool ReadProtectionFlag(object protection, string propertyName)
        {
            object raw = ReadComProperty(protection, propertyName);
            if (raw == null)
            {
                return false;
            }
            try
            {
                return Convert.ToBoolean(raw, CultureInfo.InvariantCulture);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void ApplyInitialSheets(
            Excel.Workbook workbook,
            List<string> sheetNames,
            bool preserveExisting)
        {
            if (workbook == null || sheetNames == null || sheetNames.Count == 0)
            {
                return;
            }
            if (preserveExisting)
            {
                foreach (string name in sheetNames)
                {
                    Excel.Worksheet existing = ExcelComSupport.FindWorksheet(workbook, name);
                    if (existing != null)
                    {
                        ComHelper.Release(existing);
                        continue;
                    }
                    Excel.Worksheet added = null;
                    try
                    {
                        added = (Excel.Worksheet)workbook.Worksheets.Add(
                            After: workbook.Worksheets[workbook.Worksheets.Count]);
                        added.Name = name;
                    }
                    finally
                    {
                        ComHelper.Release(added);
                    }
                }
                return;
            }

            Excel.Worksheet first = null;
            try
            {
                first = (Excel.Worksheet)workbook.Worksheets[1];
                first.Name = sheetNames[0];
            }
            finally
            {
                ComHelper.Release(first);
            }

            while (workbook.Worksheets.Count > 1)
            {
                object item = workbook.Worksheets[workbook.Worksheets.Count];
                Excel.Worksheet worksheet = item as Excel.Worksheet;
                try
                {
                    if (worksheet != null)
                    {
                        worksheet.Delete();
                    }
                }
                finally
                {
                    ComHelper.Release(worksheet, item);
                }
            }

            for (int index = 1; index < sheetNames.Count; index++)
            {
                Excel.Worksheet added = null;
                try
                {
                    added = (Excel.Worksheet)workbook.Worksheets.Add(
                        After: workbook.Worksheets[workbook.Worksheets.Count]);
                    added.Name = sheetNames[index];
                }
                finally
                {
                    ComHelper.Release(added);
                }
            }
        }

        private static List<string> GetWorksheetNames(
            Excel.Workbook workbook,
            List<string> fallback,
            bool workbookAvailable)
        {
            List<string> names = new List<string>();
            if (workbook == null || !workbookAvailable)
            {
                if (fallback != null)
                {
                    names.AddRange(fallback);
                }
                return names;
            }
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1; index <= sheets.Count; index++)
                {
                    object item = sheets[index];
                    Excel.Worksheet worksheet = item as Excel.Worksheet;
                    if (worksheet == null)
                    {
                        ComHelper.Release(item);
                        continue;
                    }
                    try
                    {
                        names.Add(worksheet.Name);
                    }
                    finally
                    {
                        ComHelper.Release(worksheet);
                    }
                }
            }
            finally
            {
                ComHelper.Release(sheets);
            }
            return names;
        }

        private static JObject BuildCreatedWorkbookData(
            Excel.Workbook workbook,
            string path,
            bool created,
            bool opened,
            string templatePath)
        {
            JObject data = new JObject();
            data["path"] = path;
            data["fullName"] = workbook == null ? path : workbook.FullName;
            data["workbookName"] = workbook == null ? Path.GetFileName(path) : workbook.Name;
            data["created"] = created;
            data["opened"] = opened;
            data["template"] = templatePath;
            data["sheets"] = new JArray(GetWorksheetNames(workbook, null, workbook != null));
            data["bytes"] = File.Exists(path) ? new FileInfo(path).Length : 0;
            return data;
        }

        private static Excel.Workbook FindOpenWorkbook(Excel.Application application, string fullPath)
        {
            Excel.Workbooks workbooks = null;
            try
            {
                workbooks = application.Workbooks;
                int count = workbooks.Count;
                for (int index = 1; index <= count; index++)
                {
                    Excel.Workbook workbook = null;
                    try
                    {
                        workbook = workbooks.Item[index];
                        if (workbook != null &&
                            string.Equals(workbook.FullName, fullPath, StringComparison.OrdinalIgnoreCase))
                        {
                            Excel.Workbook found = workbook;
                            workbook = null;
                            return found;
                        }
                    }
                    finally
                    {
                        ComHelper.Release(workbook);
                    }
                }
                return null;
            }
            finally
            {
                ComHelper.Release(workbooks);
            }
        }

        internal static int WorkbookFileFormat(string fullPath)
        {
            string extension = Path.GetExtension(fullPath).ToLowerInvariant();
            if (extension == ".xlsx")
            {
                return 51;
            }
            if (extension == ".xlsm")
            {
                return 52;
            }
            if (extension == ".xls")
            {
                return 56;
            }
            if (extension == ".csv")
            {
                return 6;
            }
            return 0;
        }

        private static string DescribeWorkbookFormat(int fileFormat, string path)
        {
            if (fileFormat == 51)
            {
                return "xlsx";
            }
            if (fileFormat == 52)
            {
                return "xlsm";
            }
            if (fileFormat == 56)
            {
                return "xls";
            }
            if (fileFormat == 6)
            {
                return "csv";
            }
            return Path.GetExtension(path).TrimStart('.');
        }
    }
}
