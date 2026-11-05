using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json.Linq;
using Office.Contracts;
using Office.Undo;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Office.Parallel
{
    /// <summary>一次池内作业的完整结果（工具结果 + 打开/保存/撤销收尾事实）。</summary>
    public sealed class ParallelJobOutcome
    {
        public OperationResult Result { get; set; }

        public bool OpenedReadOnly { get; set; }

        public bool Saved { get; set; }

        public string SaveSkippedReason { get; set; }

        public int UndoNodesRemoved { get; set; }

        public bool SaveVerified { get; set; }

        public string SaveVerificationNote { get; set; }

        public bool SaveChangedOnDisk { get; set; }

        public string SaveBeforeHash { get; set; }

        public string SaveAfterHash { get; set; }

        public long SaveBeforeLength { get; set; }

        public long SaveAfterLength { get; set; }

        public string SaveBeforeLastWriteUtc { get; set; }

        public string SaveAfterLastWriteUtc { get; set; }

        public ParallelJobOutcome()
        {
            SaveSkippedReason = string.Empty;
            SaveVerificationNote = string.Empty;
            SaveBeforeHash = string.Empty;
            SaveAfterHash = string.Empty;
            SaveBeforeLastWriteUtc = string.Empty;
            SaveAfterLastWriteUtc = string.Empty;
        }
    }

    /// <summary>
    /// 池内一次性 STA 线程的执行体：创建专属 Application → 打开文件 → 覆盖 OfficeSession
    /// 的应用实例 → 调用已注册工具 → 按 savePolicy 收尾 → Close/Quit/ReleaseComObject。
    ///
    /// 本类只允许在手动创建的一次性 STA 线程上运行；任何异常都在这里收敛成 OperationResult，
    /// 不允许把 COM 对象带回调度器线程。
    /// </summary>
    public static class ParallelStaJobRunner
    {
        public static ParallelJobOutcome Run(OfficeWorkItem item)
        {
            ComHelper.EnsureSta("ParallelStaJobRunner.Run");
            if (item == null)
            {
                return FailureOutcome(
                    OfficeErrorCode.INVALID_PARAMS,
                    "并行作业为空，未执行任何操作。");
            }

            string appType = ParallelToolCatalog.NormalizeApp(item.AppType);
            if (string.Equals(appType, ParallelToolCatalog.Excel, StringComparison.Ordinal))
            {
                return RunExcel(item);
            }
            if (string.Equals(appType, ParallelToolCatalog.Word, StringComparison.Ordinal))
            {
                return RunWord(item);
            }
            if (string.Equals(appType, ParallelToolCatalog.Ppt, StringComparison.Ordinal))
            {
                return RunPowerPoint(item);
            }
            return FailureOutcome(
                OfficeErrorCode.INVALID_PARAMS,
                "不支持的 appType「" + (item.AppType ?? string.Empty) + "」，只允许 excel、word 或 ppt。");
        }

        // ------------------------------------------------------------------
        // Excel
        // ------------------------------------------------------------------

        private static ParallelJobOutcome RunExcel(OfficeWorkItem item)
        {
            Excel.Application application = null;
            try
            {
                application = new Excel.Application();
                application.Visible = Office.Config.AppConfig.Current.ExcelVisible;
                application.DisplayAlerts = false;
                application.ScreenUpdating = false;
                return RunExcelOnApplication(item, application);
            }
            catch (Exception ex)
            {
                ParallelJobOutcome outcome = new ParallelJobOutcome();
                outcome.Result = OfficeErrorMapper.ToOperationResult(
                    ex,
                    OfficeApp.Excel,
                    "创建池内 Excel 实例失败；作业未执行。");
                return outcome;
            }
            finally
            {
                CloseExcel(application, null);
            }
        }

        /// <summary>
        /// T10：在常驻 STA worker 的 Application 上执行一个文件作业。
        /// 只关闭本次打开的文档，不 Quit Application；由 worker 生命周期统一收口。
        /// </summary>
        public static ParallelJobOutcome RunWithApplication(OfficeWorkItem item, Excel.Application application)
        {
            ComHelper.EnsureSta("ParallelStaJobRunner.RunWithApplication(Excel)");
            return RunExcelOnApplication(item, application);
        }

        private static ParallelJobOutcome RunExcelOnApplication(OfficeWorkItem item, Excel.Application application)
        {
            ParallelJobOutcome outcome = new ParallelJobOutcome();
            Excel.Workbook workbook = null;
            bool readOnlyRequested = ParallelToolCatalog.IsReadOnlyTool(ParallelToolCatalog.Excel, item.ToolName);
            DateTime undoStartUtc = DateTime.UtcNow;

            try
            {
                workbook = application.Workbooks.Open(
                    Filename: item.FilePath,
                    UpdateLinks: 0,
                    ReadOnly: readOnlyRequested,
                    IgnoreReadOnlyRecommended: true,
                    AddToMru: false,
                    Notify: false);

                if (workbook == null)
                {
                    outcome.Result = OperationResult.Failed(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "Excel 没有返回已打开的工作簿对象，作业未执行：" + item.FilePath);
                    return outcome;
                }

                outcome.OpenedReadOnly = readOnlyRequested || SafeWorkbookReadOnly(workbook);
                if (!readOnlyRequested && outcome.OpenedReadOnly)
                {
                    outcome.Result = OperationResult.Failed(
                        OfficeErrorCode.READ_ONLY,
                        "文件只能以只读方式打开（可能已被其他程序占用或文件带只读标记），写工具未执行：" + item.FilePath);
                    return outcome;
                }

                TryActivateWorkbook(workbook);

                OperationResult toolResult;
                using (OfficeSession.UseExcelApplication(application))
                {
                    toolResult = InvokeTool(item, OfficeApp.Excel);
                }

                outcome.Result = ApplyExcelSavePolicy(item, workbook, toolResult, outcome);
                PrunePoolUndoNodes(item, undoStartUtc, outcome);
                MarkUndoUnavailableForPoolWrite(item, outcome);
                return outcome;
            }
            catch (Exception ex)
            {
                outcome.Result = OfficeErrorMapper.ToOperationResult(
                    ex,
                    OfficeApp.Excel,
                    "池内一次性 Excel 实例未完成作业；不会自动重放。");
                PrunePoolUndoNodes(item, undoStartUtc, outcome);
                return outcome;
            }
            finally
            {
                // 只关闭本次文档；Application 由 RunExcel 或常驻 worker 负责 Quit。
                CloseExcel(null, workbook);
            }
        }

        private static OperationResult ApplyExcelSavePolicy(
            OfficeWorkItem item,
            Excel.Workbook workbook,
            OperationResult toolResult,
            ParallelJobOutcome outcome)
        {
            if (toolResult == null)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "Excel 工具「" + item.ToolName + "」未返回结果。");
            }

            if (!string.Equals(item.SavePolicy, ParallelToolCatalog.SavePolicySave, StringComparison.Ordinal))
            {
                outcome.Saved = false;
                outcome.SaveSkippedReason = "save_policy_none";
                return toolResult;
            }

            if (!toolResult.Success)
            {
                outcome.Saved = false;
                outcome.SaveSkippedReason = "tool_failed";
                return toolResult;
            }

            if (outcome.OpenedReadOnly)
            {
                outcome.Saved = false;
                outcome.SaveSkippedReason = "opened_read_only";
                return toolResult;
            }

            try
            {
                FileIntegritySnapshot before = FileIntegrityVerifier.Capture(
                    item.FilePath,
                    Office.Config.AppConfig.Current.WorkerSaveVerifyMaxHashMb);
                workbook.Save();
                outcome.Saved = true;
                outcome.SaveSkippedReason = string.Empty;
                if (!ApplySaveVerification(item.FilePath, before, outcome, out string verificationNote))
                {
                    outcome.SaveSkippedReason = "save_verification_failed";
                    OperationResult verificationFailure = OperationResult.Failed(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "工具已在内存中执行成功，但保存后文件校验失败：" + verificationNote,
                        "不会自动重放；请人工确认文件磁盘状态。");
                    AttachToolDataToSaveFailure(verificationFailure, toolResult);
                    return verificationFailure;
                }
                return toolResult;
            }
            catch (Exception ex)
            {
                outcome.Saved = false;
                outcome.SaveSkippedReason = "save_failed";
                OperationResult failure = OfficeErrorMapper.ToOperationResult(
                    ex,
                    OfficeApp.Excel,
                    "工具已在内存中执行成功，但保存失败；关闭池内实例会丢弃未保存改动。不会自动重放。");
                AttachToolDataToSaveFailure(failure, toolResult);
                return failure;
            }
        }

        private static void TryActivateWorkbook(Excel.Workbook workbook)
        {
            try
            {
                workbook.Activate();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[parallel] 激活池内 Excel 工作簿失败：" + ex.Message);
            }
        }

        private static bool SafeWorkbookReadOnly(Excel.Workbook workbook)
        {
            try
            {
                return workbook.ReadOnly;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[parallel] 读取 Excel 工作簿 ReadOnly 状态失败：" + ex.Message);
                return false;
            }
        }

        private static void CloseExcel(Excel.Application application, Excel.Workbook workbook)
        {
            if (workbook != null)
            {
                try
                {
                    workbook.Close(false);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[parallel] 关闭池内 Excel 工作簿失败：" + ex.Message);
                }
                ComHelper.ReleaseOne(workbook);
            }

            if (application != null)
            {
                try
                {
                    application.DisplayAlerts = false;
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[parallel] 关闭池内 Excel DisplayAlerts 失败：" + ex.Message);
                }
                try
                {
                    application.Quit();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[parallel] 退出池内 Excel 实例失败：" + ex.Message);
                }
                ComHelper.ReleaseOne(application);
            }
        }

        // ------------------------------------------------------------------
        // Word
        // ------------------------------------------------------------------

        private static ParallelJobOutcome RunWord(OfficeWorkItem item)
        {
            Word.Application application = null;
            try
            {
                application = new Word.Application();
                application.Visible = Office.Config.AppConfig.Current.WordVisible;
                application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                application.ScreenUpdating = false;
                return RunWordOnApplication(item, application);
            }
            catch (Exception ex)
            {
                ParallelJobOutcome outcome = new ParallelJobOutcome();
                outcome.Result = OfficeErrorMapper.ToOperationResult(
                    ex,
                    OfficeApp.Word,
                    "创建池内 Word 实例失败；作业未执行。");
                return outcome;
            }
            finally
            {
                CloseWord(application, null);
            }
        }

        /// <summary>
        /// T10：在常驻 STA worker 的 Word Application 上执行一个文件作业。
        /// 只关闭本次打开的文档，不 Quit Application；由 worker 生命周期统一收口。
        /// </summary>
        public static ParallelJobOutcome RunWithApplication(OfficeWorkItem item, Word.Application application)
        {
            ComHelper.EnsureSta("ParallelStaJobRunner.RunWithApplication(Word)");
            return RunWordOnApplication(item, application);
        }

        private static ParallelJobOutcome RunWordOnApplication(OfficeWorkItem item, Word.Application application)
        {
            ParallelJobOutcome outcome = new ParallelJobOutcome();
            Word.Document document = null;
            bool readOnlyRequested = ParallelToolCatalog.IsReadOnlyTool(ParallelToolCatalog.Word, item.ToolName);
            DateTime undoStartUtc = DateTime.UtcNow;

            try
            {
                document = OpenWordDocument(application, item.FilePath, readOnlyRequested);
                if (document == null)
                {
                    outcome.Result = OperationResult.Failed(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "Word 没有返回已打开的文档对象，作业未执行：" + item.FilePath);
                    return outcome;
                }

                outcome.OpenedReadOnly = readOnlyRequested || SafeWordReadOnly(document);
                if (!readOnlyRequested && outcome.OpenedReadOnly)
                {
                    outcome.Result = OperationResult.Failed(
                        OfficeErrorCode.READ_ONLY,
                        "文件只能以只读方式打开（可能已被其他程序占用或文件带只读标记），写工具未执行：" + item.FilePath);
                    return outcome;
                }

                TryActivateDocument(document);

                OperationResult toolResult;
                using (OfficeSession.UseWordApplication(application))
                {
                    toolResult = InvokeTool(item, OfficeApp.Word);
                }

                outcome.Result = ApplyWordSavePolicy(item, document, toolResult, outcome);
                PrunePoolUndoNodes(item, undoStartUtc, outcome);
                MarkUndoUnavailableForPoolWrite(item, outcome);
                return outcome;
            }
            catch (Exception ex)
            {
                outcome.Result = OfficeErrorMapper.ToOperationResult(
                    ex,
                    OfficeApp.Word,
                    "池内一次性 Word 实例未完成作业；不会自动重放。");
                PrunePoolUndoNodes(item, undoStartUtc, outcome);
                return outcome;
            }
            finally
            {
                // 只关闭本次文档；Application 由 RunWord 或常驻 worker 负责 Quit。
                CloseWord(null, document);
            }
        }

        private static Word.Document OpenWordDocument(Word.Application application, string filePath, bool readOnly)
        {
            Word.Documents documents = null;
            try
            {
                documents = application.Documents;
                object fileName = filePath;
                object confirmConversions = false;
                object readOnlyObject = readOnly;
                object addToRecentFiles = false;
                // Application 保持 Visible=false（用户不可见）；文档窗口同样不显示。
                object visibleObject = false;
                object missing = Type.Missing;
                return documents.Open(
                    ref fileName, ref confirmConversions, ref readOnlyObject, ref addToRecentFiles,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref visibleObject, ref missing, ref missing, ref missing, ref missing);
            }
            finally
            {
                ComHelper.ReleaseOne(documents);
            }
        }

        private static OperationResult ApplyWordSavePolicy(
            OfficeWorkItem item,
            Word.Document document,
            OperationResult toolResult,
            ParallelJobOutcome outcome)
        {
            if (toolResult == null)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "Word 工具「" + item.ToolName + "」未返回结果。");
            }

            if (!string.Equals(item.SavePolicy, ParallelToolCatalog.SavePolicySave, StringComparison.Ordinal))
            {
                outcome.Saved = false;
                outcome.SaveSkippedReason = "save_policy_none";
                return toolResult;
            }

            if (!toolResult.Success)
            {
                outcome.Saved = false;
                outcome.SaveSkippedReason = "tool_failed";
                return toolResult;
            }

            if (outcome.OpenedReadOnly)
            {
                outcome.Saved = false;
                outcome.SaveSkippedReason = "opened_read_only";
                return toolResult;
            }

            try
            {
                FileIntegritySnapshot before = FileIntegrityVerifier.Capture(
                    item.FilePath,
                    Office.Config.AppConfig.Current.WorkerSaveVerifyMaxHashMb);
                document.Save();
                outcome.Saved = true;
                outcome.SaveSkippedReason = string.Empty;
                if (!ApplySaveVerification(item.FilePath, before, outcome, out string verificationNote))
                {
                    outcome.SaveSkippedReason = "save_verification_failed";
                    OperationResult verificationFailure = OperationResult.Failed(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "工具已在内存中执行成功，但保存后文件校验失败：" + verificationNote,
                        "不会自动重放；请人工确认文件磁盘状态。");
                    AttachToolDataToSaveFailure(verificationFailure, toolResult);
                    return verificationFailure;
                }
                return toolResult;
            }
            catch (Exception ex)
            {
                outcome.Saved = false;
                outcome.SaveSkippedReason = "save_failed";
                OperationResult failure = OfficeErrorMapper.ToOperationResult(
                    ex,
                    OfficeApp.Word,
                    "工具已在内存中执行成功，但保存失败；关闭池内实例会丢弃未保存改动。不会自动重放。");
                AttachToolDataToSaveFailure(failure, toolResult);
                return failure;
            }
        }

        private static void TryActivateDocument(Word.Document document)
        {
            try
            {
                document.Activate();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[parallel] 激活池内 Word 文档失败：" + ex.Message);
            }
        }

        private static bool SafeWordReadOnly(Word.Document document)
        {
            try
            {
                return document.ReadOnly;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[parallel] 读取 Word 文档 ReadOnly 状态失败：" + ex.Message);
                return false;
            }
        }

        private static void CloseWord(Word.Application application, Word.Document document)
        {
            object missing = Type.Missing;
            if (document != null)
            {
                object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                try
                {
                    document.Close(ref saveChanges, ref missing, ref missing);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[parallel] 关闭池内 Word 文档失败：" + ex.Message);
                }
                ComHelper.ReleaseOne(document);
            }

            if (application != null)
            {
                try
                {
                    application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[parallel] 关闭池内 Word DisplayAlerts 失败：" + ex.Message);
                }
                try
                {
                    object quitSaveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                    application.Quit(ref quitSaveChanges, ref missing, ref missing);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[parallel] 退出池内 Word 实例失败：" + ex.Message);
                }
                ComHelper.ReleaseOne(application);
            }
        }

        // ------------------------------------------------------------------
        // PowerPoint
        // ------------------------------------------------------------------

        private static ParallelJobOutcome RunPowerPoint(OfficeWorkItem item)
        {
            PowerPointApplicationLease lease = null;
            try
            {
                lease = PowerPointApplicationLease.Acquire(
                    Office.Config.AppConfig.Current.PowerPointVisible);
                return RunPowerPointOnApplication(item, lease.Application);
            }
            catch (Exception ex)
            {
                ParallelJobOutcome outcome = new ParallelJobOutcome();
                outcome.Result = OfficeErrorMapper.ToOperationResult(
                    ex,
                    OfficeApp.PowerPoint,
                    "创建池内 PowerPoint 实例失败；作业未执行。");
                return outcome;
            }
            finally
            {
                if (lease != null)
                {
                    lease.Dispose();
                }
            }
        }

        /// <summary>
        /// T11：在常驻 STA worker 的 PowerPoint Application 上执行一个文件作业。
        /// 只关闭本次打开的演示文稿，不 Quit Application；由 worker 生命周期统一收口。
        /// </summary>
        public static ParallelJobOutcome RunWithApplication(OfficeWorkItem item, PowerPoint.Application application)
        {
            ComHelper.EnsureSta("ParallelStaJobRunner.RunWithApplication(PowerPoint)");
            return RunPowerPointOnApplication(item, application);
        }

        private static ParallelJobOutcome RunPowerPointOnApplication(
            OfficeWorkItem item,
            PowerPoint.Application application)
        {
            ParallelJobOutcome outcome = new ParallelJobOutcome();
            PowerPoint.Presentation presentation = null;
            bool readOnlyRequested = ParallelToolCatalog.IsReadOnlyTool(ParallelToolCatalog.Ppt, item.ToolName);
            DateTime undoStartUtc = DateTime.UtcNow;
            try
            {
                dynamic presentations = application.Presentations;
                object opened = presentations.Open(item.FilePath, readOnlyRequested ? -1 : 0, 0, 0);
                presentation = opened as PowerPoint.Presentation;
                if (presentation == null)
                {
                    outcome.Result = OperationResult.Failed(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "PowerPoint 没有返回已打开的演示文稿对象，作业未执行：" + item.FilePath);
                    return outcome;
                }

                outcome.OpenedReadOnly = readOnlyRequested || SafePresentationReadOnly(presentation);
                if (!readOnlyRequested && outcome.OpenedReadOnly)
                {
                    outcome.Result = OperationResult.Failed(
                        OfficeErrorCode.READ_ONLY,
                        "文件只能以只读方式打开（可能已被其他程序占用或文件带只读标记），写工具未执行：" + item.FilePath);
                    return outcome;
                }

                OperationResult toolResult;
                using (OfficeSession.UsePowerPointApplication(application))
                {
                    toolResult = InvokeTool(item, OfficeApp.PowerPoint);
                }

                outcome.Result = ApplyPowerPointSavePolicy(item, presentation, toolResult, outcome);
                PrunePoolUndoNodes(item, undoStartUtc, outcome);
                MarkUndoUnavailableForPoolWrite(item, outcome);
                return outcome;
            }
            catch (Exception ex)
            {
                outcome.Result = OfficeErrorMapper.ToOperationResult(
                    ex,
                    OfficeApp.PowerPoint,
                    "池内 PowerPoint 实例未完成作业；不会自动重放。");
                PrunePoolUndoNodes(item, undoStartUtc, outcome);
                return outcome;
            }
            finally
            {
                ClosePowerPoint(null, presentation);
            }
        }

        private static OperationResult ApplyPowerPointSavePolicy(
            OfficeWorkItem item,
            PowerPoint.Presentation presentation,
            OperationResult toolResult,
            ParallelJobOutcome outcome)
        {
            if (toolResult == null)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "PowerPoint 工具「" + item.ToolName + "」未返回结果。");
            }

            if (!string.Equals(item.SavePolicy, ParallelToolCatalog.SavePolicySave, StringComparison.Ordinal))
            {
                outcome.Saved = false;
                outcome.SaveSkippedReason = "save_policy_none";
                return toolResult;
            }
            if (!toolResult.Success)
            {
                outcome.Saved = false;
                outcome.SaveSkippedReason = "tool_failed";
                return toolResult;
            }
            if (outcome.OpenedReadOnly)
            {
                outcome.Saved = false;
                outcome.SaveSkippedReason = "opened_read_only";
                return toolResult;
            }

            try
            {
                FileIntegritySnapshot before = FileIntegrityVerifier.Capture(
                    item.FilePath,
                    Office.Config.AppConfig.Current.WorkerSaveVerifyMaxHashMb);
                presentation.Save();
                outcome.Saved = true;
                outcome.SaveSkippedReason = string.Empty;
                if (!ApplySaveVerification(item.FilePath, before, outcome, out string verificationNote))
                {
                    outcome.SaveSkippedReason = "save_verification_failed";
                    OperationResult verificationFailure = OperationResult.Failed(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "工具已在内存中执行成功，但保存后文件校验失败：" + verificationNote,
                        "不会自动重放；请人工确认文件磁盘状态。");
                    AttachToolDataToSaveFailure(verificationFailure, toolResult);
                    return verificationFailure;
                }
                return toolResult;
            }
            catch (Exception ex)
            {
                outcome.Saved = false;
                outcome.SaveSkippedReason = "save_failed";
                OperationResult failure = OfficeErrorMapper.ToOperationResult(
                    ex,
                    OfficeApp.PowerPoint,
                    "工具已在内存中执行成功，但保存失败；关闭池内实例会丢弃未保存改动。不会自动重放。");
                AttachToolDataToSaveFailure(failure, toolResult);
                return failure;
            }
        }

        private static bool SafePresentationReadOnly(PowerPoint.Presentation presentation)
        {
            try
            {
                dynamic dynamicPresentation = presentation;
                return Convert.ToInt32((object)dynamicPresentation.ReadOnly) != 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[parallel] 读取 PowerPoint 演示文稿 ReadOnly 状态失败：" + ex.Message);
                return false;
            }
        }

        private static void ClosePowerPoint(
            PowerPoint.Application application,
            PowerPoint.Presentation presentation)
        {
            if (presentation != null)
            {
                try
                {
                    presentation.Close();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[parallel] 关闭池内 PowerPoint 演示文稿失败：" + ex.Message);
                }
                ComHelper.ReleaseOne(presentation);
            }

            if (application != null)
            {
                try
                {
                    application.Quit();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[parallel] 退出池内 PowerPoint 实例失败：" + ex.Message);
                }
                ComHelper.ReleaseOne(application);
            }
        }

        // ------------------------------------------------------------------
        // 公共部分
        // ------------------------------------------------------------------

        private static OperationResult InvokeTool(OfficeWorkItem item, OfficeApp application)
        {
            Func<string, JObject, ToolResult> invoker = ParallelDispatcher.ToolInvoker;
            if (invoker == null)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "并行调度器未装配工具执行委托（Program 尚未注入 ParallelDispatcher.ToolInvoker），作业未执行。");
            }

            JObject arguments = item.Args == null ? new JObject() : (JObject)item.Args.DeepClone();
            ToolResult toolResult;
            try
            {
                toolResult = invoker(item.ToolName, arguments);
            }
            catch (Exception ex)
            {
                return OfficeErrorMapper.ToOperationResult(
                    ex,
                    application,
                    "工具在池内实例上抛出未捕获异常；关闭实例可能丢弃未保存改动。");
            }
            return FromToolResult(toolResult, application);
        }

        private static OperationResult FromToolResult(ToolResult toolResult, OfficeApp application)
        {
            if (toolResult == null)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "工具没有返回结果（执行器内部缺陷）。");
            }

            if (toolResult.Success)
            {
                return OperationResult.Succeed(toolResult.Message, toolResult.Data);
            }

            string errorCode = string.IsNullOrWhiteSpace(toolResult.ErrorCode)
                ? OfficeErrorMapper.ToCode(OfficeErrorCode.INTERNAL_ERROR)
                : toolResult.ErrorCode;
            OperationResult failure = OperationResult.Failed(
                errorCode,
                toolResult.Message,
                toolResult.Recoverable,
                null);
            failure.Data = toolResult.Data == null ? null : (JObject)toolResult.Data.DeepClone();
            failure.Diagnostic = toolResult.Diagnostic ?? string.Empty;
            return failure;
        }

        private static bool ApplySaveVerification(
            string filePath,
            FileIntegritySnapshot before,
            ParallelJobOutcome outcome,
            out string note)
        {
            FileIntegritySnapshot after = FileIntegrityVerifier.Capture(
                filePath,
                Office.Config.AppConfig.Current.WorkerSaveVerifyMaxHashMb);
            bool changedOnDisk;
            bool verified = FileIntegrityVerifier.Verify(before, after, out changedOnDisk, out note);
            outcome.SaveVerified = verified;
            outcome.SaveChangedOnDisk = changedOnDisk;
            outcome.SaveVerificationNote = note ?? string.Empty;
            outcome.SaveBeforeHash = before == null ? string.Empty : (before.Sha256 ?? string.Empty);
            outcome.SaveAfterHash = after == null ? string.Empty : (after.Sha256 ?? string.Empty);
            outcome.SaveBeforeLength = before == null ? 0 : before.Length;
            outcome.SaveAfterLength = after == null ? 0 : after.Length;
            outcome.SaveBeforeLastWriteUtc = before == null || before.LastWriteUtc == DateTime.MinValue
                ? string.Empty
                : before.LastWriteUtc.ToString("o");
            outcome.SaveAfterLastWriteUtc = after == null || after.LastWriteUtc == DateTime.MinValue
                ? string.Empty
                : after.LastWriteUtc.ToString("o");
            return verified;
        }

        private static void AttachToolDataToSaveFailure(OperationResult failure, OperationResult toolResult)
        {
            if (failure == null || toolResult == null)
            {
                return;
            }

            JObject data = failure.Data == null ? new JObject() : (JObject)failure.Data.DeepClone();
            data["toolSucceeded"] = true;
            if (toolResult.Data != null && toolResult.Data.HasValues)
            {
                data["toolData"] = toolResult.Data.DeepClone();
            }
            failure.Data = data;
        }

        /// <summary>
        /// 池内写工具会向会话撤销栈登记快照节点，但池实例随后关闭，节点无法回放。
        /// 这里按"文档名 + 本作业开始时间"移除本次作业新增的节点，避免污染主会话撤销栈。
        /// （同一文件的池作业已被文件锁串行化；主会话在批次处理期间不会并发写文档。）
        /// </summary>
        private static void PrunePoolUndoNodes(OfficeWorkItem item, DateTime startedUtc, ParallelJobOutcome outcome)
        {
            if (outcome == null || item == null)
            {
                return;
            }

            string fileName = Path.GetFileName(item.FilePath);
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return;
            }

            List<SnapshotUndoNode> nodes = new List<SnapshotUndoNode>(OfficeSession.UndoStack.ToList());
            int removed = 0;
            foreach (SnapshotUndoNode node in nodes)
            {
                if (node == null)
                {
                    continue;
                }
                if (node.CreatedUtc < startedUtc)
                {
                    continue;
                }
                if (!string.Equals(node.DocumentName, fileName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                if (OfficeSession.UndoStack.Remove(node.Id))
                {
                    removed++;
                }
            }
            outcome.UndoNodesRemoved = removed;
        }

        private static void MarkUndoUnavailableForPoolWrite(OfficeWorkItem item, ParallelJobOutcome outcome)
        {
            if (outcome == null || outcome.Result == null || item == null)
            {
                return;
            }
            if (outcome.UndoNodesRemoved <= 0)
            {
                return;
            }

            if (outcome.Result.Observation != null && outcome.Result.Observation.UndoUnits.Count > 0)
            {
                outcome.Result.Observation.UndoUnits.Clear();
            }

            JObject data = outcome.Result.Data == null ? new JObject() : (JObject)outcome.Result.Data.DeepClone();
            data["undoRestorable"] = false;
            data["undoReason"] = "池内作业使用独立 STA worker，撤销节点按 worker/任务隔离；已从本次 worker 撤销栈移除。";
            data["undoNodesRemoved"] = outcome.UndoNodesRemoved;
            outcome.Result.Data = data;
        }

        private static ParallelJobOutcome FailureOutcome(OfficeErrorCode code, string message)
        {
            ParallelJobOutcome outcome = new ParallelJobOutcome();
            outcome.Result = OperationResult.Failed(code, message);
            return outcome;
        }
    }
}
