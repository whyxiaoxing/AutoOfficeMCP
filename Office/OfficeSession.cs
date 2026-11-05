using System;
using System.Runtime.InteropServices;
using System.Threading;
using Office.Config;
using Office.Undo;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;
using WordApplication = Microsoft.Office.Interop.Word.Application;
using PowerPointApplication = Microsoft.Office.Interop.PowerPoint.Application;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace Office
{
    /// <summary>执行器所属应用。</summary>
    public enum OfficeApp
    {
        Excel,
        Word,
        PowerPoint
    }

    /// <summary>
    /// 每应用单例会话（进程级）：
    /// 1) 优先附着运行中的 Excel/Word（Marshal.GetActiveObject）；
    /// 2) MK_E_UNAVAILABLE 时按 config.json 决定是否 new Application()；
    /// 3) 所有公共 COM 入口断言 STA；
    /// 4) 会话级 SemaphoreSlim(1,1) 排队执行，超时返回 SERVER_BUSY；
    /// 5) 持有跨工具调用的撤销节点栈。
    ///
    /// 用法：OfficeSession.GetExcel() / OfficeSession.RunExcel("工具名", app => {...})；
    /// 撤销栈：OfficeSession.UndoStack。
    /// </summary>
    public static class OfficeSession
    {
        private const string ExcelProgId = "Excel.Application";
        private const string WordProgId = "Word.Application";
        private const string PowerPointProgId = "PowerPoint.Application";
        private const int DefaultTimeoutSec = 120;

        private static readonly object SyncRoot = new object();
        private static readonly SemaphoreSlim SessionGate = new SemaphoreSlim(1, 1);
        private static readonly UndoStack SessionUndoStack = new UndoStack();

        private static ExcelApplication _excel;
        private static WordApplication _word;
        private static PowerPointApplication _powerPoint;
        private static string _excelAttachmentMode = "not-connected";
        private static string _wordAttachmentMode = "not-connected";
        private static string _powerPointAttachmentMode = "not-connected";

        // T09 池内一次性 STA 作业的应用实例覆盖：线程静态，仅对设置它的线程可见。
        // 覆盖存在时 GetExcel/GetWord 直接返回覆盖实例，RunExcel/RunWord 不再等待主会话锁；
        // Dispose 后恢复上一层覆盖（支持嵌套），主会话默认行为完全不变。
        [ThreadStatic]
        private static ExcelApplication _excelApplicationOverride;

        [ThreadStatic]
        private static WordApplication _wordApplicationOverride;

        [ThreadStatic]
        private static PowerPointApplication _powerPointApplicationOverride;

        // T10：每个常驻 STA worker 持有独立撤销栈，避免 A 用户撤销 B 用户的修改。
        [ThreadStatic]
        private static UndoStack _undoStackOverride;

        /// <summary>跨工具调用的撤销节点栈；只保存 Restorable=true 的节点。</summary>
        public static UndoStack UndoStack
        {
            get { return _undoStackOverride ?? SessionUndoStack; }
        }

        /// <summary>T10：在本线程覆盖 UndoStack，仅供常驻 STA worker 使用；Dispose 恢复上一层。</summary>
        public static IDisposable UseUndoStack(UndoStack stack)
        {
            ComHelper.EnsureSta("OfficeSession.UseUndoStack");
            if (stack == null)
            {
                throw new ArgumentNullException("stack");
            }
            UndoStack previous = _undoStackOverride;
            _undoStackOverride = stack;
            return new ApplicationOverrideScope(delegate { _undoStackOverride = previous; });
        }

        /// <summary>Excel 实例来源：attached / launched / not-connected。</summary>
        public static string ExcelAttachmentMode
        {
            get { return _excelAttachmentMode; }
        }

        /// <summary>Word 实例来源：attached / launched / not-connected。</summary>
        public static string WordAttachmentMode
        {
            get { return _wordAttachmentMode; }
        }

        /// <summary>PowerPoint 实例来源：attached / launched / not-connected。</summary>
        public static string PowerPointAttachmentMode
        {
            get { return _powerPointAttachmentMode; }
        }

        /// <summary>取得 Excel 强类型 Application；未连接时按配置附着或启动。</summary>
        public static ExcelApplication GetExcel()
        {
            ComHelper.EnsureSta("OfficeSession.GetExcel");
            ExcelApplication overridden = _excelApplicationOverride;
            if (overridden != null)
            {
                return overridden;
            }
            lock (SyncRoot)
            {
                if (_excel == null)
                {
                    _excel = AttachOrLaunchExcel();
                }
                return _excel;
            }
        }

        /// <summary>取得 Word 强类型 Application；未连接时按配置附着或启动。</summary>
        public static WordApplication GetWord()
        {
            ComHelper.EnsureSta("OfficeSession.GetWord");
            WordApplication overridden = _wordApplicationOverride;
            if (overridden != null)
            {
                return overridden;
            }
            lock (SyncRoot)
            {
                if (_word == null)
                {
                    _word = AttachOrLaunchWord();
                }
                return _word;
            }
        }

        /// <summary>取得 PowerPoint Application；未连接时按配置附着或启动。</summary>
        public static PowerPointApplication GetPowerPoint()
        {
            ComHelper.EnsureSta("OfficeSession.GetPowerPoint");
            PowerPointApplication overridden = _powerPointApplicationOverride;
            if (overridden != null)
            {
                return overridden;
            }
            lock (SyncRoot)
            {
                if (_powerPoint == null)
                {
                    _powerPoint = AttachOrLaunchPowerPoint();
                }
                return _powerPoint;
            }
        }

        /// <summary>
        /// T09 池内一次性 STA 作业专用：在本线程覆盖 GetExcel 返回的实例。
        /// 作用域内 RunExcel 不再等待主会话锁；Dispose 恢复上一层覆盖（支持嵌套）。
        /// 不影响主线程或其他线程的既有会话行为。
        /// </summary>
        public static IDisposable UseExcelApplication(ExcelApplication application)
        {
            ComHelper.EnsureSta("OfficeSession.UseExcelApplication");
            if (application == null)
            {
                throw new ArgumentNullException("application");
            }
            ExcelApplication previous = _excelApplicationOverride;
            _excelApplicationOverride = application;
            return new ApplicationOverrideScope(delegate { _excelApplicationOverride = previous; });
        }

        /// <summary>T09 池内一次性 STA 作业专用：覆盖 GetWord 返回的实例，语义同 UseExcelApplication。</summary>
        public static IDisposable UseWordApplication(WordApplication application)
        {
            ComHelper.EnsureSta("OfficeSession.UseWordApplication");
            if (application == null)
            {
                throw new ArgumentNullException("application");
            }
            WordApplication previous = _wordApplicationOverride;
            _wordApplicationOverride = application;
            return new ApplicationOverrideScope(delegate { _wordApplicationOverride = previous; });
        }

        /// <summary>池内/常驻 PPT worker 专用：覆盖 GetPowerPoint 返回的实例。</summary>
        public static IDisposable UsePowerPointApplication(PowerPointApplication application)
        {
            ComHelper.EnsureSta("OfficeSession.UsePowerPointApplication");
            if (application == null)
            {
                throw new ArgumentNullException("application");
            }
            PowerPointApplication previous = _powerPointApplicationOverride;
            _powerPointApplicationOverride = application;
            return new ApplicationOverrideScope(delegate { _powerPointApplicationOverride = previous; });
        }

        /// <summary>
        /// 进程退出前的会话收口：只 Quit 由本 server 自动启动（attachmentMode=launched）的实例，
        /// 已附着的用户 Office 实例只释放 COM 代理、绝不退出。重复调用安全。
        /// 客户端断开 stdin 后 server 正常退出时会执行，避免留下隐藏的 WINWORD/EXCEL 残留进程。
        /// </summary>
        public static void Shutdown()
        {
            ComHelper.EnsureSta("OfficeSession.Shutdown");
            lock (SyncRoot)
            {
                if (_excel != null)
                {
                    ExcelApplication application = _excel;
                    _excel = null;
                    bool launched = string.Equals(_excelAttachmentMode, "launched", StringComparison.Ordinal);
                    if (launched)
                    {
                        try
                        {
                            application.DisplayAlerts = false;
                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine("[session] 退出时关闭 Excel DisplayAlerts 失败：" + ex.Message);
                        }
                        try
                        {
                            application.Quit();
                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine("[session] 退出时 Quit 自动启动的 Excel 失败：" + ex.Message);
                        }
                    }
                    ComHelper.Release(application);
                    _excelAttachmentMode = "not-connected";
                }

                if (_word != null)
                {
                    WordApplication application = _word;
                    _word = null;
                    bool launched = string.Equals(_wordAttachmentMode, "launched", StringComparison.Ordinal);
                    if (launched)
                    {
                        try
                        {
                            application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine("[session] 退出时关闭 Word DisplayAlerts 失败：" + ex.Message);
                        }
                        try
                        {
                            object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                            object missing = Type.Missing;
                            application.Quit(ref saveChanges, ref missing, ref missing);
                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine("[session] 退出时 Quit 自动启动的 Word 失败：" + ex.Message);
                        }
                    }
                    ComHelper.Release(application);
                    _wordAttachmentMode = "not-connected";
                }

                if (_powerPoint != null)
                {
                    PowerPointApplication application = _powerPoint;
                    _powerPoint = null;
                    bool launched = string.Equals(
                        _powerPointAttachmentMode, "launched", StringComparison.Ordinal);
                    if (launched)
                    {
                        try
                        {
                            application.Quit();
                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine("[session] 退出时 Quit 自动启动的 PowerPoint 失败：" + ex.Message);
                        }
                    }
                    ComHelper.Release(application);
                    _powerPointAttachmentMode = "not-connected";
                }
            }
        }

        /// <summary>
        /// 在会话锁内执行一个 Excel 操作。排队而非拒绝；超过 session.timeoutSec 返回 SERVER_BUSY。
        /// 执行器不需要自行获取锁（T09 的显式 filePath 作业走自己的 STA 线程，不经过本方法）。
        /// </summary>
        public static OperationResult RunExcel(string operationName, Func<ExcelApplication, OperationResult> action)
        {
            ComHelper.EnsureSta("OfficeSession.RunExcel");
            if (action == null)
            {
                return OperationResult.Failed(OfficeErrorCode.INVALID_PARAMS, "Excel 执行委托不能为空。");
            }
            return ExecuteExclusive(OfficeApp.Excel, operationName, delegate
            {
                return action(GetExcel());
            });
        }

        /// <summary>在会话锁内执行一个 Word 操作；错误处理与超时语义同 RunExcel。</summary>
        public static OperationResult RunWord(string operationName, Func<WordApplication, OperationResult> action)
        {
            ComHelper.EnsureSta("OfficeSession.RunWord");
            if (action == null)
            {
                return OperationResult.Failed(OfficeErrorCode.INVALID_PARAMS, "Word 执行委托不能为空。");
            }
            return ExecuteExclusive(OfficeApp.Word, operationName, delegate
            {
                return action(GetWord());
            });
        }

        /// <summary>在会话锁内执行一个 PowerPoint 操作；错误处理与超时语义同 RunExcel。</summary>
        public static OperationResult RunPowerPoint(
            string operationName,
            Func<PowerPointApplication, OperationResult> action)
        {
            ComHelper.EnsureSta("OfficeSession.RunPowerPoint");
            if (action == null)
            {
                return OperationResult.Failed(OfficeErrorCode.INVALID_PARAMS, "PowerPoint 执行委托不能为空。");
            }
            return ExecuteExclusive(OfficeApp.PowerPoint, operationName, delegate
            {
                return action(GetPowerPoint());
            });
        }

        private static OperationResult ExecuteExclusive(OfficeApp application, string operationName, Func<OperationResult> action)
        {
            string displayName = string.IsNullOrWhiteSpace(operationName)
                ? GetApplicationKey(application)
                : operationName.Trim();

            // 池内覆盖作用域：作业已经拥有专属 Application 与文件锁，
            // 不再等待主会话锁，否则 Excel 池并发会被主会话串行化。
            if (HasApplicationOverride(application))
            {
                return InvokeGuarded(application, displayName, action);
            }

            int timeoutSec = AppConfig.Current.SessionTimeoutSec;
            if (timeoutSec <= 0)
            {
                timeoutSec = DefaultTimeoutSec;
            }

            bool acquired;
            try
            {
                acquired = SessionGate.Wait(TimeSpan.FromSeconds(timeoutSec));
            }
            catch (ObjectDisposedException)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.SERVER_BUSY,
                    "Office 会话锁已释放，操作「" + displayName + "」未执行。");
            }

            if (!acquired)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.SERVER_BUSY,
                    "操作「" + displayName + "」等待 Office 会话锁超过 " + timeoutSec.ToString() +
                    " 秒，未执行。请稍后重试或减少并发调用。");
            }

            try
            {
                return InvokeGuarded(application, displayName, action);
            }
            finally
            {
                SessionGate.Release();
            }
        }

        private static bool HasApplicationOverride(OfficeApp application)
        {
            if (application == OfficeApp.Excel)
            {
                return _excelApplicationOverride != null;
            }
            if (application == OfficeApp.Word)
            {
                return _wordApplicationOverride != null;
            }
            return _powerPointApplicationOverride != null;
        }

        private static string GetApplicationKey(OfficeApp application)
        {
            if (application == OfficeApp.Excel)
            {
                return "excel";
            }
            if (application == OfficeApp.Word)
            {
                return "word";
            }
            return "ppt";
        }

        /// <summary>执行委托并统一收敛异常；不涉及会话锁，由调用方决定加锁或走线程覆盖。</summary>
        private static OperationResult InvokeGuarded(
            OfficeApp application,
            string displayName,
            Func<OperationResult> action)
        {
            try
            {
                OperationResult result = action();
                if (result == null)
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "操作「" + displayName + "」未返回结果（执行器内部缺陷）。");
                }
                return result;
            }
            catch (OfficeComException ex)
            {
                return OperationResult.Failed(
                    OfficeErrorMapper.ToCode(ex.Code),
                    ex.Message,
                    ex.Recoverable,
                    null);
            }
            catch (COMException ex)
            {
                return OperationResult.Failed(
                    OfficeErrorMapper.MapException(ex, application),
                    OfficeErrorMapper.DescribeException(ex, application));
            }
            catch (Exception ex)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "操作「" + displayName + "」发生内部错误：" + ex.Message);
            }
        }

        /// <summary>覆盖作用域：Dispose 时恢复上一层实例；重复 Dispose 安全。</summary>
        private sealed class ApplicationOverrideScope : IDisposable
        {
            private Action _restore;
            private bool _disposed;

            public ApplicationOverrideScope(Action restore)
            {
                _restore = restore;
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }
                _disposed = true;
                Action restore = _restore;
                _restore = null;
                if (restore != null)
                {
                    restore();
                }
            }
        }

        private static ExcelApplication AttachOrLaunchExcel()
        {
            try
            {
                object activeObject = Marshal.GetActiveObject(ExcelProgId);
                ExcelApplication attached = activeObject as ExcelApplication;
                if (attached == null)
                {
                    throw new OfficeComException(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "已找到运行中的 Excel，但无法转换为强类型 Excel.Application。" +
                        "可能是 WPS 兼容层、权限不一致或 Office 版本不匹配；请关闭 WPS 后重试。");
                }
                _excelAttachmentMode = "attached";
                return attached;
            }
            catch (COMException ex)
            {
                if (ex.HResult != OfficeErrorMapper.MkEUnavailable)
                {
                    throw new OfficeComException(
                        OfficeErrorMapper.MapException(ex, OfficeApp.Excel),
                        "附着运行中的 Excel 失败：" + OfficeErrorMapper.DescribeException(ex, OfficeApp.Excel),
                        ex);
                }
            }

            if (!AppConfig.Current.ExcelAutoLaunch)
            {
                throw new OfficeComException(
                    OfficeErrorCode.APP_NOT_RUNNING,
                    "未找到运行中的 Excel，且 config.json 的 office.excel.autoLaunch 为 false。" +
                    "请先手动打开 Excel，或把该配置项改为 true 后重启 MCP server。");
            }

            ExcelApplication launched = null;
            try
            {
                launched = new ExcelApplication();
                // 可见性由 office.excel.visible 控制；默认可见以避免隐藏进程长期残留。
                launched.Visible = AppConfig.Current.ExcelVisible;
                _excelAttachmentMode = "launched";
                return launched;
            }
            catch (Exception ex)
            {
                TryQuitExcel(launched);
                throw new OfficeComException(
                    OfficeErrorMapper.MapException(ex, OfficeApp.Excel),
                    "启动 Excel 失败：" + OfficeErrorMapper.DescribeException(ex, OfficeApp.Excel),
                    ex);
            }
        }

        private static WordApplication AttachOrLaunchWord()
        {
            try
            {
                object activeObject = Marshal.GetActiveObject(WordProgId);
                WordApplication attached = activeObject as WordApplication;
                if (attached == null)
                {
                    throw new OfficeComException(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "已找到运行中的 Word，但无法转换为强类型 Word.Application。" +
                        "可能是 WPS 兼容层、权限不一致或 Office 版本不匹配；请关闭 WPS 后重试。");
                }
                _wordAttachmentMode = "attached";
                return attached;
            }
            catch (COMException ex)
            {
                if (ex.HResult != OfficeErrorMapper.MkEUnavailable)
                {
                    throw new OfficeComException(
                        OfficeErrorMapper.MapException(ex, OfficeApp.Word),
                        "附着运行中的 Word 失败：" + OfficeErrorMapper.DescribeException(ex, OfficeApp.Word),
                        ex);
                }
            }

            if (!AppConfig.Current.WordAutoLaunch)
            {
                throw new OfficeComException(
                    OfficeErrorCode.APP_NOT_RUNNING,
                    "未找到运行中的 Word，且 config.json 的 office.word.autoLaunch 为 false。" +
                    "请先手动打开 Word，或把该配置项改为 true 后重启 MCP server。");
            }

            WordApplication launched = null;
            try
            {
                launched = new WordApplication();
                launched.Visible = AppConfig.Current.WordVisible;
                _wordAttachmentMode = "launched";
                return launched;
            }
            catch (Exception ex)
            {
                TryQuitWord(launched);
                throw new OfficeComException(
                    OfficeErrorMapper.MapException(ex, OfficeApp.Word),
                    "启动 Word 失败：" + OfficeErrorMapper.DescribeException(ex, OfficeApp.Word),
                    ex);
            }
        }

        private static PowerPointApplication AttachOrLaunchPowerPoint()
        {
            try
            {
                object activeObject = Marshal.GetActiveObject(PowerPointProgId);
                PowerPointApplication attached = activeObject as PowerPointApplication;
                if (attached == null)
                {
                    throw new OfficeComException(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "已找到运行中的 PowerPoint，但无法转换为强类型 PowerPoint.Application。" +
                        "可能是 WPS 兼容层、权限不一致或 Office 版本不匹配；请关闭 WPS 后重试。");
                }
                _powerPointAttachmentMode = "attached";
                return attached;
            }
            catch (COMException ex)
            {
                if (ex.HResult != OfficeErrorMapper.MkEUnavailable)
                {
                    throw new OfficeComException(
                        OfficeErrorMapper.MapException(ex, OfficeApp.PowerPoint),
                        "附着运行中的 PowerPoint 失败：" +
                        OfficeErrorMapper.DescribeException(ex, OfficeApp.PowerPoint),
                        ex);
                }
            }

            if (!AppConfig.Current.PowerPointAutoLaunch)
            {
                throw new OfficeComException(
                    OfficeErrorCode.APP_NOT_RUNNING,
                    "未找到运行中的 PowerPoint，且 config.json 的 office.ppt.autoLaunch 为 false。" +
                    "请先手动打开 PowerPoint，或把该配置项改为 true 后重启 MCP server。");
            }

            PowerPointApplication launched = null;
            try
            {
                launched = new PowerPointApplication();
                dynamic dynamicLaunched = launched;
                dynamicLaunched.Visible = AppConfig.Current.PowerPointVisible ? -1 : 0;
                _powerPointAttachmentMode = "launched";
                return launched;
            }
            catch (Exception ex)
            {
                TryQuitPowerPoint(launched);
                throw new OfficeComException(
                    OfficeErrorMapper.MapException(ex, OfficeApp.PowerPoint),
                    "启动 PowerPoint 失败：" +
                    OfficeErrorMapper.DescribeException(ex, OfficeApp.PowerPoint),
                    ex);
            }
        }

        private static void TryQuitPowerPoint(PowerPointApplication application)
        {
            if (application == null)
            {
                return;
            }
            try
            {
                application.Quit();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[session] 清理启动失败的 PowerPoint 实例时 Quit 失败：" + ex.Message);
            }
            finally
            {
                ComHelper.Release(application);
            }
        }

        private static void TryQuitExcel(ExcelApplication application)
        {
            if (application == null)
            {
                return;
            }
            try
            {
                application.DisplayAlerts = false;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[session] 清理启动失败的 Excel 实例时关闭 DisplayAlerts 失败：" + ex.Message);
            }
            try
            {
                application.Quit();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[session] 清理启动失败的 Excel 实例时 Quit 失败：" + ex.Message);
            }
            finally
            {
                ComHelper.Release(application);
            }
        }

        private static void TryQuitWord(WordApplication application)
        {
            if (application == null)
            {
                return;
            }
            try
            {
                application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[session] 清理启动失败的 Word 实例时关闭 DisplayAlerts 失败：" + ex.Message);
            }
            try
            {
                application.Quit();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[session] 清理启动失败的 Word 实例时 Quit 失败：" + ex.Message);
            }
            finally
            {
                ComHelper.Release(application);
            }
        }
    }
}
