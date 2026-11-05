using System;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
using AutoOfficeMCP.Protocol;
using AutoOfficeMCP.Tools;
using Newtonsoft.Json.Linq;
using Office;
using Office.Config;
using Office.Contracts;
using Office.Parallel;
using Office.Undo;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace AutoOfficeMCP.Worker
{
    /// <summary>
    /// T11 worker 进程宿主：一个进程 = 一个 userId + 一个 app + 一个 STA 线程 + 一个常驻 Application。
    /// 不启动 HTTP/SSE/MCP 协议；只从命名管道读取单作业请求并返回结果。
    /// </summary>
    public static class WorkerHost
    {
        public static int Run(
            ToolRegistry registry,
            ExecutorDiscovery discovery,
            WorkerSettings settings)
        {
            ComHelper.EnsureSta("WorkerHost.Run");
            if (settings == null)
            {
                Console.Error.WriteLine("worker 启动设置为空。");
                return 1;
            }
            string pipeName = settings.PipeName;
            string userId = settings.UserId;
            string appType = settings.AppType;
            if (string.IsNullOrWhiteSpace(pipeName)
                || string.IsNullOrWhiteSpace(appType))
            {
                Console.Error.WriteLine("worker 缺少 pipeName / appType 设置。");
                return 1;
            }

            appType = appType.Trim().ToLowerInvariant();
            if (!string.Equals(appType, "excel", StringComparison.Ordinal)
                && !string.Equals(appType, "word", StringComparison.Ordinal)
                && !string.Equals(appType, "ppt", StringComparison.Ordinal))
            {
                Console.Error.WriteLine("worker 不支持的 app：" + appType);
                return 1;
            }
            if (string.IsNullOrWhiteSpace(userId))
            {
                userId = "default";
            }

            McpProtocolHandler handler = new McpProtocolHandler(
                registry,
                discovery,
                userId,
                "worker",
                "pipe");
            ParallelDispatcher.ToolInvoker = handler.ExecuteToolByName;
            ParallelDispatcher.ToolExists = delegate(string toolName)
            {
                return registry.Find(toolName) != null;
            };

            Excel.Application excel = null;
            Word.Application word = null;
            PowerPoint.Application powerPoint = null;
            PowerPointApplicationLease powerPointLease = null;
            int officePid = 0;
            NamedPipeClientStream pipe = null;
            StreamReader reader = null;
            StreamWriter writer = null;
            object writerSync = new object();
            Timer heartbeatTimer = null;
            Timer comWatchdogTimer = null;
            string expectedOfficeProcess = string.Empty;
            UndoStack undoStack = new UndoStack();
            try
            {
                if (string.Equals(appType, "excel", StringComparison.Ordinal))
                {
                    excel = new Excel.Application();
                    excel.Visible = AppConfig.Current.ExcelVisible;
                    excel.DisplayAlerts = false;
                    excel.ScreenUpdating = false;
                    officePid = OfficeProcessUtil.FromHwnd(excel.Hwnd);
                    OfficeAddinGuard.Disable(excel);
                }
                else if (string.Equals(appType, "ppt", StringComparison.Ordinal))
                {
                    // PowerPoint 是单实例应用：附着用户实例时不拥有它，退出时绝不 Quit；
                    // 只有确认没有运行实例时才创建并在退出时 Quit。
                    powerPointLease = PowerPointApplicationLease.Acquire(
                        AppConfig.Current.PowerPointVisible);
                    powerPoint = powerPointLease.Application;
                    // PowerPoint 不暴露 Application.Hwnd；按进程名取得 COM 目标 PID 供轮询。
                    officePid = OfficeProcessUtil.FindProcessIdByName("POWERPNT");
                    if (powerPointLease.Owned)
                    {
                        OfficeAddinGuard.Disable(powerPoint);
                    }
                }
                else
                {
                    word = new Word.Application();
                    word.Visible = AppConfig.Current.WordVisible;
                    word.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                    word.ScreenUpdating = false;
                    // Word 没有 Application.Hwnd；从专用模板创建一个临时文档拿到 Window.Hwnd，再立即关闭。
                    Word.Document scratch = Office.WordTools.WordTemplateFactory.CreateDocument(word);
                    Word.Window scratchWindow = scratch.ActiveWindow;
                    officePid = scratchWindow == null
                        ? 0
                        : OfficeProcessUtil.FromHwnd(scratchWindow.Hwnd);
                    Release(scratchWindow);
                    scratchWindow = null;
                    object saveScratch = Word.WdSaveOptions.wdDoNotSaveChanges;
                    object missingScratch = Type.Missing;
                    scratch.Close(ref saveScratch, ref missingScratch, ref missingScratch);
                    Release(scratch);
                    scratch = null;
                    OfficeAddinGuard.Disable(word);
                }

                expectedOfficeProcess = string.Equals(appType, "excel", StringComparison.Ordinal)
                    ? "EXCEL"
                    : (string.Equals(appType, "word", StringComparison.Ordinal) ? "WINWORD" : "POWERPNT");

                pipe = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut, PipeOptions.Asynchronous);
                pipe.Connect(Math.Max(5, AppConfig.Current.WorkerStartupTimeoutSec) * 1000);
                reader = new StreamReader(pipe, new UTF8Encoding(false), false, 4096, true);
                writer = new StreamWriter(pipe, new UTF8Encoding(false), 4096, true);
                writer.AutoFlush = true;
                writer.NewLine = "\n";

                WorkerControlMessage ready = new WorkerControlMessage();
                ready.Type = "ready";
                ready.UserId = userId;
                ready.App = appType;
                ready.OfficePid = officePid;
                ready.TimestampUtc = DateTime.UtcNow.ToString("o");
                WriteLine(writer, writerSync, WorkerProtocol.Serialize(ready));

                int heartbeatIntervalSec = AppConfig.Current.WorkerHeartbeatIntervalSec;
                heartbeatTimer = new Timer(
                    delegate { WriteHeartbeat(writer, writerSync, userId, appType, officePid); },
                    null,
                    TimeSpan.FromSeconds(heartbeatIntervalSec),
                    TimeSpan.FromSeconds(heartbeatIntervalSec));

                // COM 目标轮询：Office 被用户或意外关闭后，worker 必须自行结束，
                // 否则网关会继续把作业投递给已经失联的 COM 服务器。
                ComWatchState watchState = new ComWatchState();
                watchState.Pipe = pipe;
                watchState.Writer = writer;
                watchState.WriterSync = writerSync;
                watchState.HeartbeatTimer = heartbeatTimer;
                watchState.OfficePid = officePid;
                watchState.ExpectedProcessName = expectedOfficeProcess;
                watchState.AppType = appType;
                watchState.UserId = userId;
                int watchIntervalMs = Math.Max(2000, heartbeatIntervalSec * 1000);
                comWatchdogTimer = new Timer(
                    WatchComTarget,
                    watchState,
                    watchIntervalMs,
                    watchIntervalMs);

                while (true)
                {
                    string line = reader.ReadLine();
                    if (line == null)
                    {
                        break;
                    }
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    string type = WorkerProtocol.GetMessageType(line);
                    if (string.Equals(type, "shutdown", StringComparison.OrdinalIgnoreCase))
                    {
                        break;
                    }
                    if (string.Equals(type, "cancel", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.Error.WriteLine("worker 收到 cancel；当前同步 COM 调用无法中断，将在任务结束后退出/继续。");
                        continue;
                    }
                    if (!string.Equals(type, "job", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.Error.WriteLine("worker 忽略未知消息类型：" + type);
                        continue;
                    }

                    WorkerJobRequest request = WorkerProtocol.DeserializeRequest(line);
                    if (officePid > 0 &&
                        !OfficeProcessUtil.IsAliveOfficeProcess(officePid, expectedOfficeProcess))
                    {
                        // 轮询周期之间 COM 目标可能刚被关闭：不执行作业，直接返回明确原因并退出。
                        WorkerJobResponse lostResponse = new WorkerJobResponse();
                        lostResponse.JobId = request == null ? string.Empty : (request.JobId ?? string.Empty);
                        lostResponse.Success = false;
                        lostResponse.ErrorCode = OfficeErrorMapper.ToCode(OfficeErrorCode.APP_NOT_RUNNING);
                        lostResponse.Message = "COM 目标进程（" + expectedOfficeProcess + "，pid=" +
                                               officePid.ToString() +
                                               "）已退出，本次作业未执行；worker 将自行结束，下次请求会重建。";
                        lostResponse.Recoverable = true;
                        WriteLine(writer, writerSync, WorkerProtocol.Serialize(lostResponse));
                        break;
                    }

                    WorkerJobResponse response = ExecuteJob(request, excel, word, powerPoint, undoStack, userId, appType);
                    WriteLine(writer, writerSync, WorkerProtocol.Serialize(response));
                }
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("worker 运行失败。" + Environment.NewLine + ex);
                return 1;
            }
            finally
            {
                if (comWatchdogTimer != null)
                {
                    comWatchdogTimer.Dispose();
                    comWatchdogTimer = null;
                }
                if (heartbeatTimer != null)
                {
                    heartbeatTimer.Dispose();
                    heartbeatTimer = null;
                }
                try
                {
                    if (writer != null)
                    {
                        writer.Flush();
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[worker] flush 管道失败：" + ex.Message);
                }
                DisposeQuietly(reader);
                DisposeQuietly(writer);
                DisposeQuietly(pipe);
                QuitExcel(excel);
                QuitWord(word);
                if (powerPointLease != null)
                {
                    try
                    {
                        powerPointLease.Dispose();
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[worker] 释放 PowerPoint 租约失败：" + ex.Message);
                    }
                    powerPointLease = null;
                    powerPoint = null;
                }
            }
        }

        private static WorkerJobResponse ExecuteJob(
            WorkerJobRequest request,
            Excel.Application excel,
            Word.Application word,
            PowerPoint.Application powerPoint,
            UndoStack undoStack,
            string userId,
            string appType)
        {
            WorkerJobResponse response = new WorkerJobResponse();
            response.JobId = request == null ? string.Empty : (request.JobId ?? string.Empty);
            if (request == null)
            {
                response.Success = false;
                response.ErrorCode = "INVALID_PARAMS";
                response.Message = "worker 收到空作业。";
                response.Recoverable = false;
                return response;
            }

            OfficeWorkItem item = new OfficeWorkItem();
            item.AppType = string.IsNullOrWhiteSpace(request.App) ? appType : request.App;
            item.FilePath = request.File ?? string.Empty;
            item.ToolName = request.Tool ?? string.Empty;
            item.Args = request.Arguments == null ? new JObject() : (JObject)request.Arguments.DeepClone();
            item.SavePolicy = string.IsNullOrWhiteSpace(request.SavePolicy) ? "none" : request.SavePolicy;
            item.TimeoutSec = request.TimeoutSec;
            item.AffinityKey = request.AffinityKey ?? string.Empty;
            item.Priority = request.Priority;
            item.IdempotencyKey = request.IdempotencyKey ?? string.Empty;

            // Gateway 已规范化一次；worker 再校验一次，保证直接投递到管道的作业也不能绕过工作目录规则。
            JObject normalizedArguments;
            string argumentError;
            if (!WorkspacePathResolver.TryNormalizeArguments(
                    item.AppType,
                    item.Args,
                    out normalizedArguments,
                    out argumentError))
            {
                response.Success = false;
                response.ErrorCode = "INVALID_PARAMS";
                response.Message = argumentError;
                response.Recoverable = false;
                return response;
            }
            item.Args = normalizedArguments;

            string resolvedFile;
            string fileError;
            if (!WorkspacePathResolver.TryResolveDocument(
                    item.AppType,
                    item.FilePath,
                    out resolvedFile,
                    out fileError))
            {
                response.Success = false;
                response.ErrorCode = "INVALID_PARAMS";
                response.Message = "file 必须是纯文件名、工作目录内相对子路径或本机绝对路径，且扩展名与 appType 匹配：" + fileError;
                response.Recoverable = false;
                return response;
            }
            item.FilePath = resolvedFile;

            try
            {
                using (OfficeSession.UseUndoStack(undoStack))
                {
                    ParallelJobOutcome outcome;
                    if (string.Equals(appType, "excel", StringComparison.Ordinal))
                    {
                        using (OfficeSession.UseExcelApplication(excel))
                        {
                            outcome = ParallelStaJobRunner.RunWithApplication(item, excel);
                        }
                    }
                    else if (string.Equals(appType, "ppt", StringComparison.Ordinal))
                    {
                        using (OfficeSession.UsePowerPointApplication(powerPoint))
                        {
                            outcome = ParallelStaJobRunner.RunWithApplication(item, powerPoint);
                        }
                    }
                    else
                    {
                        using (OfficeSession.UseWordApplication(word))
                        {
                            outcome = ParallelStaJobRunner.RunWithApplication(item, word);
                        }
                    }
                    FillResponse(response, outcome);
                }
            }
            catch (Exception ex)
            {
                response.Success = false;
                OfficeApp failedApp = string.Equals(appType, "excel", StringComparison.Ordinal)
                    ? OfficeApp.Excel
                    : (string.Equals(appType, "ppt", StringComparison.Ordinal)
                        ? OfficeApp.PowerPoint
                        : OfficeApp.Word);
                response.ErrorCode = OfficeErrorMapper.ToCode(
                    OfficeErrorMapper.MapException(ex, failedApp));
                response.Message = "worker 执行作业异常：" + ex.Message;
                response.Diagnostic = OfficeErrorMapper.FormatExceptionChain(ex);
                response.Recoverable = false;
            }
            finally
            {
                try
                {
                    undoStack.Clear();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[worker] 清空撤销栈失败：" + ex.Message);
                }
            }
            return response;
        }

        private static void FillResponse(WorkerJobResponse response, ParallelJobOutcome outcome)
        {
            if (outcome == null || outcome.Result == null)
            {
                response.Success = false;
                response.ErrorCode = "INTERNAL_ERROR";
                response.Message = "worker 没有返回作业结果。";
                response.Recoverable = false;
                return;
            }

            OperationResult result = outcome.Result;
            response.Success = result.Success;
            response.ErrorCode = result.ErrorCode ?? string.Empty;
            response.Message = result.UserMessage ?? string.Empty;
            response.Recoverable = result.Recoverable;
            response.Data = result.Data == null ? null : (JObject)result.Data.DeepClone();
            response.Observation = result.Observation == null ? null : result.Observation.ToJObject();
            response.RollbackNote = result.RollbackNote ?? string.Empty;
            response.Diagnostic = result.Diagnostic ?? string.Empty;
            response.Saved = outcome.Saved;
            response.OpenedReadOnly = outcome.OpenedReadOnly;
            response.SaveSkippedReason = outcome.SaveSkippedReason ?? string.Empty;
            response.UndoNodesRemoved = outcome.UndoNodesRemoved;
            response.SaveVerified = outcome.SaveVerified;
            response.SaveVerificationNote = outcome.SaveVerificationNote ?? string.Empty;
            response.SaveChangedOnDisk = outcome.SaveChangedOnDisk;
            response.SaveBeforeHash = outcome.SaveBeforeHash ?? string.Empty;
            response.SaveAfterHash = outcome.SaveAfterHash ?? string.Empty;
            response.SaveBeforeLength = outcome.SaveBeforeLength;
            response.SaveAfterLength = outcome.SaveAfterLength;
            response.SaveBeforeLastWriteUtc = outcome.SaveBeforeLastWriteUtc ?? string.Empty;
            response.SaveAfterLastWriteUtc = outcome.SaveAfterLastWriteUtc ?? string.Empty;
        }

        private static void QuitExcel(Excel.Application application)
        {
            if (application == null)
            {
                return;
            }
            try
            {
                application.DisplayAlerts = false;
                application.Quit();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[worker] Quit Excel 失败：" + ex.Message);
            }
            Release(application);
        }

        private static void QuitWord(Word.Application application)
        {
            if (application == null)
            {
                return;
            }
            try
            {
                application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                object missing = Type.Missing;
                application.Quit(ref saveChanges, ref missing, ref missing);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[worker] Quit Word 失败：" + ex.Message);
            }
            Release(application);
        }

        private static void QuitPowerPoint(PowerPoint.Application application)
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
                Console.Error.WriteLine("[worker] Quit PowerPoint 失败：" + ex.Message);
            }
            Release(application);
        }

        private static void WriteLine(StreamWriter writer, object sync, string line)
        {
            if (writer == null || line == null)
            {
                return;
            }
            lock (sync)
            {
                writer.WriteLine(line);
                writer.Flush();
            }
        }

        /// <summary>
        /// COM 目标进程轮询：按 PID + 进程名确认 Office 是否仍存活。
        /// 目标丢失时先向网关发送 error 控制消息（让在飞作业立即拿到原因），
        /// 再停止心跳、关闭管道让主循环退出；3 秒内未退出则强制结束 worker。
        /// </summary>
        private static void WatchComTarget(object state)
        {
            ComWatchState watch = state as ComWatchState;
            if (watch == null || watch.OfficePid <= 0)
            {
                return;
            }
            if (OfficeProcessUtil.IsAliveOfficeProcess(watch.OfficePid, watch.ExpectedProcessName))
            {
                return;
            }
            if (Interlocked.Exchange(ref watch.Lost, 1) != 0)
            {
                return;
            }

            string message = "COM 目标进程已退出（app=" + watch.AppType + "，pid=" +
                             watch.OfficePid.ToString() + "），worker 将自行结束；下次请求会重建 worker。";
            Console.Error.WriteLine("[worker] " + message);

            WorkerControlMessage notice = new WorkerControlMessage();
            notice.Type = "error";
            notice.Message = message;
            notice.UserId = watch.UserId;
            notice.App = watch.AppType;
            notice.OfficePid = watch.OfficePid;
            notice.TimestampUtc = DateTime.UtcNow.ToString("o");
            try
            {
                WriteLine(watch.Writer, watch.WriterSync, WorkerProtocol.Serialize(notice));
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[worker] 上报 COM 目标退出失败：" + ex.Message);
            }

            try
            {
                if (watch.HeartbeatTimer != null)
                {
                    watch.HeartbeatTimer.Change(Timeout.Infinite, Timeout.Infinite);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[worker] 停止心跳失败：" + ex.Message);
            }

            try
            {
                if (watch.Pipe != null)
                {
                    watch.Pipe.Dispose();
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[worker] 关闭管道失败：" + ex.Message);
            }

            ThreadPool.QueueUserWorkItem(delegate
            {
                Thread.Sleep(3000);
                Console.Error.WriteLine("[worker] COM 目标已退出且主循环未结束，强制退出 worker 进程。");
                Environment.Exit(0);
            });
        }

        private static void WriteHeartbeat(
            StreamWriter writer,
            object sync,
            string userId,
            string appType,
            int officePid)
        {
            try
            {
                WorkerControlMessage heartbeat = new WorkerControlMessage();
                heartbeat.Type = "heartbeat";
                heartbeat.UserId = userId;
                heartbeat.App = appType;
                heartbeat.OfficePid = officePid;
                heartbeat.TimestampUtc = DateTime.UtcNow.ToString("o");
                WriteLine(writer, sync, WorkerProtocol.Serialize(heartbeat));
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[worker] 发送心跳失败：" + ex.Message);
            }
        }

        private static void DisposeQuietly(IDisposable disposable)
        {
            if (disposable == null)
            {
                return;
            }
            try
            {
                disposable.Dispose();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[worker] 释放流失败：" + ex.Message);
            }
        }

        /// <summary>COM 目标轮询所需的 worker 上下文；只由 watchdog 定时器读取。</summary>
        private sealed class ComWatchState
        {
            public NamedPipeClientStream Pipe;
            public StreamWriter Writer;
            public object WriterSync;
            public Timer HeartbeatTimer;
            public int OfficePid;
            public string ExpectedProcessName;
            public string AppType;
            public string UserId;
            public int Lost;
        }

        private static void Release(object comObject)
        {
            if (comObject != null && System.Runtime.InteropServices.Marshal.IsComObject(comObject))
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(comObject);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[worker] 释放 COM 对象失败：" + ex.Message);
                }
            }
        }
    }
}
