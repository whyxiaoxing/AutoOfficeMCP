using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Office.Config;
using Office.Contracts;

namespace Office.Parallel
{
    /// <summary>
    /// 并行调度器（单例，手动管理生命周期）。
    ///
    /// 分工：
    /// 1) 调度线程只负责排队与异步等待，不执行 COM；
    /// 2) 每个 appType 一把 SemaphoreSlim 手动限流；
    /// 3) NamedFileLockRegistry / SqlFileLockProvider 保证同一规范化路径永远串行；
    /// 4) COM 执行段由 WorkerProcessPool 的独立进程/STA 线程完成；
    /// 5) 结果只通过 TaskCompletionSource&lt;OperationResult&gt; 回到调用方，调度线程不碰 COM。
    /// </summary>
    public sealed class ParallelDispatcher
    {
        private const int DefaultRecentErrorLimit = 30;

        private static readonly object InstanceSync = new object();
        private static ParallelDispatcher _instance;

        private readonly int _excelMaxConcurrency;
        private readonly int _wordMaxConcurrency;
        private readonly int _pptMaxConcurrency;
        private readonly int _itemTimeoutSec;
        private readonly int _maxQueue;
        private readonly string _concurrencyMode;
        private readonly SemaphoreSlim _excelGate;
        private readonly SemaphoreSlim _wordGate;
        private readonly SemaphoreSlim _pptGate;
        private readonly IFileLockProvider _fileLocks;
        private readonly WorkerProcessPool _workerPool;
        private readonly ConcurrentDictionary<long, ParallelJobContext> _inFlight;
        private readonly ConcurrentQueue<JObject> _recentErrors;
        private readonly object _queueSync;
        private readonly int _idempotencyTtlSec;
        private readonly ConcurrentDictionary<string, Task<OperationResult>> _idempotencyInFlight;
        private readonly ConcurrentDictionary<string, IdempotencyCacheEntry> _idempotencyCompleted;

        private int _pendingCount;
        private int _runningCount;
        private int _excelRunning;
        private int _wordRunning;
        private int _pptRunning;
        private long _nextJobId;
        private long _throughputSequence;
        private volatile bool _shuttingDown;

        /// <summary>
        /// 工具执行桥：Program 在构造 McpProtocolHandler 后注入 handler.ExecuteToolByName。
        /// 与 T06-Common 的 CommonToolExecutor.ToolDispatcher 并列，复用注册表查找与高危确认。
        /// </summary>
        public static Func<string, JObject, ToolResult> ToolInvoker { get; set; }

        /// <summary>
        /// 工具注册校验桥：Program 注入 registry.Find != null。未注入时跳过预校验，
        /// 由 ToolInvoker 在运行时返回 INVALID_PARAMS。
        /// </summary>
        public static Func<string, bool> ToolExists { get; set; }

        /// <summary>进程级单例；首次访问时读取 config.json 的 parallel 段。</summary>
        public static ParallelDispatcher Instance
        {
            get
            {
                lock (InstanceSync)
                {
                    if (_instance == null)
                    {
                        _instance = new ParallelDispatcher();
                    }
                    return _instance;
                }
            }
        }

        private ParallelDispatcher()
        {
            AppConfig config = AppConfig.Current;
            _excelMaxConcurrency = Math.Max(1, config.ExcelMaxConcurrency);
            _wordMaxConcurrency = Math.Max(1, config.WordMaxConcurrency);
            // PowerPoint 是单实例应用；槽位数由 parallel.ppt.maxConcurrency 控制。
            // 真正的隔离并发仍受 worker 进程/Windows 会话限制，见 README。
            _pptMaxConcurrency = Math.Max(1, config.PowerPointMaxConcurrency);
            _itemTimeoutSec = Math.Max(1, config.ParallelItemTimeoutSec);
            _maxQueue = Math.Max(1, config.ParallelMaxQueue);
            _concurrencyMode = config.WorkerConcurrencyMode;

            _excelGate = new SemaphoreSlim(_excelMaxConcurrency, _excelMaxConcurrency);
            _wordGate = new SemaphoreSlim(_wordMaxConcurrency, _wordMaxConcurrency);
            _pptGate = new SemaphoreSlim(_pptMaxConcurrency, _pptMaxConcurrency);
            if (string.Equals(config.WorkerFileLockProvider, "sql", StringComparison.Ordinal)
                && !string.IsNullOrWhiteSpace(config.WorkerSqlLockConnectionString))
            {
                _fileLocks = new SqlFileLockProvider(config.WorkerSqlLockConnectionString);
            }
            else
            {
                if (string.Equals(config.WorkerFileLockProvider, "sql", StringComparison.Ordinal))
                {
                    Console.Error.WriteLine("[parallel] fileLockProvider=sql 但 sqlLockConnectionString 为空，回退锁文件。");
                }
                _fileLocks = new NamedFileLockRegistry();
            }
            _workerPool = new WorkerProcessPool();
            _inFlight = new ConcurrentDictionary<long, ParallelJobContext>();
            _recentErrors = new ConcurrentQueue<JObject>();
            _queueSync = new object();
            _idempotencyTtlSec = config.ParallelIdempotencyTtlSec;
            _idempotencyInFlight = new ConcurrentDictionary<string, Task<OperationResult>>(StringComparer.Ordinal);
            _idempotencyCompleted = new ConcurrentDictionary<string, IdempotencyCacheEntry>(StringComparer.Ordinal);
        }

        /// <summary>规范化文件路径；空/非法路径返回空字符串。</summary>
        private static string NormalizePath(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return string.Empty;
            }

            try
            {
                return Path.GetFullPath(filePath.Trim()).ToUpperInvariant();
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>items 数量上限；office_batch_files 在入口按此拒绝超量批次。</summary>
        public int MaxQueue
        {
            get { return _maxQueue; }
        }

        /// <summary>全局默认单项超时秒数（供批次等待上限参考）。</summary>
        public int ItemTimeoutSec
        {
            get { return _itemTimeoutSec; }
        }

        // ------------------------------------------------------------------
        // 投递
        // ------------------------------------------------------------------

        /// <summary>
        /// 校验并投递一个作业。队列满、参数非法、文件不存在等立即返回失败结果，
        /// 不会抛异常。返回的 Task 与 item.Completion.Task 相同。
        /// </summary>
        public Task<OperationResult> Submit(OfficeWorkItem item)
        {
            if (item == null)
            {
                throw new ArgumentNullException("item");
            }

            string appType = ParallelToolCatalog.NormalizeApp(item.AppType);
            if (appType.Length == 0)
            {
                return Reject(item, OfficeErrorCode.INVALID_PARAMS,
                    "appType 必须是 excel、word 或 ppt，收到「" + (item.AppType ?? string.Empty) + "」。");
            }

            string toolName = string.IsNullOrWhiteSpace(item.ToolName) ? string.Empty : item.ToolName.Trim();
            if (toolName.Length == 0)
            {
                return Reject(item, OfficeErrorCode.INVALID_PARAMS, "tool 必须是非空工具名。");
            }

            if (!ParallelToolCatalog.IsToolForApp(appType, toolName))
            {
                return Reject(item, OfficeErrorCode.INVALID_PARAMS,
                    "工具「" + toolName + "」不属于 appType=" + appType +
                    "（工具名前缀必须为 " + appType + "_），已拒绝投递。");
            }

            if (ToolExists != null)
            {
                bool exists;
                try
                {
                    exists = ToolExists(toolName);
                }
                catch (Exception ex)
                {
                    Log("工具注册校验失败：" + toolName + "；" + ex.Message);
                    exists = false;
                }
                if (!exists)
                {
                    return Reject(item, OfficeErrorCode.INVALID_PARAMS,
                        "工具「" + toolName + "」未注册或已在 config.json 中禁用，已拒绝投递。");
                }
            }

            string savePolicy = ParallelToolCatalog.NormalizeSavePolicy(item.SavePolicy);
            if (savePolicy == null)
            {
                return Reject(item, OfficeErrorCode.INVALID_PARAMS,
                    "savePolicy 只允许 none 或 save，收到「" + (item.SavePolicy ?? string.Empty) + "」。");
            }

            JObject normalizedArguments;
            string argumentError;
            if (!WorkspacePathResolver.TryNormalizeArguments(
                    appType,
                    item.Args,
                    out normalizedArguments,
                    out argumentError))
            {
                return Reject(item, OfficeErrorCode.INVALID_PARAMS, argumentError);
            }
            item.Args = normalizedArguments;

            string fullPath;
            string fileError;
            if (!WorkspacePathResolver.TryResolveDocument(
                    appType,
                    item.FilePath,
                    out fullPath,
                    out fileError))
            {
                return Reject(item, OfficeErrorCode.INVALID_PARAMS,
                    "file 必须是纯文件名、工作目录内相对子路径或本机绝对路径，且扩展名与 appType 匹配：" + fileError);
            }
            if (!File.Exists(fullPath))
            {
                return Reject(item, OfficeErrorCode.FILE_NOT_FOUND,
                    "文件不存在或路径不可访问：" + fullPath);
            }

            string lockKey = NormalizePath(fullPath);
            if (lockKey.Length == 0)
            {
                return Reject(item, OfficeErrorCode.INVALID_PARAMS,
                    "无法规范化文件路径：" + fullPath);
            }

            int timeoutSec = item.TimeoutSec > 0 ? item.TimeoutSec : _itemTimeoutSec;
            item.AppType = appType;
            item.ToolName = toolName;
            item.FilePath = fullPath;
            item.SavePolicy = savePolicy;

            string idempotencyKey = NormalizeIdempotencyKey(
                ToolExecutionContext.UserId,
                appType,
                toolName,
                item.IdempotencyKey);
            if (idempotencyKey.Length > 0)
            {
                CleanupIdempotencyCache();
                IdempotencyCacheEntry cached;
                if (_idempotencyCompleted.TryGetValue(idempotencyKey, out cached) && cached != null)
                {
                    item.Completion.TrySetResult(cached.Result);
                    return Task.FromResult(cached.Result);
                }
                Task<OperationResult> inFlightTask;
                if (_idempotencyInFlight.TryGetValue(idempotencyKey, out inFlightTask) && inFlightTask != null)
                {
                    Task<OperationResult> joined = inFlightTask;
                    joined.ContinueWith(
                        delegate(Task<OperationResult> finished)
                        {
                            if (finished.Status == TaskStatus.RanToCompletion)
                            {
                                item.Completion.TrySetResult(finished.Result);
                            }
                            else
                            {
                                string reason = finished.Exception == null
                                    ? "未知原因"
                                    : finished.Exception.GetBaseException().Message;
                                item.Completion.TrySetResult(OperationResult.Failed(
                                    OfficeErrorCode.INTERNAL_ERROR,
                                    "复用幂等作业失败：" + reason));
                            }
                        },
                        TaskScheduler.Default);
                    return joined;
                }
            }

            lock (_queueSync)
            {
                if (_shuttingDown)
                {
                    return Reject(item, OfficeErrorCode.OFFICE_BUSY,
                        "并行调度器正在关闭，作业未入队。");
                }
                if (_pendingCount >= _maxQueue)
                {
                    return Reject(item, OfficeErrorCode.OFFICE_BUSY,
                        "并行队列已满（maxQueue=" + _maxQueue.ToString(CultureInfo.InvariantCulture) +
                        "），作业未入队；请稍后重试或拆分批次。");
                }
                _pendingCount++;
            }

            long jobId = Interlocked.Increment(ref _nextJobId);
            string workerUserId = ToolExecutionContext.UserId;
            if (string.Equals(_concurrencyMode, "throughput", StringComparison.Ordinal))
            {
                int slots = string.Equals(appType, ParallelToolCatalog.Excel, StringComparison.Ordinal)
                    ? _excelMaxConcurrency
                    : (string.Equals(appType, ParallelToolCatalog.Ppt, StringComparison.Ordinal)
                        ? _pptMaxConcurrency
                        : _wordMaxConcurrency);
                long sequence = Interlocked.Increment(ref _throughputSequence);
                int slot = (int)((sequence - 1) % slots);
                workerUserId = "throughput-" + appType + "-" + slot.ToString(CultureInfo.InvariantCulture);
            }

            ParallelJobContext job = new ParallelJobContext(
                jobId, item, appType, fullPath, lockKey, savePolicy, timeoutSec,
                workerUserId);
            job.AffinityKey = string.IsNullOrWhiteSpace(item.AffinityKey) ? string.Empty : item.AffinityKey.Trim();
            job.IdempotencyKey = idempotencyKey;
            _inFlight[jobId] = job;
            if (idempotencyKey.Length > 0)
            {
                _idempotencyInFlight[idempotencyKey] = item.Completion.Task;
            }

            try
            {
                // 异步调度：不占用 ThreadPool 线程等待 COM；COM 由常驻 STA worker 执行。
                Task ignored = ScheduleAsync(job);
            }
            catch (Exception ex)
            {
                Log("投递异步调度失败：" + ex.Message);
                ParallelJobContext removed;
                _inFlight.TryRemove(jobId, out removed);
                if (idempotencyKey.Length > 0)
                {
                    Task<OperationResult> ignoredTask;
                    _idempotencyInFlight.TryRemove(idempotencyKey, out ignoredTask);
                }
                Interlocked.Decrement(ref _pendingCount);
                return Reject(item, OfficeErrorCode.OFFICE_BUSY,
                    "无法投递并行作业（调度器异常）：" + ex.Message);
            }

            return item.Completion.Task;
        }

        // ------------------------------------------------------------------
        // 调度（ThreadPool 线程；await 期间不占用池线程）
        // ------------------------------------------------------------------

        private async Task ScheduleAsync(ParallelJobContext job)
        {
            OperationResult result = null;
            SemaphoreSlim appGate = null;
            bool appSlotAcquired = false;
            IDisposable fileLease = null;
            IDisposable affinityLease = null;

            try
            {
                job.Stage = "waiting-app-slot";
                appGate = GetGate(job.AppType);
                TimeSpan remaining = job.Remaining;
                if (remaining <= TimeSpan.Zero)
                {
                    result = BuildTimeout(job, "排队等待应用并发槽", true);
                    return;
                }

                appSlotAcquired = await appGate.WaitAsync(remaining).ConfigureAwait(false);
                if (!appSlotAcquired)
                {
                    result = BuildTimeout(job, "等待应用并发槽", true);
                    return;
                }

                Interlocked.Increment(ref _runningCount);
                if (string.Equals(job.AppType, ParallelToolCatalog.Excel, StringComparison.Ordinal))
                {
                    Interlocked.Increment(ref _excelRunning);
                }
                else if (string.Equals(job.AppType, ParallelToolCatalog.Ppt, StringComparison.Ordinal))
                {
                    Interlocked.Increment(ref _pptRunning);
                }
                else
                {
                    Interlocked.Increment(ref _wordRunning);
                }

                job.Stage = "waiting-file-lock";
                remaining = job.Remaining;
                if (remaining <= TimeSpan.Zero)
                {
                    result = BuildTimeout(job, "等待文件锁", true);
                    return;
                }

                fileLease = await _fileLocks.AcquireAsync(job.LockKey, job.FilePath, remaining).ConfigureAwait(false);
                if (fileLease == null)
                {
                    result = BuildTimeout(job, "等待文件锁", true);
                    return;
                }

                if (!string.IsNullOrWhiteSpace(job.AffinityKey))
                {
                    job.Stage = "waiting-affinity-lock";
                    remaining = job.Remaining;
                    if (remaining <= TimeSpan.Zero)
                    {
                        result = BuildTimeout(job, "等待亲和锁", true);
                        return;
                    }
                    string affinityLockKey = "affinity:" + job.AffinityKey;
                    affinityLease = await _fileLocks.AcquireAsync(
                        affinityLockKey,
                        affinityLockKey,
                        remaining).ConfigureAwait(false);
                    if (affinityLease == null)
                    {
                        result = BuildTimeout(job, "等待亲和锁", true);
                        return;
                    }
                }

                job.FileVersionBefore = FileIntegrityVerifier.Capture(
                    job.FilePath,
                    AppConfig.Current.WorkerSaveVerifyMaxHashMb);
                job.Stage = "running";
                job.ExecutionStartedUtc = DateTime.UtcNow;
                Log("开始 job=" + job.Id.ToString(CultureInfo.InvariantCulture) +
                    " app=" + job.AppType + " tool=" + job.Item.ToolName +
                    " file=" + job.FilePath +
                    " startedUtc=" + job.ExecutionStartedUtc.Value.ToString("o", CultureInfo.InvariantCulture));

                // 常驻 STA worker：同用户同应用串行，不同用户/应用并行。
                Task<ParallelJobOutcome> workerTask =
                    _workerPool.EnqueueAsync(job.UserId, job.AppType, job.Item);
                remaining = job.Remaining;
                if (remaining <= TimeSpan.Zero)
                {
                    result = BuildTimeout(job, "worker 排队", false);
                    return;
                }

                Task completed = await Task.WhenAny(workerTask, Task.Delay(remaining)).ConfigureAwait(false);
                if (completed == workerTask)
                {
                    job.Outcome = await workerTask.ConfigureAwait(false);
                    result = job.Outcome == null ? null : job.Outcome.Result;
                    if (result == null)
                    {
                        result = OperationResult.Failed(
                            OfficeErrorCode.INTERNAL_ERROR,
                            "常驻 STA worker 没有返回结果（调度器内部缺陷）。");
                    }
                    result = AttachMetadata(job, result, job.Outcome, true, "completed");
                }
                else
                {
                    result = BuildTimeout(job, "执行工具（常驻 STA worker）", false);
                }
            }
            catch (Exception ex)
            {
                result = OperationResult.Failed(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "并行作业调度异常：" + ex.Message);
                result = AttachMetadata(job, result, null, true, "scheduler_error");
            }
            finally
            {
                job.Stage = "finished";
                if (affinityLease != null)
                {
                    try
                    {
                        affinityLease.Dispose();
                    }
                    catch (Exception ex)
                    {
                        Log("释放亲和锁失败：" + job.AffinityKey + "；" + ex.Message);
                    }
                }
                if (fileLease != null)
                {
                    try
                    {
                        fileLease.Dispose();
                    }
                    catch (Exception ex)
                    {
                        Log("释放文件锁失败：" + job.FilePath + "；" + ex.Message);
                    }
                }
                if (appSlotAcquired && appGate != null)
                {
                    try
                    {
                        appGate.Release();
                    }
                    catch (Exception ex)
                    {
                        Log("释放应用并发槽失败：" + job.AppType + "；" + ex.Message);
                    }
                    if (string.Equals(job.AppType, ParallelToolCatalog.Excel, StringComparison.Ordinal))
                    {
                        Interlocked.Decrement(ref _excelRunning);
                    }
                    else if (string.Equals(job.AppType, ParallelToolCatalog.Ppt, StringComparison.Ordinal))
                    {
                        Interlocked.Decrement(ref _pptRunning);
                    }
                    else
                    {
                        Interlocked.Decrement(ref _wordRunning);
                    }
                    Interlocked.Decrement(ref _runningCount);
                }

                OnJobFinished(job, result);
            }
        }

        /// <summary>一次性 STA 线程入口；只做 COM 作业与结果赋值，不操作调度器锁。</summary>
        private void ExecuteOnStaThread(ParallelJobContext job)
        {
            ParallelJobOutcome outcome = null;
            try
            {
                try
                {
                    outcome = ParallelStaJobRunner.Run(job.Item);
                }
                catch (Exception ex)
                {
                    outcome = new ParallelJobOutcome();
                    OfficeApp failedApp = string.Equals(job.AppType, ParallelToolCatalog.Excel, StringComparison.Ordinal)
                        ? OfficeApp.Excel
                        : (string.Equals(job.AppType, ParallelToolCatalog.Ppt, StringComparison.Ordinal)
                            ? OfficeApp.PowerPoint
                            : OfficeApp.Word);
                    outcome.Result = OfficeErrorMapper.ToOperationResult(
                        ex,
                        failedApp,
                        "池内一次性 STA 作业发生未捕获异常；不会自动重放。");
                }
            }
            finally
            {
                job.Outcome = outcome;
                job.Done.Set();
            }
        }

        // ------------------------------------------------------------------
        // 结果、状态与生命周期
        // ------------------------------------------------------------------

        private void OnJobFinished(ParallelJobContext job, OperationResult result)
        {
            ParallelJobContext removed;
            _inFlight.TryRemove(job.Id, out removed);
            Interlocked.Decrement(ref _pendingCount);

            if (result == null)
            {
                result = OperationResult.Failed(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "并行作业未产生结果（调度器内部缺陷）。");
            }

            JObject parallel = result.Data == null ? null : result.Data["parallel"] as JObject;
            string savedText = parallel == null ? "n/a" : (parallel.Value<bool?>("saved") ?? false).ToString();
            string threadExitedText = parallel == null ? "n/a" : (parallel.Value<bool?>("threadExited") ?? false).ToString();
            Log("结束 job=" + job.Id.ToString(CultureInfo.InvariantCulture) +
                " success=" + result.Success.ToString() +
                " errorCode=" + (result.ErrorCode ?? string.Empty) +
                " saved=" + savedText +
                " threadExited=" + threadExitedText +
                " elapsedMs=" + ((long)job.Clock.Elapsed.TotalMilliseconds).ToString(CultureInfo.InvariantCulture));

            if (!result.Success)
            {
                AddRecentError(job, result);
            }

            if (!string.IsNullOrWhiteSpace(job.IdempotencyKey))
            {
                Task<OperationResult> ignoredTask;
                _idempotencyInFlight.TryRemove(job.IdempotencyKey, out ignoredTask);
                if (_idempotencyTtlSec > 0)
                {
                    _idempotencyCompleted[job.IdempotencyKey] =
                        new IdempotencyCacheEntry(result, DateTime.UtcNow);
                }
            }

            job.Item.Completion.TrySetResult(result);
        }

        private void AddRecentError(ParallelJobContext job, OperationResult result)
        {
            JObject error = new JObject();
            error["timeUtc"] = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
            error["jobId"] = job.Id;
            error["appType"] = job.AppType;
            error["tool"] = job.Item.ToolName;
            error["file"] = job.FilePath;
            error["stage"] = job.Stage;
            error["errorCode"] = result.ErrorCode ?? string.Empty;
            error["message"] = result.UserMessage ?? string.Empty;
            _recentErrors.Enqueue(error);

            while (_recentErrors.Count > DefaultRecentErrorLimit)
            {
                JObject discarded;
                _recentErrors.TryDequeue(out discarded);
            }
        }

        /// <summary>office_pool_status 的数据源：队列/在跑/文件锁/最近错误。</summary>
        public JObject GetStatus()
        {
            int pending = Volatile.Read(ref _pendingCount);
            int running = Volatile.Read(ref _runningCount);

            JObject status = new JObject();
            status["shuttingDown"] = _shuttingDown;
            status["concurrencyMode"] = _concurrencyMode;
            status["fileLockProvider"] = _fileLocks is SqlFileLockProvider ? "sql" : "file";
            status["pendingCount"] = pending;
            status["queueDepth"] = Math.Max(0, pending - running);
            status["runningCount"] = running;

            JObject runningByApp = new JObject();
            runningByApp["excel"] = Volatile.Read(ref _excelRunning);
            runningByApp["word"] = Volatile.Read(ref _wordRunning);
            runningByApp["ppt"] = Volatile.Read(ref _pptRunning);
            status["running"] = runningByApp;

            JObject limits = new JObject();
            limits["excelMaxConcurrency"] = _excelMaxConcurrency;
            limits["wordMaxConcurrency"] = _wordMaxConcurrency;
            limits["pptMaxConcurrency"] = _pptMaxConcurrency;
            limits["itemTimeoutSec"] = _itemTimeoutSec;
            limits["maxQueue"] = _maxQueue;
            status["limits"] = limits;

            status["fileLocks"] = _fileLocks.DescribeLocks();
            status["fileLockCount"] = _fileLocks.Count;

            JArray inFlight = new JArray();
            List<ParallelJobContext> jobs = new List<ParallelJobContext>(_inFlight.Values);
            jobs.Sort(delegate(ParallelJobContext left, ParallelJobContext right)
            {
                return left.Id.CompareTo(right.Id);
            });
            foreach (ParallelJobContext job in jobs)
            {
                JObject item = new JObject();
                item["jobId"] = job.Id;
                item["appType"] = job.AppType;
                item["tool"] = job.Item.ToolName;
                item["file"] = job.FilePath;
                item["stage"] = job.Stage;
                item["elapsedMs"] = (long)job.Clock.Elapsed.TotalMilliseconds;
                item["savePolicy"] = job.SavePolicy;
                item["timeoutSec"] = job.TimeoutSec;
                inFlight.Add(item);
            }
            status["inFlight"] = inFlight;

            JArray recentErrors = new JArray();
            foreach (JObject error in _recentErrors.ToArray())
            {
                recentErrors.Add(error.DeepClone());
            }
            status["recentErrors"] = recentErrors;

            status["workers"] = _workerPool.DescribeWorkers();
            status["workerCount"] = _workerPool.Count;
            status["workerEvicted"] = _workerPool.EvictedCount;
            status["comThreadsAlive"] = _workerPool.Count;
            status["toolInvokerReady"] = ToolInvoker != null;
            status["toolValidatorReady"] = ToolExists != null;
            return status;
        }

        /// <summary>
        /// 程序退出前收口：拒绝新作业，等待在飞作业与仍存活的 STA 线程（带超时）。
        /// 超时未退出时如实写 stderr，不假装清理完成。
        /// </summary>
        public void Shutdown()
        {
            lock (_queueSync)
            {
                if (_shuttingDown)
                {
                    return;
                }
                _shuttingDown = true;
            }

            int waitSeconds = Math.Min(Math.Max(_itemTimeoutSec, 5), 60);
            Log("开始收口：等待在飞作业退出，最多 " +
                waitSeconds.ToString(CultureInfo.InvariantCulture) + " 秒。");

            try
            {
                _workerPool.Shutdown();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[parallel] 收口 worker 池失败：" + ex.Message);
            }

            Stopwatch stopwatch = Stopwatch.StartNew();
            TimeSpan limit = TimeSpan.FromSeconds(waitSeconds);
            while (stopwatch.Elapsed < limit)
            {
                if (Volatile.Read(ref _pendingCount) <= 0)
                {
                    break;
                }
                Thread.Sleep(100);
            }

            int remainingJobs = Volatile.Read(ref _pendingCount);
            if (remainingJobs <= 0)
            {
                Log("收口完成：无在飞作业；worker 池已请求退出。");
                return;
            }

            Console.Error.WriteLine(
                "[parallel] 收口超时：仍有 " + remainingJobs.ToString(CultureInfo.InvariantCulture) +
                " 个作业未完成。进程退出会终止后台 worker，相关 Office 实例可能残留；" +
                "未能在超时内确认零残留，未做虚假清理声明。");
        }

        // ------------------------------------------------------------------
        // 辅助
        // ------------------------------------------------------------------

        private SemaphoreSlim GetGate(string appType)
        {
            if (string.Equals(appType, ParallelToolCatalog.Excel, StringComparison.Ordinal))
            {
                return _excelGate;
            }
            if (string.Equals(appType, ParallelToolCatalog.Ppt, StringComparison.Ordinal))
            {
                return _pptGate;
            }
            return _wordGate;
        }

        private static Task<OperationResult> Reject(OfficeWorkItem item, OfficeErrorCode code, string message)
        {
            OperationResult failure = OperationResult.Failed(code, message);
            item.Completion.TrySetResult(failure);
            return item.Completion.Task;
        }

        private OperationResult BuildTimeout(ParallelJobContext job, string stage, bool threadExited)
        {
            StringBuilder message = new StringBuilder();
            message.Append("并行作业超时：工具「");
            message.Append(job.Item.ToolName);
            message.Append("」处理文件「");
            message.Append(job.FilePath);
            message.Append("」超过 ");
            message.Append(job.TimeoutSec.ToString(CultureInfo.InvariantCulture));
            message.Append(" 秒仍未完成（阶段：");
            message.Append(stage);
            message.Append("）。");

            if (!threadExited)
            {
                message.Append("作业可能仍在常驻 STA worker 中运行；" +
                               "文件锁与应用并发槽已释放，请避免立即修改同一文件。");
            }
            else
            {
                message.Append("文件锁与应用并发槽已释放。");
            }

            OperationResult result = OperationResult.Failed(
                OfficeErrorCode.INTERNAL_ERROR,
                message.ToString(),
                "池内作业不会自动重放；若工具已部分写入，关闭实例会丢弃未保存改动。");
            JObject data = new JObject();
            data["parallel"] = BuildMetadata(job, threadExited, "timeout", null);
            result.Data = data;
            return result;
        }

        private OperationResult AttachMetadata(
            ParallelJobContext job,
            OperationResult result,
            ParallelJobOutcome outcome,
            bool threadExited,
            string status)
        {
            if (result == null)
            {
                result = OperationResult.Failed(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "并行作业未产生结果。");
            }

            JObject data = result.Data == null ? new JObject() : (JObject)result.Data.DeepClone();
            data["parallel"] = BuildMetadata(job, threadExited, status, outcome);
            result.Data = data;
            return result;
        }

        private JObject BuildMetadata(
            ParallelJobContext job,
            bool threadExited,
            string status,
            ParallelJobOutcome outcome)
        {
            DateTime finishedUtc = DateTime.UtcNow;
            JObject metadata = new JObject();
            metadata["jobId"] = job.Id;
            metadata["appType"] = job.AppType;
            metadata["tool"] = job.Item.ToolName;
            metadata["file"] = job.FilePath;
            metadata["savePolicy"] = job.SavePolicy;
            metadata["timeoutSec"] = job.TimeoutSec;
            metadata["status"] = status;
            metadata["affinityKey"] = job.AffinityKey ?? string.Empty;
            metadata["priority"] = job.Item.Priority;
            metadata["idempotencyKey"] = job.IdempotencyKey ?? string.Empty;
            metadata["submittedAtUtc"] = job.SubmittedAtUtc.ToString("o", CultureInfo.InvariantCulture);
            metadata["finishedAtUtc"] = finishedUtc.ToString("o", CultureInfo.InvariantCulture);

            if (job.ExecutionStartedUtc.HasValue)
            {
                DateTime startedUtc = job.ExecutionStartedUtc.Value;
                metadata["executionStartedAtUtc"] = startedUtc.ToString("o", CultureInfo.InvariantCulture);
                metadata["queueWaitMs"] = (long)(startedUtc - job.SubmittedAtUtc).TotalMilliseconds;
                metadata["durationMs"] = (long)(finishedUtc - startedUtc).TotalMilliseconds;
            }
            else
            {
                metadata["executionStartedAtUtc"] = JValue.CreateNull();
                metadata["queueWaitMs"] = 0;
                metadata["durationMs"] = 0;
            }

            metadata["threadExited"] = threadExited;
            metadata["fileLockReleased"] = true;
            metadata["appSlotReleased"] = true;
            if (job.FileVersionBefore != null)
            {
                FileIntegritySnapshot after = FileIntegrityVerifier.Capture(
                    job.FilePath,
                    AppConfig.Current.WorkerSaveVerifyMaxHashMb);
                bool fileChanged;
                string fileNote;
                bool fileVerified = FileIntegrityVerifier.Verify(
                    job.FileVersionBefore,
                    after,
                    out fileChanged,
                    out fileNote);
                metadata["fileVersionVerified"] = fileVerified;
                metadata["fileChangedOnDisk"] = fileChanged;
                metadata["fileVersionNote"] = fileNote;
                metadata["fileBeforeHash"] = job.FileVersionBefore.Sha256 ?? string.Empty;
                metadata["fileAfterHash"] = after == null ? string.Empty : (after.Sha256 ?? string.Empty);
                metadata["fileBeforeLength"] = job.FileVersionBefore.Length;
                metadata["fileAfterLength"] = after == null ? 0 : after.Length;
            }
            else
            {
                metadata["fileVersionVerified"] = false;
                metadata["fileChangedOnDisk"] = false;
                metadata["fileVersionNote"] = "未获取保存前文件指纹（作业未进入文件锁阶段）。";
                metadata["fileBeforeHash"] = string.Empty;
                metadata["fileAfterHash"] = string.Empty;
                metadata["fileBeforeLength"] = 0;
                metadata["fileAfterLength"] = 0;
            }

            if (outcome != null)
            {
                metadata["openedReadOnly"] = outcome.OpenedReadOnly;
                metadata["saved"] = outcome.Saved;
                metadata["saveSkippedReason"] = outcome.SaveSkippedReason ?? string.Empty;
                metadata["undoNodesRemoved"] = outcome.UndoNodesRemoved;
                metadata["saveVerified"] = outcome.SaveVerified;
                metadata["saveVerificationNote"] = outcome.SaveVerificationNote ?? string.Empty;
                metadata["saveChangedOnDisk"] = outcome.SaveChangedOnDisk;
                metadata["saveBeforeHash"] = outcome.SaveBeforeHash ?? string.Empty;
                metadata["saveAfterHash"] = outcome.SaveAfterHash ?? string.Empty;
                metadata["saveBeforeLength"] = outcome.SaveBeforeLength;
                metadata["saveAfterLength"] = outcome.SaveAfterLength;
                metadata["saveBeforeLastWriteUtc"] = outcome.SaveBeforeLastWriteUtc ?? string.Empty;
                metadata["saveAfterLastWriteUtc"] = outcome.SaveAfterLastWriteUtc ?? string.Empty;
            }
            else
            {
                metadata["openedReadOnly"] = false;
                metadata["saved"] = false;
                metadata["saveSkippedReason"] = string.Empty;
                metadata["undoNodesRemoved"] = 0;
                metadata["saveVerified"] = false;
                metadata["saveVerificationNote"] = string.Empty;
                metadata["saveChangedOnDisk"] = false;
                metadata["saveBeforeHash"] = string.Empty;
                metadata["saveAfterHash"] = string.Empty;
                metadata["saveBeforeLength"] = 0;
                metadata["saveAfterLength"] = 0;
                metadata["saveBeforeLastWriteUtc"] = string.Empty;
                metadata["saveAfterLastWriteUtc"] = string.Empty;
            }
            return metadata;
        }

        private static string NormalizeIdempotencyKey(
            string userId,
            string appType,
            string toolName,
            string rawKey)
        {
            if (string.IsNullOrWhiteSpace(rawKey))
            {
                return string.Empty;
            }
            string user = string.IsNullOrWhiteSpace(userId) ? "default" : userId.Trim();
            return user + "|" + (appType ?? string.Empty) + "|" + (toolName ?? string.Empty) + "|" + rawKey.Trim();
        }

        private void CleanupIdempotencyCache()
        {
            if (_idempotencyTtlSec <= 0)
            {
                foreach (KeyValuePair<string, IdempotencyCacheEntry> pair in _idempotencyCompleted)
                {
                    IdempotencyCacheEntry ignored;
                    _idempotencyCompleted.TryRemove(pair.Key, out ignored);
                }
                return;
            }

            DateTime cutoff = DateTime.UtcNow.AddSeconds(-_idempotencyTtlSec);
            foreach (KeyValuePair<string, IdempotencyCacheEntry> pair in _idempotencyCompleted)
            {
                if (pair.Value == null || pair.Value.CompletedUtc < cutoff)
                {
                    IdempotencyCacheEntry ignored;
                    _idempotencyCompleted.TryRemove(pair.Key, out ignored);
                }
            }
        }

        private sealed class IdempotencyCacheEntry
        {
            public readonly OperationResult Result;
            public readonly DateTime CompletedUtc;

            public IdempotencyCacheEntry(OperationResult result, DateTime completedUtc)
            {
                Result = result;
                CompletedUtc = completedUtc;
            }
        }

        private static void Log(string message)
        {
            Console.Error.WriteLine("[parallel] " + message);
        }

        /// <summary>单个作业的内部上下文；只由调度器读写。</summary>
        private sealed class ParallelJobContext
        {
            public readonly long Id;
            public readonly OfficeWorkItem Item;
            public readonly string AppType;
            public readonly string FilePath;
            public readonly string LockKey;
            public readonly string SavePolicy;
            public readonly int TimeoutSec;
            public readonly string UserId;
            public readonly DateTime SubmittedAtUtc;
            public readonly Stopwatch Clock;
            public readonly ManualResetEventSlim Done;

            public ParallelJobOutcome Outcome;
            public DateTime? ExecutionStartedUtc;
            public volatile string Stage;
            public string AffinityKey;
            public string IdempotencyKey;
            public FileIntegritySnapshot FileVersionBefore;

            public ParallelJobContext(
                long id,
                OfficeWorkItem item,
                string appType,
                string filePath,
                string lockKey,
                string savePolicy,
                int timeoutSec,
                string userId)
            {
                Id = id;
                Item = item;
                AppType = appType;
                FilePath = filePath;
                LockKey = lockKey;
                SavePolicy = savePolicy;
                TimeoutSec = timeoutSec;
                UserId = string.IsNullOrWhiteSpace(userId) ? "default" : userId;
                SubmittedAtUtc = DateTime.UtcNow;
                Clock = Stopwatch.StartNew();
                Done = new ManualResetEventSlim(false);
                Stage = "queued";
                ExecutionStartedUtc = null;
                AffinityKey = string.Empty;
                IdempotencyKey = string.Empty;
                FileVersionBefore = null;
            }

            public TimeSpan Remaining
            {
                get
                {
                    TimeSpan left = TimeSpan.FromSeconds(TimeoutSec) - Clock.Elapsed;
                    return left > TimeSpan.Zero ? left : TimeSpan.Zero;
                }
            }
        }
    }
}
