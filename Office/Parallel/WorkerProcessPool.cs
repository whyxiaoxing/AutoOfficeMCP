using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Office.Config;
using Office.Contracts;

namespace Office.Parallel
{
    /// <summary>
    /// T11：按 userId + app 管理独立 worker 进程。
    /// Gateway 只通过命名管道投递单作业，不执行 COM；worker 进程内一个 STA 线程 + 一个 Application。
    /// </summary>
    internal sealed class WorkerProcessPool
    {
        private readonly object _sync;
        private readonly ConcurrentDictionary<string, WorkerProcess> _workers;
        private readonly int _maxWarmWorkers;
        private readonly int _maxWorkersPerUser;
        private readonly int _idleTimeoutSec;
        private readonly int _startupTimeoutSec;
        private readonly int _jobTimeoutSec;
        private readonly int _memoryLimitMb;
        private readonly int _heartbeatTimeoutSec;
        private readonly string _pipePrefix;
        private readonly string _exePath;
        private readonly ConcurrentDictionary<string, int> _restartCounts;
        private readonly Timer _idleTimer;
        private int _stopping;
        private long _evicted;

        public WorkerProcessPool()
        {
            _sync = new object();
            _workers = new ConcurrentDictionary<string, WorkerProcess>(StringComparer.OrdinalIgnoreCase);
            AppConfig config = AppConfig.Current;
            _maxWarmWorkers = Math.Max(1, config.WorkerMaxWarmWorkers);
            _maxWorkersPerUser = Math.Max(1, config.WorkerMaxWorkersPerUser);
            _idleTimeoutSec = Math.Max(30, config.WorkerIdleTimeoutSec);
            _startupTimeoutSec = Math.Max(5, config.WorkerStartupTimeoutSec);
            _jobTimeoutSec = Math.Max(5, config.WorkerJobTimeoutSec);
            _memoryLimitMb = Math.Max(0, config.WorkerMemoryLimitMb);
            _heartbeatTimeoutSec = Math.Max(6, config.WorkerHeartbeatTimeoutSec);
            _restartCounts = new ConcurrentDictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            _pipePrefix = string.IsNullOrWhiteSpace(config.WorkerPipeNamePrefix)
                ? "AutoOfficeMCP.Worker"
                : config.WorkerPipeNamePrefix.Trim();
            _exePath = Process.GetCurrentProcess().MainModule.FileName;
            _idleTimer = new Timer(EvictIdleWorkers, null, 30000, 30000);
        }

        public int Count
        {
            get { return _workers.Count; }
        }

        public long EvictedCount
        {
            get { return Interlocked.Read(ref _evicted); }
        }

        public async Task<ParallelJobOutcome> EnqueueAsync(string userId, string appType, OfficeWorkItem item)
        {
            if (item == null)
            {
                return Failure(OfficeErrorCode.INVALID_PARAMS, "并行作业为空。");
            }
            if (Volatile.Read(ref _stopping) != 0)
            {
                return Failure(OfficeErrorCode.OFFICE_BUSY, "worker 进程池正在关闭，作业未执行。");
            }

            string normalizedApp = ParallelToolCatalog.NormalizeApp(appType);
            if (normalizedApp.Length == 0)
            {
                return Failure(OfficeErrorCode.INVALID_PARAMS, "appType 必须是 excel、word 或 ppt。");
            }

            string normalizedUser = string.IsNullOrWhiteSpace(userId) ? "default" : userId.Trim();
            string key = normalizedUser + "|" + normalizedApp;
            WorkerProcess worker;
            lock (_sync)
            {
                if (!_workers.TryGetValue(key, out worker) || worker.IsDead || worker.IsHeartbeatStale)
                {
                    if (worker != null)
                    {
                        WorkerProcess removed;
                        _workers.TryRemove(key, out removed);
                        if (worker.IsDead || worker.IsHeartbeatStale)
                        {
                            CountRestart(key);
                        }
                        StopReplacedWorker(worker);
                    }
                    if (_workers.Count >= _maxWarmWorkers)
                    {
                        WorkerProcess evicted = FindIdleWorker();
                        if (evicted == null)
                        {
                            return Failure(
                                OfficeErrorCode.OFFICE_BUSY,
                                "worker 进程池已满（maxWarmWorkers=" + _maxWarmWorkers +
                                "），当前没有空闲 worker 可回收；请稍后重试。");
                        }
                        RemoveAndStop(evicted);
                    }
                    if (CountWorkersForUser(normalizedUser, normalizedApp) >= _maxWorkersPerUser)
                    {
                        return Failure(
                            OfficeErrorCode.OFFICE_BUSY,
                            "用户「" + normalizedUser + "」的 " + normalizedApp +
                            " worker 数已达到 maxWorkersPerUser=" + _maxWorkersPerUser + "。");
                    }

                    worker = new WorkerProcess(
                        key,
                        normalizedUser,
                        normalizedApp,
                        _pipePrefix,
                        _exePath,
                        _startupTimeoutSec,
                        _jobTimeoutSec,
                        _memoryLimitMb,
                        _heartbeatTimeoutSec);
                    worker.SetRestartCount(GetRestartCount(key));
                    _workers[key] = worker;
                    worker.Start();
                }
            }

            try
            {
                await worker.EnsureStartedAsync().ConfigureAwait(false);
                return await worker.EnqueueAsync(item).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                WorkerProcess removed;
                _workers.TryRemove(key, out removed);
                worker.Kill();
                return Failure(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "worker 启动或执行失败（工具 " + item.ToolName + "）：" + ex.Message);
            }
        }

        private void CountRestart(string key)
        {
            _restartCounts.AddOrUpdate(
                key,
                1,
                delegate(string ignored, int current) { return current + 1; });
        }

        private int GetRestartCount(string key)
        {
            int count;
            return _restartCounts.TryGetValue(key, out count) ? count : 0;
        }

        /// <summary>替换死掉或心跳超时的 worker；收口放到后台，避免阻塞请求线程。</summary>
        private void StopReplacedWorker(WorkerProcess worker)
        {
            if (worker == null)
            {
                return;
            }
            Task.Run(delegate
            {
                try
                {
                    worker.Shutdown();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[worker-process] 收口被替换 worker 失败：" + worker.Key + "；" + ex.Message);
                }
            });
        }

        private int CountWorkersForUser(string userId, string appType)
        {
            int count = 0;
            foreach (KeyValuePair<string, WorkerProcess> pair in _workers)
            {
                if (pair.Value != null
                    && string.Equals(pair.Value.UserId, userId, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(pair.Value.AppType, appType, StringComparison.OrdinalIgnoreCase))
                {
                    count++;
                }
            }
            return count;
        }

        private WorkerProcess FindIdleWorker()
        {
            WorkerProcess oldest = null;
            DateTime oldestTime = DateTime.MaxValue;
            foreach (KeyValuePair<string, WorkerProcess> pair in _workers)
            {
                WorkerProcess candidate = pair.Value;
                if (candidate == null || !candidate.IsIdle)
                {
                    continue;
                }
                if (candidate.LastUsedUtc < oldestTime)
                {
                    oldestTime = candidate.LastUsedUtc;
                    oldest = candidate;
                }
            }
            return oldest;
        }

        private void RemoveAndStop(WorkerProcess worker)
        {
            if (worker == null)
            {
                return;
            }
            WorkerProcess removed;
            _workers.TryRemove(worker.Key, out removed);
            if (worker.IsDead || worker.IsHeartbeatStale)
            {
                CountRestart(worker.Key);
            }
            Interlocked.Increment(ref _evicted);
            Task.Run(delegate
            {
                try
                {
                    worker.Shutdown();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[worker-process] 回收失败：" + worker.Key + "；" + ex.Message);
                }
            });
        }

        private void EvictIdleWorkers(object state)
        {
            if (Volatile.Read(ref _stopping) != 0)
            {
                return;
            }

            DateTime cutoff = DateTime.UtcNow.AddSeconds(-_idleTimeoutSec);
            List<WorkerProcess> toEvict = new List<WorkerProcess>();
            lock (_sync)
            {
                foreach (KeyValuePair<string, WorkerProcess> pair in _workers)
                {
                    WorkerProcess worker = pair.Value;
                    if (worker != null
                        && ((worker.IsIdle && worker.LastUsedUtc < cutoff) || worker.IsHeartbeatStale))
                    {
                        toEvict.Add(worker);
                    }
                }
                foreach (WorkerProcess worker in toEvict)
                {
                    RemoveAndStop(worker);
                }
            }
        }

        public JArray DescribeWorkers()
        {
            JArray result = new JArray();
            foreach (KeyValuePair<string, WorkerProcess> pair in _workers)
            {
                WorkerProcess worker = pair.Value;
                if (worker != null)
                {
                    result.Add(worker.Describe());
                }
            }
            return result;
        }

        public void Shutdown()
        {
            if (Interlocked.Exchange(ref _stopping, 1) != 0)
            {
                return;
            }
            _idleTimer.Dispose();
            List<WorkerProcess> workers = new List<WorkerProcess>(_workers.Values);
            _workers.Clear();
            foreach (WorkerProcess worker in workers)
            {
                try
                {
                    worker.Shutdown();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[worker-process] 退出回收失败：" + worker.Key + "；" + ex.Message);
                }
            }
        }

        public void Dispose()
        {
            Shutdown();
        }

        private static ParallelJobOutcome Failure(OfficeErrorCode code, string message)
        {
            ParallelJobOutcome outcome = new ParallelJobOutcome();
            outcome.Result = OperationResult.Failed(code, message);
            return outcome;
        }
    }

    /// <summary>一个独立 worker 进程及其命名管道连接。</summary>
    internal sealed class WorkerProcess
    {
        private readonly string _key;
        private readonly string _userId;
        private readonly string _appType;
        private readonly string _pipeName;
        private readonly string _exePath;
        private readonly int _startupTimeoutSec;
        private readonly int _jobTimeoutSec;
        private readonly int _memoryLimitMb;
        private readonly int _heartbeatTimeoutSec;
        private readonly SemaphoreSlim _jobLock;
        private readonly TaskCompletionSource<bool> _ready;
        private readonly object _sync;
        private Task _startTask;
        private Task _readerTask;
        private Process _process;
        private JobObject _job;
        private NamedPipeServerStream _pipe;
        private StreamReader _reader;
        private StreamWriter _writer;
        private TaskCompletionSource<WorkerJobResponse> _pending;
        private int _officePid;
        private string _workerTemp;
        private int _state; // 0=stopped,1=starting,2=idle,3=busy,4=dead
        private int _stopping;
        private string _currentTool;
        private DateTime _lastUsedUtc;
        private long _completed;
        private long _failed;
        private int _restarts;
        private int _queueDepth;
        private long _lastHeartbeatTicks;
        private readonly ConcurrentQueue<string> _recentLogs = new ConcurrentQueue<string>();
        private const int MaxRecentLogs = 20;

        public WorkerProcess(
            string key,
            string userId,
            string appType,
            string pipePrefix,
            string exePath,
            int startupTimeoutSec,
            int jobTimeoutSec,
            int memoryLimitMb,
            int heartbeatTimeoutSec)
        {
            _key = key;
            _userId = userId;
            _appType = appType;
            _pipeName = pipePrefix + "." + Guid.NewGuid().ToString("N");
            _exePath = exePath;
            _startupTimeoutSec = startupTimeoutSec;
            _jobTimeoutSec = jobTimeoutSec;
            _memoryLimitMb = memoryLimitMb;
            _heartbeatTimeoutSec = heartbeatTimeoutSec;
            _jobLock = new SemaphoreSlim(1, 1);
            _ready = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            _sync = new object();
            _lastUsedUtc = DateTime.UtcNow;
            _lastHeartbeatTicks = DateTime.UtcNow.Ticks;
            _state = 0;
        }

        public string Key
        {
            get { return _key; }
        }

        public string UserId
        {
            get { return _userId; }
        }

        public string AppType
        {
            get { return _appType; }
        }

        public DateTime LastUsedUtc
        {
            get { return _lastUsedUtc; }
        }

        public DateTime LastHeartbeatUtc
        {
            get { return new DateTime(Interlocked.Read(ref _lastHeartbeatTicks), DateTimeKind.Utc); }
        }

        public int HeartbeatAgeSec
        {
            get
            {
                TimeSpan age = DateTime.UtcNow - LastHeartbeatUtc;
                return age.TotalSeconds <= 0 ? 0 : (int)age.TotalSeconds;
            }
        }

        public bool IsHeartbeatStale
        {
            get { return _heartbeatTimeoutSec > 0 && HeartbeatAgeSec > _heartbeatTimeoutSec; }
        }

        public void SetRestartCount(int value)
        {
            Volatile.Write(ref _restarts, Math.Max(0, value));
        }

        public bool IsIdle
        {
            get { return Volatile.Read(ref _state) == 2; }
        }

        public bool IsDead
        {
            get { return Volatile.Read(ref _state) == 4; }
        }

        public string State
        {
            get
            {
                switch (Volatile.Read(ref _state))
                {
                    case 0: return "stopped";
                    case 1: return "starting";
                    case 2: return "idle";
                    case 3: return "busy";
                    default: return "dead";
                }
            }
        }

        public void Start()
        {
            lock (_sync)
            {
                if (_startTask == null)
                {
                    _startTask = StartCoreAsync();
                }
            }
        }

        public async Task EnsureStartedAsync()
        {
            Start();
            Task startTask;
            lock (_sync)
            {
                startTask = _startTask;
            }
            if (startTask == null)
            {
                throw new InvalidOperationException("worker 启动任务为空。");
            }
            await startTask.ConfigureAwait(false);
            if (!_ready.Task.IsCompleted || !_ready.Task.Result)
            {
                throw new InvalidOperationException("worker 未就绪。");
            }
        }

        private async Task StartCoreAsync()
        {
            Volatile.Write(ref _state, 1);
            try
            {
                _pipe = new NamedPipeServerStream(
                    _pipeName,
                    PipeDirection.InOut,
                    1,
                    PipeTransmissionMode.Byte,
                    PipeOptions.Asynchronous,
                    0,
                    0,
                    WorkerPipeSecurity.Create());
                Task connectionTask = _pipe.WaitForConnectionAsync();

                string workerTemp = Path.Combine(
                    Path.GetTempPath(),
                    "AutoOfficeMCP",
                    "workers",
                    Math.Abs(_key.GetHashCode()).ToString(CultureInfo.InvariantCulture) + "-" +
                    Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(workerTemp);
                _workerTemp = workerTemp;

                WorkerSettings settings = new WorkerSettings();
                settings.Role = "worker";
                settings.PipeName = _pipeName;
                settings.UserId = _userId;
                settings.AppType = _appType;
                WorkerBootstrap.WriteToDirectory(workerTemp, settings);

                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = _exePath;
                // 子进程从当前目录读取 worker.settings.json；不传递任何环境变量或命令行参数。
                info.WorkingDirectory = workerTemp;
                info.UseShellExecute = false;
                info.CreateNoWindow = true;
                info.RedirectStandardOutput = true;
                info.RedirectStandardError = true;

                _process = new Process();
                _process.StartInfo = info;
                _process.EnableRaisingEvents = true;
                _process.OutputDataReceived += delegate(object sender, DataReceivedEventArgs args)
                {
                    if (!string.IsNullOrWhiteSpace(args.Data))
                    {
                        RememberLog("[out] " + args.Data);
                        Console.Error.WriteLine("[worker:" + _key + ":out] " + args.Data);
                    }
                };
                _process.ErrorDataReceived += delegate(object sender, DataReceivedEventArgs args)
                {
                    if (!string.IsNullOrWhiteSpace(args.Data))
                    {
                        RememberLog(args.Data);
                        Console.Error.WriteLine("[worker:" + _key + "] " + args.Data);
                    }
                };
                _process.Exited += delegate
                {
                    Volatile.Write(ref _state, 4);
                    _ready.TrySetResult(false);
                    // 先唤醒等待方并携带失败原因，再清理 Office 进程，避免调用方长时间无响应。
                    TaskCompletionSource<WorkerJobResponse> pending = _pending;
                    if (pending != null)
                    {
                        pending.TrySetResult(FailureResponse(DescribeFailure(
                            "worker 进程已退出（退出码 " +
                            SafeExitCode().ToString(CultureInfo.InvariantCulture) + "）")));
                    }
                    KillOfficeProcess();
                    // worker 被外部强杀时不会走 Shutdown；异步清理 per-worker 临时目录。
                    Task.Run(delegate
                    {
                        try
                        {
                            System.Threading.Thread.Sleep(2000);
                            CleanupTempDirectory();
                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine("[worker-process] 退出后清理临时目录失败：" + _key +
                                                    "；" + ex.Message);
                        }
                    });
                };
                if (!_process.Start())
                {
                    throw new InvalidOperationException("无法启动 worker 进程：" + _exePath);
                }
                _process.BeginOutputReadLine();
                _process.BeginErrorReadLine();

                _job = new JobObject(_memoryLimitMb);
                _job.Assign(_process);

                Task connected = await Task.WhenAny(
                    connectionTask,
                    Task.Delay(TimeSpan.FromSeconds(_startupTimeoutSec))).ConfigureAwait(false);
                if (connected != connectionTask)
                {
                    throw new TimeoutException("worker 命名管道连接超时（" + _startupTimeoutSec + " 秒）。");
                }

                _reader = new StreamReader(_pipe, new UTF8Encoding(false), false, 4096, true);
                _writer = new StreamWriter(_pipe, new UTF8Encoding(false), 4096, true);
                _writer.AutoFlush = true;
                _writer.NewLine = "\n";
                _readerTask = Task.Run((Func<Task>)ReadLoopAsync);

                Task<bool> readyTask = _ready.Task;
                Task readyTimeout = await Task.WhenAny(
                    readyTask,
                    Task.Delay(TimeSpan.FromSeconds(_startupTimeoutSec))).ConfigureAwait(false);
                if (readyTimeout != readyTask || !readyTask.Result)
                {
                    throw new TimeoutException("worker ready 握手超时（" + _startupTimeoutSec + " 秒）。");
                }
                Interlocked.Exchange(ref _lastHeartbeatTicks, DateTime.UtcNow.Ticks);
                Volatile.Write(ref _state, 2);
            }
            catch (Exception)
            {
                Volatile.Write(ref _state, 4);
                Kill();
                throw;
            }
        }

        public async Task<ParallelJobOutcome> EnqueueAsync(OfficeWorkItem item)
        {
            await EnsureStartedAsync().ConfigureAwait(false);
            int timeoutSec = item.TimeoutSec > 0 ? item.TimeoutSec : _jobTimeoutSec;
            bool acquired = await _jobLock.WaitAsync(TimeSpan.FromSeconds(timeoutSec)).ConfigureAwait(false);
            if (!acquired)
            {
                return Failure(
                    OfficeErrorCode.OFFICE_BUSY,
                    "worker 任务队列等待超时（工具 " + item.ToolName + "）。");
            }

            try
            {
                if (IsDead)
                {
                    return Failure(
                        OfficeErrorCode.OFFICE_BUSY,
                        DescribeFailure("worker 进程已退出，作业未执行（工具 " + item.ToolName + "）"));
                }
                if (IsHeartbeatStale)
                {
                    return Failure(
                        OfficeErrorCode.OFFICE_BUSY,
                        "worker 心跳超时（" + HeartbeatAgeSec + " 秒 > " + _heartbeatTimeoutSec +
                        " 秒），工具「" + item.ToolName + "」未执行；下一次请求会重建 worker。");
                }

                WorkerJobRequest request = new WorkerJobRequest();
                request.Type = "job";
                request.JobId = Guid.NewGuid().ToString("N");
                request.UserId = _userId;
                request.App = _appType;
                request.Tool = item.ToolName;
                request.File = item.FilePath;
                request.Arguments = item.Args == null ? new JObject() : (JObject)item.Args.DeepClone();
                request.SavePolicy = string.IsNullOrWhiteSpace(item.SavePolicy) ? "none" : item.SavePolicy;
                request.TimeoutSec = timeoutSec;
                request.AffinityKey = item.AffinityKey ?? string.Empty;
                request.Priority = item.Priority;
                request.IdempotencyKey = item.IdempotencyKey ?? string.Empty;

                TaskCompletionSource<WorkerJobResponse> responseSource =
                    new TaskCompletionSource<WorkerJobResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
                _pending = responseSource;
                _currentTool = item.ToolName;
                _lastUsedUtc = DateTime.UtcNow;
                Volatile.Write(ref _state, 3);
                Interlocked.Increment(ref _queueDepth);

                string line = WorkerProtocol.Serialize(request);
                await _writer.WriteLineAsync(line).ConfigureAwait(false);
                await _writer.FlushAsync().ConfigureAwait(false);
                Interlocked.Decrement(ref _queueDepth);

                Task<WorkerJobResponse> responseTask = responseSource.Task;
                Task completed = await Task.WhenAny(
                    responseTask,
                    Task.Delay(TimeSpan.FromSeconds(timeoutSec))).ConfigureAwait(false);
                if (completed != responseTask)
                {
                    Kill();
                    Interlocked.Increment(ref _failed);
                    return Failure(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "worker 作业超时（" + timeoutSec + " 秒），已终止进程树；工具：" + item.ToolName);
                }

                WorkerJobResponse response = await responseTask.ConfigureAwait(false);
                if (response == null)
                {
                    Interlocked.Increment(ref _failed);
                    return Failure(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "worker 返回空结果（工具 " + item.ToolName + "）。");
                }
                if (response.Success)
                {
                    Interlocked.Increment(ref _completed);
                }
                else
                {
                    Interlocked.Increment(ref _failed);
                }
                return BuildOutcome(response);
            }
            catch (Exception ex)
            {
                Interlocked.Increment(ref _failed);
                return Failure(
                    OfficeErrorCode.INTERNAL_ERROR,
                    DescribeFailure("worker 通信失败（工具 " + item.ToolName + "）：" + ex.Message));
            }
            finally
            {
                _pending = null;
                _currentTool = string.Empty;
                if (!IsDead)
                {
                    Volatile.Write(ref _state, 2);
                }
                _lastUsedUtc = DateTime.UtcNow;
                _jobLock.Release();
            }
        }

        private async Task ReadLoopAsync()
        {
            try
            {
                while (true)
                {
                    string line = await _reader.ReadLineAsync().ConfigureAwait(false);
                    if (line == null)
                    {
                        break;
                    }
                    string type = WorkerProtocol.GetMessageType(line);
                    if (string.Equals(type, "ready", StringComparison.OrdinalIgnoreCase))
                    {
                        WorkerControlMessage control = WorkerProtocol.DeserializeControl(line);
                        if (control != null)
                        {
                            _officePid = control.OfficePid;
                        }
                        _ready.TrySetResult(true);
                        continue;
                    }
                    if (string.Equals(type, "result", StringComparison.OrdinalIgnoreCase))
                    {
                        WorkerJobResponse response = WorkerProtocol.DeserializeResponse(line);
                        TaskCompletionSource<WorkerJobResponse> pending = _pending;
                        if (pending != null)
                        {
                            pending.TrySetResult(response);
                        }
                        continue;
                    }
                    if (string.Equals(type, "heartbeat", StringComparison.OrdinalIgnoreCase))
                    {
                        // 心跳只刷新存活时间，不刷新 LastUsedUtc；否则空闲回收永远无法触发。
                        Interlocked.Exchange(ref _lastHeartbeatTicks, DateTime.UtcNow.Ticks);
                        continue;
                    }
                    if (string.Equals(type, "error", StringComparison.OrdinalIgnoreCase))
                    {
                        WorkerControlMessage control = WorkerProtocol.DeserializeControl(line);
                        TaskCompletionSource<WorkerJobResponse> pending = _pending;
                        if (pending != null)
                        {
                            pending.TrySetResult(FailureResponse(DescribeFailure(
                                control == null ? "worker 错误。" : control.Message)));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!IsDead)
                {
                    Console.Error.WriteLine("[worker-process] 读管道失败：" + _key + "；" + ex.Message);
                }
            }
            finally
            {
                Volatile.Write(ref _state, 4);
                _ready.TrySetResult(false);
                TaskCompletionSource<WorkerJobResponse> pending = _pending;
                if (pending != null)
                {
                    pending.TrySetResult(FailureResponse(DescribeFailure("worker 管道已关闭")));
                }
            }
        }

        /// <summary>保留最近 worker 日志，供进程退出/管道断开时向调用方说明原因。</summary>
        private void RememberLog(string line)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                return;
            }
            _recentLogs.Enqueue(line.Trim());
            while (_recentLogs.Count > MaxRecentLogs)
            {
                string discarded;
                _recentLogs.TryDequeue(out discarded);
            }
        }

        private string RecentLogText()
        {
            if (_recentLogs.IsEmpty)
            {
                return string.Empty;
            }
            StringBuilder builder = new StringBuilder();
            foreach (string line in _recentLogs.ToArray())
            {
                if (builder.Length > 0)
                {
                    builder.Append(" | ");
                }
                builder.Append(line);
            }
            return builder.ToString();
        }

        /// <summary>
        /// 在失败消息后补当前工具与最近 worker 日志；调用方据此区分“worker 被关闭”
        /// 与“工具本身失败”，而不是只看到一句无上下文的“进程已退出”。
        /// </summary>
        private string DescribeFailure(string message)
        {
            StringBuilder builder = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(message))
            {
                builder.Append(message.Trim().TrimEnd('。'));
                builder.Append('。');
            }
            string tool = _currentTool;
            if (!string.IsNullOrWhiteSpace(tool) &&
                (message == null || message.IndexOf(tool, StringComparison.Ordinal) < 0))
            {
                builder.Append("当前工具：").Append(tool).Append('。');
            }
            string logs = RecentLogText();
            if (logs.Length > 0)
            {
                builder.Append("最近 worker 日志：").Append(logs);
            }
            return builder.ToString();
        }

        private int SafeExitCode()
        {
            try
            {
                return _process == null ? -1 : _process.ExitCode;
            }
            catch (Exception)
            {
                return -1;
            }
        }

        public JObject Describe()
        {
            JObject item = new JObject();
            item["key"] = _key;
            item["userId"] = _userId;
            item["appType"] = _appType;
            item["pid"] = _process == null ? 0 : SafePid();
            item["officePid"] = Volatile.Read(ref _officePid);
            item["state"] = State;
            item["currentTool"] = _currentTool ?? string.Empty;
            item["queueDepth"] = Volatile.Read(ref _queueDepth);
            item["completed"] = Interlocked.Read(ref _completed);
            item["failed"] = Interlocked.Read(ref _failed);
            item["restarts"] = Volatile.Read(ref _restarts);
            item["lastUsedUtc"] = _lastUsedUtc.ToString("o");
            item["lastHeartbeatUtc"] = LastHeartbeatUtc.ToString("o");
            item["heartbeatAgeSec"] = HeartbeatAgeSec;
            item["heartbeatTimeoutSec"] = _heartbeatTimeoutSec;
            item["heartbeatStale"] = IsHeartbeatStale;
            item["pipe"] = _pipeName;
            return item;
        }

        private int SafePid()
        {
            try
            {
                return _process.Id;
            }
            catch
            {
                return 0;
            }
        }

        public void Shutdown()
        {
            if (Interlocked.Exchange(ref _stopping, 1) != 0)
            {
                return;
            }
            try
            {
                if (_writer != null && !IsDead)
                {
                    WorkerControlMessage shutdown = new WorkerControlMessage();
                    shutdown.Type = "shutdown";
                    _writer.WriteLine(WorkerProtocol.Serialize(shutdown));
                    _writer.Flush();
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[worker-process] 发送 shutdown 失败：" + _key + "；" + ex.Message);
            }

            if (_process != null)
            {
                try
                {
                    if (!_process.WaitForExit(Math.Min(10000, Math.Max(2000, _startupTimeoutSec * 1000))))
                    {
                        Kill();
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[worker-process] 等待 worker 退出失败：" + _key + "；" + ex.Message);
                }
            }
            KillOfficeProcess();
            DisposeResources();
        }

        public void Kill()
        {
            try
            {
                if (_job != null)
                {
                    _job.Terminate(1);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[worker-process] TerminateJobObject 失败：" + _key + "；" + ex.Message);
            }
            try
            {
                if (_process != null && !_process.HasExited)
                {
                    _process.Kill();
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[worker-process] Kill worker 失败：" + _key + "；" + ex.Message);
            }
            Volatile.Write(ref _state, 4);
            KillOfficeProcess();
            DisposeResources();
        }

        private void KillOfficeProcess()
        {
            int pid = Volatile.Read(ref _officePid);
            if (pid <= 0)
            {
                return;
            }
            try
            {
                Process office = Process.GetProcessById(pid);
                string processName = office.ProcessName;
                if (!string.Equals(processName, "EXCEL", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(processName, "WINWORD", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(processName, "POWERPNT", StringComparison.OrdinalIgnoreCase))
                {
                    Console.Error.WriteLine("[worker-process] officePid=" + pid +
                                            " 已不是 EXCEL/WINWORD/POWERPNT（" + processName + "），跳过强杀。");
                    office.Dispose();
                    return;
                }
                if (!office.HasExited)
                {
                    office.Kill();
                    office.WaitForExit(5000);
                }
                office.Dispose();
                Console.Error.WriteLine("[worker-process] 已强杀 Office 进程：" + _key + " pid=" + pid);
            }
            catch (ArgumentException)
            {
                // 进程已经退出。
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[worker-process] 强杀 Office 进程失败：" + _key +
                                        " pid=" + pid + "；" + ex.Message);
            }
            finally
            {
                Volatile.Write(ref _officePid, 0);
            }
        }

        private void DisposeResources()
        {
            try
            {
                if (_readerTask != null)
                {
                    _readerTask.Wait(1000);
                    _readerTask = null;
                }
            }
            catch (Exception)
            {
            }
            try
            {
                if (_reader != null) { _reader.Dispose(); _reader = null; }
            }
            catch { }
            try
            {
                if (_writer != null) { _writer.Dispose(); _writer = null; }
            }
            catch { }
            try
            {
                if (_pipe != null) { _pipe.Dispose(); _pipe = null; }
            }
            catch { }
            try
            {
                if (_job != null) { _job.Dispose(); _job = null; }
            }
            catch { }
            try
            {
                if (_process != null) { _process.Dispose(); _process = null; }
            }
            catch { }
            CleanupTempDirectory();
            // _jobLock 由 EnqueueAsync 的 finally 释放；这里不 Dispose，避免超时 Kill 后 Release 抛 ObjectDisposedException。
        }

        internal void CleanupTempDirectory()
        {
            string temp = _workerTemp;
            if (string.IsNullOrWhiteSpace(temp))
            {
                return;
            }
            try
            {
                if (Directory.Exists(temp))
                {
                    Directory.Delete(temp, true);
                }
                _workerTemp = null;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[worker-process] 清理临时目录失败：" + temp + "；" + ex.Message);
            }
        }

        private static WorkerJobResponse FailureResponse(string message)
        {
            WorkerJobResponse response = new WorkerJobResponse();
            response.Success = false;
            response.ErrorCode = "INTERNAL_ERROR";
            response.Message = message ?? "worker 失败。";
            response.Recoverable = false;
            return response;
        }

        private static ParallelJobOutcome Failure(OfficeErrorCode code, string message)
        {
            ParallelJobOutcome outcome = new ParallelJobOutcome();
            outcome.Result = OperationResult.Failed(code, message);
            return outcome;
        }

        private static ParallelJobOutcome BuildOutcome(WorkerJobResponse response)
        {
            Observation observation = RebuildObservation(response.Observation);
            OperationResult result;
            if (response.Success)
            {
                result = OperationResult.Succeed(
                    response.Message,
                    response.Data,
                    observation,
                    response.RollbackNote);
            }
            else
            {
                result = OperationResult.Failed(
                    string.IsNullOrWhiteSpace(response.ErrorCode) ? "INTERNAL_ERROR" : response.ErrorCode,
                    response.Message,
                    response.Recoverable,
                    response.RollbackNote);
                result.Data = response.Data;
                result.Observation = observation;
            }
            result.Diagnostic = response.Diagnostic ?? string.Empty;

            ParallelJobOutcome outcome = new ParallelJobOutcome();
            outcome.Result = result;
            outcome.Saved = response.Saved;
            outcome.OpenedReadOnly = response.OpenedReadOnly;
            outcome.SaveSkippedReason = response.SaveSkippedReason;
            outcome.UndoNodesRemoved = response.UndoNodesRemoved;
            outcome.SaveVerified = response.SaveVerified;
            outcome.SaveVerificationNote = response.SaveVerificationNote ?? string.Empty;
            outcome.SaveChangedOnDisk = response.SaveChangedOnDisk;
            outcome.SaveBeforeHash = response.SaveBeforeHash ?? string.Empty;
            outcome.SaveAfterHash = response.SaveAfterHash ?? string.Empty;
            outcome.SaveBeforeLength = response.SaveBeforeLength;
            outcome.SaveAfterLength = response.SaveAfterLength;
            outcome.SaveBeforeLastWriteUtc = response.SaveBeforeLastWriteUtc ?? string.Empty;
            outcome.SaveAfterLastWriteUtc = response.SaveAfterLastWriteUtc ?? string.Empty;
            return outcome;
        }

        private static Observation RebuildObservation(JObject json)
        {
            if (json == null)
            {
                return null;
            }
            Observation observation = new Observation();
            observation.Kind = json.Value<string>("kind") ?? string.Empty;
            observation.Summary = json.Value<string>("summary") ?? string.Empty;
            observation.Changed = json.Value<bool?>("changed") ?? false;
            JArray targetRefs = json["targetRefs"] as JArray;
            if (targetRefs != null)
            {
                foreach (JToken token in targetRefs)
                {
                    if (token != null && token.Type == JTokenType.String)
                    {
                        observation.AddTargetRef(token.Value<string>());
                    }
                }
            }
            JArray undoUnits = json["undoUnits"] as JArray;
            if (undoUnits != null)
            {
                foreach (JToken token in undoUnits)
                {
                    if (token != null && token.Type == JTokenType.String)
                    {
                        observation.AddUndoUnit(token.Value<string>());
                    }
                }
            }
            return observation;
        }
    }
}
