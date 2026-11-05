using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;
using Office.Contracts;

namespace Office.Parallel
{
    /// <summary>
    /// 主 Office 会话的专用 STA 线程。协议线程只投递委托并 await，不再直接调用 COM。
    /// 活动文档/选区类工具仍保持“主会话串行”语义；文件级批量任务走 WorkerProcessPool。
    /// </summary>
    public sealed class MainStaSession : IDisposable
    {
        private static readonly Lazy<MainStaSession> LazyInstance =
            new Lazy<MainStaSession>(delegate { return new MainStaSession(); });

        public static MainStaSession Instance
        {
            get { return LazyInstance.Value; }
        }

        private sealed class WorkItem
        {
            public Func<ToolResult> Action;
            public TaskCompletionSource<ToolResult> Completion;
        }

        private readonly BlockingCollection<WorkItem> _queue;
        private readonly ManualResetEventSlim _ready;
        private Thread _thread;
        private int _started;
        private int _stopping;
        private int _queued;
        private int _processed;
        private string _lastError;

        private MainStaSession()
        {
            _queue = new BlockingCollection<WorkItem>(1000);
            _ready = new ManualResetEventSlim(false);
        }

        public bool IsStarted
        {
            get { return Volatile.Read(ref _started) != 0; }
        }

        public int QueuedCount
        {
            get { return Volatile.Read(ref _queued); }
        }

        public int ProcessedCount
        {
            get { return Volatile.Read(ref _processed); }
        }

        public string LastError
        {
            get { return _lastError ?? string.Empty; }
        }

        /// <summary>启动主 STA 线程；重复调用安全。启动失败抛出异常，由 Program 决定是否降级。</summary>
        public void Start()
        {
            if (Interlocked.CompareExchange(ref _started, 1, 0) != 0)
            {
                _ready.Wait(TimeSpan.FromSeconds(30));
                return;
            }

            _thread = new Thread(Run);
            _thread.IsBackground = true;
            _thread.Name = "office-main-session";
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();

            int timeoutSec = Math.Max(1, Config.AppConfig.Current.WorkerStartupTimeoutSec);
            if (!_ready.Wait(TimeSpan.FromSeconds(timeoutSec)))
            {
                throw new InvalidOperationException(
                    "主 Office STA 会话启动超时（" + timeoutSec + " 秒）。");
            }
        }

        /// <summary>投递一个主会话委托；COM 只会在专用 STA 线程上执行。</summary>
        public Task<ToolResult> ExecuteAsync(Func<ToolResult> action)
        {
            if (action == null)
            {
                return Task.FromResult(ToolResult.Fail(
                    OfficeErrorMapper.ToCode(OfficeErrorCode.INVALID_PARAMS),
                    "主会话操作不能为空。",
                    false));
            }
            if (Volatile.Read(ref _stopping) != 0)
            {
                return Task.FromResult(ToolResult.Fail(
                    OfficeErrorMapper.ToCode(OfficeErrorCode.SERVER_BUSY),
                    "主 Office 会话正在关闭，操作未执行。",
                    true));
            }
            if (!IsStarted)
            {
                try
                {
                    Start();
                }
                catch (Exception ex)
                {
                    return Task.FromResult(ToolResult.Fail(
                        OfficeErrorMapper.ToCode(OfficeErrorCode.INTERNAL_ERROR),
                        "主 Office STA 会话启动失败：" + ex.Message,
                        true));
                }
            }

            WorkItem item = new WorkItem();
            item.Action = action;
            item.Completion = new TaskCompletionSource<ToolResult>(
                TaskCreationOptions.RunContinuationsAsynchronously);
            Interlocked.Increment(ref _queued);
            try
            {
                _queue.Add(item);
            }
            catch (Exception ex)
            {
                Interlocked.Decrement(ref _queued);
                return Task.FromResult(ToolResult.Fail(
                    OfficeErrorMapper.ToCode(OfficeErrorCode.SERVER_BUSY),
                    "主 Office 会话队列已关闭：" + ex.Message,
                    true));
            }
            return item.Completion.Task;
        }

        private void Run()
        {
            ComHelper.EnsureSta("MainStaSession.Run");
            _ready.Set();

            foreach (WorkItem item in _queue.GetConsumingEnumerable())
            {
                Interlocked.Decrement(ref _queued);
                ToolResult result;
                try
                {
                    result = item.Action();
                }
                catch (Exception ex)
                {
                    _lastError = ex.Message;
                    Console.Error.WriteLine("[main-session] 执行异常：" + ex.Message);
                    result = ToolResult.Fail(
                        OfficeErrorMapper.ToCode(OfficeErrorCode.INTERNAL_ERROR),
                        "主 Office 会话执行异常：" + ex.Message,
                        true);
                }

                if (result == null)
                {
                    result = ToolResult.Fail(
                        OfficeErrorMapper.ToCode(OfficeErrorCode.INTERNAL_ERROR),
                        "主 Office 会话没有返回结果（内部缺陷）。",
                        false);
                }
                Interlocked.Increment(ref _processed);
                item.Completion.TrySetResult(result);
            }

            try
            {
                OfficeSession.Shutdown();
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
                Console.Error.WriteLine("[main-session] 收口 Office 会话失败：" + ex.Message);
            }
        }

        /// <summary>停止接收新委托，排空队列，在主 STA 线程上 Quit 自动启动的 Office。</summary>
        public void Shutdown()
        {
            if (Interlocked.Exchange(ref _stopping, 1) != 0)
            {
                return;
            }
            try
            {
                _queue.CompleteAdding();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[main-session] 关闭队列失败：" + ex.Message);
            }

            Thread thread = _thread;
            if (thread != null && thread.IsAlive)
            {
                int timeoutSec = Math.Min(Math.Max(Config.AppConfig.Current.SessionTimeoutSec, 5), 60);
                if (!thread.Join(TimeSpan.FromSeconds(timeoutSec)))
                {
                    Console.Error.WriteLine("[main-session] 收口超时：STA 线程仍在运行。");
                }
            }
        }

        public void Dispose()
        {
            Shutdown();
            _queue.Dispose();
            _ready.Dispose();
        }
    }
}
