using System;
using System.Text;
using System.Threading;
using AutoOfficeMCP.Common;
using AutoOfficeMCP.Protocol;
using AutoOfficeMCP.Tools;
using AutoOfficeMCP.Transport;
using AutoOfficeMCP.Worker;
using Office;
using Office.Config;
using Office.Parallel;

namespace AutoOfficeMCP
{
    /// <summary>
    /// MCP server 入口（HTTP/SSE）。所有日志走 stderr。
    /// COM 纪律要求主线程保持 STA，因此 Main 标记 [STAThread]，执行器也在该线程上调用。
    /// </summary>
    internal static class Program
    {
        [STAThread]
        private static int Main()
        {
            ConfigureConsoleEncoding();

            try
            {
                AppConfig config = AppConfig.Current;
                ToolRegistry registry = ToolRegistry.Load();
                ExecutorDiscovery discovery = new ExecutorDiscovery();

                WorkerSettings workerSettings = WorkerBootstrap.TryReadFromDirectory(Environment.CurrentDirectory);
                if (workerSettings != null)
                {
                    return WorkerHost.Run(registry, discovery, workerSettings);
                }

                // 工具说明直接生成到工作目录根目录。
                ToolDocumentationService.GenerateAll(registry);

                McpProtocolHandler primaryHandler = new McpProtocolHandler(
                    registry,
                    discovery,
                    "server",
                    "server",
                    "internal");
                // office_batch 调度桥：复用协议层的 registry 查找、高危确认与 discovery 路由。
                CommonToolExecutor.ToolDispatcher = primaryHandler.ExecuteToolByName;
                // 并行调度桥：office_batch_files 逐项执行同样复用协议层门禁与发现路由。
                ParallelDispatcher.ToolInvoker = primaryHandler.ExecuteToolByName;
                ParallelDispatcher.ToolExists = delegate(string toolName)
                {
                    return registry.Find(toolName) != null;
                };

                // PPT 的原生撤销与活动演示文稿语义要求主 STA 会话始终可用；
                // Excel/Word 走进程 worker，PPT 与 PPT 相关 common 工具走本会话。
                try
                {
                    MainStaSession.Instance.Start();
                }
                catch (Exception sessionError)
                {
                    Console.Error.WriteLine(
                        "主 Office STA 会话启动失败；直接工具将返回错误，批量 worker 仍可尝试启动。" +
                        Environment.NewLine + sessionError);
                }

                using (ManualResetEventSlim shutdownEvent = new ManualResetEventSlim(false))
                {
                    ConsoleCancelEventHandler cancelHandler = delegate(object sender, ConsoleCancelEventArgs cancelArgs)
                    {
                        cancelArgs.Cancel = true;
                        shutdownEvent.Set();
                    };
                    Console.CancelKeyPress += cancelHandler;

                    HttpTransportHost httpTransport = null;
                    try
                    {
                        httpTransport = new HttpTransportHost(registry, discovery, config);
                        httpTransport.ShutdownRequested += delegate { shutdownEvent.Set(); };
                        httpTransport.Start();
                        shutdownEvent.Wait();
                    }
                    finally
                    {
                        if (httpTransport != null)
                        {
                            httpTransport.Stop();
                            httpTransport.Dispose();
                        }
                        Console.CancelKeyPress -= cancelHandler;
                    }
                }

                // 退出前收口进程 worker 与主会话；最后再尝试收口进程级 Office 代理。
                ParallelDispatcher.Instance.Shutdown();
                MainStaSession.Instance.Shutdown();
                ShutdownOfficeSession();
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("MCP 主循环发生未处理异常。" + Environment.NewLine + ex);
                ShutdownParallelDispatcher();
                ShutdownOfficeSession();
                return 1;
            }
        }

        /// <summary>正常/异常退出都尝试关闭 server 自动启动的 Office；失败只写 stderr。</summary>
        private static void ShutdownOfficeSession()
        {
            try
            {
                OfficeSession.Shutdown();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[session] 退出时收口 Office 会话失败：" + ex.Message);
            }
        }

        /// <summary>异常退出路径也要尝试收口并行池；收口失败只写 stderr，不覆盖原始异常。</summary>
        private static void ShutdownParallelDispatcher()
        {
            try
            {
                ParallelDispatcher.Instance.Shutdown();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[parallel] 异常退出时收口并行调度器失败：" + ex.Message);
            }
        }

        /// <summary>强制控制台输出 UTF-8 且不带 BOM。</summary>
        private static void ConfigureConsoleEncoding()
        {
            try
            {
                Console.OutputEncoding = new UTF8Encoding(false);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[console] 设置控制台 UTF-8 编码失败：" + ex.Message);
            }
        }
    }
}
