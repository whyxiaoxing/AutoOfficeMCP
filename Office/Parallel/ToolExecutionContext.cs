using System.Threading;

namespace Office.Parallel
{
    /// <summary>
    /// 当前 MCP 请求的用户身份。HTTP 会话在协议线程设置，异步流经 await 后仍可读取；
    /// 并行批次据此把任务路由到同一用户的 STA worker。
    /// </summary>
    public static class ToolExecutionContext
    {
        private static readonly AsyncLocal<string> UserIdHolder = new AsyncLocal<string>();
        private static readonly AsyncLocal<string> ClientAddressHolder = new AsyncLocal<string>();
        private static readonly AsyncLocal<string> TransportKindHolder = new AsyncLocal<string>();

        public static string UserId
        {
            get
            {
                string value = UserIdHolder.Value;
                return string.IsNullOrWhiteSpace(value) ? "default" : value;
            }
            set
            {
                UserIdHolder.Value = string.IsNullOrWhiteSpace(value) ? "default" : value.Trim();
            }
        }

        /// <summary>
        /// 当前请求的客户端地址；用于 office_batch 等嵌套调用复用同一日志上下文。
        /// 未设置时返回空字符串，由调用方回退到 handler 自身字段。
        /// </summary>
        public static string ClientAddress
        {
            get { return ClientAddressHolder.Value ?? string.Empty; }
            set { ClientAddressHolder.Value = (value ?? string.Empty).Trim(); }
        }

        /// <summary>当前请求的传输方式（HTTP/SSE）；未设置时返回空字符串。</summary>
        public static string TransportKind
        {
            get { return TransportKindHolder.Value ?? string.Empty; }
            set { TransportKindHolder.Value = (value ?? string.Empty).Trim(); }
        }
    }
}
