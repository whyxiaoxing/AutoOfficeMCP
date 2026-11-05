using Newtonsoft.Json.Linq;

namespace Office.Contracts
{
    /// <summary>
    /// MCP 工具调用的统一返回模型。
    /// 写信任合同约束：Success=true 必须代表操作真实生效；被拒绝的写操作一律 Success=false。
    /// </summary>
    public sealed class ToolResult
    {
        /// <summary>操作是否真实成功。</summary>
        public bool Success { get; set; }

        /// <summary>给用户/LLM 看的中文说明。</summary>
        public string Message { get; set; }

        /// <summary>结构化数据（Observation、行集、统计等）。</summary>
        public JObject Data { get; set; }

        /// <summary>稳定错误码；成功时为空字符串。</summary>
        public string ErrorCode { get; set; }

        /// <summary>调用方是否可以调整参数后重试。</summary>
        public bool Recoverable { get; set; }

        /// <summary>截图等二进制内容；非空时协议层追加 image content。</summary>
        public byte[] ImageBytes { get; set; }

        /// <summary>
        /// 内部诊断信息（异常链类型/消息/堆栈）。只写入 stderr 错误日志，
        /// 不进入 MCP 响应，避免把内部堆栈暴露给远端调用方。
        /// </summary>
        public string Diagnostic { get; set; }

        public static ToolResult Ok(string message, JObject data)
        {
            ToolResult result = new ToolResult();
            result.Success = true;
            result.Message = message ?? string.Empty;
            result.Data = data;
            result.ErrorCode = string.Empty;
            result.Recoverable = true;
            result.ImageBytes = null;
            result.Diagnostic = string.Empty;
            return result;
        }

        public static ToolResult Fail(string errorCode, string message, bool recoverable)
        {
            ToolResult result = new ToolResult();
            result.Success = false;
            // 失败结果必须在结构化载荷中暴露稳定错误码；调用方（含 golden 负样例）
            // 不解析中文消息，只认 errorCode。
            result.ErrorCode = string.IsNullOrWhiteSpace(errorCode) ? "INTERNAL_ERROR" : errorCode.Trim();
            // 任何工具失败都必须携带可展示的信息，禁止空消息进入协议层。
            result.Message = string.IsNullOrWhiteSpace(message)
                ? "工具执行失败（错误码 " + result.ErrorCode + "）。"
                : message.Trim();
            JObject data = new JObject();
            data["errorCode"] = result.ErrorCode;
            data["recoverable"] = recoverable;
            result.Data = data;
            result.Recoverable = recoverable;
            result.ImageBytes = null;
            result.Diagnostic = string.Empty;
            return result;
        }
    }
}
