using System;
using System.Diagnostics;
using System.Threading.Tasks;
using AutoOfficeMCP.Tools;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Office;
using Office.Config;
using Office.Contracts;
using Office.Parallel;
using Office.Security;

namespace AutoOfficeMCP.Protocol
{
    /// <summary>
    /// MCP 方法分发器：负责 JSON-RPC 2.0 行协议、initialize/tools/list/tools/call/ping/shutdown
    /// 以及错误码映射。stdout 只由 Program 写入本类返回的响应字符串。
    /// </summary>
    public sealed class McpProtocolHandler
    {
        public const string DefaultProtocolVersion = "2024-11-05";
        public const string ServerName = "AutoOfficeMCP";
        public const string ServerVersion = "0.1.0";

        private readonly ToolRegistry _registry;
        private readonly ExecutorDiscovery _discovery;
        private readonly string _userId;
        private readonly string _clientAddress;
        private readonly string _transportKind;

        private string _protocolVersion;
        private bool _shutdownRequested;

        public McpProtocolHandler(
            ToolRegistry registry,
            ExecutorDiscovery discovery,
            string userId,
            string clientAddress,
            string transportKind)
        {
            if (registry == null)
            {
                throw new ArgumentNullException("registry");
            }
            if (discovery == null)
            {
                throw new ArgumentNullException("discovery");
            }

            _registry = registry;
            _discovery = discovery;
            _userId = string.IsNullOrWhiteSpace(userId) ? "default" : userId.Trim();
            _clientAddress = string.IsNullOrWhiteSpace(clientAddress) ? "unknown" : clientAddress.Trim();
            _transportKind = string.IsNullOrWhiteSpace(transportKind) ? "unknown" : transportKind.Trim();
            _protocolVersion = DefaultProtocolVersion;
        }

        /// <summary>收到 shutdown 后置位；Program 写完响应帧即退出并返回 0。</summary>
        public bool ShutdownRequested
        {
            get { return _shutdownRequested; }
        }

        /// <summary>
        /// 按工具名执行已注册工具，供 office_batch 等跨工具调度使用。
        /// 与 tools/call 共用注册表查找、高危 confirm 门禁和执行器发现路由；
        /// 未知/已禁用工具返回 INVALID_PARAMS 的 ToolResult，而不是 JSON-RPC error。
        /// </summary>
        public ToolResult ExecuteToolByName(string toolName, JObject arguments)
        {
            string normalizedName = string.IsNullOrWhiteSpace(toolName) ? string.Empty : toolName.Trim();
            if (normalizedName.Length == 0)
            {
                return ToolResult.Fail("INVALID_PARAMS", "工具名必须是非空字符串。", false);
            }

            ToolDescriptor descriptor = _registry.Find(normalizedName);
            if (descriptor == null)
            {
                return ToolResult.Fail(
                    "INVALID_PARAMS",
                    "未知或已禁用的工具 " + normalizedName + "。",
                    false);
            }

            return ExecuteDescriptor(descriptor, arguments == null ? new JObject() : arguments);
        }

        /// <summary>
        /// 处理一行输入。返回 null 表示无需响应（通知、空行、未识别 notifications/*）。
        /// 返回字符串时必须是一行紧凑 JSON-RPC 响应。
        /// </summary>
        public string HandleLine(string line)
        {
            return HandleLineAsync(line).ConfigureAwait(false).GetAwaiter().GetResult();
        }

        /// <summary>
        /// 异步处理一行输入；传输入口 await 本方法，COM 由 MainStaSession/WorkerProcessPool 执行。
        /// 返回 null 表示无需响应（通知、空行、未识别 notifications/*）。
        /// </summary>
        public async Task<string> HandleLineAsync(string line)
        {
            ToolExecutionContext.UserId = _userId;
            ToolExecutionContext.ClientAddress = _clientAddress;
            ToolExecutionContext.TransportKind = _transportKind;
            if (string.IsNullOrWhiteSpace(line))
            {
                return null;
            }

            JToken parsed;
            try
            {
                parsed = JToken.Parse(line);
            }
            catch (JsonException ex)
            {
                Console.Error.WriteLine("JSON 解析失败：" + ex.Message);
                return Serialize(JsonRpcResponse.CreateError(
                    null,
                    JsonRpcErrorCodes.ParseError,
                    "Parse error: " + ex.Message,
                    null));
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("读取 JSON-RPC 报文失败。" + Environment.NewLine + ex);
                return Serialize(JsonRpcResponse.CreateError(
                    null,
                    JsonRpcErrorCodes.ParseError,
                    "Parse error: " + ex.Message,
                    null));
            }

            JObject root = parsed as JObject;
            if (root == null)
            {
                Console.Error.WriteLine("JSON-RPC 报文不是对象，返回 -32600。");
                return Serialize(JsonRpcResponse.CreateError(
                    null,
                    JsonRpcErrorCodes.InvalidRequest,
                    "Invalid Request: 报文必须是 JSON 对象（不支持批量数组）。",
                    null));
            }

            JsonRpcRequest request;
            JsonRpcError parseError;
            if (!JsonRpcRequest.TryParse(root, out request, out parseError))
            {
                Console.Error.WriteLine("无效 JSON-RPC 请求：" + parseError.Message);
                return Serialize(JsonRpcResponse.CreateError(parseError));
            }

            // notifications/* 一律不响应；初始化和 shutdown 通知仍更新内部状态。
            if (request.Method.StartsWith("notifications/", StringComparison.OrdinalIgnoreCase)
                || request.IsNotification)
            {
                HandleNotification(request);
                return null;
            }

            try
            {
                return Serialize(await HandleRequestAsync(request).ConfigureAwait(false));
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("处理请求时发生内部异常：" + request.Method +
                                        Environment.NewLine + ex);
                return Serialize(JsonRpcResponse.CreateError(
                    request.Id,
                    JsonRpcErrorCodes.InternalError,
                    "Internal error: " + ex.Message,
                    null));
            }
        }

        private void HandleNotification(JsonRpcRequest request)
        {
            if (string.Equals(request.Method, "notifications/initialized", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            if (string.Equals(request.Method, "shutdown", StringComparison.OrdinalIgnoreCase))
            {
                _shutdownRequested = true;
                return;
            }

            // 未识别的 notifications/* 静默丢弃；其它无 id 报文同样不需要响应。
        }

        private async Task<JObject> HandleRequestAsync(JsonRpcRequest request)
        {
            switch (request.Method)
            {
                case "initialize":
                    return HandleInitialize(request);
                case "tools/list":
                    return JsonRpcResponse.CreateResult(request.Id, _registry.CreateToolsListResult());
                case "tools/call":
                    return await HandleToolsCallAsync(request).ConfigureAwait(false);
                case "ping":
                    return JsonRpcResponse.CreateResult(request.Id, new JObject());
                case "shutdown":
                    _shutdownRequested = true;
                    return JsonRpcResponse.CreateResult(request.Id, new JObject());
                default:
                    Console.Error.WriteLine("未知请求方法：" + request.Method);
                    return JsonRpcResponse.CreateError(
                        request.Id,
                        JsonRpcErrorCodes.MethodNotFound,
                        "Method not found: " + request.Method,
                        null);
            }
        }

        private JObject HandleInitialize(JsonRpcRequest request)
        {
            string requestedVersion = DefaultProtocolVersion;
            JObject parameters = request.Params;

            if (parameters != null)
            {
                JToken versionToken = parameters["protocolVersion"];
                if (versionToken != null && versionToken.Type != JTokenType.Null)
                {
                    if (versionToken.Type != JTokenType.String)
                    {
                        return JsonRpcResponse.CreateError(
                            request.Id,
                            JsonRpcErrorCodes.InvalidParams,
                            "Invalid params: protocolVersion 必须是字符串。",
                            null);
                    }

                    string value = versionToken.Value<string>();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        requestedVersion = value.Trim();
                    }
                }

            }

            _protocolVersion = requestedVersion;

            JObject toolsCapability = new JObject();
            toolsCapability["listChanged"] = false;

            JObject capabilities = new JObject();
            capabilities["tools"] = toolsCapability;

            JObject serverInfo = new JObject();
            serverInfo["name"] = ServerName;
            serverInfo["version"] = ServerVersion;

            JObject result = new JObject();
            result["protocolVersion"] = _protocolVersion;
            result["capabilities"] = capabilities;
            result["serverInfo"] = serverInfo;
            return JsonRpcResponse.CreateResult(request.Id, result);
        }

        private async Task<JObject> HandleToolsCallAsync(JsonRpcRequest request)
        {
            JObject parameters = request.Params;
            if (parameters == null)
            {
                return JsonRpcResponse.CreateError(
                    request.Id,
                    JsonRpcErrorCodes.InvalidParams,
                    "Invalid params: tools/call 需要 params.name 与 params.arguments。",
                    null);
            }

            JToken nameToken = parameters["name"];
            if (nameToken == null
                || nameToken.Type != JTokenType.String
                || string.IsNullOrWhiteSpace(nameToken.Value<string>()))
            {
                return JsonRpcResponse.CreateError(
                    request.Id,
                    JsonRpcErrorCodes.InvalidParams,
                    "Invalid params: 工具名必须是非空字符串。",
                    null);
            }
            string toolName = nameToken.Value<string>().Trim();

            JObject arguments = new JObject();
            JToken argumentsToken = parameters["arguments"];
            if (argumentsToken != null && argumentsToken.Type != JTokenType.Null)
            {
                arguments = argumentsToken as JObject;
                if (arguments == null)
                {
                    return JsonRpcResponse.CreateError(
                        request.Id,
                        JsonRpcErrorCodes.InvalidParams,
                        "Invalid params: arguments 必须是 JSON 对象。",
                        null);
                }
            }

            ToolDescriptor descriptor = _registry.Find(toolName);
            if (descriptor == null)
            {
                Console.Error.WriteLine("tools/call 未知或已禁用工具：" + toolName);
                return JsonRpcResponse.CreateError(
                    request.Id,
                    JsonRpcErrorCodes.InvalidParams,
                    "Invalid params: 未知或已禁用的工具 " + toolName + "。",
                    null);
            }

            ToolResult toolResult = await ExecuteDescriptorAsync(descriptor, arguments).ConfigureAwait(false);
            return JsonRpcResponse.CreateResult(request.Id, BuildToolCallResult(toolResult));
        }

        private ToolResult ExecuteDescriptor(ToolDescriptor descriptor, JObject arguments)
        {
            JObject normalizedArguments;
            string pathError;
            if (!WorkspacePathResolver.TryNormalizeArguments(
                    descriptor.AppType,
                    arguments,
                    out normalizedArguments,
                    out pathError))
            {
                return FailTool(descriptor, "INVALID_PARAMS", pathError, false);
            }
            arguments = normalizedArguments;

            if (descriptor.IsHighRisk && !IsConfirmed(arguments))
            {
                return FailTool(
                    descriptor,
                    "CONFIRMATION_REQUIRED",
                    "高风险工具 " + descriptor.ToolName + " 需要显式确认：请在 arguments 中设置 confirm=true 后重试。",
                    true);
            }

            string clipboardMessage;
            if (!ClipboardGuard.EnsureAllowed(descriptor.ToolName, out clipboardMessage))
            {
                return FailTool(descriptor, "CLIPBOARD_DISABLED", clipboardMessage, false);
            }

            Stopwatch stopwatch = Stopwatch.StartNew();
            ToolResult result = _discovery.Execute(descriptor, arguments);
            stopwatch.Stop();
            LogToolResult(descriptor, result, stopwatch.ElapsedMilliseconds);
            return result;
        }

        /// <summary>
        /// T10 异步执行路径：批处理走 IAsyncToolExecutor；其余工具投递到主 STA 会话。
        /// 协议线程只 await，不直接执行 COM，也不阻塞读循环。
        /// </summary>
        private async Task<ToolResult> ExecuteDescriptorAsync(ToolDescriptor descriptor, JObject arguments)
        {
            JObject normalizedArguments;
            string pathError;
            if (!WorkspacePathResolver.TryNormalizeArguments(
                    descriptor.AppType,
                    arguments,
                    out normalizedArguments,
                    out pathError))
            {
                return FailTool(descriptor, "INVALID_PARAMS", pathError, false);
            }
            arguments = normalizedArguments;

            if (descriptor.IsHighRisk && !IsConfirmed(arguments))
            {
                return FailTool(
                    descriptor,
                    "CONFIRMATION_REQUIRED",
                    "高风险工具 " + descriptor.ToolName + " 需要显式确认：请在 arguments 中设置 confirm=true 后重试。",
                    true);
            }

            string clipboardMessage;
            if (!ClipboardGuard.EnsureAllowed(descriptor.ToolName, out clipboardMessage))
            {
                return FailTool(descriptor, "CLIPBOARD_DISABLED", clipboardMessage, false);
            }

            bool localTool = _discovery.IsLocalTool(descriptor, arguments);
            bool mainSessionTool = RequiresMainSession(descriptor);
            bool fileTool = string.Equals(descriptor.AppType, "excel", StringComparison.OrdinalIgnoreCase)
                            || string.Equals(descriptor.AppType, "word", StringComparison.OrdinalIgnoreCase);
            if (string.Equals(AppConfig.Current.WorkerConcurrencyMode, "throughput", StringComparison.Ordinal)
                && fileTool
                && !localTool
                && ReadFilePath(arguments).Length == 0)
            {
                return FailTool(
                    descriptor,
                    "INVALID_PARAMS",
                    "throughput 模式只允许显式带 file 的无状态工具；请使用 office_batch_files 并显式传 file。",
                    false);
            }
            if (!_discovery.HasAsyncExecutor(descriptor)
                && !localTool
                && !mainSessionTool)
            {
                Stopwatch workerWatch = Stopwatch.StartNew();
                ToolResult workerResult = await ExecuteFileToolThroughWorkerAsync(descriptor, arguments)
                    .ConfigureAwait(false);
                workerWatch.Stop();
                LogToolResult(descriptor, workerResult, workerWatch.ElapsedMilliseconds);
                return workerResult;
            }

            Stopwatch stopwatch = Stopwatch.StartNew();
            ToolResult result;
            if (_discovery.HasAsyncExecutor(descriptor))
            {
                result = await _discovery.ExecuteAsync(descriptor, arguments).ConfigureAwait(false);
            }
            else if (localTool)
            {
                result = _discovery.Execute(descriptor, arguments);
            }
            else
            {
                result = await MainStaSession.Instance.ExecuteAsync(delegate
                {
                    return _discovery.Execute(descriptor, arguments);
                }).ConfigureAwait(false);
            }
            stopwatch.Stop();
            LogToolResult(descriptor, result, stopwatch.ElapsedMilliseconds);
            return result;
        }

        /// <summary>
        /// 工具结果日志：成功只写 debug；失败写 error，并携带客户端 IP、传输方式
        /// （HTTP/SSE）和诊断异常链。诊断只进 stderr，不进入 MCP 响应。
        /// </summary>
        private ToolResult FailTool(
            ToolDescriptor descriptor,
            string errorCode,
            string message,
            bool recoverable)
        {
            ToolResult failure = ToolResult.Fail(errorCode, message, recoverable);
            LogToolResult(descriptor, failure, 0);
            return failure;
        }

        private void LogToolResult(ToolDescriptor descriptor, ToolResult result, long elapsedMilliseconds)
        {
            if (result == null)
            {
                Console.Error.WriteLine("tools/call " + descriptor.ToolName +
                                        " 返回空结果；client=" + _clientAddress +
                                        "；transport=" + _transportKind + "。");
                return;
            }

            if (result.Success)
            {
                return;
            }

            string clientAddress = ToolExecutionContext.ClientAddress;
            if (string.IsNullOrWhiteSpace(clientAddress))
            {
                clientAddress = _clientAddress;
            }
            string transportKind = ToolExecutionContext.TransportKind;
            if (string.IsNullOrWhiteSpace(transportKind))
            {
                transportKind = _transportKind;
            }

            string message = "tools/call 失败：" + descriptor.ToolName +
                             "；client=" + clientAddress +
                             "；transport=" + transportKind +
                             "；errorCode=" + (result.ErrorCode ?? string.Empty) +
                             "；recoverable=" + result.Recoverable +
                             "；message=" + (result.Message ?? string.Empty) +
                             "；耗时 " + elapsedMilliseconds.ToString() + " ms。";
            if (!string.IsNullOrWhiteSpace(result.Diagnostic))
            {
                message = message + Environment.NewLine + result.Diagnostic;
            }
            Console.Error.WriteLine(message);
        }

        /// <summary>
        /// 进程 worker 模式：直连 Excel/Word 工具必须带显式 file，包装成单作业投递 worker。
        /// Gateway 不执行 COM；活动文档/选区工具没有 file，直接拒绝。
        /// </summary>
        private async Task<ToolResult> ExecuteFileToolThroughWorkerAsync(
            ToolDescriptor descriptor,
            JObject arguments)
        {
            if (!string.Equals(descriptor.AppType, "excel", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(descriptor.AppType, "word", StringComparison.OrdinalIgnoreCase))
            {
                return ToolResult.Fail(
                    "INVALID_PARAMS",
                    "进程 worker 模式下不支持 common/活动文档类直连工具「" + descriptor.ToolName +
                    "」；请使用 office_batch_files 并显式传 file。",
                    false);
            }
            if (string.Equals(descriptor.ToolName, "excel_profile_files", StringComparison.Ordinal)
                || string.Equals(descriptor.ToolName, "excel_query_files", StringComparison.Ordinal))
            {
                return ToolResult.Fail(
                    "INVALID_PARAMS",
                    "进程 worker 模式下多文件只读工具请设置 backend=openxml；" +
                    "COM 后端需要主 STA 会话，Gateway 不执行 COM。",
                    false);
            }

            string filePath = ReadFilePath(arguments);
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return ToolResult.Fail(
                    "INVALID_PARAMS",
                    "进程 worker 模式下工具「" + descriptor.ToolName +
                    "」需要显式 file 参数（当前活动文档语义在远端多用户模式下不可用）。",
                    false);
            }

            string savePolicy = "none";
            if (arguments != null)
            {
                JToken saveToken = arguments["savePolicy"];
                if (saveToken != null && saveToken.Type == JTokenType.String)
                {
                    string value = saveToken.Value<string>();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        savePolicy = value.Trim().ToLowerInvariant();
                    }
                }
                else if (!string.Equals(descriptor.Category, "read", StringComparison.OrdinalIgnoreCase))
                {
                    savePolicy = "save";
                }
            }

            OfficeWorkItem item = new OfficeWorkItem();
            item.AppType = descriptor.AppType;
            item.ToolName = descriptor.ToolName;
            item.FilePath = filePath;
            item.Args = arguments == null ? new JObject() : (JObject)arguments.DeepClone();
            item.SavePolicy = savePolicy;
            item.TimeoutSec = AppConfig.Current.ParallelItemTimeoutSec;

            OperationResult operation = await ParallelDispatcher.Instance.Submit(item).ConfigureAwait(false);
            return operation.ToToolResult();
        }

        /// <summary>
        /// 是否在 Gateway 主 STA 会话执行：
        /// PPT 直连工具依赖活动演示文稿与 PowerPoint 原生撤销；Common 工具需要操作活动文档
        /// 或主会话撤销栈。Excel/Word 文件工具继续走进程 worker。
        /// </summary>
        private static bool RequiresMainSession(ToolDescriptor descriptor)
        {
            if (descriptor == null)
            {
                return false;
            }
            return string.Equals(descriptor.AppType, "ppt", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(descriptor.AppType, "common", StringComparison.OrdinalIgnoreCase);
        }

        private static string ReadFilePath(JObject arguments)
        {
            if (arguments == null)
            {
                return string.Empty;
            }
            JToken token = arguments["file"];
            if (token == null || token.Type == JTokenType.Null)
            {
                token = arguments["filePath"];
            }
            if (token == null || token.Type != JTokenType.String)
            {
                return string.Empty;
            }
            string value = token.Value<string>();
            return value == null ? string.Empty : value.Trim();
        }

        /// <summary>
        /// ToolResult → MCP content[]。Data 以紧凑 JSON 追加在文本尾部，避免结构化结果丢失；
        /// ImageBytes 非空时追加 image content（参考客户端支持 data + mimeType）。
        /// </summary>
        private static JObject BuildToolCallResult(ToolResult toolResult)
        {
            ToolResult result = toolResult;
            if (result == null)
            {
                result = ToolResult.Fail("INTERNAL_ERROR", "执行器未返回结果。", true);
            }

            string text = result.Message ?? string.Empty;
            if (result.Data != null && result.Data.HasValues)
            {
                string dataText = result.Data.ToString(Formatting.None);
                text = string.IsNullOrEmpty(text) ? dataText : text + "\n" + dataText;
            }
            if (string.IsNullOrEmpty(text))
            {
                text = result.Success ? "操作完成。" : "操作失败。";
            }

            JObject textItem = new JObject();
            textItem["type"] = "text";
            textItem["text"] = text;

            JArray content = new JArray();
            content.Add(textItem);

            if (result.ImageBytes != null && result.ImageBytes.Length > 0)
            {
                JObject imageItem = new JObject();
                imageItem["type"] = "image";
                imageItem["data"] = Convert.ToBase64String(result.ImageBytes);
                imageItem["mimeType"] = "image/png";
                content.Add(imageItem);
            }

            JObject payload = new JObject();
            payload["content"] = content;
            payload["isError"] = !result.Success;
            if (!result.Success)
            {
                string errorCode = string.IsNullOrWhiteSpace(result.ErrorCode)
                    ? "INTERNAL_ERROR"
                    : result.ErrorCode;
                // 部分 MCP 客户端（如 ShareRibbon）只读取 result.errorMessage / result.error，
                // 不解析 content；这里同时提供文本与错误码，保证每个工具的错误都能被展示。
                string errorMessage = string.IsNullOrWhiteSpace(result.Message)
                    ? "工具执行失败（错误码 " + errorCode + "）。"
                    : result.Message.Trim();
                payload["errorMessage"] = errorMessage;
                // 兼容只读取 result.error 的旧客户端；标准客户端仍以 content + isError 为准。
                payload["error"] = errorMessage;
                payload["errorCode"] = errorCode;
                payload["recoverable"] = result.Recoverable;
            }
            return payload;
        }

        private static bool IsConfirmed(JObject arguments)
        {
            if (arguments == null)
            {
                return false;
            }

            JToken confirm = arguments["confirm"];
            if (confirm == null || confirm.Type != JTokenType.Boolean)
            {
                return false;
            }
            return confirm.Value<bool>();
        }

        private static string GetString(JObject root, string propertyName)
        {
            if (root == null)
            {
                return string.Empty;
            }

            JToken token = root[propertyName];
            if (token == null || token.Type == JTokenType.Null)
            {
                return string.Empty;
            }
            return token.ToString();
        }

        private static string Serialize(JObject response)
        {
            return response.ToString(Formatting.None);
        }
    }
}
