using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using AutoOfficeMCP.Protocol;
using AutoOfficeMCP.Tools;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Office.Config;
using Office.Parallel;

namespace AutoOfficeMCP.Transport
{
    /// <summary>
    /// 远端 MCP HTTP 入口（Streamable HTTP 的 JSON 子集）：
    /// POST /mcp 是唯一 MCP 端点（全量工具注册）；Mcp-Session-Id 绑定单会话；可选 transport.tokens 做用户隔离。
    /// /xlsx /word /ppt 只提供工具文档（GET 总览 + GET /&lt;工具名称&gt; 单工具页），不承载 MCP 协议。
    /// 不在此类执行 COM；tools/call 由协议层 await MainStaSession/WorkerProcessPool。
    /// </summary>
    public sealed class HttpTransportHost : IDisposable
    {
        private sealed class Session
        {
            public McpProtocolHandler Handler;
            public string UserId;
            public string ClientAddress;
            public string TransportKind;
            public SemaphoreSlim Gate;
            public DateTime LastSeenUtc;
        }

        /// <summary>文档路由（/xlsx、/word、/ppt）绑定的 appType 与过滤后的注册表视图（总览/单工具页查找）。</summary>
        private sealed class RouteInfo
        {
            public ToolRoute Route;
            public ToolRegistry Registry;
        }

        /// <summary>旧版 HTTP+SSE 兼容会话：GET /sse 保持流，POST /messages 把响应推回该流。</summary>
        private sealed class SseSession
        {
            public string SessionId;
            public string UserId;
            public McpProtocolHandler Handler;
            public SemaphoreSlim Gate;
            public SemaphoreSlim Signal;
            public ConcurrentQueue<string> Queue;
            public CancellationTokenSource Cancellation;
            public HttpListenerResponse Response;
        }

        private sealed class HttpError : Exception
        {
            public readonly int StatusCode;

            public HttpError(int statusCode, string message)
                : base(message)
            {
                StatusCode = statusCode;
            }
        }

        private readonly ToolRegistry _registry;
        private readonly ExecutorDiscovery _discovery;
        private readonly AppConfig _config;
        private readonly ConcurrentDictionary<string, Session> _sessions;
        private readonly ConcurrentDictionary<string, SseSession> _sseSessions;
        private readonly System.Collections.Generic.Dictionary<string, RouteInfo> _routes;
        private readonly string _prefix;
        private readonly System.Collections.Generic.Dictionary<string, string> _tokens;
        private readonly int _maxBodyBytes;
        private readonly int _maxInflightPerSession;
        private HttpListener _listener;
        private Task _acceptTask;
        private volatile bool _stopping;
        private long _requests;
        private long _errors;

        public HttpTransportHost(ToolRegistry registry, ExecutorDiscovery discovery, AppConfig config)
        {
            _registry = registry;
            _discovery = discovery;
            _config = config ?? AppConfig.Current;
            _sessions = new ConcurrentDictionary<string, Session>(StringComparer.Ordinal);
            _sseSessions = new ConcurrentDictionary<string, SseSession>(StringComparer.Ordinal);
            _routes = new System.Collections.Generic.Dictionary<string, RouteInfo>(StringComparer.OrdinalIgnoreCase);
            foreach (ToolRoute route in ToolRouteCatalog.Routes)
            {
                RouteInfo info = new RouteInfo();
                info.Route = route;
                info.Registry = _registry.CreateScoped(route.AppType, true);
                _routes[route.Key] = info;
            }
            string bind = string.IsNullOrWhiteSpace(_config.TransportBind) ? "127.0.0.1" : _config.TransportBind.Trim();
            if (string.Equals(bind, "*", StringComparison.Ordinal)
                || string.Equals(bind, "0.0.0.0", StringComparison.Ordinal))
            {
                bind = "+";
            }
            _prefix = "http://" + bind + ":" + _config.TransportPort.ToString() + "/";
            _tokens = _config.TransportTokens;
            _maxBodyBytes = Math.Max(1024, _config.TransportMaxBodyBytes);
            _maxInflightPerSession = Math.Max(1, _config.ServerMaxInflightPerSession);
        }

        public event Action ShutdownRequested;

        public long RequestCount
        {
            get { return Interlocked.Read(ref _requests); }
        }

        public long ErrorCount
        {
            get { return Interlocked.Read(ref _errors); }
        }

        public void Start()
        {
            _listener = new HttpListener();
            _listener.Prefixes.Add(_prefix);
            _listener.Start();
            _acceptTask = Task.Run((Func<Task>)AcceptLoopAsync);
            string authMode = _tokens.Count > 0 ? "tokens(" + _tokens.Count + ")" : "none";
            string workspace = WorkspacePathResolver.WorkspaceRoot;
            if (workspace.Length == 0)
            {
                workspace = "不可用";
            }
            Console.Error.WriteLine("[NOTICE] HTTP 监听 " + _prefix + "（IP 即 bind 地址；auth=" + authMode + "）" +
                                    "；HTTP: " + _prefix + "mcp" +
                                    "；SSE: " + _prefix + "sse" +
                                    "；工具文档: /xlsx /word /ppt（单工具: /xlsx/<工具名>）" +
                                    "；文件: /files/upload /files/download" +
                                    "；工作目录: " + workspace + "。");
        }

        private async Task AcceptLoopAsync()
        {
            while (!_stopping)
            {
                HttpListenerContext context;
                try
                {
                    context = await _listener.GetContextAsync().ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    if (!_stopping)
                    {
                        Console.Error.WriteLine("HTTP 接受连接失败：" + ex.Message);
                    }
                    break;
                }

                Task ignored = HandleContextAsync(context);
            }
        }

        private async Task HandleContextAsync(HttpListenerContext context)
        {
            HttpListenerRequest request = context.Request;
            HttpListenerResponse response = context.Response;
            Interlocked.Increment(ref _requests);
            try
            {
                string path = request.Url == null ? "/" : request.Url.AbsolutePath;
                if (string.Equals(path, "/health", StringComparison.OrdinalIgnoreCase))
                {
                    await HandleHealthAsync(request, response).ConfigureAwait(false);
                    return;
                }
                if (string.Equals(path, "/shutdown", StringComparison.OrdinalIgnoreCase))
                {
                    await HandleShutdownAsync(request, response).ConfigureAwait(false);
                    return;
                }
                if (string.Equals(path, "/sse", StringComparison.OrdinalIgnoreCase))
                {
                    await HandleSseAsync(request, response).ConfigureAwait(false);
                    return;
                }
                if (string.Equals(path, "/messages", StringComparison.OrdinalIgnoreCase))
                {
                    await HandleMessagesAsync(request, response).ConfigureAwait(false);
                    return;
                }
                if (string.Equals(path, "/files/upload", StringComparison.OrdinalIgnoreCase))
                {
                    await HandleFileUploadAsync(request, response).ConfigureAwait(false);
                    return;
                }
                if (string.Equals(path, "/files/download", StringComparison.OrdinalIgnoreCase))
                {
                    await HandleFileDownloadAsync(request, response).ConfigureAwait(false);
                    return;
                }

                RouteInfo routeInfo;
                string routeRest;
                if (TryResolveRoute(path, out routeInfo, out routeRest))
                {
                    await HandleRouteAsync(routeInfo, routeRest, request, response).ConfigureAwait(false);
                    return;
                }

                if (string.Equals(path, "/mcp", StringComparison.OrdinalIgnoreCase))
                {
                    await HandleMcpAsync(request, response).ConfigureAwait(false);
                    return;
                }
                await WriteJsonAsync(response, 404, ErrorObject("NOT_FOUND", "未知路径：" + path), null)
                    .ConfigureAwait(false);
            }
            catch (HttpError ex)
            {
                Interlocked.Increment(ref _errors);
                await WriteJsonAsync(response, ex.StatusCode, ErrorObject("HTTP_ERROR", ex.Message), null)
                    .ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                Interlocked.Increment(ref _errors);
                string errorPath = request.Url == null ? string.Empty : request.Url.AbsolutePath;
                bool errorSse = string.Equals(errorPath, "/sse", StringComparison.OrdinalIgnoreCase)
                                || string.Equals(errorPath, "/messages", StringComparison.OrdinalIgnoreCase)
                                || (request.AcceptTypes != null && Array.Exists(
                                    request.AcceptTypes,
                                    delegate(string value)
                                    {
                                        return value != null && value.IndexOf(
                                            "text/event-stream",
                                            StringComparison.OrdinalIgnoreCase) >= 0;
                                    }));
                string transportKind = errorSse ? "SSE" : "HTTP";
                Console.Error.WriteLine("HTTP 请求处理失败：client=" + GetClientAddress(request) +
                                        "；transport=" + transportKind +
                                        "；url=" + (request.Url == null ? string.Empty : request.Url.ToString()) +
                                        Environment.NewLine + ex);
                try
                {
                    await WriteJsonAsync(response, 500, ErrorObject("INTERNAL_ERROR", ex.Message), null)
                        .ConfigureAwait(false);
                }
                catch (Exception writeError)
                {
                    Console.Error.WriteLine("[http] 写错误响应失败：" + writeError.Message);
                }
            }
            finally
            {
                try
                {
                    response.Close();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[http] 关闭响应失败：" + ex.Message);
                }
            }
        }

        private async Task HandleHealthAsync(HttpListenerRequest request, HttpListenerResponse response)
        {
            string userId;
            int status;
            if (!Authorize(request, out userId, out status))
            {
                await WriteJsonAsync(response, status, ErrorObject("UNAUTHORIZED", "认证失败。"), null)
                    .ConfigureAwait(false);
                return;
            }

            JObject health = new JObject();
            health["ok"] = true;
            health["transport"] = "http";
            health["prefix"] = _prefix;
            health["sessions"] = _sessions.Count;
            health["sseSessions"] = _sseSessions.Count;
            health["requests"] = RequestCount;
            health["errors"] = ErrorCount;
            health["mainSession"] = MainStaSession.Instance.IsStarted;
            health["mainQueued"] = MainStaSession.Instance.QueuedCount;
            health["mainProcessed"] = MainStaSession.Instance.ProcessedCount;
            health["workspace"] = WorkspacePathResolver.WorkspaceRoot;
            JArray routeList = new JArray();
            foreach (ToolRoute route in ToolRouteCatalog.Routes)
            {
                JObject routeItem = new JObject();
                routeItem["key"] = route.Key;
                routeItem["appType"] = route.AppType;
                routeItem["docs"] = "/" + route.Key;
                routeItem["tool"] = "/" + route.Key + "/<工具名称>";
                routeList.Add(routeItem);
            }
            health["routes"] = routeList;
            health["parallel"] = ParallelDispatcher.Instance.GetStatus();
            await WriteJsonAsync(response, 200, health, null).ConfigureAwait(false);
        }

        private async Task HandleShutdownAsync(HttpListenerRequest request, HttpListenerResponse response)
        {
            string userId;
            int status;
            if (!Authorize(request, out userId, out status))
            {
                await WriteJsonAsync(response, status, ErrorObject("UNAUTHORIZED", "认证失败。"), null)
                    .ConfigureAwait(false);
                return;
            }
            if (!string.Equals(request.HttpMethod, "POST", StringComparison.OrdinalIgnoreCase))
            {
                await WriteJsonAsync(response, 405, ErrorObject("METHOD_NOT_ALLOWED", "请使用 POST /shutdown。"), null)
                    .ConfigureAwait(false);
                return;
            }

            JObject payload = new JObject();
            payload["ok"] = true;
            payload["message"] = "server shutdown requested";
            await WriteJsonAsync(response, 200, payload, null).ConfigureAwait(false);
            Action handler = ShutdownRequested;
            if (handler != null)
            {
                Task shutdownTask = Task.Run(handler);
            }
        }

        /// <summary>旧版 HTTP+SSE：GET /sse 建立长连接，先发 endpoint 事件，再推送 message 事件。</summary>
        private async Task HandleSseAsync(HttpListenerRequest request, HttpListenerResponse response)
        {
            string userId;
            int status;
            if (!Authorize(request, out userId, out status))
            {
                await WriteJsonAsync(response, status, ErrorObject("UNAUTHORIZED", "认证失败。"), null)
                    .ConfigureAwait(false);
                return;
            }
            if (!string.Equals(request.HttpMethod, "GET", StringComparison.OrdinalIgnoreCase))
            {
                response.Headers["Allow"] = "GET";
                await WriteJsonAsync(response, 405, ErrorObject("METHOD_NOT_ALLOWED", "请使用 GET /sse。"), null)
                    .ConfigureAwait(false);
                return;
            }

            string sessionId = Guid.NewGuid().ToString("N");
            SseSession session = new SseSession();
            session.SessionId = sessionId;
            session.UserId = string.IsNullOrWhiteSpace(userId) ? "anonymous" : userId;
            session.Handler = new McpProtocolHandler(
                _registry,
                _discovery,
                session.UserId,
                GetClientAddress(request),
                "SSE");
            session.Gate = new SemaphoreSlim(_maxInflightPerSession, _maxInflightPerSession);
            session.Signal = new SemaphoreSlim(0);
            session.Queue = new ConcurrentQueue<string>();
            session.Cancellation = new CancellationTokenSource();
            session.Response = response;
            _sseSessions[sessionId] = session;

            response.StatusCode = 200;
            response.ContentType = "text/event-stream; charset=utf-8";
            response.ContentEncoding = new UTF8Encoding(false);
            response.SendChunked = true;
            response.KeepAlive = true;
            response.Headers["Cache-Control"] = "no-cache";
            response.Headers["Connection"] = "keep-alive";
            response.Headers["Mcp-Session-Id"] = sessionId;

            try
            {
                await WriteSseEventAsync(response, "endpoint", "/messages?sessionId=" + sessionId)
                    .ConfigureAwait(false);
                while (!_stopping && !session.Cancellation.IsCancellationRequested)
                {
                    Task<bool> wait = session.Signal.WaitAsync(TimeSpan.FromSeconds(15));
                    bool signaled = await wait.ConfigureAwait(false);
                    if (session.Cancellation.IsCancellationRequested)
                    {
                        break;
                    }
                    bool wrote = false;
                    string message;
                    while (session.Queue.TryDequeue(out message))
                    {
                        await WriteSseEventAsync(response, "message", message).ConfigureAwait(false);
                        wrote = true;
                    }
                    if (!signaled && !wrote)
                    {
                        await WriteSseCommentAsync(response, "keep-alive").ConfigureAwait(false);
                    }
                }
            }
            catch (Exception ex)
            {
                if (!_stopping && !session.Cancellation.IsCancellationRequested)
                {
                    Console.Error.WriteLine("SSE 会话断开：" + sessionId + "；" + ex.Message);
                }
            }
            finally
            {
                SseSession removed;
                _sseSessions.TryRemove(sessionId, out removed);
                try
                {
                    session.Gate.Dispose();
                    session.Signal.Dispose();
                    session.Cancellation.Dispose();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[http] 释放 SSE 会话失败：" + ex.Message);
                }
            }
        }

        /// <summary>旧版 HTTP+SSE：POST /messages?sessionId=...，响应通过 /sse 流返回。</summary>
        private async Task HandleMessagesAsync(HttpListenerRequest request, HttpListenerResponse response)
        {
            string userId;
            int status;
            if (!Authorize(request, out userId, out status))
            {
                await WriteJsonAsync(response, status, ErrorObject("UNAUTHORIZED", "认证失败。"), null)
                    .ConfigureAwait(false);
                return;
            }
            if (!string.Equals(request.HttpMethod, "POST", StringComparison.OrdinalIgnoreCase))
            {
                response.Headers["Allow"] = "POST";
                await WriteJsonAsync(response, 405, ErrorObject("METHOD_NOT_ALLOWED", "请使用 POST /messages。"), null)
                    .ConfigureAwait(false);
                return;
            }

            string sessionId = GetQueryValue(request, "sessionId");
            if (string.IsNullOrWhiteSpace(sessionId))
            {
                await WriteJsonAsync(response, 400, ErrorObject("INVALID_PARAMS", "缺少 sessionId 查询参数。"), null)
                    .ConfigureAwait(false);
                return;
            }
            SseSession session;
            if (!_sseSessions.TryGetValue(sessionId, out session))
            {
                await WriteJsonAsync(response, 404, ErrorObject("SESSION_NOT_FOUND", "未知 SSE sessionId：" + sessionId), null)
                    .ConfigureAwait(false);
                return;
            }
            if (!string.Equals(session.UserId, userId, StringComparison.OrdinalIgnoreCase))
            {
                await WriteJsonAsync(response, 403, ErrorObject("FORBIDDEN", "该 SSE 会话属于其他用户。"), null)
                    .ConfigureAwait(false);
                return;
            }

            string body = await ReadBodyAsync(request).ConfigureAwait(false);
            await session.Gate.WaitAsync().ConfigureAwait(false);
            string responseText;
            try
            {
                responseText = await session.Handler.HandleLineAsync(body).ConfigureAwait(false);
            }
            finally
            {
                session.Gate.Release();
            }

            if (responseText != null)
            {
                session.Queue.Enqueue(responseText);
                session.Signal.Release();
            }

            response.StatusCode = 202;
            response.ContentLength64 = 0;
            response.Headers["Mcp-Session-Id"] = sessionId;
        }

        // ------------------------------------------------------------------
        // /xlsx、/word、/ppt 工具文档路由（总览 + 单工具页）；MCP 统一走 POST /mcp
        // ------------------------------------------------------------------

        private bool TryResolveRoute(string path, out RouteInfo route, out string rest)
        {
            route = null;
            rest = string.Empty;
            string trimmed = (path ?? string.Empty).Trim('/');
            if (trimmed.Length == 0)
            {
                return false;
            }

            int slash = trimmed.IndexOf('/');
            string head = slash < 0 ? trimmed : trimmed.Substring(0, slash);
            RouteInfo found;
            if (!_routes.TryGetValue(head, out found))
            {
                return false;
            }

            route = found;
            rest = slash < 0 ? string.Empty : trimmed.Substring(slash + 1).Trim('/');
            return true;
        }

        /// <summary>
        /// 文档路由：GET /xlsx 返回总览（?refresh=true 按最新描述符重建），
        /// GET /xlsx/&lt;工具名称&gt; 返回单工具说明页；common 工具在任意路由下可查。
        /// 非 GET 返回 405，未知子路径返回 404；MCP 调用统一走 POST /mcp。
        /// </summary>
        private async Task HandleRouteAsync(
            RouteInfo routeInfo,
            string rest,
            HttpListenerRequest request,
            HttpListenerResponse response)
        {
            if (!string.Equals(request.HttpMethod, "GET", StringComparison.OrdinalIgnoreCase))
            {
                response.Headers["Allow"] = "GET";
                await WriteJsonAsync(response, 405,
                    ErrorObject("METHOD_NOT_ALLOWED", "工具文档路由只支持 GET；MCP 调用请使用 POST /mcp。"),
                    null).ConfigureAwait(false);
                return;
            }

            if (rest.Length == 0)
            {
                await HandleToolMdAsync(request, response, routeInfo).ConfigureAwait(false);
                return;
            }

            ToolDescriptor descriptor = routeInfo.Registry.Find(rest);
            if (descriptor != null)
            {
                await HandleToolDocAsync(request, response, routeInfo.Route, descriptor).ConfigureAwait(false);
                return;
            }

            await WriteJsonAsync(response, 404,
                ErrorObject("NOT_FOUND",
                    "未知路径：/" + routeInfo.Route.Key + "/" + rest +
                    "；工具文档请访问 /" + routeInfo.Route.Key + " 或 /" + routeInfo.Route.Key + "/<工具名称>。"),
                null).ConfigureAwait(false);
        }

        private async Task HandleToolMdAsync(
            HttpListenerRequest request,
            HttpListenerResponse response,
            RouteInfo routeInfo)
        {
            string userId;
            int status;
            if (!Authorize(request, out userId, out status))
            {
                await WriteJsonAsync(response, status, ErrorObject("UNAUTHORIZED", "认证失败。"), null)
                    .ConfigureAwait(false);
                return;
            }

            string path;
            string error;
            bool refresh = GetQueryBool(request, "refresh") || GetQueryBool(request, "regenerate");
            string markdown = ToolDocumentationService.GetMarkdown(
                routeInfo.Route,
                routeInfo.Registry,
                refresh,
                out path,
                out error);
            if (!string.IsNullOrWhiteSpace(error))
            {
                Console.Error.WriteLine(error);
            }

            response.Headers["Content-Disposition"] =
                "inline; filename=\"" + routeInfo.Route.Key + ".md\"";
            await WriteTextAsync(response, 200, markdown, "text/markdown; charset=utf-8")
                .ConfigureAwait(false);
        }

        /// <summary>GET /xlsx/&lt;工具名称&gt;：单工具 Markdown 说明，动态生成、不落盘。</summary>
        private async Task HandleToolDocAsync(
            HttpListenerRequest request,
            HttpListenerResponse response,
            ToolRoute route,
            ToolDescriptor descriptor)
        {
            string userId;
            int status;
            if (!Authorize(request, out userId, out status))
            {
                await WriteJsonAsync(response, status, ErrorObject("UNAUTHORIZED", "认证失败。"), null)
                    .ConfigureAwait(false);
                return;
            }

            string markdown = ToolDocumentationService.BuildToolMarkdown(route, descriptor);
            response.Headers["Content-Disposition"] =
                "inline; filename=\"" + descriptor.ToolName + ".md\"";
            await WriteTextAsync(response, 200, markdown, "text/markdown; charset=utf-8")
                .ConfigureAwait(false);
        }

        // ------------------------------------------------------------------
        // 文件上传 / 下载（全局 /files/*，代理 workspace.root）
        // ------------------------------------------------------------------

        private async Task HandleFileDownloadAsync(
            HttpListenerRequest request,
            HttpListenerResponse response)
        {
            string userId;
            int status;
            if (!Authorize(request, out userId, out status))
            {
                await WriteJsonAsync(response, status, ErrorObject("UNAUTHORIZED", "认证失败。"), null)
                    .ConfigureAwait(false);
                return;
            }
            if (!string.Equals(request.HttpMethod, "GET", StringComparison.OrdinalIgnoreCase))
            {
                response.Headers["Allow"] = "GET";
                await WriteJsonAsync(response, 405,
                    ErrorObject("METHOD_NOT_ALLOWED", "文件下载只支持 GET。"), null).ConfigureAwait(false);
                return;
            }

            string name = GetRequestedFileName(request);
            string fullPath;
            string pathError;
            if (!WorkspacePathResolver.TryResolveWorkspaceFileName(name, out fullPath, out pathError))
            {
                await WriteJsonAsync(response, 400, ErrorObject("INVALID_PARAMS", pathError), null)
                    .ConfigureAwait(false);
                return;
            }

            if (!WorkspacePathResolver.IsAllowedUploadExtension(null, fullPath))
            {
                await WriteJsonAsync(response, 403,
                    ErrorObject("FORBIDDEN",
                        "不允许下载扩展名「" + Path.GetExtension(fullPath) + "」。"),
                    null).ConfigureAwait(false);
                return;
            }
            if (!File.Exists(fullPath))
            {
                await WriteJsonAsync(response, 404,
                    ErrorObject("FILE_NOT_FOUND", "工作目录中不存在文件：" + name),
                    null).ConfigureAwait(false);
                return;
            }

            FileInfo fileInfo = new FileInfo(fullPath);
            response.StatusCode = 200;
            response.ContentType = GetContentType(fullPath);
            response.ContentLength64 = fileInfo.Length;
            response.Headers["Content-Disposition"] =
                "attachment; filename*=UTF-8''" + Uri.EscapeDataString(fileInfo.Name);

            using (FileStream input = new FileStream(
                       fullPath,
                       FileMode.Open,
                       FileAccess.Read,
                       FileShare.Read,
                       81920,
                       true))
            {
                byte[] buffer = new byte[81920];
                int read;
                while ((read = await input.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false)) > 0)
                {
                    await response.OutputStream.WriteAsync(buffer, 0, read).ConfigureAwait(false);
                }
                await response.OutputStream.FlushAsync().ConfigureAwait(false);
            }
        }

        private async Task HandleFileUploadAsync(
            HttpListenerRequest request,
            HttpListenerResponse response)
        {
            string userId;
            int status;
            if (!Authorize(request, out userId, out status))
            {
                await WriteJsonAsync(response, status, ErrorObject("UNAUTHORIZED", "认证失败。"), null)
                    .ConfigureAwait(false);
                return;
            }

            string method = request.HttpMethod ?? string.Empty;
            if (!string.Equals(method, "POST", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(method, "PUT", StringComparison.OrdinalIgnoreCase))
            {
                response.Headers["Allow"] = "POST, PUT";
                await WriteJsonAsync(response, 405,
                    ErrorObject("METHOD_NOT_ALLOWED", "文件上传只支持 POST 或 PUT。"), null).ConfigureAwait(false);
                return;
            }

            string name = GetRequestedFileName(request);
            string fullPath;
            string pathError;
            if (!WorkspacePathResolver.TryResolveWorkspaceFileName(name, out fullPath, out pathError))
            {
                await WriteJsonAsync(response, 400, ErrorObject("INVALID_PARAMS", pathError), null)
                    .ConfigureAwait(false);
                return;
            }

            if (!WorkspacePathResolver.IsAllowedUploadExtension(null, fullPath))
            {
                await WriteJsonAsync(response, 403,
                    ErrorObject("FORBIDDEN",
                        "不允许上传扩展名「" + Path.GetExtension(fullPath) + "」。"),
                    null).ConfigureAwait(false);
                return;
            }

            bool overwrite = GetQueryBool(request, "overwrite") || GetQueryBool(request, "confirm");
            if (File.Exists(fullPath) && !overwrite)
            {
                await WriteJsonAsync(response, 409,
                    ErrorObject("FILE_EXISTS",
                        "目标文件已存在：" + name + "。如需覆盖请传 overwrite=true。"),
                    null).ConfigureAwait(false);
                return;
            }

            long maxBytes = Math.Max(1024L, _config.WorkspaceMaxUploadBytes);
            if (request.ContentLength64 > maxBytes)
            {
                await WriteJsonAsync(response, 413,
                    ErrorObject("PAYLOAD_TOO_LARGE",
                        "上传超过上限 " + maxBytes.ToString() + " 字节。"),
                    null).ConfigureAwait(false);
                return;
            }

            string directory = Path.GetDirectoryName(fullPath);
            string tempPath = Path.Combine(
                directory ?? WorkspacePathResolver.WorkspaceRoot,
                "." + Path.GetFileName(fullPath) + "." + Guid.NewGuid().ToString("N") + ".upload");
            long totalBytes = 0;
            try
            {
                using (FileStream output = new FileStream(
                           tempPath,
                           FileMode.CreateNew,
                           FileAccess.Write,
                           FileShare.None,
                           81920,
                           true))
                {
                    byte[] buffer = new byte[81920];
                    int read;
                    while ((read = await request.InputStream.ReadAsync(buffer, 0, buffer.Length)
                               .ConfigureAwait(false)) > 0)
                    {
                        totalBytes += read;
                        if (totalBytes > maxBytes)
                        {
                            throw new HttpError(413,
                                "上传超过上限 " + maxBytes.ToString() + " 字节。");
                        }
                        await output.WriteAsync(buffer, 0, read).ConfigureAwait(false);
                    }
                    await output.FlushAsync().ConfigureAwait(false);
                }

                try
                {
                    if (overwrite)
                    {
                        File.Copy(tempPath, fullPath, true);
                    }
                    else
                    {
                        File.Move(tempPath, fullPath);
                    }
                }
                catch (IOException)
                {
                    throw new HttpError(409,
                        "目标文件已存在：" + name + "。如需覆盖请传 overwrite=true。");
                }
            }
            catch (HttpError)
            {
                throw;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("文件上传失败：name=" + name + "；client=" +
                                        GetClientAddress(request) + "；transport=HTTP" +
                                        Environment.NewLine + ex);
                throw new HttpError(500, "文件上传失败：" + ex.Message);
            }
            finally
            {
                try
                {
                    if (File.Exists(tempPath))
                    {
                        File.Delete(tempPath);
                    }
                }
                catch (Exception cleanupError)
                {
                    Console.Error.WriteLine("清理上传临时文件失败：" + tempPath + "；" + cleanupError.Message);
                }
            }

            JObject payload = new JObject();
            payload["ok"] = true;
            payload["name"] = Path.GetFileName(fullPath);
            payload["path"] = fullPath;
            payload["size"] = totalBytes;
            await WriteJsonAsync(response, 200, payload, null).ConfigureAwait(false);
        }

        private static string GetRequestedFileName(HttpListenerRequest request)
        {
            string name = GetQueryValue(request, "name");
            if (string.IsNullOrWhiteSpace(name))
            {
                name = GetQueryValue(request, "file");
            }
            if (!string.IsNullOrWhiteSpace(name))
            {
                return name.Trim();
            }

            string disposition = request.Headers["Content-Disposition"];
            if (string.IsNullOrWhiteSpace(disposition))
            {
                return string.Empty;
            }

            string value = ExtractDispositionFileName(disposition);
            if (value.Length == 0)
            {
                return string.Empty;
            }
            // 头部文件名不做表单式解码：'+' 必须保持为普通字符。
            return DecodeQueryComponent(value, false).Trim();
        }

        /// <summary>
        /// 读取查询参数。HttpListener 在中文系统上会按系统 ANSI 解码请求行/百分号字节，
        /// UTF-8 中文因此变成乱码；这里从 RawUrl 取原始查询串并按 UTF-8 重新解码。
        /// </summary>
        private static string GetQueryValue(HttpListenerRequest request, string name)
        {
            string rawQuery = GetRawQuery(request);
            if (rawQuery.Length > 0)
            {
                string[] pairs = rawQuery.Split('&');
                foreach (string pair in pairs)
                {
                    if (pair.Length == 0)
                    {
                        continue;
                    }
                    int equals = pair.IndexOf('=');
                    string rawKey = equals < 0 ? pair : pair.Substring(0, equals);
                    string rawValue = equals < 0 ? string.Empty : pair.Substring(equals + 1);
                    string key = DecodeQueryComponent(rawKey);
                    if (string.Equals(key, name, StringComparison.OrdinalIgnoreCase))
                    {
                        return DecodeQueryComponent(rawValue);
                    }
                }
            }

            // 极端情况下 RawUrl 不可用，退回 HttpListener 的解析结果并做乱码修复。
            string fallback = request.QueryString[name];
            return fallback == null ? string.Empty : DecodeQueryComponent(fallback);
        }

        private static string GetRawQuery(HttpListenerRequest request)
        {
            string rawUrl = request.RawUrl ?? string.Empty;
            int question = rawUrl.IndexOf('?');
            if (question < 0 || question == rawUrl.Length - 1)
            {
                return string.Empty;
            }
            return rawUrl.Substring(question + 1);
        }

        /// <summary>
        /// 查询串组件解码：处理 %XX（按 UTF-8）、'+' 与已按系统 ANSI 解码的原始 UTF-8 字节。
        /// UTF-8 解码失败时回退系统 ANSI，兼容 GBK 客户端。
        /// </summary>
        private static string DecodeQueryComponent(string value)
        {
            return DecodeQueryComponent(value, true);
        }

        private static string DecodeQueryComponent(string value, bool plusAsSpace)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            string prepared = RepairMojibake(value);
            List<byte> bytes = new List<byte>(prepared.Length);
            for (int index = 0; index < prepared.Length; index++)
            {
                char character = prepared[index];
                if (character == '%' && index + 2 < prepared.Length &&
                    IsHex(prepared[index + 1]) && IsHex(prepared[index + 2]))
                {
                    bytes.Add((byte)((HexValue(prepared[index + 1]) << 4) | HexValue(prepared[index + 2])));
                    index += 2;
                    continue;
                }
                if (character == '+' && plusAsSpace)
                {
                    bytes.Add(0x20);
                    continue;
                }

                int length = char.IsSurrogatePair(prepared, index) ? 2 : 1;
                bytes.AddRange(Encoding.UTF8.GetBytes(prepared.Substring(index, length)));
                if (length == 2)
                {
                    index++;
                }
            }

            try
            {
                return new UTF8Encoding(false, true).GetString(bytes.ToArray());
            }
            catch (DecoderFallbackException)
            {
                return Encoding.Default.GetString(bytes.ToArray());
            }
        }

        /// <summary>
        /// 修复“UTF-8 字节被按系统 ANSI 解码”产生的乱码；仅在还原结果仍是有效非 ASCII
        /// 文本时采用，避免误改正常中文（例如英文系统代码页无法表示中文时）。
        /// </summary>
        private static string RepairMojibake(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return value ?? string.Empty;
            }
            if (!HasNonAscii(value))
            {
                return value;
            }

            try
            {
                byte[] bytes = Encoding.Default.GetBytes(value);
                string repaired = new UTF8Encoding(false, true).GetString(bytes);
                if (!string.Equals(repaired, value, StringComparison.Ordinal) && HasNonAscii(repaired))
                {
                    return repaired;
                }
            }
            catch (Exception)
            {
                // 不是 UTF-8 字节序列：保持原样，后续按系统 ANSI 处理。
            }
            return value;
        }

        private static bool HasNonAscii(string value)
        {
            for (int index = 0; index < value.Length; index++)
            {
                if (value[index] > 0x7F)
                {
                    return true;
                }
            }
            return false;
        }

        private static bool IsHex(char character)
        {
            return (character >= '0' && character <= '9')
                   || (character >= 'a' && character <= 'f')
                   || (character >= 'A' && character <= 'F');
        }

        private static int HexValue(char character)
        {
            if (character >= '0' && character <= '9')
            {
                return character - '0';
            }
            if (character >= 'a' && character <= 'f')
            {
                return character - 'a' + 10;
            }
            return character - 'A' + 10;
        }

        private static string ExtractDispositionFileName(string disposition)
        {
            int extendedIndex = disposition.IndexOf("filename*=", StringComparison.OrdinalIgnoreCase);
            if (extendedIndex >= 0)
            {
                string value = disposition.Substring(extendedIndex + "filename*=".Length).Trim();
                int semicolon = value.IndexOf(';');
                if (semicolon >= 0)
                {
                    value = value.Substring(0, semicolon);
                }
                value = value.Trim().Trim('"');
                // RFC 5987：charset'language'value，例如 UTF-8''%E4%B8%AD 或 UTF-8'en'%E4%B8%AD。
                int firstQuote = value.IndexOf('\'');
                int secondQuote = firstQuote >= 0 ? value.IndexOf('\'', firstQuote + 1) : -1;
                if (firstQuote >= 0 && secondQuote > firstQuote)
                {
                    value = value.Substring(secondQuote + 1);
                }
                return value;
            }

            int plainIndex = disposition.IndexOf("filename=", StringComparison.OrdinalIgnoreCase);
            if (plainIndex >= 0)
            {
                string value = disposition.Substring(plainIndex + "filename=".Length).Trim();
                int semicolon = value.IndexOf(';');
                if (semicolon >= 0)
                {
                    value = value.Substring(0, semicolon);
                }
                return value.Trim().Trim('"');
            }
            return string.Empty;
        }

        private static bool GetQueryBool(HttpListenerRequest request, string name)
        {
            string value = GetQueryValue(request, name);
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }
            string normalized = value.Trim().ToLowerInvariant();
            return normalized == "1" || normalized == "true" || normalized == "yes" || normalized == "on";
        }

        private static string GetContentType(string filePath)
        {
            string extension = (Path.GetExtension(filePath) ?? string.Empty).ToLowerInvariant();
            switch (extension)
            {
                case ".xlsx":
                    return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                case ".xlsm":
                    return "application/vnd.ms-excel.sheet.macroEnabled.12";
                case ".xlsb":
                    return "application/vnd.ms-excel.sheet.binary.macroEnabled.12";
                case ".xls":
                case ".csv":
                    return "application/vnd.ms-excel";
                case ".docx":
                    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                case ".docm":
                    return "application/vnd.ms-word.document.macroEnabled.12";
                case ".doc":
                    return "application/msword";
                case ".pptx":
                    return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
                case ".pptm":
                    return "application/vnd.ms-powerpoint.presentation.macroEnabled.12";
                case ".pdf":
                    return "application/pdf";
                case ".json":
                    return "application/json; charset=utf-8";
                case ".xml":
                    return "application/xml; charset=utf-8";
                case ".txt":
                case ".md":
                case ".rtf":
                    return "text/plain; charset=utf-8";
                case ".png":
                    return "image/png";
                case ".jpg":
                case ".jpeg":
                    return "image/jpeg";
                case ".gif":
                    return "image/gif";
                case ".bmp":
                    return "image/bmp";
                case ".tif":
                case ".tiff":
                    return "image/tiff";
                case ".emf":
                case ".wmf":
                    return "image/x-emf";
                default:
                    return "application/octet-stream";
            }
        }

        private static string GetClientAddress(HttpListenerRequest request)
        {
            if (request == null || request.RemoteEndPoint == null)
            {
                return "unknown";
            }
            try
            {
                return request.RemoteEndPoint.Address.ToString();
            }
            catch
            {
                return "unknown";
            }
        }

        private static async Task WriteSseEventAsync(
            HttpListenerResponse response,
            string eventName,
            string data)
        {
            string frame = "event: " + eventName + "\ndata: " + (data ?? string.Empty) + "\n\n";
            byte[] bytes = Encoding.UTF8.GetBytes(frame);
            await response.OutputStream.WriteAsync(bytes, 0, bytes.Length).ConfigureAwait(false);
            await response.OutputStream.FlushAsync().ConfigureAwait(false);
        }

        private static async Task WriteSseCommentAsync(HttpListenerResponse response, string comment)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(": " + (comment ?? "keep-alive") + "\n\n");
            await response.OutputStream.WriteAsync(bytes, 0, bytes.Length).ConfigureAwait(false);
            await response.OutputStream.FlushAsync().ConfigureAwait(false);
        }

        /// <summary>
        /// POST /mcp（唯一 MCP 端点）：处理 JSON-RPC；无 Mcp-Session-Id 时创建会话，
        /// DELETE /mcp 删除会话；本 server 不提供服务器主动 SSE，GET 返回 405。
        /// </summary>
        private async Task HandleMcpAsync(
            HttpListenerRequest request,
            HttpListenerResponse response)
        {
            string userId;
            int status;
            if (!Authorize(request, out userId, out status))
            {
                await WriteJsonAsync(response, status, ErrorObject("UNAUTHORIZED", "认证失败。"), null)
                    .ConfigureAwait(false);
                return;
            }

            string clientAddress = GetClientAddress(request);
            bool sse = request.AcceptTypes != null && Array.Exists(
                request.AcceptTypes,
                delegate(string value)
                {
                    return value != null &&
                           value.IndexOf("text/event-stream", StringComparison.OrdinalIgnoreCase) >= 0;
                });
            string transportKind = sse ? "SSE" : "HTTP";

            string method = request.HttpMethod ?? string.Empty;
            if (string.Equals(method, "OPTIONS", StringComparison.OrdinalIgnoreCase))
            {
                response.StatusCode = 204;
                response.Headers["Allow"] = "POST, DELETE, OPTIONS";
                response.Headers["Access-Control-Allow-Headers"] = "Authorization, Content-Type, Mcp-Session-Id, MCP-Protocol-Version";
                response.Headers["Access-Control-Allow-Methods"] = "POST, DELETE, OPTIONS";
                return;
            }
            if (string.Equals(method, "DELETE", StringComparison.OrdinalIgnoreCase))
            {
                string deleteSessionId = request.Headers["Mcp-Session-Id"];
                Session deleted;
                if (!string.IsNullOrWhiteSpace(deleteSessionId))
                {
                    _sessions.TryRemove(deleteSessionId, out deleted);
                }
                await WriteJsonAsync(response, 200, new JObject { ["ok"] = true }, deleteSessionId)
                    .ConfigureAwait(false);
                return;
            }
            if (string.Equals(method, "GET", StringComparison.OrdinalIgnoreCase))
            {
                response.Headers["Allow"] = "POST, DELETE, OPTIONS";
                await WriteJsonAsync(response, 405,
                    ErrorObject("METHOD_NOT_ALLOWED", "当前 server 不提供服务器主动 SSE；请使用 POST /mcp。"),
                    null).ConfigureAwait(false);
                return;
            }
            if (!string.Equals(method, "POST", StringComparison.OrdinalIgnoreCase))
            {
                response.Headers["Allow"] = "POST, DELETE, OPTIONS";
                await WriteJsonAsync(response, 405, ErrorObject("METHOD_NOT_ALLOWED", "不支持的方法。"), null)
                    .ConfigureAwait(false);
                return;
            }

            string body = await ReadBodyAsync(request).ConfigureAwait(false);
            string sessionId = request.Headers["Mcp-Session-Id"];
            bool isNewSession = false;
            Session session;
            if (string.IsNullOrWhiteSpace(sessionId))
            {
                sessionId = Guid.NewGuid().ToString("N");
                session = CreateSession(userId, clientAddress, transportKind);
                _sessions[sessionId] = session;
                isNewSession = true;
            }
            else if (!_sessions.TryGetValue(sessionId, out session))
            {
                await WriteJsonAsync(response, 404,
                    ErrorObject("SESSION_NOT_FOUND", "未知 Mcp-Session-Id：" + sessionId),
                    null).ConfigureAwait(false);
                return;
            }

            session.LastSeenUtc = DateTime.UtcNow;
            await session.Gate.WaitAsync().ConfigureAwait(false);
            string responseText;
            try
            {
                responseText = await session.Handler.HandleLineAsync(body).ConfigureAwait(false);
            }
            finally
            {
                session.Gate.Release();
            }

            if (responseText == null)
            {
                response.StatusCode = 202;
                response.Headers["Mcp-Session-Id"] = sessionId;
                return;
            }

            if (sse)
            {
                await WriteSseAsync(response, responseText, sessionId).ConfigureAwait(false);
            }
            else
            {
                await WriteJsonRawAsync(response, 200, responseText, sessionId).ConfigureAwait(false);
            }

            if (session.Handler.ShutdownRequested)
            {
                Session removed;
                _sessions.TryRemove(sessionId, out removed);
            }
            if (isNewSession && response.StatusCode >= 400)
            {
                Session removed;
                _sessions.TryRemove(sessionId, out removed);
            }
        }

        private Session CreateSession(
            string userId,
            string clientAddress,
            string transportKind)
        {
            Session session = new Session();
            session.UserId = string.IsNullOrWhiteSpace(userId) ? "anonymous" : userId;
            session.ClientAddress = string.IsNullOrWhiteSpace(clientAddress) ? "unknown" : clientAddress.Trim();
            session.TransportKind = string.IsNullOrWhiteSpace(transportKind) ? "HTTP" : transportKind.Trim();
            session.Handler = new McpProtocolHandler(
                _registry,
                _discovery,
                session.UserId,
                session.ClientAddress,
                session.TransportKind);
            session.Gate = new SemaphoreSlim(_maxInflightPerSession, _maxInflightPerSession);
            session.LastSeenUtc = DateTime.UtcNow;
            return session;
        }

        private async Task<string> ReadBodyAsync(HttpListenerRequest request)
        {
            if (request.ContentLength64 > _maxBodyBytes)
            {
                throw new HttpError(413, "请求体超过上限 " + _maxBodyBytes + " 字节。");
            }

            string body;
            using (StreamReader reader = new StreamReader(
                       request.InputStream,
                       new UTF8Encoding(false),
                       false,
                       4096,
                       true))
            {
                body = await reader.ReadToEndAsync().ConfigureAwait(false);
            }

            if (Encoding.UTF8.GetByteCount(body) > _maxBodyBytes)
            {
                throw new HttpError(413, "请求体超过上限 " + _maxBodyBytes + " 字节。");
            }
            return body;
        }

        private bool Authorize(HttpListenerRequest request, out string userId, out int statusCode)
        {
            userId = "anonymous";
            statusCode = 200;

            if (_tokens == null || _tokens.Count == 0)
            {
                return true;
            }

            string authorization = request.Headers["Authorization"];
            if (string.IsNullOrWhiteSpace(authorization)
                || !authorization.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
            {
                statusCode = 401;
                return false;
            }
            string provided = authorization.Substring("Bearer ".Length).Trim();
            foreach (System.Collections.Generic.KeyValuePair<string, string> pair in _tokens)
            {
                if (FixedTimeEquals(provided, pair.Key))
                {
                    userId = string.IsNullOrWhiteSpace(pair.Value) ? "user" : pair.Value;
                    return true;
                }
            }
            statusCode = 403;
            return false;
        }

        private static bool FixedTimeEquals(string left, string right)
        {
            if (left == null || right == null || left.Length != right.Length)
            {
                return false;
            }
            int diff = 0;
            for (int index = 0; index < left.Length; index++)
            {
                diff |= left[index] ^ right[index];
            }
            return diff == 0;
        }

        private static JObject ErrorObject(string code, string message)
        {
            JObject error = new JObject();
            error["errorCode"] = code;
            error["message"] = message;
            return error;
        }

        private async Task WriteJsonAsync(
            HttpListenerResponse response,
            int statusCode,
            JObject payload,
            string sessionId)
        {
            await WriteJsonRawAsync(response, statusCode, payload.ToString(Formatting.None), sessionId)
                .ConfigureAwait(false);
        }

        private async Task WriteJsonRawAsync(
            HttpListenerResponse response,
            int statusCode,
            string json,
            string sessionId)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(json ?? string.Empty);
            response.StatusCode = statusCode;
            response.ContentType = "application/json; charset=utf-8";
            response.ContentEncoding = new UTF8Encoding(false);
            response.ContentLength64 = bytes.Length;
            if (!string.IsNullOrEmpty(sessionId))
            {
                response.Headers["Mcp-Session-Id"] = sessionId;
            }
            await response.OutputStream.WriteAsync(bytes, 0, bytes.Length).ConfigureAwait(false);
            await response.OutputStream.FlushAsync().ConfigureAwait(false);
        }

        private async Task WriteTextAsync(
            HttpListenerResponse response,
            int statusCode,
            string text,
            string contentType)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(text ?? string.Empty);
            response.StatusCode = statusCode;
            response.ContentType = contentType;
            response.ContentLength64 = bytes.Length;
            await response.OutputStream.WriteAsync(bytes, 0, bytes.Length).ConfigureAwait(false);
            await response.OutputStream.FlushAsync().ConfigureAwait(false);
        }

        private async Task WriteSseAsync(
            HttpListenerResponse response,
            string json,
            string sessionId)
        {
            response.StatusCode = 200;
            response.ContentType = "text/event-stream; charset=utf-8";
            response.ContentEncoding = new UTF8Encoding(false);
            response.SendChunked = true;
            if (!string.IsNullOrEmpty(sessionId))
            {
                response.Headers["Mcp-Session-Id"] = sessionId;
            }
            byte[] bytes = Encoding.UTF8.GetBytes("event: message\ndata: " + json + "\n\n");
            await response.OutputStream.WriteAsync(bytes, 0, bytes.Length).ConfigureAwait(false);
            await response.OutputStream.FlushAsync().ConfigureAwait(false);
        }

        public void Stop()
        {
            if (_stopping)
            {
                return;
            }
            _stopping = true;
            try
            {
                if (_listener != null)
                {
                    _listener.Stop();
                    _listener.Close();
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[http] 停止监听失败：" + ex.Message);
            }

            foreach (System.Collections.Generic.KeyValuePair<string, SseSession> pair in _sseSessions)
            {
                try
                {
                    pair.Value.Cancellation.Cancel();
                    pair.Value.Response.Close();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[http] 关闭 SSE 会话失败：" + ex.Message);
                }
            }
            _sseSessions.Clear();

            foreach (System.Collections.Generic.KeyValuePair<string, Session> pair in _sessions)
            {
                try
                {
                    pair.Value.Gate.Dispose();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[http] 释放会话锁失败：" + ex.Message);
                }
            }
            _sessions.Clear();
            Task task = _acceptTask;
            if (task != null)
            {
                try
                {
                    task.Wait(TimeSpan.FromSeconds(5));
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[http] 等待接受循环退出失败：" + ex.Message);
                }
            }
        }

        public void Dispose()
        {
            Stop();
        }
    }
}
