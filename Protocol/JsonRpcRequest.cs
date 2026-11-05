using Newtonsoft.Json.Linq;

namespace AutoOfficeMCP.Protocol
{
    /// <summary>
    /// JSON-RPC 2.0 请求/通知的最小模型。
    /// MCP 的 params 始终是对象；这里只做协议层基础校验，业务参数由工具注册表与执行器校验。
    /// </summary>
    internal sealed class JsonRpcRequest
    {
        /// <summary>请求 id；通知没有 id（HasId=false）。</summary>
        public JToken Id { get; private set; }

        /// <summary>原始报文中是否出现 id 字段（显式 null 也算出现）。</summary>
        public bool HasId { get; private set; }

        /// <summary>方法名，例如 initialize / tools/list / tools/call。</summary>
        public string Method { get; private set; }

        /// <summary>params 对象；未提供时为 null。</summary>
        public JObject Params { get; private set; }

        /// <summary>没有 id 的报文按 JSON-RPC 通知处理，不产生响应。</summary>
        public bool IsNotification
        {
            get { return !HasId; }
        }

        /// <summary>
        /// 解析并校验 JSON-RPC 基本形状。失败时通过 error 返回可回显的 id 与错误码。
        /// </summary>
        public static bool TryParse(JObject root, out JsonRpcRequest request, out JsonRpcError error)
        {
            request = null;
            error = null;

            if (root == null)
            {
                error = JsonRpcError.Create(JsonRpcErrorCodes.InvalidRequest, "Invalid Request: 报文必须是 JSON 对象。", null, null);
                return false;
            }

            JToken rawId = root["id"];
            bool hasId = root.Property("id") != null;
            JToken safeId = NormalizeId(rawId);
            if (hasId && safeId == null)
            {
                error = JsonRpcError.Create(JsonRpcErrorCodes.InvalidRequest, "Invalid Request: id 必须是字符串、数字或 null。", null, null);
                return false;
            }

            JToken version = root["jsonrpc"];
            if (version == null
                || version.Type != JTokenType.String
                || !string.Equals(version.Value<string>(), "2.0", System.StringComparison.Ordinal))
            {
                error = JsonRpcError.Create(JsonRpcErrorCodes.InvalidRequest, "Invalid Request: jsonrpc 必须为 \"2.0\"。", safeId, null);
                return false;
            }

            JToken method = root["method"];
            if (method == null
                || method.Type != JTokenType.String
                || string.IsNullOrWhiteSpace(method.Value<string>()))
            {
                error = JsonRpcError.Create(JsonRpcErrorCodes.InvalidRequest, "Invalid Request: method 必须是非空字符串。", safeId, null);
                return false;
            }

            JObject parameters = null;
            JToken rawParams = root["params"];
            if (rawParams != null && rawParams.Type != JTokenType.Null)
            {
                parameters = rawParams as JObject;
                if (parameters == null)
                {
                    error = JsonRpcError.Create(JsonRpcErrorCodes.InvalidParams, "Invalid params: params 必须是 JSON 对象。", safeId, null);
                    return false;
                }
            }

            request = new JsonRpcRequest();
            request.Id = safeId;
            request.HasId = hasId;
            request.Method = method.Value<string>();
            request.Params = parameters;
            return true;
        }

        /// <summary>只接受 JSON-RPC 允许的 string/number/null id，其余类型视为无效请求。</summary>
        private static JToken NormalizeId(JToken token)
        {
            if (token == null)
            {
                return null;
            }

            switch (token.Type)
            {
                case JTokenType.String:
                case JTokenType.Integer:
                case JTokenType.Float:
                    return token;
                case JTokenType.Null:
                    return JValue.CreateNull();
                default:
                    return null;
            }
        }
    }
}
