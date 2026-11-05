using Newtonsoft.Json.Linq;

namespace AutoOfficeMCP.Protocol
{
    /// <summary>标准 JSON-RPC 2.0 错误码。</summary>
    public static class JsonRpcErrorCodes
    {
        public const int ParseError = -32700;
        public const int InvalidRequest = -32600;
        public const int MethodNotFound = -32601;
        public const int InvalidParams = -32602;
        public const int InternalError = -32603;
    }

    /// <summary>解析/分发阶段的错误载体；Id 仅用于把请求 id 回显到错误帧。</summary>
    public sealed class JsonRpcError
    {
        public int Code { get; set; }

        public string Message { get; set; }

        public JToken Data { get; set; }

        public JToken Id { get; set; }

        public static JsonRpcError Create(int code, string message, JToken id, JToken data)
        {
            JsonRpcError error = new JsonRpcError();
            error.Code = code;
            error.Message = message ?? string.Empty;
            error.Id = id;
            error.Data = data;
            return error;
        }
    }

    /// <summary>JSON-RPC 2.0 响应帧工厂。所有响应都以单行紧凑 JSON 输出。</summary>
    public static class JsonRpcResponse
    {
        public static JObject CreateResult(JToken id, JObject result)
        {
            JObject response = new JObject();
            response["jsonrpc"] = "2.0";
            response["id"] = id == null ? JValue.CreateNull() : id.DeepClone();
            response["result"] = result ?? new JObject();
            return response;
        }

        public static JObject CreateError(JToken id, int code, string message, JToken data)
        {
            JObject error = new JObject();
            error["code"] = code;
            error["message"] = message ?? "Unknown error";
            if (data != null)
            {
                error["data"] = data.DeepClone();
            }

            JObject response = new JObject();
            response["jsonrpc"] = "2.0";
            response["id"] = id == null ? JValue.CreateNull() : id.DeepClone();
            response["error"] = error;
            return response;
        }

        public static JObject CreateError(JsonRpcError error)
        {
            if (error == null)
            {
                return CreateError(null, JsonRpcErrorCodes.InternalError, "Internal error", null);
            }
            return CreateError(error.Id, error.Code, error.Message, error.Data);
        }
    }
}
