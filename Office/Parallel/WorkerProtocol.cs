using System;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Office.Parallel
{
    /// <summary>Gateway → worker 的单作业请求（一行 JSON）。</summary>
    public sealed class WorkerJobRequest
    {
        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("jobId")]
        public string JobId { get; set; }

        [JsonProperty("userId")]
        public string UserId { get; set; }

        [JsonProperty("app")]
        public string App { get; set; }

        [JsonProperty("tool")]
        public string Tool { get; set; }

        [JsonProperty("file")]
        public string File { get; set; }

        [JsonProperty("arguments")]
        public JObject Arguments { get; set; }

        [JsonProperty("savePolicy")]
        public string SavePolicy { get; set; }

        [JsonProperty("timeoutSec")]
        public int TimeoutSec { get; set; }

        [JsonProperty("affinityKey")]
        public string AffinityKey { get; set; }

        [JsonProperty("priority")]
        public int Priority { get; set; }

        [JsonProperty("idempotencyKey")]
        public string IdempotencyKey { get; set; }

        public WorkerJobRequest()
        {
            Type = "job";
            JobId = string.Empty;
            UserId = string.Empty;
            App = string.Empty;
            Tool = string.Empty;
            File = string.Empty;
            Arguments = new JObject();
            SavePolicy = "none";
            AffinityKey = string.Empty;
            IdempotencyKey = string.Empty;
        }
    }

    /// <summary>worker → Gateway 的作业结果（一行 JSON）。</summary>
    public sealed class WorkerJobResponse
    {
        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("jobId")]
        public string JobId { get; set; }

        [JsonProperty("success")]
        public bool Success { get; set; }

        [JsonProperty("errorCode")]
        public string ErrorCode { get; set; }

        [JsonProperty("message")]
        public string Message { get; set; }

        [JsonProperty("recoverable")]
        public bool Recoverable { get; set; }

        [JsonProperty("data")]
        public JObject Data { get; set; }

        [JsonProperty("observation")]
        public JObject Observation { get; set; }

        [JsonProperty("rollbackNote")]
        public string RollbackNote { get; set; }

        /// <summary>内部诊断（异常链与堆栈）；只用于 Gateway 写入 stderr，不返回客户端。</summary>
        [JsonProperty("diagnostic")]
        public string Diagnostic { get; set; }

        [JsonProperty("saved")]
        public bool Saved { get; set; }

        [JsonProperty("openedReadOnly")]
        public bool OpenedReadOnly { get; set; }

        [JsonProperty("saveSkippedReason")]
        public string SaveSkippedReason { get; set; }

        [JsonProperty("undoNodesRemoved")]
        public int UndoNodesRemoved { get; set; }

        [JsonProperty("saveVerified")]
        public bool SaveVerified { get; set; }

        [JsonProperty("saveVerificationNote")]
        public string SaveVerificationNote { get; set; }

        [JsonProperty("saveChangedOnDisk")]
        public bool SaveChangedOnDisk { get; set; }

        [JsonProperty("saveBeforeHash")]
        public string SaveBeforeHash { get; set; }

        [JsonProperty("saveAfterHash")]
        public string SaveAfterHash { get; set; }

        [JsonProperty("saveBeforeLength")]
        public long SaveBeforeLength { get; set; }

        [JsonProperty("saveAfterLength")]
        public long SaveAfterLength { get; set; }

        [JsonProperty("saveBeforeLastWriteUtc")]
        public string SaveBeforeLastWriteUtc { get; set; }

        [JsonProperty("saveAfterLastWriteUtc")]
        public string SaveAfterLastWriteUtc { get; set; }

        public WorkerJobResponse()
        {
            Type = "result";
            JobId = string.Empty;
            ErrorCode = string.Empty;
            Message = string.Empty;
            Recoverable = true;
            Diagnostic = string.Empty;
            SaveSkippedReason = string.Empty;
            SaveVerificationNote = string.Empty;
            SaveBeforeHash = string.Empty;
            SaveAfterHash = string.Empty;
            SaveBeforeLastWriteUtc = string.Empty;
            SaveAfterLastWriteUtc = string.Empty;
        }
    }

    /// <summary>ready/heartbeat/shutdown/error 控制消息。</summary>
    public sealed class WorkerControlMessage
    {
        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("jobId")]
        public string JobId { get; set; }

        [JsonProperty("message")]
        public string Message { get; set; }

        [JsonProperty("userId")]
        public string UserId { get; set; }

        [JsonProperty("app")]
        public string App { get; set; }

        [JsonProperty("officePid")]
        public int OfficePid { get; set; }

        [JsonProperty("timestampUtc")]
        public string TimestampUtc { get; set; }

        public WorkerControlMessage()
        {
            Type = string.Empty;
            JobId = string.Empty;
            Message = string.Empty;
            UserId = string.Empty;
            App = string.Empty;
            OfficePid = 0;
            TimestampUtc = string.Empty;
        }
    }

    public static class WorkerProtocol
    {
        public static string Serialize(object value)
        {
            return JsonConvert.SerializeObject(value, Formatting.None);
        }

        public static WorkerJobRequest DeserializeRequest(string line)
        {
            return JsonConvert.DeserializeObject<WorkerJobRequest>(line);
        }

        public static WorkerJobResponse DeserializeResponse(string line)
        {
            return JsonConvert.DeserializeObject<WorkerJobResponse>(line);
        }

        public static WorkerControlMessage DeserializeControl(string line)
        {
            return JsonConvert.DeserializeObject<WorkerControlMessage>(line);
        }

        public static string GetMessageType(string line)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                return string.Empty;
            }
            try
            {
                JObject root = JObject.Parse(line);
                return root.Value<string>("type") ?? string.Empty;
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }
    }
}
