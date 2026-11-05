using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Office.Contracts;

namespace Office
{
    /// <summary>
    /// 统一观察形状：{kind, summary, changed, targetRefs[], undoUnits[]}。
    /// 每次写操作都必须说明"改了什么、是否真的改了、改到哪里、哪些节点可用于撤销"。
    /// </summary>
    public sealed class Observation
    {
        public string Kind { get; set; }

        public string Summary { get; set; }

        public bool Changed { get; set; }

        public List<string> TargetRefs { get; private set; }

        public List<string> UndoUnits { get; private set; }

        public Observation()
        {
            Kind = string.Empty;
            Summary = string.Empty;
            TargetRefs = new List<string>();
            UndoUnits = new List<string>();
        }

        public static Observation Create(string kind, string summary, bool changed)
        {
            Observation observation = new Observation();
            observation.Kind = kind ?? string.Empty;
            observation.Summary = summary ?? string.Empty;
            observation.Changed = changed;
            return observation;
        }

        public Observation AddTargetRef(string targetRef)
        {
            if (!string.IsNullOrWhiteSpace(targetRef))
            {
                TargetRefs.Add(targetRef);
            }
            return this;
        }

        public Observation AddTargetRefs(IEnumerable<string> targetRefs)
        {
            if (targetRefs != null)
            {
                foreach (string targetRef in targetRefs)
                {
                    AddTargetRef(targetRef);
                }
            }
            return this;
        }

        public Observation AddUndoUnit(string undoUnitId)
        {
            if (!string.IsNullOrWhiteSpace(undoUnitId))
            {
                UndoUnits.Add(undoUnitId);
            }
            return this;
        }

        /// <summary>序列化为协议层统一观察形状。</summary>
        public JObject ToJObject()
        {
            JObject json = new JObject();
            json["kind"] = Kind ?? string.Empty;
            json["summary"] = Summary ?? string.Empty;
            json["changed"] = Changed;

            JArray targetRefs = new JArray();
            foreach (string item in TargetRefs)
            {
                if (!string.IsNullOrWhiteSpace(item))
                {
                    targetRefs.Add(item);
                }
            }
            json["targetRefs"] = targetRefs;

            JArray undoUnits = new JArray();
            foreach (string item in UndoUnits)
            {
                if (!string.IsNullOrWhiteSpace(item))
                {
                    undoUnits.Add(item);
                }
            }
            json["undoUnits"] = undoUnits;
            return json;
        }
    }

    /// <summary>
    /// 写信任合同落地：Success=true 必须代表操作真实生效；
    /// 部分应用只能返回 PARTIAL_APPLY + Success=false，并附带回滚说明。
    /// </summary>
    public sealed class OperationResult
    {
        /// <summary>操作是否真实成功。</summary>
        public bool Success { get; private set; }

        /// <summary>稳定错误码；成功时为空字符串。</summary>
        public string ErrorCode { get; private set; }

        /// <summary>给用户/LLM 看的中文说明。</summary>
        public string UserMessage { get; private set; }

        /// <summary>结构化数据（Observation 会由 ToToolResult 合并进来）。</summary>
        public JObject Data { get; set; }

        /// <summary>统一观察结果。</summary>
        public Observation Observation { get; set; }

        /// <summary>回滚说明（部分应用时必须说明已完成/未完成/如何人工恢复）。</summary>
        public string RollbackNote { get; set; }

        /// <summary>调用方是否可以调整参数或稍后重试。</summary>
        public bool Recoverable { get; set; }

        /// <summary>
        /// 内部诊断信息（异常链类型/消息/堆栈）。由异常映射点填充，
        /// 最终只进入 stderr 错误日志，不进入 MCP 响应。
        /// </summary>
        public string Diagnostic { get; set; }

        /// <summary>截图/缩略图等二进制内容；成功结果转 ToolResult 时原样传递。</summary>
        public byte[] ImageBytes { get; set; }

        private OperationResult()
        {
            ErrorCode = string.Empty;
            UserMessage = string.Empty;
            RollbackNote = string.Empty;
            Recoverable = true;
            Diagnostic = string.Empty;
        }

        public static OperationResult Succeed(string userMessage, JObject data, Observation observation, string rollbackNote)
        {
            OperationResult result = new OperationResult();
            result.Success = true;
            result.ErrorCode = string.Empty;
            result.UserMessage = userMessage ?? string.Empty;
            result.Data = data;
            result.Observation = observation;
            result.RollbackNote = rollbackNote ?? string.Empty;
            result.Recoverable = true;
            return result;
        }

        public static OperationResult Succeed(string userMessage)
        {
            return Succeed(userMessage, null, null, null);
        }

        public static OperationResult Succeed(string userMessage, JObject data)
        {
            return Succeed(userMessage, data, null, null);
        }

        public static OperationResult Succeed(string userMessage, JObject data, Observation observation)
        {
            return Succeed(userMessage, data, observation, null);
        }

        public static OperationResult Failed(OfficeErrorCode code, string userMessage)
        {
            return Failed(code, userMessage, (string)null);
        }

        public static OperationResult Failed(OfficeErrorCode code, string userMessage, string rollbackNote)
        {
            OperationResult result = CreateFailure(
                OfficeErrorMapper.ToCode(code),
                string.IsNullOrWhiteSpace(userMessage) ? OfficeErrorMapper.GetUserMessage(code) : userMessage,
                OfficeErrorMapper.IsRecoverable(code),
                rollbackNote);
            return result;
        }

        public static OperationResult Failed(OfficeErrorCode code, string userMessage, Observation observation, JObject data, string rollbackNote)
        {
            return Failed(code, userMessage, observation, data, rollbackNote, null);
        }

        public static OperationResult Failed(OfficeErrorCode code, string userMessage, Observation observation, JObject data, string rollbackNote, bool? recoverable)
        {
            OperationResult result = CreateFailure(
                OfficeErrorMapper.ToCode(code),
                string.IsNullOrWhiteSpace(userMessage) ? OfficeErrorMapper.GetUserMessage(code) : userMessage,
                recoverable.HasValue ? recoverable.Value : OfficeErrorMapper.IsRecoverable(code),
                rollbackNote);
            result.Observation = observation;
            result.Data = data;
            return result;
        }

        public static OperationResult Failed(string errorCode, string userMessage)
        {
            return CreateFailure(errorCode, userMessage, true, null);
        }

        public static OperationResult Failed(string errorCode, string userMessage, string rollbackNote)
        {
            return CreateFailure(errorCode, userMessage, true, rollbackNote);
        }

        public static OperationResult Failed(string errorCode, string userMessage, bool recoverable, string rollbackNote)
        {
            return CreateFailure(errorCode, userMessage, recoverable, rollbackNote);
        }

        /// <summary>
        /// 部分应用：Success 必须为 false，错误码固定 PARTIAL_APPLY，禁止调用方自动重放。
        /// succeededCount 是"确实已生效"的步骤数；rollbackNote 必须说明如何人工恢复。
        /// </summary>
        public static OperationResult PartialApplied(int succeededCount, string rollbackNote)
        {
            return PartialApplied(succeededCount, rollbackNote, null, null);
        }

        public static OperationResult PartialApplied(int succeededCount, string rollbackNote, Observation observation, JObject data)
        {
            string note = string.IsNullOrWhiteSpace(rollbackNote)
                ? "已完成的改动不会自动回滚，请对照变更清单人工确认。"
                : rollbackNote.Trim();
            string message = "部分应用：已完成 " + succeededCount.ToString() + " 项，后续步骤未执行；不会自动重放。";
            if (note.Length > 0)
            {
                message = message + " 回滚说明：" + note;
            }

            JObject payload = data == null ? new JObject() : (JObject)data.DeepClone();
            payload["succeededCount"] = succeededCount;

            OperationResult result = CreateFailure(
                OfficeErrorMapper.ToCode(OfficeErrorCode.PARTIAL_APPLY),
                message,
                false,
                note);
            result.Observation = observation;
            result.Data = payload;
            return result;
        }

        /// <summary>预览被拒绝、用户取消等场景：确实没有修改文档。</summary>
        public static OperationResult Cancelled(string userMessage)
        {
            return Failed(OfficeErrorCode.CANCELLED, userMessage);
        }

        /// <summary>
        /// 转换到协议层 ToolResult：Observation 序列化进 Data["observation"]，
        /// RollbackNote 同时写入 Data["rollbackNote"]，失败时保留结构化数据。
        /// </summary>
        public ToolResult ToToolResult()
        {
            JObject data = BuildData();
            if (Success)
            {
                ToolResult success = ToolResult.Ok(UserMessage, data);
                if (ImageBytes != null && ImageBytes.Length > 0)
                {
                    success.ImageBytes = ImageBytes;
                }
                return success;
            }

            ToolResult result = ToolResult.Fail(ErrorCode, UserMessage, Recoverable);
            result.Data = data;
            result.Diagnostic = Diagnostic ?? string.Empty;
            return result;
        }

        private JObject BuildData()
        {
            JObject data = Data == null ? new JObject() : (JObject)Data.DeepClone();
            if (!Success)
            {
                // 负样例/安全拒绝必须在结构化载荷中带稳定错误码，客户端不能只靠中文消息判断。
                data["errorCode"] = ErrorCode ?? string.Empty;
                data["recoverable"] = Recoverable;
            }
            if (Observation != null)
            {
                data["observation"] = Observation.ToJObject();
            }
            if (!string.IsNullOrWhiteSpace(RollbackNote))
            {
                data["rollbackNote"] = RollbackNote;
            }
            if (!data.HasValues)
            {
                return null;
            }
            return data;
        }

        private static OperationResult CreateFailure(string errorCode, string userMessage, bool recoverable, string rollbackNote)
        {
            OperationResult result = new OperationResult();
            result.Success = false;
            result.ErrorCode = string.IsNullOrWhiteSpace(errorCode)
                ? OfficeErrorMapper.ToCode(OfficeErrorCode.INTERNAL_ERROR)
                : errorCode.Trim();

            if (!string.IsNullOrWhiteSpace(userMessage))
            {
                result.UserMessage = userMessage;
            }
            else
            {
                OfficeErrorCode parsedCode;
                result.UserMessage = OfficeErrorMapper.TryParseCode(result.ErrorCode, out parsedCode)
                    ? OfficeErrorMapper.GetUserMessage(parsedCode)
                    : OfficeErrorMapper.GetUserMessage(OfficeErrorCode.INTERNAL_ERROR);
            }

            result.Recoverable = recoverable;
            result.RollbackNote = rollbackNote ?? string.Empty;
            return result;
        }
    }
}
