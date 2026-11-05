using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Office.Contracts;

namespace Office.Parallel
{
    /// <summary>
    /// 并行批次执行器（common，appType=common）：
    /// office_batch_files —— 把 {app, tool, file, arguments, savePolicy?, timeoutSec?} 逐项投递到
    /// ParallelDispatcher，再按原始顺序汇总 {results[], succeeded, failed}。
    /// 单项失败/超时不影响其他项；写信任合同要求逐项如实报告错误码与消息。
    /// </summary>
    public sealed class ParallelToolExecutor : IToolExecutor, IAsyncToolExecutor
    {
        public const string BatchFilesToolName = "office_batch_files";
        public const string PoolStatusToolName = "office_pool_status";

        private static readonly string[] ToolNames = new string[]
        {
            BatchFilesToolName,
            PoolStatusToolName
        };

        private static readonly HashSet<string> ToolNameSet =
            new HashSet<string>(ToolNames, StringComparer.Ordinal);

        public string AppType
        {
            get { return "common"; }
        }

        public bool CanExecute(string toolName)
        {
            if (string.IsNullOrWhiteSpace(toolName))
            {
                return false;
            }
            return ToolNameSet.Contains(toolName.Trim());
        }

        public ToolResult Execute(string toolName, JObject arguments)
        {
            if (!CanExecute(toolName))
            {
                return ToolResult.Fail(
                    "INVALID_PARAMS",
                    "并行执行器不支持工具「" + (toolName ?? string.Empty) + "」。",
                    false);
            }

            string normalized = toolName.Trim();
            try
            {
                if (string.Equals(normalized, PoolStatusToolName, StringComparison.Ordinal))
                {
                    return ExecutePoolStatus(arguments);
                }
                return ExecuteBatchFilesAsync(arguments).ConfigureAwait(false).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                return ToolResult.Fail(
                    "INTERNAL_ERROR",
                    "并行执行器发生未捕获异常：" + ex.Message,
                    false);
            }
        }

        /// <summary>T10：异步批处理入口；协议线程 await 本方法，不阻塞读循环。</summary>
        public async Task<ToolResult> ExecuteAsync(string toolName, JObject arguments)
        {
            if (!CanExecute(toolName))
            {
                return ToolResult.Fail(
                    "INVALID_PARAMS",
                    "并行执行器不支持工具「" + (toolName ?? string.Empty) + "」。",
                    false);
            }

            string normalized = toolName.Trim();
            try
            {
                if (string.Equals(normalized, PoolStatusToolName, StringComparison.Ordinal))
                {
                    return ExecutePoolStatus(arguments);
                }
                return await ExecuteBatchFilesAsync(arguments).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                return ToolResult.Fail(
                    "INTERNAL_ERROR",
                    "并行执行器发生未捕获异常：" + ex.Message,
                    false);
            }
        }

        // ------------------------------------------------------------------
        // office_pool_status
        // ------------------------------------------------------------------

        private static ToolResult ExecutePoolStatus(JObject arguments)
        {
            JObject status = ParallelDispatcher.Instance.GetStatus();
            int queueDepth = status.Value<int?>("queueDepth") ?? 0;
            int running = status.Value<int?>("runningCount") ?? 0;
            int fileLocks = status.Value<int?>("fileLockCount") ?? 0;
            int recentErrors = status["recentErrors"] == null ? 0 : ((JArray)status["recentErrors"]).Count;

            string message = "并行调度器状态：队列深度 " +
                             queueDepth.ToString(CultureInfo.InvariantCulture) + "，在跑 " +
                             running.ToString(CultureInfo.InvariantCulture) + "，文件锁条目 " +
                             fileLocks.ToString(CultureInfo.InvariantCulture) + "，最近错误 " +
                             recentErrors.ToString(CultureInfo.InvariantCulture) + " 条。";
            return ToolResult.Ok(message, status);
        }

        // ------------------------------------------------------------------
        // office_batch_files
        // ------------------------------------------------------------------

        private static async Task<ToolResult> ExecuteBatchFilesAsync(JObject arguments)
        {
            if (arguments == null)
            {
                return InvalidParameters("office_batch_files 需要 items 数组。");
            }

            JArray rawItems = arguments["items"] as JArray;
            if (rawItems == null)
            {
                return InvalidParameters(
                    "items 必须是数组，元素形状为 {\"app\":\"excel|word|ppt\",\"tool\":\"...\",\"file\":\"显式路径\",\"arguments\":{...},\"savePolicy\":\"none|save\",\"timeoutSec\":300}。");
            }
            if (rawItems.Count == 0)
            {
                return InvalidParameters("items 不能为空。");
            }

            ParallelDispatcher dispatcher = ParallelDispatcher.Instance;
            int maxQueue = dispatcher.MaxQueue;
            if (rawItems.Count > maxQueue)
            {
                return InvalidParameters(
                    "items 最多 " + maxQueue.ToString(CultureInfo.InvariantCulture) +
                    " 项（等于 parallel.maxQueue），收到 " +
                    rawItems.Count.ToString(CultureInfo.InvariantCulture) + " 项；请拆分批次。");
            }

            List<BatchRequest> requests = new List<BatchRequest>();
            List<Task<OperationResult>> pendingTasks = new List<Task<OperationResult>>();
            int maxTimeoutSec = dispatcher.ItemTimeoutSec;

            for (int index = 0; index < rawItems.Count; index++)
            {
                JObject rawItem = rawItems[index] as JObject;
                BatchRequest request;
                string errorCode;
                string errorMessage;
                if (!TryParseItem(rawItem, index, out request, out errorCode, out errorMessage))
                {
                    request = new BatchRequest();
                    request.Index = index;
                    request.ValidationFailure = OperationResult.Failed(errorCode, errorMessage);
                    requests.Add(request);
                    continue;
                }

                requests.Add(request);
                if (request.WorkItem.TimeoutSec > maxTimeoutSec)
                {
                    maxTimeoutSec = request.WorkItem.TimeoutSec;
                }
                request.Task = dispatcher.Submit(request.WorkItem);
                pendingTasks.Add(request.Task);
            }

            // 每项都有自己的 itemTimeoutSec；这里再加 30 秒安全余量，防止调度器自身缺陷导致永久等待。
            int waitSeconds = maxTimeoutSec + 30;
            if (pendingTasks.Count > 0)
            {
                Task all = Task.WhenAll(pendingTasks);
                await Task.WhenAny(all, Task.Delay(TimeSpan.FromSeconds(waitSeconds))).ConfigureAwait(false);
            }

            JArray results = new JArray();
            int succeeded = 0;
            int failed = 0;
            for (int index = 0; index < requests.Count; index++)
            {
                BatchRequest request = requests[index];
                OperationResult result;
                if (request.ValidationFailure != null)
                {
                    result = request.ValidationFailure;
                }
                else
                {
                    result = ReadTaskResult(request.Task);
                }

                if (result.Success)
                {
                    succeeded++;
                }
                else
                {
                    failed++;
                }
                results.Add(BuildResultRecord(request, result));
            }

            JObject data = new JObject();
            data["requestedCount"] = requests.Count;
            data["results"] = results;
            data["succeeded"] = succeeded;
            data["failed"] = failed;

            Observation observation = Observation.Create(
                "parallel_batch",
                "并行批次完成 " + succeeded.ToString(CultureInfo.InvariantCulture) + " 项，失败 " +
                failed.ToString(CultureInfo.InvariantCulture) + " 项。",
                succeeded > 0);
            for (int index = 0; index < requests.Count; index++)
            {
                BatchRequest request = requests[index];
                if (!string.IsNullOrWhiteSpace(request.File))
                {
                    observation.AddTargetRef(
                        (request.App ?? string.Empty) + ":" + request.File + "#" + (request.Tool ?? string.Empty));
                }
            }

            string rollbackNote =
                "并行批次各项彼此独立；成功项不会自动回滚。savePolicy=save 且 saved=true 的项已写入磁盘；" +
                "savePolicy=none 的项只修改池内一次性实例的内存，实例关闭后改动丢弃。" +
                "失败项不会自动重放，请按 results[].errorCode 与 message 人工处理。";

            if (failed == 0)
            {
                string successMessage = "office_batch_files 已完成 " +
                                        succeeded.ToString(CultureInfo.InvariantCulture) + " 项并行作业。";
                return OperationResult.Succeed(successMessage, data, observation).ToToolResult();
            }

            string partialMessage = "office_batch_files 完成 " +
                                    succeeded.ToString(CultureInfo.InvariantCulture) + " 项，失败 " +
                                    failed.ToString(CultureInfo.InvariantCulture) +
                                    " 项；各项独立执行，失败项未自动重放。";
            OperationResult partial = OperationResult.Failed(
                OfficeErrorCode.PARTIAL_APPLY,
                partialMessage,
                observation,
                data,
                rollbackNote,
                false);
            return partial.ToToolResult();
        }

        private static bool TryParseItem(
            JObject rawItem,
            int index,
            out BatchRequest request,
            out string errorCode,
            out string errorMessage)
        {
            request = new BatchRequest();
            request.Index = index;
            errorCode = OfficeErrorMapper.ToCode(OfficeErrorCode.INVALID_PARAMS);
            errorMessage = string.Empty;

            if (rawItem == null)
            {
                errorMessage = "items[" + index.ToString(CultureInfo.InvariantCulture) + "] 必须是对象。";
                return false;
            }

            string app = ReadString(rawItem, "app");
            string normalizedApp = ParallelToolCatalog.NormalizeApp(app);
            if (normalizedApp.Length == 0)
            {
                errorMessage = "items[" + index.ToString(CultureInfo.InvariantCulture) +
                               "].app 必须是 excel、word 或 ppt。";
                return false;
            }
            request.App = normalizedApp;

            string tool = ReadString(rawItem, "tool").Trim();
            request.Tool = tool;
            if (tool.Length == 0)
            {
                errorMessage = "items[" + index.ToString(CultureInfo.InvariantCulture) +
                               "].tool 必须是非空工具名。";
                return false;
            }
            if (string.Equals(tool, BatchFilesToolName, StringComparison.OrdinalIgnoreCase))
            {
                errorMessage = "items[" + index.ToString(CultureInfo.InvariantCulture) +
                               "].tool 不允许递归调用 " + BatchFilesToolName + "。";
                return false;
            }
            if (!ParallelToolCatalog.IsToolForApp(normalizedApp, tool))
            {
                errorMessage = "items[" + index.ToString(CultureInfo.InvariantCulture) + "].tool「" + tool +
                               "」与 app=" + normalizedApp + " 不匹配（工具名前缀必须为 " +
                               normalizedApp + "_）。";
                return false;
            }
            if (string.Equals(tool, "ppt_execute_vba", StringComparison.OrdinalIgnoreCase))
            {
                errorMessage = "items[" + index.ToString(CultureInfo.InvariantCulture) +
                               "].tool 不允许使用 ppt_execute_vba：PowerPoint COM 无法超时中断已执行的 VBA，" +
                               "会挂住池内 PowerPoint 实例。请在主会话直接调用并自行确认风险。";
                return false;
            }
            if (ParallelDispatcher.ToolExists != null)
            {
                bool exists;
                try
                {
                    exists = ParallelDispatcher.ToolExists(tool);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[parallel] 工具注册校验失败：" + tool + "；" + ex.Message);
                    exists = false;
                }
                if (!exists)
                {
                    errorMessage = "items[" + index.ToString(CultureInfo.InvariantCulture) + "].tool「" + tool +
                                   "」未注册或已在 config.json 中禁用。";
                    return false;
                }
            }

            string file = ReadString(rawItem, "file");
            request.File = file.Trim();
            request.ReadOnly = ParallelToolCatalog.IsReadOnlyTool(normalizedApp, tool);
            if (string.IsNullOrWhiteSpace(file))
            {
                errorMessage = "items[" + index.ToString(CultureInfo.InvariantCulture) +
                               "].file 必须是显式文件路径（不支持“当前活动文档”）。";
                return false;
            }

            JObject itemArguments = null;
            JToken argumentsToken = rawItem["arguments"];
            if (argumentsToken != null && argumentsToken.Type != JTokenType.Null)
            {
                itemArguments = argumentsToken as JObject;
                if (itemArguments == null)
                {
                    errorMessage = "items[" + index.ToString(CultureInfo.InvariantCulture) +
                                   "].arguments 必须是 JSON 对象。";
                    return false;
                }
            }

            string savePolicyText = ReadString(rawItem, "savePolicy");
            request.SavePolicy = savePolicyText;
            string savePolicy = ParallelToolCatalog.NormalizeSavePolicy(savePolicyText);
            if (savePolicy == null)
            {
                errorMessage = "items[" + index.ToString(CultureInfo.InvariantCulture) +
                               "].savePolicy 只允许 none 或 save。";
                return false;
            }

            int timeoutSec = 0;
            JToken timeoutToken = rawItem["timeoutSec"];
            if (timeoutToken != null && timeoutToken.Type != JTokenType.Null)
            {
                if (timeoutToken.Type != JTokenType.Integer)
                {
                    errorMessage = "items[" + index.ToString(CultureInfo.InvariantCulture) +
                                   "].timeoutSec 必须是正整数秒。";
                    return false;
                }
                try
                {
                    timeoutSec = timeoutToken.Value<int>();
                }
                catch
                {
                    errorMessage = "items[" + index.ToString(CultureInfo.InvariantCulture) +
                                   "].timeoutSec 超出整数范围。";
                    return false;
                }
                if (timeoutSec <= 0)
                {
                    errorMessage = "items[" + index.ToString(CultureInfo.InvariantCulture) +
                                   "].timeoutSec 必须大于 0。";
                    return false;
                }
            }

            int priority = 0;
            JToken priorityToken = rawItem["priority"];
            if (priorityToken != null && priorityToken.Type != JTokenType.Null)
            {
                if (priorityToken.Type != JTokenType.Integer)
                {
                    errorMessage = "items[" + index.ToString(CultureInfo.InvariantCulture) +
                                   "].priority 必须是整数。";
                    return false;
                }
                try
                {
                    priority = priorityToken.Value<int>();
                }
                catch
                {
                    errorMessage = "items[" + index.ToString(CultureInfo.InvariantCulture) +
                                   "].priority 超出整数范围。";
                    return false;
                }
            }

            OfficeWorkItem workItem = new OfficeWorkItem();
            workItem.AppType = normalizedApp;
            workItem.ToolName = tool;
            workItem.FilePath = file.Trim();
            workItem.Args = itemArguments == null ? new JObject() : (JObject)itemArguments.DeepClone();
            workItem.SavePolicy = savePolicy;
            workItem.TimeoutSec = timeoutSec;
            workItem.AffinityKey = ReadString(rawItem, "affinityKey").Trim();
            workItem.Priority = priority;
            workItem.IdempotencyKey = ReadString(rawItem, "idempotencyKey").Trim();

            request.App = normalizedApp;
            request.Tool = tool;
            request.File = file.Trim();
            request.SavePolicy = savePolicy;
            request.ReadOnly = ParallelToolCatalog.IsReadOnlyTool(normalizedApp, tool);
            request.WorkItem = workItem;
            return true;
        }

        private static OperationResult ReadTaskResult(Task<OperationResult> task)
        {
            if (task == null)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "并行作业没有返回 Task。");
            }
            if (!task.IsCompleted)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "office_batch_files 等待超时：该作业仍未完成，调度器会在自己的 itemTimeoutSec 到期时返回错误。");
            }
            if (task.Status != TaskStatus.RanToCompletion)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "并行作业 Task 状态异常：" + task.Status.ToString() + "。");
            }
            return task.Result;
        }

        private static JObject BuildResultRecord(BatchRequest request, OperationResult result)
        {
            OperationResult safeResult = result ?? OperationResult.Failed(
                OfficeErrorCode.INTERNAL_ERROR,
                "并行作业未返回结果。");

            JObject record = new JObject();
            record["index"] = request.Index;
            record["app"] = request.App ?? string.Empty;
            record["tool"] = request.Tool ?? string.Empty;
            record["file"] = request.File ?? string.Empty;
            record["readOnlyTool"] = request.ReadOnly;
            record["savePolicy"] = request.SavePolicy ?? ParallelToolCatalog.SavePolicyNone;
            record["affinityKey"] = request.WorkItem == null ? string.Empty : (request.WorkItem.AffinityKey ?? string.Empty);
            record["priority"] = request.WorkItem == null ? 0 : request.WorkItem.Priority;
            record["idempotencyKey"] = request.WorkItem == null ? string.Empty : (request.WorkItem.IdempotencyKey ?? string.Empty);
            record["success"] = safeResult.Success;
            record["errorCode"] = safeResult.ErrorCode ?? string.Empty;
            record["message"] = safeResult.UserMessage ?? string.Empty;
            record["recoverable"] = safeResult.Recoverable;

            if (safeResult.Data != null && safeResult.Data.HasValues)
            {
                record["data"] = safeResult.Data.DeepClone();
                JObject parallel = safeResult.Data["parallel"] as JObject;
                if (parallel != null)
                {
                    record["openedReadOnly"] = parallel["openedReadOnly"] ?? false;
                    record["saved"] = parallel["saved"] ?? false;
                    record["saveSkippedReason"] = parallel["saveSkippedReason"] ?? string.Empty;
                    record["saveVerified"] = parallel["saveVerified"] ?? false;
                    record["saveVerificationNote"] = parallel["saveVerificationNote"] ?? string.Empty;
                    record["saveChangedOnDisk"] = parallel["saveChangedOnDisk"] ?? false;
                    record["saveBeforeHash"] = parallel["saveBeforeHash"] ?? string.Empty;
                    record["saveAfterHash"] = parallel["saveAfterHash"] ?? string.Empty;
                    record["saveBeforeLength"] = parallel["saveBeforeLength"] ?? 0;
                    record["saveAfterLength"] = parallel["saveAfterLength"] ?? 0;
                    record["saveBeforeLastWriteUtc"] = parallel["saveBeforeLastWriteUtc"] ?? string.Empty;
                    record["saveAfterLastWriteUtc"] = parallel["saveAfterLastWriteUtc"] ?? string.Empty;
                    record["durationMs"] = parallel["durationMs"] ?? 0;
                    record["queueWaitMs"] = parallel["queueWaitMs"] ?? 0;
                    record["threadExited"] = parallel["threadExited"] ?? false;
                    record["executionStartedAtUtc"] = parallel["executionStartedAtUtc"] ?? JValue.CreateNull();
                    record["finishedAtUtc"] = parallel["finishedAtUtc"] ?? JValue.CreateNull();
                }
            }
            else
            {
                record["openedReadOnly"] = false;
                record["saved"] = false;
                record["saveSkippedReason"] = string.Empty;
                record["durationMs"] = 0;
                record["queueWaitMs"] = 0;
                record["threadExited"] = true;
            }
            return record;
        }

        private static string ReadString(JObject root, string propertyName)
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
            return token.Type == JTokenType.String ? (token.Value<string>() ?? string.Empty) : token.ToString();
        }

        private static ToolResult InvalidParameters(string message)
        {
            return ToolResult.Fail(
                OfficeErrorMapper.ToCode(OfficeErrorCode.INVALID_PARAMS),
                message,
                false);
        }

        private sealed class BatchRequest
        {
            public int Index;
            public string App;
            public string Tool;
            public string File;
            public string SavePolicy;
            public bool ReadOnly;
            public OfficeWorkItem WorkItem;
            public OperationResult ValidationFailure;
            public Task<OperationResult> Task;
        }
    }
}
