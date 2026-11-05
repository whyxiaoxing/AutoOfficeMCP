using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace Office.Parallel
{
    /// <summary>
    /// 并行调度单元：纯数据，不包含调度或 COM 逻辑。
    /// AppType 只允许 excel / word / ppt；FilePath 必须是显式文件路径；
    /// SavePolicy 只允许 none（默认）或 save；Completion 由调用方等待。
    /// </summary>
    public sealed class OfficeWorkItem
    {
        /// <summary>excel、word 或 ppt。</summary>
        public string AppType { get; set; }

        /// <summary>显式文件路径（提交时会被解析为全路径）。</summary>
        public string FilePath { get; set; }

        /// <summary>已注册工具名，必须与 AppType 前缀一致。</summary>
        public string ToolName { get; set; }

        /// <summary>工具参数；调度器会在执行前深拷贝，避免调用方后续修改。</summary>
        public JObject Args { get; set; }

        /// <summary>本项超时秒数；小于等于 0 时使用 config.json 的 parallel.itemTimeoutSec。</summary>
        public int TimeoutSec { get; set; }

        /// <summary>none（默认，不保存）或 save（工具成功后保存）。</summary>
        public string SavePolicy { get; set; }

        /// <summary>
        /// 亲和键：相同值的作业即使文件不同也串行执行（用于同一客户/同一批数据的互斥）。
        /// 为空表示不额外加锁。
        /// </summary>
        public string AffinityKey { get; set; }

        /// <summary>调度优先级：数值越大越优先；当前用于请求透传与状态展示，调度器按到达顺序执行。</summary>
        public int Priority { get; set; }

        /// <summary>
        /// 幂等键：同一用户下相同键的重复提交复用首次结果；完成后在 parallel.idempotencyTtlSec 内缓存。
        /// 为空表示不去重。
        /// </summary>
        public string IdempotencyKey { get; set; }

        /// <summary>作业结果；调度器只把 OperationResult 推入，回调内不碰 COM。</summary>
        public TaskCompletionSource<OperationResult> Completion { get; private set; }

        public OfficeWorkItem()
        {
            AppType = string.Empty;
            FilePath = string.Empty;
            ToolName = string.Empty;
            Args = new JObject();
            TimeoutSec = 0;
            SavePolicy = "none";
            AffinityKey = string.Empty;
            Priority = 0;
            IdempotencyKey = string.Empty;
            Completion = new TaskCompletionSource<OperationResult>(
                TaskCreationOptions.RunContinuationsAsynchronously);
        }
    }
}
