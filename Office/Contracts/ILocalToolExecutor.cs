using Newtonsoft.Json.Linq;

namespace Office.Contracts
{
    /// <summary>
    /// 标记执行器中有“纯本地/无 COM”工具。协议层在固定进程 worker 模式下对这些工具直接执行，
    /// 不要求显式 file，也不投递 worker（例如 OpenXML 只读查询、静态函数目录）。
    /// </summary>
    public interface ILocalToolExecutor
    {
        bool IsLocalTool(string toolName, JObject arguments);
    }
}
