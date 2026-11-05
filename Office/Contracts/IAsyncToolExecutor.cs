using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace Office.Contracts
{
    /// <summary>
    /// 异步工具执行器：传输入口在协议线程上 await，不阻塞读循环。
    /// 同步 IToolExecutor 保持原签名；需要异步批处理/远端调度的执行器额外实现本接口。
    /// </summary>
    public interface IAsyncToolExecutor
    {
        string AppType { get; }

        bool CanExecute(string toolName);

        Task<ToolResult> ExecuteAsync(string toolName, JObject arguments);
    }
}
