using Newtonsoft.Json.Linq;

namespace Office.Contracts
{
    /// <summary>
    /// 执行器统一接口。主程序通过反射发现 Office/Xlsx 程序集中的实现并注册。
    /// </summary>
    public interface IToolExecutor
    {
        /// <summary>执行器所属应用：excel / word / common。</summary>
        string AppType { get; }

        /// <summary>该执行器是否认领指定工具名（同一 appType 允许多个执行器并存）。</summary>
        bool CanExecute(string toolName);

        /// <summary>
        /// 执行一个已注册工具。参数由协议层完成 JSON Schema 基础校验后传入。
        /// 实现必须自行捕获 COM 异常并转换为 ToolResult.Fail，不允许把异常抛回协议层。
        /// </summary>
        ToolResult Execute(string toolName, JObject arguments);
    }
}
