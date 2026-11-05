using System;
using Newtonsoft.Json.Linq;
using Office.Contracts;

namespace AutoOfficeMCP.Tools
{
    /// <summary>
    /// T01 占位执行器：认领 stub 探针工具，并在没有任何真实执行器认领时兜底。
    /// 真实工具未接入前统一返回可读的 isError 结果，不做假成功。
    /// </summary>
    public sealed class StubToolExecutor : IToolExecutor
    {
        public const string ProbeToolName = "office_stub_probe";

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

            return string.Equals(toolName, ProbeToolName, StringComparison.OrdinalIgnoreCase);
        }

        public ToolResult Execute(string toolName, JObject arguments)
        {
            // T01 不实现任何 Office 操作：即使路由到 stub 探针，也如实返回未接入，禁止假成功。
            return ToolResult.Fail(
                "EXECUTOR_NOT_AVAILABLE",
                "执行器未接入：工具 " + toolName + " 暂无对应执行器（T01 仅提供 stub 占位）。",
                false);
        }
    }
}
