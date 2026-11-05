using System;
using Office.Config;

namespace Office
{
    /// <summary>需要 config.json 显式放行的高危操作。</summary>
    public enum SecurityRisk
    {
        Vba,
        DeleteSheet,
        CloseWithoutSave
    }

    /// <summary>安全门判定结果。</summary>
    public sealed class SecurityDecision
    {
        public bool Allowed { get; private set; }

        public OfficeErrorCode ErrorCode { get; private set; }

        public string Message { get; private set; }

        private SecurityDecision(bool allowed, OfficeErrorCode errorCode, string message)
        {
            Allowed = allowed;
            ErrorCode = errorCode;
            Message = message ?? string.Empty;
        }

        public static SecurityDecision Allow()
        {
            return new SecurityDecision(true, OfficeErrorCode.CANCELLED, string.Empty);
        }

        public static SecurityDecision Deny(OfficeErrorCode errorCode, string message)
        {
            return new SecurityDecision(false, errorCode, message);
        }
    }

    /// <summary>
    /// 高危操作门禁：默认全部拒绝，只有 config.json 显式放行才通过。
    /// 拒绝消息必须说明"改哪个配置项、改成什么、何时生效"。
    /// </summary>
    public static class SecurityGate
    {
        public static SecurityDecision Evaluate(SecurityRisk risk, string toolName)
        {
            AppConfig config = AppConfig.Current;
            string tool = string.IsNullOrWhiteSpace(toolName) ? "该操作" : "工具「" + toolName.Trim() + "」";
            string configPath = AppConfig.ConfigPath;

            switch (risk)
            {
                case SecurityRisk.Vba:
                    if (config.AllowVba)
                    {
                        return SecurityDecision.Allow();
                    }
                    return SecurityDecision.Deny(
                        OfficeErrorCode.VBA_DISABLED,
                        tool + " 被安全策略拒绝：config.json 中 security.allowVba 为 false。" +
                        "若要放行，请编辑 " + configPath + "，把 security.allowVba 改为 true，然后重启 MCP server。");

                case SecurityRisk.DeleteSheet:
                    if (config.AllowDeleteSheet)
                    {
                        return SecurityDecision.Allow();
                    }
                    return SecurityDecision.Deny(
                        OfficeErrorCode.CANCELLED,
                        tool + " 被安全策略取消：config.json 中 security.allowDeleteSheet 为 false，删除工作表属于不可逆高危操作。" +
                        "若要放行，请编辑 " + configPath + "，把 security.allowDeleteSheet 改为 true，然后重启 MCP server；" +
                        "也可以先让用户确认并手动删除。");

                case SecurityRisk.CloseWithoutSave:
                    if (config.AllowCloseWithoutSave)
                    {
                        return SecurityDecision.Allow();
                    }
                    return SecurityDecision.Deny(
                        OfficeErrorCode.CANCELLED,
                        tool + " 被安全策略取消：config.json 中 security.allowCloseWithoutSave 为 false，关闭文档但不保存可能丢失用户内容。" +
                        "若要放行，请编辑 " + configPath + "，把 security.allowCloseWithoutSave 改为 true，然后重启 MCP server。");

                default:
                    return SecurityDecision.Deny(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "未知的安全风险类型：" + risk.ToString() + "，已按拒绝处理。");
            }
        }

        public static SecurityDecision Evaluate(SecurityRisk risk)
        {
            return Evaluate(risk, null);
        }

        public static bool IsAllowed(SecurityRisk risk)
        {
            return Evaluate(risk, null).Allowed;
        }

        public static SecurityDecision CheckVba(string toolName)
        {
            return Evaluate(SecurityRisk.Vba, toolName);
        }

        public static SecurityDecision CheckDeleteSheet(string toolName)
        {
            return Evaluate(SecurityRisk.DeleteSheet, toolName);
        }

        public static SecurityDecision CheckCloseWithoutSave(string toolName)
        {
            return Evaluate(SecurityRisk.CloseWithoutSave, toolName);
        }

        /// <summary>
        /// 允许时返回 null；拒绝时返回可直接交给协议层的失败结果。
        /// 用法：OperationResult denial = SecurityGate.CreateDenial(...); if (denial != null) return denial;
        /// </summary>
        public static OperationResult CreateDenial(SecurityRisk risk, string toolName)
        {
            SecurityDecision decision = Evaluate(risk, toolName);
            if (decision.Allowed)
            {
                return null;
            }
            return OperationResult.Failed(decision.ErrorCode, decision.Message);
        }

        public static OperationResult CreateDenial(SecurityRisk risk)
        {
            return CreateDenial(risk, null);
        }
    }
}
