using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

namespace Office
{
    /// <summary>
    /// 全项目统一错误码。只允许追加成员，禁止修改既有成员的语义。
    /// SERVER_BUSY 是 T02 会话锁超时专用码（README 的枚举表未列出，但 T02 任务书要求）。
    /// </summary>
    public enum OfficeErrorCode
    {
        APP_NOT_RUNNING,
        DOC_NOT_OPEN,
        FILE_NOT_FOUND,
        INVALID_RANGE,
        INVALID_PARAMS,
        SHEET_NOT_FOUND,
        SHEET_PROTECTED,
        DOC_PROTECTED,
        READ_ONLY,
        PARTIAL_APPLY,
        CANCELLED,
        VBA_DISABLED,
        CLIPBOARD_DISABLED,
        OFFICE_BUSY,
        INTERNAL_ERROR,
        SERVER_BUSY
    }

    /// <summary>
    /// 带稳定错误码的 Office COM 异常。会话层用它向执行器表达"有明确语义的失败"，
    /// 执行器只需转换成 OperationResult，不必再猜 HRESULT。
    /// </summary>
    public sealed class OfficeComException : Exception
    {
        public OfficeErrorCode Code { get; private set; }

        public bool Recoverable { get; private set; }

        public OfficeComException(OfficeErrorCode code, string message)
            : this(code, message, null, OfficeErrorMapper.IsRecoverable(code))
        {
        }

        public OfficeComException(OfficeErrorCode code, string message, Exception innerException)
            : this(code, message, innerException, OfficeErrorMapper.IsRecoverable(code))
        {
        }

        public OfficeComException(OfficeErrorCode code, string message, Exception innerException, bool recoverable)
            : base(message, innerException)
        {
            Code = code;
            Recoverable = recoverable;
        }
    }

    /// <summary>
    /// COMException / 普通异常 → 稳定错误码 + 中文用户消息。
    /// 映射策略：先按 HRESULT 精确映射；未知 HRESULT 再按消息关键词兜底，
    /// 任何情况下都返回确定值，不把裸 COMException 抛给协议层。
    /// </summary>
    public static class OfficeErrorMapper
    {
        // 常用 HRESULT 常量（int 为有符号值，和 COMException.HResult 比较）
        public const int MkEUnavailable = unchecked((int)0x800401E3);
        public const int CoEObjNotConnected = unchecked((int)0x800401FD);
        public const int RpcEDisconnected = unchecked((int)0x80010108);
        public const int RpcECallRejected = unchecked((int)0x80010001);
        public const int RpcEServerCallRetryLater = unchecked((int)0x8001010A);
        public const int RpcEServerCallRejected = unchecked((int)0x8001010B);
        public const int CoEServerExecFailure = unchecked((int)0x80080005);
        public const int CoEClassString = unchecked((int)0x800401F3);
        public const int RegDbEClassNotReg = unchecked((int)0x80040154);
        public const int ExcelGeneric = unchecked((int)0x800A03EC);
        public const int ExcelSubscriptOutOfRange = unchecked((int)0x800A0009);
        public const int WordCollectionMemberNotFound = unchecked((int)0x800A1735);
        public const int WordObjectRequired = unchecked((int)0x800A01A8);
        public const int WordMethodFailed = unchecked((int)0x800A01B6);
        public const int VbaIgnore = unchecked((int)0x800AC472);
        public const int VbaProjectNotTrusted = unchecked((int)0x800A9C68);
        public const int Win32FileNotFound = unchecked((int)0x80070002);
        public const int Win32PathNotFound = unchecked((int)0x80070003);
        public const int Win32AccessDenied = unchecked((int)0x80070005);
        public const int Win32SharingViolation = unchecked((int)0x80070020);
        public const int StgFileNotFound = unchecked((int)0x80030002);
        public const int StgAccessDenied = unchecked((int)0x80030005);
        public const int StgShareViolation = unchecked((int)0x80030020);

        /// <summary>稳定错误码字符串（与枚举成员同名，全大写）。</summary>
        public static string ToCode(OfficeErrorCode code)
        {
            return code.ToString();
        }

        /// <summary>大小写不敏感地解析错误码字符串。</summary>
        public static bool TryParseCode(string value, out OfficeErrorCode code)
        {
            code = OfficeErrorCode.INTERNAL_ERROR;
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            string trimmed = value.Trim();
            foreach (OfficeErrorCode candidate in Enum.GetValues(typeof(OfficeErrorCode)))
            {
                if (string.Equals(candidate.ToString(), trimmed, StringComparison.OrdinalIgnoreCase))
                {
                    code = candidate;
                    return true;
                }
            }
            return false;
        }

        /// <summary>每个错误码都有非空中文消息（协议层可安全直接展示）。</summary>
        public static string GetUserMessage(OfficeErrorCode code)
        {
            switch (code)
            {
                case OfficeErrorCode.APP_NOT_RUNNING:
                    return "Office 应用未运行，且当前配置不允许自动启动。请先手动打开 Excel/Word/PowerPoint，或启用 office.excel.autoLaunch / office.word.autoLaunch / office.ppt.autoLaunch 后重试。";
                case OfficeErrorCode.DOC_NOT_OPEN:
                    return "当前没有打开的目标文档，或文档未处于活动状态。请先打开并激活目标文档后重试。";
                case OfficeErrorCode.FILE_NOT_FOUND:
                    return "指定的文件不存在或路径不可访问。请检查路径、文件名与扩展名后重试。";
                case OfficeErrorCode.INVALID_RANGE:
                    return "目标区域无效或超出工作表范围。请检查区域地址（例如 Sheet1!A1:C10）后重试。";
                case OfficeErrorCode.INVALID_PARAMS:
                    return "工具参数无效或缺少必填项。请检查参数名称、类型与取值范围。";
                case OfficeErrorCode.SHEET_NOT_FOUND:
                    return "找不到指定的工作表。请先列出工作表并确认名称（注意空格与大小写）后重试。";
                case OfficeErrorCode.SHEET_PROTECTED:
                    return "工作表受保护，当前操作无法写入。请先取消工作表保护（需要密码时由用户提供）后重试。";
                case OfficeErrorCode.DOC_PROTECTED:
                    return "文档受保护（只读或限制编辑），当前操作无法写入。请解除保护或改用可写副本后重试。";
                case OfficeErrorCode.READ_ONLY:
                    return "文件或文档处于只读状态，无法保存或修改。请检查文件属性、是否被其他程序占用，或另存为可写副本。";
                case OfficeErrorCode.PARTIAL_APPLY:
                    return "操作只完成了一部分，后续步骤未执行。已完成的改动不会自动回放，请根据回滚说明人工确认。";
                case OfficeErrorCode.CANCELLED:
                    return "操作已取消，未对文档做任何修改。";
                case OfficeErrorCode.VBA_DISABLED:
                    return "VBA 执行被安全策略拒绝。若要放行，请编辑 exe 同目录的 config.json，将 security.allowVba 设为 true 后重启 MCP server。";
                case OfficeErrorCode.CLIPBOARD_DISABLED:
                    return "该工具依赖 Windows 剪贴板，已被安全策略禁用。请改用不依赖剪贴板的工具，或在受控环境下将 security.allowClipboard 设为 true 后重启 worker。";
                case OfficeErrorCode.OFFICE_BUSY:
                    return "Office 正忙（可能处于单元格编辑状态、弹窗阻塞或正在处理其他请求），请稍后重试。";
                case OfficeErrorCode.SERVER_BUSY:
                    return "服务器繁忙：等待 Office 会话锁超时。请稍后重试，或减少并发请求。";
                case OfficeErrorCode.INTERNAL_ERROR:
                    return "发生内部错误，操作可能未完成。请查看 stderr 日志，并提供错误信息与操作步骤。";
                default:
                    return "发生未知错误，操作可能未完成。请查看 stderr 日志。";
            }
        }

        /// <summary>错误是否属于"调整参数/稍后重试即可"的可恢复类别。</summary>
        public static bool IsRecoverable(OfficeErrorCode code)
        {
            switch (code)
            {
                case OfficeErrorCode.PARTIAL_APPLY:
                case OfficeErrorCode.VBA_DISABLED:
                case OfficeErrorCode.CLIPBOARD_DISABLED:
                case OfficeErrorCode.INTERNAL_ERROR:
                    return false;
                default:
                    return true;
            }
        }

        /// <summary>COMException HRESULT → 稳定错误码。未知 HRESULT 统一落到 INTERNAL_ERROR。</summary>
        public static OfficeErrorCode MapComException(COMException exception)
        {
            if (exception == null)
            {
                return OfficeErrorCode.INTERNAL_ERROR;
            }

            switch (exception.HResult)
            {
                case MkEUnavailable:
                case CoEObjNotConnected:
                case RpcEDisconnected:
                case CoEClassString:
                case RegDbEClassNotReg:
                    return OfficeErrorCode.APP_NOT_RUNNING;
                case RpcECallRejected:
                case RpcEServerCallRetryLater:
                case RpcEServerCallRejected:
                case CoEServerExecFailure:
                case VbaIgnore:
                    return OfficeErrorCode.OFFICE_BUSY;
                case VbaProjectNotTrusted:
                    return OfficeErrorCode.VBA_DISABLED;
                case ExcelGeneric:
                    return OfficeErrorCode.INVALID_PARAMS;
                case ExcelSubscriptOutOfRange:
                    return OfficeErrorCode.INVALID_RANGE;
                case WordCollectionMemberNotFound:
                case WordObjectRequired:
                case WordMethodFailed:
                    return OfficeErrorCode.INVALID_PARAMS;
                case Win32FileNotFound:
                case Win32PathNotFound:
                case StgFileNotFound:
                    return OfficeErrorCode.FILE_NOT_FOUND;
                case Win32AccessDenied:
                case Win32SharingViolation:
                case StgAccessDenied:
                case StgShareViolation:
                    return OfficeErrorCode.READ_ONLY;
                default:
                    return OfficeErrorCode.INTERNAL_ERROR;
            }
        }

        /// <summary>
        /// 异常 → 错误码。先按 HRESULT；无法确定时按消息关键词兜底，
        /// 并借助应用类型区分"工作表受保护"和"文档受保护"。
        /// </summary>
        public static OfficeErrorCode MapException(Exception exception, OfficeApp application)
        {
            if (exception == null)
            {
                return OfficeErrorCode.INTERNAL_ERROR;
            }

            OfficeComException officeException = exception as OfficeComException;
            if (officeException != null)
            {
                return officeException.Code;
            }

            COMException comException = exception as COMException;
            if (comException != null)
            {
                OfficeErrorCode hresultCode = MapComException(comException);
                OfficeErrorCode messageCode = MapMessage(comException.Message, application);
                if (messageCode != OfficeErrorCode.INTERNAL_ERROR)
                {
                    return messageCode;
                }
                return hresultCode;
            }

            OfficeErrorCode fallback = MapMessage(exception.Message, application);
            return fallback == OfficeErrorCode.INTERNAL_ERROR ? OfficeErrorCode.INTERNAL_ERROR : fallback;
        }

        /// <summary>中文用户消息 + 原始错误细节（保留 HRESULT 便于排查）。</summary>
        public static string DescribeException(Exception exception, OfficeApp application)
        {
            if (exception == null)
            {
                return GetUserMessage(OfficeErrorCode.INTERNAL_ERROR);
            }

            OfficeComException officeException = exception as OfficeComException;
            if (officeException != null)
            {
                string officeDetail = DescribeInnerMessages(officeException);
                return officeDetail.Length == 0
                    ? officeException.Message
                    : officeException.Message + "（内层错误：" + officeDetail + "）";
            }

            OfficeErrorCode code = MapException(exception, application);
            string baseMessage = GetUserMessage(code);
            string detail = string.IsNullOrWhiteSpace(exception.Message) ? string.Empty : exception.Message.Trim();
            string innerDetail = DescribeInnerMessages(exception);

            COMException comException = exception as COMException;
            if (comException != null)
            {
                string hresult = "0x" + comException.HResult.ToString("X8", CultureInfo.InvariantCulture);
                if (detail.Length == 0)
                {
                    return baseMessage + "（HRESULT: " + hresult + "）";
                }
                string innerSuffix = innerDetail.Length == 0 ? string.Empty : "；内层错误：" + innerDetail;
                return baseMessage + "（HRESULT: " + hresult + "；原始信息：" + detail + innerSuffix + "）";
            }

            if (detail.Length == 0)
            {
                return baseMessage;
            }
            return baseMessage + "（原始信息：" + detail +
                   (innerDetail.Length == 0 ? string.Empty : "；内层错误：" + innerDetail) + "）";
        }

        /// <summary>
        /// 递归收集 InnerException 的消息，避免 COM 包装异常只暴露外层“调用失败”。
        /// </summary>
        public static string DescribeInnerMessages(Exception exception)
        {
            if (exception == null || exception.InnerException == null)
            {
                return string.Empty;
            }

            StringBuilder builder = new StringBuilder();
            Exception current = exception.InnerException;
            int depth = 0;
            while (current != null)
            {
                if (builder.Length > 0)
                {
                    builder.Append(" <- ");
                }
                builder.Append("[").Append(depth.ToString()).Append("] ");
                builder.Append(current.GetType().Name);
                if (!string.IsNullOrWhiteSpace(current.Message))
                {
                    builder.Append(": ").Append(current.Message.Trim());
                }
                current = current.InnerException;
                depth++;
            }
            return builder.ToString();
        }

        /// <summary>
        /// 格式化完整异常链（类型/消息/HRESULT/堆栈），供 stderr 错误日志使用。
        /// 只允许写入日志，不进入 MCP 响应。
        /// </summary>
        public static string FormatExceptionChain(Exception exception)
        {
            if (exception == null)
            {
                return string.Empty;
            }

            StringBuilder builder = new StringBuilder();
            Exception current = exception;
            int depth = 0;
            while (current != null)
            {
                if (depth > 0)
                {
                    builder.AppendLine("--- 内层异常 " + depth.ToString() + " ---");
                }
                builder.AppendLine("类型：" + current.GetType().FullName);
                builder.AppendLine("消息：" + (current.Message ?? string.Empty));

                COMException comException = current as COMException;
                if (comException != null)
                {
                    builder.AppendLine("HRESULT：0x" +
                                       comException.HResult.ToString("X8", CultureInfo.InvariantCulture));
                }

                if (!string.IsNullOrWhiteSpace(current.StackTrace))
                {
                    builder.AppendLine("堆栈：");
                    builder.AppendLine(current.StackTrace);
                }

                current = current.InnerException;
                depth++;
            }
            return builder.ToString().TrimEnd();
        }

        /// <summary>把异常直接转换成失败结果；执行器 catch 块可一行复用。</summary>
        public static OperationResult ToOperationResult(Exception exception, OfficeApp application, string rollbackNote = null)
        {
            OfficeErrorCode code = MapException(exception, application);
            OperationResult result = OperationResult.Failed(code, DescribeException(exception, application), rollbackNote);
            result.Diagnostic = FormatExceptionChain(exception);
            return result;
        }

        private static OfficeErrorCode MapMessage(string message, OfficeApp application)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return OfficeErrorCode.INTERNAL_ERROR;
            }

            string text = message.ToLowerInvariant();

            if (ContainsAny(text, "not trusted", "vbproject", "visual basic project", "信任", "信任中心"))
            {
                return OfficeErrorCode.VBA_DISABLED;
            }
            if (ContainsAny(text, "read-only", "readonly", "只读"))
            {
                return OfficeErrorCode.READ_ONLY;
            }
            if (ContainsAny(text, "protected", "protected view", "受保护", "保护视图"))
            {
                return application == OfficeApp.Excel ? OfficeErrorCode.SHEET_PROTECTED : OfficeErrorCode.DOC_PROTECTED;
            }
            if (ContainsAny(text, "sheet", "工作表") &&
                ContainsAny(text, "not found", "does not exist", "找不到", "不存在", "无效"))
            {
                return OfficeErrorCode.SHEET_NOT_FOUND;
            }
            if (ContainsAny(text, "file not found", "path not found", "找不到文件", "文件不存在", "路径不存在"))
            {
                return OfficeErrorCode.FILE_NOT_FOUND;
            }
            if (ContainsAny(text, "not open", "未打开", "没有打开", "not available", "不可用"))
            {
                return OfficeErrorCode.DOC_NOT_OPEN;
            }
            if (ContainsAny(text, "range", "区域", "单元格", "cell"))
            {
                return OfficeErrorCode.INVALID_RANGE;
            }
            if (ContainsAny(text, "busy", "忙", "call was rejected", "被拒绝", "retry later", "稍后重试"))
            {
                return OfficeErrorCode.OFFICE_BUSY;
            }
            if (ContainsAny(text, "invalid", "无效", "参数", "parameter", "argument", "参数错误"))
            {
                return OfficeErrorCode.INVALID_PARAMS;
            }

            return OfficeErrorCode.INTERNAL_ERROR;
        }

        private static bool ContainsAny(string text, params string[] words)
        {
            foreach (string word in words)
            {
                if (text.IndexOf(word, StringComparison.Ordinal) >= 0)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
