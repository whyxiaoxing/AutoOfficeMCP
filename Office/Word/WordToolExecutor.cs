using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Office.Contracts;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    /// <summary>
    /// Word 主执行器（T04）。工具处理按只读/文字/表格/文档/修订/公式/图片拆分到 partial 文件，
    /// 本文件只负责接口、路由与统一的读/写外壳。
    ///
    /// 生产环境的无参构造通过 OfficeSession.RunWord 取得 STA 会话与进程级互斥；
    /// 第二个构造仅给冒烟/集成测试注入独立 Word 实例，避免影响用户已打开的文档。
    /// </summary>
    public sealed partial class WordToolExecutor : IToolExecutor
    {
        private readonly Word.Application _applicationOverride;
        private readonly object _overrideLock = new object();
        private readonly Dictionary<string, Func<Word.Application, JObject, OperationResult>> _handlers;

        public WordToolExecutor()
        {
            _handlers = BuildHandlers();
            if (_handlers.Count != WordToolNames.Count)
            {
                Console.Error.WriteLine("[word] 警告：处理器数量 " + _handlers.Count +
                                        " 与工具名清单 " + WordToolNames.Count + " 不一致。");
            }
        }

        /// <summary>仅供本机 STA 冒烟/集成测试使用：注入独立 Word 实例，不经过 OfficeSession。</summary>
        public WordToolExecutor(Word.Application application)
            : this()
        {
            if (application == null)
            {
                throw new ArgumentNullException("application");
            }
            _applicationOverride = application;
        }

        public string AppType
        {
            get { return "word"; }
        }

        public bool CanExecute(string toolName)
        {
            return WordToolNames.Contains(toolName);
        }

        public ToolResult Execute(string toolName, JObject arguments)
        {
            if (!CanExecute(toolName))
            {
                return ToolResult.Fail(
                    "INVALID_PARAMS",
                    "Word 执行器不认领工具「" + (toolName ?? string.Empty) + "」。",
                    false);
            }

            JObject safeArguments = arguments ?? new JObject();
            if (_applicationOverride != null)
            {
                lock (_overrideLock)
                {
                    try
                    {
                        OperationResult result = Dispatch(_applicationOverride, toolName.Trim(), safeArguments);
                        return result == null
                            ? ToolResult.Fail("INTERNAL_ERROR", "Word 执行器返回空结果。", false)
                            : result.ToToolResult();
                    }
                    catch (Exception ex)
                    {
                        return ToolResult.Fail(
                            OfficeErrorMapper.ToCode(OfficeErrorMapper.MapException(ex, OfficeApp.Word)),
                            OfficeErrorMapper.DescribeException(ex, OfficeApp.Word),
                            OfficeErrorMapper.IsRecoverable(OfficeErrorMapper.MapException(ex, OfficeApp.Word)));
                    }
                }
            }

            OperationResult sessionResult = OfficeSession.RunWord(
                toolName,
                delegate(Word.Application application)
                {
                    return Dispatch(application, toolName.Trim(), safeArguments);
                });
            return sessionResult.ToToolResult();
        }

        private OperationResult Dispatch(Word.Application application, string toolName, JObject arguments)
        {
            Func<Word.Application, JObject, OperationResult> handler;
            if (!_handlers.TryGetValue(toolName, out handler))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "Word 执行器没有为工具「" + toolName + "」注册处理器（清单漂移）。");
            }
            return handler(application, arguments);
        }

        private Dictionary<string, Func<Word.Application, JObject, OperationResult>> BuildHandlers()
        {
            Dictionary<string, Func<Word.Application, JObject, OperationResult>> handlers =
                new Dictionary<string, Func<Word.Application, JObject, OperationResult>>(StringComparer.Ordinal);
            RegisterReadHandlers(handlers);
            RegisterTextHandlers(handlers);
            RegisterTableHandlers(handlers);
            RegisterDocumentHandlers(handlers);
            RegisterRevisionHandlers(handlers);
            RegisterEquationHandlers(handlers);
            RegisterImageHandlers(handlers);
            RegisterTodoReadHandlers(handlers);
            RegisterTodoWriteHandlers(handlers);
            return handlers;
        }

        /// <summary>只读工具的统一入口：取得活动文档并保证异常转稳定错误码。</summary>
        private OperationResult RunReadTool(
            Word.Application application,
            Func<Word.Document, OperationResult> action)
        {
            Word.Document document = null;
            try
            {
                document = application.ActiveDocument;
                if (document == null)
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.DOC_NOT_OPEN,
                        "当前没有活动 Word 文档。请先打开并激活目标文档后重试。");
                }
                return action(document);
            }
            catch (Exception ex)
            {
                return OfficeErrorMapper.ToOperationResult(ex, OfficeApp.Word);
            }
            finally
            {
                ComHelper.Release(document);
            }
        }

        /// <summary>写入工具的统一入口：目标选择、只读/保护检查、写后观察都由这里兜底。</summary>
        private OperationResult RunWriteTool(
            Word.Application application,
            string toolName,
            bool allowReadOnly,
            Func<Word.Document, WordWriteTarget> selectTarget,
            Func<Word.Document, WordWriteTarget, WordWriteResult> action,
            bool useNativeUndo = false)
        {
            Word.Document document = null;
            try
            {
                document = application.ActiveDocument;
                if (document == null)
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.DOC_NOT_OPEN,
                        "当前没有活动 Word 文档。请先打开并激活目标文档后重试。");
                }
                if (!allowReadOnly)
                {
                    WordComUtil.CheckWritable(document, toolName);
                }

                WordWriteTarget target = selectTarget == null
                    ? WordWriteTargets.None("Word 文档", "Word:Document")
                    : selectTarget(document);
                if (target == null)
                {
                    target = WordWriteTargets.None("Word 文档", "Word:Document");
                }

                return WordWriteObservation.Run(
                    application,
                    document,
                    toolName,
                    target,
                    delegate
                    {
                        return action(document, target);
                    },
                    useNativeUndo);
            }
            catch (Exception ex)
            {
                return OfficeErrorMapper.ToOperationResult(ex, OfficeApp.Word);
            }
            finally
            {
                ComHelper.Release(document);
            }
        }

        private static OperationResult ReadSucceed(
            string message,
            JObject data,
            params string[] targetRefs)
        {
            Observation observation = Observation.Create("read", message, false);
            if (targetRefs != null)
            {
                foreach (string targetRef in targetRefs)
                {
                    observation.AddTargetRef(targetRef);
                }
            }
            return OperationResult.Succeed(message, data, observation);
        }
    }
}
