using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Office.Contracts;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    /// <summary>
    /// T06 Word 扩展能力执行器。
    /// 与 T04 的 WordToolExecutor 并存：工具名不重叠，主程序按 CanExecute 精确路由。
    /// 处理器按文字/文档/内容拆分到 WordExtensions*.cs，本文件只负责接口、路由与统一读/写外壳。
    /// 第二个构造仅给本机 STA 冒烟/集成测试注入独立 Word 实例，避免影响用户已打开的文档。
    /// </summary>
    public sealed partial class WordExtensionsExecutor : IToolExecutor
    {
        private readonly Word.Application _applicationOverride;
        private readonly Word.Document _documentOverride;
        private readonly object _overrideLock = new object();
        private readonly Dictionary<string, Func<Word.Application, JObject, OperationResult>> _handlers;

        public WordExtensionsExecutor()
        {
            _handlers = BuildHandlers();
            if (_handlers.Count != WordExtensionsToolNames.Count)
            {
                Console.Error.WriteLine("[word-ext] 警告：处理器数量 " + _handlers.Count +
                                        " 与扩展工具名清单 " + WordExtensionsToolNames.Count + " 不一致。");
            }
        }

        /// <summary>仅供本机 STA 冒烟/集成测试使用：注入独立 Word 实例，不经过 OfficeSession。</summary>
        public WordExtensionsExecutor(Word.Application application)
            : this()
        {
            if (application == null)
            {
                throw new ArgumentNullException("application");
            }
            _applicationOverride = application;
        }

        /// <summary>
        /// 仅供本机 STA 冒烟/集成测试使用：同时注入 Word 实例与目标文档，
        /// 避免同一 Word 进程里其他自动化任务切换 ActiveDocument 时误操作。
        /// </summary>
        public WordExtensionsExecutor(Word.Application application, Word.Document document)
            : this(application)
        {
            if (document == null)
            {
                throw new ArgumentNullException("document");
            }
            _documentOverride = document;
        }

        public string AppType
        {
            get { return "word"; }
        }

        public bool CanExecute(string toolName)
        {
            return WordExtensionsToolNames.Contains(toolName);
        }

        public ToolResult Execute(string toolName, JObject arguments)
        {
            if (!CanExecute(toolName))
            {
                return ToolResult.Fail(
                    "INVALID_PARAMS",
                    "Word 扩展执行器不认领工具「" + (toolName ?? string.Empty) + "」。",
                    false);
            }

            JObject safeArguments = arguments ?? new JObject();
            if (_applicationOverride != null)
            {
                lock (_overrideLock)
                {
                    try
                    {
                        if (_documentOverride != null)
                        {
                            // 测试注入口用：Word 的页眉 Shapes / ContentControls 等操作要求文档窗口处于活动状态。
                            // 生产路径通过 ActiveDocument 获得目标文档，无需在这里切换。
                            try
                            {
                                _documentOverride.Activate();
                            }
                            catch (Exception activationError)
                            {
                                Console.Error.WriteLine("[word-ext] 激活注入文档失败：" +
                                                        activationError.Message);
                            }
                        }
                        OperationResult result = Dispatch(
                            _applicationOverride, toolName.Trim(), safeArguments);
                        return result == null
                            ? ToolResult.Fail("INTERNAL_ERROR", "Word 扩展执行器返回空结果。", false)
                            : result.ToToolResult();
                    }
                    catch (Exception ex)
                    {
                        return ToolResult.Fail(
                            OfficeErrorMapper.ToCode(OfficeErrorMapper.MapException(ex, OfficeApp.Word)),
                            OfficeErrorMapper.DescribeException(ex, OfficeApp.Word),
                            OfficeErrorMapper.IsRecoverable(
                                OfficeErrorMapper.MapException(ex, OfficeApp.Word)));
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
                    "Word 扩展执行器没有为工具「" + toolName + "」注册处理器（清单漂移）。");
            }
            return handler(application, arguments);
        }

        private Dictionary<string, Func<Word.Application, JObject, OperationResult>> BuildHandlers()
        {
            Dictionary<string, Func<Word.Application, JObject, OperationResult>> handlers =
                new Dictionary<string, Func<Word.Application, JObject, OperationResult>>(
                    StringComparer.Ordinal);
            RegisterTextExtensionHandlers(handlers);
            RegisterDocumentExtensionHandlers(handlers);
            RegisterContentExtensionHandlers(handlers);
            RegisterP3MoreHandlers(handlers);
            return handlers;
        }

        /// <summary>只读工具的统一入口：取得活动文档并保证异常转稳定错误码。</summary>
        private OperationResult RunReadTool(
            Word.Application application,
            Func<Word.Document, OperationResult> action)
        {
            Word.Document document = null;
            bool releaseDocument = _documentOverride == null;
            try
            {
                document = _documentOverride ?? application.ActiveDocument;
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
                Console.Error.WriteLine("[word-ext] 只读工具异常：" + ex);
                return OfficeErrorMapper.ToOperationResult(ex, OfficeApp.Word);
            }
            finally
            {
                if (releaseDocument)
                {
                    ComHelper.Release(document);
                }
            }
        }

        /// <summary>写入工具的统一入口：目标选择、只读/保护检查、写后观察都由这里兜底。</summary>
        private OperationResult RunWriteTool(
            Word.Application application,
            string toolName,
            bool allowReadOnly,
            Func<Word.Document, WordWriteTarget> selectTarget,
            Func<Word.Document, WordWriteTarget, WordWriteResult> action)
        {
            Word.Document document = null;
            bool releaseDocument = _documentOverride == null;
            try
            {
                document = _documentOverride ?? application.ActiveDocument;
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
                    });
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[word-ext] 写入工具异常：" + ex);
                return OfficeErrorMapper.ToOperationResult(ex, OfficeApp.Word);
            }
            finally
            {
                if (releaseDocument)
                {
                    ComHelper.Release(document);
                }
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

    /// <summary>
    /// T06 Word 扩展工具名清单。必须与 Tools\word\*.json 描述符、处理器表三处一致。
    /// </summary>
    internal static class WordExtensionsToolNames
    {
        public const string StyleAudit = "word_style_audit";
        public const string PunctuationNormalize = "word_punctuation_normalize";
        public const string CjkSpacing = "word_cjk_spacing";
        public const string NormalizeHeadings = "word_normalize_headings";
        public const string ExtractImages = "word_extract_images";
        public const string DocMerge = "word_doc_merge";

        public const string MdExport = "word_md_export";
        public const string MdImport = "word_md_import";
        public const string TableConvert = "word_table_convert";
        public const string CaptionXref = "word_caption_xref";
        public const string SectionManage = "word_section_manage";
        public const string Watermark = "word_watermark";
        public const string RepeatCheck = "word_repeat_check";
        public const string HiddenTextReport = "word_hidden_text_report";
        public const string FootnoteManage = "word_footnote_manage";

        public const string ContentControls = "word_content_controls";
        public const string LineNumbering = "word_line_numbering";
        public const string ReadingTime = "word_reading_time";
        public const string VerticalText = "word_vertical_text";

        private static readonly string[] ToolNames = new string[]
        {
            StyleAudit,
            PunctuationNormalize,
            CjkSpacing,
            NormalizeHeadings,
            ExtractImages,
            DocMerge,
            MdExport,
            MdImport,
            TableConvert,
            CaptionXref,
            SectionManage,
            Watermark,
            RepeatCheck,
            HiddenTextReport,
            FootnoteManage,
            ContentControls,
            LineNumbering,
            ReadingTime,
            VerticalText
        };

        private static readonly HashSet<string> ToolSet =
            new HashSet<string>(ToolNames, StringComparer.Ordinal);

        public static int Count
        {
            get { return ToolSet.Count; }
        }

        public static bool Contains(string toolName)
        {
            return !string.IsNullOrWhiteSpace(toolName) && ToolSet.Contains(toolName.Trim());
        }

        public static string[] GetAllNames()
        {
            string[] names = new string[ToolSet.Count];
            ToolSet.CopyTo(names);
            Array.Sort(names, StringComparer.Ordinal);
            return names;
        }
    }
}
