using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Office;
using Office.Contracts;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>
    /// T06 Excel 扩展能力执行器。
    /// 工具名事实源是本类的 Handlers 字典：描述符 id/name、CanExecute、Execute 共用同一组键。
    /// 所有会话型调用统一走 OfficeSession.RunExcel；本类不持有 Application。
    /// </summary>
    public sealed partial class ExcelExtensionsExecutor : IToolExecutor
    {
        private static readonly Dictionary<string, Func<Excel.Application, JObject, OperationResult>> Handlers =
            new Dictionary<string, Func<Excel.Application, JObject, OperationResult>>(StringComparer.Ordinal)
            {
                // P1
                { "excel_search_all", ExecuteSearchAll },
                { "excel_error_scan", ExecuteErrorScan },
                { "excel_regex_extract", ExecuteRegexExtract },
                { "excel_rmb_upper", ExecuteRmbUpper },
                { "excel_cross_sheet_consolidate", ExecuteCrossSheetConsolidate },
                { "excel_diff_sheets", ExecuteDiffSheets },
                { "excel_sheet_toc", ExecuteSheetToc },
                { "excel_numbers_as_text_fix", ExecuteNumbersAsTextFix },
                // P2
                { "excel_export", ExecuteExport },
                { "excel_import", ExecuteImport },
                { "excel_named_range_manage", ExecuteNamedRangeManage },
                { "excel_data_validation_manage", ExecuteDataValidationManage },
                { "excel_table_manage", ExecuteTableManage },
                { "excel_freeze_and_view", ExecuteFreezeAndView },
                { "excel_formula_audit", ExecuteFormulaAudit },
                { "excel_print_setup", ExecutePrintSetup },
                { "excel_goal_seek", ExecuteGoalSeek },
                { "excel_batch_rename_sheets", ExecuteBatchRenameSheets },
                { "excel_comment_manage", ExecuteCommentManage },
                // P3（只读清单）
                { "excel_hidden_inventory", ExecuteHiddenInventory },
                { "excel_macro_list", ExecuteMacroList },
                // P3（补充操作）
                { "excel_sparkline_manage", ExecuteSparklineManage },
                { "excel_camera_picture", ExecuteCameraPicture },
                { "excel_powerquery_refresh", ExecutePowerQueryRefresh }
            };

        public string AppType
        {
            get { return "excel"; }
        }

        public bool CanExecute(string toolName)
        {
            if (string.IsNullOrWhiteSpace(toolName))
            {
                return false;
            }
            return Handlers.ContainsKey(toolName);
        }

        public ToolResult Execute(string toolName, JObject arguments)
        {
            if (!CanExecute(toolName))
            {
                return ToolResult.Fail(
                    "INVALID_PARAMS",
                    "Excel 扩展执行器不支持工具「" + (toolName ?? string.Empty) + "」。",
                    false);
            }

            JObject normalized = ExcelArgumentReader.Normalize(arguments);
            OperationResult result;
            try
            {
                Func<Excel.Application, JObject, OperationResult> handler = Handlers[toolName];
                result = OfficeSession.RunExcel(
                    toolName,
                    delegate(Excel.Application application)
                    {
                        return handler(application, normalized);
                    });
            }
            catch (Exception ex)
            {
                result = OfficeErrorMapper.ToOperationResult(ex, OfficeApp.Excel);
            }

            if (result == null)
            {
                return ToolResult.Fail(
                    "INTERNAL_ERROR",
                    "Excel 扩展工具「" + toolName + "」未返回结果。",
                    false);
            }
            return result.ToToolResult();
        }

        internal static OperationResult Fail(OfficeErrorCode code, string message)
        {
            return OperationResult.Failed(code, message);
        }

        internal static OperationResult Fail(string errorCode, string message)
        {
            return OperationResult.Failed(errorCode, message);
        }

        /// <summary>只读工具统一附带 changed=false 的 Observation，满足观察契约。</summary>
        internal static Observation ReadOnlyObservation(string kind, string summary)
        {
            return Observation.Create(kind, summary, false);
        }

        internal static int ClampInt(int value, int min, int max)
        {
            if (value < min)
            {
                return min;
            }
            if (value > max)
            {
                return max;
            }
            return value;
        }
    }
}
