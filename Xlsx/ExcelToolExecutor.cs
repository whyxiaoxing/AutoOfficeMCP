using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Office;
using Office.Contracts;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>
    /// Excel MCP 工具执行器。工具分发表是唯一的工具名事实源：
    /// 描述符、CanExecute、Execute 均引用同一组键，防止清单漂移。
    /// 会话型工具统一走 OfficeSession.RunExcel；多文件只读查询使用独立隐藏实例。
    /// </summary>
    public sealed partial class ExcelToolExecutor : IToolExecutor, ILocalToolExecutor
    {
        private static readonly Dictionary<string, Func<Excel.Application, JObject, OperationResult>> SessionHandlers =
            new Dictionary<string, Func<Excel.Application, JObject, OperationResult>>(StringComparer.Ordinal)
            {
                { "excel_write_data", ExecuteWriteData },
                { "excel_apply_formula", ExecuteApplyFormula },
                { "excel_format_range", ExecuteFormatRange },
                { "excel_create_chart", ExecuteCreateChart },
                { "excel_clean_data", ExecuteCleanData },
                { "excel_clear_range", ExecuteClearRange },
                { "excel_sort_data", ExecuteSortData },
                { "excel_filter_data", ExecuteFilterData },
                { "excel_remove_duplicates", ExecuteRemoveDuplicates },
                { "excel_conditional_format", ExecuteConditionalFormat },
                { "excel_merge_cells", ExecuteMergeCells },
                { "excel_autofit", ExecuteAutoFit },
                { "excel_find_replace", ExecuteFindReplace },
                { "excel_create_pivot_table", ExecuteCreatePivotTable },
                { "excel_create_sheet", ExecuteCreateSheet },
                { "excel_delete_sheet", ExecuteDeleteSheet },
                { "excel_rename_sheet", ExecuteRenameSheet },
                { "excel_copy_sheet", ExecuteCopySheet },
                { "excel_insert_row_col", ExecuteInsertRowCol },
                { "excel_delete_row_col", ExecuteDeleteRowCol },
                { "excel_hide_row_col", ExecuteHideRowCol },
                { "excel_protect_sheet", ExecuteProtectSheet },
                { "excel_transform_data", ExecuteTransformData },
                { "excel_analyze_data", ExecuteAnalyzeData },
                { "excel_generate_report", ExecuteGenerateReport },
                { "excel_run_vba", ExecuteRunVba },
                { "excel_get_context", ExecuteGetContext },
                { "excel_read_range", ExecuteReadRange },
                { "excel_get_selection", ExecuteGetSelection },
                { "excel_list_sheets", ExecuteListSheets },
                { "excel_range_profile", ExecuteRangeProfile },
                { "excel_evaluate", ExecuteEvaluate },
                { "excel_recalculate", ExecuteRecalculate },
                { "excel_formula_reference_audit", ExecuteFormulaReferenceAudit },
                { "excel_precision_check", ExecutePrecisionCheck },
                { "excel_paste_special", ExecutePasteSpecial },
                { "excel_formula_to_value", ExecuteFormulaToValue },
                { "excel_copy_format", ExecuteCopyFormat },
                { "excel_hyperlink_manage", ExecuteHyperlinkManage },
                { "excel_list_merged_cells", ExecuteListMergedCells },
                { "excel_create_workbook", ExecuteCreateWorkbook },
                { "excel_save_as", ExecuteSaveAs },
                { "excel_get_protection_status", ExecuteGetProtectionStatus },
                { "excel_protect_workbook", ExecuteProtectWorkbook },
                { "excel_list_charts", ExecuteListCharts },
                { "excel_update_chart", ExecuteUpdateChart },
                { "excel_delete_chart", ExecuteDeleteChart }
            };

        private static readonly Dictionary<string, Func<JObject, OperationResult>> LocalHandlers =
            new Dictionary<string, Func<JObject, OperationResult>>(StringComparer.Ordinal)
            {
                { "excel_function_catalog", ExcelFunctionCatalog.Query },
                { "excel_profile_files", ExcelExternalWorkbookService.ProfileFiles },
                { "excel_query_files", ExcelExternalWorkbookService.QueryFiles }
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
            return SessionHandlers.ContainsKey(toolName) || LocalHandlers.ContainsKey(toolName);
        }

        /// <summary>
        /// 纯本地工具：函数目录；以及 backend=openxml 的多文件只读查询。
        /// 这些工具不要求 STA/file，固定进程 worker 模式下可在 Gateway 直接执行。
        /// </summary>
        public bool IsLocalTool(string toolName, JObject arguments)
        {
            if (string.IsNullOrWhiteSpace(toolName))
            {
                return false;
            }
            if (string.Equals(toolName, "excel_function_catalog", StringComparison.Ordinal))
            {
                return true;
            }
            if (string.Equals(toolName, "excel_profile_files", StringComparison.Ordinal)
                || string.Equals(toolName, "excel_query_files", StringComparison.Ordinal))
            {
                string backend = arguments == null ? string.Empty : ReadString(arguments, "backend", "readBackend");
                return string.Equals(backend, "openxml", StringComparison.OrdinalIgnoreCase);
            }
            return false;
        }

        private static string ReadString(JObject arguments, params string[] names)
        {
            if (arguments == null || names == null)
            {
                return string.Empty;
            }
            foreach (string name in names)
            {
                JToken token = arguments[name];
                if (token != null && token.Type == JTokenType.String)
                {
                    string value = token.Value<string>();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        return value.Trim();
                    }
                }
            }
            return string.Empty;
        }

        public ToolResult Execute(string toolName, JObject arguments)
        {
            if (!CanExecute(toolName))
            {
                return ToolResult.Fail(
                    "INVALID_PARAMS",
                    "Excel 执行器不支持工具「" + (toolName ?? string.Empty) + "」。",
                    false);
            }

            JObject normalized = ExcelArgumentReader.Normalize(arguments);
            OperationResult result;
            Func<Excel.Application, JObject, OperationResult> sessionHandler;
            if (SessionHandlers.TryGetValue(toolName, out sessionHandler))
            {
                try
                {
                    result = OfficeSession.RunExcel(
                        toolName,
                        delegate(Excel.Application application)
                        {
                            return sessionHandler(application, normalized);
                        });
                }
                catch (Exception ex)
                {
                    result = OfficeErrorMapper.ToOperationResult(ex, OfficeApp.Excel);
                }
            }
            else
            {
                Func<JObject, OperationResult> localHandler = LocalHandlers[toolName];
                try
                {
                    result = localHandler(normalized);
                }
                catch (Exception ex)
                {
                    result = OfficeErrorMapper.ToOperationResult(ex, OfficeApp.Excel);
                }
            }

            if (result == null)
            {
                return ToolResult.Fail(
                    "INTERNAL_ERROR",
                    "Excel 工具「" + toolName + "」未返回结果。",
                    false);
            }
            return result.ToToolResult();
        }

        internal static OperationResult Fail(OfficeErrorCode code, string message)
        {
            return OperationResult.Failed(code, message);
        }

    }
}
