using System;
using System.Collections.Generic;

namespace Office.Parallel
{
    /// <summary>
    /// 池内作业的静态分类规则：
    /// 1) appType 只接受 excel / word / ppt；
    /// 2) 工具名前缀必须与 appType 一致（拒绝把 Word 工具投到 Excel 实例）；
    /// 3) 明确只读的工具以 ReadOnly=true 打开文件，其余按可读写打开。
    ///
    /// 说明：这里只做保守分类。像 excel_evaluate、excel_analyze_data 这类
    /// 带可选 writeTarget 的"读"工具按可读写处理，避免在只读工作簿上执行写入。
    /// </summary>
    internal static class ParallelToolCatalog
    {
        public const string Excel = "excel";
        public const string Word = "word";
        public const string Ppt = "ppt";
        public const string SavePolicyNone = "none";
        public const string SavePolicySave = "save";

        private static readonly HashSet<string> ExcelReadOnlyTools = new HashSet<string>(StringComparer.Ordinal)
        {
            "excel_diff_sheets",
            "excel_error_scan",
            "excel_formula_audit",
            "excel_get_context",
            "excel_get_selection",
            "excel_hidden_inventory",
            "excel_list_sheets",
            "excel_macro_list",
            "excel_range_profile",
            "excel_read_range",
            "excel_search_all",
            "excel_function_catalog",
            "excel_profile_files"
        };

        private static readonly HashSet<string> WordReadOnlyTools = new HashSet<string>(StringComparer.Ordinal)
        {
            "word_comments_read",
            "word_document_stats",
            "word_equations_list",
            "word_extract_images",
            "word_find_and_select",
            "word_get_context",
            "word_get_paragraph_info",
            "word_hidden_text_report",
            "word_images_list",
            "word_list_paragraphs",
            "word_md_export",
            "word_read_selection",
            "word_reading_time",
            "word_repeat_check",
            "word_revisions_list",
            "word_style_audit",
            "word_table_read",
            "word_tables_list"
        };

        private static readonly HashSet<string> PptReadOnlyTools = new HashSet<string>(StringComparer.Ordinal)
        {
            "ppt_get_presentation_info",
            "ppt_get_slide_content",
            "ppt_list_shapes",
            "ppt_export_images",
            "ppt_export_pdf",
            "ppt_get_thumbnail",
            "ppt_list_layouts",
            "ppt_list_themes",
            "ppt_list_animations",
            "ppt_get_transition",
            "ppt_get_table",
            "ppt_validate_layout"
        };

        public static bool IsSupportedApp(string appType)
        {
            string normalized = NormalizeApp(appType);
            return normalized.Length > 0;
        }

        public static string NormalizeApp(string appType)
        {
            if (string.IsNullOrWhiteSpace(appType))
            {
                return string.Empty;
            }

            string normalized = appType.Trim().ToLowerInvariant();
            if (string.Equals(normalized, Excel, StringComparison.Ordinal)
                || string.Equals(normalized, Word, StringComparison.Ordinal)
                || string.Equals(normalized, Ppt, StringComparison.Ordinal))
            {
                return normalized;
            }
            return string.Empty;
        }

        /// <summary>工具名是否属于指定应用（excel_ / word_ 前缀）。</summary>
        public static bool IsToolForApp(string appType, string toolName)
        {
            string app = NormalizeApp(appType);
            if (app.Length == 0 || string.IsNullOrWhiteSpace(toolName))
            {
                return false;
            }

            string tool = toolName.Trim();
            if (string.Equals(app, Excel, StringComparison.Ordinal))
            {
                return tool.StartsWith("excel_", StringComparison.Ordinal) && tool.Length > "excel_".Length;
            }
            if (string.Equals(app, Ppt, StringComparison.Ordinal))
            {
                return tool.StartsWith("ppt_", StringComparison.Ordinal) && tool.Length > "ppt_".Length;
            }
            return tool.StartsWith("word_", StringComparison.Ordinal) && tool.Length > "word_".Length;
        }

        public static bool IsReadOnlyTool(string appType, string toolName)
        {
            string app = NormalizeApp(appType);
            if (app.Length == 0 || string.IsNullOrWhiteSpace(toolName))
            {
                return false;
            }

            string tool = toolName.Trim();
            if (string.Equals(app, Excel, StringComparison.Ordinal))
            {
                return ExcelReadOnlyTools.Contains(tool);
            }
            if (string.Equals(app, Ppt, StringComparison.Ordinal))
            {
                return PptReadOnlyTools.Contains(tool);
            }
            return WordReadOnlyTools.Contains(tool);
        }

        /// <summary>返回 none / save；无法识别时返回 null（调用方按 INVALID_PARAMS 拒绝）。</summary>
        public static string NormalizeSavePolicy(string savePolicy)
        {
            if (string.IsNullOrWhiteSpace(savePolicy))
            {
                return SavePolicyNone;
            }

            string normalized = savePolicy.Trim().ToLowerInvariant();
            if (string.Equals(normalized, SavePolicyNone, StringComparison.Ordinal))
            {
                return SavePolicyNone;
            }
            if (string.Equals(normalized, SavePolicySave, StringComparison.Ordinal))
            {
                return SavePolicySave;
            }
            return null;
        }
    }
}
