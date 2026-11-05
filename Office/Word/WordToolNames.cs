using System;
using System.Collections.Generic;

namespace Office.WordTools
{
    /// <summary>
    /// Word 执行器认领的 53 个 MCP 工具名。
    /// 这里必须与 Tools\word\*.json 描述符、Execute 的处理器表三处保持一致，
    /// 避免参考项目里"能力清单漂移导致 Agent 空转重试"的问题。
    /// </summary>
    internal static class WordToolNames
    {
        public const string GetContext = "word_get_context";
        public const string ListParagraphs = "word_list_paragraphs";
        public const string GetParagraphInfo = "word_get_paragraph_info";
        public const string FindAndSelect = "word_find_and_select";
        public const string ReadSelection = "word_read_selection";
        public const string DocumentStats = "word_document_stats";

        public const string InsertText = "word_insert_text";
        public const string FormatText = "word_format_text";
        public const string ReplaceText = "word_replace_text";
        public const string DeleteText = "word_delete_text";
        public const string ApplyStyle = "word_apply_style";
        public const string SetParagraphFormat = "word_set_paragraph_format";
        public const string InsertTable = "word_insert_table";
        public const string GenerateToc = "word_generate_toc";
        public const string Beautify = "word_beautify";

        public const string RebuildNumbering = "word_rebuild_numbering";
        public const string SetPageSetup = "word_set_page_setup";
        public const string HeaderFooter = "word_header_footer";
        public const string InsertImage = "word_insert_image";
        public const string Export = "word_export";
        public const string DocProperties = "word_doc_properties";
        public const string CommentAdd = "word_comment_add";
        public const string CommentsRead = "word_comments_read";
        public const string CommentResolve = "word_comment_resolve";
        public const string CommentDelete = "word_comment_delete";

        public const string PageMap = "word_page_map";
        public const string GotoPage = "word_goto_page";
        public const string StylesList = "word_styles_list";
        public const string StyleModify = "word_style_modify";
        public const string FieldsList = "word_fields_list";

        public const string TrackChangesToggle = "word_track_changes_toggle";
        public const string RevisionsList = "word_revisions_list";
        public const string RevisionAccept = "word_revision_accept";
        public const string RevisionReject = "word_revision_reject";
        public const string ApplyAsRevision = "word_apply_as_revision";
        public const string CompareDocuments = "word_compare_documents";
        public const string FieldUpdateAll = "word_field_update_all";

        public const string ClearFormat = "word_clear_format";
        public const string InsertEquation = "word_insert_equation";
        public const string EquationsList = "word_equations_list";
        public const string TableFormula = "word_table_formula";
        public const string TablesList = "word_tables_list";
        public const string TableRead = "word_table_read";
        public const string TableSetCell = "word_table_set_cell";
        public const string TableSetRow = "word_table_set_row";
        public const string TableSetCol = "word_table_set_col";
        public const string TableAddRowCol = "word_table_add_row_col";
        public const string TableDeleteRowCol = "word_table_delete_row_col";
        public const string TableMergeCells = "word_table_merge_cells";
        public const string TableFormat = "word_table_format";
        public const string ImagesList = "word_images_list";
        public const string ImageResize = "word_image_resize";
        public const string ImageLayout = "word_image_layout";

        private static readonly string[] ReadToolNames = new string[]
        {
            GetContext,
            ListParagraphs,
            GetParagraphInfo,
            FindAndSelect,
            ReadSelection,
            DocumentStats,
            CommentsRead,
            RevisionsList,
            EquationsList,
            TablesList,
            TableRead,
            ImagesList,
            PageMap,
            GotoPage,
            StylesList,
            FieldsList
        };

        private static readonly string[] WriteToolNames = new string[]
        {
            InsertText,
            FormatText,
            ReplaceText,
            DeleteText,
            ApplyStyle,
            SetParagraphFormat,
            InsertTable,
            GenerateToc,
            Beautify,
            RebuildNumbering,
            SetPageSetup,
            HeaderFooter,
            InsertImage,
            Export,
            DocProperties,
            CommentAdd,
            TrackChangesToggle,
            RevisionAccept,
            RevisionReject,
            ApplyAsRevision,
            CompareDocuments,
            FieldUpdateAll,
            ClearFormat,
            InsertEquation,
            TableFormula,
            TableSetCell,
            TableSetRow,
            TableSetCol,
            TableAddRowCol,
            TableDeleteRowCol,
            TableMergeCells,
            TableFormat,
            ImageResize,
            ImageLayout,
            CommentResolve,
            CommentDelete,
            StyleModify
        };

        private static readonly HashSet<string> ReadSet =
            new HashSet<string>(ReadToolNames, StringComparer.Ordinal);

        private static readonly HashSet<string> WriteSet =
            new HashSet<string>(WriteToolNames, StringComparer.Ordinal);

        private static readonly HashSet<string> AllSet = BuildAllSet();

        public static int Count
        {
            get { return AllSet.Count; }
        }

        public static bool Contains(string toolName)
        {
            return !string.IsNullOrWhiteSpace(toolName) && AllSet.Contains(toolName.Trim());
        }

        public static bool IsReadOnly(string toolName)
        {
            return !string.IsNullOrWhiteSpace(toolName) && ReadSet.Contains(toolName.Trim());
        }

        public static bool IsWrite(string toolName)
        {
            return !string.IsNullOrWhiteSpace(toolName) && WriteSet.Contains(toolName.Trim());
        }

        public static string[] GetAllNames()
        {
            string[] names = new string[AllSet.Count];
            AllSet.CopyTo(names);
            Array.Sort(names, StringComparer.Ordinal);
            return names;
        }

        private static HashSet<string> BuildAllSet()
        {
            HashSet<string> all = new HashSet<string>(StringComparer.Ordinal);
            foreach (string name in ReadToolNames)
            {
                all.Add(name);
            }
            foreach (string name in WriteToolNames)
            {
                if (!all.Add(name))
                {
                    throw new InvalidOperationException("Word 工具名重复：" + name);
                }
            }
            return all;
        }
    }
}
