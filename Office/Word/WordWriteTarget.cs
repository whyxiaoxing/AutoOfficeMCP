using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    /// <summary>一次写入操作的目标区段快照（使用字符位置，不跨调用持有 Range）。</summary>
    internal sealed class WordWriteTarget
    {
        public int Start { get; private set; }

        public int End { get; private set; }

        public string Label { get; private set; }

        public bool SnapshotCapable { get; private set; }

        public bool Structural { get; private set; }

        public List<string> TargetRefs { get; private set; }

        public WordWriteTarget(
            int start,
            int end,
            string label,
            string targetRef,
            bool snapshotCapable,
            bool structural)
        {
            Start = start;
            End = end;
            Label = label ?? "Word 目标区段";
            SnapshotCapable = snapshotCapable;
            Structural = structural;
            TargetRefs = new List<string>();
            if (!string.IsNullOrWhiteSpace(targetRef))
            {
                TargetRefs.Add(targetRef);
            }
        }

        public WordWriteTarget AddTargetRef(string targetRef)
        {
            if (!string.IsNullOrWhiteSpace(targetRef))
            {
                TargetRefs.Add(targetRef);
            }
            return this;
        }
    }

    /// <summary>常用写入目标选择器；只返回位置，不持有 COM Range。</summary>
    internal static class WordWriteTargets
    {
        public static WordWriteTarget Selection(
            Word.Application application,
            string label,
            string targetRef,
            bool checkPlainText,
            bool structural)
        {
            Word.Selection selection = null;
            Word.Range range = null;
            try
            {
                selection = application.Selection;
                if (selection == null)
                {
                    throw WordArguments.Invalid("当前没有有效 Word 选区，无法确定写入目标。");
                }
                range = selection.Range;
                bool snapshotCapable = !structural && (!checkPlainText || WordComUtil.IsPlainTextRange(range));
                return new WordWriteTarget(
                    range.Start,
                    range.End,
                    label,
                    targetRef,
                    snapshotCapable,
                    structural);
            }
            finally
            {
                ComHelper.Release(range, selection);
            }
        }

        public static WordWriteTarget Document(
            Word.Document document,
            string label,
            string targetRef,
            bool checkPlainText,
            bool structural)
        {
            Word.Range range = null;
            try
            {
                range = document.Content;
                bool snapshotCapable = !structural && (!checkPlainText || WordComUtil.IsPlainTextRange(range));
                return new WordWriteTarget(
                    range.Start,
                    range.End,
                    label,
                    targetRef,
                    snapshotCapable,
                    structural);
            }
            finally
            {
                ComHelper.Release(range);
            }
        }

        public static WordWriteTarget Paragraph(
            Word.Document document,
            int paragraphIndex,
            string targetRef,
            bool structural)
        {
            Word.Paragraph paragraph = null;
            Word.Range range = null;
            try
            {
                paragraph = WordComUtil.GetParagraph(document, paragraphIndex);
                range = paragraph.Range;
                bool snapshotCapable = !structural && WordComUtil.IsPlainTextRange(range);
                return new WordWriteTarget(
                    range.Start,
                    range.End,
                    "第 " + paragraphIndex + " 段",
                    targetRef,
                    snapshotCapable,
                    structural);
            }
            finally
            {
                ComHelper.Release(range, paragraph);
            }
        }

        public static WordWriteTarget DocumentStart()
        {
            return new WordWriteTarget(0, 0, "文档开头", "Word:DocumentStart", false, true);
        }

        public static WordWriteTarget DocumentEnd(Word.Document document)
        {
            int end = Math.Max(0, WordComUtil.GetContentEnd(document) - 1);
            return new WordWriteTarget(end, end, "文档末尾", "Word:DocumentEnd", false, true);
        }

        public static WordWriteTarget Cursor(Word.Application application)
        {
            return Selection(application, "当前光标位置", "Word:Selection", false, false);
        }

        public static WordWriteTarget Table(Word.Document document, int tableIndex)
        {
            Word.Table table = null;
            Word.Range range = null;
            try
            {
                table = WordComUtil.GetTable(document, tableIndex);
                range = table.Range;
                return new WordWriteTarget(
                    range.Start,
                    range.End,
                    "第 " + tableIndex + " 个表格",
                    "Word:Table:" + tableIndex,
                    false,
                    true);
            }
            finally
            {
                ComHelper.Release(range, table);
            }
        }

        public static WordWriteTarget None(string label, string targetRef)
        {
            return new WordWriteTarget(-1, -1, label, targetRef, false, true);
        }
    }
}
