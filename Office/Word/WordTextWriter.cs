using System;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    /// <summary>
    /// 大文本写入与查找替换的 COM 细节：
    /// - 超过 32K 字符按 30000 字符分块；
    /// - 光标插入三级降级 TypeText → Range.Text → InsertAfter；
    /// - 文档开头/末尾使用带位置跟踪的 Range 分块写入，保证顺序与长度可回读。
    /// </summary>
    internal static class WordTextWriter
    {
        public const int ChunkSize = 30000;

        public static void TypeText(Word.Selection selection, string text)
        {
            if (selection == null)
            {
                throw WordArguments.Invalid("当前没有有效选区，无法在光标处写入文本。");
            }
            string normalized = WordFormatUtil.NormalizeForWord(text);
            int position = 0;
            while (position < normalized.Length)
            {
                int length = Math.Min(ChunkSize, normalized.Length - position);
                string chunk = normalized.Substring(position, length);
                WriteChunkAtSelection(selection, chunk);
                position += length;
            }
        }

        public static int InsertAtStart(Word.Document document, string text)
        {
            string normalized = WordFormatUtil.NormalizeForWord(text);
            int position = 0;
            int offset = 0;
            while (offset < normalized.Length)
            {
                int length = Math.Min(ChunkSize, normalized.Length - offset);
                string chunk = normalized.Substring(offset, length);
                InsertChunk(document, position, chunk, true);
                position += length;
                offset += length;
            }
            return normalized.Length;
        }

        public static int InsertAtEnd(Word.Document document, string text)
        {
            string normalized = WordFormatUtil.NormalizeForWord(text);
            int offset = 0;
            while (offset < normalized.Length)
            {
                int length = Math.Min(ChunkSize, normalized.Length - offset);
                string chunk = normalized.Substring(offset, length);
                int position = Math.Max(0, WordComUtil.GetContentEnd(document) - 1);
                InsertChunk(document, position, chunk, false);
                offset += length;
            }
            return normalized.Length;
        }

        public static bool ReplaceAll(
            Word.Find find,
            string findText,
            string replaceText,
            bool matchCase,
            bool matchWholeWord)
        {
            if (find == null)
            {
                throw WordArguments.Invalid("Word Find 对象不可用，无法执行查找替换。");
            }

            find.ClearFormatting();
            find.Replacement.ClearFormatting();
            find.Text = findText;
            find.Replacement.Text = replaceText ?? string.Empty;
            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindContinue;
            find.MatchCase = matchCase;
            find.MatchWholeWord = matchWholeWord;
            find.MatchWildcards = false;
            find.Format = false;

            object missing = Type.Missing;
            object replaceAll = Word.WdReplace.wdReplaceAll;
            return find.Execute(
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref replaceAll, ref missing, ref missing, ref missing, ref missing);
        }

        public static bool FindFirst(Word.Find find, string findText)
        {
            if (find == null)
            {
                throw WordArguments.Invalid("Word Find 对象不可用，无法定位文本。");
            }
            find.ClearFormatting();
            find.Text = findText;
            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            find.MatchCase = false;
            find.MatchWholeWord = false;
            find.MatchWildcards = false;
            find.Format = false;

            object missing = Type.Missing;
            return find.Execute(
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing);
        }

        private static void WriteChunkAtSelection(Word.Selection selection, string chunk)
        {
            try
            {
                selection.TypeText(chunk);
                return;
            }
            catch (Exception firstFailure)
            {
                Console.Error.WriteLine("[word] TypeText 失败，降级 Range.Text：" + firstFailure.Message);
            }

            Word.Range range = null;
            try
            {
                range = selection.Range;
                range.Text = chunk;
                try
                {
                    selection.SetRange(range.End, range.End);
                }
                catch
                {
                    // 降级路径下 selection 可能已经自动折叠；SetRange 失败不影响已写入内容。
                }
                return;
            }
            catch (Exception secondFailure)
            {
                Console.Error.WriteLine("[word] Range.Text 失败，降级 InsertAfter：" + secondFailure.Message);
            }
            finally
            {
                ComHelper.Release(range);
            }

            Word.Range fallbackRange = null;
            try
            {
                fallbackRange = selection.Range;
                fallbackRange.InsertAfter(chunk);
            }
            catch (Exception thirdFailure)
            {
                throw new OfficeComException(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "文本写入三级降级均失败（TypeText → Range.Text → InsertAfter）：" +
                    thirdFailure.Message);
            }
            finally
            {
                ComHelper.Release(fallbackRange);
            }
        }

        private static void InsertChunk(Word.Document document, int position, string chunk, bool atStart)
        {
            Word.Range range = null;
            try
            {
                range = WordComUtil.CreateRange(document, position, position);
                if (atStart)
                {
                    range.InsertAfter(chunk);
                }
                else
                {
                    range.InsertAfter(chunk);
                }
            }
            catch (Exception firstFailure)
            {
                Console.Error.WriteLine("[word] Range.InsertAfter 失败，降级 Range.Text：" + firstFailure.Message);
                ComHelper.Release(range);
                range = null;
                try
                {
                    range = WordComUtil.CreateRange(document, position, position);
                    range.Text = chunk;
                }
                catch (Exception secondFailure)
                {
                    throw new OfficeComException(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "在文档" + (atStart ? "开头" : "末尾") + "写入文本失败（InsertAfter → Range.Text）：" +
                        secondFailure.Message);
                }
            }
            finally
            {
                ComHelper.Release(range);
            }
        }
    }
}
