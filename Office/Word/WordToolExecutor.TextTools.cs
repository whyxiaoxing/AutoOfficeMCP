using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    public sealed partial class WordToolExecutor
    {
        private void RegisterTextHandlers(
            Dictionary<string, Func<Word.Application, JObject, OperationResult>> handlers)
        {
            handlers[WordToolNames.InsertText] = HandleInsertText;
            handlers[WordToolNames.FormatText] = HandleFormatText;
            handlers[WordToolNames.ReplaceText] = HandleReplaceText;
            handlers[WordToolNames.DeleteText] = HandleDeleteText;
            handlers[WordToolNames.ApplyStyle] = HandleApplyStyle;
            handlers[WordToolNames.SetParagraphFormat] = HandleSetParagraphFormat;
            handlers[WordToolNames.ClearFormat] = HandleClearFormat;
            handlers[WordToolNames.ApplyAsRevision] = HandleApplyAsRevision;
        }

        private OperationResult HandleInsertText(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string content = args.RequireText("content");
            string position = (args.GetOptionalString("position", "cursor") ?? "cursor").Trim().ToLowerInvariant();
            if (position != "cursor" && position != "start" && position != "end")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "position 只支持 cursor/start/end，收到 " + position + "。");
            }

            string normalized = WordFormatUtil.NormalizeForWord(content);
            return RunWriteTool(
                application,
                WordToolNames.InsertText,
                false,
                delegate(Word.Document document)
                {
                    if (position == "start")
                    {
                        return new WordWriteTarget(0, 0, "文档开头", "Word:DocumentStart", true, false);
                    }
                    if (position == "end")
                    {
                        int end = Math.Max(0, WordComUtil.GetContentEnd(document) - 1);
                        return new WordWriteTarget(end, end, "文档末尾", "Word:DocumentEnd", true, false);
                    }
                    return WordWriteTargets.Selection(application, "当前光标位置", "Word:Selection", false, false);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Selection selection = null;
                    Word.Range verificationRange = null;
                    try
                    {
                        int afterStart = target.Start;
                        int afterEnd = target.End;
                        if (position == "start")
                        {
                            int length = WordTextWriter.InsertAtStart(document, content);
                            afterStart = 0;
                            afterEnd = length;
                        }
                        else if (position == "end")
                        {
                            int beforeEnd = WordComUtil.GetContentEnd(document);
                            int length = WordTextWriter.InsertAtEnd(document, content);
                            afterStart = Math.Max(0, beforeEnd - 1);
                            afterEnd = afterStart + length;
                        }
                        else
                        {
                            selection = application.Selection;
                            WordTextWriter.TypeText(selection, content);
                            Word.Range selectionRange = null;
                            try
                            {
                                selectionRange = selection.Range;
                                afterEnd = selectionRange.End;
                                afterStart = Math.Max(0, afterEnd - normalized.Length);
                            }
                            finally
                            {
                                ComHelper.Release(selectionRange);
                            }
                        }

                        string actualText = string.Empty;
                        try
                        {
                            verificationRange = WordComUtil.CreateRange(document, afterStart, afterEnd);
                            actualText = verificationRange.Text ?? string.Empty;
                        }
                        finally
                        {
                            ComHelper.Release(verificationRange);
                        }

                        bool verified = string.Equals(
                            WordFormatUtil.NormalizeForCompare(actualText),
                            WordFormatUtil.NormalizeForCompare(normalized),
                            StringComparison.Ordinal);

                        JObject data = new JObject();
                        data["position"] = position;
                        data["requestedCharacters"] = normalized.Length;
                        data["insertedRangeStart"] = afterStart;
                        data["insertedRangeEnd"] = afterEnd;
                        data["verified"] = verified;
                        data["actualPreview"] = WordComUtil.NormalizePreview(actualText, 160);

                        if (!verified)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.PARTIAL_APPLY,
                                "文本写入后回读不一致：期望 " + normalized.Length + " 个字符，实际回读「" +
                                WordComUtil.NormalizePreview(actualText, 80) + "」。请检查选区是否仍有效、文档是否受保护。",
                                data,
                                true);
                        }

                        string message = position == "start"
                            ? "已在文档开头插入 " + normalized.Length + " 个字符。"
                            : (position == "end"
                                ? "已在文档末尾追加 " + normalized.Length + " 个字符（Content.InsertAfter，不额外新建段落）。"
                                : "已在光标处写入 " + normalized.Length + " 个字符。");
                        WordWriteResult result = WordWriteResult.Ok(message, data, normalized.Length > 0);
                        result.WithAfterRange(afterStart, afterEnd);
                        result.AddTargetRef(position == "cursor" ? "Word:Selection" : "Word:Document");
                        return result;
                    }
                    finally
                    {
                        ComHelper.Release(verificationRange, selection);
                    }
                });
        }

        private OperationResult HandleFormatText(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            WordFontRequest request = new WordFontRequest();
            request.Bold = ReadNullableBool(args, "bold");
            request.Italic = ReadNullableBool(args, "italic");
            request.UnderlineStyle = args.Has("underlineStyle")
                ? args.GetOptionalString("underlineStyle", null)
                : null;
            request.UnderlineEnabled = ReadNullableBool(args, "underline");
            request.Subscript = ReadNullableBool(args, "subscript");
            request.Superscript = ReadNullableBool(args, "superscript");
            request.StrikeThrough = ReadNullableBool(args, "strikethrough");
            request.DoubleStrikeThrough = ReadNullableBool(args, "doubleStrikethrough");
            request.SmallCaps = ReadNullableBool(args, "smallCaps");
            request.AllCaps = ReadNullableBool(args, "allCaps");
            request.FontSize = args.GetOptionalDouble("fontSize");
            request.FontName = args.Has("fontName") ? args.GetOptionalString("fontName", null) : null;
            request.CharacterSpacing = args.GetOptionalDouble("characterSpacing");
            if (args.Has("color"))
            {
                request.Color = WordFormatUtil.ParseColor(args.GetOptionalString("color", null));
            }
            if (args.Has("highlightColor"))
            {
                request.HighlightColor = WordFormatUtil.ParseHighlight(args.GetOptionalString("highlightColor", null));
            }
            if (request.Subscript == true && request.Superscript == true)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "subscript 与 superscript 不能同时为 true。");
            }
            if (!request.HasAny)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "word_format_text 至少需要一个格式属性。收到 keys=" + args.DescribeKeys() +
                    "；支持 bold/italic/underlineStyle/subscript/superscript/strikethrough/" +
                    "doubleStrikethrough/fontSize/fontName/color/highlightColor/smallCaps/allCaps/characterSpacing");
            }

            string rangeMode = (args.GetOptionalString("range", "selection") ?? "selection")
                .Trim().ToLowerInvariant();
            if (rangeMode != "selection" && rangeMode != "all" && rangeMode != "document" && rangeMode != "全文")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "range 只支持 selection/all，收到 " + rangeMode + "。");
            }
            bool all = rangeMode == "all" || rangeMode == "document" || rangeMode == "全文";
            bool snapshotCapable = !request.HighlightColor.HasValue;

            return RunWriteTool(
                application,
                WordToolNames.FormatText,
                false,
                delegate(Word.Document document)
                {
                    if (all)
                    {
                        Word.Range content = null;
                        try
                        {
                            content = document.Content;
                            return new WordWriteTarget(
                                content.Start, content.End, "全文", "Word:Document",
                                snapshotCapable, false);
                        }
                        finally
                        {
                            ComHelper.Release(content);
                        }
                    }
                    Word.Selection selection = null;
                    Word.Range range = null;
                    try
                    {
                        selection = application.Selection;
                        range = selection == null ? null : selection.Range;
                        if (range == null)
                        {
                            throw WordArguments.Invalid("当前没有有效选区，无法设置字体格式。");
                        }
                        return new WordWriteTarget(
                            range.Start, range.End, "当前选区", "Word:Selection",
                            snapshotCapable, false);
                    }
                    finally
                    {
                        ComHelper.Release(range, selection);
                    }
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Selection selection = null;
                    Word.Range range = null;
                    try
                    {
                        if (all)
                        {
                            range = document.Content;
                        }
                        else
                        {
                            selection = application.Selection;
                            range = selection == null ? null : selection.Range;
                        }
                        if (range == null)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "当前没有有效选区，range=selection 无法设置格式；可改用 range=all。",
                                null,
                                false);
                        }

                        WordFontApplyResult applied = WordFontApplier.ApplyAndVerify(range, request);
                        JObject data = new JObject();
                        data["range"] = all ? "all" : "selection";
                        data["rangeLength"] = range.End - range.Start;
                        data["applied"] = applied.Applied;
                        data["verification"] = applied.Verification;
                        data["allMatched"] = applied.AllMatched;

                        if (!applied.AllMatched)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "字体格式回读不一致：" + applied.FirstFailure +
                                "。请根据 verification 中 expected/actual 调整参数或检查选区。",
                                data,
                                applied.Changed);
                        }

                        string message = "已对" + (all ? "全文" : "当前选区") + "设置 " +
                                         (int)applied.Verification.Count + " 项字体格式，回读全部一致。";
                        return WordWriteResult.Ok(message, data, applied.Changed);
                    }
                    finally
                    {
                        ComHelper.Release(range, selection);
                    }
                });
        }

        private OperationResult HandleReplaceText(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string findText = args.RequireString("find");
            string replaceText = args.Has("replace") ? args.GetOptionalString("replace", string.Empty) : string.Empty;
            bool matchCase = args.GetOptionalBool("matchCase", false);
            bool matchWholeWord = args.GetOptionalBool("matchWholeWord", false);
            JObject format = args.GetOptionalObject("format");
            if (format != null)
            {
                return HandleReplaceTextWithFormat(
                    application, findText, replaceText, matchCase, matchWholeWord, format);
            }

            return RunWriteTool(
                application,
                WordToolNames.ReplaceText,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.Document(
                        document, "文档正文", "Word:Document", true, false);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Range content = null;
                    Word.Find find = null;
                    try
                    {
                        content = document.Content;
                        string text = content.Text ?? string.Empty;
                        int expectedCount = WordComUtil.CountOccurrences(
                            text, findText, matchCase, matchWholeWord);
                        find = content.Find;
                        bool executed = WordTextWriter.ReplaceAll(
                            find, findText, replaceText, matchCase, matchWholeWord);

                        // ReplaceAll 返回后回读正文：changed 与替换次数都必须由写后事实支撑，
                        // 不能仅凭 Find.Execute 的返回值宣称成功。
                        string afterText = content.Text ?? string.Empty;
                        bool changed = !string.Equals(text, afterText, StringComparison.Ordinal);
                        bool replacementContainsFind = !string.IsNullOrEmpty(replaceText) &&
                                                       WordComUtil.CountOccurrences(
                                                           replaceText, findText, matchCase, matchWholeWord) > 0;
                        int remaining = changed && !replacementContainsFind
                            ? WordComUtil.CountOccurrences(afterText, findText, matchCase, matchWholeWord)
                            : 0;
                        int actualCount = expectedCount;
                        bool countVerified = false;
                        string countSource = "pre-scan";
                        if (changed)
                        {
                            if (!replacementContainsFind)
                            {
                                actualCount = Math.Max(0, expectedCount - remaining);
                                countVerified = remaining == 0;
                                countSource = countVerified ? "content-diff" : "content-diff-partial";
                            }
                            else
                            {
                                int unit = replaceText.Length - findText.Length;
                                int delta = afterText.Length - text.Length;
                                if (unit != 0 && delta % unit == 0 && delta / unit >= 0)
                                {
                                    actualCount = delta / unit;
                                    countVerified = true;
                                    countSource = "length-delta";
                                }
                                else if (actualCount == 0)
                                {
                                    // 替换文本包含查找文本且长度相同（如大小写变化），无法用
                                    // 内容或长度差独立计数；至少发生了一次真实替换。
                                    actualCount = 1;
                                    countSource = "lower-bound";
                                }
                                else
                                {
                                    countSource = "pre-scan";
                                }
                            }
                        }
                        else
                        {
                            actualCount = 0;
                            countVerified = true;
                            countSource = "content-diff";
                        }

                        JObject data = new JObject();
                        data["find"] = findText;
                        data["replace"] = replaceText;
                        data["matchCase"] = matchCase;
                        data["matchWholeWord"] = matchWholeWord;
                        data["replacementCount"] = actualCount;
                        data["replacementCountVerified"] = countVerified;
                        data["replacementCountSource"] = countSource;
                        data["preScanCount"] = expectedCount;
                        data["remainingMatches"] = remaining;
                        data["replaceAllExecuted"] = executed;
                        data["documentTextLengthBefore"] = text.Length;
                        data["documentTextLengthAfter"] = afterText.Length;

                        if (!changed)
                        {
                            if (executed && expectedCount > 0)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.PARTIAL_APPLY,
                                    "Word 报告 ReplaceAll 已执行，但正文文本前后完全一致；未确认任何替换生效。" +
                                    "请用 word_read_selection 核对实际内容。",
                                    data,
                                    false);
                            }
                            return WordWriteResult.Ok(
                                "未找到匹配文本「" + WordComUtil.Truncate(findText, 40) + "」，文档未修改。",
                                data,
                                false);
                        }

                        if (!executed)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.PARTIAL_APPLY,
                                "正文已发生变化，但 Word ReplaceAll 返回 false；实际替换约 " + actualCount +
                                " 处。请用 word_read_selection 核对实际内容。",
                                data,
                                true);
                        }

                        if (remaining > 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.PARTIAL_APPLY,
                                "ReplaceAll 执行后正文仍残留 " + remaining + " 处匹配「" +
                                WordComUtil.Truncate(findText, 40) + "」，已替换约 " + actualCount +
                                " 处；请核对 matchCase/matchWholeWord 或手动复查。",
                                data,
                                true);
                        }

                        string countNote = countVerified
                            ? string.Empty
                            : "（替换次数未能独立复核，来源：" + countSource + "）";
                        return WordWriteResult.Ok(
                            "已替换 " + actualCount + " 处「" + WordComUtil.Truncate(findText, 30) +
                            "」→「" + WordComUtil.Truncate(replaceText, 30) + "」。" + countNote,
                            data,
                            true);
                    }
                    finally
                    {
                        ComHelper.Release(find, content);
                    }
                },
                true);
        }

        private OperationResult HandleDeleteText(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string rangeMode = (args.GetOptionalString("range", "selection") ?? "selection")
                .Trim().ToLowerInvariant();
            bool all = rangeMode == "all" || rangeMode == "document" || rangeMode == "全文";
            if (!all && rangeMode != "selection")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "range 只支持 selection/all，收到 " + rangeMode + "。");
            }

            return RunWriteTool(
                application,
                WordToolNames.DeleteText,
                false,
                delegate(Word.Document document)
                {
                    if (all)
                    {
                        Word.Range content = null;
                        try
                        {
                            content = document.Content;
                            return new WordWriteTarget(
                                content.Start, content.End, "全文", "Word:Document",
                                WordComUtil.IsPlainTextRange(content), false);
                        }
                        finally
                        {
                            ComHelper.Release(content);
                        }
                    }
                    return WordWriteTargets.Selection(
                        application, "当前选区", "Word:Selection", true, false);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Selection selection = null;
                    Word.Range range = null;
                    try
                    {
                        if (all)
                        {
                            range = document.Content;
                        }
                        else
                        {
                            selection = application.Selection;
                            range = selection == null ? null : selection.Range;
                        }
                        if (range == null)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "当前没有有效选区，range=selection 无法删除；可改用 range=all。",
                                null,
                                false);
                        }

                        string deletedText = range.Text ?? string.Empty;
                        int deletedCharacters = deletedText.Length;
                        if (deletedCharacters > 0)
                        {
                            object missing = Type.Missing;
                            range.Delete(ref missing, ref missing);
                        }

                        JObject data = new JObject();
                        data["range"] = all ? "all" : "selection";
                        data["deletedCharacters"] = deletedCharacters;
                        data["deletedPreview"] = WordComUtil.NormalizePreview(deletedText, 80);

                        WordWriteResult result = deletedCharacters == 0
                            ? WordWriteResult.Ok("目标选区为空，未删除任何内容。", data, false)
                            : WordWriteResult.Ok(
                                "已删除" + (all ? "全文" : "当前选区") + " " + deletedCharacters + " 个字符。",
                                data,
                                true);
                        result.WithAfterRange(target.Start, target.Start);
                        return result;
                    }
                    finally
                    {
                        ComHelper.Release(range, selection);
                    }
                });
        }

        private OperationResult HandleApplyStyle(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string styleRequested = args.RequireString("styleName");
            string rangeMode = (args.GetOptionalString("range", "selection") ?? "selection")
                .Trim().ToLowerInvariant();
            int? paragraphIndex = args.GetOptionalInt("paragraphIndex");
            List<int> paragraphIndexes = args.GetOptionalIntList("paragraphIndexes");
            bool paragraphBatch = paragraphIndexes != null && paragraphIndexes.Count > 0;
            if (!paragraphBatch && (rangeMode == "paragraph" || paragraphIndex.HasValue))
            {
                if (!paragraphIndex.HasValue)
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.INVALID_PARAMS,
                        "range=paragraph 时必须提供 paragraphIndex（1 基）。");
                }
                paragraphBatch = true;
                paragraphIndexes = new List<int>();
                paragraphIndexes.Add(paragraphIndex.Value);
            }
            bool all = rangeMode == "all" || rangeMode == "document" || rangeMode == "全文";

            return RunWriteTool(
                application,
                WordToolNames.ApplyStyle,
                false,
                delegate(Word.Document document)
                {
                    if (paragraphBatch)
                    {
                        return WordWriteTargets.None("批量段落样式", "Word:Paragraphs");
                    }
                    if (all)
                    {
                        return WordWriteTargets.Document(
                            document, "全文", "Word:Document", false, true);
                    }
                    return WordWriteTargets.Selection(
                        application, "当前选区", "Word:Selection", false, true);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    string resolvedStyle = WordStyleCatalog.ResolveStyleName(document, styleRequested);
                    JArray failures = new JArray();
                    JArray appliedTargets = new JArray();
                    int appliedCount = 0;

                    if (paragraphBatch)
                    {
                        List<int> sorted = new List<int>(paragraphIndexes);
                        sorted.Sort();
                        sorted.Reverse();
                        int paragraphCount = WordComUtil.GetParagraphCount(document);
                        foreach (int index in sorted)
                        {
                            if (index < 1 || index > paragraphCount)
                            {
                                failures.Add(new JObject
                                {
                                    { "paragraphIndex", index },
                                    { "message", "第 " + index + " 段不存在（当前共 " + paragraphCount + " 段）" }
                                });
                                continue;
                            }
                            Word.Paragraph paragraph = null;
                            Word.Range range = null;
                            try
                            {
                                paragraph = WordComUtil.GetParagraph(document, index);
                                range = paragraph.Range;
                                object styleObject = resolvedStyle;
                                range.set_Style(ref styleObject);
                                string actual = WordComUtil.ReadStyleName(range);
                                if (string.Equals(actual, resolvedStyle, StringComparison.OrdinalIgnoreCase))
                                {
                                    appliedCount++;
                                    appliedTargets.Add(new JObject
                                    {
                                        { "paragraphIndex", index },
                                        { "style", actual }
                                    });
                                }
                                else
                                {
                                    failures.Add(new JObject
                                    {
                                        { "paragraphIndex", index },
                                        { "message", "应用后回读样式为「" + actual + "」，期望「" + resolvedStyle + "」" }
                                    });
                                }
                            }
                            catch (Exception ex)
                            {
                                failures.Add(new JObject
                                {
                                    { "paragraphIndex", index },
                                    { "message", OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) }
                                });
                            }
                            finally
                            {
                                ComHelper.Release(range, paragraph);
                            }
                        }
                    }
                    else
                    {
                        Word.Selection selection = null;
                        Word.Range range = null;
                        try
                        {
                            if (all)
                            {
                                range = document.Content;
                            }
                            else
                            {
                                selection = application.Selection;
                                range = selection == null ? null : selection.Range;
                            }
                            if (range == null)
                            {
                                failures.Add(new JObject { { "message", "当前没有有效选区。" } });
                            }
                            else
                            {
                                object styleObject = resolvedStyle;
                                range.set_Style(ref styleObject);
                                string actual = WordComUtil.ReadStyleName(range);
                                if (string.Equals(actual, resolvedStyle, StringComparison.OrdinalIgnoreCase))
                                {
                                    appliedCount = 1;
                                    appliedTargets.Add(new JObject
                                    {
                                        { "range", all ? "all" : "selection" },
                                        { "style", actual }
                                    });
                                }
                                else
                                {
                                    failures.Add(new JObject
                                    {
                                        { "message", "应用后回读样式为「" + actual + "」，期望「" + resolvedStyle + "」" }
                                    });
                                }
                            }
                        }
                        finally
                        {
                            ComHelper.Release(range, selection);
                        }
                    }

                    JObject data = new JObject();
                    data["styleRequested"] = styleRequested;
                    data["styleApplied"] = resolvedStyle;
                    data["appliedCount"] = appliedCount;
                    data["failedCount"] = failures.Count;
                    data["failures"] = failures;
                    data["targets"] = appliedTargets;

                    if (appliedCount == 0)
                    {
                        return WordWriteResult.Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "样式「" + styleRequested + "」已解析为「" + resolvedStyle +
                            "」，但没有段落成功应用。首个失败：" +
                            (failures.Count > 0 ? failures[0]["message"].ToString() : "未知原因") +
                            "。请用 word_get_paragraph_info 核对真实段号/样式。",
                            data,
                            false);
                    }
                    if (failures.Count > 0)
                    {
                        return WordWriteResult.Fail(
                            OfficeErrorCode.PARTIAL_APPLY,
                            "批量应用样式部分成功：成功 " + appliedCount + " 段，失败 " + failures.Count +
                            " 段；首个失败：" + failures[0]["message"].ToString() + "。失败项未中断后续段落。",
                            data,
                            true);
                    }

                    return WordWriteResult.Ok(
                        "已将" + (paragraphBatch ? " " + appliedCount + " 个段落" : (all ? "全文" : "当前选区")) +
                        "应用样式「" + resolvedStyle + "」（请求：" + styleRequested + "）。",
                        data,
                        true);
                });
        }

        private OperationResult HandleSetParagraphFormat(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int paragraphIndex = args.RequireInt("paragraphIndex", 1, int.MaxValue);
            Word.WdParagraphAlignment? alignment = null;
            if (args.Has("alignment"))
            {
                alignment = WordFormatUtil.ParseAlignment(args.GetOptionalString("alignment", null));
            }
            double? firstLineIndent = args.GetOptionalDouble("firstLineIndent");
            double? beforeSpacing = args.GetOptionalDouble("beforeSpacing");
            double? afterSpacing = args.GetOptionalDouble("afterSpacing");
            double? fontSize = args.GetOptionalDouble("fontSize");
            string fontName = args.Has("fontName") ? args.GetOptionalString("fontName", null) : null;
            bool? bold = ReadNullableBool(args, "bold");
            bool? italic = ReadNullableBool(args, "italic");
            int? color = args.Has("color")
                ? WordFormatUtil.ParseColor(args.GetOptionalString("color", null))
                : (int?)null;

            bool hasAny = alignment.HasValue || firstLineIndent.HasValue || beforeSpacing.HasValue ||
                          afterSpacing.HasValue || fontSize.HasValue || !string.IsNullOrWhiteSpace(fontName) ||
                          bold.HasValue || italic.HasValue || color.HasValue;
            if (!hasAny)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "word_set_paragraph_format 至少需要一个格式属性。收到 keys=" + args.DescribeKeys() +
                    "；支持 alignment/firstLineIndent/beforeSpacing/afterSpacing/fontSize/fontName/bold/italic/color");
            }

            return RunWriteTool(
                application,
                WordToolNames.SetParagraphFormat,
                false,
                delegate(Word.Document document)
                {
                    WordComUtil.ValidateParagraphIndex(document, paragraphIndex);
                    return WordWriteTargets.Paragraph(
                        document, paragraphIndex, "Word:Paragraph:" + paragraphIndex, true);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range range = null;
                    Word.Font font = null;
                    Word.ParagraphFormat format = null;
                    try
                    {
                        paragraph = WordComUtil.GetParagraph(document, paragraphIndex);
                        range = paragraph.Range;
                        format = paragraph.Format;
                        font = range.Font;
                        JObject before = CaptureParagraphFormatState(paragraph, font, format);
                        bool changed = false;

                        if (alignment.HasValue)
                        {
                            paragraph.Alignment = alignment.Value;
                        }
                        if (firstLineIndent.HasValue)
                        {
                            format.CharacterUnitFirstLineIndent = (float)firstLineIndent.Value;
                        }
                        if (beforeSpacing.HasValue)
                        {
                            format.SpaceBefore = (float)beforeSpacing.Value;
                        }
                        if (afterSpacing.HasValue)
                        {
                            format.SpaceAfter = (float)afterSpacing.Value;
                        }
                        if (fontSize.HasValue)
                        {
                            font.Size = (float)fontSize.Value;
                        }
                        if (!string.IsNullOrWhiteSpace(fontName))
                        {
                            font.Name = fontName;
                            try
                            {
                                font.NameFarEast = fontName;
                            }
                            catch (Exception ex)
                            {
                                Console.Error.WriteLine("[word] 段落字体 NameFarEast 设置失败：" + ex.Message);
                            }
                        }
                        if (bold.HasValue)
                        {
                            font.Bold = bold.Value ? -1 : 0;
                        }
                        if (italic.HasValue)
                        {
                            font.Italic = italic.Value ? -1 : 0;
                        }
                        if (color.HasValue)
                        {
                            font.Color = (Word.WdColor)color.Value;
                        }

                        JObject after = CaptureParagraphFormatState(paragraph, font, format);
                        JArray verification = new JArray();
                        string firstFailure = string.Empty;
                        if (alignment.HasValue)
                        {
                            AddStringVerification(verification, ref firstFailure, "alignment",
                                WordFormatUtil.GetAlignmentName(alignment.Value),
                                after.Value<string>("alignment"));
                        }
                        if (firstLineIndent.HasValue)
                        {
                            AddDoubleVerification(verification, ref firstFailure, "firstLineIndent",
                                firstLineIndent.Value,
                                after.Value<double?>("characterUnitFirstLineIndent"),
                                0.05);
                        }
                        if (beforeSpacing.HasValue)
                        {
                            AddDoubleVerification(verification, ref firstFailure, "beforeSpacing",
                                beforeSpacing.Value, after.Value<double?>("spaceBefore"), 0.2);
                        }
                        if (afterSpacing.HasValue)
                        {
                            AddDoubleVerification(verification, ref firstFailure, "afterSpacing",
                                afterSpacing.Value, after.Value<double?>("spaceAfter"), 0.2);
                        }
                        if (fontSize.HasValue)
                        {
                            AddDoubleVerification(verification, ref firstFailure, "fontSize",
                                fontSize.Value, after.Value<double?>("fontSize"), 0.2);
                        }
                        if (!string.IsNullOrWhiteSpace(fontName))
                        {
                            string actualName = after.Value<string>("fontName");
                            string actualFarEast = after.Value<string>("fontNameFarEast");
                            bool matched = string.Equals(actualName, fontName, StringComparison.OrdinalIgnoreCase) ||
                                           string.Equals(actualFarEast, fontName, StringComparison.OrdinalIgnoreCase);
                            AddVerification(verification, ref firstFailure, "fontName", fontName,
                                new JObject { { "fontName", actualName }, { "fontNameFarEast", actualFarEast } },
                                matched);
                        }
                        if (bold.HasValue)
                        {
                            AddBoolVerification(verification, ref firstFailure, "bold",
                                bold.Value, after["bold"]);
                        }
                        if (italic.HasValue)
                        {
                            AddBoolVerification(verification, ref firstFailure, "italic",
                                italic.Value, after["italic"]);
                        }
                        if (color.HasValue)
                        {
                            string expectedHex = WordFormatUtil.ColorToHex(color.Value);
                            string actualHex = after.Value<string>("colorHex");
                            bool matched = string.Equals(expectedHex, actualHex, StringComparison.OrdinalIgnoreCase);
                            AddVerification(verification, ref firstFailure, "color", expectedHex,
                                new JValue(actualHex), matched);
                        }

                        changed = !JToken.DeepEquals(before, after);
                        bool allMatched = string.IsNullOrWhiteSpace(firstFailure);
                        JObject data = new JObject();
                        data["paragraphIndex"] = paragraphIndex;
                        data["before"] = before;
                        data["after"] = after;
                        data["verification"] = verification;
                        data["allMatched"] = allMatched;
                        data["actual"] = after;

                        if (!allMatched)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "第 " + paragraphIndex + " 段格式回读不一致：" + firstFailure +
                                "。请根据 verification 调整参数。",
                                data,
                                changed);
                        }

                        return WordWriteResult.Ok(
                            "已设置第 " + paragraphIndex + " 段格式，" + verification.Count +
                            " 项回读一致。",
                            data,
                            changed);
                    }
                    finally
                    {
                        ComHelper.Release(format, font, range, paragraph);
                    }
                });
        }

        private OperationResult HandleClearFormat(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string rangeMode = (args.GetOptionalString("range", "selection") ?? "selection")
                .Trim().ToLowerInvariant();
            bool all = rangeMode == "all" || rangeMode == "document" || rangeMode == "全文";
            if (!all && rangeMode != "selection")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "range 只支持 selection/all，收到 " + rangeMode + "。");
            }

            return RunWriteTool(
                application,
                WordToolNames.ClearFormat,
                false,
                delegate(Word.Document document)
                {
                    if (all)
                    {
                        return WordWriteTargets.Document(document, "全文", "Word:Document", false, true);
                    }
                    return WordWriteTargets.Selection(
                        application, "当前选区", "Word:Selection", false, true);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Selection selection = null;
                    Word.Range range = null;
                    Word.Font font = null;
                    Word.ParagraphFormat format = null;
                    Word.Paragraphs paragraphs = null;
                    try
                    {
                        if (all)
                        {
                            range = document.Content;
                        }
                        else
                        {
                            selection = application.Selection;
                            range = selection == null ? null : selection.Range;
                        }
                        if (range == null)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "当前没有有效选区，range=selection 无法清除格式；可改用 range=all。",
                                null,
                                false);
                        }

                        int characters = range.Text == null ? 0 : range.Text.Length;
                        paragraphs = range.Paragraphs;
                        int paragraphCount = paragraphs == null ? 0 : paragraphs.Count;
                        JObject before = WordFormatUtil.CaptureRangeFormat(range);
                        font = range.Font;
                        format = range.ParagraphFormat;
                        font.Reset();
                        format.Reset();
                        JObject after = WordFormatUtil.CaptureRangeFormat(range);
                        bool changed = !JToken.DeepEquals(before, after);

                        JObject data = new JObject();
                        data["range"] = all ? "all" : "selection";
                        data["clearedCharacters"] = characters;
                        data["clearedParagraphs"] = paragraphCount;
                        data["before"] = before;
                        data["after"] = after;
                        return WordWriteResult.Ok(
                            "已清除" + (all ? "全文" : "当前选区") + "的直接字符/段落格式（" +
                            characters + " 字符，" + paragraphCount + " 段）。",
                            data,
                            changed);
                    }
                    finally
                    {
                        ComHelper.Release(paragraphs, format, font, range, selection);
                    }
                });
        }

        private OperationResult HandleApplyAsRevision(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int paragraphIndex = args.RequireInt("paragraphIndex", 1, int.MaxValue);
            string original = args.RequireString("original");
            if (!args.Has("corrected"))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "缺少 corrected 参数（替换后的文本，可为空字符串）。");
            }
            string corrected = args.GetOptionalString("corrected", string.Empty);
            bool matchCase = args.GetOptionalBool("matchCase", true);

            return RunWriteTool(
                application,
                WordToolNames.ApplyAsRevision,
                false,
                delegate(Word.Document document)
                {
                    WordComUtil.ValidateParagraphIndex(document, paragraphIndex);
                    return WordWriteTargets.Paragraph(
                        document, paragraphIndex, "Word:Paragraph:" + paragraphIndex, true);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range paragraphRange = null;
                    Word.Range replaceRange = null;
                    Word.Revisions revisions = null;
                    bool originalTracking = false;
                    bool trackingCaptured = false;
                    try
                    {
                        paragraph = WordComUtil.GetParagraph(document, paragraphIndex);
                        paragraphRange = paragraph.Range;
                        string paragraphText = paragraphRange.Text ?? string.Empty;
                        StringComparison comparison = matchCase
                            ? StringComparison.Ordinal
                            : StringComparison.OrdinalIgnoreCase;
                        int index = paragraphText.IndexOf(original, comparison);
                        if (index < 0)
                        {
                            JObject failureData = new JObject();
                            failureData["paragraphIndex"] = paragraphIndex;
                            failureData["original"] = original;
                            failureData["paragraphPreview"] = WordComUtil.NormalizePreview(paragraphText, 160);
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "第 " + paragraphIndex + " 段中未找到原文「" +
                                WordComUtil.Truncate(original, 60) + "」。当前段文本：" +
                                WordComUtil.NormalizePreview(paragraphText, 120) +
                                "。请先用 word_get_paragraph_info 核对原文与 matchCase。",
                                failureData,
                                false);
                        }

                        revisions = document.Revisions;
                        int beforeCount = revisions == null ? 0 : revisions.Count;
                        trackingCaptured = true;
                        originalTracking = document.TrackRevisions;
                        document.TrackRevisions = true;

                        replaceRange = WordComUtil.CreateRange(
                            document,
                            paragraphRange.Start + index,
                            paragraphRange.Start + index + original.Length);
                        replaceRange.Text = corrected;

                        revisions = document.Revisions;
                        int afterCount = revisions == null ? 0 : revisions.Count;
                        JObject data = new JObject();
                        data["paragraphIndex"] = paragraphIndex;
                        data["original"] = original;
                        data["corrected"] = corrected;
                        data["revisionCountBefore"] = beforeCount;
                        data["revisionCountAfter"] = afterCount;
                        data["revisionDelta"] = afterCount - beforeCount;

                        if (afterCount <= beforeCount)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INTERNAL_ERROR,
                                "文本替换已提交，但修订集合没有增加（前 " + beforeCount + "，后 " +
                                afterCount + "）。请确认 Word 修订模式未被策略/文档保护禁用；" +
                                "本次改动可能已按直接编辑落入文档，请先核对。",
                                data,
                                true);
                        }

                        return WordWriteResult.Ok(
                            "已以修订模式替换第 " + paragraphIndex + " 段中的原文，产生 " +
                            (afterCount - beforeCount) + " 个修订。",
                            data,
                            true);
                    }
                    finally
                    {
                        if (trackingCaptured)
                        {
                            try
                            {
                                document.TrackRevisions = originalTracking;
                            }
                            catch (Exception ex)
                            {
                                Console.Error.WriteLine("[word] 恢复 TrackRevisions 失败：" + ex.Message);
                            }
                        }
                        ComHelper.Release(revisions, replaceRange, paragraphRange, paragraph);
                    }
                });
        }

        private static bool? ReadNullableBool(WordArguments args, string name)
        {
            return args.Has(name) ? new bool?(args.GetOptionalBool(name, false)) : (bool?)null;
        }

        private static JObject CaptureParagraphFormatState(
            Word.Paragraph paragraph,
            Word.Font font,
            Word.ParagraphFormat format)
        {
            JObject state = new JObject();
            state["alignment"] = WordFormatUtil.GetAlignmentName(paragraph.Alignment);
            state["firstLineIndent"] = (double)format.FirstLineIndent;
            state["characterUnitFirstLineIndent"] = (double)format.CharacterUnitFirstLineIndent;
            state["spaceBefore"] = (double)format.SpaceBefore;
            state["spaceAfter"] = (double)format.SpaceAfter;
            if (font != null)
            {
                JObject fontState = WordFormatUtil.CaptureFont(font);
                state["fontSize"] = fontState["fontSize"];
                state["fontName"] = fontState["fontName"];
                state["fontNameFarEast"] = fontState["fontNameFarEast"];
                state["bold"] = fontState["bold"];
                state["italic"] = fontState["italic"];
                state["color"] = fontState["color"];
                state["colorHex"] = fontState["colorHex"];
            }
            return state;
        }

        private static void AddStringVerification(
            JArray verification,
            ref string firstFailure,
            string property,
            string expected,
            string actual)
        {
            bool matched = string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase);
            AddVerification(verification, ref firstFailure, property, expected, new JValue(actual), matched);
        }

        private static void AddDoubleVerification(
            JArray verification,
            ref string firstFailure,
            string property,
            double expected,
            double? actual,
            double tolerance)
        {
            bool matched = actual.HasValue && WordFormatUtil.CloseEnough(actual.Value, expected, tolerance);
            AddVerification(verification, ref firstFailure, property, expected,
                actual.HasValue ? new JValue(actual.Value) : JValue.CreateNull(), matched);
        }

        private static void AddBoolVerification(
            JArray verification,
            ref string firstFailure,
            string property,
            bool expected,
            JToken actual)
        {
            bool matched = actual != null && actual.Type == JTokenType.Boolean &&
                           actual.Value<bool>() == expected;
            AddVerification(verification, ref firstFailure, property, expected, actual, matched);
        }

        private static void AddVerification(
            JArray verification,
            ref string firstFailure,
            string property,
            JToken expected,
            JToken actual,
            bool matched)
        {
            JObject item = new JObject();
            item["property"] = property;
            item["expected"] = expected == null ? JValue.CreateNull() : expected.DeepClone();
            item["actual"] = actual == null ? JValue.CreateNull() : actual.DeepClone();
            item["matched"] = matched;
            verification.Add(item);
            if (!matched && string.IsNullOrWhiteSpace(firstFailure))
            {
                firstFailure = property + " 期望 " + (expected == null ? "null" : expected.ToString()) +
                               "，回读 " + (actual == null ? "null" : actual.ToString());
            }
        }
    }
}
