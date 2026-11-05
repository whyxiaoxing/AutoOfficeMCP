using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Office.Contracts;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    /// <summary>
    /// Word TODO 写入能力：批注完成/删除、样式定义修改，以及带格式的查找替换。
    /// 成功必须由 Word 回读支撑；中途失败且文档已变化时返回 PARTIAL_APPLY。
    /// </summary>
    public sealed partial class WordToolExecutor
    {
        private sealed class WordReplacementFormat
        {
            public bool? Bold { get; set; }

            public int? Color { get; set; }

            public double? FontSize { get; set; }

            public string FontName { get; set; }

            public Word.WdUnderline? UnderlineStyle { get; set; }

            public bool HasAny
            {
                get
                {
                    return Bold.HasValue || Color.HasValue || FontSize.HasValue ||
                           !string.IsNullOrWhiteSpace(FontName) || UnderlineStyle.HasValue;
                }
            }
        }

        private void RegisterTodoWriteHandlers(
            Dictionary<string, Func<Word.Application, JObject, OperationResult>> handlers)
        {
            handlers[WordToolNames.CommentResolve] = HandleCommentResolve;
            handlers[WordToolNames.CommentDelete] = HandleCommentDelete;
            handlers[WordToolNames.StyleModify] = HandleStyleModify;
        }

        private OperationResult HandleCommentResolve(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int commentIndex = args.RequireInt("commentIndex", 1, int.MaxValue);

            return RunWriteTool(
                application,
                WordToolNames.CommentResolve,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.None(
                        "批注 #" + commentIndex, "Word:Comment:" + commentIndex);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Comments comments = null;
                    Word.Comment comment = null;
                    try
                    {
                        comments = document.Comments;
                        int total = comments == null ? 0 : comments.Count;
                        comment = FindComment(comments, commentIndex);
                        if (comment == null)
                        {
                            JObject missingData = new JObject();
                            missingData["commentIndex"] = commentIndex;
                            missingData["totalComments"] = total;
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "批注 #" + commentIndex + " 不存在（当前共 " + total + " 个批注）。",
                                missingData,
                                false);
                        }

                        bool before = comment.Done;
                        if (!before)
                        {
                            comment.Done = true;
                        }
                        bool after = comment.Done;
                        string preview;
                        Word.Range commentRange = null;
                        try
                        {
                            commentRange = comment.Range;
                            preview = WordComUtil.NormalizePreview(
                                commentRange == null ? string.Empty : commentRange.Text, 120);
                        }
                        finally
                        {
                            ComHelper.Release(commentRange);
                        }
                        JObject data = new JObject();
                        data["commentIndex"] = commentIndex;
                        data["doneBefore"] = before;
                        data["doneAfter"] = after;
                        data["done"] = after;
                        data["author"] = comment.Author ?? string.Empty;
                        data["textPreview"] = preview;
                        if (!after)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.PARTIAL_APPLY,
                                "已请求将批注 #" + commentIndex + " 标记为完成，但回读 Done=false；" +
                                "请检查批注状态后重试。",
                                data,
                                before != after);
                        }

                        WordWriteResult result = WordWriteResult.Ok(
                            "批注 #" + commentIndex + " 已完成" +
                            (before ? "（操作前已是完成状态）。" : "，回读 Done=true。"),
                            data,
                            before != after);
                        result.AddTargetRef("Word:Comment:" + commentIndex);
                        return result;
                    }
                    finally
                    {
                        ComHelper.Release(comment, comments);
                    }
                });
        }

        private OperationResult HandleCommentDelete(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            JArray rawIndexes = args.RequireArray("commentIndexes");
            bool confirm = args.GetOptionalBool("confirm", false);
            if (!confirm)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "word_comment_delete 是高危操作，必须显式传入 confirm=true；本次未删除任何批注。");
            }

            List<int> requested = new List<int>();
            HashSet<int> seen = new HashSet<int>();
            for (int index = 0; index < rawIndexes.Count; index++)
            {
                JToken token = rawIndexes[index];
                int value;
                try
                {
                    value = token.Value<int>();
                }
                catch (Exception)
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.INVALID_PARAMS,
                        "commentIndexes[" + index + "] 必须是整数，实际为 " + token);
                }
                if (value < 1)
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.INVALID_PARAMS,
                        "commentIndexes[" + index + "] 必须从 1 开始，收到 " + value + "。");
                }
                if (seen.Add(value))
                {
                    requested.Add(value);
                }
            }
            if (requested.Count == 0)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "commentIndexes 不能为空数组。");
            }

            requested.Sort();
            requested.Reverse();

            return RunWriteTool(
                application,
                WordToolNames.CommentDelete,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.None("删除批注", "Word:Comments");
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Comments comments = null;
                    JArray deletedIndexes = new JArray();
                    JArray missingIndexes = new JArray();
                    int deletedCount = 0;
                    string failure = string.Empty;
                    try
                    {
                        comments = document.Comments;
                        int beforeCount = comments == null ? 0 : comments.Count;

                        // 先完整校验，避免部分删除后才发现后续索引不存在。
                        foreach (int commentIndex in requested)
                        {
                            Word.Comment existing = null;
                            try
                            {
                                existing = FindComment(comments, commentIndex);
                                if (existing == null)
                                {
                                    missingIndexes.Add(commentIndex);
                                }
                            }
                            finally
                            {
                                ComHelper.Release(existing);
                            }
                        }
                        if (missingIndexes.Count > 0)
                        {
                            JObject invalidData = new JObject();
                            invalidData["requestedIndexes"] = new JArray(requested);
                            invalidData["missingIndexes"] = missingIndexes;
                            invalidData["totalComments"] = beforeCount;
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "以下批注索引不存在：" + missingIndexes.ToString() +
                                "（当前共 " + beforeCount + " 个批注）。未删除任何批注。",
                                invalidData,
                                false);
                        }

                        foreach (int commentIndex in requested)
                        {
                            Word.Comment comment = null;
                            try
                            {
                                comment = FindComment(comments, commentIndex);
                                if (comment == null)
                                {
                                    missingIndexes.Add(commentIndex);
                                    failure = "删除过程中批注 #" + commentIndex + " 已不存在。";
                                    break;
                                }
                                comment.Delete();
                                deletedCount++;
                                deletedIndexes.Add(commentIndex);
                            }
                            catch (Exception ex)
                            {
                                failure = "删除批注 #" + commentIndex + " 失败：" +
                                          OfficeErrorMapper.DescribeException(ex, OfficeApp.Word);
                                break;
                            }
                            finally
                            {
                                ComHelper.Release(comment);
                            }
                        }

                        int afterCount = comments == null ? 0 : comments.Count;
                        JObject data = new JObject();
                        data["requestedIndexes"] = new JArray(requested);
                        data["deletedIndexes"] = deletedIndexes;
                        data["deletedCount"] = deletedCount;
                        data["requestedCount"] = requested.Count;
                        data["commentsBefore"] = beforeCount;
                        data["commentsAfter"] = afterCount;
                        data["remainingMissingIndexes"] = missingIndexes;
                        if (!string.IsNullOrWhiteSpace(failure))
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.PARTIAL_APPLY,
                                failure + " 已删除 " + deletedCount + " 个批注；不会自动重放。",
                                data,
                                deletedCount > 0);
                        }
                        if (afterCount != beforeCount - deletedCount)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.PARTIAL_APPLY,
                                "批注删除数量回读不一致：删除前 " + beforeCount + "，请求删除 " +
                                deletedCount + "，删除后 " + afterCount + "。",
                                data,
                                deletedCount > 0);
                        }

                        WordWriteResult result = WordWriteResult.Ok(
                            "已按索引倒序删除 " + deletedCount + " 个批注（" + beforeCount +
                            "→" + afterCount + "）。",
                            data,
                            deletedCount > 0);
                        foreach (JToken item in deletedIndexes)
                        {
                            result.AddTargetRef("Word:Comment:" + item.Value<int>());
                        }
                        return result;
                    }
                    finally
                    {
                        ComHelper.Release(comments);
                    }
                });
        }

        private static Word.Comment FindComment(Word.Comments comments, int commentIndex)
        {
            if (comments == null)
            {
                return null;
            }
            int count = comments.Count;
            for (int position = 1; position <= count; position++)
            {
                Word.Comment candidate = null;
                try
                {
                    candidate = comments[position];
                    if (candidate.Index == commentIndex)
                    {
                        Word.Comment found = candidate;
                        candidate = null;
                        return found;
                    }
                }
                finally
                {
                    ComHelper.Release(candidate);
                }
            }
            return null;
        }

        private OperationResult HandleStyleModify(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string styleName = args.RequireString("name");
            string fontName = args.Has("font") ? args.GetOptionalString("font", null) : null;
            double? fontSize = args.GetOptionalDouble("size");
            int? color = args.Has("color")
                ? new int?(ParseStyleColor(args.GetOptionalString("color", null)))
                : (int?)null;
            bool? bold = ReadNullableBool(args, "bold");
            bool? italic = ReadNullableBool(args, "italic");
            double? lineSpacing = args.GetOptionalDouble("lineSpacing");
            string basedOn = args.Has("basedOn") ? args.GetOptionalString("basedOn", null) : null;

            if (!string.IsNullOrEmpty(fontName) && string.IsNullOrWhiteSpace(fontName))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "font 不能为空字符串。");
            }
            fontName = string.IsNullOrWhiteSpace(fontName) ? null : fontName.Trim();
            if (fontSize.HasValue && (fontSize.Value < 1d || fontSize.Value > 500d))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "size 允许 1-500，收到 " + fontSize.Value.ToString(CultureInfo.InvariantCulture) + "。");
            }
            if (lineSpacing.HasValue && (lineSpacing.Value <= 0d || lineSpacing.Value > 100d))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "lineSpacing 是倍数，允许 0-100（不含 0），收到 " +
                    lineSpacing.Value.ToString(CultureInfo.InvariantCulture) + "。");
            }
            if (basedOn != null)
            {
                basedOn = basedOn.Trim();
                if (basedOn.Length == 0)
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.INVALID_PARAMS,
                        "basedOn 必须是非空样式名；不传该参数表示不修改基准样式。");
                }
                if (string.Equals(basedOn, styleName, StringComparison.OrdinalIgnoreCase))
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.INVALID_PARAMS,
                        "basedOn 不能是样式自身（" + styleName + "），否则会形成循环基准。");
                }
            }

            bool hasAny = !string.IsNullOrWhiteSpace(fontName) || fontSize.HasValue ||
                          color.HasValue || bold.HasValue || italic.HasValue ||
                          lineSpacing.HasValue || basedOn != null;
            if (!hasAny)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "word_style_modify 至少需要一个样式定义属性。收到 keys=" + args.DescribeKeys() +
                    "；支持 font/size/color/bold/italic/lineSpacing/basedOn。");
            }

            return RunWriteTool(
                application,
                WordToolNames.StyleModify,
                false,
                delegate(Word.Document document)
                {
                    if (!WordComUtil.StyleExists(document, styleName))
                    {
                        throw WordArguments.Invalid(
                            "样式「" + styleName + "」不存在。请先用 word_styles_list 查看可用样式名。");
                    }
                    if (basedOn != null && !WordComUtil.StyleExists(document, basedOn))
                    {
                        throw WordArguments.Invalid(
                            "basedOn 指向的样式「" + basedOn + "」不存在。请先用 word_styles_list 确认。");
                    }
                    return WordWriteTargets.None(
                        "样式「" + styleName + "」", "Word:Style:" + styleName);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Styles styles = null;
                    Word.Style style = null;
                    Word.Style baseStyle = null;
                    Word.Font font = null;
                    Word.ParagraphFormat format = null;
                    try
                    {
                        styles = document.Styles;
                        style = styles[styleName] as Word.Style;
                        if (style == null)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "无法取得样式「" + styleName + "」的定义对象。",
                                null,
                                false);
                        }
                        if (lineSpacing.HasValue &&
                            style.Type != Word.WdStyleType.wdStyleTypeParagraph)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "lineSpacing 只能用于段落样式；样式「" + styleName +
                                "」类型为 " + WordStyleTypeName(style.Type) + "。",
                                null,
                                false);
                        }
                        if (basedOn != null)
                        {
                            baseStyle = styles[basedOn] as Word.Style;
                            if (baseStyle == null)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INVALID_PARAMS,
                                    "无法取得基准样式「" + basedOn + "」的定义对象。",
                                    null,
                                    false);
                            }
                        }

                        JObject before = CaptureStyleDefinition(style);
                        font = style.Font;
                        if (!string.IsNullOrWhiteSpace(fontName))
                        {
                            font.Name = fontName;
                            try
                            {
                                font.NameFarEast = fontName;
                            }
                            catch (Exception ex)
                            {
                                Console.Error.WriteLine("[word] 样式字体 NameFarEast 设置失败：" + ex.Message);
                            }
                        }
                        if (fontSize.HasValue)
                        {
                            font.Size = (float)fontSize.Value;
                        }
                        if (color.HasValue)
                        {
                            font.Color = (Word.WdColor)color.Value;
                        }
                        if (bold.HasValue)
                        {
                            font.Bold = bold.Value ? -1 : 0;
                        }
                        if (italic.HasValue)
                        {
                            font.Italic = italic.Value ? -1 : 0;
                        }
                        if (lineSpacing.HasValue)
                        {
                            format = style.ParagraphFormat;
                            format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
                            format.LineSpacing = (float)(12d * lineSpacing.Value);
                        }
                        if (baseStyle != null)
                        {
                            object baseStyleObject = baseStyle;
                            style.set_BaseStyle(ref baseStyleObject);
                        }

                        JObject after = CaptureStyleDefinition(style);
                        JArray verification = new JArray();
                        string firstFailure = string.Empty;
                        if (!string.IsNullOrWhiteSpace(fontName))
                        {
                            string actualName = after.Value<string>("fontName") ?? string.Empty;
                            string actualFarEast = after.Value<string>("fontNameFarEast") ?? string.Empty;
                            bool matched = string.Equals(
                                               actualName, fontName, StringComparison.OrdinalIgnoreCase) ||
                                           string.Equals(
                                               actualFarEast, fontName, StringComparison.OrdinalIgnoreCase);
                            AddVerification(
                                verification,
                                ref firstFailure,
                                "font",
                                new JValue(fontName),
                                new JObject
                                {
                                    { "fontName", actualName },
                                    { "fontNameFarEast", actualFarEast }
                                },
                                matched);
                        }
                        if (fontSize.HasValue)
                        {
                            AddDoubleVerification(
                                verification, ref firstFailure, "size", fontSize.Value,
                                after.Value<double?>("fontSize"), 0.2);
                        }
                        if (color.HasValue)
                        {
                            string expectedHex = StyleColorToHex(color.Value);
                            string actualHex = after.Value<string>("colorHex") ?? string.Empty;
                            AddVerification(
                                verification,
                                ref firstFailure,
                                "color",
                                new JValue(expectedHex),
                                new JValue(actualHex),
                                string.Equals(expectedHex, actualHex, StringComparison.OrdinalIgnoreCase));
                        }
                        if (bold.HasValue)
                        {
                            AddBoolVerification(verification, ref firstFailure, "bold", bold.Value, after["bold"]);
                        }
                        if (italic.HasValue)
                        {
                            AddBoolVerification(verification, ref firstFailure, "italic", italic.Value, after["italic"]);
                        }
                        if (lineSpacing.HasValue)
                        {
                            AddDoubleVerification(
                                verification, ref firstFailure, "lineSpacing", lineSpacing.Value,
                                after.Value<double?>("lineSpacingMultiplier"), 0.01);
                        }
                        if (basedOn != null)
                        {
                            string actualBasedOn = after.Value<string>("basedOn") ?? string.Empty;
                            AddStringVerification(
                                verification, ref firstFailure, "basedOn", basedOn, actualBasedOn);
                        }

                        bool changed = !JToken.DeepEquals(before, after);
                        bool allMatched = string.IsNullOrWhiteSpace(firstFailure);
                        JObject data = new JObject();
                        data["styleName"] = styleName;
                        data["styleType"] = WordStyleTypeName(style.Type);
                        data["before"] = before;
                        data["after"] = after;
                        data["verification"] = verification;
                        data["allMatched"] = allMatched;
                        if (!allMatched)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.PARTIAL_APPLY,
                                "样式「" + styleName + "」定义修改后回读不一致：" + firstFailure +
                                "。已发生的定义变化不会自动回滚，请用 word_styles_list 复核。",
                                data,
                                changed);
                        }

                        WordWriteResult result = WordWriteResult.Ok(
                            "已修改样式「" + styleName + "」定义，" + verification.Count +
                            " 项回读一致。",
                            data,
                            changed);
                        result.AddTargetRef("Word:Style:" + styleName);
                        return result;
                    }
                    finally
                    {
                        ComHelper.Release(format, font, baseStyle, style, styles);
                    }
                });
        }

        private static JObject CaptureStyleDefinition(Word.Style style)
        {
            JObject state = new JObject();
            if (style == null)
            {
                return state;
            }
            Word.Font font = null;
            Word.ParagraphFormat format = null;
            try
            {
                state["name"] = style.NameLocal ?? string.Empty;
                state["type"] = WordStyleTypeName(style.Type);
                state["typeCode"] = (int)style.Type;
                state["basedOn"] = ReadBaseStyleName(style);
                font = style.Font;
                state["fontName"] = font.Name ?? string.Empty;
                state["fontNameFarEast"] = font.NameFarEast ?? string.Empty;
                state["fontSize"] = WordFormatUtil.IsUndefined(font.Size)
                    ? JValue.CreateNull()
                    : new JValue((double)font.Size);
                state["color"] = (int)font.Color;
                state["colorHex"] = StyleColorToHex((int)font.Color);
                state["bold"] = WordFormatUtil.TriStateToken(font.Bold);
                state["italic"] = WordFormatUtil.TriStateToken(font.Italic);
                try
                {
                    format = style.ParagraphFormat;
                    state["lineSpacingRule"] = format.LineSpacingRule.ToString();
                    state["lineSpacingPt"] = (double)format.LineSpacing;
                    // Word 会把 1.5 倍行距规范化为 wdLineSpace1pt5（而不是 Multiple），
                    // 因此按“磅/12”换算的倍数在 Single/Multiple/1pt5/Double 规则下都成立。
                    bool multipleRule = format.LineSpacingRule == Word.WdLineSpacing.wdLineSpaceSingle ||
                                        format.LineSpacingRule == Word.WdLineSpacing.wdLineSpaceMultiple ||
                                        format.LineSpacingRule == Word.WdLineSpacing.wdLineSpace1pt5 ||
                                        format.LineSpacingRule == Word.WdLineSpacing.wdLineSpaceDouble;
                    state["lineSpacingMultiplier"] = multipleRule
                        ? new JValue(Math.Round((double)format.LineSpacing / 12d, 4))
                        : JValue.CreateNull();
                }
                catch (Exception ex)
                {
                    state["lineSpacingRule"] = string.Empty;
                    state["lineSpacingPt"] = JValue.CreateNull();
                    state["lineSpacingMultiplier"] = JValue.CreateNull();
                    state["lineSpacingReadError"] = ex.Message;
                }
                return state;
            }
            finally
            {
                ComHelper.Release(format, font);
            }
        }

        private static int ParseStyleColor(string raw)
        {
            string text = (raw ?? string.Empty).Trim();
            string hex = text.StartsWith("#", StringComparison.Ordinal)
                ? text.Substring(1)
                : text;
            if (hex.Length == 3)
            {
                hex = new string(new char[]
                {
                    hex[0], hex[0], hex[1], hex[1], hex[2], hex[2]
                });
            }
            if (hex.Length == 6)
            {
                int rgb;
                if (int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out rgb))
                {
                    // Word 的 WdColor 是 BGR 顺序，而用户输入习惯为 #RRGGBB。
                    return ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF);
                }
            }
            return WordFormatUtil.ParseColor(text);
        }

        private static string StyleColorToHex(int color)
        {
            if (color == (int)Word.WdColor.wdColorAutomatic)
            {
                return "auto";
            }
            int bgr = color & 0xFFFFFF;
            int red = bgr & 0xFF;
            int green = (bgr >> 8) & 0xFF;
            int blue = (bgr >> 16) & 0xFF;
            int rgb = (red << 16) | (green << 8) | blue;
            return "#" + rgb.ToString("X6", CultureInfo.InvariantCulture);
        }

        private OperationResult HandleReplaceTextWithFormat(
            Word.Application application,
            string findText,
            string replaceText,
            bool matchCase,
            bool matchWholeWord,
            JObject formatObject)
        {
            WordReplacementFormat format = ParseReplacementFormat(formatObject);
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
                    try
                    {
                        content = document.Content;
                        string beforeText = content.Text ?? string.Empty;
                        int expectedCount = WordComUtil.CountOccurrences(
                            beforeText, findText, matchCase, matchWholeWord);
                        int position = content.Start;
                        int replacedCount = 0;
                        int formatVerifiedCount = 0;
                        int guard = 0;
                        int maxReplacements = Math.Max(1000, expectedCount + 10);
                        string formatFailure = string.Empty;
                        JArray failures = new JArray();

                        while (guard < maxReplacements)
                        {
                            guard++;
                            int currentEnd = WordComUtil.GetContentEnd(document);
                            if (position >= currentEnd)
                            {
                                break;
                            }

                            Word.Range searchRange = null;
                            Word.Find find = null;
                            try
                            {
                                searchRange = WordComUtil.CreateRange(
                                    document, position, currentEnd);
                                find = searchRange.Find;
                                find.ClearFormatting();
                                find.Replacement.ClearFormatting();
                                find.Text = findText;
                                find.Replacement.Text = replaceText ?? string.Empty;
                                find.Forward = true;
                                find.Wrap = Word.WdFindWrap.wdFindStop;
                                find.MatchCase = matchCase;
                                find.MatchWholeWord = matchWholeWord;
                                find.MatchWildcards = false;
                                find.Format = false;
                                object missing = Type.Missing;
                                object replaceOne = Word.WdReplace.wdReplaceOne;
                                bool found = find.Execute(
                                    ref missing, ref missing, ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing, ref missing, ref missing,
                                    ref replaceOne, ref missing, ref missing, ref missing, ref missing);
                                if (!found)
                                {
                                    break;
                                }

                                int matchStart = searchRange.Start;
                                int matchEnd = searchRange.End;
                                JArray verification;
                                bool matched;
                                ApplyReplacementFormat(searchRange, format, out verification, out matched);
                                replacedCount++;
                                if (matched)
                                {
                                    formatVerifiedCount++;
                                }
                                else if (string.IsNullOrWhiteSpace(formatFailure))
                                {
                                    formatFailure = "第 " + replacedCount +
                                                    " 处替换的格式回读不一致。";
                                    failures.Add(new JObject
                                    {
                                        { "replacementIndex", replacedCount },
                                        { "rangeStart", matchStart },
                                        { "rangeEnd", matchEnd },
                                        { "verification", verification }
                                    });
                                }
                                if (!matched)
                                {
                                    break;
                                }
                                position = matchEnd;
                            }
                            finally
                            {
                                ComHelper.Release(find, searchRange);
                            }
                        }

                        Word.Range afterContent = null;
                        string afterText;
                        try
                        {
                            afterContent = document.Content;
                            afterText = afterContent == null ? string.Empty : (afterContent.Text ?? string.Empty);
                        }
                        finally
                        {
                            ComHelper.Release(afterContent);
                        }

                        bool changed = replacedCount > 0 ||
                                       !string.Equals(beforeText, afterText, StringComparison.Ordinal);
                        JObject data = new JObject();
                        data["find"] = findText;
                        data["replace"] = replaceText ?? string.Empty;
                        data["matchCase"] = matchCase;
                        data["matchWholeWord"] = matchWholeWord;
                        data["format"] = formatObject.DeepClone();
                        data["replacementCount"] = replacedCount;
                        data["preScanCount"] = expectedCount;
                        data["formatVerifiedCount"] = formatVerifiedCount;
                        data["formatVerified"] = replacedCount > 0 &&
                                                  formatVerifiedCount == replacedCount;
                        data["formatFailures"] = failures;
                        data["documentTextLengthBefore"] = beforeText.Length;
                        data["documentTextLengthAfter"] = afterText.Length;
                        data["partialApply"] = !string.IsNullOrWhiteSpace(formatFailure);

                        if (!string.IsNullOrWhiteSpace(formatFailure))
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.PARTIAL_APPLY,
                                "带格式替换已执行 " + replacedCount + " 处，但后续格式核验失败：" +
                                formatFailure + "。不会自动重放，请用 word_read_selection 复核。",
                                data,
                                changed);
                        }
                        if (replacedCount == 0)
                        {
                            return WordWriteResult.Ok(
                                "未找到匹配文本「" + WordComUtil.Truncate(findText, 40) +
                                "」，文档未修改。",
                                data,
                                false);
                        }
                        if (replacedCount >= maxReplacements)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.PARTIAL_APPLY,
                                "带格式替换达到安全上限 " + maxReplacements +
                                " 处后停止；请检查替换文本是否再次包含查找文本。",
                                data,
                                changed);
                        }

                        WordWriteResult result = WordWriteResult.Ok(
                            "已带格式替换 " + replacedCount + " 处「" +
                            WordComUtil.Truncate(findText, 30) + "」→「" +
                            WordComUtil.Truncate(replaceText, 30) + "」，" +
                            formatVerifiedCount + " 处格式回读一致。",
                            data,
                            true);
                        result.AddTargetRef("Word:Document");
                        return result;
                    }
                    finally
                    {
                        ComHelper.Release(content);
                    }
                },
                true);
        }

        private static WordReplacementFormat ParseReplacementFormat(JObject formatObject)
        {
            WordReplacementFormat request = new WordReplacementFormat();
            foreach (JProperty property in formatObject.Properties())
            {
                JToken token = property.Value;
                if (token == null || token.Type == JTokenType.Null)
                {
                    continue;
                }
                switch (property.Name)
                {
                    case "bold":
                        if (token.Type != JTokenType.Boolean)
                        {
                            throw WordArguments.Invalid("format.bold 必须是布尔值。");
                        }
                        request.Bold = new bool?(token.Value<bool>());
                        break;
                    case "color":
                        if (token.Type != JTokenType.String)
                        {
                            throw WordArguments.Invalid("format.color 必须是颜色字符串，如 #C00000。");
                        }
                        request.Color = new int?(ParseStyleColor(token.Value<string>()));
                        break;
                    case "fontSize":
                        if (token.Type != JTokenType.Integer && token.Type != JTokenType.Float)
                        {
                            throw WordArguments.Invalid("format.fontSize 必须是数字。");
                        }
                        double size = token.Value<double>();
                        if (size < 1d || size > 500d)
                        {
                            throw WordArguments.Invalid("format.fontSize 允许 1-500，收到 " + size + "。");
                        }
                        request.FontSize = new double?(size);
                        break;
                    case "fontName":
                        if (token.Type != JTokenType.String ||
                            string.IsNullOrWhiteSpace(token.Value<string>()))
                        {
                            throw WordArguments.Invalid("format.fontName 必须是非空字符串。");
                        }
                        request.FontName = token.Value<string>().Trim();
                        break;
                    case "underlineStyle":
                        if (token.Type != JTokenType.String)
                        {
                            throw WordArguments.Invalid("format.underlineStyle 必须是字符串。");
                        }
                        request.UnderlineStyle = new Word.WdUnderline?(
                            WordFormatUtil.ParseUnderlineStyle(token.Value<string>()));
                        break;
                    default:
                        throw WordArguments.Invalid(
                            "format 不支持属性 " + property.Name +
                            "；支持 bold/color/fontSize/fontName/underlineStyle。");
                }
            }
            if (!request.HasAny)
            {
                throw WordArguments.Invalid(
                    "format 至少需要 bold/color/fontSize/fontName/underlineStyle 之一。");
            }
            return request;
        }

        private static void ApplyReplacementFormat(
            Word.Range range,
            WordReplacementFormat request,
            out JArray verification,
            out bool matched)
        {
            verification = new JArray();
            Word.Font font = null;
            try
            {
                font = range.Font;
                if (request.Bold.HasValue)
                {
                    font.Bold = request.Bold.Value ? -1 : 0;
                }
                if (request.Color.HasValue)
                {
                    font.Color = (Word.WdColor)request.Color.Value;
                }
                if (request.FontSize.HasValue)
                {
                    font.Size = (float)request.FontSize.Value;
                }
                if (!string.IsNullOrWhiteSpace(request.FontName))
                {
                    font.Name = request.FontName;
                    try
                    {
                        font.NameFarEast = request.FontName;
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[word] 替换格式 NameFarEast 设置失败：" + ex.Message);
                    }
                }
                if (request.UnderlineStyle.HasValue)
                {
                    font.Underline = request.UnderlineStyle.Value;
                }

                string firstFailure = string.Empty;
                if (request.Bold.HasValue)
                {
                    AddBoolVerification(
                        verification, ref firstFailure, "bold", request.Bold.Value,
                        WordFormatUtil.TriStateToken(font.Bold));
                }
                if (request.Color.HasValue)
                {
                    string expectedHex = StyleColorToHex(request.Color.Value);
                    string actualHex = StyleColorToHex((int)font.Color);
                    AddVerification(
                        verification,
                        ref firstFailure,
                        "color",
                        new JValue(expectedHex),
                        new JValue(actualHex),
                        string.Equals(expectedHex, actualHex, StringComparison.OrdinalIgnoreCase));
                }
                if (request.FontSize.HasValue)
                {
                    AddDoubleVerification(
                        verification, ref firstFailure, "fontSize", request.FontSize.Value,
                        WordFormatUtil.IsUndefined(font.Size)
                            ? (double?)null
                            : new double?((double)font.Size),
                        0.2);
                }
                if (!string.IsNullOrWhiteSpace(request.FontName))
                {
                    string actualName = font.Name ?? string.Empty;
                    string actualFarEast = font.NameFarEast ?? string.Empty;
                    bool nameMatched = string.Equals(
                                           actualName, request.FontName, StringComparison.OrdinalIgnoreCase) ||
                                       string.Equals(
                                           actualFarEast, request.FontName, StringComparison.OrdinalIgnoreCase);
                    AddVerification(
                        verification,
                        ref firstFailure,
                        "fontName",
                        new JValue(request.FontName),
                        new JObject
                        {
                            { "fontName", actualName },
                            { "fontNameFarEast", actualFarEast }
                        },
                        nameMatched);
                }
                if (request.UnderlineStyle.HasValue)
                {
                    Word.WdUnderline actualUnderline = font.Underline;
                    AddVerification(
                        verification,
                        ref firstFailure,
                        "underlineStyle",
                        new JValue(request.UnderlineStyle.Value.ToString()),
                        new JValue(actualUnderline.ToString()),
                        actualUnderline == request.UnderlineStyle.Value);
                }
                matched = string.IsNullOrWhiteSpace(firstFailure);
            }
            finally
            {
                ComHelper.Release(font);
            }
        }
    }
}
