using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    /// <summary>
    /// word_format_text 的字体属性请求。所有布尔属性都是 nullable：未传不改。
    /// ApplyAndVerify 逐属性回读，返回实际生效值与 matched 结果。
    /// </summary>
    internal sealed class WordFontRequest
    {
        public bool? Bold { get; set; }

        public bool? Italic { get; set; }

        public string UnderlineStyle { get; set; }

        public bool? UnderlineEnabled { get; set; }

        public bool? Subscript { get; set; }

        public bool? Superscript { get; set; }

        public bool? StrikeThrough { get; set; }

        public bool? DoubleStrikeThrough { get; set; }

        public double? FontSize { get; set; }

        public string FontName { get; set; }

        public int? Color { get; set; }

        public Word.WdColorIndex? HighlightColor { get; set; }

        public bool? SmallCaps { get; set; }

        public bool? AllCaps { get; set; }

        public double? CharacterSpacing { get; set; }

        public bool HasAny
        {
            get
            {
                return Bold.HasValue || Italic.HasValue ||
                       !string.IsNullOrWhiteSpace(UnderlineStyle) || UnderlineEnabled.HasValue ||
                       Subscript.HasValue || Superscript.HasValue ||
                       StrikeThrough.HasValue || DoubleStrikeThrough.HasValue ||
                       FontSize.HasValue || !string.IsNullOrWhiteSpace(FontName) ||
                       Color.HasValue || HighlightColor.HasValue ||
                       SmallCaps.HasValue || AllCaps.HasValue || CharacterSpacing.HasValue;
            }
        }
    }

    internal sealed class WordFontApplyResult
    {
        public JObject Applied { get; internal set; }

        public JArray Verification { get; private set; }

        public bool AllMatched { get; private set; }

        public bool Changed { get; internal set; }

        public string FirstFailure { get; private set; }

        public WordFontApplyResult()
        {
            Applied = new JObject();
            Verification = new JArray();
            AllMatched = true;
            FirstFailure = string.Empty;
        }

        internal void AddVerification(string property, JToken expected, JToken actual, bool matched)
        {
            JObject item = new JObject();
            item["property"] = property;
            item["expected"] = expected == null ? JValue.CreateNull() : expected.DeepClone();
            item["actual"] = actual == null ? JValue.CreateNull() : actual.DeepClone();
            item["matched"] = matched;
            Verification.Add(item);
            if (!matched)
            {
                AllMatched = false;
                if (string.IsNullOrWhiteSpace(FirstFailure))
                {
                    FirstFailure = property + " 期望 " + (expected == null ? "null" : expected.ToString()) +
                                   "，回读 " + (actual == null ? "null" : actual.ToString());
                }
            }
        }
    }

    internal static class WordFontApplier
    {
        public static WordFontApplyResult ApplyAndVerify(Word.Range range, WordFontRequest request)
        {
            if (range == null)
            {
                throw WordArguments.Invalid("目标范围为空，无法设置字体格式。");
            }
            if (request == null || !request.HasAny)
            {
                throw WordArguments.Invalid(
                    "word_format_text 至少需要一个格式属性。支持 " +
                    "bold/italic/underlineStyle/subscript/superscript/strikethrough/" +
                    "doubleStrikethrough/fontSize/fontName/color/highlightColor/smallCaps/allCaps/characterSpacing");
            }

            WordFontApplyResult result = new WordFontApplyResult();
            Word.Font font = null;
            try
            {
                font = range.Font;
                JObject beforeFont = WordFormatUtil.CaptureFont(font);
                Word.WdColorIndex beforeHighlight = range.HighlightColorIndex;

                if (request.Bold.HasValue)
                {
                    font.Bold = request.Bold.Value ? -1 : 0;
                }
                if (request.Italic.HasValue)
                {
                    font.Italic = request.Italic.Value ? -1 : 0;
                }
                if (!string.IsNullOrWhiteSpace(request.UnderlineStyle))
                {
                    font.Underline = WordFormatUtil.ParseUnderlineStyle(request.UnderlineStyle);
                }
                else if (request.UnderlineEnabled.HasValue)
                {
                    font.Underline = request.UnderlineEnabled.Value
                        ? Word.WdUnderline.wdUnderlineSingle
                        : Word.WdUnderline.wdUnderlineNone;
                }
                if (request.Subscript.HasValue || request.Superscript.HasValue)
                {
                    if (request.Subscript == true && request.Superscript == true)
                    {
                        throw WordArguments.Invalid("subscript 与 superscript 不能同时为 true。");
                    }
                    if (request.Subscript.HasValue)
                    {
                        font.Subscript = request.Subscript.Value ? -1 : 0;
                    }
                    if (request.Superscript.HasValue)
                    {
                        font.Superscript = request.Superscript.Value ? -1 : 0;
                    }
                    if (request.Subscript == true)
                    {
                        font.Superscript = 0;
                    }
                    if (request.Superscript == true)
                    {
                        font.Subscript = 0;
                    }
                }
                if (request.StrikeThrough.HasValue)
                {
                    font.StrikeThrough = request.StrikeThrough.Value ? -1 : 0;
                }
                if (request.DoubleStrikeThrough.HasValue)
                {
                    font.DoubleStrikeThrough = request.DoubleStrikeThrough.Value ? -1 : 0;
                }
                if (request.SmallCaps.HasValue)
                {
                    font.SmallCaps = request.SmallCaps.Value ? -1 : 0;
                }
                if (request.AllCaps.HasValue)
                {
                    font.AllCaps = request.AllCaps.Value ? -1 : 0;
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
                        Console.Error.WriteLine("[word] 设置 NameFarEast 失败（已保留 Name）：" + ex.Message);
                    }
                }
                if (request.Color.HasValue)
                {
                    font.Color = (Word.WdColor)request.Color.Value;
                }
                if (request.HighlightColor.HasValue)
                {
                    range.HighlightColorIndex = request.HighlightColor.Value;
                }
                if (request.CharacterSpacing.HasValue)
                {
                    font.Spacing = (float)request.CharacterSpacing.Value;
                }

                JObject afterFont = WordFormatUtil.CaptureFont(font);
                Word.WdColorIndex afterHighlight = range.HighlightColorIndex;
                result.Applied = afterFont;

                VerifyBoolean(result, beforeFont, afterFont, "bold", request.Bold);
                VerifyBoolean(result, beforeFont, afterFont, "italic", request.Italic);
                if (!string.IsNullOrWhiteSpace(request.UnderlineStyle))
                {
                    VerifyString(result, beforeFont, afterFont, "underlineStyle",
                        new JValue(WordFormatUtil.GetUnderlineStyleName(
                            WordFormatUtil.ParseUnderlineStyle(request.UnderlineStyle))));
                }
                else if (request.UnderlineEnabled.HasValue)
                {
                    VerifyString(result, beforeFont, afterFont, "underlineStyle",
                        new JValue(request.UnderlineEnabled.Value ? "single" : "none"));
                }
                VerifyBoolean(result, beforeFont, afterFont, "subscript", request.Subscript);
                VerifyBoolean(result, beforeFont, afterFont, "superscript", request.Superscript);
                VerifyBoolean(result, beforeFont, afterFont, "strikeThrough", request.StrikeThrough);
                VerifyBoolean(result, beforeFont, afterFont, "doubleStrikeThrough", request.DoubleStrikeThrough);
                VerifyBoolean(result, beforeFont, afterFont, "smallCaps", request.SmallCaps);
                VerifyBoolean(result, beforeFont, afterFont, "allCaps", request.AllCaps);
                VerifyDouble(result, beforeFont, afterFont, "fontSize", request.FontSize, 0.2);
                VerifyFontName(result, beforeFont, afterFont, request.FontName);
                if (request.Color.HasValue)
                {
                    string expectedHex = WordFormatUtil.ColorToHex(request.Color.Value);
                    JToken actual = afterFont["colorHex"];
                    bool matched = actual != null && string.Equals(
                        actual.ToString(), expectedHex, StringComparison.OrdinalIgnoreCase);
                    result.AddVerification("color", new JValue(expectedHex), actual, matched);
                    result.Changed = result.Changed ||
                                     !JToken.DeepEquals(beforeFont["colorHex"], afterFont["colorHex"]);
                }
                if (request.HighlightColor.HasValue)
                {
                    JToken expected = new JValue((int)request.HighlightColor.Value);
                    JToken actual = new JValue((int)afterHighlight);
                    bool matched = (int)afterHighlight == (int)request.HighlightColor.Value;
                    result.AddVerification("highlightColor", expected, actual, matched);
                    result.Changed = result.Changed || beforeHighlight != afterHighlight;
                }
                VerifyDouble(result, beforeFont, afterFont, "characterSpacing", request.CharacterSpacing, 0.2);
                return result;
            }
            finally
            {
                ComHelper.Release(font);
            }
        }

        private static void VerifyBoolean(
            WordFontApplyResult result,
            JObject before,
            JObject after,
            string property,
            bool? expected)
        {
            if (!expected.HasValue)
            {
                return;
            }
            JToken expectedToken = new JValue(expected.Value);
            JToken actual = after[property];
            bool matched = actual != null && actual.Type == JTokenType.Boolean &&
                           actual.Value<bool>() == expected.Value;
            result.AddVerification(property, expectedToken, actual, matched);
            result.Changed = result.Changed || !JToken.DeepEquals(before[property], after[property]);
        }

        private static void VerifyString(
            WordFontApplyResult result,
            JObject before,
            JObject after,
            string property,
            JToken expected)
        {
            JToken actual = after[property];
            bool matched = actual != null && string.Equals(
                actual.ToString(), expected.ToString(), StringComparison.OrdinalIgnoreCase);
            result.AddVerification(property, expected, actual, matched);
            result.Changed = result.Changed || !JToken.DeepEquals(before[property], after[property]);
        }

        private static void VerifyDouble(
            WordFontApplyResult result,
            JObject before,
            JObject after,
            string property,
            double? expected,
            double tolerance)
        {
            if (!expected.HasValue)
            {
                return;
            }
            JToken expectedToken = new JValue(expected.Value);
            JToken actual = after[property];
            bool matched = false;
            if (actual != null && actual.Type != JTokenType.Null)
            {
                try
                {
                    matched = WordFormatUtil.CloseEnough(actual.Value<double>(), expected.Value, tolerance);
                }
                catch
                {
                    matched = false;
                }
            }
            result.AddVerification(property, expectedToken, actual, matched);
            result.Changed = result.Changed || !JToken.DeepEquals(before[property], after[property]);
        }

        private static void VerifyFontName(
            WordFontApplyResult result,
            JObject before,
            JObject after,
            string expected)
        {
            if (string.IsNullOrWhiteSpace(expected))
            {
                return;
            }
            JToken actualName = after["fontName"];
            JToken actualFarEast = after["fontNameFarEast"];
            bool matched = false;
            if (actualName != null &&
                string.Equals(actualName.ToString(), expected, StringComparison.OrdinalIgnoreCase))
            {
                matched = true;
            }
            if (actualFarEast != null &&
                string.Equals(actualFarEast.ToString(), expected, StringComparison.OrdinalIgnoreCase))
            {
                matched = true;
            }
            JObject actual = new JObject();
            actual["fontName"] = actualName == null ? JValue.CreateNull() : actualName.DeepClone();
            actual["fontNameFarEast"] = actualFarEast == null ? JValue.CreateNull() : actualFarEast.DeepClone();
            result.AddVerification("fontName", new JValue(expected), actual, matched);
            result.Changed = result.Changed ||
                             !JToken.DeepEquals(before["fontName"], after["fontName"]) ||
                             !JToken.DeepEquals(before["fontNameFarEast"], after["fontNameFarEast"]);
        }
    }
}
