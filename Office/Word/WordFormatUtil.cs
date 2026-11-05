using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    /// <summary>
    /// Word 字体/段落格式的读取、回读核验与名词映射。
    /// 所有"已应用"的判断都以 Word 对象模型回读为准，不凭设置语句未抛异常就报成功。
    /// </summary>
    internal static class WordFormatUtil
    {
        public const int UndefinedValue = 9999999;
        public const float PointsPerCentimeter = 28.35f;

        private static readonly Dictionary<string, int> ColorNames =
            new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                { "black", 0 },
                { "黑色", 0 },
                { "red", 255 },
                { "红色", 255 },
                { "darkred", 128 },
                { "深红", 128 },
                { "green", 32768 },
                { "绿色", 32768 },
                { "darkgreen", 13056 },
                { "深绿", 13056 },
                { "blue", 16711680 },
                { "蓝色", 16711680 },
                { "darkblue", 8388608 },
                { "深蓝", 8388608 },
                { "yellow", 65535 },
                { "黄色", 65535 },
                { "orange", 26367 },
                { "橙色", 26367 },
                { "gold", 52479 },
                { "金色", 52479 },
                { "purple", 8388736 },
                { "紫色", 8388736 },
                { "pink", 16711935 },
                { "粉色", 16711935 },
                { "brown", 13209 },
                { "棕色", 13209 },
                { "gray", 8421504 },
                { "grey", 8421504 },
                { "灰色", 8421504 },
                { "white", 16777215 },
                { "白色", 16777215 },
                { "teal", 8421376 },
                { "青色", 8421376 },
                { "auto", -16777216 },
                { "automatic", -16777216 },
                { "自动", -16777216 }
            };

        private static readonly Dictionary<string, Word.WdColorIndex> HighlightNames =
            new Dictionary<string, Word.WdColorIndex>(StringComparer.OrdinalIgnoreCase)
            {
                { "none", Word.WdColorIndex.wdNoHighlight },
                { "auto", Word.WdColorIndex.wdNoHighlight },
                { "无", Word.WdColorIndex.wdNoHighlight },
                { "black", Word.WdColorIndex.wdBlack },
                { "黑色", Word.WdColorIndex.wdBlack },
                { "blue", Word.WdColorIndex.wdBlue },
                { "蓝色", Word.WdColorIndex.wdBlue },
                { "turquoise", Word.WdColorIndex.wdTurquoise },
                { "青绿", Word.WdColorIndex.wdTurquoise },
                { "brightgreen", Word.WdColorIndex.wdBrightGreen },
                { "亮绿", Word.WdColorIndex.wdBrightGreen },
                { "pink", Word.WdColorIndex.wdPink },
                { "粉色", Word.WdColorIndex.wdPink },
                { "red", Word.WdColorIndex.wdRed },
                { "红色", Word.WdColorIndex.wdRed },
                { "yellow", Word.WdColorIndex.wdYellow },
                { "黄色", Word.WdColorIndex.wdYellow },
                { "white", Word.WdColorIndex.wdWhite },
                { "白色", Word.WdColorIndex.wdWhite },
                { "darkblue", Word.WdColorIndex.wdDarkBlue },
                { "深蓝", Word.WdColorIndex.wdDarkBlue },
                { "teal", Word.WdColorIndex.wdTeal },
                { "青色", Word.WdColorIndex.wdTeal },
                { "green", Word.WdColorIndex.wdGreen },
                { "绿色", Word.WdColorIndex.wdGreen },
                { "violet", Word.WdColorIndex.wdViolet },
                { "紫色", Word.WdColorIndex.wdViolet },
                { "darkred", Word.WdColorIndex.wdDarkRed },
                { "深红", Word.WdColorIndex.wdDarkRed },
                { "darkyellow", Word.WdColorIndex.wdDarkYellow },
                { "深黄", Word.WdColorIndex.wdDarkYellow },
                { "gray50", Word.WdColorIndex.wdGray50 },
                { "gray25", Word.WdColorIndex.wdGray25 }
            };

        public static bool IsUndefined(int value)
        {
            return value == UndefinedValue;
        }

        public static bool IsUndefined(float value)
        {
            return Math.Abs(value - (float)UndefinedValue) < 0.5f;
        }

        public static JToken TriStateToken(int value)
        {
            if (value == -1)
            {
                return new JValue(true);
            }
            if (value == 0)
            {
                return new JValue(false);
            }
            return JValue.CreateNull();
        }

        public static JObject CaptureFont(Word.Font font)
        {
            if (font == null)
            {
                return new JObject();
            }

            JObject json = new JObject();
            json["fontName"] = font.Name ?? string.Empty;
            json["fontNameFarEast"] = font.NameFarEast ?? string.Empty;
            json["fontNameAscii"] = font.NameAscii ?? string.Empty;
            json["fontSize"] = IsUndefined(font.Size) ? JValue.CreateNull() : new JValue((double)font.Size);
            json["bold"] = TriStateToken(font.Bold);
            json["italic"] = TriStateToken(font.Italic);
            json["underlineStyle"] = GetUnderlineStyleName(font.Underline);
            json["underlineCode"] = (int)font.Underline;
            json["color"] = (int)font.Color;
            json["colorHex"] = ColorToHex((int)font.Color);
            json["subscript"] = TriStateToken(font.Subscript);
            json["superscript"] = TriStateToken(font.Superscript);
            json["strikeThrough"] = TriStateToken(font.StrikeThrough);
            json["doubleStrikeThrough"] = TriStateToken(font.DoubleStrikeThrough);
            json["smallCaps"] = TriStateToken(font.SmallCaps);
            json["allCaps"] = TriStateToken(font.AllCaps);
            json["characterSpacing"] = IsUndefined(font.Spacing) ? JValue.CreateNull() : new JValue((double)font.Spacing);
            return json;
        }

        public static JObject CaptureRangeFormat(Word.Range range)
        {
            JObject json = new JObject();
            if (range == null)
            {
                return json;
            }

            Word.Font font = null;
            Word.ParagraphFormat paragraphFormat = null;
            try
            {
                font = range.Font;
                json["font"] = CaptureFont(font);
                json["style"] = WordComUtil.ReadStyleName(range);
                json["highlightColorIndex"] = (int)range.HighlightColorIndex;
                json["highlightColor"] = HighlightIndexToName(range.HighlightColorIndex);
                paragraphFormat = range.ParagraphFormat;
                json["alignment"] = GetAlignmentName(paragraphFormat.Alignment);
                json["alignmentCode"] = (int)paragraphFormat.Alignment;
                json["firstLineIndent"] = (double)paragraphFormat.FirstLineIndent;
                json["characterUnitFirstLineIndent"] = (double)paragraphFormat.CharacterUnitFirstLineIndent;
                json["spaceBefore"] = (double)paragraphFormat.SpaceBefore;
                json["spaceAfter"] = (double)paragraphFormat.SpaceAfter;
                json["lineSpacingRule"] = paragraphFormat.LineSpacingRule.ToString();
                json["lineSpacing"] = (double)paragraphFormat.LineSpacing;
            }
            finally
            {
                ComHelper.Release(paragraphFormat, font);
            }
            return json;
        }

        public static string GetUnderlineStyleName(Word.WdUnderline underline)
        {
            switch (underline)
            {
                case Word.WdUnderline.wdUnderlineNone:
                    return "none";
                case Word.WdUnderline.wdUnderlineSingle:
                    return "single";
                case Word.WdUnderline.wdUnderlineDouble:
                    return "double";
                case Word.WdUnderline.wdUnderlineThick:
                    return "thick";
                case Word.WdUnderline.wdUnderlineWavy:
                    return "wavy";
                case Word.WdUnderline.wdUnderlineDotted:
                    return "dotted";
                case Word.WdUnderline.wdUnderlineDash:
                    return "dash";
                case Word.WdUnderline.wdUnderlineDotDash:
                    return "dotdash";
                default:
                    return underline.ToString();
            }
        }

        public static Word.WdUnderline ParseUnderlineStyle(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                throw WordArguments.Invalid("underlineStyle 不能为空；支持 none/single/double/wavy/thick/dotted/dash/dotdash");
            }
            switch (raw.Trim().ToLowerInvariant())
            {
                case "none":
                case "无":
                    return Word.WdUnderline.wdUnderlineNone;
                case "single":
                case "单线":
                    return Word.WdUnderline.wdUnderlineSingle;
                case "double":
                case "双线":
                    return Word.WdUnderline.wdUnderlineDouble;
                case "wavy":
                case "波浪":
                    return Word.WdUnderline.wdUnderlineWavy;
                case "thick":
                case "粗线":
                    return Word.WdUnderline.wdUnderlineThick;
                case "dotted":
                case "点线":
                    return Word.WdUnderline.wdUnderlineDotted;
                case "dash":
                case "虚线":
                    return Word.WdUnderline.wdUnderlineDash;
                case "dotdash":
                case "点划线":
                    return Word.WdUnderline.wdUnderlineDotDash;
                default:
                    throw WordArguments.Invalid("不支持的 underlineStyle=" + raw +
                                                "；支持 none/single/double/wavy/thick/dotted/dash/dotdash");
            }
        }

        public static string GetAlignmentName(Word.WdParagraphAlignment alignment)
        {
            switch (alignment)
            {
                case Word.WdParagraphAlignment.wdAlignParagraphLeft:
                    return "left";
                case Word.WdParagraphAlignment.wdAlignParagraphCenter:
                    return "center";
                case Word.WdParagraphAlignment.wdAlignParagraphRight:
                    return "right";
                case Word.WdParagraphAlignment.wdAlignParagraphJustify:
                    return "justify";
                case Word.WdParagraphAlignment.wdAlignParagraphDistribute:
                    return "distribute";
                default:
                    return alignment.ToString();
            }
        }

        public static Word.WdParagraphAlignment ParseAlignment(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                throw WordArguments.Invalid("alignment 不能为空；支持 left/center/right/justify/distribute");
            }
            switch (raw.Trim().ToLowerInvariant())
            {
                case "left":
                case "左":
                case "左对齐":
                    return Word.WdParagraphAlignment.wdAlignParagraphLeft;
                case "center":
                case "centre":
                case "居中":
                case "中":
                    return Word.WdParagraphAlignment.wdAlignParagraphCenter;
                case "right":
                case "右":
                case "右对齐":
                    return Word.WdParagraphAlignment.wdAlignParagraphRight;
                case "justify":
                case "两端对齐":
                    return Word.WdParagraphAlignment.wdAlignParagraphJustify;
                case "distribute":
                case "分散对齐":
                    return Word.WdParagraphAlignment.wdAlignParagraphDistribute;
                default:
                    throw WordArguments.Invalid("不支持的 alignment=" + raw +
                                                "；支持 left/center/right/justify/distribute");
            }
        }

        public static int ParseColor(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                throw WordArguments.Invalid("color 不能为空；支持 #RRGGBB、red/green/blue/black 等颜色名");
            }

            string text = raw.Trim();
            int namedColor;
            if (ColorNames.TryGetValue(text, out namedColor))
            {
                return namedColor;
            }

            if (text.StartsWith("#", StringComparison.Ordinal))
            {
                text = text.Substring(1);
            }
            if (text.Length == 3)
            {
                text = new string(new char[]
                {
                    text[0], text[0], text[1], text[1], text[2], text[2]
                });
            }
            if (text.Length == 6)
            {
                int rgb;
                if (int.TryParse(text, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out rgb))
                {
                    // 用户输入为 #RRGGBB，Word 的 WdColor 使用 BGR 内存顺序。
                    return ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF);
                }
            }
            throw WordArguments.Invalid("不支持的 color=" + raw +
                                        "；支持 #RRGGBB、red/green/blue/black/yellow/gray 等颜色名或 auto");
        }

        public static string ColorToHex(int color)
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

        public static Word.WdColorIndex ParseHighlight(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                throw WordArguments.Invalid("highlightColor 不能为空；支持 none/yellow/brightgreen/turquoise/pink/red 等色名或 wdColorIndex 整数");
            }

            string text = raw.Trim();
            Word.WdColorIndex index;
            if (HighlightNames.TryGetValue(text, out index))
            {
                return index;
            }

            int numeric;
            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out numeric))
            {
                return (Word.WdColorIndex)numeric;
            }
            throw WordArguments.Invalid("不支持的 highlightColor=" + raw +
                                        "；支持 none/yellow/brightgreen/turquoise/pink/red/blue/green 等色名或 wdColorIndex 整数");
        }

        public static string HighlightIndexToName(Word.WdColorIndex index)
        {
            switch (index)
            {
                case Word.WdColorIndex.wdNoHighlight:
                    return "none";
                case Word.WdColorIndex.wdBlack:
                    return "black";
                case Word.WdColorIndex.wdBlue:
                    return "blue";
                case Word.WdColorIndex.wdTurquoise:
                    return "turquoise";
                case Word.WdColorIndex.wdBrightGreen:
                    return "brightgreen";
                case Word.WdColorIndex.wdPink:
                    return "pink";
                case Word.WdColorIndex.wdRed:
                    return "red";
                case Word.WdColorIndex.wdYellow:
                    return "yellow";
                case Word.WdColorIndex.wdWhite:
                    return "white";
                case Word.WdColorIndex.wdDarkBlue:
                    return "darkblue";
                case Word.WdColorIndex.wdTeal:
                    return "teal";
                case Word.WdColorIndex.wdGreen:
                    return "green";
                case Word.WdColorIndex.wdViolet:
                    return "violet";
                case Word.WdColorIndex.wdDarkRed:
                    return "darkred";
                case Word.WdColorIndex.wdDarkYellow:
                    return "darkyellow";
                case Word.WdColorIndex.wdGray50:
                    return "gray50";
                case Word.WdColorIndex.wdGray25:
                    return "gray25";
                default:
                    return index.ToString();
            }
        }

        public static bool CloseEnough(double left, double right, double tolerance)
        {
            return Math.Abs(left - right) <= tolerance;
        }

        public static string NormalizeForWord(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }
            return text.Replace("\r\n", "\r").Replace("\n", "\r");
        }

        public static string NormalizeForCompare(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }
            return text.Replace("\r\n", "\n").Replace('\r', '\n');
        }

        public static double RoundPoints(double value)
        {
            return Math.Round(value, 2, MidpointRounding.AwayFromZero);
        }
    }
}
