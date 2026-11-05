using System;
using System.Reflection;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    internal sealed class WordSizeResult
    {
        public double WidthPt { get; set; }

        public double HeightPt { get; set; }

        public string Warning { get; set; }

        public WordSizeResult()
        {
            Warning = string.Empty;
        }
    }

    /// <summary>图片尺寸计算与布局枚举映射。尺寸统一用磅，默认按原比例计算以防拉伸。</summary>
    internal static class WordImageUtil
    {
        public static WordSizeResult ComputeSize(
            double currentWidth,
            double currentHeight,
            double? widthPt,
            double? heightPt,
            double? scalePercent,
            double? widthLimitPt,
            bool lockAspect)
        {
            if (currentWidth <= 0 || currentHeight <= 0)
            {
                throw WordArguments.Invalid(
                    "图片当前宽高无效（" + currentWidth + "×" + currentHeight +
                    " 磅），无法计算缩放。");
            }
            if (scalePercent.HasValue && (widthPt.HasValue || heightPt.HasValue || widthLimitPt.HasValue))
            {
                throw WordArguments.Invalid("scalePercent 不能与 widthPt/heightPt/widthLimitPt 同时使用。");
            }
            if (scalePercent.HasValue && scalePercent.Value <= 0)
            {
                throw WordArguments.Invalid("scalePercent 必须大于 0。");
            }
            if (widthLimitPt.HasValue && widthLimitPt.Value <= 0)
            {
                throw WordArguments.Invalid("widthLimitPt 必须大于 0。");
            }

            double aspect = currentWidth / currentHeight;
            double newWidth = currentWidth;
            double newHeight = currentHeight;
            string warning = string.Empty;

            if (scalePercent.HasValue)
            {
                newWidth = currentWidth * scalePercent.Value / 100d;
                newHeight = currentHeight * scalePercent.Value / 100d;
            }
            else if (widthPt.HasValue || heightPt.HasValue)
            {
                if (widthPt.HasValue && heightPt.HasValue)
                {
                    if (lockAspect)
                    {
                        newWidth = widthPt.Value;
                        newHeight = widthPt.Value / aspect;
                        double requestedRatio = widthPt.Value / heightPt.Value;
                        if (Math.Abs(requestedRatio - aspect) > 0.01)
                        {
                            warning = "lockAspect=true：请求的宽高比与原始图片不一致，已按 widthPt=" +
                                      widthPt.Value.ToString("0.##") + " 磅和原比例计算高度为 " +
                                      newHeight.ToString("0.##") + " 磅。";
                        }
                    }
                    else
                    {
                        newWidth = widthPt.Value;
                        newHeight = heightPt.Value;
                    }
                }
                else if (widthPt.HasValue)
                {
                    newWidth = widthPt.Value;
                    newHeight = widthPt.Value / aspect;
                }
                else
                {
                    newHeight = heightPt.Value;
                    newWidth = heightPt.Value * aspect;
                }
            }

            if (widthLimitPt.HasValue && newWidth > widthLimitPt.Value)
            {
                double scale = widthLimitPt.Value / newWidth;
                newWidth = widthLimitPt.Value;
                newHeight = newHeight * scale;
                warning = string.IsNullOrWhiteSpace(warning)
                    ? "已按 widthLimitPt=" + widthLimitPt.Value.ToString("0.##") + " 磅等比缩小。"
                    : warning + " 另按 widthLimitPt 等比缩小。";
            }
            if (newWidth <= 0 || newHeight <= 0)
            {
                throw WordArguments.Invalid("缩放计算结果无效：" + newWidth + "×" + newHeight + " 磅。");
            }

            WordSizeResult result = new WordSizeResult();
            result.WidthPt = newWidth;
            result.HeightPt = newHeight;
            result.Warning = warning;
            return result;
        }

        public static Word.WdWrapType ParseWrapType(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return Word.WdWrapType.wdWrapSquare;
            }
            switch (raw.Trim().ToLowerInvariant())
            {
                case "inline":
                case "嵌入":
                    return Word.WdWrapType.wdWrapInline;
                case "topbottom":
                case "上下型":
                    return Word.WdWrapType.wdWrapTopBottom;
                case "square":
                case "四周型":
                    return Word.WdWrapType.wdWrapSquare;
                case "tight":
                case "紧密型":
                    return Word.WdWrapType.wdWrapTight;
                case "through":
                case "穿越型":
                    return Word.WdWrapType.wdWrapThrough;
                case "behind":
                case "衬于文字下方":
                    return Word.WdWrapType.wdWrapBehind;
                case "front":
                case "浮于文字上方":
                    return Word.WdWrapType.wdWrapFront;
                default:
                    throw WordArguments.Invalid(
                        "wrap 只支持 inline/topbottom/square/tight/through/behind/front，收到 " + raw + "。");
            }
        }

        public static Word.WdRelativeHorizontalPosition ParseRelativeHorizontal(string raw)
        {
            switch ((raw ?? "margin").Trim().ToLowerInvariant())
            {
                case "page":
                case "页面":
                    return Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                case "paragraph":
                case "段落":
                    return Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionColumn;
                case "margin":
                case "页边距":
                default:
                    return Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
            }
        }

        public static Word.WdRelativeVerticalPosition ParseRelativeVertical(string raw)
        {
            switch ((raw ?? "margin").Trim().ToLowerInvariant())
            {
                case "page":
                case "页面":
                    return Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                case "paragraph":
                case "段落":
                    return Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph;
                case "margin":
                case "页边距":
                default:
                    return Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin;
            }
        }

        public static int ReadFloatingTypeCode(Word.Shape shape)
        {
            if (shape == null)
            {
                return 0;
            }
            try
            {
                object value = shape.GetType().InvokeMember(
                    "Type", BindingFlags.GetProperty, null, shape, null);
                if (value == null)
                {
                    return 0;
                }
                return Convert.ToInt32(value);
            }
            catch
            {
                return 0;
            }
        }

        public static string ReadFloatingType(Word.Shape shape)
        {
            int code = ReadFloatingTypeCode(shape);
            switch (code)
            {
                case 1: return "msoAutoShape";
                case 2: return "msoCallout";
                case 3: return "msoChart";
                case 4: return "msoComment";
                case 5: return "msoFreeform";
                case 6: return "msoGroup";
                case 7: return "msoEmbeddedOLEObject";
                case 8: return "msoFormControl";
                case 9: return "msoLine";
                case 10: return "msoLinkedOLEObject";
                case 11: return "msoLinkedPicture";
                case 12: return "msoOLEControlObject";
                case 13: return "msoPicture";
                case 14: return "msoPlaceholder";
                case 15: return "msoTextEffect";
                case 16: return "msoMedia";
                case 17: return "msoTextBox";
                case 18: return "msoScriptAnchor";
                case 19: return "msoTable";
                case 20: return "msoCanvas";
                case 21: return "msoDiagram";
                case 22: return "msoInk";
                case 23: return "msoInkComment";
                case 24: return "msoSmartArt";
                case 25: return "msoSlicer";
                case 26: return "msoWebVideo";
                case 27: return "msoContentApp";
                case 28: return "msoGraphic";
                case 29: return "msoLinkedGraphic";
                case 30: return "mso3DModel";
                case 31: return "msoLinked3DModel";
                default: return code == 0 ? "unknown" : code.ToString();
            }
        }

        public static double ApplyInlineSize(Word.InlineShape shape, WordSizeResult size)
        {
            shape.Width = (float)size.WidthPt;
            shape.Height = (float)size.HeightPt;
            return shape.Width;
        }

        public static double ApplyFloatingSize(Word.Shape shape, WordSizeResult size)
        {
            shape.Width = (float)size.WidthPt;
            shape.Height = (float)size.HeightPt;
            return shape.Width;
        }

        public static bool SizeMatches(double actualWidth, double actualHeight,
            double expectedWidth, double expectedHeight, double tolerance)
        {
            return WordFormatUtil.CloseEnough(actualWidth, expectedWidth, tolerance) &&
                   WordFormatUtil.CloseEnough(actualHeight, expectedHeight, tolerance);
        }

        public static JObject CaptureInline(Word.Document document, Word.InlineShape shape, int index)
        {
            JObject json = new JObject();
            Word.Range range = null;
            try
            {
                range = shape.Range;
                json["collection"] = "inline";
                json["index"] = index;
                json["type"] = shape.Type.ToString();
                json["typeCode"] = (int)shape.Type;
                json["widthPt"] = (double)shape.Width;
                json["heightPt"] = (double)shape.Height;
                json["alternativeText"] = shape.AlternativeText ?? string.Empty;
                json["rangeStart"] = range == null ? 0 : range.Start;
                json["rangeEnd"] = range == null ? 0 : range.End;
                json["anchorParagraphIndex"] = WordComUtil.FindParagraphIndexAt(
                    document, range == null ? 0 : range.Start);
                return json;
            }
            finally
            {
                ComHelper.Release(range);
            }
        }

        public static JObject CaptureFloating(Word.Document document, Word.Shape shape, int index)
        {
            JObject json = new JObject();
            Word.Range anchor = null;
            Word.WrapFormat wrap = null;
            try
            {
                anchor = shape.Anchor;
                wrap = shape.WrapFormat;
                json["collection"] = "floating";
                json["index"] = index;
                json["name"] = shape.Name ?? string.Empty;
                json["type"] = ReadFloatingType(shape);
                json["typeCode"] = ReadFloatingTypeCode(shape);
                json["widthPt"] = (double)shape.Width;
                json["heightPt"] = (double)shape.Height;
                json["leftPt"] = (double)shape.Left;
                json["topPt"] = (double)shape.Top;
                json["alternativeText"] = shape.AlternativeText ?? string.Empty;
                json["wrapType"] = wrap == null ? string.Empty : wrap.Type.ToString();
                json["wrapTypeCode"] = wrap == null ? 0 : (int)wrap.Type;
                json["anchorParagraphIndex"] = WordComUtil.FindParagraphIndexAt(
                    document, anchor == null ? 0 : anchor.Start);
                json["anchorStart"] = anchor == null ? 0 : anchor.Start;
                return json;
            }
            finally
            {
                ComHelper.Release(wrap, anchor);
            }
        }
    }
}
