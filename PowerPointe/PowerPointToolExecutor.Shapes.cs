using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net.Http;
using Newtonsoft.Json.Linq;
using Office;
using Office.Config;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointe
{
    public sealed partial class PowerPointToolExecutor
    {
        private const long MaxRemoteImageBytes = 52428800L;

        private static readonly HashSet<string> ImageExtensions =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf"
            };

        private static OperationResult ExecuteInsertImage(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.Shape shape = null;
            string downloadedFile = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                string source = PowerPointArgumentReader.RequireString(arguments, "source");
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                bool allowRemote = PowerPointArgumentReader.GetBool(arguments, false, "allowRemote");
                string localPath;
                string sourceError;
                if (IsRemoteUrl(source))
                {
                    if (!allowRemote)
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "source 是远程 URL；如确认允许联网下载，请显式传 allowRemote=true。");
                    }
                    if (!TryDownloadImage(source, out localPath, out downloadedFile, out sourceError))
                    {
                        return Fail(OfficeErrorCode.FILE_NOT_FOUND, sourceError);
                    }
                }
                else if (!WorkspacePathResolver.TryResolve(source, out localPath, out sourceError))
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, sourceError);
                }
                if (!File.Exists(localPath))
                {
                    return Fail(OfficeErrorCode.FILE_NOT_FOUND, "图片文件不存在：" + localPath);
                }
                string extension = Path.GetExtension(localPath);
                if (!ImageExtensions.Contains(extension))
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "不支持的图片扩展名「" + extension + "」。支持 png/jpg/jpeg/gif/bmp/tif/tiff/emf/wmf。");
                }

                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                dynamic shapes = slide.Shapes;
                object pictureObject = shapes.AddPicture(localPath, 0, -1, 0, 0, -1, -1);
                shape = pictureObject as PowerPoint.Shape;
                if (shape == null)
                {
                    return Fail(OfficeErrorCode.INTERNAL_ERROR, "PowerPoint 未返回已插入的图片对象。");
                }

                float originalWidth = shape.Width;
                float originalHeight = shape.Height;
                if (originalWidth <= 0 || originalHeight <= 0)
                {
                    return Fail(OfficeErrorCode.INTERNAL_ERROR, "无法读取图片原始尺寸。");
                }

                float boxWidth = PowerPointArgumentReader.GetFloat(arguments, 0f, "width");
                float boxHeight = PowerPointArgumentReader.GetFloat(arguments, 0f, "height");
                string fit = PowerPointArgumentReader.GetTrimmedString(arguments, "fit").ToLowerInvariant();
                if (fit.Length == 0)
                {
                    fit = "contain";
                }
                if (fit != "contain" && fit != "cover" && fit != "stretch")
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "fit 只能是 contain、cover 或 stretch。");
                }

                float targetWidth;
                float targetHeight;
                float cropLeft = 0f;
                float cropTop = 0f;
                float cropRight = 0f;
                float cropBottom = 0f;
                bool coverCropped = false;
                if (boxWidth > 0 && boxHeight > 0)
                {
                    if (fit == "stretch")
                    {
                        targetWidth = boxWidth;
                        targetHeight = boxHeight;
                    }
                    else if (fit == "cover")
                    {
                        float scale = Math.Max(boxWidth / originalWidth, boxHeight / originalHeight);
                        float displayWidth = originalWidth * scale;
                        float displayHeight = originalHeight * scale;
                        cropLeft = (displayWidth - boxWidth) / 2f / scale;
                        cropRight = cropLeft;
                        cropTop = (displayHeight - boxHeight) / 2f / scale;
                        cropBottom = cropTop;
                        if (cropLeft > 0.01f || cropTop > 0.01f)
                        {
                            dynamic cropPicture = shape.PictureFormat;
                            cropPicture.CropLeft = cropLeft;
                            cropPicture.CropRight = cropRight;
                            cropPicture.CropTop = cropTop;
                            cropPicture.CropBottom = cropBottom;
                            coverCropped = true;
                        }
                        targetWidth = boxWidth;
                        targetHeight = boxHeight;
                    }
                    else
                    {
                        float scale = Math.Min(boxWidth / originalWidth, boxHeight / originalHeight);
                        targetWidth = originalWidth * scale;
                        targetHeight = originalHeight * scale;
                    }
                }
                else if (boxWidth > 0)
                {
                    targetWidth = boxWidth;
                    targetHeight = originalHeight * (boxWidth / originalWidth);
                }
                else if (boxHeight > 0)
                {
                    targetHeight = boxHeight;
                    targetWidth = originalWidth * (boxHeight / originalHeight);
                }
                else
                {
                    targetWidth = originalWidth;
                    targetHeight = originalHeight;
                }

                float slideWidth = PowerPointComUtil.SafeFloat(delegate { return presentation.PageSetup.SlideWidth; });
                float slideHeight = PowerPointComUtil.SafeFloat(delegate { return presentation.PageSetup.SlideHeight; });
                bool hasX = HasArgument(arguments, "x");
                bool hasY = HasArgument(arguments, "y");
                float left = hasX
                    ? PowerPointArgumentReader.GetFloat(arguments, 0f, "x")
                    : (slideWidth - targetWidth) / 2f;
                float top = hasY
                    ? PowerPointArgumentReader.GetFloat(arguments, 0f, "y")
                    : (slideHeight - targetHeight) / 2f;

                shape.Width = targetWidth;
                shape.Height = targetHeight;
                shape.Left = left;
                shape.Top = top;

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["shapeId"] = PowerPointComUtil.SafeInt(delegate
                {
                    dynamic dynamicShape = shape;
                    return Convert.ToInt32((object)dynamicShape.Id);
                });
                data["shapeName"] = shape.Name;
                data["source"] = source;
                data["fit"] = fit;
                data["x"] = shape.Left;
                data["y"] = shape.Top;
                data["actualWidth"] = shape.Width;
                data["actualHeight"] = shape.Height;
                data["originalWidth"] = originalWidth;
                data["originalHeight"] = originalHeight;
                data["coverCropped"] = coverCropped;
                if (coverCropped)
                {
                    data["cropLeft"] = cropLeft;
                    data["cropTop"] = cropTop;
                    data["cropRight"] = cropRight;
                    data["cropBottom"] = cropBottom;
                }
                return Success(
                    "已在第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片插入图片：" + Path.GetFileName(localPath) + "。",
                    "write",
                    true,
                    SlideRef(presentation, slide) + "!" + shape.Name,
                    data);
            }
            finally
            {
                ComHelper.Release(shape, slide, presentation);
                if (!string.IsNullOrWhiteSpace(downloadedFile))
                {
                    TryDeleteFile(downloadedFile);
                }
            }
        }

        private static OperationResult ExecuteSetShapeFormat(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.Shape shape = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                int shapeId = PowerPointArgumentReader.GetInt(arguments, -1, "shapeId");
                if (shapeId < 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少必填参数 shapeId。");
                }
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                shape = PowerPointComUtil.FindShape(slide, shapeId);
                if (shape == null)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                        " 张幻灯片上找不到 shapeId=" + shapeId.ToString(CultureInfo.InvariantCulture) + "。");
                }

                JArray applied = new JArray();
                JArray warnings = new JArray();
                string fillColor = PowerPointArgumentReader.GetTrimmedString(arguments, "fillColor");
                if (fillColor.Length > 0)
                {
                    int oleColor;
                    if (!PowerPointComUtil.TryParseColor(fillColor, out oleColor))
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "fillColor 不是有效颜色：" + fillColor);
                    }
                    dynamic fill = shape.Fill;
                    fill.Solid();
                    fill.Visible = -1;
                    fill.ForeColor.RGB = oleColor;
                    applied.Add("fillColor");
                }

                if (HasArgument(arguments, "fillTransparency"))
                {
                    float transparency = PowerPointArgumentReader.GetFloat(arguments, 0f, "fillTransparency");
                    if (transparency < 0f || transparency > 1f)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "fillTransparency 必须在 0..1 之间。");
                    }
                    dynamic fill = shape.Fill;
                    fill.Transparency = transparency;
                    applied.Add("fillTransparency");
                }

                string lineColor = PowerPointArgumentReader.GetTrimmedString(arguments, "lineColor");
                if (lineColor.Length > 0)
                {
                    int oleColor;
                    if (!PowerPointComUtil.TryParseColor(lineColor, out oleColor))
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "lineColor 不是有效颜色：" + lineColor);
                    }
                    dynamic line = shape.Line;
                    line.Visible = -1;
                    line.ForeColor.RGB = oleColor;
                    applied.Add("lineColor");
                }

                if (HasArgument(arguments, "lineWidth"))
                {
                    float lineWidth = PowerPointArgumentReader.GetFloat(arguments, 1f, "lineWidth");
                    if (lineWidth <= 0)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "lineWidth 必须大于 0。");
                    }
                    dynamic line = shape.Line;
                    line.Weight = lineWidth;
                    applied.Add("lineWidth");
                }

                string lineDash = PowerPointArgumentReader.GetTrimmedString(arguments, "lineDash");
                if (lineDash.Length > 0)
                {
                    dynamic line = shape.Line;
                    line.DashStyle = PowerPointComUtil.MapLineDashStyle(lineDash);
                    applied.Add("lineDash");
                }

                if (HasArgument(arguments, "cornerRadius"))
                {
                    float radius = PowerPointArgumentReader.GetFloat(arguments, 0f, "cornerRadius");
                    if (radius <= 0)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "cornerRadius 必须大于 0。");
                    }
                    try
                    {
                        dynamic adjustments = shape.Adjustments;
                        int count = Convert.ToInt32((object)adjustments.Count);
                        if (count <= 0)
                        {
                            warnings.Add("该形状没有可调圆角（仅圆角矩形等形状支持）。");
                        }
                        else
                        {
                            float ratio = radius / Math.Min(shape.Width, shape.Height);
                            adjustments[1] = Math.Max(0f, Math.Min(0.5f, ratio));
                            applied.Add("cornerRadius");
                        }
                    }
                    catch (Exception ex)
                    {
                        warnings.Add("设置圆角失败：" + ex.Message);
                    }
                }

                if (HasArgument(arguments, "rotation"))
                {
                    float rotation = PowerPointArgumentReader.GetFloat(arguments, 0f, "rotation");
                    shape.Rotation = rotation;
                    applied.Add("rotation");
                }

                if (applied.Count == 0 && warnings.Count == 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "没有提供任何可应用的格式属性。");
                }

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["shapeId"] = shapeId;
                data["shapeName"] = shape.Name;
                data["applied"] = applied;
                if (warnings.Count > 0)
                {
                    data["warnings"] = warnings;
                }
                return Success(
                    "已修改第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片上的形状 " + shape.Name + "（" +
                    applied.Count.ToString(CultureInfo.InvariantCulture) + " 项属性）。",
                    "write",
                    applied.Count > 0,
                    SlideRef(presentation, slide) + "!" + shape.Name,
                    data);
            }
            finally
            {
                ComHelper.Release(shape, slide, presentation);
            }
        }

        private static OperationResult ExecuteFormatText(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                int shapeId = PowerPointArgumentReader.GetInt(arguments, -1, "shapeId");
                List<PowerPoint.Shape> shapes = new List<PowerPoint.Shape>();
                try
                {
                    if (shapeId >= 0)
                    {
                        PowerPoint.Shape target = PowerPointComUtil.FindShape(slide, shapeId);
                        if (target == null)
                        {
                            return Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "找不到 shapeId=" + shapeId.ToString(CultureInfo.InvariantCulture) + "。");
                        }
                        shapes.Add(target);
                    }
                    else
                    {
                        PowerPoint.Shapes collection = slide.Shapes;
                        try
                        {
                            for (int index = 1; index <= collection.Count; index++)
                            {
                                PowerPoint.Shape item = collection[index];
                                if (PowerPointComUtil.ShapeHasTextFrame(item))
                                {
                                    shapes.Add(item);
                                }
                                else
                                {
                                    ComHelper.Release(item);
                                }
                            }
                        }
                        finally
                        {
                            ComHelper.Release(collection);
                        }
                    }

                    JArray applied = new JArray();
                    int changedShapes = 0;
                    JArray runs = PowerPointArgumentReader.GetArray(arguments, "runs");
                    if (runs != null && shapeId < 0)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "使用 runs 时必须同时提供 shapeId。");
                    }
                    int changedRuns = 0;
                    foreach (PowerPoint.Shape shape in shapes)
                    {
                        PowerPoint.TextFrame textFrame = null;
                        PowerPoint.TextRange textRange = null;
                        try
                        {
                            textFrame = shape.TextFrame;
                            textRange = textFrame.TextRange;
                            bool changed = ApplyTextFormat(textRange, arguments, applied);
                            if (changed)
                            {
                                changedShapes++;
                            }
                            if (runs != null && runs.Count > 0)
                            {
                                int runChanges = ApplyTextRuns(textRange, runs, applied);
                                changedRuns += runChanges;
                                if (runChanges > 0)
                                {
                                    changedShapes++;
                                }
                            }
                        }
                        finally
                        {
                            ComHelper.Release(textRange, textFrame);
                        }
                    }

                    if (applied.Count == 0)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "没有提供任何可应用的文字格式属性。");
                    }

                    JObject data = new JObject();
                    data["slideIndex"] = slide.SlideIndex;
                    data["shapeId"] = shapeId >= 0 ? shapeId : -1;
                    data["changedShapeCount"] = changedShapes;
                    data["changedRunCount"] = changedRuns;
                    data["applied"] = applied;
                    return Success(
                        "已对第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                        " 张幻灯片的 " + changedShapes.ToString(CultureInfo.InvariantCulture) +
                        " 个文本框应用文字格式。",
                        "write",
                        changedShapes > 0,
                        SlideRef(presentation, slide),
                        data);
                }
                finally
                {
                    foreach (PowerPoint.Shape shape in shapes)
                    {
                        ComHelper.Release(shape);
                    }
                }
            }
            finally
            {
                ComHelper.Release(slide, presentation);
            }
        }

        private static bool ApplyTextFormat(
            PowerPoint.TextRange textRange,
            JObject arguments,
            JArray applied)
        {
            if (textRange == null)
            {
                return false;
            }
            bool changed = false;
            dynamic font = textRange.Font;
            string fontName = PowerPointArgumentReader.GetTrimmedString(arguments, "fontName");
            if (fontName.Length > 0)
            {
                font.Name = fontName;
                applied.Add("fontName");
                changed = true;
            }
            if (HasArgument(arguments, "fontSize"))
            {
                float fontSize = PowerPointArgumentReader.GetFloat(arguments, 18f, "fontSize");
                if (fontSize <= 0)
                {
                    throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "fontSize 必须大于 0。");
                }
                font.Size = fontSize;
                applied.Add("fontSize");
                changed = true;
            }
            string color = PowerPointArgumentReader.GetTrimmedString(arguments, "color");
            if (color.Length > 0)
            {
                int oleColor;
                if (!PowerPointComUtil.TryParseColor(color, out oleColor))
                {
                    throw new OfficeComException(
                        OfficeErrorCode.INVALID_PARAMS,
                        "color 不是有效颜色：" + color);
                }
                font.Color.RGB = oleColor;
                applied.Add("color");
                changed = true;
            }
            changed = ApplyFontFlag(arguments, "bold", font, applied, changed);
            changed = ApplyFontFlag(arguments, "italic", font, applied, changed);
            changed = ApplyFontFlag(arguments, "underline", font, applied, changed);

            dynamic paragraphFormat = textRange.ParagraphFormat;
            string align = PowerPointArgumentReader.GetTrimmedString(arguments, "align");
            if (align.Length > 0)
            {
                paragraphFormat.Alignment = PowerPointComUtil.MapParagraphAlignment(align);
                applied.Add("align");
                changed = true;
            }
            if (HasArgument(arguments, "lineSpacing"))
            {
                float lineSpacing = PowerPointArgumentReader.GetFloat(arguments, 1f, "lineSpacing");
                if (lineSpacing <= 0)
                {
                    throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "lineSpacing 必须大于 0。");
                }
                paragraphFormat.LineRuleWithin = -1;
                paragraphFormat.SpaceWithin = lineSpacing;
                applied.Add("lineSpacing");
                changed = true;
            }
            if (HasArgument(arguments, "spaceBefore"))
            {
                float spaceBefore = PowerPointArgumentReader.GetFloat(arguments, 0f, "spaceBefore");
                paragraphFormat.SpaceBefore = spaceBefore;
                applied.Add("spaceBefore");
                changed = true;
            }
            if (HasArgument(arguments, "spaceAfter"))
            {
                float spaceAfter = PowerPointArgumentReader.GetFloat(arguments, 0f, "spaceAfter");
                paragraphFormat.SpaceAfter = spaceAfter;
                applied.Add("spaceAfter");
                changed = true;
            }
            return changed;
        }

        private static int ApplyTextRuns(PowerPoint.TextRange textRange, JArray runs, JArray applied)
        {
            if (textRange == null || runs == null || runs.Count == 0)
            {
                return 0;
            }
            int textLength = 0;
            try
            {
                textLength = textRange.Length;
            }
            catch (Exception)
            {
                textLength = 0;
            }
            int changed = 0;
            bool recorded = false;
            for (int index = 0; index < runs.Count; index++)
            {
                JObject run = runs[index] as JObject;
                if (run == null)
                {
                    throw new OfficeComException(
                        OfficeErrorCode.INVALID_PARAMS,
                        "runs[" + index.ToString(CultureInfo.InvariantCulture) + "] 必须是对象。");
                }
                int start = PowerPointArgumentReader.GetInt(run, 1, "start");
                int length = PowerPointArgumentReader.GetInt(run, 0, "length");
                if (start < 1)
                {
                    throw new OfficeComException(
                        OfficeErrorCode.INVALID_PARAMS,
                        "runs[" + index.ToString(CultureInfo.InvariantCulture) + "].start 必须大于等于 1。");
                }
                if (textLength <= 0 || start > textLength)
                {
                    throw new OfficeComException(
                        OfficeErrorCode.INVALID_PARAMS,
                        "runs[" + index.ToString(CultureInfo.InvariantCulture) + "] 的 start 超出文本长度 " +
                        textLength.ToString(CultureInfo.InvariantCulture) + "。");
                }
                if (length <= 0)
                {
                    length = textLength - start + 1;
                }
                else if (length > textLength - start + 1)
                {
                    length = textLength - start + 1;
                }
                PowerPoint.TextRange runRange = null;
                try
                {
                    runRange = textRange.Characters(start, length);
                    JArray runApplied = new JArray();
                    if (ApplyTextFormat(runRange, run, runApplied))
                    {
                        changed++;
                        if (!recorded)
                        {
                            applied.Add("runs");
                            recorded = true;
                        }
                    }
                }
                finally
                {
                    ComHelper.Release(runRange);
                }
            }
            return changed;
        }

        private static bool ApplyFontFlag(
            JObject arguments,
            string name,
            dynamic font,
            JArray applied,
            bool changed)
        {
            if (!HasArgument(arguments, name))
            {
                return changed;
            }
            bool value = PowerPointArgumentReader.GetBool(arguments, false, name);
            if (string.Equals(name, "bold", StringComparison.Ordinal))
            {
                font.Bold = value ? -1 : 0;
            }
            else if (string.Equals(name, "italic", StringComparison.Ordinal))
            {
                font.Italic = value ? -1 : 0;
            }
            else
            {
                font.Underline = value ? -1 : 0;
            }
            applied.Add(name);
            return true;
        }

        private static OperationResult ExecuteSetShapeText(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.Shape shape = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                string content = PowerPointArgumentReader.RequireString(arguments, "content");
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                int shapeId = PowerPointArgumentReader.GetInt(arguments, -1, "shapeId");
                if (shapeId < 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少必填参数 shapeId。");
                }
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                shape = PowerPointComUtil.FindShape(slide, shapeId);
                if (shape == null)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "找不到 shapeId=" + shapeId.ToString(CultureInfo.InvariantCulture) + "。");
                }
                if (!PowerPointComUtil.ShapeHasTextFrame(shape))
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "目标形状没有文本框，无法设置文字。");
                }

                PowerPoint.TextFrame textFrame = null;
                PowerPoint.TextRange textRange = null;
                try
                {
                    textFrame = shape.TextFrame;
                    textRange = textFrame.TextRange;
                    string normalized = PowerPointComUtil.NormalizeText(content);
                    textRange.Text = normalized;
                    string applied = textRange.Text ?? string.Empty;
                    if (!string.Equals(
                            NormalizeForCompare(applied),
                            NormalizeForCompare(normalized),
                            StringComparison.Ordinal))
                    {
                        return Fail(OfficeErrorCode.PARTIAL_APPLY, "文字写入后回读不一致，未确认成功。");
                    }

                    JObject data = new JObject();
                    data["slideIndex"] = slide.SlideIndex;
                    data["shapeId"] = shapeId;
                    data["shapeName"] = shape.Name;
                    data["textLength"] = applied.Length;
                    return Success(
                        "已修改第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                        " 张幻灯片形状 " + shape.Name + " 的文字。",
                        "write",
                        true,
                        SlideRef(presentation, slide) + "!" + shape.Name,
                        data);
                }
                finally
                {
                    ComHelper.Release(textRange, textFrame);
                }
            }
            finally
            {
                ComHelper.Release(shape, slide, presentation);
            }
        }

        private static OperationResult ExecuteDeleteShape(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.Shape shape = null;
            try
            {
                if (!PowerPointArgumentReader.GetBool(arguments, false, "confirm"))
                {
                    return OperationResult.Cancelled(
                        "ppt_delete_shape 需要显式 confirm=true 才会删除；本次未修改演示文稿。");
                }
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                int shapeId = PowerPointArgumentReader.GetInt(arguments, -1, "shapeId");
                if (shapeId < 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少必填参数 shapeId。");
                }
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                shape = PowerPointComUtil.FindShape(slide, shapeId);
                if (shape == null)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "找不到 shapeId=" + shapeId.ToString(CultureInfo.InvariantCulture) + "。");
                }
                string shapeName = shape.Name;
                int targetSlideIndex = slide.SlideIndex;
                shape.Delete();
                shape = null;

                JObject data = new JObject();
                data["slideIndex"] = targetSlideIndex;
                data["shapeId"] = shapeId;
                data["shapeName"] = shapeName;
                data["deleted"] = true;
                return Success(
                    "已删除第 " + targetSlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片上的形状 " + shapeName + "。",
                    "write",
                    true,
                    "ppt:" + presentation.Name + "!Slide" + targetSlideIndex.ToString(CultureInfo.InvariantCulture) +
                    "!" + shapeName,
                    data);
            }
            finally
            {
                ComHelper.Release(shape, slide, presentation);
            }
        }

        private static OperationResult ExecuteMoveResizeShape(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.Shape shape = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                int shapeId = PowerPointArgumentReader.GetInt(arguments, -1, "shapeId");
                if (shapeId < 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少必填参数 shapeId。");
                }
                bool hasX = HasArgument(arguments, "x");
                bool hasY = HasArgument(arguments, "y");
                bool hasWidth = HasArgument(arguments, "width");
                bool hasHeight = HasArgument(arguments, "height");
                if (!hasX && !hasY && !hasWidth && !hasHeight)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "x/y/width/height 至少需要提供一项。");
                }

                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                shape = PowerPointComUtil.FindShape(slide, shapeId);
                if (shape == null)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "找不到 shapeId=" + shapeId.ToString(CultureInfo.InvariantCulture) + "。");
                }

                JArray applied = new JArray();
                if (hasX)
                {
                    shape.Left = PowerPointArgumentReader.GetFloat(arguments, shape.Left, "x");
                    applied.Add("x");
                }
                if (hasY)
                {
                    shape.Top = PowerPointArgumentReader.GetFloat(arguments, shape.Top, "y");
                    applied.Add("y");
                }
                if (hasWidth)
                {
                    float width = PowerPointArgumentReader.GetFloat(arguments, shape.Width, "width");
                    if (width <= 0)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "width 必须大于 0。");
                    }
                    shape.Width = width;
                    applied.Add("width");
                }
                if (hasHeight)
                {
                    float height = PowerPointArgumentReader.GetFloat(arguments, shape.Height, "height");
                    if (height <= 0)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "height 必须大于 0。");
                    }
                    shape.Height = height;
                    applied.Add("height");
                }

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["shapeId"] = shapeId;
                data["shapeName"] = shape.Name;
                data["x"] = shape.Left;
                data["y"] = shape.Top;
                data["width"] = shape.Width;
                data["height"] = shape.Height;
                data["applied"] = applied;
                return Success(
                    "已调整第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片形状 " + shape.Name + " 的位置/尺寸。",
                    "write",
                    true,
                    SlideRef(presentation, slide) + "!" + shape.Name,
                    data);
            }
            finally
            {
                ComHelper.Release(shape, slide, presentation);
            }
        }

        private static OperationResult ExecuteInsertChart(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.Shape shape = null;
            object chartDataObject = null;
            object workbookObject = null;
            object worksheetObject = null;
            object rangeObject = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                string chartType = PowerPointArgumentReader.RequireString(arguments, "chartType");
                JArray data = PowerPointArgumentReader.GetArray(arguments, "data");
                if (data == null || data.Count < 2)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "data 必须至少包含表头行和一行数据。");
                }
                int rows = data.Count;
                int columns = 0;
                for (int index = 0; index < rows; index++)
                {
                    JArray row = data[index] as JArray;
                    if (row == null)
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "data[" + index.ToString(CultureInfo.InvariantCulture) + "] 必须是数组。");
                    }
                    columns = Math.Max(columns, row.Count);
                }
                if (columns < 2)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "data 至少需要两列（类别列和数值列）。");
                }

                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                float x = PowerPointArgumentReader.GetFloat(arguments, 50f, "x");
                float y = PowerPointArgumentReader.GetFloat(arguments, 150f, "y");
                float width = PowerPointArgumentReader.GetFloat(arguments, 600f, "width");
                float height = PowerPointArgumentReader.GetFloat(arguments, 350f, "height");
                if (width <= 0 || height <= 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "width/height 必须大于 0。");
                }

                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                dynamic shapes = slide.Shapes;
                object chartShapeObject = shapes.AddChart2(
                    -1,
                    PowerPointComUtil.MapChartType(chartType),
                    x,
                    y,
                    width,
                    height,
                    0);
                shape = chartShapeObject as PowerPoint.Shape;
                if (shape == null)
                {
                    return Fail(OfficeErrorCode.INTERNAL_ERROR, "PowerPoint 未返回图表形状对象。");
                }
                PowerPoint.Chart chart = shape.Chart;
                PowerPoint.ChartData chartData = chart.ChartData;
                chartDataObject = chartData;
                chartData.Activate();
                workbookObject = TryGetChartWorkbook(chartDataObject, chartData);
                if (workbookObject == null)
                {
                    return Fail(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "无法打开 PowerPoint 图表数据工作簿（Excel 自动化不可用或图表初始化失败）。");
                }
                dynamic workbook = workbookObject;
                worksheetObject = workbook.Worksheets[1];
                dynamic worksheet = worksheetObject;
                dynamic cells = worksheet.Cells;
                cells.Clear();

                for (int rowIndex = 0; rowIndex < rows; rowIndex++)
                {
                    JArray row = data[rowIndex] as JArray;
                    for (int columnIndex = 0; columnIndex < columns; columnIndex++)
                    {
                        dynamic cell = worksheet.Cells[rowIndex + 1, columnIndex + 1];
                        JToken value = row != null && columnIndex < row.Count ? row[columnIndex] : null;
                        cell.Value2 = value == null ? null : value.ToString();
                        ComHelper.Release(cell);
                    }
                }

                rangeObject = worksheet.Range(
                    worksheet.Cells[1, 1],
                    worksheet.Cells[rows, columns]);
                // ChartData 工作簿与图表默认数据源已绑定；写入单元格后图表会随数据源刷新。
                // 不再调用 SetSourceData：部分 Office 版本对隐藏/新建图表的字符串地址会返回 E_FAIL。

                string title = PowerPointArgumentReader.GetTrimmedString(arguments, "title");
                if (title.Length > 0)
                {
                    dynamic dynamicChart = chart;
                    dynamicChart.HasTitle = -1;
                    dynamicChart.ChartTitle.Text = title;
                }

                try
                {
                    workbook.Close(false);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[ppt] 关闭图表数据工作簿失败：" + ex.Message);
                }

                JObject resultData = new JObject();
                resultData["slideIndex"] = slide.SlideIndex;
                resultData["shapeId"] = PowerPointComUtil.SafeInt(delegate
                {
                    dynamic dynamicShape = shape;
                    return Convert.ToInt32((object)dynamicShape.Id);
                });
                resultData["shapeName"] = shape.Name;
                resultData["chartType"] = chartType;
                resultData["rows"] = rows - 1;
                resultData["seriesCount"] = columns - 1;
                return Success(
                    "已在第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片插入图表（" + (rows - 1).ToString(CultureInfo.InvariantCulture) + " 行，" +
                    (columns - 1).ToString(CultureInfo.InvariantCulture) + " 个系列）。",
                    "write",
                    true,
                    SlideRef(presentation, slide) + "!" + shape.Name,
                    resultData);
            }
            catch (Exception ex)
            {
                return OfficeErrorMapper.ToOperationResult(ex, OfficeApp.PowerPoint);
            }
            finally
            {
                ComHelper.Release(rangeObject, worksheetObject, workbookObject, chartDataObject, shape, slide, presentation);
            }
        }

        private static object TryGetChartWorkbook(object chartDataObject, dynamic chartData)
        {
            for (int attempt = 0; attempt < 6; attempt++)
            {
                try
                {
                    return chartData.Workbook;
                }
                catch (Exception ex)
                {
                    if (attempt == 5)
                    {
                        Console.Error.WriteLine("[ppt] 读取图表数据工作簿失败：" + ex.Message);
                        return null;
                    }
                    System.Threading.Thread.Sleep(250);
                }
            }
            return null;
        }

        private static OperationResult ExecuteSetBullet(
            PowerPoint.Application application,
            JObject arguments)
        {
            return ApplyParagraphProperties(application, arguments, true);
        }

        private static OperationResult ExecuteSetParagraph(
            PowerPoint.Application application,
            JObject arguments)
        {
            return ApplyParagraphProperties(application, arguments, false);
        }

        private static OperationResult ApplyParagraphProperties(
            PowerPoint.Application application,
            JObject arguments,
            bool bulletMode)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                int shapeId = PowerPointArgumentReader.GetInt(arguments, -1, "shapeId");
                List<PowerPoint.Shape> shapes = new List<PowerPoint.Shape>();
                try
                {
                    if (shapeId >= 0)
                    {
                        PowerPoint.Shape target = PowerPointComUtil.FindShape(slide, shapeId);
                        if (target == null)
                        {
                            return Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "找不到 shapeId=" + shapeId.ToString(CultureInfo.InvariantCulture) + "。");
                        }
                        shapes.Add(target);
                    }
                    else
                    {
                        PowerPoint.Shapes collection = slide.Shapes;
                        try
                        {
                            for (int index = 1; index <= collection.Count; index++)
                            {
                                PowerPoint.Shape item = collection[index];
                                if (PowerPointComUtil.ShapeHasTextFrame(item))
                                {
                                    shapes.Add(item);
                                }
                                else
                                {
                                    ComHelper.Release(item);
                                }
                            }
                        }
                        finally
                        {
                            ComHelper.Release(collection);
                        }
                    }

                    JArray applied = new JArray();
                    int changedShapes = 0;
                    foreach (PowerPoint.Shape shape in shapes)
                    {
                        PowerPoint.TextFrame textFrame = null;
                        PowerPoint.TextRange textRange = null;
                        try
                        {
                            textFrame = shape.TextFrame;
                            textRange = textFrame.TextRange;
                            bool changed = bulletMode
                                ? ApplyBullet(textRange, arguments, applied)
                                : ApplyParagraph(textRange, arguments, applied);
                            if (changed)
                            {
                                changedShapes++;
                            }
                        }
                        finally
                        {
                            ComHelper.Release(textRange, textFrame);
                        }
                    }

                    if (applied.Count == 0)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "没有提供任何可应用的段落属性。");
                    }

                    JObject data = new JObject();
                    data["slideIndex"] = slide.SlideIndex;
                    data["shapeId"] = shapeId >= 0 ? shapeId : -1;
                    data["changedShapeCount"] = changedShapes;
                    data["applied"] = applied;
                    return Success(
                        "已对第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                        " 张幻灯片的 " + changedShapes.ToString(CultureInfo.InvariantCulture) +
                        " 个文本框应用段落设置。",
                        "write",
                        changedShapes > 0,
                        SlideRef(presentation, slide),
                        data);
                }
                finally
                {
                    foreach (PowerPoint.Shape shape in shapes)
                    {
                        ComHelper.Release(shape);
                    }
                }
            }
            finally
            {
                ComHelper.Release(slide, presentation);
            }
        }

        private static bool ApplyBullet(
            PowerPoint.TextRange textRange,
            JObject arguments,
            JArray applied)
        {
            dynamic paragraphFormat = textRange.ParagraphFormat;
            bool enabled = PowerPointArgumentReader.GetBool(arguments, true, "enabled");
            dynamic bullet = paragraphFormat.Bullet;
            bullet.Visible = enabled ? -1 : 0;
            applied.Add("bulletVisible");

            if (HasArgument(arguments, "level"))
            {
                int level = PowerPointArgumentReader.GetInt(arguments, 1, "level");
                dynamic dynamicRange = textRange;
                dynamicRange.IndentLevel = Math.Max(1, Math.Min(5, level));
                applied.Add("level");
            }
            string bulletChar = PowerPointArgumentReader.GetTrimmedString(arguments, "bulletChar");
            if (bulletChar.Length > 0 && enabled)
            {
                if (string.Equals(bulletChar, "number", StringComparison.OrdinalIgnoreCase))
                {
                    bullet.Type = 2;
                }
                else
                {
                    bullet.Type = 1;
                    if (string.Equals(bulletChar, "dash", StringComparison.OrdinalIgnoreCase))
                    {
                        bullet.Character = 45;
                    }
                    else if (string.Equals(bulletChar, "dot", StringComparison.OrdinalIgnoreCase))
                    {
                        bullet.Character = 8226;
                    }
                    else
                    {
                        bullet.Character = Convert.ToInt32(bulletChar[0]);
                    }
                }
                applied.Add("bulletChar");
            }
            return true;
        }

        private static bool ApplyParagraph(
            PowerPoint.TextRange textRange,
            JObject arguments,
            JArray applied)
        {
            dynamic paragraphFormat = textRange.ParagraphFormat;
            bool changed = false;
            if (HasArgument(arguments, "indentLevel"))
            {
                int level = PowerPointArgumentReader.GetInt(arguments, 1, "indentLevel");
                dynamic dynamicRange = textRange;
                dynamicRange.IndentLevel = Math.Max(1, Math.Min(5, level));
                applied.Add("indentLevel");
                changed = true;
            }
            if (HasArgument(arguments, "spaceBefore"))
            {
                paragraphFormat.SpaceBefore = PowerPointArgumentReader.GetFloat(arguments, 0f, "spaceBefore");
                applied.Add("spaceBefore");
                changed = true;
            }
            if (HasArgument(arguments, "spaceAfter"))
            {
                paragraphFormat.SpaceAfter = PowerPointArgumentReader.GetFloat(arguments, 0f, "spaceAfter");
                applied.Add("spaceAfter");
                changed = true;
            }
            if (HasArgument(arguments, "lineSpacing"))
            {
                float lineSpacing = PowerPointArgumentReader.GetFloat(arguments, 1f, "lineSpacing");
                if (lineSpacing <= 0)
                {
                    throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "lineSpacing 必须大于 0。");
                }
                paragraphFormat.LineRuleWithin = -1;
                paragraphFormat.SpaceWithin = lineSpacing;
                applied.Add("lineSpacing");
                changed = true;
            }
            string align = PowerPointArgumentReader.GetTrimmedString(arguments, "align");
            if (align.Length > 0)
            {
                paragraphFormat.Alignment = PowerPointComUtil.MapParagraphAlignment(align);
                applied.Add("align");
                changed = true;
            }
            return changed;
        }

        private static bool HasArgument(JObject arguments, string name)
        {
            if (arguments == null)
            {
                return false;
            }
            JToken token = arguments[name];
            return token != null && token.Type != JTokenType.Null;
        }

        private static bool IsRemoteUrl(string source)
        {
            Uri uri;
            return Uri.TryCreate(source, UriKind.Absolute, out uri) &&
                   (string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase));
        }

        private static bool TryDownloadImage(
            string url,
            out string localPath,
            out string downloadedFile,
            out string error)
        {
            localPath = string.Empty;
            downloadedFile = string.Empty;
            error = string.Empty;
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(30);
                    byte[] bytes = client.GetByteArrayAsync(url).GetAwaiter().GetResult();
                    if (bytes == null || bytes.Length == 0)
                    {
                        error = "远程图片为空：" + url;
                        return false;
                    }
                    if (bytes.Length > MaxRemoteImageBytes)
                    {
                        error = "远程图片超过 50 MB 限制：" + url;
                        return false;
                    }
                    string extension = GetRemoteImageExtension(url);
                    string temp = Path.Combine(
                        Path.GetTempPath(),
                        "AutoOfficeMCP-ppt-image-" + Guid.NewGuid().ToString("N") + extension);
                    File.WriteAllBytes(temp, bytes);
                    localPath = temp;
                    downloadedFile = temp;
                    return true;
                }
            }
            catch (Exception ex)
            {
                error = "下载远程图片失败：" + ex.Message;
                return false;
            }
        }

        private static string GetRemoteImageExtension(string url)
        {
            try
            {
                string extension = Path.GetExtension(new Uri(url).AbsolutePath);
                if (ImageExtensions.Contains(extension))
                {
                    return extension.ToLowerInvariant();
                }
            }
            catch (Exception)
            {
            }
            return ".png";
        }
    }
}
