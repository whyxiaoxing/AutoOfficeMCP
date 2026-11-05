using System;
using System.Globalization;
using System.IO;
using Newtonsoft.Json.Linq;
using Office;
using Office.Config;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointe
{
    public sealed partial class PowerPointToolExecutor
    {
        private static OperationResult ExecuteCropImage(
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
                        "找不到 shapeId=" + shapeId.ToString(CultureInfo.InvariantCulture) + "。");
                }
                if (PowerPointComUtil.ShapeType(shape) != 13)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "目标形状不是图片，无法裁切。");
                }

                float leftValue = PowerPointArgumentReader.GetFloat(arguments, 0f, "cropLeft");
                float topValue = PowerPointArgumentReader.GetFloat(arguments, 0f, "cropTop");
                float rightValue = PowerPointArgumentReader.GetFloat(arguments, 0f, "cropRight");
                float bottomValue = PowerPointArgumentReader.GetFloat(arguments, 0f, "cropBottom");
                if (leftValue < 0f || topValue < 0f || rightValue < 0f || bottomValue < 0f)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "裁切值不能为负数。");
                }
                string unit = PowerPointArgumentReader.GetTrimmedString(arguments, "unit", "cropUnit");
                if (unit.Length == 0)
                {
                    unit = "ratio";
                }
                if (unit != "ratio" && unit != "points")
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "unit 只能是 ratio 或 points。");
                }

                float currentWidth = PowerPointComUtil.SafeFloat(delegate { return shape.Width; });
                float currentHeight = PowerPointComUtil.SafeFloat(delegate { return shape.Height; });
                if (currentWidth <= 0f || currentHeight <= 0f)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "无法读取图片当前尺寸。");
                }
                float cropLeft = leftValue;
                float cropTop = topValue;
                float cropRight = rightValue;
                float cropBottom = bottomValue;
                if (string.Equals(unit, "ratio", StringComparison.OrdinalIgnoreCase))
                {
                    if (leftValue + rightValue >= 1f)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "cropLeft + cropRight 必须小于 1。");
                    }
                    if (topValue + bottomValue >= 1f)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "cropTop + cropBottom 必须小于 1。");
                    }
                    cropLeft = leftValue * currentWidth;
                    cropTop = topValue * currentHeight;
                    cropRight = rightValue * currentWidth;
                    cropBottom = bottomValue * currentHeight;
                }
                else if (cropLeft + cropRight >= currentWidth || cropTop + cropBottom >= currentHeight)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "裁切总量不能大于图片当前尺寸。");
                }

                dynamic picture = shape.PictureFormat;
                if (cropLeft > 0f)
                {
                    picture.CropLeft = cropLeft;
                }
                if (cropTop > 0f)
                {
                    picture.CropTop = cropTop;
                }
                if (cropRight > 0f)
                {
                    picture.CropRight = cropRight;
                }
                if (cropBottom > 0f)
                {
                    picture.CropBottom = cropBottom;
                }

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["shapeId"] = shapeId;
                data["shapeName"] = shape.Name;
                data["unit"] = unit;
                data["cropLeft"] = cropLeft;
                data["cropTop"] = cropTop;
                data["cropRight"] = cropRight;
                data["cropBottom"] = cropBottom;
                data["width"] = PowerPointComUtil.SafeFloat(delegate { return shape.Width; });
                data["height"] = PowerPointComUtil.SafeFloat(delegate { return shape.Height; });
                return Success(
                    "已裁切第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片上的图片 " + shape.Name + "。",
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

        private static OperationResult ExecuteReplaceImage(
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
                string source = PowerPointArgumentReader.RequireString(arguments, "source");
                string localPath;
                string sourceError;
                if (!WorkspacePathResolver.TryResolve(source, out localPath, out sourceError))
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, sourceError);
                }
                if (!File.Exists(localPath))
                {
                    return Fail(OfficeErrorCode.FILE_NOT_FOUND, "图片文件不存在：" + localPath);
                }
                if (!ImageExtensions.Contains(Path.GetExtension(localPath)))
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "不支持的图片扩展名：" + Path.GetExtension(localPath));
                }

                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                shape = PowerPointComUtil.FindShape(slide, shapeId);
                if (shape == null)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "找不到 shapeId=" + shapeId.ToString(CultureInfo.InvariantCulture) + "。");
                }
                if (PowerPointComUtil.ShapeType(shape) != 13)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "目标形状不是图片，无法替换。");
                }
                float left = PowerPointComUtil.SafeFloat(delegate { return shape.Left; });
                float top = PowerPointComUtil.SafeFloat(delegate { return shape.Top; });
                float width = PowerPointComUtil.SafeFloat(delegate { return shape.Width; });
                float height = PowerPointComUtil.SafeFloat(delegate { return shape.Height; });

                dynamic fill = shape.Fill;
                fill.UserPicture(localPath);
                shape.Left = left;
                shape.Top = top;
                shape.Width = width;
                shape.Height = height;

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["shapeId"] = shapeId;
                data["shapeName"] = shape.Name;
                data["source"] = localPath;
                data["left"] = left;
                data["top"] = top;
                data["width"] = width;
                data["height"] = height;
                return Success(
                    "已替换第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片上的图片 " + shape.Name + "。",
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

        private static OperationResult ExecuteSetImageFormat(
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
                        "找不到 shapeId=" + shapeId.ToString(CultureInfo.InvariantCulture) + "。");
                }
                if (PowerPointComUtil.ShapeType(shape) != 13)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "目标形状不是图片，无法设置图片格式。");
                }

                JArray applied = new JArray();
                JArray warnings = new JArray();
                string borderColor = PowerPointArgumentReader.GetTrimmedString(arguments, "borderColor");
                float borderWidth = PowerPointArgumentReader.GetFloat(arguments, 1f, "borderWidth");
                if (borderColor.Length > 0)
                {
                    int oleColor;
                    if (!PowerPointComUtil.TryParseColor(borderColor, out oleColor))
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "borderColor 不是有效颜色：" + borderColor);
                    }
                    if (borderWidth <= 0f)
                    {
                        borderWidth = 1f;
                    }
                    dynamic line = shape.Line;
                    line.Visible = -1;
                    line.ForeColor.RGB = oleColor;
                    line.Weight = borderWidth;
                    applied.Add("border");
                }

                JObject shadow = PowerPointArgumentReader.GetObject(arguments, "shadow");
                if (shadow != null)
                {
                    dynamic dynamicShadow = shape.Shadow;
                    dynamicShadow.Visible = -1;
                    ApplyShadowValue(dynamicShadow, shadow, "blur", "Blur");
                    ApplyShadowValue(dynamicShadow, shadow, "transparency", "Transparency");
                    ApplyShadowValue(dynamicShadow, shadow, "offsetX", "OffsetX");
                    ApplyShadowValue(dynamicShadow, shadow, "offsetY", "OffsetY");
                    string shadowColor = PowerPointArgumentReader.GetTrimmedString(shadow, "color");
                    int shadowOleColor;
                    if (shadowColor.Length > 0 && PowerPointComUtil.TryParseColor(shadowColor, out shadowOleColor))
                    {
                        dynamicShadow.ForeColor.RGB = shadowOleColor;
                    }
                    applied.Add("shadow");
                }

                float cornerRadius = PowerPointArgumentReader.GetFloat(arguments, 0f, "cornerRadius");
                if (cornerRadius > 0f)
                {
                    PowerPoint.Shape converted = ConvertPictureToRounded(
                        slide, shape, cornerRadius, borderColor, borderWidth, shadow, warnings);
                    if (converted == null)
                    {
                        warnings.Add("圆角转换失败，已保留原图片。");
                    }
                    else
                    {
                        applied.Add("cornerRadius");
                        JObject data = BuildImageFormatResult(slide, shapeId, applied, warnings);
                        data["converted"] = true;
                        data["newShapeId"] = PowerPointComUtil.SafeInt(delegate
                        {
                            dynamic dynamicShape = converted;
                            return Convert.ToInt32((object)dynamicShape.Id);
                        });
                        data["newShapeName"] = converted.Name;
                        ComHelper.Release(converted);
                        return Success(
                            "已把第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                            " 张幻灯片的图片转换为圆角图片。",
                            "write",
                            true,
                            SlideRef(presentation, slide),
                            data);
                    }
                }

                if (applied.Count == 0 && warnings.Count == 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "没有提供任何可应用的图片格式属性。");
                }
                JObject result = BuildImageFormatResult(slide, shapeId, applied, warnings);
                result["converted"] = false;
                return Success(
                    "已设置第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片图片 " + shape.Name + " 的格式。",
                    "write",
                    applied.Count > 0,
                    SlideRef(presentation, slide) + "!" + shape.Name,
                    result);
            }
            finally
            {
                ComHelper.Release(shape, slide, presentation);
            }
        }

        private static void ApplyShadowValue(
            dynamic shadow,
            JObject arguments,
            string argumentName,
            string propertyName)
        {
            try
            {
                if (!HasArgument(arguments, argumentName))
                {
                    return;
                }
                float value = PowerPointArgumentReader.GetFloat(arguments, 0f, argumentName);
                if (string.Equals(propertyName, "Blur", StringComparison.Ordinal))
                {
                    shadow.Blur = value;
                }
                else if (string.Equals(propertyName, "Transparency", StringComparison.Ordinal))
                {
                    shadow.Transparency = value;
                }
                else if (string.Equals(propertyName, "OffsetX", StringComparison.Ordinal))
                {
                    shadow.OffsetX = value;
                }
                else
                {
                    shadow.OffsetY = value;
                }
            }
            catch (Exception)
            {
            }
        }

        private static PowerPoint.Shape ConvertPictureToRounded(
            PowerPoint.Slide slide,
            PowerPoint.Shape source,
            float cornerRadius,
            string borderColor,
            float borderWidth,
            JObject shadow,
            JArray warnings)
        {
            string tempFile = Path.Combine(
                Path.GetTempPath(),
                "AutoOfficeMCP-ppt-rounded-" + Guid.NewGuid().ToString("N") + ".png");
            PowerPoint.Shape rounded = null;
            try
            {
                dynamic dynamicSource = source;
                dynamicSource.Export(tempFile, 2);
                if (!File.Exists(tempFile))
                {
                    return null;
                }
                float left = PowerPointComUtil.SafeFloat(delegate { return source.Left; });
                float top = PowerPointComUtil.SafeFloat(delegate { return source.Top; });
                float width = PowerPointComUtil.SafeFloat(delegate { return source.Width; });
                float height = PowerPointComUtil.SafeFloat(delegate { return source.Height; });
                rounded = PowerPointComUtil.AddShape(slide, 5, left, top, width, height);
                if (rounded == null)
                {
                    return null;
                }
                dynamic fill = rounded.Fill;
                fill.UserPicture(tempFile);
                try
                {
                    dynamic adjustments = rounded.Adjustments;
                    float ratio = cornerRadius / Math.Min(width, height);
                    adjustments[1] = Math.Max(0f, Math.Min(0.5f, ratio));
                }
                catch (Exception ex)
                {
                    warnings.Add("圆角半径设置失败：" + ex.Message);
                }
                if (borderColor.Length > 0)
                {
                    int oleColor;
                    if (PowerPointComUtil.TryParseColor(borderColor, out oleColor))
                    {
                        dynamic line = rounded.Line;
                        line.Visible = -1;
                        line.ForeColor.RGB = oleColor;
                        line.Weight = borderWidth > 0f ? borderWidth : 1f;
                    }
                }
                if (shadow != null)
                {
                    dynamic dynamicShadow = rounded.Shadow;
                    dynamicShadow.Visible = -1;
                    ApplyShadowValue(dynamicShadow, shadow, "blur", "Blur");
                    ApplyShadowValue(dynamicShadow, shadow, "transparency", "Transparency");
                    ApplyShadowValue(dynamicShadow, shadow, "offsetX", "OffsetX");
                    ApplyShadowValue(dynamicShadow, shadow, "offsetY", "OffsetY");
                }
                source.Delete();
                PowerPoint.Shape result = rounded;
                rounded = null;
                return result;
            }
            catch (Exception ex)
            {
                warnings.Add("圆角转换失败：" + ex.Message);
                ComHelper.Release(rounded);
                return null;
            }
            finally
            {
                try
                {
                    if (File.Exists(tempFile))
                    {
                        File.Delete(tempFile);
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        private static JObject BuildImageFormatResult(
            PowerPoint.Slide slide,
            int shapeId,
            JArray applied,
            JArray warnings)
        {
            JObject data = new JObject();
            data["slideIndex"] = slide.SlideIndex;
            data["shapeId"] = shapeId;
            data["applied"] = applied;
            if (warnings != null && warnings.Count > 0)
            {
                data["warnings"] = warnings;
            }
            return data;
        }
    }
}
