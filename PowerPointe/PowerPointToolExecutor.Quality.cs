using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using Newtonsoft.Json.Linq;
using Office;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointe
{
    public sealed partial class PowerPointToolExecutor
    {
        private sealed class LayoutBox
        {
            public int SlideIndex;
            public int ShapeId;
            public string ShapeName;
            public float Left;
            public float Top;
            public float Width;
            public float Height;
        }

        private static OperationResult ExecuteValidateLayout(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int requestedSlideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                bool allSlides = requestedSlideIndex < 0;
                HashSet<string> checks = ReadLayoutChecks(arguments);
                float minImageDpi = PowerPointArgumentReader.GetFloat(arguments, 96f, "minImageDpi");
                if (minImageDpi <= 0f)
                {
                    minImageDpi = 96f;
                }

                float slideWidth = PowerPointComUtil.SafeFloat(delegate
                {
                    return presentation.PageSetup.SlideWidth;
                });
                float slideHeight = PowerPointComUtil.SafeFloat(delegate
                {
                    return presentation.PageSetup.SlideHeight;
                });
                HashSet<string> installedFonts = ReadInstalledFonts();
                JArray issues = new JArray();
                JArray checkedSlides = new JArray();
                int checkedShapeCount = 0;

                PowerPoint.Slides slides = null;
                try
                {
                    slides = presentation.Slides;
                    if (allSlides)
                    {
                        for (int index = 1; index <= slides.Count; index++)
                        {
                            PowerPoint.Slide slide = null;
                            try
                            {
                                slide = slides[index];
                                ValidateSlideLayout(
                                    slide, checks, slideWidth, slideHeight, minImageDpi,
                                    installedFonts, issues, ref checkedShapeCount);
                                checkedSlides.Add(slide.SlideIndex);
                            }
                            finally
                            {
                                ComHelper.Release(slide);
                            }
                        }
                    }
                    else
                    {
                        PowerPoint.Slide slide = PowerPointComUtil.ResolveSlide(
                            application, presentation, requestedSlideIndex);
                        try
                        {
                            ValidateSlideLayout(
                                slide, checks, slideWidth, slideHeight, minImageDpi,
                                installedFonts, issues, ref checkedShapeCount);
                            checkedSlides.Add(slide.SlideIndex);
                        }
                        finally
                        {
                            ComHelper.Release(slide);
                        }
                    }
                }
                finally
                {
                    ComHelper.Release(slides);
                }

                JObject data = new JObject();
                data["checkedSlideCount"] = checkedSlides.Count;
                data["checkedSlideIndices"] = checkedSlides;
                data["checkedShapeCount"] = checkedShapeCount;
                data["checks"] = new JArray(checks);
                data["issueCount"] = issues.Count;
                data["issues"] = issues;
                return Success(
                    "布局自检完成：检查 " + checkedSlides.Count.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片、共 " + checkedShapeCount.ToString(CultureInfo.InvariantCulture) +
                    " 个形状，发现 " + issues.Count.ToString(CultureInfo.InvariantCulture) + " 个问题。",
                    "read",
                    false,
                    "ppt:" + presentation.Name,
                    data);
            }
            finally
            {
                ComHelper.Release(presentation);
            }
        }

        private static HashSet<string> ReadLayoutChecks(JObject arguments)
        {
            string[] defaultChecks = new string[]
            {
                "textOverflow", "overlap", "outOfBounds", "missingFont", "lowResolution"
            };
            HashSet<string> checks = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            JArray requested = PowerPointArgumentReader.GetArray(arguments, "checks");
            if (requested == null || requested.Count == 0)
            {
                foreach (string check in defaultChecks)
                {
                    checks.Add(check);
                }
                return checks;
            }
            foreach (JToken token in requested)
            {
                string value = token == null ? string.Empty : token.ToString().Trim();
                if (value.Length == 0)
                {
                    continue;
                }
                bool known = false;
                foreach (string check in defaultChecks)
                {
                    if (string.Equals(check, value, StringComparison.OrdinalIgnoreCase))
                    {
                        checks.Add(check);
                        known = true;
                        break;
                    }
                }
                if (!known)
                {
                    throw new OfficeComException(
                        OfficeErrorCode.INVALID_PARAMS,
                        "checks 含未知检查项「" + value + "」；支持 textOverflow/overlap/outOfBounds/missingFont/lowResolution。");
                }
            }
            if (checks.Count == 0)
            {
                throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "checks 不能为空。");
            }
            return checks;
        }

        private static HashSet<string> ReadInstalledFonts()
        {
            HashSet<string> fonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (InstalledFontCollection collection = new InstalledFontCollection())
                {
                    foreach (FontFamily family in collection.Families)
                    {
                        if (!string.IsNullOrWhiteSpace(family.Name))
                        {
                            fonts.Add(family.Name);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[ppt] 读取系统字体失败：" + ex.Message);
            }
            return fonts;
        }

        private static void ValidateSlideLayout(
            PowerPoint.Slide slide,
            HashSet<string> checks,
            float slideWidth,
            float slideHeight,
            float minImageDpi,
            HashSet<string> installedFonts,
            JArray issues,
            ref int checkedShapeCount)
        {
            List<LayoutBox> boxes = new List<LayoutBox>();
            PowerPoint.Shapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                int count = shapes.Count;
                for (int index = 1; index <= count; index++)
                {
                    PowerPoint.Shape shape = null;
                    try
                    {
                        shape = shapes[index];
                        dynamic dynamicShape = shape;
                        bool visible = true;
                        try
                        {
                            visible = Convert.ToInt32((object)dynamicShape.Visible) != 0;
                        }
                        catch (Exception)
                        {
                        }
                        if (!visible)
                        {
                            continue;
                        }

                        LayoutBox box = new LayoutBox();
                        box.SlideIndex = slide.SlideIndex;
                        box.ShapeId = PowerPointComUtil.SafeInt(delegate
                        {
                            return Convert.ToInt32((object)dynamicShape.Id);
                        });
                        box.ShapeName = PowerPointComUtil.SafeString(delegate { return shape.Name; });
                        box.Left = PowerPointComUtil.SafeFloat(delegate { return shape.Left; });
                        box.Top = PowerPointComUtil.SafeFloat(delegate { return shape.Top; });
                        box.Width = PowerPointComUtil.SafeFloat(delegate { return shape.Width; });
                        box.Height = PowerPointComUtil.SafeFloat(delegate { return shape.Height; });
                        boxes.Add(box);
                        checkedShapeCount++;

                        if (checks.Contains("outOfBounds") &&
                            (box.Left < -1f || box.Top < -1f ||
                             box.Left + box.Width > slideWidth + 1f ||
                             box.Top + box.Height > slideHeight + 1f))
                        {
                            AddLayoutIssue(
                                issues, box, "outOfBounds", "error",
                                "形状超出页面边界（left=" + box.Left.ToString("0.##", CultureInfo.InvariantCulture) +
                                ", top=" + box.Top.ToString("0.##", CultureInfo.InvariantCulture) +
                                ", width=" + box.Width.ToString("0.##", CultureInfo.InvariantCulture) +
                                ", height=" + box.Height.ToString("0.##", CultureInfo.InvariantCulture) +
                                "；页面 " + slideWidth.ToString("0.##", CultureInfo.InvariantCulture) + "×" +
                                slideHeight.ToString("0.##", CultureInfo.InvariantCulture) + "）。");
                        }

                        if (checks.Contains("textOverflow") && PowerPointComUtil.ShapeHasTextFrame(shape))
                        {
                            CheckTextOverflow(box, shape, issues);
                        }
                        if (checks.Contains("missingFont"))
                        {
                            CheckMissingFont(box, shape, installedFonts, issues);
                        }
                        if (checks.Contains("lowResolution") &&
                            (PowerPointComUtil.ShapeType(shape) == 13 || ShapeHasPictureFill(shape)))
                        {
                            CheckImageResolution(box, shape, minImageDpi, issues);
                        }
                    }
                    finally
                    {
                        ComHelper.Release(shape);
                    }
                }
            }
            finally
            {
                ComHelper.Release(shapes);
            }

            if (checks.Contains("overlap"))
            {
                for (int leftIndex = 0; leftIndex < boxes.Count; leftIndex++)
                {
                    for (int rightIndex = leftIndex + 1; rightIndex < boxes.Count; rightIndex++)
                    {
                        LayoutBox left = boxes[leftIndex];
                        LayoutBox right = boxes[rightIndex];
                        float intersectionWidth = Math.Min(left.Left + left.Width, right.Left + right.Width) -
                                                  Math.Max(left.Left, right.Left);
                        float intersectionHeight = Math.Min(left.Top + left.Height, right.Top + right.Height) -
                                                   Math.Max(left.Top, right.Top);
                        if (intersectionWidth <= 4f || intersectionHeight <= 4f)
                        {
                            continue;
                        }
                        float intersectionArea = intersectionWidth * intersectionHeight;
                        float minArea = Math.Min(left.Width * left.Height, right.Width * right.Height);
                        if (minArea <= 0f || intersectionArea < minArea * 0.05f)
                        {
                            continue;
                        }
                        AddLayoutIssue(
                            issues, left, "overlap", "warning",
                            "形状与「" + right.ShapeName + "」(id=" +
                            right.ShapeId.ToString(CultureInfo.InvariantCulture) + ") 重叠约 " +
                            (intersectionArea / minArea * 100f).ToString("0.#", CultureInfo.InvariantCulture) + "%。");
                    }
                }
            }
        }

        private static void CheckTextOverflow(
            LayoutBox box,
            PowerPoint.Shape shape,
            JArray issues)
        {
            PowerPoint.TextFrame textFrame = null;
            PowerPoint.TextRange textRange = null;
            try
            {
                textFrame = shape.TextFrame;
                dynamic dynamicFrame = textFrame;
                try
                {
                    // 1 = ppAutoSizeShapeToFitText：文本框会随文字长高，本身不会溢出。
                    // 0 = ppAutoSizeNone；-2 = ppAutoSizeMixed（占位符常见），仍按固定框检查。
                    if (Convert.ToInt32((object)dynamicFrame.AutoSize) == 1)
                    {
                        return;
                    }
                }
                catch (Exception)
                {
                }
                textRange = textFrame.TextRange;
                string text = textRange.Text ?? string.Empty;
                if (text.Trim().Length == 0)
                {
                    return;
                }
                float boundHeight = PowerPointComUtil.SafeFloat(delegate
                {
                    dynamic dynamicRange = textRange;
                    return Convert.ToSingle((object)dynamicRange.BoundHeight);
                });
                float boundWidth = PowerPointComUtil.SafeFloat(delegate
                {
                    dynamic dynamicRange = textRange;
                    return Convert.ToSingle((object)dynamicRange.BoundWidth);
                });
                float marginTop = PowerPointComUtil.SafeFloat(delegate
                {
                    return Convert.ToSingle((object)dynamicFrame.MarginTop);
                });
                float marginBottom = PowerPointComUtil.SafeFloat(delegate
                {
                    return Convert.ToSingle((object)dynamicFrame.MarginBottom);
                });
                float marginLeft = PowerPointComUtil.SafeFloat(delegate
                {
                    return Convert.ToSingle((object)dynamicFrame.MarginLeft);
                });
                float marginRight = PowerPointComUtil.SafeFloat(delegate
                {
                    return Convert.ToSingle((object)dynamicFrame.MarginRight);
                });
                bool wordWrap = PowerPointComUtil.SafeBool(delegate
                {
                    return Convert.ToInt32((object)dynamicFrame.WordWrap) != 0;
                });
                float availableHeight = box.Height - marginTop - marginBottom;
                float availableWidth = box.Width - marginLeft - marginRight;
                bool heightOverflow = boundHeight > 0f && availableHeight > 0f && boundHeight > availableHeight + 4f;
                bool widthOverflow = !wordWrap && boundWidth > 0f && availableWidth > 0f &&
                                     boundWidth > availableWidth + 4f;
                if (heightOverflow || widthOverflow)
                {
                    AddLayoutIssue(
                        issues, box, "textOverflow", "warning",
                        "文字可能溢出文本框（文本 " + boundWidth.ToString("0.##", CultureInfo.InvariantCulture) +
                        "×" + boundHeight.ToString("0.##", CultureInfo.InvariantCulture) +
                        "，可用 " + availableWidth.ToString("0.##", CultureInfo.InvariantCulture) +
                        "×" + availableHeight.ToString("0.##", CultureInfo.InvariantCulture) + "）。");
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                ComHelper.Release(textRange, textFrame);
            }
        }

        private static void CheckMissingFont(
            LayoutBox box,
            PowerPoint.Shape shape,
            HashSet<string> installedFonts,
            JArray issues)
        {
            if (installedFonts.Count == 0 || !PowerPointComUtil.ShapeHasTextFrame(shape))
            {
                return;
            }
            PowerPoint.TextFrame textFrame = null;
            PowerPoint.TextRange textRange = null;
            try
            {
                textFrame = shape.TextFrame;
                textRange = textFrame.TextRange;
                if ((textRange.Text ?? string.Empty).Trim().Length == 0)
                {
                    return;
                }
                dynamic font = textRange.Font;
                string fontName = PowerPointComUtil.SafeString(delegate
                {
                    return Convert.ToString((object)font.Name);
                }).Trim();
                if (fontName.Length == 0 || fontName.StartsWith("+", StringComparison.Ordinal))
                {
                    return;
                }
                if (!installedFonts.Contains(fontName))
                {
                    AddLayoutIssue(
                        issues, box, "missingFont", "warning",
                        "字体「" + fontName + "」未安装，目标端可能回退为其他字体。");
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                ComHelper.Release(textRange, textFrame);
            }
        }

        private static bool ShapeHasPictureFill(PowerPoint.Shape shape)
        {
            object fill = null;
            try
            {
                dynamic dynamicShape = shape;
                fill = dynamicShape.Fill;
                dynamic dynamicFill = fill;
                return Convert.ToInt32((object)dynamicFill.Type) == 6; // msoFillPicture
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                ComHelper.Release(fill);
            }
        }

        private static void CheckImageResolution(
            LayoutBox box,
            PowerPoint.Shape shape,
            float minImageDpi,
            JArray issues)
        {
            if (box.Width <= 0f || box.Height <= 0f)
            {
                return;
            }
            string tempFile = Path.Combine(
                Path.GetTempPath(),
                "AutoOfficeMCP-ppt-quality-" + Guid.NewGuid().ToString("N") + ".png");
            try
            {
                dynamic dynamicShape = shape;
                dynamicShape.Export(tempFile, 2);
                if (!File.Exists(tempFile))
                {
                    return;
                }
                using (Image image = Image.FromFile(tempFile))
                {
                    float dpiX = image.Width / (box.Width / 72f);
                    float dpiY = image.Height / (box.Height / 72f);
                    float dpi = Math.Min(dpiX, dpiY);
                    if (dpi < minImageDpi)
                    {
                        AddLayoutIssue(
                            issues, box, "lowResolution", "warning",
                            "图片有效分辨率约 " + dpi.ToString("0", CultureInfo.InvariantCulture) +
                            " DPI（阈值 " + minImageDpi.ToString("0", CultureInfo.InvariantCulture) +
                            "），放大后可能模糊。");
                    }
                }
            }
            catch (Exception)
            {
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

        private static void AddLayoutIssue(
            JArray issues,
            LayoutBox box,
            string issueType,
            string severity,
            string detail)
        {
            JObject issue = new JObject();
            issue["slideIndex"] = box.SlideIndex;
            issue["shapeId"] = box.ShapeId;
            issue["shapeName"] = box.ShapeName;
            issue["issueType"] = issueType;
            issue["severity"] = severity;
            issue["detail"] = detail;
            issues.Add(issue);
        }
    }
}
