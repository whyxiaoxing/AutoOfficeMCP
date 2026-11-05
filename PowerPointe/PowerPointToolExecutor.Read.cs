using System;
using System.Collections.Generic;
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
        private static OperationResult ExecuteGetPresentationInfo(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide activeSlide = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideCount = presentation.Slides.Count;
                int activeSlideIndex = -1;
                try
                {
                    PowerPoint.DocumentWindow window = application.ActiveWindow;
                    if (window != null)
                    {
                        PowerPoint.View view = window.View;
                        if (view != null)
                        {
                            activeSlide = view.Slide as PowerPoint.Slide;
                            if (activeSlide != null)
                            {
                                activeSlideIndex = activeSlide.SlideIndex;
                            }
                        }
                        ComHelper.Release(view, window);
                    }
                }
                catch (Exception)
                {
                    activeSlideIndex = -1;
                }

                float slideWidth = PowerPointComUtil.SafeFloat(delegate
                {
                    return presentation.PageSetup.SlideWidth;
                });
                float slideHeight = PowerPointComUtil.SafeFloat(delegate
                {
                    return presentation.PageSetup.SlideHeight;
                });
                string filePath = PowerPointComUtil.SafeString(delegate { return presentation.FullName; });
                if (!IsAbsolutePath(filePath))
                {
                    filePath = string.Empty;
                }

                JObject data = new JObject();
                data["filePath"] = filePath;
                data["name"] = PowerPointComUtil.SafeString(delegate { return presentation.Name; });
                data["slideCount"] = slideCount;
                data["activeSlideIndex"] = activeSlideIndex;
                data["slideWidth"] = slideWidth;
                data["slideHeight"] = slideHeight;
                data["aspectRatio"] = DescribeAspectRatio(slideWidth, slideHeight);
                data["saved"] = PowerPointComUtil.SafeBool(delegate
                {
                    dynamic dynamicPresentation = presentation;
                    return Convert.ToInt32((object)dynamicPresentation.Saved) != 0;
                });
                data["readOnly"] = PowerPointComUtil.SafeBool(delegate
                {
                    dynamic dynamicPresentation = presentation;
                    return Convert.ToInt32((object)dynamicPresentation.ReadOnly) != 0;
                });
                data["hasUnsavedChanges"] = !(data.Value<bool?>("saved") ?? false);
                return Success(
                    "已读取演示文稿信息：" + data.Value<string>("name") + "，共 " +
                    slideCount.ToString(CultureInfo.InvariantCulture) + " 张幻灯片。",
                    "read",
                    false,
                    "ppt:" + data.Value<string>("name"),
                    data);
            }
            finally
            {
                ComHelper.Release(activeSlide, presentation);
            }
        }

        private static OperationResult ExecuteGetSlideContent(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                bool includeShapes = PowerPointArgumentReader.GetBool(arguments, true, "includeShapes");
                bool includeFormat = PowerPointArgumentReader.GetBool(arguments, false, "includeFormat");
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);

                string title = string.Empty;
                string content = string.Empty;
                ReadSlideTextPlaceholders(slide, out title, out content);

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["layout"] = PowerPointComUtil.GetSlideLayoutName(slide);
                data["title"] = title;
                data["content"] = content;
                data["notes"] = PowerPointComUtil.GetNotesText(slide);
                if (includeShapes)
                {
                    data["shapes"] = ReadShapes(slide, includeFormat);
                    data["shapesIncludeFormat"] = includeFormat;
                }
                return Success(
                    "已读取第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片内容。",
                    "read",
                    false,
                    SlideRef(presentation, slide),
                    data);
            }
            finally
            {
                ComHelper.Release(slide, presentation);
            }
        }

        private static OperationResult ExecuteListShapes(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                bool includeFormat = PowerPointArgumentReader.GetBool(arguments, false, "includeFormat");
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                JArray shapes = ReadShapes(slide, includeFormat);
                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["shapeCount"] = shapes.Count;
                data["includeFormat"] = includeFormat;
                data["shapes"] = shapes;
                return Success(
                    "第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片共有 " + shapes.Count.ToString(CultureInfo.InvariantCulture) + " 个形状。",
                    "read",
                    false,
                    SlideRef(presentation, slide),
                    data);
            }
            finally
            {
                ComHelper.Release(slide, presentation);
            }
        }

        private static OperationResult ExecuteListLayouts(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                JArray layouts = new JArray();
                PowerPoint.Designs designs = null;
                try
                {
                    designs = presentation.Designs;
                    int designCount = designs.Count;
                    for (int designIndex = 1; designIndex <= designCount; designIndex++)
                    {
                        PowerPoint.Design design = null;
                        PowerPoint.Master master = null;
                        PowerPoint.CustomLayouts customLayouts = null;
                        try
                        {
                            design = designs[designIndex];
                            master = design.SlideMaster;
                            customLayouts = master.CustomLayouts;
                            string designName = PowerPointComUtil.SafeString(delegate { return design.Name; });
                            int layoutCount = customLayouts.Count;
                            for (int layoutIndex = 1; layoutIndex <= layoutCount; layoutIndex++)
                            {
                                PowerPoint.CustomLayout layout = null;
                                try
                                {
                                    layout = customLayouts[layoutIndex];
                                    JObject item = new JObject();
                                    item["index"] = layoutIndex;
                                    item["name"] = PowerPointComUtil.SafeString(delegate { return layout.Name; });
                                    item["design"] = designName;
                                    item["isBuiltIn"] = IsBuiltInLayout(layout.Name);
                                    item["placeholderCount"] = PowerPointComUtil.SafeInt(delegate
                                    {
                                        dynamic dynamicLayout = layout;
                                        return Convert.ToInt32((object)dynamicLayout.Placeholders.Count);
                                    });
                                    layouts.Add(item);
                                }
                                finally
                                {
                                    ComHelper.Release(layout);
                                }
                            }
                        }
                        finally
                        {
                            ComHelper.Release(customLayouts, master, design);
                        }
                    }
                }
                finally
                {
                    ComHelper.Release(designs);
                }

                JObject data = new JObject();
                data["layoutCount"] = layouts.Count;
                data["layouts"] = layouts;
                return Success(
                    "已读取当前演示文稿可用版式，共 " + layouts.Count.ToString(CultureInfo.InvariantCulture) + " 个。",
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

        private static OperationResult ExecuteListThemes(
            PowerPoint.Application application,
            JObject arguments)
        {
            string root;
            string rootError;
            if (!WorkspacePathResolver.TryGetWorkspaceRoot(out root, out rootError))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, rootError);
            }

            JArray themes = new JArray();
            try
            {
                string[] files = Directory.GetFiles(root, "*.thmx", SearchOption.TopDirectoryOnly);
                Array.Sort(files, StringComparer.OrdinalIgnoreCase);
                foreach (string file in files)
                {
                    FileInfo info = new FileInfo(file);
                    JObject item = new JObject();
                    item["fileName"] = info.Name;
                    item["path"] = info.FullName;
                    item["modifiedTime"] = info.LastWriteTimeUtc.ToString("o", CultureInfo.InvariantCulture);
                    item["sizeBytes"] = info.Length;
                    themes.Add(item);
                }
            }
            catch (Exception ex)
            {
                return Fail(OfficeErrorCode.INTERNAL_ERROR, "读取工作目录主题文件失败：" + ex.Message);
            }

            JObject data = new JObject();
            data["themeCount"] = themes.Count;
            data["themes"] = themes;
            return Success(
                "工作目录共有 " + themes.Count.ToString(CultureInfo.InvariantCulture) + " 个 .thmx 主题文件。",
                "read",
                false,
                "ppt:themes",
                data);
        }

        private static OperationResult ExecuteGetThumbnail(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                int width = PowerPointArgumentReader.GetInt(arguments, 640, "width");
                if (width <= 0 || width > 4096)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "width 必须在 1..4096 之间。");
                }
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                float slideWidth = PowerPointComUtil.SafeFloat(delegate { return presentation.PageSetup.SlideWidth; });
                float slideHeight = PowerPointComUtil.SafeFloat(delegate { return presentation.PageSetup.SlideHeight; });
                int height = ComputeHeight(width, slideWidth, slideHeight);

                string tempFile = Path.Combine(
                    Path.GetTempPath(),
                    "AutoOfficeMCP-ppt-thumb-" + Guid.NewGuid().ToString("N") + ".png");
                try
                {
                    dynamic dynamicSlide = slide;
                    try
                    {
                        dynamicSlide.Export(tempFile, "PNG", width, height);
                    }
                    catch (Exception exportError)
                    {
                        return Fail(
                            OfficeErrorCode.INTERNAL_ERROR,
                            "生成缩略图失败：PowerPoint 窗口可能已最小化、无可用窗口或 COM 导出被阻塞；" +
                            "请恢复/激活 PowerPoint 窗口后重试。原始错误：" + exportError.Message);
                    }
                    if (!File.Exists(tempFile))
                    {
                        return Fail(OfficeErrorCode.INTERNAL_ERROR, "PowerPoint 未生成缩略图文件。");
                    }
                    byte[] imageBytes = File.ReadAllBytes(tempFile);
                    JObject data = new JObject();
                    data["slideIndex"] = slide.SlideIndex;
                    data["width"] = width;
                    data["height"] = height;
                    data["mimeType"] = "image/png";
                    data["pngBytes"] = imageBytes.Length;
                    OperationResult result = Success(
                        "已生成第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                        " 张幻灯片缩略图（" + width.ToString(CultureInfo.InvariantCulture) + "×" +
                        height.ToString(CultureInfo.InvariantCulture) + "）。",
                        "read",
                        false,
                        SlideRef(presentation, slide),
                        data);
                    result.ImageBytes = imageBytes;
                    return result;
                }
                finally
                {
                    TryDeleteFile(tempFile);
                }
            }
            finally
            {
                ComHelper.Release(slide, presentation);
            }
        }

        private static OperationResult ExecuteExportImages(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                string scope = PowerPointArgumentReader.GetTrimmedString(arguments, "scope");
                if (scope.Length == 0)
                {
                    scope = "all";
                }
                string outputDir = PowerPointArgumentReader.GetTrimmedString(arguments, "outputDir");
                string outputPath;
                string pathError;
                if (outputDir.Length == 0)
                {
                    string root;
                    string rootError;
                    if (!WorkspacePathResolver.TryGetWorkspaceRoot(out root, out rootError))
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, rootError);
                    }
                    outputPath = Path.Combine(root, "export");
                }
                else if (!WorkspacePathResolver.TryResolve(outputDir, out outputPath, out pathError))
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, pathError);
                }
                string directoryError;
                if (!WorkspacePathResolver.TryEnsureOutputDirectory(outputPath, out directoryError))
                {
                    return Fail(OfficeErrorCode.READ_ONLY, directoryError);
                }

                int width = PowerPointArgumentReader.GetInt(arguments, 1920, "width");
                if (width <= 0 || width > 8192)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "width 必须在 1..8192 之间。");
                }
                string format = PowerPointArgumentReader.GetTrimmedString(arguments, "format").ToLowerInvariant();
                if (format.Length == 0)
                {
                    format = "png";
                }
                if (format != "png" && format != "jpg" && format != "jpeg")
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "format 只能是 png 或 jpg。");
                }
                string extension = format == "png" ? ".png" : ".jpg";
                string filter = format == "png" ? "PNG" : "JPG";

                float slideWidth = PowerPointComUtil.SafeFloat(delegate { return presentation.PageSetup.SlideWidth; });
                float slideHeight = PowerPointComUtil.SafeFloat(delegate { return presentation.PageSetup.SlideHeight; });
                int height = ComputeHeight(width, slideWidth, slideHeight);

                PowerPoint.Slides slides = null;
                JArray files = new JArray();
                try
                {
                    slides = presentation.Slides;
                    if (string.Equals(scope, "current", StringComparison.OrdinalIgnoreCase))
                    {
                        int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                        PowerPoint.Slide slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                        try
                        {
                            files.Add(ExportOneSlide(slide, outputPath, extension, filter, width, height));
                        }
                        finally
                        {
                            ComHelper.Release(slide);
                        }
                    }
                    else
                    {
                        int slideCount = slides.Count;
                        for (int index = 1; index <= slideCount; index++)
                        {
                            PowerPoint.Slide slide = null;
                            try
                            {
                                slide = slides[index];
                                files.Add(ExportOneSlide(slide, outputPath, extension, filter, width, height));
                            }
                            finally
                            {
                                ComHelper.Release(slide);
                            }
                        }
                    }
                }
                finally
                {
                    ComHelper.Release(slides);
                }

                JObject data = new JObject();
                data["outputDir"] = outputPath;
                data["format"] = format == "jpeg" ? "jpg" : format;
                data["width"] = width;
                data["height"] = height;
                data["fileCount"] = files.Count;
                data["files"] = files;
                return Success(
                    "已导出 " + files.Count.ToString(CultureInfo.InvariantCulture) + " 张幻灯片图片到：" + outputPath,
                    "read",
                    files.Count > 0,
                    "ppt:" + presentation.Name,
                    data);
            }
            finally
            {
                ComHelper.Release(presentation);
            }
        }

        private static OperationResult ExecuteExportPdf(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                string scope = PowerPointArgumentReader.GetTrimmedString(arguments, "scope");
                if (scope.Length == 0)
                {
                    scope = "all";
                }
                string outputFile = PowerPointArgumentReader.GetTrimmedString(arguments, "outputFile");
                string outputPath;
                if (outputFile.Length == 0)
                {
                    string root;
                    string rootError;
                    if (!WorkspacePathResolver.TryGetWorkspaceRoot(out root, out rootError))
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, rootError);
                    }
                    string baseName = PowerPointComUtil.SafeString(delegate { return presentation.Name; });
                    baseName = Path.GetFileNameWithoutExtension(baseName);
                    if (string.IsNullOrWhiteSpace(baseName))
                    {
                        baseName = "presentation";
                    }
                    outputPath = Path.Combine(root, baseName + ".pdf");
                }
                else
                {
                    string pathError;
                    if (!WorkspacePathResolver.TryResolve(outputFile, out outputPath, out pathError))
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, pathError);
                    }
                    if (Directory.Exists(outputPath))
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "outputFile 是目录，不是 PDF 文件：" + outputPath);
                    }
                    if (!string.Equals(Path.GetExtension(outputPath), ".pdf", StringComparison.OrdinalIgnoreCase))
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "outputFile 必须以 .pdf 结尾：" + outputPath);
                    }
                }

                string parent = Path.GetDirectoryName(outputPath);
                string directoryError;
                if (!WorkspacePathResolver.TryEnsureOutputDirectory(parent, out directoryError))
                {
                    return Fail(OfficeErrorCode.READ_ONLY, directoryError);
                }
                int pageCount = presentation.Slides.Count;
                bool scopeCurrent = string.Equals(scope, "current", StringComparison.OrdinalIgnoreCase);
                int targetSlideIndex = -1;
                if (scopeCurrent)
                {
                    int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                    PowerPoint.Slide slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                    try
                    {
                        pageCount = 1;
                        targetSlideIndex = slide.SlideIndex;
                    }
                    finally
                    {
                        ComHelper.Release(slide);
                    }
                }

                string exportError;
                if (!ExportPdfViaCopy(application, presentation, outputPath, targetSlideIndex, out exportError))
                {
                    return Fail(OfficeErrorCode.INTERNAL_ERROR, exportError);
                }

                if (!File.Exists(outputPath))
                {
                    return Fail(OfficeErrorCode.INTERNAL_ERROR, "PowerPoint 未生成 PDF 文件：" + outputPath);
                }
                FileInfo info = new FileInfo(outputPath);
                JObject data = new JObject();
                data["filePath"] = outputPath;
                data["pageCount"] = pageCount;
                data["scope"] = scope;
                data["sizeBytes"] = info.Length;
                return Success(
                    "已导出 PDF：" + outputPath + "（" + pageCount.ToString(CultureInfo.InvariantCulture) + " 页）。",
                    "read",
                    true,
                    "ppt:" + presentation.Name,
                    data);
            }
            finally
            {
                ComHelper.Release(presentation);
            }
        }

        private static bool ExportPdfViaCopy(
            PowerPoint.Application application,
            PowerPoint.Presentation source,
            string outputPath,
            int targetSlideIndex,
            out string error)
        {
            error = string.Empty;
            string tempPptx = Path.Combine(
                Path.GetTempPath(),
                "AutoOfficeMCP-ppt-pdf-" + Guid.NewGuid().ToString("N") + ".pptx");
            string tempPdf = Path.Combine(
                Path.GetTempPath(),
                "AutoOfficeMCP-ppt-pdf-" + Guid.NewGuid().ToString("N") + ".pdf");
            PowerPoint.Presentation tempPresentation = null;
            try
            {
                dynamic dynamicSource = source;
                dynamicSource.SaveCopyAs(tempPptx, 24);
                dynamic presentations = application.Presentations;
                object opened = presentations.Open(tempPptx, 0, 0, 0);
                tempPresentation = opened as PowerPoint.Presentation;
                if (tempPresentation == null)
                {
                    error = "PowerPoint 无法打开 PDF 导出临时副本。";
                    return false;
                }

                if (targetSlideIndex > 0)
                {
                    PowerPoint.Slides slides = tempPresentation.Slides;
                    try
                    {
                        int count = slides.Count;
                        if (targetSlideIndex > count)
                        {
                            error = "导出的目标页超出临时副本页数。";
                            return false;
                        }
                        for (int index = count; index >= 1; index--)
                        {
                            if (index == targetSlideIndex)
                            {
                                continue;
                            }
                            PowerPoint.Slide slide = null;
                            try
                            {
                                slide = slides[index];
                                slide.Delete();
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
                }

                dynamic dynamicTemp = tempPresentation;
                dynamicTemp.SaveCopyAs(tempPdf, 32);
                if (!File.Exists(tempPdf))
                {
                    error = "PowerPoint 未生成 PDF 临时文件。";
                    return false;
                }
                File.Copy(tempPdf, outputPath, true);
                return true;
            }
            catch (Exception ex)
            {
                error = "PowerPoint 导出 PDF 失败：" + ex.Message;
                return false;
            }
            finally
            {
                if (tempPresentation != null)
                {
                    try
                    {
                        tempPresentation.Close();
                    }
                    catch (Exception cleanupError)
                    {
                        Console.Error.WriteLine("[ppt] 关闭 PDF 临时演示文稿失败：" + cleanupError.Message);
                    }
                    ComHelper.Release(tempPresentation);
                }
                try
                {
                    dynamic windows = source.Windows;
                    if (Convert.ToInt32((object)windows.Count) > 0)
                    {
                        dynamic firstWindow = windows[1];
                        firstWindow.Activate();
                    }
                }
                catch (Exception)
                {
                }
                TryDeleteFile(tempPptx);
                TryDeleteFile(tempPdf);
            }
        }

        private static JArray ReadShapes(PowerPoint.Slide slide, bool includeFormat)
        {
            JArray shapes = new JArray();
            PowerPoint.Shapes shapeCollection = null;
            try
            {
                shapeCollection = slide.Shapes;
                int count = shapeCollection.Count;
                for (int index = 1; index <= count; index++)
                {
                    PowerPoint.Shape shape = null;
                    try
                    {
                        shape = shapeCollection[index];
                        JObject info = PowerPointComUtil.BuildShapeInfo(shape);
                        if (includeFormat)
                        {
                            PowerPointComUtil.AppendShapeFormat(info, shape);
                        }
                        shapes.Add(info);
                    }
                    finally
                    {
                        ComHelper.Release(shape);
                    }
                }
            }
            finally
            {
                ComHelper.Release(shapeCollection);
            }
            return shapes;
        }

        private static void ReadSlideTextPlaceholders(
            PowerPoint.Slide slide,
            out string title,
            out string content)
        {
            title = string.Empty;
            content = string.Empty;
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
                        if (Convert.ToInt32((object)dynamicShape.Type) != 14)
                        {
                            continue;
                        }
                        int placeholderType = Convert.ToInt32((object)shape.PlaceholderFormat.Type);
                        if (placeholderType == 1 || placeholderType == 3)
                        {
                            if (title.Length == 0)
                            {
                                title = PowerPointComUtil.GetShapeText(shape);
                            }
                        }
                        else if (placeholderType == 2 || placeholderType == 4)
                        {
                            if (content.Length == 0)
                            {
                                content = PowerPointComUtil.GetShapeText(shape);
                            }
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
        }

        private static JObject ExportOneSlide(
            PowerPoint.Slide slide,
            string outputDir,
            string extension,
            string filter,
            int width,
            int height)
        {
            string fileName = "slide-" + slide.SlideIndex.ToString("D3", CultureInfo.InvariantCulture) + extension;
            string filePath = Path.Combine(outputDir, fileName);
            TryDeleteFile(filePath);
            dynamic dynamicSlide = slide;
            dynamicSlide.Export(filePath, filter, width, height);
            if (!File.Exists(filePath))
            {
                throw new OfficeComException(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "PowerPoint 未生成幻灯片图片：" + filePath);
            }
            FileInfo info = new FileInfo(filePath);
            JObject item = new JObject();
            item["slideIndex"] = slide.SlideIndex;
            item["fileName"] = fileName;
            item["filePath"] = filePath;
            item["width"] = width;
            item["height"] = height;
            item["sizeBytes"] = info.Length;
            return item;
        }

        private static int ComputeHeight(int width, float slideWidth, float slideHeight)
        {
            if (width <= 0 || slideWidth <= 0 || slideHeight <= 0)
            {
                return 0;
            }
            return Math.Max(1, (int)Math.Round(width * (slideHeight / slideWidth)));
        }

        private static string DescribeAspectRatio(float width, float height)
        {
            if (width <= 0 || height <= 0)
            {
                return "custom";
            }
            float ratio = width / height;
            if (Math.Abs(ratio - (16f / 9f)) < 0.02f)
            {
                return "16:9";
            }
            if (Math.Abs(ratio - (4f / 3f)) < 0.02f)
            {
                return "4:3";
            }
            return "custom";
        }

        private static bool IsBuiltInLayout(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }
            string[] builtInNames = new string[]
            {
                "Title Slide", "Title and Content", "Section Header", "Two Content",
                "Comparison", "Title Only", "Blank", "Content with Caption", "Picture with Caption",
                "标题幻灯片", "标题和内容", "节标题", "两栏内容", "比较", "仅标题", "空白",
                "内容与标题", "图片与标题"
            };
            foreach (string builtIn in builtInNames)
            {
                if (string.Equals(name.Trim(), builtIn, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        private static bool IsAbsolutePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }
            try
            {
                return Path.IsPathRooted(path);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void TryDeleteFile(string path)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[ppt] 清理临时文件失败：" + path + "；" + ex.Message);
            }
        }
    }
}
