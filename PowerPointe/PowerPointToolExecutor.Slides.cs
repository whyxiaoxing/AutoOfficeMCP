using System;
using Newtonsoft.Json.Linq;
using Office;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointe
{
    public sealed partial class PowerPointToolExecutor
    {
        private static OperationResult ExecuteInsertSlide(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.Slides slides = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                slides = presentation.Slides;
                string position = PowerPointArgumentReader.GetTrimmedString(arguments, "position");
                string layout = PowerPointArgumentReader.GetTrimmedString(arguments, "layout");
                string title = PowerPointArgumentReader.GetString(arguments, "title");
                string content = PowerPointArgumentReader.GetString(arguments, "content");

                int insertIndex;
                if (string.Equals(position, "current", StringComparison.OrdinalIgnoreCase))
                {
                    PowerPoint.Slide current = PowerPointComUtil.ResolveSlide(application, presentation, -1);
                    try
                    {
                        insertIndex = current.SlideIndex + 1;
                    }
                    finally
                    {
                        ComHelper.Release(current);
                    }
                }
                else if (string.Equals(position, "index", StringComparison.OrdinalIgnoreCase))
                {
                    int index = PowerPointArgumentReader.GetInt(arguments, -1, "index");
                    if (index < 0 || index > slides.Count)
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "position=index 时 index 必须在 0.." + slides.Count.ToString() +
                            " 之间（0=插到第一页之前，N=插到第 N 页之后）。");
                    }
                    insertIndex = index + 1;
                }
                else
                {
                    insertIndex = slides.Count + 1;
                }

                slide = slides.Add(insertIndex, PowerPointComUtil.MapLayout(layout));
                PowerPointComUtil.SetPlaceholderText(
                    slide,
                    title,
                    PowerPoint.PpPlaceholderType.ppPlaceholderTitle,
                    PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle);
                if (!string.IsNullOrWhiteSpace(content))
                {
                    bool filled = PowerPointComUtil.SetPlaceholderText(
                        slide,
                        content,
                        PowerPoint.PpPlaceholderType.ppPlaceholderBody,
                        PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle);
                    if (!filled)
                    {
                        PowerPoint.Shape fallbackText = PowerPointComUtil.AddTextBox(
                            slide, content, 50, 150, 600, 300, 18);
                        ComHelper.Release(fallbackText);
                    }
                }

                int slideIndex = slide.SlideIndex;
                JObject data = new JObject();
                data["slideIndex"] = slideIndex;
                data["position"] = string.IsNullOrWhiteSpace(position) ? "end" : position;
                data["index"] = string.Equals(position, "index", StringComparison.OrdinalIgnoreCase)
                    ? PowerPointArgumentReader.GetInt(arguments, -1, "index")
                    : -1;
                data["layout"] = string.IsNullOrWhiteSpace(layout) ? "titleAndContent" : layout;
                data["slideCount"] = slides.Count;
                return Success(
                    "已插入第 " + slideIndex.ToString() + " 张幻灯片。",
                    "write",
                    true,
                    SlideRef(presentation, slide),
                    data);
            }
            finally
            {
                ComHelper.Release(slide, slides, presentation);
            }
        }

        private static OperationResult ExecuteCreateSlides(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slides slides = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                JArray slideSpecs = PowerPointArgumentReader.GetArray(arguments, "slides");
                if (slideSpecs == null || slideSpecs.Count == 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "slides 必须是非空数组。");
                }
                if (slideSpecs.Count > 50)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "slides 最多 50 页。");
                }

                slides = presentation.Slides;
                JArray createdIndices = new JArray();
                JArray warnings = new JArray();
                int created = 0;
                int nextIndex = slides.Count + 1;

                for (int index = 0; index < slideSpecs.Count; index++)
                {
                    JObject spec = slideSpecs[index] as JObject;
                    if (spec == null)
                    {
                        warnings.Add("slides[" + index.ToString() + "] 不是对象，已跳过。");
                        continue;
                    }

                    string title = PowerPointArgumentReader.GetString(spec, "title");
                    string content = PowerPointArgumentReader.GetString(spec, "content");
                    string layout = PowerPointArgumentReader.GetString(spec, "layout");
                    PowerPoint.Slide slide = null;
                    try
                    {
                        slide = slides.Add(nextIndex, PowerPointComUtil.MapLayout(layout));
                        PowerPointComUtil.SetPlaceholderText(
                            slide,
                            title,
                            PowerPoint.PpPlaceholderType.ppPlaceholderTitle,
                            PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle);
                        if (!string.IsNullOrWhiteSpace(content))
                        {
                            bool filled = PowerPointComUtil.SetPlaceholderText(
                                slide,
                                content,
                                PowerPoint.PpPlaceholderType.ppPlaceholderBody,
                                PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle);
                            if (!filled)
                            {
                                PowerPoint.Shape fallbackText = PowerPointComUtil.AddTextBox(
                                    slide, content, 50, 150, 620, 350, 18);
                                ComHelper.Release(fallbackText);
                            }
                        }
                        createdIndices.Add(slide.SlideIndex);
                        created++;
                        nextIndex++;
                    }
                    catch (Exception ex)
                    {
                        JObject partialData = new JObject();
                        partialData["createdCount"] = created;
                        partialData["createdSlideIndices"] = createdIndices;
                        partialData["warnings"] = warnings;
                        partialData["failedIndex"] = index;
                        return OperationResult.PartialApplied(
                            created,
                            "第 " + (index + 1).ToString() + " 页创建失败：" + ex.Message +
                            "；已创建的 " + created.ToString() + " 页不会自动删除。",
                            ObservationFor("write", "批量创建幻灯片只完成 " + created.ToString() + " 页。", true),
                            partialData);
                    }
                    finally
                    {
                        ComHelper.Release(slide);
                    }
                }

                JObject data = new JObject();
                data["createdCount"] = created;
                data["createdSlideIndices"] = createdIndices;
                data["slideCount"] = slides.Count;
                if (warnings.Count > 0)
                {
                    data["warnings"] = warnings;
                }
                return Success(
                    "已创建 " + created.ToString() + " 张幻灯片。",
                    "write",
                    created > 0,
                    "ppt:" + presentation.Name,
                    data);
            }
            finally
            {
                ComHelper.Release(slides, presentation);
            }
        }

        private static OperationResult ExecuteDeleteSlide(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slides slides = null;
            PowerPoint.Slide slide = null;
            try
            {
                if (!PowerPointArgumentReader.GetBool(arguments, false, "confirm"))
                {
                    return OperationResult.Cancelled(
                        "ppt_delete_slide 需要显式 confirm=true 才会删除；本次未修改演示文稿。");
                }

                presentation = PowerPointComUtil.RequireActivePresentation(application);
                slides = presentation.Slides;
                if (slides.Count <= 1)
                {
                    return Fail(OfficeErrorCode.CANCELLED, "演示文稿必须至少保留一张幻灯片。");
                }

                int requestedIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                slide = PowerPointComUtil.ResolveSlide(application, presentation, requestedIndex);
                int deletedIndex = slide.SlideIndex;
                slide.Delete();
                int remaining = slides.Count;

                JObject data = new JObject();
                data["deletedIndex"] = deletedIndex;
                data["slideCount"] = remaining;
                return Success(
                    "已删除第 " + deletedIndex.ToString() + " 张幻灯片。",
                    "write",
                    true,
                    "ppt:" + presentation.Name + "!Slide" + deletedIndex.ToString(),
                    data);
            }
            finally
            {
                ComHelper.Release(slide, slides, presentation);
            }
        }

        private static OperationResult ExecuteDuplicateSlide(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slides slides = null;
            PowerPoint.Slide slide = null;
            PowerPoint.SlideRange pasted = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                slides = presentation.Slides;
                int requestedIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                slide = PowerPointComUtil.ResolveSlide(application, presentation, requestedIndex);
                int sourceIndex = slide.SlideIndex;
                int insertAfter = PowerPointArgumentReader.GetInt(arguments, sourceIndex, "insertAfter");
                if (insertAfter < 1 || insertAfter > slides.Count)
                {
                    insertAfter = sourceIndex;
                }

                slide.Copy();
                pasted = slides.Paste(insertAfter + 1);
                int newIndex = Math.Min(insertAfter + 1, slides.Count);

                JObject data = new JObject();
                data["sourceIndex"] = sourceIndex;
                data["newIndex"] = newIndex;
                data["slideCount"] = slides.Count;
                return Success(
                    "已复制第 " + sourceIndex.ToString() + " 张幻灯片到第 " +
                    newIndex.ToString() + " 张。",
                    "write",
                    true,
                    "ppt:" + presentation.Name + "!Slide" + newIndex.ToString(),
                    data);
            }
            finally
            {
                ComHelper.Release(pasted, slide, slides, presentation);
            }
        }

        private static OperationResult ExecuteMoveSlide(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slides slides = null;
            PowerPoint.Slide slide = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                slides = presentation.Slides;
                int fromIndex = PowerPointArgumentReader.GetInt(arguments, -1, "fromIndex");
                int toIndex = PowerPointArgumentReader.GetInt(arguments, -1, "toIndex");
                if (fromIndex < 1 || fromIndex > slides.Count ||
                    toIndex < 1 || toIndex > slides.Count)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "fromIndex/toIndex 必须在 1.." + slides.Count.ToString() + " 范围内。");
                }

                slide = slides[fromIndex];
                slide.MoveTo(toIndex);

                JObject data = new JObject();
                data["fromIndex"] = fromIndex;
                data["toIndex"] = toIndex;
                data["slideCount"] = slides.Count;
                return Success(
                    "已将第 " + fromIndex.ToString() + " 张幻灯片移动到第 " +
                    toIndex.ToString() + " 张。",
                    "write",
                    true,
                    "ppt:" + presentation.Name + "!Slide" + toIndex.ToString(),
                    data);
            }
            finally
            {
                ComHelper.Release(slide, slides, presentation);
            }
        }

        private static OperationResult ExecuteSetSlideLayout(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                string layout = PowerPointArgumentReader.RequireString(arguments, "layout");
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                PowerPoint.PpSlideLayout requested = PowerPointComUtil.MapLayout(layout);
                slide.Layout = requested;
                PowerPoint.PpSlideLayout applied = slide.Layout;
                if (applied != requested)
                {
                    return Fail(
                        OfficeErrorCode.PARTIAL_APPLY,
                        "版式「" + layout + "」应用后读回为 " + applied.ToString() + "，未能确认成功。");
                }

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["layout"] = layout;
                data["appliedLayout"] = applied.ToString();
                return Success(
                    "已将第 " + slide.SlideIndex.ToString() + " 张幻灯片版式设置为 " + layout + "。",
                    "write",
                    true,
                    SlideRef(presentation, slide),
                    data);
            }
            finally
            {
                ComHelper.Release(slide, presentation);
            }
        }

        private static Observation ObservationFor(string kind, string summary, bool changed)
        {
            return Observation.Create(kind, summary, changed);
        }
    }
}
