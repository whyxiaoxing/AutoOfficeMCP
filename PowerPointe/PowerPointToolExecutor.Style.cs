using System;
using System.IO;
using Newtonsoft.Json.Linq;
using Office;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointe
{
    public sealed partial class PowerPointToolExecutor
    {
        private static OperationResult ExecuteFormatSlide(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                string background = PowerPointArgumentReader.GetTrimmedString(arguments, "background");
                string layout = PowerPointArgumentReader.GetTrimmedString(arguments, "layout");
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);

                bool changed = false;
                if (background.Length > 0)
                {
                    PowerPointComUtil.ApplyBackground(slide, background);
                    changed = true;
                }
                if (layout.Length > 0)
                {
                    PowerPoint.PpSlideLayout requested = PowerPointComUtil.MapLayout(layout);
                    slide.Layout = requested;
                    changed = true;
                }

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["background"] = background;
                data["layout"] = layout;
                return Success(
                    "已格式化第 " + slide.SlideIndex.ToString() + " 张幻灯片。",
                    "write",
                    changed,
                    SlideRef(presentation, slide),
                    data);
            }
            finally
            {
                ComHelper.Release(slide, presentation);
            }
        }

        private static OperationResult ExecuteAddAnimation(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.Shapes shapes = null;
            PowerPoint.TimeLine timeline = null;
            PowerPoint.Sequence sequence = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                string effectName = PowerPointArgumentReader.GetTrimmedString(arguments, "effect");
                if (effectName.Length == 0)
                {
                    effectName = "fadeIn";
                }
                string targetShapes = PowerPointArgumentReader.GetTrimmedString(arguments, "targetShapes");
                if (targetShapes.Length == 0)
                {
                    targetShapes = "all";
                }
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                float duration = PowerPointArgumentReader.GetFloat(arguments, 0f, "duration");
                float delay = PowerPointArgumentReader.GetFloat(arguments, 0f, "delay");
                int order = PowerPointArgumentReader.GetInt(arguments, 0, "order");
                string triggerText = PowerPointArgumentReader.GetTrimmedString(
                    arguments, "trigger", "startMode");
                int trigger = PowerPointComUtil.MapTriggerType(triggerText);
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                timeline = slide.TimeLine;
                sequence = timeline.MainSequence;
                shapes = slide.Shapes;

                int effect = PowerPointComUtil.MapAnimationEffect(effectName);
                int added = 0;
                JArray warnings = new JArray();
                for (int index = 1; index <= shapes.Count; index++)
                {
                    PowerPoint.Shape shape = null;
                    object animationEffect = null;
                    try
                    {
                        shape = shapes[index];
                        if (!ShouldAnimateShape(shape, targetShapes))
                        {
                            continue;
                        }
                        try
                        {
                            dynamic dynamicSequence = sequence;
                            int insertIndex = order >= 1 ? order + added : -1;
                            animationEffect = dynamicSequence.AddEffect(
                                shape, effect, 0, trigger, insertIndex);
                            if (duration > 0f || delay > 0f)
                            {
                                dynamic effectObject = animationEffect;
                                dynamic timing = effectObject.Timing;
                                if (duration > 0f)
                                {
                                    timing.Duration = duration;
                                }
                                if (delay > 0f)
                                {
                                    timing.TriggerDelayTime = delay;
                                }
                            }
                            added++;
                        }
                        catch (Exception ex)
                        {
                            // 某些形状类型不支持动画；跳过，不把整页判失败。
                            warnings.Add("形状 " + PowerPointComUtil.SafeString(delegate { return shape.Name; }) +
                                         " 添加动画失败：" + ex.Message);
                        }
                    }
                    finally
                    {
                        ComHelper.Release(animationEffect, shape);
                    }
                }

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["effect"] = effectName;
                data["targetShapes"] = targetShapes;
                data["duration"] = duration;
                data["delay"] = delay;
                data["order"] = order;
                data["trigger"] = PowerPointComUtil.MapTriggerName(trigger);
                data["animatedShapeCount"] = added;
                if (warnings.Count > 0)
                {
                    data["warnings"] = warnings;
                }
                return Success(
                    "已为第 " + slide.SlideIndex.ToString() + " 张幻灯片的 " +
                    added.ToString() + " 个形状添加动画。",
                    "write",
                    added > 0,
                    SlideRef(presentation, slide),
                    data);
            }
            finally
            {
                ComHelper.Release(sequence, timeline, shapes, slide, presentation);
            }
        }

        private static bool ShouldAnimateShape(PowerPoint.Shape shape, string targetShapes)
        {
            if (shape == null)
            {
                return false;
            }
            dynamic dynamicShape = shape;
            if (!string.Equals(targetShapes, "all", StringComparison.OrdinalIgnoreCase))
            {
                if (Convert.ToInt32(dynamicShape.Type) != 14) // msoPlaceholder
                {
                    return false;
                }
                PowerPoint.PpPlaceholderType type = shape.PlaceholderFormat.Type;
                if (string.Equals(targetShapes, "title", StringComparison.OrdinalIgnoreCase))
                {
                    return type == PowerPoint.PpPlaceholderType.ppPlaceholderTitle ||
                           type == PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle;
                }
                if (string.Equals(targetShapes, "content", StringComparison.OrdinalIgnoreCase))
                {
                    return type == PowerPoint.PpPlaceholderType.ppPlaceholderBody ||
                           type == PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle;
                }
                return false;
            }
            return Convert.ToInt32(dynamicShape.HasTextFrame) != 0; // msoTrue = -1
        }

        private static OperationResult ExecuteApplyTransition(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slides slides = null;
            PowerPoint.Slide slide = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                string transitionType = PowerPointArgumentReader.RequireString(arguments, "transitionType");
                string scope = PowerPointArgumentReader.GetTrimmedString(arguments, "scope");
                if (scope.Length == 0)
                {
                    scope = "all";
                }
                float duration = PowerPointArgumentReader.GetFloat(arguments, 1.0f, "duration");
                if (duration <= 0)
                {
                    duration = 1.0f;
                }
                PowerPoint.PpEntryEffect effect =
                    (PowerPoint.PpEntryEffect)PowerPointComUtil.MapTransitionEffect(transitionType);
                slides = presentation.Slides;
                int processed = 0;

                if (string.Equals(scope, "all", StringComparison.OrdinalIgnoreCase))
                {
                    for (int index = 1; index <= slides.Count; index++)
                    {
                        PowerPoint.Slide item = null;
                        PowerPoint.SlideShowTransition transition = null;
                        try
                        {
                            item = slides[index];
                            transition = item.SlideShowTransition;
                            transition.EntryEffect = effect;
                            transition.Duration = duration;
                            dynamic dynamicTransition = transition;
                            dynamicTransition.AdvanceOnClick = -1; // msoTrue
                            processed++;
                        }
                        finally
                        {
                            ComHelper.Release(transition, item);
                        }
                    }
                }
                else
                {
                    int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                    slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                    PowerPoint.SlideShowTransition transition = null;
                    try
                    {
                        transition = slide.SlideShowTransition;
                        transition.EntryEffect = effect;
                        transition.Duration = duration;
                        dynamic dynamicTransition = transition;
                        dynamicTransition.AdvanceOnClick = -1; // msoTrue
                        processed = 1;
                    }
                    finally
                    {
                        ComHelper.Release(transition);
                    }
                }

                JObject data = new JObject();
                data["transitionType"] = transitionType;
                data["scope"] = scope;
                data["duration"] = duration;
                data["processedCount"] = processed;
                return Success(
                    "已为 " + processed.ToString() + " 张幻灯片应用切换效果 " + transitionType + "。",
                    "write",
                    processed > 0,
                    "ppt:" + presentation.Name,
                    data);
            }
            finally
            {
                ComHelper.Release(slide, slides, presentation);
            }
        }

        private static OperationResult ExecuteBeautifySlides(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slides slides = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                string scope = PowerPointArgumentReader.GetTrimmedString(arguments, "scope");
                if (scope.Length == 0)
                {
                    scope = "all";
                }
                JObject theme = PowerPointArgumentReader.GetObject(arguments, "theme");
                if (theme == null)
                {
                    theme = BuildDefaultTheme();
                }

                slides = presentation.Slides;
                int processed = 0;
                if (string.Equals(scope, "all", StringComparison.OrdinalIgnoreCase))
                {
                    for (int index = 1; index <= slides.Count; index++)
                    {
                        PowerPoint.Slide slide = null;
                        try
                        {
                            slide = slides[index];
                            BeautifySlide(slide, theme);
                            processed++;
                        }
                        finally
                        {
                            ComHelper.Release(slide);
                        }
                    }
                }
                else
                {
                    PowerPoint.Slide slide = null;
                    try
                    {
                        int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                        slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                        BeautifySlide(slide, theme);
                        processed = 1;
                    }
                    finally
                    {
                        ComHelper.Release(slide);
                    }
                }

                JObject data = new JObject();
                data["scope"] = scope;
                data["processedCount"] = processed;
                return Success(
                    "已美化 " + processed.ToString() + " 张幻灯片。",
                    "write",
                    processed > 0,
                    "ppt:" + presentation.Name,
                    data);
            }
            finally
            {
                ComHelper.Release(slides, presentation);
            }
        }

        private static void BeautifySlide(PowerPoint.Slide slide, JObject theme)
        {
            if (slide == null || theme == null)
            {
                return;
            }
            string background = theme.Value<string>("background");
            PowerPointComUtil.ApplyBackground(slide, background);

            JObject titleFont = theme["titleFont"] as JObject;
            JObject bodyFont = theme["bodyFont"] as JObject;
            PowerPoint.Shapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                for (int index = 1; index <= shapes.Count; index++)
                {
                    PowerPoint.Shape shape = null;
                    PowerPoint.TextRange textRange = null;
                    try
                    {
                        shape = shapes[index];
                        dynamic dynamicShape = shape;
                        if (Convert.ToInt32(dynamicShape.HasTextFrame) == 0)
                        {
                            continue;
                        }
                        bool isTitle = Convert.ToInt32(dynamicShape.Type) == 14 &&
                                       (shape.PlaceholderFormat.Type ==
                                        PowerPoint.PpPlaceholderType.ppPlaceholderTitle ||
                                        shape.PlaceholderFormat.Type ==
                                        PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle);
                        JObject font = isTitle ? titleFont : bodyFont;
                        if (font == null)
                        {
                            continue;
                        }
                        textRange = shape.TextFrame.TextRange;
                        string fontName = font.Value<string>("name");
                        float fontSize = font.Value<float?>("size") ?? 0;
                        string color = font.Value<string>("color");
                        bool bold = font.Value<bool?>("bold") ?? false;
                        PowerPointComUtil.ApplyFont(textRange, fontName, fontSize, color, bold);
                    }
                    finally
                    {
                        ComHelper.Release(textRange, shape);
                    }
                }
            }
            finally
            {
                ComHelper.Release(shapes);
            }
        }

        private static JObject BuildDefaultTheme()
        {
            JObject theme = new JObject();
            theme["background"] = "#F7F3EA";
            JObject titleFont = new JObject();
            titleFont["name"] = "Microsoft YaHei";
            titleFont["size"] = 30;
            titleFont["bold"] = true;
            titleFont["color"] = "#7A3E2E";
            theme["titleFont"] = titleFont;
            JObject bodyFont = new JObject();
            bodyFont["name"] = "Microsoft YaHei";
            bodyFont["size"] = 20;
            bodyFont["color"] = "#27323A";
            theme["bodyFont"] = bodyFont;
            return theme;
        }

        private static OperationResult ExecuteApplyTheme(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                string themeFile = PowerPointArgumentReader.GetTrimmedString(arguments, "themeFile");
                string themeName = PowerPointArgumentReader.GetTrimmedString(arguments, "themeName");
                if (themeFile.Length > 0)
                {
                    if (!File.Exists(themeFile))
                    {
                        return Fail(
                            OfficeErrorCode.FILE_NOT_FOUND,
                            "主题文件不存在：" + themeFile);
                    }
                    presentation.ApplyTheme(themeFile);
                    JObject data = new JObject();
                    data["themeFile"] = themeFile;
                    return Success(
                        "已应用主题文件：" + Path.GetFileName(themeFile) + "。",
                        "write",
                        true,
                        "ppt:" + presentation.Name,
                        data);
                }
                if (themeName.Length > 0)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "PowerPoint COM 不支持按内置主题名直接应用；请提供 .thmx 主题文件路径（themeFile），" +
                        "或在 PowerPoint 设计菜单中手动应用「" + themeName + "」。");
                }
                return Fail(
                    OfficeErrorCode.INVALID_PARAMS,
                    "需要 themeFile 或 themeName 参数。");
            }
            finally
            {
                ComHelper.Release(presentation);
            }
        }
    }
}
