using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Office;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointe
{
    public sealed partial class PowerPointToolExecutor
    {
        private static OperationResult ExecuteHideSlide(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                bool hidden = PowerPointArgumentReader.GetBool(arguments, true, "hidden");
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                dynamic transition = slide.SlideShowTransition;
                transition.Hidden = hidden ? -1 : 0;
                bool applied = Convert.ToInt32((object)transition.Hidden) != 0;

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["hidden"] = applied;
                return Success(
                    "已" + (applied ? "隐藏" : "取消隐藏") + "第 " +
                    slide.SlideIndex.ToString(CultureInfo.InvariantCulture) + " 张幻灯片。",
                    "write",
                    applied == hidden,
                    SlideRef(presentation, slide),
                    data);
            }
            finally
            {
                ComHelper.Release(slide, presentation);
            }
        }

        private static OperationResult ExecuteReorderSlides(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slides slides = null;
            try
            {
                if (!PowerPointArgumentReader.GetBool(arguments, false, "confirm"))
                {
                    return OperationResult.Cancelled(
                        "ppt_reorder_slides 需要显式 confirm=true 才会重排；本次未修改演示文稿。");
                }
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                JArray order = PowerPointArgumentReader.GetArray(arguments, "order");
                if (order == null || order.Count == 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "order 必须是非空数组。");
                }
                slides = presentation.Slides;
                int count = slides.Count;
                if (order.Count != count)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "order 长度必须等于幻灯片总数 " + count.ToString(CultureInfo.InvariantCulture) + "。");
                }

                bool[] seen = new bool[count + 1];
                int[] desiredIds = new int[count];
                for (int index = 0; index < count; index++)
                {
                    JToken token = order[index];
                    int sourceIndex;
                    try
                    {
                        sourceIndex = token.Value<int>();
                    }
                    catch (Exception)
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "order[" + index.ToString(CultureInfo.InvariantCulture) + "] 必须是整数。");
                    }
                    if (sourceIndex < 1 || sourceIndex > count || seen[sourceIndex])
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "order 必须是 1.." + count.ToString(CultureInfo.InvariantCulture) +
                            " 的不重复排列，收到第 " + (index + 1).ToString(CultureInfo.InvariantCulture) +
                            " 项 " + sourceIndex.ToString(CultureInfo.InvariantCulture) + "。");
                    }
                    seen[sourceIndex] = true;
                    PowerPoint.Slide source = null;
                    try
                    {
                        source = slides[sourceIndex];
                        dynamic dynamicSource = source;
                        desiredIds[index] = Convert.ToInt32((object)dynamicSource.SlideID);
                    }
                    finally
                    {
                        ComHelper.Release(source);
                    }
                }

                for (int targetIndex = 1; targetIndex <= count; targetIndex++)
                {
                    int desiredId = desiredIds[targetIndex - 1];
                    int currentIndex = -1;
                    for (int index = 1; index <= count; index++)
                    {
                        PowerPoint.Slide item = null;
                        try
                        {
                            item = slides[index];
                            dynamic dynamicItem = item;
                            if (Convert.ToInt32((object)dynamicItem.SlideID) == desiredId)
                            {
                                currentIndex = index;
                                break;
                            }
                        }
                        finally
                        {
                            ComHelper.Release(item);
                        }
                    }
                    if (currentIndex < 0)
                    {
                        return Fail(OfficeErrorCode.INTERNAL_ERROR, "重排过程中找不到目标幻灯片。");
                    }
                    if (currentIndex != targetIndex)
                    {
                        PowerPoint.Slide moving = null;
                        try
                        {
                            moving = slides[currentIndex];
                            moving.MoveTo(targetIndex);
                        }
                        finally
                        {
                            ComHelper.Release(moving);
                        }
                    }
                }

                JArray newOrder = new JArray();
                for (int index = 1; index <= count; index++)
                {
                    newOrder.Add(index);
                }
                JObject data = new JObject();
                data["applied"] = true;
                data["newOrder"] = newOrder;
                data["slideCount"] = count;
                return Success(
                    "已按指定顺序重排 " + count.ToString(CultureInfo.InvariantCulture) + " 张幻灯片。",
                    "write",
                    true,
                    "ppt:" + presentation.Name,
                    data);
            }
            finally
            {
                ComHelper.Release(slides, presentation);
            }
        }

        private static OperationResult ExecuteClearSlide(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            try
            {
                if (!PowerPointArgumentReader.GetBool(arguments, false, "confirm"))
                {
                    return OperationResult.Cancelled(
                        "ppt_clear_slide 需要显式 confirm=true 才会清空；本次未修改演示文稿。");
                }
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                bool keepLayout = PowerPointArgumentReader.GetBool(arguments, true, "keepLayout");
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);

                PowerPoint.Shapes shapes = null;
                int removed = 0;
                try
                {
                    shapes = slide.Shapes;
                    removed = shapes.Count;
                    // 一次 Range().Delete() 作为单个原生撤销组，避免逐形状删除导致撤销只恢复部分形状。
                    if (removed > 0)
                    {
                        dynamic dynamicShapes = shapes;
                        dynamicShapes.Range().Delete();
                    }
                }
                finally
                {
                    ComHelper.Release(shapes);
                }
                if (!keepLayout)
                {
                    slide.Layout = PowerPointComUtil.MapLayout("blank");
                }

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["cleared"] = true;
                data["removedShapeCount"] = removed;
                data["keepLayout"] = keepLayout;
                return Success(
                    "已清空第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片，共删除 " + removed.ToString(CultureInfo.InvariantCulture) + " 个形状。",
                    "write",
                    removed > 0,
                    SlideRef(presentation, slide),
                    data);
            }
            finally
            {
                ComHelper.Release(slide, presentation);
            }
        }

        private static OperationResult ExecuteSetSlideSize(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            try
            {
                if (!PowerPointArgumentReader.GetBool(arguments, false, "confirm"))
                {
                    return OperationResult.Cancelled(
                        "ppt_set_slide_size 会改变所有页面的布局，需要显式 confirm=true；本次未修改演示文稿。");
                }
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                string size = PowerPointArgumentReader.GetTrimmedString(arguments, "size").ToLowerInvariant();
                if (size.Length == 0)
                {
                    size = "16:9";
                }

                float width;
                float height;
                switch (size)
                {
                    case "16:9":
                        width = 960f;
                        height = 540f;
                        break;
                    case "4:3":
                        width = 720f;
                        height = 540f;
                        break;
                    case "a4":
                        width = 842f;
                        height = 595f;
                        break;
                    case "custom":
                        width = PowerPointArgumentReader.GetFloat(arguments, 0f, "width");
                        height = PowerPointArgumentReader.GetFloat(arguments, 0f, "height");
                        if (width <= 0 || height <= 0)
                        {
                            return Fail(OfficeErrorCode.INVALID_PARAMS, "size=custom 时 width/height 必须大于 0。");
                        }
                        break;
                    default:
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "size 只能是 16:9、4:3、a4 或 custom。");
                }

                presentation.PageSetup.SlideWidth = width;
                presentation.PageSetup.SlideHeight = height;

                JObject data = new JObject();
                data["size"] = size;
                data["slideWidth"] = presentation.PageSetup.SlideWidth;
                data["slideHeight"] = presentation.PageSetup.SlideHeight;
                return Success(
                    "已把页面尺寸设置为 " + size + "（" +
                    width.ToString(CultureInfo.InvariantCulture) + "×" +
                    height.ToString(CultureInfo.InvariantCulture) + " 磅）。",
                    "write",
                    true,
                    "ppt:" + presentation.Name,
                    data);
            }
            finally
            {
                ComHelper.Release(presentation);
            }
        }
    }
}
