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
        private sealed class AlignShapeInfo
        {
            public PowerPoint.Shape Shape;
            public int Id;
            public string Name;
            public float Left;
            public float Top;
            public float Width;
            public float Height;
        }

        private static OperationResult ExecuteSetZOrder(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.Shape shape = null;
            PowerPoint.Shapes shapes = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                int shapeId = PowerPointArgumentReader.GetInt(arguments, -1, "shapeId");
                if (shapeId < 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少必填参数 shapeId。");
                }
                string position = PowerPointArgumentReader.RequireString(arguments, "position").ToLowerInvariant();
                int command;
                switch (position)
                {
                    case "front":
                        command = 0;
                        break;
                    case "back":
                        command = 1;
                        break;
                    case "forward":
                        command = 2;
                        break;
                    case "backward":
                        command = 3;
                        break;
                    default:
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "position 只能是 front、back、forward 或 backward。");
                }

                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                shape = PowerPointComUtil.FindShape(slide, shapeId);
                if (shape == null)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "找不到 shapeId=" + shapeId.ToString(CultureInfo.InvariantCulture) + "。");
                }
                shapes = slide.Shapes;
                int before = PowerPointComUtil.SafeInt(delegate
                {
                    dynamic dynamicShape = shape;
                    return Convert.ToInt32((object)dynamicShape.ZOrderPosition);
                });
                dynamic dynamicTarget = shape;
                dynamicTarget.ZOrder(command);
                int after = PowerPointComUtil.SafeInt(delegate
                {
                    dynamic dynamicShape = shape;
                    return Convert.ToInt32((object)dynamicShape.ZOrderPosition);
                });

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["shapeId"] = shapeId;
                data["shapeName"] = shape.Name;
                data["position"] = position;
                data["beforeZOrder"] = before;
                data["afterZOrder"] = after;
                data["shapeCount"] = shapes.Count;
                return Success(
                    "已调整第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片形状 " + shape.Name + " 的层级（" + before.ToString(CultureInfo.InvariantCulture) +
                    " → " + after.ToString(CultureInfo.InvariantCulture) + "）。",
                    "write",
                    before != after,
                    SlideRef(presentation, slide) + "!" + shape.Name,
                    data);
            }
            finally
            {
                ComHelper.Release(shapes, shape, slide, presentation);
            }
        }

        private static OperationResult ExecuteGroupShapes(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.Shapes shapes = null;
            PowerPoint.ShapeRange range = null;
            PowerPoint.Shape groupShape = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                JArray shapeIds = PowerPointArgumentReader.GetArray(arguments, "shapeIds");
                if (shapeIds == null || shapeIds.Count < 2)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "shapeIds 至少需要两个形状。");
                }
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                shapes = slide.Shapes;
                object[] indexes = new object[shapeIds.Count];
                JArray memberIds = new JArray();
                for (int index = 0; index < shapeIds.Count; index++)
                {
                    int shapeId;
                    try
                    {
                        shapeId = shapeIds[index].Value<int>();
                    }
                    catch (Exception)
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "shapeIds[" + index.ToString(CultureInfo.InvariantCulture) + "] 必须是整数。");
                    }
                    int position = FindShapePosition(shapes, shapeId);
                    if (position < 1)
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "找不到 shapeId=" + shapeId.ToString(CultureInfo.InvariantCulture) + "。");
                    }
                    indexes[index] = (object)position;
                    memberIds.Add(shapeId);
                }

                dynamic dynamicShapes = shapes;
                range = dynamicShapes.Range(indexes) as PowerPoint.ShapeRange;
                if (range == null)
                {
                    return Fail(OfficeErrorCode.INTERNAL_ERROR, "PowerPoint 未返回可组合的形状范围。");
                }
                dynamic dynamicRange = range;
                object grouped = dynamicRange.Group();
                groupShape = grouped as PowerPoint.Shape;
                if (groupShape == null)
                {
                    return Fail(OfficeErrorCode.INTERNAL_ERROR, "PowerPoint 未返回组合后的形状对象。");
                }

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["groupShapeId"] = PowerPointComUtil.SafeInt(delegate
                {
                    dynamic dynamicShape = groupShape;
                    return Convert.ToInt32((object)dynamicShape.Id);
                });
                data["groupShapeName"] = groupShape.Name;
                data["memberIds"] = memberIds;
                data["memberCount"] = memberIds.Count;
                return Success(
                    "已组合第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片上的 " + memberIds.Count.ToString(CultureInfo.InvariantCulture) + " 个形状。",
                    "write",
                    true,
                    SlideRef(presentation, slide) + "!" + groupShape.Name,
                    data);
            }
            finally
            {
                ComHelper.Release(groupShape, range, shapes, slide, presentation);
            }
        }

        private static OperationResult ExecuteUngroupShapes(
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

                PowerPoint.ShapeRange members = null;
                try
                {
                    dynamic dynamicShape = shape;
                    object ungrouped = dynamicShape.Ungroup();
                    members = ungrouped as PowerPoint.ShapeRange;
                    if (members == null)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "目标形状不是组合形状，无法解组。");
                    }
                    JArray memberIds = new JArray();
                    JArray memberNames = new JArray();
                    int count = members.Count;
                    for (int index = 1; index <= count; index++)
                    {
                        PowerPoint.Shape member = null;
                        try
                        {
                            member = members[index];
                            memberIds.Add(PowerPointComUtil.SafeInt(delegate
                            {
                                dynamic dynamicMember = member;
                                return Convert.ToInt32((object)dynamicMember.Id);
                            }));
                            memberNames.Add(member.Name);
                        }
                        finally
                        {
                            ComHelper.Release(member);
                        }
                    }

                    JObject data = new JObject();
                    data["slideIndex"] = slide.SlideIndex;
                    data["sourceShapeId"] = shapeId;
                    data["memberIds"] = memberIds;
                    data["memberNames"] = memberNames;
                    data["memberCount"] = memberIds.Count;
                    return Success(
                        "已解开第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                        " 张幻灯片上的组合，得到 " + memberIds.Count.ToString(CultureInfo.InvariantCulture) +
                        " 个形状。",
                        "write",
                        true,
                        SlideRef(presentation, slide),
                        data);
                }
                finally
                {
                    ComHelper.Release(members);
                }
            }
            finally
            {
                ComHelper.Release(shape, slide, presentation);
            }
        }

        private static OperationResult ExecuteAlignShapes(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            List<AlignShapeInfo> selected = new List<AlignShapeInfo>();
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                string align = PowerPointArgumentReader.GetTrimmedString(arguments, "align").ToLowerInvariant();
                string distribute = PowerPointArgumentReader.GetTrimmedString(
                    arguments, "distribute").ToLowerInvariant();
                if (align.Length == 0 && distribute.Length == 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "align 与 distribute 至少需要提供一项。");
                }
                if (align.Length > 0 &&
                    align != "left" && align != "centerh" && align != "right" &&
                    align != "top" && align != "centerv" && align != "bottom")
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "align 只能是 left/centerH/right/top/centerV/bottom。");
                }
                if (distribute.Length > 0 &&
                    distribute != "horizontal" && distribute != "vertical")
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "distribute 只能是 horizontal 或 vertical。");
                }

                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                JArray shapeIds = PowerPointArgumentReader.GetArray(arguments, "shapeIds");
                if (shapeIds != null && shapeIds.Count > 0)
                {
                    for (int index = 0; index < shapeIds.Count; index++)
                    {
                        int shapeId;
                        try
                        {
                            shapeId = shapeIds[index].Value<int>();
                        }
                        catch (Exception)
                        {
                            return Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "shapeIds[" + index.ToString(CultureInfo.InvariantCulture) + "] 必须是整数。");
                        }
                        PowerPoint.Shape shape = PowerPointComUtil.FindShape(slide, shapeId);
                        if (shape == null)
                        {
                            return Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "找不到 shapeId=" + shapeId.ToString(CultureInfo.InvariantCulture) + "。");
                        }
                        selected.Add(BuildAlignShapeInfo(shape));
                    }
                }
                else
                {
                    PowerPoint.Shapes shapes = null;
                    try
                    {
                        shapes = slide.Shapes;
                        for (int index = 1; index <= shapes.Count; index++)
                        {
                            PowerPoint.Shape shape = shapes[index];
                            selected.Add(BuildAlignShapeInfo(shape));
                        }
                    }
                    finally
                    {
                        ComHelper.Release(shapes);
                    }
                }

                if (selected.Count < 2)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "至少需要两个形状才能对齐。");
                }

                JArray applied = new JArray();
                if (align.Length > 0)
                {
                    ApplyAlignment(selected, align);
                    applied.Add("align:" + align);
                }
                if (distribute.Length > 0)
                {
                    if (selected.Count < 3)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "distribute 至少需要三个形状。");
                    }
                    ApplyDistribution(selected, distribute);
                    applied.Add("distribute:" + distribute);
                }

                JArray items = new JArray();
                foreach (AlignShapeInfo item in selected)
                {
                    JObject shapeInfo = new JObject();
                    shapeInfo["shapeId"] = item.Id;
                    shapeInfo["shapeName"] = item.Name;
                    shapeInfo["left"] = PowerPointComUtil.SafeFloat(delegate { return item.Shape.Left; });
                    shapeInfo["top"] = PowerPointComUtil.SafeFloat(delegate { return item.Shape.Top; });
                    shapeInfo["width"] = item.Width;
                    shapeInfo["height"] = item.Height;
                    items.Add(shapeInfo);
                }

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["shapeCount"] = selected.Count;
                data["applied"] = applied;
                data["shapes"] = items;
                return Success(
                    "已对齐第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片上的 " + selected.Count.ToString(CultureInfo.InvariantCulture) + " 个形状。",
                    "write",
                    true,
                    SlideRef(presentation, slide),
                    data);
            }
            finally
            {
                foreach (AlignShapeInfo item in selected)
                {
                    ComHelper.Release(item.Shape);
                }
                ComHelper.Release(slide, presentation);
            }
        }

        private static int FindShapePosition(PowerPoint.Shapes shapes, int shapeId)
        {
            if (shapes == null)
            {
                return -1;
            }
            int count = shapes.Count;
            for (int position = 1; position <= count; position++)
            {
                PowerPoint.Shape shape = null;
                try
                {
                    shape = shapes[position];
                    dynamic dynamicShape = shape;
                    if (Convert.ToInt32((object)dynamicShape.Id) == shapeId)
                    {
                        return position;
                    }
                }
                catch (Exception)
                {
                }
                finally
                {
                    ComHelper.Release(shape);
                }
            }
            return -1;
        }

        private static AlignShapeInfo BuildAlignShapeInfo(PowerPoint.Shape shape)
        {
            AlignShapeInfo info = new AlignShapeInfo();
            info.Shape = shape;
            info.Id = PowerPointComUtil.SafeInt(delegate
            {
                dynamic dynamicShape = shape;
                return Convert.ToInt32((object)dynamicShape.Id);
            });
            info.Name = PowerPointComUtil.SafeString(delegate { return shape.Name; });
            info.Left = PowerPointComUtil.SafeFloat(delegate { return shape.Left; });
            info.Top = PowerPointComUtil.SafeFloat(delegate { return shape.Top; });
            info.Width = PowerPointComUtil.SafeFloat(delegate { return shape.Width; });
            info.Height = PowerPointComUtil.SafeFloat(delegate { return shape.Height; });
            return info;
        }

        private static void ApplyAlignment(List<AlignShapeInfo> shapes, string align)
        {
            float minLeft = float.MaxValue;
            float maxRight = float.MinValue;
            float minTop = float.MaxValue;
            float maxBottom = float.MinValue;
            foreach (AlignShapeInfo item in shapes)
            {
                minLeft = Math.Min(minLeft, item.Left);
                maxRight = Math.Max(maxRight, item.Left + item.Width);
                minTop = Math.Min(minTop, item.Top);
                maxBottom = Math.Max(maxBottom, item.Top + item.Height);
            }

            foreach (AlignShapeInfo item in shapes)
            {
                float left = item.Left;
                float top = item.Top;
                switch (align)
                {
                    case "left":
                        left = minLeft;
                        break;
                    case "centerh":
                        left = (minLeft + maxRight) / 2f - item.Width / 2f;
                        break;
                    case "right":
                        left = maxRight - item.Width;
                        break;
                    case "top":
                        top = minTop;
                        break;
                    case "centerv":
                        top = (minTop + maxBottom) / 2f - item.Height / 2f;
                        break;
                    case "bottom":
                        top = maxBottom - item.Height;
                        break;
                }
                if (Math.Abs(left - item.Left) > 0.01f)
                {
                    item.Shape.Left = left;
                    item.Left = left;
                }
                if (Math.Abs(top - item.Top) > 0.01f)
                {
                    item.Shape.Top = top;
                    item.Top = top;
                }
            }
        }

        private static void ApplyDistribution(List<AlignShapeInfo> shapes, string direction)
        {
            shapes.Sort(delegate(AlignShapeInfo left, AlignShapeInfo right)
            {
                float leftValue = string.Equals(direction, "horizontal", StringComparison.Ordinal)
                    ? left.Left
                    : left.Top;
                float rightValue = string.Equals(direction, "horizontal", StringComparison.Ordinal)
                    ? right.Left
                    : right.Top;
                return leftValue.CompareTo(rightValue);
            });

            AlignShapeInfo first = shapes[0];
            AlignShapeInfo last = shapes[shapes.Count - 1];
            float totalSize = 0f;
            foreach (AlignShapeInfo item in shapes)
            {
                totalSize += string.Equals(direction, "horizontal", StringComparison.Ordinal)
                    ? item.Width
                    : item.Height;
            }
            float span = string.Equals(direction, "horizontal", StringComparison.Ordinal)
                ? (last.Left + last.Width) - first.Left
                : (last.Top + last.Height) - first.Top;
            float gap = (span - totalSize) / (shapes.Count - 1);
            if (gap < 0f)
            {
                gap = 0f;
            }

            float cursor = string.Equals(direction, "horizontal", StringComparison.Ordinal)
                ? first.Left
                : first.Top;
            foreach (AlignShapeInfo item in shapes)
            {
                if (string.Equals(direction, "horizontal", StringComparison.Ordinal))
                {
                    if (Math.Abs(cursor - item.Left) > 0.01f)
                    {
                        item.Shape.Left = cursor;
                        item.Left = cursor;
                    }
                    cursor += item.Width + gap;
                }
                else
                {
                    if (Math.Abs(cursor - item.Top) > 0.01f)
                    {
                        item.Shape.Top = cursor;
                        item.Top = cursor;
                    }
                    cursor += item.Height + gap;
                }
            }
        }

        private static OperationResult ExecuteCopySlideFromFile(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slides slides = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                string sourceFile = PowerPointArgumentReader.RequireString(arguments, "sourceFile");
                string sourcePath;
                string sourceError;
                if (!WorkspacePathResolver.TryResolve(sourceFile, out sourcePath, out sourceError))
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, sourceError);
                }
                if (!File.Exists(sourcePath))
                {
                    return Fail(OfficeErrorCode.FILE_NOT_FOUND, "源演示文稿不存在：" + sourcePath);
                }
                string extension = Path.GetExtension(sourcePath);
                if (!string.Equals(extension, ".pptx", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(extension, ".pptm", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(extension, ".ppt", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(extension, ".ppsx", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(extension, ".ppsm", StringComparison.OrdinalIgnoreCase))
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "sourceFile 必须是 PowerPoint 文件（.pptx/.pptm/.ppt/.ppsx/.ppsm）：" + sourcePath);
                }

                int sourceSlideIndex = PowerPointArgumentReader.GetInt(arguments, 1, "sourceSlideIndex");
                if (sourceSlideIndex < 1)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "sourceSlideIndex 必须大于等于 1。");
                }
                slides = presentation.Slides;
                int insertAfter = PowerPointArgumentReader.GetInt(arguments, slides.Count, "insertAfter");
                if (insertAfter < 0 || insertAfter > slides.Count)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "insertAfter 必须在 0.." + slides.Count.ToString(CultureInfo.InvariantCulture) + " 之间。");
                }

                int beforeCount = slides.Count;
                dynamic dynamicSlides = slides;
                // 该 PIA 版本 InsertFromFile 返回 int（不是 SlideRange）；插入后按目标位置回读。
                dynamicSlides.InsertFromFile(
                    sourcePath,
                    insertAfter + 1,
                    sourceSlideIndex,
                    sourceSlideIndex);
                int afterCount = slides.Count;
                if (afterCount != beforeCount + 1)
                {
                    return Fail(
                        OfficeErrorCode.PARTIAL_APPLY,
                        "PowerPoint 插入源页后页数从 " + beforeCount.ToString(CultureInfo.InvariantCulture) +
                        " 变为 " + afterCount.ToString(CultureInfo.InvariantCulture) + "，未能确认只插入一页。");
                }
                int newIndex = insertAfter + 1;
                PowerPoint.Slide newSlide = null;
                try
                {
                    newSlide = slides[newIndex];
                    newIndex = newSlide.SlideIndex;
                }
                finally
                {
                    ComHelper.Release(newSlide);
                }

                JObject data = new JObject();
                data["sourceFile"] = sourcePath;
                data["sourceSlideIndex"] = sourceSlideIndex;
                data["insertAfter"] = insertAfter;
                data["newIndex"] = newIndex;
                data["slideCount"] = slides.Count;
                return Success(
                    "已从 " + Path.GetFileName(sourcePath) + " 复制第 " +
                    sourceSlideIndex.ToString(CultureInfo.InvariantCulture) + " 页到第 " +
                    newIndex.ToString(CultureInfo.InvariantCulture) + " 页。",
                    "write",
                    true,
                    "ppt:" + presentation.Name + "!Slide" + newIndex.ToString(CultureInfo.InvariantCulture),
                    data);
            }
            finally
            {
                ComHelper.Release(slides, presentation);
            }
        }
    }
}
