using System;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Office;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointe
{
    public sealed partial class PowerPointToolExecutor
    {
        private static OperationResult ExecuteListAnimations(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.TimeLine timeline = null;
            PowerPoint.Sequence sequence = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                timeline = slide.TimeLine;
                sequence = timeline.MainSequence;

                JArray animations = new JArray();
                int count = sequence.Count;
                for (int index = 1; index <= count; index++)
                {
                    PowerPoint.Effect effect = null;
                    PowerPoint.Shape shape = null;
                    try
                    {
                        effect = sequence[index];
                        JObject item = new JObject();
                        item["order"] = index;
                        try
                        {
                            shape = effect.Shape;
                            item["shapeId"] = PowerPointComUtil.SafeInt(delegate
                            {
                                dynamic dynamicShape = shape;
                                return Convert.ToInt32((object)dynamicShape.Id);
                            });
                            item["shapeName"] = PowerPointComUtil.SafeString(delegate
                            {
                                return shape.Name;
                            });
                        }
                        catch (Exception)
                        {
                            item["shapeId"] = 0;
                            item["shapeName"] = string.Empty;
                        }

                        int effectType = PowerPointComUtil.SafeInt(delegate
                        {
                            dynamic dynamicEffect = effect;
                            return Convert.ToInt32((object)dynamicEffect.EffectType);
                        });
                        item["effectType"] = effectType;
                        item["effect"] = PowerPointComUtil.MapAnimationEffectName(effectType);

                        PowerPoint.Timing timing = null;
                        try
                        {
                            timing = effect.Timing;
                            float duration = PowerPointComUtil.SafeFloat(delegate
                            {
                                return Convert.ToSingle((object)timing.Duration);
                            });
                            float delay = PowerPointComUtil.SafeFloat(delegate
                            {
                                return Convert.ToSingle((object)timing.TriggerDelayTime);
                            });
                            int trigger = PowerPointComUtil.SafeInt(delegate
                            {
                                return Convert.ToInt32((object)timing.TriggerType);
                            });
                            item["duration"] = duration;
                            item["delay"] = delay;
                            item["trigger"] = PowerPointComUtil.MapTriggerName(trigger);
                            item["startMode"] = PowerPointComUtil.MapTriggerName(trigger);
                        }
                        finally
                        {
                            ComHelper.Release(timing);
                        }
                        animations.Add(item);
                    }
                    finally
                    {
                        ComHelper.Release(shape, effect);
                    }
                }

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["animationCount"] = animations.Count;
                data["animations"] = animations;
                return Success(
                    "第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片共有 " + animations.Count.ToString(CultureInfo.InvariantCulture) + " 个动画。",
                    "read",
                    false,
                    SlideRef(presentation, slide),
                    data);
            }
            finally
            {
                ComHelper.Release(sequence, timeline, slide, presentation);
            }
        }

        private static OperationResult ExecuteDeleteAnimation(
            PowerPoint.Application application,
            JObject arguments)
        {
            if (!PowerPointArgumentReader.GetBool(arguments, false, "confirm"))
            {
                return OperationResult.Cancelled(
                    "ppt_delete_animation 需要显式 confirm=true 才会删除动画；本次未修改演示文稿。");
            }

            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.TimeLine timeline = null;
            PowerPoint.Sequence sequence = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                int animationIndex = PowerPointArgumentReader.GetInt(arguments, -1, "animationIndex", "order");
                int shapeId = PowerPointArgumentReader.GetInt(arguments, -1, "shapeId");
                if (animationIndex < 1 && shapeId < 0)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "需要提供 animationIndex（1 基动画序号）或 shapeId。");
                }

                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                timeline = slide.TimeLine;
                sequence = timeline.MainSequence;
                int before = sequence.Count;
                int deleted = 0;
                if (animationIndex >= 1)
                {
                    if (animationIndex > sequence.Count)
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "animationIndex 超出范围：当前共 " + sequence.Count.ToString(CultureInfo.InvariantCulture) +
                            " 个动画。");
                    }
                    PowerPoint.Effect effect = null;
                    try
                    {
                        effect = sequence[animationIndex];
                        effect.Delete();
                        deleted = 1;
                    }
                    finally
                    {
                        ComHelper.Release(effect);
                    }
                }
                else
                {
                    for (int index = sequence.Count; index >= 1; index--)
                    {
                        PowerPoint.Effect effect = null;
                        PowerPoint.Shape shape = null;
                        try
                        {
                            effect = sequence[index];
                            try
                            {
                                shape = effect.Shape;
                                dynamic dynamicShape = shape;
                                if (Convert.ToInt32((object)dynamicShape.Id) == shapeId)
                                {
                                    effect.Delete();
                                    deleted++;
                                }
                            }
                            catch (Exception)
                            {
                                // 无法读取目标形状的动画跳过。
                            }
                        }
                        finally
                        {
                            ComHelper.Release(shape, effect);
                        }
                    }
                }

                int remaining = sequence.Count;
                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["deletedCount"] = deleted;
                data["beforeCount"] = before;
                data["remainingCount"] = remaining;
                data["animationIndex"] = animationIndex;
                data["shapeId"] = shapeId;
                if (deleted == 0)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "没有找到可删除的动画（shapeId=" + shapeId.ToString(CultureInfo.InvariantCulture) + "）。");
                }
                return Success(
                    "已删除第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片上的 " + deleted.ToString(CultureInfo.InvariantCulture) + " 个动画。",
                    "write",
                    true,
                    SlideRef(presentation, slide),
                    data);
            }
            finally
            {
                ComHelper.Release(sequence, timeline, slide, presentation);
            }
        }

        private static OperationResult ExecuteGetTransition(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.SlideShowTransition transition = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                transition = slide.SlideShowTransition;
                dynamic dynamicTransition = transition;
                int entryEffect = PowerPointComUtil.SafeInt(delegate
                {
                    return Convert.ToInt32((object)transition.EntryEffect);
                });
                float duration = PowerPointComUtil.SafeFloat(delegate
                {
                    return Convert.ToSingle((object)transition.Duration);
                });
                bool advanceOnClick = PowerPointComUtil.SafeBool(delegate
                {
                    return Convert.ToInt32((object)dynamicTransition.AdvanceOnClick) != 0;
                });
                float advanceTime = PowerPointComUtil.SafeFloat(delegate
                {
                    return Convert.ToSingle((object)dynamicTransition.AdvanceTime);
                });
                bool hidden = PowerPointComUtil.SafeBool(delegate
                {
                    return Convert.ToInt32((object)dynamicTransition.Hidden) != 0;
                });

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["entryEffect"] = entryEffect;
                data["entryEffectName"] = PowerPointComUtil.MapTransitionEffectName(entryEffect);
                data["duration"] = duration;
                data["advanceOnClick"] = advanceOnClick;
                data["advanceTime"] = advanceTime;
                data["hidden"] = hidden;
                return Success(
                    "已读取第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片的切换效果。",
                    "read",
                    false,
                    SlideRef(presentation, slide),
                    data);
            }
            finally
            {
                ComHelper.Release(transition, slide, presentation);
            }
        }
    }
}
