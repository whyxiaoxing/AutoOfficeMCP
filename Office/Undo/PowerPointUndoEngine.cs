using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Office.Undo
{
    /// <summary>
    /// PowerPoint 撤销节点：利用 PowerPoint 原生撤销栈。
    /// 每次写工具执行前调用 StartNewUndoEntry，把该工具的全部 COM 改动归并为一个原生撤销组；
    /// 节点保存执行前/执行后的演示文稿状态签名，回放前校验文档与当前状态，回放后再次校验，
    /// 避免用户已另行修改或切换演示文稿时误撤销。
    /// </summary>
    public static class PowerPointUndoEngine
    {
        public const string PptKind = "ppt_presentation_state";

        /// <summary>采集当前活动演示文稿的可比较状态；失败时抛出异常，由调用方收敛。</summary>
        public static JObject Capture(PowerPoint.Presentation presentation)
        {
            if (presentation == null)
            {
                throw new ArgumentNullException("presentation");
            }

            JObject state = new JObject();
            state["kind"] = PptKind;
            state["document"] = SafeString(delegate { return presentation.Name; });
            state["slideWidth"] = SafeFloat(delegate { return presentation.PageSetup.SlideWidth; });
            state["slideHeight"] = SafeFloat(delegate { return presentation.PageSetup.SlideHeight; });

            JArray slides = new JArray();
            PowerPoint.Slides slideCollection = null;
            try
            {
                slideCollection = presentation.Slides;
                int slideCount = slideCollection.Count;
                for (int index = 1; index <= slideCount; index++)
                {
                    PowerPoint.Slide slide = null;
                    try
                    {
                        slide = slideCollection[index];
                        slides.Add(CaptureSlide(slide, index));
                    }
                    finally
                    {
                        ComHelper.Release(slide);
                    }
                }
            }
            finally
            {
                ComHelper.Release(slideCollection);
            }

            state["slides"] = slides;
            state["slideCount"] = slides.Count;
            state["signature"] = ComputeSignature(state);
            return state;
        }

        /// <summary>纯逻辑构建可撤销节点；签名未变化或结构不完整时返回 null。</summary>
        public static SnapshotUndoNode BuildNode(JObject beforeState, JObject afterState, string summary)
        {
            if (beforeState == null || afterState == null)
            {
                return null;
            }
            if (!string.Equals(ReadString(beforeState, "kind"), PptKind, StringComparison.Ordinal) ||
                !string.Equals(ReadString(afterState, "kind"), PptKind, StringComparison.Ordinal))
            {
                return null;
            }

            string documentName = ReadString(beforeState, "document");
            string beforeSignature = ReadString(beforeState, "signature");
            string afterSignature = ReadString(afterState, "signature");
            if (documentName.Length == 0 || beforeSignature.Length == 0 || afterSignature.Length == 0)
            {
                return null;
            }
            if (string.Equals(beforeSignature, afterSignature, StringComparison.Ordinal))
            {
                return null;
            }
            if (!string.Equals(documentName, ReadString(afterState, "document"), StringComparison.Ordinal))
            {
                return null;
            }

            JObject snapshot = (JObject)beforeState.DeepClone();
            snapshot["afterSignature"] = afterSignature;
            return new SnapshotUndoNode(
                SnapshotUndoNode.NewId(),
                PptKind,
                "ppt",
                documentName,
                string.Empty,
                "presentation",
                string.IsNullOrWhiteSpace(summary) ? "PowerPoint 操作" : summary.Trim(),
                true,
                string.Empty,
                DateTime.UtcNow,
                snapshot);
        }

        /// <summary>回放前校验：文档必须是执行时的演示文稿，当前状态必须等于执行后状态。</summary>
        public static ReplayValidation Validate(
            JObject state,
            string activeDocumentName,
            string currentSignature)
        {
            ReplayValidation result = new ReplayValidation();
            if (state == null)
            {
                result.Reason = "PPT 撤销快照缺失，已拒绝回放。";
                return result;
            }
            if (!string.Equals(ReadString(state, "kind"), PptKind, StringComparison.Ordinal))
            {
                result.Reason = "快照不是 PowerPoint 演示文稿状态，已拒绝回放。";
                return result;
            }

            string documentName = ReadString(state, "document");
            string beforeSignature = ReadString(state, "signature");
            string afterSignature = ReadString(state, "afterSignature");
            if (documentName.Length == 0 || beforeSignature.Length == 0 || afterSignature.Length == 0)
            {
                result.Reason = "PPT 撤销快照缺少文档名或状态签名，已拒绝回放。";
                return result;
            }
            if (string.IsNullOrWhiteSpace(activeDocumentName))
            {
                result.Reason = "当前没有活动演示文稿，无法撤销。";
                return result;
            }
            if (!string.Equals(activeDocumentName, documentName, StringComparison.Ordinal))
            {
                result.Reason = "当前活动演示文稿不是本轮改动所在的「" + documentName + "」，已拒绝撤销。";
                return result;
            }

            result.DocumentName = documentName;
            result.BeforeSignature = beforeSignature;
            result.AfterSignature = afterSignature;
            if (string.Equals(currentSignature, beforeSignature, StringComparison.Ordinal))
            {
                result.AlreadyRestored = true;
                result.CanReplay = false;
                result.Reason = "当前演示文稿已经处于该操作执行前的状态，无需再次撤销。";
                return result;
            }
            if (!string.Equals(currentSignature, afterSignature, StringComparison.Ordinal))
            {
                result.Reason = "演示文稿在本轮操作后又被修改，原生撤销顺序已无法可靠回放；已拒绝撤销。";
                return result;
            }

            result.CanReplay = true;
            return result;
        }

        /// <summary>
        /// 执行 PowerPoint 原生 Undo 并回读校验。
        /// 返回 true 表示已恢复到执行前状态；alreadyRestored=true 表示本来就已经是执行前状态。
        /// </summary>
        public static bool Replay(
            PowerPoint.Application application,
            JObject state,
            out bool alreadyRestored,
            out string reason)
        {
            alreadyRestored = false;
            reason = string.Empty;
            ComHelper.EnsureSta("PowerPointUndoEngine.Replay");
            if (application == null)
            {
                reason = "PowerPoint 应用不可用，无法撤销。";
                return false;
            }

            PowerPoint.Presentation presentation = null;
            JObject currentState = null;
            try
            {
                presentation = application.ActivePresentation;
                if (presentation == null)
                {
                    reason = "当前没有活动演示文稿，无法撤销。";
                    return false;
                }
                currentState = Capture(presentation);
                ReplayValidation validation = Validate(
                    state,
                    currentState.Value<string>("document"),
                    currentState.Value<string>("signature"));
                if (validation.AlreadyRestored)
                {
                    alreadyRestored = true;
                    return true;
                }
                if (!validation.CanReplay)
                {
                    reason = validation.Reason;
                    return false;
                }

                dynamic app = application;
                dynamic commandBars = app.CommandBars;
                string beforeSignature = ReadString(state, "signature");
                string afterSignature = ReadString(state, "afterSignature");
                string previousSignature = afterSignature;
                int unchangedAttempts = 0;
                for (int attempt = 0; attempt < 12; attempt++)
                {
                    commandBars.ExecuteMso("Undo");
                    System.Threading.Thread.Sleep(120);
                    JObject restoredState = Capture(presentation);
                    string restoredSignature = ReadString(restoredState, "signature");
                    if (string.Equals(beforeSignature, restoredSignature, StringComparison.Ordinal))
                    {
                        return true;
                    }
                    if (string.Equals(restoredSignature, previousSignature, StringComparison.Ordinal))
                    {
                        unchangedAttempts++;
                        if (unchangedAttempts >= 2)
                        {
                            reason = "PowerPoint 原生撤销连续无效果，未恢复到执行前状态。";
                            return false;
                        }
                    }
                    else
                    {
                        unchangedAttempts = 0;
                    }
                    previousSignature = restoredSignature;
                }

                JObject finalState = Capture(presentation);
                string difference = DescribeDifference(state, finalState);
                Console.Error.WriteLine("[ppt-undo] 撤销后状态不一致：" + difference);
                reason = "PowerPoint 原生撤销已执行，但回读状态与执行前不一致（可能撤销组未覆盖全部改动）：" +
                         difference;
                return false;
            }
            catch (Exception ex)
            {
                reason = "PowerPoint 撤销回放失败：" + ex.Message;
                return false;
            }
            finally
            {
                ComHelper.Release(presentation);
            }
        }

        /// <summary>计算状态签名；忽略 signature 字段本身。</summary>
        public static string ComputeSignature(JObject state)
        {
            if (state == null)
            {
                return string.Empty;
            }
            JObject clone = (JObject)state.DeepClone();
            clone.Remove("signature");
            string text = clone.ToString(Formatting.None);
            using (SHA256 sha = SHA256.Create())
            {
                byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(text));
                StringBuilder builder = new StringBuilder(bytes.Length * 2);
                foreach (byte value in bytes)
                {
                    builder.Append(value.ToString("x2", CultureInfo.InvariantCulture));
                }
                return builder.ToString();
            }
        }

        private static JObject CaptureSlide(PowerPoint.Slide slide, int index)
        {
            JObject state = new JObject();
            state["index"] = index;
            state["layout"] = SafeString(delegate { return slide.CustomLayout.Name; });
            state["hidden"] = SafeBool(delegate
            {
                dynamic transition = slide.SlideShowTransition;
                return Convert.ToInt32((object)transition.Hidden) != 0;
            });
            state["transition"] = ReadTransitionState(slide);
            state["animation"] = ReadAnimationState(slide);
            state["notes"] = ReadNotesText(slide);

            JArray shapes = new JArray();
            PowerPoint.Shapes shapeCollection = null;
            try
            {
                shapeCollection = slide.Shapes;
                int shapeCount = shapeCollection.Count;
                for (int shapeIndex = 1; shapeIndex <= shapeCount; shapeIndex++)
                {
                    PowerPoint.Shape shape = null;
                    try
                    {
                        shape = shapeCollection[shapeIndex];
                        shapes.Add(CaptureShape(shape));
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
            state["shapes"] = shapes;
            return state;
        }

        private static JObject CaptureShape(PowerPoint.Shape shape)
        {
            JObject state = new JObject();
            // shape.Id 在原生撤销后可能重新分配，不能进入状态签名；返回给用户的 shapeId 由工具结果单独提供。
            state["name"] = SafeString(delegate { return shape.Name; });
            state["type"] = SafeInt(delegate { return Convert.ToInt32((object)((dynamic)shape).Type); });
            state["left"] = SafeFloat(delegate { return shape.Left; });
            state["top"] = SafeFloat(delegate { return shape.Top; });
            state["width"] = SafeFloat(delegate { return shape.Width; });
            state["height"] = SafeFloat(delegate { return shape.Height; });
            state["rotation"] = SafeFloat(delegate { return shape.Rotation; });

            bool hasText = SafeBool(delegate { return Convert.ToInt32((object)((dynamic)shape).HasTextFrame) != 0; });
            state["hasText"] = hasText;
            if (hasText)
            {
                PowerPoint.TextFrame textFrame = null;
                PowerPoint.TextRange textRange = null;
                try
                {
                    textFrame = shape.TextFrame;
                    textRange = textFrame.TextRange;
                    state["text"] = textRange.Text ?? string.Empty;
                    state["font"] = ReadFontState(textRange);
                    state["paragraph"] = ReadParagraphState(textRange);
                }
                catch (Exception)
                {
                }
                finally
                {
                    ComHelper.Release(textRange, textFrame);
                }
            }

            state["fillColor"] = SafeInt(delegate
            {
                dynamic dynamicShape = shape;
                dynamic fill = dynamicShape.Fill;
                if (Convert.ToInt32((object)fill.Visible) == 0)
                {
                    return -1;
                }
                return Convert.ToInt32((object)fill.ForeColor.RGB);
            });
            state["lineColor"] = SafeInt(delegate
            {
                dynamic dynamicShape = shape;
                dynamic line = dynamicShape.Line;
                return Convert.ToInt32((object)line.ForeColor.RGB);
            });
            state["lineWeight"] = SafeFloat(delegate
            {
                dynamic dynamicShape = shape;
                dynamic line = dynamicShape.Line;
                return Convert.ToSingle((object)line.Weight, CultureInfo.InvariantCulture);
            });

            JArray table = ReadTable(shape);
            if (table != null)
            {
                state["table"] = table;
            }
            return state;
        }

        private static JObject ReadFontState(PowerPoint.TextRange textRange)
        {
            JObject state = new JObject();
            dynamic font = textRange.Font;
            state["name"] = SafeString(delegate { return (string)font.Name; });
            state["size"] = SafeFloat(delegate { return Convert.ToSingle((object)font.Size, CultureInfo.InvariantCulture); });
            state["bold"] = SafeInt(delegate { return Convert.ToInt32((object)font.Bold); });
            state["italic"] = SafeInt(delegate { return Convert.ToInt32((object)font.Italic); });
            state["underline"] = SafeInt(delegate { return Convert.ToInt32((object)font.Underline); });
            state["color"] = SafeInt(delegate
            {
                dynamic color = font.Color;
                return Convert.ToInt32((object)color.RGB);
            });
            return state;
        }

        private static JObject ReadParagraphState(PowerPoint.TextRange textRange)
        {
            JObject state = new JObject();
            dynamic paragraphFormat = textRange.ParagraphFormat;
            dynamic dynamicRange = textRange;
            state["alignment"] = SafeInt(delegate { return Convert.ToInt32((object)paragraphFormat.Alignment); });
            state["level"] = SafeInt(delegate { return Convert.ToInt32((object)dynamicRange.IndentLevel); });
            state["spaceBefore"] = SafeFloat(delegate
            {
                return Convert.ToSingle((object)paragraphFormat.SpaceBefore, CultureInfo.InvariantCulture);
            });
            state["spaceAfter"] = SafeFloat(delegate
            {
                return Convert.ToSingle((object)paragraphFormat.SpaceAfter, CultureInfo.InvariantCulture);
            });
            state["lineRuleWithin"] = SafeInt(delegate
            {
                return Convert.ToInt32((object)paragraphFormat.LineRuleWithin);
            });
            state["spaceWithin"] = SafeFloat(delegate
            {
                return Convert.ToSingle((object)paragraphFormat.SpaceWithin, CultureInfo.InvariantCulture);
            });
            return state;
        }

        private static JObject ReadTransitionState(PowerPoint.Slide slide)
        {
            JObject state = new JObject();
            try
            {
                dynamic transition = slide.SlideShowTransition;
                state["entryEffect"] = SafeInt(delegate { return Convert.ToInt32((object)transition.EntryEffect); });
                state["duration"] = SafeFloat(delegate
                {
                    return Convert.ToSingle((object)transition.Duration, CultureInfo.InvariantCulture);
                });
                state["advanceOnClick"] = SafeInt(delegate
                {
                    return Convert.ToInt32((object)transition.AdvanceOnClick);
                });
            }
            catch (Exception)
            {
            }
            return state;
        }

        private static JObject ReadAnimationState(PowerPoint.Slide slide)
        {
            JObject state = new JObject();
            PowerPoint.TimeLine timeline = null;
            PowerPoint.Sequence sequence = null;
            try
            {
                timeline = slide.TimeLine;
                sequence = timeline.MainSequence;
                int count = sequence.Count;
                state["count"] = count;
                JArray effects = new JArray();
                for (int index = 1; index <= count; index++)
                {
                    PowerPoint.Effect effect = null;
                    try
                    {
                        effect = sequence[index];
                        effects.Add(SafeInt(delegate
                        {
                            dynamic dynamicEffect = effect;
                            return Convert.ToInt32((object)dynamicEffect.EffectType);
                        }));
                    }
                    finally
                    {
                        ComHelper.Release(effect);
                    }
                }
                state["effects"] = effects;
            }
            catch (Exception)
            {
            }
            finally
            {
                ComHelper.Release(sequence, timeline);
            }
            return state;
        }

        private static JArray ReadTable(PowerPoint.Shape shape)
        {
            try
            {
                dynamic dynamicShape = shape;
                if (Convert.ToInt32((object)dynamicShape.HasTable) == 0)
                {
                    return null;
                }
                PowerPoint.Table table = shape.Table;
                try
                {
                    int rows = table.Rows.Count;
                    int columns = table.Columns.Count;
                    JArray result = new JArray();
                    for (int row = 1; row <= rows; row++)
                    {
                        JArray rowArray = new JArray();
                        for (int column = 1; column <= columns; column++)
                        {
                            PowerPoint.Cell cell = null;
                            PowerPoint.Shape cellShape = null;
                            try
                            {
                                cell = table.Cell(row, column);
                                cellShape = cell.Shape;
                                rowArray.Add(cellShape.TextFrame.TextRange.Text ?? string.Empty);
                            }
                            finally
                            {
                                ComHelper.Release(cellShape, cell);
                            }
                        }
                        result.Add(rowArray);
                    }
                    return result;
                }
                finally
                {
                    ComHelper.Release(table);
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string ReadNotesText(PowerPoint.Slide slide)
        {
            PowerPoint.SlideRange notesPage = null;
            try
            {
                notesPage = slide.NotesPage;
                dynamic shapes = notesPage.Shapes;
                int count = Convert.ToInt32((object)shapes.Count);
                for (int index = 1; index <= count; index++)
                {
                    PowerPoint.Shape shape = null;
                    try
                    {
                        shape = shapes[index];
                        dynamic dynamicShape = shape;
                        if (Convert.ToInt32((object)dynamicShape.Type) == 14 &&
                            Convert.ToInt32((object)shape.PlaceholderFormat.Type) == 2)
                        {
                            return shape.TextFrame.TextRange.Text ?? string.Empty;
                        }
                    }
                    finally
                    {
                        ComHelper.Release(shape);
                    }
                }
                return string.Empty;
            }
            catch (Exception)
            {
                return string.Empty;
            }
            finally
            {
                ComHelper.Release(notesPage);
            }
        }

        /// <summary>递归定位第一个不同字段，供撤销失败时诊断；只输出路径与截断后的值。</summary>
        private static string DescribeDifference(JToken expected, JToken actual)
        {
            return DescribeDifference(expected, actual, "$");
        }

        private static string DescribeDifference(JToken expected, JToken actual, string path)
        {
            if (expected == null && actual == null)
            {
                return "无差异";
            }
            if (expected == null || actual == null)
            {
                return path + " 一侧为空";
            }
            if (expected.Type != actual.Type)
            {
                return path + " 类型不同（" + expected.Type + " / " + actual.Type + "）";
            }
            if (expected is JObject && actual is JObject)
            {
                JObject expectedObject = (JObject)expected;
                JObject actualObject = (JObject)actual;
                foreach (JProperty property in expectedObject.Properties())
                {
                    if (actualObject[property.Name] == null)
                    {
                        return path + "." + property.Name + " 缺失";
                    }
                    string difference = DescribeDifference(
                        property.Value, actualObject[property.Name], path + "." + property.Name);
                    if (!string.Equals(difference, "无差异", StringComparison.Ordinal))
                    {
                        return difference;
                    }
                }
                foreach (JProperty property in actualObject.Properties())
                {
                    if (expectedObject[property.Name] == null)
                    {
                        return path + "." + property.Name + " 多出";
                    }
                }
                return "无差异";
            }
            if (expected is JArray && actual is JArray)
            {
                JArray expectedArray = (JArray)expected;
                JArray actualArray = (JArray)actual;
                if (expectedArray.Count != actualArray.Count)
                {
                    return path + " 数组长度不同（" + expectedArray.Count + " / " + actualArray.Count + "）";
                }
                for (int index = 0; index < expectedArray.Count; index++)
                {
                    string difference = DescribeDifference(
                        expectedArray[index], actualArray[index], path + "[" + index + "]");
                    if (!string.Equals(difference, "无差异", StringComparison.Ordinal))
                    {
                        return difference;
                    }
                }
                return "无差异";
            }
            if (string.Equals(expected.ToString(), actual.ToString(), StringComparison.Ordinal))
            {
                return "无差异";
            }
            string expectedText = expected.ToString();
            string actualText = actual.ToString();
            if (expectedText.Length > 160) { expectedText = expectedText.Substring(0, 160) + "…"; }
            if (actualText.Length > 160) { actualText = actualText.Substring(0, 160) + "…"; }
            return path + " 期望「" + expectedText + "」实际「" + actualText + "」";
        }

        private static string ReadString(JObject root, string propertyName)
        {
            if (root == null)
            {
                return string.Empty;
            }
            JToken token = root[propertyName];
            if (token == null || token.Type == JTokenType.Null)
            {
                return string.Empty;
            }
            return token.Type == JTokenType.String
                ? (token.Value<string>() ?? string.Empty)
                : token.ToString();
        }

        private static string SafeString(Func<string> reader)
        {
            try
            {
                return reader() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static float SafeFloat(Func<float> reader)
        {
            try
            {
                float value = reader();
                if (float.IsNaN(value) || float.IsInfinity(value))
                {
                    return 0f;
                }
                return (float)Math.Round(value, 3);
            }
            catch
            {
                return 0f;
            }
        }

        private static int SafeInt(Func<int> reader)
        {
            try
            {
                return reader();
            }
            catch
            {
                return 0;
            }
        }

        private static bool SafeBool(Func<bool> reader)
        {
            try
            {
                return reader();
            }
            catch
            {
                return false;
            }
        }
    }
}
