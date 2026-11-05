using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Office.Undo;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    /// <summary>
    /// Word 写操作统一外壳：写前/写后指纹 + 目标区段快照 + 自定义撤销记录 + 真实 Observation。
    /// 成功必须由 Word 回读或明确副作用支撑；失败但文档已变化时只能上报 PARTIAL_APPLY。
    /// </summary>
    internal static class WordWriteObservation
    {
        public static OperationResult Run(
            Word.Application application,
            Word.Document document,
            string toolName,
            WordWriteTarget target,
            Func<WordWriteResult> action)
        {
            return Run(application, document, toolName, target, action, false);
        }

        /// <summary>
        /// useNativeUndo=true 时不开启 Word 自定义 UndoRecord，改用 Word 原生撤销。
        /// 查找替换类工具必须走此路径：Office 2024 上自定义 UndoRecord 期间执行
        /// Find.Execute(wdReplaceAll) 会导致 WINWORD 崩溃（0x800706BE/0x800706BA）。
        /// </summary>
        public static OperationResult Run(
            Word.Application application,
            Word.Document document,
            string toolName,
            WordWriteTarget target,
            Func<WordWriteResult> action,
            bool useNativeUndo)
        {
            WordWriteTarget effectiveTarget = target ?? WordWriteTargets.None("Word 文档", "Word:Document");
            JObject beforeState = null;
            JObject afterState = null;
            SnapshotCapture beforeSnapshot = null;
            SnapshotCapture afterSnapshot = null;
            Word.Range beforeRange = null;
            Word.Range afterRange = null;
            WordWriteSession session = null;

            try
            {
                beforeState = WordDocumentState.Capture(application, document);
                if (effectiveTarget.SnapshotCapable && effectiveTarget.Start >= 0)
                {
                    beforeRange = WordComUtil.CreateRange(document, effectiveTarget.Start, effectiveTarget.End);
                    beforeSnapshot = SnapshotUndoEngine.CaptureWord(
                        document, beforeRange, effectiveTarget.Label);
                }

                session = WordWriteSession.Start(application, toolName, !useNativeUndo);

                WordWriteResult result = null;
                Exception actionException = null;
                try
                {
                    result = action == null ? null : action();
                }
                catch (Exception ex)
                {
                    actionException = ex;
                }

                int afterStart = effectiveTarget.Start;
                int afterEnd = effectiveTarget.End;
                if (result != null)
                {
                    if (result.AfterStart.HasValue)
                    {
                        afterStart = result.AfterStart.Value;
                    }
                    if (result.AfterEnd.HasValue)
                    {
                        afterEnd = result.AfterEnd.Value;
                    }
                }

                afterState = WordDocumentState.Capture(application, document);
                if (effectiveTarget.SnapshotCapable && afterStart >= 0 && afterEnd >= afterStart)
                {
                    try
                    {
                        afterRange = WordComUtil.CreateRange(document, afterStart, afterEnd);
                        afterSnapshot = SnapshotUndoEngine.CaptureWord(
                            document, afterRange, effectiveTarget.Label + "（写后）");
                    }
                    catch (Exception ex)
                    {
                        afterSnapshot = SnapshotCapture.Failed(
                            "写入后目标区段快照采集失败：" +
                            OfficeErrorMapper.DescribeException(ex, OfficeApp.Word));
                    }
                }

                bool changed = DetermineChanged(result, beforeState, afterState, beforeSnapshot, afterSnapshot);
                if (result == null && actionException == null)
                {
                    actionException = new InvalidOperationException("Word 写处理器未返回结果（执行器内部缺陷）。");
                }

                if (actionException != null)
                {
                    result = WordWriteResult.Fail(
                        OfficeErrorMapper.MapException(actionException, OfficeApp.Word),
                        OfficeErrorMapper.DescribeException(actionException, OfficeApp.Word),
                        null,
                        changed);
                }

                bool resultStructural = effectiveTarget.Structural || (result != null && result.Structural);
                string undoUnitId;
                string undoNote = BuildUndoNode(
                    beforeSnapshot,
                    afterSnapshot,
                    effectiveTarget,
                    resultStructural,
                    changed,
                    toolName,
                    out undoUnitId);

                Observation observation = Observation.Create(
                    "write",
                    BuildObservationSummary(effectiveTarget, result, changed),
                    changed);
                AddTargetRefs(observation, effectiveTarget, result);
                if (!string.IsNullOrWhiteSpace(undoUnitId))
                {
                    // undoUnits 只放真实可回放的 T02 节点 ID；原生 UndoRecord 不伪装成 T02 节点。
                    observation.AddUndoUnit(undoUnitId);
                }

                JObject data = result != null && result.Data != null
                    ? (JObject)result.Data.DeepClone()
                    : new JObject();
                data["snapshotBefore"] = beforeState == null ? new JObject() : beforeState.DeepClone();
                data["snapshotAfter"] = afterState == null ? new JObject() : afterState.DeepClone();
                data["snapshotDiff"] = WordDocumentState.BuildDiff(beforeState, afterState);
                data["documentTextHashBefore"] = beforeState == null
                    ? string.Empty : (beforeState.Value<string>("documentTextHash") ?? string.Empty);
                data["documentTextHashAfter"] = afterState == null
                    ? string.Empty : (afterState.Value<string>("documentTextHash") ?? string.Empty);
                data["paragraphCountBefore"] = beforeState == null
                    ? 0 : (beforeState.Value<int?>("paragraphCount") ?? 0);
                data["paragraphCountAfter"] = afterState == null
                    ? 0 : (afterState.Value<int?>("paragraphCount") ?? 0);
                data["selectionStartBefore"] = beforeState == null
                    ? 0 : (beforeState.Value<int?>("selectionStart") ?? 0);
                data["selectionStartAfter"] = afterState == null
                    ? 0 : (afterState.Value<int?>("selectionStart") ?? 0);
                data["selectionEndBefore"] = beforeState == null
                    ? 0 : (beforeState.Value<int?>("selectionEnd") ?? 0);
                data["selectionEndAfter"] = afterState == null
                    ? 0 : (afterState.Value<int?>("selectionEnd") ?? 0);
                data["undoNote"] = undoNote ?? string.Empty;
                data["nativeUndoMode"] = session.NativeUndoEnabled ? "word-native" : "custom-record";
                data["nativeUndoRecord"] = session.NativeUndoEnabled
                    ? "未开启自定义记录（该工具使用 Word 原生撤销，规避 Office 2024 UndoRecord + Find.Execute 崩溃）"
                    : (session.UndoRecordStarted
                        ? session.CustomRecordName
                        : "未开启（Word 拒绝或已有其他自定义记录）");
                data["targetRefs"] = BuildRefArray(observation.TargetRefs);

                if (result == null)
                {
                    result = WordWriteResult.Fail(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "Word 写处理器未返回结果。",
                        null,
                        changed);
                }

                if (result.Success)
                {
                    string successMessage = string.IsNullOrWhiteSpace(result.Message)
                        ? "Word 操作完成。"
                        : result.Message;
                    if (!changed)
                    {
                        successMessage = successMessage +
                                         "（写后观察未检测到文档内容或目标格式变化；目标可能已经是该状态）";
                    }
                    return OperationResult.Succeed(successMessage, data, observation);
                }

                string failureMessage = string.IsNullOrWhiteSpace(result.Message)
                    ? OfficeErrorMapper.GetUserMessage(result.ErrorCode)
                    : result.Message;
                OfficeErrorCode failureCode = result.ErrorCode == OfficeErrorCode.INTERNAL_ERROR &&
                                              actionException == null
                    ? OfficeErrorCode.INTERNAL_ERROR
                    : result.ErrorCode;

                if (changed)
                {
                    data["partialApply"] = new JObject
                    {
                        { "appliedObserved", true },
                        { "failure", failureMessage },
                        { "errorCode", OfficeErrorMapper.ToCode(failureCode) }
                    };
                    string rollback = "已检测到文档实际变化，未自动回放。原始失败：" + failureMessage +
                                      "。撤销信息：" + undoNote;
                    return OperationResult.PartialApplied(1, rollback, observation, data);
                }

                string rollbackNote = string.IsNullOrWhiteSpace(undoNote) ? null : undoNote;
                return OperationResult.Failed(failureCode, failureMessage, observation, data, rollbackNote);
            }
            catch (Exception ex)
            {
                return OfficeErrorMapper.ToOperationResult(ex, OfficeApp.Word);
            }
            finally
            {
                if (session != null)
                {
                    session.Dispose();
                }
                ComHelper.Release(afterRange, beforeRange);
            }
        }

        private static bool DetermineChanged(
            WordWriteResult result,
            JObject beforeState,
            JObject afterState,
            SnapshotCapture beforeSnapshot,
            SnapshotCapture afterSnapshot)
        {
            bool observed = WordDocumentState.HasChanged(beforeState, afterState) ||
                            SnapshotSignaturesDiffer(beforeSnapshot, afterSnapshot);
            if (result != null && result.Changed.HasValue)
            {
                if (result.Changed.Value)
                {
                    return true;
                }

                // 处理器明确回读为“未改变”时，只接受文档内容/结构或目标快照的变化，
                // 排除 Find/Select 造成的纯选区位移（否则未命中的替换会假报 changed=true）。
                return WordDocumentState.HasContentChanged(beforeState, afterState) ||
                       SnapshotSignaturesDiffer(beforeSnapshot, afterSnapshot);
            }
            return observed;
        }

        private static bool SnapshotSignaturesDiffer(SnapshotCapture before, SnapshotCapture after)
        {
            if (before == null || after == null || !before.Success || !after.Success ||
                before.Snapshot == null || after.Snapshot == null)
            {
                return false;
            }
            string beforeSignature = before.Snapshot.Value<string>("signature");
            string afterSignature = after.Snapshot.Value<string>("signature");
            return !string.IsNullOrWhiteSpace(beforeSignature) &&
                   !string.IsNullOrWhiteSpace(afterSignature) &&
                   !string.Equals(beforeSignature, afterSignature, StringComparison.Ordinal);
        }

        private static string BuildUndoNode(
            SnapshotCapture beforeSnapshot,
            SnapshotCapture afterSnapshot,
            WordWriteTarget target,
            bool structural,
            bool changed,
            string toolName,
            out string undoUnitId)
        {
            undoUnitId = string.Empty;
            if (!target.SnapshotCapable)
            {
                return "该工具修改的是结构/段落格式/域/文档属性等 T02 快照引擎未覆盖的状态，未登记可回放撤销节点。" +
                       "本轮已尝试开启 Word 自定义撤销记录（AutoOfficeMCP/" + toolName + "）。";
            }
            if (structural)
            {
                return "该操作会改变文档结构（表格/图片/域/段落或对象布局），T02 Word 快照只覆盖文本与字体，" +
                       "为避免假撤销未登记可回放节点。本轮已尝试开启 Word 自定义撤销记录。";
            }
            if (!changed)
            {
                return "写后未检测到目标区段内容或字体变化，无需登记撤销节点。";
            }
            if (beforeSnapshot == null || !beforeSnapshot.Success)
            {
                return "执行前快照采集失败，未登记可回放撤销节点：" +
                       (beforeSnapshot == null ? "快照缺失" : beforeSnapshot.Reason);
            }
            if (afterSnapshot == null || !afterSnapshot.Success)
            {
                return "执行后快照采集失败，未登记可回放撤销节点：" +
                       (afterSnapshot == null ? "快照缺失" : afterSnapshot.Reason);
            }

            SnapshotUndoNode node = SnapshotUndoEngine.BuildNode(
                beforeSnapshot,
                afterSnapshot,
                true,
                toolName + "：" + target.Label);
            if (node == null)
            {
                return "T02 快照引擎未返回撤销节点。";
            }
            if (!node.Restorable)
            {
                return "T02 快照引擎判定不可回放：" + node.Reason;
            }

            OfficeSession.UndoStack.Push(node);
            undoUnitId = node.Id;
            return "已登记 T02 可回放撤销节点 " + node.Id + "。";
        }

        private static string BuildObservationSummary(
            WordWriteTarget target,
            WordWriteResult result,
            bool changed)
        {
            string prefix = result != null && result.Success ? string.Empty : "未完成：";
            string message = result == null || string.IsNullOrWhiteSpace(result.Message)
                ? "Word 写操作"
                : result.Message;
            return prefix + message + "（目标：" + target.Label + "；changed=" +
                   (changed ? "true" : "false") + "）";
        }

        private static void AddTargetRefs(
            Observation observation,
            WordWriteTarget target,
            WordWriteResult result)
        {
            HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
            if (target != null && target.TargetRefs != null)
            {
                foreach (string item in target.TargetRefs)
                {
                    if (seen.Add(item))
                    {
                        observation.AddTargetRef(item);
                    }
                }
            }
            if (result != null && result.TargetRefs != null)
            {
                foreach (string item in result.TargetRefs)
                {
                    if (seen.Add(item))
                    {
                        observation.AddTargetRef(item);
                    }
                }
            }
        }

        private static JArray BuildRefArray(List<string> refs)
        {
            JArray array = new JArray();
            if (refs == null)
            {
                return array;
            }
            foreach (string item in refs)
            {
                if (!string.IsNullOrWhiteSpace(item))
                {
                    array.Add(item);
                }
            }
            return array;
        }
    }
}
