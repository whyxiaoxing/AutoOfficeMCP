using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Office.Undo
{
    /// <summary>纯逻辑自检报告（不依赖测试框架、不触碰 Office/COM）。</summary>
    public sealed class SelfTestReport
    {
        public int PassedCount { get; private set; }

        public int FailedCount { get; private set; }

        public IList<string> Failures { get; private set; }

        public bool AllPassed
        {
            get { return FailedCount == 0; }
        }

        public string Summary
        {
            get
            {
                return "SnapshotUndoEngine.SelfTest: " + PassedCount.ToString() + " passed, " +
                       FailedCount.ToString() + " failed";
            }
        }

        internal SelfTestReport()
        {
            Failures = new List<string>();
        }

        internal void Pass()
        {
            PassedCount++;
        }

        internal void Fail(string name, string detail)
        {
            FailedCount++;
            Failures.Add(name + "：" + (string.IsNullOrWhiteSpace(detail) ? "断言失败" : detail));
        }
    }

    /// <summary>
    /// SnapshotUndoEngine 的纯逻辑自检：序列化 / 签名 / 形状比对 / 五道防御链。
    /// 由 T07 或人工通过反射调用 Run()；本任务不在 Main 中自动调用。
    /// </summary>
    public static class SnapshotUndoEngineSelfTest
    {
        public static SelfTestReport Run()
        {
            SelfTestReport report = new SelfTestReport();
            CheckSerializationAndSignature(report);
            CheckShapeComparison(report);
            CheckBuildNode(report);
            CheckExcelReplayDefenseChain(report);
            CheckWordReplayDefenseChain(report);
            CheckFormulaDefense(report);
            CheckUndoStack(report);
            return report;
        }

        private static void CheckSerializationAndSignature(SelfTestReport report)
        {
            JArray values = Matrix(Row(1, 2), Row(3, 4));
            JArray formulas = Matrix(Row("=A1", "文本"), Row(string.Empty, string.Empty));
            JArray formats = Matrix(Row("General", "@"), Row("General", "General"));

            string first = SnapshotUndoEngine.ComputeSignature(values, formulas, formats);
            string second = SnapshotUndoEngine.ComputeSignature(values, formulas, formats);
            Check(report, "签名确定性", string.Equals(first, second, StringComparison.Ordinal),
                "相同矩阵两次签名不一致");
            Check(report, "签名长度", first.Length == 64 && IsHex(first),
                "SHA256 应为 64 位十六进制，实际：" + first);

            JArray changedValues = Matrix(Row(1, 2), Row(3, 5));
            string changed = SnapshotUndoEngine.ComputeSignature(changedValues, formulas, formats);
            Check(report, "签名敏感性", !string.Equals(first, changed, StringComparison.Ordinal),
                "Value2 变化后签名未变化");

            JObject container = new JObject();
            container["values"] = values;
            container["formulas"] = formulas;
            container["numberFormats"] = formats;
            JObject roundTrip = JObject.Parse(container.ToString(Formatting.None));
            string roundTripSignature = SnapshotUndoEngine.ComputeSignature(
                roundTrip["values"] as JArray,
                roundTrip["formulas"] as JArray,
                roundTrip["numberFormats"] as JArray);
            Check(report, "JArray 序列化往返", string.Equals(first, roundTripSignature, StringComparison.Ordinal),
                "JSON 序列化往返后签名不一致");

            JArray oneByOne = Matrix(Row("1-1"));
            JObject wordFont = SampleFont(false);
            string wordFirst = SnapshotUndoEngine.ComputeWordSignature("1-1", wordFont);
            JObject wordFontRoundTrip = JObject.Parse(wordFont.ToString(Formatting.None));
            string wordSecond = SnapshotUndoEngine.ComputeWordSignature("1-1", wordFontRoundTrip);
            Check(report, "Word 签名确定性", string.Equals(wordFirst, wordSecond, StringComparison.Ordinal),
                "Word 文本/字体签名往返后不一致");
            Check(report, "矩阵形状工具", SnapshotUndoEngine.MatrixHasShape(oneByOne, 1, 1) &&
                                         !SnapshotUndoEngine.MatrixHasShape(oneByOne, 2, 1),
                "MatrixHasShape 判断错误");
        }

        private static void CheckShapeComparison(SelfTestReport report)
        {
            JObject before = BuildExcelSnapshot("Book1.xlsx", "Sheet1", "A1:B2",
                Matrix(Row(1, 2), Row(3, 4)),
                Matrix(Row(string.Empty, string.Empty), Row(string.Empty, string.Empty)),
                Matrix(Row("General", "General"), Row("General", "General")));
            JObject after = (JObject)before.DeepClone();
            after["signature"] = SnapshotUndoEngine.ComputeSignature(
                Matrix(Row(1, 2), Row(3, 5)),
                after["formulas"] as JArray,
                after["numberFormats"] as JArray);

            string reason;
            Check(report, "形状相同", SnapshotUndoEngine.ShapesMatch(before, after, out reason),
                "相同结构被判为不一致：" + reason);

            JObject otherWorkbook = (JObject)after.DeepClone();
            otherWorkbook["workbook"] = "Book2.xlsx";
            Check(report, "工作簿名不一致", !SnapshotUndoEngine.ShapesMatch(before, otherWorkbook, out reason) &&
                                             Contains(reason, "工作簿"),
                "工作簿名变化未被识别：" + reason);

            JObject otherSheet = (JObject)after.DeepClone();
            otherSheet["sheet"] = "Sheet2";
            Check(report, "工作表名不一致", !SnapshotUndoEngine.ShapesMatch(before, otherSheet, out reason) &&
                                             Contains(reason, "工作表"),
                "工作表名变化未被识别：" + reason);

            JObject otherShape = (JObject)after.DeepClone();
            otherShape["rows"] = 3;
            Check(report, "行列数不一致", !SnapshotUndoEngine.ShapesMatch(before, otherShape, out reason) &&
                                           Contains(reason, "行列数"),
                "行列数变化未被识别：" + reason);

            JObject otherKind = (JObject)after.DeepClone();
            otherKind["kind"] = SnapshotUndoEngine.WordKind;
            Check(report, "快照类型不一致", !SnapshotUndoEngine.ShapesMatch(before, otherKind, out reason) &&
                                             Contains(reason, "类型"),
                "快照类型变化未被识别：" + reason);
        }

        private static void CheckBuildNode(SelfTestReport report)
        {
            JObject before = BuildExcelSnapshot("Book1.xlsx", "Sheet1", "A1:B2",
                Matrix(Row(1, 2), Row(3, 4)),
                Matrix(Row(string.Empty, string.Empty), Row(string.Empty, string.Empty)),
                Matrix(Row("General", "General"), Row("General", "General")));
            JObject after = (JObject)before.DeepClone();
            JArray afterValues = Matrix(Row(1, 2), Row(3, 9));
            after["values"] = afterValues;
            after["signature"] = SnapshotUndoEngine.ComputeSignature(
                afterValues, after["formulas"] as JArray, after["numberFormats"] as JArray);

            SnapshotUndoNode node = SnapshotUndoEngine.BuildNode(
                SnapshotCapture.Ok(before), SnapshotCapture.Ok(after), true, "写测试区域");
            Check(report, "BuildNode 可还原", node.Restorable && node.Id.StartsWith("undo-", StringComparison.Ordinal),
                "内容变化且结构稳定时应得到可还原节点：" + node.Reason);
            Check(report, "BuildNode 记录 afterSignature",
                node.Snapshot != null &&
                string.Equals(node.Snapshot.Value<string>("afterSignature"), after.Value<string>("signature"), StringComparison.Ordinal),
                "节点快照未记录执行后签名");
            Check(report, "BuildNode 文档信息", node.DocumentName == "Book1.xlsx" && node.SheetName == "Sheet1" &&
                                                 node.Address == "A1:B2",
                "节点文档/工作表/区域信息不正确");

            SnapshotUndoNode unchanged = SnapshotUndoEngine.BuildNode(
                SnapshotCapture.Ok(before), SnapshotCapture.Ok(before), true, "无变化");
            Check(report, "BuildNode 无变化拒绝", !unchanged.Restorable && Contains(unchanged.Reason, "未发生变化"),
                "内容未变化时应返回不可还原原因：" + unchanged.Reason);

            SnapshotUndoNode unstable = SnapshotUndoEngine.BuildNode(
                SnapshotCapture.Ok(before), SnapshotCapture.Ok(after), false, "结构变化");
            Check(report, "BuildNode 结构不稳定拒绝", !unstable.Restorable && Contains(unstable.Reason, "结构"),
                "structureStable=false 时应拒绝登记：" + unstable.Reason);

            SnapshotUndoNode beforeFailed = SnapshotUndoEngine.BuildNode(
                SnapshotCapture.Failed("模拟采集失败"), SnapshotCapture.Ok(after), true, "采集失败");
            Check(report, "BuildNode 执行前采集失败拒绝",
                !beforeFailed.Restorable && Contains(beforeFailed.Reason, "执行前快照采集失败"),
                "执行前采集失败应如实返回原因：" + beforeFailed.Reason);

            SnapshotUndoNode afterFailed = SnapshotUndoEngine.BuildNode(
                SnapshotCapture.Ok(before), SnapshotCapture.Failed("模拟采集失败"), true, "采集失败");
            Check(report, "BuildNode 执行后采集失败拒绝",
                !afterFailed.Restorable && Contains(afterFailed.Reason, "执行后快照采集失败"),
                "执行后采集失败应如实返回原因：" + afterFailed.Reason);

            JObject wordBefore = BuildWordSnapshot("Doc1.docx", 5, 5, "abc", SampleFont(false), false);
            JObject wordAfter = BuildWordSnapshot("Doc1.docx", 5, 8, "abcdef", SampleFont(false), false);
            SnapshotUndoNode wordNode = SnapshotUndoEngine.BuildNode(
                SnapshotCapture.Ok(wordBefore), SnapshotCapture.Ok(wordAfter), true, "插入文本");
            Check(report, "Word BuildNode 可还原", wordNode.Restorable &&
                                                    ReadInt(wordNode.Snapshot, "afterStart") == 5 &&
                                                    ReadInt(wordNode.Snapshot, "afterEnd") == 8,
                "Word 快照应记录执行后区段位置：" + wordNode.Reason);

            JObject mixedBefore = BuildWordSnapshot("Doc1.docx", 5, 5, "abc", SampleFont(true), true);
            JObject mixedAfter = BuildWordSnapshot("Doc1.docx", 5, 8, "abcdef", SampleFont(true), true);
            SnapshotUndoNode mixedNode = SnapshotUndoEngine.BuildNode(
                SnapshotCapture.Ok(mixedBefore), SnapshotCapture.Ok(mixedAfter), true, "混合字体");
            Check(report, "Word 混合字体拒绝", !mixedNode.Restorable && Contains(mixedNode.Reason, "混合"),
                "混合字体应如实拒绝还原：" + mixedNode.Reason);

            JObject otherDoc = BuildWordSnapshot("Doc2.docx", 5, 8, "abcdef", SampleFont(false), false);
            SnapshotUndoNode otherDocNode = SnapshotUndoEngine.BuildNode(
                SnapshotCapture.Ok(wordBefore), SnapshotCapture.Ok(otherDoc), true, "换文档");
            Check(report, "Word 文档名不一致拒绝", !otherDocNode.Restorable && Contains(otherDocNode.Reason, "文档名"),
                "Word 文档名变化应拒绝登记：" + otherDocNode.Reason);

            JObject roundTripNode = node.ToJObject();
            SnapshotUndoNode restoredNode = SnapshotUndoNode.FromJObject(roundTripNode);
            Check(report, "节点 ToJObject/FromJObject 往返",
                restoredNode != null && restoredNode.Id == node.Id && restoredNode.Restorable &&
                restoredNode.Kind == node.Kind && restoredNode.Snapshot != null,
                "节点序列化往返丢失字段");
        }

        private static void CheckExcelReplayDefenseChain(SelfTestReport report)
        {
            JObject before = BuildExcelSnapshot("Book1.xlsx", "Sheet1", "A1:B2",
                Matrix(Row(1, 2), Row(3, 4)),
                Matrix(Row("=A1", string.Empty), Row(string.Empty, string.Empty)),
                Matrix(Row("General", "General"), Row("General", "General")));
            JObject after = (JObject)before.DeepClone();
            JArray afterValues = Matrix(Row(1, 2), Row(3, 9));
            after["values"] = afterValues;
            after["signature"] = SnapshotUndoEngine.ComputeSignature(
                afterValues, after["formulas"] as JArray, after["numberFormats"] as JArray);

            SnapshotUndoNode node = SnapshotUndoEngine.BuildNode(
                SnapshotCapture.Ok(before), SnapshotCapture.Ok(after), true, "回放测试");
            JObject state = node.Snapshot;
            string beforeSignature = state.Value<string>("signature");
            string afterSignature = state.Value<string>("afterSignature");

            ReplayValidation missing = SnapshotUndoEngine.ValidateExcelReplay(null, "Book1.xlsx", 2, 2, afterSignature);
            Check(report, "防御 1 快照缺失", !missing.CanReplay && Contains(missing.Reason, "缺失"),
                "空快照未被拒绝：" + missing.Reason);

            ReplayValidation wrongWorkbook = SnapshotUndoEngine.ValidateExcelReplay(
                state, "Book2.xlsx", 2, 2, afterSignature);
            Check(report, "防御 2 活动工作簿不匹配",
                !wrongWorkbook.CanReplay && Contains(wrongWorkbook.Reason, "工作簿"),
                "活动工作簿不匹配未被拒绝：" + wrongWorkbook.Reason);

            ReplayValidation wrongShape = SnapshotUndoEngine.ValidateExcelReplay(
                state, "Book1.xlsx", 3, 2, afterSignature);
            Check(report, "防御 3 结构形状不一致",
                !wrongShape.CanReplay && Contains(wrongShape.Reason, "形状"),
                "区域形状不一致未被拒绝：" + wrongShape.Reason);

            ReplayValidation userChanged = SnapshotUndoEngine.ValidateExcelReplay(
                state, "Book1.xlsx", 2, 2, "deadbeef");
            Check(report, "防御 4 签名不符",
                !userChanged.CanReplay && Contains(userChanged.Reason, "其他改动"),
                "用户后续改动未被识别：" + userChanged.Reason);

            ReplayValidation already = SnapshotUndoEngine.ValidateExcelReplay(
                state, "Book1.xlsx", 2, 2, beforeSignature);
            Check(report, "已处于执行前状态", already.AlreadyRestored && !already.CanReplay,
                "当前签名等于执行前签名时应报告 already-restored");

            ReplayValidation allowed = SnapshotUndoEngine.ValidateExcelReplay(
                state, "Book1.xlsx", 2, 2, afterSignature);
            Check(report, "全部防御通过", allowed.CanReplay && allowed.Rows == 2 && allowed.Columns == 2 &&
                                           allowed.Address == "A1:B2" && allowed.SheetName == "Sheet1",
                "合法快照未通过校验：" + allowed.Reason);

            JObject badMatrix = (JObject)state.DeepClone();
            badMatrix["values"] = Matrix(Row(1));
            ReplayValidation badMatrixResult = SnapshotUndoEngine.ValidateExcelReplay(
                badMatrix, "Book1.xlsx", 2, 2, afterSignature);
            Check(report, "快照矩阵形状损坏",
                !badMatrixResult.CanReplay && Contains(badMatrixResult.Reason, "矩阵"),
                "矩阵形状损坏未被拒绝：" + badMatrixResult.Reason);

            JObject missingValues = (JObject)state.DeepClone();
            missingValues.Remove("values");
            ReplayValidation missingValuesResult = SnapshotUndoEngine.ValidateExcelReplay(
                missingValues, "Book1.xlsx", 2, 2, afterSignature);
            Check(report, "快照内容缺失",
                !missingValuesResult.CanReplay && Contains(missingValuesResult.Reason, "内容不完整"),
                "缺失 Value2 矩阵未被拒绝：" + missingValuesResult.Reason);

            JObject missingAfterSignature = (JObject)state.DeepClone();
            missingAfterSignature.Remove("afterSignature");
            ReplayValidation missingAfterResult = SnapshotUndoEngine.ValidateExcelReplay(
                missingAfterSignature, "Book1.xlsx", 2, 2, afterSignature);
            Check(report, "执行后签名缺失",
                !missingAfterResult.CanReplay && Contains(missingAfterResult.Reason, "执行后签名"),
                "缺少执行后签名未被拒绝：" + missingAfterResult.Reason);
        }

        private static void CheckWordReplayDefenseChain(SelfTestReport report)
        {
            JObject before = BuildWordSnapshot("Doc1.docx", 5, 5, "abc", SampleFont(false), false);
            JObject after = BuildWordSnapshot("Doc1.docx", 5, 8, "abcdef", SampleFont(false), false);
            SnapshotUndoNode node = SnapshotUndoEngine.BuildNode(
                SnapshotCapture.Ok(before), SnapshotCapture.Ok(after), true, "Word 回放测试");
            JObject state = node.Snapshot;
            string beforeSignature = state.Value<string>("signature");
            string afterSignature = state.Value<string>("afterSignature");

            ReplayValidation wrongDocument = SnapshotUndoEngine.ValidateWordReplay(
                state, "Doc2.docx", afterSignature, 5, 8);
            Check(report, "Word 防御 文档名不匹配",
                !wrongDocument.CanReplay && Contains(wrongDocument.Reason, "文档"),
                "活动文档不匹配未被拒绝：" + wrongDocument.Reason);

            ReplayValidation wrongPosition = SnapshotUndoEngine.ValidateWordReplay(
                state, "Doc1.docx", afterSignature, 6, 9);
            Check(report, "Word 防御 区段位置变化",
                !wrongPosition.CanReplay && Contains(wrongPosition.Reason, "位置"),
                "区段位置变化未被拒绝：" + wrongPosition.Reason);

            ReplayValidation changedByUser = SnapshotUndoEngine.ValidateWordReplay(
                state, "Doc1.docx", "deadbeef", 5, 8);
            Check(report, "Word 防御 签名不符",
                !changedByUser.CanReplay && Contains(changedByUser.Reason, "其他改动"),
                "用户后续改动未被识别：" + changedByUser.Reason);

            ReplayValidation already = SnapshotUndoEngine.ValidateWordReplay(
                state, "Doc1.docx", beforeSignature, 5, 8);
            Check(report, "Word 已处于执行前状态", already.AlreadyRestored,
                "当前签名等于执行前签名时应报告 already-restored");

            ReplayValidation allowed = SnapshotUndoEngine.ValidateWordReplay(
                state, "Doc1.docx", afterSignature, 5, 8);
            Check(report, "Word 全部防御通过", allowed.CanReplay && allowed.Text == "abc",
                "合法 Word 快照未通过校验：" + allowed.Reason);

            JObject missingAfter = (JObject)state.DeepClone();
            missingAfter.Remove("afterSignature");
            ReplayValidation missingAfterResult = SnapshotUndoEngine.ValidateWordReplay(
                missingAfter, "Doc1.docx", afterSignature, 5, 8);
            Check(report, "Word 执行后签名缺失",
                !missingAfterResult.CanReplay && Contains(missingAfterResult.Reason, "执行后签名"),
                "缺少执行后签名未被拒绝：" + missingAfterResult.Reason);

            JObject mixedState = (JObject)state.DeepClone();
            mixedState["fontMixed"] = true;
            ReplayValidation mixedResult = SnapshotUndoEngine.ValidateWordReplay(
                mixedState, "Doc1.docx", afterSignature, 5, 8);
            Check(report, "Word 混合字体防御",
                !mixedResult.CanReplay && Contains(mixedResult.Reason, "混合"),
                "混合字体快照未被拒绝：" + mixedResult.Reason);

            JObject missingText = (JObject)state.DeepClone();
            missingText.Remove("text");
            ReplayValidation missingTextResult = SnapshotUndoEngine.ValidateWordReplay(
                missingText, "Doc1.docx", afterSignature, 5, 8);
            Check(report, "Word 文本缺失",
                !missingTextResult.CanReplay && Contains(missingTextResult.Reason, "文本"),
                "缺失文本未被拒绝：" + missingTextResult.Reason);
        }

        private static void CheckFormulaDefense(SelfTestReport report)
        {
            Check(report, "公式回写 正常公式",
                SnapshotUndoEngine.ShouldWriteFormula("=SUM(A1:A2)") &&
                SnapshotUndoEngine.ShouldWriteFormula("=A1"),
                "正常公式应以 Formula 回写");
            Check(report, "公式回写 日期样式文本",
                !SnapshotUndoEngine.ShouldWriteFormula("1-1") &&
                !SnapshotUndoEngine.ShouldWriteFormula("001"),
                "非公式文本禁止用 Formula 回写");
            Check(report, "公式回写 空值与空白",
                !SnapshotUndoEngine.ShouldWriteFormula(null) &&
                !SnapshotUndoEngine.ShouldWriteFormula(string.Empty) &&
                !SnapshotUndoEngine.ShouldWriteFormula(" =A1"),
                "空值/前导空格的伪公式不应以 Formula 回写");
        }

        private static void CheckUndoStack(SelfTestReport report)
        {
            JObject before = BuildExcelSnapshot("Book1.xlsx", "Sheet1", "A1:B2",
                Matrix(Row(1, 2), Row(3, 4)),
                Matrix(Row(string.Empty, string.Empty), Row(string.Empty, string.Empty)),
                Matrix(Row("General", "General"), Row("General", "General")));
            JObject after = (JObject)before.DeepClone();
            JArray afterValues = Matrix(Row(1, 2), Row(3, 8));
            after["values"] = afterValues;
            after["signature"] = SnapshotUndoEngine.ComputeSignature(
                afterValues, after["formulas"] as JArray, after["numberFormats"] as JArray);

            SnapshotUndoNode first = SnapshotUndoEngine.BuildNode(
                SnapshotCapture.Ok(before), SnapshotCapture.Ok(after), true, "栈测试 1");
            SnapshotUndoNode second = SnapshotUndoEngine.BuildNode(
                SnapshotCapture.Ok(before), SnapshotCapture.Ok(after), true, "栈测试 2");
            SnapshotUndoNode nonRestorable = SnapshotUndoEngine.BuildNode(
                SnapshotCapture.Ok(before), SnapshotCapture.Ok(before), true, "不可还原");

            UndoStack stack = new UndoStack();
            stack.Push(first);
            stack.Push(second);
            Check(report, "UndoStack 入栈计数", stack.Count == 2, "入栈后计数应为 2");

            SnapshotUndoNode peeked;
            Check(report, "UndoStack 栈顶", stack.TryPeek(out peeked) && peeked.Id == second.Id,
                "栈顶节点不是最后入栈的节点");
            Check(report, "UndoStack 查找", stack.Find(first.Id) == first,
                "按 Id 查找失败");
            Check(report, "UndoStack 按 Id 移除", stack.Remove(first.Id) && stack.Count == 1,
                "按 Id 移除失败");
            Check(report, "UndoStack 弹出", stack.TryTakeLast(out peeked) && peeked.Id == second.Id && stack.Count == 0,
                "弹出栈顶失败");

            bool rejected = false;
            try
            {
                stack.Push(nonRestorable);
            }
            catch (InvalidOperationException ex)
            {
                rejected = Contains(ex.Message, "不可还原");
            }
            Check(report, "UndoStack 拒绝不可还原节点", rejected,
                "不可还原节点不得入栈");

            stack.Push(first);
            IList<SnapshotUndoNode> copy = stack.ToList();
            copy.Clear();
            Check(report, "UndoStack 列表为副本", stack.Count == 1,
                "ToList 返回的列表修改影响了内部栈");
            stack.Clear();
            Check(report, "UndoStack 清空", stack.Count == 0, "Clear 后计数应为 0");
        }

        private static JObject BuildExcelSnapshot(
            string workbook,
            string sheet,
            string address,
            JArray values,
            JArray formulas,
            JArray numberFormats)
        {
            int rows = values == null ? 0 : values.Count;
            int columns = rows == 0 ? 0 : ((JArray)values[0]).Count;

            JObject snapshot = new JObject();
            snapshot["kind"] = SnapshotUndoEngine.ExcelKind;
            snapshot["workbook"] = workbook;
            snapshot["sheet"] = sheet;
            snapshot["address"] = address;
            snapshot["rows"] = rows;
            snapshot["columns"] = columns;
            snapshot["values"] = values;
            snapshot["formulas"] = formulas;
            snapshot["numberFormats"] = numberFormats;
            snapshot["signature"] = SnapshotUndoEngine.ComputeSignature(values, formulas, numberFormats);
            return snapshot;
        }

        private static JObject BuildWordSnapshot(
            string document,
            int start,
            int end,
            string text,
            JObject font,
            bool fontMixed)
        {
            JObject snapshot = new JObject();
            snapshot["kind"] = SnapshotUndoEngine.WordKind;
            snapshot["document"] = document;
            snapshot["rangeLabel"] = "测试区段";
            snapshot["start"] = start;
            snapshot["end"] = end;
            snapshot["text"] = text;
            snapshot["font"] = font;
            snapshot["fontMixed"] = fontMixed;
            snapshot["signature"] = SnapshotUndoEngine.ComputeWordSignature(text, font);
            return snapshot;
        }

        private static JObject SampleFont(bool mixed)
        {
            JObject font = new JObject();
            font["name"] = mixed ? string.Empty : "宋体";
            font["nameFarEast"] = mixed ? string.Empty : "宋体";
            font["nameAscii"] = mixed ? string.Empty : "Calibri";
            font["size"] = mixed ? 9999999d : 12d;
            font["bold"] = mixed ? 9999999 : 0;
            font["italic"] = 0;
            font["underline"] = 0;
            font["color"] = 0;
            font["subscript"] = 0;
            font["superscript"] = 0;
            font["strikeThrough"] = 0;
            font["doubleStrikeThrough"] = 0;
            font["smallCaps"] = 0;
            font["allCaps"] = 0;
            font["spacing"] = 0d;
            font["scaling"] = 100;
            return font;
        }

        private static JArray Row(params object[] values)
        {
            JArray row = new JArray();
            foreach (object value in values)
            {
                row.Add(value);
            }
            return row;
        }

        private static JArray Matrix(params JArray[] rows)
        {
            JArray matrix = new JArray();
            foreach (JArray row in rows)
            {
                matrix.Add(row);
            }
            return matrix;
        }

        private static int ReadInt(JObject json, string name)
        {
            if (json == null || json[name] == null)
            {
                return 0;
            }
            try
            {
                return json.Value<int>(name);
            }
            catch
            {
                return 0;
            }
        }

        private static bool IsHex(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }
            foreach (char character in text)
            {
                bool digit = character >= '0' && character <= '9';
                bool lower = character >= 'a' && character <= 'f';
                if (!digit && !lower)
                {
                    return false;
                }
            }
            return true;
        }

        private static bool Contains(string text, string keyword)
        {
            return text != null && text.IndexOf(keyword, StringComparison.Ordinal) >= 0;
        }

        private static void Check(SelfTestReport report, string name, bool condition, string detail)
        {
            if (condition)
            {
                report.Pass();
            }
            else
            {
                report.Fail(name, detail);
            }
        }
    }
}
