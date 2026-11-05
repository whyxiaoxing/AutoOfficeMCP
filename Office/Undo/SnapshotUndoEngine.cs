using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Office.Undo
{
    /// <summary>
    /// 一次快照采集的结果。采集失败也必须给出真实原因，调用方不得据此伪造撤销节点。
    /// </summary>
    public sealed class SnapshotCapture
    {
        public bool Success { get; private set; }

        public string Reason { get; private set; }

        public JObject Snapshot { get; private set; }

        private SnapshotCapture(bool success, string reason, JObject snapshot)
        {
            Success = success;
            Reason = reason ?? string.Empty;
            Snapshot = snapshot;
        }

        public static SnapshotCapture Ok(JObject snapshot)
        {
            if (snapshot == null)
            {
                return Failed("快照内容为空，采集失败。");
            }
            return new SnapshotCapture(true, string.Empty, (JObject)snapshot.DeepClone());
        }

        public static SnapshotCapture Failed(string reason)
        {
            string safeReason = string.IsNullOrWhiteSpace(reason) ? "快照采集失败（原因未知）。" : reason.Trim();
            return new SnapshotCapture(false, safeReason, null);
        }
    }

    /// <summary>
    /// 回放前校验结果（纯逻辑，不触碰 COM）。
    /// AlreadyRestored=true 表示当前状态已经等于执行前快照，无需写回；CanReplay=true 表示可以安全写回。
    /// </summary>
    public sealed class ReplayValidation
    {
        public bool CanReplay { get; internal set; }

        public bool AlreadyRestored { get; internal set; }

        public string Reason { get; internal set; }

        public JObject Snapshot { get; internal set; }

        public string WorkbookName { get; internal set; }

        public string SheetName { get; internal set; }

        public string Address { get; internal set; }

        public string DocumentName { get; internal set; }

        public string BeforeSignature { get; internal set; }

        public string AfterSignature { get; internal set; }

        public int Rows { get; internal set; }

        public int Columns { get; internal set; }

        public JArray Values { get; internal set; }

        public JArray Formulas { get; internal set; }

        public JArray NumberFormats { get; internal set; }

        public string Text { get; internal set; }

        public JObject Font { get; internal set; }

        public int AfterStart { get; internal set; }

        public int AfterEnd { get; internal set; }

        public ReplayValidation()
        {
            Reason = string.Empty;
        }
    }

    /// <summary>
    /// 通用快照撤销引擎（Excel 区域 / Word 区段）。
    ///
    /// Excel 采集 Value2 + Formula + NumberFormat 三组二维矩阵；
    /// Word 采集 Range.Text + 关键字体属性。
    /// 所有快照都带 SHA256 签名，回放前依次执行五道防御：
    /// 1) 快照缺失；2) 活动文档名不匹配；3) 结构形状不一致；
    /// 4) 签名不符（用户又改过）；5) 公式逐格回写（非公式单元格禁止 Formula 文本回写）。
    ///
    /// 本引擎的 Capture/Replay 只负责 COM 读写，不获取会话锁；执行器应通过
    /// OfficeSession.RunExcel/RunWord 调用，确保"当前活动文档"类操作串行。
    /// </summary>
    public static class SnapshotUndoEngine
    {
        public const int MaxCells = 4000;
        public const int MaxWordParagraphs = 4000;
        public const string ExcelKind = "excel_range_state";
        public const string WordKind = "word_range_state";

        private const int WdUndefinedValue = 9999999;
        private const string AfterSignatureKey = "afterSignature";

        // ------------------------------------------------------------------
        // Excel 采集
        // ------------------------------------------------------------------

        /// <summary>从当前活动工作簿/活动工作表抓取指定地址的区域快照。</summary>
        public static SnapshotCapture CaptureExcel(Excel.Application application, string address)
        {
            ComHelper.EnsureSta("SnapshotUndoEngine.CaptureExcel");
            if (application == null)
            {
                return SnapshotCapture.Failed("Excel 应用不可用，无法采集快照。");
            }
            if (string.IsNullOrWhiteSpace(address))
            {
                return SnapshotCapture.Failed("目标区域地址为空，无法采集快照。");
            }

            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Range range = null;
            try
            {
                workbook = application.ActiveWorkbook;
                if (workbook == null)
                {
                    return SnapshotCapture.Failed("没有活动工作簿，无法采集快照。");
                }
                object activeSheet = application.ActiveSheet;
                worksheet = activeSheet as Excel.Worksheet;
                if (worksheet == null)
                {
                    ComHelper.Release(activeSheet);
                    return SnapshotCapture.Failed("活动工作表不可用（当前可能是图表页或其他非工作表对象）。");
                }
                range = worksheet.Range[address.Trim()];
                return CaptureExcelCore(workbook, worksheet, range, address.Trim());
            }
            catch (Exception ex)
            {
                return SnapshotCapture.Failed("Excel 区域快照采集失败：" + OfficeErrorMapper.DescribeException(ex, OfficeApp.Excel));
            }
            finally
            {
                ComHelper.Release(range, worksheet, workbook);
            }
        }

        /// <summary>从调用方已有的工作簿/工作表/区域对象抓取快照（不释放调用方对象）。</summary>
        public static SnapshotCapture CaptureExcel(Excel.Workbook workbook, Excel.Worksheet worksheet, Excel.Range range, string address)
        {
            ComHelper.EnsureSta("SnapshotUndoEngine.CaptureExcel");
            if (workbook == null || worksheet == null || range == null)
            {
                return SnapshotCapture.Failed("工作簿、工作表或目标区域为空，无法采集快照。");
            }
            if (string.IsNullOrWhiteSpace(address))
            {
                return SnapshotCapture.Failed("目标区域地址为空，无法采集快照。");
            }
            return CaptureExcelCore(workbook, worksheet, range, address.Trim());
        }

        private static SnapshotCapture CaptureExcelCore(Excel.Workbook workbook, Excel.Worksheet worksheet, Excel.Range range, string address)
        {
            try
            {
                int rows;
                int columns;
                ReadRangeShape(range, out rows, out columns);
                long total = (long)rows * (long)columns;
                if (rows <= 0 || columns <= 0)
                {
                    return SnapshotCapture.Failed("目标区域 " + address + " 的行列数无效（" + rows + " 行 × " + columns + " 列）。");
                }
                if (total > MaxCells)
                {
                    return SnapshotCapture.Failed(
                        "目标区域 " + address + " 共 " + rows + " 行 × " + columns + " 列，超过 " + MaxCells +
                        " 单元格快照上限，未采集撤销快照。");
                }

                JArray values;
                JArray formulas;
                JArray numberFormats;
                if (!TryReadExcelMatrices(range, rows, columns, out values, out formulas, out numberFormats))
                {
                    return SnapshotCapture.Failed(
                        "无法忠实读取区域 " + address +
                        " 的 Value2/Formula/NumberFormat（可能是特殊单元格），未采集撤销快照。");
                }

                JObject snapshot = new JObject();
                snapshot["kind"] = ExcelKind;
                snapshot["workbook"] = workbook.Name;
                snapshot["sheet"] = worksheet.Name;
                snapshot["address"] = address;
                snapshot["rows"] = rows;
                snapshot["columns"] = columns;
                snapshot["values"] = values;
                snapshot["formulas"] = formulas;
                snapshot["numberFormats"] = numberFormats;
                snapshot["signature"] = ComputeSignature(values, formulas, numberFormats);
                snapshot["capturedUtc"] = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
                return SnapshotCapture.Ok(snapshot);
            }
            catch (Exception ex)
            {
                return SnapshotCapture.Failed("Excel 区域快照采集失败：" + OfficeErrorMapper.DescribeException(ex, OfficeApp.Excel));
            }
        }

        // ------------------------------------------------------------------
        // Word 采集
        // ------------------------------------------------------------------

        /// <summary>抓取当前活动文档当前选区的快照。</summary>
        public static SnapshotCapture CaptureWord(Word.Application application, string label)
        {
            ComHelper.EnsureSta("SnapshotUndoEngine.CaptureWord");
            if (application == null)
            {
                return SnapshotCapture.Failed("Word 应用不可用，无法采集快照。");
            }

            Word.Document document = null;
            Word.Selection selection = null;
            Word.Range range = null;
            try
            {
                document = application.ActiveDocument;
                if (document == null)
                {
                    return SnapshotCapture.Failed("没有活动文档，无法采集快照。");
                }
                selection = application.Selection;
                if (selection == null)
                {
                    return SnapshotCapture.Failed("当前没有有效的 Word 选区，无法采集快照。");
                }
                range = selection.Range;
                return CaptureWordCore(document, range, label);
            }
            catch (Exception ex)
            {
                return SnapshotCapture.Failed("Word 区段快照采集失败：" + OfficeErrorMapper.DescribeException(ex, OfficeApp.Word));
            }
            finally
            {
                ComHelper.Release(range, selection, document);
            }
        }

        /// <summary>从调用方已有的文档/区段对象抓取快照（不释放调用方对象）。</summary>
        public static SnapshotCapture CaptureWord(Word.Document document, Word.Range range, string label)
        {
            ComHelper.EnsureSta("SnapshotUndoEngine.CaptureWord");
            if (document == null || range == null)
            {
                return SnapshotCapture.Failed("文档或目标区段为空，无法采集快照。");
            }
            return CaptureWordCore(document, range, label);
        }

        private static SnapshotCapture CaptureWordCore(Word.Document document, Word.Range range, string label)
        {
            try
            {
                string text = range.Text ?? string.Empty;
                string rangeLabel = string.IsNullOrWhiteSpace(label) ? "Word 选区" : label.Trim();

                int paragraphs = 0;
                Word.Paragraphs paragraphCollection = null;
                try
                {
                    paragraphCollection = range.Paragraphs;
                    if (paragraphCollection != null)
                    {
                        paragraphs = paragraphCollection.Count;
                    }
                }
                catch
                {
                    paragraphs = 0;
                }
                finally
                {
                    ComHelper.Release(paragraphCollection);
                }

                if (paragraphs > MaxWordParagraphs)
                {
                    return SnapshotCapture.Failed(
                        "Word 区段「" + rangeLabel + "」共 " + paragraphs + " 段，超过 " + MaxWordParagraphs +
                        " 段快照上限，未采集撤销快照。");
                }

                JObject fontState = null;
                bool fontMixed = false;
                Word.Font font = null;
                try
                {
                    font = range.Font;
                    fontState = CaptureWordFont(font, out fontMixed);
                }
                finally
                {
                    ComHelper.Release(font);
                }
                if (fontState == null)
                {
                    return SnapshotCapture.Failed("无法读取 Word 区段「" + rangeLabel + "」的字体属性，未采集撤销快照。");
                }

                JObject snapshot = new JObject();
                snapshot["kind"] = WordKind;
                snapshot["document"] = document.Name;
                snapshot["rangeLabel"] = rangeLabel;
                snapshot["start"] = range.Start;
                snapshot["end"] = range.End;
                snapshot["paragraphs"] = paragraphs;
                snapshot["text"] = text;
                snapshot["font"] = fontState;
                snapshot["fontMixed"] = fontMixed;
                snapshot["signature"] = ComputeWordSignature(text, fontState);
                snapshot["capturedUtc"] = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
                return SnapshotCapture.Ok(snapshot);
            }
            catch (Exception ex)
            {
                return SnapshotCapture.Failed("Word 区段快照采集失败：" + OfficeErrorMapper.DescribeException(ex, OfficeApp.Word));
            }
        }

        // ------------------------------------------------------------------
        // 节点构建
        // ------------------------------------------------------------------

        public static SnapshotUndoNode BuildNode(SnapshotCapture before, SnapshotCapture after, bool structureStable, string summary)
        {
            string summaryText = NormalizeSummary(summary);
            if (before == null || !before.Success)
            {
                string reason = before == null
                    ? "执行前快照缺失，未登记可撤销节点。"
                    : "执行前快照采集失败，未登记可撤销节点：" + before.Reason;
                return CreateUnrestorable(ResolveKind(null, after), summaryText, reason);
            }
            if (after == null || !after.Success)
            {
                string reason = after == null
                    ? "执行后快照缺失，未登记可撤销节点。"
                    : "执行后快照采集失败，未登记可撤销节点：" + after.Reason;
                return CreateUnrestorable(ResolveKind(before, null), summaryText, reason);
            }
            return BuildNode(before.Snapshot, after.Snapshot, structureStable, summaryText);
        }

        /// <summary>
        /// 纯逻辑节点构建：只有内容确实变化且结构稳定、快照完整时才返回 Restorable=true。
        /// 否则返回 Restorable=false + 真实中文原因，调用方不得把该节点压入 UndoStack。
        /// </summary>
        public static SnapshotUndoNode BuildNode(JObject beforeSnapshot, JObject afterSnapshot, bool structureStable, string summary)
        {
            string summaryText = NormalizeSummary(summary);
            if (beforeSnapshot == null || afterSnapshot == null)
            {
                return CreateUnrestorable("unknown", summaryText, "执行前或执行后快照缺失，未登记可撤销节点。");
            }

            string kind = ReadString(beforeSnapshot, "kind");
            if (kind.Length == 0)
            {
                kind = ReadString(afterSnapshot, "kind");
            }
            if (kind.Length == 0)
            {
                return CreateUnrestorable("unknown", summaryText, "快照缺少类型信息，未登记可撤销节点。");
            }

            if (!structureStable)
            {
                return CreateUnrestorable(
                    kind,
                    summaryText,
                    "调用方判定文档结构已发生变化（可能插入/删除了行列、工作表或段落），未登记可撤销节点。");
            }

            string shapeReason;
            if (!ShapesMatch(beforeSnapshot, afterSnapshot, out shapeReason))
            {
                return CreateUnrestorable(kind, summaryText, shapeReason);
            }

            string beforeSignature = ReadString(beforeSnapshot, "signature");
            string afterSignature = ReadString(afterSnapshot, "signature");
            if (beforeSignature.Length == 0 || afterSignature.Length == 0)
            {
                return CreateUnrestorable(kind, summaryText, "快照缺少内容签名，未登记可撤销节点。");
            }
            if (string.Equals(beforeSignature, afterSignature, StringComparison.Ordinal))
            {
                return CreateUnrestorable(kind, summaryText, "目标区域内容未发生变化，无需撤销。");
            }

            string documentName;
            string sheetName = string.Empty;
            string address;
            if (string.Equals(kind, ExcelKind, StringComparison.Ordinal))
            {
                int rows = ReadInt(beforeSnapshot, "rows");
                int columns = ReadInt(beforeSnapshot, "columns");
                JArray values = beforeSnapshot["values"] as JArray;
                JArray formulas = beforeSnapshot["formulas"] as JArray;
                JArray formats = beforeSnapshot["numberFormats"] as JArray;
                JArray afterValues = afterSnapshot["values"] as JArray;
                JArray afterFormulas = afterSnapshot["formulas"] as JArray;
                JArray afterFormats = afterSnapshot["numberFormats"] as JArray;
                if (rows <= 0 || columns <= 0 ||
                    !MatrixHasShape(values, rows, columns) ||
                    !MatrixHasShape(formulas, rows, columns) ||
                    !MatrixHasShape(formats, rows, columns) ||
                    !MatrixHasShape(afterValues, rows, columns) ||
                    !MatrixHasShape(afterFormulas, rows, columns) ||
                    !MatrixHasShape(afterFormats, rows, columns))
                {
                    return CreateUnrestorable(kind, summaryText, "Excel 快照矩阵不完整或形状不一致，未登记可撤销节点。");
                }

                documentName = ReadString(beforeSnapshot, "workbook");
                sheetName = ReadString(beforeSnapshot, "sheet");
                address = ReadString(beforeSnapshot, "address");
                if (documentName.Length == 0 || sheetName.Length == 0 || address.Length == 0)
                {
                    return CreateUnrestorable(kind, summaryText, "Excel 快照缺少工作簿/工作表/区域信息，未登记可撤销节点。");
                }
            }
            else if (string.Equals(kind, WordKind, StringComparison.Ordinal))
            {
                string textValue;
                string afterTextValue;
                JObject font = beforeSnapshot["font"] as JObject;
                if (!TryReadString(beforeSnapshot, "text", out textValue) || font == null)
                {
                    return CreateUnrestorable(kind, summaryText, "Word 执行前快照缺少文本或字体属性，未登记可撤销节点。");
                }
                if (!TryReadString(afterSnapshot, "text", out afterTextValue) ||
                    (afterSnapshot["font"] as JObject) == null)
                {
                    return CreateUnrestorable(kind, summaryText, "Word 执行后快照缺少文本或字体属性，未登记可撤销节点。");
                }
                if (ReadBool(beforeSnapshot, "fontMixed") || ReadBool(afterSnapshot, "fontMixed"))
                {
                    return CreateUnrestorable(
                        kind,
                        summaryText,
                        "Word 区段的字体属性为混合值（部分文字格式不同），无法用整段字体属性忠实还原，未登记可撤销节点。");
                }

                int beforeStart;
                int beforeEnd;
                int afterStart;
                int afterEnd;
                if (!TryReadInt(beforeSnapshot, "start", out beforeStart) ||
                    !TryReadInt(beforeSnapshot, "end", out beforeEnd) ||
                    !TryReadInt(afterSnapshot, "start", out afterStart) ||
                    !TryReadInt(afterSnapshot, "end", out afterEnd) ||
                    beforeStart < 0 || beforeEnd < beforeStart ||
                    afterStart < 0 || afterEnd < afterStart)
                {
                    return CreateUnrestorable(kind, summaryText, "Word 快照缺少执行前/执行后的有效区段位置，未登记可撤销节点。");
                }

                documentName = ReadString(beforeSnapshot, "document");
                address = ReadString(beforeSnapshot, "rangeLabel");
                if (documentName.Length == 0)
                {
                    return CreateUnrestorable(kind, summaryText, "Word 快照缺少文档名，未登记可撤销节点。");
                }
            }
            else
            {
                return CreateUnrestorable(kind, summaryText, "未知快照类型「" + kind + "」，未登记可撤销节点。");
            }

            JObject state = (JObject)beforeSnapshot.DeepClone();
            state[AfterSignatureKey] = afterSignature;
            if (string.Equals(kind, WordKind, StringComparison.Ordinal))
            {
                state["afterStart"] = ReadInt(afterSnapshot, "start");
                state["afterEnd"] = ReadInt(afterSnapshot, "end");
            }


            string appType = string.Equals(kind, ExcelKind, StringComparison.Ordinal) ? "excel" : "word";
            return new SnapshotUndoNode(
                SnapshotUndoNode.NewId(),
                kind,
                appType,
                documentName,
                sheetName,
                address,
                summaryText,
                true,
                string.Empty,
                DateTime.UtcNow,
                state);
        }

        // ------------------------------------------------------------------
        // 回放前校验（纯逻辑，自检可脱离 Office）
        // ------------------------------------------------------------------

        public static ReplayValidation ValidateExcelReplay(
            JObject state,
            string activeWorkbookName,
            int currentRows,
            int currentColumns,
            string currentSignature)
        {
            ReplayValidation result = new ReplayValidation();
            if (state == null)
            {
                result.Reason = "还原快照缺失，已拒绝自动回退。";
                return result;
            }
            if (!string.Equals(ReadString(state, "kind"), ExcelKind, StringComparison.Ordinal))
            {
                result.Reason = "快照类型不是 Excel 区域快照，已拒绝还原。";
                return result;
            }

            string workbookName = ReadString(state, "workbook");
            string sheetName = ReadString(state, "sheet");
            string address = ReadString(state, "address");
            int rows = ReadInt(state, "rows");
            int columns = ReadInt(state, "columns");
            string beforeSignature = ReadString(state, "signature");
            string afterSignature = ReadString(state, AfterSignatureKey);

            if (workbookName.Length == 0 || sheetName.Length == 0 || address.Length == 0 ||
                rows <= 0 || columns <= 0 || beforeSignature.Length == 0)
            {
                result.Reason = "还原快照缺少工作簿/工作表/区域/形状/执行前签名信息，已拒绝自动回退。";
                return result;
            }

            if (string.IsNullOrWhiteSpace(activeWorkbookName))
            {
                result.Reason = "没有活动工作簿，无法还原本轮改动。";
                return result;
            }
            if (!string.Equals(activeWorkbookName, workbookName, StringComparison.OrdinalIgnoreCase))
            {
                result.Reason = "当前活动工作簿不是本轮改动所在的「" + workbookName + "」，已拒绝还原。";
                return result;
            }

            if (currentRows != rows || currentColumns != columns)
            {
                result.Reason = "目标区域 " + address + " 的形状已与本轮记录不一致（当前 " + currentRows +
                                " 行 × " + currentColumns + " 列，记录 " + rows + " 行 × " + columns +
                                " 列），直接写回会串位，已拒绝还原。";
                return result;
            }

            JArray values = state["values"] as JArray;
            JArray formulas = state["formulas"] as JArray;
            JArray formats = state["numberFormats"] as JArray;
            if (values == null || formulas == null || formats == null)
            {
                result.Reason = "还原快照内容不完整（缺少 Value2/Formula/NumberFormat），已拒绝自动回退。";
                return result;
            }
            if (!MatrixHasShape(values, rows, columns) ||
                !MatrixHasShape(formulas, rows, columns) ||
                !MatrixHasShape(formats, rows, columns))
            {
                result.Reason = "还原快照矩阵形状与记录的行列数不一致，已拒绝自动回退。";
                return result;
            }

            if (string.IsNullOrWhiteSpace(currentSignature))
            {
                result.Reason = "无法读取目标区域当前内容，已拒绝自动回退。";
                return result;
            }
            if (string.Equals(currentSignature, beforeSignature, StringComparison.Ordinal))
            {
                result.AlreadyRestored = true;
                return result;
            }
            if (afterSignature.Length == 0)
            {
                result.Reason = "还原快照缺少执行后签名，无法确认期间是否被用户修改，已拒绝自动回退。";
                return result;
            }
            if (!string.Equals(currentSignature, afterSignature, StringComparison.Ordinal))
            {
                result.Reason = "区域 " + address + " 在本轮之后又有其他改动，已拒绝还原以免覆盖你的内容。";
                return result;
            }

            result.CanReplay = true;
            result.Snapshot = state;
            result.WorkbookName = workbookName;
            result.SheetName = sheetName;
            result.Address = address;
            result.BeforeSignature = beforeSignature;
            result.AfterSignature = afterSignature;
            result.Rows = rows;
            result.Columns = columns;
            result.Values = values;
            result.Formulas = formulas;
            result.NumberFormats = formats;
            return result;
        }

        public static ReplayValidation ValidateWordReplay(
            JObject state,
            string activeDocumentName,
            string currentSignature,
            int currentStart,
            int currentEnd)
        {
            ReplayValidation result = new ReplayValidation();
            if (state == null)
            {
                result.Reason = "还原快照缺失，已拒绝自动回退。";
                return result;
            }
            if (!string.Equals(ReadString(state, "kind"), WordKind, StringComparison.Ordinal))
            {
                result.Reason = "快照类型不是 Word 区段快照，已拒绝还原。";
                return result;
            }

            string documentName = ReadString(state, "document");
            int start;
            int end;
            int afterStart;
            int afterEnd;
            string beforeSignature = ReadString(state, "signature");
            string afterSignature = ReadString(state, AfterSignatureKey);

            if (!TryReadInt(state, "start", out start) ||
                !TryReadInt(state, "end", out end) ||
                !TryReadInt(state, "afterStart", out afterStart) ||
                !TryReadInt(state, "afterEnd", out afterEnd) ||
                documentName.Length == 0 || start < 0 || end < start || afterStart < 0 ||
                afterEnd < afterStart || beforeSignature.Length == 0)
            {
                result.Reason = "还原快照缺少文档名/区段位置/执行前签名信息，已拒绝自动回退。";
                return result;
            }
            if (string.IsNullOrWhiteSpace(activeDocumentName))
            {
                result.Reason = "没有活动文档，无法还原本轮改动。";
                return result;
            }
            if (!string.Equals(activeDocumentName, documentName, StringComparison.OrdinalIgnoreCase))
            {
                result.Reason = "当前活动文档不是本轮改动所在的「" + documentName + "」，已拒绝还原。";
                return result;
            }
            if (currentStart != afterStart || currentEnd != afterEnd)
            {
                result.Reason = "目标区段的位置已发生变化（当前 " + currentStart + "-" + currentEnd +
                                "，记录 " + afterStart + "-" + afterEnd + "），已拒绝还原以免误伤其他内容。";
                return result;
            }

            string textValue;
            JObject font = state["font"] as JObject;
            if (!TryReadString(state, "text", out textValue) || font == null)
            {
                result.Reason = "还原快照缺少文本或字体属性，已拒绝自动回退。";
                return result;
            }
            if (ReadBool(state, "fontMixed"))
            {
                result.Reason = "还原快照的字体属性为混合值，无法忠实还原，已拒绝自动回退。";
                return result;
            }

            if (string.IsNullOrWhiteSpace(currentSignature))
            {
                result.Reason = "无法读取目标区段当前内容，已拒绝自动回退。";
                return result;
            }
            if (string.Equals(currentSignature, beforeSignature, StringComparison.Ordinal))
            {
                result.AlreadyRestored = true;
                return result;
            }
            if (afterSignature.Length == 0)
            {
                result.Reason = "还原快照缺少执行后签名，无法确认期间是否被用户修改，已拒绝自动回退。";
                return result;
            }
            if (!string.Equals(currentSignature, afterSignature, StringComparison.Ordinal))
            {
                result.Reason = "文档「" + documentName + "」在本轮之后又有其他改动，已拒绝还原以免覆盖你的内容。";
                return result;
            }

            result.CanReplay = true;
            result.Snapshot = state;
            result.DocumentName = documentName;
            result.BeforeSignature = beforeSignature;
            result.AfterSignature = afterSignature;
            result.Text = textValue;
            result.Font = font;
            result.AfterStart = afterStart;
            result.AfterEnd = afterEnd;
            return result;
        }

        // ------------------------------------------------------------------
        // Excel 回放
        // ------------------------------------------------------------------

        /// <summary>返回空字符串表示已还原并核对通过；其他返回值都是"没有还原"的真实原因。</summary>
        public static string ReplayExcel(Excel.Application application, JObject state)
        {
            ComHelper.EnsureSta("SnapshotUndoEngine.ReplayExcel");
            if (application == null)
            {
                return "Excel 应用不可用，无法还原。";
            }
            if (state == null)
            {
                return "还原快照缺失，已拒绝自动回退。";
            }
            if (!string.Equals(ReadString(state, "kind"), ExcelKind, StringComparison.Ordinal))
            {
                return "快照类型不是 Excel 区域快照，已拒绝还原。";
            }

            Excel.Workbook workbook = null;
            Excel.Sheets sheets = null;
            Excel.Worksheet worksheet = null;
            Excel.Range range = null;
            try
            {
                workbook = application.ActiveWorkbook;
                if (workbook == null)
                {
                    return "没有活动工作簿，无法还原本轮改动。";
                }

                string expectedWorkbook = ReadString(state, "workbook");
                if (expectedWorkbook.Length == 0)
                {
                    return "还原快照缺少工作簿名，已拒绝自动回退。";
                }
                if (!string.Equals(workbook.Name, expectedWorkbook, StringComparison.OrdinalIgnoreCase))
                {
                    return "当前活动工作簿不是本轮改动所在的「" + expectedWorkbook + "」，已拒绝还原。";
                }

                string sheetName = ReadString(state, "sheet");
                string address = ReadString(state, "address");
                if (sheetName.Length == 0 || address.Length == 0)
                {
                    return "还原快照缺少工作表或区域信息，已拒绝自动回退。";
                }

                sheets = workbook.Worksheets;
                object sheetObject = sheets.Item[sheetName];
                worksheet = sheetObject as Excel.Worksheet;
                if (worksheet == null)
                {
                    ComHelper.Release(sheetObject);
                    return "工作表「" + sheetName + "」不是普通工作表（可能是图表页），已拒绝还原。";
                }

                range = worksheet.Range[address];
                int currentRows;
                int currentColumns;
                ReadRangeShape(range, out currentRows, out currentColumns);
                string currentSignature = ReadExcelSignature(range, currentRows, currentColumns);
                if (currentSignature == null)
                {
                    return "无法读取目标区域当前内容，已拒绝自动回退。";
                }

                ReplayValidation validation = ValidateExcelReplay(
                    state, workbook.Name, currentRows, currentColumns, currentSignature);
                if (validation.AlreadyRestored)
                {
                    return string.Empty;
                }
                if (!validation.CanReplay)
                {
                    return validation.Reason;
                }

                object[,] numberFormats = ToObjectMatrix(validation.NumberFormats, validation.Rows, validation.Columns);
                object[,] values = ToObjectMatrix(validation.Values, validation.Rows, validation.Columns);
                if (numberFormats == null || values == null)
                {
                    return "还原快照内容不完整，已拒绝自动回退。";
                }

                range.NumberFormat = numberFormats;
                range.Value2 = values;
                string formulaError = WriteFormulas(range, validation.Formulas, validation.Rows, validation.Columns);
                if (formulaError.Length > 0)
                {
                    return formulaError;
                }

                string restored = ReadExcelSignature(range, validation.Rows, validation.Columns);
                if (restored == null)
                {
                    return "写回后无法核对目标区域，请对照变更清单人工确认。";
                }
                if (!string.Equals(restored, validation.BeforeSignature, StringComparison.Ordinal))
                {
                    return "已写回 " + validation.SheetName + "!" + validation.Address +
                           "，但与执行前的记录仍不一致，请对照变更清单人工核对。";
                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                return "还原失败：" + OfficeErrorMapper.DescribeException(ex, OfficeApp.Excel);
            }
            finally
            {
                ComHelper.Release(range, worksheet, sheets, workbook);
            }
        }

        // ------------------------------------------------------------------
        // Word 回放
        // ------------------------------------------------------------------

        /// <summary>返回空字符串表示已还原并核对通过；其他返回值都是"没有还原"的真实原因。</summary>
        public static string ReplayWord(Word.Application application, JObject state)
        {
            ComHelper.EnsureSta("SnapshotUndoEngine.ReplayWord");
            if (application == null)
            {
                return "Word 应用不可用，无法还原。";
            }
            if (state == null)
            {
                return "还原快照缺失，已拒绝自动回退。";
            }
            if (!string.Equals(ReadString(state, "kind"), WordKind, StringComparison.Ordinal))
            {
                return "快照类型不是 Word 区段快照，已拒绝还原。";
            }

            Word.Document document = null;
            Word.Range targetRange = null;
            Word.Font font = null;
            try
            {
                document = application.ActiveDocument;
                if (document == null)
                {
                    return "没有活动文档，无法还原本轮改动。";
                }

                string expectedDocument = ReadString(state, "document");
                if (expectedDocument.Length == 0)
                {
                    return "还原快照缺少文档名，已拒绝自动回退。";
                }
                if (!string.Equals(document.Name, expectedDocument, StringComparison.OrdinalIgnoreCase))
                {
                    return "当前活动文档不是本轮改动所在的「" + expectedDocument + "」，已拒绝还原。";
                }

                int afterStart = ReadInt(state, "afterStart");
                int afterEnd = ReadInt(state, "afterEnd");
                if (afterStart < 0 || afterEnd < afterStart)
                {
                    return "还原快照缺少执行后的有效区段位置，已拒绝自动回退。";
                }

                targetRange = GetWordRange(document, afterStart, afterEnd);
                string currentText = targetRange.Text ?? string.Empty;
                font = targetRange.Font;
                bool currentFontMixed;
                JObject currentFont = CaptureWordFont(font, out currentFontMixed);
                if (currentFont == null)
                {
                    return "无法读取目标区段当前字体属性，已拒绝自动回退。";
                }
                string currentSignature = ComputeWordSignature(currentText, currentFont);

                ReplayValidation validation = ValidateWordReplay(
                    state, document.Name, currentSignature, targetRange.Start, targetRange.End);
                if (validation.AlreadyRestored)
                {
                    return string.Empty;
                }
                if (!validation.CanReplay)
                {
                    return validation.Reason;
                }

                targetRange.Text = validation.Text;
                ComHelper.Release(font);
                font = targetRange.Font;
                ApplyWordFont(font, validation.Font);

                string restoredText = targetRange.Text ?? string.Empty;
                bool restoredMixed;
                JObject restoredFont = CaptureWordFont(font, out restoredMixed);
                string restoredSignature = restoredFont == null
                    ? string.Empty
                    : ComputeWordSignature(restoredText, restoredFont);
                if (restoredSignature.Length == 0 ||
                    !string.Equals(restoredSignature, validation.BeforeSignature, StringComparison.Ordinal))
                {
                    return "已写回文档「" + validation.DocumentName +
                           "」的目标区段，但与执行前的记录仍不一致，请对照变更清单人工核对。";
                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                return "还原失败：" + OfficeErrorMapper.DescribeException(ex, OfficeApp.Word);
            }
            finally
            {
                ComHelper.Release(font, targetRange, document);
            }
        }

        // ------------------------------------------------------------------
        // 纯逻辑工具
        // ------------------------------------------------------------------

        /// <summary>Excel 三矩阵 SHA256 签名（JSON 无空白序列化，跨调用稳定）。</summary>
        public static string ComputeSignature(JArray values, JArray formulas, JArray numberFormats)
        {
            string valuesText = values == null ? "[]" : values.ToString(Formatting.None);
            string formulasText = formulas == null ? "[]" : formulas.ToString(Formatting.None);
            string formatsText = numberFormats == null ? "[]" : numberFormats.ToString(Formatting.None);
            return Sha256(valuesText + "|" + formulasText + "|" + formatsText);
        }

        /// <summary>Word 文本 + 关键字体属性的签名。</summary>
        public static string ComputeWordSignature(string text, JObject font)
        {
            string fontText = font == null ? "{}" : font.ToString(Formatting.None);
            return Sha256((text ?? string.Empty) + "|" + fontText);
        }

        /// <summary>
        /// 第五道防御的核心判定：只有非空且以 "=" 开头的文本才允许用 Formula 写回。
        /// "1-1"、"001" 等文本必须走 Value2，不能被 Excel 重新解释成日期或数字。
        /// </summary>
        public static bool ShouldWriteFormula(string formulaText)
        {
            return !string.IsNullOrEmpty(formulaText) &&
                   formulaText.StartsWith("=", StringComparison.Ordinal);
        }

        /// <summary>比较两个快照的结构（工作表名/文档名、行列数、区域地址）。</summary>
        public static bool ShapesMatch(JObject before, JObject after, out string reason)
        {
            reason = string.Empty;
            if (before == null || after == null)
            {
                reason = "执行前或执行后快照缺失，无法比较结构。";
                return false;
            }

            string kind = ReadString(before, "kind");
            string otherKind = ReadString(after, "kind");
            if (!string.Equals(kind, otherKind, StringComparison.Ordinal))
            {
                reason = "快照类型不一致（" + kind + " → " + otherKind + "），未登记可撤销节点。";
                return false;
            }

            if (string.Equals(kind, ExcelKind, StringComparison.Ordinal))
            {
                string workbookBefore = ReadString(before, "workbook");
                string workbookAfter = ReadString(after, "workbook");
                if (!string.Equals(workbookBefore, workbookAfter, StringComparison.OrdinalIgnoreCase))
                {
                    reason = "工作簿名不一致（" + workbookBefore + " → " + workbookAfter + "），未登记可撤销节点。";
                    return false;
                }

                string sheetBefore = ReadString(before, "sheet");
                string sheetAfter = ReadString(after, "sheet");
                if (!string.Equals(sheetBefore, sheetAfter, StringComparison.OrdinalIgnoreCase))
                {
                    reason = "工作表名不一致（" + sheetBefore + " → " + sheetAfter + "），未登记可撤销节点。";
                    return false;
                }

                string addressBefore = ReadString(before, "address");
                string addressAfter = ReadString(after, "address");
                if (!string.Equals(addressBefore, addressAfter, StringComparison.OrdinalIgnoreCase))
                {
                    reason = "目标区域不一致（" + addressBefore + " → " + addressAfter + "），未登记可撤销节点。";
                    return false;
                }

                int rowsBefore = ReadInt(before, "rows");
                int columnsBefore = ReadInt(before, "columns");
                int rowsAfter = ReadInt(after, "rows");
                int columnsAfter = ReadInt(after, "columns");
                if (rowsBefore != rowsAfter || columnsBefore != columnsAfter)
                {
                    reason = "目标区域行列数不一致（" + rowsBefore + "×" + columnsBefore +
                             " → " + rowsAfter + "×" + columnsAfter + "），未登记可撤销节点。";
                    return false;
                }
                return true;
            }

            if (string.Equals(kind, WordKind, StringComparison.Ordinal))
            {
                string documentBefore = ReadString(before, "document");
                string documentAfter = ReadString(after, "document");
                if (!string.Equals(documentBefore, documentAfter, StringComparison.OrdinalIgnoreCase))
                {
                    reason = "文档名不一致（" + documentBefore + " → " + documentAfter + "），未登记可撤销节点。";
                    return false;
                }
                return true;
            }

            reason = "未知快照类型「" + kind + "」，无法比较结构。";
            return false;
        }

        public static bool MatrixHasShape(JArray matrix, int rows, int columns)
        {
            if (matrix == null || rows <= 0 || columns <= 0 || matrix.Count != rows)
            {
                return false;
            }
            for (int rowIndex = 0; rowIndex < rows; rowIndex++)
            {
                JArray row = matrix[rowIndex] as JArray;
                if (row == null || row.Count != columns)
                {
                    return false;
                }
            }
            return true;
        }

        // ------------------------------------------------------------------
        // COM 读写细节
        // ------------------------------------------------------------------

        private static void ReadRangeShape(Excel.Range range, out int rows, out int columns)
        {
            rows = 0;
            columns = 0;
            Excel.Range rowsRange = null;
            Excel.Range columnsRange = null;
            try
            {
                rowsRange = range.Rows;
                columnsRange = range.Columns;
                rows = rowsRange.Count;
                columns = columnsRange.Count;
            }
            finally
            {
                ComHelper.Release(rowsRange, columnsRange);
            }
        }

        private static bool TryReadExcelMatrices(
            Excel.Range range,
            int rows,
            int columns,
            out JArray values,
            out JArray formulas,
            out JArray numberFormats)
        {
            values = ToJagged(range.Value2, rows, columns);
            formulas = ToJagged(range.Formula, rows, columns);
            numberFormats = ToJagged(range.NumberFormat, rows, columns);
            if (numberFormats == null)
            {
                // NumberFormat 在混合格式区域会返回 DBNull（参考项目因此放弃快照）。
                // 这里逐格读取兜底：最多 4000 格，换来混合格式区域也能真实撤销。
                numberFormats = ReadNumberFormatsPerCell(range, rows, columns);
            }
            return values != null && formulas != null && numberFormats != null;
        }

        private static JArray ReadNumberFormatsPerCell(Excel.Range range, int rows, int columns)
        {
            Excel.Range cells = null;
            try
            {
                cells = range.Cells;
                JArray matrix = new JArray();
                for (int rowIndex = 0; rowIndex < rows; rowIndex++)
                {
                    JArray row = new JArray();
                    for (int columnIndex = 0; columnIndex < columns; columnIndex++)
                    {
                        Excel.Range cell = null;
                        try
                        {
                            cell = cells.Item[rowIndex + 1, columnIndex + 1] as Excel.Range;
                            if (cell == null)
                            {
                                return null;
                            }
                            row.Add(ToToken(cell.NumberFormat));
                        }
                        finally
                        {
                            ComHelper.Release(cell);
                        }
                    }
                    matrix.Add(row);
                }
                return matrix;
            }
            catch
            {
                return null;
            }
            finally
            {
                ComHelper.Release(cells);
            }
        }

        private static string ReadExcelSignature(Excel.Range range, int rows, int columns)
        {
            try
            {
                JArray values;
                JArray formulas;
                JArray numberFormats;
                if (!TryReadExcelMatrices(range, rows, columns, out values, out formulas, out numberFormats))
                {
                    return null;
                }
                return ComputeSignature(values, formulas, numberFormats);
            }
            catch
            {
                return null;
            }
        }

        private static string WriteFormulas(Excel.Range range, JArray formulas, int rows, int columns)
        {
            if (range == null || formulas == null)
            {
                return "公式写回失败：公式矩阵缺失。";
            }

            Excel.Range cells = null;
            try
            {
                cells = range.Cells;
                for (int rowIndex = 0; rowIndex < rows; rowIndex++)
                {
                    JArray row = formulas[rowIndex] as JArray;
                    if (row == null || row.Count != columns)
                    {
                        return "公式写回失败：公式矩阵形状不一致。";
                    }
                    for (int columnIndex = 0; columnIndex < columns; columnIndex++)
                    {
                        string formulaText = ReadCellText(row, columnIndex);
                        if (!ShouldWriteFormula(formulaText))
                        {
                            continue;
                        }

                        Excel.Range cell = null;
                        try
                        {
                            cell = cells.Item[rowIndex + 1, columnIndex + 1] as Excel.Range;
                            if (cell == null)
                            {
                                return "公式写回失败：无法取得单元格 R" + (rowIndex + 1) + "C" + (columnIndex + 1) + "。";
                            }
                            cell.Formula = formulaText;
                        }
                        finally
                        {
                            ComHelper.Release(cell);
                        }
                    }
                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                return "公式写回失败：" + OfficeErrorMapper.DescribeException(ex, OfficeApp.Excel);
            }
            finally
            {
                ComHelper.Release(cells);
            }
        }

        private static Word.Range GetWordRange(Word.Document document, int start, int end)
        {
            object startObject = start;
            object endObject = end;
            return document.Range(ref startObject, ref endObject);
        }

        private static JObject CaptureWordFont(Word.Font font, out bool mixed)
        {
            mixed = false;
            if (font == null)
            {
                return null;
            }

            string name = font.Name;
            string nameFarEast = font.NameFarEast;
            string nameAscii = font.NameAscii;
            float size = font.Size;
            int bold = font.Bold;
            int italic = font.Italic;
            int underline = (int)font.Underline;
            int color = (int)font.Color;
            int subscript = font.Subscript;
            int superscript = font.Superscript;
            int strikeThrough = font.StrikeThrough;
            int doubleStrikeThrough = font.DoubleStrikeThrough;
            int smallCaps = font.SmallCaps;
            int allCaps = font.AllCaps;
            float spacing = font.Spacing;
            int scaling = font.Scaling;

            mixed = IsUndefined(bold) || IsUndefined(italic) || IsUndefined(underline) ||
                    IsUndefined(color) || IsUndefined(subscript) || IsUndefined(superscript) ||
                    IsUndefined(strikeThrough) || IsUndefined(doubleStrikeThrough) ||
                    IsUndefined(smallCaps) || IsUndefined(allCaps) || IsUndefined(spacing) ||
                    IsUndefined(scaling) || IsUndefined(size) || string.IsNullOrEmpty(name);

            JObject state = new JObject();
            state["name"] = name ?? string.Empty;
            state["nameFarEast"] = nameFarEast ?? string.Empty;
            state["nameAscii"] = nameAscii ?? string.Empty;
            state["size"] = Convert.ToDouble(size, CultureInfo.InvariantCulture);
            state["bold"] = bold;
            state["italic"] = italic;
            state["underline"] = underline;
            state["color"] = color;
            state["subscript"] = subscript;
            state["superscript"] = superscript;
            state["strikeThrough"] = strikeThrough;
            state["doubleStrikeThrough"] = doubleStrikeThrough;
            state["smallCaps"] = smallCaps;
            state["allCaps"] = allCaps;
            state["spacing"] = Convert.ToDouble(spacing, CultureInfo.InvariantCulture);
            state["scaling"] = scaling;
            return state;
        }

        private static void ApplyWordFont(Word.Font font, JObject state)
        {
            if (font == null || state == null)
            {
                return;
            }

            string name = ReadString(state, "name");
            if (name.Length > 0)
            {
                font.Name = name;
            }
            string nameFarEast = ReadString(state, "nameFarEast");
            if (nameFarEast.Length > 0)
            {
                font.NameFarEast = nameFarEast;
            }
            string nameAscii = ReadString(state, "nameAscii");
            if (nameAscii.Length > 0)
            {
                font.NameAscii = nameAscii;
            }
            if (state["size"] != null)
            {
                font.Size = (float)ReadDouble(state, "size");
            }
            if (state["bold"] != null)
            {
                font.Bold = ReadInt(state, "bold");
            }
            if (state["italic"] != null)
            {
                font.Italic = ReadInt(state, "italic");
            }
            if (state["underline"] != null)
            {
                font.Underline = (Word.WdUnderline)ReadInt(state, "underline");
            }
            if (state["color"] != null)
            {
                font.Color = (Word.WdColor)ReadInt(state, "color");
            }
            if (state["subscript"] != null)
            {
                font.Subscript = ReadInt(state, "subscript");
            }
            if (state["superscript"] != null)
            {
                font.Superscript = ReadInt(state, "superscript");
            }
            if (state["strikeThrough"] != null)
            {
                font.StrikeThrough = ReadInt(state, "strikeThrough");
            }
            if (state["doubleStrikeThrough"] != null)
            {
                font.DoubleStrikeThrough = ReadInt(state, "doubleStrikeThrough");
            }
            if (state["smallCaps"] != null)
            {
                font.SmallCaps = ReadInt(state, "smallCaps");
            }
            if (state["allCaps"] != null)
            {
                font.AllCaps = ReadInt(state, "allCaps");
            }
            if (state["spacing"] != null)
            {
                font.Spacing = (float)ReadDouble(state, "spacing");
            }
            if (state["scaling"] != null)
            {
                font.Scaling = ReadInt(state, "scaling");
            }
        }

        private static object[,] ToObjectMatrix(JArray matrix, int rows, int columns)
        {
            if (!MatrixHasShape(matrix, rows, columns))
            {
                return null;
            }

            object[,] values = new object[rows, columns];
            for (int rowIndex = 0; rowIndex < rows; rowIndex++)
            {
                JArray row = matrix[rowIndex] as JArray;
                for (int columnIndex = 0; columnIndex < columns; columnIndex++)
                {
                    values[rowIndex, columnIndex] = ToCellValue(row[columnIndex]);
                }
            }
            return values;
        }

        private static JArray ToJagged(object raw, int rows, int columns)
        {
            if (rows <= 0 || columns <= 0)
            {
                return null;
            }

            Array array = raw as Array;
            if (array == null)
            {
                if (rows != 1 || columns != 1)
                {
                    return null;
                }
                JArray row = new JArray();
                row.Add(ToToken(raw));
                JArray single = new JArray();
                single.Add(row);
                return single;
            }

            if (array.Rank == 1)
            {
                int length = array.GetLength(0);
                int lower = array.GetLowerBound(0);
                if (rows == 1 && length >= columns)
                {
                    JArray row = new JArray();
                    for (int columnIndex = 0; columnIndex < columns; columnIndex++)
                    {
                        row.Add(ToToken(array.GetValue(lower + columnIndex)));
                    }
                    JArray single = new JArray();
                    single.Add(row);
                    return single;
                }
                if (columns == 1 && length >= rows)
                {
                    JArray result = new JArray();
                    for (int rowIndex = 0; rowIndex < rows; rowIndex++)
                    {
                        JArray row = new JArray();
                        row.Add(ToToken(array.GetValue(lower + rowIndex)));
                        result.Add(row);
                    }
                    return result;
                }
                return null;
            }

            if (array.Rank < 2)
            {
                return null;
            }

            int lowerRow = array.GetLowerBound(0);
            int lowerColumn = array.GetLowerBound(1);
            int height = array.GetUpperBound(0) - lowerRow + 1;
            int width = array.GetUpperBound(1) - lowerColumn + 1;
            if (height < rows || width < columns)
            {
                return null;
            }

            JArray matrix = new JArray();
            for (int rowIndex = 0; rowIndex < rows; rowIndex++)
            {
                JArray row = new JArray();
                for (int columnIndex = 0; columnIndex < columns; columnIndex++)
                {
                    row.Add(ToToken(array.GetValue(lowerRow + rowIndex, lowerColumn + columnIndex)));
                }
                matrix.Add(row);
            }
            return matrix;
        }

        private static JToken ToToken(object value)
        {
            if (value == null || value is DBNull)
            {
                return JValue.CreateNull();
            }
            if (value is DateTime)
            {
                return new JValue(((DateTime)value).ToOADate());
            }
            if (value is string)
            {
                return new JValue((string)value);
            }
            if (value is bool)
            {
                return new JValue((bool)value);
            }
            if (value is double || value is float || value is decimal ||
                value is byte || value is short || value is int || value is long ||
                value is ushort || value is uint || value is ulong)
            {
                try
                {
                    return new JValue(Convert.ToDouble(value, CultureInfo.InvariantCulture));
                }
                catch
                {
                    return new JValue(value.ToString());
                }
            }
            return new JValue(value.ToString());
        }

        private static object ToCellValue(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                return DBNull.Value;
            }
            switch (token.Type)
            {
                case JTokenType.String:
                    return token.Value<string>();
                case JTokenType.Boolean:
                    return token.Value<bool>();
                case JTokenType.Date:
                    return token.Value<DateTime>().ToOADate();
                case JTokenType.Integer:
                case JTokenType.Float:
                    return token.Value<double>();
                default:
                    return token.ToString();
            }
        }

        private static string ReadCellText(JArray row, int columnIndex)
        {
            if (row == null || columnIndex < 0 || columnIndex >= row.Count)
            {
                return null;
            }
            JToken token = row[columnIndex];
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }
            return token.Type == JTokenType.String ? token.Value<string>() : null;
        }

        private static string Sha256(string value)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(bytes);
                StringBuilder builder = new StringBuilder(hash.Length * 2);
                foreach (byte item in hash)
                {
                    builder.Append(item.ToString("x2", CultureInfo.InvariantCulture));
                }
                return builder.ToString();
            }
        }

        private static string NormalizeSummary(string summary)
        {
            return string.IsNullOrWhiteSpace(summary) ? "未命名操作" : summary.Trim();
        }

        private static string ResolveKind(SnapshotCapture before, SnapshotCapture after)
        {
            if (before != null && before.Success && before.Snapshot != null)
            {
                string kind = ReadString(before.Snapshot, "kind");
                if (kind.Length > 0)
                {
                    return kind;
                }
            }
            if (after != null && after.Success && after.Snapshot != null)
            {
                string kind = ReadString(after.Snapshot, "kind");
                if (kind.Length > 0)
                {
                    return kind;
                }
            }
            return "unknown";
        }

        private static SnapshotUndoNode CreateUnrestorable(string kind, string summary, string reason)
        {
            string safeKind = string.IsNullOrWhiteSpace(kind) ? "unknown" : kind;
            string safeReason = string.IsNullOrWhiteSpace(reason) ? "未登记可撤销节点：原因未知。" : reason;
            string appType = string.Equals(safeKind, ExcelKind, StringComparison.Ordinal)
                ? "excel"
                : (string.Equals(safeKind, WordKind, StringComparison.Ordinal) ? "word" : string.Empty);
            return new SnapshotUndoNode(
                SnapshotUndoNode.NewId(),
                safeKind,
                appType,
                string.Empty,
                string.Empty,
                string.Empty,
                NormalizeSummary(summary),
                false,
                safeReason,
                DateTime.UtcNow,
                null);
        }

        private static string ReadString(JObject json, string name)
        {
            if (json == null)
            {
                return string.Empty;
            }
            JToken token = json[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return string.Empty;
            }
            if (token.Type == JTokenType.String)
            {
                return token.Value<string>() ?? string.Empty;
            }
            return token.ToString();
        }

        private static bool TryReadString(JObject json, string name, out string value)
        {
            value = null;
            if (json == null)
            {
                return false;
            }
            JToken token = json[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return false;
            }
            value = token.Type == JTokenType.String ? token.Value<string>() : token.ToString();
            return true;
        }

        private static int ReadInt(JObject json, string name)
        {
            int value;
            return TryReadInt(json, name, out value) ? value : 0;
        }

        private static bool TryReadInt(JObject json, string name, out int value)
        {
            value = 0;
            if (json == null)
            {
                return false;
            }
            JToken token = json[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return false;
            }
            try
            {
                value = token.Value<int>();
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static double ReadDouble(JObject json, string name)
        {
            if (json == null)
            {
                return 0d;
            }
            JToken token = json[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return 0d;
            }
            try
            {
                return token.Value<double>();
            }
            catch
            {
                return 0d;
            }
        }

        private static bool ReadBool(JObject json, string name)
        {
            if (json == null)
            {
                return false;
            }
            JToken token = json[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return false;
            }
            try
            {
                return token.Value<bool>();
            }
            catch
            {
                return false;
            }
        }

        private static bool IsUndefined(int value)
        {
            return value == WdUndefinedValue;
        }

        private static bool IsUndefined(float value)
        {
            return Math.Abs(value - (float)WdUndefinedValue) < 0.5f;
        }
    }
}
