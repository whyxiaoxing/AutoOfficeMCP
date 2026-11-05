using System;
using Newtonsoft.Json.Linq;
using Office;
using Office.Undo;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>
    /// 写操作撤销采集适配器：
    /// 1) 写前采集区域快照；
    /// 2) 写后采集并交给 T02 的 SnapshotUndoEngine.BuildNode；
    /// 3) 只有 Restorable=true 的节点才入 OfficeSession.UndoStack，并把 Id 写入 Observation；
    /// 4) 不可还原时在 Data.undoReason 中写真实原因，绝不伪造 undoUnits。
    /// </summary>
    internal sealed class ExcelUndoScope
    {
        private readonly SnapshotCapture _before;

        private ExcelUndoScope(SnapshotCapture before)
        {
            _before = before;
        }

        public static ExcelUndoScope Capture(
            Excel.Workbook workbook,
            Excel.Worksheet worksheet,
            Excel.Range range,
            string address)
        {
            SnapshotCapture before;
            if (workbook == null || worksheet == null || range == null)
            {
                before = SnapshotCapture.Failed("执行前目标区域不存在或不可用，未采集撤销快照。");
            }
            else
            {
                string snapshotAddress = string.IsNullOrWhiteSpace(address)
                    ? range.AddressLocal()
                    : address;
                before = SnapshotUndoEngine.CaptureExcel(workbook, worksheet, range, snapshotAddress);
            }
            return new ExcelUndoScope(before);
        }

        public static ExcelUndoScope Failed(string reason)
        {
            return new ExcelUndoScope(SnapshotCapture.Failed(reason));
        }

        public SnapshotUndoNode Complete(
            Excel.Workbook workbook,
            Excel.Worksheet worksheet,
            Excel.Range afterRange,
            string afterAddress,
            bool structureStable,
            string summary,
            Observation observation,
            JObject data,
            string fallbackReason)
        {
            if (data == null)
            {
                data = new JObject();
            }

            SnapshotCapture after = null;
            if (afterRange != null)
            {
                string snapshotAddress = string.IsNullOrWhiteSpace(afterAddress)
                    ? afterRange.AddressLocal()
                    : afterAddress;
                after = SnapshotUndoEngine.CaptureExcel(workbook, worksheet, afterRange, snapshotAddress);
            }

            SnapshotUndoNode node = SnapshotUndoEngine.BuildNode(_before, after, structureStable, summary);
            if (node == null)
            {
                data["undoRestorable"] = false;
                data["undoReason"] = "撤销节点构建失败，未登记可撤销节点。";
                return null;
            }

            if (node.Restorable)
            {
                OfficeSession.UndoStack.Push(node);
                if (observation != null)
                {
                    observation.AddUndoUnit(node.Id);
                }
                data["undoRestorable"] = true;
                data["undoScope"] = "区域值/公式/数字格式";
                if (!string.IsNullOrWhiteSpace(fallbackReason))
                {
                    data["undoCaveat"] = fallbackReason.Trim();
                }
                return node;
            }

            string reason = string.IsNullOrWhiteSpace(node.Reason) ? "未登记可撤销节点。" : node.Reason;
            if (!string.IsNullOrWhiteSpace(fallbackReason) &&
                reason.IndexOf("内容未发生变化", StringComparison.Ordinal) >= 0)
            {
                reason = fallbackReason.Trim();
            }
            data["undoRestorable"] = false;
            data["undoReason"] = reason;
            return node;
        }

        /// <summary>对没有可采集区域的结构性/外部副作用操作，如实记录不可撤销原因。</summary>
        public static void RecordUnavailable(JObject data, string reason)
        {
            if (data == null)
            {
                return;
            }
            data["undoRestorable"] = false;
            data["undoReason"] = string.IsNullOrWhiteSpace(reason)
                ? "该操作无法用区域快照完整还原，未登记可撤销节点。"
                : reason.Trim();
        }
    }
}
