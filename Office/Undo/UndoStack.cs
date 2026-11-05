using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using Newtonsoft.Json.Linq;

namespace Office.Undo
{
    /// <summary>
    /// 一个可还原（或不可还原）的撤销节点。
    /// Id 在创建时生成一次，之后在整个会话栈中保持稳定，Observation.undoUnits 引用它。
    /// </summary>
    public sealed class SnapshotUndoNode
    {
        private static int _sequence;

        public string Id { get; private set; }

        public string Kind { get; private set; }

        public string AppType { get; private set; }

        public string DocumentName { get; private set; }

        public string SheetName { get; private set; }

        public string Address { get; private set; }

        public string Summary { get; private set; }

        public bool Restorable { get; private set; }

        public string Reason { get; private set; }

        public DateTime CreatedUtc { get; private set; }

        /// <summary>Restorable=true 时非空；不可还原节点不保存快照，避免被误回放。</summary>
        public JObject Snapshot { get; private set; }

        internal SnapshotUndoNode(
            string id,
            string kind,
            string appType,
            string documentName,
            string sheetName,
            string address,
            string summary,
            bool restorable,
            string reason,
            DateTime createdUtc,
            JObject snapshot)
        {
            Id = id;
            Kind = kind ?? string.Empty;
            AppType = appType ?? string.Empty;
            DocumentName = documentName ?? string.Empty;
            SheetName = sheetName ?? string.Empty;
            Address = address ?? string.Empty;
            Summary = summary ?? string.Empty;
            Restorable = restorable;
            Reason = reason ?? string.Empty;
            CreatedUtc = createdUtc;
            Snapshot = snapshot;
        }

        internal static string NewId()
        {
            int sequence = Interlocked.Increment(ref _sequence);
            return "undo-" +
                   DateTime.UtcNow.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture) +
                   "-" + sequence.ToString("D4", CultureInfo.InvariantCulture);
        }

        public JObject ToJObject()
        {
            JObject json = new JObject();
            json["id"] = Id;
            json["kind"] = Kind;
            json["appType"] = AppType;
            json["document"] = DocumentName;
            json["sheet"] = SheetName;
            json["address"] = Address;
            json["summary"] = Summary;
            json["restorable"] = Restorable;
            json["reason"] = Reason;
            json["createdUtc"] = CreatedUtc.ToString("o", CultureInfo.InvariantCulture);
            if (Snapshot == null)
            {
                json["snapshot"] = JValue.CreateNull();
            }
            else
            {
                json["snapshot"] = Snapshot.DeepClone();
            }
            return json;
        }

        /// <summary>从持久化 JSON 还原节点；Id 缺失时重新生成。</summary>
        public static SnapshotUndoNode FromJObject(JObject json)
        {
            if (json == null)
            {
                return null;
            }

            string id = json.Value<string>("id");
            if (string.IsNullOrWhiteSpace(id))
            {
                id = NewId();
            }

            DateTime createdUtc;
            string createdText = json.Value<string>("createdUtc");
            if (string.IsNullOrWhiteSpace(createdText) ||
                !DateTime.TryParse(createdText, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out createdUtc))
            {
                createdUtc = DateTime.UtcNow;
            }

            JObject snapshot = json["snapshot"] as JObject;
            SnapshotUndoNode node = new SnapshotUndoNode(
                id,
                json.Value<string>("kind"),
                json.Value<string>("appType"),
                json.Value<string>("document"),
                json.Value<string>("sheet"),
                json.Value<string>("address"),
                json.Value<string>("summary"),
                json.Value<bool?>("restorable") ?? false,
                json.Value<string>("reason"),
                createdUtc,
                snapshot);
            return node;
        }
    }

    /// <summary>
    /// 会话级撤销节点栈。线程安全；只接受 Restorable=true 的节点，
    /// 从根上杜绝"不可还原节点被当成可撤销"的假成功。
    /// </summary>
    public sealed class UndoStack
    {
        private readonly object _syncRoot = new object();
        private readonly List<SnapshotUndoNode> _nodes = new List<SnapshotUndoNode>();

        public int Count
        {
            get
            {
                lock (_syncRoot)
                {
                    return _nodes.Count;
                }
            }
        }

        public void Push(SnapshotUndoNode node)
        {
            if (node == null)
            {
                throw new ArgumentNullException("node");
            }
            if (!node.Restorable)
            {
                throw new InvalidOperationException("不可还原节点不能入栈：" + node.Reason);
            }
            lock (_syncRoot)
            {
                _nodes.Add(node);
            }
        }

        public bool TryPeek(out SnapshotUndoNode node)
        {
            lock (_syncRoot)
            {
                if (_nodes.Count == 0)
                {
                    node = null;
                    return false;
                }
                node = _nodes[_nodes.Count - 1];
                return true;
            }
        }

        /// <summary>弹出栈顶节点。调用方应在回放成功后使用；回放失败请先 TryPeek 再决定是否保留。</summary>
        public bool TryTakeLast(out SnapshotUndoNode node)
        {
            lock (_syncRoot)
            {
                if (_nodes.Count == 0)
                {
                    node = null;
                    return false;
                }
                node = _nodes[_nodes.Count - 1];
                _nodes.RemoveAt(_nodes.Count - 1);
                return true;
            }
        }

        public SnapshotUndoNode Find(string nodeId)
        {
            if (string.IsNullOrWhiteSpace(nodeId))
            {
                return null;
            }
            lock (_syncRoot)
            {
                foreach (SnapshotUndoNode node in _nodes)
                {
                    if (string.Equals(node.Id, nodeId, StringComparison.Ordinal))
                    {
                        return node;
                    }
                }
            }
            return null;
        }

        public bool Remove(string nodeId)
        {
            if (string.IsNullOrWhiteSpace(nodeId))
            {
                return false;
            }
            lock (_syncRoot)
            {
                for (int index = _nodes.Count - 1; index >= 0; index--)
                {
                    if (string.Equals(_nodes[index].Id, nodeId, StringComparison.Ordinal))
                    {
                        _nodes.RemoveAt(index);
                        return true;
                    }
                }
            }
            return false;
        }

        public IList<SnapshotUndoNode> ToList()
        {
            lock (_syncRoot)
            {
                return new List<SnapshotUndoNode>(_nodes);
            }
        }

        public void Clear()
        {
            lock (_syncRoot)
            {
                _nodes.Clear();
            }
        }
    }
}
