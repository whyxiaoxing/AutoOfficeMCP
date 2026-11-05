using System;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    /// <summary>
    /// 写前/写后指纹：文档文本哈希、段落/表格/图片/批注/修订/域/公式计数、选区边界与选区格式。
    /// 这些快照会原样进入 OperationResult.Data，供调用方确认 changed 是真实观察而不是自报。
    /// </summary>
    internal static class WordDocumentState
    {
        public static JObject Capture(Word.Application application, Word.Document document)
        {
            JObject state = new JObject();
            if (document == null)
            {
                state["error"] = "没有活动文档。";
                return state;
            }

            Word.Range content = null;
            try
            {
                state["documentName"] = document.Name ?? string.Empty;
                state["documentPath"] = document.FullName ?? string.Empty;
                state["readOnly"] = document.ReadOnly;
                state["protectionType"] = document.ProtectionType.ToString();

                content = document.Content;
                string text = content == null ? string.Empty : (content.Text ?? string.Empty);
                state["documentTextHash"] = WordComUtil.HashText(text);
                state["documentCharacterCount"] = text.Length;
                state["documentTextPreview"] = WordComUtil.NormalizePreview(text, 160);

                state["paragraphCount"] = WordComUtil.GetParagraphCount(document);
                state["tableCount"] = CountTables(document);
                state["inlineShapeCount"] = CountInlineShapes(document);
                state["floatingShapeCount"] = CountShapes(document);
                state["commentCount"] = CountComments(document);
                state["revisionCount"] = CountRevisions(document);
                state["fieldCount"] = CountFields(document);
                state["equationCount"] = CountEquations(document);
                state["sectionCount"] = CountSections(document);

                JObject selection = WordComUtil.GetSelectionSnapshot(application, document);
                state["selection"] = selection;
                state["selectionStart"] = ReadInt(selection, "start");
                state["selectionEnd"] = ReadInt(selection, "end");
                state["selectionTextHash"] = ReadString(selection, "textHash");
                state["selectionFormatHash"] = WordComUtil.HashText(
                    (selection["format"] == null ? string.Empty : selection["format"].ToString()));
                return state;
            }
            catch (Exception ex)
            {
                state["error"] = OfficeErrorMapper.DescribeException(ex, OfficeApp.Word);
                return state;
            }
            finally
            {
                ComHelper.Release(content);
            }
        }

        public static bool HasChanged(JObject before, JObject after)
        {
            if (HasContentChanged(before, after))
            {
                return true;
            }
            if (before == null || after == null)
            {
                return false;
            }

            string[] selectionKeys = new string[]
            {
                "selectionStart",
                "selectionEnd",
                "selectionTextHash",
                "selectionFormatHash"
            };
            foreach (string key in selectionKeys)
            {
                if (!JToken.DeepEquals(before[key], after[key]))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 只比较文档内容/结构指纹，忽略光标选区变化。
        /// 查找类工具即使未替换也会移动选区，不能用选区差异伪造 Observation.changed=true。
        /// </summary>
        public static bool HasContentChanged(JObject before, JObject after)
        {
            if (before == null || after == null)
            {
                return false;
            }

            string[] keys = new string[]
            {
                "documentTextHash",
                "paragraphCount",
                "tableCount",
                "inlineShapeCount",
                "floatingShapeCount",
                "commentCount",
                "revisionCount",
                "fieldCount",
                "equationCount",
                "sectionCount"
            };
            foreach (string key in keys)
            {
                if (!JToken.DeepEquals(before[key], after[key]))
                {
                    return true;
                }
            }
            return false;
        }

        public static JObject BuildDiff(JObject before, JObject after)
        {
            JObject diff = new JObject();
            if (before == null || after == null)
            {
                diff["available"] = false;
                return diff;
            }

            diff["documentTextChanged"] = !JToken.DeepEquals(
                before["documentTextHash"], after["documentTextHash"]);
            diff["selectionChanged"] = !JToken.DeepEquals(
                before["selectionTextHash"], after["selectionTextHash"]);
            diff["selectionFormatChanged"] = !JToken.DeepEquals(
                before["selectionFormatHash"], after["selectionFormatHash"]);
            diff["paragraphDelta"] = ReadInt(after, "paragraphCount") - ReadInt(before, "paragraphCount");
            diff["tableDelta"] = ReadInt(after, "tableCount") - ReadInt(before, "tableCount");
            diff["inlineShapeDelta"] = ReadInt(after, "inlineShapeCount") - ReadInt(before, "inlineShapeCount");
            diff["floatingShapeDelta"] = ReadInt(after, "floatingShapeCount") - ReadInt(before, "floatingShapeCount");
            diff["commentDelta"] = ReadInt(after, "commentCount") - ReadInt(before, "commentCount");
            diff["revisionDelta"] = ReadInt(after, "revisionCount") - ReadInt(before, "revisionCount");
            diff["fieldDelta"] = ReadInt(after, "fieldCount") - ReadInt(before, "fieldCount");
            diff["equationDelta"] = ReadInt(after, "equationCount") - ReadInt(before, "equationCount");
            return diff;
        }

        private static int CountTables(Word.Document document)
        {
            Word.Tables collection = null;
            try
            {
                collection = document.Tables;
                return collection == null ? 0 : collection.Count;
            }
            catch
            {
                return 0;
            }
            finally
            {
                ComHelper.Release(collection);
            }
        }

        private static int CountInlineShapes(Word.Document document)
        {
            Word.InlineShapes collection = null;
            try
            {
                collection = document.InlineShapes;
                return collection == null ? 0 : collection.Count;
            }
            catch
            {
                return 0;
            }
            finally
            {
                ComHelper.Release(collection);
            }
        }

        private static int CountShapes(Word.Document document)
        {
            Word.Shapes collection = null;
            try
            {
                collection = document.Shapes;
                return collection == null ? 0 : collection.Count;
            }
            catch
            {
                return 0;
            }
            finally
            {
                ComHelper.Release(collection);
            }
        }

        private static int CountComments(Word.Document document)
        {
            Word.Comments collection = null;
            try
            {
                collection = document.Comments;
                return collection == null ? 0 : collection.Count;
            }
            catch
            {
                return 0;
            }
            finally
            {
                ComHelper.Release(collection);
            }
        }

        private static int CountRevisions(Word.Document document)
        {
            Word.Revisions collection = null;
            try
            {
                collection = document.Revisions;
                return collection == null ? 0 : collection.Count;
            }
            catch
            {
                return 0;
            }
            finally
            {
                ComHelper.Release(collection);
            }
        }

        private static int CountFields(Word.Document document)
        {
            Word.Fields collection = null;
            try
            {
                collection = document.Fields;
                return collection == null ? 0 : collection.Count;
            }
            catch
            {
                return 0;
            }
            finally
            {
                ComHelper.Release(collection);
            }
        }

        private static int CountEquations(Word.Document document)
        {
            Word.OMaths collection = null;
            try
            {
                collection = document.OMaths;
                return collection == null ? 0 : collection.Count;
            }
            catch
            {
                return 0;
            }
            finally
            {
                ComHelper.Release(collection);
            }
        }

        private static int CountSections(Word.Document document)
        {
            Word.Sections collection = null;
            try
            {
                collection = document.Sections;
                return collection == null ? 0 : collection.Count;
            }
            catch
            {
                return 0;
            }
            finally
            {
                ComHelper.Release(collection);
            }
        }

        private static int ReadInt(JObject json, string name)
        {
            if (json == null)
            {
                return 0;
            }
            JToken token = json[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return 0;
            }
            try
            {
                return token.Value<int>();
            }
            catch
            {
                return 0;
            }
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
            return token.Type == JTokenType.String ? token.Value<string>() : token.ToString();
        }
    }
}
