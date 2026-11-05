using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    /// <summary>
    /// T06 P3 补充：竖排文字/文字方向。Word PIA 15 的 PageSetup 类型未暴露 TextDirection，
    /// 这里用 IDispatch 晚期绑定访问，兼容 Office 2016+/Microsoft 365。
    /// </summary>
    public sealed partial class WordExtensionsExecutor
    {
        private void RegisterP3MoreHandlers(
            Dictionary<string, Func<Word.Application, JObject, OperationResult>> handlers)
        {
            handlers[WordExtensionsToolNames.VerticalText] = HandleVerticalText;
        }

        private OperationResult HandleVerticalText(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string action = (args.GetOptionalString("action", "get") ?? "get").Trim().ToLowerInvariant();
            if (action != "get" && action != "set")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "action 只支持 get/set，收到「" + action + "」。");
            }

            string target = (args.GetOptionalString("target", "shape") ?? "shape").Trim().ToLowerInvariant();
            string shapeName = args.GetOptionalString("shapeName", null) ?? string.Empty;
            int? requestedShape = args.GetOptionalInt("shapeIndex");
            int shapeIndex = requestedShape ?? 0;
            if (target != "shape" && target != "document")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "target 只支持 shape/document，收到「" + target + "」。");
            }
            if (target == "document" && shapeName.Length == 0 && shapeIndex <= 0)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "Word COM（PIA 15 / Office 2024）不暴露 PageSetup.TextDirection，" +
                    "无法设置文档级竖排文字；请用 target=shape 并指定 shapeName/shapeIndex 对文本框或形状设置方向。");
            }
            if (shapeName.Length == 0 && shapeIndex <= 0)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "shape 模式需要 shapeName 或 shapeIndex（1 基）。");
            }

            if (action == "get")
            {
                return RunReadTool(application, delegate(Word.Document document)
                {
                    return ReadShapeVerticalText(document, shapeName, shapeIndex);
                });
            }

            string directionText = (args.GetOptionalString("direction", "downward") ?? "downward")
                .Trim().ToLowerInvariant();
            int orientationValue;
            if (!TryParseShapeOrientation(directionText, out orientationValue))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "shape 的 direction 仅支持 horizontal/upward/downward/vertical/verticalFarEast/" +
                    "horizontalRotatedFarEast，收到「" + directionText + "」。");
            }

            return RunWriteTool(
                application,
                WordExtensionsToolNames.VerticalText,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.None("Word 形状", "Word:Shape");
                },
                delegate(Word.Document document, WordWriteTarget writeTarget)
                {
                    return SetShapeVerticalText(document, shapeName, shapeIndex, orientationValue, directionText);
                });
        }

        private static OperationResult ReadShapeVerticalText(
            Word.Document document,
            string shapeName,
            int shapeIndex)
        {
            Word.Shapes shapes = null;
            Word.Shape shape = null;
            object textFrame = null;
            try
            {
                shapes = document.Shapes;
                string error;
                if (!TryGetShape(shapes, shapeName, shapeIndex, out shape, out error))
                {
                    return OperationResult.Failed(OfficeErrorCode.INVALID_PARAMS, error);
                }
                textFrame = GetTextFrame(shape, out error);
                if (textFrame == null)
                {
                    return OperationResult.Failed(OfficeErrorCode.INVALID_PARAMS, error);
                }

                int orientation = ReadComInt(textFrame, "Orientation");
                JObject data = new JObject();
                data["document"] = document.Name;
                data["shapeName"] = Convert.ToString(shape.Name, CultureInfo.InvariantCulture) ?? string.Empty;
                data["orientation"] = orientation;
                data["direction"] = ShapeOrientationText(orientation);
                return ReadSucceed(
                    "读取形状文字方向：" + data.Value<string>("direction") + "。",
                    data,
                    "Word:Shape#" + data.Value<string>("shapeName"));
            }
            finally
            {
                ComHelper.Release(textFrame, shape, shapes);
            }
        }

        private static WordWriteResult SetShapeVerticalText(
            Word.Document document,
            string shapeName,
            int shapeIndex,
            int orientationValue,
            string directionText)
        {
            Word.Shapes shapes = null;
            Word.Shape shape = null;
            object textFrame = null;
            try
            {
                shapes = document.Shapes;
                string error;
                if (!TryGetShape(shapes, shapeName, shapeIndex, out shape, out error))
                {
                    return WordWriteResult.Fail(OfficeErrorCode.INVALID_PARAMS, error, null);
                }
                textFrame = GetTextFrame(shape, out error);
                if (textFrame == null)
                {
                    return WordWriteResult.Fail(OfficeErrorCode.INVALID_PARAMS, error, null);
                }

                int before = ReadComInt(textFrame, "Orientation");
                TrySetComProperty(textFrame, "Orientation", orientationValue);
                int after = ReadComInt(textFrame, "Orientation");
                bool changed = before != after;
                if (after != orientationValue)
                {
                    return WordWriteResult.Fail(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "设置形状「" + shapeName + "」文字方向失败：写入 " + orientationValue +
                        "，回读 " + after + "。",
                        null,
                        changed);
                }

                JObject data = new JObject();
                data["document"] = document.Name;
                data["shapeName"] = Convert.ToString(shape.Name, CultureInfo.InvariantCulture) ?? string.Empty;
                data["requestedDirection"] = directionText;
                data["before"] = ShapeOrientationText(before);
                data["after"] = ShapeOrientationText(after);
                data["undoNote"] = "形状文字方向属于版式设置，区域文本快照不覆盖 TextFrame 属性。";
                WordWriteResult result = WordWriteResult.Ok(
                    "形状文字方向设置完成：" + ShapeOrientationText(after) + "。",
                    data,
                    changed);
                result.AsStructural("形状文字方向属于版式设置，区域文本快照不覆盖 TextFrame 属性。");
                return result;
            }
            finally
            {
                ComHelper.Release(textFrame, shape, shapes);
            }
        }

        private static bool TryGetShape(
            Word.Shapes shapes,
            string shapeName,
            int shapeIndex,
            out Word.Shape shape,
            out string error)
        {
            shape = null;
            error = string.Empty;
            try
            {
                if (!string.IsNullOrWhiteSpace(shapeName))
                {
                    shape = shapes[shapeName];
                }
                else if (shapeIndex >= 1 && shapeIndex <= shapes.Count)
                {
                    shape = shapes[shapeIndex];
                }
                else
                {
                    error = "shapeIndex=" + shapeIndex.ToString(CultureInfo.InvariantCulture) +
                            " 超出范围；当前文档共 " + shapes.Count.ToString(CultureInfo.InvariantCulture) +
                            " 个形状。";
                    return false;
                }
            }
            catch (Exception ex)
            {
                error = "找不到形状「" + shapeName + "」：" + ex.Message;
                return false;
            }
            if (shape == null)
            {
                error = "找不到形状「" + shapeName + "」。";
                return false;
            }
            return true;
        }

        private static object GetTextFrame(Word.Shape shape, out string error)
        {
            error = string.Empty;
            try
            {
                object textFrame = shape.GetType().InvokeMember(
                    "TextFrame",
                    BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    shape,
                    null);
                if (textFrame == null)
                {
                    error = "形状「" + Convert.ToString(shape.Name, CultureInfo.InvariantCulture) +
                            "」没有 TextFrame（可能不是文本框/自选图形）。";
                }
                return textFrame;
            }
            catch (Exception ex)
            {
                error = "形状「" + Convert.ToString(shape.Name, CultureInfo.InvariantCulture) +
                        "」不支持 TextFrame：" + ex.Message;
                return null;
            }
        }

        private static bool TryParseShapeOrientation(string text, out int value)
        {
            switch (text)
            {
                case "horizontal":
                    value = 1;
                    return true;
                case "upward":
                case "up":
                    value = 2;
                    return true;
                case "downward":
                case "down":
                    value = 3;
                    return true;
                case "verticalfareast":
                    value = 4;
                    return true;
                case "vertical":
                    value = 5;
                    return true;
                case "horizontalrotatedfareast":
                    value = 6;
                    return true;
                default:
                    value = -1;
                    return false;
            }
        }

        private static string ShapeOrientationText(int value)
        {
            switch (value)
            {
                case 1:
                    return "horizontal";
                case 2:
                    return "upward";
                case 3:
                    return "downward";
                case 4:
                    return "verticalFarEast";
                case 5:
                    return "vertical";
                case 6:
                    return "horizontalRotatedFarEast";
                default:
                    return "unknown(" + value.ToString(CultureInfo.InvariantCulture) + ")";
            }
        }

        private static OperationResult ReadVerticalText(Word.Document document, int sectionIndex)
        {
            Word.Sections sections = null;
            List<Word.Section> targets = new List<Word.Section>();
            JArray items = new JArray();
            try
            {
                sections = document.Sections;
                int count = sections.Count;
                if (!TryCollectSections(sections, count, sectionIndex, targets))
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.INVALID_PARAMS,
                        "sectionIndex=" + sectionIndex.ToString(CultureInfo.InvariantCulture) +
                        " 超出范围；当前文档共 " + count.ToString(CultureInfo.InvariantCulture) + " 节。");
                }

                foreach (Word.Section section in targets)
                {
                    Word.PageSetup setup = null;
                    try
                    {
                        setup = section.PageSetup;
                        int direction = ReadComInt(setup, "TextDirection");
                        JObject item = new JObject();
                        item["sectionIndex"] = SectionNumber(section);
                        item["textDirection"] = direction;
                        item["direction"] = DirectionText(direction);
                        items.Add(item);
                    }
                    finally
                    {
                        ComHelper.Release(setup);
                    }
                }

                JObject data = new JObject();
                data["document"] = document.Name;
                data["sectionCount"] = sections.Count;
                data["sections"] = items;
                return ReadSucceed(
                    "读取文字方向：共 " + items.Count + " 节。",
                    data,
                    "Word:Document#" + document.Name);
            }
            finally
            {
                foreach (Word.Section section in targets)
                {
                    ComHelper.Release(section);
                }
                ComHelper.Release(sections);
            }
        }

        private static WordWriteResult SetVerticalText(
            Word.Document document,
            int sectionIndex,
            int directionValue,
            string directionText)
        {
            Word.Sections sections = null;
            List<Word.Section> targets = new List<Word.Section>();
            try
            {
                sections = document.Sections;
                int count = sections.Count;
                if (!TryCollectSections(sections, count, sectionIndex, targets))
                {
                    return WordWriteResult.Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "sectionIndex=" + sectionIndex.ToString(CultureInfo.InvariantCulture) +
                        " 超出范围；当前文档共 " + count.ToString(CultureInfo.InvariantCulture) + " 节。",
                        null);
                }

                int changedSections = 0;
                JArray items = new JArray();
                foreach (Word.Section section in targets)
                {
                    Word.PageSetup setup = null;
                    try
                    {
                        setup = section.PageSetup;
                        int before = ReadComInt(setup, "TextDirection");
                        TrySetComProperty(setup, "TextDirection", directionValue);
                        int after = ReadComInt(setup, "TextDirection");
                        if (after != directionValue)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INTERNAL_ERROR,
                                "设置第 " + SectionNumber(section).ToString(CultureInfo.InvariantCulture) +
                                " 节文字方向失败：写入 " + directionValue + "，回读 " + after + "。",
                                null,
                                before != after);
                        }
                        if (before != after)
                        {
                            changedSections++;
                        }
                        JObject item = new JObject();
                        item["sectionIndex"] = SectionNumber(section);
                        item["before"] = DirectionText(before);
                        item["after"] = DirectionText(after);
                        items.Add(item);
                    }
                    finally
                    {
                        ComHelper.Release(setup);
                    }
                }

                JObject data = new JObject();
                data["document"] = document.Name;
                data["requestedDirection"] = directionText;
                data["changedSectionCount"] = changedSections;
                data["sections"] = items;
                data["undoNote"] = "文字方向属于节版式设置，区域文本快照不覆盖节属性。";
                WordWriteResult result = WordWriteResult.Ok(
                    "文字方向设置完成：变更 " + changedSections + " 节。",
                    data,
                    changedSections > 0);
                result.AsStructural("文字方向属于节版式设置，区域文本快照不覆盖节属性。");
                return result;
            }
            finally
            {
                foreach (Word.Section section in targets)
                {
                    ComHelper.Release(section);
                }
                ComHelper.Release(sections);
            }
        }

        private static bool TryCollectSections(
            Word.Sections sections,
            int count,
            int sectionIndex,
            List<Word.Section> targets)
        {
            if (sectionIndex <= 0)
            {
                for (int index = 1; index <= count; index++)
                {
                    targets.Add((Word.Section)sections[index]);
                }
                return true;
            }
            if (sectionIndex > count)
            {
                return false;
            }
            targets.Add((Word.Section)sections[sectionIndex]);
            return true;
        }

        private static int SectionNumber(Word.Section section)
        {
            if (section == null)
            {
                return 0;
            }
            try
            {
                return Convert.ToInt32(section.Index, CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[word-ext] 读取节索引失败：" + ex.Message);
                return 0;
            }
        }

        private static int ReadComInt(object target, string propertyName)
        {
            if (target == null)
            {
                return -1;
            }
            try
            {
                object value = target.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    target,
                    null);
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[word-ext] 读取 COM 属性 " + propertyName + " 失败：" + ex.Message);
                return -1;
            }
        }

        private static bool TryParseDirection(string text, out int value)
        {
            switch (text)
            {
                case "lefttoright":
                case "ltr":
                case "horizontal":
                    value = 0;
                    return true;
                case "righttoleft":
                case "rtl":
                    value = 1;
                    return true;
                case "upward":
                case "up":
                    value = 2;
                    return true;
                case "downward":
                case "down":
                case "vertical":
                    value = 3;
                    return true;
                default:
                    value = -1;
                    return false;
            }
        }

        private static string DirectionText(int value)
        {
            switch (value)
            {
                case 0:
                    return "leftToRight";
                case 1:
                    return "rightToLeft";
                case 2:
                    return "upward";
                case 3:
                    return "downward";
                default:
                    return "unknown(" + value.ToString(CultureInfo.InvariantCulture) + ")";
            }
        }
    }
}
