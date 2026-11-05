using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    public sealed partial class WordToolExecutor
    {
        private void RegisterImageHandlers(
            Dictionary<string, Func<Word.Application, JObject, OperationResult>> handlers)
        {
            handlers[WordToolNames.InsertImage] = HandleInsertImage;
            handlers[WordToolNames.ImagesList] = HandleImagesList;
            handlers[WordToolNames.ImageResize] = HandleImageResize;
            handlers[WordToolNames.ImageLayout] = HandleImageLayout;
        }

        private OperationResult HandleInsertImage(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string filePath = args.RequireString("filePath");
            if (!File.Exists(filePath))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.FILE_NOT_FOUND,
                    "图片文件不存在：" + filePath);
            }
            string position = (args.GetOptionalString("position", "cursor") ?? "cursor")
                .Trim().ToLowerInvariant();
            bool floating = args.GetOptionalBool("floating", false);
            bool caption = args.GetOptionalBool("caption", false);
            bool lockAspect = args.GetOptionalBool("lockAspect", true);
            double? widthPt = args.GetOptionalDouble("widthPt");
            double? heightPt = args.GetOptionalDouble("heightPt");
            double? widthLimitPt = args.GetOptionalDouble("widthLimitPt");
            double? scalePercent = args.GetOptionalDouble("scalePercent");
            double? xPt = args.GetOptionalDouble("xPt");
            double? yPt = args.GetOptionalDouble("yPt");
            string altText = args.Has("altText") ? args.GetOptionalString("altText", string.Empty) : string.Empty;
            string wrapText = args.Has("wrap") ? args.GetOptionalString("wrap", null) : null;
            if (position != "cursor" && position != "start" && position != "end")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "position 只支持 cursor/start/end，收到 " + position + "。");
            }
            Word.WdWrapType wrapType = Word.WdWrapType.wdWrapSquare;
            if (!string.IsNullOrWhiteSpace(wrapText))
            {
                wrapType = WordImageUtil.ParseWrapType(wrapText);
            }
            if (floating && wrapType == Word.WdWrapType.wdWrapInline)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "floating=true 时 wrap 不能是 inline；请用 wrap=square/topbottom/tight/through/behind/front。");
            }

            return RunWriteTool(
                application,
                WordToolNames.InsertImage,
                false,
                delegate(Word.Document document)
                {
                    return SelectImageInsertionTarget(application, document, position);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Range insertionRange = null;
                    Word.InlineShapes inlineShapes = null;
                    Word.InlineShape inlineShape = null;
                    Word.Shapes shapes = null;
                    Word.Shape floatingShape = null;
                    Word.Range anchorRange = null;
                    JArray warnings = new JArray();
                    try
                    {
                        insertionRange = WordComUtil.CreateRange(document, target.Start, target.Start);
                        int beforeInline = CountInlineShapes(document);
                        int beforeFloating = CountFloatingShapes(document);
                        int? index = null;
                        double actualWidth;
                        double actualHeight;
                        string captionNote = string.Empty;
                        bool captionInserted = false;
                        double? requestedWidth = widthPt;
                        double? requestedHeight = heightPt;
                        double? requestedScale = scalePercent;
                        double? requestedLimit = widthLimitPt;

                        if (floating)
                        {
                            shapes = document.Shapes;
                            object linkToFile = false;
                            object saveWithDocument = true;
                            object left = (float)(xPt.HasValue ? xPt.Value : 0d);
                            object top = (float)(yPt.HasValue ? yPt.Value : 0d);
                            object width = Type.Missing;
                            object height = Type.Missing;
                            object anchorObject = insertionRange;
                            floatingShape = shapes.AddPicture(
                                filePath,
                                ref linkToFile,
                                ref saveWithDocument,
                                ref left,
                                ref top,
                                ref width,
                                ref height,
                                ref anchorObject);
                            if (floatingShape == null)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INTERNAL_ERROR,
                                    "Shapes.AddPicture 返回 null，图片未插入。",
                                    null,
                                    false);
                            }
                            WordSizeResult size = WordImageUtil.ComputeSize(
                                floatingShape.Width, floatingShape.Height,
                                widthPt, heightPt, scalePercent, widthLimitPt, lockAspect);
                            floatingShape.Width = (float)size.WidthPt;
                            floatingShape.Height = (float)size.HeightPt;
                            Word.WrapFormat wrap = null;
                            try
                            {
                                wrap = floatingShape.WrapFormat;
                                wrap.Type = wrapType;
                            }
                            finally
                            {
                                ComHelper.Release(wrap);
                            }
                            floatingShape.Left = (float)(xPt.HasValue ? xPt.Value : 0d);
                            floatingShape.Top = (float)(yPt.HasValue ? yPt.Value : 0d);
                            floatingShape.AlternativeText = altText ?? string.Empty;
                            actualWidth = floatingShape.Width;
                            actualHeight = floatingShape.Height;
                            index = WordComUtil.FindShapeIndex(document, floatingShape);
                            if (!string.IsNullOrWhiteSpace(size.Warning))
                            {
                                warnings.Add(size.Warning);
                            }
                            if (caption)
                            {
                                anchorRange = floatingShape.Anchor;
                                captionInserted = InsertImageCaption(
                                    document, anchorRange, altText, index.Value, warnings);
                                if (!captionInserted)
                                {
                                    captionNote = "浮动图片已插入，但图注插入失败。";
                                }
                            }
                            if (CountFloatingShapes(document) <= beforeFloating)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INTERNAL_ERROR,
                                    "浮动图片插入后 Shapes 数量未增加。",
                                    null,
                                    false);
                            }
                        }
                        else
                        {
                            inlineShapes = document.InlineShapes;
                            inlineShape = inlineShapes.AddPicture(
                                filePath, false, true, insertionRange);
                            if (inlineShape == null)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INTERNAL_ERROR,
                                    "InlineShapes.AddPicture 返回 null，图片未插入。",
                                    null,
                                    false);
                            }
                            WordSizeResult size = WordImageUtil.ComputeSize(
                                inlineShape.Width, inlineShape.Height,
                                widthPt, heightPt, scalePercent, widthLimitPt, lockAspect);
                            inlineShape.Width = (float)size.WidthPt;
                            inlineShape.Height = (float)size.HeightPt;
                            inlineShape.AlternativeText = altText ?? string.Empty;
                            actualWidth = inlineShape.Width;
                            actualHeight = inlineShape.Height;
                            index = WordComUtil.FindInlineShapeIndex(document, inlineShape);
                            if (!string.IsNullOrWhiteSpace(size.Warning))
                            {
                                warnings.Add(size.Warning);
                            }
                            if (caption)
                            {
                                anchorRange = inlineShape.Range;
                                captionInserted = InsertImageCaption(
                                    document, anchorRange, altText, index.Value, warnings);
                                if (!captionInserted)
                                {
                                    captionNote = "图片已插入，但图注插入失败。";
                                }
                            }
                            if (CountInlineShapes(document) <= beforeInline)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INTERNAL_ERROR,
                                    "内嵌图片插入后 InlineShapes 数量未增加。",
                                    null,
                                    false);
                            }
                        }

                        JObject data = new JObject();
                        data["collection"] = floating ? "floating" : "inline";
                        data["index"] = index ?? 0;
                        data["floating"] = floating;
                        data["widthPt"] = actualWidth;
                        data["heightPt"] = actualHeight;
                        data["requestedWidthPt"] = widthPt.HasValue ? new JValue(widthPt.Value) : JValue.CreateNull();
                        data["requestedHeightPt"] = heightPt.HasValue ? new JValue(heightPt.Value) : JValue.CreateNull();
                        data["scalePercent"] = scalePercent.HasValue ? new JValue(scalePercent.Value) : JValue.CreateNull();
                        data["widthLimitPt"] = widthLimitPt.HasValue ? new JValue(widthLimitPt.Value) : JValue.CreateNull();
                        data["lockAspect"] = lockAspect;
                        data["altText"] = altText ?? string.Empty;
                        data["captionRequested"] = caption;
                        data["captionInserted"] = captionInserted;
                        data["warnings"] = warnings;
                        data["filePath"] = Path.GetFullPath(filePath);
                        if (caption && !captionInserted)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.PARTIAL_APPLY,
                                captionNote + " 图片本身已插入；请检查插入点/锚点是否有效。",
                                data,
                                true);
                        }
                        WordWriteResult result = WordWriteResult.Ok(
                            "已插入" + (floating ? "浮动" : "内嵌") + "图片（索引 " + (index ?? 0) +
                            "，实际 " + actualWidth.ToString("0.##", CultureInfo.InvariantCulture) + "×" +
                            actualHeight.ToString("0.##", CultureInfo.InvariantCulture) + " 磅）" +
                            (captionInserted ? "，已插入 9pt 斜体图注。" : "。"),
                            data,
                            true);
                        result.AddTargetRef((floating ? "Word:Shape:" : "Word:InlineShape:") + (index ?? 0));
                        return result;
                    }
                    finally
                    {
                        ComHelper.Release(anchorRange, floatingShape, shapes, inlineShape, inlineShapes, insertionRange);
                    }
                });
        }

        private OperationResult HandleImagesList(Word.Application application, JObject arguments)
        {
            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    Word.InlineShapes inlineShapes = null;
                    Word.Shapes shapes = null;
                    JArray inlineItems = new JArray();
                    JArray floatingItems = new JArray();
                    try
                    {
                        inlineShapes = document.InlineShapes;
                        int inlineCount = inlineShapes == null ? 0 : inlineShapes.Count;
                        for (int index = 1; index <= inlineCount; index++)
                        {
                            Word.InlineShape shape = null;
                            try
                            {
                                shape = inlineShapes[index];
                                inlineItems.Add(WordImageUtil.CaptureInline(document, shape, index));
                            }
                            finally
                            {
                                ComHelper.Release(shape);
                            }
                        }
                        shapes = document.Shapes;
                        int floatingCount = shapes == null ? 0 : shapes.Count;
                        for (int index = 1; index <= floatingCount; index++)
                        {
                            Word.Shape shape = null;
                            try
                            {
                                shape = shapes[index];
                                floatingItems.Add(WordImageUtil.CaptureFloating(document, shape, index));
                            }
                            finally
                            {
                                ComHelper.Release(shape);
                            }
                        }
                        JObject data = new JObject();
                        data["inlineShapes"] = inlineItems;
                        data["floatingShapes"] = floatingItems;
                        data["inlineCount"] = inlineCount;
                        data["floatingCount"] = floatingCount;
                        data["note"] = "InlineShapes 与 Shapes 是两套集合，必须分开枚举。Shapes 可能包含非图片形状；" +
                                        "可通过 type/name 区分。";
                        return ReadSucceed(
                            "读取 " + inlineCount + " 个内嵌图片、" + floatingCount + " 个浮动对象。",
                            data, "Word:InlineShapes", "Word:Shapes");
                    }
                    finally
                    {
                        ComHelper.Release(shapes, inlineShapes);
                    }
                });
        }

        private OperationResult HandleImageResize(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string collection = (args.GetOptionalString("collection", "inline") ?? "inline")
                .Trim().ToLowerInvariant();
            bool useSelection = args.GetOptionalBool("useSelection", false);
            int? index = args.GetOptionalInt("index");
            double? widthPt = args.GetOptionalDouble("widthPt");
            double? heightPt = args.GetOptionalDouble("heightPt");
            double? scalePercent = args.GetOptionalDouble("scalePercent");
            double? widthLimitPt = args.GetOptionalDouble("widthLimitPt");
            bool lockAspect = args.GetOptionalBool("lockAspect", true);
            if (collection != "inline" && collection != "floating")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "collection 只支持 inline/floating，收到 " + collection + "。");
            }
            if (!widthPt.HasValue && !heightPt.HasValue && !scalePercent.HasValue && !widthLimitPt.HasValue)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "至少需要 widthPt/heightPt/scalePercent/widthLimitPt 之一。");
            }
            if (!useSelection && !index.HasValue)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "需要 index（1 基）或 useSelection=true 指定要缩放的图片。");
            }

            return RunWriteTool(
                application,
                WordToolNames.ImageResize,
                false,
                delegate(Word.Document document)
                {
                    ValidateImageTarget(application, document, collection, useSelection, index);
                    return WordWriteTargets.None("图片缩放",
                        "Word:Image:" + (useSelection ? "selection" : index.Value.ToString()));
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Selection selection = null;
                    Word.InlineShapes inlineShapes = null;
                    Word.InlineShape inlineShape = null;
                    Word.ShapeRange shapeRange = null;
                    Word.Shape floatingShape = null;
                    try
                    {
                        int actualIndex;
                        double beforeWidth;
                        double beforeHeight;
                        if (collection == "inline")
                        {
                            if (useSelection)
                            {
                                selection = application.Selection;
                                inlineShapes = selection == null ? null : selection.InlineShapes;
                                if (inlineShapes == null || inlineShapes.Count == 0)
                                {
                                    return WordWriteResult.Fail(
                                        OfficeErrorCode.INVALID_PARAMS,
                                        "当前选区没有内嵌图片。请先选中图片，或改用 index。",
                                        null,
                                        false);
                                }
                                inlineShape = inlineShapes[1];
                                actualIndex = WordComUtil.FindInlineShapeIndex(document, inlineShape);
                            }
                            else
                            {
                                inlineShapes = document.InlineShapes;
                                if (index.Value > inlineShapes.Count || index.Value < 1)
                                {
                                    return WordWriteResult.Fail(
                                        OfficeErrorCode.INVALID_PARAMS,
                                        "内嵌图片索引 " + index.Value + " 不存在（当前共 " +
                                        inlineShapes.Count + " 个，1 基）。",
                                        null,
                                        false);
                                }
                                inlineShape = inlineShapes[index.Value];
                                actualIndex = index.Value;
                            }
                            beforeWidth = inlineShape.Width;
                            beforeHeight = inlineShape.Height;
                            WordSizeResult size = WordImageUtil.ComputeSize(
                                beforeWidth, beforeHeight, widthPt, heightPt,
                                scalePercent, widthLimitPt, lockAspect);
                            inlineShape.Width = (float)size.WidthPt;
                            inlineShape.Height = (float)size.HeightPt;
                            double afterWidth = inlineShape.Width;
                            double afterHeight = inlineShape.Height;
                            bool verified = WordImageUtil.SizeMatches(
                                afterWidth, afterHeight, size.WidthPt, size.HeightPt, 0.5);
                            JObject data = BuildResizeData(
                                "inline", actualIndex, beforeWidth, beforeHeight,
                                afterWidth, afterHeight, size, 0.5, verified, lockAspect);
                            if (!verified)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INTERNAL_ERROR,
                                    "内嵌图片缩放后回读超差：期望 " + size.WidthPt.ToString("0.##") +
                                    "×" + size.HeightPt.ToString("0.##") + " 磅，实际 " +
                                    afterWidth.ToString("0.##") + "×" + afterHeight.ToString("0.##") + " 磅。",
                                    data,
                                    true);
                            }
                            bool changed = !WordFormatUtil.CloseEnough(beforeWidth, afterWidth, 0.1) ||
                                           !WordFormatUtil.CloseEnough(beforeHeight, afterHeight, 0.1);
                            WordWriteResult result = WordWriteResult.Ok(
                                "已缩放内嵌图片 " + actualIndex + " 至 " +
                                afterWidth.ToString("0.##") + "×" + afterHeight.ToString("0.##") + " 磅。",
                                data,
                                changed);
                            result.AddTargetRef("Word:InlineShape:" + actualIndex);
                            if (!string.IsNullOrWhiteSpace(size.Warning))
                            {
                                data["warning"] = size.Warning;
                            }
                            return result;
                        }

                        if (useSelection)
                        {
                            selection = application.Selection;
                            shapeRange = selection == null ? null : selection.ShapeRange;
                            if (shapeRange == null || shapeRange.Count == 0)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INVALID_PARAMS,
                                    "当前选区没有浮动图片/形状。请先选中浮动图片，或改用 index。",
                                    null,
                                    false);
                            }
                            floatingShape = shapeRange[1];
                            actualIndex = WordComUtil.FindShapeIndex(document, floatingShape);
                        }
                        else
                        {
                            Word.Shapes shapeCollection = null;
                            try
                            {
                                shapeCollection = document.Shapes;
                                if (index.Value > shapeCollection.Count || index.Value < 1)
                                {
                                    return WordWriteResult.Fail(
                                        OfficeErrorCode.INVALID_PARAMS,
                                        "浮动图片索引 " + index.Value + " 不存在（当前共 " +
                                        shapeCollection.Count + " 个 Shapes，1 基；可能含非图片形状）。",
                                        null,
                                        false);
                                }
                                floatingShape = shapeCollection[index.Value];
                                actualIndex = index.Value;
                            }
                            finally
                            {
                                ComHelper.Release(shapeCollection);
                            }
                        }
                        beforeWidth = floatingShape.Width;
                        beforeHeight = floatingShape.Height;
                        WordSizeResult floatingSize = WordImageUtil.ComputeSize(
                            beforeWidth, beforeHeight, widthPt, heightPt,
                            scalePercent, widthLimitPt, lockAspect);
                        floatingShape.Width = (float)floatingSize.WidthPt;
                        floatingShape.Height = (float)floatingSize.HeightPt;
                        double floatingAfterWidth = floatingShape.Width;
                        double floatingAfterHeight = floatingShape.Height;
                        bool floatingVerified = WordImageUtil.SizeMatches(
                            floatingAfterWidth, floatingAfterHeight,
                            floatingSize.WidthPt, floatingSize.HeightPt, 0.5);
                        JObject floatingData = BuildResizeData(
                            "floating", actualIndex, beforeWidth, beforeHeight,
                            floatingAfterWidth, floatingAfterHeight, floatingSize, 0.5,
                            floatingVerified, lockAspect);
                        if (!floatingVerified)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INTERNAL_ERROR,
                                "浮动图片缩放后回读超差：期望 " + floatingSize.WidthPt.ToString("0.##") +
                                "×" + floatingSize.HeightPt.ToString("0.##") + " 磅，实际 " +
                                floatingAfterWidth.ToString("0.##") + "×" +
                                floatingAfterHeight.ToString("0.##") + " 磅。",
                                floatingData,
                                true);
                        }
                        bool floatingChanged = !WordFormatUtil.CloseEnough(beforeWidth, floatingAfterWidth, 0.1) ||
                                               !WordFormatUtil.CloseEnough(beforeHeight, floatingAfterHeight, 0.1);
                        WordWriteResult floatingResult = WordWriteResult.Ok(
                            "已缩放浮动图片 " + actualIndex + " 至 " +
                            floatingAfterWidth.ToString("0.##") + "×" +
                            floatingAfterHeight.ToString("0.##") + " 磅。",
                            floatingData,
                            floatingChanged);
                        floatingResult.AddTargetRef("Word:Shape:" + actualIndex);
                        return floatingResult;
                    }
                    finally
                    {
                        ComHelper.Release(floatingShape, shapeRange, inlineShape, inlineShapes, selection);
                    }
                });
        }

        private OperationResult HandleImageLayout(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int index = args.RequireInt("index", 1, int.MaxValue);
            string collection = (args.GetOptionalString("collection", "inline") ?? "inline")
                .Trim().ToLowerInvariant();
            if (collection != "inline" && collection != "floating")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "collection 只支持 inline/floating，收到 " + collection + "。");
            }
            string wrapText = args.Has("wrap") ? args.GetOptionalString("wrap", null) : null;
            Word.WdWrapType? wrapType = !string.IsNullOrWhiteSpace(wrapText)
                ? WordImageUtil.ParseWrapType(wrapText)
                : (Word.WdWrapType?)null;
            double? xPt = args.GetOptionalDouble("xPt");
            double? yPt = args.GetOptionalDouble("yPt");
            string relativeTo = (args.GetOptionalString("relativeTo", "margin") ?? "margin")
                .Trim().ToLowerInvariant();
            if (relativeTo != "margin" && relativeTo != "page" && relativeTo != "paragraph" &&
                relativeTo != "页边距" && relativeTo != "页面" && relativeTo != "段落")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "relativeTo 只支持 margin/page/paragraph，收到 " + relativeTo + "。");
            }
            int? anchorParagraph = args.GetOptionalInt("anchorParagraphIndex");
            string altText = args.Has("altText") ? args.GetOptionalString("altText", string.Empty) : null;
            if (!wrapType.HasValue && !xPt.HasValue && !yPt.HasValue &&
                !anchorParagraph.HasValue && altText == null)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "word_image_layout 至少需要 wrap/xPt/yPt/anchorParagraphIndex/altText 之一。");
            }

            return RunWriteTool(
                application,
                WordToolNames.ImageLayout,
                false,
                delegate(Word.Document document)
                {
                    ValidateImageTarget(application, document, collection, false, index);
                    if (anchorParagraph.HasValue)
                    {
                        WordComUtil.ValidateParagraphIndex(document, anchorParagraph.Value);
                    }
                    return WordWriteTargets.None("图片布局", "Word:Image:" + index);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.InlineShapes inlineShapes = null;
                    Word.InlineShape inlineShape = null;
                    Word.Shapes shapes = null;
                    Word.Shape floatingShape = null;
                    string tempFile = null;
                    try
                    {
                        if (collection == "inline")
                        {
                            inlineShapes = document.InlineShapes;
                            if (index > inlineShapes.Count)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INVALID_PARAMS,
                                    "内嵌图片索引 " + index + " 不存在（当前共 " + inlineShapes.Count +
                                    " 个，1 基）。",
                                    null,
                                    false);
                            }
                            inlineShape = inlineShapes[index];
                            int inlineAnchor = FindInlineAnchorParagraph(document, inlineShape);
                            int targetAnchor = anchorParagraph.HasValue ? anchorParagraph.Value : inlineAnchor;
                            bool needsConversion = (wrapType.HasValue && wrapType.Value != Word.WdWrapType.wdWrapInline) ||
                                                   xPt.HasValue || yPt.HasValue || anchorParagraph.HasValue;
                            if (!needsConversion)
                            {
                                if (altText != null)
                                {
                                    string inlineAlt = inlineShape.AlternativeText ?? string.Empty;
                                    inlineShape.AlternativeText = altText;
                                    bool altChanged = !string.Equals(inlineAlt, altText, StringComparison.Ordinal);
                                    JObject altData = new JObject();
                                    altData["collection"] = "inline";
                                    altData["index"] = index;
                                    altData["altText"] = altText;
                                    altData["conversion"] = "none";
                                    WordWriteResult altResult = WordWriteResult.Ok(
                                        "已更新内嵌图片 " + index + " 的 Alt 文本；未改变布局。",
                                        altData,
                                        altChanged);
                                    altResult.AddTargetRef("Word:InlineShape:" + index);
                                    return altResult;
                                }
                                return WordWriteResult.Ok(
                                    "内嵌图片 " + index + " 已是 wrap=inline，且未提供其它布局参数，无需修改。",
                                    new JObject { { "collection", "inline" }, { "index", index }, { "changed", false } },
                                    false);
                            }

                            double inlineWidth = inlineShape.Width;
                            double inlineHeight = inlineShape.Height;
                            string inlineAltText = inlineShape.AlternativeText ?? string.Empty;
                            string exportError;
                            tempFile = WordImageLayoutConverter.ExportToTempFile(
                                application, document, inlineShape, null, out exportError);
                            if (tempFile == null)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INVALID_PARAMS,
                                    "inline→floating 转换需要先导出原图，但导出失败：" + exportError +
                                    "。为避免丢失图片，未删除原图，也未使用 PIA 直接转换。",
                                    null,
                                    false);
                            }

                            Word.Range anchorRange = GetAnchorRange(document, targetAnchor);
                            Word.WdWrapType targetWrap = wrapType.HasValue &&
                                                        wrapType.Value != Word.WdWrapType.wdWrapInline
                                ? wrapType.Value
                                : Word.WdWrapType.wdWrapSquare;
                            Word.WdRelativeHorizontalPosition relativeHorizontal =
                                WordImageUtil.ParseRelativeHorizontal(relativeTo);
                            Word.WdRelativeVerticalPosition relativeVertical =
                                WordImageUtil.ParseRelativeVertical(relativeTo);
                            int newIndex;
                            Word.Shape newShape = WordImageLayoutConverter.CreateFloatingFromFile(
                                document,
                                tempFile,
                                anchorRange,
                                xPt.HasValue ? xPt.Value : 0d,
                                yPt.HasValue ? yPt.Value : 0d,
                                targetWrap,
                                relativeHorizontal,
                                relativeVertical,
                                altText ?? inlineAltText,
                                inlineWidth,
                                inlineHeight,
                                out newIndex);
                            // 新建成功后再删除旧图；创建失败时原图仍保留。
                            inlineShape.Delete();
                            ComHelper.Release(newShape);
                            ComHelper.Release(anchorRange);
                            JObject data = new JObject();
                            data["collectionBefore"] = "inline";
                            data["collectionAfter"] = "floating";
                            data["oldIndex"] = index;
                            data["newIndex"] = newIndex;
                            data["widthPt"] = inlineWidth;
                            data["heightPt"] = inlineHeight;
                            data["anchorParagraphIndex"] = targetAnchor;
                            data["wrap"] = targetWrap.ToString();
                            data["xPt"] = xPt.HasValue ? new JValue(xPt.Value) : JValue.CreateNull();
                            data["yPt"] = yPt.HasValue ? new JValue(yPt.Value) : JValue.CreateNull();
                            data["conversionNote"] =
                                "inline→floating 使用 EMF 临时文件删除重建（未使用 PIA ConvertToShape）；" +
                                "EMF 可能丢失部分原始编码/透明信息。";
                            WordWriteResult result = WordWriteResult.Ok(
                                "已把内嵌图片转换为浮动图片（删除重建），新 Shape 索引 " + newIndex + "。",
                                data,
                                true);
                            result.AddTargetRef("Word:Shape:" + newIndex);
                            return result;
                        }

                        shapes = document.Shapes;
                        if (index > shapes.Count)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "浮动对象索引 " + index + " 不存在（当前共 " + shapes.Count +
                                " 个 Shapes，1 基）。",
                                null,
                                false);
                        }
                        floatingShape = shapes[index];
                        Word.WrapFormat currentWrap = null;
                        Word.WdWrapType oldWrapType;
                        double oldLeft = floatingShape.Left;
                        double oldTop = floatingShape.Top;
                        double oldWidth = floatingShape.Width;
                        double oldHeight = floatingShape.Height;
                        string oldAltText = floatingShape.AlternativeText ?? string.Empty;
                        int oldAnchor = FindFloatingAnchorParagraph(document, floatingShape);
                        try
                        {
                            currentWrap = floatingShape.WrapFormat;
                            oldWrapType = currentWrap.Type;
                        }
                        finally
                        {
                            ComHelper.Release(currentWrap);
                        }

                        bool convertToInline = wrapType.HasValue &&
                                               wrapType.Value == Word.WdWrapType.wdWrapInline;
                        bool reAnchor = anchorParagraph.HasValue;
                        if (!convertToInline && !reAnchor)
                        {
                            bool layoutChanged = false;
                            if (wrapType.HasValue)
                            {
                                Word.WrapFormat wrap = null;
                                try
                                {
                                    wrap = floatingShape.WrapFormat;
                                    wrap.Type = wrapType.Value;
                                    layoutChanged = oldWrapType != wrapType.Value;
                                }
                                finally
                                {
                                    ComHelper.Release(wrap);
                                }
                            }
                            Word.WdRelativeHorizontalPosition horizontal =
                                WordImageUtil.ParseRelativeHorizontal(relativeTo);
                            Word.WdRelativeVerticalPosition vertical =
                                WordImageUtil.ParseRelativeVertical(relativeTo);
                            floatingShape.RelativeHorizontalPosition = horizontal;
                            floatingShape.RelativeVerticalPosition = vertical;
                            if (xPt.HasValue)
                            {
                                floatingShape.Left = (float)xPt.Value;
                                layoutChanged = layoutChanged ||
                                                !WordFormatUtil.CloseEnough(oldLeft, xPt.Value, 0.5);
                            }
                            if (yPt.HasValue)
                            {
                                floatingShape.Top = (float)yPt.Value;
                                layoutChanged = layoutChanged ||
                                                !WordFormatUtil.CloseEnough(oldTop, yPt.Value, 0.5);
                            }
                            if (altText != null)
                            {
                                floatingShape.AlternativeText = altText;
                                layoutChanged = layoutChanged ||
                                                !string.Equals(oldAltText, altText, StringComparison.Ordinal);
                            }
                            JObject data = WordImageUtil.CaptureFloating(document, floatingShape, index);
                            data["conversion"] = "none";
                            data["relativeTo"] = relativeTo;
                            WordWriteResult result = WordWriteResult.Ok(
                                "已更新浮动对象 " + index + " 的布局（wrap/位置/相对基准/Alt 文本）。",
                                data,
                                layoutChanged);
                            result.AddTargetRef("Word:Shape:" + index);
                            return result;
                        }

                        string tempError;
                        tempFile = WordImageLayoutConverter.ExportToTempFile(
                            application, document, null, floatingShape, out tempError);
                        if (tempFile == null)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "浮动图片转换需要先导出原图，但导出失败：" + tempError +
                                "。为避免丢失图片，未删除原对象，也未使用 PIA 直接转换。",
                                null,
                                false);
                        }
                        int targetParagraph = anchorParagraph.HasValue ? anchorParagraph.Value : oldAnchor;
                        Word.Range newAnchorRange = GetAnchorRange(document, targetParagraph);
                        if (convertToInline)
                        {
                            int newInlineIndex;
                            Word.InlineShape newInline = WordImageLayoutConverter.CreateInlineFromFile(
                                document,
                                tempFile,
                                newAnchorRange,
                                altText ?? oldAltText,
                                oldWidth,
                                oldHeight,
                                out newInlineIndex);
                            floatingShape.Delete();
                            ComHelper.Release(newInline);
                            ComHelper.Release(newAnchorRange);
                            JObject data = new JObject();
                            data["collectionBefore"] = "floating";
                            data["collectionAfter"] = "inline";
                            data["oldIndex"] = index;
                            data["newIndex"] = newInlineIndex;
                            data["widthPt"] = oldWidth;
                            data["heightPt"] = oldHeight;
                            data["anchorParagraphIndex"] = targetParagraph;
                            data["conversionNote"] =
                                "floating→inline 使用 EMF 临时文件删除重建（未使用 PIA ConvertToInlineShape）；" +
                                "EMF 可能丢失部分原始编码/透明信息。";
                            WordWriteResult result = WordWriteResult.Ok(
                                "已把浮动对象转换为内嵌图片（删除重建），新 InlineShape 索引 " +
                                newInlineIndex + "。",
                                data,
                                true);
                            result.AddTargetRef("Word:InlineShape:" + newInlineIndex);
                            return result;
                        }

                        Word.WdWrapType rebuildWrap = wrapType.HasValue &&
                                                      wrapType.Value != Word.WdWrapType.wdWrapInline
                            ? wrapType.Value
                            : oldWrapType;
                        int rebuiltIndex;
                        Word.Shape rebuilt = WordImageLayoutConverter.CreateFloatingFromFile(
                            document,
                            tempFile,
                            newAnchorRange,
                            xPt.HasValue ? xPt.Value : oldLeft,
                            yPt.HasValue ? yPt.Value : oldTop,
                            rebuildWrap,
                            WordImageUtil.ParseRelativeHorizontal(relativeTo),
                            WordImageUtil.ParseRelativeVertical(relativeTo),
                            altText ?? oldAltText,
                            oldWidth,
                            oldHeight,
                            out rebuiltIndex);
                        floatingShape.Delete();
                        ComHelper.Release(rebuilt);
                        ComHelper.Release(newAnchorRange);
                        JObject rebuildData = new JObject();
                        rebuildData["collection"] = "floating";
                        rebuildData["oldIndex"] = index;
                        rebuildData["newIndex"] = rebuiltIndex;
                        rebuildData["anchorParagraphIndex"] = targetParagraph;
                        rebuildData["wrap"] = rebuildWrap.ToString();
                        rebuildData["relativeTo"] = relativeTo;
                        rebuildData["conversionNote"] =
                            "修改锚定段落/对象位置使用 EMF 临时文件删除重建（Word 无直接修改 Anchor 的 API）；" +
                            "EMF 可能丢失部分原始编码/透明信息。";
                        WordWriteResult rebuildResult = WordWriteResult.Ok(
                            "已重建浮动对象并改锚定到第 " + targetParagraph + " 段，新 Shape 索引 " +
                            rebuiltIndex + "。",
                            rebuildData,
                            true);
                        rebuildResult.AddTargetRef("Word:Shape:" + rebuiltIndex);
                        return rebuildResult;
                    }
                    finally
                    {
                        ComHelper.Release(floatingShape, shapes, inlineShape, inlineShapes);
                        if (!string.IsNullOrWhiteSpace(tempFile))
                        {
                            try
                            {
                                File.Delete(tempFile);
                            }
                            catch (Exception ex)
                            {
                                Console.Error.WriteLine("[word] 删除图片临时文件失败：" + ex.Message);
                            }
                        }
                    }
                });
        }

        private static void ValidateImageTarget(
            Word.Application application,
            Word.Document document,
            string collection,
            bool useSelection,
            int? index)
        {
            if (useSelection)
            {
                Word.Selection selection = null;
                Word.InlineShapes inlineShapes = null;
                Word.ShapeRange shapeRange = null;
                try
                {
                    selection = application.Selection;
                    if (collection == "inline")
                    {
                        inlineShapes = selection == null ? null : selection.InlineShapes;
                        if (inlineShapes == null || inlineShapes.Count == 0)
                        {
                            throw WordArguments.Invalid("当前选区没有内嵌图片。");
                        }
                    }
                    else
                    {
                        shapeRange = selection == null ? null : selection.ShapeRange;
                        if (shapeRange == null || shapeRange.Count == 0)
                        {
                            throw WordArguments.Invalid("当前选区没有浮动图片/形状。");
                        }
                    }
                }
                finally
                {
                    ComHelper.Release(shapeRange, inlineShapes, selection);
                }
                return;
            }
            if (!index.HasValue)
            {
                throw WordArguments.Invalid("缺少 index 或 useSelection=true。");
            }
            if (collection == "inline")
            {
                Word.InlineShapes shapes = null;
                try
                {
                    shapes = document.InlineShapes;
                    if (index.Value < 1 || index.Value > shapes.Count)
                    {
                        throw WordArguments.Invalid(
                            "内嵌图片索引 " + index.Value + " 不存在（当前共 " + shapes.Count + " 个，1 基）。");
                    }
                }
                finally
                {
                    ComHelper.Release(shapes);
                }
            }
            else
            {
                Word.Shapes shapes = null;
                try
                {
                    shapes = document.Shapes;
                    if (index.Value < 1 || index.Value > shapes.Count)
                    {
                        throw WordArguments.Invalid(
                            "浮动对象索引 " + index.Value + " 不存在（当前共 " + shapes.Count +
                            " 个 Shapes，1 基；可能含非图片形状）。");
                    }
                }
                finally
                {
                    ComHelper.Release(shapes);
                }
            }
        }

        private static JObject BuildResizeData(
            string collection,
            int index,
            double beforeWidth,
            double beforeHeight,
            double afterWidth,
            double afterHeight,
            WordSizeResult size,
            double tolerance,
            bool verified,
            bool lockAspect)
        {
            JObject data = new JObject();
            data["collection"] = collection;
            data["index"] = index;
            data["beforeWidthPt"] = beforeWidth;
            data["beforeHeightPt"] = beforeHeight;
            data["afterWidthPt"] = afterWidth;
            data["afterHeightPt"] = afterHeight;
            data["expectedWidthPt"] = size.WidthPt;
            data["expectedHeightPt"] = size.HeightPt;
            data["tolerancePt"] = tolerance;
            data["verified"] = verified;
            data["lockAspect"] = lockAspect;
            data["warning"] = size.Warning ?? string.Empty;
            return data;
        }

        private static int CountInlineShapes(Word.Document document)
        {
            Word.InlineShapes shapes = null;
            try
            {
                shapes = document.InlineShapes;
                return shapes == null ? 0 : shapes.Count;
            }
            finally
            {
                ComHelper.Release(shapes);
            }
        }

        private static int CountFloatingShapes(Word.Document document)
        {
            Word.Shapes shapes = null;
            try
            {
                shapes = document.Shapes;
                return shapes == null ? 0 : shapes.Count;
            }
            finally
            {
                ComHelper.Release(shapes);
            }
        }

        private static int FindInlineAnchorParagraph(Word.Document document, Word.InlineShape shape)
        {
            Word.Range range = null;
            try
            {
                range = shape.Range;
                return WordComUtil.FindParagraphIndexAt(document, range == null ? 0 : range.Start);
            }
            finally
            {
                ComHelper.Release(range);
            }
        }

        private static int FindFloatingAnchorParagraph(Word.Document document, Word.Shape shape)
        {
            Word.Range anchor = null;
            try
            {
                anchor = shape.Anchor;
                return WordComUtil.FindParagraphIndexAt(document, anchor == null ? 0 : anchor.Start);
            }
            finally
            {
                ComHelper.Release(anchor);
            }
        }

        private static Word.Range GetAnchorRange(Word.Document document, int paragraphIndex)
        {
            Word.Paragraph paragraph = null;
            Word.Range paragraphRange = null;
            try
            {
                paragraph = WordComUtil.GetParagraph(document, paragraphIndex);
                paragraphRange = paragraph.Range;
                return WordComUtil.CreateRange(document, paragraphRange.Start, paragraphRange.Start);
            }
            finally
            {
                ComHelper.Release(paragraphRange, paragraph);
            }
        }

        private static WordWriteTarget SelectImageInsertionTarget(
            Word.Application application,
            Word.Document document,
            string position)
        {
            if (position == "start")
            {
                return new WordWriteTarget(0, 0, "文档开头", "Word:Image", false, true);
            }
            if (position == "end")
            {
                int end = Math.Max(0, WordComUtil.GetContentEnd(document) - 1);
                return new WordWriteTarget(end, end, "文档末尾", "Word:Image", false, true);
            }
            Word.Selection selection = null;
            Word.Range range = null;
            try
            {
                selection = application.Selection;
                range = selection == null ? null : selection.Range;
                if (range == null)
                {
                    throw WordArguments.Invalid("position=cursor 时没有有效选区，无法插入图片。");
                }
                return new WordWriteTarget(range.Start, range.Start, "当前光标位置", "Word:Image", false, true);
            }
            finally
            {
                ComHelper.Release(range, selection);
            }
        }

        private static bool InsertImageCaption(
            Word.Document document,
            Word.Range anchorRange,
            string altText,
            int index,
            JArray warnings)
        {
            Word.Range captionRange = null;
            Word.Font font = null;
            Word.ParagraphFormat format = null;
            try
            {
                if (anchorRange == null)
                {
                    warnings.Add("图注插入失败：图片锚点为空。");
                    return false;
                }
                int position = anchorRange.End;
                captionRange = WordComUtil.CreateRange(document, position, position);
                string caption = "图 " + index + (string.IsNullOrWhiteSpace(altText) ? "" : " " + altText);
                captionRange.InsertAfter(WordFormatUtil.NormalizeForWord("\r" + caption));
                font = captionRange.Font;
                font.Italic = -1;
                font.Size = 9f;
                format = captionRange.ParagraphFormat;
                format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                string actual = captionRange.Text ?? string.Empty;
                if (actual.IndexOf("图", StringComparison.Ordinal) < 0)
                {
                    warnings.Add("图注写入后回读不含「图」。");
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                warnings.Add("图注插入失败：" + ex.Message);
                return false;
            }
            finally
            {
                ComHelper.Release(format, font, captionRange);
            }
        }
    }
}
