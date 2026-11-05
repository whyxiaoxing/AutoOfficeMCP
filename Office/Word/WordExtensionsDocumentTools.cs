using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    public sealed partial class WordExtensionsExecutor
    {
        private void RegisterDocumentExtensionHandlers(
            Dictionary<string, Func<Word.Application, JObject, OperationResult>> handlers)
        {
            handlers[WordExtensionsToolNames.StyleAudit] = HandleStyleAudit;
            handlers[WordExtensionsToolNames.ExtractImages] = HandleExtractImages;
            handlers[WordExtensionsToolNames.DocMerge] = HandleDocMerge;
            RegisterStructureExtensionHandlers(handlers);
        }

        // ------------------------------------------------------------------
        // word_style_audit（只读）
        // ------------------------------------------------------------------

        private OperationResult HandleStyleAudit(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int maxFindings = args.GetOptionalInt("maxFindings", 100);
            int maxDirectSamples = args.GetOptionalInt("maxDirectSamples", 50);
            if (maxFindings < 1 || maxFindings > 1000)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "maxFindings 允许 1-1000，收到 " + maxFindings + "。");
            }
            if (maxDirectSamples < 1 || maxDirectSamples > 500)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "maxDirectSamples 允许 1-500，收到 " + maxDirectSamples + "。");
            }

            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    List<int> indexes = WordExtensionsUtil.ResolveParagraphIndexes(
                        application, document, args, out string scopeLabel);
                    Dictionary<string, int> styleCounts =
                        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    Dictionary<string, StyleFormatSnapshot> styleCache =
                        new Dictionary<string, StyleFormatSnapshot>(StringComparer.OrdinalIgnoreCase);
                    JArray suggestions = new JArray();
                    JArray directSamples = new JArray();
                    JArray warnings = new JArray();
                    int paragraphsScanned = 0;
                    int emptyParagraphs = 0;
                    int styledParagraphs = 0;
                    int directFormattedParagraphs = 0;
                    int headingLikeBodyParagraphs = 0;
                    int suggestionCount = 0;

                    foreach (int paragraphIndex in indexes)
                    {
                        Word.Paragraph paragraph = null;
                        Word.Range range = null;
                        Word.Font font = null;
                        Word.ParagraphFormat paragraphFormat = null;
                        try
                        {
                            paragraph = WordComUtil.GetParagraph(document, paragraphIndex);
                            string text = WordExtensionsUtil.ReadParagraphText(paragraph);
                            if (text.Length == 0)
                            {
                                emptyParagraphs++;
                                continue;
                            }
                            paragraphsScanned++;
                            range = paragraph.Range;
                            string styleName = WordComUtil.ReadStyleName(range);
                            if (string.IsNullOrWhiteSpace(styleName))
                            {
                                styleName = "（未设置样式）";
                            }
                            int currentCount;
                            styleCounts.TryGetValue(styleName, out currentCount);
                            styleCounts[styleName] = currentCount + 1;
                            if (!string.IsNullOrWhiteSpace(styleName) &&
                                styleName != "（未设置样式）")
                            {
                                styledParagraphs++;
                            }

                            StyleFormatSnapshot styleSnapshot = GetStyleSnapshot(
                                document, styleName, styleCache);
                            font = range.Font;
                            paragraphFormat = range.ParagraphFormat;
                            JArray directFlags = DetectDirectFormatting(
                                font, paragraphFormat, styleSnapshot, styleName);

                            if (directFlags.Count > 0)
                            {
                                directFormattedParagraphs++;
                                if (directSamples.Count < maxDirectSamples)
                                {
                                    directSamples.Add(new JObject
                                    {
                                        { "paragraphIndex", paragraphIndex },
                                        { "style", styleName },
                                        { "textPreview", WordExtensionsUtil.Preview(text, 100) },
                                        { "flags", directFlags },
                                        { "fontSize", WordFormatUtil.IsUndefined(font.Size)
                                            ? JValue.CreateNull() : new JValue((double)font.Size) },
                                        { "fontName", font.Name ?? string.Empty },
                                        { "bold", WordFormatUtil.TriStateToken(font.Bold) }
                                    });
                                }
                            }

                            int headingLevel = WordExtensionsUtil.TryGetHeadingLevel(styleName);
                            int suggestedLevel;
                            int prefixLength;
                            string prefixText;
                            bool manualHeading = WordExtensionsUtil.TryMatchManualHeading(
                                text, out suggestedLevel, out prefixLength, out prefixText);
                            bool bodyLike = WordExtensionsUtil.IsBodyLikeStyle(styleName) &&
                                            !WordExtensionsUtil.IsCaptionOrTocStyle(styleName);
                            if (headingLevel == 0 && bodyLike)
                            {
                                JArray reasons = new JArray();
                                int levelToSuggest = 0;
                                if (manualHeading)
                                {
                                    levelToSuggest = suggestedLevel;
                                    reasons.Add("存在手工编号前缀「" +
                                                WordExtensionsUtil.Preview(prefixText.Trim(), 30) + "」");
                                }
                                if (!WordFormatUtil.IsUndefined(font.Bold) && font.Bold == -1)
                                {
                                    reasons.Add("整段加粗");
                                }
                                if (!WordFormatUtil.IsUndefined(font.Size) && font.Size >= 16f)
                                {
                                    reasons.Add("字号较大（" +
                                                ((double)font.Size).ToString("0.#",
                                                    CultureInfo.InvariantCulture) + " 磅）");
                                }
                                if (!manualHeading && reasons.Count >= 2 && text.Length <= 100)
                                {
                                    levelToSuggest = 1;
                                }
                                if (levelToSuggest > 0)
                                {
                                    headingLikeBodyParagraphs++;
                                    string suggestedStyle = WordExtensionsUtil.ResolveHeadingStyleName(
                                        document, levelToSuggest);
                                    if (string.IsNullOrWhiteSpace(suggestedStyle))
                                    {
                                        WordExtensionsUtil.AddWarning(
                                            warnings,
                                            "第 " + paragraphIndex + " 段疑似标题，但文档中找不到「标题 " +
                                            levelToSuggest + "」样式，未生成建议。");
                                    }
                                    else if (suggestionCount < maxFindings)
                                    {
                                        suggestionCount++;
                                        suggestions.Add(new JObject
                                        {
                                            { "paragraphIndex", paragraphIndex },
                                            { "currentStyle", styleName },
                                            { "suggestedStyle", suggestedStyle },
                                            { "suggestedLevel", levelToSuggest },
                                            { "textPreview", WordExtensionsUtil.Preview(text, 100) },
                                            { "reasons", reasons }
                                        });
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            WordExtensionsUtil.AddWarning(
                                warnings,
                                "第 " + paragraphIndex + " 段审计失败：" +
                                OfficeErrorMapper.DescribeException(ex, OfficeApp.Word));
                        }
                        finally
                        {
                            ComHelper.Release(paragraphFormat, font, range, paragraph);
                        }
                    }

                    JArray styleUsage = BuildStyleUsage(styleCounts);
                    JObject data = new JObject();
                    data["scope"] = scopeLabel;
                    data["paragraphsScanned"] = paragraphsScanned;
                    data["emptyParagraphs"] = emptyParagraphs;
                    data["styledParagraphs"] = styledParagraphs;
                    data["directFormattedParagraphs"] = directFormattedParagraphs;
                    data["headingLikeBodyParagraphs"] = headingLikeBodyParagraphs;
                    data["styleUsage"] = styleUsage;
                    data["suggestions"] = suggestions;
                    data["directFormattingSamples"] = directSamples;
                    data["warnings"] = warnings;
                    data["maxFindings"] = maxFindings;
                    data["maxDirectSamples"] = maxDirectSamples;
                    data["suggestionTruncated"] = suggestionCount >= maxFindings;
                    data["note"] = "直接格式化通过对比段落实际字体/段落属性与样式定义得出（Word 对象模型不提供直接格式标记），" +
                                    "混合值（9999999）不参与比较；建议样式仅针对正文类样式中带手工编号、整段加粗且字号较大的段，" +
                                    "属于启发式结果，请先用 word_get_paragraph_info 复核再批量套用。";

                    return ReadSucceed(
                        "样式审计完成：扫描 " + paragraphsScanned + " 段，" +
                        directFormattedParagraphs + " 段存在直接格式化，" +
                        headingLikeBodyParagraphs + " 段疑似应套用标题样式。",
                        data, "Word:Document");
                });
        }

        private static JArray BuildStyleUsage(Dictionary<string, int> styleCounts)
        {
            List<KeyValuePair<string, int>> list =
                new List<KeyValuePair<string, int>>(styleCounts);
            list.Sort(delegate(KeyValuePair<string, int> left, KeyValuePair<string, int> right)
            {
                int countCompare = right.Value.CompareTo(left.Value);
                return countCompare != 0
                    ? countCompare
                    : string.Compare(left.Key, right.Key, StringComparison.Ordinal);
            });
            JArray result = new JArray();
            foreach (KeyValuePair<string, int> pair in list)
            {
                result.Add(new JObject
                {
                    { "style", pair.Key },
                    { "count", pair.Value }
                });
            }
            return result;
        }

        private static StyleFormatSnapshot GetStyleSnapshot(
            Word.Document document,
            string styleName,
            Dictionary<string, StyleFormatSnapshot> cache)
        {
            if (cache != null)
            {
                StyleFormatSnapshot cached;
                if (cache.TryGetValue(styleName, out cached))
                {
                    return cached;
                }
            }

            StyleFormatSnapshot snapshot = null;
            object rawStyle = null;
            Word.Style style = null;
            Word.Font font = null;
            Word.ParagraphFormat paragraphFormat = null;
            try
            {
                rawStyle = document.Styles[styleName];
                style = rawStyle as Word.Style;
                if (style != null)
                {
                    snapshot = new StyleFormatSnapshot();
                    font = style.Font;
                    snapshot.FontName = font.Name ?? string.Empty;
                    snapshot.FontNameFarEast = font.NameFarEast ?? string.Empty;
                    snapshot.Size = WordFormatUtil.IsUndefined(font.Size)
                        ? (float?)null : (float)font.Size;
                    snapshot.Bold = WordFormatUtil.IsUndefined(font.Bold)
                        ? (int?)null : font.Bold;
                    snapshot.Italic = WordFormatUtil.IsUndefined(font.Italic)
                        ? (int?)null : font.Italic;
                    snapshot.Color = WordFormatUtil.IsUndefined((int)font.Color)
                        ? (int?)null : (int)font.Color;
                    snapshot.Underline = WordFormatUtil.IsUndefined((int)font.Underline)
                        ? (int?)null : (int)font.Underline;
                    paragraphFormat = style.ParagraphFormat;
                    snapshot.Alignment = WordFormatUtil.IsUndefined(
                        (int)paragraphFormat.Alignment)
                        ? (int?)null : (int)paragraphFormat.Alignment;
                    snapshot.SpaceBefore = WordFormatUtil.IsUndefined(paragraphFormat.SpaceBefore)
                        ? (float?)null : (float)paragraphFormat.SpaceBefore;
                    snapshot.SpaceAfter = WordFormatUtil.IsUndefined(paragraphFormat.SpaceAfter)
                        ? (float?)null : (float)paragraphFormat.SpaceAfter;
                    snapshot.FirstLineIndent = WordFormatUtil.IsUndefined(
                        paragraphFormat.FirstLineIndent)
                        ? (float?)null : (float)paragraphFormat.FirstLineIndent;
                }
            }
            catch
            {
                snapshot = null;
            }
            finally
            {
                ComHelper.Release(paragraphFormat, font, style, rawStyle);
            }

            if (cache != null)
            {
                cache[styleName] = snapshot;
            }
            return snapshot;
        }

        private static JArray DetectDirectFormatting(
            Word.Font font,
            Word.ParagraphFormat paragraphFormat,
            StyleFormatSnapshot styleSnapshot,
            string styleName)
        {
            JArray flags = new JArray();
            if (font == null || styleSnapshot == null)
            {
                return flags;
            }

            if (styleSnapshot.Size.HasValue && !WordFormatUtil.IsUndefined(font.Size) &&
                Math.Abs((double)font.Size - styleSnapshot.Size.Value) > 0.1)
            {
                flags.Add("fontSize");
            }
            string actualName = font.Name ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(actualName) &&
                !string.IsNullOrWhiteSpace(styleSnapshot.FontName) &&
                !string.Equals(actualName, styleSnapshot.FontName,
                    StringComparison.OrdinalIgnoreCase))
            {
                flags.Add("fontName");
            }
            string actualFarEast = font.NameFarEast ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(actualFarEast) &&
                !string.IsNullOrWhiteSpace(styleSnapshot.FontNameFarEast) &&
                !string.Equals(actualFarEast, styleSnapshot.FontNameFarEast,
                    StringComparison.OrdinalIgnoreCase))
            {
                flags.Add("fontNameFarEast");
            }
            if (styleSnapshot.Bold.HasValue && !WordFormatUtil.IsUndefined(font.Bold) &&
                font.Bold != styleSnapshot.Bold.Value)
            {
                flags.Add("bold");
            }
            if (styleSnapshot.Italic.HasValue && !WordFormatUtil.IsUndefined(font.Italic) &&
                font.Italic != styleSnapshot.Italic.Value)
            {
                flags.Add("italic");
            }
            if (styleSnapshot.Color.HasValue && !WordFormatUtil.IsUndefined((int)font.Color) &&
                (int)font.Color != styleSnapshot.Color.Value)
            {
                flags.Add("fontColor");
            }
            if (styleSnapshot.Underline.HasValue &&
                !WordFormatUtil.IsUndefined((int)font.Underline) &&
                (int)font.Underline != styleSnapshot.Underline.Value)
            {
                flags.Add("underline");
            }

            if (paragraphFormat != null)
            {
                if (styleSnapshot.Alignment.HasValue &&
                    !WordFormatUtil.IsUndefined((int)paragraphFormat.Alignment) &&
                    (int)paragraphFormat.Alignment != styleSnapshot.Alignment.Value)
                {
                    flags.Add("alignment");
                }
                if (styleSnapshot.SpaceBefore.HasValue &&
                    !WordFormatUtil.IsUndefined(paragraphFormat.SpaceBefore) &&
                    Math.Abs((double)paragraphFormat.SpaceBefore -
                             styleSnapshot.SpaceBefore.Value) > 0.5)
                {
                    flags.Add("spaceBefore");
                }
                if (styleSnapshot.SpaceAfter.HasValue &&
                    !WordFormatUtil.IsUndefined(paragraphFormat.SpaceAfter) &&
                    Math.Abs((double)paragraphFormat.SpaceAfter -
                             styleSnapshot.SpaceAfter.Value) > 0.5)
                {
                    flags.Add("spaceAfter");
                }
                if (styleSnapshot.FirstLineIndent.HasValue &&
                    !WordFormatUtil.IsUndefined(paragraphFormat.FirstLineIndent) &&
                    Math.Abs((double)paragraphFormat.FirstLineIndent -
                             styleSnapshot.FirstLineIndent.Value) > 0.5)
                {
                    flags.Add("firstLineIndent");
                }
            }

            if (flags.Count > 0 && WordExtensionsUtil.IsCaptionOrTocStyle(styleName))
            {
                // 题注/目录段的直接格式通常由域结果维护，不作为排版治理建议。
                return flags;
            }
            return flags;
        }

        private sealed class StyleFormatSnapshot
        {
            public string FontName { get; set; }

            public string FontNameFarEast { get; set; }

            public float? Size { get; set; }

            public int? Bold { get; set; }

            public int? Italic { get; set; }

            public int? Color { get; set; }

            public int? Underline { get; set; }

            public int? Alignment { get; set; }

            public float? SpaceBefore { get; set; }

            public float? SpaceAfter { get; set; }

            public float? FirstLineIndent { get; set; }
        }

        // ------------------------------------------------------------------
        // word_extract_images（只读文档 + 导出文件）
        // ------------------------------------------------------------------

        private OperationResult HandleExtractImages(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string requestedDirectory = args.RequireString("outputDirectory");
            string prefix = args.GetOptionalString("prefix", "image");
            bool confirm = args.GetOptionalBool("confirm", false);
            bool includeInline = args.GetOptionalBool("includeInline", true);
            bool includeFloating = args.GetOptionalBool("includeFloating", true);
            int maxImages = args.GetOptionalInt("maxImages", 500);
            if (prefix == null)
            {
                prefix = "image";
            }
            prefix = SanitizeFileNamePrefix(prefix);
            if (maxImages < 1 || maxImages > 2000)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "maxImages 允许 1-2000，收到 " + maxImages + "。");
            }
            if (!includeInline && !includeFloating)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "includeInline 与 includeFloating 不能同时为 false。");
            }

            string outputDirectory;
            OperationResult directoryFailure =
                WordExtensionsUtil.PrepareOutputDirectory(requestedDirectory, out outputDirectory);
            if (directoryFailure != null)
            {
                return directoryFailure;
            }

            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    string documentPath = document.FullName ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(documentPath) || !File.Exists(documentPath))
                    {
                        return OperationResult.Failed(
                            OfficeErrorCode.FILE_NOT_FOUND,
                            "当前活动文档尚未保存到磁盘（FullName=「" + documentPath +
                            "」），无法通过临时副本导出图片。请先保存文档后重试。");
                    }

                    List<ImageAnchorInfo> anchors;
                    JArray warnings = new JArray();
                    int nonPictureInline = 0;
                    int nonPictureFloating = 0;
                    try
                    {
                        anchors = CollectImageAnchors(
                            document, includeInline, includeFloating,
                            warnings, out nonPictureInline, out nonPictureFloating);
                    }
                    catch (Exception ex)
                    {
                        return OfficeErrorMapper.ToOperationResult(ex, OfficeApp.Word);
                    }

                    if (anchors.Count == 0)
                    {
                        JObject emptyData = new JObject();
                        emptyData["outputDirectory"] = outputDirectory;
                        emptyData["images"] = new JArray();
                        emptyData["imageCount"] = 0;
                        emptyData["nonPictureInlineShapes"] = nonPictureInline;
                        emptyData["nonPictureFloatingShapes"] = nonPictureFloating;
                        emptyData["documentModified"] = false;
                        emptyData["note"] = "文档中没有可导出的内嵌/浮动图片。";
                        return ReadSucceed(
                            "文档中没有可导出的图片；未创建任何文件。",
                            emptyData, "Word:Document");
                    }
                    if (anchors.Count > maxImages)
                    {
                        anchors = anchors.GetRange(0, maxImages);
                        WordExtensionsUtil.AddWarning(
                            warnings,
                            "图片数量超过 maxImages=" + maxImages + "，仅导出前 " + maxImages + " 张。");
                    }
                    anchors.Sort(delegate(ImageAnchorInfo left, ImageAnchorInfo right)
                    {
                        return left.DocumentOrder.CompareTo(right.DocumentOrder);
                    });

                    string tempRoot = Path.Combine(
                        Path.GetTempPath(),
                        "AutoOfficeMCP_WordImages_" + Guid.NewGuid().ToString("N"));
                    try
                    {
                        Directory.CreateDirectory(tempRoot);
                        string htmlPath = Path.Combine(tempRoot, "extract.html");
                        JObject extraction = ExtractDocumentHtmlCopy(
                            application, document, htmlPath, warnings);
                        if (extraction == null)
                        {
                            return OperationResult.Failed(
                                OfficeErrorCode.INTERNAL_ERROR,
                                "临时副本另存为 HTML 失败，未导出任何图片。请检查 Word HTML 导出组件与临时目录权限：" +
                                tempRoot);
                        }

                        string filesDirectory = Path.Combine(tempRoot, "extract.files");
                        if (!Directory.Exists(filesDirectory))
                        {
                            JObject noFiles = new JObject();
                            noFiles["outputDirectory"] = outputDirectory;
                            noFiles["images"] = new JArray();
                            noFiles["imageCount"] = 0;
                            noFiles["documentModified"] = false;
                            noFiles["warnings"] = warnings;
                            noFiles["note"] = "Word 已生成 HTML，但未生成 extract.files 目录；" +
                                               "可能是文档图片为链接图片或 OLE 对象。";
                            return ReadSucceed(
                                "未从 Word HTML 副本中发现图片文件；未写入输出目录。",
                                noFiles, "Word:Document");
                        }

                        List<string> exportedFiles = EnumerateImageFiles(filesDirectory);
                        exportedFiles.Sort(StringComparer.OrdinalIgnoreCase);
                        if (exportedFiles.Count == 0)
                        {
                            JObject noFiles = new JObject();
                            noFiles["outputDirectory"] = outputDirectory;
                            noFiles["images"] = new JArray();
                            noFiles["imageCount"] = 0;
                            noFiles["documentModified"] = false;
                            noFiles["warnings"] = warnings;
                            noFiles["note"] = "extract.files 目录存在，但没有可识别的图片文件。";
                            return ReadSucceed(
                                "Word HTML 副本中没有图片文件；未写入输出目录。",
                                noFiles, "Word:Document");
                        }

                        int copyCount = Math.Min(exportedFiles.Count, anchors.Count);
                        List<string> targetPaths = new List<string>();
                        for (int index = 0; index < copyCount; index++)
                        {
                            string extension = Path.GetExtension(exportedFiles[index]);
                            string fileName = prefix + "_" +
                                (index + 1).ToString("000", CultureInfo.InvariantCulture) +
                                extension.ToLowerInvariant();
                            targetPaths.Add(Path.Combine(outputDirectory, fileName));
                        }

                        JArray collisions = new JArray();
                        foreach (string targetPath in targetPaths)
                        {
                            if (File.Exists(targetPath))
                            {
                                collisions.Add(targetPath);
                            }
                        }
                        if (collisions.Count > 0 && !confirm)
                        {
                            JObject collisionData = new JObject();
                            collisionData["outputDirectory"] = outputDirectory;
                            collisionData["collisions"] = collisions;
                            return OperationResult.Cancelled(
                                "输出目录中已有 " + collisions.Count +
                                " 个同名图片文件，未覆盖。如需覆盖，请显式传 confirm=true。首个冲突：" +
                                collisions[0].ToString());
                        }

                        JArray images = new JArray();
                        for (int index = 0; index < copyCount; index++)
                        {
                            string targetPath = targetPaths[index];
                            File.Copy(exportedFiles[index], targetPath, true);
                            FileInfo info = new FileInfo(targetPath);
                            ImageAnchorInfo anchor = anchors[index];
                            JObject item = new JObject();
                            item["fileName"] = Path.GetFileName(targetPath);
                            item["path"] = targetPath;
                            item["bytes"] = info.Length;
                            item["extension"] = Path.GetExtension(targetPath);
                            item["collection"] = anchor.Collection;
                            item["index"] = anchor.Index;
                            item["anchorParagraphIndex"] = anchor.AnchorParagraphIndex;
                            item["widthPt"] = anchor.WidthPt;
                            item["heightPt"] = anchor.HeightPt;
                            item["altText"] = anchor.AltText ?? string.Empty;
                            item["wrapType"] = anchor.WrapType ?? string.Empty;
                            images.Add(item);
                        }

                        JArray unmatchedExported = new JArray();
                        for (int index = copyCount; index < exportedFiles.Count; index++)
                        {
                            unmatchedExported.Add(Path.GetFileName(exportedFiles[index]));
                        }
                        JArray unmatchedAnchors = new JArray();
                        for (int index = copyCount; index < anchors.Count; index++)
                        {
                            unmatchedAnchors.Add(new JObject
                            {
                                { "collection", anchors[index].Collection },
                                { "index", anchors[index].Index },
                                { "anchorParagraphIndex", anchors[index].AnchorParagraphIndex }
                            });
                        }
                        if (unmatchedExported.Count > 0 || unmatchedAnchors.Count > 0)
                        {
                            WordExtensionsUtil.AddWarning(
                                warnings,
                                "Word HTML 导出的图片数（" + exportedFiles.Count +
                                "）与文档图片对象数（" + anchors.Count +
                                "）不一致，已按文档顺序配对；未配对项见 unmatchedExported/unmatchedAnchors。");
                        }

                        JObject data = new JObject();
                        data["outputDirectory"] = outputDirectory;
                        data["prefix"] = prefix;
                        data["imageCount"] = images.Count;
                        data["images"] = images;
                        data["documentImageObjects"] = anchors.Count;
                        data["exportedFileCount"] = exportedFiles.Count;
                        data["nonPictureInlineShapes"] = nonPictureInline;
                        data["nonPictureFloatingShapes"] = nonPictureFloating;
                        data["unmatchedExportedFiles"] = unmatchedExported;
                        data["unmatchedAnchors"] = unmatchedAnchors;
                        data["documentModified"] = false;
                        data["sourceSaved"] = document.Saved;
                        data["warnings"] = warnings;
                        data["note"] = "通过临时副本另存为 HTML，再复制 Word 生成的 extract.files 图片；" +
                                        "原文档不修改、不另存。未保存的编辑不会进入导出结果。" +
                                        "锚定段号为文档正文 1 基段号；HTML 导出文件与图片对象按文档顺序配对。";

                        return ReadSucceed(
                            "已导出 " + images.Count + " 张图片到 " + outputDirectory +
                            "（文档图片对象 " + anchors.Count + " 个，原文档未修改）。",
                            data, "Word:Document", "Directory:" + outputDirectory);
                    }
                    catch (Exception ex)
                    {
                        return OfficeErrorMapper.ToOperationResult(ex, OfficeApp.Word);
                    }
                    finally
                    {
                        TryDeleteDirectory(tempRoot);
                    }
                });
        }

        private static List<ImageAnchorInfo> CollectImageAnchors(
            Word.Document document,
            bool includeInline,
            bool includeFloating,
            JArray warnings,
            out int nonPictureInline,
            out int nonPictureFloating)
        {
            nonPictureInline = 0;
            nonPictureFloating = 0;
            List<ImageAnchorInfo> anchors = new List<ImageAnchorInfo>();
            Word.InlineShapes inlineShapes = null;
            Word.Shapes shapes = null;
            try
            {
                if (includeInline)
                {
                    inlineShapes = document.InlineShapes;
                    int count = inlineShapes == null ? 0 : inlineShapes.Count;
                    for (int index = 1; index <= count; index++)
                    {
                        Word.InlineShape shape = null;
                        Word.Range range = null;
                        try
                        {
                            shape = inlineShapes[index];
                            bool isPicture = shape.Type == Word.WdInlineShapeType.wdInlineShapePicture ||
                                             shape.Type == Word.WdInlineShapeType.wdInlineShapeLinkedPicture;
                            if (!isPicture)
                            {
                                nonPictureInline++;
                                continue;
                            }
                            range = shape.Range;
                            ImageAnchorInfo info = new ImageAnchorInfo();
                            info.Collection = "inline";
                            info.Index = index;
                            info.DocumentOrder = range == null ? index : range.Start;
                            info.AnchorParagraphIndex = WordComUtil.FindParagraphIndexAt(
                                document, range == null ? 0 : range.Start);
                            info.WidthPt = (double)shape.Width;
                            info.HeightPt = (double)shape.Height;
                            info.AltText = WordExtensionsUtil.SafeReadAlternativeText(shape);
                            info.WrapType = "inline";
                            anchors.Add(info);
                        }
                        catch (Exception ex)
                        {
                            WordExtensionsUtil.AddWarning(
                                warnings,
                                "第 " + index + " 个内嵌对象读取失败，已跳过：" +
                                OfficeErrorMapper.DescribeException(ex, OfficeApp.Word));
                        }
                        finally
                        {
                            ComHelper.Release(range, shape);
                        }
                    }
                }
                if (includeFloating)
                {
                    shapes = document.Shapes;
                    int count = shapes == null ? 0 : shapes.Count;
                    for (int index = 1; index <= count; index++)
                    {
                        Word.Shape shape = null;
                        Word.Range anchor = null;
                        try
                        {
                            shape = shapes[index];
                            int typeCode = WordImageUtil.ReadFloatingTypeCode(shape);
                            bool isPicture = typeCode == 13 || typeCode == 11 ||
                                             typeCode == 28 || typeCode == 29;
                            if (!isPicture)
                            {
                                nonPictureFloating++;
                                continue;
                            }
                            anchor = shape.Anchor;
                            ImageAnchorInfo info = new ImageAnchorInfo();
                            info.Collection = "floating";
                            info.Index = index;
                            info.DocumentOrder = anchor == null ? index : anchor.Start;
                            info.AnchorParagraphIndex = WordComUtil.FindParagraphIndexAt(
                                document, anchor == null ? 0 : anchor.Start);
                            info.WidthPt = (double)shape.Width;
                            info.HeightPt = (double)shape.Height;
                            info.AltText = WordExtensionsUtil.SafeReadAlternativeText(shape);
                            info.WrapType = WordImageUtil.ReadFloatingType(shape);
                            anchors.Add(info);
                        }
                        catch (Exception ex)
                        {
                            WordExtensionsUtil.AddWarning(
                                warnings,
                                "第 " + index + " 个浮动对象读取失败，已跳过：" +
                                OfficeErrorMapper.DescribeException(ex, OfficeApp.Word));
                        }
                        finally
                        {
                            ComHelper.Release(anchor, shape);
                        }
                    }
                }
                return anchors;
            }
            finally
            {
                ComHelper.Release(shapes, inlineShapes);
            }
        }

        private static JObject ExtractDocumentHtmlCopy(
            Word.Application application,
            Word.Document sourceDocument,
            string htmlPath,
            JArray warnings)
        {
            return application.Guard(delegate
            {
                Word.Documents documents = null;
                Word.Document tempDocument = null;
                object missing = Type.Missing;
                object visible = false;
                object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                object fileFormat = Word.WdSaveFormat.wdFormatHTML;
                try
                {
                    if (!sourceDocument.Saved)
                    {
                        WordExtensionsUtil.AddWarning(
                            warnings,
                            "当前文档有未保存修改；图片导出基于磁盘上最后保存的版本。");
                    }
                    documents = application.Documents;
                    tempDocument = documents.Add(ref missing, ref missing, ref missing, ref visible);
                    if (tempDocument == null)
                    {
                        return null;
                    }
                    Word.Range content = null;
                    try
                    {
                        content = tempDocument.Content;
                        content.InsertFile(
                            sourceDocument.FullName,
                            ref missing, ref missing, ref missing, ref missing);
                    }
                    finally
                    {
                        ComHelper.Release(content);
                    }
                    object htmlFile = htmlPath;
                    tempDocument.SaveAs2(
                        ref htmlFile, ref fileFormat, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing);
                    JObject result = new JObject();
                    result["htmlPath"] = htmlPath;
                    return result;
                }
                catch (Exception ex)
                {
                    WordExtensionsUtil.AddWarning(
                        warnings,
                        "临时副本 HTML 导出失败：" +
                        OfficeErrorMapper.DescribeException(ex, OfficeApp.Word));
                    return null;
                }
                finally
                {
                    if (tempDocument != null)
                    {
                        try
                        {
                            tempDocument.Close(ref saveChanges, ref missing, ref missing);
                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine("[word-ext] 关闭图片导出临时文档失败：" + ex.Message);
                        }
                    }
                    ComHelper.Release(tempDocument, documents);
                }
            });
        }

        private static List<string> EnumerateImageFiles(string directory)
        {
            List<string> result = new List<string>();
            if (!Directory.Exists(directory))
            {
                return result;
            }
            string[] files = Directory.GetFiles(directory, "*", SearchOption.TopDirectoryOnly);
            foreach (string file in files)
            {
                string extension = Path.GetExtension(file).ToLowerInvariant();
                if (extension == ".png" || extension == ".jpg" || extension == ".jpeg" ||
                    extension == ".gif" || extension == ".bmp" || extension == ".emf" ||
                    extension == ".wmf" || extension == ".tif" || extension == ".tiff" ||
                    extension == ".svg")
                {
                    result.Add(file);
                }
            }
            return result;
        }

        private static string SanitizeFileNamePrefix(string prefix)
        {
            string value = string.IsNullOrWhiteSpace(prefix) ? "image" : prefix.Trim();
            char[] invalid = Path.GetInvalidFileNameChars();
            char[] buffer = value.ToCharArray();
            for (int index = 0; index < buffer.Length; index++)
            {
                foreach (char invalidChar in invalid)
                {
                    if (buffer[index] == invalidChar)
                    {
                        buffer[index] = '_';
                        break;
                    }
                }
            }
            string result = new string(buffer).Trim('.', ' ');
            return string.IsNullOrWhiteSpace(result) ? "image" : result;
        }

        private static void TryDeleteDirectory(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }
            try
            {
                if (Directory.Exists(path))
                {
                    Directory.Delete(path, true);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[word-ext] 清理临时目录失败：" + path + "；" + ex.Message);
            }
        }

        private sealed class ImageAnchorInfo
        {
            public string Collection { get; set; }

            public int Index { get; set; }

            public int DocumentOrder { get; set; }

            public int AnchorParagraphIndex { get; set; }

            public double WidthPt { get; set; }

            public double HeightPt { get; set; }

            public string AltText { get; set; }

            public string WrapType { get; set; }

            public ImageAnchorInfo()
            {
                Collection = string.Empty;
                AltText = string.Empty;
                WrapType = string.Empty;
            }
        }

        // ------------------------------------------------------------------
        // word_doc_merge
        // ------------------------------------------------------------------

        private OperationResult HandleDocMerge(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string sourcePath = args.RequireString("sourcePath");
            string position = (args.GetOptionalString("position", "cursor") ?? "cursor")
                .Trim().ToLowerInvariant();
            string mode = (args.GetOptionalString("mode", "preserve") ?? "preserve")
                .Trim().ToLowerInvariant();
            if (!File.Exists(sourcePath))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.FILE_NOT_FOUND,
                    "待插入文档不存在：" + sourcePath + "。请检查路径与文件名后重试。");
            }
            if (position != "start" && position != "end" &&
                position != "cursor" && position != "paragraph")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "position 只支持 start/end/cursor/paragraph，收到「" + position +
                    "」；position=paragraph 时必须同时提供 paragraphIndex（1 基）。");
            }
            if (mode != "preserve" && mode != "match")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "mode 只支持 preserve（保留源格式）/match（套用当前文档样式），收到「" +
                    mode + "」。");
            }
            int paragraphIndex = args.GetOptionalInt("paragraphIndex", 0);
            if (position == "paragraph" && paragraphIndex <= 0)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "position=paragraph 时必须提供 paragraphIndex（1 基，指向插入位置之前的段落）。" +
                    "收到 paragraphIndex=" + paragraphIndex + "。");
            }

            return RunWriteTool(
                application,
                WordExtensionsToolNames.DocMerge,
                false,
                delegate(Word.Document document)
                {
                    if (position == "start")
                    {
                        return WordWriteTargets.DocumentStart();
                    }
                    if (position == "end")
                    {
                        return WordWriteTargets.DocumentEnd(document);
                    }
                    if (position == "paragraph")
                    {
                        WordComUtil.ValidateParagraphIndex(document, paragraphIndex);
                        return WordWriteTargets.Paragraph(
                            document, paragraphIndex, "Word:Paragraph:" + paragraphIndex, true);
                    }
                    return WordWriteTargets.Cursor(application);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Range insertion = null;
                    Word.Range content = null;
                    try
                    {
                        content = document.Content;
                        int beforeCharacters = content == null ? 0 : (content.Text ?? string.Empty).Length;
                        int beforeParagraphs = WordComUtil.GetParagraphCount(document);
                        int insertionStart = target.Start;
                        bool needsLeadingBreak = position == "cursor" || position == "paragraph";
                        int separatorCharacters = 0;

                        if (needsLeadingBreak)
                        {
                            insertion = WordComUtil.CreateRange(
                                document, insertionStart, insertionStart);
                            insertion.InsertAfter("\r");
                            insertionStart++;
                            separatorCharacters = 1;
                            ComHelper.Release(insertion);
                            insertion = null;
                        }

                        insertion = WordComUtil.CreateRange(
                            document, insertionStart, insertionStart);
                        object missing = Type.Missing;
                        insertion.InsertFile(
                            Path.GetFullPath(sourcePath),
                            ref missing, ref missing, ref missing, ref missing);

                        content = document.Content;
                        int afterCharacters = content == null ? 0 : (content.Text ?? string.Empty).Length;
                        int afterParagraphs = WordComUtil.GetParagraphCount(document);
                        int totalInsertedCharacters = afterCharacters - beforeCharacters;
                        int insertedCharacters = Math.Max(
                            0, totalInsertedCharacters - separatorCharacters);
                        int insertedParagraphs = afterParagraphs - beforeParagraphs;
                        int matchedParagraphs = 0;
                        JArray warnings = new JArray();

                        if (mode == "match" && insertedCharacters > 0)
                        {
                            matchedParagraphs = MatchInsertedStyles(
                                document,
                                insertionStart,
                                insertionStart + insertedCharacters,
                                warnings);
                        }

                        bool changed = insertedCharacters > 0 || insertedParagraphs > 0 ||
                                       afterCharacters != beforeCharacters;
                        JObject data = new JObject();
                        data["sourcePath"] = Path.GetFullPath(sourcePath);
                        data["sourceBytes"] = new FileInfo(sourcePath).Length;
                        data["position"] = position;
                        data["paragraphIndex"] = paragraphIndex > 0 ? paragraphIndex : 0;
                        data["mode"] = mode;
                        data["charactersBefore"] = beforeCharacters;
                        data["charactersAfter"] = afterCharacters;
                        data["totalInsertedCharacters"] = totalInsertedCharacters;
                        data["separatorCharacters"] = separatorCharacters;
                        data["insertedCharacters"] = insertedCharacters;
                        data["paragraphsBefore"] = beforeParagraphs;
                        data["paragraphsAfter"] = afterParagraphs;
                        data["insertedParagraphs"] = insertedParagraphs;
                        data["matchedParagraphs"] = matchedParagraphs;
                        data["warnings"] = warnings;
                        data["note"] = mode == "preserve"
                            ? "使用 Word Range.InsertFile 插入源文档，保留源文档样式与直接格式。"
                            : "使用 Word Range.InsertFile 插入后，对插入区段清除直接格式并按同名标题级别套用当前文档样式；" +
                              "源文档中的表格/图片对象保留。";

                        if (!changed)
                        {
                            return WordWriteResult.Ok(
                                "源文档插入调用已完成，但未观察到文档字符/段落数变化（源文档可能为空）。",
                                data,
                                false);
                        }

                        WordWriteResult result = WordWriteResult.Ok(
                            "已把 " + Path.GetFileName(sourcePath) + " 插入到 " + target.Label +
                            "（插入 " + insertedCharacters + " 字符、" + insertedParagraphs +
                            " 段；mode=" + mode + "）。",
                            data,
                            true);
                        result.AddTargetRef("Word:Document");
                        return result;
                    }
                    finally
                    {
                        ComHelper.Release(content, insertion);
                    }
                });
        }

        private static int MatchInsertedStyles(
            Word.Document document,
            int insertedStart,
            int insertedEnd,
            JArray warnings)
        {
            if (document == null || insertedEnd <= insertedStart)
            {
                return 0;
            }
            string bodyStyle = ResolveBodyStyleName(document);
            Word.Paragraphs paragraphs = null;
            List<int> indexes = new List<int>();
            try
            {
                paragraphs = document.Paragraphs;
                int count = paragraphs == null ? 0 : paragraphs.Count;
                for (int index = 1; index <= count; index++)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range range = null;
                    try
                    {
                        paragraph = paragraphs[index];
                        range = paragraph.Range;
                        if (range.Start >= insertedStart && range.End <= insertedEnd + 1)
                        {
                            indexes.Add(index);
                        }
                    }
                    finally
                    {
                        ComHelper.Release(range, paragraph);
                    }
                }
            }
            finally
            {
                ComHelper.Release(paragraphs);
            }

            indexes.Sort();
            indexes.Reverse();
            int matched = 0;
            foreach (int index in indexes)
            {
                Word.Paragraph paragraph = null;
                Word.Range range = null;
                Word.Font font = null;
                Word.ParagraphFormat paragraphFormat = null;
                try
                {
                    paragraph = WordComUtil.GetParagraph(document, index);
                    range = paragraph.Range;
                    string text = WordExtensionsUtil.ReadParagraphText(paragraph);
                    if (text.Length == 0)
                    {
                        continue;
                    }
                    string currentStyle = WordComUtil.ReadStyleName(range);
                    int level = WordExtensionsUtil.TryGetHeadingLevel(currentStyle);
                    string targetStyle = level > 0
                        ? WordExtensionsUtil.ResolveHeadingStyleName(document, level)
                        : bodyStyle;
                    if (!string.IsNullOrWhiteSpace(targetStyle))
                    {
                        WordExtensionsUtil.ApplyParagraphStyle(
                            document, paragraph, targetStyle, true, true);
                        matched++;
                    }
                    else
                    {
                        WordExtensionsUtil.AddWarning(
                            warnings,
                            "第 " + index + " 段找不到目标样式（level=" + level + "），保留原格式。");
                    }
                }
                catch (Exception ex)
                {
                    WordExtensionsUtil.AddWarning(
                        warnings,
                        "第 " + index + " 段匹配当前样式失败：" +
                        OfficeErrorMapper.DescribeException(ex, OfficeApp.Word));
                }
                finally
                {
                    ComHelper.Release(paragraphFormat, font, range, paragraph);
                }
            }
            return matched;
        }

        private static string ResolveBodyStyleName(Word.Document document)
        {
            string[] candidates = new string[] { "正文", "Normal", "Body Text" };
            foreach (string candidate in candidates)
            {
                string resolved;
                if (WordStyleCatalog.TryGetStyleName(document, candidate, out resolved))
                {
                    return resolved;
                }
            }
            return null;
        }
    }
}
