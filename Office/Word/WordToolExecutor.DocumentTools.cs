using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Newtonsoft.Json.Linq;
using Office.Config;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    public sealed partial class WordToolExecutor
    {
        private void RegisterDocumentHandlers(
            Dictionary<string, Func<Word.Application, JObject, OperationResult>> handlers)
        {
            handlers[WordToolNames.GenerateToc] = HandleGenerateToc;
            handlers[WordToolNames.Beautify] = HandleBeautify;
            handlers[WordToolNames.RebuildNumbering] = HandleRebuildNumbering;
            handlers[WordToolNames.SetPageSetup] = HandleSetPageSetup;
            handlers[WordToolNames.HeaderFooter] = HandleHeaderFooter;
            handlers[WordToolNames.Export] = HandleExport;
            handlers[WordToolNames.DocProperties] = HandleDocProperties;
            handlers[WordToolNames.FieldUpdateAll] = HandleFieldUpdateAll;
        }

        private OperationResult HandleGenerateToc(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string position = (args.GetOptionalString("position", "start") ?? "start").Trim().ToLowerInvariant();
            int levels = args.GetOptionalInt("levels", 3);
            bool includePageNumbers = args.GetOptionalBool("includePageNumbers", true);
            if (position != "start" && position != "cursor" && position != "end")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "position 只支持 start/cursor/end，收到 " + position + "。");
            }
            if (levels < 1 || levels > 9)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "levels 必须在 1-9 之间，收到 " + levels + "。");
            }

            return RunWriteTool(
                application,
                WordToolNames.GenerateToc,
                false,
                delegate(Word.Document document)
                {
                    return SelectTocTarget(application, document, position);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Range insertionRange = null;
                    Word.TablesOfContents existing = null;
                    Word.TablesOfContents contents = null;
                    Word.TableOfContents toc = null;
                    Word.Range tocRange = null;
                    try
                    {
                        insertionRange = WordComUtil.CreateRange(document, target.Start, target.End);
                        existing = document.TablesOfContents;
                        int existingCount = existing == null ? 0 : existing.Count;
                        for (int index = existingCount; index >= 1; index--)
                        {
                            Word.TableOfContents oldToc = null;
                            try
                            {
                                oldToc = existing[index];
                                oldToc.Delete();
                            }
                            finally
                            {
                                ComHelper.Release(oldToc);
                            }
                        }

                        object useHeadingStyles = true;
                        object upperHeadingLevel = 1;
                        object lowerHeadingLevel = levels;
                        object useFields = false;
                        object tableId = Type.Missing;
                        object rightAlignPageNumbers = true;
                        object includePageNumbersValue = includePageNumbers;
                        object addedStyles = Type.Missing;
                        object useHyperlinks = true;
                        object hidePageNumbersInWeb = false;
                        object useOutlineLevels = false;

                        contents = document.TablesOfContents;
                        toc = contents.Add(
                            insertionRange,
                            ref useHeadingStyles,
                            ref upperHeadingLevel,
                            ref lowerHeadingLevel,
                            ref useFields,
                            ref tableId,
                            ref rightAlignPageNumbers,
                            ref includePageNumbersValue,
                            ref addedStyles,
                            ref useHyperlinks,
                            ref hidePageNumbersInWeb,
                            ref useOutlineLevels);
                        toc.Update();
                        toc.UpdatePageNumbers();

                        int newCount = contents.Count;
                        tocRange = toc.Range;
                        string tocText = tocRange == null ? string.Empty : (tocRange.Text ?? string.Empty);
                        JObject data = new JObject();
                        data["position"] = position;
                        data["levels"] = levels;
                        data["includePageNumbers"] = includePageNumbers;
                        data["tablesOfContentsCount"] = newCount;
                        data["previousTocCount"] = existingCount;
                        data["tocRangeStart"] = tocRange == null ? 0 : tocRange.Start;
                        data["tocRangeEnd"] = tocRange == null ? 0 : tocRange.End;
                        data["tocTextPreview"] = WordComUtil.NormalizePreview(tocText, 240);

                        bool changed = true;
                        WordWriteResult result = WordWriteResult.Ok(
                            "已生成 " + levels + " 级目录（原有目录 " + existingCount + " 个已先删除），并调用 Update/UpdatePageNumbers。",
                            data,
                            changed);
                        result.AddTargetRef("Word:TOC");
                        return result;
                    }
                    finally
                    {
                        ComHelper.Release(tocRange, toc, contents, existing, insertionRange);
                    }
                });
        }

        private OperationResult HandleSetPageSetup(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            double? topMarginPt = ResolveMargin(args, "topMarginPt", "topMarginCm");
            double? bottomMarginPt = ResolveMargin(args, "bottomMarginPt", "bottomMarginCm");
            double? leftMarginPt = ResolveMargin(args, "leftMarginPt", "leftMarginCm");
            double? rightMarginPt = ResolveMargin(args, "rightMarginPt", "rightMarginCm");
            Word.WdPaperSize? paperSize = args.Has("paperSize")
                ? WordComUtil.ParsePaperSize(args.GetOptionalString("paperSize", null))
                : (Word.WdPaperSize?)null;
            Word.WdOrientation? orientation = null;
            if (args.Has("orientation"))
            {
                string orientationText = (args.GetOptionalString("orientation", null) ?? string.Empty)
                    .Trim().ToLowerInvariant();
                if (orientationText == "portrait" || orientationText == "纵向")
                {
                    orientation = Word.WdOrientation.wdOrientPortrait;
                }
                else if (orientationText == "landscape" || orientationText == "横向")
                {
                    orientation = Word.WdOrientation.wdOrientLandscape;
                }
                else
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.INVALID_PARAMS,
                        "orientation 只支持 portrait/landscape，收到 " + orientationText + "。");
                }
            }
            int sectionIndex = args.GetOptionalInt("sectionIndex", 0);

            if (!topMarginPt.HasValue && !bottomMarginPt.HasValue && !leftMarginPt.HasValue &&
                !rightMarginPt.HasValue && !paperSize.HasValue && !orientation.HasValue)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "word_set_page_setup 至少需要一个设置：topMarginPt/topMarginCm/bottomMarginPt/bottomMarginCm/" +
                    "leftMarginPt/leftMarginCm/rightMarginPt/rightMarginCm/paperSize/orientation。收到 keys=" +
                    args.DescribeKeys());
            }

            return RunWriteTool(
                application,
                WordToolNames.SetPageSetup,
                false,
                delegate(Word.Document document)
                {
                    if (sectionIndex > 0)
                    {
                        WordComUtil.ValidateSectionIndex(document, sectionIndex);
                    }
                    return WordWriteTargets.None("页面设置", "Word:PageSetup");
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Sections sections = null;
                    JArray sectionResults = new JArray();
                    JArray failures = new JArray();
                    bool changed = false;
                    try
                    {
                        sections = document.Sections;
                        int count = sections == null ? 0 : sections.Count;
                        if (count == 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INTERNAL_ERROR,
                                "文档没有可用于设置页面的节（Sections.Count=0）。",
                                null,
                                false);
                        }
                        int start = sectionIndex > 0 ? sectionIndex : 1;
                        int end = sectionIndex > 0 ? sectionIndex : count;
                        for (int index = end; index >= start; index--)
                        {
                            Word.Section section = null;
                            Word.PageSetup pageSetup = null;
                            try
                            {
                                section = sections[index];
                                pageSetup = section.PageSetup;
                                JObject before = WordComUtil.ReadPageSetup(pageSetup);
                                if (paperSize.HasValue)
                                {
                                    pageSetup.PaperSize = paperSize.Value;
                                }
                                if (orientation.HasValue)
                                {
                                    pageSetup.Orientation = orientation.Value;
                                }
                                if (topMarginPt.HasValue)
                                {
                                    pageSetup.TopMargin = (float)topMarginPt.Value;
                                }
                                if (bottomMarginPt.HasValue)
                                {
                                    pageSetup.BottomMargin = (float)bottomMarginPt.Value;
                                }
                                if (leftMarginPt.HasValue)
                                {
                                    pageSetup.LeftMargin = (float)leftMarginPt.Value;
                                }
                                if (rightMarginPt.HasValue)
                                {
                                    pageSetup.RightMargin = (float)rightMarginPt.Value;
                                }
                                JObject after = WordComUtil.ReadPageSetup(pageSetup);
                                if (!JToken.DeepEquals(before, after))
                                {
                                    changed = true;
                                }
                                JArray verification = VerifyPageSetup(
                                    before, after, topMarginPt, bottomMarginPt,
                                    leftMarginPt, rightMarginPt, paperSize, orientation);
                                bool sectionMatched = true;
                                foreach (JToken item in verification)
                                {
                                    if (item["matched"] != null && !item["matched"].Value<bool>())
                                    {
                                        sectionMatched = false;
                                        failures.Add(new JObject
                                        {
                                            { "sectionIndex", index },
                                            { "message", item["property"].ToString() + " 回读不一致：" +
                                                         "期望 " + item["expected"] + "，实际 " + item["actual"] }
                                        });
                                        break;
                                    }
                                }
                                sectionResults.Add(new JObject
                                {
                                    { "sectionIndex", index },
                                    { "before", before },
                                    { "after", after },
                                    { "matched", sectionMatched }
                                });
                            }
                            catch (Exception ex)
                            {
                                failures.Add(new JObject
                                {
                                    { "sectionIndex", index },
                                    { "message", OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) }
                                });
                            }
                            finally
                            {
                                ComHelper.Release(pageSetup, section);
                            }
                        }
                    }
                    finally
                    {
                        ComHelper.Release(sections);
                    }

                    JObject data = new JObject();
                    data["sections"] = sectionResults;
                    data["failures"] = failures;
                    data["sectionIndex"] = sectionIndex;
                    if (failures.Count > 0)
                    {
                        return WordWriteResult.Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "页面设置部分失败：" + failures[0]["message"].ToString() +
                            "。请先确认节索引/纸张名/页边距范围。",
                            data,
                            changed);
                    }
                    return WordWriteResult.Ok(
                        "页面设置已应用到 " + sectionResults.Count + " 个节，关键值回读一致。",
                        data,
                        changed);
                });
        }

        private OperationResult HandleHeaderFooter(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string headerText = args.Has("headerText") ? args.GetOptionalString("headerText", string.Empty) : null;
            string footerText = args.Has("footerText") ? args.GetOptionalString("footerText", string.Empty) : null;
            string pageNumbers = (args.GetOptionalString("pageNumbers", "none") ?? "none").Trim();
            string pageNumberTarget = (args.GetOptionalString("pageNumberTarget", "footer") ?? "footer")
                .Trim().ToLowerInvariant();
            string headerType = (args.GetOptionalString("headerType", "primary") ?? "primary")
                .Trim().ToLowerInvariant();
            int sectionIndex = args.GetOptionalInt("sectionIndex", 1);
            bool? linkToPrevious = ReadNullableBool(args, "linkToPrevious");

            bool pageNumberRequested = !string.IsNullOrWhiteSpace(pageNumbers) &&
                                       !string.Equals(pageNumbers, "none", StringComparison.OrdinalIgnoreCase);
            if (headerText == null && footerText == null && !pageNumberRequested && !linkToPrevious.HasValue)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "word_header_footer 至少需要 headerText/footerText/pageNumbers/linkToPrevious 之一。");
            }
            if (pageNumbers != "none" && pageNumbers != "page" && pageNumbers != "x" &&
                pageNumbers != "x/y" && pageNumbers != "xOfY" && pageNumbers != "pageOfPages" &&
                pageNumbers != "第X页共Y页" && pageNumbers != "chinese")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "pageNumbers 只支持 none/page/x/xOfY/pageOfPages/第X页共Y页，收到 " + pageNumbers + "。");
            }
            if (pageNumberTarget != "header" && pageNumberTarget != "footer")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "pageNumberTarget 只支持 header/footer，收到 " + pageNumberTarget + "。");
            }

            return RunWriteTool(
                application,
                WordToolNames.HeaderFooter,
                false,
                delegate(Word.Document document)
                {
                    WordComUtil.ValidateSectionIndex(document, sectionIndex);
                    return WordWriteTargets.None("页眉页脚", "Word:Section:" + sectionIndex);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Sections sections = null;
                    Word.Section section = null;
                    Word.HeadersFooters headers = null;
                    Word.HeadersFooters footers = null;
                    Word.HeaderFooter header = null;
                    Word.HeaderFooter footer = null;
                    Word.Range headerRange = null;
                    Word.Range footerRange = null;
                    try
                    {
                        sections = document.Sections;
                        section = sections[sectionIndex];
                        headers = section.Headers;
                        footers = section.Footers;
                        Word.WdHeaderFooterIndex footerIndex = MapHeaderFooterIndex(headerType);
                        header = headers[footerIndex];
                        footer = footers[footerIndex];
                        headerRange = header.Range;
                        footerRange = footer.Range;

                        string beforeHeader = WordComUtil.NormalizePreview(headerRange.Text, 200);
                        string beforeFooter = WordComUtil.NormalizePreview(footerRange.Text, 200);
                        int beforeHeaderFields = CountFields(headerRange);
                        int beforeFooterFields = CountFields(footerRange);
                        JArray warnings = new JArray();

                        if (linkToPrevious.HasValue)
                        {
                            TrySetLinkToPrevious(header, linkToPrevious.Value, warnings);
                            TrySetLinkToPrevious(footer, linkToPrevious.Value, warnings);
                        }
                        if (headerText != null)
                        {
                            header.Range.Text = WordFormatUtil.NormalizeForWord(headerText);
                        }
                        if (footerText != null)
                        {
                            footer.Range.Text = WordFormatUtil.NormalizeForWord(footerText);
                        }
                        if (pageNumberRequested)
                        {
                            Word.HeaderFooter targetFooter =
                                pageNumberTarget == "header" ? header : footer;
                            AppendPageNumberFields(document, targetFooter, pageNumbers);
                        }

                        ComHelper.Release(headerRange);
                        ComHelper.Release(footerRange);
                        headerRange = header.Range;
                        footerRange = footer.Range;
                        string afterHeader = WordComUtil.NormalizePreview(headerRange.Text, 200);
                        string afterFooter = WordComUtil.NormalizePreview(footerRange.Text, 200);
                        int afterHeaderFields = CountFields(headerRange);
                        int afterFooterFields = CountFields(footerRange);

                        JObject data = new JObject();
                        data["sectionIndex"] = sectionIndex;
                        data["headerType"] = headerType;
                        data["headerTextBefore"] = beforeHeader;
                        data["headerTextAfter"] = afterHeader;
                        data["footerTextBefore"] = beforeFooter;
                        data["footerTextAfter"] = afterFooter;
                        data["headerFieldCountBefore"] = beforeHeaderFields;
                        data["headerFieldCountAfter"] = afterHeaderFields;
                        data["footerFieldCountBefore"] = beforeFooterFields;
                        data["footerFieldCountAfter"] = afterFooterFields;
                        data["pageNumbers"] = pageNumbers;
                        data["pageNumberTarget"] = pageNumberTarget;
                        data["warnings"] = warnings;

                        bool changed = !string.Equals(beforeHeader, afterHeader, StringComparison.Ordinal) ||
                                       !string.Equals(beforeFooter, afterFooter, StringComparison.Ordinal) ||
                                       beforeHeaderFields != afterHeaderFields ||
                                       beforeFooterFields != afterFooterFields;
                        if (headerText != null &&
                            afterHeader.IndexOf(WordComUtil.NormalizePreview(headerText, 200),
                                StringComparison.Ordinal) < 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "页眉写入后回读不含请求文本。期望「" +
                                WordComUtil.NormalizePreview(headerText, 80) + "」，实际「" + afterHeader +
                                "」。请确认该节页眉未链接到上一节或被锁定。",
                                data,
                                changed);
                        }
                        if (footerText != null &&
                            afterFooter.IndexOf(WordComUtil.NormalizePreview(footerText, 200),
                                StringComparison.Ordinal) < 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "页脚写入后回读不含请求文本。期望「" +
                                WordComUtil.NormalizePreview(footerText, 80) + "」，实际「" + afterFooter +
                                "」。请确认该节页脚未链接到上一节或被锁定。",
                                data,
                                changed);
                        }
                        if (pageNumberRequested && afterFooterFields + afterHeaderFields <=
                            beforeFooterFields + beforeHeaderFields)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "页码域写入后未检测到域数量增加（页眉 " + beforeHeaderFields + "→" +
                                afterHeaderFields + "，页脚 " + beforeFooterFields + "→" +
                                afterFooterFields + "）。请检查页眉页脚是否受保护。",
                                data,
                                changed);
                        }

                        WordWriteResult result = WordWriteResult.Ok(
                            "已更新第 " + sectionIndex + " 节" + headerType + "页眉页脚" +
                            (pageNumberRequested ? "并插入页码域（" + pageNumbers + "）" : "") + "。",
                            data,
                            changed);
                        result.AddTargetRef("Word:Section:" + sectionIndex);
                        return result;
                    }
                    finally
                    {
                        ComHelper.Release(footerRange, headerRange, footer, header, footers, headers, section, sections);
                    }
                });
        }

        private OperationResult HandleExport(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string requestedPath = args.RequireString("path");
            string format = (args.GetOptionalString("format", "pdf") ?? "pdf").Trim().ToLowerInvariant();
            bool confirm = args.GetOptionalBool("confirm", false);
            if (format != "pdf" && format != "xps" && format != "txt")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "format 只支持 pdf/xps/txt，收到 " + format + "。");
            }

            string path = NormalizeExportPath(requestedPath, format);
            string directory = Path.GetDirectoryName(path);
            if (string.IsNullOrWhiteSpace(directory))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "导出路径缺少有效目录：" + path + "。");
            }
            string directoryError;
            if (!WorkspacePathResolver.TryEnsureOutputDirectory(directory, out directoryError))
            {
                return OperationResult.Failed(OfficeErrorCode.READ_ONLY, directoryError);
            }
            bool existed = File.Exists(path);
            if (existed && !confirm)
            {
                return OperationResult.Cancelled(
                    "导出目标已存在：" + path + "。如需覆盖，请显式传 confirm=true；本次未写入任何文件。");
            }

            return RunWriteTool(
                application,
                WordToolNames.Export,
                true,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.None("导出文件", "File:" + path);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Range content = null;
                    try
                    {
                        if (format == "pdf" || format == "xps")
                        {
                            object missing = Type.Missing;
                            Word.WdExportFormat exportFormat = format == "pdf"
                                ? Word.WdExportFormat.wdExportFormatPDF
                                : Word.WdExportFormat.wdExportFormatXPS;
                            document.ExportAsFixedFormat(
                                path,
                                exportFormat,
                                false,
                                Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                                Word.WdExportRange.wdExportAllDocument,
                                1,
                                1,
                                Word.WdExportItem.wdExportDocumentContent,
                                true,
                                true,
                                Word.WdExportCreateBookmarks.wdExportCreateHeadingBookmarks,
                                true,
                                true,
                                false,
                                ref missing);
                        }
                        else
                        {
                            content = document.Content;
                            content.ExportFragment(path, Word.WdSaveFormat.wdFormatUnicodeText);
                        }

                        if (!File.Exists(path))
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INTERNAL_ERROR,
                                "Word 导出调用已返回，但目标文件不存在：" + path + "。请检查路径权限或 Word 导出组件。",
                                null,
                                false);
                        }
                        FileInfo info = new FileInfo(path);
                        if (info.Length <= 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INTERNAL_ERROR,
                                "导出文件已创建但大小为 0：" + path + "。",
                                null,
                                false);
                        }

                        JObject data = new JObject();
                        data["path"] = path;
                        data["format"] = format;
                        data["bytes"] = info.Length;
                        data["overwroteExisting"] = existed;
                        data["activeDocumentModified"] = false;
                        WordWriteResult result = WordWriteResult.Ok(
                            "已导出 " + format.ToUpperInvariant() + "：" + path + "（" + info.Length +
                            " 字节）。当前活动文档未被修改。",
                            data,
                            false);
                        result.AddTargetRef("File:" + path);
                        return result;
                    }
                    finally
                    {
                        ComHelper.Release(content);
                    }
                });
        }

        private OperationResult HandleDocProperties(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            List<string> requested = new List<string>();
            for (int index = 0; index < WordDocumentProperties.Count; index++)
            {
                string jsonName = WordDocumentProperties.GetJsonName(index);
                if (args.Has(jsonName))
                {
                    requested.Add(jsonName);
                }
            }

            if (requested.Count == 0)
            {
                return RunReadTool(
                    application,
                    delegate(Word.Document document)
                    {
                        JObject properties = WordDocumentProperties.ReadAll(document);
                        JObject data = new JObject();
                        data["properties"] = properties;
                        data["documentName"] = document.Name;
                        return ReadSucceed("已读取文档属性。", data, "Word:DocumentProperties");
                    });
            }

            JObject requestedValues = new JObject();
            for (int index = 0; index < WordDocumentProperties.Count; index++)
            {
                string jsonName = WordDocumentProperties.GetJsonName(index);
                if (requested.Contains(jsonName))
                {
                    requestedValues[jsonName] = args.GetOptionalString(jsonName, string.Empty) ?? string.Empty;
                }
            }

            return RunWriteTool(
                application,
                WordToolNames.DocProperties,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.None("文档属性", "Word:DocumentProperties");
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    JObject before = WordDocumentProperties.ReadAll(document);
                    for (int index = 0; index < WordDocumentProperties.Count; index++)
                    {
                        string jsonName = WordDocumentProperties.GetJsonName(index);
                        if (requested.Contains(jsonName))
                        {
                            WordDocumentProperties.Write(
                                document,
                                WordDocumentProperties.GetBuiltInName(index),
                                requestedValues.Value<string>(jsonName));
                        }
                    }
                    JObject after = WordDocumentProperties.ReadAll(document);
                    JArray mismatches = new JArray();
                    for (int index = 0; index < WordDocumentProperties.Count; index++)
                    {
                        string jsonName = WordDocumentProperties.GetJsonName(index);
                        if (!requested.Contains(jsonName))
                        {
                            continue;
                        }
                        string expected = requestedValues.Value<string>(jsonName) ?? string.Empty;
                        string actual = after.Value<string>(jsonName) ?? string.Empty;
                        if (!string.Equals(expected, actual, StringComparison.Ordinal))
                        {
                            mismatches.Add(new JObject
                            {
                                { "property", jsonName },
                                { "expected", expected },
                                { "actual", actual }
                            });
                        }
                    }

                    JObject data = new JObject();
                    data["before"] = before;
                    data["after"] = after;
                    data["requested"] = requestedValues;
                    data["mismatches"] = mismatches;
                    bool changed = !JToken.DeepEquals(before, after);
                    if (mismatches.Count > 0)
                    {
                        return WordWriteResult.Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "文档属性写入后回读不一致：" + mismatches[0]["property"].ToString() +
                            " 期望「" + mismatches[0]["expected"] + "」，实际「" +
                            mismatches[0]["actual"] + "」。请检查文档是否受保护。",
                            data,
                            changed);
                    }
                    return WordWriteResult.Ok(
                        "已更新文档属性：" + string.Join("/", requested.ToArray()) + "，全部回读一致。",
                        data,
                        changed);
                });
        }

        private OperationResult HandleFieldUpdateAll(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            bool includeToc = args.GetOptionalBool("includeToc", true);
            return RunWriteTool(
                application,
                WordToolNames.FieldUpdateAll,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.None("域与目录刷新", "Word:Fields");
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Fields fields = null;
                    Word.TablesOfContents contents = null;
                    JArray failures = new JArray();
                    JArray tocPreviews = new JArray();
                    int fieldCount = 0;
                    int fieldUpdateResult = 0;
                    int tocCount = 0;
                    try
                    {
                        fields = document.Fields;
                        fieldCount = fields == null ? 0 : fields.Count;
                        if (fields != null)
                        {
                            try
                            {
                                fieldUpdateResult = fields.Update();
                            }
                            catch (Exception ex)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INTERNAL_ERROR,
                                    "Fields.Update 失败：" +
                                    OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) +
                                    "。请检查文档保护/域锁定；本次未继续刷新区。",
                                    null,
                                    null);
                            }
                        }

                        if (includeToc)
                        {
                            contents = document.TablesOfContents;
                            tocCount = contents == null ? 0 : contents.Count;
                            for (int index = tocCount; index >= 1; index--)
                            {
                                Word.TableOfContents toc = null;
                                Word.Range range = null;
                                try
                                {
                                    toc = contents[index];
                                    toc.Update();
                                    toc.UpdatePageNumbers();
                                    range = toc.Range;
                                    tocPreviews.Add(new JObject
                                    {
                                        { "index", index },
                                        { "textPreview", WordComUtil.NormalizePreview(
                                            range == null ? string.Empty : range.Text, 160) }
                                    });
                                }
                                catch (Exception ex)
                                {
                                    failures.Add(new JObject
                                    {
                                        { "index", index },
                                        { "message", OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) }
                                    });
                                }
                                finally
                                {
                                    ComHelper.Release(range, toc);
                                }
                            }
                        }

                        try
                        {
                            document.Repaginate();
                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine("[word] Repaginate 失败：" + ex.Message);
                        }

                        JObject data = new JObject();
                        data["fieldCount"] = fieldCount;
                        data["fieldsUpdateResult"] = fieldUpdateResult;
                        data["tocCount"] = tocCount;
                        data["tocPreviews"] = tocPreviews;
                        data["failures"] = failures;
                        if (failures.Count > 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "目录刷新部分失败：" + failures[0]["message"].ToString() +
                                "。请检查目录是否受保护或被锁定。",
                                data,
                                null);
                        }
                        return WordWriteResult.Ok(
                            "已刷新 " + fieldCount + " 个域和 " + tocCount + " 个目录。",
                            data,
                            null);
                    }
                    finally
                    {
                        ComHelper.Release(contents, fields);
                    }
                });
        }

        private OperationResult HandleBeautify(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            JObject themeObject = args.GetOptionalObject("theme");
            JObject marginsObject = args.GetOptionalObject("margins");
            if (themeObject == null && marginsObject == null)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "word_beautify 至少需要 theme 或 margins。收到 keys=" + args.DescribeKeys() +
                    "；theme 支持 h1/h2/h3/body:{font,size,color,bold,italic,lineSpacing}，" +
                    "margins 支持 {top,bottom,left,right}（厘米，可用 marginsUnit=pt 改为磅）。");
            }

            Dictionary<string, WordThemeStyle> theme = WordThemeSettings.ParseTheme(themeObject);
            string marginsUnit = (args.GetOptionalString("marginsUnit", "cm") ?? "cm")
                .Trim().ToLowerInvariant();
            if (marginsUnit != "cm" && marginsUnit != "pt")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "marginsUnit 只支持 cm/pt，收到 " + marginsUnit + "。");
            }
            double? top = WordThemeSettings.ReadMargin(marginsObject, "top", marginsUnit);
            double? bottom = WordThemeSettings.ReadMargin(marginsObject, "bottom", marginsUnit);
            double? left = WordThemeSettings.ReadMargin(marginsObject, "left", marginsUnit);
            double? right = WordThemeSettings.ReadMargin(marginsObject, "right", marginsUnit);
            bool hasMargins = top.HasValue || bottom.HasValue || left.HasValue || right.HasValue;
            if (!hasMargins && theme.Count == 0)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "theme/margins 中没有可应用的字段。theme 示例：{\"h1\":{\"font\":\"微软雅黑\",\"size\":18}," +
                    "\"body\":{\"size\":12,\"lineSpacing\":1.5}}；margins 示例：{\"top\":2.54,\"bottom\":2.54}。");
            }

            return RunWriteTool(
                application,
                WordToolNames.Beautify,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.None("文档美化", "Word:Document");
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    JArray failures = new JArray();
                    JArray styleResults = new JArray();
                    JObject marginsBefore = null;
                    JObject marginsAfter = null;
                    bool changed = false;

                    if (hasMargins)
                    {
                        Word.PageSetup pageSetup = null;
                        try
                        {
                            pageSetup = document.PageSetup;
                            marginsBefore = WordComUtil.ReadPageSetup(pageSetup);
                            if (top.HasValue)
                            {
                                pageSetup.TopMargin = (float)top.Value;
                            }
                            if (bottom.HasValue)
                            {
                                pageSetup.BottomMargin = (float)bottom.Value;
                            }
                            if (left.HasValue)
                            {
                                pageSetup.LeftMargin = (float)left.Value;
                            }
                            if (right.HasValue)
                            {
                                pageSetup.RightMargin = (float)right.Value;
                            }
                            marginsAfter = WordComUtil.ReadPageSetup(pageSetup);
                            changed = !JToken.DeepEquals(marginsBefore, marginsAfter);
                            VerifyBeautifyMargins(failures, marginsAfter, top, bottom, left, right);
                        }
                        catch (Exception ex)
                        {
                            failures.Add(new JObject
                            {
                                { "target", "margins" },
                                { "message", OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) }
                            });
                        }
                        finally
                        {
                            ComHelper.Release(pageSetup);
                        }
                    }

                    string[] themeKeys = new string[] { "h1", "h2", "h3", "body" };
                    string[] styleNames = new string[] { "标题1", "标题2", "标题3", "正文" };
                    for (int index = 0; index < themeKeys.Length; index++)
                    {
                        WordThemeStyle style;
                        if (!theme.TryGetValue(themeKeys[index], out style) || style == null)
                        {
                            continue;
                        }
                        bool styleChanged;
                        JObject before;
                        JObject after;
                        bool applied = ApplyThemeStyle(
                            document, styleNames[index], style, failures, out styleChanged,
                            out before, out after);
                        if (applied)
                        {
                            changed = changed || styleChanged;
                            styleResults.Add(new JObject
                            {
                                { "key", themeKeys[index] },
                                { "style", styleNames[index] },
                                { "before", before },
                                { "after", after },
                                { "changed", styleChanged }
                            });
                        }
                    }

                    WordThemeStyle bodyStyle;
                    int lineSpacingParagraphs = 0;
                    if (theme.TryGetValue("body", out bodyStyle) && bodyStyle != null &&
                        bodyStyle.LineSpacing.HasValue)
                    {
                        bool lineSpacingChanged;
                        lineSpacingParagraphs = ApplyBodyLineSpacing(
                            document, bodyStyle.LineSpacing.Value, failures, out lineSpacingChanged);
                        changed = changed || lineSpacingChanged;
                    }

                    JObject data = new JObject();
                    data["marginsUnit"] = marginsUnit;
                    data["marginsBefore"] = marginsBefore == null ? new JObject() : marginsBefore;
                    data["marginsAfter"] = marginsAfter == null ? new JObject() : marginsAfter;
                    data["styles"] = styleResults;
                    data["lineSpacingParagraphs"] = lineSpacingParagraphs;
                    data["failures"] = failures;
                    if (failures.Count > 0)
                    {
                        return WordWriteResult.Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "文档美化部分失败：" + failures[0]["message"].ToString() +
                            "。成功项保持生效；请按 failures 调整 theme/margins。",
                            data,
                            changed);
                    }
                    return WordWriteResult.Ok(
                        "文档美化完成：页边距 " + (hasMargins ? "已设置" : "未改") +
                        "，样式主题 " + styleResults.Count + " 项，正文行距段 " +
                        lineSpacingParagraphs + " 个。",
                        data,
                        changed);
                });
        }

        private OperationResult HandleRebuildNumbering(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string scope = (args.GetOptionalString("scope", "all") ?? "all").Trim().ToLowerInvariant();
            bool selectionScope = scope == "selection" || scope == "选区";
            if (!selectionScope && scope != "all" && scope != "document" && scope != "全文")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "scope 只支持 all/selection，收到 " + scope + "。");
            }

            return RunWriteTool(
                application,
                WordToolNames.RebuildNumbering,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.None("自动编号重排", "Word:Numbering");
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Selection selection = null;
                    Word.Range scopeRange = null;
                    Word.Paragraphs scopeParagraphs = null;
                    Word.ListGalleries galleries = null;
                    Word.ListGallery gallery = null;
                    Word.ListTemplates templates = null;
                    Word.ListTemplate template = null;
                    try
                    {
                        if (selectionScope)
                        {
                            selection = application.Selection;
                            scopeRange = selection == null ? null : selection.Range;
                        }
                        else
                        {
                            scopeRange = document.Content;
                        }
                        if (scopeRange == null)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "scope=selection 时没有有效选区。",
                                null,
                                false);
                        }

                        scopeParagraphs = scopeRange.Paragraphs;
                        int paragraphCount = scopeParagraphs == null ? 0 : scopeParagraphs.Count;
                        List<int> targetIndexes = new List<int>();
                        List<string> beforeStrings = new List<string>();
                        List<int> beforeValues = new List<int>();
                        for (int index = 1; index <= paragraphCount; index++)
                        {
                            Word.Paragraph paragraph = null;
                            Word.Range range = null;
                            Word.ListFormat listFormat = null;
                            try
                            {
                                paragraph = scopeParagraphs[index];
                                range = paragraph.Range;
                                listFormat = range.ListFormat;
                                if (IsAutomaticNumbering(listFormat.ListType))
                                {
                                    targetIndexes.Add(index);
                                    beforeStrings.Add(listFormat.ListString ?? string.Empty);
                                    beforeValues.Add(listFormat.ListValue);
                                }
                            }
                            finally
                            {
                                ComHelper.Release(listFormat, range, paragraph);
                            }
                        }

                        JObject data = new JObject();
                        data["scope"] = selectionScope ? "selection" : "all";
                        data["detectedCount"] = targetIndexes.Count;
                        data["note"] = "只处理自动编号（simple/outline/mixed/listnum），不处理项目符号与手工编号；" +
                                       "为防索引漂移按文档倒序重排，验证按文档正序读取。";
                        if (targetIndexes.Count == 0)
                        {
                            data["appliedCount"] = 0;
                            return WordWriteResult.Ok(
                                "未检测到自动编号段落，文档未修改。若编号是手工输入的，本工具不会改动。",
                                data,
                                false);
                        }

                        galleries = application.ListGalleries;
                        gallery = galleries[Word.WdListGalleryType.wdNumberGallery];
                        templates = gallery.ListTemplates;
                        template = templates[1];

                        JArray failures = new JArray();
                        int appliedCount = 0;
                        for (int position = targetIndexes.Count - 1; position >= 0; position--)
                        {
                            int paragraphIndex = targetIndexes[position];
                            Word.Paragraph paragraph = null;
                            Word.Range range = null;
                            Word.ListFormat listFormat = null;
                            try
                            {
                                paragraph = scopeParagraphs[paragraphIndex];
                                range = paragraph.Range;
                                listFormat = range.ListFormat;
                                listFormat.RemoveNumbers(Word.WdNumberType.wdNumberParagraph);
                                object continuePrevious = position > 0;
                                object applyTo = Word.WdListApplyTo.wdListApplyToSelection;
                                object behavior = Word.WdDefaultListBehavior.wdWord10ListBehavior;
                                object level = 1;
                                listFormat.ApplyListTemplateWithLevel(
                                    template,
                                    ref continuePrevious,
                                    ref applyTo,
                                    ref behavior,
                                    ref level);
                                appliedCount++;
                            }
                            catch (Exception ex)
                            {
                                failures.Add(new JObject
                                {
                                    { "paragraphIndex", paragraphIndex },
                                    { "message", OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) }
                                });
                            }
                            finally
                            {
                                ComHelper.Release(listFormat, range, paragraph);
                            }
                        }

                        JArray numbering = new JArray();
                        bool changed = false;
                        for (int position = 0; position < targetIndexes.Count; position++)
                        {
                            int paragraphIndex = targetIndexes[position];
                            Word.Paragraph paragraph = null;
                            Word.Range range = null;
                            Word.ListFormat listFormat = null;
                            try
                            {
                                paragraph = scopeParagraphs[paragraphIndex];
                                range = paragraph.Range;
                                listFormat = range.ListFormat;
                                int expected = position + 1;
                                string listString = listFormat.ListString ?? string.Empty;
                                int listValue = listFormat.ListValue;
                                numbering.Add(new JObject
                                {
                                    { "paragraphIndex", paragraphIndex },
                                    { "expected", expected },
                                    { "listString", listString },
                                    { "listValue", listValue }
                                });
                                if (listValue != expected)
                                {
                                    failures.Add(new JObject
                                    {
                                        { "paragraphIndex", paragraphIndex },
                                        { "message", "重排后 ListValue=" + listValue +
                                                     "，期望 " + expected + "；ListString=" + listString }
                                    });
                                }
                                if (!string.Equals(beforeStrings[position], listString, StringComparison.Ordinal) ||
                                    beforeValues[position] != listValue)
                                {
                                    changed = true;
                                }
                            }
                            finally
                            {
                                ComHelper.Release(listFormat, range, paragraph);
                            }
                        }

                        data["appliedCount"] = appliedCount;
                        data["numbering"] = numbering;
                        data["failures"] = failures;
                        if (failures.Count > 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "自动编号重排部分失败：" + failures[0]["message"].ToString() +
                                "。含合并单元格/表格内编号时，建议先在小范围验证。",
                                data,
                                changed);
                        }
                        return WordWriteResult.Ok(
                            "已将 " + targetIndexes.Count + " 个自动编号段落重排为连续 1..n，回读 ListValue 一致。",
                            data,
                            changed);
                    }
                    finally
                    {
                        ComHelper.Release(template, templates, gallery, galleries, scopeParagraphs, scopeRange, selection);
                    }
                });
        }

        private static bool IsAutomaticNumbering(Word.WdListType listType)
        {
            return listType == Word.WdListType.wdListSimpleNumbering ||
                   listType == Word.WdListType.wdListOutlineNumbering ||
                   listType == Word.WdListType.wdListMixedNumbering ||
                   listType == Word.WdListType.wdListListNumOnly;
        }

        private static bool ApplyThemeStyle(
            Word.Document document,
            string requestedStyle,
            WordThemeStyle themeStyle,
            JArray failures,
            out bool changed,
            out JObject before,
            out JObject after)
        {
            changed = false;
            before = new JObject();
            after = new JObject();
            Word.Style style = null;
            Word.Font font = null;
            Word.ParagraphFormat format = null;
            Word.Styles styles = null;
            object styleObject = null;
            try
            {
                string resolved = WordStyleCatalog.ResolveStyleName(document, requestedStyle);
                styles = document.Styles;
                styleObject = styles[resolved];
                style = styleObject as Word.Style;
                if (style == null)
                {
                    failures.Add(new JObject
                    {
                        { "target", requestedStyle },
                        { "message", "解析到样式名「" + resolved + "」但无法取得 Style 对象。" }
                    });
                    return false;
                }

                before = CaptureStyleState(style);
                font = style.Font;
                format = style.ParagraphFormat;
                if (!string.IsNullOrWhiteSpace(themeStyle.FontName))
                {
                    font.Name = themeStyle.FontName;
                    try
                    {
                        font.NameFarEast = themeStyle.FontName;
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[word] 美化样式 NameFarEast 设置失败：" + ex.Message);
                    }
                }
                if (themeStyle.Size.HasValue)
                {
                    font.Size = (float)themeStyle.Size.Value;
                }
                if (themeStyle.Color.HasValue)
                {
                    font.Color = (Word.WdColor)themeStyle.Color.Value;
                }
                if (themeStyle.Bold.HasValue)
                {
                    font.Bold = themeStyle.Bold.Value ? -1 : 0;
                }
                if (themeStyle.Italic.HasValue)
                {
                    font.Italic = themeStyle.Italic.Value ? -1 : 0;
                }
                if (themeStyle.LineSpacing.HasValue)
                {
                    format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
                    format.LineSpacing = (float)(12d * themeStyle.LineSpacing.Value);
                }
                after = CaptureStyleState(style);
                changed = !JToken.DeepEquals(before, after);
                return true;
            }
            catch (Exception ex)
            {
                failures.Add(new JObject
                {
                    { "target", requestedStyle },
                    { "message", OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) }
                });
                return false;
            }
            finally
            {
                ComHelper.Release(format, font, style, styleObject, styles);
            }
        }

        private static JObject CaptureStyleState(Word.Style style)
        {
            JObject state = new JObject();
            Word.Font font = null;
            Word.ParagraphFormat format = null;
            try
            {
                font = style.Font;
                format = style.ParagraphFormat;
                state["styleName"] = style.NameLocal;
                state["fontName"] = font.Name;
                state["fontNameFarEast"] = font.NameFarEast;
                state["fontSize"] = (double)font.Size;
                state["color"] = (int)font.Color;
                state["colorHex"] = WordFormatUtil.ColorToHex((int)font.Color);
                state["bold"] = WordFormatUtil.TriStateToken(font.Bold);
                state["italic"] = WordFormatUtil.TriStateToken(font.Italic);
                state["lineSpacingRule"] = format.LineSpacingRule.ToString();
                state["lineSpacing"] = (double)format.LineSpacing;
                return state;
            }
            finally
            {
                ComHelper.Release(format, font);
            }
        }

        private static int ApplyBodyLineSpacing(
            Word.Document document,
            double multiplier,
            JArray failures,
            out bool changed)
        {
            changed = false;
            if (multiplier <= 0)
            {
                throw WordArguments.Invalid("theme.body.lineSpacing 必须大于 0。");
            }

            Word.Paragraphs paragraphs = null;
            JObject beforeSample = null;
            JObject afterSample = null;
            int applied = 0;
            try
            {
                paragraphs = document.Paragraphs;
                int count = paragraphs == null ? 0 : paragraphs.Count;
                if (count > 0)
                {
                    Word.Paragraph sample = null;
                    try
                    {
                        sample = paragraphs[1];
                        beforeSample = CaptureParagraphLineSpacing(sample);
                    }
                    finally
                    {
                        ComHelper.Release(sample);
                    }
                }
                for (int index = count; index >= 1; index--)
                {
                    Word.Paragraph paragraph = null;
                    Word.ParagraphFormat format = null;
                    try
                    {
                        paragraph = paragraphs[index];
                        format = paragraph.Format;
                        format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
                        format.LineSpacing = (float)(12d * multiplier);
                        applied++;
                    }
                    catch (Exception ex)
                    {
                        failures.Add(new JObject
                        {
                            { "target", "lineSpacing:paragraph" + index },
                            { "message", OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) }
                        });
                    }
                    finally
                    {
                        ComHelper.Release(format, paragraph);
                    }
                }
                if (count > 0)
                {
                    Word.Paragraph sample = null;
                    try
                    {
                        sample = paragraphs[1];
                        afterSample = CaptureParagraphLineSpacing(sample);
                    }
                    finally
                    {
                        ComHelper.Release(sample);
                    }
                }
                changed = !JToken.DeepEquals(beforeSample, afterSample);
                return applied;
            }
            finally
            {
                ComHelper.Release(paragraphs);
            }
        }

        private static JObject CaptureParagraphLineSpacing(Word.Paragraph paragraph)
        {
            JObject state = new JObject();
            Word.ParagraphFormat format = null;
            try
            {
                format = paragraph.Format;
                state["lineSpacingRule"] = format.LineSpacingRule.ToString();
                state["lineSpacing"] = (double)format.LineSpacing;
                return state;
            }
            finally
            {
                ComHelper.Release(format);
            }
        }

        private static void VerifyBeautifyMargins(
            JArray failures,
            JObject after,
            double? top,
            double? bottom,
            double? left,
            double? right)
        {
            VerifyBeautifyMargin(failures, after, "topMarginPt", top);
            VerifyBeautifyMargin(failures, after, "bottomMarginPt", bottom);
            VerifyBeautifyMargin(failures, after, "leftMarginPt", left);
            VerifyBeautifyMargin(failures, after, "rightMarginPt", right);
        }

        private static void VerifyBeautifyMargin(
            JArray failures,
            JObject after,
            string property,
            double? expected)
        {
            if (!expected.HasValue)
            {
                return;
            }
            double? actual = after.Value<double?>(property);
            if (!actual.HasValue || !WordFormatUtil.CloseEnough(actual.Value, expected.Value, 0.2))
            {
                failures.Add(new JObject
                {
                    { "target", property },
                    { "message", "页边距回读不一致：期望 " + expected.Value + " 磅，实际 " +
                                 (actual.HasValue ? actual.Value.ToString(CultureInfo.InvariantCulture) : "null") }
                });
            }
        }

        private static WordWriteTarget SelectTocTarget(
            Word.Application application,
            Word.Document document,
            string position)
        {
            if (position == "start")
            {
                return new WordWriteTarget(0, 0, "文档开头", "Word:TOC", false, true);
            }
            if (position == "end")
            {
                int end = Math.Max(0, WordComUtil.GetContentEnd(document) - 1);
                return new WordWriteTarget(end, end, "文档末尾", "Word:TOC", false, true);
            }
            Word.Selection selection = null;
            Word.Range range = null;
            try
            {
                selection = application.Selection;
                range = selection == null ? null : selection.Range;
                if (range == null)
                {
                    throw WordArguments.Invalid("position=cursor 时没有有效选区，无法插入目录。");
                }
                return new WordWriteTarget(range.Start, range.End, "当前光标位置", "Word:TOC", false, true);
            }
            finally
            {
                ComHelper.Release(range, selection);
            }
        }

        private static double? ResolveMargin(WordArguments args, string pointsName, string centimetersName)
        {
            bool hasPoints = args.Has(pointsName);
            bool hasCentimeters = args.Has(centimetersName);
            if (hasPoints && hasCentimeters)
            {
                throw WordArguments.Invalid(
                    "同一页边距不能同时用 " + pointsName + " 与 " + centimetersName + " 指定，请只保留一个。");
            }
            if (hasPoints)
            {
                double value = args.RequireDouble(pointsName);
                if (value < 0)
                {
                    throw WordArguments.Invalid(pointsName + " 不能为负数。");
                }
                return value;
            }
            if (hasCentimeters)
            {
                double value = args.RequireDouble(centimetersName);
                if (value < 0)
                {
                    throw WordArguments.Invalid(centimetersName + " 不能为负数。");
                }
                return value * WordFormatUtil.PointsPerCentimeter;
            }
            return null;
        }

        private static JArray VerifyPageSetup(
            JObject before,
            JObject after,
            double? top,
            double? bottom,
            double? left,
            double? right,
            Word.WdPaperSize? paperSize,
            Word.WdOrientation? orientation)
        {
            JArray verification = new JArray();
            AddPageSetupDouble(verification, "topMarginPt", top, before, after, "topMarginPt");
            AddPageSetupDouble(verification, "bottomMarginPt", bottom, before, after, "bottomMarginPt");
            AddPageSetupDouble(verification, "leftMarginPt", left, before, after, "leftMarginPt");
            AddPageSetupDouble(verification, "rightMarginPt", right, before, after, "rightMarginPt");
            if (paperSize.HasValue)
            {
                JToken actual = after["paperSizeCode"];
                bool matched = actual != null && actual.Value<int>() == (int)paperSize.Value;
                verification.Add(new JObject
                {
                    { "property", "paperSize" },
                    { "expected", (int)paperSize.Value },
                    { "actual", actual == null ? JValue.CreateNull() : actual.DeepClone() },
                    { "matched", matched }
                });
            }
            if (orientation.HasValue)
            {
                JToken actual = after["orientation"];
                bool matched = actual != null &&
                               string.Equals(actual.ToString(), orientation.Value.ToString(),
                                   StringComparison.OrdinalIgnoreCase);
                verification.Add(new JObject
                {
                    { "property", "orientation" },
                    { "expected", orientation.Value.ToString() },
                    { "actual", actual == null ? JValue.CreateNull() : actual.DeepClone() },
                    { "matched", matched }
                });
            }
            return verification;
        }

        private static void AddPageSetupDouble(
            JArray verification,
            string property,
            double? expected,
            JObject before,
            JObject after,
            string afterKey)
        {
            if (!expected.HasValue)
            {
                return;
            }
            double? actual = after.Value<double?>(afterKey);
            bool matched = actual.HasValue &&
                           WordFormatUtil.CloseEnough(actual.Value, expected.Value, 0.2);
            verification.Add(new JObject
            {
                { "property", property },
                { "expected", expected.Value },
                { "actual", actual.HasValue ? new JValue(actual.Value) : JValue.CreateNull() },
                { "matched", matched }
            });
        }

        private static Word.WdHeaderFooterIndex MapHeaderFooterIndex(string headerType)
        {
            switch ((headerType ?? "primary").Trim().ToLowerInvariant())
            {
                case "first":
                case "firstpage":
                case "首页":
                    return Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage;
                case "even":
                case "evenpages":
                case "偶数页":
                    return Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages;
                default:
                    return Word.WdHeaderFooterIndex.wdHeaderFooterPrimary;
            }
        }

        private static void TrySetLinkToPrevious(
            Word.HeaderFooter headerFooter,
            bool enabled,
            JArray warnings)
        {
            if (headerFooter == null)
            {
                return;
            }
            try
            {
                headerFooter.LinkToPrevious = enabled;
            }
            catch (Exception ex)
            {
                warnings.Add("设置 LinkToPrevious 失败：" + ex.Message);
            }
        }

        private static int CountFields(Word.Range range)
        {
            if (range == null)
            {
                return 0;
            }
            Word.Fields fields = null;
            try
            {
                fields = range.Fields;
                return fields == null ? 0 : fields.Count;
            }
            catch
            {
                return 0;
            }
            finally
            {
                ComHelper.Release(fields);
            }
        }

        private static void AppendPageNumberFields(
            Word.Document document,
            Word.HeaderFooter headerFooter,
            string format)
        {
            if (headerFooter == null || string.IsNullOrWhiteSpace(format) ||
                string.Equals(format, "none", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            string normalized = format.Trim();
            if (string.Equals(normalized, "page", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "x", StringComparison.OrdinalIgnoreCase))
            {
                AppendHeaderFooterField(document, headerFooter, Word.WdFieldType.wdFieldPage);
                return;
            }
            if (string.Equals(normalized, "x/y", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "xOfY", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "pageOfPages", StringComparison.OrdinalIgnoreCase))
            {
                AppendHeaderFooterField(document, headerFooter, Word.WdFieldType.wdFieldPage);
                AppendHeaderFooterText(document, headerFooter, " / ");
                AppendHeaderFooterField(document, headerFooter, Word.WdFieldType.wdFieldNumPages);
                return;
            }
            if (string.Equals(normalized, "第X页共Y页", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "chinese", StringComparison.OrdinalIgnoreCase))
            {
                AppendHeaderFooterText(document, headerFooter, "第 ");
                AppendHeaderFooterField(document, headerFooter, Word.WdFieldType.wdFieldPage);
                AppendHeaderFooterText(document, headerFooter, " 页 共 ");
                AppendHeaderFooterField(document, headerFooter, Word.WdFieldType.wdFieldNumPages);
                AppendHeaderFooterText(document, headerFooter, " 页");
                return;
            }
            throw WordArguments.Invalid("不支持的 pageNumbers=" + format +
                                        "；支持 none/page/x/xOfY/pageOfPages/第X页共Y页");
        }

        private static void AppendHeaderFooterText(
            Word.Document document,
            Word.HeaderFooter headerFooter,
            string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }
            Word.Range headerFooterRange = null;
            Word.Range range = null;
            try
            {
                // 页眉页脚是独立 Story，不能用 document.Range 在同一位置创建范围，
                // 否则插入会落到正文 Story，Fields.Count 也不会增加。
                headerFooterRange = headerFooter.Range;
                int position = Math.Max(0, headerFooterRange.End - 1);
                range = headerFooterRange.Duplicate;
                range.SetRange(position, position);
                range.InsertAfter(text);
            }
            finally
            {
                ComHelper.Release(range, headerFooterRange);
            }
        }

        private static void AppendHeaderFooterField(
            Word.Document document,
            Word.HeaderFooter headerFooter,
            Word.WdFieldType fieldType)
        {
            Word.Range headerFooterRange = null;
            Word.Range range = null;
            Word.Fields fields = null;
            Word.Field field = null;
            try
            {
                headerFooterRange = headerFooter.Range;
                int position = Math.Max(0, headerFooterRange.End - 1);
                range = headerFooterRange.Duplicate;
                range.SetRange(position, position);
                fields = range.Fields;
                object type = fieldType;
                object text = Type.Missing;
                object preserve = false;
                field = fields.Add(range, ref type, ref text, ref preserve);
                if (field != null)
                {
                    field.Update();
                }
            }
            finally
            {
                ComHelper.Release(field, fields, range, headerFooterRange);
            }
        }

        private static string NormalizeExportPath(string requestedPath, string format)
        {
            string path;
            string pathError;
            if (!Office.Config.WorkspacePathResolver.TryResolve(requestedPath, out path, out pathError))
            {
                throw WordArguments.Invalid(pathError);
            }
            if (string.IsNullOrEmpty(Path.GetExtension(path)))
            {
                path = path + "." + format;
            }
            return path;
        }
    }
}
