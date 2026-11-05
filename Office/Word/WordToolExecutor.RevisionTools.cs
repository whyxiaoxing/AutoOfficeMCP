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
        private void RegisterRevisionHandlers(
            Dictionary<string, Func<Word.Application, JObject, OperationResult>> handlers)
        {
            handlers[WordToolNames.TrackChangesToggle] = HandleTrackChangesToggle;
            handlers[WordToolNames.RevisionsList] = HandleRevisionsList;
            handlers[WordToolNames.RevisionAccept] = HandleRevisionAccept;
            handlers[WordToolNames.RevisionReject] = HandleRevisionReject;
            handlers[WordToolNames.CommentAdd] = HandleCommentAdd;
            handlers[WordToolNames.CommentsRead] = HandleCommentsRead;
            handlers[WordToolNames.CompareDocuments] = HandleCompareDocuments;
        }

        private OperationResult HandleTrackChangesToggle(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            if (!args.Has("enabled"))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "缺少 enabled 布尔参数（true=开启修订，false=关闭修订）。");
            }
            bool enabled = args.GetOptionalBool("enabled", false);

            return RunWriteTool(
                application,
                WordToolNames.TrackChangesToggle,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.None("修订模式", "Word:TrackChanges");
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    bool before = document.TrackRevisions;
                    document.TrackRevisions = enabled;
                    bool after = document.TrackRevisions;
                    bool changed = before != after;
                    JObject data = new JObject();
                    data["before"] = before;
                    data["requested"] = enabled;
                    data["after"] = after;
                    if (after != enabled)
                    {
                        return WordWriteResult.Fail(
                            OfficeErrorCode.INTERNAL_ERROR,
                            "修订模式设置后回读为 " + after + "，与请求 " + enabled +
                            " 不一致。请检查文档保护/策略限制。",
                            data,
                            changed);
                    }
                    return WordWriteResult.Ok(
                        "修订模式已" + (after ? "开启" : "关闭") + "（原值 " + before + "）。",
                        data,
                        changed);
                });
        }

        private OperationResult HandleRevisionsList(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int maxCount = args.GetOptionalInt("maxCount", 100);
            if (maxCount < 1)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "maxCount 必须大于等于 1。");
            }

            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    Word.Revisions revisions = null;
                    JArray items = new JArray();
                    try
                    {
                        revisions = document.Revisions;
                        int total = revisions == null ? 0 : revisions.Count;
                        int returned = Math.Min(total, maxCount);
                        for (int index = 1; index <= returned; index++)
                        {
                            Word.Revision revision = null;
                            Word.Range range = null;
                            try
                            {
                                revision = revisions[index];
                                range = revision.Range;
                                items.Add(new JObject
                                {
                                    { "index", index },
                                    { "type", revision.Type.ToString() },
                                    { "typeCode", (int)revision.Type },
                                    { "author", revision.Author ?? string.Empty },
                                    { "date", revision.Date.ToString("o", CultureInfo.InvariantCulture) },
                                    { "paragraphIndex", WordComUtil.FindParagraphIndexAt(
                                        document, range == null ? 0 : range.Start) },
                                    { "rangeStart", range == null ? 0 : range.Start },
                                    { "rangeEnd", range == null ? 0 : range.End },
                                    { "textPreview", WordComUtil.NormalizePreview(
                                        range == null ? string.Empty : range.Text, 120) },
                                    { "formatDescription", revision.FormatDescription ?? string.Empty }
                                });
                            }
                            finally
                            {
                                ComHelper.Release(range, revision);
                            }
                        }
                        JObject data = new JObject();
                        data["total"] = total;
                        data["returned"] = returned;
                        data["truncated"] = total > returned;
                        data["trackRevisions"] = document.TrackRevisions;
                        data["items"] = items;
                        return ReadSucceed(
                            "读取 " + returned + "/" + total + " 个修订。", data, "Word:Revisions");
                    }
                    finally
                    {
                        ComHelper.Release(revisions);
                    }
                });
        }

        private OperationResult HandleRevisionAccept(Word.Application application, JObject arguments)
        {
            return ChangeRevision(application, arguments, true);
        }

        private OperationResult HandleRevisionReject(Word.Application application, JObject arguments)
        {
            return ChangeRevision(application, arguments, false);
        }

        private OperationResult ChangeRevision(
            Word.Application application,
            JObject arguments,
            bool accept)
        {
            WordArguments args = new WordArguments(arguments);
            int revisionIndex = args.RequireInt("revisionIndex", 1, int.MaxValue);
            string toolName = accept
                ? WordToolNames.RevisionAccept
                : WordToolNames.RevisionReject;

            return RunWriteTool(
                application,
                toolName,
                false,
                delegate(Word.Document document)
                {
                    return WordWriteTargets.None("修订", "Word:Revision:" + revisionIndex);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Revisions revisions = null;
                    Word.Revision revision = null;
                    try
                    {
                        revisions = document.Revisions;
                        int beforeCount = revisions == null ? 0 : revisions.Count;
                        if (revisionIndex > beforeCount)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "修订索引 " + revisionIndex + " 不存在（当前共 " + beforeCount +
                                " 个修订，1 基）。请先调用 word_revisions_list。",
                                null,
                                false);
                        }
                        revision = revisions[revisionIndex];
                        string type = revision.Type.ToString();
                        string author = revision.Author ?? string.Empty;
                        if (accept)
                        {
                            revision.Accept();
                        }
                        else
                        {
                            revision.Reject();
                        }
                        ComHelper.Release(revision);
                        revision = null;
                        ComHelper.Release(revisions);
                        revisions = document.Revisions;
                        int afterCount = revisions == null ? 0 : revisions.Count;
                        bool changed = afterCount != beforeCount;
                        JObject data = new JObject();
                        data["revisionIndex"] = revisionIndex;
                        data["action"] = accept ? "accept" : "reject";
                        data["revisionType"] = type;
                        data["author"] = author;
                        data["beforeCount"] = beforeCount;
                        data["afterCount"] = afterCount;
                        if (!changed)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "修订 " + revisionIndex + " 已执行 " + (accept ? "接受" : "拒绝") +
                                "，但修订数量未变化（前 " + beforeCount + "，后 " + afterCount +
                                "）。请确认该修订仍存在且文档未受保护。",
                                data,
                                false);
                        }
                        return WordWriteResult.Ok(
                            "已" + (accept ? "接受" : "拒绝") + "修订 " + revisionIndex +
                            "（" + type + "），修订数 " + beforeCount + "→" + afterCount + "。",
                            data,
                            true);
                    }
                    finally
                    {
                        ComHelper.Release(revision, revisions);
                    }
                });
        }

        private OperationResult HandleCommentAdd(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string text = args.RequireText("text");
            int? paragraphIndex = args.GetOptionalInt("paragraphIndex");
            string author = args.Has("author") ? args.GetOptionalString("author", null) : null;
            string initials = args.Has("initials") ? args.GetOptionalString("initials", null) : null;

            return RunWriteTool(
                application,
                WordToolNames.CommentAdd,
                false,
                delegate(Word.Document document)
                {
                    if (paragraphIndex.HasValue)
                    {
                        WordComUtil.ValidateParagraphIndex(document, paragraphIndex.Value);
                        return WordWriteTargets.Paragraph(
                            document,
                            paragraphIndex.Value,
                            "Word:Paragraph:" + paragraphIndex.Value,
                            true);
                    }
                    return WordWriteTargets.Selection(
                        application, "当前选区", "Word:Selection", false, true);
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Selection selection = null;
                    Word.Paragraph paragraph = null;
                    Word.Range anchorRange = null;
                    Word.Comments comments = null;
                    Word.Comment comment = null;
                    try
                    {
                        if (paragraphIndex.HasValue)
                        {
                            paragraph = WordComUtil.GetParagraph(document, paragraphIndex.Value);
                            anchorRange = paragraph.Range;
                        }
                        else
                        {
                            selection = application.Selection;
                            anchorRange = selection == null ? null : selection.Range;
                        }
                        if (anchorRange == null)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "批注锚点无效：请先选中文本，或传 paragraphIndex 指定段落。",
                                null,
                                false);
                        }
                        string anchorText = anchorRange.Text ?? string.Empty;
                        if (WordComUtil.CleanParagraphText(anchorText).Length == 0)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "批注锚点为空。请选中非空文本，或改用包含文本的 paragraphIndex。",
                                null,
                                false);
                        }

                        comments = document.Comments;
                        int beforeCount = comments == null ? 0 : comments.Count;
                        object textObject = text;
                        comment = comments.Add(anchorRange, ref textObject);
                        if (comment == null)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INTERNAL_ERROR,
                                "Comments.Add 返回 null，批注未创建。",
                                null,
                                false);
                        }
                        if (!string.IsNullOrWhiteSpace(author))
                        {
                            comment.Author = author;
                        }
                        if (!string.IsNullOrWhiteSpace(initials))
                        {
                            comment.Initial = initials;
                        }
                        int afterCount = comments.Count;
                        JObject data = new JObject();
                        data["commentIndex"] = comment.Index;
                        data["author"] = comment.Author ?? string.Empty;
                        data["initials"] = comment.Initial ?? string.Empty;
                        data["text"] = text;
                        data["paragraphIndex"] = WordComUtil.FindParagraphIndexAt(
                            document, anchorRange.Start);
                        data["anchorStart"] = anchorRange.Start;
                        data["anchorEnd"] = anchorRange.End;
                        data["anchorPreview"] = WordComUtil.NormalizePreview(anchorText, 120);
                        data["beforeCount"] = beforeCount;
                        data["afterCount"] = afterCount;
                        if (afterCount <= beforeCount)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INTERNAL_ERROR,
                                "批注创建后评论数量未增加（前 " + beforeCount + "，后 " + afterCount + "）。",
                                data,
                                false);
                        }
                        WordWriteResult result = WordWriteResult.Ok(
                            "已在第 " + data["paragraphIndex"] + " 段添加批注 #" + comment.Index + "。",
                            data,
                            true);
                        result.AddTargetRef("Word:Comment:" + comment.Index);
                        return result;
                    }
                    finally
                    {
                        ComHelper.Release(comment, comments, anchorRange, paragraph, selection);
                    }
                });
        }

        private OperationResult HandleCommentsRead(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            int maxCount = args.GetOptionalInt("maxCount", 100);
            if (maxCount < 1)
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "maxCount 必须大于等于 1。");
            }

            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    Word.Comments comments = null;
                    JArray items = new JArray();
                    try
                    {
                        comments = document.Comments;
                        int total = comments == null ? 0 : comments.Count;
                        int returned = Math.Min(total, maxCount);
                        for (int index = 1; index <= returned; index++)
                        {
                            Word.Comment comment = null;
                            Word.Range range = null;
                            Word.Range scope = null;
                            try
                            {
                                comment = comments[index];
                                range = comment.Range;
                                scope = comment.Scope;
                                items.Add(new JObject
                                {
                                    { "index", comment.Index },
                                    { "author", comment.Author ?? string.Empty },
                                    { "initials", comment.Initial ?? string.Empty },
                                    { "date", comment.Date.ToString("o", CultureInfo.InvariantCulture) },
                                    { "done", comment.Done },
                                    { "text", WordComUtil.CleanParagraphText(
                                        range == null ? string.Empty : range.Text) },
                                    { "anchorParagraphIndex", WordComUtil.FindParagraphIndexAt(
                                        document, scope == null ? 0 : scope.Start) },
                                    { "anchorStart", scope == null ? 0 : scope.Start },
                                    { "anchorEnd", scope == null ? 0 : scope.End },
                                    { "anchorPreview", WordComUtil.NormalizePreview(
                                        scope == null ? string.Empty : scope.Text, 120) }
                                });
                            }
                            finally
                            {
                                ComHelper.Release(scope, range, comment);
                            }
                        }
                        JObject data = new JObject();
                        data["total"] = total;
                        data["returned"] = returned;
                        data["truncated"] = total > returned;
                        data["items"] = items;
                        return ReadSucceed(
                            "读取 " + returned + "/" + total + " 个批注。", data, "Word:Comments");
                    }
                    finally
                    {
                        ComHelper.Release(comments);
                    }
                });
        }

        private OperationResult HandleCompareDocuments(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string originalPath = args.RequireString("originalPath");
            string revisedPath = args.RequireString("revisedPath");
            string outputPath = args.RequireString("outputPath");
            bool confirm = args.GetOptionalBool("confirm", false);
            string author = args.Has("author")
                ? args.GetOptionalString("author", "AutoOfficeMCP")
                : "AutoOfficeMCP";

            if (!File.Exists(originalPath))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.FILE_NOT_FOUND,
                    "原始文档不存在：" + originalPath);
            }
            if (!File.Exists(revisedPath))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.FILE_NOT_FOUND,
                    "修订文档不存在：" + revisedPath);
            }

            string output = NormalizeCompareOutputPath(originalPath, outputPath);
            string directory = Path.GetDirectoryName(output);
            if (string.IsNullOrWhiteSpace(directory))
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "对比稿输出路径缺少有效目录：" + output);
            }
            string directoryError;
            if (!WorkspacePathResolver.TryEnsureOutputDirectory(directory, out directoryError))
            {
                return OperationResult.Failed(OfficeErrorCode.READ_ONLY, directoryError);
            }
            bool existed = File.Exists(output);
            if (existed && !confirm)
            {
                return OperationResult.Cancelled(
                    "对比稿已存在：" + output + "。如需覆盖，请显式传 confirm=true；本次未写入文件。");
            }

            Word.Documents documents = null;
            Word.Document original = null;
            Word.Document revised = null;
            Word.Document compared = null;
            Word.Document previous = null;
            bool screenCaptured = false;
            bool screenOriginal = false;
            bool alertsCaptured = false;
            Word.WdAlertLevel alertsOriginal = Word.WdAlertLevel.wdAlertsAll;
            try
            {
                previous = application.ActiveDocument;
            }
            catch
            {
                previous = null;
            }

            try
            {
                try
                {
                    screenOriginal = application.ScreenUpdating;
                    screenCaptured = true;
                    application.ScreenUpdating = false;
                }
                catch
                {
                }
                try
                {
                    alertsOriginal = application.DisplayAlerts;
                    alertsCaptured = true;
                    application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                }
                catch
                {
                }

                documents = application.Documents;
                object missing = Type.Missing;
                object readOnly = true;
                object addToRecentFiles = false;
                object visible = false;

                object originalFile = originalPath;
                original = documents.Open(
                    ref originalFile, ref missing, ref readOnly, ref addToRecentFiles,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref visible, ref missing, ref missing,
                    ref missing, ref missing);
                object revisedFile = revisedPath;
                revised = documents.Open(
                    ref revisedFile, ref missing, ref readOnly, ref addToRecentFiles,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref visible, ref missing, ref missing,
                    ref missing, ref missing);

                compared = application.CompareDocuments(
                    original,
                    revised,
                    Word.WdCompareDestination.wdCompareDestinationNew,
                    Word.WdGranularity.wdGranularityWordLevel,
                    true,
                    true,
                    true,
                    true,
                    true,
                    true,
                    true,
                    true,
                    true,
                    true,
                    author,
                    true);

                int revisions = 0;
                Word.Revisions revisionCollection = null;
                try
                {
                    revisionCollection = compared.Revisions;
                    revisions = revisionCollection == null ? 0 : revisionCollection.Count;
                }
                finally
                {
                    ComHelper.Release(revisionCollection);
                }

                SaveComparedDocument(compared, output);
                if (!File.Exists(output))
                {
                    return OperationResult.Failed(
                        OfficeErrorCode.INTERNAL_ERROR,
                        "CompareDocuments 已返回，但对比稿文件不存在：" + output);
                }
                FileInfo info = new FileInfo(output);
                JObject data = new JObject();
                data["originalPath"] = Path.GetFullPath(originalPath);
                data["revisedPath"] = Path.GetFullPath(revisedPath);
                data["outputPath"] = output;
                data["revisionCount"] = revisions;
                data["bytes"] = info.Length;
                data["overwroteExisting"] = existed;
                data["activeDocumentModified"] = false;
                data["note"] = "对比稿已保存并关闭；原始/修订文档以只读方式临时打开，当前活动文档未修改。";

                Observation observation = Observation.Create(
                    "write", "已生成 Word 修订对比稿并保存到 " + output, true);
                observation.AddTargetRef("File:" + output);
                return OperationResult.Succeed(
                    "已生成修订对比稿：" + output + "（" + revisions + " 个修订，" +
                    info.Length + " 字节）。", data, observation);
            }
            catch (Exception ex)
            {
                return OfficeErrorMapper.ToOperationResult(ex, OfficeApp.Word);
            }
            finally
            {
                CloseDocumentQuietly(compared, false);
                CloseDocumentQuietly(revised, false);
                CloseDocumentQuietly(original, false);
                ComHelper.Release(compared, revised, original, documents);
                if (previous != null)
                {
                    try
                    {
                        previous.Activate();
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[word] 恢复活动文档失败：" + ex.Message);
                    }
                }
                ComHelper.Release(previous);
                if (alertsCaptured)
                {
                    try
                    {
                        application.DisplayAlerts = alertsOriginal;
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[word] 恢复 DisplayAlerts 失败：" + ex.Message);
                    }
                }
                if (screenCaptured)
                {
                    try
                    {
                        application.ScreenUpdating = screenOriginal;
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[word] 恢复 ScreenUpdating 失败：" + ex.Message);
                    }
                }
            }
        }

        private static void SaveComparedDocument(Word.Document document, string outputPath)
        {
            object fileName = outputPath;
            object fileFormat = Word.WdSaveFormat.wdFormatXMLDocument;
            object missing = Type.Missing;
            document.SaveAs2(
                ref fileName, ref fileFormat, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing);
        }

        private static void CloseDocumentQuietly(Word.Document document, bool saveChanges)
        {
            if (document == null)
            {
                return;
            }
            try
            {
                object save = saveChanges
                    ? Word.WdSaveOptions.wdSaveChanges
                    : Word.WdSaveOptions.wdDoNotSaveChanges;
                object missing = Type.Missing;
                document.Close(ref save, ref missing, ref missing);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[word] 关闭对比文档失败：" + ex.Message);
            }
        }

        private static string NormalizeCompareOutputPath(string originalPath, string outputPath)
        {
            string path = outputPath.Trim();
            if (!Path.IsPathRooted(path))
            {
                string originalDirectory = Path.GetDirectoryName(Path.GetFullPath(originalPath));
                path = Path.Combine(originalDirectory ?? string.Empty, path);
            }
            path = Path.GetFullPath(path);
            if (string.IsNullOrEmpty(Path.GetExtension(path)))
            {
                path = path + ".docx";
            }
            return path;
        }
    }
}
