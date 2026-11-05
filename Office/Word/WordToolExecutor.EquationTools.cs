using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    public sealed partial class WordToolExecutor
    {
        private void RegisterEquationHandlers(
            Dictionary<string, Func<Word.Application, JObject, OperationResult>> handlers)
        {
            handlers[WordToolNames.InsertEquation] = HandleInsertEquation;
            handlers[WordToolNames.EquationsList] = HandleEquationsList;
        }

        private OperationResult HandleInsertEquation(Word.Application application, JObject arguments)
        {
            WordArguments args = new WordArguments(arguments);
            string equation = args.RequireText("equation");
            int? paragraphIndex = args.GetOptionalInt("paragraphIndex");
            string position = (args.GetOptionalString("position", "cursor") ?? "cursor")
                .Trim().ToLowerInvariant();
            if (position != "cursor" && position != "start" && position != "end")
            {
                return OperationResult.Failed(
                    OfficeErrorCode.INVALID_PARAMS,
                    "position 只支持 cursor/start/end，收到 " + position + "。");
            }

            return RunWriteTool(
                application,
                WordToolNames.InsertEquation,
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
                    return WordWriteTargets.None("公式插入点", "Word:Equation");
                },
                delegate(Word.Document document, WordWriteTarget target)
                {
                    Word.Selection selection = null;
                    Word.Paragraph paragraph = null;
                    Word.Range paragraphRange = null;
                    Word.Range insertionRange = null;
                    Word.Range equationRange = null;
                    Word.OMaths omaths = null;
                    Word.OMaths equationCollection = null;
                    Word.OMath math = null;
                    Word.Range mathRange = null;
                    try
                    {
                        int insertPosition;
                        if (paragraphIndex.HasValue)
                        {
                            paragraph = WordComUtil.GetParagraph(document, paragraphIndex.Value);
                            paragraphRange = paragraph.Range;
                            insertPosition = position == "start"
                                ? paragraphRange.Start
                                : Math.Max(paragraphRange.Start, paragraphRange.End - 1);
                        }
                        else if (position == "start")
                        {
                            insertPosition = 0;
                        }
                        else if (position == "end")
                        {
                            insertPosition = Math.Max(0, WordComUtil.GetContentEnd(document) - 1);
                        }
                        else
                        {
                            selection = application.Selection;
                            insertionRange = selection == null ? null : selection.Range;
                            if (insertionRange == null)
                            {
                                return WordWriteResult.Fail(
                                    OfficeErrorCode.INVALID_PARAMS,
                                    "position=cursor 时没有有效选区，无法插入公式。",
                                    null,
                                    false);
                            }
                            insertPosition = insertionRange.Start;
                        }

                        if (insertionRange == null)
                        {
                            insertionRange = WordComUtil.CreateRange(document, insertPosition, insertPosition);
                        }
                        omaths = document.OMaths;
                        int beforeCount = omaths == null ? 0 : omaths.Count;
                        equationRange = omaths.Add(insertionRange);
                        if (equationRange == null)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INTERNAL_ERROR,
                                "OMaths.Add 返回 null，公式未创建。",
                                null,
                                false);
                        }

                        bool usedFallback = false;
                        string fallbackWarning = string.Empty;
                        int currentCount = CountEquations(document);
                        equationCollection = equationRange.OMaths;
                        if (currentCount <= beforeCount ||
                            equationCollection == null || equationCollection.Count == 0)
                        {
                            // Word 2016 在隐藏窗口/非零插入点（尤其目标段落含 TOC 域）时
                            // OMaths.Add 可能静默不生成 OMath。可靠降级：在独立临时文档的
                            // 0 位置创建并 BuildUp，复制公式，关闭临时文档，再粘贴到目标位置。
                            ComHelper.Release(equationCollection);
                            equationCollection = null;
                            ComHelper.Release(equationRange);
                            equationRange = null;
                            Word.Document scratch = null;
                            Word.Range scratchRange = null;
                            Word.OMaths scratchOmaths = null;
                            Word.Range scratchEquation = null;
                            Word.OMaths scratchEquationCollection = null;
                            Word.OMath scratchMath = null;
                            Word.Range fallbackTarget = null;
                            try
                            {
                                object missing = Type.Missing;
                                scratch = WordTemplateFactory.CreateDocument(application);
                                scratchRange = WordComUtil.CreateRange(scratch, 0, 0);
                                scratchOmaths = scratch.OMaths;
                                scratchEquation = scratchOmaths.Add(scratchRange);
                                scratchEquationCollection = scratchEquation == null
                                    ? null
                                    : scratchEquation.OMaths;
                                if (scratchEquationCollection == null || scratchEquationCollection.Count == 0)
                                {
                                    return WordWriteResult.Fail(
                                        OfficeErrorCode.INVALID_PARAMS,
                                        "OMaths.Add 未生成公式（原位置与独立临时文档降级均失败）。" +
                                        "请改用 UnicodeMath 线性格式 x^2+y^2=z^2 并确认 Word 数学功能可用。",
                                        null,
                                        false);
                                }
                                scratchMath = scratchEquationCollection[1];
                                scratchMath.Range.Text = equation;
                                scratchMath.BuildUp();
                                mathRange = scratchMath.Range;
                                mathRange.Copy();
                                object save = Word.WdSaveOptions.wdDoNotSaveChanges;
                                scratch.Close(ref save, ref missing, ref missing);
                                scratch = null;
                                document.Activate();
                                fallbackTarget = WordComUtil.CreateRange(
                                    document, insertPosition, insertPosition);
                                fallbackTarget.Paste();
                                usedFallback = true;
                                fallbackWarning =
                                    "Word 2016 在隐藏窗口/非零插入点（尤其目标含 TOC 域）时 OMaths.Add 未生成 OMath；" +
                                    "已降级为“独立临时文档创建→BuildUp→复制→目标位置粘贴”。";
                            }
                            finally
                            {
                                if (scratch != null)
                                {
                                    try
                                    {
                                        object save = Word.WdSaveOptions.wdDoNotSaveChanges;
                                        object missing = Type.Missing;
                                        scratch.Close(ref save, ref missing, ref missing);
                                    }
                                    catch (Exception closeError)
                                    {
                                        Console.Error.WriteLine("[word] 关闭公式临时文档失败：" +
                                                                closeError.Message);
                                    }
                                }
                                ComHelper.Release(fallbackTarget, scratchMath,
                                    scratchEquationCollection, scratchEquation, scratchOmaths,
                                    scratchRange, scratch);
                                try
                                {
                                    document.Activate();
                                }
                                catch (Exception activateError)
                                {
                                    Console.Error.WriteLine("[word] 恢复公式目标文档失败：" +
                                                            activateError.Message);
                                }
                            }
                        }
                        else
                        {
                            math = equationCollection[1];
                            math.Range.Text = equation;
                            math.BuildUp();
                        }

                        int afterCount = CountEquations(document);
                        JObject data = new JObject();
                        data["equation"] = equation;
                        data["omathsCountBefore"] = beforeCount;
                        data["omathsCountAfter"] = afterCount;
                        data["usedFallback"] = usedFallback;
                        data["formatHint"] =
                            "使用 Word UnicodeMath 线性格式（如 x^2+y^2=z^2）。COM 没有可靠的 LaTeX 解析入口；" +
                            "\\alpha、\\frac 等 LaTeX 命令不保证被 Word 识别。";
                        if (equation.IndexOf("\\", StringComparison.Ordinal) >= 0)
                        {
                            data["warning"] = "输入含反斜杠/LaTeX 命令；若 BuildUp 后显示异常，请改用 UnicodeMath 线性格式。";
                        }
                        if (!string.IsNullOrWhiteSpace(fallbackWarning))
                        {
                            data["fallbackWarning"] = fallbackWarning;
                        }
                        data["paragraphIndex"] = WordComUtil.FindParagraphIndexAt(
                            document, insertPosition);

                        if (afterCount <= beforeCount)
                        {
                            return WordWriteResult.Fail(
                                OfficeErrorCode.INVALID_PARAMS,
                                "OMaths.Add/BuildUp 后公式数量未增加（前 " + beforeCount + "，后 " +
                                afterCount + "）。请改用 UnicodeMath 线性格式，例如 x^2+y^2=z^2。",
                                data,
                                false);
                        }

                        int equationIndex = FindEquationIndex(document, insertPosition);
                        if (equationIndex <= 0)
                        {
                            equationIndex = FindEquationIndexByText(document, equation);
                        }
                        if (equationIndex <= 0)
                        {
                            equationIndex = FindNearestEquationIndex(document, insertPosition);
                        }
                        if (equationIndex <= 0 && afterCount == 1)
                        {
                            equationIndex = 1;
                        }
                        data["equationIndex"] = equationIndex;
                        data["omathsCount"] = afterCount;
                        WordWriteResult result = WordWriteResult.Ok(
                            "已插入公式 " + equation + "（OMaths.Count=" + afterCount +
                            "，索引 " + equationIndex + "，已调用 BuildUp）。",
                            data,
                            true);
                        result.AddTargetRef("Word:Equation:" + equationIndex);
                        return result;
                    }
                    catch (Exception ex)
                    {
                        throw new OfficeComException(
                            OfficeErrorCode.INVALID_PARAMS,
                            "公式插入/构建失败：" + OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) +
                            "。COM 通道优先支持 Word UnicodeMath 线性格式；LaTeX 命令（\\alpha、\\frac 等）" +
                            "不保证可解析，请改用 x^2+y^2=z^2 这类线性格式后重试。",
                            ex);
                    }
                    finally
                    {
                        ComHelper.Release(mathRange, math, equationCollection, omaths,
                            equationRange, insertionRange, paragraphRange, paragraph, selection);
                    }
                });
        }

        private OperationResult HandleEquationsList(Word.Application application, JObject arguments)
        {
            return RunReadTool(
                application,
                delegate(Word.Document document)
                {
                    Word.OMaths omaths = null;
                    JArray items = new JArray();
                    try
                    {
                        omaths = document.OMaths;
                        int count = omaths == null ? 0 : omaths.Count;
                        for (int index = 1; index <= count; index++)
                        {
                            Word.OMath math = null;
                            Word.Range range = null;
                            try
                            {
                                math = omaths[index];
                                range = math.Range;
                                items.Add(new JObject
                                {
                                    { "index", index },
                                    { "paragraphIndex", WordComUtil.FindParagraphIndexAt(
                                        document, range == null ? 0 : range.Start) },
                                    { "type", math.Type.ToString() },
                                    { "typeCode", (int)math.Type },
                                    { "rangeStart", range == null ? 0 : range.Start },
                                    { "rangeEnd", range == null ? 0 : range.End },
                                    { "textPreview", WordComUtil.NormalizePreview(
                                        range == null ? string.Empty : range.Text, 160) }
                                });
                            }
                            finally
                            {
                                ComHelper.Release(range, math);
                            }
                        }
                        JObject data = new JObject();
                        data["total"] = count;
                        data["items"] = items;
                        return ReadSucceed("共 " + count + " 个公式（OMath）。",
                            data, "Word:Equations");
                    }
                    finally
                    {
                        ComHelper.Release(omaths);
                    }
                });
        }

        private static int CountEquations(Word.Document document)
        {
            Word.OMaths omaths = null;
            try
            {
                omaths = document.OMaths;
                return omaths == null ? 0 : omaths.Count;
            }
            finally
            {
                ComHelper.Release(omaths);
            }
        }

        private static int FindEquationIndex(Word.Document document, int start)
        {
            Word.OMaths omaths = null;
            try
            {
                omaths = document.OMaths;
                int count = omaths == null ? 0 : omaths.Count;
                for (int index = 1; index <= count; index++)
                {
                    Word.OMath math = null;
                    Word.Range range = null;
                    try
                    {
                        math = omaths[index];
                        range = math.Range;
                        if (range != null && range.Start == start)
                        {
                            return index;
                        }
                    }
                    finally
                    {
                        ComHelper.Release(range, math);
                    }
                }
                return 0;
            }
            finally
            {
                ComHelper.Release(omaths);
            }
        }

        private static int FindEquationIndexByText(Word.Document document, string equation)
        {
            if (string.IsNullOrWhiteSpace(equation))
            {
                return 0;
            }
            string normalizedEquation = NormalizeEquationText(equation);
            Word.OMaths omaths = null;
            try
            {
                omaths = document.OMaths;
                int count = omaths == null ? 0 : omaths.Count;
                for (int index = 1; index <= count; index++)
                {
                    Word.OMath math = null;
                    Word.Range range = null;
                    try
                    {
                        math = omaths[index];
                        range = math.Range;
                        string text = range == null ? string.Empty : NormalizeEquationText(range.Text);
                        if (text.IndexOf(normalizedEquation, StringComparison.Ordinal) >= 0 ||
                            normalizedEquation.IndexOf(text, StringComparison.Ordinal) >= 0)
                        {
                            return index;
                        }
                    }
                    finally
                    {
                        ComHelper.Release(range, math);
                    }
                }
                return 0;
            }
            finally
            {
                ComHelper.Release(omaths);
            }
        }

        private static int FindNearestEquationIndex(Word.Document document, int position)
        {
            Word.OMaths omaths = null;
            try
            {
                omaths = document.OMaths;
                int count = omaths == null ? 0 : omaths.Count;
                int bestIndex = 0;
                int bestDistance = int.MaxValue;
                for (int index = 1; index <= count; index++)
                {
                    Word.OMath math = null;
                    Word.Range range = null;
                    try
                    {
                        math = omaths[index];
                        range = math.Range;
                        if (range == null)
                        {
                            continue;
                        }
                        if (position >= range.Start && position <= range.End)
                        {
                            return index;
                        }
                        int distance = Math.Abs(range.Start - position);
                        if (distance < bestDistance)
                        {
                            bestDistance = distance;
                            bestIndex = index;
                        }
                    }
                    finally
                    {
                        ComHelper.Release(range, math);
                    }
                }
                return bestIndex;
            }
            finally
            {
                ComHelper.Release(omaths);
            }
        }

        private static string NormalizeEquationText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }
            return text.Replace("\r", string.Empty)
                       .Replace("\n", string.Empty)
                       .Replace("\t", string.Empty)
                       .Replace(" ", string.Empty)
                       .Trim();
        }
    }
}
