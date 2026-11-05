using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Xlsx
{
    /// <summary>公式中的 A1 引用信息，用于引用平移审计。</summary>
    internal sealed class ExcelFormulaReference
    {
        public string Text { get; set; }

        public string Column { get; set; }

        public int Row { get; set; }

        public bool ColumnAbsolute { get; set; }

        public bool RowAbsolute { get; set; }
    }

    /// <summary>
    /// 公式引用平移引擎：只处理 A1 引用与整列/整行引用，跳过字符串字面量与函数名。
    /// 复制/粘贴公式时的相对引用平移、引用审计和独立求值都复用这里的规则。
    /// </summary>
    internal static class ExcelFormulaReferenceEngine
    {
        private static readonly Regex A1Reference = new Regex(
            @"(?<![A-Za-z0-9_.])(?<colAbs>\$?)(?<col>[A-Za-z]{1,3})(?<rowAbs>\$?)(?<row>[0-9]+)(?![A-Za-z0-9_])",
            RegexOptions.Compiled);

        private static readonly Regex ColumnRange = new Regex(
            @"(?<![A-Za-z0-9_.])(?<startAbs>\$?)(?<start>[A-Za-z]{1,3}):(?<endAbs>\$?)(?<end>[A-Za-z]{1,3})(?![A-Za-z0-9_])",
            RegexOptions.Compiled);

        private static readonly Regex RowRange = new Regex(
            @"(?<![A-Za-z0-9_.])(?<startAbs>\$?)(?<start>[0-9]+):(?<endAbs>\$?)(?<end>[0-9]+)(?![A-Za-z0-9_])",
            RegexOptions.Compiled);

        /// <summary>按行列偏移平移公式中的相对引用；绝对引用保持不变。</summary>
        public static string ShiftFormula(string formula, int rowDelta, int columnDelta)
        {
            if (string.IsNullOrWhiteSpace(formula) || formula[0] != '=')
            {
                return formula;
            }
            if (rowDelta == 0 && columnDelta == 0)
            {
                return formula;
            }

            StringBuilder builder = new StringBuilder(formula.Length + 16);
            int index = 0;
            while (index < formula.Length)
            {
                char current = formula[index];
                if (current == '"')
                {
                    int end = FindStringEnd(formula, index);
                    builder.Append(formula, index, end - index);
                    index = end;
                    continue;
                }

                int nextQuote = formula.IndexOf('"', index);
                int segmentEnd = nextQuote < 0 ? formula.Length : nextQuote;
                string segment = formula.Substring(index, segmentEnd - index);
                builder.Append(ShiftSegment(segment, rowDelta, columnDelta));
                index = segmentEnd;
            }
            return builder.ToString();
        }

        /// <summary>
        /// 将公式中 A1 引用限定到指定工作表并转为绝对引用，供 Application.Evaluate 独立求值。
        /// 公式含工作表分隔符、结构化引用或整行/整列引用时返回 false，避免错误改写。
        /// </summary>
        public static bool TryQualifyForEvaluate(string formula, string sheetName, out string qualified)
        {
            qualified = string.Empty;
            if (string.IsNullOrWhiteSpace(formula) || formula[0] != '=')
            {
                return false;
            }
            if (formula.IndexOf('!') >= 0 || formula.IndexOf('[') >= 0 || formula.IndexOf(']') >= 0)
            {
                return false;
            }
            if (ColumnRange.IsMatch(formula) || RowRange.IsMatch(formula))
            {
                return false;
            }

            StringBuilder builder = new StringBuilder(formula.Length + 32);
            int index = 0;
            while (index < formula.Length)
            {
                char current = formula[index];
                if (current == '"')
                {
                    int end = FindStringEnd(formula, index);
                    builder.Append(formula, index, end - index);
                    index = end;
                    continue;
                }

                int nextQuote = formula.IndexOf('"', index);
                int segmentEnd = nextQuote < 0 ? formula.Length : nextQuote;
                string segment = formula.Substring(index, segmentEnd - index);
                string rewritten = A1Reference.Replace(segment, delegate(Match match)
                {
                    if (IsNonCellToken(segment, match))
                    {
                        return match.Value;
                    }
                    string column = match.Groups["col"].Value.ToUpperInvariant();
                    string row = match.Groups["row"].Value;
                    return ExcelComSupport.QuoteSheetName(sheetName) + "!$" + column + "$" + row;
                });
                builder.Append(rewritten);
                index = segmentEnd;
            }
            qualified = builder.ToString();
            return true;
        }

        /// <summary>枚举公式中的 A1 引用（不含字符串字面量与函数名）。</summary>
        public static List<ExcelFormulaReference> EnumerateReferences(string formula)
        {
            List<ExcelFormulaReference> references = new List<ExcelFormulaReference>();
            if (string.IsNullOrWhiteSpace(formula))
            {
                return references;
            }
            int index = 0;
            while (index < formula.Length)
            {
                char current = formula[index];
                if (current == '"')
                {
                    index = FindStringEnd(formula, index);
                    continue;
                }

                int nextQuote = formula.IndexOf('"', index);
                int segmentEnd = nextQuote < 0 ? formula.Length : nextQuote;
                string segment = formula.Substring(index, segmentEnd - index);
                foreach (Match match in A1Reference.Matches(segment))
                {
                    if (IsNonCellToken(segment, match))
                    {
                        continue;
                    }
                    ExcelFormulaReference reference = new ExcelFormulaReference();
                    reference.Text = match.Value;
                    reference.Column = match.Groups["col"].Value.ToUpperInvariant();
                    reference.Row = int.Parse(match.Groups["row"].Value, CultureInfo.InvariantCulture);
                    reference.ColumnAbsolute = match.Groups["colAbs"].Value.Length > 0;
                    reference.RowAbsolute = match.Groups["rowAbs"].Value.Length > 0;
                    references.Add(reference);
                }
                index = segmentEnd;
            }
            return references;
        }

        /// <summary>去掉字符串外空白并统一大写，便于比较平移后的公式文本。</summary>
        public static string NormalizeFormula(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
            {
                return string.Empty;
            }
            StringBuilder builder = new StringBuilder(formula.Length);
            bool inString = false;
            foreach (char character in formula)
            {
                if (character == '"')
                {
                    inString = !inString;
                    builder.Append(character);
                    continue;
                }
                if (!inString && char.IsWhiteSpace(character))
                {
                    continue;
                }
                builder.Append(inString ? character : char.ToUpperInvariant(character));
            }
            return builder.ToString();
        }

        /// <summary>取公式中第一个 A1 引用的文本；没有引用时返回空字符串。</summary>
        public static string FirstReferenceText(string formula)
        {
            List<ExcelFormulaReference> references = EnumerateReferences(formula);
            return references.Count == 0 ? string.Empty : references[0].Text;
        }

        private static string ShiftSegment(string segment, int rowDelta, int columnDelta)
        {
            string result = segment;
            if (columnDelta != 0)
            {
                result = ColumnRange.Replace(result, delegate(Match match)
                {
                    string start = match.Groups["start"].Value.ToUpperInvariant();
                    string end = match.Groups["end"].Value.ToUpperInvariant();
                    bool startAbs = match.Groups["startAbs"].Value.Length > 0;
                    bool endAbs = match.Groups["endAbs"].Value.Length > 0;
                    string shiftedStart = startAbs ? start : ShiftColumn(start, columnDelta);
                    string shiftedEnd = endAbs ? end : ShiftColumn(end, columnDelta);
                    return (startAbs ? "$" : string.Empty) + shiftedStart + ":" +
                           (endAbs ? "$" : string.Empty) + shiftedEnd;
                });
            }
            if (rowDelta != 0)
            {
                result = RowRange.Replace(result, delegate(Match match)
                {
                    int start = int.Parse(match.Groups["start"].Value, CultureInfo.InvariantCulture);
                    int end = int.Parse(match.Groups["end"].Value, CultureInfo.InvariantCulture);
                    bool startAbs = match.Groups["startAbs"].Value.Length > 0;
                    bool endAbs = match.Groups["endAbs"].Value.Length > 0;
                    int shiftedStart = startAbs ? start : Math.Max(1, start + rowDelta);
                    int shiftedEnd = endAbs ? end : Math.Max(1, end + rowDelta);
                    return (startAbs ? "$" : string.Empty) + shiftedStart.ToString(CultureInfo.InvariantCulture) + ":" +
                           (endAbs ? "$" : string.Empty) + shiftedEnd.ToString(CultureInfo.InvariantCulture);
                });
            }

            result = A1Reference.Replace(result, delegate(Match match)
            {
                if (IsNonCellToken(result, match))
                {
                    return match.Value;
                }
                string column = match.Groups["col"].Value.ToUpperInvariant();
                string rowText = match.Groups["row"].Value;
                bool columnAbs = match.Groups["colAbs"].Value.Length > 0;
                bool rowAbs = match.Groups["rowAbs"].Value.Length > 0;
                int row = int.Parse(rowText, CultureInfo.InvariantCulture);
                string shiftedColumn = columnAbs ? column : ShiftColumn(column, columnDelta);
                int shiftedRow = rowAbs ? row : row + rowDelta;
                if (shiftedRow < 1)
                {
                    shiftedRow = 1;
                }
                return (columnAbs ? "$" : string.Empty) + shiftedColumn +
                       (rowAbs ? "$" : string.Empty) + shiftedRow.ToString(CultureInfo.InvariantCulture);
            });
            return result;
        }

        private static string ShiftColumn(string column, int delta)
        {
            if (delta == 0)
            {
                return column;
            }
            int index = ExcelComSupport.ColumnIndex(column);
            if (index <= 0)
            {
                return column;
            }
            int shifted = Math.Max(1, index + delta);
            return ExcelComSupport.ColumnLetter(shifted);
        }

        private static bool IsNonCellToken(string segment, Match match)
        {
            int index = match.Index + match.Length;
            while (index < segment.Length && char.IsWhiteSpace(segment[index]))
            {
                index++;
            }
            if (index < segment.Length && segment[index] == '(')
            {
                return true;
            }
            if (index < segment.Length && segment[index] == '!')
            {
                return true;
            }
            return index + 1 < segment.Length &&
                   segment[index] == '\'' &&
                   segment[index + 1] == '!';
        }

        private static int FindStringEnd(string formula, int start)
        {
            int index = start + 1;
            while (index < formula.Length)
            {
                if (formula[index] == '"')
                {
                    if (index + 1 < formula.Length && formula[index + 1] == '"')
                    {
                        index += 2;
                        continue;
                    }
                    return index + 1;
                }
                index++;
            }
            return formula.Length;
        }
    }
}
