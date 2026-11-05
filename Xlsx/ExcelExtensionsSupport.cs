using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Office;
using Office.Config;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>
    /// T06 Excel 扩展工具的公共辅助：
    /// 工作表遍历/释放、矩阵补齐、单元格文本、CSV/JSON 解析、人民币大写转换、路径与输出列解析。
    /// 所有方法只做纯计算或返回调用方负责释放的 COM 对象。
    /// </summary>
    internal static class ExcelExtensionSupport
    {
        private static readonly string[] RmbDigits =
        {
            "零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"
        };

        private static readonly string[] RmbPositionUnits = { "", "拾", "佰", "仟" };

        private static readonly string[] RmbGroupUnits = { "", "万", "亿", "万亿" };

        private const decimal RmbMaximum = 9999999999999999m;

        public static List<Excel.Worksheet> CollectWorksheets(
            Excel.Workbook workbook,
            bool includeHidden,
            string excludeName)
        {
            List<Excel.Worksheet> result = new List<Excel.Worksheet>();
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                int count = sheets.Count;
                for (int index = 1; index <= count; index++)
                {
                    object item = sheets[index];
                    Excel.Worksheet worksheet = item as Excel.Worksheet;
                    if (worksheet == null)
                    {
                        ComHelper.Release(item);
                        continue;
                    }

                    bool visible = worksheet.Visible == Excel.XlSheetVisibility.xlSheetVisible;
                    bool excluded = !string.IsNullOrWhiteSpace(excludeName) &&
                                    string.Equals(worksheet.Name, excludeName, StringComparison.OrdinalIgnoreCase);
                    if (excluded || (!includeHidden && !visible))
                    {
                        ComHelper.Release(worksheet);
                        continue;
                    }
                    result.Add(worksheet);
                }
                return result;
            }
            finally
            {
                ComHelper.Release(sheets);
            }
        }

        public static void ReleaseWorksheets(List<Excel.Worksheet> worksheets)
        {
            if (worksheets == null)
            {
                return;
            }
            foreach (Excel.Worksheet worksheet in worksheets)
            {
                ComHelper.Release(worksheet);
            }
            worksheets.Clear();
        }

        public static Excel.Worksheet FindWorksheet(Excel.Workbook workbook, string sheetName)
        {
            return ExcelComSupport.FindWorksheet(workbook, sheetName);
        }

        public static Excel.Worksheet ResolveWorksheet(
            Excel.Application application,
            Excel.Workbook workbook,
            string sheetName,
            bool createSheet)
        {
            return RangeResolver.ResolveWorksheet(application, workbook, sheetName, createSheet);
        }

        public static Excel.Range GetUsedRange(Excel.Worksheet worksheet)
        {
            if (worksheet == null)
            {
                return null;
            }
            return worksheet.UsedRange;
        }

        public static object[,] ReadMatrix(Excel.Range range, int rows, int columns)
        {
            return ExcelComSupport.ReadMatrix(range, rows, columns);
        }

        public static object[,] ReadFormulaMatrix(Excel.Range range, int rows, int columns)
        {
            object raw = range.Formula;
            return ExcelComSupport.NormalizeMatrix(raw, rows, columns);
        }

        public static object[,] PadMatrix(
            object[,] source,
            int sourceRows,
            int sourceColumns,
            int targetRows,
            int targetColumns)
        {
            object[,] result = new object[targetRows, targetColumns];
            if (source == null)
            {
                return result;
            }
            int copyRows = Math.Min(sourceRows, targetRows);
            int copyColumns = Math.Min(sourceColumns, targetColumns);
            for (int row = 0; row < copyRows; row++)
            {
                for (int column = 0; column < copyColumns; column++)
                {
                    result[row, column] = source[row, column];
                }
            }
            return result;
        }

        public static string CellText(object value)
        {
            if (value == null || value is DBNull)
            {
                return string.Empty;
            }
            if (value is string)
            {
                return (string)value;
            }
            if (value is bool)
            {
                return (bool)value ? "TRUE" : "FALSE";
            }
            if (value is DateTime)
            {
                return ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            }
            if (value is double)
            {
                double number = (double)value;
                if (double.IsNaN(number) || double.IsInfinity(number))
                {
                    return number.ToString(CultureInfo.InvariantCulture);
                }
                return number.ToString("G15", CultureInfo.InvariantCulture);
            }
            if (value is float)
            {
                return ((float)value).ToString("G7", CultureInfo.InvariantCulture);
            }
            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        public static string ErrorTextFromValue(object value)
        {
            if (value == null || value is DBNull)
            {
                return string.Empty;
            }
            if (value is int)
            {
                string mapped = ExcelComSupport.ExcelErrorText((int)value);
                if (mapped.Length > 0)
                {
                    return mapped;
                }
            }
            string text = CellText(value);
            return text.StartsWith("#", StringComparison.Ordinal) ? text : string.Empty;
        }

        public static string CellAddress(int row, int column)
        {
            return ExcelComSupport.ColumnLetter(column) + row.ToString(CultureInfo.InvariantCulture);
        }

        public static string FullCellAddress(string sheetName, int row, int column)
        {
            return ExcelComSupport.FullAddress(sheetName, CellAddress(row, column));
        }

        public static void EnsureRangeCapacity(int row, int column, int rows, int columns)
        {
            if (row < 1 || column < 1 || rows < 1 || columns < 1)
            {
                throw new OfficeComException(OfficeErrorCode.INVALID_RANGE, "目标区域起止行列必须大于 0。");
            }
            if ((long)row + rows - 1 > 1048576L)
            {
                throw new OfficeComException(OfficeErrorCode.INVALID_RANGE, "写入区域超出 Excel 最大行数 1048576。");
            }
            if ((long)column + columns - 1 > 16384L)
            {
                throw new OfficeComException(OfficeErrorCode.INVALID_RANGE, "写入区域超出 Excel 最大列数 16384。");
            }
        }

        /// <summary>
        /// 解析写出的目标列：优先 outputColumn/targetColumn（列字母、数字或源表头名），
        /// 否则默认紧邻源区域右侧一列。
        /// </summary>
        public static int ResolveOutputColumn(JObject arguments, Excel.Range sourceRange, int defaultOffset)
        {
            string token = ExcelArgumentReader.GetTrimmedString(
                arguments,
                "outputColumn",
                "targetColumn",
                "outputCol");
            if (token.Length == 0)
            {
                return sourceRange.Column + sourceRange.Columns.Count + defaultOffset;
            }

            int column = ExcelComSupport.ParseColumnToken(token, 0);
            if (column > 0)
            {
                return column;
            }

            Excel.Worksheet worksheet = null;
            try
            {
                worksheet = sourceRange.Worksheet as Excel.Worksheet;
                string[] headers = RangeResolver.GetHeaders(
                    worksheet,
                    sourceRange,
                    Math.Min(sourceRange.Columns.Count, 50));
                for (int index = 0; index < headers.Length; index++)
                {
                    if (string.Equals(headers[index], token, StringComparison.OrdinalIgnoreCase))
                    {
                        return sourceRange.Column + index;
                    }
                }
            }
            finally
            {
                ComHelper.Release(worksheet);
            }

            throw new OfficeComException(
                OfficeErrorCode.INVALID_PARAMS,
                "无法解析输出列「" + token + "」。请使用列字母（如 F）、列号（6）或源区域表头名。");
        }

        public static string NormalizeDelimiter(string delimiter)
        {
            if (string.IsNullOrEmpty(delimiter))
            {
                return ",";
            }
            if (string.Equals(delimiter, "tab", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(delimiter, "\\t", StringComparison.Ordinal))
            {
                return "\t";
            }
            return delimiter;
        }

        public static Encoding ResolveEncoding(string encodingName, bool withBomForUtf8)
        {
            string name = (encodingName ?? string.Empty).Trim().ToLowerInvariant();
            if (name == "utf8bom" || name == "utf-8-bom" || name == "utf8-bom")
            {
                return new UTF8Encoding(true);
            }
            if (name == "gbk" || name == "gb2312" || name == "cp936" || name == "936")
            {
                return Encoding.GetEncoding(936);
            }
            if (name == "default" || name == "ansi")
            {
                return Encoding.Default;
            }
            return new UTF8Encoding(withBomForUtf8);
        }

        /// <summary>读取文本文件：未显式指定编码时，BOM → UTF-8 严格校验 → GBK 兜底。</summary>
        public static string ReadTextFile(string path, string encodingName)
        {
            byte[] bytes = File.ReadAllBytes(path);
            string name = (encodingName ?? string.Empty).Trim().ToLowerInvariant();
            if (name.Length > 0 && name != "auto")
            {
                return ResolveEncoding(name, false).GetString(bytes);
            }

            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            {
                return new UTF8Encoding(true).GetString(bytes, 3, bytes.Length - 3);
            }
            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
            {
                return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
            }
            if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
            {
                return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
            }

            try
            {
                UTF8Encoding strict = new UTF8Encoding(false, true);
                return strict.GetString(bytes);
            }
            catch (DecoderFallbackException)
            {
                return Encoding.GetEncoding(936).GetString(bytes);
            }
        }

        public static string ResolveFilePath(JObject arguments, bool mustExist, bool overwrite)
        {
            string path = ExcelArgumentReader.GetTrimmedString(
                arguments,
                "path",
                "filePath",
                "outputPath",
                "inputPath");
            if (path.Length == 0)
            {
                throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "缺少 path/filePath 参数。");
            }

            string fullPath;
            string pathError;
            if (!WorkspacePathResolver.TryResolve(path, out fullPath, out pathError))
            {
                throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, pathError);
            }
            if (mustExist)
            {
                if (!File.Exists(fullPath))
                {
                    throw new OfficeComException(
                        OfficeErrorCode.FILE_NOT_FOUND,
                        "文件不存在：" + fullPath);
                }
                return fullPath;
            }

            string directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                string directoryError;
                if (!WorkspacePathResolver.TryEnsureOutputDirectory(directory, out directoryError))
                {
                    throw new OfficeComException(OfficeErrorCode.READ_ONLY, directoryError);
                }
            }
            if (File.Exists(fullPath) && !overwrite)
            {
                throw new OfficeComException(
                    OfficeErrorCode.INVALID_PARAMS,
                    "目标文件已存在：" + fullPath + "。如需覆盖请传 overwrite=true。");
            }
            return fullPath;
        }

        public static object[,] ParseCsv(string text, string delimiter, bool inferTypes, out int rows, out int columns)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "CSV 内容为空。");
            }
            if (delimiter == null || delimiter.Length != 1)
            {
                throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "CSV 分隔符必须是单个字符。");
            }
            char separator = delimiter[0];

            List<List<string>> records = new List<List<string>>();
            List<string> currentRow = new List<string>();
            StringBuilder field = new StringBuilder();
            bool inQuotes = false;

            for (int index = 0; index < text.Length; index++)
            {
                char character = text[index];
                if (inQuotes)
                {
                    if (character == '"')
                    {
                        if (index + 1 < text.Length && text[index + 1] == '"')
                        {
                            field.Append('"');
                            index++;
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        field.Append(character);
                    }
                    continue;
                }

                if (character == '"' && field.Length == 0)
                {
                    inQuotes = true;
                }
                else if (character == separator)
                {
                    currentRow.Add(field.ToString());
                    field.Length = 0;
                }
                else if (character == '\r' || character == '\n')
                {
                    if (character == '\r' && index + 1 < text.Length && text[index + 1] == '\n')
                    {
                        index++;
                    }
                    currentRow.Add(field.ToString());
                    field.Length = 0;
                    records.Add(currentRow);
                    currentRow = new List<string>();
                }
                else
                {
                    field.Append(character);
                }
            }

            if (field.Length > 0 || currentRow.Count > 0)
            {
                currentRow.Add(field.ToString());
                records.Add(currentRow);
            }

            while (records.Count > 0 && records[records.Count - 1].Count == 1 &&
                   records[records.Count - 1][0].Length == 0)
            {
                records.RemoveAt(records.Count - 1);
            }
            if (records.Count == 0)
            {
                throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "CSV 中没有可导入的数据行。");
            }

            rows = records.Count;
            columns = 1;
            foreach (List<string> record in records)
            {
                if (record.Count > columns)
                {
                    columns = record.Count;
                }
            }

            object[,] matrix = new object[rows, columns];
            for (int row = 0; row < rows; row++)
            {
                List<string> record = records[row];
                for (int column = 0; column < columns; column++)
                {
                    string raw = column < record.Count ? record[column] : string.Empty;
                    matrix[row, column] = inferTypes ? InferCsvValue(raw) : raw;
                }
            }
            return matrix;
        }

        private static object InferCsvValue(string raw)
        {
            if (string.IsNullOrEmpty(raw))
            {
                return null;
            }
            if (LooksLikeLeadingZero(raw))
            {
                return raw;
            }

            long integer;
            if (long.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out integer))
            {
                if (Math.Abs((double)integer) < 1e15)
                {
                    return integer;
                }
                return raw;
            }

            double number;
            if (double.TryParse(raw, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number))
            {
                if (Math.Abs(number) < 1e15)
                {
                    return number;
                }
            }
            return raw;
        }

        public static string BuildCsv(object[,] values, string delimiter)
        {
            if (values == null)
            {
                return string.Empty;
            }
            if (delimiter == null || delimiter.Length != 1)
            {
                delimiter = ",";
            }
            int rows = values.GetLength(0);
            int columns = values.GetLength(1);
            StringBuilder builder = new StringBuilder();
            for (int row = 0; row < rows; row++)
            {
                for (int column = 0; column < columns; column++)
                {
                    if (column > 0)
                    {
                        builder.Append(delimiter);
                    }
                    builder.Append(EscapeCsvField(CellText(values[row, column]), delimiter[0]));
                }
                builder.Append("\r\n");
            }
            return builder.ToString();
        }

        private static string EscapeCsvField(string text, char delimiter)
        {
            string value = text ?? string.Empty;
            bool needQuotes = value.IndexOf(delimiter) >= 0 ||
                              value.IndexOf('"') >= 0 ||
                              value.IndexOf('\r') >= 0 ||
                              value.IndexOf('\n') >= 0;
            if (!needQuotes)
            {
                return value;
            }
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        /// <summary>JSON → 二维数组。支持 [{...}] 记录、[[...]]、[值...]、{values:[...]}。</summary>
        public static object[,] ParseJsonMatrix(string text, out int rows, out int columns, out string layout)
        {
            rows = 0;
            columns = 0;
            layout = "scalar";
            JToken token = JToken.Parse(text);
            return BuildJsonMatrix(token, ref rows, ref columns, ref layout);
        }

        private static object[,] BuildJsonMatrix(JToken token, ref int rows, ref int columns, ref string layout)
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "JSON 内容为空。");
            }

            JObject rootObject = token as JObject;
            if (rootObject != null)
            {
                JArray nestedValues = rootObject["values"] as JArray;
                if (nestedValues != null)
                {
                    return BuildJsonMatrix(nestedValues, ref rows, ref columns, ref layout);
                }
                return BuildRecordMatrix(new JArray(rootObject), ref rows, ref columns, ref layout);
            }

            JArray array = token as JArray;
            if (array == null)
            {
                rows = 1;
                columns = 1;
                layout = "scalar";
                object[,] single = new object[1, 1];
                single[0, 0] = ExcelComSupport.JTokenToValue(token);
                return single;
            }
            if (array.Count == 0)
            {
                throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "JSON 数组为空。");
            }

            if (array[0] is JObject)
            {
                return BuildRecordMatrix(array, ref rows, ref columns, ref layout);
            }

            if (array[0] is JArray)
            {
                rows = array.Count;
                columns = 1;
                foreach (JToken item in array)
                {
                    JArray child = item as JArray;
                    if (child != null && child.Count > columns)
                    {
                        columns = child.Count;
                    }
                }
                object[,] matrix = new object[rows, columns];
                for (int row = 0; row < rows; row++)
                {
                    JArray child = array[row] as JArray;
                    for (int column = 0; column < columns; column++)
                    {
                        matrix[row, column] = child != null && column < child.Count
                            ? ExcelComSupport.JTokenToValue(child[column])
                            : null;
                    }
                }
                layout = "arrays";
                return matrix;
            }

            rows = 1;
            columns = array.Count;
            object[,] flat = new object[rows, columns];
            for (int column = 0; column < columns; column++)
            {
                flat[0, column] = ExcelComSupport.JTokenToValue(array[column]);
            }
            layout = "scalars";
            return flat;
        }

        private static object[,] BuildRecordMatrix(JArray array, ref int rows, ref int columns, ref string layout)
        {
            List<string> headers = new List<string>();
            foreach (JToken item in array)
            {
                JObject record = item as JObject;
                if (record == null)
                {
                    continue;
                }
                foreach (JProperty property in record.Properties())
                {
                    if (!headers.Contains(property.Name))
                    {
                        headers.Add(property.Name);
                    }
                }
            }

            if (headers.Count == 0)
            {
                throw new OfficeComException(
                    OfficeErrorCode.INVALID_PARAMS,
                    "JSON 记录为空，无法确定列名。");
            }

            rows = array.Count + 1;
            columns = headers.Count;
            object[,] matrix = new object[rows, columns];
            for (int column = 0; column < columns; column++)
            {
                matrix[0, column] = headers[column];
            }

            for (int rowIndex = 0; rowIndex < array.Count; rowIndex++)
            {
                JObject record = array[rowIndex] as JObject;
                for (int column = 0; column < columns; column++)
                {
                    JToken value = record == null ? null : record[headers[column]];
                    matrix[rowIndex + 1, column] = ExcelComSupport.JTokenToValue(value);
                }
            }
            layout = "records";
            return matrix;
        }

        public static JArray MatrixToJson(object[,] values)
        {
            JArray result = new JArray();
            if (values == null)
            {
                return result;
            }
            int rows = values.GetLength(0);
            int columns = values.GetLength(1);
            for (int row = 0; row < rows; row++)
            {
                JArray line = new JArray();
                for (int column = 0; column < columns; column++)
                {
                    line.Add(ExcelComSupport.ToJToken(values[row, column]));
                }
                result.Add(line);
            }
            return result;
        }

        public static bool TryConvertRmb(object value, out string upper, out string reason)
        {
            upper = string.Empty;
            reason = string.Empty;
            decimal amount;
            if (!TryParseAmount(value, out amount))
            {
                reason = "无法识别为金额";
                return false;
            }
            string converted = ConvertRmb(amount);
            if (converted == null)
            {
                reason = "金额超出支持范围（最大 9999999999999999.99）";
                return false;
            }
            upper = converted;
            return true;
        }

        public static string ConvertRmb(decimal amount)
        {
            decimal rounded = decimal.Round(amount, 2, MidpointRounding.AwayFromZero);
            if (rounded == 0m)
            {
                return "零元整";
            }

            decimal absolute = Math.Abs(rounded);
            long integerPart = (long)decimal.Truncate(absolute);
            if (integerPart > (long)RmbMaximum)
            {
                return null;
            }

            decimal fraction = absolute - integerPart;
            int cents = (int)decimal.Round(fraction * 100m, 0, MidpointRounding.AwayFromZero);
            if (cents == 100)
            {
                integerPart++;
                cents = 0;
            }

            StringBuilder builder = new StringBuilder();
            if (rounded < 0)
            {
                builder.Append("负");
            }
            builder.Append(integerPart == 0 ? "零" : ConvertIntegerPart(integerPart));
            builder.Append("元");

            if (cents == 0)
            {
                builder.Append("整");
                return builder.ToString();
            }

            int jiao = cents / 10;
            int fen = cents % 10;
            if (jiao > 0)
            {
                builder.Append(RmbDigits[jiao]);
                builder.Append("角");
            }
            if (fen > 0)
            {
                if (jiao == 0)
                {
                    builder.Append("零");
                }
                builder.Append(RmbDigits[fen]);
                builder.Append("分");
            }
            return builder.ToString();
        }

        private static string ConvertIntegerPart(long value)
        {
            List<int> groups = new List<int>();
            long remaining = value;
            while (remaining > 0)
            {
                groups.Add((int)(remaining % 10000));
                remaining = remaining / 10000;
            }

            StringBuilder builder = new StringBuilder();
            bool zeroPending = false;
            for (int index = groups.Count - 1; index >= 0; index--)
            {
                int group = groups[index];
                if (group == 0)
                {
                    if (builder.Length > 0)
                    {
                        zeroPending = true;
                    }
                    continue;
                }

                bool endsWithZero = builder.Length > 0 &&
                                    builder[builder.Length - 1] == '零';
                if (zeroPending && builder.Length > 0 && !endsWithZero)
                {
                    builder.Append("零");
                }
                else if (!zeroPending && builder.Length > 0 && group < 1000 && !endsWithZero)
                {
                    builder.Append("零");
                }

                builder.Append(ConvertGroup(group));
                builder.Append(RmbGroupUnits[index]);
                zeroPending = false;
            }
            return builder.ToString();
        }

        private static string ConvertGroup(int group)
        {
            StringBuilder builder = new StringBuilder();
            bool pendingZero = false;
            int[] divisors = { 1000, 100, 10, 1 };
            for (int position = 3; position >= 0; position--)
            {
                int digit = (group / divisors[3 - position]) % 10;
                if (digit == 0)
                {
                    if (builder.Length > 0 && position > 0)
                    {
                        pendingZero = true;
                    }
                    continue;
                }
                if (pendingZero)
                {
                    builder.Append("零");
                    pendingZero = false;
                }
                builder.Append(RmbDigits[digit]);
                builder.Append(RmbPositionUnits[position]);
            }
            return builder.ToString();
        }

        public static bool TryParseAmount(object value, out decimal amount)
        {
            amount = 0m;
            if (value == null || value is DBNull || value is bool || value is DateTime)
            {
                return false;
            }
            if (value is decimal)
            {
                amount = (decimal)value;
                return Math.Abs(amount) <= RmbMaximum;
            }
            if (value is double)
            {
                double number = (double)value;
                if (double.IsNaN(number) || double.IsInfinity(number))
                {
                    return false;
                }
                try
                {
                    amount = Convert.ToDecimal(number, CultureInfo.InvariantCulture);
                }
                catch (OverflowException)
                {
                    return false;
                }
                return Math.Abs(amount) <= RmbMaximum;
            }

            string text = NormalizeFullWidth(CellText(value)).Trim();
            if (text.Length == 0)
            {
                return false;
            }

            bool negativeParentheses = text.StartsWith("(", StringComparison.Ordinal) &&
                                       text.EndsWith(")", StringComparison.Ordinal);
            if (negativeParentheses)
            {
                text = text.Substring(1, text.Length - 2);
            }
            if (text.StartsWith("人民币", StringComparison.Ordinal))
            {
                text = text.Substring(3);
            }
            text = text.Replace("￥", string.Empty)
                       .Replace("¥", string.Empty)
                       .Replace(",", string.Empty)
                       .Replace("，", string.Empty)
                       .Replace(" ", string.Empty)
                       .Trim();
            if (text.EndsWith("元整", StringComparison.Ordinal))
            {
                text = text.Substring(0, text.Length - 2);
            }
            else if (text.EndsWith("元", StringComparison.Ordinal) ||
                     text.EndsWith("圆", StringComparison.Ordinal))
            {
                text = text.Substring(0, text.Length - 1);
            }
            if (text.Length == 0)
            {
                return false;
            }

            decimal parsed;
            if (!decimal.TryParse(
                    text,
                    NumberStyles.Number | NumberStyles.AllowLeadingSign,
                    CultureInfo.InvariantCulture,
                    out parsed))
            {
                return false;
            }
            if (negativeParentheses)
            {
                parsed = -Math.Abs(parsed);
            }
            if (Math.Abs(parsed) > RmbMaximum)
            {
                return false;
            }
            amount = parsed;
            return true;
        }

        public static string NormalizeFullWidth(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }
            StringBuilder builder = new StringBuilder(text.Length);
            foreach (char character in text)
            {
                if (character >= '０' && character <= '９')
                {
                    builder.Append((char)('0' + (character - '０')));
                }
                else if (character == '．')
                {
                    builder.Append('.');
                }
                else if (character == '－')
                {
                    builder.Append('-');
                }
                else if (character == '，')
                {
                    builder.Append(',');
                }
                else if (character == '＋')
                {
                    builder.Append('+');
                }
                else
                {
                    builder.Append(character);
                }
            }
            return builder.ToString();
        }

        public static bool LooksLikeLeadingZero(string text)
        {
            string value = NormalizeFullWidth(text).Trim();
            if (value.StartsWith("+", StringComparison.Ordinal) ||
                value.StartsWith("-", StringComparison.Ordinal))
            {
                value = value.Substring(1);
            }
            if (value.Length <= 1 || value[0] != '0')
            {
                return false;
            }
            char next = value[1];
            return next >= '0' && next <= '9';
        }

        public static bool TryParseTextNumber(string text, out decimal value, out bool isInteger)
        {
            value = 0m;
            isInteger = false;
            string normalized = NormalizeFullWidth(text).Trim();
            if (normalized.Length == 0 || normalized.Length > 40)
            {
                return false;
            }
            if (LooksLikeLeadingZero(normalized))
            {
                return false;
            }
            if (!decimal.TryParse(
                    normalized,
                    NumberStyles.Float | NumberStyles.AllowThousands,
                    CultureInfo.InvariantCulture,
                    out value))
            {
                return false;
            }
            if (Math.Abs(value) >= 1000000000000000m)
            {
                return false;
            }
            isInteger = normalized.IndexOf('.') < 0 &&
                        normalized.IndexOf('e') < 0 &&
                        normalized.IndexOf('E') < 0;
            return true;
        }

        public static object ToExcelNumber(decimal value, bool isInteger)
        {
            if (isInteger && value >= long.MinValue && value <= long.MaxValue)
            {
                return (long)value;
            }
            return (double)value;
        }

        public static bool IsBlankValue(object value)
        {
            if (value == null || value is DBNull)
            {
                return true;
            }
            string text = value as string;
            return text != null && text.Length == 0;
        }

        public static bool MatrixHasData(object[,] values)
        {
            if (values == null)
            {
                return false;
            }
            int rows = values.GetLength(0);
            int columns = values.GetLength(1);
            for (int row = 0; row < rows; row++)
            {
                for (int column = 0; column < columns; column++)
                {
                    if (!IsBlankValue(values[row, column]))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        public static string ToColumnTokenText(JObject arguments, params string[] names)
        {
            return ExcelArgumentReader.GetTrimmedString(arguments, names);
        }

        public static JArray BuildAddressList(IEnumerable<string> addresses, int limit)
        {
            JArray result = new JArray();
            if (addresses == null)
            {
                return result;
            }
            int count = 0;
            foreach (string address in addresses)
            {
                if (count >= limit)
                {
                    break;
                }
                result.Add(address);
                count++;
            }
            return result;
        }

        public static string ToRegexPatternError(ArgumentException exception)
        {
            return "正则表达式无效：" + exception.Message;
        }

        public static string QuoteSheetName(string sheetName)
        {
            return ExcelComSupport.QuoteSheetName(sheetName);
        }

        public static string SheetInternalReference(string sheetName, string cellAddress)
        {
            return ExcelComSupport.QuoteSheetName(sheetName) + "!" + cellAddress;
        }

        public static string TruncateText(string text, int maxLength)
        {
            if (text == null)
            {
                return string.Empty;
            }
            if (maxLength <= 0 || text.Length <= maxLength)
            {
                return text;
            }
            return text.Substring(0, maxLength);
        }

        public static string ContextAround(string text, int matchIndex, int matchLength, int contextLength)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }
            int start = Math.Max(0, matchIndex - contextLength);
            int end = Math.Min(text.Length, matchIndex + matchLength + contextLength);
            string context = text.Substring(start, end - start);
            if (start > 0)
            {
                context = "…" + context;
            }
            if (end < text.Length)
            {
                context = context + "…";
            }
            return context;
        }
    }
}
