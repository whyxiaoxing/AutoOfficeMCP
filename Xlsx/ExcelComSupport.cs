using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>
    /// Excel 执行器通用工具：数组转换、颜色、COM 对象查找、公式写入与单元格值序列化。
    /// 本类不持有 Application 会话，所有调用方仍须遵守 finally 释放纪律。
    /// </summary>
    /// <summary>PIA 15 的 Address 是带索引参数的属性，这里集中封装调用形状。</summary>
    internal static class ExcelRangeExtensions
    {
        public static string AddressLocal(this Excel.Range range)
        {
            if (range == null)
            {
                return string.Empty;
            }
            return range.Address[
                false,
                false,
                Excel.XlReferenceStyle.xlA1,
                Type.Missing,
                Type.Missing];
        }
    }

    internal static class ExcelComSupport
    {
        private static readonly Dictionary<string, int> NamedColors = BuildNamedColors();

        public static Excel.Workbook RequireActiveWorkbook(Excel.Application application)
        {
            if (application == null)
            {
                throw new OfficeComException(OfficeErrorCode.APP_NOT_RUNNING, "Excel 应用不可用。");
            }
            Excel.Workbook workbook = application.ActiveWorkbook;
            if (workbook == null)
            {
                throw new OfficeComException(OfficeErrorCode.DOC_NOT_OPEN, "当前没有打开的工作簿，请先打开或新建工作簿。");
            }
            return workbook;
        }

        public static Excel.Worksheet RequireActiveWorksheet(Excel.Application application)
        {
            object activeSheet = application.ActiveSheet;
            Excel.Worksheet worksheet = activeSheet as Excel.Worksheet;
            if (worksheet == null)
            {
                ComHelper.Release(activeSheet);
                throw new OfficeComException(OfficeErrorCode.DOC_NOT_OPEN, "当前活动对象不是普通工作表。");
            }
            return worksheet;
        }

        public static void EnsureWritable(Excel.Application application, Excel.Workbook workbook, Excel.Worksheet worksheet, string toolName)
        {
            if (application == null || workbook == null || worksheet == null)
            {
                throw new OfficeComException(OfficeErrorCode.DOC_NOT_OPEN, "工作簿或工作表不可用，无法执行写入。");
            }
            if (workbook.ReadOnly)
            {
                throw new OfficeComException(
                    OfficeErrorCode.READ_ONLY,
                    "工具「" + toolName + "」无法写入：工作簿以只读方式打开。请另存为可写副本后重试。");
            }
            if (worksheet.ProtectContents)
            {
                throw new OfficeComException(
                    OfficeErrorCode.SHEET_PROTECTED,
                    "工具「" + toolName + "」无法写入：工作表「" + worksheet.Name + "」受保护。请先取消工作表保护。");
            }
        }

        public static Excel.Worksheet FindWorksheet(Excel.Workbook workbook, string sheetName)
        {
            if (workbook == null || string.IsNullOrWhiteSpace(sheetName))
            {
                return null;
            }
            try
            {
                object value = workbook.Worksheets[sheetName];
                Excel.Worksheet worksheet = value as Excel.Worksheet;
                if (worksheet == null)
                {
                    ComHelper.Release(value);
                }
                return worksheet;
            }
            catch (COMException)
            {
                // 名称大小写/全半角差异在少数环境下会走索引失败，下面按实际名称再找一次。
            }

            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                int count = sheets.Count;
                for (int index = 1; index <= count; index++)
                {
                    object item = sheets[index];
                    Excel.Worksheet candidate = item as Excel.Worksheet;
                    if (candidate == null)
                    {
                        ComHelper.Release(item);
                        continue;
                    }
                    string name = candidate.Name;
                    if (string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        return candidate;
                    }
                    ComHelper.Release(candidate);
                }
                return null;
            }
            finally
            {
                ComHelper.Release(sheets);
            }
        }

        public static string QuoteSheetName(string sheetName)
        {
            string safe = sheetName ?? string.Empty;
            return "'" + safe.Replace("'", "''") + "'";
        }

        public static string FullAddress(string sheetName, string localAddress)
        {
            return QuoteSheetName(sheetName) + "!" + (localAddress ?? string.Empty);
        }

        public static string RangeAddress(Excel.Range range, string sheetName)
        {
            if (range == null)
            {
                return string.Empty;
            }
            string local = range.AddressLocal();
            return string.IsNullOrWhiteSpace(sheetName) ? local : FullAddress(sheetName, local);
        }

        public static long CellCount(Excel.Range range)
        {
            if (range == null)
            {
                return 0;
            }
            return (long)range.Rows.Count * (long)range.Columns.Count;
        }

        public static Excel.Range ResizeRange(Excel.Range range, int rows, int columns)
        {
            if (range == null)
            {
                return null;
            }
            return range.Resize[rows, columns];
        }

        public static string OffsetAddress(Excel.Range source, int rowOffset, int columnOffset)
        {
            if (source == null)
            {
                return string.Empty;
            }
            int row = source.Row + rowOffset;
            int column = source.Column + columnOffset;
            return ColumnLetter(column) + row.ToString(CultureInfo.InvariantCulture);
        }

        public static string DefaultOutputAddress(Excel.Range source)
        {
            if (source == null)
            {
                return "A1";
            }
            int columns = source.Columns.Count;
            return OffsetAddress(source, 0, columns + 1);
        }

        public static string ColumnLetter(int columnNumber)
        {
            if (columnNumber <= 0)
            {
                return string.Empty;
            }
            int dividend = columnNumber;
            StringBuilder builder = new StringBuilder();
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                builder.Insert(0, (char)('A' + modulo));
                dividend = (dividend - modulo) / 26;
            }
            return builder.ToString();
        }

        public static int ColumnIndex(string column)
        {
            if (string.IsNullOrWhiteSpace(column))
            {
                return 0;
            }
            string text = column.Trim();
            int number = 0;
            foreach (char character in text)
            {
                char upper = char.ToUpperInvariant(character);
                if (upper < 'A' || upper > 'Z')
                {
                    return 0;
                }
                number = number * 26 + (upper - 'A' + 1);
            }
            return number;
        }

        public static string ColumnOffset(string column, int offset)
        {
            int index = ColumnIndex(column);
            if (index <= 0)
            {
                return string.Empty;
            }
            return ColumnLetter(index + offset);
        }

        public static int ParseColumnToken(string token, int fallback)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return fallback;
            }
            int number;
            if (int.TryParse(token.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out number))
            {
                return number;
            }
            int index = ColumnIndex(token.Trim());
            return index > 0 ? index : fallback;
        }

        public static int ParseColor(string colorText, int defaultColor)
        {
            if (string.IsNullOrWhiteSpace(colorText))
            {
                return defaultColor;
            }

            string text = colorText.Trim();
            int named;
            if (NamedColors.TryGetValue(text, out named))
            {
                return named;
            }

            if (text.StartsWith("#", StringComparison.Ordinal))
            {
                text = text.Substring(1);
            }
            else if (text.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            {
                text = text.Substring(2);
            }

            if (text.Length == 6)
            {
                int red;
                int green;
                int blue;
                if (int.TryParse(text.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out red) &&
                    int.TryParse(text.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out green) &&
                    int.TryParse(text.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out blue))
                {
                    return ToOleColor(red, green, blue);
                }
            }

            throw new OfficeComException(
                OfficeErrorCode.INVALID_PARAMS,
                "无法识别颜色「" + colorText + "」。请使用 #RRGGBB 或 red/blue/green/yellow/orange/gray/white/black。");
        }

        public static int ToOleColor(int red, int green, int blue)
        {
            return (blue << 16) | (green << 8) | red;
        }

        public static bool TrySetFormula2(Excel.Range range, object formula, out string engine)
        {
            engine = string.Empty;
            if (range == null)
            {
                return false;
            }
            try
            {
                dynamic dynamicRange = range;
                dynamicRange.Formula2 = formula;
                engine = "Formula2";
                return true;
            }
            catch (Exception)
            {
                // 旧版 Excel 没有 Formula2，调用方回退到 Formula；原公式错误会在回退时真实抛出。
                engine = string.Empty;
                return false;
            }
        }

        public static void SetFormula(Excel.Range range, object formula, out string engine)
        {
            if (TrySetFormula2(range, formula, out engine))
            {
                return;
            }
            range.Formula = formula;
            engine = "Formula";
        }

        public static object[,] JsonToMatrix(JToken token, out int rows, out int columns, out bool scalar)
        {
            rows = 0;
            columns = 0;
            scalar = false;
            if (token == null || token.Type == JTokenType.Null)
            {
                throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "data 不能为空。");
            }

            JArray array = token as JArray;
            if (array == null)
            {
                scalar = true;
                rows = 1;
                columns = 1;
                object[,] single = new object[1, 1];
                single[0, 0] = JTokenToValue(token);
                return single;
            }

            if (array.Count == 0)
            {
                throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "data 数组不能为空。");
            }

            bool nested = false;
            foreach (JToken item in array)
            {
                if (item is JArray)
                {
                    nested = true;
                    break;
                }
            }

            if (!nested)
            {
                rows = 1;
                columns = array.Count;
                object[,] row = new object[rows, columns];
                for (int column = 0; column < columns; column++)
                {
                    row[0, column] = JTokenToValue(array[column]);
                }
                return row;
            }

            rows = array.Count;
            columns = 0;
            foreach (JToken item in array)
            {
                JArray child = item as JArray;
                if (child != null && child.Count > columns)
                {
                    columns = child.Count;
                }
            }
            if (columns == 0)
            {
                throw new OfficeComException(OfficeErrorCode.INVALID_PARAMS, "data 二维数组每行都为空。");
            }

            object[,] matrix = new object[rows, columns];
            for (int rowIndex = 0; rowIndex < rows; rowIndex++)
            {
                JArray child = array[rowIndex] as JArray;
                for (int columnIndex = 0; columnIndex < columns; columnIndex++)
                {
                    matrix[rowIndex, columnIndex] = child != null && columnIndex < child.Count
                        ? JTokenToValue(child[columnIndex])
                        : null;
                }
            }
            return matrix;
        }

        public static object JTokenToValue(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null || token.Type == JTokenType.Undefined)
            {
                return null;
            }
            switch (token.Type)
            {
                case JTokenType.Integer:
                    return token.Value<long>();
                case JTokenType.Float:
                    return token.Value<double>();
                case JTokenType.Boolean:
                    return token.Value<bool>();
                case JTokenType.Date:
                    return token.Value<DateTime>();
                case JTokenType.String:
                    return token.Value<string>();
                case JTokenType.Object:
                case JTokenType.Array:
                    return token.ToString(Formatting.None);
                default:
                    return token.ToString();
            }
        }

        public static JToken ToJToken(object value)
        {
            if (value == null || value is DBNull)
            {
                return JValue.CreateNull();
            }
            if (value is object[,])
            {
                return MatrixToJArray((object[,])value);
            }
            if (value is Array)
            {
                Array array = (Array)value;
                JArray result = new JArray();
                foreach (object item in array)
                {
                    result.Add(ToJToken(item));
                }
                return result;
            }
            if (value is DateTime)
            {
                return new JValue(((DateTime)value).ToString("o", CultureInfo.InvariantCulture));
            }
            if (value is double)
            {
                double number = (double)value;
                if (double.IsNaN(number) || double.IsInfinity(number))
                {
                    return new JValue(number.ToString(CultureInfo.InvariantCulture));
                }
                return new JValue(number);
            }
            if (value is float)
            {
                return new JValue(Convert.ToDouble(value, CultureInfo.InvariantCulture));
            }
            if (value is int)
            {
                string errorText = ExcelErrorText((int)value);
                return errorText.Length > 0 ? new JValue(errorText) : new JValue((int)value);
            }
            if (value is long)
            {
                return new JValue((long)value);
            }
            if (value is bool)
            {
                return new JValue((bool)value);
            }
            return new JValue(Convert.ToString(value, CultureInfo.InvariantCulture));
        }

        public static JArray MatrixToJArray(object[,] matrix)
        {
            JArray result = new JArray();
            if (matrix == null)
            {
                return result;
            }
            int rowLower = matrix.GetLowerBound(0);
            int columnLower = matrix.GetLowerBound(1);
            int rows = matrix.GetLength(0);
            int columns = matrix.GetLength(1);
            for (int row = 0; row < rows; row++)
            {
                JArray rowArray = new JArray();
                for (int column = 0; column < columns; column++)
                {
                    rowArray.Add(ToJToken(matrix[rowLower + row, columnLower + column]));
                }
                result.Add(rowArray);
            }
            return result;
        }

        public static object[,] NormalizeMatrix(object raw, int rows, int columns)
        {
            object[,] normalized = new object[rows, columns];
            if (raw is object[,])
            {
                object[,] source = (object[,])raw;
                int sourceRowLower = source.GetLowerBound(0);
                int sourceColumnLower = source.GetLowerBound(1);
                int sourceRows = source.GetLength(0);
                int sourceColumns = source.GetLength(1);
                for (int row = 0; row < rows; row++)
                {
                    for (int column = 0; column < columns; column++)
                    {
                        if (row < sourceRows && column < sourceColumns)
                        {
                            normalized[row, column] = source[sourceRowLower + row, sourceColumnLower + column];
                        }
                    }
                }
                return normalized;
            }

            // 单单元格读取时 Excel 返回标量而不是二维数组。
            normalized[0, 0] = raw;
            return normalized;
        }

        public static object[,] ReadMatrix(Excel.Range range, int rows, int columns)
        {
            object raw = range.Value2;
            return NormalizeMatrix(raw, rows, columns);
        }

        public static string CellText(Excel.Range cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }
            try
            {
                object value = cell.Value2;
                if (value != null && !(value is DBNull))
                {
                    return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
                }
            }
            catch (COMException)
            {
                return string.Empty;
            }
            return string.Empty;
        }

        public static int CountMatches(object[,] values, string findText, bool matchCase, bool entireCell)
        {
            if (values == null || string.IsNullOrEmpty(findText))
            {
                return 0;
            }
            StringComparison comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            int count = 0;
            int rows = values.GetLength(0);
            int columns = values.GetLength(1);
            for (int row = 0; row < rows; row++)
            {
                for (int column = 0; column < columns; column++)
                {
                    string text = Convert.ToString(values[row, column], CultureInfo.InvariantCulture) ?? string.Empty;
                    if (entireCell)
                    {
                        if (string.Equals(text, findText, comparison))
                        {
                            count++;
                        }
                    }
                    else if (text.IndexOf(findText, comparison) >= 0)
                    {
                        count++;
                    }
                }
            }
            return count;
        }

        public static int ResolveColumnIndex(JObject arguments, JToken token, string[] headers, bool hasHeader, int fallback)
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                return fallback;
            }
            if (token.Type == JTokenType.Integer)
            {
                return token.Value<int>();
            }

            string text = token.ToString().Trim();
            int numeric;
            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out numeric))
            {
                return numeric;
            }

            if (hasHeader && headers != null)
            {
                for (int index = 0; index < headers.Length; index++)
                {
                    if (string.Equals(headers[index], text, StringComparison.OrdinalIgnoreCase))
                    {
                        return index + 1;
                    }
                }
            }

            int column = ColumnIndex(text);
            return column > 0 ? column : fallback;
        }

        public static string ExcelErrorText(int value)
        {
            switch (value)
            {
                case -2146826281:
                    return "#DIV/0!";
                case -2146826246:
                    return "#N/A";
                case -2146826259:
                    return "#NAME?";
                case -2146826288:
                    return "#NULL!";
                case -2146826252:
                    return "#NUM!";
                case -2146826265:
                    return "#REF!";
                case -2146826273:
                    return "#VALUE!";
                default:
                    return string.Empty;
            }
        }

        public static bool IsExcelErrorToken(JToken token)
        {
            if (token == null || token.Type != JTokenType.String)
            {
                return false;
            }
            string text = token.Value<string>() ?? string.Empty;
            return text.StartsWith("#", StringComparison.Ordinal);
        }

        public static void ApplyBorder(Excel.Range range, string borders)
        {
            if (range == null || string.IsNullOrWhiteSpace(borders))
            {
                return;
            }
            string value = borders.Trim().ToLowerInvariant();
            Excel.Borders collection = null;
            try
            {
                collection = range.Borders;
                if (value == "true" || value == "all")
                {
                    collection.LineStyle = Excel.XlLineStyle.xlContinuous;
                    collection.Weight = Excel.XlBorderWeight.xlThin;
                }
                else if (value == "outline")
                {
                    range.BorderAround(
                        Excel.XlLineStyle.xlContinuous,
                        Excel.XlBorderWeight.xlThin,
                        Excel.XlColorIndex.xlColorIndexAutomatic,
                        Type.Missing);
                }
                else if (value == "none" || value == "false")
                {
                    collection.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                }
                else
                {
                    throw new OfficeComException(
                        OfficeErrorCode.INVALID_PARAMS,
                        "borders 仅支持 true/all/outline/none/false，当前值：" + borders);
                }
            }
            finally
            {
                ComHelper.Release(collection);
            }
        }

        public static void SetBackgroundColor(Excel.Range range, string colorText)
        {
            if (range == null || string.IsNullOrWhiteSpace(colorText))
            {
                return;
            }
            Excel.Interior interior = null;
            try
            {
                interior = range.Interior;
                interior.Color = ParseColor(colorText, ToOleColor(255, 255, 255));
            }
            finally
            {
                ComHelper.Release(interior);
            }
        }

        public static void SetFontColor(Excel.Range range, string colorText)
        {
            if (range == null || string.IsNullOrWhiteSpace(colorText))
            {
                return;
            }
            Excel.Font font = null;
            try
            {
                font = range.Font;
                font.Color = ParseColor(colorText, ToOleColor(0, 0, 0));
            }
            finally
            {
                ComHelper.Release(font);
            }
        }

        public static void AddSheetSummary(Excel.Worksheet worksheet, List<JObject> target)
        {
            if (worksheet == null || target == null)
            {
                return;
            }
            Excel.Range usedRange = null;
            try
            {
                usedRange = worksheet.UsedRange;
                JObject summary = new JObject();
                summary["name"] = worksheet.Name;
                summary["index"] = worksheet.Index;
                summary["visible"] = VisibilityText(worksheet.Visible);
                summary["protected"] = worksheet.ProtectContents;
                summary["usedRange"] = usedRange.AddressLocal();
                summary["rows"] = usedRange.Rows.Count;
                summary["columns"] = usedRange.Columns.Count;
                target.Add(summary);
            }
            finally
            {
                ComHelper.Release(usedRange);
            }
        }

        public static string VisibilityText(object visibility)
        {
            try
            {
                int value = Convert.ToInt32(visibility, CultureInfo.InvariantCulture);
                if (value == -1)
                {
                    return "visible";
                }
                if (value == 0)
                {
                    return "hidden";
                }
                if (value == 2)
                {
                    return "veryHidden";
                }
                return value.ToString(CultureInfo.InvariantCulture);
            }
            catch
            {
                return "unknown";
            }
        }

        private static Dictionary<string, int> BuildNamedColors()
        {
            Dictionary<string, int> colors = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            colors["red"] = ToOleColor(255, 0, 0);
            colors["green"] = ToOleColor(0, 128, 0);
            colors["blue"] = ToOleColor(0, 0, 255);
            colors["yellow"] = ToOleColor(255, 255, 0);
            colors["orange"] = ToOleColor(255, 165, 0);
            colors["gray"] = ToOleColor(128, 128, 128);
            colors["grey"] = ToOleColor(128, 128, 128);
            colors["white"] = ToOleColor(255, 255, 255);
            colors["black"] = ToOleColor(0, 0, 0);
            colors["purple"] = ToOleColor(128, 0, 128);
            colors["cyan"] = ToOleColor(0, 255, 255);
            colors["pink"] = ToOleColor(255, 192, 203);
            return colors;
        }
    }
}
