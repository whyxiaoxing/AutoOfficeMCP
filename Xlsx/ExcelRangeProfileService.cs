using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>
    /// 范围画像：表头识别、逐列类型统计（numeric/date/text/empty/+formula）与采样。
    /// 同时供 excel_range_profile、excel_get_context 和 excel_profile_files 复用。
    /// </summary>
    internal static class ExcelRangeProfileService
    {
        public static JObject Build(
            Excel.Worksheet worksheet,
            Excel.Range range,
            int maxRows,
            int maxColumns)
        {
            if (range == null)
            {
                throw new OfficeComException(OfficeErrorCode.INVALID_RANGE, "范围为空，无法生成画像。");
            }

            int totalRows = range.Rows.Count;
            int totalColumns = range.Columns.Count;
            int rowsToRead = Math.Min(totalRows, Math.Max(1, maxRows));
            int columnsToRead = Math.Min(totalColumns, Math.Max(1, maxColumns));

            Excel.Range sampleRange = null;
            try
            {
                sampleRange = range.Resize[rowsToRead, columnsToRead];
                object[,] values = ExcelComSupport.ReadMatrix(sampleRange, rowsToRead, columnsToRead);
                object rawFormulas = sampleRange.Formula;
                object[,] formulas = ExcelComSupport.NormalizeMatrix(rawFormulas, rowsToRead, columnsToRead);

                string[] headers = new string[columnsToRead];
                for (int column = 0; column < columnsToRead; column++)
                {
                    headers[column] = Convert.ToString(values[0, column], CultureInfo.InvariantCulture) ?? string.Empty;
                }
                bool hasHeader = rowsToRead > 1 && HasAnyText(headers);

                JArray samples = new JArray();
                for (int row = 0; row < rowsToRead; row++)
                {
                    JArray sampleRow = new JArray();
                    for (int column = 0; column < columnsToRead; column++)
                    {
                        sampleRow.Add(ExcelComSupport.ToJToken(values[row, column]));
                    }
                    samples.Add(sampleRow);
                }

                JArray columnProfiles = new JArray();
                for (int column = 0; column < columnsToRead; column++)
                {
                    int numeric = 0;
                    int date = 0;
                    int text = 0;
                    int empty = 0;
                    int formulaCount = 0;
                    int firstDataRow = hasHeader ? 1 : 0;
                    for (int row = firstDataRow; row < rowsToRead; row++)
                    {
                        string formula = Convert.ToString(formulas[row, column], CultureInfo.InvariantCulture) ?? string.Empty;
                        if (formula.StartsWith("=", StringComparison.Ordinal))
                        {
                            formulaCount++;
                        }

                        object value = values[row, column];
                        if (value == null || value is DBNull ||
                            (value is string && ((string)value).Length == 0))
                        {
                            empty++;
                        }
                        else if (value is DateTime)
                        {
                            date++;
                        }
                        else
                        {
                            double number;
                            if (IsNumeric(value, out number))
                            {
                                numeric++;
                            }
                            else
                            {
                                text++;
                            }
                        }
                    }

                    string kind = "empty";
                    if (numeric + date + text > 0)
                    {
                        if (numeric >= text && numeric >= date)
                        {
                            kind = "numeric";
                        }
                        else if (date >= numeric && date >= text)
                        {
                            kind = "date";
                        }
                        else
                        {
                            kind = "text";
                        }
                    }

                    JObject profile = new JObject();
                    profile["column"] = ExcelComSupport.ColumnLetter(range.Column + column);
                    profile["header"] = headers[column];
                    profile["type"] = formulaCount > 0 ? kind + "+formula" : kind;
                    profile["numeric"] = numeric;
                    profile["date"] = date;
                    profile["text"] = text;
                    profile["empty"] = empty;
                    profile["formula"] = formulaCount;
                    columnProfiles.Add(profile);
                }

                JObject data = new JObject();
                data["sheet"] = worksheet == null ? string.Empty : worksheet.Name;
                data["address"] = range.AddressLocal();
                data["totalRows"] = totalRows;
                data["totalColumns"] = totalColumns;
                data["sampledRows"] = rowsToRead;
                data["sampledColumns"] = columnsToRead;
                data["truncated"] = rowsToRead < totalRows || columnsToRead < totalColumns;
                data["hasHeader"] = hasHeader;
                data["headers"] = new JArray(headers);
                data["columns"] = columnProfiles;
                data["sample"] = samples;
                return data;
            }
            finally
            {
                ComHelper.Release(sampleRange);
            }
        }

        private static bool HasAnyText(string[] headers)
        {
            if (headers == null)
            {
                return false;
            }
            foreach (string header in headers)
            {
                if (!string.IsNullOrWhiteSpace(header))
                {
                    return true;
                }
            }
            return false;
        }

        private static bool IsNumeric(object value, out double number)
        {
            number = 0;
            if (value == null || value is DBNull || value is DateTime)
            {
                return false;
            }
            if (value is double)
            {
                number = (double)value;
                return true;
            }
            if (value is int)
            {
                number = (int)value;
                return true;
            }
            if (value is long)
            {
                number = (long)value;
                return true;
            }
            if (value is decimal)
            {
                number = (double)(decimal)value;
                return true;
            }
            return false;
        }
    }
}
