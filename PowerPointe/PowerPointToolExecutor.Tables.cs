using System;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Office;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointe
{
    public sealed partial class PowerPointToolExecutor
    {
        private static OperationResult ExecuteGetTable(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.Shape tableShape = null;
            PowerPoint.Table table = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                int shapeId = PowerPointArgumentReader.GetInt(arguments, -1, "shapeId");
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                tableShape = FindTableShape(slide, shapeId);
                if (tableShape == null)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        shapeId >= 0
                            ? "找不到 shapeId=" + shapeId.ToString(CultureInfo.InvariantCulture) + " 的表格。"
                            : "当前幻灯片上没有表格。");
                }

                table = tableShape.Table;
                int rows = table.Rows.Count;
                int cols = table.Columns.Count;
                JArray data = new JArray();
                for (int rowIndex = 1; rowIndex <= rows; rowIndex++)
                {
                    JArray row = new JArray();
                    for (int colIndex = 1; colIndex <= cols; colIndex++)
                    {
                        row.Add(ReadTableCellText(table, rowIndex, colIndex));
                    }
                    data.Add(row);
                }

                JArray columnWidths = new JArray();
                for (int colIndex = 1; colIndex <= cols; colIndex++)
                {
                    PowerPoint.Column column = null;
                    try
                    {
                        column = table.Columns[colIndex];
                        columnWidths.Add(PowerPointComUtil.SafeFloat(delegate { return column.Width; }));
                    }
                    finally
                    {
                        ComHelper.Release(column);
                    }
                }

                JArray rowHeights = new JArray();
                for (int rowIndex = 1; rowIndex <= rows; rowIndex++)
                {
                    PowerPoint.Row row = null;
                    try
                    {
                        row = table.Rows[rowIndex];
                        rowHeights.Add(PowerPointComUtil.SafeFloat(delegate { return row.Height; }));
                    }
                    finally
                    {
                        ComHelper.Release(row);
                    }
                }

                JObject headerStyle = new JObject();
                headerStyle["firstRow"] = PowerPointComUtil.SafeBool(delegate
                {
                    return Convert.ToInt32((object)table.FirstRow) != 0;
                });
                headerStyle["firstCol"] = PowerPointComUtil.SafeBool(delegate
                {
                    return Convert.ToInt32((object)table.FirstCol) != 0;
                });
                headerStyle["lastRow"] = PowerPointComUtil.SafeBool(delegate
                {
                    return Convert.ToInt32((object)table.LastRow) != 0;
                });
                headerStyle["lastCol"] = PowerPointComUtil.SafeBool(delegate
                {
                    return Convert.ToInt32((object)table.LastCol) != 0;
                });
                headerStyle["horizBanding"] = PowerPointComUtil.SafeBool(delegate
                {
                    return Convert.ToInt32((object)table.HorizBanding) != 0;
                });
                headerStyle["vertBanding"] = PowerPointComUtil.SafeBool(delegate
                {
                    return Convert.ToInt32((object)table.VertBanding) != 0;
                });

                JObject result = new JObject();
                result["slideIndex"] = slide.SlideIndex;
                result["shapeId"] = PowerPointComUtil.SafeInt(delegate
                {
                    dynamic dynamicShape = tableShape;
                    return Convert.ToInt32((object)dynamicShape.Id);
                });
                result["shapeName"] = tableShape.Name;
                result["rows"] = rows;
                result["cols"] = cols;
                result["data"] = data;
                result["columnWidths"] = columnWidths;
                result["rowHeights"] = rowHeights;
                result["headerStyle"] = headerStyle;
                return Success(
                    "已读取第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片的 " + rows.ToString(CultureInfo.InvariantCulture) + "×" +
                    cols.ToString(CultureInfo.InvariantCulture) + " 表格。",
                    "read",
                    false,
                    SlideRef(presentation, slide) + "!" + tableShape.Name,
                    result);
            }
            finally
            {
                ComHelper.Release(table, tableShape, slide, presentation);
            }
        }

        private static OperationResult ExecuteSetTableCell(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.Shape tableShape = null;
            PowerPoint.Table table = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                int shapeId = PowerPointArgumentReader.GetInt(arguments, -1, "shapeId");
                if (shapeId < 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少必填参数 shapeId。");
                }
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                tableShape = FindTableShape(slide, shapeId);
                if (tableShape == null)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "找不到 shapeId=" + shapeId.ToString(CultureInfo.InvariantCulture) + " 的表格。");
                }
                table = tableShape.Table;
                int rows = table.Rows.Count;
                int cols = table.Columns.Count;

                JArray cells = PowerPointArgumentReader.GetArray(arguments, "cells");
                if (cells == null || cells.Count == 0)
                {
                    int row = PowerPointArgumentReader.GetInt(arguments, -1, "row");
                    int col = PowerPointArgumentReader.GetInt(arguments, -1, "col");
                    if (row < 1 || col < 1)
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "需要提供 cells 数组，或同时提供 row/col 与 value。");
                    }
                    JObject single = new JObject();
                    single["row"] = row;
                    single["col"] = col;
                    single["value"] = arguments["value"] == null
                        ? arguments["text"]
                        : arguments["value"];
                    cells = new JArray();
                    cells.Add(single);
                }

                JArray updated = new JArray();
                for (int index = 0; index < cells.Count; index++)
                {
                    JObject cellSpec = cells[index] as JObject;
                    if (cellSpec == null)
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "cells[" + index.ToString(CultureInfo.InvariantCulture) + "] 必须是对象。");
                    }
                    int row = PowerPointArgumentReader.GetInt(cellSpec, -1, "row");
                    int col = PowerPointArgumentReader.GetInt(cellSpec, -1, "col");
                    if (row < 1 || row > rows || col < 1 || col > cols)
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "cells[" + index.ToString(CultureInfo.InvariantCulture) +
                            "] 的 row/col 超出表格范围（" + rows.ToString(CultureInfo.InvariantCulture) +
                            "×" + cols.ToString(CultureInfo.InvariantCulture) + "）。");
                    }
                    JToken valueToken = cellSpec["value"] ?? cellSpec["text"];
                    string text = valueToken == null || valueToken.Type == JTokenType.Null
                        ? string.Empty
                        : valueToken.ToString();
                    PowerPoint.Cell cell = null;
                    PowerPoint.Shape cellShape = null;
                    PowerPoint.TextRange textRange = null;
                    try
                    {
                        cell = table.Cell(row, col);
                        cellShape = cell.Shape;
                        textRange = cellShape.TextFrame.TextRange;
                        textRange.Text = PowerPointComUtil.NormalizeText(text);
                        JObject item = new JObject();
                        item["row"] = row;
                        item["col"] = col;
                        item["text"] = text;
                        updated.Add(item);
                    }
                    finally
                    {
                        ComHelper.Release(textRange, cellShape, cell);
                    }
                }

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["shapeId"] = shapeId;
                data["updatedCount"] = updated.Count;
                data["cells"] = updated;
                return Success(
                    "已修改第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片表格的 " + updated.Count.ToString(CultureInfo.InvariantCulture) + " 个单元格。",
                    "write",
                    updated.Count > 0,
                    SlideRef(presentation, slide) + "!" + tableShape.Name,
                    data);
            }
            finally
            {
                ComHelper.Release(table, tableShape, slide, presentation);
            }
        }

        private static OperationResult ExecuteFormatTable(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.Shape tableShape = null;
            PowerPoint.Table table = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                int shapeId = PowerPointArgumentReader.GetInt(arguments, -1, "shapeId");
                if (shapeId < 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少必填参数 shapeId。");
                }
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                tableShape = FindTableShape(slide, shapeId);
                if (tableShape == null)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "找不到 shapeId=" + shapeId.ToString(CultureInfo.InvariantCulture) + " 的表格。");
                }
                table = tableShape.Table;
                int rows = table.Rows.Count;
                int cols = table.Columns.Count;
                JArray applied = new JArray();
                JArray warnings = new JArray();

                if (HasArgument(arguments, "firstRow"))
                {
                    table.FirstRow = PowerPointArgumentReader.GetBool(arguments, true, "firstRow");
                    applied.Add("firstRow");
                }
                if (HasArgument(arguments, "horizBanding"))
                {
                    table.HorizBanding = PowerPointArgumentReader.GetBool(arguments, true, "horizBanding");
                    applied.Add("horizBanding");
                }
                if (HasArgument(arguments, "firstCol"))
                {
                    table.FirstCol = PowerPointArgumentReader.GetBool(arguments, false, "firstCol");
                    applied.Add("firstCol");
                }
                if (HasArgument(arguments, "lastRow"))
                {
                    table.LastRow = PowerPointArgumentReader.GetBool(arguments, false, "lastRow");
                    applied.Add("lastRow");
                }
                if (HasArgument(arguments, "lastCol"))
                {
                    table.LastCol = PowerPointArgumentReader.GetBool(arguments, false, "lastCol");
                    applied.Add("lastCol");
                }

                bool headerBold = PowerPointArgumentReader.GetBool(arguments, false, "headerBold");
                if (headerBold && rows > 0)
                {
                    for (int col = 1; col <= cols; col++)
                    {
                        ApplyCellFont(table, 1, col, -1, 0f, string.Empty, true);
                    }
                    applied.Add("headerBold");
                }

                string headerFill = PowerPointArgumentReader.GetTrimmedString(arguments, "headerFillColor");
                if (headerFill.Length > 0 && rows > 0)
                {
                    int color;
                    if (!PowerPointComUtil.TryParseColor(headerFill, out color))
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "headerFillColor 不是有效颜色：" + headerFill);
                    }
                    for (int col = 1; col <= cols; col++)
                    {
                        ApplyCellFill(table, 1, col, color);
                    }
                    applied.Add("headerFillColor");
                }

                string zebraFill = PowerPointArgumentReader.GetTrimmedString(arguments, "zebraFillColor");
                if (zebraFill.Length > 0)
                {
                    int color;
                    if (!PowerPointComUtil.TryParseColor(zebraFill, out color))
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "zebraFillColor 不是有效颜色：" + zebraFill);
                    }
                    for (int row = 3; row <= rows; row += 2)
                    {
                        for (int col = 1; col <= cols; col++)
                        {
                            ApplyCellFill(table, row, col, color);
                        }
                    }
                    applied.Add("zebraFillColor");
                }

                string borderColor = PowerPointArgumentReader.GetTrimmedString(arguments, "borderColor");
                float borderWidth = PowerPointArgumentReader.GetFloat(arguments, 1f, "borderWidth");
                if (borderColor.Length > 0)
                {
                    int color;
                    if (!PowerPointComUtil.TryParseColor(borderColor, out color))
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "borderColor 不是有效颜色：" + borderColor);
                    }
                    if (borderWidth <= 0)
                    {
                        borderWidth = 1f;
                    }
                    for (int row = 1; row <= rows; row++)
                    {
                        for (int col = 1; col <= cols; col++)
                        {
                            try
                            {
                                ApplyCellBorders(table, row, col, color, borderWidth);
                            }
                            catch (Exception ex)
                            {
                                warnings.Add("单元格 " + row.ToString(CultureInfo.InvariantCulture) + "," +
                                             col.ToString(CultureInfo.InvariantCulture) + " 边框设置失败：" + ex.Message);
                            }
                        }
                    }
                    applied.Add("borderColor");
                }

                JArray cellFormats = PowerPointArgumentReader.GetArray(arguments, "cells");
                if (cellFormats != null && cellFormats.Count > 0)
                {
                    for (int index = 0; index < cellFormats.Count; index++)
                    {
                        JObject spec = cellFormats[index] as JObject;
                        if (spec == null)
                        {
                            warnings.Add("cells[" + index.ToString(CultureInfo.InvariantCulture) + "] 不是对象，已跳过。");
                            continue;
                        }
                        int row = PowerPointArgumentReader.GetInt(spec, -1, "row");
                        int col = PowerPointArgumentReader.GetInt(spec, -1, "col");
                        if (row < 1 || row > rows || col < 1 || col > cols)
                        {
                            warnings.Add("cells[" + index.ToString(CultureInfo.InvariantCulture) + "] 超出表格范围。");
                            continue;
                        }
                        string fillColor = PowerPointArgumentReader.GetTrimmedString(spec, "fillColor");
                        if (fillColor.Length > 0)
                        {
                            int color;
                            if (PowerPointComUtil.TryParseColor(fillColor, out color))
                            {
                                ApplyCellFill(table, row, col, color);
                            }
                        }
                        string fontColor = PowerPointArgumentReader.GetTrimmedString(spec, "fontColor");
                        bool bold = PowerPointArgumentReader.GetBool(spec, false, "bold");
                        float fontSize = PowerPointArgumentReader.GetFloat(spec, 0f, "fontSize");
                        if (fontColor.Length > 0 || bold || fontSize > 0)
                        {
                            ApplyCellFont(table, row, col, bold ? 1 : 0, fontSize, fontColor, bold);
                        }
                    }
                    applied.Add("cells");
                }

                if (applied.Count == 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "没有提供任何可应用的表格格式属性。");
                }

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["shapeId"] = shapeId;
                data["applied"] = applied;
                if (warnings.Count > 0)
                {
                    data["warnings"] = warnings;
                }
                return Success(
                    "已格式化第 " + slide.SlideIndex.ToString(CultureInfo.InvariantCulture) +
                    " 张幻灯片的表格（" + applied.Count.ToString(CultureInfo.InvariantCulture) + " 项）。",
                    "write",
                    true,
                    SlideRef(presentation, slide) + "!" + tableShape.Name,
                    data);
            }
            finally
            {
                ComHelper.Release(table, tableShape, slide, presentation);
            }
        }

        private static PowerPoint.Shape FindTableShape(PowerPoint.Slide slide, int shapeId)
        {
            if (shapeId >= 0)
            {
                PowerPoint.Shape shape = PowerPointComUtil.FindShape(slide, shapeId);
                if (shape == null)
                {
                    return null;
                }
                dynamic dynamicShape = shape;
                bool hasTable = false;
                try
                {
                    hasTable = Convert.ToInt32((object)dynamicShape.HasTable) != 0;
                }
                catch (Exception)
                {
                    hasTable = false;
                }
                if (!hasTable)
                {
                    ComHelper.Release(shape);
                    return null;
                }
                return shape;
            }

            PowerPoint.Shapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                for (int index = 1; index <= shapes.Count; index++)
                {
                    PowerPoint.Shape shape = null;
                    bool matched = false;
                    try
                    {
                        shape = shapes[index];
                        dynamic dynamicShape = shape;
                        if (Convert.ToInt32((object)dynamicShape.HasTable) != 0)
                        {
                            matched = true;
                            return shape;
                        }
                    }
                    catch (Exception)
                    {
                    }
                    finally
                    {
                        if (!matched)
                        {
                            ComHelper.Release(shape);
                        }
                    }
                }
            }
            finally
            {
                ComHelper.Release(shapes);
            }
            return null;
        }

        private static string ReadTableCellText(PowerPoint.Table table, int row, int col)
        {
            PowerPoint.Cell cell = null;
            PowerPoint.Shape shape = null;
            PowerPoint.TextRange textRange = null;
            try
            {
                cell = table.Cell(row, col);
                shape = cell.Shape;
                textRange = shape.TextFrame.TextRange;
                return textRange.Text ?? string.Empty;
            }
            catch (Exception)
            {
                return string.Empty;
            }
            finally
            {
                ComHelper.Release(textRange, shape, cell);
            }
        }

        private static void ApplyCellFill(PowerPoint.Table table, int row, int col, int oleColor)
        {
            PowerPoint.Cell cell = null;
            PowerPoint.Shape shape = null;
            try
            {
                cell = table.Cell(row, col);
                shape = cell.Shape;
                dynamic fill = shape.Fill;
                fill.Solid();
                fill.Visible = -1;
                fill.ForeColor.RGB = oleColor;
            }
            finally
            {
                ComHelper.Release(shape, cell);
            }
        }

        private static void ApplyCellFont(
            PowerPoint.Table table,
            int row,
            int col,
            int bold,
            float fontSize,
            string fontColor,
            bool setBold)
        {
            PowerPoint.Cell cell = null;
            PowerPoint.Shape shape = null;
            PowerPoint.TextRange textRange = null;
            try
            {
                cell = row == 0 ? null : table.Cell(row, col);
                if (cell == null)
                {
                    return;
                }
                shape = cell.Shape;
                textRange = shape.TextFrame.TextRange;
                dynamic font = textRange.Font;
                if (setBold)
                {
                    font.Bold = bold != 0 ? -1 : 0;
                }
                if (fontSize > 0)
                {
                    font.Size = fontSize;
                }
                int oleColor;
                if (fontColor.Length > 0 && PowerPointComUtil.TryParseColor(fontColor, out oleColor))
                {
                    font.Color.RGB = oleColor;
                }
            }
            finally
            {
                ComHelper.Release(textRange, shape, cell);
            }
        }

        private static void ApplyCellBorders(
            PowerPoint.Table table,
            int row,
            int col,
            int oleColor,
            float width)
        {
            PowerPoint.Cell cell = null;
            try
            {
                cell = table.Cell(row, col);
                dynamic borders = cell.Borders;
                for (int borderIndex = 1; borderIndex <= 4; borderIndex++)
                {
                    try
                    {
                        dynamic border = borders.Item(borderIndex);
                        border.Enable = -1;
                        border.ForeColor.RGB = oleColor;
                        border.Weight = width;
                        ComHelper.Release(border);
                    }
                    catch (Exception)
                    {
                    }
                }
                ComHelper.Release(borders);
            }
            finally
            {
                ComHelper.Release(cell);
            }
        }
    }
}
