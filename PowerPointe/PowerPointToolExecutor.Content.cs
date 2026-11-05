using System;
using Newtonsoft.Json.Linq;
using Office;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointe
{
    public sealed partial class PowerPointToolExecutor
    {
        private static OperationResult ExecuteInsertText(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                string content = PowerPointArgumentReader.RequireString(arguments, "content");
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                float x = PowerPointArgumentReader.GetFloat(arguments, 100, "x");
                float y = PowerPointArgumentReader.GetFloat(arguments, 200, "y");
                float width = PowerPointArgumentReader.GetFloat(arguments, 400, "width");
                float height = PowerPointArgumentReader.GetFloat(arguments, 100, "height");
                float fontSize = PowerPointArgumentReader.GetFloat(arguments, 18, "fontSize");
                if (width <= 0 || height <= 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "width/height 必须大于 0。");
                }

                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                PowerPoint.Shape textBox = PowerPointComUtil.AddTextBox(
                    slide, content, x, y, width, height, fontSize);
                try
                {
                    JObject data = new JObject();
                    data["slideIndex"] = slide.SlideIndex;
                    data["shapeId"] = textBox == null
                        ? 0
                        : PowerPointComUtil.SafeInt(delegate
                        {
                            dynamic dynamicShape = textBox;
                            return Convert.ToInt32((object)dynamicShape.Id);
                        });
                    data["shapeName"] = textBox == null ? string.Empty : textBox.Name;
                    data["x"] = x;
                    data["y"] = y;
                    data["width"] = width;
                    data["height"] = height;
                    return Success(
                        "已向第 " + slide.SlideIndex.ToString() + " 张幻灯片插入文本。",
                        "write",
                        true,
                        SlideRef(presentation, slide),
                        data);
                }
                finally
                {
                    ComHelper.Release(textBox);
                }
            }
            finally
            {
                ComHelper.Release(slide, presentation);
            }
        }

        private static OperationResult ExecuteInsertShape(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.Shape shape = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                string shapeType = PowerPointArgumentReader.GetTrimmedString(arguments, "shapeType");
                if (shapeType.Length == 0)
                {
                    shapeType = "rectangle";
                }
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                float x = PowerPointArgumentReader.GetFloat(arguments, 100, "x");
                float y = PowerPointArgumentReader.GetFloat(arguments, 100, "y");
                float width = PowerPointArgumentReader.GetFloat(arguments, 100, "width");
                float height = PowerPointArgumentReader.GetFloat(arguments, 100, "height");
                if (width <= 0 || height <= 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "width/height 必须大于 0。");
                }

                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                shape = PowerPointComUtil.AddShape(
                    slide,
                    PowerPointComUtil.MapShapeType(shapeType),
                    x,
                    y,
                    width,
                    height);

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["shapeId"] = PowerPointComUtil.SafeInt(delegate
                {
                    dynamic dynamicShape = shape;
                    return Convert.ToInt32((object)dynamicShape.Id);
                });
                data["shapeName"] = shape.Name;
                data["shapeType"] = shapeType;
                data["x"] = x;
                data["y"] = y;
                data["width"] = width;
                data["height"] = height;
                return Success(
                    "已在第 " + slide.SlideIndex.ToString() + " 张幻灯片插入形状 " +
                    shape.Name + "。",
                    "write",
                    true,
                    SlideRef(presentation, slide) + "!" + shape.Name,
                    data);
            }
            finally
            {
                ComHelper.Release(shape, slide, presentation);
            }
        }

        private static OperationResult ExecuteInsertTable(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.Shape tableShape = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                int rows = PowerPointArgumentReader.GetInt(arguments, 0, "rows");
                int cols = PowerPointArgumentReader.GetInt(arguments, 0, "cols");
                if (rows <= 0 || cols <= 0)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "rows/cols 必须大于 0。");
                }

                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                float x = PowerPointArgumentReader.GetFloat(arguments, 50, "x");
                float y = PowerPointArgumentReader.GetFloat(arguments, 150, "y");
                float width = PowerPointArgumentReader.GetFloat(arguments, 600, "width");
                float height = PowerPointArgumentReader.GetFloat(arguments, 300, "height");
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                tableShape = slide.Shapes.AddTable(rows, cols, x, y, width, height);
                PowerPoint.Table table = tableShape.Table;
                try
                {
                    JArray data = PowerPointArgumentReader.GetArray(arguments, "data");
                    int filled = FillTable(table, data, rows, cols);
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
                    result["filledCells"] = filled;
                    return Success(
                        "已在第 " + slide.SlideIndex.ToString() + " 张幻灯片插入 " +
                        rows.ToString() + "×" + cols.ToString() + " 表格。",
                        "write",
                        true,
                        SlideRef(presentation, slide) + "!" + tableShape.Name,
                        result);
                }
                finally
                {
                    ComHelper.Release(table);
                }
            }
            finally
            {
                ComHelper.Release(tableShape, slide, presentation);
            }
        }

        private static int FillTable(PowerPoint.Table table, JArray data, int rows, int cols)
        {
            if (table == null || data == null || data.Count == 0)
            {
                return 0;
            }

            int filled = 0;
            int rowCount = Math.Min(rows, data.Count);
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                JArray row = data[rowIndex] as JArray;
                if (row == null)
                {
                    continue;
                }
                int colCount = Math.Min(cols, row.Count);
                for (int colIndex = 0; colIndex < colCount; colIndex++)
                {
                    PowerPoint.Cell cell = null;
                    PowerPoint.Shape cellShape = null;
                    PowerPoint.TextRange textRange = null;
                    try
                    {
                        cell = table.Cell(rowIndex + 1, colIndex + 1);
                        cellShape = cell.Shape;
                        textRange = cellShape.TextFrame.TextRange;
                        JToken value = row[colIndex];
                        textRange.Text = value == null
                            ? string.Empty
                            : PowerPointComUtil.NormalizeText(value.ToString());
                        filled++;
                    }
                    finally
                    {
                        ComHelper.Release(textRange, cellShape, cell);
                    }
                }
            }
            return filled;
        }
    }
}
