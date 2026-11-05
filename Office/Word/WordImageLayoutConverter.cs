using System;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    /// <summary>
    /// inline↔floating 转换的删除重建辅助：
    /// 先把原图片导出为临时 EMF，再删除旧对象并按目标集合重建。
    /// Word PIA 虽有 ConvertToShape/ConvertToInlineShape，但任务要求使用删除重建；
    /// 若无法导出原图会明确失败，不会退化为不可靠的隐式转换。
    /// </summary>
    internal static class WordImageLayoutConverter
    {
        public static string ExportToTempFile(
            Word.Application application,
            Word.Document document,
            Word.InlineShape inlineShape,
            Word.Shape floatingShape,
            out string error)
        {
            error = string.Empty;
            string directory = Path.Combine(Path.GetTempPath(), "AutoOfficeMCP-WordImages");
            try
            {
                Directory.CreateDirectory(directory);
            }
            catch (Exception ex)
            {
                error = "无法创建图片临时目录：" + ex.Message;
                return null;
            }
            string file = Path.Combine(directory, "word-image-" + Guid.NewGuid().ToString("N") + ".emf");
            byte[] bytes = null;
            if (inlineShape != null)
            {
                Word.Range range = null;
                try
                {
                    range = inlineShape.Range;
                    object bits = range.EnhMetaFileBits;
                    bytes = bits as byte[];
                }
                catch (Exception ex)
                {
                    error = "读取内嵌图片 EMF 失败：" + ex.Message;
                }
                finally
                {
                    ComHelper.Release(range);
                }
            }
            else if (floatingShape != null)
            {
                bytes = CopyFloatingToEmf(application, floatingShape, out error);
            }
            else
            {
                error = "没有指定要导出的图片对象。";
            }

            if (bytes == null || bytes.Length == 0)
            {
                if (string.IsNullOrWhiteSpace(error))
                {
                    error = "无法从 Word 图片导出 EMF 字节（该对象可能不是图片）。";
                }
                return null;
            }

            try
            {
                File.WriteAllBytes(file, bytes);
                return file;
            }
            catch (Exception ex)
            {
                error = "写入图片临时文件失败：" + ex.Message;
                return null;
            }
        }

        public static Word.Shape CreateFloatingFromFile(
            Word.Document document,
            string filePath,
            Word.Range anchorRange,
            double left,
            double top,
            Word.WdWrapType wrapType,
            Word.WdRelativeHorizontalPosition relativeHorizontal,
            Word.WdRelativeVerticalPosition relativeVertical,
            string alternativeText,
            double width,
            double height,
            out int shapeIndex)
        {
            Word.Shapes shapes = null;
            Word.Shape shape = null;
            try
            {
                shapes = document.Shapes;
                object linkToFile = false;
                object saveWithDocument = true;
                object leftObject = (float)left;
                object topObject = (float)top;
                object widthObject = (float)width;
                object heightObject = (float)height;
                object anchorObject = anchorRange;
                shape = shapes.AddPicture(
                    filePath,
                    ref linkToFile,
                    ref saveWithDocument,
                    ref leftObject,
                    ref topObject,
                    ref widthObject,
                    ref heightObject,
                    ref anchorObject);
                if (shape == null)
                {
                    throw new OfficeComException(OfficeErrorCode.INTERNAL_ERROR, "Shapes.AddPicture 返回 null。");
                }
                Word.WrapFormat wrap = null;
                try
                {
                    wrap = shape.WrapFormat;
                    wrap.Type = wrapType;
                }
                finally
                {
                    ComHelper.Release(wrap);
                }
                shape.RelativeHorizontalPosition = relativeHorizontal;
                shape.RelativeVerticalPosition = relativeVertical;
                shape.Left = (float)left;
                shape.Top = (float)top;
                shape.Width = (float)width;
                shape.Height = (float)height;
                if (alternativeText != null)
                {
                    shape.AlternativeText = alternativeText;
                }
                shapeIndex = WordComUtil.FindShapeIndex(document, shape);
                return shape;
            }
            finally
            {
                ComHelper.Release(shapes);
            }
        }

        public static Word.InlineShape CreateInlineFromFile(
            Word.Document document,
            string filePath,
            Word.Range anchorRange,
            string alternativeText,
            double width,
            double height,
            out int inlineIndex)
        {
            Word.InlineShapes shapes = null;
            Word.InlineShape shape = null;
            try
            {
                shapes = document.InlineShapes;
                object linkToFile = false;
                object saveWithDocument = true;
                object anchorObject = anchorRange;
                shape = shapes.AddPicture(
                    filePath,
                    ref linkToFile,
                    ref saveWithDocument,
                    ref anchorObject);
                if (shape == null)
                {
                    throw new OfficeComException(OfficeErrorCode.INTERNAL_ERROR, "InlineShapes.AddPicture 返回 null。");
                }
                shape.Width = (float)width;
                shape.Height = (float)height;
                if (alternativeText != null)
                {
                    shape.AlternativeText = alternativeText;
                }
                inlineIndex = WordComUtil.FindInlineShapeIndex(document, shape);
                return shape;
            }
            finally
            {
                ComHelper.Release(shapes);
            }
        }

        private static byte[] CopyFloatingToEmf(
            Word.Application application,
            Word.Shape floatingShape,
            out string error)
        {
            error = string.Empty;
            Word.Document previous = null;
            Word.Document scratch = null;
            Word.Selection selection = null;
            Word.Range content = null;
            Word.InlineShapes inlineShapes = null;
            Word.InlineShape pasted = null;
            Word.Range range = null;
            try
            {
                try
                {
                    previous = application.ActiveDocument;
                }
                catch
                {
                    previous = null;
                }
                object missing = Type.Missing;
                floatingShape.Select(ref missing);
                selection = application.Selection;
                if (selection == null)
                {
                    error = "复制浮动图片失败：没有有效选区。";
                    return null;
                }
                selection.CopyAsPicture();
                scratch = WordTemplateFactory.CreateDocument(application);
                content = scratch.Content;
                content.Paste();
                inlineShapes = scratch.InlineShapes;
                if (inlineShapes == null || inlineShapes.Count == 0)
                {
                    Word.Shapes scratchShapes = null;
                    Word.Shape scratchShape = null;
                    Word.InlineShape converted = null;
                    Word.Range convertedRange = null;
                    try
                    {
                        scratchShapes = scratch.Shapes;
                        if (scratchShapes == null || scratchShapes.Count == 0)
                        {
                            error = "复制浮动图片到临时文档后未生成 InlineShape/Shape。";
                            return null;
                        }
                        scratchShape = scratchShapes[1];
                        converted = scratchShape.ConvertToInlineShape();
                        if (converted == null)
                        {
                            error = "临时文档中的浮动对象无法转为 InlineShape。";
                            return null;
                        }
                        convertedRange = converted.Range;
                        object convertedBits = convertedRange.EnhMetaFileBits;
                        return convertedBits as byte[];
                    }
                    finally
                    {
                        ComHelper.Release(convertedRange, converted, scratchShape, scratchShapes);
                    }
                }
                pasted = inlineShapes[1];
                range = pasted.Range;
                object bits = range.EnhMetaFileBits;
                return bits as byte[];
            }
            catch (Exception ex)
            {
                error = "导出浮动图片 EMF 失败：" + ex.Message;
                return null;
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
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[word] 关闭图片临时文档失败：" + ex.Message);
                    }
                }
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
                ComHelper.Release(range, pasted, inlineShapes, content, selection, scratch, previous);
            }
        }
    }
}
