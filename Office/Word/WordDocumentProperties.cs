using System;
using System.Reflection;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    /// <summary>
    /// Word 内置文档属性读写。Office.csproj 不额外引用 office.dll，
    /// 因此对 BuiltInDocumentProperties 只使用 IDispatch 反射调用 Item/Value，避免引入 Microsoft.Office.Core 类型。
    /// </summary>
    internal static class WordDocumentProperties
    {
        private static readonly string[,] PropertyMap = new string[,]
        {
            { "title", "Title" },
            { "author", "Author" },
            { "keywords", "Keywords" },
            { "subject", "Subject" },
            { "comments", "Comments" }
        };

        public static JObject ReadAll(Word.Document document)
        {
            JObject result = new JObject();
            for (int index = 0; index < PropertyMap.GetLength(0); index++)
            {
                result[PropertyMap[index, 0]] = Read(document, PropertyMap[index, 1]);
            }
            return result;
        }

        public static string Read(Word.Document document, string builtInName)
        {
            object properties = null;
            object item = null;
            try
            {
                properties = document.BuiltInDocumentProperties;
                item = GetItem(properties, builtInName);
                if (item == null)
                {
                    return string.Empty;
                }
                object value = item.GetType().InvokeMember(
                    "Value", BindingFlags.GetProperty, null, item, null);
                return value == null ? string.Empty : Convert.ToString(value);
            }
            catch (Exception ex)
            {
                throw new OfficeComException(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "读取文档属性 " + builtInName + " 失败：" +
                    OfficeErrorMapper.DescribeException(ex, OfficeApp.Word));
            }
            finally
            {
                ComHelper.Release(item, properties);
            }
        }

        public static void Write(Word.Document document, string builtInName, string value)
        {
            object properties = null;
            object item = null;
            try
            {
                properties = document.BuiltInDocumentProperties;
                item = GetItem(properties, builtInName);
                if (item == null)
                {
                    throw new OfficeComException(
                        OfficeErrorCode.INVALID_PARAMS,
                        "Word 内置文档属性 " + builtInName + " 不存在，无法写入。");
                }
                item.GetType().InvokeMember(
                    "Value",
                    BindingFlags.SetProperty,
                    null,
                    item,
                    new object[] { value ?? string.Empty });
            }
            catch (OfficeComException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new OfficeComException(
                    OfficeErrorCode.DOC_PROTECTED,
                    "写入文档属性 " + builtInName + " 失败：" +
                    OfficeErrorMapper.DescribeException(ex, OfficeApp.Word) +
                    " 请检查文档是否受保护或只读。");
            }
            finally
            {
                ComHelper.Release(item, properties);
            }
        }

        public static string GetJsonName(int index)
        {
            return PropertyMap[index, 0];
        }

        public static string GetBuiltInName(int index)
        {
            return PropertyMap[index, 1];
        }

        public static int Count
        {
            get { return PropertyMap.GetLength(0); }
        }

        private static object GetItem(object properties, string builtInName)
        {
            if (properties == null)
            {
                return null;
            }
            return properties.GetType().InvokeMember(
                "Item",
                BindingFlags.GetProperty,
                null,
                properties,
                new object[] { builtInName });
        }
    }
}
