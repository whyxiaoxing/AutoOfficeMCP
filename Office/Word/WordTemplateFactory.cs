using System;
using System.IO;
using Office.Config;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    /// <summary>
    /// T11：Word 新文档只使用显式文档模板（.dotx），不依赖/不修改 Normal.dotm。
    /// Word 启动仍会加载 Normal.dotm（COM 无法卸载全局模板），但本类保证业务文档从专用模板创建。
    /// </summary>
    public static class WordTemplateFactory
    {
        public static string TemplatePath
        {
            get
            {
                string configured = AppConfig.Current.WordTemplatePath;
                if (string.IsNullOrWhiteSpace(configured))
                {
                    return string.Empty;
                }

                string path = configured.Trim();
                if (!Path.IsPathRooted(path))
                {
                    path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path);
                }
                return Path.GetFullPath(path);
            }
        }

        public static bool TemplateExists
        {
            get
            {
                string path = TemplatePath;
                return !string.IsNullOrWhiteSpace(path) && File.Exists(path);
            }
        }

        /// <summary>
        /// 从专用模板创建文档；模板缺失时抛 FILE_NOT_FOUND，避免静默退回 Normal.dotm。
        /// </summary>
        public static Word.Document CreateDocument(Word.Application application)
        {
            if (application == null)
            {
                throw new OfficeComException(OfficeErrorCode.APP_NOT_RUNNING, "Word 应用不可用。");
            }

            string templatePath = TemplatePath;
            if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
            {
                throw new OfficeComException(
                    OfficeErrorCode.FILE_NOT_FOUND,
                    "缺少 Word 文档模板：" + (string.IsNullOrWhiteSpace(templatePath)
                        ? "(security.wordTemplatePath 为空)"
                        : templatePath) +
                    "。T11 要求使用专用 .dotx 模板，不退回 Normal.dotm。");
            }

            object template = templatePath;
            object missing = Type.Missing;
            Word.Document document = application.Documents.Add(
                ref template,
                ref missing,
                ref missing,
                ref missing);
            if (document == null)
            {
                throw new OfficeComException(
                    OfficeErrorCode.INTERNAL_ERROR,
                    "Word 未返回从模板创建的新文档：" + templatePath);
            }
            return document;
        }
    }
}
