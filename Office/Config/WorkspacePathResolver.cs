using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json.Linq;

namespace Office.Config
{
    /// <summary>
    /// 工作目录与文件参数解析器。
    ///
    /// 规则（用户约定）：
    /// 1) 工具的文件参数接受纯文件名、工作目录内的安全相对子路径或本机绝对路径；
    /// 2) 纯文件名与相对子路径从 config.json 的 workspace.root 解析，绝对路径原样使用；
    /// 3) 文档类参数（file / files 等）按 appType 校验扩展名，防止把错误类型文件投给执行器；
    /// 4) HTTP 上传/下载只允许工作目录内的纯文件名，禁止路径穿越。
    /// </summary>
    public static class WorkspacePathResolver
    {
        private static readonly object SyncRoot = new object();
        private static string _cachedRoot;

        private static readonly Dictionary<string, string[]> DocumentExtensions =
            new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
            {
                { "excel", new string[] { ".xlsx", ".xlsm", ".xlsb", ".xls", ".csv" } },
                { "word", new string[] { ".docx", ".docm", ".doc", ".rtf", ".txt", ".pdf" } },
                { "ppt", new string[] { ".pptx", ".pptm", ".ppsx", ".pps", ".potx", ".potm", ".ppt" } }
            };

        private static readonly HashSet<string> GlobalUploadExtensions =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                ".xlsx", ".xlsm", ".xlsb", ".xls", ".csv",
                ".docx", ".docm", ".doc", ".rtf", ".txt", ".pdf", ".md",
                ".pptx", ".pptm", ".ppsx", ".pps", ".potx", ".potm", ".ppt", ".thmx",
                ".json", ".xml",
                ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf"
            };

        private static readonly HashSet<string> SinglePathParameters =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "file", "filePath", "path", "inputPath", "outputPath", "outputDirectory",
                "outputDir", "saveAs", "sourcePath", "originalPath", "revisedPath",
                "directory", "targetPath", "savePath", "outputFile", "inputFile",
                "templatePath", "folder", "folderPath", "sourceFile", "targetFile", "themeFile"
            };

        private static readonly HashSet<string> ArrayPathParameters =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "files", "filePaths", "sourcePaths", "inputFiles", "outputFiles"
            };

        /// <summary>这些参数代表“将被 Office 打开的文档”，必须校验 appType 对应扩展名。</summary>
        private static readonly HashSet<string> DocumentPathParameters =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "file", "files", "filePaths", "sourcePaths", "inputFiles",
                "sourcePath", "originalPath", "revisedPath"
            };

        /// <summary>当前工作目录绝对路径；配置缺失时使用 exe 同目录的 workspace。</summary>
        public static string WorkspaceRoot
        {
            get
            {
                string root;
                string error;
                if (!TryGetWorkspaceRoot(out root, out error))
                {
                    return string.Empty;
                }
                return root;
            }
        }

        /// <summary>工作目录不存在时自动创建。</summary>
        public static bool TryGetWorkspaceRoot(out string root, out string error)
        {
            root = string.Empty;
            error = string.Empty;

            lock (SyncRoot)
            {
                if (!string.IsNullOrWhiteSpace(_cachedRoot))
                {
                    root = _cachedRoot;
                    return true;
                }

                string configured = AppConfig.Current.WorkspaceRoot;
                if (string.IsNullOrWhiteSpace(configured))
                {
                    configured = "workspace";
                }

                try
                {
                    string expanded = Environment.ExpandEnvironmentVariables(configured.Trim().Trim('"'));
                    string fullPath = Path.IsPathRooted(expanded)
                        ? expanded
                        : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, expanded);
                    fullPath = Path.GetFullPath(fullPath);
                    if (!Directory.Exists(fullPath))
                    {
                        Directory.CreateDirectory(fullPath);
                    }
                    _cachedRoot = fullPath.TrimEnd('\\', '/');
                    root = _cachedRoot;
                    return true;
                }
                catch (Exception ex)
                {
                    error = "工作目录不可用（workspace.root=" + configured + "）：" + ex.Message;
                    return false;
                }
            }
        }

        /// <summary>判断绝对路径是否位于工作目录内（含工作目录本身）。</summary>
        public static bool IsInsideWorkspace(string fullPath)
        {
            if (string.IsNullOrWhiteSpace(fullPath))
            {
                return false;
            }
            string root;
            string rootError;
            if (!TryGetWorkspaceRoot(out root, out rootError))
            {
                return false;
            }
            try
            {
                string normalizedRoot = root.TrimEnd('\\', '/');
                string candidate = Path.GetFullPath(fullPath).TrimEnd('\\', '/');
                if (string.Equals(candidate, normalizedRoot, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
                return candidate.StartsWith(
                    normalizedRoot + Path.DirectorySeparatorChar,
                    StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 创建输出目录。自动创建只允许发生在 workspace.root 内；
        /// 工作目录外的绝对路径必须由用户预先创建，避免工具在任意位置建目录。
        /// </summary>
        public static bool TryEnsureOutputDirectory(string directory, out string error)
        {
            error = string.Empty;
            if (string.IsNullOrWhiteSpace(directory))
            {
                error = "输出目录不能为空。";
                return false;
            }

            try
            {
                string fullDirectory = Path.GetFullPath(directory);
                if (Directory.Exists(fullDirectory))
                {
                    return true;
                }
                if (!IsInsideWorkspace(fullDirectory))
                {
                    error = "仅允许在工作目录（workspace.root=" + WorkspaceRoot +
                            "）内自动创建目录；工作目录外的目录请先手动创建：" + fullDirectory;
                    return false;
                }
                Directory.CreateDirectory(fullDirectory);
                return true;
            }
            catch (Exception ex)
            {
                error = "无法创建输出目录：" + directory + "；" + ex.Message;
                return false;
            }
        }

        /// <summary>清空缓存，便于测试或 config 热更新后重新解析。</summary>
        public static void ResetCache()
        {
            lock (SyncRoot)
            {
                _cachedRoot = null;
            }
        }

        /// <summary>
        /// 解析工具文件参数：纯文件名或安全相对子路径 → 工作目录；绝对路径 → 原样使用。
        /// 相对子路径禁止 ..、盘符与非法文件名段；写入类工具负责创建缺失的父目录。
        /// </summary>
        public static bool TryResolve(string rawPath, out string fullPath, out string error)
        {
            fullPath = string.Empty;
            error = string.Empty;

            string trimmed = (rawPath ?? string.Empty).Trim().Trim('"');
            if (trimmed.Length == 0)
            {
                error = "文件路径不能为空。";
                return false;
            }

            try
            {
                string expanded = Environment.ExpandEnvironmentVariables(trimmed);
                if (IsBareFileName(expanded))
                {
                    string root;
                    string rootError;
                    if (!TryGetWorkspaceRoot(out root, out rootError))
                    {
                        error = rootError;
                        return false;
                    }
                    if (!IsValidFileName(expanded))
                    {
                        error = "文件名称包含无效字符：" + expanded;
                        return false;
                    }
                    fullPath = Path.GetFullPath(Path.Combine(root, expanded));
                    return true;
                }

                if (!Path.IsPathRooted(expanded))
                {
                    string relativeRoot;
                    string relativeRootError;
                    if (!TryGetWorkspaceRoot(out relativeRoot, out relativeRootError))
                    {
                        error = relativeRootError;
                        return false;
                    }
                    if (!IsSafeRelativePath(expanded))
                    {
                        error = "相对路径只能位于工作目录内，且不能包含 ..、盘符或非法文件名段：「" +
                                trimmed + "」。";
                        return false;
                    }
                    string candidate = Path.GetFullPath(Path.Combine(relativeRoot, expanded));
                    string normalizedRoot = relativeRoot.TrimEnd('\\', '/') + Path.DirectorySeparatorChar;
                    if (!candidate.StartsWith(normalizedRoot, StringComparison.OrdinalIgnoreCase))
                    {
                        error = "路径越出工作目录：「" + trimmed + "」。";
                        return false;
                    }
                    fullPath = candidate;
                    return true;
                }

                // 拒绝 "C:foo.xlsx" 这类驱动器相对路径：它不是完整的本机绝对路径。
                string pathRoot = Path.GetPathRoot(expanded);
                if (pathRoot != null && pathRoot.Length == 2 && pathRoot[1] == ':' &&
                    expanded.Length > 2 && expanded[2] != '\\' && expanded[2] != '/')
                {
                    error = "只支持完整的本机绝对路径（如 D:\\docs\\a.xlsx），收到「" + trimmed + "」。";
                    return false;
                }

                fullPath = Path.GetFullPath(expanded);
                return true;
            }
            catch (Exception ex)
            {
                error = "文件路径无效：「" + trimmed + "」；" + ex.Message;
                return false;
            }
        }

        /// <summary>解析文档参数并校验 appType 对应的扩展名。</summary>
        public static bool TryResolveDocument(
            string appType,
            string rawPath,
            out string fullPath,
            out string error)
        {
            if (!TryResolve(rawPath, out fullPath, out error))
            {
                return false;
            }

            if (!IsAllowedDocumentExtension(appType, fullPath))
            {
                error = "文件扩展名「" + Path.GetExtension(fullPath) +
                        "」不符合 " + DescribeAppType(appType) + " 文档要求（支持：" +
                        string.Join(" / ", GetDocumentExtensions(appType)) + "）：" + fullPath;
                fullPath = string.Empty;
                return false;
            }
            return true;
        }

        /// <summary>
        /// HTTP 上传/下载专用：只允许工作目录内的纯文件名，禁止任何目录段与路径穿越。
        /// </summary>
        public static bool TryResolveWorkspaceFileName(
            string rawName,
            out string fullPath,
            out string error)
        {
            fullPath = string.Empty;
            error = string.Empty;

            string trimmed = (rawName ?? string.Empty).Trim().Trim('"');
            if (trimmed.Length == 0)
            {
                error = "文件名称不能为空。";
                return false;
            }
            if (!IsBareFileName(trimmed) || !IsValidFileName(trimmed))
            {
                error = "上传/下载只接受工作目录内的纯文件名，禁止目录或路径穿越：「" + trimmed + "」。";
                return false;
            }

            string root;
            string rootError;
            if (!TryGetWorkspaceRoot(out root, out rootError))
            {
                error = rootError;
                return false;
            }

            try
            {
                fullPath = Path.GetFullPath(Path.Combine(root, trimmed));
                string normalizedRoot = root.TrimEnd('\\', '/') + Path.DirectorySeparatorChar;
                if (!fullPath.StartsWith(normalizedRoot, StringComparison.OrdinalIgnoreCase))
                {
                    error = "文件名称越出工作目录：" + trimmed;
                    fullPath = string.Empty;
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                error = "文件名称无效：" + ex.Message;
                return false;
            }
        }

        public static bool IsAllowedDocumentExtension(string appType, string filePath)
        {
            string extension = GetExtension(filePath);
            if (extension.Length == 0)
            {
                return false;
            }

            return IsAllowedUploadExtension(appType, filePath);
        }

        /// <summary>
        /// 路由上传/下载扩展名白名单：指定 appType 时只允许该应用文档类型；
        /// appType 为空时使用通用安全清单（文档/数据/图片），不包含可执行文件。
        /// </summary>
        public static bool IsAllowedUploadExtension(string appType, string fileName)
        {
            string extension = GetExtension(fileName);
            if (extension.Length == 0)
            {
                return false;
            }

            string normalizedApp = NormalizeAppType(appType);
            if (normalizedApp.Length > 0 && normalizedApp != "common")
            {
                string[] allowed;
                if (!DocumentExtensions.TryGetValue(normalizedApp, out allowed))
                {
                    return false;
                }
                return ContainsExtension(allowed, extension);
            }

            if (GlobalUploadExtensions.Contains(extension))
            {
                return true;
            }

            // common/未知路由：允许所有 office 文档扩展名的并集。
            foreach (KeyValuePair<string, string[]> pair in DocumentExtensions)
            {
                if (ContainsExtension(pair.Value, extension))
                {
                    return true;
                }
            }
            return false;
        }

        public static string[] GetDocumentExtensions(string appType)
        {
            string normalizedApp = NormalizeAppType(appType);
            string[] allowed;
            if (normalizedApp.Length > 0 && DocumentExtensions.TryGetValue(normalizedApp, out allowed))
            {
                return allowed;
            }

            List<string> merged = new List<string>();
            foreach (KeyValuePair<string, string[]> pair in DocumentExtensions)
            {
                foreach (string extension in pair.Value)
                {
                    if (!ContainsExtension(merged.ToArray(), extension))
                    {
                        merged.Add(extension);
                    }
                }
            }
            return merged.ToArray();
        }

        /// <summary>
        /// 克隆并规范化工具参数中的所有路径字段。纯文件名与安全相对子路径会解析到工作目录，
        /// file/files 等文档字段同时校验扩展名。解析失败返回 false 与中文错误。
        /// </summary>
        public static bool TryNormalizeArguments(
            string appType,
            JObject arguments,
            out JObject normalized,
            out string error)
        {
            normalized = arguments == null ? new JObject() : (JObject)arguments.DeepClone();
            error = string.Empty;
            if (normalized == null)
            {
                normalized = new JObject();
                return true;
            }

            foreach (JProperty property in normalized.Properties())
            {
                string name = property.Name;
                bool enforceDocument = DocumentPathParameters.Contains(name);

                JArray array = property.Value as JArray;
                if (array != null && ArrayPathParameters.Contains(name))
                {
                    for (int index = 0; index < array.Count; index++)
                    {
                        JToken item = array[index];
                        if (item == null || item.Type != JTokenType.String)
                        {
                            continue;
                        }

                        string raw = item.Value<string>();
                        if (string.IsNullOrWhiteSpace(raw))
                        {
                            continue;
                        }

                        string resolved;
                        string pathError;
                        bool ok = enforceDocument
                            ? TryResolveDocument(appType, raw, out resolved, out pathError)
                            : TryResolve(raw, out resolved, out pathError);
                        if (!ok)
                        {
                            error = "参数 " + name + "[" + index.ToString() + "]" + "：" + pathError;
                            return false;
                        }
                        array[index] = new JValue(resolved);
                    }
                    continue;
                }

                if (property.Value == null || property.Value.Type != JTokenType.String)
                {
                    continue;
                }
                if (!SinglePathParameters.Contains(name))
                {
                    continue;
                }

                string rawValue = property.Value.Value<string>();
                if (string.IsNullOrWhiteSpace(rawValue))
                {
                    continue;
                }

                string fullPath;
                string resolveError;
                bool resolvedOk = enforceDocument
                    ? TryResolveDocument(appType, rawValue, out fullPath, out resolveError)
                    : TryResolve(rawValue, out fullPath, out resolveError);
                if (!resolvedOk)
                {
                    error = "参数 " + name + "：" + resolveError;
                    return false;
                }
                property.Value = new JValue(fullPath);
            }

            return true;
        }

        /// <summary>
        /// 安全相对子路径判定：允许 a、a\b、a/b；拒绝 ..、.、盘符、根路径与非法文件名段。
        /// 解引用前必须先通过本校验，再由调用方确认结果仍位于工作目录内。
        /// </summary>
        public static bool IsSafeRelativePath(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }
            string trimmed = value.Trim();
            if (Path.IsPathRooted(trimmed))
            {
                return false;
            }
            string normalized = trimmed.Replace('/', '\\');
            string[] segments = normalized.Split('\\');
            foreach (string segment in segments)
            {
                if (segment.Length == 0)
                {
                    continue;
                }
                if (segment == "." || segment == "..")
                {
                    return false;
                }
                if (segment.IndexOf(':') >= 0)
                {
                    return false;
                }
                if (!IsValidFileName(segment))
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>纯文件名判定：不含目录分隔符、不是根路径、不是 . 或 ..。</summary>
        public static bool IsBareFileName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }
            string trimmed = value.Trim();
            if (trimmed == "." || trimmed == "..")
            {
                return false;
            }
            if (trimmed.IndexOf('\\') >= 0 || trimmed.IndexOf('/') >= 0)
            {
                return false;
            }
            if (Path.IsPathRooted(trimmed))
            {
                return false;
            }
            try
            {
                return string.Equals(Path.GetFileName(trimmed), trimmed, StringComparison.Ordinal);
            }
            catch
            {
                return false;
            }
        }

        private static bool IsValidFileName(string value)
        {
            char[] invalid = Path.GetInvalidFileNameChars();
            foreach (char character in invalid)
            {
                if (value.IndexOf(character) >= 0)
                {
                    return false;
                }
            }
            return true;
        }

        private static string NormalizeAppType(string appType)
        {
            string normalized = (appType ?? string.Empty).Trim().ToLowerInvariant();
            if (normalized == "xlsx")
            {
                return "excel";
            }
            if (normalized == "docx" || normalized == "doc")
            {
                return "word";
            }
            return normalized;
        }

        private static string DescribeAppType(string appType)
        {
            string normalized = NormalizeAppType(appType);
            if (normalized == "excel")
            {
                return "Excel";
            }
            if (normalized == "word")
            {
                return "Word";
            }
            if (normalized == "ppt")
            {
                return "PowerPoint";
            }
            return "Office";
        }

        private static string GetExtension(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }
            try
            {
                return Path.GetExtension(path).ToLowerInvariant();
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool ContainsExtension(string[] extensions, string extension)
        {
            if (extensions == null || string.IsNullOrWhiteSpace(extension))
            {
                return false;
            }
            foreach (string item in extensions)
            {
                if (string.Equals(item, extension, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
