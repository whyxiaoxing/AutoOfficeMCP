using System;
using System.IO;
using System.Text;
using Newtonsoft.Json.Linq;

namespace Office
{
    /// <summary>
    /// worker 进程启动参数。由 Gateway 写入 worker 专用工作目录下的
    /// worker.settings.json，子进程从当前目录读取；不使用环境变量，也不使用命令行参数。
    /// </summary>
    public sealed class WorkerSettings
    {
        public string Role { get; set; }

        public string PipeName { get; set; }

        public string UserId { get; set; }

        public string AppType { get; set; }

        public WorkerSettings()
        {
            Role = string.Empty;
            PipeName = string.Empty;
            UserId = "default";
            AppType = string.Empty;
        }
    }

    /// <summary>worker.settings.json 的读写器；协议字段集中在本类，避免两处各写一套。</summary>
    public static class WorkerBootstrap
    {
        public const string SettingsFileName = "worker.settings.json";

        /// <summary>从指定目录读取 worker 设置；文件不存在或不是 worker 角色时返回 null。</summary>
        public static WorkerSettings TryReadFromDirectory(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory))
            {
                return null;
            }

            string path = Path.Combine(directory, SettingsFileName);
            if (!File.Exists(path))
            {
                return null;
            }

            try
            {
                string text = File.ReadAllText(path, Encoding.UTF8);
                if (string.IsNullOrWhiteSpace(text))
                {
                    return null;
                }
                JObject root = JObject.Parse(text);
                string role = ReadString(root, "role");
                if (!string.Equals(role, "worker", StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }

                WorkerSettings settings = new WorkerSettings();
                settings.Role = "worker";
                settings.PipeName = ReadString(root, "pipeName");
                settings.UserId = ReadString(root, "userId");
                settings.AppType = ReadString(root, "appType").ToLowerInvariant();
                if (string.IsNullOrWhiteSpace(settings.UserId))
                {
                    settings.UserId = "default";
                }
                return settings;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[worker-bootstrap] 读取 " + path + " 失败：" + ex.Message);
                return null;
            }
        }

        /// <summary>把 worker 设置写入指定目录；调用方负责创建目录与后续清理。</summary>
        public static void WriteToDirectory(string directory, WorkerSettings settings)
        {
            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new ArgumentException("worker 目录不能为空。", "directory");
            }
            if (settings == null)
            {
                throw new ArgumentNullException("settings");
            }

            JObject root = new JObject();
            root["role"] = "worker";
            root["pipeName"] = settings.PipeName ?? string.Empty;
            root["userId"] = settings.UserId ?? "default";
            root["appType"] = settings.AppType ?? string.Empty;
            string path = Path.Combine(directory, SettingsFileName);
            File.WriteAllText(path, root.ToString(), new UTF8Encoding(false));
        }

        private static string ReadString(JObject root, string propertyName)
        {
            JToken token = root[propertyName];
            if (token == null || token.Type == JTokenType.Null)
            {
                return string.Empty;
            }
            return token.Type == JTokenType.String
                ? (token.Value<string>() ?? string.Empty).Trim()
                : token.ToString().Trim();
        }
    }
}
