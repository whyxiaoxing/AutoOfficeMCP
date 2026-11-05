using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json.Linq;

namespace Office.Config
{
    /// <summary>
    /// exe 同目录 config.json 的强类型只读快照（缺失字段全部回退默认值）。
    /// 所有执行器共享同一份配置；加载失败时记录 stderr 并继续用默认值，保证 server 可启动。
    /// </summary>
    public sealed class AppConfig
    {
        private static readonly object SyncRoot = new object();
        private static AppConfig _current;

        private readonly JObject _root;

        private AppConfig(JObject root)
        {
            _root = root;
        }

        /// <summary>当前配置单例（首次访问时加载）。</summary>
        public static AppConfig Current
        {
            get
            {
                lock (SyncRoot)
                {
                    if (_current == null)
                    {
                        _current = Load();
                    }
                    return _current;
                }
            }
        }

        /// <summary>config.json 的绝对路径：exe 同目录。</summary>
        public static string ConfigPath
        {
            get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json"); }
        }

        /// <summary>重新从磁盘加载（测试/热更新用）。</summary>
        public static AppConfig Reload()
        {
            lock (SyncRoot)
            {
                _current = Load();
                return _current;
            }
        }

        private static AppConfig Load()
        {
            JObject root = new JObject();
            string path = ConfigPath;
            try
            {
                if (File.Exists(path))
                {
                    string text = File.ReadAllText(path);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        root = JObject.Parse(text);
                    }
                }
                else
                {
                    Console.Error.WriteLine("[config] 未找到 " + path + "，使用默认配置。");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[config] 读取失败，使用默认配置：" + ex.Message);
                root = new JObject();
            }
            return new AppConfig(root);
        }

        private JToken ReadPath(string dottedPath)
        {
            string[] parts = dottedPath.Split('.');
            JToken current = _root;
            for (int i = 0; i < parts.Length; i++)
            {
                if (current == null || current.Type != JTokenType.Object)
                {
                    return null;
                }
                current = current[parts[i]];
            }
            return current;
        }

        private bool GetBool(string dottedPath, bool defaultValue)
        {
            JToken token = ReadPath(dottedPath);
            if (token == null || token.Type == JTokenType.Null)
            {
                return defaultValue;
            }
            try
            {
                return token.Value<bool>();
            }
            catch
            {
                return defaultValue;
            }
        }

        private int GetInt(string dottedPath, int defaultValue)
        {
            JToken token = ReadPath(dottedPath);
            if (token == null || token.Type == JTokenType.Null)
            {
                return defaultValue;
            }
            try
            {
                return token.Value<int>();
            }
            catch
            {
                return defaultValue;
            }
        }

        private string GetString(string dottedPath, string defaultValue)
        {
            JToken token = ReadPath(dottedPath);
            if (token == null || token.Type == JTokenType.Null)
            {
                return defaultValue;
            }
            try
            {
                return token.Value<string>();
            }
            catch
            {
                return defaultValue;
            }
        }

        public bool ExcelAutoLaunch
        {
            get { return GetBool("office.excel.autoLaunch", true); }
        }

        public bool ExcelVisible
        {
            get { return GetBool("office.excel.visible", true); }
        }

        public bool WordAutoLaunch
        {
            get { return GetBool("office.word.autoLaunch", true); }
        }

        public bool WordVisible
        {
            get { return GetBool("office.word.visible", false); }
        }

        public bool PowerPointAutoLaunch
        {
            get { return GetBool("office.ppt.autoLaunch", true); }
        }

        public bool PowerPointVisible
        {
            get { return GetBool("office.ppt.visible", true); }
        }

        public int SessionTimeoutSec
        {
            get { return GetInt("session.timeoutSec", 120); }
        }

        public bool AllowVba
        {
            get { return GetBool("security.allowVba", false); }
        }

        public bool AllowDeleteSheet
        {
            get { return GetBool("security.allowDeleteSheet", false); }
        }

        public bool AllowCloseWithoutSave
        {
            get { return GetBool("security.allowCloseWithoutSave", false); }
        }

        /// <summary>tools.disabled 名单（大小写不敏感）。</summary>
        public HashSet<string> DisabledTools
        {
            get
            {
                HashSet<string> set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                JToken token = ReadPath("tools.disabled");
                if (token != null && token.Type == JTokenType.Array)
                {
                    foreach (JToken item in token.Children())
                    {
                        if (item.Type == JTokenType.String)
                        {
                            set.Add(item.Value<string>());
                        }
                    }
                }
                return set;
            }
        }

        public int ExcelMaxConcurrency
        {
            get { return GetInt("parallel.excel.maxConcurrency", 2); }
        }

        public int WordMaxConcurrency
        {
            get { return GetInt("parallel.word.maxConcurrency", 1); }
        }

        /// <summary>PPT 并行作业槽位数；PowerPoint 单实例，实际并发受 worker/会话隔离限制。</summary>
        public int PowerPointMaxConcurrency
        {
            get { return Math.Max(1, GetInt("parallel.ppt.maxConcurrency", 2)); }
        }

        public int ParallelItemTimeoutSec
        {
            get { return GetInt("parallel.itemTimeoutSec", 300); }
        }

        public int ParallelMaxQueue
        {
            get { return GetInt("parallel.maxQueue", 200); }
        }

        /// <summary>幂等结果缓存秒数；0 表示只做在飞去重，不缓存完成结果。</summary>
        public int ParallelIdempotencyTtlSec
        {
            get { return Math.Max(0, GetInt("parallel.idempotencyTtlSec", 300)); }
        }

        public string TransportBind
        {
            get { return GetString("transport.bind", "127.0.0.1"); }
        }

        public int TransportPort
        {
            get { return GetInt("transport.port", 8765); }
        }

        /// <summary>多用户 token → userId 映射；按用户路由 worker。</summary>
        public Dictionary<string, string> TransportTokens
        {
            get
            {
                Dictionary<string, string> map = new Dictionary<string, string>(StringComparer.Ordinal);
                JToken token = ReadPath("transport.tokens");
                JObject root = token as JObject;
                if (root == null)
                {
                    return map;
                }
                foreach (JProperty property in root.Properties())
                {
                    if (!string.IsNullOrWhiteSpace(property.Name)
                        && property.Value != null
                        && property.Value.Type == JTokenType.String
                        && !string.IsNullOrWhiteSpace(property.Value.Value<string>()))
                    {
                        map[property.Name] = property.Value.Value<string>().Trim();
                    }
                }
                return map;
            }
        }

        public int TransportMaxBodyBytes
        {
            get { return GetInt("transport.maxBodyBytes", 8388608); }
        }

        public int ServerMaxInflightPerSession
        {
            get { return GetInt("server.maxInflightPerSession", 1); }
        }

        public int WorkerMaxWarmWorkers
        {
            get { return GetInt("worker.maxWarmWorkers", 8); }
        }

        public int WorkerIdleTimeoutSec
        {
            get { return GetInt("worker.idleTimeoutSec", 600); }
        }

        public int WorkerStartupTimeoutSec
        {
            get { return GetInt("worker.startupTimeoutSec", 30); }
        }

        /// <summary>并发模式：balanced（默认）/ throughput；未知值回退 balanced。</summary>
        public string WorkerConcurrencyMode
        {
            get
            {
                string value = GetString("worker.concurrencyMode", "balanced").Trim().ToLowerInvariant();
                return string.Equals(value, "throughput", StringComparison.Ordinal) ? value : "balanced";
            }
        }

        public int WorkerMaxWorkersPerUser
        {
            get { return GetInt("worker.maxWorkersPerUser", 1); }
        }

        public int WorkerMemoryLimitMb
        {
            get { return GetInt("worker.memoryLimitMb", 0); }
        }

        public int WorkerJobTimeoutSec
        {
            get { return GetInt("worker.jobTimeoutSec", 300); }
        }

        public string WorkerPipeNamePrefix
        {
            get { return GetString("worker.pipeNamePrefix", "AutoOfficeMCP.Worker"); }
        }

        /// <summary>worker → Gateway 心跳间隔（秒）；小于 2 按 2 处理。</summary>
        public int WorkerHeartbeatIntervalSec
        {
            get { return Math.Max(2, GetInt("worker.heartbeatIntervalSec", 10)); }
        }

        /// <summary>超过该秒数未收到心跳视为 worker 失联；至少 6 秒。</summary>
        public int WorkerHeartbeatTimeoutSec
        {
            get
            {
                int timeout = GetInt("worker.heartbeatTimeoutSec", 30);
                int interval = WorkerHeartbeatIntervalSec;
                return Math.Max(Math.Max(6, interval * 2), timeout);
            }
        }

        /// <summary>保存后 SHA256 校验的最大文件大小（MB）；0 表示只用长度 + mtime。</summary>
        public int WorkerSaveVerifyMaxHashMb
        {
            get { return Math.Max(0, GetInt("worker.saveVerifyMaxHashMb", 64)); }
        }

        /// <summary>跨进程锁文件空闲 TTL（秒）；仅清理确认未被持有的陈旧锁文件。</summary>
        public int WorkerFileLockTtlSec
        {
            get { return Math.Max(60, GetInt("worker.fileLockTtlSec", 3600)); }
        }

        /// <summary>文件锁提供者：file（默认，锁文件）/ sql（SQL Server 分布式锁）。</summary>
        public string WorkerFileLockProvider
        {
            get
            {
                string value = GetString("worker.fileLockProvider", "file").Trim().ToLowerInvariant();
                return string.Equals(value, "sql", StringComparison.Ordinal) ? "sql" : "file";
            }
        }

        /// <summary>锁文件根目录；留空使用 %TEMP%\AutoOfficeMCP\locks，可填 UNC 共享路径实现多机互斥。</summary>
        public string WorkerFileLockRoot
        {
            get { return GetString("worker.fileLockRoot", string.Empty).Trim(); }
        }

        /// <summary>SQL Server 分布式锁连接字符串；仅 fileLockProvider=sql 时使用。</summary>
        public string WorkerSqlLockConnectionString
        {
            get { return GetString("worker.sqlLockConnectionString", string.Empty).Trim(); }
        }

        /// <summary>是否允许使用剪贴板的工具；默认 false（T11 多进程安全）。</summary>
        public bool AllowClipboard
        {
            get { return GetBool("security.allowClipboard", false); }
        }

        /// <summary>worker 启动时是否断开 COM/AddIns 加载项。</summary>
        public bool DisableAddins
        {
            get { return GetBool("security.disableAddins", true); }
        }

        /// <summary>Word 新文档使用的文档模板；留空则退回 Normal.dotm（不推荐）。</summary>
        public string WordTemplatePath
        {
            get { return GetString("security.wordTemplatePath", "Templates\\AutoOfficeMCP.dotx"); }
        }

        /// <summary>
        /// 工作目录：文件参数纯文件名与 HTTP 上传/下载的根目录。
        /// 相对路径以 exe 同目录为基准；缺省为 workspace。
        /// </summary>
        public string WorkspaceRoot
        {
            get { return GetString("workspace.root", "workspace").Trim(); }
        }

        /// <summary>HTTP 单次上传大小上限（字节）；默认 256 MB，最小 1 KB。</summary>
        public long WorkspaceMaxUploadBytes
        {
            get
            {
                long value = GetLong("workspace.maxUploadBytes", 268435456L);
                return Math.Max(1024L, value);
            }
        }

        private long GetLong(string dottedPath, long defaultValue)
        {
            JToken token = ReadPath(dottedPath);
            if (token == null || token.Type == JTokenType.Null)
            {
                return defaultValue;
            }
            try
            {
                return token.Value<long>();
            }
            catch
            {
                return defaultValue;
            }
        }
    }
}
