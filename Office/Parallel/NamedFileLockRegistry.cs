using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace Office.Parallel
{
    /// <summary>
    /// T11：同机跨进程文件锁。使用锁文件的 FileShare.None 句柄，不依赖互斥量线程亲和性；
    /// 进程崩溃时 OS 自动关闭句柄，下一个 worker 可以重新获取。锁文件本身可以残留，不代表锁被持有。
    /// `~$` owner file 只做辅助提示，不作为互斥依据。
    /// </summary>
    internal sealed class NamedFileLockRegistry : IFileLockProvider
    {
        private sealed class LockEntry
        {
            public readonly string DisplayPath;
            public int Users;
            public int ActiveLeases;

            public LockEntry(string displayPath)
            {
                DisplayPath = displayPath ?? string.Empty;
            }
        }

        private readonly object _sync;
        private readonly string _lockRoot;
        private readonly ConcurrentDictionary<string, LockEntry> _entries;
        private readonly int _lockTtlSec;
        private DateTime _lastCleanupUtc;

        public NamedFileLockRegistry()
        {
            _sync = new object();
            _entries = new ConcurrentDictionary<string, LockEntry>(StringComparer.OrdinalIgnoreCase);
            string configuredRoot = Office.Config.AppConfig.Current.WorkerFileLockRoot;
            _lockRoot = string.IsNullOrWhiteSpace(configuredRoot)
                ? Path.Combine(Path.GetTempPath(), "AutoOfficeMCP", "locks")
                : configuredRoot;
            _lockTtlSec = Office.Config.AppConfig.Current.WorkerFileLockTtlSec;
            _lastCleanupUtc = DateTime.MinValue;
            Directory.CreateDirectory(_lockRoot);
        }

        public async Task<IDisposable> AcquireAsync(
            string normalizedPath,
            string displayPath,
            TimeSpan timeout)
        {
            if (string.IsNullOrWhiteSpace(normalizedPath) || timeout <= TimeSpan.Zero)
            {
                return null;
            }

            CleanupStaleLockFilesIfDue();
            LockEntry entry;
            lock (_sync)
            {
                entry = _entries.GetOrAdd(normalizedPath, delegate { return new LockEntry(displayPath); });
                entry.Users++;
            }

            string lockPath = Path.Combine(_lockRoot, Hash(normalizedPath) + ".lock");
            DateTime deadline = DateTime.UtcNow.Add(timeout);
            try
            {
                while (true)
                {
                    try
                    {
                        FileStream stream = new FileStream(
                            lockPath,
                            FileMode.OpenOrCreate,
                            FileAccess.ReadWrite,
                            FileShare.None,
                            1,
                            FileOptions.None);
                        OfficeOwnerFile.LogIfPresent(displayPath);
                        lock (_sync)
                        {
                            entry.ActiveLeases++;
                        }
                        return new Lease(this, normalizedPath, entry, stream);
                    }
                    catch (IOException)
                    {
                        // 另一个 worker/进程持有锁；等待重试。
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        Console.Error.WriteLine("[file-lock] 打开锁文件失败：" + lockPath + "；" + ex.Message);
                        ReleaseReference(normalizedPath, entry);
                        return null;
                    }

                    if (DateTime.UtcNow >= deadline)
                    {
                        ReleaseReference(normalizedPath, entry);
                        return null;
                    }
                    await Task.Delay(50).ConfigureAwait(false);
                }
            }
            catch
            {
                ReleaseReference(normalizedPath, entry);
                throw;
            }
        }

        private void ReleaseLease(string normalizedPath, LockEntry entry, FileStream stream)
        {
            try
            {
                if (stream != null)
                {
                    stream.Dispose();
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[file-lock] 释放锁文件失败：" + normalizedPath + "；" + ex.Message);
            }
            lock (_sync)
            {
                entry.ActiveLeases--;
            }
            ReleaseReference(normalizedPath, entry);
        }

        private void ReleaseReference(string normalizedPath, LockEntry entry)
        {
            lock (_sync)
            {
                entry.Users--;
                if (entry.Users > 0 || entry.ActiveLeases > 0)
                {
                    return;
                }
                LockEntry removed;
                _entries.TryRemove(normalizedPath, out removed);
            }
        }

        public JArray DescribeLocks()
        {
            JArray result = new JArray();
            lock (_sync)
            {
                foreach (KeyValuePair<string, LockEntry> pair in _entries)
                {
                    JObject item = new JObject();
                    item["path"] = pair.Value.DisplayPath;
                    item["normalizedPath"] = pair.Key;
                    item["users"] = pair.Value.Users;
                    item["held"] = pair.Value.ActiveLeases > 0;
                    item["waiting"] = Math.Max(0, pair.Value.Users - pair.Value.ActiveLeases);
                    result.Add(item);
                }
            }
            return result;
        }

        public int Count
        {
            get
            {
                lock (_sync)
                {
                    return _entries.Count;
                }
            }
        }

        /// <summary>
        /// 清理超过 TTL 且确认未被持有的锁文件：以 FileShare.None 打开成功说明没有活动句柄。
        /// 只删除陈旧文件，不删除仍被任何进程持有的锁；扫描节流为每分钟一次。
        /// </summary>
        private void CleanupStaleLockFilesIfDue()
        {
            DateTime now = DateTime.UtcNow;
            if ((now - _lastCleanupUtc).TotalSeconds < 60)
            {
                return;
            }
            _lastCleanupUtc = now;
            DateTime cutoff = now.AddSeconds(-_lockTtlSec);
            string[] files;
            try
            {
                files = Directory.GetFiles(_lockRoot, "*.lock");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[file-lock] 枚举陈旧锁文件失败：" + ex.Message);
                return;
            }
            foreach (string lockPath in files)
            {
                try
                {
                    FileInfo info = new FileInfo(lockPath);
                    if (!info.Exists || info.LastWriteTimeUtc >= cutoff)
                    {
                        continue;
                    }
                    using (new FileStream(
                               lockPath,
                               FileMode.Open,
                               FileAccess.ReadWrite,
                               FileShare.None,
                               1,
                               FileOptions.None))
                    {
                        // 能独占打开说明没有活动锁句柄。
                    }
                    File.Delete(lockPath);
                }
                catch (IOException)
                {
                    // 仍被持有；保留。
                }
                catch (UnauthorizedAccessException)
                {
                    // 权限不足；保留。
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[file-lock] 清理陈旧锁文件失败：" + lockPath + "；" + ex.Message);
                }
            }
        }

        private static string Hash(string value)
        {
            using (SHA256 sha = SHA256.Create())
            {
                byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty));
                StringBuilder builder = new StringBuilder(bytes.Length * 2);
                for (int index = 0; index < bytes.Length; index++)
                {
                    builder.Append(bytes[index].ToString("x2"));
                }
                return builder.ToString();
            }
        }

        private sealed class Lease : IDisposable
        {
            private readonly NamedFileLockRegistry _owner;
            private readonly string _normalizedPath;
            private readonly LockEntry _entry;
            private readonly FileStream _stream;
            private int _disposed;

            public Lease(
                NamedFileLockRegistry owner,
                string normalizedPath,
                LockEntry entry,
                FileStream stream)
            {
                _owner = owner;
                _normalizedPath = normalizedPath;
                _entry = entry;
                _stream = stream;
            }

            public void Dispose()
            {
                if (Interlocked.Exchange(ref _disposed, 1) != 0)
                {
                    return;
                }
                _owner.ReleaseLease(_normalizedPath, _entry, _stream);
            }
        }
    }

    internal static class OfficeOwnerFile
    {
        public static void LogIfPresent(string filePath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath))
                {
                    return;
                }
                string directory = Path.GetDirectoryName(filePath);
                string name = Path.GetFileName(filePath);
                if (string.IsNullOrWhiteSpace(directory)
                    || string.IsNullOrWhiteSpace(name)
                    || name.StartsWith("~$", StringComparison.Ordinal))
                {
                    return;
                }
                string ownerPath = Path.Combine(directory, "~$" + name);
                if (File.Exists(ownerPath))
                {
                    Console.Error.WriteLine("[file-lock] 辅助提示：检测到 Office owner file " + ownerPath +
                                            "；实际互斥仍由锁文件句柄保证。");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[file-lock] 检查 ~$ 文件失败：" + ex.Message);
            }
        }
    }
}
