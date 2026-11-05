using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace Office.Parallel
{
    /// <summary>
    /// T10 P2.2：SQL Server 分布式文件锁（多机部署）。
    /// 使用 sp_getapplock(@LockOwner='Session')；连接保持打开即持有锁，连接断开/进程崩溃时 SQL Server 自动释放。
    /// 不创建表，只需要执行权限；锁资源名 = AutoOfficeMCP: + SHA256(规范化路径)。
    /// </summary>
    internal sealed class SqlFileLockProvider : IFileLockProvider
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
        private readonly string _connectionString;
        private readonly ConcurrentDictionary<string, LockEntry> _entries;

        public SqlFileLockProvider(string connectionString)
        {
            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new ArgumentException("SQL 文件锁连接字符串为空。", "connectionString");
            }
            _sync = new object();
            _connectionString = connectionString;
            _entries = new ConcurrentDictionary<string, LockEntry>(StringComparer.OrdinalIgnoreCase);
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

            LockEntry entry;
            lock (_sync)
            {
                entry = _entries.GetOrAdd(normalizedPath, delegate { return new LockEntry(displayPath); });
                entry.Users++;
            }

            SqlConnection connection = null;
            try
            {
                connection = new SqlConnection(_connectionString);
                await connection.OpenAsync().ConfigureAwait(false);

                int timeoutMs = (int)Math.Min(int.MaxValue, Math.Max(1000, timeout.TotalMilliseconds));
                using (SqlCommand command = new SqlCommand("sp_getapplock", connection))
                {
                    command.CommandType = System.Data.CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@Resource", "AutoOfficeMCP:" + Hash(normalizedPath));
                    command.Parameters.AddWithValue("@LockMode", "Exclusive");
                    command.Parameters.AddWithValue("@LockOwner", "Session");
                    command.Parameters.AddWithValue("@LockTimeout", timeoutMs);
                    SqlParameter result = command.Parameters.Add("@Result", System.Data.SqlDbType.Int);
                    result.Direction = System.Data.ParameterDirection.ReturnValue;
                    await command.ExecuteNonQueryAsync().ConfigureAwait(false);
                    int code = result.Value == null ? -999 : Convert.ToInt32(result.Value);
                    if (code < 0)
                    {
                        connection.Dispose();
                        ReleaseReference(normalizedPath, entry);
                        return null;
                    }
                }

                lock (_sync)
                {
                    entry.ActiveLeases++;
                }
                return new SqlLease(this, normalizedPath, entry, connection);
            }
            catch (Exception ex)
            {
                if (connection != null)
                {
                    try { connection.Dispose(); } catch { }
                }
                ReleaseReference(normalizedPath, entry);
                Console.Error.WriteLine("[sql-lock] 获取分布式锁失败：" + displayPath + "；" + ex.Message);
                return null;
            }
        }

        private void ReleaseLease(string normalizedPath, LockEntry entry, SqlConnection connection)
        {
            try
            {
                using (SqlCommand command = new SqlCommand("sp_releaseapplock", connection))
                {
                    command.CommandType = System.Data.CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@Resource", "AutoOfficeMCP:" + Hash(normalizedPath));
                    command.Parameters.AddWithValue("@LockOwner", "Session");
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[sql-lock] 释放分布式锁失败：" + normalizedPath + "；" + ex.Message);
            }
            try { connection.Dispose(); } catch { }
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
                    item["provider"] = "sql";
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

        private sealed class SqlLease : IDisposable
        {
            private readonly SqlFileLockProvider _owner;
            private readonly string _normalizedPath;
            private readonly LockEntry _entry;
            private readonly SqlConnection _connection;
            private int _disposed;

            public SqlLease(
                SqlFileLockProvider owner,
                string normalizedPath,
                LockEntry entry,
                SqlConnection connection)
            {
                _owner = owner;
                _normalizedPath = normalizedPath;
                _entry = entry;
                _connection = connection;
            }

            public void Dispose()
            {
                if (Interlocked.Exchange(ref _disposed, 1) != 0)
                {
                    return;
                }
                _owner.ReleaseLease(_normalizedPath, _entry, _connection);
            }
        }
    }
}
