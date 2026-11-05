using System;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace Office.Parallel
{
    /// <summary>文件锁抽象：跨进程锁文件或 SQL Server 分布式锁。</summary>
    internal interface IFileLockProvider
    {
        Task<IDisposable> AcquireAsync(string normalizedPath, string displayPath, TimeSpan timeout);

        JArray DescribeLocks();

        int Count { get; }
    }
}
