using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;

namespace Office.Parallel
{
    /// <summary>
    /// T11 收尾：保存前后文件指纹。优先 SHA256；文件过大或无法读取时退回长度 + 最后写入时间。
    /// 只读文件属性，不修改源文件；校验失败如实报告，不伪造 saved 事实。
    /// </summary>
    internal sealed class FileIntegritySnapshot
    {
        public bool Exists { get; set; }

        public long Length { get; set; }

        public DateTime LastWriteUtc { get; set; }

        public string Sha256 { get; set; }

        public string Method { get; set; }

        public string Note { get; set; }

        public FileIntegritySnapshot()
        {
            Sha256 = string.Empty;
            Method = string.Empty;
            Note = string.Empty;
        }
    }

    internal static class FileIntegrityVerifier
    {
        /// <summary>
        /// 抓取文件指纹。maxHashMb 小于等于 0 表示只使用长度 + mtime；
        /// 超过上限或文件被 Office 独占时同样退回长度 + mtime，并写明原因。
        /// </summary>
        public static FileIntegritySnapshot Capture(string filePath, int maxHashMb)
        {
            FileIntegritySnapshot snapshot = new FileIntegritySnapshot();
            if (string.IsNullOrWhiteSpace(filePath))
            {
                snapshot.Note = "文件路径为空。";
                return snapshot;
            }

            try
            {
                FileInfo info = new FileInfo(filePath);
                info.Refresh();
                snapshot.Exists = info.Exists;
                if (!info.Exists)
                {
                    snapshot.Note = "文件不存在。";
                    return snapshot;
                }

                snapshot.Length = info.Length;
                snapshot.LastWriteUtc = info.LastWriteTimeUtc;
                if (maxHashMb <= 0)
                {
                    snapshot.Method = "mtime+length";
                    snapshot.Note = "已配置不计算 SHA256。";
                    return snapshot;
                }

                long maxBytes = (long)maxHashMb * 1024L * 1024L;
                if (snapshot.Length > maxBytes)
                {
                    snapshot.Method = "mtime+length";
                    snapshot.Note = "文件大小超过 SHA256 上限（" +
                                    maxHashMb.ToString(CultureInfo.InvariantCulture) + " MB）。";
                    return snapshot;
                }

                try
                {
                    snapshot.Sha256 = ComputeSha256(filePath);
                    snapshot.Method = "sha256";
                    return snapshot;
                }
                catch (Exception ex)
                {
                    snapshot.Method = "mtime+length";
                    snapshot.Note = "读取文件计算 SHA256 失败，退回长度 + mtime：" + ex.Message;
                    return snapshot;
                }
            }
            catch (Exception ex)
            {
                snapshot.Note = "读取文件属性失败：" + ex.Message;
                return snapshot;
            }
        }

        /// <summary>
        /// 校验保存结果：文件必须存在且非空；changedOnDisk 表示指纹是否变化。
        /// 返回 false 只用于“文件不存在/为空”这类确定失败；无变化不算失败，由 changedOnDisk 如实表达。
        /// </summary>
        public static bool Verify(
            FileIntegritySnapshot before,
            FileIntegritySnapshot after,
            out bool changedOnDisk,
            out string note)
        {
            changedOnDisk = false;
            note = string.Empty;
            if (after == null || !after.Exists)
            {
                note = after == null ? "保存后未取得文件指纹。" : ("保存后文件不存在：" + after.Note);
                return false;
            }
            if (after.Length <= 0)
            {
                note = "保存后文件长度为 0。";
                return false;
            }
            if (before == null)
            {
                note = "保存前未取得文件指纹，只能确认保存后文件存在且非空。";
                return true;
            }

            bool beforeHash = !string.IsNullOrWhiteSpace(before.Sha256);
            bool afterHash = !string.IsNullOrWhiteSpace(after.Sha256);
            if (beforeHash && afterHash)
            {
                changedOnDisk = !string.Equals(before.Sha256, after.Sha256, StringComparison.OrdinalIgnoreCase);
                note = "SHA256 校验：" + (changedOnDisk ? "文件内容已变化。" : "文件内容未变化。");
                return true;
            }

            changedOnDisk = after.LastWriteUtc > before.LastWriteUtc || after.Length != before.Length;
            note = "长度 + mtime 校验：" + (changedOnDisk ? "文件已变化。" : "文件未变化。") +
                   "（before=" + before.Length.ToString(CultureInfo.InvariantCulture) + "B/" +
                   before.LastWriteUtc.ToString("o", CultureInfo.InvariantCulture) +
                   "，after=" + after.Length.ToString(CultureInfo.InvariantCulture) + "B/" +
                   after.LastWriteUtc.ToString("o", CultureInfo.InvariantCulture) + "）";
            return true;
        }

        private static string ComputeSha256(string filePath)
        {
            using (FileStream stream = new FileStream(
                       filePath,
                       FileMode.Open,
                       FileAccess.Read,
                       FileShare.ReadWrite | FileShare.Delete,
                       81920,
                       FileOptions.SequentialScan))
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(stream);
                return BitConverter.ToString(hash).Replace("-", string.Empty).ToLowerInvariant();
            }
        }
    }
}
