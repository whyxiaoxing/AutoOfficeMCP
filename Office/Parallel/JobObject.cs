using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Office.Parallel
{
    /// <summary>
    /// T11：Windows Job Object，绑定 worker 进程及其 Office 子进程。
    /// 默认 KILL_ON_JOB_CLOSE；worker 退出/被强杀时不会留下 EXCEL/WINWORD。
    /// 可选进程内存上限；memoryLimitMb=0 表示不限制。
    /// </summary>
    internal sealed class JobObject : IDisposable
    {
        private const int JobObjectExtendedLimitInformation = 9;
        private const uint JobObjectLimitProcessMemory = 0x00000100;
        private const uint JobObjectLimitKillOnJobClose = 0x00002000;

        [StructLayout(LayoutKind.Sequential)]
        private struct JobObjectBasicLimitInformation
        {
            public long PerProcessUserTimeLimit;
            public long PerJobUserTimeLimit;
            public uint LimitFlags;
            public UIntPtr MinimumWorkingSetSize;
            public UIntPtr MaximumWorkingSetSize;
            public uint ActiveProcessLimit;
            public UIntPtr Affinity;
            public uint PriorityClass;
            public uint SchedulingClass;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct IoCounters
        {
            public ulong ReadOperationCount;
            public ulong WriteOperationCount;
            public ulong OtherOperationCount;
            public ulong ReadTransferCount;
            public ulong WriteTransferCount;
            public ulong OtherTransferCount;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct JobObjectExtendedLimitInformationStruct
        {
            public JobObjectBasicLimitInformation BasicLimitInformation;
            public IoCounters IoInfo;
            public UIntPtr ProcessMemoryLimit;
            public UIntPtr JobMemoryLimit;
            public UIntPtr PeakProcessMemoryUsed;
            public UIntPtr PeakJobMemoryUsed;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateJobObject(IntPtr jobAttributes, string name);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetInformationJobObject(
            IntPtr job,
            int infoClass,
            IntPtr info,
            uint infoLength);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool TerminateJobObject(IntPtr job, uint exitCode);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr handle);

        private IntPtr _handle;

        public JobObject(int memoryLimitMb)
        {
            _handle = CreateJobObject(IntPtr.Zero, null);
            if (_handle == IntPtr.Zero)
            {
                throw new InvalidOperationException("CreateJobObject 失败，Win32Error=" + Marshal.GetLastWin32Error());
            }

            JobObjectExtendedLimitInformationStruct info = new JobObjectExtendedLimitInformationStruct();
            info.BasicLimitInformation.LimitFlags = JobObjectLimitKillOnJobClose;
            if (memoryLimitMb > 0)
            {
                info.BasicLimitInformation.LimitFlags |= JobObjectLimitProcessMemory;
                info.ProcessMemoryLimit = new UIntPtr((ulong)memoryLimitMb * 1024UL * 1024UL);
            }

            int size = Marshal.SizeOf(typeof(JobObjectExtendedLimitInformationStruct));
            IntPtr pointer = Marshal.AllocHGlobal(size);
            try
            {
                Marshal.StructureToPtr(info, pointer, false);
                if (!SetInformationJobObject(_handle, JobObjectExtendedLimitInformation, pointer, (uint)size))
                {
                    throw new InvalidOperationException(
                        "SetInformationJobObject 失败，Win32Error=" + Marshal.GetLastWin32Error());
                }
            }
            finally
            {
                Marshal.FreeHGlobal(pointer);
            }
        }

        public void Assign(Process process)
        {
            if (process == null)
            {
                throw new ArgumentNullException("process");
            }
            if (!AssignProcessToJobObject(_handle, process.Handle))
            {
                throw new InvalidOperationException(
                    "AssignProcessToJobObject 失败，Win32Error=" + Marshal.GetLastWin32Error());
            }
        }

        public void Terminate(uint exitCode)
        {
            if (_handle == IntPtr.Zero)
            {
                return;
            }
            TerminateJobObject(_handle, exitCode);
        }

        public void Dispose()
        {
            IntPtr handle = _handle;
            _handle = IntPtr.Zero;
            if (handle != IntPtr.Zero)
            {
                try
                {
                    CloseHandle(handle);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[worker-process] 关闭 Job Object 失败：" + ex.Message);
                }
            }
        }
    }
}
