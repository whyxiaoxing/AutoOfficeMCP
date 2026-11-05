using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace AutoOfficeMCP.Worker
{
    /// <summary>从 Office 窗口句柄取得 EXCEL.EXE/WINWORD.EXE 的 PID，供 Supervisor 强杀兜底。</summary>
    public static class OfficeProcessUtil
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        public static int FromHwnd(int hwnd)
        {
            if (hwnd == 0)
            {
                return 0;
            }
            uint pid;
            GetWindowThreadProcessId(new IntPtr(hwnd), out pid);
            return pid > int.MaxValue ? 0 : (int)pid;
        }

        /// <summary>
        /// 判断 COM 所控制的 Office 进程是否仍存活。按 PID + 进程名双重校验，
        /// 避免 PID 被系统复用后误判；查询异常时保守返回 true，不误杀正常实例。
        /// </summary>
        public static bool IsAliveOfficeProcess(int pid, string expectedProcessName)
        {
            if (pid <= 0)
            {
                return false;
            }
            try
            {
                using (Process process = Process.GetProcessById(pid))
                {
                    if (process.HasExited)
                    {
                        return false;
                    }
                    if (!string.IsNullOrWhiteSpace(expectedProcessName))
                    {
                        string name = process.ProcessName;
                        if (!string.Equals(name, expectedProcessName.Trim(), StringComparison.OrdinalIgnoreCase))
                        {
                            Console.Error.WriteLine("[worker] officePid=" + pid +
                                                    " 已被其他进程复用（" + name + "），视为 COM 目标丢失。");
                            return false;
                        }
                    }
                    return true;
                }
            }
            catch (ArgumentException)
            {
                // 进程不存在。
                return false;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[worker] 检查 Office 进程存活失败：pid=" + pid + "；" + ex.Message);
                return true;
            }
        }

        /// <summary>按进程名查找 PID；多实例时优先取带主窗口的进程，找不到返回 0。</summary>
        public static int FindProcessIdByName(string processName)
        {
            if (string.IsNullOrWhiteSpace(processName))
            {
                return 0;
            }
            try
            {
                Process[] processes = Process.GetProcessesByName(processName.Trim());
                int fallback = 0;
                foreach (Process process in processes)
                {
                    try
                    {
                        if (fallback == 0)
                        {
                            fallback = process.Id;
                        }
                        if (process.MainWindowHandle != IntPtr.Zero)
                        {
                            return process.Id;
                        }
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }
                return fallback;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[worker] 查找 Office 进程失败：" + processName + "；" + ex.Message);
                return 0;
            }
        }
    }
}
