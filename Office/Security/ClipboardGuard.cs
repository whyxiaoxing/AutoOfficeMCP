using System;
using System.Collections.Generic;
using Office.Config;

namespace Office.Security
{
    /// <summary>
    /// T11：多用户 worker 下禁止依赖 Windows 剪贴板的工具。
    /// 剪贴板按窗口站/会话共享，不是按进程隔离；两个 worker 同时 Copy/Paste 会互相覆盖。
    /// 这里做工具级开关；真正不依赖剪贴板的实现应在工具内部改为直接赋值。
    /// </summary>
    public static class ClipboardGuard
    {
        private static readonly HashSet<string> ClipboardDependentTools =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "excel_transform_data",   // Range.Copy + PasteSpecial(Transpose)
                "excel_export",           // CopyPicture + Chart.Paste
                "excel_camera_picture",   // Range.CopyPicture + Pictures.Paste
                "word_image_layout",      // CopyAsPicture + Paste
                "word_insert_equation"    // 降级路径 Copy + Paste
            };

        public static bool IsClipboardDependent(string toolName)
        {
            return !string.IsNullOrWhiteSpace(toolName)
                   && ClipboardDependentTools.Contains(toolName.Trim());
        }

        /// <summary>返回 false 时调用方应返回 CLIPBOARD_DISABLED，不得继续执行工具。</summary>
        public static bool EnsureAllowed(string toolName, out string message)
        {
            message = string.Empty;
            if (!IsClipboardDependent(toolName))
            {
                return true;
            }
            if (AppConfig.Current.AllowClipboard)
            {
                return true;
            }
            message = "工具「" + (toolName ?? string.Empty) +
                      "」依赖 Windows 剪贴板；当前 security.allowClipboard=false，已拒绝执行。" +
                      "多用户进程 worker 模式下请改用不依赖剪贴板的工具，或在受控环境显式放行。";
            return false;
        }
    }
}
