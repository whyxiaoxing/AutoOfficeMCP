using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Office
{
    /// <summary>
    /// COM 纪律工具：STA 断言、逐对象释放、DisplayAlerts/ScreenUpdating 守卫。
    /// 注意：会话缓存中的 Application 对象由 OfficeSession 持有，禁止传给 Release。
    /// </summary>
    public static class ComHelper
    {
        /// <summary>断言当前线程是 STA。所有公共 COM 入口都必须先调用本方法。</summary>
        public static void EnsureSta(string memberName)
        {
            ApartmentState state = Thread.CurrentThread.GetApartmentState();
            if (state != ApartmentState.STA)
            {
                string where = string.IsNullOrWhiteSpace(memberName) ? "COM 调用" : memberName.Trim();
                throw new InvalidOperationException(
                    where + " 必须在 STA 线程执行（当前线程套间为 " + state.ToString() +
                    "）。请从标有 [STAThread] 的主线程或 T09 创建的一次性 STA 线程调用。");
            }
        }

        /// <summary>
        /// 逐个安全释放 COM 对象：空引用跳过、非 COM 对象跳过、重复引用只释放一次，
        /// 释放失败只写 stderr，不覆盖调用方正在抛出的原始异常。
        /// </summary>
        public static void Release(params object[] comObjects)
        {
            if (comObjects == null || comObjects.Length == 0)
            {
                return;
            }

            HashSet<object> seen = new HashSet<object>(ReferenceComparer.Instance);
            foreach (object comObject in comObjects)
            {
                if (comObject == null || comObject is ValueType)
                {
                    continue;
                }
                if (!seen.Add(comObject))
                {
                    continue;
                }
                ReleaseOne(comObject);
            }
        }

        /// <summary>释放单个 COM 对象；非 COM 对象与 null 直接跳过。</summary>
        public static void ReleaseOne(object comObject)
        {
            if (comObject == null || comObject is ValueType)
            {
                return;
            }

            try
            {
                if (!Marshal.IsComObject(comObject))
                {
                    return;
                }
                Marshal.ReleaseComObject(comObject);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[com] ReleaseComObject 失败：" + ex.Message);
            }
        }

        private sealed class ReferenceComparer : IEqualityComparer<object>
        {
            public static readonly ReferenceComparer Instance = new ReferenceComparer();

            public new bool Equals(object x, object y)
            {
                return object.ReferenceEquals(x, y);
            }

            public int GetHashCode(object obj)
            {
                return RuntimeHelpers.GetHashCode(obj);
            }
        }
    }

    /// <summary>Excel 应用级临时设置守卫：DisplayAlerts=false + ScreenUpdating=false，finally 恢复原值。</summary>
    public static class ExcelComGuardExtensions
    {
        public static void Guard(this Excel.Application application, Action action)
        {
            if (action == null)
            {
                throw new ArgumentNullException("action");
            }
            Guard(application, delegate
            {
                action();
                return true;
            });
        }

        public static T Guard<T>(this Excel.Application application, Func<T> action)
        {
            if (application == null)
            {
                throw new ArgumentNullException("application");
            }
            if (action == null)
            {
                throw new ArgumentNullException("action");
            }
            ComHelper.EnsureSta("ExcelComGuardExtensions.Guard");

            bool originalDisplayAlerts = application.DisplayAlerts;
            bool originalScreenUpdating = application.ScreenUpdating;
            try
            {
                application.DisplayAlerts = false;
                application.ScreenUpdating = false;
                return action();
            }
            finally
            {
                TryRestoreBoolean("Excel.DisplayAlerts", delegate { application.DisplayAlerts = originalDisplayAlerts; });
                TryRestoreBoolean("Excel.ScreenUpdating", delegate { application.ScreenUpdating = originalScreenUpdating; });
            }
        }

        private static void TryRestoreBoolean(string name, Action restore)
        {
            try
            {
                restore();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[com] 恢复 " + name + " 失败：" + ex.Message);
            }
        }
    }

    /// <summary>Word 应用级临时设置守卫：DisplayAlerts=wdAlertsNone + ScreenUpdating=false，finally 恢复原值。</summary>
    public static class WordComGuardExtensions
    {
        public static void Guard(this Word.Application application, Action action)
        {
            if (action == null)
            {
                throw new ArgumentNullException("action");
            }
            Guard(application, delegate
            {
                action();
                return true;
            });
        }

        public static T Guard<T>(this Word.Application application, Func<T> action)
        {
            if (application == null)
            {
                throw new ArgumentNullException("application");
            }
            if (action == null)
            {
                throw new ArgumentNullException("action");
            }
            ComHelper.EnsureSta("WordComGuardExtensions.Guard");

            Word.WdAlertLevel originalDisplayAlerts = application.DisplayAlerts;
            bool originalScreenUpdating = application.ScreenUpdating;
            try
            {
                application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                application.ScreenUpdating = false;
                return action();
            }
            finally
            {
                TryRestoreDisplayAlerts(application, originalDisplayAlerts);
                TryRestoreScreenUpdating(application, originalScreenUpdating);
            }
        }

        private static void TryRestoreDisplayAlerts(Word.Application application, Word.WdAlertLevel value)
        {
            try
            {
                application.DisplayAlerts = value;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[com] 恢复 Word.DisplayAlerts 失败：" + ex.Message);
            }
        }

        private static void TryRestoreScreenUpdating(Word.Application application, bool value)
        {
            try
            {
                application.ScreenUpdating = value;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[com] 恢复 Word.ScreenUpdating 失败：" + ex.Message);
            }
        }
    }
}
