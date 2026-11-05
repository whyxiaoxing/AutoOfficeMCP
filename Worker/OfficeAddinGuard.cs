using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Office.Config;

namespace AutoOfficeMCP.Worker
{
    /// <summary>
    /// T11：worker 启动后断开用户级 Office 加载项，降低多进程共享同一 Windows 配置带来的插件干扰。
    /// 使用 late-bound 反射，避免额外依赖 Microsoft.Office.Core PIA。
    /// 注意：加载项可能在 guard 执行前已加载并运行过启动代码；强隔离仍需独立 Windows 用户/会话。
    /// </summary>
    public static class OfficeAddinGuard
    {
        public static void Disable(object application)
        {
            if (application == null || !AppConfig.Current.DisableAddins)
            {
                return;
            }

            DisableCollection(application, "COMAddIns", "Connect");
            DisableCollection(application, "AddIns", "Installed");
        }

        private static void DisableCollection(object application, string collectionName, string flagName)
        {
            object collection = null;
            try
            {
                collection = GetProperty(application, collectionName);
                if (collection == null)
                {
                    return;
                }
                int count = GetIntProperty(collection, "Count");
                for (int index = 1; index <= count; index++)
                {
                    object item = null;
                    try
                    {
                        item = GetItem(collection, index);
                        if (item == null)
                        {
                            continue;
                        }
                        string name = GetStringProperty(item, "ProgId");
                        if (string.IsNullOrWhiteSpace(name))
                        {
                            name = GetStringProperty(item, "Description");
                        }
                        if (string.IsNullOrWhiteSpace(name))
                        {
                            name = GetStringProperty(item, "Name");
                        }
                        SetProperty(item, flagName, false);
                        Console.Error.WriteLine("[worker] 已断开 Office 加载项：" + collectionName +
                                                "[" + index + "] " + (name ?? "(unknown)"));
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("[worker] 断开 Office 加载项失败：" + collectionName +
                                                "[" + index + "]；" + ex.Message);
                    }
                    finally
                    {
                        Release(item);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[worker] 读取 " + collectionName + " 失败：" + ex.Message);
            }
            finally
            {
                Release(collection);
            }
        }

        private static object GetProperty(object target, string name)
        {
            return target.GetType().InvokeMember(
                name,
                BindingFlags.GetProperty,
                null,
                target,
                null);
        }

        private static int GetIntProperty(object target, string name)
        {
            object value = GetProperty(target, name);
            return value == null ? 0 : Convert.ToInt32(value);
        }

        private static string GetStringProperty(object target, string name)
        {
            try
            {
                object value = GetProperty(target, name);
                return value == null ? string.Empty : Convert.ToString(value);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static object GetItem(object collection, int index)
        {
            try
            {
                return collection.GetType().InvokeMember(
                    "Item",
                    BindingFlags.GetProperty,
                    null,
                    collection,
                    new object[] { index });
            }
            catch (TargetInvocationException)
            {
                return collection.GetType().InvokeMember(
                    "Item",
                    BindingFlags.InvokeMethod,
                    null,
                    collection,
                    new object[] { index });
            }
        }

        private static void SetProperty(object target, string name, object value)
        {
            target.GetType().InvokeMember(
                name,
                BindingFlags.SetProperty,
                null,
                target,
                new object[] { value });
        }

        private static void Release(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                try
                {
                    Marshal.ReleaseComObject(comObject);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[worker] 释放加载项 COM 对象失败：" + ex.Message);
                }
            }
        }
    }
}
