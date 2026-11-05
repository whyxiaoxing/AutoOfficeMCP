using System;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Office.Parallel
{
    /// <summary>
    /// PowerPoint 是单实例 COM 应用：worker 不能通过 new 可靠地获得独立进程。
    /// 本租约先附着已运行的 PowerPoint（不拥有，退出时绝不 Quit 用户实例）；
    /// 只有确认没有运行实例时才 new，并在释放时 Quit 自己创建的实例。
    /// </summary>
    public sealed class PowerPointApplicationLease : IDisposable
    {
        private const string ProgId = "PowerPoint.Application";
        private bool _disposed;

        public PowerPoint.Application Application { get; private set; }

        public bool Owned { get; private set; }

        private PowerPointApplicationLease(PowerPoint.Application application, bool owned)
        {
            Application = application;
            Owned = owned;
        }

        public static PowerPointApplicationLease Acquire(bool visible)
        {
            object activeObject = null;
            try
            {
                activeObject = Marshal.GetActiveObject(ProgId);
                PowerPoint.Application attached = activeObject as PowerPoint.Application;
                if (attached != null)
                {
                    return new PowerPointApplicationLease(attached, false);
                }
                ComHelper.Release(activeObject);
            }
            catch (COMException ex)
            {
                ComHelper.Release(activeObject);
                if (ex.HResult != OfficeErrorMapper.MkEUnavailable)
                {
                    throw;
                }
            }
            catch (Exception)
            {
                ComHelper.Release(activeObject);
                throw;
            }

            PowerPoint.Application launched = new PowerPoint.Application();
            try
            {
                dynamic dynamicApplication = launched;
                dynamicApplication.Visible = visible ? -1 : 0;
                return new PowerPointApplicationLease(launched, true);
            }
            catch (Exception)
            {
                try
                {
                    launched.Quit();
                }
                catch (Exception)
                {
                }
                ComHelper.Release(launched);
                throw;
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            PowerPoint.Application application = Application;
            Application = null;
            if (application == null)
            {
                return;
            }
            if (Owned)
            {
                try
                {
                    application.Quit();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[ppt-lease] Quit 自动启动的 PowerPoint 失败：" + ex.Message);
                }
            }
            ComHelper.Release(application);
        }
    }
}
