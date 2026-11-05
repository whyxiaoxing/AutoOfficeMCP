using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;

namespace Office.Parallel
{
    /// <summary>
    /// T11 收尾：worker 命名管道 ACL。
    /// 只允许当前 Windows 用户、SYSTEM 与 Administrators 访问；不添加 Everyone/Anonymous。
    /// 管道名本身仍带随机 GUID，ACL 是第二道防线。
    /// </summary>
    internal static class WorkerPipeSecurity
    {
        public static PipeSecurity Create()
        {
            PipeSecurity security = new PipeSecurity();
            using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
            {
                if (identity != null && identity.User != null)
                {
                    security.AddAccessRule(new PipeAccessRule(
                        identity.User,
                        PipeAccessRights.FullControl,
                        AccessControlType.Allow));
                }
            }

            security.AddAccessRule(new PipeAccessRule(
                new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null),
                PipeAccessRights.FullControl,
                AccessControlType.Allow));
            security.AddAccessRule(new PipeAccessRule(
                new SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid, null),
                PipeAccessRights.FullControl,
                AccessControlType.Allow));
            return security;
        }
    }
}
