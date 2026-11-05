using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Office;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointe
{
    public sealed partial class PowerPointToolExecutor
    {
        private const int VbextCtStdModule = 1;

        private static readonly Regex ProcedureRegex = new Regex(
            @"^\s*(?:(?:Public|Private)\s+)?(?:Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)",
            RegexOptions.Multiline | RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static OperationResult ExecuteAddSpeakerNotes(
            PowerPoint.Application application,
            JObject arguments)
        {
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.SlideRange notesPage = null;
            PowerPoint.Shapes shapes = null;
            PowerPoint.Shape body = null;
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                string notes = PowerPointArgumentReader.RequireString(arguments, "notes");
                int slideIndex = PowerPointArgumentReader.GetInt(arguments, -1, "slideIndex");
                slide = PowerPointComUtil.ResolveSlide(application, presentation, slideIndex);
                notesPage = slide.NotesPage;
                shapes = notesPage.Shapes;

                for (int index = 1; index <= shapes.Count; index++)
                {
                    PowerPoint.Shape shape = null;
                    try
                    {
                        shape = shapes[index];
                        dynamic dynamicShape = shape;
                        if (Convert.ToInt32(dynamicShape.Type) == 14 &&
                            shape.PlaceholderFormat.Type ==
                            PowerPoint.PpPlaceholderType.ppPlaceholderBody)
                        {
                            body = shape;
                            shape = null;
                            break;
                        }
                    }
                    finally
                    {
                        ComHelper.Release(shape);
                    }
                }

                if (body == null && shapes.Count >= 2)
                {
                    PowerPoint.Shape fallback = null;
                    try
                    {
                        fallback = shapes[2];
                        dynamic dynamicFallback = fallback;
                        if (Convert.ToInt32(dynamicFallback.HasTextFrame) != 0)
                        {
                            body = fallback;
                            fallback = null;
                        }
                    }
                    finally
                    {
                        ComHelper.Release(fallback);
                    }
                }

                if (body == null)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "未找到备注占位符，无法写入演讲备注。");
                }

                PowerPoint.TextRange textRange = null;
                try
                {
                    textRange = body.TextFrame.TextRange;
                    string normalized = PowerPointComUtil.NormalizeText(notes);
                    textRange.Text = normalized;
                    string applied = textRange.Text ?? string.Empty;
                    if (!string.Equals(
                            NormalizeForCompare(applied),
                            NormalizeForCompare(normalized),
                            StringComparison.Ordinal))
                    {
                        return Fail(
                            OfficeErrorCode.PARTIAL_APPLY,
                            "演讲备注写入后回读不一致，未确认成功。");
                    }
                }
                finally
                {
                    ComHelper.Release(textRange);
                }

                JObject data = new JObject();
                data["slideIndex"] = slide.SlideIndex;
                data["notesLength"] = notes.Length;
                return Success(
                    "已为第 " + slide.SlideIndex.ToString() + " 张幻灯片添加演讲备注。",
                    "write",
                    true,
                    SlideRef(presentation, slide),
                    data);
            }
            finally
            {
                ComHelper.Release(body, shapes, notesPage, slide, presentation);
            }
        }

        private static OperationResult ExecuteVba(
            PowerPoint.Application application,
            JObject arguments)
        {
            if (!PowerPointArgumentReader.GetBool(arguments, false, "confirm"))
            {
                return OperationResult.Cancelled(
                    "ppt_execute_vba 需要显式 confirm=true 才会执行；本次未执行任何 VBA 代码。");
            }

            OperationResult denied = SecurityGate.CreateDenial(SecurityRisk.Vba, "ppt_execute_vba");
            if (denied != null)
            {
                return denied;
            }

            string code = PowerPointArgumentReader.GetString(arguments, "code", "vbaCode");
            if (string.IsNullOrWhiteSpace(code))
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 code/vbaCode 参数。");
            }
            float timeoutSec = PowerPointArgumentReader.GetFloat(arguments, 30f, "timeoutSec");
            if (timeoutSec < 1f || timeoutSec > 300f)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "timeoutSec 必须在 1..300 之间。");
            }

            PowerPoint.Presentation presentation = null;
            object vbeObject = null;
            object projectObject = null;
            object componentsObject = null;
            object componentObject = null;
            bool moduleAdded = false;
            string moduleName = "McpPpt_" +
                                DateTime.Now.ToString("HHmmssfff", CultureInfo.InvariantCulture);
            try
            {
                presentation = PowerPointComUtil.RequireActivePresentation(application);
                dynamic dynamicApplication = application;
                vbeObject = dynamicApplication.VBE;
                if (vbeObject == null)
                {
                    return Fail(
                        OfficeErrorCode.VBA_DISABLED,
                        "无法访问 PowerPoint VBE。请在信任中心启用" +
                        "“信任对 VBA 工程对象模型的访问”。");
                }

                dynamic vbe = vbeObject;
                projectObject = vbe.ActiveVBProject;
                if (projectObject == null)
                {
                    return Fail(
                        OfficeErrorCode.VBA_DISABLED,
                        "当前演示文稿没有可用的 VBA 工程。");
                }

                dynamic project = projectObject;
                componentsObject = project.VBComponents;
                dynamic components = componentsObject;
                componentObject = components.Add(VbextCtStdModule);
                moduleAdded = true;
                dynamic component = componentObject;
                component.Name = moduleName;

                string procedureName;
                string moduleCode = BuildModuleCode(code, out procedureName);
                object codeModuleObject = component.CodeModule;
                try
                {
                    dynamic codeModule = codeModuleObject;
                    codeModule.AddFromString(moduleCode);
                }
                finally
                {
                    ComHelper.Release(codeModuleObject);
                }

                dynamic app = application;
                object runResult = app.Run(moduleName + "." + procedureName);

                JObject data = new JObject();
                data["moduleName"] = moduleName;
                data["procedureName"] = procedureName;
                data["timeoutSec"] = timeoutSec;
                data["timeoutEnforced"] = false;
                data["returnValue"] = ToReturnToken(runResult);
                data["debugOutput"] = new JArray();
                data["debugOutputSupported"] = false;
                data["debugOutputNote"] =
                    "PowerPoint COM 无法读取 VBA 立即窗口；Debug.Print 输出不会回传，请用 Function 返回值或写入单元格。";
                data["note"] = "PowerPoint VBA 可修改任意对象；写操作会登记原生撤销组，可用 office_undo_last 尝试撤销。";
                return Success(
                    "PowerPoint VBA 已执行：" + moduleName + "." + procedureName + "。",
                    "run_vba",
                    true,
                    "ppt:" + presentation.Name,
                    data);
            }
            catch (COMException ex)
            {
                OfficeErrorCode errorCode = OfficeErrorMapper.MapException(ex, OfficeApp.PowerPoint);
                if (errorCode == OfficeErrorCode.VBA_DISABLED)
                {
                    return Fail(
                        OfficeErrorCode.VBA_DISABLED,
                        "VBA 执行被拒绝：security.allowVba 已放行，但 PowerPoint 拒绝访问 VBA 工程。" +
                        "请在“文件 > 选项 > 信任中心 > 信任中心设置 > 宏设置”中启用" +
                        "“信任对 VBA 工程对象模型的访问”，然后重启 PowerPoint。原始信息：" + ex.Message);
                }
                return Fail(
                    errorCode,
                    "PowerPoint VBA 执行失败：" +
                    OfficeErrorMapper.DescribeException(ex, OfficeApp.PowerPoint));
            }
            finally
            {
                if (moduleAdded && componentsObject != null && componentObject != null)
                {
                    try
                    {
                        dynamic components = componentsObject;
                        components.Remove(componentObject);
                    }
                    catch (Exception cleanupError)
                    {
                        Console.Error.WriteLine("[ppt] 删除 VBA 临时模块失败：" + cleanupError.Message);
                    }
                }
                ComHelper.Release(componentObject, componentsObject, projectObject, vbeObject, presentation);
            }
        }

        private static JToken ToReturnToken(object value)
        {
            if (value == null)
            {
                return JValue.CreateNull();
            }
            if (value is JToken)
            {
                return (JToken)value;
            }
            if (value is string || value is bool)
            {
                return new JValue(value);
            }
            if (value is byte || value is short || value is int || value is long ||
                value is float || value is double || value is decimal)
            {
                return new JValue(value);
            }
            Array array = value as Array;
            if (array != null)
            {
                JArray result = new JArray();
                foreach (object item in array)
                {
                    result.Add(ToReturnToken(item));
                }
                return result;
            }
            try
            {
                if (Marshal.IsComObject(value))
                {
                    string text = Convert.ToString(value, CultureInfo.InvariantCulture);
                    Marshal.ReleaseComObject(value);
                    return new JValue(text ?? string.Empty);
                }
            }
            catch (Exception)
            {
            }
            return new JValue(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
        }

        private static string BuildModuleCode(string code, out string procedureName)
        {
            Match match = ProcedureRegex.Match(code ?? string.Empty);
            if (match.Success)
            {
                procedureName = match.Groups[1].Value;
                return code;
            }

            procedureName = "Auto_Run";
            return "Sub Auto_Run()" + Environment.NewLine +
                   (code ?? string.Empty) + Environment.NewLine +
                   "End Sub";
        }

        private static string NormalizeForCompare(string value)
        {
            if (value == null)
            {
                return string.Empty;
            }
            return value.Replace("\r\n", "\r").Replace("\n", "\r").TrimEnd();
        }
    }
}
