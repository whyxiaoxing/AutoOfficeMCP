using System;
using System.Globalization;
using System.Reflection;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>
    /// T06 P3 补充工具：迷你图、照相机图片、Power Query 刷新。
    /// 全部通过 OfficeSession.RunExcel 进入，COM 对象在本方法内释放。
    /// </summary>
    public sealed partial class ExcelExtensionsExecutor
    {
        private static OperationResult ExecuteSparklineManage(Excel.Application application, JObject arguments)
        {
            string action = ExcelArgumentReader.GetTrimmedString(arguments, "action", "operation").ToLowerInvariant();
            if (action.Length == 0)
            {
                action = "add";
            }
            if (action != "add" && action != "create" && action != "clear" && action != "list")
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "action 仅支持 add/clear/list。");
            }

            string targetAddress = ExcelArgumentReader.GetRangeAddress(
                arguments,
                "targetRange",
                "range",
                "location");
            if (targetAddress.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 targetRange 参数（迷你图放置位置）。");
            }

            using (ExcelRangeContext target = RangeResolver.Resolve(
                       application,
                       targetAddress,
                       false,
                       "excel_sparkline_manage"))
            {
                if (action != "list")
                {
                    ExcelComSupport.EnsureWritable(application, target.Workbook, target.Worksheet, "excel_sparkline_manage");
                }

                dynamic groups = null;
                try
                {
                    groups = target.Range.SparklineGroups;

                    if (action == "clear")
                    {
                        int before = Convert.ToInt32(groups.Count, CultureInfo.InvariantCulture);
                        groups.Clear();
                        JObject clearData = new JObject();
                        clearData["removedGroupCount"] = before;
                        clearData["target"] = ExcelComSupport.RangeAddress(target.Range, target.SheetName);
                        Observation clearObservation = Observation.Create(
                            "sparkline_manage",
                            "已清除 " + before + " 组迷你图。",
                            before > 0);
                        clearObservation.AddTargetRef(ExcelComSupport.RangeAddress(target.Range, target.SheetName));
                        return OperationResult.Succeed(
                            "已清除 " + before + " 组迷你图。",
                            clearData,
                            clearObservation);
                    }

                    if (action == "list")
                    {
                        JArray items = new JArray();
                        int count = Convert.ToInt32(groups.Count, CultureInfo.InvariantCulture);
                        for (int index = 1; index <= count; index++)
                        {
                            object groupObject = null;
                            try
                            {
                                groupObject = groups[index];
                                dynamic group = groupObject;
                                JObject item = new JObject();
                                item["index"] = index;
                                item["type"] = SparklineTypeText(Convert.ToInt32(group.Type, CultureInfo.InvariantCulture));
                                item["sourceData"] = Convert.ToString(group.SourceData, CultureInfo.InvariantCulture) ?? string.Empty;
                                items.Add(item);
                            }
                            finally
                            {
                                ComHelper.Release(groupObject);
                            }
                        }

                        JObject listData = new JObject();
                        listData["target"] = ExcelComSupport.RangeAddress(target.Range, target.SheetName);
                        listData["groupCount"] = items.Count;
                        listData["groups"] = items;
                        string listSummary = "迷你图清单：共 " + items.Count + " 组。";
                        Observation listObservation = ReadOnlyObservation("sparkline_manage", listSummary);
                        return OperationResult.Succeed(listSummary, listData, listObservation);
                    }

                    string sourceAddress = ExcelArgumentReader.GetRangeAddress(
                        arguments,
                        "sourceRange",
                        "source",
                        "dataRange");
                    if (sourceAddress.Length == 0)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "add 需要 sourceRange 参数。");
                    }

                    string typeText = ExcelArgumentReader.GetTrimmedString(arguments, "type", "sparklineType")
                        .ToLowerInvariant();
                    int typeValue = SparklineTypeValue(typeText);
                    if (typeValue == 0)
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "type 仅支持 line/column/winloss。");
                    }

                    using (ExcelRangeContext source = RangeResolver.Resolve(
                               application,
                               sourceAddress,
                               false,
                               "excel_sparkline_manage"))
                    {
                        string sourceRef = ExcelComSupport.RangeAddress(source.Range, source.SheetName);
                        dynamic group = groups.Add(typeValue, sourceRef);
                        JArray warnings = new JArray();
                        ApplySparklineOptions(group, arguments, warnings);

                        JObject data = new JObject();
                        data["target"] = ExcelComSupport.RangeAddress(target.Range, target.SheetName);
                        data["source"] = sourceRef;
                        data["type"] = SparklineTypeText(typeValue);
                        data["groupCount"] = Convert.ToInt32(groups.Count, CultureInfo.InvariantCulture);
                        if (warnings.Count > 0)
                        {
                            data["warnings"] = warnings;
                        }

                        int changedCount = 1;
                        Observation observation = Observation.Create(
                            "sparkline_manage",
                            "已添加 " + changedCount + " 组迷你图。",
                            true);
                        observation.AddTargetRef(ExcelComSupport.RangeAddress(target.Range, target.SheetName));
                        ComHelper.Release(group);
                        return OperationResult.Succeed(
                            "已添加 " + changedCount + " 组迷你图。",
                            data,
                            observation);
                    }
                }
                finally
                {
                    ComHelper.Release(groups);
                }
            }
        }

        private static void ApplySparklineOptions(dynamic group, JObject arguments, JArray warnings)
        {
            if (group == null)
            {
                return;
            }

            object points = null;
            object markers = null;
            object highpoint = null;
            object lowpoint = null;
            object negative = null;
            object firstpoint = null;
            object lastpoint = null;
            try
            {
                points = group.Points;
                if (points == null)
                {
                    return;
                }
                markers = ((dynamic)points).Markers;
                highpoint = ((dynamic)points).Highpoint;
                lowpoint = ((dynamic)points).Lowpoint;
                negative = ((dynamic)points).Negative;
                firstpoint = ((dynamic)points).Firstpoint;
                lastpoint = ((dynamic)points).Lastpoint;

                TrySetBoolProperty(markers, "Visible", arguments, "showMarkers", warnings);
                TrySetBoolProperty(highpoint, "Visible", arguments, "showHighPoint", warnings);
                TrySetBoolProperty(lowpoint, "Visible", arguments, "showLowPoint", warnings);
                TrySetBoolProperty(negative, "Visible", arguments, "showNegativePoints", warnings);
                TrySetBoolProperty(firstpoint, "Visible", arguments, "showFirstPoint", warnings);
                TrySetBoolProperty(lastpoint, "Visible", arguments, "showLastPoint", warnings);
            }
            catch (Exception ex)
            {
                warnings.Add("读取/设置迷你图 Points 失败：" + ex.Message);
            }
            finally
            {
                ComHelper.Release(
                    lastpoint,
                    firstpoint,
                    negative,
                    lowpoint,
                    highpoint,
                    markers,
                    points);
            }

            string dateAddress = ExcelArgumentReader.GetRangeAddress(arguments, "dateAxisRange", "dateRange");
            if (dateAddress.Length > 0)
            {
                try
                {
                    group.DateRange = dateAddress;
                }
                catch (Exception ex)
                {
                    warnings.Add("设置 DateRange 失败：" + ex.Message);
                }
            }
        }

        private static void TrySetBoolProperty(
            object target,
            string propertyName,
            JObject arguments,
            string argumentName,
            JArray warnings)
        {
            if (target == null || arguments == null || arguments[argumentName] == null)
            {
                return;
            }
            bool value = ExcelArgumentReader.GetBool(arguments, false, argumentName);
            try
            {
                target.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    target,
                    new object[] { value });
            }
            catch (Exception ex)
            {
                warnings.Add(argumentName + " 设置失败：" + ex.Message);
            }
        }

        private static int SparklineTypeValue(string text)
        {
            switch (text)
            {
                case "":
                case "line":
                case "sparkline":
                    return 1; // xlSparkLine
                case "column":
                case "columnstacked":
                    return 2; // xlSparkColumn
                case "winloss":
                case "win_loss":
                case "columnstacked100":
                    return 3; // xlSparkColumnStacked100
                default:
                    return 0;
            }
        }

        private static string SparklineTypeText(int value)
        {
            switch (value)
            {
                case 1:
                    return "line";
                case 2:
                    return "column";
                case 3:
                    return "winloss";
                default:
                    return "type" + value.ToString(CultureInfo.InvariantCulture);
            }
        }

        // ------------------------------------------------------------------
        // 照相机图片
        // ------------------------------------------------------------------

        private static OperationResult ExecuteCameraPicture(Excel.Application application, JObject arguments)
        {
            string sourceAddress = ExcelArgumentReader.GetRangeAddress(
                arguments,
                "sourceRange",
                "range",
                "source");
            if (sourceAddress.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 sourceRange 参数。");
            }
            string targetAddress = ExcelArgumentReader.GetRangeAddress(
                arguments,
                "targetCell",
                "targetRange",
                "destination");
            if (targetAddress.Length == 0)
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 targetCell/targetRange 参数（图片左上角位置）。");
            }

            bool linked = ExcelArgumentReader.GetBool(arguments, true, "link", "linked");
            string shapeName = ExcelArgumentReader.GetTrimmedString(arguments, "name", "shapeName");

            using (ExcelRangeContext source = RangeResolver.Resolve(
                       application,
                       sourceAddress,
                       false,
                       "excel_camera_picture"))
            using (ExcelRangeContext target = RangeResolver.Resolve(
                       application,
                       targetAddress,
                       false,
                       "excel_camera_picture"))
            {
                ExcelComSupport.EnsureWritable(application, target.Workbook, target.Worksheet, "excel_camera_picture");
                source.Range.CopyPicture(
                    Excel.XlPictureAppearance.xlScreen,
                    Excel.XlCopyPictureFormat.xlPicture);

                object picturesObject = null;
                object pictureObject = null;
                try
                {
                    // PIA 将 Worksheet.Pictures 暴露为带可选索引的方法；Missing 返回 Pictures 集合。
                    picturesObject = target.Worksheet.Pictures(Type.Missing);
                    JArray warnings = new JArray();
                    bool linkedApplied = linked;
                    try
                    {
                        pictureObject = InvokePaste(picturesObject, linked);
                    }
                    catch (Exception ex)
                    {
                        if (!linked)
                        {
                            throw;
                        }
                        // Office 2024 实测：Pictures.Paste(Link:=True) 返回“方法无效”；
                        // 退化为静态图片，并如实报告，不宣称已创建链接图片。
                        warnings.Add("Pictures.Paste(Link=true) 失败，已退化为静态图片：" + ex.Message);
                        pictureObject = InvokePaste(picturesObject, false);
                        linkedApplied = false;
                    }

                    dynamic picture = pictureObject;
                    picture.Top = target.Range.Top;
                    picture.Left = target.Range.Left;
                    if (shapeName.Length > 0)
                    {
                        picture.Name = shapeName;
                    }

                    JObject data = new JObject();
                    data["source"] = ExcelComSupport.RangeAddress(source.Range, source.SheetName);
                    data["target"] = ExcelComSupport.RangeAddress(target.Range, target.SheetName);
                    data["linkRequested"] = linked;
                    data["linked"] = linkedApplied;
                    data["shapeName"] = Convert.ToString(picture.Name, CultureInfo.InvariantCulture) ?? string.Empty;
                    data["widthPt"] = Convert.ToDouble(picture.Width, CultureInfo.InvariantCulture);
                    data["heightPt"] = Convert.ToDouble(picture.Height, CultureInfo.InvariantCulture);
                    data["clipboardRequired"] = true;
                    if (warnings.Count > 0)
                    {
                        data["warnings"] = warnings;
                    }

                    string summary = "已生成照相机图片「" + data.Value<string>("shapeName") + "」" +
                                     (linkedApplied ? "（链接图片）。" : "（静态图片）。");
                    Observation observation = Observation.Create("camera_picture", summary, true);
                    observation.AddTargetRef(ExcelComSupport.RangeAddress(target.Range, target.SheetName));
                    return OperationResult.Succeed(summary, data, observation);
                }
                finally
                {
                    ComHelper.Release(pictureObject, picturesObject);
                }
            }
        }

        private static object InvokePaste(object pictures, bool linked)
        {
            if (pictures == null)
            {
                throw new InvalidOperationException("Excel 没有返回 Pictures 集合。");
            }
            // 直接 dynamic 调用在 Office 2024 上会把 bool 参数错误绑定为属性访问；
            // 用 IDispatch InvokeMethod 显式传 Variant 参数。
            return pictures.GetType().InvokeMember(
                "Paste",
                BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                null,
                pictures,
                new object[] { linked });
        }

        // ------------------------------------------------------------------
        // Power Query 刷新
        // ------------------------------------------------------------------

        private static OperationResult ExecutePowerQueryRefresh(Excel.Application application, JObject arguments)
        {
            string action = ExcelArgumentReader.GetTrimmedString(arguments, "action", "operation").ToLowerInvariant();
            if (action.Length == 0)
            {
                action = "refresh";
            }
            if (action != "refresh" && action != "list")
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "action 仅支持 refresh/list。");
            }

            Excel.Workbook workbook = null;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                if (action == "refresh" && workbook.ReadOnly)
                {
                    return Fail(OfficeErrorCode.READ_ONLY, "工作簿以只读方式打开，无法刷新 Power Query。");
                }

                object queriesObject = null;
                try
                {
                    dynamic dynamicWorkbook = workbook;
                    queriesObject = dynamicWorkbook.Queries;
                    if (queriesObject == null)
                    {
                        return Fail(
                            OfficeErrorCode.INVALID_PARAMS,
                            "当前 Excel 版本不提供 Workbook.Queries；无法读取/刷新 Power Query。");
                    }

                    dynamic queries = queriesObject;
                    int count = Convert.ToInt32(queries.Count, CultureInfo.InvariantCulture);
                    string queryName = ExcelArgumentReader.GetTrimmedString(arguments, "queryName", "name");
                    bool refreshAll = ExcelArgumentReader.GetBool(
                        arguments,
                        action == "refresh" && queryName.Length == 0,
                        "refreshAll");

                    JArray items = new JArray();
                    int refreshed = 0;
                    int failed = 0;
                    JArray errors = new JArray();

                    for (int index = 1; index <= count; index++)
                    {
                        object queryObject = null;
                        try
                        {
                            queryObject = queries[index];
                            dynamic query = queryObject;
                            string name = Convert.ToString(query.Name, CultureInfo.InvariantCulture) ?? string.Empty;
                            JObject item = new JObject();
                            item["name"] = name;
                            item["formula"] = Convert.ToString(query.Formula, CultureInfo.InvariantCulture) ?? string.Empty;

                            bool selected = queryName.Length == 0
                                || string.Equals(name, queryName, StringComparison.OrdinalIgnoreCase);
                            if (action == "refresh" && selected)
                            {
                                try
                                {
                                    query.Refresh();
                                    refreshed++;
                                    item["refreshed"] = true;
                                }
                                catch (Exception ex)
                                {
                                    failed++;
                                    item["refreshed"] = false;
                                    item["error"] = ex.Message;
                                    JObject error = new JObject();
                                    error["query"] = name;
                                    error["error"] = ex.Message;
                                    errors.Add(error);
                                }
                            }
                            else
                            {
                                item["refreshed"] = false;
                            }
                            items.Add(item);
                        }
                        finally
                        {
                            ComHelper.Release(queryObject);
                        }
                    }

                    if (action == "refresh" && refreshAll && count == 0)
                    {
                        workbook.RefreshAll();
                    }

                    JObject data = new JObject();
                    data["action"] = action;
                    data["queryCount"] = count;
                    data["refreshedCount"] = refreshed;
                    data["failedCount"] = failed;
                    data["queries"] = items;
                    data["errors"] = errors;
                    data["clipboardRequired"] = false;

                    if (failed > 0)
                    {
                        string partialMessage = "Power Query 刷新完成 " + refreshed + " 个，失败 " + failed + " 个。";
                        return OperationResult.Failed(
                            OfficeErrorCode.PARTIAL_APPLY,
                            partialMessage,
                            null,
                            data,
                            "刷新失败项保持原状态；不会自动重放。",
                            true);
                    }

                    string summary = action == "list"
                        ? "Power Query 清单：共 " + count + " 个。"
                        : "Power Query 刷新完成：成功 " + refreshed + " 个。";
                    Observation observation = action == "list"
                        ? ReadOnlyObservation("powerquery_refresh", summary)
                        : Observation.Create("powerquery_refresh", summary, refreshed > 0);
                    return OperationResult.Succeed(summary, data, observation);
                }
                finally
                {
                    ComHelper.Release(queriesObject);
                }
            }
            catch (Exception ex)
            {
                OfficeErrorCode code = OfficeErrorMapper.MapException(ex, OfficeApp.Excel);
                return Fail(code, "Power Query 操作失败：" + ex.Message);
            }
            finally
            {
                ComHelper.Release(workbook);
            }
        }
    }
}
