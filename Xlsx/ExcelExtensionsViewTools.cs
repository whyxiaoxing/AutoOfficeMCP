using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Newtonsoft.Json.Linq;
using Office;
using Excel = Microsoft.Office.Interop.Excel;

namespace Xlsx
{
    /// <summary>P2 视图、打印与批注管理。</summary>
    public sealed partial class ExcelExtensionsExecutor
    {
        private static OperationResult ExecuteFreezeAndView(Excel.Application application, JObject arguments)
        {
            string action = ExcelArgumentReader.GetTrimmedString(arguments, "action", "operation")
                .ToLowerInvariant();
            if (action.Length == 0)
            {
                action = "get";
            }
            if (action != "get" && action != "set")
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "action 仅支持 get/set。");
            }

            string sheetName = ExcelArgumentReader.GetSheetName(arguments);
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Window window = null;
            object activeSheetObject = null;
            bool activated = false;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                worksheet = string.IsNullOrWhiteSpace(sheetName)
                    ? ExcelComSupport.RequireActiveWorksheet(application)
                    : RangeResolver.ResolveWorksheet(application, workbook, sheetName, false);

                activeSheetObject = application.ActiveSheet;
                Excel.Worksheet activeSheet = activeSheetObject as Excel.Worksheet;
                bool isActive = activeSheet != null &&
                                string.Equals(activeSheet.Name, worksheet.Name, StringComparison.Ordinal);
                ComHelper.Release(activeSheet);

                if (!isActive)
                {
                    if (action == "set" && workbook.ReadOnly)
                    {
                        return Fail(OfficeErrorCode.READ_ONLY, "工作簿以只读方式打开，无法改变视图设置。");
                    }
                    application.Guard(delegate { worksheet.Activate(); });
                    activated = true;
                }

                window = application.ActiveWindow;
                if (window == null)
                {
                    return Fail(OfficeErrorCode.INTERNAL_ERROR, "当前没有可用的 Excel 窗口（请确认 Excel 可见）。");
                }

                JObject before = ReadWindowState(worksheet, window);
                if (action == "get")
                {
                    JObject getData = new JObject();
                    getData["sheet"] = worksheet.Name;
                    getData["activated"] = activated;
                    getData["state"] = before;
                    string summary = "已读取工作表「" + worksheet.Name + "」视图状态。";
                    Observation getObservation = ReadOnlyObservation("freeze_and_view", summary);
                    return OperationResult.Succeed(summary, getData, getObservation);
                }

                bool hasUnfreeze = ExcelArgumentReader.Has(arguments, "unfreeze");
                bool unfreeze = ExcelArgumentReader.GetBool(arguments, false, "unfreeze");
                string freezeCell = ExcelArgumentReader.GetTrimmedString(arguments, "freezeCell", "freezeAt");
                bool hasFreezeRows = ExcelArgumentReader.Has(arguments, "freezeRows");
                bool hasFreezeColumns = ExcelArgumentReader.Has(arguments, "freezeColumns");
                int freezeRows = ExcelArgumentReader.GetInt(arguments, 0, "freezeRows", "rows");
                int freezeColumns = ExcelArgumentReader.GetInt(arguments, 0, "freezeColumns", "columns");
                bool hasGridlines = ExcelArgumentReader.Has(arguments, "gridlines") ||
                                    ExcelArgumentReader.Has(arguments, "displayGridlines");
                bool gridlines = ExcelArgumentReader.GetBool(arguments, false, "gridlines", "displayGridlines");
                bool hasHeadings = ExcelArgumentReader.Has(arguments, "showHeadings") ||
                                   ExcelArgumentReader.Has(arguments, "displayHeadings");
                bool showHeadings = ExcelArgumentReader.GetBool(arguments, true, "showHeadings", "displayHeadings");
                int zoom = ExcelArgumentReader.GetInt(arguments, 0, "zoom");
                string view = ExcelArgumentReader.GetTrimmedString(arguments, "view", "viewMode").ToLowerInvariant();

                bool anyChange = hasUnfreeze || freezeCell.Length > 0 || hasFreezeRows || hasFreezeColumns ||
                                 hasGridlines || hasHeadings || zoom > 0 || view.Length > 0;
                if (!anyChange)
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "set 操作未提供任何视图参数。");
                }

                application.Guard(delegate
                {
                    if (unfreeze)
                    {
                        window.FreezePanes = false;
                        window.SplitRow = 0;
                        window.SplitColumn = 0;
                    }
                    else if (freezeCell.Length > 0 || hasFreezeRows || hasFreezeColumns)
                    {
                        int rowsToFreeze = Math.Max(0, freezeRows);
                        int columnsToFreeze = Math.Max(0, freezeColumns);
                        if (freezeCell.Length > 0)
                        {
                            Excel.Range cell = null;
                            try
                            {
                                cell = (Excel.Range)worksheet.Range[freezeCell];
                                rowsToFreeze = Math.Max(0, cell.Row - 1);
                                columnsToFreeze = Math.Max(0, cell.Column - 1);
                            }
                            finally
                            {
                                ComHelper.Release(cell);
                            }
                        }
                        window.FreezePanes = false;
                        window.SplitRow = rowsToFreeze;
                        window.SplitColumn = columnsToFreeze;
                        window.FreezePanes = true;
                    }

                    if (hasGridlines)
                    {
                        window.DisplayGridlines = gridlines;
                    }
                    if (hasHeadings)
                    {
                        window.DisplayHeadings = showHeadings;
                    }
                    if (zoom > 0)
                    {
                        window.Zoom = ClampInt(zoom, 10, 400);
                    }
                    if (view.Length > 0)
                    {
                        window.View = MapWindowView(view);
                    }
                });

                JObject after = ReadWindowState(worksheet, window);
                bool changed = !JToken.DeepEquals(before, after);
                Observation observation = Observation.Create(
                    "freeze_and_view",
                    changed
                        ? "已更新工作表「" + worksheet.Name + "」视图设置。"
                        : "视图参数与当前状态一致，未产生实际变化。",
                    changed);
                if (changed)
                {
                    observation.AddTargetRef(ExcelComSupport.FullAddress(worksheet.Name, "A1"));
                }
                JObject data = new JObject();
                data["sheet"] = worksheet.Name;
                data["activated"] = activated;
                data["before"] = before;
                data["after"] = after;
                ExcelUndoScope.RecordUnavailable(
                    data,
                    "冻结窗格/网格线/缩放/视图模式属于窗口状态，区域快照无法撤销。");
                return OperationResult.Succeed(
                    changed ? "视图设置已更新。" : "视图设置未发生变化。",
                    data,
                    observation);
            }
            finally
            {
                ComHelper.Release(activeSheetObject, window, worksheet, workbook);
            }
        }

        private static JObject ReadWindowState(Excel.Worksheet worksheet, Excel.Window window)
        {
            JObject state = new JObject();
            state["freezePanes"] = window.FreezePanes;
            state["splitRow"] = window.SplitRow;
            state["splitColumn"] = window.SplitColumn;
            state["displayGridlines"] = window.DisplayGridlines;
            state["displayHeadings"] = window.DisplayHeadings;
            state["zoom"] = Convert.ToInt32(window.Zoom, CultureInfo.InvariantCulture);
            state["view"] = WindowViewText(window.View);
            return state;
        }

        private static string WindowViewText(Excel.XlWindowView view)
        {
            switch (view)
            {
                case Excel.XlWindowView.xlPageBreakPreview:
                    return "pageBreakPreview";
                case Excel.XlWindowView.xlPageLayoutView:
                    return "pageLayout";
                default:
                    return "normal";
            }
        }

        private static Excel.XlWindowView MapWindowView(string view)
        {
            switch (view)
            {
                case "pagebreak":
                case "pagebreakpreview":
                case "page_break_preview":
                    return Excel.XlWindowView.xlPageBreakPreview;
                case "pagelayout":
                case "page_layout":
                    return Excel.XlWindowView.xlPageLayoutView;
                default:
                    return Excel.XlWindowView.xlNormalView;
            }
        }

        private static OperationResult ExecutePrintSetup(Excel.Application application, JObject arguments)
        {
            string action = ExcelArgumentReader.GetTrimmedString(arguments, "action", "operation")
                .ToLowerInvariant();
            if (action.Length == 0)
            {
                action = "get";
            }
            if (action != "get" && action != "set")
            {
                return Fail(OfficeErrorCode.INVALID_PARAMS, "action 仅支持 get/set。");
            }

            string sheetName = ExcelArgumentReader.GetSheetName(arguments);
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            object activeSheetObject = null;
            bool activated = false;
            try
            {
                workbook = ExcelComSupport.RequireActiveWorkbook(application);
                worksheet = string.IsNullOrWhiteSpace(sheetName)
                    ? ExcelComSupport.RequireActiveWorksheet(application)
                    : RangeResolver.ResolveWorksheet(application, workbook, sheetName, false);

                if (action == "set")
                {
                    if (workbook.ReadOnly)
                    {
                        return Fail(OfficeErrorCode.READ_ONLY, "工作簿以只读方式打开，无法修改打印设置。");
                    }
                    ExcelComSupport.EnsureWritable(application, workbook, worksheet, "excel_print_setup");
                }

                activeSheetObject = application.ActiveSheet;
                Excel.Worksheet activeSheet = activeSheetObject as Excel.Worksheet;
                bool isActive = activeSheet != null &&
                                string.Equals(activeSheet.Name, worksheet.Name, StringComparison.Ordinal);
                ComHelper.Release(activeSheet);
                if (!isActive)
                {
                    application.Guard(delegate { worksheet.Activate(); });
                    activated = true;
                }

                JObject before = ReadPageSetup(worksheet);
                if (action == "get")
                {
                    JObject getData = new JObject();
                    getData["sheet"] = worksheet.Name;
                    getData["activated"] = activated;
                    getData["pageSetup"] = before;
                    string summary = "已读取工作表「" + worksheet.Name + "」打印设置。";
                    Observation getObservation = ReadOnlyObservation("print_setup", summary);
                    return OperationResult.Succeed(summary, getData, getObservation);
                }

                Excel.PageSetup pageSetup = null;
                List<string> warnings = new List<string>();
                int applied = 0;
                try
                {
                    pageSetup = worksheet.PageSetup;

                    string printArea = ExcelArgumentReader.GetString(arguments, "printArea");
                    if (printArea.Length > 0)
                    {
                        applied += TryApplyPageSetup(
                            warnings,
                            "printArea",
                            delegate
                            {
                                pageSetup.PrintArea = IsClearToken(printArea) ? string.Empty : printArea;
                            });
                    }
                    string titleRows = ExcelArgumentReader.GetString(arguments, "printTitleRows", "titleRows");
                    if (titleRows.Length > 0)
                    {
                        applied += TryApplyPageSetup(
                            warnings,
                            "printTitleRows",
                            delegate
                            {
                                pageSetup.PrintTitleRows = IsClearToken(titleRows) ? string.Empty : titleRows;
                            });
                    }
                    string titleColumns = ExcelArgumentReader.GetString(
                        arguments,
                        "printTitleColumns",
                        "titleColumns");
                    if (titleColumns.Length > 0)
                    {
                        applied += TryApplyPageSetup(
                            warnings,
                            "printTitleColumns",
                            delegate
                            {
                                pageSetup.PrintTitleColumns = IsClearToken(titleColumns) ? string.Empty : titleColumns;
                            });
                    }

                    string orientation = ExcelArgumentReader.GetTrimmedString(arguments, "orientation")
                        .ToLowerInvariant();
                    if (orientation.Length > 0)
                    {
                        applied += TryApplyPageSetup(
                            warnings,
                            "orientation",
                            delegate
                            {
                                pageSetup.Orientation = orientation == "landscape"
                                    ? Excel.XlPageOrientation.xlLandscape
                                    : Excel.XlPageOrientation.xlPortrait;
                            });
                    }

                    string paperSize = ExcelArgumentReader.GetTrimmedString(arguments, "paperSize");
                    if (paperSize.Length > 0)
                    {
                        applied += TryApplyPageSetup(
                            warnings,
                            "paperSize",
                            delegate { pageSetup.PaperSize = MapPaperSize(paperSize); });
                    }

                    int fitWide;
                    bool hasFitWide = ExcelArgumentReader.TryGetInt(
                        arguments,
                        out fitWide,
                        "fitToPagesWide",
                        "fitWide");
                    int fitTall;
                    bool hasFitTall = ExcelArgumentReader.TryGetInt(
                        arguments,
                        out fitTall,
                        "fitToPagesTall",
                        "fitTall");
                    if (hasFitWide || hasFitTall)
                    {
                        applied += TryApplyPageSetup(
                            warnings,
                            "fitToPages",
                            delegate
                            {
                                pageSetup.Zoom = false;
                                if (hasFitWide)
                                {
                                    pageSetup.FitToPagesWide = fitWide > 0 ? (object)fitWide : false;
                                }
                                if (hasFitTall)
                                {
                                    pageSetup.FitToPagesTall = fitTall > 0 ? (object)fitTall : false;
                                }
                            });
                    }
                    else
                    {
                        int zoom;
                        if (ExcelArgumentReader.TryGetInt(arguments, out zoom, "zoom") && zoom > 0)
                        {
                            applied += TryApplyPageSetup(
                                warnings,
                                "zoom",
                                delegate { pageSetup.Zoom = ClampInt(zoom, 10, 400); });
                        }
                    }

                    applied += ApplyMargin(pageSetup, arguments, "topMargin", "top", warnings);
                    applied += ApplyMargin(pageSetup, arguments, "bottomMargin", "bottom", warnings);
                    applied += ApplyMargin(pageSetup, arguments, "leftMargin", "left", warnings);
                    applied += ApplyMargin(pageSetup, arguments, "rightMargin", "right", warnings);
                    applied += ApplyMargin(pageSetup, arguments, "headerMargin", "header", warnings);
                    applied += ApplyMargin(pageSetup, arguments, "footerMargin", "footer", warnings);

                    if (ExcelArgumentReader.Has(arguments, "centerHorizontally") ||
                        ExcelArgumentReader.Has(arguments, "centerHorizontal"))
                    {
                        bool center = ExcelArgumentReader.GetBool(
                            arguments,
                            false,
                            "centerHorizontally",
                            "centerHorizontal");
                        applied += TryApplyPageSetup(
                            warnings,
                            "centerHorizontally",
                            delegate { pageSetup.CenterHorizontally = center; });
                    }
                    if (ExcelArgumentReader.Has(arguments, "centerVertically") ||
                        ExcelArgumentReader.Has(arguments, "centerVertical"))
                    {
                        bool center = ExcelArgumentReader.GetBool(
                            arguments,
                            false,
                            "centerVertically",
                            "centerVertical");
                        applied += TryApplyPageSetup(
                            warnings,
                            "centerVertically",
                            delegate { pageSetup.CenterVertically = center; });
                    }
                    if (ExcelArgumentReader.Has(arguments, "printGridlines"))
                    {
                        bool printGridlines = ExcelArgumentReader.GetBool(
                            arguments,
                            false,
                            "printGridlines");
                        applied += TryApplyPageSetup(
                            warnings,
                            "printGridlines",
                            delegate { pageSetup.PrintGridlines = printGridlines; });
                    }
                }
                finally
                {
                    ComHelper.Release(pageSetup);
                }

                JObject after = ReadPageSetup(worksheet);
                bool changed = !JToken.DeepEquals(before, after);
                Observation observation = Observation.Create(
                    "print_setup",
                    changed ? "打印设置已更新。" : "打印设置未产生实际变化。",
                    changed);
                if (changed)
                {
                    observation.AddTargetRef(ExcelComSupport.FullAddress(worksheet.Name, "A1"));
                }

                JObject data = new JObject();
                data["sheet"] = worksheet.Name;
                data["activated"] = activated;
                data["before"] = before;
                data["after"] = after;
                data["appliedCount"] = applied;
                data["warnings"] = new JArray(warnings.ToArray());
                ExcelUndoScope.RecordUnavailable(
                    data,
                    "打印设置属于工作表页面元数据，区域快照无法撤销。");

                if (warnings.Count > 0)
                {
                    return OperationResult.PartialApplied(
                        applied,
                        "部分打印设置未生效：" + string.Join("；", warnings.ToArray()) +
                        " 请先确认打印机可用、工作表未保护，再人工修正。",
                        observation,
                        data);
                }

                return OperationResult.Succeed(
                    changed ? "打印设置已更新。" : "打印设置未产生实际变化。",
                    data,
                    observation);
            }
            finally
            {
                ComHelper.Release(activeSheetObject, worksheet, workbook);
            }
        }

        private static int TryApplyPageSetup(
            List<string> warnings,
            string name,
            Action action)
        {
            try
            {
                action();
                return 1;
            }
            catch (Exception ex)
            {
                warnings.Add(name + "：" + ex.Message);
                return 0;
            }
        }

        private static int ApplyMargin(
            Excel.PageSetup pageSetup,
            JObject arguments,
            string propertyName,
            string alias,
            List<string> warnings)
        {
            double value;
            if (!ExcelArgumentReader.TryGetDouble(arguments, out value, propertyName, alias))
            {
                return 0;
            }
            return TryApplyPageSetup(
                warnings,
                propertyName,
                delegate
                {
                    if (propertyName == "topMargin")
                    {
                        pageSetup.TopMargin = value;
                    }
                    else if (propertyName == "bottomMargin")
                    {
                        pageSetup.BottomMargin = value;
                    }
                    else if (propertyName == "leftMargin")
                    {
                        pageSetup.LeftMargin = value;
                    }
                    else if (propertyName == "rightMargin")
                    {
                        pageSetup.RightMargin = value;
                    }
                    else if (propertyName == "headerMargin")
                    {
                        pageSetup.HeaderMargin = value;
                    }
                    else
                    {
                        pageSetup.FooterMargin = value;
                    }
                });
        }

        private static JObject ReadPageSetup(Excel.Worksheet worksheet)
        {
            Excel.PageSetup pageSetup = null;
            try
            {
                pageSetup = worksheet.PageSetup;
                JObject state = new JObject();
                state["printArea"] = pageSetup.PrintArea ?? string.Empty;
                state["printTitleRows"] = pageSetup.PrintTitleRows ?? string.Empty;
                state["printTitleColumns"] = pageSetup.PrintTitleColumns ?? string.Empty;
                state["orientation"] = pageSetup.Orientation == Excel.XlPageOrientation.xlLandscape
                    ? "landscape"
                    : "portrait";
                state["paperSize"] = Convert.ToInt32(pageSetup.PaperSize, CultureInfo.InvariantCulture);
                state["zoom"] = Convert.ToInt32(pageSetup.Zoom, CultureInfo.InvariantCulture);
                state["fitToPagesWide"] = Convert.ToInt32(pageSetup.FitToPagesWide, CultureInfo.InvariantCulture);
                state["fitToPagesTall"] = Convert.ToInt32(pageSetup.FitToPagesTall, CultureInfo.InvariantCulture);
                state["topMargin"] = pageSetup.TopMargin;
                state["bottomMargin"] = pageSetup.BottomMargin;
                state["leftMargin"] = pageSetup.LeftMargin;
                state["rightMargin"] = pageSetup.RightMargin;
                state["headerMargin"] = pageSetup.HeaderMargin;
                state["footerMargin"] = pageSetup.FooterMargin;
                state["centerHorizontally"] = pageSetup.CenterHorizontally;
                state["centerVertically"] = pageSetup.CenterVertically;
                state["printGridlines"] = pageSetup.PrintGridlines;
                return state;
            }
            finally
            {
                ComHelper.Release(pageSetup);
            }
        }

        private static bool IsClearToken(string text)
        {
            string value = (text ?? string.Empty).Trim().ToLowerInvariant();
            return value == "clear" || value == "none" || value == "empty";
        }

        private static Excel.XlPaperSize MapPaperSize(string paperSize)
        {
            string value = (paperSize ?? string.Empty).Trim().ToLowerInvariant();
            switch (value)
            {
                case "a3":
                    return Excel.XlPaperSize.xlPaperA3;
                case "a4":
                    return Excel.XlPaperSize.xlPaperA4;
                case "a5":
                    return Excel.XlPaperSize.xlPaperA5;
                case "letter":
                    return Excel.XlPaperSize.xlPaperLetter;
                case "legal":
                    return Excel.XlPaperSize.xlPaperLegal;
                case "b5":
                    return Excel.XlPaperSize.xlPaperB5;
                default:
                    int numeric;
                    if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out numeric))
                    {
                        return (Excel.XlPaperSize)numeric;
                    }
                    return Excel.XlPaperSize.xlPaperA4;
            }
        }

        private static OperationResult ExecuteCommentManage(Excel.Application application, JObject arguments)
        {
            string action = ExcelArgumentReader.GetTrimmedString(arguments, "action", "operation")
                .ToLowerInvariant();
            if (action.Length == 0)
            {
                action = "list";
            }
            string address = ExcelArgumentReader.GetRangeAddress(arguments, "range", "targetRange", "address");
            string sheetName = ExcelArgumentReader.GetSheetName(arguments);
            if (address.Length == 0 && sheetName.Length > 0)
            {
                address = ExcelComSupport.QuoteSheetName(sheetName) + "!usedrange";
            }
            if (address.Length == 0)
            {
                address = "usedrange";
            }
            string text = ExcelArgumentReader.GetString(arguments, "text", "comment");
            int maxResults = ClampInt(
                ExcelArgumentReader.GetInt(arguments, 1000, "maxResults", "limit"),
                1,
                5000);

            using (ExcelRangeContext context = RangeResolver.Resolve(
                       application,
                       address,
                       false,
                       "excel_comment_manage"))
            {
                if (action == "list")
                {
                    JArray comments = new JArray();
                    bool truncated = false;
                    CollectComments(context.Worksheet, context.Range, comments, maxResults, ref truncated);
                    JObject listData = new JObject();
                    listData["sheet"] = context.SheetName;
                    listData["address"] = context.Range.AddressLocal();
                    listData["count"] = comments.Count;
                    listData["truncated"] = truncated;
                    listData["comments"] = comments;
                    string summary = "批注清单：共 " + comments.Count + " 条。";
                    Observation listObservation = ReadOnlyObservation("comment_manage", summary);
                    return OperationResult.Succeed(summary, listData, listObservation);
                }

                ExcelComSupport.EnsureWritable(
                    application,
                    context.Workbook,
                    context.Worksheet,
                    "excel_comment_manage");

                if (action == "add" || action == "update" || action == "delete")
                {
                    if (text.Length == 0 && (action == "add" || action == "update"))
                    {
                        return Fail(OfficeErrorCode.INVALID_PARAMS, "缺少 text 参数。");
                    }
                }
                else
                {
                    return Fail(OfficeErrorCode.INVALID_PARAMS, "action 仅支持 list/add/update/delete。");
                }

                int rows = context.Range.Rows.Count;
                int columns = context.Range.Columns.Count;
                long totalCells = (long)rows * columns;
                if (totalCells > 20000)
                {
                    return Fail(
                        OfficeErrorCode.INVALID_PARAMS,
                        "批注操作区域过大（" + totalCells + " 个单元格），请缩小 range。");
                }

                int changedCount = 0;
                int skippedCount = 0;
                if (action == "delete")
                {
                    int beforeCount = 0;
                    CollectComments(context.Worksheet, context.Range, new JArray(), int.MaxValue, ref beforeCount);
                    context.Range.ClearComments();
                    changedCount = beforeCount;
                    skippedCount = 0;
                }
                else
                {
                    for (int row = 1; row <= rows; row++)
                    {
                        for (int column = 1; column <= columns; column++)
                        {
                            Excel.Range cell = null;
                            Excel.Comment comment = null;
                            try
                            {
                                cell = (Excel.Range)context.Range.Cells[row, column];
                                comment = cell.Comment;
                                if (action == "add")
                                {
                                    if (comment != null)
                                    {
                                        skippedCount++;
                                        continue;
                                    }
                                    Excel.Comment added = null;
                                    try
                                    {
                                        added = cell.AddComment(text);
                                        changedCount++;
                                    }
                                    finally
                                    {
                                        ComHelper.Release(added);
                                    }
                                }
                                else
                                {
                                    if (comment == null)
                                    {
                                        skippedCount++;
                                        continue;
                                    }
                                    comment.Text(text, Type.Missing, Type.Missing);
                                    changedCount++;
                                }
                            }
                            finally
                            {
                                ComHelper.Release(comment, cell);
                            }
                        }
                    }
                }

                string targetRef = context.FullAddress;
                Observation observation = Observation.Create(
                    "comment_manage",
                    "批注操作 " + action + " 完成：" + changedCount + " 处已修改。",
                    changedCount > 0);
                if (changedCount > 0)
                {
                    observation.AddTargetRef(targetRef);
                }
                JObject resultData = new JObject();
                resultData["action"] = action;
                resultData["sheet"] = context.SheetName;
                resultData["address"] = context.Range.AddressLocal();
                resultData["changedCount"] = changedCount;
                resultData["skippedCount"] = skippedCount;
                ExcelUndoScope.RecordUnavailable(
                    resultData,
                    "批注内容不在区域快照中，未登记自动撤销。");
                return OperationResult.Succeed(
                    "批注操作完成：" + changedCount + " 处已修改。",
                    resultData,
                    observation);
            }
        }

        private static void CollectComments(
            Excel.Worksheet worksheet,
            Excel.Range range,
            JArray comments,
            int maxResults,
            ref bool truncated)
        {
            Excel.Range special = null;
            try
            {
                special = range.SpecialCells(Excel.XlCellType.xlCellTypeComments, Type.Missing);
            }
            catch (COMException)
            {
                special = null;
            }
            if (special == null)
            {
                return;
            }

            Excel.Areas areas = null;
            try
            {
                areas = special.Areas;
                int areaCount = areas.Count;
                for (int areaIndex = 1; areaIndex <= areaCount; areaIndex++)
                {
                    Excel.Range area = null;
                    try
                    {
                        area = (Excel.Range)areas[areaIndex];
                        int rows = area.Rows.Count;
                        int columns = area.Columns.Count;
                        for (int row = 1; row <= rows; row++)
                        {
                            for (int column = 1; column <= columns; column++)
                            {
                                Excel.Range cell = null;
                                Excel.Comment comment = null;
                                try
                                {
                                    cell = (Excel.Range)area.Cells[row, column];
                                    comment = cell.Comment;
                                    if (comment == null)
                                    {
                                        continue;
                                    }
                                    if (comments.Count >= maxResults)
                                    {
                                        truncated = true;
                                        return;
                                    }
                                    JObject item = new JObject();
                                    item["address"] = ExcelComSupport.FullAddress(
                                        worksheet.Name,
                                        cell.AddressLocal());
                                    item["text"] = comment.Text(Type.Missing, Type.Missing, Type.Missing) ?? string.Empty;
                                    item["author"] = comment.Author ?? string.Empty;
                                    item["visible"] = comment.Visible;
                                    comments.Add(item);
                                }
                                catch (COMException)
                                {
                                    // 线程批注等对象可能无法通过旧版 Comment 接口读取。
                                }
                                finally
                                {
                                    ComHelper.Release(comment, cell);
                                }
                            }
                        }
                    }
                    finally
                    {
                        ComHelper.Release(area);
                    }
                }
            }
            finally
            {
                ComHelper.Release(areas, special);
            }
        }

        private static void CollectComments(
            Excel.Worksheet worksheet,
            Excel.Range range,
            JArray comments,
            int maxResults,
            ref int count)
        {
            bool ignored = false;
            CollectComments(worksheet, range, comments, maxResults, ref ignored);
            count = comments.Count;
        }
    }
}
