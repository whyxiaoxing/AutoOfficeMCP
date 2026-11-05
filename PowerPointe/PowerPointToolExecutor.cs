using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Office;
using Office.Contracts;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointe
{
    /// <summary>
    /// PowerPoint MCP 工具执行器（54 个工具，按读取/写入/格式/动画/表格/图片/版面自检划分）。
    /// 所有 COM 调用在 OfficeSession.RunPowerPoint 的 STA 主会话内执行；
    /// 工具名是唯一事实源，描述符与 CanExecute 共用同一组键。
    /// </summary>
    public sealed partial class PowerPointToolExecutor : IToolExecutor
    {
        /// <summary>工具名静态清单：供 T07 注册表漂移门禁与人工核对使用。</summary>
        private static readonly string[] ToolNames = new string[]
        {
            "ppt_insert_slide",
            "ppt_insert_text",
            "ppt_insert_shape",
            "ppt_insert_table",
            "ppt_format_slide",
            "ppt_create_slides",
            "ppt_add_animation",
            "ppt_apply_transition",
            "ppt_beautify_slides",
            "ppt_delete_slide",
            "ppt_duplicate_slide",
            "ppt_move_slide",
            "ppt_set_slide_layout",
            "ppt_apply_theme",
            "ppt_add_speaker_notes",
            "ppt_execute_vba",
            "ppt_get_presentation_info",
            "ppt_get_slide_content",
            "ppt_list_shapes",
            "ppt_insert_image",
            "ppt_set_shape_format",
            "ppt_format_text",
            "ppt_set_shape_text",
            "ppt_delete_shape",
            "ppt_move_resize_shape",
            "ppt_export_images",
            "ppt_export_pdf",
            "ppt_get_thumbnail",
            "ppt_list_layouts",
            "ppt_list_themes",
            "ppt_insert_chart",
            "ppt_set_bullet",
            "ppt_set_paragraph",
            "ppt_hide_slide",
            "ppt_reorder_slides",
            "ppt_clear_slide",
            "ppt_set_slide_size",
            "ppt_set_zorder",
            "ppt_list_animations",
            "ppt_delete_animation",
            "ppt_get_transition",
            "ppt_get_table",
            "ppt_set_table_cell",
            "ppt_format_table",
            "ppt_group_shapes",
            "ppt_ungroup_shapes",
            "ppt_align_shapes",
            "ppt_crop_image",
            "ppt_replace_image",
            "ppt_set_image_format",
            "ppt_copy_slide_from_file",
            "ppt_validate_layout",
            "ppt_insert_table_from_excel",
            "ppt_insert_chart_from_excel"
        };

        /// <summary>会修改演示文稿的工具；执行前后采集原生撤销组与状态签名。</summary>
        private static readonly HashSet<string> WriteToolNames =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "ppt_insert_slide",
                "ppt_insert_text",
                "ppt_insert_shape",
                "ppt_insert_table",
                "ppt_format_slide",
                "ppt_create_slides",
                "ppt_add_animation",
                "ppt_apply_transition",
                "ppt_beautify_slides",
                "ppt_delete_slide",
                "ppt_duplicate_slide",
                "ppt_move_slide",
                "ppt_set_slide_layout",
                "ppt_apply_theme",
                "ppt_add_speaker_notes",
                "ppt_insert_image",
                "ppt_set_shape_format",
                "ppt_format_text",
                "ppt_set_shape_text",
                "ppt_delete_shape",
                "ppt_move_resize_shape",
                "ppt_insert_chart",
                "ppt_set_bullet",
                "ppt_set_paragraph",
                "ppt_hide_slide",
                "ppt_reorder_slides",
                "ppt_clear_slide",
                "ppt_set_slide_size",
                "ppt_set_zorder",
                "ppt_delete_animation",
                "ppt_set_table_cell",
                "ppt_format_table",
                "ppt_group_shapes",
                "ppt_ungroup_shapes",
                "ppt_align_shapes",
                "ppt_crop_image",
                "ppt_replace_image",
                "ppt_set_image_format",
                "ppt_copy_slide_from_file",
                "ppt_insert_table_from_excel",
                "ppt_insert_chart_from_excel"
            };

        private static readonly Dictionary<string, Func<PowerPoint.Application, JObject, OperationResult>> Handlers =
            BuildHandlers();

        public string AppType
        {
            get { return "ppt"; }
        }

        public bool CanExecute(string toolName)
        {
            return !string.IsNullOrWhiteSpace(toolName) && Handlers.ContainsKey(toolName);
        }

        public ToolResult Execute(string toolName, JObject arguments)
        {
            if (!CanExecute(toolName))
            {
                return ToolResult.Fail(
                    "INVALID_PARAMS",
                    "PowerPoint 执行器不认领工具「" + (toolName ?? string.Empty) + "」。",
                    false);
            }

            JObject safeArguments = arguments ?? new JObject();
            Func<PowerPoint.Application, JObject, OperationResult> handler = Handlers[toolName];
            bool writeTool = WriteToolNames.Contains(toolName);
            OperationResult result = OfficeSession.RunPowerPoint(
                toolName,
                delegate(PowerPoint.Application application)
                {
                    PowerPoint.Presentation presentation = null;
                    JObject beforeState = null;
                    string documentName = string.Empty;
                    bool undoGroupStarted = false;
                    try
                    {
                        if (writeTool)
                        {
                            presentation = PowerPointComUtil.TryGetActivePresentation(application);
                            if (presentation != null)
                            {
                                undoGroupStarted = PowerPointComUtil.TryStartUndoEntry(application);
                                try
                                {
                                    beforeState = Office.Undo.PowerPointUndoEngine.Capture(presentation);
                                    documentName = beforeState.Value<string>("document") ?? string.Empty;
                                }
                                catch (Exception captureError)
                                {
                                    Console.Error.WriteLine("[ppt] 采集撤销前状态失败：" + captureError.Message);
                                    beforeState = null;
                                }
                            }
                        }

                        OperationResult operationResult;
                        try
                        {
                            operationResult = handler(application, safeArguments);
                        }
                        catch (Exception handlerError)
                        {
                            if (writeTool)
                            {
                                TryRegisterUndo(
                                    presentation, beforeState, documentName, undoGroupStarted, null);
                            }
                            throw handlerError;
                        }

                        if (writeTool)
                        {
                            TryRegisterUndo(
                                presentation, beforeState, documentName, undoGroupStarted, operationResult);
                        }
                        return operationResult;
                    }
                    finally
                    {
                        ComHelper.Release(presentation);
                    }
                });
            if (result == null)
            {
                return ToolResult.Fail(
                    "INTERNAL_ERROR",
                    "PowerPoint 工具「" + toolName + "」未返回结果。",
                    false);
            }
            return result.ToToolResult();
        }

        private static Dictionary<string, Func<PowerPoint.Application, JObject, OperationResult>> BuildHandlers()
        {
            Dictionary<string, Func<PowerPoint.Application, JObject, OperationResult>> handlers =
                new Dictionary<string, Func<PowerPoint.Application, JObject, OperationResult>>(
                    StringComparer.Ordinal);
            handlers.Add("ppt_insert_slide", ExecuteInsertSlide);
            handlers.Add("ppt_insert_text", ExecuteInsertText);
            handlers.Add("ppt_insert_shape", ExecuteInsertShape);
            handlers.Add("ppt_insert_table", ExecuteInsertTable);
            handlers.Add("ppt_format_slide", ExecuteFormatSlide);
            handlers.Add("ppt_create_slides", ExecuteCreateSlides);
            handlers.Add("ppt_add_animation", ExecuteAddAnimation);
            handlers.Add("ppt_apply_transition", ExecuteApplyTransition);
            handlers.Add("ppt_beautify_slides", ExecuteBeautifySlides);
            handlers.Add("ppt_delete_slide", ExecuteDeleteSlide);
            handlers.Add("ppt_duplicate_slide", ExecuteDuplicateSlide);
            handlers.Add("ppt_move_slide", ExecuteMoveSlide);
            handlers.Add("ppt_set_slide_layout", ExecuteSetSlideLayout);
            handlers.Add("ppt_apply_theme", ExecuteApplyTheme);
            handlers.Add("ppt_add_speaker_notes", ExecuteAddSpeakerNotes);
            handlers.Add("ppt_execute_vba", ExecuteVba);
            handlers.Add("ppt_get_presentation_info", ExecuteGetPresentationInfo);
            handlers.Add("ppt_get_slide_content", ExecuteGetSlideContent);
            handlers.Add("ppt_list_shapes", ExecuteListShapes);
            handlers.Add("ppt_insert_image", ExecuteInsertImage);
            handlers.Add("ppt_set_shape_format", ExecuteSetShapeFormat);
            handlers.Add("ppt_format_text", ExecuteFormatText);
            handlers.Add("ppt_set_shape_text", ExecuteSetShapeText);
            handlers.Add("ppt_delete_shape", ExecuteDeleteShape);
            handlers.Add("ppt_move_resize_shape", ExecuteMoveResizeShape);
            handlers.Add("ppt_export_images", ExecuteExportImages);
            handlers.Add("ppt_export_pdf", ExecuteExportPdf);
            handlers.Add("ppt_get_thumbnail", ExecuteGetThumbnail);
            handlers.Add("ppt_list_layouts", ExecuteListLayouts);
            handlers.Add("ppt_list_themes", ExecuteListThemes);
            handlers.Add("ppt_insert_chart", ExecuteInsertChart);
            handlers.Add("ppt_set_bullet", ExecuteSetBullet);
            handlers.Add("ppt_set_paragraph", ExecuteSetParagraph);
            handlers.Add("ppt_hide_slide", ExecuteHideSlide);
            handlers.Add("ppt_reorder_slides", ExecuteReorderSlides);
            handlers.Add("ppt_clear_slide", ExecuteClearSlide);
            handlers.Add("ppt_set_slide_size", ExecuteSetSlideSize);
            handlers.Add("ppt_set_zorder", ExecuteSetZOrder);
            handlers.Add("ppt_list_animations", ExecuteListAnimations);
            handlers.Add("ppt_delete_animation", ExecuteDeleteAnimation);
            handlers.Add("ppt_get_transition", ExecuteGetTransition);
            handlers.Add("ppt_get_table", ExecuteGetTable);
            handlers.Add("ppt_set_table_cell", ExecuteSetTableCell);
            handlers.Add("ppt_format_table", ExecuteFormatTable);
            handlers.Add("ppt_group_shapes", ExecuteGroupShapes);
            handlers.Add("ppt_ungroup_shapes", ExecuteUngroupShapes);
            handlers.Add("ppt_align_shapes", ExecuteAlignShapes);
            handlers.Add("ppt_crop_image", ExecuteCropImage);
            handlers.Add("ppt_replace_image", ExecuteReplaceImage);
            handlers.Add("ppt_set_image_format", ExecuteSetImageFormat);
            handlers.Add("ppt_copy_slide_from_file", ExecuteCopySlideFromFile);
            handlers.Add("ppt_validate_layout", ExecuteValidateLayout);
            handlers.Add("ppt_insert_table_from_excel", ExecuteInsertTableFromExcel);
            handlers.Add("ppt_insert_chart_from_excel", ExecuteInsertChartFromExcel);
            return handlers;
        }

        internal static OperationResult Success(
            string summary,
            string kind,
            bool changed,
            string targetRef,
            JObject data)
        {
            Observation observation = Observation.Create(kind, summary, changed);
            if (!string.IsNullOrWhiteSpace(targetRef))
            {
                observation.AddTargetRef(targetRef);
            }
            return OperationResult.Succeed(summary, data, observation);
        }

        internal static OperationResult Fail(OfficeErrorCode code, string message)
        {
            return OperationResult.Failed(code, message);
        }

        /// <summary>
        /// 写工具结束后采集执行后状态并登记原生撤销节点。
        /// 状态签名未变化或执行前快照缺失时不登记，避免产生“撤销会误伤上一操作”的假节点。
        /// </summary>
        private static void TryRegisterUndo(
            PowerPoint.Presentation presentation,
            JObject beforeState,
            string documentName,
            bool undoGroupStarted,
            OperationResult result)
        {
            if (!undoGroupStarted || presentation == null || beforeState == null)
            {
                return;
            }

            JObject afterState;
            try
            {
                afterState = Office.Undo.PowerPointUndoEngine.Capture(presentation);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[ppt] 采集撤销后状态失败：" + ex.Message);
                return;
            }

            string summary = result == null ? "PowerPoint 操作" : result.UserMessage;
            Office.Undo.SnapshotUndoNode node = Office.Undo.PowerPointUndoEngine.BuildNode(
                beforeState,
                afterState,
                summary);
            if (node == null)
            {
                return;
            }

            try
            {
                OfficeSession.UndoStack.Push(node);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[ppt] 登记撤销节点失败：" + ex.Message);
                return;
            }

            if (result != null)
            {
                if (result.Observation == null)
                {
                    result.Observation = Observation.Create(
                        "write",
                        string.IsNullOrWhiteSpace(summary) ? "PowerPoint 操作" : summary,
                        true);
                }
                result.Observation.AddUndoUnit(node.Id);
                if (result.Data == null)
                {
                    result.Data = new JObject();
                }
                result.Data["undoRestorable"] = true;
                result.Data["undoNodeId"] = node.Id;
                result.Data["documentName"] = documentName;
            }
        }

        internal static string SlideRef(PowerPoint.Presentation presentation, PowerPoint.Slide slide)
        {
            if (presentation == null || slide == null)
            {
                return string.Empty;
            }
            try
            {
                return "ppt:" + presentation.Name + "!Slide" + slide.SlideIndex.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
