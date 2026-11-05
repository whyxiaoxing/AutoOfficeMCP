using System;
using System.Drawing;
using Newtonsoft.Json.Linq;
using Office;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointe
{
    /// <summary>
    /// PowerPoint COM 公共操作：活动演示文稿/幻灯片解析、占位符写入、版式与枚举映射。
    /// 所有返回的 COM 对象都由调用方在 finally 中 Release。
    /// </summary>
    internal static class PowerPointComUtil
    {
        /// <summary>尝试取得活动演示文稿；没有时返回 null，不抛异常。</summary>
        public static PowerPoint.Presentation TryGetActivePresentation(PowerPoint.Application application)
        {
            if (application == null)
            {
                return null;
            }
            try
            {
                return application.ActivePresentation;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static PowerPoint.Presentation RequireActivePresentation(PowerPoint.Application application)
        {
            PowerPoint.Presentation presentation = null;
            try
            {
                presentation = application.ActivePresentation;
            }
            catch (Exception)
            {
                presentation = null;
            }
            if (presentation == null)
            {
                throw new OfficeComException(
                    OfficeErrorCode.DOC_NOT_OPEN,
                    "当前没有打开的 PowerPoint 演示文稿。请先打开或新建一个演示文稿后重试。");
            }
            return presentation;
        }

        /// <summary>
        /// 解析 slideIndex：-1 或省略 = 当前页；0 = 第一页；1..N = 1 基索引。
        /// </summary>
        public static PowerPoint.Slide ResolveSlide(
            PowerPoint.Application application,
            PowerPoint.Presentation presentation,
            int slideIndex)
        {
            if (slideIndex < 0)
            {
                PowerPoint.Slide active = TryGetActiveSlide(application);
                if (active != null)
                {
                    return active;
                }
                int count = presentation.Slides.Count;
                if (count <= 0)
                {
                    throw new OfficeComException(OfficeErrorCode.DOC_NOT_OPEN, "演示文稿中没有幻灯片。");
                }
                return presentation.Slides[count];
            }

            int index = slideIndex == 0 ? 1 : slideIndex;
            int slideCount = presentation.Slides.Count;
            if (index < 1 || index > slideCount)
            {
                throw new OfficeComException(
                    OfficeErrorCode.INVALID_PARAMS,
                    "slideIndex 超出范围：" + slideIndex.ToString() +
                    "（当前共 " + slideCount.ToString() + " 张；-1/省略=当前，0=第一张，1..N=1 基索引）。");
            }
            return presentation.Slides[index];
        }

        private static PowerPoint.Slide TryGetActiveSlide(PowerPoint.Application application)
        {
            try
            {
                PowerPoint.DocumentWindow window = application.ActiveWindow;
                if (window == null)
                {
                    return null;
                }
                PowerPoint.View view = window.View;
                if (view == null)
                {
                    return null;
                }
                return view.Slide as PowerPoint.Slide;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static bool SetPlaceholderText(
            PowerPoint.Slide slide,
            string text,
            params PowerPoint.PpPlaceholderType[] placeholderTypes)
        {
            if (slide == null || string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            PowerPoint.Shapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                for (int index = 1; index <= shapes.Count; index++)
                {
                    PowerPoint.Shape shape = null;
                    PowerPoint.TextRange textRange = null;
                    try
                    {
                        shape = shapes[index];
                        dynamic dynamicShape = shape;
                        if (Convert.ToInt32(dynamicShape.Type) != 14)
                        {
                            // 14 = MsoShapeType.msoPlaceholder；避免引入 Microsoft.Office.Core 编译依赖。
                            continue;
                        }
                        PowerPoint.PpPlaceholderType placeholderType = shape.PlaceholderFormat.Type;
                        if (!ContainsPlaceholderType(placeholderTypes, placeholderType))
                        {
                            continue;
                        }
                        textRange = shape.TextFrame.TextRange;
                        textRange.Text = NormalizeText(text);
                        return true;
                    }
                    finally
                    {
                        ComHelper.Release(textRange, shape);
                    }
                }
            }
            finally
            {
                ComHelper.Release(shapes);
            }
            return false;
        }

        public static PowerPoint.Shape AddTextBox(
            PowerPoint.Slide slide,
            string text,
            float x,
            float y,
            float width,
            float height,
            float fontSize)
        {
            PowerPoint.Shape textBox = null;
            PowerPoint.TextRange textRange = null;
            try
            {
                dynamic shapes = slide.Shapes;
                // 1 = MsoTextOrientation.msoTextOrientationHorizontal。
                object textBoxObject = shapes.AddTextbox(1, x, y, width, height);
                textBox = textBoxObject as PowerPoint.Shape;
                if (textBox == null)
                {
                    return null;
                }
                textRange = textBox.TextFrame.TextRange;
                textRange.Text = NormalizeText(text);
                if (fontSize > 0)
                {
                    textRange.Font.Size = fontSize;
                }
                PowerPoint.Shape result = textBox;
                textBox = null;
                return result;
            }
            finally
            {
                ComHelper.Release(textRange, textBox);
            }
        }

        public static PowerPoint.Shape AddShape(
            PowerPoint.Slide slide,
            int shapeType,
            float x,
            float y,
            float width,
            float height)
        {
            dynamic shapes = slide.Shapes;
            object shapeObject = shapes.AddShape(shapeType, x, y, width, height);
            return shapeObject as PowerPoint.Shape;
        }

        /// <summary>按 shape.Id 查找形状；返回的 COM 对象由调用方 Release。找不到返回 null。</summary>
        public static PowerPoint.Shape FindShape(PowerPoint.Slide slide, int shapeId)
        {
            if (slide == null)
            {
                return null;
            }
            PowerPoint.Shapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                int count = shapes.Count;
                for (int index = 1; index <= count; index++)
                {
                    PowerPoint.Shape shape = null;
                    bool matched = false;
                    try
                    {
                        shape = shapes[index];
                        dynamic dynamicShape = shape;
                        int currentId = Convert.ToInt32((object)dynamicShape.Id);
                        if (currentId == shapeId)
                        {
                            matched = true;
                            return shape;
                        }
                    }
                    finally
                    {
                        if (!matched)
                        {
                            ComHelper.Release(shape);
                        }
                    }
                }
            }
            finally
            {
                ComHelper.Release(shapes);
            }
            return null;
        }

        public static bool ShapeHasTextFrame(PowerPoint.Shape shape)
        {
            if (shape == null)
            {
                return false;
            }
            try
            {
                dynamic dynamicShape = shape;
                return Convert.ToInt32((object)dynamicShape.HasTextFrame) != 0;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static int ShapeType(PowerPoint.Shape shape)
        {
            if (shape == null)
            {
                return 0;
            }
            try
            {
                dynamic dynamicShape = shape;
                return Convert.ToInt32((object)dynamicShape.Type);
            }
            catch (Exception)
            {
                return 0;
            }
        }

        public static string GetShapeText(PowerPoint.Shape shape)
        {
            if (shape == null || !ShapeHasTextFrame(shape))
            {
                return string.Empty;
            }
            PowerPoint.TextFrame textFrame = null;
            try
            {
                textFrame = shape.TextFrame;
                return textFrame.TextRange.Text ?? string.Empty;
            }
            catch (Exception)
            {
                return string.Empty;
            }
            finally
            {
                ComHelper.Release(textFrame);
            }
        }

        /// <summary>返回形状 id、名称、类型、坐标、尺寸、文本等可序列化信息。</summary>
        public static JObject BuildShapeInfo(PowerPoint.Shape shape)
        {
            JObject info = new JObject();
            if (shape == null)
            {
                return info;
            }
            info["id"] = SafeInt(delegate
            {
                dynamic dynamicShape = shape;
                return Convert.ToInt32((object)dynamicShape.Id);
            });
            info["name"] = SafeString(delegate { return shape.Name; });
            info["type"] = ShapeType(shape);
            info["left"] = SafeFloat(delegate { return shape.Left; });
            info["top"] = SafeFloat(delegate { return shape.Top; });
            info["width"] = SafeFloat(delegate { return shape.Width; });
            info["height"] = SafeFloat(delegate { return shape.Height; });
            bool hasText = ShapeHasTextFrame(shape);
            info["hasText"] = hasText;
            info["text"] = hasText ? GetShapeText(shape) : string.Empty;
            info["hasTable"] = SafeBool(delegate
            {
                dynamic dynamicShape = shape;
                return Convert.ToInt32((object)dynamicShape.HasTable) != 0;
            });
            info["hasChart"] = SafeBool(delegate
            {
                dynamic dynamicShape = shape;
                return Convert.ToInt32((object)dynamicShape.HasChart) != 0;
            });
            info["shapeId"] = info["id"];
            info["zOrder"] = SafeInt(delegate
            {
                dynamic dynamicShape = shape;
                return Convert.ToInt32((object)dynamicShape.ZOrderPosition);
            });
            return info;
        }

        /// <summary>
        /// 回读形状格式：填充、线条、字体和层级。仅在 includeFormat=true 时调用，
        /// 避免批量读取形状时产生大量不必要的 COM 往返。
        /// </summary>
        public static void AppendShapeFormat(JObject info, PowerPoint.Shape shape)
        {
            if (info == null || shape == null)
            {
                return;
            }

            JObject format = new JObject();
            try
            {
                dynamic fill = shape.Fill;
                int fillVisible = Convert.ToInt32((object)fill.Visible);
                format["fillVisible"] = fillVisible != 0;
                if (fillVisible != 0)
                {
                    string fillColor = SafeColor(delegate
                    {
                        return Convert.ToInt32((object)fill.ForeColor.RGB);
                    });
                    if (fillColor.Length > 0)
                    {
                        format["fillColor"] = fillColor;
                    }
                    format["fillTransparency"] = SafeFloat(delegate
                    {
                        return Convert.ToSingle((object)fill.Transparency);
                    });
                }
            }
            catch (Exception)
            {
            }

            try
            {
                dynamic line = shape.Line;
                int lineVisible = Convert.ToInt32((object)line.Visible);
                format["lineVisible"] = lineVisible != 0;
                if (lineVisible != 0)
                {
                    string lineColor = SafeColor(delegate
                    {
                        return Convert.ToInt32((object)line.ForeColor.RGB);
                    });
                    if (lineColor.Length > 0)
                    {
                        format["lineColor"] = lineColor;
                    }
                    format["lineWidth"] = SafeFloat(delegate
                    {
                        return Convert.ToSingle((object)line.Weight);
                    });
                    format["lineDash"] = SafeString(delegate
                    {
                        return MapDashStyleName(Convert.ToInt32((object)line.DashStyle));
                    });
                }
            }
            catch (Exception)
            {
            }

            if (ShapeHasTextFrame(shape))
            {
                PowerPoint.TextFrame textFrame = null;
                PowerPoint.TextRange textRange = null;
                try
                {
                    textFrame = shape.TextFrame;
                    textRange = textFrame.TextRange;
                    dynamic font = textRange.Font;
                    string fontName = SafeString(delegate { return Convert.ToString((object)font.Name); });
                    if (fontName.Length > 0)
                    {
                        format["fontName"] = fontName;
                    }
                    float fontSize = SafeFloat(delegate { return Convert.ToSingle((object)font.Size); });
                    if (fontSize > 0)
                    {
                        format["fontSize"] = fontSize;
                    }
                    string fontColor = SafeColor(delegate
                    {
                        return Convert.ToInt32((object)font.Color.RGB);
                    });
                    if (fontColor.Length > 0)
                    {
                        format["fontColor"] = fontColor;
                    }
                    format["bold"] = SafeBool(delegate
                    {
                        return Convert.ToInt32((object)font.Bold) != 0;
                    });
                }
                catch (Exception)
                {
                }
                finally
                {
                    ComHelper.Release(textRange, textFrame);
                }
            }

            if (format.Count > 0)
            {
                info["format"] = format;
                foreach (JProperty property in format.Properties())
                {
                    if (info[property.Name] == null)
                    {
                        info[property.Name] = property.Value.DeepClone();
                    }
                }
            }
        }

        public static string ColorToHex(int oleColor)
        {
            try
            {
                Color color = ColorTranslator.FromOle(oleColor);
                return "#" + color.R.ToString("X2", System.Globalization.CultureInfo.InvariantCulture) +
                       color.G.ToString("X2", System.Globalization.CultureInfo.InvariantCulture) +
                       color.B.ToString("X2", System.Globalization.CultureInfo.InvariantCulture);
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        private static string SafeColor(Func<int> reader)
        {
            try
            {
                return ColorToHex(reader());
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        public static string MapDashStyleName(int dashStyle)
        {
            switch (dashStyle)
            {
                case 2:
                    return "dot";
                case 4:
                    return "dash";
                case 5:
                    return "dashDot";
                case 7:
                    return "longDash";
                default:
                    return "solid";
            }
        }

        public static string MapAnimationEffectName(int effectType)
        {
            switch (effectType)
            {
                case 1:
                    return "appear";
                case 2:
                    return "flyIn";
                case 10:
                    return "fadeIn";
                case 22:
                    return "wipe";
                case 42:
                    return "float";
                case 53:
                    return "zoom";
                default:
                    return "effect" + effectType.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        public static int MapTriggerType(string trigger)
        {
            string key = NormalizeKey(trigger);
            switch (key)
            {
                case "withprevious":
                case "with":
                    return 2; // msoAnimTriggerWithPrevious
                case "afterprevious":
                case "after":
                    return 3; // msoAnimTriggerAfterPrevious
                case "onclick":
                case "click":
                default:
                    return 1; // msoAnimTriggerOnPageClick
            }
        }

        public static string MapTriggerName(int triggerType)
        {
            switch (triggerType)
            {
                case 2:
                    return "withPrevious";
                case 3:
                    return "afterPrevious";
                default:
                    return "onClick";
            }
        }

        public static string MapTransitionEffectName(int entryEffect)
        {
            switch (entryEffect)
            {
                case 257:
                    return "fade";
                case 769:
                    return "wipe";
                case 2817:
                    return "split";
                case 3073:
                    return "reveal";
                case 3844:
                    return "push";
                case 513:
                    return "random";
                case 0:
                    return "none";
                default:
                    return "effect" + entryEffect.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        public static string GetSlideLayoutName(PowerPoint.Slide slide)
        {
            if (slide == null)
            {
                return string.Empty;
            }
            try
            {
                return slide.CustomLayout.Name ?? string.Empty;
            }
            catch (Exception)
            {
                try
                {
                    return slide.Layout.ToString();
                }
                catch (Exception)
                {
                    return string.Empty;
                }
            }
        }

        /// <summary>读取备注页正文占位符；没有备注时返回空字符串。</summary>
        public static string GetNotesText(PowerPoint.Slide slide)
        {
            if (slide == null)
            {
                return string.Empty;
            }
            PowerPoint.SlideRange notesPage = null;
            PowerPoint.Shapes shapes = null;
            try
            {
                notesPage = slide.NotesPage;
                shapes = notesPage.Shapes;
                int count = shapes.Count;
                for (int index = 1; index <= count; index++)
                {
                    PowerPoint.Shape shape = null;
                    try
                    {
                        shape = shapes[index];
                        dynamic dynamicShape = shape;
                        if (Convert.ToInt32((object)dynamicShape.Type) == 14 &&
                            Convert.ToInt32((object)shape.PlaceholderFormat.Type) == 2)
                        {
                            return GetShapeText(shape);
                        }
                    }
                    finally
                    {
                        ComHelper.Release(shape);
                    }
                }
            }
            catch (Exception)
            {
                return string.Empty;
            }
            finally
            {
                ComHelper.Release(shapes, notesPage);
            }
            return string.Empty;
        }

        /// <summary>开始一个 PowerPoint 原生撤销组；旧版本不支持时返回 false。</summary>
        public static bool TryStartUndoEntry(PowerPoint.Application application)
        {
            if (application == null)
            {
                return false;
            }
            try
            {
                dynamic dynamicApplication = application;
                dynamicApplication.StartNewUndoEntry();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static PowerPoint.PpParagraphAlignment MapParagraphAlignment(string alignment)
        {
            string key = NormalizeKey(alignment);
            switch (key)
            {
                case "center":
                case "居中":
                    return PowerPoint.PpParagraphAlignment.ppAlignCenter;
                case "right":
                case "右对齐":
                    return PowerPoint.PpParagraphAlignment.ppAlignRight;
                case "justify":
                case "两端对齐":
                    return PowerPoint.PpParagraphAlignment.ppAlignJustify;
                case "left":
                case "左对齐":
                default:
                    return PowerPoint.PpParagraphAlignment.ppAlignLeft;
            }
        }

        /// <summary>MsoLineDashStyle 数值映射。</summary>
        public static int MapLineDashStyle(string dashStyle)
        {
            string key = NormalizeKey(dashStyle);
            switch (key)
            {
                case "dash":
                    return 4;
                case "dot":
                    return 2;
                case "dashdot":
                    return 5;
                case "longdash":
                    return 7;
                case "solid":
                default:
                    return 1;
            }
        }

        /// <summary>XlChartType 数值映射。</summary>
        public static int MapChartType(string chartType)
        {
            string key = NormalizeKey(chartType);
            switch (key)
            {
                case "column":
                case "clusteredcolumn":
                case "柱形图":
                    return 51;
                case "bar":
                case "条形图":
                    return 57;
                case "line":
                case "折线图":
                    return 4;
                case "pie":
                case "饼图":
                    return 5;
                case "doughnut":
                case "圆环图":
                    return -4120;
                case "area":
                case "面积图":
                    return 1;
                case "scatter":
                case "散点图":
                    return -4169;
                default:
                    return 51;
            }
        }

        public static PowerPoint.PpSlideLayout MapLayout(string layout)
        {
            string key = NormalizeKey(layout);
            switch (key)
            {
                case "title":
                case "titleslide":
                case "sectionheader":
                case "标题":
                    return PowerPoint.PpSlideLayout.ppLayoutTitle;
                case "titleonly":
                case "仅标题":
                    return PowerPoint.PpSlideLayout.ppLayoutTitleOnly;
                case "blank":
                case "空白":
                    return PowerPoint.PpSlideLayout.ppLayoutBlank;
                case "twocontent":
                case "twotext":
                case "twocolumn":
                case "两栏":
                    return PowerPoint.PpSlideLayout.ppLayoutTwoColumnText;
                case "comparison":
                case "比较":
                    return PowerPoint.PpSlideLayout.ppLayoutTwoColumnText;
                case "titleandcontent":
                case "titleandbody":
                case "content":
                case "标题和内容":
                    return PowerPoint.PpSlideLayout.ppLayoutText;
                default:
                    return PowerPoint.PpSlideLayout.ppLayoutText;
            }
        }

        /// <summary>
        /// MsoAutoShapeType 数值映射（避免引入 Microsoft.Office.Core 编译依赖）。
        /// </summary>
        public static int MapShapeType(string shapeType)
        {
            string key = NormalizeKey(shapeType);
            switch (key)
            {
                case "oval":
                case "circle":
                case "椭圆":
                    return 9; // msoShapeOval
                case "triangle":
                case "三角形":
                    return 7; // msoShapeIsoscelesTriangle
                case "arrow":
                case "rightarrow":
                case "右箭头":
                    return 13; // msoShapeRightArrow
                case "roundedrectangle":
                case "圆角矩形":
                    return 5; // msoShapeRoundedRectangle
                case "diamond":
                case "菱形":
                    return 4; // msoShapeDiamond
                case "chevron":
                case "箭头形":
                    return 52; // msoShapeChevron
                case "pentagon":
                case "五边形":
                    return 51; // msoShapePentagon
                case "hexagon":
                case "六边形":
                    return 10; // msoShapeHexagon
                case "rectangle":
                case "rect":
                case "矩形":
                default:
                    return 1; // msoShapeRectangle
            }
        }

        /// <summary>MsoAnimEffect 数值映射。</summary>
        public static int MapAnimationEffect(string effect)
        {
            string key = NormalizeKey(effect);
            switch (key)
            {
                case "flyin":
                case "fly":
                    return 2; // msoAnimEffectFly
                case "zoom":
                    return 53; // msoAnimEffectGrowAndTurn
                case "wipe":
                    return 22; // msoAnimEffectWipe
                case "appear":
                    return 1; // msoAnimEffectAppear
                case "float":
                    return 42; // msoAnimEffectFloat
                case "fadein":
                case "fade":
                default:
                    return 10; // msoAnimEffectFade
            }
        }

        /// <summary>PpEntryEffect 数值映射（与 Office 常量一致）。</summary>
        public static int MapTransitionEffect(string transitionType)
        {
            string key = NormalizeKey(transitionType);
            switch (key)
            {
                case "push":
                    return 3844; // ppEffectPush
                case "wipe":
                    return 769; // ppEffectWipe
                case "split":
                    return 2817; // ppEffectSplit
                case "reveal":
                    return 3073; // ppEffectReveal
                case "random":
                    return 513; // ppEffectRandom
                case "fade":
                default:
                    return 257; // ppEffectFade
            }
        }

        /// <summary>解析 #RRGGBB / 颜色名；失败返回 false。</summary>
        public static bool TryParseColor(string value, out int oleColor)
        {
            oleColor = 0;
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }
            try
            {
                Color color = ColorTranslator.FromHtml(value.Trim());
                oleColor = ColorTranslator.ToOle(color);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static void ApplyBackground(PowerPoint.Slide slide, string background)
        {
            int oleColor;
            if (slide == null || !TryParseColor(background, out oleColor))
            {
                return;
            }
            dynamic dynamicSlide = slide;
            dynamicSlide.FollowMasterBackground = 0; // msoFalse
            slide.Background.Fill.Solid();
            slide.Background.Fill.ForeColor.RGB = oleColor;
        }

        public static void ApplyFont(
            PowerPoint.TextRange textRange,
            string fontName,
            float fontSize,
            string color,
            bool bold)
        {
            if (textRange == null)
            {
                return;
            }
            if (!string.IsNullOrWhiteSpace(fontName))
            {
                textRange.Font.Name = fontName;
            }
            if (fontSize > 0)
            {
                textRange.Font.Size = fontSize;
            }
            int oleColor;
            if (TryParseColor(color, out oleColor))
            {
                textRange.Font.Color.RGB = oleColor;
            }
            dynamic font = textRange.Font;
            font.Bold = bold ? -1 : 0; // msoTrue / msoFalse
        }

        /// <summary>PowerPoint 段落分隔使用 \r；把用户输入的 \n 统一转换。</summary>
        public static string NormalizeText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }
            return text.Replace("\r\n", "\r").Replace("\n", "\r");
        }

        private static bool ContainsPlaceholderType(
            PowerPoint.PpPlaceholderType[] types,
            PowerPoint.PpPlaceholderType value)
        {
            if (types == null)
            {
                return false;
            }
            foreach (PowerPoint.PpPlaceholderType type in types)
            {
                if (type == value)
                {
                    return true;
                }
            }
            return false;
        }

        private static string NormalizeKey(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }
            return value.Trim().ToLowerInvariant()
                .Replace(" ", string.Empty)
                .Replace("-", string.Empty)
                .Replace("_", string.Empty);
        }

        internal static string SafeString(Func<string> reader)
        {
            try
            {
                return reader() ?? string.Empty;
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        internal static int SafeInt(Func<int> reader)
        {
            try
            {
                return reader();
            }
            catch (Exception)
            {
                return 0;
            }
        }

        internal static float SafeFloat(Func<float> reader)
        {
            try
            {
                float value = reader();
                if (float.IsNaN(value) || float.IsInfinity(value))
                {
                    return 0f;
                }
                return (float)Math.Round(value, 3);
            }
            catch (Exception)
            {
                return 0f;
            }
        }

        internal static bool SafeBool(Func<bool> reader)
        {
            try
            {
                return reader();
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
