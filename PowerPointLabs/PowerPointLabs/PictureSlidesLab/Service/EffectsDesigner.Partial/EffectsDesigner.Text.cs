using System;
using System.Globalization;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.PictureSlidesLab.Util;
using PowerPointLabs.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        // apply text formats to textbox & placeholer
        public void ApplyTextEffect(string fontFamily, string fontColor, int fontSizeToIncrease, int fontTransparency)
        {
            var shape = ShapeUtil.GetTextShapeToProcess(Shapes);
            if (shape == null)
            {
                return;
            }

            // a hack to enable the textbox layout
            shape.Width = shape.Width;

            shape.Fill.Visible = MsoTriState.msoFalse;
            shape.Line.Visible = MsoTriState.msoFalse;

            var font = shape.TextFrame2.TextRange.TrimText().Font;

            if (!string.IsNullOrEmpty(fontColor))
            {
                font.Fill.ForeColor.RGB = Graphics.ConvertColorToRgb(StringUtil.GetColorFromHexValue(fontColor));
            }

            if (!StringUtil.IsEmpty(fontFamily))
            {
                shape.TextEffect.FontName = fontFamily;
            }

            if (StringUtil.IsEmpty(shape.Tags[Tag.OriginalFontSize]))
            {
                shape.Tags.Add(Tag.OriginalFontSize, shape.TextEffect.FontSize.ToString(CultureInfo.InvariantCulture));
            }

            if (fontSizeToIncrease != -1)
            {
                shape.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;
                shape.TextEffect.FontSize = float.Parse(shape.Tags[Tag.OriginalFontSize]) + fontSizeToIncrease;
            }

            if (fontTransparency != -1)
            {
                font.Fill.Transparency = (float)fontTransparency / 100;
            }
        }

        public void ApplyTextGlowEffect(bool isUseTextGlow, string textGlowColor)
        {
            var shape = ShapeUtil.GetTextShapeToProcess(Shapes);
            if (shape == null)
            {
                return;
            }

            if (isUseTextGlow)
            {
                shape.TextFrame2.TextRange.Font.Glow.Radius = 8;
                shape.TextFrame2.TextRange.Font.Glow.Color.RGB =
                    Graphics.ConvertColorToRgb(StringUtil.GetColorFromHexValue(textGlowColor));
                shape.TextFrame2.TextRange.Font.Glow.Transparency = 0.6f;
            }
            else
            {
                shape.TextFrame2.TextRange.Font.Glow.Radius = 0;
                shape.TextFrame2.TextRange.Font.Glow.Color.RGB =
                    Graphics.ConvertColorToRgb(StringUtil.GetColorFromHexValue(textGlowColor));
                shape.TextFrame2.TextRange.Font.Glow.Transparency = 0.0f;
            }
        }

        public void ApplyTextPositionAndAlignment(Position pos, Alignment alignment)
        {
            new TextBoxes(Shapes.Range(), SlideWidth, SlideHeight)
                .SetPosition(pos)
                .SetAlignment(alignment)
                .StartBoxing();
        }

        public void ApplyPseudoTextWhenNoTextShapes()
        {
            var isTextShapesEmpty = new TextBoxes(
                Shapes.Range(), SlideWidth, SlideHeight)
                .IsTextShapesEmpty();

            if (!isTextShapesEmpty)
            {
                return;
            }

            try
            {
                if (_slide.Layout == PowerPoint.PpSlideLayout.ppLayoutBlank)
                {
                    _slide.Layout = PowerPoint.PpSlideLayout.ppLayoutTitleOnly;
                }
                Shapes.AddTitle().TextFrame2.TextRange.Text = "Picture Slides Lab";
            }
            catch
            {
                // title already exist
                foreach (PowerPoint.Shape shape in Shapes)
                {
                    try
                    {
                        if (shape.Type != MsoShapeType.msoPlaceholder)
                        {
                            continue;
                        }

                        switch (shape.PlaceholderFormat.Type)
                        {
                            case PowerPoint.PpPlaceholderType.ppPlaceholderTitle:
                            case PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle:
                            case PowerPoint.PpPlaceholderType.ppPlaceholderVerticalTitle:
                                shape.TextFrame2.TextRange.Text = "Picture Slides Lab";
                                break;
                        }
                    }
                    catch (Exception e)
                    {
                        Logger.LogException(e, "ApplyPseudoTextWhenNoTextShapes");
                    }
                }
            }
        }

        public void ApplyTextWrapping()
        {
            new TextBoxes(Shapes.Range(), SlideWidth, SlideHeight)
                .StartTextWrapping();
        }

        public void RecoverTextWrapping(Position position, Alignment alignment)
        {
            if (position == Position.NoEffect
                || alignment == Alignment.NoEffect)
            {
                return;
            }

            new TextBoxes(Shapes.Range(), SlideWidth, SlideHeight)
                .RecoverTextWrapping();
        }
    }
}
