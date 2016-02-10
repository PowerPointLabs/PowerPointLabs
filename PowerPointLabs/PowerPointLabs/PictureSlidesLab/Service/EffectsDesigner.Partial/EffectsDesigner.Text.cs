using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        // apply text formats to textbox & placeholer
        public void ApplyTextEffect(string fontFamily, string fontColor, int fontSizeToIncrease)
        {
            foreach (PowerPoint.Shape shape in Shapes)
            {
                if ((shape.Type != MsoShapeType.msoPlaceholder
                        && shape.Type != MsoShapeType.msoTextBox)
                        || shape.TextFrame.HasText == MsoTriState.msoFalse)
                {
                    continue;
                }

                AddTag(shape, Tag.OriginalFillVisible, BoolUtil.ToBool(shape.Fill.Visible).ToString());
                shape.Fill.Visible = MsoTriState.msoFalse;

                AddTag(shape, Tag.OriginalLineVisible, BoolUtil.ToBool(shape.Line.Visible).ToString());
                shape.Line.Visible = MsoTriState.msoFalse;

                var font = shape.TextFrame2.TextRange.TrimText().Font;

                AddTag(shape, Tag.OriginalFontColor, StringUtil.GetHexValue(Graphics.ConvertRgbToColor(font.Fill.ForeColor.RGB)));
                font.Fill.ForeColor.RGB = Graphics.ConvertColorToRgb(StringUtil.GetColorFromHexValue(fontColor));

                AddTag(shape, Tag.OriginalFontFamily, font.Name);
                if (StringUtil.IsEmpty(fontFamily))
                {
                    shape.TextEffect.FontName = shape.Tags[Tag.OriginalFontFamily];
                    shape.Tags.Add(Tag.OriginalFontFamily, "");
                }
                else
                {
                    shape.TextEffect.FontName = fontFamily;
                }

                if (StringUtil.IsEmpty(shape.Tags[Tag.OriginalFontSize]))
                {
                    shape.Tags.Add(Tag.OriginalFontSize, shape.TextEffect.FontSize.ToString(CultureInfo.InvariantCulture));
                }
                shape.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;
                shape.TextEffect.FontSize = float.Parse(shape.Tags[Tag.OriginalFontSize]) + fontSizeToIncrease;
            }
        }

        public void ApplyTextGlowEffect(bool isUseTextGlow, string textGlowColor)
        {
            foreach (PowerPoint.Shape shape in Shapes)
            {
                if ((shape.Type != MsoShapeType.msoPlaceholder
                     && shape.Type != MsoShapeType.msoTextBox)
                    || shape.TextFrame.HasText == MsoTriState.msoFalse)
                {
                    continue;
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
        }

        public void ApplyOriginalTextEffect()
        {
            foreach (PowerPoint.Shape shape in Shapes)
            {
                if ((shape.Type != MsoShapeType.msoPlaceholder
                        && shape.Type != MsoShapeType.msoTextBox)
                        || shape.TextFrame.HasText == MsoTriState.msoFalse)
                {
                    continue;
                }

                if (StringUtil.IsNotEmpty(shape.Tags[Tag.OriginalFillVisible]))
                {
                    shape.Fill.Visible = BoolUtil.ToMsoTriState(bool.Parse(shape.Tags[Tag.OriginalFillVisible]));
                    shape.Tags.Add(Tag.OriginalFillVisible, "");
                }
                if (StringUtil.IsNotEmpty(shape.Tags[Tag.OriginalLineVisible]))
                {
                    shape.Line.Visible = BoolUtil.ToMsoTriState(bool.Parse(shape.Tags[Tag.OriginalLineVisible]));
                    shape.Tags.Add(Tag.OriginalLineVisible, "");
                }

                var font = shape.TextFrame2.TextRange.TrimText().Font;

                if (StringUtil.IsNotEmpty(shape.Tags[Tag.OriginalFontColor]))
                {
                    font.Fill.ForeColor.RGB
                        = Graphics.ConvertColorToRgb(StringUtil.GetColorFromHexValue(shape.Tags[Tag.OriginalFontColor]));
                    shape.Tags.Add(Tag.OriginalFontColor, "");
                }
                if (StringUtil.IsNotEmpty(shape.Tags[Tag.OriginalFontFamily]))
                {
                    font.Name = shape.Tags[Tag.OriginalFontFamily];
                    shape.Tags.Add(Tag.OriginalFontFamily, "");
                }
                if (StringUtil.IsNotEmpty(shape.Tags[Tag.OriginalFontSize]))
                {
                    shape.TextEffect.FontSize = float.Parse(shape.Tags[Tag.OriginalFontSize]);
                    shape.Tags.Add(Tag.OriginalFontSize, "");
                }
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

            if (!isTextShapesEmpty) return;

            try
            {
                Shapes.AddTitle().TextFrame2.TextRange.Text = "Picture Slides Lab";
            }
            catch
            {
                // title already exist
                foreach (PowerPoint.Shape shape in Shapes)
                {
                    try
                    {
                        switch (shape.PlaceholderFormat.Type)
                        {
                            case PowerPoint.PpPlaceholderType.ppPlaceholderTitle:
                            case PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle:
                            case PowerPoint.PpPlaceholderType.ppPlaceholderVerticalTitle:
                                shape.TextFrame2.TextRange.Text = "Picture Slides Lab";
                                break;
                        }
                    }
                    catch (COMException)
                    {
                        // non-placeholder shapes don't have PlaceholderFormat
                        // and will cause exception
                    }
                }
            }
        }

        public void ApplyTextWrapping()
        {
            new TextBoxes(Shapes.Range(), SlideWidth, SlideHeight)
                .StartTextWrapping();
        }

        public void RecoverTextWrapping()
        {
            new TextBoxes(Shapes.Range(), SlideWidth, SlideHeight)
                .RecoverTextWrapping();
        }
    }
}
