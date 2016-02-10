using System;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.Utils;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        public void ApplyImageReference(string contextLink)
        {
            if (StringUtil.IsEmpty(contextLink)) return;

            RemovePreviousImageReference();
            NotesPageText = "Background image taken from " + contextLink + " on " + DateTime.Now + "\n" +
                            NotesPageText;
        }

        public void ApplyImageReferenceInsertion(string contextLink, string fontFamily, string fontColor,
            int fontSize, string textBoxColor, Alignment citationTextAlignment)
        {
            var imageRefShape = Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, SlideWidth,
                20);
            imageRefShape.TextFrame2.TextRange.Text = "Image From: " + contextLink;

            if (!StringUtil.IsEmpty(textBoxColor))
            {
                imageRefShape.Fill.BackColor.RGB
                    = Graphics.ConvertColorToRgb(StringUtil.GetColorFromHexValue(textBoxColor));
                imageRefShape.Fill.Transparency = 0.2f;
            }
            imageRefShape.TextFrame2.TextRange.TrimText().Font.Fill.ForeColor.RGB
                = Graphics.ConvertColorToRgb(StringUtil.GetColorFromHexValue(fontColor));
            imageRefShape.TextEffect.FontName = StringUtil.IsEmpty(fontFamily) ? "Tahoma" : fontFamily;
            imageRefShape.TextEffect.FontSize = fontSize;
            imageRefShape.TextEffect.Alignment = AlignmentToMsoTextEffectAlignment(citationTextAlignment);
            imageRefShape.Top = SlideHeight - imageRefShape.Height;

            AddTag(imageRefShape, Tag.ImageReference, "true");
            ChangeName(imageRefShape, EffectName.ImageReference);
        }

        private void RemovePreviousImageReference()
        {
            NotesPageText = Regex.Replace(NotesPageText, @"^Background image taken from .* on .*\n", "");
        }

        private static MsoTextEffectAlignment AlignmentToMsoTextEffectAlignment(Alignment align)
        {
            switch (align)
            {
                case Alignment.Centre:
                    return MsoTextEffectAlignment.msoTextEffectAlignmentCentered;
                case Alignment.Left:
                    return MsoTextEffectAlignment.msoTextEffectAlignmentLeft;
                case Alignment.Right:
                    return MsoTextEffectAlignment.msoTextEffectAlignmentRight;
                // case Alignment.Auto:
                default:
                    return MsoTextEffectAlignment.msoTextEffectAlignmentWordJustify;
            }
        }
    }
}
