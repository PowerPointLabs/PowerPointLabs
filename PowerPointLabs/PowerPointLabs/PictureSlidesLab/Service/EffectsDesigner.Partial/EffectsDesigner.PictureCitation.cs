using System;
using System.Text.RegularExpressions;

using Microsoft.Office.Core;

using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.Utils;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        public const string RegexForPictureCitation = @"^\[\[Picture taken from .* on .*\]\]\n";

        public void ApplyImageReferenceToSlideNote(string source)
        {
            if (StringUtil.IsEmpty(source))
            {
                return;
            }

            RemoveImageReference();
            NotesPageText = "[[Picture taken from " + source + " on " + DateTime.Now + "]]\n" +
                            NotesPageText;
        }

        public void ApplyImageReferenceInsertion(string source, string fontFamily, string fontColor,
            int fontSize, string textBoxColor, Alignment citationTextAlignment)
        {
            Microsoft.Office.Interop.PowerPoint.Shape imageRefShape = Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, SlideWidth,
                20);
            imageRefShape.TextFrame2.TextRange.Text = "Picture taken from: " + source;

            if (!StringUtil.IsEmpty(textBoxColor))
            {
                imageRefShape.Fill.BackColor.RGB
                    = GraphicsUtil.ConvertColorToRgb(StringUtil.GetColorFromHexValue(textBoxColor));
                imageRefShape.Fill.Transparency = 0.2f;
            }
            imageRefShape.TextFrame2.TextRange.TrimText().Font.Fill.ForeColor.RGB
                = GraphicsUtil.ConvertColorToRgb(StringUtil.GetColorFromHexValue(fontColor));
            imageRefShape.TextEffect.FontName = StringUtil.IsEmpty(fontFamily) ? "Tahoma" : fontFamily;
            imageRefShape.TextEffect.FontSize = fontSize;
            imageRefShape.TextEffect.Alignment = AlignmentToMsoTextEffectAlignment(citationTextAlignment);
            imageRefShape.Top = SlideHeight - imageRefShape.Height;

            AddTag(imageRefShape, Tag.ImageReference, "true");
            ChangeName(imageRefShape, EffectName.ImageReference);
        }

        public void RemoveImageReference()
        {
            NotesPageText = Regex.Replace(NotesPageText, RegexForPictureCitation, "");
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
