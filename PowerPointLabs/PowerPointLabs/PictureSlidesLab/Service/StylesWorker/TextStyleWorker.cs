using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.Effect;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    class TextStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source,
            Shape imageShape)
        {
            designer.ApplyPseudoTextWhenNoTextShapes();

            if (option.IsUseBannerStyle
                && (option.GetTextBoxPosition() == Position.Left
                    || option.GetTextBoxPosition() == Position.Centre
                    || option.GetTextBoxPosition() == Position.Right))
            {
                designer.ApplyTextWrapping();
            }
            else if (option.IsUseCircleStyle
                     || option.IsUseOutlineStyle)
            {
                designer.ApplyTextWrapping();
            }
            else
            {
                designer.RecoverTextWrapping();
            }

            ApplyTextEffect(option, designer);
            designer.ApplyTextGlowEffect(option.IsUseTextGlow, option.TextGlowColor);

            return new List<Shape>();
        }

        private void ApplyTextEffect(StyleOption option, EffectsDesigner effectsDesigner)
        {
            if (option.IsUseTextFormat)
            {
                effectsDesigner.ApplyTextEffect(option.GetFontFamily(), option.FontColor, option.FontSizeIncrease);
                effectsDesigner.ApplyTextPositionAndAlignment(option.GetTextBoxPosition(), option.GetTextAlignment());
            }
            else
            {
                effectsDesigner.ApplyOriginalTextEffect();
                effectsDesigner.ApplyTextPositionAndAlignment(Position.Original, Alignment.Auto);
            }
        }
    }
}
