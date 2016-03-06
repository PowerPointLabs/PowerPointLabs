using System.Collections.Generic;
using System.ComponentModel.Composition;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Interface;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    [Export(typeof(IStyleWorker))]
    [ExportMetadata("WorkerOrder", 0)]
    class TextStyleWorker : IStyleWorker
    {
        public IList<Shape> Execute(StyleOption option, EffectsDesigner designer, ImageItem source, Shape imageShape, Settings settings)
        {
            designer.ApplyPseudoTextWhenNoTextShapes();

            if ((option.IsUseBannerStyle 
                || option.IsUseFrostedGlassBannerStyle)
                    && (option.GetTextBoxPosition() == Position.Left
                        || (option.GetTextBoxPosition() == Position.Centre 
                            && option.GetBannerDirection() != BannerDirection.Horizontal)
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
                designer.RecoverTextWrapping(option.GetTextBoxPosition(), option.GetTextAlignment());
            }

            ApplyTextEffect(option, designer);
            designer.ApplyTextGlowEffect(option.IsUseTextGlow, option.TextGlowColor);

            return new List<Shape>();
        }

        private void ApplyTextEffect(StyleOption option, EffectsDesigner effectsDesigner)
        {
            effectsDesigner.ApplyTextEffect(option.GetFontFamily(), option.FontColor, option.FontSizeIncrease);
            effectsDesigner.ApplyTextPositionAndAlignment(option.GetTextBoxPosition(), option.GetTextAlignment());
        }
    }
}
