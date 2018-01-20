﻿using System.Collections.Generic;
using System.ComponentModel.Composition;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    [Export(typeof(IStyleOptions))]
    [ExportMetadata("StyleOrder", 8)]
    class OverlayStyleOptions : BaseStyleOptions
    {
        public override List<StyleOption> GetOptionsForVariation()
        {
            List<StyleOption> result = GetOptionsWithSuitableFontColor();
            foreach (StyleOption styleOption in result)
            {
                styleOption.OverlayTransparency = 35;
                styleOption.FontFamily = "Trebuchet MS";
            }
            return UpdateStyleName(result,
                PictureSlidesLabText.StyleNameOverlay);
        }

        public override StyleOption GetDefaultOptionForPreview()
        {
            return new StyleOption
            {
                StyleName = PictureSlidesLabText.StyleNameOverlay,
                IsUseOverlayStyle = true,
                OverlayTransparency = 35,
                OverlayColor = "#007FFF", // blue
                IsUseSpecialEffectStyle = true,
                SpecialEffect = 0,
                FontFamily = "Trebuchet MS"
            };
        }
    }
}
