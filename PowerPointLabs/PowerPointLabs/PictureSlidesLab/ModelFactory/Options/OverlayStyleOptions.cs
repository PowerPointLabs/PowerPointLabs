using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPointLabs.PictureSlidesLab.Model;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    class OverlayStyleOptions : BaseStyleOptions
    {
        public override List<StyleOption> GetOptionsForVariation()
        {
            var result = GetOptionsWithSuitableFontColor();
            foreach (var styleOption in result)
            {
                styleOption.Transparency = 35;
            }
            return UpdateStyleName(result,
                TextCollection.PictureSlidesLabText.StyleNameOverlay);
        }

        public override StyleOption GetDefaultOptionForPreview()
        {
            return new StyleOption
            {
                StyleName = TextCollection.PictureSlidesLabText.StyleNameOverlay,
                IsUseOverlayStyle = true,
                Transparency = 35,
                OverlayColor = "#007FFF", // blue
                IsUseSpecialEffectStyle = true,
                SpecialEffect = 0
            };
        }
    }
}
