using System.Collections.Generic;
using System.ComponentModel.Composition;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    [Export(typeof(IStyleOptions))]
    [ExportMetadata("StyleOrder", 6)]
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
