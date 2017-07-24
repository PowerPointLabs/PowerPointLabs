using System.Collections.Generic;
using System.ComponentModel.Composition;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    [Export(typeof(IStyleOptions))]
    [ExportMetadata("StyleOrder", 7)]
    class SpecialEffectStyleOptions : BaseStyleOptions
    {
        public override List<StyleOption> GetOptionsForVariation()
        {
            var result = GetOptions();
            foreach (var styleOption in result)
            {
                styleOption.IsUseTextGlow = true;
                styleOption.TextGlowColor = "#000000";
                styleOption.FontFamily = "Arial Black";
            }
            return UpdateStyleName(
                result,
                PictureSlidesLabText.StyleNameSpecialEffect);
        }

        public override StyleOption GetDefaultOptionForPreview()
        {
            return new StyleOption
            {
                StyleName = PictureSlidesLabText.StyleNameSpecialEffect,
                IsUseSpecialEffectStyle = true,
                SpecialEffect = 0,
                IsUseTextGlow = true,
                TextGlowColor = "#000000",
                FontFamily = "Arial Black"
            };
        }
    }
}
