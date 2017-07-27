using System.Collections.Generic;
using System.ComponentModel.Composition;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    [Export(typeof(IStyleOptions))]
    [ExportMetadata("StyleOrder", 2)]
    class BlurStyleOptions : BaseStyleOptions
    {
        public override List<StyleOption> GetOptionsForVariation()
        {
            var result = GetOptions();
            foreach (var styleOption in result)
            {
                styleOption.IsUseTextGlow = true;
                styleOption.TextGlowColor = "#000000";
                styleOption.FontFamily = "Segoe UI";
            }
            return UpdateStyleName(
                result,
                PictureSlidesLabText.StyleNameBlur);
        }

        public override StyleOption GetDefaultOptionForPreview()
        {
            return new StyleOption
            {
                StyleName = PictureSlidesLabText.StyleNameBlur,
                IsUseBlurStyle = true,
                BlurDegree = 85,
                TextBoxPosition = 5,
                IsUseTextGlow = true,
                TextGlowColor = "#000000",
                FontFamily = "Segoe UI"
            };
        }
    }
}
