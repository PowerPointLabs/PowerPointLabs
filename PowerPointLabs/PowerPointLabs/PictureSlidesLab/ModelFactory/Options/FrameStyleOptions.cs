using System.Collections.Generic;
using System.ComponentModel.Composition;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    [Export(typeof(IStyleOptions))]
    [ExportMetadata("StyleOrder", 10)]
    class FrameStyleOptions : BaseStyleOptions
    {
        public override List<StyleOption> GetOptionsForVariation()
        {
            var result = GetOptions();
            foreach (var styleOption in result)
            {
                styleOption.IsUseFrameStyle = true;
                styleOption.IsUseTextGlow = true;
                styleOption.TextGlowColor = "#000000";
                styleOption.FontFamily = "Courier New";
            }
            return UpdateStyleName(
                result,
                PictureSlidesLabText.StyleNameFrame);
        }

        public override StyleOption GetDefaultOptionForPreview()
        {
            return new StyleOption()
            {
                StyleName = PictureSlidesLabText.StyleNameFrame,
                IsUseFrameStyle = true,
                IsUseTextGlow = true,
                TextGlowColor = "#000000",
                FontFamily = "Courier New"
            };
        }
    }
}
