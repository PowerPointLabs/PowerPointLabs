using System.Collections.Generic;
using System.ComponentModel.Composition;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    [Export(typeof(IStyleOptions))]
    [ExportMetadata("StyleOrder", 8)]
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
            }
            return UpdateStyleName(
                result,
                TextCollection.PictureSlidesLabText.StyleNameFrame);
        }

        public override StyleOption GetDefaultOptionForPreview()
        {
            return new StyleOption()
            {
                StyleName = TextCollection.PictureSlidesLabText.StyleNameFrame,
                IsUseFrameStyle = true,
                IsUseTextGlow = true,
                TextGlowColor = "#000000"
            };
        }
    }
}
