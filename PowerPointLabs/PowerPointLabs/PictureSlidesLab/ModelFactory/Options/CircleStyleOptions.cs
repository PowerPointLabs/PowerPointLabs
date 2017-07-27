using System.Collections.Generic;
using System.ComponentModel.Composition;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    [Export(typeof(IStyleOptions))]
    [ExportMetadata("StyleOrder", 11)]
    class CircleStyleOptions : BaseStyleOptions
    {
        public override List<StyleOption> GetOptionsForVariation()
        {
            var result = GetOptionsWithSuitableFontColor();
            foreach (var styleOption in result)
            {
                styleOption.IsUseCircleStyle = true;
                styleOption.CircleTransparency = 25;
                styleOption.FontFamily = "Impact";
            }
            return UpdateStyleName(
                result,
                PictureSlidesLabText.StyleNameCircle);
        }

        public override StyleOption GetDefaultOptionForPreview()
        {
            return new StyleOption()
            {
                StyleName = PictureSlidesLabText.StyleNameCircle,
                IsUseCircleStyle = true,
                FontColor = "#000000",
                CircleTransparency = 25,
                FontFamily = "Impact"
            };
        }
    }
}
