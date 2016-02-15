using System.Collections.Generic;
using System.ComponentModel.Composition;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    [Export(typeof(IStyleOptions))]
    [ExportMetadata("StyleOrder", 9)]
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
                TextCollection.PictureSlidesLabText.StyleNameCircle);
        }

        public override StyleOption GetDefaultOptionForPreview()
        {
            return new StyleOption()
            {
                StyleName = TextCollection.PictureSlidesLabText.StyleNameCircle,
                IsUseCircleStyle = true,
                FontColor = "#000000",
                CircleTransparency = 25,
                FontFamily = "Impact"
            };
        }
    }
}
