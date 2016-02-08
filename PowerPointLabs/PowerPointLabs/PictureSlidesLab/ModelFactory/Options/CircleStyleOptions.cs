using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPointLabs.PictureSlidesLab.Model;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    class CircleStyleOptions : BaseStyleOptions
    {
        public override List<StyleOption> GetOptionsForVariation()
        {
            var result = GetOptionsWithSuitableFontColor();
            foreach (var styleOption in result)
            {
                styleOption.IsUseCircleStyle = true;
                styleOption.CircleTransparency = 25;
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
                CircleTransparency = 25
            };
        }
    }
}
