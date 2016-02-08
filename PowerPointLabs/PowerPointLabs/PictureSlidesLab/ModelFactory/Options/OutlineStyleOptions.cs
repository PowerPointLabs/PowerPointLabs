using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPointLabs.PictureSlidesLab.Model;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    class OutlineStyleOptions : BaseStyleOptions
    {
        public override List<StyleOption> GetOptionsForVariation()
        {
            var result = GetOptionsWithSuitableFontColor();
            foreach (var styleOption in result)
            {
                styleOption.IsUseOutlineStyle = true;
                styleOption.IsUseTextGlow = true;
                styleOption.TextGlowColor = "#000000";
            }
            return UpdateStyleName(
                result,
                TextCollection.PictureSlidesLabText.StyleNameOutline);
        }

        public override StyleOption GetDefaultOptionForPreview()
        {
            return new StyleOption()
            {
                StyleName = TextCollection.PictureSlidesLabText.StyleNameOutline,
                IsUseOutlineStyle = true,
                IsUseTextGlow = true,
                TextGlowColor = "#000000"
            };
        }
    }
}
