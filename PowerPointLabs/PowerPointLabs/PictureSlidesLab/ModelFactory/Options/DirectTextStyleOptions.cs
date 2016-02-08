using System.Collections.Generic;
using PowerPointLabs.PictureSlidesLab.Model;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    class DirectTextStyleOptions : BaseStyleOptions
    {
        public override List<StyleOption> GetOptionsForVariation()
        {
            var result = GetOptions();
            foreach (var styleOption in result)
            {
                styleOption.IsUseTextGlow = true;
                styleOption.TextGlowColor = "#000000";
            }
            UpdateStyleName(
                result,
                TextCollection.PictureSlidesLabText.StyleNameDirectText);
            return result;
        }

        public override StyleOption GetDefaultOptionForPreview()
        {
            return new StyleOption
            {
                StyleName = TextCollection.PictureSlidesLabText.StyleNameDirectText,
                TextBoxPosition = 5,
                IsUseTextGlow = true,
                TextGlowColor = "#000000"
            };
        }
    }
}
