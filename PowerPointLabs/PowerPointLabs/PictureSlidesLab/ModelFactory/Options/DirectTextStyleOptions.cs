using System.Collections.Generic;
using System.ComponentModel.Composition;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    [Export(typeof(IStyleOptions))]
    [ExportMetadata("StyleOrder", 1)]
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
