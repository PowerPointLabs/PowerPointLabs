using System.Collections.Generic;
using System.ComponentModel.Composition;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    [Export(typeof(IStyleOptions))]
    [ExportMetadata("StyleOrder", 7)]
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
