using System.Collections.Generic;
using System.ComponentModel.Composition;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    [Export(typeof(IStyleOptions))]
    [ExportMetadata("StyleOrder", 11)]
    class FrostedGlassBannerStyleOptions : BaseStyleOptions
    {
        public override List<StyleOption> GetOptionsForVariation()
        {
            var options = GetOptions();
            foreach (var option in options)
            {
                option.IsUseFrostedGlassBannerStyle = true;
                option.FontFamily = "Segoe UI";
                option.TextBoxPosition = 4; // left
            }
            return UpdateStyleName(
                options,
                TextCollection.PictureSlidesLabText.StyleNameFrostedGlassBanner);
        }

        public override StyleOption GetDefaultOptionForPreview()
        {
            return new StyleOption
            {
                StyleName = TextCollection.PictureSlidesLabText.StyleNameFrostedGlassBanner,
                IsUseFrostedGlassBannerStyle = true,
                TextBoxPosition = 4
            };
        }
    }
}
