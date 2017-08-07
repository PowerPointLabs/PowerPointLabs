using System.Collections.Generic;
using System.ComponentModel.Composition;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    [Export(typeof(IStyleOptions))]
    [ExportMetadata("StyleOrder", 4)]
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
                PictureSlidesLabText.StyleNameFrostedGlassBanner);
        }

        public override StyleOption GetDefaultOptionForPreview()
        {
            return new StyleOption
            {
                StyleName = PictureSlidesLabText.StyleNameFrostedGlassBanner,
                IsUseFrostedGlassBannerStyle = true,
                FontFamily = "Segoe UI",
                TextBoxPosition = 4
            };
        }
    }
}
