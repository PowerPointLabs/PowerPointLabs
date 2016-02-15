using System.Collections.Generic;
using System.ComponentModel.Composition;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    [Export(typeof(IStyleOptions))]
    [ExportMetadata("StyleOrder", 4)]
    class BannerStyleOptions : BaseStyleOptions
    {
        public override List<StyleOption> GetOptionsForVariation()
        {
            var options = new TextBoxStyleOptions().GetOptionsForVariation();
            foreach (var option in options)
            {
                option.FontFamily = "Times New Roman";
            }
            return UpdateStyleName(
                options,
                TextCollection.PictureSlidesLabText.StyleNameBanner);
        }

        public override StyleOption GetDefaultOptionForPreview()
        {
            return new StyleOption
            {
                StyleName = TextCollection.PictureSlidesLabText.StyleNameBanner,
                IsUseBannerStyle = true,
                TextBoxPosition = 7,
                TextBoxColor = "#000000",
                FontColor = "#FFD700",
                FontFamily = "Times New Roman"
            };
        }
    }
}
