using System.Collections.Generic;
using System.ComponentModel.Composition;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    [Export(typeof(IStyleOptions))]
    [ExportMetadata("StyleOrder", 9)]
    class OutlineStyleOptions : BaseStyleOptions
    {
        public override List<StyleOption> GetOptionsForVariation()
        {
            var result = GetOptionsWithSuitableFontColorForOutline();
            foreach (var styleOption in result)
            {
                styleOption.IsUseTextGlow = true;
                styleOption.TextGlowColor = "#000000";
                styleOption.FontFamily = "Segoe UI";
            }
            return UpdateStyleName(
                result,
                PictureSlidesLabText.StyleNameOutline);
        }

        public override StyleOption GetDefaultOptionForPreview()
        {
            return new StyleOption()
            {
                StyleName = PictureSlidesLabText.StyleNameOutline,
                IsUseOutlineStyle = true,
                IsUseTextGlow = true,
                TextGlowColor = "#000000",
                FontFamily = "Segoe UI"
            };
        }

        private List<StyleOption> GetOptionsWithSuitableFontColorForOutline()
        {
            var result = GetOptions();
            result[0].FontColor = "#FFFFFF"; //white
            result[1].FontColor = "#000000"; //black
            result[2].FontColor = "#FFD700"; //yellow
            result[3].FontColor = "#FF0000"; //red
            result[4].FontColor = "#3DFF8F"; //green
            result[5].FontColor = "#007FFF"; //blue
            result[6].FontColor = "#001550"; //dark blue
            result[7].FontColor = ""; //no effect
            return result;
        }
    }
}
