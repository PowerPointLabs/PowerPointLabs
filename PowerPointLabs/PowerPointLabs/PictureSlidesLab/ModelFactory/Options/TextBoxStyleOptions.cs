using System.Collections.Generic;
using System.ComponentModel.Composition;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    [Export(typeof(IStyleOptions))]
    [ExportMetadata("StyleOrder", 5)]
    class TextBoxStyleOptions : BaseStyleOptions
    {
        public override List<StyleOption> GetOptionsForVariation()
        {
            var result = GetOptionsWithSuitableFontColor();
            foreach (var option in result)
            {
                option.TextBoxPosition = 7; //bottom-left;
                option.FontFamily = "Calibri";
            }
            UpdateStyleName(
                result,
                PictureSlidesLabText.StyleNameTextBox);
            return result;
        }

        public override StyleOption GetDefaultOptionForPreview()
        {
            return new StyleOption
            {
                StyleName = PictureSlidesLabText.StyleNameTextBox,
                IsUseTextBoxStyle = true,
                TextBoxPosition = 7,
                TextBoxColor = "#000000",
                FontColor = "#FFD700",
                FontFamily = "Calibri"
            };
        }
    }
}
