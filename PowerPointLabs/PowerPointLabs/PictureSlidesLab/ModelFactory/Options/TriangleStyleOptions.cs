using System.Collections.Generic;
using System.ComponentModel.Composition;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    [Export(typeof(IStyleOptions))]
    [ExportMetadata("StyleOrder", 12)]
    class TriangleStyleOptions : BaseStyleOptions
    {
        public override List<StyleOption> GetOptionsForVariation()
        {
            var result = GetOptionsWithSuitableFontColor();
            foreach (var styleOption in result)
            {
                styleOption.IsUseTriangleStyle = true;
                styleOption.TextBoxPosition = 4; // left
                styleOption.TriangleTransparency = 25;
                styleOption.FontFamily = "Times New Roman Italic";
            }
            UpdateStyleName(
                result,
                PictureSlidesLabText.StyleNameTriangle);
            return result;
        }

        public override StyleOption GetDefaultOptionForPreview()
        {
            return new StyleOption()
            {
                StyleName = PictureSlidesLabText.StyleNameTriangle,
                IsUseTriangleStyle = true,
                TriangleColor = "#007FFF", // blue
                TextBoxPosition = 4, // left
                TriangleTransparency = 25,
                FontFamily = "Times New Roman Italic"
            };
        }
    }
}
