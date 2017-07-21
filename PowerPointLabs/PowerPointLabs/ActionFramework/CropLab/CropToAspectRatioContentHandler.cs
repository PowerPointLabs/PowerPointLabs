using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.CropLab
{
    [ExportContentRibbonId(TextCollection1.CropToAspectRatioTag + TextCollection1.RibbonMenu)]
    class CropToAspectRatioContentHandler : ContentHandler
    {
        private static readonly string[] PRESET_ASPECT_RATIOS = { "1:1", "4:3", "16:9" };

        protected override string GetContent(string ribbonId)
        {
            var feature = ribbonId.Replace(TextCollection1.DynamicMenuId, "");

            var xmlString = new System.Text.StringBuilder();

            for (int i = 0; i < PRESET_ASPECT_RATIOS.Length; i++)
            {
                string idFriendlyAspectRatio = PRESET_ASPECT_RATIOS[i].Replace(':', '_');
                xmlString.Append(string.Format(TextCollection1.DynamicMenuXmlButton, 
                                feature + TextCollection1.DynamicMenuOptionId + idFriendlyAspectRatio,
                                TextCollection1.CropToAspectRatioTag + TextCollection1.RibbonMenu));
            }

            xmlString.Append(string.Format(TextCollection1.DynamicMenuXmlMenuSeparator, feature));
            xmlString.Append(string.Format(TextCollection1.DynamicMenuXmlButton, 
                            feature + TextCollection1.DynamicMenuButtonId + "Custom",
                            TextCollection1.CropToAspectRatioTag + TextCollection1.RibbonMenu));

            return string.Format(TextCollection1.DynamicMenuXmlMenu, xmlString);
        }
    }
}
