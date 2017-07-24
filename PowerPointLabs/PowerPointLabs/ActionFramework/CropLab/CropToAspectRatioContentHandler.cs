using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.CropLab
{
    [ExportContentRibbonId(CropLabText.CropToAspectRatioTag + CommonText.RibbonMenu)]
    class CropToAspectRatioContentHandler : ContentHandler
    {
        private static readonly string[] PRESET_ASPECT_RATIOS = { "1:1", "4:3", "16:9" };

        protected override string GetContent(string ribbonId)
        {
            var feature = ribbonId.Replace(CommonText.DynamicMenuId, "");

            var xmlString = new System.Text.StringBuilder();

            for (int i = 0; i < PRESET_ASPECT_RATIOS.Length; i++)
            {
                string idFriendlyAspectRatio = PRESET_ASPECT_RATIOS[i].Replace(':', '_');
                xmlString.Append(string.Format(CommonText.DynamicMenuXmlButton, 
                                feature + CommonText.DynamicMenuOptionId + idFriendlyAspectRatio,
                                CropLabText.CropToAspectRatioTag + CommonText.RibbonMenu));
            }

            xmlString.Append(string.Format(CommonText.DynamicMenuXmlMenuSeparator, feature));
            xmlString.Append(string.Format(CommonText.DynamicMenuXmlButton, 
                            feature + CommonText.DynamicMenuButtonId + "Custom",
                            CropLabText.CropToAspectRatioTag + CommonText.RibbonMenu));

            return string.Format(CommonText.DynamicMenuXmlMenu, xmlString);
        }
    }
}
