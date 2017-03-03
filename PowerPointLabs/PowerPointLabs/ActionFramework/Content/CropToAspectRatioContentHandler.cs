using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Content
{
    [ExportContentRibbonId("CropToAspectRatioDynamicMenu")]
    class CropToAspectRatioContentHandler : ContentHandler
    {
        private static readonly string[] PRESET_ASPECT_RATIOS = { "1:1", "4:3", "16:9" };

        protected override string GetContent(string ribbonId)
        {
            var feature = ribbonId.Replace(TextCollection.DynamicMenuId, "");

            var xmlString = new System.Text.StringBuilder();

            for (int i = 0; i < PRESET_ASPECT_RATIOS.Length; i++)
            {
                string idFriendlyAspectRatio = PRESET_ASPECT_RATIOS[i].Replace(':', '_');
                xmlString.Append(string.Format(TextCollection.DynamicMenuXmlButton, 
                                feature + TextCollection.DynamicMenuOptionId + idFriendlyAspectRatio,
                                TextCollection.CropToAspectRatioTag));
            }

            xmlString.Append(string.Format(TextCollection.DynamicMenuXmlMenuSeparator, feature));
            xmlString.Append(string.Format(TextCollection.DynamicMenuXmlButton, 
                            feature + TextCollection.DynamicMenuButtonId + "Custom",
                            TextCollection.CropToAspectRatioTag));

            return string.Format(TextCollection.DynamicMenuXmlMenu, xmlString);
        }
    }
}
