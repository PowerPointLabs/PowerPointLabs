using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Content
{
    [ExportContentRibbonId(
        TextCollection.EffectsLabBlurrinessFeatureSelected + TextCollection.DynamicMenuId,
        TextCollection.EffectsLabBlurrinessFeatureRemainder + TextCollection.DynamicMenuId,
        TextCollection.EffectsLabBlurrinessFeatureBackground + TextCollection.DynamicMenuId)]
    class EffectsLabBlurrinessContentHandler : ContentHandler
    {
        protected override string GetContent(string ribbonId)
        {
            var feature = ribbonId.Replace(TextCollection.DynamicMenuId, "");

            var xmlString = new System.Text.StringBuilder();

            for (int i = 0; i < 7; i++)
            {
                xmlString.Append(string.Format(TextCollection.DynamicMenuXmlButton, feature + TextCollection.DynamicMenuOptionId + (i + 4) + "0",
                    TextCollection.EffectsLabBlurrinessTag));
            }

            xmlString.Append(string.Format(TextCollection.DynamicMenuXmlMenuSeparator, feature));
            xmlString.Append(string.Format(TextCollection.DynamicMenuXmlCheckBox, feature + TextCollection.DynamicMenuCheckBoxId,
                TextCollection.EffectsLabBlurrinessTag));
            xmlString.Append(string.Format(TextCollection.DynamicMenuXmlButton, feature + TextCollection.DynamicMenuButtonId,
                TextCollection.EffectsLabBlurrinessTag));

            return string.Format(TextCollection.DynamicMenuXmlMenu, xmlString);
        }
    }
}
