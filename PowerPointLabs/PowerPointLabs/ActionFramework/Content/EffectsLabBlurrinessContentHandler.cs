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
            string feature = ribbonId.Replace(TextCollection.DynamicMenuId, "");

            System.Text.StringBuilder xmlString = new System.Text.StringBuilder();

            for (int i = 40; i <= 100; i += 10)
            {
                xmlString.Append(string.Format(TextCollection.DynamicMenuXmlButton, 
                    feature + TextCollection.DynamicMenuOptionId + i,
                    TextCollection.EffectsLabBlurrinessTag));
            }
            xmlString.Append(string.Format(TextCollection.DynamicMenuXmlButton,
                feature + TextCollection.DynamicMenuOptionId + TextCollection.EffectsLabBlurrinessCustom,
                TextCollection.EffectsLabBlurrinessTag));

            xmlString.Append(string.Format(TextCollection.DynamicMenuXmlMenuSeparator, 
                feature + TextCollection.DynamicMenuOptionId + TextCollection.DynamicMenuSeparatorId));
            xmlString.Append(string.Format(TextCollection.DynamicMenuXmlButton, 
                feature + TextCollection.DynamicMenuButtonId,
                TextCollection.EffectsLabBlurrinessTag));

            return string.Format(TextCollection.DynamicMenuXmlMenu, xmlString);
        }
    }
}
