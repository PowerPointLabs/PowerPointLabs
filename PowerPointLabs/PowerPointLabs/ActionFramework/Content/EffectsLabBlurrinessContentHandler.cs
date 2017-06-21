using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.EffectsLab;

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

            for (int i = 0; i < 7; i++)
            {
                xmlString.Append(string.Format(TextCollection.DynamicMenuXmlButton, 
                    feature + TextCollection.DynamicMenuOptionId + (i + 4) + "0",
                    TextCollection.EffectsLabBlurrinessTag));
            }

            int customPercentage = 0;
            switch (feature)
            {
                case TextCollection.EffectsLabBlurrinessFeatureSelected:
                    customPercentage = EffectsLabBlurSelected.CustomPercentageSelected;
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureRemainder:
                    customPercentage = EffectsLabBlurSelected.CustomPercentageRemainder;
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureBackground:
                    customPercentage = EffectsLabBlurSelected.CustomPercentageBackground;
                    break;
                default:
                    Logger.Log(feature + " does not exist!", Common.Logger.LogType.Error);
                    break;
            }

            xmlString.Append(string.Format(TextCollection.DynamicMenuXmlMenuSeparator, 
                feature + TextCollection.DynamicMenuSeparatorId));
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
