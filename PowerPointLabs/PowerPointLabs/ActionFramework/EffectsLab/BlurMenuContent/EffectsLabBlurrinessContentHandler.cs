using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportContentRibbonId(
        EffectsLabText.BlurSelectedMenuId,
        EffectsLabText.BlurRemainderMenuId,
        EffectsLabText.BlurBackgroundMenuId)]
    class EffectsLabBlurrinessContentHandler : ContentHandler
    {
        protected override string GetContent(string ribbonId)
        {
            string feature = ribbonId.Replace("Menu", "");

            System.Text.StringBuilder xmlString = new System.Text.StringBuilder();

            for (int i = 40; i <= 100; i += 10)
            {
                xmlString.Append(string.Format(CommonText.DynamicMenuXmlButton, 
                    feature + CommonText.DynamicMenuOptionId + i,
                    EffectsLabText.BlurrinessTag));
            }
            xmlString.Append(string.Format(CommonText.DynamicMenuXmlButton,
                feature + CommonText.DynamicMenuOptionId + EffectsLabText.BlurrinessCustom,
                EffectsLabText.BlurrinessTag));

            xmlString.Append(string.Format(CommonText.DynamicMenuXmlMenuSeparator, 
                feature + CommonText.DynamicMenuOptionId + CommonText.DynamicMenuSeparatorId));
            xmlString.Append(string.Format(CommonText.DynamicMenuXmlButton, 
                feature + CommonText.DynamicMenuButtonId,
                EffectsLabText.BlurrinessTag));

            return string.Format(CommonText.DynamicMenuXmlMenu, xmlString);
        }
    }
}
