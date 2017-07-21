using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportContentRibbonId(
        TextCollection1.BlurSelectedMenuId,
        TextCollection1.BlurRemainderMenuId,
        TextCollection1.BlurBackgroundMenuId)]
    class EffectsLabBlurrinessContentHandler : ContentHandler
    {
        protected override string GetContent(string ribbonId)
        {
            string feature = ribbonId.Replace("Menu", "");

            System.Text.StringBuilder xmlString = new System.Text.StringBuilder();

            for (int i = 40; i <= 100; i += 10)
            {
                xmlString.Append(string.Format(TextCollection1.DynamicMenuXmlButton, 
                    feature + TextCollection1.DynamicMenuOptionId + i,
                    TextCollection1.EffectsLabBlurrinessTag));
            }
            xmlString.Append(string.Format(TextCollection1.DynamicMenuXmlButton,
                feature + TextCollection1.DynamicMenuOptionId + TextCollection1.EffectsLabBlurrinessCustom,
                TextCollection1.EffectsLabBlurrinessTag));

            xmlString.Append(string.Format(TextCollection1.DynamicMenuXmlMenuSeparator, 
                feature + TextCollection1.DynamicMenuOptionId + TextCollection1.DynamicMenuSeparatorId));
            xmlString.Append(string.Format(TextCollection1.DynamicMenuXmlButton, 
                feature + TextCollection1.DynamicMenuButtonId,
                TextCollection1.EffectsLabBlurrinessTag));

            return string.Format(TextCollection1.DynamicMenuXmlMenu, xmlString);
        }
    }
}
