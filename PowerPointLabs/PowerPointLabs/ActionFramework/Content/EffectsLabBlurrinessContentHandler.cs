using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Content
{
    [ExportContentRibbonId("EffectsLabBlurSelectedDynamicMenu", "EffectsLabBlurRemainderDynamicMenu", "EffectsLabBlurBackgroundDynamicMenu")]
    class EffectsLabBlurrinessContentHandler : ContentHandler
    {
        protected override string GetContent(string ribbonId, string ribbonTag)
        {
            var feature = ribbonId.Replace("DynamicMenu", "");

            var xmlString = new System.Text.StringBuilder("<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">");

            for (int i = 0; i < 7; i++)
            {
                xmlString.Append("<button id=\"");
                xmlString.Append(feature);
                xmlString.Append("Option");
                xmlString.Append(i + 4);
                xmlString.Append("0\" tag=\"");
                xmlString.Append(TextCollection.EffectsLabBlurrinessTag);
                xmlString.Append("\" getLabel=\"GetLabel\" onAction=\"OnAction\"/>");
            }

            xmlString.Append("<menuSeparator id=\"");
            xmlString.Append(feature);
            xmlString.Append("Separator\"/>");

            xmlString.Append("<checkBox id=\"");
            xmlString.Append(feature);
            xmlString.Append("CheckBox\" tag=\"");
            xmlString.Append(TextCollection.EffectsLabBlurrinessTag);
            xmlString.Append("\" getLabel=\"GetLabel\" getPressed=\"GetPressed\" onAction=\"OnCheckBoxAction\"/>");

            xmlString.Append("<button id=\"");
            xmlString.Append(feature);
            xmlString.Append("Button\" tag=\"");
            xmlString.Append(TextCollection.EffectsLabBlurrinessTag);
            xmlString.Append("\" getLabel=\"GetLabel\" onAction=\"OnAction\"/>");

            xmlString.Append("</menu>");

            return xmlString.ToString();
        }
    }
}
