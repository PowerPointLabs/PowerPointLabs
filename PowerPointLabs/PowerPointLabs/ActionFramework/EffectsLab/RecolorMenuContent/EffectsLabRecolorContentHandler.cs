using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportContentRibbonId(TextCollection1.RecolorRemainderMenuId, TextCollection1.RecolorBackgroundMenuId)]
    class EffectsLabRecolorContentHandler : ContentHandler
    {
        private static string[] features =
        {
            TextCollection1.GrayScaleTag, TextCollection1.BlackWhiteTag,
            TextCollection1.GothamTag, TextCollection1.SepiaTag
        };

        protected override string GetContent(string ribbonId)
        {
            var xmlString = new System.Text.StringBuilder();

            foreach (string feature in features)
            {
                xmlString.Append(string.Format(TextCollection1.DynamicMenuXmlButton,
                    feature + ribbonId, TextCollection1.EffectsLabRecolorTag));
            }
           
            return string.Format(TextCollection1.DynamicMenuXmlMenu, xmlString);
        }
    }
}
