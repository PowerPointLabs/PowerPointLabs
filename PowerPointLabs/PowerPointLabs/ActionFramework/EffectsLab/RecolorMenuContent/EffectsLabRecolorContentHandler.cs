using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportContentRibbonId(TextCollection.RecolorRemainderMenuId, TextCollection.RecolorBackgroundMenuId)]
    class EffectsLabRecolorContentHandler : ContentHandler
    {
        private static string[] features =
        {
            TextCollection.GrayScaleTag, TextCollection.BlackWhiteTag,
            TextCollection.GothamTag, TextCollection.SepiaTag
        };

        protected override string GetContent(string ribbonId)
        {
            var xmlString = new System.Text.StringBuilder();

            foreach (string feature in features)
            {
                xmlString.Append(string.Format(TextCollection.DynamicMenuXmlButton,
                    feature + ribbonId, TextCollection.EffectsLabRecolorTag));
            }
           
            return string.Format(TextCollection.DynamicMenuXmlMenu, xmlString);
        }
    }
}
