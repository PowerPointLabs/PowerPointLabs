using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportContentRibbonId(EffectsLabText.RecolorRemainderMenuId, EffectsLabText.RecolorBackgroundMenuId)]
    class EffectsLabRecolorContentHandler : ContentHandler
    {
        private static string[] features =
        {
            EffectsLabText.GrayScaleTag, EffectsLabText.BlackWhiteTag,
            EffectsLabText.GothamTag, EffectsLabText.SepiaTag
        };

        protected override string GetContent(string ribbonId)
        {
            var xmlString = new System.Text.StringBuilder();

            foreach (string feature in features)
            {
                xmlString.Append(string.Format(CommonText.DynamicMenuXmlButton,
                    feature + ribbonId, EffectsLabText.RecolorTag));
            }
           
            return string.Format(CommonText.DynamicMenuXmlMenu, xmlString);
        }
    }
}
