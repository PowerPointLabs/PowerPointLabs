using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.HighlightLab
{
    [ExportSupertipRibbonId(TextCollection.HighlightPointsTag)]
    class HighlightPointsSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.HighlightBulletsTextButtonSupertip;
        }
    }
}
