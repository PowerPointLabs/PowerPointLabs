using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.HighlightLab
{
    [ExportSupertipRibbonId(HighlightLabText.HighlightPointsTag)]
    class HighlightPointsSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return HighlightLabText.HighlightBulletsTextButtonSupertip;
        }
    }
}
