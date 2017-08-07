using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.HighlightLab
{
    [ExportSupertipRibbonId(HighlightLabText.HighlightBackgroundTag)]
    class HighlightBackgroundSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return HighlightLabText.HighlightBulletsBackgroundButtonSupertip;
        }
    }
}
