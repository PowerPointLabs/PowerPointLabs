using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.HighlightLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.HighlightLab
{
    [ExportEnabledRibbonId(HighlightLabText.HighlightBackgroundTag)]
    class HighlightBackgroundEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return HighlightBulletsBackground.IsHighlightBackgroundEnabled;
        }
    }
}