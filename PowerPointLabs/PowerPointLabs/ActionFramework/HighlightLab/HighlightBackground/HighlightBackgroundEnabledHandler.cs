using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.HighlightLab
{
    [ExportEnabledRibbonId(HighlightLabText.HighlightBackgroundTag)]
    class HighlightBackgroundEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return this.GetRibbonUi().HighlightBulletsEnabled;
        }
    }
}