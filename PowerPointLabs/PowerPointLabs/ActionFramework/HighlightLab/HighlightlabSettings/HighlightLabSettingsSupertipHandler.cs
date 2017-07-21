using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.Supertip.HighlightLab
{
    [ExportSupertipRibbonId(TextCollection1.HighlightLabSettingsTag)]
    class HighlightLabSettingsSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return HighlightLabText.SettingsButtonSupertip;
        }
    }
}
