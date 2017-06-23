using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip.HighlightLab
{
    [ExportSupertipRibbonId(TextCollection.HighlightLabSettingsTag)]
    class HighlightLabSettingsSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.HighlightLabSettingsSupertip;
        }
    }
}
