using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.TooltipsLab.Settings
{
    [ExportSupertipRibbonId(TooltipsLabText.SettingsTag)]
    class TooltipsLabSettingsSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TooltipsLabText.SettingsButtonSupertip;
        }
    }
}