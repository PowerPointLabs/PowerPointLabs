using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip.ZoomLab
{
    [ExportSupertipRibbonId("ZoomLabSettingsButton")]
    class ZoomLabSettingsSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.ZoomLabSettingsSupertip;
        }
    }
}
