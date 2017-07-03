using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ZoomLab
{
    [ExportSupertipRibbonId(TextCollection.ZoomLabSettingsTag)]
    class ZoomLabSettingsSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.ZoomLabSettingsSupertip;
        }
    }
}
