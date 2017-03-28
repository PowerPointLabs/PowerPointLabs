using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip
{
    [ExportSupertipRibbonId("CropLabSettingsButton")]
    class CropLabSettingsSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.CropLabSettingsSupertip;
        }
    }
}
