using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.SaveLab
{
    [ExportSupertipRibbonId(SaveLabText.SaveLabSettingsButtonTag)]
    class SaveLabSettingsSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return SaveLabText.SaveLabSettingsButtonSupertip;
        }
    }
}
