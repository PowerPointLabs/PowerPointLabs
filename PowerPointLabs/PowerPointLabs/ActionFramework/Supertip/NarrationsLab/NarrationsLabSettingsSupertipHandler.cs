using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip.NarrationsLab
{
    [ExportSupertipRibbonId("NarrationsLabSettingsButton")]
    class NarrationsLabSettingsSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.NarrationsLabSettingsSupertip;
        }
    }
}
