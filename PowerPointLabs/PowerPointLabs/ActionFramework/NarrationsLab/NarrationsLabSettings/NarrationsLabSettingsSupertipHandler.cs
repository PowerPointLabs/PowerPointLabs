using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.NarrationsLab
{
    [ExportSupertipRibbonId(TextCollection.NarrationsLabSettingsTag)]
    class NarrationsLabSettingsSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.NarrationsLabSettingsSupertip;
        }
    }
}
