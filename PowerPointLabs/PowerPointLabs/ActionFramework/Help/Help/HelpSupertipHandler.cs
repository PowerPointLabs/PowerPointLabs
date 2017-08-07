using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.Supertip.Help
{
    [ExportSupertipRibbonId(HelpText.TutorialTag)]
    class TutorialSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return HelpText.UserGuideButtonSupertip;
        }
    }
}
