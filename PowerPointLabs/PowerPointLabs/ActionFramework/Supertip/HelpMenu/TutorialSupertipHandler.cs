using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip.HelpMenu
{
    [ExportSupertipRibbonId("TutorialButton")]
    class TutorialSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.TutorialButtonSupertip;
        }
    }
}
