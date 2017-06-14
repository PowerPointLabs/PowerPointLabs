using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip.HelpMenu
{
    [ExportSupertipRibbonId("HelpMenu")]
    class HelpMenuSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.HelpMenuSupertip;
        }
    }
}
