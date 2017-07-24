using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.Help
{
    [ExportSupertipRibbonId(HelpText.AboutTag)]
    class AboutSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return HelpText.AboutButtonSupertip;
        }
    }
}
