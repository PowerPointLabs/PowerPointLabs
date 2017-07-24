using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.NarrationsLab
{
    [ExportSupertipRibbonId(NarrationsLabText.RibbonMenuId)]
    class NarrationsLabMenuSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return NarrationsLabText.RibbonMenuSupertip;
        }
    }
}
