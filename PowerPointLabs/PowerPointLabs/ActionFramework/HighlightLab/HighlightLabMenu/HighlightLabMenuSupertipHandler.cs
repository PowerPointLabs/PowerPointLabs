using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.Supertip.HighlightLab
{
    [ExportSupertipRibbonId(HighlightLabText.RibbonMenuId)]
    class HighlightLabMenuSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return HighlightLabText.RibbonMenuSupertip;
        }
    }
}
