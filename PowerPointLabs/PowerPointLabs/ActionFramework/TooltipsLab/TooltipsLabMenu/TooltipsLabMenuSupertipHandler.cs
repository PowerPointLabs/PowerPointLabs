using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.TooltipsLab
{
    [ExportSupertipRibbonId(TooltipsLabText.RibbonMenuId)]
    class TooltipsLabMenuSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return AnimationLabText.RibbonMenuSupertip;
        }
    }
}
