using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.TooltipsLab
{
    [ExportSupertipRibbonId(TooltipsLabText.AddTextboxTag)]
    class AddTextboxSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return AnimationLabText.AddAnimationButtonSupertip;
        }
    }
}
