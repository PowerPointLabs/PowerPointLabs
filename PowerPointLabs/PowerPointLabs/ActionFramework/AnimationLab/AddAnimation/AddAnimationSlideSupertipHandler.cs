using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.AnimationLab
{
    [ExportSupertipRibbonId(AnimationLabText.AddAnimationSlideTag)]
    class AddAnimationSlideSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return AnimationLabText.AddAnimationButtonSupertip;
        }
    }
}
