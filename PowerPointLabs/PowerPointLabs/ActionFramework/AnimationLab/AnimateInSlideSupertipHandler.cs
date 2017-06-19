using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.AnimationLab
{
    [ExportSupertipRibbonId("AnimateInSlide")]
    class AnimateInSlideSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.InSlideAnimateButtonSupertip;
        }
    }
}
