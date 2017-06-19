using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.AnimationsLab
{
    [ExportSupertipRibbonId("AddAnimationSlide")]
    class AddAnimationSlideSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.AddAnimationButtonSupertip;
        }
    }
}
