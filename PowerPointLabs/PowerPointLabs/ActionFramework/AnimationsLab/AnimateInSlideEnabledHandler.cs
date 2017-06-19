using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.AnimationsLab
{
    [ExportEnabledRibbonId("AnimateInSlide")]
    class AnimateInSlideEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return IsSelectionMultipleSameShapeType();
        }
    }
}