using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.AnimationLab
{
    [ExportEnabledRibbonId(AnimationLabText.AnimateInSlideTag)]
    class AnimateInSlideEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return IsSelectionMultipleSameShapeType();
        }
    }
}