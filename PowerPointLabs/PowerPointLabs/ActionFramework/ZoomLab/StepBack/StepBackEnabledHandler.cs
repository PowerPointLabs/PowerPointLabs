using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.AnimationLab
{
    [ExportEnabledRibbonId(ZoomLabText.StepBackTag)]
    class StepBackEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return IsSelectionSingleShape() && IsSelectionAllRectangle();
        }
    }
}