using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.AnimationsLab
{
    [ExportLabelRibbonId("AnimateInSlide")]
    class AnimateInSlideLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.AddAnimationInSlideAnimateButtonLabel;
        }
    }
}
