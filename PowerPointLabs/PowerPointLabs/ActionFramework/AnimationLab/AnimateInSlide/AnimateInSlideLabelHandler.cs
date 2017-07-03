using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.AnimationLab
{
    [ExportLabelRibbonId(TextCollection.AnimateInSlideTag)]
    class AnimateInSlideLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.AddAnimationInSlideAnimateButtonLabel;
        }
    }
}
