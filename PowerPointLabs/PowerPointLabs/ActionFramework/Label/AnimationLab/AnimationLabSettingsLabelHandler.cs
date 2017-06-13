using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.AnimationLab
{
    [ExportLabelRibbonId("AnimationLabSettingsButton")]
    class AnimationLabSettingsLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.AnimationLabSettingsButtonLabel;
        }
    }
}
