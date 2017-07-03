using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.AnimationLab
{
    [ExportLabelRibbonId(TextCollection.AnimationLabSettingsTag)]
    class AnimationLabSettingsLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.AnimationLabSettingsButtonLabel;
        }
    }
}
