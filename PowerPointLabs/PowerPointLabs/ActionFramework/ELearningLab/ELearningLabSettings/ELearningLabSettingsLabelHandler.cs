using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ELearningLab
{
    [ExportLabelRibbonId(ELearningLabText.ELearningLabSettingsTag)]
    class ELearningLabSettingsLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return NarrationsLabText.SettingsButtonLabel;
        }
    }
}
