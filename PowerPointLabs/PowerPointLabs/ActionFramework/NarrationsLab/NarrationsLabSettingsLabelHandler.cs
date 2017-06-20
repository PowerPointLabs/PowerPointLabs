using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.NarrationsLab
{
    [ExportLabelRibbonId(TextCollection.NarrationsLabSettingsTag)]
    class NarrationsLabSettingsLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.NarrationsLabSettingsButtonLabel;
        }
    }
}
