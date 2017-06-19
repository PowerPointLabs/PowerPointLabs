using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.CropLab
{
    [ExportLabelRibbonId(TextCollection.CropLabSettingsTag)]
    class CropLabSettingsLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.CropLabSettingsButtonLabel;
        }
    }
}
