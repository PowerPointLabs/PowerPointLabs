using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ZoomLab
{
    [ExportLabelRibbonId(TextCollection.ZoomLabSettingsTag)]
    class ZoomLabSettingsLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.ZoomLabSettingsButtonLabel;
        }
    }
}
