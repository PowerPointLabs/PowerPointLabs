using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.ZoomLab
{
    [ExportLabelRibbonId("ZoomLabSettingsButton")]
    class ZoomLabSettingsLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.ZoomLabSettingsButtonLabel;
        }
    }
}
