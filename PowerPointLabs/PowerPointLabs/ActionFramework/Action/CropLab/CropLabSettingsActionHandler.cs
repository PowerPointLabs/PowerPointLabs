using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.CropLab;

namespace PowerPointLabs.ActionFramework.Action.CropLab
{
    [ExportActionRibbonId("CropLabSettingsButton")]
    class CropLabSettingsActionHandler : CropLabActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var dialog = new CropLabSettingsDialog();
            dialog.ShowDialog();
        }
    }
}
