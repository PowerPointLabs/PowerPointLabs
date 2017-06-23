using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.CropLab;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("CropLabSettingsButton")]
    class CropLabSettingsActionHandler : CropLabActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            CropLabSettingsDialogBox dialog = new CropLabSettingsDialogBox();
            dialog.ShowDialog();
        }
    }
}
