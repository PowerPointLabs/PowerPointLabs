using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.CropLab.Views;

namespace PowerPointLabs.ActionFramework.CropLab
{
    [ExportActionRibbonId(TextCollection.CropLabSettingsTag)]
    class CropLabSettingsActionHandler : CropLabActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            CropLabSettingsDialogBox dialog = new CropLabSettingsDialogBox();
            dialog.ShowDialog();
        }
    }
}
