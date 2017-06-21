using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.CropLab;

namespace PowerPointLabs.ActionFramework.CropLab
{
    [ExportActionRibbonId(TextCollection.CropLabSettingsTag)]
    class CropLabSettingsActionHandler : CropLabActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var dialog = new CropLabSettingsDialog();
            dialog.ShowDialog();
        }
    }
}
