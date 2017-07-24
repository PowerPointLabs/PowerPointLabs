using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.CropLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.CropLab
{
    [ExportActionRibbonId(CropLabText.SettingsTag)]
    class CropLabSettingsActionHandler : CropLabActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            CropLabSettings.ShowSettingsDialog();
        }
    }
}
