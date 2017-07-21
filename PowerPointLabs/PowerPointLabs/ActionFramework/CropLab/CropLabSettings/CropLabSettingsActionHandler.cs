using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.CropLab;

namespace PowerPointLabs.ActionFramework.CropLab
{
    [ExportActionRibbonId(TextCollection1.CropLabSettingsTag)]
    class CropLabSettingsActionHandler : CropLabActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            CropLabSettings.ShowSettingsDialog();
        }
    }
}
