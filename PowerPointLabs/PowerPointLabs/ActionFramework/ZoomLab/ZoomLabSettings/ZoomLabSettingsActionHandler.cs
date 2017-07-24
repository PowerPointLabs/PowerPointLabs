using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;
using PowerPointLabs.ZoomLab;

namespace PowerPointLabs.ActionFramework.ZoomLab
{
    [ExportActionRibbonId(ZoomLabText.SettingsTag)]
    class ZoomLabSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            ZoomLabSettings.ShowSettingsDialog();
        }
    }
}
