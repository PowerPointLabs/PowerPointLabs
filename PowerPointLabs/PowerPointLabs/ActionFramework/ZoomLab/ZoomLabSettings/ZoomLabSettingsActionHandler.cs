using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ZoomLab;

namespace PowerPointLabs.ActionFramework.ZoomLab
{
    [ExportActionRibbonId(TextCollection.ZoomLabSettingsTag)]
    class ZoomLabSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            ZoomLabSettings.OpenZoomLabSettingsDialog();
        }
    }
}
