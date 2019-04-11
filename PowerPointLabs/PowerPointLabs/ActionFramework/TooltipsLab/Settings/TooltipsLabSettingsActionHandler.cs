using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;
using PowerPointLabs.TooltipsLab;

namespace PowerPointLabs.ActionFramework.TooltipsLab.Settings
{
    [ExportActionRibbonId(TooltipsLabText.SettingsTag)]
    class TooltipsLabSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            TooltipsLabSettings.ShowSettingsDialog();
        }
    }
}