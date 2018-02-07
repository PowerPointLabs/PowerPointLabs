using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.SaveLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.SaveLab
{
    [ExportActionRibbonId(SaveLabText.SaveLabSettingsButtonTag)]
    class SaveLabSettingsActionHandler : Common.Interface.ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            // Action for Save Lab settings here
            SaveLabSettings.ShowSettingsDialog();
        }
    }
}
