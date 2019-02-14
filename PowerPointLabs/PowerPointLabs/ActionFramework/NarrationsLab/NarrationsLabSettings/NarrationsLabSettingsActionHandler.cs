using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.NarrationsLab
{
    [ExportActionRibbonId(NarrationsLabText.SettingsTag)]
    class NarrationsLabSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            AzureAccountStorageService.LoadUserAccount();
            AudioSettingService.ShowSettingsDialog();
        }
    }
}
