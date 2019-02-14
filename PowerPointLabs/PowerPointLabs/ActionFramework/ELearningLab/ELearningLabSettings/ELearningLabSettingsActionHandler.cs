using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ELearningLab
{
    [ExportActionRibbonId(ELearningLabText.ELearningLabSettingsTag)]
    class ELearningLabSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            AzureAccountStorageService.LoadUserAccount();
            AudioSettingService.ShowSettingsDialog();
        }
    }
}
